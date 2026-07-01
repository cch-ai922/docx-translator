# -*- coding: utf-8 -*-
"""
chm_reader.py  --  pure-stdlib CHM (Compiled HTML Help) reader / decompiler.

Parses the Microsoft ITSS/ITSF container (ITSF -> ITSP -> PMGL directory),
decompresses the LZX content section, and extracts every internal file to a
folder.  This is the *read* half of a CHM translator: decompile -> translate
the HTML with html_translator -> (recompile with the external HHW `hhc.exe`,
which a pure-Python `.chm` writer cannot replace -- LZX *compression* isn't
shipped in the stdlib).

Public API:
    CHMFile(path)          -- open and parse
      .names()             -- list internal file names
      .read(name)          -- bytes of one internal file
      .extract_all(dst)    -- write every internal file under dst/
    extract_chm(path, dst) -- convenience wrapper

Standard library only (struct, os).  Python 3.7+.

Scope / limits
--------------
* LZX "Intel E8" call-translation is not implemented; it is essentially never
  enabled for HTML help content (the stream flag is 0). If a CHM does enable it
  the reader raises NotImplementedError rather than emit corrupt bytes.
* PMGI index chunks are not needed: PMGL listing chunks are walked directly.
"""

import os
import struct

# --------------------------------------------------------------------------
# little-endian scalar helpers + CHM's variable-length ENCINT
# --------------------------------------------------------------------------

def _u16(d, o): return struct.unpack_from("<H", d, o)[0]
def _u32(d, o): return struct.unpack_from("<I", d, o)[0]
def _i32(d, o): return struct.unpack_from("<i", d, o)[0]
def _u64(d, o): return struct.unpack_from("<Q", d, o)[0]


def _encint(d, p):
    """Big-endian base-128 varint (high bit = continue). Returns (value, newpos)."""
    v = 0
    while True:
        b = d[p]
        p += 1
        v = (v << 7) | (b & 0x7F)
        if not (b & 0x80):
            return v, p


# ==========================================================================
# LZX decompression (the CHM/CAB variant)
# ==========================================================================

_LZX_NUM_CHARS = 256
_LZX_PRETREE_NUM = 20
_LZX_NUM_PRIMARY_LENGTHS = 7
_LZX_NUM_SECONDARY_LENGTHS = 249
_LZX_MIN_MATCH = 2

# position-slot base + extra-bit tables (51 slots, enough through a 2 MB window)
_EXTRA_BITS = []
_j = 0
for _i in range(0, 52, 2):
    _EXTRA_BITS.append(_j)
    _EXTRA_BITS.append(_j)
    if _i != 0 and _j < 17:
        _j += 1
_EXTRA_BITS = _EXTRA_BITS[:51]

_POSITION_BASE = []
_j = 0
for _i in range(51):
    _POSITION_BASE.append(_j)
    _j += 1 << _EXTRA_BITS[_i]


class _BitReader(object):
    """LZX bitstream: 16-bit little-endian words, bits consumed MSB-first."""
    __slots__ = ("d", "pos", "buf", "n")

    def __init__(self, data):
        self.d = data
        self.pos = 0
        self.buf = 0
        self.n = 0

    def _fill(self):
        d, p = self.d, self.pos
        if p + 1 < len(d):
            word = d[p] | (d[p + 1] << 8)
            self.pos = p + 2
        elif p < len(d):
            word = d[p]
            self.pos = p + 1
        else:
            word = 0
        self.buf = (self.buf << 16) | word
        self.n += 16

    def bits(self, k):
        if k == 0:
            return 0
        while self.n < k:
            self._fill()
        self.n -= k
        val = (self.buf >> self.n) & ((1 << k) - 1)
        self.buf &= (1 << self.n) - 1
        return val

    def frame_align(self):
        """Realign to the next 16-bit boundary (LZX pads the bitstream at every
        32768-byte output frame boundary)."""
        drop = self.n % 16
        if drop:
            self.bits(drop)

    def align16(self):
        """Discard to the next 16-bit boundary and expose a raw byte cursor."""
        drop = self.n % 16
        if drop:
            self.bits(drop)
        self.pos -= self.n // 8      # un-read the fully buffered words
        self.buf = 0
        self.n = 0

    def read_u32(self):
        v = self.d[self.pos] | (self.d[self.pos + 1] << 8) \
            | (self.d[self.pos + 2] << 16) | (self.d[self.pos + 3] << 24)
        self.pos += 4
        return v

    def read_byte(self):
        b = self.d[self.pos]
        self.pos += 1
        return b


def _build_huff(lengths):
    """Canonical Huffman decode table from a list of code lengths (0 = unused)."""
    maxlen = 0
    for l in lengths:
        if l > maxlen:
            maxlen = l
    cnt = [0] * (maxlen + 1)
    by_len = [[] for _ in range(maxlen + 1)]
    for sym, l in enumerate(lengths):
        if l > 0:
            cnt[l] += 1
            by_len[l].append(sym)
    first_code = [0] * (maxlen + 2)
    first_sym = [0] * (maxlen + 2)
    sorted_syms = []
    code = 0
    s = 0
    for l in range(1, maxlen + 1):
        first_code[l] = code
        first_sym[l] = s
        for sym in by_len[l]:
            sorted_syms.append(sym)
            s += 1
        code = (code + cnt[l]) << 1
    return (maxlen, first_code, first_sym, sorted_syms, cnt)


def _decode_huff(br, table):
    maxlen, first_code, first_sym, sorted_syms, cnt = table
    code = 0
    for l in range(1, maxlen + 1):
        code = (code << 1) | br.bits(1)
        c = cnt[l]
        if c:
            idx = code - first_code[l]
            if 0 <= idx < c:
                return sorted_syms[first_sym[l] + idx]
    raise ValueError("bad LZX Huffman code")


class _Lzx(object):
    """Decode one independent (reset) LZX interval."""

    def __init__(self, window_bits):
        self.window_bits = window_bits
        if window_bits == 20:
            posn_slots = 42
        elif window_bits == 21:
            posn_slots = 50
        else:
            posn_slots = window_bits * 2
        self.main_maxsym = _LZX_NUM_CHARS + posn_slots * 8

    def _read_lens(self, br, lens, first, last):
        pre = [br.bits(4) for _ in range(_LZX_PRETREE_NUM)]
        ptab = _build_huff(pre)
        x = first
        while x < last:
            z = _decode_huff(br, ptab)
            if z == 17:
                y = br.bits(4) + 4
                for _ in range(y):
                    lens[x] = 0
                    x += 1
            elif z == 18:
                y = br.bits(5) + 20
                for _ in range(y):
                    lens[x] = 0
                    x += 1
            elif z == 19:
                y = br.bits(1) + 4
                z2 = _decode_huff(br, ptab)
                val = (lens[x] - z2) % 17
                for _ in range(y):
                    lens[x] = val
                    x += 1
            else:
                lens[x] = (lens[x] - z) % 17
                x += 1

    def decode(self, comp, out_len):
        br = _BitReader(comp)
        if br.bits(1):                       # Intel E8 translation flag
            raise NotImplementedError("LZX Intel E8 translation not supported")

        win = bytearray(out_len)
        pos = 0
        frame_end = _LZX_FRAME             # realign the bitstream at each frame
        R0 = R1 = R2 = 1
        main_lens = [0] * self.main_maxsym
        len_lens = [0] * _LZX_NUM_SECONDARY_LENGTHS
        main_tab = len_tab = aligned_tab = None
        block_type = 0
        block_remaining = 0

        while pos < out_len:
            if pos == frame_end:
                br.frame_align()
                frame_end += _LZX_FRAME

            if block_remaining == 0:
                block_type = br.bits(3)
                block_remaining = (br.bits(16) << 8) | br.bits(8)
                if block_type in (1, 2):                 # verbatim / aligned
                    if block_type == 2:
                        aligned_tab = _build_huff([br.bits(3) for _ in range(8)])
                    self._read_lens(br, main_lens, 0, _LZX_NUM_CHARS)
                    self._read_lens(br, main_lens, _LZX_NUM_CHARS, self.main_maxsym)
                    main_tab = _build_huff(main_lens)
                    self._read_lens(br, len_lens, 0, _LZX_NUM_SECONDARY_LENGTHS)
                    len_tab = _build_huff(len_lens)
                elif block_type == 3:                    # uncompressed
                    br.align16()
                    R0 = br.read_u32()
                    R1 = br.read_u32()
                    R2 = br.read_u32()
                else:
                    raise ValueError("bad LZX block type %d" % block_type)

            if block_type == 3:
                take = min(block_remaining, frame_end - pos, out_len - pos)
                for _ in range(take):
                    win[pos] = br.read_byte()
                    pos += 1
                block_remaining -= take
                if block_remaining == 0 and (take & 1):
                    br.read_byte()                       # pad to 16-bit boundary
                    br.buf = 0
                    br.n = 0
                continue

            # verbatim / aligned block: decode symbols up to the frame boundary
            limit = min(frame_end, out_len)
            while block_remaining > 0 and pos < limit:
                sym = _decode_huff(br, main_tab)
                if sym < _LZX_NUM_CHARS:
                    win[pos] = sym
                    pos += 1
                    block_remaining -= 1
                    continue
                sym -= _LZX_NUM_CHARS
                match_len = sym & _LZX_NUM_PRIMARY_LENGTHS
                if match_len == _LZX_NUM_PRIMARY_LENGTHS:
                    match_len += _decode_huff(br, len_tab)
                match_len += _LZX_MIN_MATCH
                slot = sym >> 3
                if slot == 0:
                    off = R0
                elif slot == 1:
                    off = R1
                    R1 = R0
                    R0 = off
                elif slot == 2:
                    off = R2
                    R2 = R0
                    R0 = off
                else:
                    extra = _EXTRA_BITS[slot]
                    if block_type == 2 and extra >= 3:
                        vb = (br.bits(extra - 3) << 3) if extra > 3 else 0
                        ab = _decode_huff(br, aligned_tab)
                        off = _POSITION_BASE[slot] - 2 + vb + ab
                    else:
                        off = _POSITION_BASE[slot] - 2 + (br.bits(extra) if extra else 0)
                    R2 = R1
                    R1 = R0
                    R0 = off
                src = pos - off
                block_remaining -= match_len
                for _ in range(match_len):
                    win[pos] = win[src]
                    pos += 1
                    src += 1

        return bytes(win)


# ==========================================================================
# CHM container
# ==========================================================================

_RESET_TABLE = ("::DataSpace/Storage/MSCompressed/Transform/"
                "{7FC28940-9D31-11D0-9B27-00A0C91E9C7C}/InstanceData/ResetTable")
_CONTROL_DATA = "::DataSpace/Storage/MSCompressed/ControlData"
_CONTENT = "::DataSpace/Storage/MSCompressed/Content"
_LZX_FRAME = 0x8000


class CHMFile(object):
    def __init__(self, path):
        with open(path, "rb") as f:
            self.data = f.read()
        d = self.data
        if d[0:4] != b"ITSF":
            raise ValueError("not a CHM file (no ITSF header)")
        self.version = _u32(d, 4)
        sec1_off = _u64(d, 72)
        self.content_off = _u64(d, 88)
        if d[sec1_off:sec1_off + 4] != b"ITSP":
            raise ValueError("bad CHM directory (no ITSP header)")
        itsp_hdrlen = _u32(d, sec1_off + 8)
        chunk_size = _u32(d, sec1_off + 16)
        first_pmgl = _i32(d, sec1_off + 32)
        last_pmgl = _i32(d, sec1_off + 36)
        dir_off = sec1_off + itsp_hdrlen

        self.entries = {}          # name -> (section, offset, length)
        for ci in range(first_pmgl, last_pmgl + 1):
            base = dir_off + ci * chunk_size
            if d[base:base + 4] != b"PMGL":
                continue
            free = _u32(d, base + 4)
            end = base + chunk_size - free
            p = base + 20
            while p < end:
                nl, p = _encint(d, p)
                name = d[p:p + nl].decode("utf-8", "replace")
                p += nl
                sec, p = _encint(d, p)
                off, p = _encint(d, p)
                ln, p = _encint(d, p)
                self.entries[name] = (sec, off, ln)

        self._lzx_blob = None      # decompressed MSCompressed section (lazy)

    # -- raw section-0 (uncompressed) access -------------------------------
    def _read_raw(self, name):
        sec, off, ln = self.entries[name]
        if sec != 0:
            raise ValueError("%s is not in the uncompressed section" % name)
        base = self.content_off + off
        return self.data[base:base + ln]

    # -- LZX section -------------------------------------------------------
    def _decompress_section(self):
        cd = self._read_raw(_CONTROL_DATA)
        if cd[4:8] != b"LZXC":
            raise ValueError("unexpected compression (not LZXC)")
        version = _u32(cd, 8)
        reset_interval = _u32(cd, 12)
        window_size = _u32(cd, 16)
        if version == 2:
            reset_interval *= _LZX_FRAME
            window_size *= _LZX_FRAME
        window_bits = window_size.bit_length() - 1     # log2 of a power of two

        rt = self._read_raw(_RESET_TABLE)
        n_entries = _u32(rt, 4)
        tbl_hdr = _u32(rt, 12)
        uncomp_len = _u64(rt, 16)
        comp_len = _u64(rt, 24)
        frame = _u64(rt, 32) or _LZX_FRAME
        reset_off = [_u64(rt, tbl_hdr + i * 8) for i in range(n_entries)]

        comp = self._read_raw(_CONTENT)
        ri_frames = max(1, reset_interval // frame)
        n_frames = (uncomp_len + frame - 1) // frame

        out = bytearray()
        i = 0
        while i < n_frames:
            start = reset_off[i]
            nxt = i + ri_frames
            stop = reset_off[nxt] if nxt < n_frames else comp_len
            out_len = min(ri_frames * frame, uncomp_len - len(out))
            out += _Lzx(window_bits).decode(comp[start:stop], out_len)
            i = nxt
        return bytes(out[:uncomp_len])

    def _blob(self):
        if self._lzx_blob is None:
            self._lzx_blob = self._decompress_section()
        return self._lzx_blob

    # -- public ------------------------------------------------------------
    def read(self, name):
        sec, off, ln = self.entries[name]
        if ln == 0:
            return b""
        if sec == 0:
            base = self.content_off + off
            return self.data[base:base + ln]
        return self._blob()[off:off + ln]

    def names(self, content_only=True):
        out = []
        for name, (sec, off, ln) in self.entries.items():
            if name.endswith("/"):
                continue                       # directory marker
            if content_only and (name.startswith("::") or name.startswith("/#")
                                 or name in ("/", "/$OBJINST")):
                continue
            out.append(name)
        return sorted(out)

    def extract_all(self, dst_dir, content_only=True):
        written = []
        for name in self.names(content_only=content_only):
            rel = name.lstrip("/")
            target = os.path.join(dst_dir, *rel.split("/"))
            os.makedirs(os.path.dirname(target) or dst_dir, exist_ok=True)
            with open(target, "wb") as f:
                f.write(self.read(name))
            written.append(target)
        return written


def extract_chm(chm_path, dst_dir, content_only=True):
    """Decompile a .chm into dst_dir; returns the list of written paths."""
    os.makedirs(dst_dir, exist_ok=True)
    return CHMFile(chm_path).extract_all(dst_dir, content_only=content_only)


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------

def _main(argv):
    if len(argv) == 1 and argv[0] not in ("-h", "--help"):
        chm = CHMFile(argv[0])
        for n in chm.names():
            print("%10d  %s" % (chm.entries[n][2], n))
        return 0
    if len(argv) == 2:
        paths = extract_chm(argv[0], argv[1])
        print("extracted %d files to %s" % (len(paths), argv[1]))
        return 0
    print("usage:\n"
          "  python chm_reader.py file.chm             # list internal files\n"
          "  python chm_reader.py file.chm  out_dir/   # decompile to a folder")
    return 0


if __name__ == "__main__":
    import sys
    sys.exit(_main(sys.argv[1:]))
