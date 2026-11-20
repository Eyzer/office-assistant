"""sjz 文件解密。

sjz 容器本质是标准 AES-256-ECB：
  密文 = AES-256-ECB.encrypt(明文补齐到 16 字节倍数)
  明文长度记录在尾部定长记录的魔术 "\\x02.SJZ" 之前 20 字节（u32 小端）。

判定三层，任一层命中即视为 sjz 密文：
  1. 魔术  —— 密文尾部含 "\\x02.SJZ"
  2. FILL  —— 密文内出现连续 >= 4 个 FILL 块（FILL = AES(零块)，填充区特征）
  3. 试解  —— 文件头不是任何已知明文签名时试解一次，解出来是已知格式才认

AES 后端按速度优先级选择：
  cryptography 库 > pycryptodome 库 > openssl 命令 > 内置纯 Python
最后一级只用于兜底，约 1~2 MB/s；其余三级都在数十 MB/s 以上。
"""
import os
import shutil
import subprocess
import struct

KEY = b"xiaobaohappyxiaohaohappyxiaobaoh"
FILL = bytes.fromhex("db550aa6e23f0ad5a34998bcb40655c1")
SJZ_MAGIC = b"\x02.SJZ"

NEED = 4                             # 连续 FILL 块数：4 块 = 64 字节
FULL_SCAN_MAX = 8 * 1024 * 1024      # 超过此体积改头尾 + 抽样扫描
PROBE_MAX = 64 * 1024 * 1024         # 超过此体积不做试解

# 已知明文文件头签名
_PLAIN_SIGS = (
    b"PK\x03\x04", b"PK\x05\x06", b"PK\x07\x08",
    b"%PDF", b"\x89PNG\r\n\x1a\n", b"\xff\xd8\xff",
    b"GIF87a", b"GIF89a", b"D0CF11E0A1B11AE1",
    b"Rar!\x1a\x07", b"7z\xbc\xaf\x27\x1c", b"\x1f\x8b",
    b"\xef\xbb\xbf", b"{\\rtf", b"\x00\x00\x01\x00",
)


def _starts_with_sig(head):
    return any(head.startswith(s) for s in _PLAIN_SIGS)


# ---------------------------------------------------------------- AES 后端
_openssl_bin = None
_openssl_checked = False


def _openssl_path():
    global _openssl_bin, _openssl_checked
    if not _openssl_checked:
        _openssl_checked = True
        _openssl_bin = shutil.which("openssl")
    return _openssl_bin


def _openssl_decrypt(data):
    """走 openssl 命令，实测比纯 Python 快三个数量级。"""
    bin_ = _openssl_path()
    if not bin_:
        return None
    cmd = [bin_, "enc", "-d", "-aes-256-ecb", "-K", KEY.hex(), "-nopad", "-nosalt"]
    try:
        proc = subprocess.run(cmd, input=data, stdout=subprocess.PIPE,
                              stderr=subprocess.DEVNULL, timeout=600)
    except (OSError, subprocess.SubprocessError):
        return None
    if proc.returncode != 0 or len(proc.stdout) != len(data):
        return None
    return proc.stdout


def _aes_ecb_decrypt(data):
    """AES-256-ECB 解密（调用方保证长度是 16 的倍数）。"""
    if not data:
        return b""
    try:
        from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
        dec = Cipher(algorithms.AES(KEY), modes.ECB()).decryptor()
        return dec.update(data) + dec.finalize()
    except Exception:
        pass
    try:
        from Crypto.Cipher import AES as _AES
        return _AES.new(KEY, _AES.MODE_ECB).decrypt(data)
    except Exception:
        pass
    out = _openssl_decrypt(data)
    if out is not None:
        return out
    return _py_aes_ecb_decrypt(data)


# ------------------------------------------------- 纯 Python AES-256 兜底
def _gmul(a, b):
    r = 0
    for _ in range(8):
        if b & 1:
            r ^= a
        hi = a & 0x80
        a = (a << 1) & 0xFF
        if hi:
            a ^= 0x1B
        b >>= 1
    return r


def _build_tables():
    inv = [0] * 256
    for a in range(1, 256):
        for b in range(1, 256):
            if _gmul(a, b) == 1:
                inv[a] = b
                break
    sbox = bytearray(256)
    for a in range(256):
        x = inv[a]
        s = x
        for _ in range(4):
            x = ((x << 1) | (x >> 7)) & 0xFF
            s ^= x
        sbox[a] = s ^ 0x63
    inv_sbox = bytearray(256)
    for i in range(256):
        inv_sbox[sbox[i]] = i
    return bytes(sbox), bytes(inv_sbox)


_SBOX, _INV_SBOX = _build_tables()
_MUL9 = bytes(_gmul(i, 9) for i in range(256))
_MUL11 = bytes(_gmul(i, 11) for i in range(256))
_MUL13 = bytes(_gmul(i, 13) for i in range(256))
_MUL14 = bytes(_gmul(i, 14) for i in range(256))
_RCON = (0x01, 0x02, 0x04, 0x08, 0x10, 0x20, 0x40)


def _expand_key(key):
    """AES-256 密钥扩展 → 15 个 16 字节轮密钥。"""
    nk, nr = 8, 14
    w = [list(key[i * 4:i * 4 + 4]) for i in range(nk)]
    for i in range(nk, 4 * (nr + 1)):
        temp = list(w[i - 1])
        if i % nk == 0:
            temp = temp[1:] + temp[:1]
            temp = [_SBOX[b] for b in temp]
            temp[0] ^= _RCON[i // nk - 1]
        elif i % nk == 4:
            temp = [_SBOX[b] for b in temp]
        w.append([w[i - nk][j] ^ temp[j] for j in range(4)])
    return [bytes(b for word in w[r * 4:(r + 1) * 4] for b in word) for r in range(nr + 1)]


def _inv_shift_rows(s):
    return [s[0], s[13], s[10], s[7],
            s[4], s[1], s[14], s[11],
            s[8], s[5], s[2], s[15],
            s[12], s[9], s[6], s[3]]


def _inv_mix_columns(s):
    for c in range(4):
        i = c * 4
        a0, a1, a2, a3 = s[i], s[i + 1], s[i + 2], s[i + 3]
        s[i] = _MUL14[a0] ^ _MUL11[a1] ^ _MUL13[a2] ^ _MUL9[a3]
        s[i + 1] = _MUL9[a0] ^ _MUL14[a1] ^ _MUL11[a2] ^ _MUL13[a3]
        s[i + 2] = _MUL13[a0] ^ _MUL9[a1] ^ _MUL14[a2] ^ _MUL11[a3]
        s[i + 3] = _MUL11[a0] ^ _MUL13[a1] ^ _MUL9[a2] ^ _MUL14[a3]
    return s


def _py_aes_ecb_decrypt(data, rk=None):
    rk = rk or _expand_key(KEY)
    inv = _INV_SBOX
    out = bytearray(len(data))
    for off in range(0, len(data), 16):
        blk = data[off:off + 16]
        s = [blk[i] ^ rk[14][i] for i in range(16)]
        for rnd in range(13, -1, -1):
            s = _inv_shift_rows(s)
            s = [inv[b] for b in s]
            s = [s[i] ^ rk[rnd][i] for i in range(16)]
            if rnd:
                s = _inv_mix_columns(s)
        out[off:off + 16] = bytes(s)
    return bytes(out)


# ---------------------------------------------------------------- 尾部裁剪
def _sjz_true_len(buf):
    mk = buf.rfind(SJZ_MAGIC)
    if mk < 24:
        return None
    v = struct.unpack("<I", buf[mk - 20:mk - 16])[0]
    return v if 0 < v <= len(buf) else None


def _zip_true_len(buf):
    e = buf.rfind(b"PK\x05\x06")
    if e < 0 or e + 22 > len(buf):
        return None
    cmt = struct.unpack("<H", buf[e + 20:e + 22])[0]
    n = e + 22 + cmt
    return n if 0 < n <= len(buf) else None


def _sig_true_len(buf):
    n = len(buf)
    if n < 16:
        return None
    head = buf[:8]
    if head[:2] == b"\xff\xd8":
        p = buf.rfind(b"\xff\xd9")
        return p + 2 if 0 <= p <= n - 2 else None
    if head == b"\x89PNG\r\n\x1a\n":
        p = buf.rfind(b"IEND\xaeB\x60\x82")
        return p + 8 if 0 <= p <= n - 8 else None
    if head[:4] == b"%PDF":
        p = buf.rfind(b"%%EOF")
        if p < 0:
            return None
        t = p + 5
        while t < n and buf[t:t + 1] in (b"\r", b"\n"):
            t += 1
        return t
    if head[:6] in (b"GIF87a", b"GIF89a"):
        p = buf.rfind(b"\x3b")
        return p + 1 if 0 <= p < n else None
    return None


def plain_true_len(buf):
    """多级兜底求原始明文长度。"""
    for fn in (_sjz_true_len, _zip_true_len, _sig_true_len):
        v = fn(buf)
        if v is not None:
            return v
    return len(buf)


# ---------------------------------------------------------------- 判定
def _has_fill(buf):
    """连续 >= NEED 个 FILL 块。用 bytes.find 定位，比逐块比对快一个量级。"""
    step = NEED * 16
    if len(buf) < step:
        return False
    start = 0
    while True:
        i = buf.find(FILL, start)
        if i < 0 or i + step > len(buf):
            return False
        if i % 16 == 0:
            if all(buf[i + k * 16:i + (k + 1) * 16] == FILL for k in range(1, NEED)):
                return True
        start = i + 16


def _scan_fill(path, size):
    win = 65536
    blk = NEED * 16
    with open(path, "rb") as f:
        if size <= FULL_SCAN_MAX:
            pos = 0
            while pos < size:
                f.seek(pos)
                buf = f.read(win)
                if not buf:
                    break
                if _has_fill(buf):
                    return True
                if len(buf) < win:
                    break
                pos += win - blk
        else:
            ranges = [(0, 1024 * 1024), (max(0, size - 1024 * 1024), 1024 * 1024)]
            p = 1024 * 1024
            while p < size - 1024 * 1024:
                ranges.append((p, win))
                p += 4 * 1024 * 1024
            for start, length in ranges:
                length = min(length, size - start)
                if length < blk:
                    continue
                f.seek(start)
                if _has_fill(f.read(length)):
                    return True
    return False


def detect(path, probe=True):
    """判定是否为 sjz 密文，返回 (bool, 依据)。

    依据取值：magic / fill / probe / none
    """
    try:
        size = os.path.getsize(path)
    except OSError:
        return False, "none"
    if size < 64 or size % 16 != 0:
        return False, "none"
    try:
        with open(path, "rb") as f:
            head = f.read(8)
            f.seek(max(0, size - 8192))
            tail = f.read()
    except OSError:
        return False, "none"
    if SJZ_MAGIC in tail:
        return True, "magic"
    if _scan_fill(path, size):
        return True, "fill"
    if not probe or size > PROBE_MAX:
        return False, "none"
    # 已经是已知明文格式，不必试解
    if _starts_with_sig(head):
        return False, "none"
    try:
        with open(path, "rb") as f:
            plain = decrypt_bytes(f.read(), trim=False)
    except Exception:
        return False, "none"
    if plain and _starts_with_sig(plain[:8]):
        return True, "probe"
    return False, "none"


def is_encrypted(path):
    return detect(path, probe=False)[0]


# ---------------------------------------------------------------- 公开接口
def encrypt_bytes(data):
    """按 sjz 容器规则加密（仅供测试与自检，正常运行不需要）。"""
    pad = (-len(data)) % 16
    if pad:
        data = data + b"\x00" * pad
    try:
        from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
        enc = Cipher(algorithms.AES(KEY), modes.ECB()).encryptor()
        return enc.update(data) + enc.finalize()
    except Exception:
        pass
    try:
        from Crypto.Cipher import AES as _AES
        return _AES.new(KEY, _AES.MODE_ECB).encrypt(data)
    except Exception:
        pass
    bin_ = _openssl_path()
    if bin_:
        cmd = [bin_, "enc", "-e", "-aes-256-ecb", "-K", KEY.hex(), "-nopad", "-nosalt"]
        try:
            proc = subprocess.run(cmd, input=data, stdout=subprocess.PIPE,
                                  stderr=subprocess.DEVNULL, timeout=600)
            if proc.returncode == 0 and len(proc.stdout) == len(data):
                return proc.stdout
        except (OSError, subprocess.SubprocessError):
            pass
    raise RuntimeError("当前环境缺少可用的 AES 加密后端")


def decrypt_bytes(data, trim=True):
    """解密整份密文，返回明文。"""
    if not data:
        return b""
    pad = (-len(data)) % 16
    if pad:
        data = data + b"\x00" * pad
    out = _aes_ecb_decrypt(data)
    if trim:
        n = plain_true_len(out)
        if 0 < n < len(out):
            out = out[:n]
    return out


def decrypt_file(src, dst=None, probe=True):
    """解密单个文件。

    返回 dict：ok / changed / src / dst / size / reason / via
    changed=False 表示不是 sjz 密文，文件原样保留。
    """
    res = {"ok": False, "changed": False, "src": src, "dst": src,
           "size": 0, "reason": "", "via": "none"}
    if not os.path.isfile(src):
        res["reason"] = "文件不存在"
        return res
    hit, via = detect(src, probe=probe)
    if not hit:
        res.update({"ok": True, "size": os.path.getsize(src)})
        return res
    with open(src, "rb") as f:
        raw = f.read()
    try:
        plain = decrypt_bytes(raw, trim=False)
    except Exception as exc:
        res["reason"] = "解密异常：%s" % exc
        return res
    # 裁剪尾部填充
    n = plain_true_len(plain)
    if 0 < n < len(plain):
        plain = plain[:n]
    if not plain:
        res["reason"] = "解密结果为空"
        return res
    target = dst or src
    tmp = target + ".sjztmp"
    try:
        with open(tmp, "wb") as f:
            f.write(plain)
        os.replace(tmp, target)
    except OSError as exc:
        try:
            os.remove(tmp)
        except OSError:
            pass
        res["reason"] = "写入失败：%s" % exc
        return res
    res.update({"ok": True, "changed": True, "dst": target,
                "size": len(plain), "via": via})
    return res


def backend_name():
    try:
        import cryptography
        return "cryptography %s" % cryptography.__version__
    except Exception:
        pass
    try:
        import Crypto
        return "pycryptodome %s" % getattr(Crypto, "__version__", "")
    except Exception:
        pass
    if _openssl_path():
        return "openssl 命令"
    return "纯 Python（慢）"


def self_check():
    """自检：后端身份 + 一次标准向量往返。"""
    import binascii
    ok = True
    try:
        pt = binascii.unhexlify("00112233445566778899aabbccddeeff")
        ct = binascii.unhexlify("8ea2b7ca516745bfeafc49904b496089")
        key = binascii.unhexlify("000102030405060708090a0b0c0d0e0f"
                                 "101112131415161718191a1b1c1d1e1f")
        ok = _py_aes_ecb_decrypt(ct, _expand_key(key)) == pt
    except Exception:
        ok = False
    return {
        "backend": backend_name(),
        "aes_vector": "PASS" if ok else "FAIL",
        "key_len": len(KEY),
        "fill_hex": binascii.hexlify(FILL).decode(),
    }
