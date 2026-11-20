"""sjz 解密模块的单元测试。

样本不依赖外部文件：用模块自带的 encrypt_bytes 现造密文。
"""
import binascii
import os
import tempfile
import unittest

from utils import sjz


class AesVectorTest(unittest.TestCase):
    """纯 Python 实现必须过 FIPS-197 标准向量，否则整个功能不可信。"""

    def test_fips197_aes256(self):
        pt = binascii.unhexlify("00112233445566778899aabbccddeeff")
        ct = binascii.unhexlify("8ea2b7ca516745bfeafc49904b496089")
        key = binascii.unhexlify(
            "000102030405060708090a0b0c0d0e0f"
            "101112131415161718191a1b1c1d1e1f")
        rk = sjz._expand_key(key)
        self.assertEqual(sjz._py_aes_ecb_decrypt(ct, rk), pt)

    def test_fill_is_encrypted_zero_block(self):
        """FILL 必须等于零块的密文，否则判定逻辑没有依据。"""
        self.assertEqual(sjz.encrypt_bytes(b"\x00" * 16), sjz.FILL)

    def test_self_check(self):
        info = sjz.self_check()
        self.assertEqual(info["aes_vector"], "PASS")
        self.assertEqual(info["key_len"], 32)


class DetectTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="sjzt_")

    def tearDown(self):
        import shutil
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _write(self, name, data):
        p = os.path.join(self.tmp, name)
        with open(p, "wb") as f:
            f.write(data)
        return p

    def test_zip_like_payload_roundtrip(self):
        """造一个 zip 头 + 尾部 4KB 零填充的文件，模拟 sjz 容器。"""
        body = b"PK\x03\x04" + os.urandom(3000)
        plain = body + b"\x00" * 4096
        p = self._write("a.docx.sjz", sjz.encrypt_bytes(plain))
        hit, via = sjz.detect(p)
        self.assertTrue(hit, "尾部 4KB 零填充应命中 FILL 判据")
        self.assertIn(via, ("fill", "magic", "probe"))
        r = sjz.decrypt_file(p)
        self.assertTrue(r["ok"])
        self.assertTrue(r["changed"])
        with open(p, "rb") as f:
            out = f.read()
        self.assertTrue(out.startswith(b"PK\x03\x04"))

    def test_plain_file_untouched(self):
        for name, data in (
            ("a.docx", b"PK\x03\x04" + os.urandom(500)),
            ("a.pdf", b"%PDF-1.4\n" + os.urandom(500)),
            ("a.csv", "序号,名称\n1,测试\n".encode("utf-8") * 20),
        ):
            p = self._write(name, data)
            hit, _ = sjz.detect(p)
            self.assertFalse(hit, name + " 不该被判定为 sjz 密文")
            r = sjz.decrypt_file(p)
            self.assertTrue(r["ok"])
            self.assertFalse(r["changed"])
            with open(p, "rb") as f:
                self.assertEqual(f.read(), data, name + " 内容被改动了")

    def test_tiny_file_skipped(self):
        p = self._write("tiny.bin", b"\x00" * 8)
        self.assertFalse(sjz.is_encrypted(p))

    def test_missing_file(self):
        r = sjz.decrypt_file(os.path.join(self.tmp, "nope.bin"))
        self.assertFalse(r["ok"])
        self.assertIn("不存在", r["reason"])

    def test_zip_true_len(self):
        # EOCD 后无注释，长度 = EOCD 起点 + 22
        blob = b"PK\x03\x04" + b"x" * 100 + b"PK\x05\x06" + b"\x00" * 18
        self.assertEqual(sjz._zip_true_len(blob), len(blob))

    def test_sjz_magic_true_len(self):
        # 布局：… [u32 原长][16 字节][\x02.SJZ]，原长取魔术前 20 字节
        import struct
        body = b"A" * 100
        buf = body + struct.pack("<I", 100) + b"\x00" * 16 + sjz.SJZ_MAGIC
        self.assertEqual(sjz._sjz_true_len(buf), 100)


if __name__ == "__main__":
    unittest.main()
