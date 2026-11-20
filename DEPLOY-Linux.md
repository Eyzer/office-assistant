# Linux 部署与测试

## 一、解压

```sh
tar xzf office-assistant-linux-1.3.2.tar.gz
cd office-assistant
```

## 二、自检

```sh
sh check.sh
```

正常输出长这样（关键看最后一行 `结论：就绪`）：

```
Python: Python 3.9.2  ->  /usr/bin/python3
openpyxl      3.1.5
python-docx   1.2.0

sjz 解密后端  openssl 命令
AES 标准向量  PASS
密钥长度      32 字节
FILL 判据     db550aa6e23f0ad5a34998bcb40655c1

结论：就绪
```

`sjz 解密后端` 如果不是 `openssl 命令` 或 `cryptography`，说明系统没装 openssl，解密会退回纯 Python，**能用但很慢**（约 1MB/s）。装一下即可：`sudo apt install openssl`。

## 三、启动

**本机浏览器用**（桌面麒麟）：

```sh
sh start.sh
```

**从另一台机器访问**（虚拟机 / 服务器）：

```sh
sh start.sh --host 0.0.0.0 --port 8765
```

然后浏览器打开 `http://服务器IP:8765/?t=<启动时打印的令牌>`。

启动脚本会打印完整带令牌的地址，直接复制。

**让它常驻**（后台跑）：

```sh
nohup sh start.sh --host 0.0.0.0 --port 8765 --token 你的令牌 > office.log 2>&1 &
```

`--token` 固定下来，才能把地址存成书签。不指定则每次启动随机生成。

## 四、sjz 解密怎么测

功能已经接进上传流程，**不需要任何开关**：把 sjz 加密的文件拖进页面，服务端检测到就自动解密，然后用明文参与后续处理。

判定依据有三个，命中任一即解：

| 依据 | 说明 |
|---|---|
| `magic` | 文件尾部有 `\x02.SJZ` 标记 |
| `fill` | 文件里有连续 4 个及以上 FILL 块（sjz 填充区特征） |
| `probe` | 文件头不是任何已知格式，试解一次，解出来是 docx/xlsx/pdf 等才认 |

测试步骤：

1. 拖一个 sjz 加密的 docx 到页面
2. 底部日志抽屉会出一条：`sjz 已解密：xxx.docx（判据 fill，104901B）`
3. 页面上该文件的体积会变成解密后的大小
4. 接着随便跑一个功能（比如「列出文件清单」），能正常处理说明解出来的确实是有效文件

**普通文件不受影响**：未加密的 docx / xlsx / pdf / csv 一律零改动，日志里也不会有 sjz 字样。

## 五、已验证的事实

- AES 实现过 FIPS-197 标准向量（AES-256）
- 用站上真实样本验证过：3317760 字节密文 → 104901 字节，明文头是 `PK\x03\x04`，与 PHP 版结果一致
- 解密耗时 0.034 秒（openssl 后端）
- 121 个单元测试全绿，其中 9 个专测 sjz

## 六、出问题怎么查

| 现象 | 原因 |
|---|---|
| 页面空白 / 一直转圈 | 服务没起来，看终端有没有报错 |
| 所有接口 403 | 地址里的 `?t=` 令牌丢了，回终端复制完整地址 |
| 提示「后台服务是旧版本」 | 浏览器缓存了旧页面，Ctrl+F5 强刷 |
| sjz 文件没被解 | 看日志有没有 `sjz 识别到但未解密`，有的话括号里是原因 |
| 大文件解得很慢 | 没装 openssl，`sudo apt install openssl` |
