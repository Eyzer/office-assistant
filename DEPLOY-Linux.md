# Linux 部署与测试

## 一、解压

```sh
tar xzf office-assistant-intranet-aarch64-1.3.3.tar.gz
cd office-assistant
```

如果拿的是内网版（`office-assistant-intranet-*.tar.gz`），解压出来的目录里已经带 `offline-deps/wheels/`，依赖不用另外准备，直接看第二节。

## 二、依赖（内网必看）

程序只依赖 Python 标准库，第三方库全部可选离线安装。有三种情况：

### 情况 1：用内网版压缩包

依赖已经打包在里面，直接执行：

```sh
python3 offline-deps/install-deps.py
```

它会按本机的 Python 版本和 CPU 架构挑对应的 wheel，解包到 `vendor/`，不联网、不调用 pip。输出最后一行是 `结论：就绪` 就成了。

`start.sh` 也会自动做这件事：只要检测到缺依赖，并且目录里有 `offline-deps/wheels/`，就会先离线装一遍，不需要手工执行。

### 情况 2：程序包里没有依赖

看自检输出里本机需要哪一组：

```sh
sh check.sh        # 会打印「本机组合 py39-aarch64 在包里，执行下面这条即可补齐」
```

然后在**能联网的机器**上生成对应组合：

```sh
python3 offline-deps/build-wheels.py --arch aarch64 --py 3.9
```

把生成的 `offline-deps/wheels/py39-aarch64/` 整个目录拷进目标机同名位置，再跑 `install-deps.py`。

### 情况 3：内网有 apt 源

那就不用管离线包，直接：

```sh
sudo apt install python3-openpyxl python3-docx python3-lxml python3-pil \
                 python3-qrcode python3-pypinyin
```

程序优先用系统包，`vendor/` 里的只是兜底。

### 目录结构说明

```
offline-deps/
  install-deps.py            离线安装器（只用标准库）
  build-wheels.py            在有网机器上生成 wheel 的脚本
  wheels/
    common/                  纯 Python 包，任何架构/版本通用
    py37-aarch64/            按「Python 版本-架构」分组
    py38-aarch64/
    ...
```

`lxml` 和 `Pillow` 带编译产物，必须与目标的 Python ABI 和 CPU 架构完全一致，所以按组合分开；其余全是纯 Python，只放一份。

wheel 全部基于 manylinux2014（glibc 2.17 基线），麒麟 V10 的 glibc 2.28 完全兼容，SP1/SP2/SP3 都能用。

## 三、自检

```sh
sh check.sh
```

正常输出长这样（关键看最后一行 `结论：就绪`）：

```
Python: Python 3.9.2  ->  /usr/bin/python3
平台: aarch64 / Python 3.9.2

openpyxl       必需         3.1.2
python-docx    必需         1.1.2
lxml           必需         4.9.4
Pillow         可选         9.5.0
qrcode         可选         7.4.2
pypinyin       可选         0.49.0
pypdf          可选         3.17.4
python-barcode 可选         0.15.1
lunardate      可选         0.2.2

sjz 解密后端  openssl 命令
AES 标准向量  PASS
密钥长度      32 字节
FILL 判据     db550aa6e23f0ad5a34998bcb40655c1

结论：就绪（依赖齐全）
```

- 缺**必需**（openpyxl / python-docx / lxml）→ 整个程序用不了
- 缺**可选** → 只有相关功能报缺依赖，其余正常
- `sjz 解密后端` 如果不是 `openssl 命令` 或 `cryptography`，说明系统没装 openssl，解密会退回纯 Python，**能用但很慢**（约 1MB/s）。装一下即可：`sudo apt install openssl`

## 四、启动

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

## 五、sjz 解密怎么测

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

## 六、已验证的事实

- AES 实现过 FIPS-197 标准向量（AES-256）
- 用真实样本验证过：3317760 字节密文 → 104901 字节，明文头是 `PK\x03\x04`，与 PHP 版结果一致
- 解密耗时 0.034 秒（openssl 后端）
- 140 项单元测试，其中 9 项专测 sjz、19 项专测人员信息提取

## 七、出问题怎么查

| 现象 | 原因 |
|---|---|
| 页面空白 / 一直转圈 | 服务没起来，看终端有没有报错 |
| 所有接口 403 | 地址里的 `?t=` 令牌丢了，回终端复制完整地址 |
| 提示「后台服务是旧版本」 | 浏览器缓存了旧页面，Ctrl+F5 强刷 |
| 启动时报「依赖仍然缺失」 | 离线包与本机 Python 版本/架构不匹配，`sh check.sh` 看需要哪组 |
| 解包时报 `invalid ELF header` | 装错架构了（x86_64 的包装到了 arm64 上），删掉 `vendor/` 重装正确的 |
| 工具报「缺少依赖 Pillow」 | 可选依赖没装，执行 `offline-deps/install-deps.py` 或 apt 装 |
| sjz 文件没被解 | 看日志有没有 `sjz 识别到但未解密`，有的话括号里是原因 |
| 大文件解得很慢 | 没装 openssl，`sudo apt install openssl` |
