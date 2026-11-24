# office办公助手

银河麒麟 V10 / arm64 上的批量文档与表格处理工具。独立桌面程序，通过本地网页界面操作，不嵌入 WPS/Office 插件，不改源文件。

对标易用宝、方方格子、Kutools 的常用批处理能力，但脱离 Office 宿主进程，纯文件级处理。

> 作者：**Eyzer** <eyzer@qq.com> · 主页：https://github.com/Eyzer · 许可：MIT
> 名称、版本、作者信息集中在 `version.py`，改一处全站生效（启动横幅、界面标题、关于页、deb 控制信息）。

**下载**：到 [Releases](https://github.com/Eyzer/office-assistant/releases) 拿最新 `office-assistant-linux-x.y.z.tar.gz`，解压后 `sh start.sh` 即可，不需要装任何东西（除 Python 3）。

## 特性

- **93 个功能**：通用文件 12 / 表格 56 / 文档 25，首页按「大类 → 二级分类」铺成按钮墙，点击即进入操作页
- 界面在浏览器里，后端是 Python 标准库 http.server，零 WebView 依赖，不受 glibc 版本限制
- **系统文件对话框**：能调 `zenity` / `kdialog` 就弹真正的系统选择框（拿到真实路径，不用上传副本）；调不到自动退回浏览器自带选择框
- **拖拽即用**：文件或整个文件夹拖进窗口任意位置即可加入待处理列表
- **日志抽屉**：底部可收起的抽屉 + 悬浮按钮，不再常驻占屏；收起时用角标提示新增条数
- **依赖内嵌，内网可用**：第三方库按「Python 版本 + 架构」打成离线 wheel 包，`python3 offline-deps/install-deps.py` 解包即用，不联网、不调 pip；qrcode / Pillow / pypinyin / pypdf / python-barcode / lunardate 全在内
- **sjz 加密文件自动解密**：上传时识别 sjz 加密的 docx/xlsx/PDF 并原地解密，后续功能直接可用；非加密文件零误判
- **人员信息提取**：从一堆 docx 表格里抓姓名、身份证号等字段，汇总成一张带核验标记的 Excel 台账
- 全部结果写入输出目录，源文件只读
- 单文件失败不中断批量，日志可查
- 支持方案保存与复用、只预览（跑完清产物）、功能页深链（`#tool=<id>`）

## 环境要求

主目标平台：

- Linux arm64（银河麒麟 V10 SP1 已验证适用），glibc 2.27+
- Python 3.7+

开发调试也支持 Windows / macOS（Python 3.8+），除了少数依赖外部程序的工具外功能一致。

- 必需：openpyxl、python-docx、lxml、Pillow
- 内嵌（打包时自动带）：qrcode、pypinyin、pypdf、python-barcode、lunardate
- 可选：LibreOffice（转 PDF / PPT 转换）、poppler-utils（PDF 转图片）、zenity 或 kdialog（Linux 原生文件对话框；Windows 用系统自带的 WinForms 对话框，无需额外安装）

用界面上「环境自检」按钮可以直接看目标机缺什么、该装哪个包。

## Windows 上启动

**双击 `start-windows.bat`** 即可。脚本会自己找 Python、缺依赖时自动 `pip install`，然后启动服务并打开浏览器。关掉那个黑窗口就是退出程序。

命令行方式等价：

```bat
python main.py
python main.py --port 8080     :: 指定端口
python main.py --no-browser    :: 不自动开浏览器，终端会打印带令牌的地址
python main.py --version       :: 看版本与作者
```

在 Windows 上「添加文件」/「添加文件夹」/「选择输出目录」调的是 Windows 原生对话框，拿到的是真实路径，不会产生上传副本。

少数工具在 Windows 上需要额外装外部程序：转 PDF / PPT 转 PDF 要 LibreOffice（`soffice` 在 PATH 里），PDF 转图片要 poppler 的 `pdftoppm`。其余工具纯 Python，开箱可用。

## 安装（压缩包，推荐）

从 [Releases](https://github.com/Eyzer/office-assistant/releases) 下载对应版本：

| 包 | 用途 |
|---|---|
| `office-assistant-intranet-aarch64-<版本>.tar.gz` | **内网 / 离线**，arm64（麒麟飞腾、鲲鹏）。含全部第三方依赖 |
| `office-assistant-intranet-x86_64-<版本>.tar.gz` | **内网 / 离线**，x86_64（麒麟兆芯、海光、Intel/AMD） |
| `office-assistant-linux-<版本>.tar.gz` | 只要程序，依赖自己装（有 apt 源或能联网时用） |

```bash
tar xzf office-assistant-intranet-aarch64-1.3.3.tar.gz
cd office-assistant
python3 offline-deps/install-deps.py     # 离线装依赖，输出「结论：就绪」
sh check.sh                              # 环境自检
sh start.sh                              # 启动并自动打开浏览器
sh start.sh --host 0.0.0.0 --port 8765   # 供局域网访问
```

内网包里的依赖覆盖 Python 3.7 / 3.8 / 3.9 / 3.10 / 3.11，装的时候按本机版本自动挑，不需要你选。`start.sh` 也会在发现缺依赖时自动跑这一步。

万一打包的 Python 版本跟目标机对不上，`sh check.sh` 会直接告诉你需要哪一组、怎么生成。完整说明见 [DEPLOY-Linux.md](DEPLOY-Linux.md)。

## 安装（deb）

> Release 里只提供压缩包。deb 需要在 arm64 环境自行构建（见下方「打包 deb」），构建产物为 `office-assistant_<版本>_arm64.deb`。

```bash
sudo dpkg -i office-assistant_1.3.3_arm64.deb
sudo apt-get -f install          # 缺依赖时补齐
office-assistant                # 启动，自动打开浏览器
```

从旧包名升级：`sudo dpkg -r docbatchtool` 再装新包。

## 源码运行

```bash
pip install -r requirements.txt
python3 main.py                  # 自动选端口并打开浏览器
python3 main.py --port 8080      # 指定端口
python3 main.py --no-browser     # 不自动开浏览器
python3 main.py --version        # 看版本与作者
```

启动后终端会打印带令牌的地址，例如 `http://127.0.0.1:43210/?t=xxxx`。令牌只对本机有效，防止其他程序调用接口。

## 使用

1. 首页按分类浏览功能按钮（顶部可按大类筛选、可搜索），点一个进入操作页
2. 加文件：拖进窗口 / 「添加文件」/「添加文件夹」（都优先走系统对话框）/「浏览目录」（服务器端目录树）
3. 填参数（每个功能都有说明）
4. 「开始处理」，进度条在页面内，日志在底部抽屉里（运行时自动展开）
5. 跑完点「打开输出目录」直接跳到产物所在文件夹

拖进来的文件会先暂存到程序目录下的 `uploads/`（浏览器不给真实路径，只能由本机服务落盘一次），启动时自动清理 3 天前的暂存。用系统对话框或「浏览目录」加进来的文件是**真实路径，不产生副本**。

**大文件或大目录优先用「添加文件夹」**：拖拽必须把内容完整复制一份到 `uploads/` 再处理，几百 MB 的目录会明显慢；系统对话框只传路径，不复制。

上传是 4 路并发 + 单个连接复用，速度按字节显示。实测 1000 个几百字节的小文件约 3 秒，40 MB 的 20 个文件约 0.3 秒。

**系统对话框是立即弹出的**：Windows 上原本每次点击都要新起一个 PowerShell 并加载 WinForms（约 1.3 秒，看着像点了没反应），现在改成常驻一个对话框进程，服务启动时就预热好，点击时只发一条指令。实测「点击 → 对话框出现」从 1310~1560 ms 降到 **220~250 ms**（每次启动后的第一次约 600 ms）。该进程约占用 100 MB 内存，关掉程序时随之退出。

## 两个专项能力

### sjz 加密文件自动解密

上传时自动识别并解密，不需要额外操作。识别依据三层（命中任一即认定）：

1. 尾部有 `\x02.SJZ` 标记（记录原始长度）
2. 密文中出现连续 ≥4 个 FILL 块（`db550aa6e23f0ad5a34998bcb40655c1`，即 AES 零块的密文）
3. 文件头不是已知格式时，试解一次，解出来是 `PK` / `%PDF` / JPEG 等已知签名才认

解密走 AES-256-ECB，优先 `cryptography` → `pycryptodome` → **`openssl` 命令** → 内置纯 Python。前三者不可用时纯 Python 也能正确解，但慢约 2000 倍（3.3 MB 约 24 秒），麒麟上务必确认 `openssl` 存在（`check.sh` 会检查）。

普通 docx/xlsx/csv/pdf **不会被误判**，文件一个字节都不动。自检接口 `/api/sjz` 可查当前用的是哪个后端。

### 人员信息提取

输入一批**表格填写人员信息**的 docx，输出一张 Excel 台账（一人一行）。

认四种排版：横向 `[姓名][张三][性别][男]`、单元格内冒号 `[姓名：张三]`、上下两行（上行标签、下行值）、单列竖排（一格标签一格值往下走），以及正文段落里的 `姓名：张三　性别：男`。标签和值中间夹空列也能对上。

一份文档里有多人（表头下多行）会自动拆成多条。身份证号除了按标签取，还会全文正则兜底（18 位/15 位），校验位算错也照样抓出来，只在「身份证核验」列标红等人工确认。

参数：① 要提取的字段（逗号分隔，留空=内置 25 项全抓）；② 额外字段（词典里没有的标签，按字面匹配）；输出格式 xlsx/csv；是否每份另存明细。

**限制**：只认 docx（`.doc` 先另存为 docx）；扫描件/图片表格抓不到（没有文字层）；值是左、标签在右的反向布局不支持。

## 功能清单

### 通用文件（12）

| 分类 | 工具 |
|---|---|
| 文件整理 | 批量复制/移动、按规则分拣、摊平目录、删除空文件夹、**批量创建文件夹**、**列出文件清单**、**查找重复文件** |
| 批量改名 | 批量改名（支持规则与序号） |
| 压缩与编码 | 批量压缩/解压、**批量转文件编码** |
| 图片处理 | 图片格式转换 |
| 格式转换 | **Office 批量转 PDF**（PPT/Word/Excel） |

### 表格（56）

| 分类 | 工具 |
|---|---|
| 拆分与合并 | 按列/按行数/按工作表/按区域拆分、**按关键字拆表**、合并到总表、多模式合并、按文件夹合并 |
| 数据整理 | 数据清洗、查重、高级排序、随机抽样、二维转一维、**行列转置/一维转二维/单行单列转多行多列/交换区域**、**层级编号/拆解/合并/提取/汇总/比较**、**行列批量插入删除**、**分类汇总** |
| 单元格与区域 | 取消合并填充、合并同值、多列合一、一列拆多列、隔行插入/换色、**特殊求和（合计/小计/合并格）**、**超链接批处理** |
| 文本处理 | **文本增删改（前后缀/间隔插入/删空格特殊字符/大小写）**、**信息提取（手机号/邮箱/网址/身份证/数字/汉字/去重）**、**文字数值互转（拼音/英文金额/大写金额/算式求值/进制/IP）** |
| 数值与计算 | 数字格式转换、**取整与分级（四舍五入/四舍六入五留单/分级）**、**公式填充与引用转换**、公式转数值、**日期时间处理（成分/时间戳/标准化/农历/工作日）**、**身份证手机号脱敏校验** |
| 格式与样式 | 标准化格式、跨文件套格式、格式互转 |
| 工作表与工作簿 | 批量管理、批量改名、删空表、导出清单、**排序/隐藏/标签色/删除/保护/冻结/按模板新建/生成目录**、按规则改文件名 |
| 查找与对比 | 跨表跨簿查找、工作簿对比、**两列相同与不同项 / 两表逐格对比** |
| 图片与条码 | 提取/导入图片、二维码生成、**条形码生成（Code128/Code39/EAN13）**、**图片批处理（缩放/水印/打码/拼接/转PDF/EXIF/主色调）** |
| 批注 | **批注提取/删除/改内容/改作者/加前缀** |
| 导入与导出 | 转 PDF、**JSON↔表格**、**文本文件互转（拆分/合并/转表格）**、**导出 VCF 通讯录 / DBF**、**PDF 拆分/合并/转图片** |

### 文档（25）

| 分类 | 工具 |
|---|---|
| 内容处理 | 清理（空段落/批注/超链接）、查找替换、修订接受拒绝、元数据清理、数据套打 |
| 格式排版 | 页边距纸张、正文样式统一、页眉页脚页码、图片布局、表格转文本/翻转、**加水印** |
| 提取与转换 | 提取正文、提取图片、提取表格、转 PDF、压缩瘦身、**字数统计台账**、**人员信息提取（汇总成 Excel 台账）** |
| 文档结构 | 多文档合并、按标题拆分、按页提取/删除、只读保护、文档对比、删图片、**批量插入目录** |

## 打包 deb

在 arm64 环境（或 qemu arm64 rootfs）执行：

```bash
bash packaging/build-deb.sh
PYVER=3.7 bash packaging/build-deb.sh     # 指定目标机 Python 版本
```

包名 `office-assistant_<版本>_arm64.deb`，版本号与 Maintainer 从 `version.py` 读。脚本会调用 `packaging/fetch-vendor.sh`，优先用 `offline-deps/wheels` 里现成的 wheel（没有时才联网下载），装到目标机后无需联网。

```bash
PYVER=3.7 ARCH=aarch64 bash packaging/fetch-vendor.sh
```

## 离线依赖包

目录约定：

```
offline-deps/
  install-deps.py            离线安装器，只用标准库，按本机版本挑 wheel 解到 vendor/
  build-wheels.py            在有网机器上生成 wheel
  wheels/
    common/                  纯 Python 包，任何架构与版本通用
    py37-aarch64/            二进制包按「Python 版本-架构」分组
    py38-aarch64/  py39-*/  py310-*/  py311-*/
    py39-x86_64/  ...
```

二进制包（`lxml`、`Pillow`）必须与目标 ABI 和 CPU 架构严格一致，所以按组合分开；纯 Python 只存一份。全部 wheel 基于 manylinux2014（glibc 2.17 基线），麒麟 V10 的 glibc 2.28 完全兼容。

重新生成某一组（有网机器）：

```bash
python3 offline-deps/build-wheels.py --arch aarch64 --py 3.9
python3 offline-deps/build-wheels.py --arch x86_64 --py 3.7 3.9 3.11
python3 offline-deps/build-wheels.py --clean          # 全部重下
```

换 Python 版本或换架构后记得同步 `build-wheels.py` 里的版本清单：3.7 用 python-docx 1.0.1（1.1.x 依赖的 typing-extensions 已不支持 3.7），3.12 用 lxml 5.x + Pillow 10.x。

## 目录

```
main.py         入口（--version 看版本与作者）
version.py      名称 / 版本 / 作者 / 关于页信息，改这里全站生效
start-windows.bat  Windows 双击启动（自动找 Python、缺依赖自动装）
server.py       本地服务与接口（/api/upload 拖拽落盘、/api/pick 系统对话框）
web/            前端页面（首页功能墙 + 工作区 + 日志抽屉）
utils/          底层封装（batch.py 批处理公共骨架，deps.py 依赖回退）
service/        任务调度与分类索引
tools/          各功能，按 common/excel/word 分组
vendor/         第三方依赖解包后的位置（离线安装器写入）
offline-deps/   离线依赖：install-deps.py、build-wheels.py、wheels/
uploads/        拖拽暂存（自动清理）
profiles/       保存的方案
samples/        样例与生成脚本
tests/          单元、接口与冒烟测试
packaging/      deb 构建、依赖拉取、桌面图标与启动器
```

## 主要接口

| 接口 | 说明 |
|---|---|
| `GET /api/tools` | 功能清单（含分组与二级分类） |
| `POST /api/run` | 启动任务，返回 task_id |
| `GET /api/task?id=&since=` | 轮询进度与增量日志 |
| `POST /api/upload?sid=&name=` | 拖拽文件落盘（原始字节流） |
| `POST /api/open` | 用系统文件管理器打开目录（传文件路径则打开其所在目录） |
| `GET /api/pick?mode=files\|dir` | 弹系统文件对话框，返回真实路径 |
| `GET /api/pick?mode=warm` | 预热对话框进程（页面加载时调用，不弹窗） |
| `GET /api/fs?path=` | 服务器端目录树（对话框不可用时的兜底） |
| `GET /api/deps` | 依赖自检 |
| `GET /api/sjz` | sjz 解密模块自检（后端、密钥指纹、FILL 判据） |
| `GET /api/ping` / `GET /api/about` | 版本握手 / 关于信息 |

页面加载时会用 `/api/ping` 和后端对一次接口版本；对不上会横条提示「后台服务版本过旧，请重启程序」，避免出现「按钮点了没反应」的闷包。

## 开发

新增工具：在 `tools/<组>/` 下建文件，定义 `META`（id/name/desc/fields/cat）与 `run_task(inputs, output_dir, params, ctx)`，调度层自动加载，无需改主程序。分类写到 `service/catalog.py`。

简单工具直接用 `utils/batch.py` 的 `excel_edit()` 或 `process()`，行循环、进度、失败隔离都封装好了。

需要第三方库时：`from utils import deps` 然后 `deps.load("qrcode")`，缺失会抛带中文提示的 `MissingDep`；加新依赖要同时写进 `requirements.txt` 和 `offline-deps/build-wheels.py` 的清单，再跑一次 `build-wheels.py`。

测试：

```bash
python -m unittest discover -s tests -p "test_*.py"   # 140 个：工具 + 接口 + 版本信息
python -m unittest tests.test_doc_info_extract -q      # 只跑人员信息提取（19 个）
python samples/make_samples.py                        # 重新生成样例
```

改名字 / 版本 / 作者：只动 `version.py`。前端占位符 `%%APP_NAME%%` 等由 `server.py` 在返回 index.html 时替换，deb 的 Version/Maintainer 由 `build-deb.sh` 从同一文件读。

## 关于

- 名称：office办公助手
- 作者：Eyzer
- 邮箱：eyzer@qq.com
- 主页：https://github.com/Eyzer
- 许可：MIT
- 当前版本：1.3.3（`python3 main.py --version` 可查）

界面右上角「关于」里能看到同一份信息与第三方组件清单。本程序只在 127.0.0.1 上开服务，不联网、不上传文件。

## 限制

- `.et` 不直接支持：LibreOffice 预转有损，建议先另存为 xlsx
- 打开密码、单元格加密：需 Office 环境，不做
- 转 PDF 依赖 LibreOffice；PDF 转图片依赖 poppler-utils；农历依赖 lunardate
- 依赖 Office 宿主的功能不做：透视表刷新、PQ 刷新、多簿群控、脚本管理器、AI 翻译、OCR、发票识别、录屏
- 按页操作基于手动分页符，无分页符的文档按整篇处理
- openpyxl 读写会丢图片/图表/条件格式的部分属性；图片导出走 zip 底层解包，不受影响
