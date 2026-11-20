#!/usr/bin/env bash
# office办公助手 启动器：进入安装目录并运行主程序
set -e
APP_DIR="${OFFICE_ASSISTANT_HOME:-/opt/office-assistant}"
cd "$APP_DIR"
exec python3 main.py "$@"
