#!/bin/bash
cd "$(dirname "$0")"

echo "========================================"
echo "   通途订单整理工具"
echo "========================================"
echo ""

# 停掉旧进程
lsof -ti:5099 | xargs kill -9 2>/dev/null
sleep 0.5

# 检查并安装依赖
echo "正在检查依赖..."
pip3 install -q flask openpyxl 2>&1 | tail -1

echo "正在启动服务..."
echo ""

# 打开浏览器
sleep 1
open "http://127.0.0.1:5099" 2>/dev/null &

# 启动 Flask
python3 app.py
