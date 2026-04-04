#!/bin/bash

# 切换到脚本所在目录
cd "$(dirname "$0")"

# 检查嵌入式Python是否存在
if [ ! -f "runtime/python3" ]; then
    echo "Error: Embedded Python not found in runtime directory!"
    exit 1
fi

echo "Starting Gradio app with embedded Python..."

# 运行应用
runtime/python3 gradio_app.py