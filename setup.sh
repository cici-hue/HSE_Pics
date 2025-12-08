#!/bin/bash
# setup.sh - 安装脚本

# 更新包列表
apt-get update

# 安装系统依赖
apt-get install -y libgl1-mesa-glx libglib2.0-0 libsm6 libxext6 libxrender-dev

# 安装Python包
pip install --no-cache-dir -r requirements.txt

# 确保PyMuPDF正确安装
pip install --force-reinstall pymupdf
