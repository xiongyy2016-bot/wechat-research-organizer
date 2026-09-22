#!/bin/bash
set -euo pipefail
cd "$(dirname "$0")"
export MACOSX_DEPLOYMENT_TARGET=13.0
python3 -m pip install -r requirements-build-macos.txt
python3 -m PyInstaller \
  --noconfirm \
  --clean \
  --windowed \
  --name "微信研究资料整理器" \
  --osx-bundle-identifier "org.local.wechat-research-collector" \
  --add-data "categories.json:." \
  --collect-data certifi \
  --collect-all Vision \
  --collect-all Foundation \
  --collect-all Quartz \
  app.py
codesign --force --deep --sign - "dist/微信研究资料整理器.app"
