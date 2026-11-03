#!/bin/sh
# 打包扩展：生成 dist/ChannelsExport.zip（压缩包内用英文名，避免 Windows 解压乱码），解压后是一个包含 manifest.json 的文件夹
set -e
cd "$(dirname "$0")"
rm -rf dist && mkdir -p "dist/ChannelsExport"
cp -R extension/ "dist/ChannelsExport/"
find dist -name .DS_Store -delete
(cd dist && zip -qr ChannelsExport.zip ChannelsExport && rm -rf ChannelsExport)
ls -la dist
