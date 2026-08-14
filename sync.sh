#!/bin/bash
# 方太看板 一键同步脚本（公司/家里电脑通用）
# 用法：
#   ./sync.sh            -> 拉取最新 + 提交本地改动 + 推送（默认提交信息）
#   ./sync.sh "改了XX数据" -> 自定义提交信息
# 注意：运行前请确保已把最新数据生成进 index.html（或手动改好文件）
set -e
cd "$(dirname "$0")"

echo "== 1/4 拉取远端最新（rebase，避免分叉）=="
git pull --rebase origin main || { echo "拉取失败，请检查网络/权限"; exit 1; }

echo "== 2/4 暂存所有改动 =="
git add -A

if git diff --cached --quiet; then
  echo "（无本地改动，仅同步完成）"
  exit 0
fi

MSG="${1:-自动同步 $(date +%Y-%m-%d_%H:%M)}"
echo "== 3/4 提交: $MSG =="
git commit -m "$MSG"

echo "== 4/4 推送到 main =="
git push origin main

echo "✅ 完成。稍等约 1 分钟，线上看板即更新："
echo "   https://mouren2580.github.io/fangtai-dashboard/"
