#!/bin/bash
# 论文格式助手一键更新网站脚本
# 用法: bash deploy.sh "更新说明"
cd "$(dirname "$0")"
git add -A
git commit -m "${1:-网站更新}"
git push origin master
echo "✓ 已推送,网站 1-3 分钟内自动更新: https://ala-yxr.github.io/thesis-formatter/"
