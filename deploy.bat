@echo off
rem 论文格式助手一键更新网站(Windows 双击使用)
cd /d %~dp0
set /p msg=请输入更新说明(直接回车跳过):
git add -A
git commit -m "%msg%"
git push origin master
echo.
echo 已推送,网站 1-3 分钟内自动更新: https://ala-yxr.github.io/thesis-formatter/
pause
