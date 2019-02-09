@echo off
cls
echo このスクリプトがあるフォルダーとタスクを削除します
pause

SchTasks /Delete /TN "Cleanup-WsusContents"
del "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Wsusコンテンツのクリーンアップ.lnk"
rd /s /q "%~dp0"
