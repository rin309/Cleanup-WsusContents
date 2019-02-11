@echo off
cls
set InstallDirectory=C:\Tools\Scripts\Wsus

echo %InstallDirectory% へ Cleanup-WsusContents.ps1 をコピーします (上書きインストール)
echo - 今後のWindows 10機能更新プログラムのリリースのたびにメンテナンスが必要であることを理解してください
echo.
echo.
echo インストールを始めてもよい場合は何かキーを押してください...
pause > nul

xcopy /erchy "%~dp0*" "%InstallDirectory%\"
cd /d "%InstallDirectory%\"
del Cleanup-WsusContents.pssproj
del Install.bat
del Upgrades.bat
move "Assets\Uninstall.bat" "Uninstall.bat"
cls

echo 環境・運用に応じて Settings.Current.json を記述してください
echo.
echo - FeatureUpdatesFilter.FileNames: 既定でリテール版のWindows 10の機能更新プログラムとWindows 10, バージョン 1809以外の機能更新プログラムを拒否します
echo -- %InstallDirectory%\Filters\FeatureUpdates\ のファイル名を追加することにより、対象を増やすことができます
echo - QualityUpdatesFilter.FileNames: 既定でWindows 7 Service Pack 1 32ビット版, Windows 8.1 64ビット版, Windows 10, バージョン 1809 64ビット版以外の品質更新プログラムを拒否します
echo -- %InstallDirectory%\Filters\QualityUpdates\ のファイル名を追加することにより、対象を増やすことができます
echo - IsDeclineMsOfficeUpdates: 既定では TargetMsOfficeArchitecture で指定したOffice向け更新プログラムの拒否が有効です
echo - TargetMsOfficeArchitecture: 既定では64ビット版のOffice向け更新プログラムが拒否されます
echo - ReservedFile: 設定は暫定処置です。同パーティション内にほかのシステムが同居する場合はFSRMによるクォーターなどを検討してください。
echo.
echo.
ping localhost -n 4 > nul
notepad "Settings.Current.json"
cls
