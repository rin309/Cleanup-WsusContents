@echo off
cls
set InstallDirectory=C:\Tools\Scripts\Wsus

echo SQLのメンテナンスなどが動作しませんので、このツールはWSUSがインストールされているサーバーにて実行されることをおすすめします
echo また、SQLのメンテナンスにはODBC Driverとsqlcmd Utilityが必要になります
echo.
echo 事前にインストールを済ませてください
echo - https://docs.microsoft.com/ja-jp/sql/tools/sqlcmd-utility
echo △sqlcmd 15をインストールするにはODBC 13.1とODBC 17のインストールが必要なようです
echo.
echo.
pause
cls
echo %InstallDirectory% へ Cleanup-WsusContents.ps1 をコピーします
echo - 今後のWindows 10機能更新プログラムのリリースのたびにメンテナンスが必要であることを理解してください
echo - 日曜日2:00にバックグラウンドで実行するタスクを登録します
echo - Cleanup-WsusContents.ps1 のログファイルは %InstallDirectory%\Logs\ に保存されますので、定期的に確認されることをおすすめします
echo.
echo インストールディレクトリの変更をした場合は、以下のファイルの修正が必要です
echo - Wsusコンテンツのクリーンアップ.lnk
echo - Assets\Task-Cleanup-WsusContents (Monthly).xml
echo.
echo (下記ファイルは通常は使用しませんので適宜修正してください)
echo - WSUS DB インデックスの再構成.lnk
echo - WSUS メモリーサイズ調整 (メモリー実装容量が8GB～環境向け).lnk
echo - Assets\Task-Cleanup-WsusContents (Weekly).xml
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

echo 既に設定ファイルがある場合、上書きされないことをおすすめします
copy "Assets\Settings.Current.json" "Settings.Current.json"
cls
@rem explorer /n,"%InstallDirectory%\Filters\FeatureUpdates\"
@rem explorer /n,"%InstallDirectory%\Filters\QualityUpdates\"

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

SchTasks /Create /Xml "%InstallDirectory%\Assets\Task-Cleanup-WsusContents (Weekly).xml" /TN "Cleanup-WsusContents"
copy /y "Wsusコンテンツのクリーンアップ.lnk" "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Wsusコンテンツのクリーンアップ.lnk"
cls

echo 初回実行には時間がかかる場合があります。検証も含め、あらかじめ実行しておくことをおすすめします。
ping localhost -n 4 > nul
explorer /n,"%InstallDirectory%\Wsusコンテンツのクリーンアップ.lnk"
