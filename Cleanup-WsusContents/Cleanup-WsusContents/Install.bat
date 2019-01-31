@echo off
cls
set InstallDirectory=C:\Tools\Scripts\Wsus
echo このスクリプトを %InstallDirectory% にコピーします
pause

xcopy /erchy "%~dp0*" "%InstallDirectory%\"
cd /d "%InstallDirectory%\"
cls

echo 既に設定ファイルがある場合、上書きされないことをおすすめします
copy "Assets\Settings.Default.json" "Settings.Current.json"
cls
explorer /n,"%InstallDirectory%\Filters\FeatureUpdates\
explorer /n,"%InstallDirectory%\Filters\QualityUpdates\
echo 環境・運用に応じて設定を記述してください
echo .
echo - FeatureUpdatesFilter.FileNames: 既定でWindows 10, バージョン 1809以外の機能更新プログラムを拒否します
echo - QualityUpdatesFilter.FileNames: 既定でWindows 7 Service Pack 1 32ビット版, Windows 8.1 64ビット版, Windows 10, バージョン 1809 64ビット版以外の品質更新プログラムを拒否します
echo - IsDeclineMsOfficeUpdates: 既定では TargetMsOfficeArchitecture で指定したOffice向け更新プログラムの拒否が無効です
echo - TargetMsOfficeArchitecture: 既定では64ビット版のOffice向け更新プログラムが拒否されます
echo - ReservedFile.Path: [空き領域が少なくなりがちな環境向け] WSUSのコンテンツのルートディレクトリを指定します
echo - ReservedFile.Size: [空き領域が少なくなりがちな環境向け] 生成するファイルサイズ (既定で4GB)
notepad "Settings.Current.json"
cls

SchTasks /Create /Xml "%InstallDirectory%\Assets\Task-Cleanup-WsusContents (Weekly).xml" /TN "Cleanup-WsusContents"
copy /y "Wsusコンテンツのクリーンアップ.lnk" "%UserProfile%\Desktop\Wsusコンテンツのクリーンアップ.lnk"
