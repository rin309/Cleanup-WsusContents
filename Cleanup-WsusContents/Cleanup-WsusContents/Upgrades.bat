@echo off
cls
set InstallDirectory=C:\Tools\Scripts\Wsus

echo %InstallDirectory% へ Cleanup-WsusContents.ps1 をコピーします (上書きインストール)
echo - 今後のWindows 10機能更新プログラムのリリースのたびにメンテナンスが必要であることを理解してください
echo - 日曜日2:00にバックグラウンドで実行するタスクを登録します
echo - ログファイルが %InstallDirectory%\Logs\ に保存されますので、実行前後で参考にしてください
echo.
echo インストールディレクトリの変更をした場合は、以下のファイルの修正が必要です
echo - Wsusコンテンツのクリーンアップ.lnk
echo - Assets\Task-Cleanup-WsusContents (Monthly).xml
echo.
echo.
echo インストールを始めてもよい場合は何かキーを押してください...
pause > nul

xcopy /erchy "%~dp0*" "%InstallDirectory%\"
cd /d "%InstallDirectory%\"
cls

@rem echo 既に設定ファイルがある場合、上書きされないことをおすすめします
@rem copy "Assets\Settings.Default.json" "Settings.Current.json"
@rem cls
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

@rem SchTasks /Create /Xml "%InstallDirectory%\Assets\Task-Cleanup-WsusContents (Weekly).xml" /TN "Cleanup-WsusContents"
@rem copy /y "Wsusコンテンツのクリーンアップ.lnk" "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Wsusコンテンツのクリーンアップ.lnk"
@rem cls

@rem echo 初回実行には時間がかかる場合があります。検証も含め、あらかじめ実行しておくことをおすすめします。
@rem ping localhost -n 4 > nul
@rem explorer /n,"%InstallDirectory%\Wsusコンテンツのクリーンアップ.lnk"