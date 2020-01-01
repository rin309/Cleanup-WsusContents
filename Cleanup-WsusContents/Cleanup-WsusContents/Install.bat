@echo off
cls
set InstallDirectory=C:\Tools\Scripts\Wsus

cls
echo コピー済みのODBC Driverとsqlcmd Utilityをインストールしています...
Installers\VC_redist.x64.exe /install /passive
msiexec /i Installers\msodbcsql_17.4.2.1_x64.msi /passive IACCEPTMSODBCSQLLICENSETERMS=YES
msiexec /i Installers\MsSqlCmdLnUtils.msi /passive IACCEPTMSSQLCMDLNUTILSLICENSETERMS=YES

cls
echo (インストールの有無に関係なく表示されています)
echo SQLのメンテナンスなどを動作させるために、WSUSがインストールされているサーバーにて実行されることをおすすめします
echo.
echo また、SQLのメンテナンスにはODBC Driverとsqlcmd Utilityが必要になります
echo 事前にインストールを済ませてください
echo - https://docs.microsoft.com/ja-jp/sql/tools/sqlcmd-utility
echo - https://www.microsoft.com/ja-JP/download/details.aspx?id=56567
echo - https://aka.ms/vs/15/release/vc_redist.x64.exe
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
echo (その他ショートカットは直接プログラムで参照しませんので、適宜ご活用ください)
echo.
echo.
echo インストールを始めてもよい場合は何かキーを押してください...
pause > nul
cls

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

echo 標準で以下の設定がされていますので、環境・運用に応じて Settings.Current.json を記述してください
echo.
echo * 機能更新プログラム
echo - 拒否: Windows 10, バージョン 1903を含む機能更新プログラム
echo - 拒否: 64ビット版以外の機能更新プログラム
echo - 拒否: コンシューマー エディション
echo * 品質更新プログラム
echo - 拒否: Windows 10, バージョン 1909 64ビット版以外の品質更新プログラム
echo - 拒否: Windows 8.1 64ビット版以外の品質更新プログラム
echo * Office
echo - 拒否: 64ビット版向けの更新プログラム
echo * 【BETA】クライアントから必要とされた更新プログラムに対して、自動承認するグループ
echo - 品質更新プログラム: 更新プログラムリリースから15日経過後、"すべてのコンピューター"に承認されます
echo - 品質更新プログラム: 更新プログラムリリースから5日経過後、"先行適用-品質更新プログラム"に承認されます
echo - 機能更新プログラム: 更新プログラムリリースから60日経過後、"先行適用-機能更新プログラム"に承認されます
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
