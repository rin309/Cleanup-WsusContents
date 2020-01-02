@echo off
cls
set InstallDirectory=C:\Tools\Scripts\Wsus

echo %InstallDirectory% へ Cleanup-WsusContents.ps1 をコピーします (上書きインストール)
echo - 今後のWindows 10機能更新プログラムのリリースのたびにメンテナンスが必要であることを理解してください
echo.
echo.
echo インストールを始めてもよい場合は何かキーを押してください...
pause > nul

copy /y Assets\Settings.Current.json Assets\Settings.Current.json.old
xcopy /erchy "%~dp0*" "%InstallDirectory%\"
cd /d "%InstallDirectory%\"
del Cleanup-WsusContents.pssproj
del Install.bat
del Upgrades.bat
move "Assets\Uninstall.bat" "Uninstall.bat"
cls

echo 環境・運用に応じて Settings.Current.json を記述してください
echo.
echo - 置き換えられた更新プログラムを拒否
echo - 【BETA】クライアントから必要とされた更新プログラムに対して、指定したグループに指定した期間を経過後に承認する
echo - WSUSのクリーンアップ (削除された古い更新プログラム, 圧縮された更新プログラム, 削除された古い更新プログラム, 解放されたディスク領域)
echo - WSUS DB インデックスの再構成 (https://gallery.technet.microsoft.com/scriptcenter/6f8cde49-5c52-4abd-9820-f1d270ddea61)
echo - 空き領域が少なくなりがちな環境で、スクリプトが正常に動作するためのダミーファイル (4GB) を作成
echo ## 機能更新プログラム
echo - 拒否: Windows 10, バージョン 1903を含む機能更新プログラム
echo - 拒否: 64ビット版以外の機能更新プログラム
echo - 拒否: コンシューマー エディション
echo ## 品質更新プログラム
echo - 拒否: Windows 10, バージョン 1909 64ビット版以外の品質更新プログラム
echo - 拒否: Windows 8.1 64ビット版以外の品質更新プログラム
echo ## Office
echo - 拒否: 64ビット版向けの更新プログラム

echo.
echo.
ping localhost -n 4 > nul
notepad "Assets\Settings.Current.json"
cls
