@echo off
cls
set InstallDirectory=C:\Tools\Scripts\Wsus

cls
echo コピー済みのODBC Driverとsqlcmd Utilityをインストールしています...
Installers\VC_redist.x64.exe /install /passive
msiexec /i Installers\msodbcsql_17.4.2.1_x64.msi /passive IACCEPTMSODBCSQLLICENSETERMS=YES
msiexec /i Installers\MsSqlCmdLnUtils.msi /passive IACCEPTMSSQLCMDLNUTILSLICENSETERMS=YES

cls
xcopy /erchy "%~dp0*" "%InstallDirectory%\"
cd /d "%InstallDirectory%\"
rd /s /q bin\*
rd /s /q obj\*
del Cleanup-WsusContents.pssproj
del Install.bat
del Silent-Install.bat
del Upgrades.bat
move "Assets\Uninstall.bat" "Uninstall.bat"
cls

SchTasks /Create /Xml "%InstallDirectory%\Assets\Task-Cleanup-WsusContents (Weekly).xml" /TN "Cleanup-WsusContents"
copy /y "Wsusコンテンツのクリーンアップ.lnk" "%ProgramData%\Microsoft\Windows\Start Menu\Programs\Wsusコンテンツのクリーンアップ.lnk"
cls
