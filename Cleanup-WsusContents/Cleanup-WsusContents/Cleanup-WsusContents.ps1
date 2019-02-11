#Requires -Version 4.0
#Requires -RunAsAdministrator
$Host.PrivateData.VerboseForegroundColor = "Cyan"
Set-Location -Path (Split-Path -Parent ($MyInvocation.MyCommand.Path))
#
# 20190211 WSUS から不要な更新プログラムを拒否する
# Cleanup-WsusContents (CWS)
#
# このスクリプトは現状ベースで作成されたものです。今後の更新プログラムに対応するには、WSUSコンソールかSettings.Current.jsonかスクリプトのメンテナンスが必要になることを理解してください。
# このスクリプトを利用したことによる問題に対する責任は一切負いません。実行する前に必ず検証をしてください。

#最初に Settings.Current.json をメンテナンスしてください
$CuttentSettingsPath = "Settings.Current.json"
$DefaultSettingsPath = "Assets\Settings.Default.json"


Function Load-Settings(){
    $SqlServerName = $SqlServerPath.Replace("$SqlServerName" ,(Get-Item -Path "Registry::HKLM\SOFTWARE\Microsoft\Update Services\Server\Setup").GetValue("SqlServerName"))
	
	If (Test-Path $CuttentSettingsPath){
        $Settings = Get-Content $CuttentSettingsPath -Encoding UTF8 -Raw | ConvertFrom-Json
        $Script:FeatureUpdatesFilterFileNames = $Settings.DeclineRule.FeatureUpdatesFilter.FileNames
        $Script:QualityUpdatesFilterFileNames = $Settings.DeclineRule.QualityUpdatesFilter.FileNames
        $Script:DummyFileName = $Settings.ReservedFile.Name
        $Script:DummyFileSize = $Settings.ReservedFile.Size
        $Script:WsusDBMaintenanceScriptPath = $Settings.MaintenanceSql.ScriptPath
        $Script:SqlCmdPath = $Settings.MaintenanceSql.SqlCmdPath
        $Script:SqlServerPath = $Settings.MaintenanceSql.ServerPath
        $Script:IsDeclineMsOfficeUpdates = $Settings.DeclineRule.IsDeclineMsOfficeUpdates
        $Script:TargetMsOfficeArchitecture = $Settings.DeclineRule.TargetMsOfficeArchitecture
        $Script:WsusServer = $Settings.Wsus.Server
        $Script:WsusPort = $Settings.Wsus.Port
		$Script:WsusInstallDirectory = $Settings.Wsus.InstallDirectory
        $Script:IsLogging = $Settings.Log.IsLogging
        $Script:LogMaximumCount = $Settings.Log.MaximumCount
    }
	If ($WsusInstallDirectory -eq $null) {
		$Script:WsusInstallDirectory = (Get-Item -Path "Registry::HKLM\SOFTWARE\Microsoft\Update Services\Server\Setup").GetValue("ContentDir")
	}
    If ($SqlCmdPath -eq $null) {
		$ODBCToolsPath = (Get-Item -Path "Registry::HKLM\SOFTWARE\Microsoft\Microsoft SQL Server\150\Tools\ClientSetup").GetValue("ODBCToolsPath")
		$Script:SqlCmdPath = Join-Path $ODBCToolsPath "SQLCMD.EXE"
	}
	$Script:SqlServerPath = $SqlServerPath.Replace("$SqlServerName",$SqlServerName)

	$DefaultSettings = Get-Content $DefaultSettingsPath -Encoding UTF8 -Raw | ConvertFrom-Json
	If ($FeatureUpdatesFilterFileNames -eq $null) {$FeatureUpdatesFilterFileNames = $DefaultSettings.DeclineRule.FeatureUpdatesFilter.FileNames}
	If ($QualityUpdatesFilterFileNames -eq $null) {$QualityUpdatesFilterFileNames = $DefaultSettings.DeclineRule.QualityUpdatesFilter.FileNames}
	If ($DummyFileName -eq $null) {$DummyFileName = $DefaultSettings.ReservedFile.Name}
	If ($DummyFileSize -eq $null) {$DummyFileSize = $DefaultSettings.ReservedFile.Size}
	If ($WsusDBMaintenanceScriptPath -eq $null) {$WsusDBMaintenanceScriptPath = $DefaultSettings.MaintenanceSql.ScriptPath}
	If ($SqlCmdPath -eq $null) {$SqlCmdPath = $DefaultSettings.MaintenanceSql.SqlCmdPath}
	If ($SqlServerPath -eq $null) {$SqlServerPath = $DefaultSettings.MaintenanceSql.ServerPath}
	If ($IsDeclineMsOfficeUpdates -eq $null) {$IsDeclineMsOfficeUpdates = $DefaultSettings.DeclineRule.IsDeclineMsOfficeUpdates}
	If ($TargetMsOfficeArchitecture -eq $null) {$TargetMsOfficeArchitecture = $DefaultSettings.DeclineRule.TargetMsOfficeArchitecture}
	If ($WsusServer -eq $null) {$WsusServer = $DefaultSettings.Wsus.Server}
	If ($WsusPort -eq $null) {$WsusPort = $DefaultSettings.Wsus.Port}
	If ($IsLogging -eq $null) {$IsLogging = $DefaultSettings.Log.IsLogging}
	If ($LogMaximumCount -eq $null) {$LogMaximumCount = $DefaultSettings.Log.MaximumCount}

	$Script:DummyFilePath = Join-Path $WsusInstallDirectory $DummyFileName
}
Function Check-Settings(){
	Write-Host "FeatureUpdatesFilterFileNames: $FeatureUpdatesFilterFileNames" -Verbose
	Write-Host "QualityUpdatesFilterFileNames: $QualityUpdatesFilterFileNames" -Verbose
	Write-Host "DummyFileName: $DummyFileName" -Verbose
	Write-Host "WsusDBMaintenanceScriptPath: $WsusDBMaintenanceScriptPath" -Verbose
	Write-Host "SqlCmdPath: $SqlCmdPath" -Verbose
	Write-Host "SqlServerPath: $SqlServerPath" -Verbose
	Write-Host "IsDeclineMsOfficeUpdates: $IsDeclineMsOfficeUpdates" -Verbose
	Write-Host "TargetMsOfficeArchitecture: $TargetMsOfficeArchitecture" -Verbose
	Write-Host "WsusServer: $WsusServer" -Verbose
	Write-Host "WsusPort: $WsusPort" -Verbose
	Write-Host "WsusInstallDirectory: $WsusInstallDirectory" -Verbose
	Write-Host "IsLogging: $IsLogging" -Verbose
	Write-Host "LogMaximumCount: $LogMaximumCount" -Verbose

    Clear-Host
}
#特定の機能更新プログラムを拒否
Function Decline-FeatureUpdates($FilteredUpdates){
    If ($FilteredUpdates -ne $null){
	    $DeclineUpdatesCount = 0
	    $FilteredUpdates | ForEach-Object {
		    $_.Decline()
            If ($DeclineUpdatesCount -eq 0){
                Write-Progress -Activity "機能更新プログラムを拒否" -Status $_.Title -PercentComplete 0
            }
            Else{
                Write-Progress -Activity "機能更新プログラムを拒否" -Status $_.Title -PercentComplete ($DeclineUpdatesCount / $FilteredUpdates.Count * 100)
		    }
            $DeclineUpdatesCount++
	    }
    }
}
#特定の品質更新プログラムを拒否
Function Decline-QualityUpdates($FilteredUpdates){
    If ($FilteredUpdates -ne $null){
	    $DeclineUpdatesCount = 0
	    $FilteredUpdates | ForEach-Object {
		    $_.Decline()
            If ($DeclineUpdatesCount -eq 0){
                Write-Progress -Activity "品質更新プログラムを拒否" -Status $_.Title -CurrentOperation $_.LegacyName -PercentComplete 0
            }
            Else{
                Write-Progress -Activity "品質更新プログラムを拒否" -Status $_.Title -CurrentOperation $_.LegacyName -PercentComplete ($DeclineUpdatesCount / $FilteredUpdates.Count * 100)
		    }
            $DeclineUpdatesCount++
	    }
    }
}
#Wsusのクリーンアップ
Function Cleanup-Wsus(){
	Write-Progress -Activity "クリーンアップしています" -Status "1/4 - 削除された古い更新プログラム" -CurrentOperation $_.LegacyName -PercentComplete (0 / 4 * 100)
	$Wsus | Invoke-WsusServerCleanup -DeclineSupersededUpdates
	Write-Progress -Activity "クリーンアップしています" -Status "2/4 - 圧縮された更新プログラム" -CurrentOperation $_.LegacyName -PercentComplete (1 / 4 * 100)
	$Wsus | Invoke-WsusServerCleanup -CompressUpdates
	Write-Progress -Activity "クリーンアップしています" -Status "3/4 - 削除された古い更新プログラム" -CurrentOperation $_.LegacyName -PercentComplete (2 / 4 * 100)
	$Wsus | Invoke-WsusServerCleanup -CleanupObsoleteUpdates
	Write-Progress -Activity "クリーンアップしています" -Status "4/4 - 解放されたディスク領域" -CurrentOperation $_.LegacyName -PercentComplete (3 / 4 * 100)
	$Wsus | Invoke-WsusServerCleanup -CleanupUnneededContentFiles
}
Function Start-Logging(){
	If ($IsLogging){
        $StartTime = (Get-Date –F s).Replace(':','')
		$Script:LogDirectory = "Logs\$StartTime\"
		New-Item $LogDirectory -ItemType Directory -Force | Out-Null
        Start-Transcript "$LogDirectory\1 Transcript.log"

		Copy-Item $CuttentSettingsPath "$LogDirectory\Settings.Current.json" -Force | Out-Null
		Copy-Item $DefaultSettingsPath "$LogDirectory\Settings.Default.json" -Force | Out-Null
		$LogsDirectoryChildItems = (Get-ChildItem "Logs\" -Directory -Filter "20*")
		If ($LogsDirectoryChildItems.Length -gt $LogMaximumCount){
			ForEach ($LogsDirectoryChildItem in $LogsDirectoryChildItems){
				$LogsDirectoryChildItem | Remove-Item
				#西暦上2桁"20"から始まるディレクトリを検索
                If ((Get-ChildItem "Logs\" -Directory -Filter "20*").Length -le $LogMaximumCount){
					break
				}
			}
		}

		Get-PSDrive -Name $WsusInstallDirectory.SubString(0,1)
	}
}

Load-Settings
Start-Logging
Check-Settings

Write-Host "* WSUSへ接続" -Verbose
$Wsus = Get-WsusServer -Name $WsusServer -PortNumber $WsusPort
Clear-Host

Write-Host "空き領域が少なくなりがちな環境で、スクリプトが正常に動作するためのダミーファイルを削除する" -Verbose
If (Test-Path $DummyFilePath){
	Remove-Item -Path $DummyFilePath -Force | Out-Null
}
Clear-Host


Write-Host "* 更新プログラムを拒否" -Verbose
Write-Host "** 置き換えられた更新プログラムを拒否" -Verbose
$FilteredUpdates = @()
$FilteredUpdates = $Wsus.GetUpdates() | Where-Object {$_.IsDeclined -eq $False -and $_.IsSuperseded -eq $True -and $_.HasSupersededUpdates -eq $False}
Decline-QualityUpdates($FilteredUpdates)
If ($IsLogging){
	$FilteredUpdates | Where-Object {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Select Title, @{Name="ProductTitles";Expression={($_.ProductTitles)}}, CreationDate, LegacyName | Export-Csv -NoTypeInformation "$LogDirectory\2-1 拒否済み - 置き換えられた更新プログラム.csv" -Encoding UTF8
}


Write-Host "** 機能更新プログラムを拒否" -Verbose
$FilteredUpdates = @()
$AllUpdates = $Wsus.GetUpdates() | Where-Object {$_.IsDeclined -eq $False -and $_.UpdateClassificationTitle -eq "Upgrades"}
ForEach ($FeatureUpdatesFilterFileName in $FeatureUpdatesFilterFileNames){
	$FeatureUpdatesFilters = Get-Content -Path "Filters\FeatureUpdates\$FeatureUpdatesFilterFileName"
	$AllUpdates | ForEach-Object {
		$UpdateInformation = $_
		$FeatureUpdatesFilters | ForEach-Object {
			If ($UpdateInformation.GetInstallableItems().Files.Name -like $_){
				$FilteredUpdates += $UpdateInformation
			}
		}
	}
}
Decline-FeatureUpdates($FilteredUpdates)
If ($IsLogging){
	$FilteredUpdates | Where-Object {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Select Title, @{Name="ProductTitles";Expression={($_.ProductTitles)}}, CreationDate, LegacyName | Export-Csv -NoTypeInformation "$LogDirectory\2-2 拒否済み - 機能更新プログラム.csv" -Encoding UTF8
}

Write-Host "** 品質更新プログラムを拒否" -Verbose
$FilteredUpdates = @()
$AllUpdates = $Wsus.GetUpdates() | Where-Object IsDeclined -eq $False
ForEach ($QualityUpdatesFilterFileName in $QualityUpdatesFilterFileNames){
	Get-Content -Path "Filters\QualityUpdates\$QualityUpdatesFilterFileName" | ForEach-Object {
		$FilteredUpdates += $AllUpdates | Where-Object LegacyName -like $_
	}
}
Write-Host "*** Office向け更新プログラム" -Verbose
If ($IsDeclineMsOfficeUpdates){
	$FilteredUpdates += $AllUpdates | Where-Object {$_.Title -like "*$TargetMsOfficeArchitecture*" -and $_.ProductTitles -like "Office *"}
}
If ($IsLogging){
	$FilteredUpdates | Where-Object {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Select Title, @{Name="ProductTitles";Expression={($_.ProductTitles)}}, CreationDate, LegacyName | Export-Csv -NoTypeInformation "$LogDirectory\2-3 拒否済み - 品質更新プログラム.csv" -Encoding UTF8
}
Decline-QualityUpdates($FilteredUpdates)
Clear-Host


Write-Host "* WSUSのクリーンアップ" -Verbose
Cleanup-Wsus
if (Test-Path $SqlCmdPath){
	Start $SqlCmdPath ("-S", $SqlServerPath, "-i", $WsusDBMaintenanceScriptPath)
}
Else{
	Write-Host "sqlcmd Utility が見つかりませんでしたので、SQLデータベースの最適化をしていません"
}

Clear-Host


#更新プログラムの一覧をGridViewで確認する
#$Wsus.GetUpdates() | Where-Object IsDeclined -eq $False | Select Title, ProductTitles, CreationDate, LegacyName | Out-GridView -Title "拒否された更新以外のすべて"
#$Wsus.GetUpdates() | Where-Object {$_.IsDeclined -eq $False -and $_.IsApproved -eq $True} | Select Title, ProductTitles, CreationDate, LegacyName | Out-GridView -Title "承認済みの更新プログラム"
#$Wsus.GetUpdates() | Where-Object {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Select Title, ProductTitles, CreationDate, LegacyName | Out-GridView -Title "未承認の更新プログラム"
If ($IsLogging){
	$Wsus.GetUpdates() | Where-Object {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Select Title, @{Name="ProductTitles";Expression={($_.ProductTitles)}}, CreationDate, LegacyName | Export-Csv -NoTypeInformation "$LogDirectory\3 未承認の更新プログラム.csv" -Encoding UTF8
	Get-PSDrive -Name $WsusInstallDirectory.SubString(0,1)
}

If ($DummyFileSize -ne 0){
	Write-Host "空き領域が少なくなりがちな環境で、スクリプトが正常に動作するためのダミーファイルを作成する" -Verbose
	FsUtil File CreateNew $DummyFilePath $DummyFileSize | Out-Null
}
Else{
	#ファイルサイズが0の場合は実行しない
}

If ($IsLogging){
	Stop-Transcript
}