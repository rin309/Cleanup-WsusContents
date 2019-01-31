#Requires -Version 4.0
#Requires -RunAsAdministrator
$Host.PrivateData.VerboseForegroundColor = "Cyan"
#
# 20190201 WSUS から不要な更新プログラムを拒否する
#
# このスクリプトは現状ベースで作成されたものです。今後の更新プログラムに対応するには、直接WSUSかスクリプトのメンテナンスが必要になることを理解してください。
# このスクリプトを利用したことによる問題に対する責任は一切負いません。実行する前に必ず検証をしてください。

#最初に Settings.Current.json をメンテナンスしてください
$CuttentSettingsPath = "Settings.Current.json"
$DefaultSettingsPath = "Assets\Settings.Default.json"
If (!(Test-Path $CuttentSettingsPath)){
	Copy-Item $DefaultSettingsPath $CuttentSettingsPath
}

Function Load-Settings(){
	$Settings = Get-Content $CuttentSettingsPath -Encoding UTF8 -Raw | ConvertFrom-Json
	$DefaultSettings = Get-Content $DefaultSettingsPath -Encoding UTF8 -Raw | ConvertFrom-Json
	
	$FeatureUpdatesFilterFileNames = $Settings.DeclineRule.FeatureUpdatesFilter.FileNames
	If ($FeatureUpdatesFilterFileNames -eq $null) {$FeatureUpdatesFilterFileNames = $DefaultSettings.DeclineRule.FeatureUpdatesFilter.FileNames}
	$QualityUpdatesFilterFileNames = $Settings.DeclineRule.QualityUpdatesFilter.FileNames
	If ($QualityUpdatesFilterFileNames -eq $null) {$QualityUpdatesFilterFileNames = $DefaultSettings.DeclineRule.QualityUpdatesFilter.FileNames}
	$DummyFilePath = $Settings.ReservedFile.Path
	If ($DummyFilePath -eq $null) {$DummyFilePath = $DefaultSettings.ReservedFile.Path}
	$DummyFileSize = $Settings.ReservedFile.Size
	If ($DummyFileSize -eq $null) {$DummyFileSize = $DefaultSettings.ReservedFile.Size}
	$WsusDBMaintenanceScriptPath = $Settings.MaintenanceSql.ScriptPath
	If ($WsusDBMaintenanceScriptPath -eq $null) {$WsusDBMaintenanceScriptPath = $DefaultSettings.MaintenanceSql.ScriptPath}
	$SqlCmdPath = $Settings.MaintenanceSql.SqlCmdPath
	If ($SqlCmdPath -eq $null) {$SqlCmdPath = $DefaultSettings.MaintenanceSql.SqlCmdPath}
	$SqlServerPath = $Settings.MaintenanceSql.ServerPath
	If ($SqlServerPath -eq $null) {$SqlServerPath = $DefaultSettings.MaintenanceSql.ServerPath}
	$IsDeclineMsOfficeUpdates = $Settings.DeclineRule.IsDeclineMsOfficeUpdates
	If ($IsDeclineMsOfficeUpdates -eq $null) {$IsDeclineMsOfficeUpdates = $DefaultSettings.DeclineRule.IsDeclineMsOfficeUpdates}
	$TargetMsOfficeArchitecture = $Settings.DeclineRule.TargetMsOfficeArchitecture
	If ($TargetMsOfficeArchitecture -eq $null) {$TargetMsOfficeArchitecture = $DefaultSettings.DeclineRule.TargetMsOfficeArchitecture}
	$WsusServer = $Settings.Wsus.Server
	If ($WsusServer -eq $null) {$WsusServer = $DefaultSettings.Wsus.Server}
	$WsusPort = $Settings.Wsus.Port
	If ($WsusPort -eq $null) {$WsusPort = $DefaultSettings.Wsus.Port}
	$IsLogging = $Settings.Log.IsLogging
	If ($IsLogging -eq $null) {$IsLogging = $DefaultSettings.Log.IsLogging}
	$LogMaximumCount = $Settings.Log.MaximumCount
	If ($LogMaximumCount -eq $null) {$LogMaximumCount = $DefaultSettings.Log.MaximumCount}

	Write-Host $FeatureUpdatesFilterFileNames -Verbose
	Write-Host $QualityUpdatesFilterFileNames -Verbose
	Write-Host $DummyFilePath -Verbose
	Write-Host $WsusDBMaintenanceScriptPath -Verbose
	Write-Host $SqlCmdPath -Verbose
	Write-Host $SqlServerPath -Verbose
	Write-Host $IsDeclineMsOfficeUpdates -Verbose
	Write-Host $TargetMsOfficeArchitecture -Verbose
	Write-Host $WsusServer -Verbose
	Write-Host $WsusPort -Verbose
	Write-Host $IsLogging -Verbose
	Write-Host $LogMaximumCount -Verbose
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
		Start-Transcript "Logs\$StartTime\1 Transcript.log"
		Copy-Item $CuttentSettingsPath "Logs\$StartTime\Settings.Current.json" -Force | Out-Null
		Copy-Item $DefaultSettingsPath "Logs\$StartTime\Settings.Default.json" -Force | Out-Null
		$StartTime = (Get-Date –F s).Replace(':','')
		$LogDirectory = "Logs\$StartTime\"
		New-Item $LogDirectory -ItemType Directory -Force | Out-Null
		$LogsDirectoryChildItems = (Get-ChildItem "Logs\" -Directory -Filter "20*")
		If ($LogsDirectoryChildItems.Length -gt $LogMaximumCount){
			ForEach ($LogsDirectoryChildItem in $LogsDirectoryChildItems){
				$LogsDirectoryChildItem | Remove-Item
				If ((Get-ChildItem "Logs\" -Directory -Filter "20*").Length -le $LogMaximumCount){
					break
				}
			}
		}
	}
}

Start-Logging
Load-Settings

Write-Host "* WSUSへ接続" -Verbose
$Wsus = Get-WsusServer -Name $Settings.Wsus.Server -PortNumber $Settings.Wsus.Port

Write-Host "空き領域が少なくなりがちな環境で、スクリプトが正常に動作するためのダミーファイルを削除する" -Verbose
If (Test-Path $DummyFilePath){
	Remove-Item -Path $DummyFilePath -Force | Out-Null
}


Write-Host "* 更新プログラムを拒否" -Verbose
Write-Host "** 置き換えられた更新プログラムを拒否" -Verbose
$FilteredUpdates = @()
$FilteredUpdates = $Wsus.GetUpdates() | Where-Object {$_.IsDeclined -eq $False -and $_.IsSuperseded -eq $True -and $_.HasSupersededUpdates -eq $False}
Decline-QualityUpdates($FilteredUpdates)
If ($IsLogging){
	$FilteredUpdates | Where-Object {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Select Title, @{Name="ProductTitles";Expression={($_.ProductTitles)}}, CreationDate, LegacyName | Export-Csv -NoTypeInformation "Logs\$StartTime\2-1 拒否済み - 置き換えられた更新プログラム.csv" -Encoding UTF8
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
	$FilteredUpdates | Where-Object {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Select Title, @{Name="ProductTitles";Expression={($_.ProductTitles)}}, CreationDate, LegacyName | Export-Csv -NoTypeInformation "Logs\$StartTime\2-2 拒否済み - 機能更新プログラム.csv" -Encoding UTF8
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
	$FilteredUpdates | Where-Object {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Select Title, @{Name="ProductTitles";Expression={($_.ProductTitles)}}, CreationDate, LegacyName | Export-Csv -NoTypeInformation "Logs\$StartTime\2-3 拒否済み - 品質更新プログラム.csv" -Encoding UTF8
}
Decline-QualityUpdates($FilteredUpdates)


Write-Host "* WSUSのクリーンアップ" -Verbose
Cleanup-Wsus
Start $SqlCmdPath ("-S", $SqlServerPath, "-i", $WsusDBMaintenanceScriptPath)


#更新プログラムの一覧
#$Wsus.GetUpdates() | Where-Object IsDeclined -eq $False | Select Title, ProductTitles, CreationDate, LegacyName | Out-GridView -Title "拒否された更新以外のすべて"
#$Wsus.GetUpdates() | Where-Object {$_.IsDeclined -eq $False -and $_.IsApproved -eq $True} | Select Title, ProductTitles, CreationDate, LegacyName | Out-GridView -Title "承認済みの更新プログラム"
#$Wsus.GetUpdates() | Where-Object {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Select Title, ProductTitles, CreationDate, LegacyName | Out-GridView -Title "未承認の更新プログラム"
If ($IsLogging){
	$Wsus.GetUpdates() | Where-Object {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Select Title, @{Name="ProductTitles";Expression={($_.ProductTitles)}}, CreationDate, LegacyName | Export-Csv -NoTypeInformation "Logs\$StartTime\2-4 未承認の更新プログラム.csv" -Encoding UTF8
}

Write-Host "空き領域が少なくなりがちな環境で、スクリプトが正常に動作するためのダミーファイルを作成する" -Verbose
FsUtil File CreateNew $DummyFilePath $DummyFileSize | Out-Null

Stop-Transcript