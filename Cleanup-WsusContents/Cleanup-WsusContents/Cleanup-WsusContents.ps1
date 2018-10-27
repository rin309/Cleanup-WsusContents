#Requires -Version 4.0
#Requires -RunAsAdministrator
#
# 20181027 WSUS から不要な更新プログラムを拒否する
#
# このスクリプトは現状ベースで作成されたものです。今後の更新プログラムに対応するには、直接WSUSかスクリプトのメンテナンスが必要になることを理解してください。
# このスクリプトを利用したことによる問題に対する責任は一切負いません。実行する前に必ず検証をしてください。

#最初に変数をメンテナンスしてください
$FeatureUpdatesFilterFileName = "C:\Tools\Scripts\Wsus\Filter-FU-対象の定義ファイル名.txt"
$QualityUpdatesFilterFileName = "C:\Tools\Scripts\Wsus\Filter-QU-対象の定義ファイル名.txt"
$DummyFilePath = "C:\Tools\Scripts\Wsus\Dummy.tmp"
$DummyFileSize = 4294967296
$WsusDBMaintenanceScriptPath = "C:\Tools\Scripts\Wsus\Scripts-WsusDBMaintenance.sql"

$SqlCmdPath = "C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\130\Tools\Binn\SQLCMD.EXE"
$SqlServerPath = "np:\\.\pipe\Microsoft##WID\tsql\query"
$Host.PrivateData.VerboseForegroundColor = "Cyan"
$Wsus = Get-WsusServer -Name localhost -PortNumber 8530


#特定の機能更新プログラムを拒否
Function Decline-FeatureUpdates($FilteredUpdates){
    If ($FilteredUpdates -ne $null){
	    $DeclineUpdatesCount = 0
	    $FilteredUpdates | ForEach-Object {
		    $_.Decline()
		    #Write-Host ("*** 拒否済み: " + $_.Title)
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
		    #Write-Host ("*** 拒否済み: " + $_.Title)
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


#スクリプトが正常に動作するためのダミーファイルを削除する
If (Test-Path $DummyFilePath){
	Remove-Item -Path $DummyFilePath -Force | Out-Null
}


#Write-Host "* 更新プログラムを拒否"
Write-Host ("** 置き換えられた更新プログラムを拒否")
$FilteredUpdates = @()
$FilteredUpdates = $Wsus.GetUpdates() | Where-Object {$_.IsDeclined -eq $False -and $_.IsSuperseded -eq $True -and $_.HasSupersededUpdates -eq $False}
Decline-QualityUpdates($FilteredUpdates)


Write-Host ("** 機能更新プログラムを拒否")
$FilteredUpdates = @()
$AllUpdates = $Wsus.GetUpdates() | Where-Object {$_.IsDeclined -eq $False -and $_.UpdateClassificationTitle -eq "Upgrades"}
$FeatureUpdatesFilters = Get-Content -Path $FeatureUpdatesFilterFileName
$AllUpdates | ForEach-Object {
	$UpdateInformation = $_
	$FeatureUpdatesFilters | ForEach-Object {
		If ($UpdateInformation.GetInstallableItems().Files.Name -like $_){
			$FilteredUpdates += $UpdateInformation
		}
	}
}
Decline-FeatureUpdates($FilteredUpdates)


Write-Host "** 品質更新プログラムを拒否"
$FilteredUpdates = @()
$AllUpdates = $Wsus.GetUpdates() | Where-Object IsDeclined -eq $False
#Officeの更新プログラムのメンテナンス
#Office 2016 64ビット版 の更新プログラムをすべて拒否
# $FilteredUpdates += $AllUpdates | Where-Object {$_.Title -like "*64 ビット版*" -and $_.Title -notlike "*32 ビット版*" -and $_.ProductTitles -eq "Office 2016"}
Get-Content -Path $QualityUpdatesFilterFileName | ForEach-Object {
    $FilteredUpdates += $AllUpdates | Where-Object LegacyName -like $_
}
Decline-QualityUpdates($FilteredUpdates)


Cleanup-Wsus
Start $SqlCmdPath ("-S", $SqlServerPath, "-i", $WsusDBMaintenanceScriptPath)


#更新プログラムの一覧
#$Wsus.GetUpdates() | Where-Object IsDeclined -eq $False | Select Title, ProductTitles, CreationDate, LegacyName | Out-GridView -Title "拒否された更新以外のすべて"
#$Wsus.GetUpdates() | Where-Object {$_.IsDeclined -eq $False -and $_.IsApproved -eq $True} | Select Title, ProductTitles, CreationDate, LegacyName | Out-GridView -Title "承認済みの更新プログラム"
#$Wsus.GetUpdates() | Where-Object {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Select Title, ProductTitles, CreationDate, LegacyName | Out-GridView -Title "未承認の更新プログラム"
#$Wsus.GetUpdates() | Where-Object {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Export-Csv 未承認の更新プログラム.csv -Encoding UTF8

#スクリプトが正常に動作するために、作業用のダミーファイルを作成する
FsUtil File CreateNew $DummyFilePath $DummyFileSize | Out-Null
