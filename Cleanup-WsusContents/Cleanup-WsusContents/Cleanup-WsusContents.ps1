#Requires -RunAsAdministrator
#
# 20180401 WSUS から不要な更新プログラムを拒否する
#
# このスクリプトは現状ベースで作成されたものです。今後の更新プログラムに対応するには、直接WSUSかスクリプトのメンテナンスが必要になることを理解してください。
# このスクリプトを利用したことによる問題に対する責任は一切負いません。実行する前に必ず検証をしてください。
$Host.PrivateData.VerboseForegroundColor = "Cyan"
$Wsus = Get-WsusServer -Name localhost -PortNumber 8530


Function Decline-WsusUpdates($FilteredUpdates){
	Write-Verbose "** 指定済みの LegacyName を含む更新プログラムを拒否" #-Verbose
	$DeclineUpdatesCount = 0
	$FilteredUpdates | Foreach-Object {
		$_.Decline()
		Write-Verbose ("*** 拒否済み: " + $_.Title) #-Verbose
		Write-Progress -Activity "指定済みの LegacyName を含む更新プログラムを拒否" -Status $_.Title -CurrentOperation $_.LegacyName -PercentComplete ($DeclineUpdatesCount++ / $FilteredUpdates.Count * 100)
	}
}
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



$FilteredUpdates = @()
$FilteredUpdates = $AllUpdates | Where {IsDeclined -eq $False -and $_.IsSuperseded -eq $True -and $_.HasSupersededUpdates -eq $False}
Decline-WsusUpdates($FilteredUpdates)

$FilteredUpdates = @()
$AllUpdates = $Wsus.GetUpdates() | Where IsDeclined -eq $False

Write-Verbose "* 更新プログラムを拒否" #-Verbose
Write-Verbose "** 最新から4つより古い機能更新プログラムを削除する" #-Verbose
$AllUpdates | Where UpdateClassificationTitle -eq "Upgrades" | Sort CreationDate -Descending | Foreach-Object {
    If ($WsusUpgradesProgramsCount++ -gt 4){
        $_.Decline()
        Write-Verbose ("*** 拒否済み: " + $_.Title) #-Verbose
    }
}

Write-Verbose "** 更新プログラムを拒否するためのフィルターを作成" #-Verbose
#Upgrades に含まれる更新プログラムをすべて拒否
# $FilteredUpdates += $AllUpdates | Where UpdateClassificationTitle -eq "Upgrades"
#Office 2016 に含まれ、64ビット版 を含む更新プログラムをすべて拒否
# $FilteredUpdates += $AllUpdates | Where {$_.Title -like "*64 ビット版*" -and $_.Title -notlike "*32 ビット版*" -and $_.ProductTitles -eq "Office 2016"}
Get-Content -Path "Windows 10, バージョン 1703 64ビット版.txt" | ForEach{
    $FilteredUpdates += $AllUpdates | Where LegacyName -like $_
}

Decline-WsusUpdates($FilteredUpdates)



#更新プログラムの一覧
#$Wsus.GetUpdates() | Where IsDeclined -eq $False | Select Title, ProductTitles, CreationDate, LegacyName | Out-GridView -Title "拒否された更新以外のすべて"
#$Wsus.GetUpdates() | Where {$_.IsDeclined -eq $False -and $_.IsApproved -eq $True} | Select Title, ProductTitles, CreationDate, LegacyName | Out-GridView -Title "承認済みの更新プログラム"
#$Wsus.GetUpdates() | Where {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Select Title, ProductTitles, CreationDate, LegacyName | Out-GridView -Title "未承認の更新プログラム"
#$Wsus.GetUpdates() | Where {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Export-Csv 未承認の更新プログラム.csv -Encoding UTF8
