#Requires -Version 4.0
#Requires -RunAsAdministrator
#
# 20180515 WSUS から不要な更新プログラムを拒否する
#
# このスクリプトは現状ベースで作成されたものです。今後の更新プログラムに対応するには、直接WSUSかスクリプトのメンテナンスが必要になることを理解してください。
# このスクリプトを利用したことによる問題に対する責任は一切負いません。実行する前に必ず検証をしてください。
$FilterFileName = "C:\Tools\Scripts\Wsus\Filter-対象の定義ファイル名.txt"
$DeclineUpgradesProgramsLeft = 4

$Host.PrivateData.VerboseForegroundColor = "Cyan"
$Wsus = Get-WsusServer -Name localhost -PortNumber 8530


Function Decline-WsusUpdates($FilteredUpdates){
    If ($FilteredUpdates -ne $null){
	    Write-Host "** 指定済みの LegacyName を含む更新プログラムを拒否"
	    $DeclineUpdatesCount = 0
	    $FilteredUpdates | Foreach-Object {
		    $_.Decline()
		    #Write-Host ("*** 拒否済み: " + $_.Title)
            If ($DeclineUpdatesCount -eq 0){
                Write-Progress -Activity "指定済みの LegacyName を含む更新プログラムを拒否" -Status $_.Title -CurrentOperation $_.LegacyName -PercentComplete 0
            }
            Else{
                Write-Progress -Activity "指定済みの LegacyName を含む更新プログラムを拒否" -Status $_.Title -CurrentOperation $_.LegacyName -PercentComplete ($DeclineUpdatesCount / $FilteredUpdates.Count * 100)
		    }
            $DeclineUpdatesCount++
	    }
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
$FilteredUpdates = $Wsus.GetUpdates() | Where {$_.IsDeclined -eq $False -and $_.IsSuperseded -eq $True -and $_.HasSupersededUpdates -eq $False}
Decline-WsusUpdates($FilteredUpdates)

$FilteredUpdates = @()
$AllUpdates = $Wsus.GetUpdates() | Where IsDeclined -eq $False

Write-Host "* 更新プログラムを拒否"
Write-Host ("** 最新から" + $DeclineUpgradesProgramsLeft + "つより古い機能更新プログラムを削除する")
$WsusUpgradesProgramsCount = 1
$AllUpdates | Where UpdateClassificationTitle -eq "Upgrades" | Sort CreationDate -Descending | Foreach-Object {
    If ($WsusUpgradesProgramsCount++ -gt $DeclineUpgradesProgramsLeft){
        $_.Decline()
        #Write-Host ("*** 拒否済み: " + $_.Title)
    }
}

Write-Host "** 更新プログラムを拒否するためのフィルターを作成"
#Upgrades に含まれる更新プログラムをすべて拒否
# $FilteredUpdates += $AllUpdates | Where UpdateClassificationTitle -eq "Upgrades"
#Office 2016 に含まれ、64ビット版 を含む更新プログラムをすべて拒否
# $FilteredUpdates += $AllUpdates | Where {$_.Title -like "*64 ビット版*" -and $_.Title -notlike "*32 ビット版*" -and $_.ProductTitles -eq "Office 2016"}
Get-Content -Path $FilterFileName | ForEach{
    $FilteredUpdates += $AllUpdates | Where LegacyName -like $_
}

Decline-WsusUpdates($FilteredUpdates)
Cleanup-Wsus
start "C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\130\Tools\Binn\SQLCMD.EXE" ("-S", "np:\\.\pipe\Microsoft##WID\tsql\query", "-i", "C:\Tools\Scripts\Wsus\Scripts-WsusDBMaintenance.sql")


#更新プログラムの一覧
#$Wsus.GetUpdates() | Where IsDeclined -eq $False | Select Title, ProductTitles, CreationDate, LegacyName | Out-GridView -Title "拒否された更新以外のすべて"
#$Wsus.GetUpdates() | Where {$_.IsDeclined -eq $False -and $_.IsApproved -eq $True} | Select Title, ProductTitles, CreationDate, LegacyName | Out-GridView -Title "承認済みの更新プログラム"
#$Wsus.GetUpdates() | Where {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Select Title, ProductTitles, CreationDate, LegacyName | Out-GridView -Title "未承認の更新プログラム"
#$Wsus.GetUpdates() | Where {$_.IsDeclined -eq $False -and $_.IsApproved -eq $False} | Export-Csv 未承認の更新プログラム.csv -Encoding UTF8
