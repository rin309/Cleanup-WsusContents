#
# Set_ValueToWsusPool.ps1
#
#Windows Server Update Services のベスト プラクティス
#https://support.microsoft.com/ja-jp/help/4490414/windows-server-update-services-best-practices

Import-Module WebAdministration
$Script:WsusPoolPath = "iis:\AppPools\WsusPool"

Function Set-ValueWsusPool($ItemName, $DisplayName, $NewValue){
	$BeforeValue = (Get-ItemProperty $WsusPoolPath -Name $ItemName).Value
	Set-ItemProperty $WsusPoolPath -Name ([System.String]$ItemName) -Value $NewValue
	$AfterValue = (Get-ItemProperty $WsusPoolPath -Name $ItemName).Value
	Write-Host "$DisplayName (変更前: $BeforeValue, 変更後: $AfterValue)"
}

#リサイクルを無効にし、メモリ制限を構成する
Set-ValueWsusPool "queueLength" "(全般)\キューの長さ" 2000
Set-ValueWsusPool "processModel.idleTimeout" "プロセスモデル\アイドル状態のタイムアウト" "0:00:00"
Set-ValueWsusPool "processModel.pingingEnabled" "プロセスモデル\Pingの有効化" $false
Set-ValueWsusPool "recycling.periodicRestart.time" "リサイクル\定期的な期間" "0:00:00"
Set-ValueWsusPool "recycling.periodicRestart.privateMemory" "リサイクル\プライベートメモリ制限" 4000000

#recycling.periodicRestart.privateMemoryの値は暫定値
#上記においても支障がある場合は増やす・メモリ増設にて対処する
#https://social.msdn.microsoft.com/Forums/ja-JP/0dc69153-1d4e-4e91-bf91-df311424a8be/12300wsus?forum=jpsccmwsus
