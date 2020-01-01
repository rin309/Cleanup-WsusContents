#Requires -Version 4.0
#Requires -RunAsAdministrator
$Version = New-Object System.Version("1.2020.101")
# Cleanup-WsusContents (CWS)
#
# このスクリプトは現状ベースで作成されたものです。今後の更新プログラムに対応するには、WSUSコンソールかSettings.Current.jsonかスクリプトのメンテナンスが必要になることを理解してください。
# このスクリプトを利用したことによる問題に対する責任は一切負いません。実行する前に必ず検証をしてください。

#最初に Settings.Current.json をメンテナンスしてください
$CuttentSettingsPath = "Assets\Settings.Current.json"
$DefaultSettingsPath = "Assets\Settings.Default.json"
Set-Location -Path (Split-Path -Parent ($MyInvocation.MyCommand.Path))
$Host.PrivateData.VerboseForegroundColor = "Cyan"

Function Load-Settings(){
    $SqlServerName = (Get-Item -Path "Registry::HKLM\SOFTWARE\Microsoft\Update Services\Server\Setup").GetValue("SqlServerName")
	
	If (Test-Path $CuttentSettingsPath){
        $Settings = Get-Content $CuttentSettingsPath -Encoding UTF8 -Raw | ConvertFrom-Json
        $Script:FeatureUpdatesFilterFileNames = $Settings.DeclineRule.FeatureUpdatesFilter
        $Script:FeatureUpdatesOutcludeFilter = $Settings.DeclineRule.FeatureUpdatesOutcludeFilter
        $Script:IsDeclineFeatureUpdatesClientBusiness = $Settings.DeclineRule.IsDeclineFeatureUpdatesClientBusiness
        $Script:IsDeclineFeatureUpdatesClientConsumer = $Settings.DeclineRule.IsDeclineFeatureUpdatesClientConsumer
        $Script:QualityUpdatesFilterFileNames = $Settings.DeclineRule.QualityUpdatesFilter
        $Script:DummyFileName = $Settings.ReservedFile.Name
        $Script:DummyFileSize = $Settings.ReservedFile.Size
        $Script:WsusDBMaintenanceScriptPath = $Settings.MaintenanceSql.ScriptPath
        $Script:SqlCmdPath = $Settings.MaintenanceSql.SqlCmdPath
        $Script:SqlServerPath = $Settings.MaintenanceSql.ServerPath
        $Script:IsDeclineMsOfficeUpdates = $Settings.DeclineRule.IsDeclineMsOfficeUpdates
        $Script:TargetMsOfficeArchitecture = $Settings.DeclineRule.TargetMsOfficeArchitecture
        $Script:WsusServer = $Settings.Wsus.Server
        $Script:ApproveNeededUpdatesRule = $Settings.ApproveNeededUpdatesRule
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

	$DefaultSettings = Get-Content $DefaultSettingsPath -Encoding UTF8 -Raw | ConvertFrom-Json
	If ($FeatureUpdatesFilterFileNames -eq $null) {$Script:FeatureUpdatesFilterFileNames = $DefaultSettings.DeclineRule.FeatureUpdatesFilter}
	If ($FeatureUpdatesOutcludeFilter -eq $null) {$Script:FeatureUpdatesOutcludeFilter = $DefaultSettings.DeclineRule.FeatureUpdatesOutcludeFilter}
	If ($IsDeclineFeatureUpdatesClientBusiness -eq $null) {$Script:IsDeclineFeatureUpdatesClientBusiness = $DefaultSettings.DeclineRule.IsDeclineFeatureUpdatesClientBusiness}
	If ($IsDeclineFeatureUpdatesClientConsumer -eq $null) {$Script:IsDeclineFeatureUpdatesClientConsumer = $DefaultSettings.DeclineRule.IsDeclineFeatureUpdatesClientConsumer}
	If ($QualityUpdatesFilterFileNames -eq $null) {$Script:QualityUpdatesFilterFileNames = $DefaultSettings.DeclineRule.QualityUpdatesFilter}
	If ($DummyFileName -eq $null) {$Script:DummyFileName = $DefaultSettings.ReservedFile.Name}
	If ($DummyFileSize -eq $null) {$Script:DummyFileSize = $DefaultSettings.ReservedFile.Size}
	If ($WsusDBMaintenanceScriptPath -eq $null) {$Script:WsusDBMaintenanceScriptPath = $DefaultSettings.MaintenanceSql.ScriptPath}
	If ($SqlCmdPath -eq $null) {$Script:SqlCmdPath = $DefaultSettings.MaintenanceSql.SqlCmdPath}
	If ($SqlServerPath -eq $null) {$Script:SqlServerPath = $DefaultSettings.MaintenanceSql.ServerPath}
	If ($IsDeclineMsOfficeUpdates -eq $null) {$Script:IsDeclineMsOfficeUpdates = $DefaultSettings.DeclineRule.IsDeclineMsOfficeUpdates}
	If ($TargetMsOfficeArchitecture -eq $null) {$Script:TargetMsOfficeArchitecture = $DefaultSettings.DeclineRule.TargetMsOfficeArchitecture}
	If ($WsusServer -eq $null) {$Script:WsusServer = $DefaultSettings.Wsus.Server}
	If ($ApproveNeededUpdatesRule -eq $null) {$Script:ApproveNeededUpdatesRule = $DefaultSettings.ApproveNeededUpdatesRule}
	If ($WsusPort -eq $null) {$Script:WsusPort = $DefaultSettings.Wsus.Port}
	If ($IsLogging -eq $null) {$Script:IsLogging = $DefaultSettings.Log.IsLogging}
	If ($LogMaximumCount -eq $null) {$Script:LogMaximumCount = $DefaultSettings.Log.MaximumCount}

	$Script:SqlServerPath = $SqlServerPath.Replace("$SqlServerName",$SqlServerName)
	$Script:DummyFilePath = Join-Path $WsusInstallDirectory $DummyFileName
}
Function Check-Settings(){
	Write-Host "FeatureUpdatesFilterFileNames: $FeatureUpdatesFilterFileNames" -Verbose
	Write-Host "FeatureUpdatesOutcludeFilter: $FeatureUpdatesOutcludeFilter" -Verbose
	Write-Host "IsDeclineFeatureUpdatesClientBusiness: $IsDeclineFeatureUpdatesClientBusiness" -Verbose
	Write-Host "IsDeclineFeatureUpdatesClientConsumer: $IsDeclineFeatureUpdatesClientConsumer" -Verbose
	Write-Host "QualityUpdatesFilterFileNames: $QualityUpdatesFilterFileNames" -Verbose
	Write-Host "DummyFileName: $DummyFileName" -Verbose
	Write-Host "WsusDBMaintenanceScriptPath: $WsusDBMaintenanceScriptPath" -Verbose
	Write-Host "SqlCmdPath: $SqlCmdPath" -Verbose
	Write-Host "SqlServerPath: $SqlServerPath" -Verbose
	Write-Host "IsDeclineMsOfficeUpdates: $IsDeclineMsOfficeUpdates" -Verbose
	Write-Host "TargetMsOfficeArchitecture: $TargetMsOfficeArchitecture" -Verbose
	Write-Host "WsusServer: $WsusServer" -Verbose
	Write-Host "ApproveNeededUpdatesRule: $ApproveNeededUpdatesRule" -Verbose
	Write-Host "WsusPort: $WsusPort" -Verbose
	Write-Host "WsusInstallDirectory: $WsusInstallDirectory" -Verbose
	Write-Host "IsLogging: $IsLogging" -Verbose
	Write-Host "LogMaximumCount: $LogMaximumCount" -Verbose

    Clear-Host
}
#特定の更新プログラムを承認
Function Approve-Updates($FilteredUpdates, $TargetGroup){
    If ($FilteredUpdates -ne $null){
	    $DeclineUpdatesCount = 0
	    $FilteredUpdates | ForEach-Object {
		    $_.Approve("install",$TargetGroup)
            If ($DeclineUpdatesCount -eq 0){
                Write-Progress -Activity "プログラムを承認" -Status $_.Title -PercentComplete 0
            }
            Else{
                Write-Progress -Activity "プログラムを承認" -Status $_.Title -PercentComplete ($DeclineUpdatesCount / $FilteredUpdates.Count * 100)
		    }
            $DeclineUpdatesCount++
	    }
    }
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
	Write-Progress -Activity "クリーンアップしています" -Status "1/4 - 削除された古い更新プログラム" -PercentComplete (0 / 4 * 100)
	$Wsus | Invoke-WsusServerCleanup -DeclineSupersededUpdates
	Write-Progress -Activity "クリーンアップしています" -Status "2/4 - 圧縮された更新プログラム" -PercentComplete (1 / 4 * 100)
	$Wsus | Invoke-WsusServerCleanup -CompressUpdates
	Write-Progress -Activity "クリーンアップしています" -Status "3/4 - 削除された古い更新プログラム" -PercentComplete (2 / 4 * 100)
	$Wsus | Invoke-WsusServerCleanup -CleanupObsoleteUpdates
	Write-Progress -Activity "クリーンアップしています" -Status "4/4 - 解放されたディスク領域" -PercentComplete (3 / 4 * 100)
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
				$LogsDirectoryChildItem | Remove-Item -Recurse -Force
				#西暦上2桁"20"から始まるディレクトリを検索
                If ((Get-ChildItem "Logs\" -Directory -Filter "20*").Length -le $LogMaximumCount){
					break
				}
			}
		}

		Get-PSDrive -Name $WsusInstallDirectory.SubString(0,1)
	}
}
Function Export-CsvFromWsusUpdates{
	param (
		[Parameter(ValueFromPipeline=$true,Mandatory=$true)]$Updates,
		[Parameter(Mandatory=$true)][String]$FileName
	)
	$Updates | Select-Object Title, @{Name="ProductTitles";Expression={($_.ProductTitles)}}, CreationDate, LegacyName | Export-Csv -NoTypeInformation "$LogDirectory\$FileName.csv" -Encoding UTF8
}

Load-Settings
Start-Logging
Check-Settings

Write-Host "* WSUSへ接続" -Verbose
$Wsus = Get-WsusServer -Name $WsusServer -PortNumber $WsusPort
Clear-Host

If (Test-Path $DummyFilePath){
	Write-Host "空き領域が少なくなりがちな環境で、スクリプトが正常に動作するためのダミーファイルを削除する" -Verbose
	Remove-Item -Path $DummyFilePath -Force | Out-Null
	Clear-Host
}


Write-Host "* 更新プログラムを拒否" -Verbose
Write-Host "** 置き換えられた更新プログラムを拒否" -Verbose
$FilteredUpdates = @()
$FilteredUpdates = $Wsus.GetUpdates() | Where-Object {!($_.IsDeclined) -and $_.IsSuperseded -and !($_.HasSupersededUpdates)}
Decline-QualityUpdates($FilteredUpdates)
If ($IsLogging){
	$FilteredUpdates | Where-Object {!($_.IsDeclined) -and !($_.IsApproved)} | Export-CsvFromWsusUpdates -FileName "2-1 拒否済み - 置き換えられた更新プログラム"
}


Write-Host "** 機能更新プログラムを拒否" -Verbose
$FilteredUpdates = @()
$AllUpdates = $Wsus.GetUpdates() | Where-Object {!($_.IsDeclined) -and $_.UpdateClassificationTitle -eq "Upgrades"}
ForEach ($FeatureUpdatesFilterFileName in $FeatureUpdatesFilterFileNames){
	If ($IsDeclineFeatureUpdatesClientBusiness){
		$FeatureUpdatesFilters += "*CLIENTBUSINESS*`r`n"
	}
	If ($IsDeclineFeatureUpdatesClientConsumer){
		$FeatureUpdatesFilters += "*CLIENTCONSUMER*`r`n"
	}
	$FeatureUpdatesFilters = Get-Content -Path "Filters\FeatureUpdates\$FeatureUpdatesFilterFileName"
	$AllUpdates | ForEach-Object {
		$UpdateInformation = $_
		$FeatureUpdatesFilters -split "`r`n" | ForEach-Object {
			If ($UpdateInformation.GetInstallableItems().Files.Name -like $_){
				$FilteredUpdates += $UpdateInformation
			}
		}
	}
}
If ($FeatureUpdatesOutcludeFilter){
	$AllUpdates | ForEach-Object {
		$UpdateInformation = $_
		If (($UpdateInformation.GetInstallableItems().Files.Name -match $FeatureUpdatesOutcludeFilter).Count -eq 0){
			$FilteredUpdates += $UpdateInformation
		}
	}
}
Decline-FeatureUpdates($FilteredUpdates)
If ($IsLogging){
	$FilteredUpdates | Where-Object {!($_.IsDeclined) -and !($_.IsApproved)} | Export-CsvFromWsusUpdates -FileName "2-2 拒否済み - 機能更新プログラム"
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
	$FilteredUpdates | Where-Object {!($_.IsDeclined) -and !($_.IsApproved)} | Export-CsvFromWsusUpdates -FileName "2-3 拒否済み - 品質更新プログラム"
}
Decline-QualityUpdates($FilteredUpdates)
Clear-Host


Write-Host "* クライアントから必要とされた未承認の更新プログラムを承認" -Verbose
$ApproveNeededUpdatesRule | ForEach-Object {
	If ($_.TargetGroupName){
		$TargetGroupName = $_.TargetGroupName
	}
	Else{
		$TargetGroupName = "すべてのコンピューター"
	}
	If (($Wsus.GetComputerTargetGroups() | Where-Object Name -eq $TargetGroupName).Count -eq 0){
		$Wsus.CreateComputerTargetGroup($TargetGroupName) | Out-Null
	}

	$UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
	#$UpdateScope.ApprovedStates = [Microsoft.UpdateServices.Administration.ApprovedStates]::NotApproved #未承認
	$UpdateScope.IncludedInstallationStates = [Microsoft.UpdateServices.Commands.WsusUpdateInstallationState]::Needed #必要とされている
	#$UpdateScope.ApprovedComputerTargetGroups.Add(($Wsus.GetComputerTargetGroups() | Where-Object Name -eq $TargetGroupName)) #ターゲットとなるグループ 
	#承認されたプログラムとグループ
	#$Wsus.GetUpdateApprovals($UpdateScope) | ForEach-Object { ($Wsus.GetComputerTargetGroup($_.ComputerTargetGroupId).Name + " にて承認: " + $Wsus.GetUpdate($_.UpdateId).Title)}

	$FilteredUpdates = @()
	If ($_.QualityUpdates){
		$FilteredUpdates += $Wsus.GetUpdates($UpdateScope) | Where-Object {!($_.Update.IsDeclined) -and !($_.Update.IsApproved) -and $_.Update.CreationDate -le [DateTime]::Now.AddDays(-($_.MinimumWaitDays))}
	}
	If ($_.FeatureUpdates){
		$FilteredUpdates += $Wsus.GetUpdates($UpdateScope) | Where-Object {!($_.Update.IsDeclined) -and !($_.Update.IsApproved) -and $_.Update.UpdateClassificationTitle -ne "Upgrades" -and $_.Update.CreationDate -le [DateTime]::Now.AddDays(-($_.MinimumWaitDays))}
	}
	$FilteredUpdates | Export-CsvFromWsusUpdates -FileName "3 承認 - $TargetGroupName"
	Approve-Updates $FilteredUpdates ($Wsus.GetComputerTargetGroups() | Where-Object Name -eq $TargetGroupName)
}
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
	$Wsus.GetUpdates() | Where-Object {!($_.IsDeclined) -and !($_.IsApproved)} | Export-CsvFromWsusUpdates -FileName "4 未承認の更新プログラム"
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