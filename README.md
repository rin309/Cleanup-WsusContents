# Cleanup-WsusContents

## これらのスクリプトの実行は、実行した本人が責任を負うものとします。


## 以下のコンポーネントをインストールしてください
- Microsoft® ODBC Driver 13.1 for SQL Server https://www.microsoft.com/ja-JP/download/details.aspx?id=53339
- Microsoft Command Line Utilities 14.0 for SQL Server https://www.microsoft.com/en-us/download/details.aspx?id=53591


## インストール先
以下のファイルをすべてダウンロードします
- https://github.com/morokoshi/Cleanup-WsusContents/tree/master/Cleanup-WsusContents/Cleanup-WsusContents
一部ショートカットなどから絶対パスで記述している項目がありますので、以下のフォルダーに保存してください。
- C:\Tools\Scripts\Wsus

## はじめに
対象のOS用のフィルターを作成する必要があります。
サンプルとして、OS・プラットフォームごとに"Filter-"から始まるファイルを作成しました。
以下のファイルのうち、"$FilterFileName"の値を作成したフィルターのパスを指定してください。
- https://github.com/morokoshi/Cleanup-WsusContents/blob/master/Cleanup-WsusContents/Cleanup-WsusContents/Cleanup-WsusContents.ps1

Office製品を指定する場合のサンプルも記述しております。
Cleanup-WsusContents.ps1 のコメントアウトを1つずつ見ていくと Office 2016 などの文字列を確認できます。
このコード内では Office 2016 製品の64ビット版向け更新プログラムを拒否できるようにしております。


## 実行する
以下のコマンドを実行するとスクリプトを実行できます。
－PowerShell -ExecutionPolicy RemoteSigned -File "C:\Tools\Scripts\Wsus\Cleanup-WsusContents.ps1"


## タスクでスケジュールする
難しい話ではありませんが、インポートすれば使えるようにしております。
以下のいずれかのファイルをタスクスケジューラでインポートしてください。
- Task-Cleanup-WsusContents (Monthly).xml
- Task-Cleanup-WsusContents (Weekly).xml
