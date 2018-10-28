**これらのスクリプトの実行は、実行した本人が責任を負うものとします。**

# Cleanup WsusContents.ps1 とは
公式ブログの方法ですとWindows 10の各機能更新プログラムごとの品質更新プログラムが自動でメンテナンスできません。
WSUS メンテナンス ガイド – Japan WSUS Support Team Blog
https://blogs.technet.microsoft.com/jpwsus/2018/03/08/maintenance-guide/

<img width="1024" alt="image" src="https://user-images.githubusercontent.com/760251/47601797-a5c0ad80-da10-11e8-81b3-fdb10d5e1fca.png">
このスクリプトでは、指定した古いビルドの品質更新プログラムを自動で拒否することを目的としています。
このスクリプトは現状ベースで提供され、今後正常に動作することや必要に応じてスクリプトが更新されることは保証されません。

# 機能
- 置き換えられた更新プログラムを拒否
- 指定された機能更新プログラム (Upgrades) を拒否 <$FeatureUpdatesFilterFileName で作成したフィルターを指定>
- 指定された品質更新プログラム (Upgrades) を拒否 <$QualityUpdatesFilterFileName で作成したフィルターを指定>
- #Office 2016 64ビット版 の更新プログラムをすべて拒否 <既定ではコメントアウト>
- WSUSのクリーンアップ (削除された古い更新プログラム, 圧縮された更新プログラム, 削除された古い更新プログラム, 解放されたディスク領域)
- WSUS DB インデックスの再構成 (https://gallery.technet.microsoft.com/scriptcenter/6f8cde49-5c52-4abd-9820-f1d270ddea61)
- 4GBの一時ファイルを作成

# 詳しくは Wiki をご覧ください
https://github.com/morokoshi/Cleanup-WsusContents/wiki
