# Altium / Tool / PCBA-JLC

JLCPCB PCBA向けのツール集です。


## Convert_BOM_File

Altium Designer BOM ファイルのコンバーターです。

部品の在庫数を取得し、BOMに書き足します。

（オンライン取得処理のため、ネット接続必須、JLCPCBが落ちているときは使えません）

サーバー負荷防止のため処理ごとに1秒待機するのでデータ取得には時間がかかります。

JLCPCBの部品データは「[Parts](Parts)」フォルダ中にMarkdownで定義されています。




## Convert_Centroid_File

Altium Designer Pick and Place Locations ファイルのコンバーターです。

PnPに使用するCSVファイルの回転角をJLCPCBのPCBA向けに合わせて書き出します。

回転データは「[Rotation.md](Rotation.md)」を使用しています。







