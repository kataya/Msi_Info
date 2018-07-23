
# Msi_Info

MSIファイルやMSPファイルのアンインストールは、[プログラムのアンインストールまたは変更]から行うのが一般的ですが、アプリケーション開発時には、特定のプログラムをアンインストールすることが頻繁に発生します。

また、サイレントアンインストールを行うときには、MSIファイルやMSPファイルの GUID を指定してコマンドを実行すると、イントールしたときの*.msiや *.msp のファイルがなくても、アンインストールを行うことが可能です。


**MSI ファイルのアンインストールコマンド：**

>mxiexec /X "Pruduct GUID"

**MSP ファイルのアンインストールコマンド：**

>msiexec /I "Product GUID" MSIPATCHREMOVE="Patch GUID"


MSIファイルやMSPファイルのプロパティを確認するには Windows SDK に含まれる Orca を利用することが多いと思いますが、Orcaを利用するためだけに Windows SDK をダウンロードしてインストールするのも面倒なので、.NETでアンインストール用コマンドを取得するためのサンプルコードを作成しました。
コンパイル済みのアプリケーションは[こちら](https://github.com/kataya/Msi_Info/files/2218401/MsiInfo_v1.0.zip) にあります。

なお、同様のことをPowershellで実装したサンプルは[こちら](https://github.com/EsriJapan/arcgis-install-batch/tree/master/Ps_Scripts) にあります。


**関連情報）**　[How To: Silently uninstall ArcGIS products](https://support.esri.com/en/technical-article/000013200)

