# taxi-logger.xlsmのソースコード

src/taxi-logger.xlsmディレクトリは[taxi-logger](https://github.com/yoshinari/taxi-logger/blob/master/README.md)アプリでexportしたjsonファイルを読み込んで日報を作成するExcelマクロのソースコードです。
実行マクロはbinの下に有ります。
ソースコードは[Ariawaseのvbac.wsf](https://github.com/vbaidiot/Ariawase)を使ってdecombineしています。VBAの世界ではShift-JISで記述されているため、そのまま登録すると、Gitの世界では文字化けします。
そのため、decombine後、以下のコマンドでUTF-8に変換したファイルを追加登録しています。
```bash
iconv -f SJIS -t UTF8 taxi_logger.bas > taxi_logger_utf8.bas
```

##### 開発環境の設定
* Windows10 / Excel 2016の環境で開発しています。
- [VBA-JSON v2.2.3](https://github.com/VBA-tools/VBA-JSON/releases)をダウンロードし、解凍する。

- [ファイル]->[ファイルのインポート]
JsonConverter.basをインポートする

- [ツール]->[参照設定]
Microsoft Scripting Runtimeをチェックする
---