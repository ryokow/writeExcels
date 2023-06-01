# 本スクリプトについて
本スクリプトは複数のエクセルファイルに同じ内容を書き込むpythonスクリプトです。

書き込むシートとセル、内容はjsonファイルで指定します。

### 対応している形式
- .xlsx
- .xlsm
- .xls

# インストール
1. [リリースページ](https://github.com/ryokow/writeExcels/releases)から最新版をダウンロードする
2. zipファイルを任意の場所に展開する

### pywin32のインストール
本スクリプトの実行にはpywin32ライブラリが必要です。実行前に以下コマンドでインストールして下さい。
```
pip install pywin32
```

# 使い方
以下のコマンドでスクリプトを実行
```
python writeExcels.py <フォルダパス>
```
※pythonのバージョンによっては"python3"とする

- フォルダパスには処理したいエクセルファイルの格納場所を指定
- フォルダパス未指定の場合はスクリプトがあるフォルダ内のエクセルが処理対象となる

# jsonファイル
当スクリプトでは以下の情報をdata.jsonファイルから読み込み使用します。
- 対象シート名
- 対象セル＋内容（任意の数）

以下はSheet1のA6セルにXXX、B3セルにYYY、C2セルにZZZを書き込むサンプルです。
~~~data.json
{
  "sheet_name": "Sheet1",
  "cell_data": {
    "A6": "XXX",
    "B3": "YYY",
    "C2": "ZZZ"
  }
}
~~~
cell_dataは"A1:D10"のように範囲で指定することもできます。

data.jsonは必ず**writeExcels.pyと同じフォルダ**に置いてください。

# 便利な使い方（Windows）
処理対象エクセルを格納したフォルダを右クリックしコンテキストメニューの「送る」から当スクリプトを実行すると簡単にフォルダパスの指定ができます。

1. 実行スクリプトをbatファイル化する（<フォルダパス>を%1に置き換える）
2. batファイルのショートカットファイルを作成する
3. Win+Rで表示されたダイアログに"shell:sendto"と入力しOKを押す
4. 3で開かれたSendToフォルダに2のショートカットをコピーする

これでコンテキストメニューの「送る」からwriteExcels実行batを選択でき、対象のフォルダがスクリプトの引数として渡されます。
