# csv2excel

csvファイルから単純にエクセルファイルを生成するPythonスクリプトです。csvファイルの内容はエクセルファイルにコピーされます。エクセルファイルのファイル名はcsvファイルのものを引き継ぎます。

## Windowsでの使い方

Pythonはインストールされているとします。このスクリプトは
[openpyxl](https://pypi.org/project/openpyxl/)を使うので，もしまだインストールしていなければ
```
py -m pip install openpyxl
```
を実行します。たとえば
```
py csv2excel.py sample.csv
```
とすると，同じディレクトリに`sample.xlsx`ができます。同様に
```
py csv2excel.py sample.csv sample2.csv
```
とすると，`sample.xlsx`と`sample2.xlsx`とが同じディレクトリにできます。拡張子が`.csv`でないファイルはスキップするようになっています。

## Windowsで単体で実行可能にする手順

[Pyinstaller](https://pypi.org/project/pyinstaller/)が有名ですが，[Nuitka](https://github.com/Nuitka/Nuitka)を用いることをお勧めします。
```
py -m pip install nuitka
```
とします。初回のコンパイル時には途中，YesまたはNoの質問が何回かありますが，Yesと答えます。これで`mingw64`(cコンパイラ環境)を利用するための準備をしてくれます。また，実行ファイルにアイコン画像を指定したいときは
```
py -m pip install imageio
```
によって[imageio](https://pypi.org/project/imageio/)を入れておきます。

次を実行すれば，`csv2excel.py`に対応する全部入りの実行ファイルができます。
```
py -m nuitka --disable-console --onefile --windows-icon-from-ico=icon.png csv2excel.py
```
数分後，`csv2excel.exe`ができているはずです。（`dist`の中にも同じものができているが、こちらは`python???.dll`がないと動かないようです）

これによってPythonが入っていない環境でも実行できるようになります。
```
csv2excel.exe sample.csv sample2.csv
```
とすれば，Excelファイルが２つできます。なお，アンチ・ウィルスソフトが動いている場合は，コンパイルが途中で止まってしまうことがあります。その場合にはアンチ・ウィルスソフトを一時的に停止させておくとよいでしょう。

## ドラッグ・アンド・ドロップ

Windowsのデスクトップ上では，2つのcsvファイルをマウスで選択し，`csv2excel.exe`のアイコンにドラッグ＆ドロップすることでExcelファイルを生成することができます。これはコマンドラインの別の方法で，土曜にcsvファイルと同じディレクトリに `sample.xlsx` と `sample2.xlsx` ができます。