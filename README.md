# msg2eml

複数環境でのファイル互換性を保つために、
Outlook の .msg ファイルから、
情報交換に最低限必要な情報のみ抜き出して .eml ファイルに変換する。

* Microsoft Outlook 環境必須。
* 日本標準時(JST)のみ対応。
* ファイル名は ディレクトリ内序数・date・sender・subject 情報から自動生成。
  * 書き換える必要がある場合はソースコード要変更。
  * (余裕があればフォーマッタを実装するかも。)

## インストール

```
pip install pywin32 tkinterdnd2 nuitka pillow chardet click PyAutoGUI rich
.\make.ps1 # exe が欲しい場合
```

## 使い方

以下のいずれかを実行する。

* CUI から `py main.py` を実行する。
  * 表示されるウィンドウに *.msg ファイルを D&D して使用。
  * 一度に複数のファイルを処理可能。
  * 終了する場合はダブルクリックする。
* `ms2eml.ps1` をダブルクリックで起動する。
  * `ms2eml_dd.ps1` はデバッグモード。
* `make.ps1` 実行後に生成される `msg2eml.exe` を実行する。
  * `msg2eml_debug.exe` はコンソール付き。
  * onefile ではないので python 環境やパッケージ不足で起動しない可能性がある。
  * exe 特有のバグが起きる可能性があるので注意。

実行時に Outlook でデータ読み込みを承認するよう要求される場合は許可する。
