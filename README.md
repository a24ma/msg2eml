# msg2eml

複数環境でのファイル互換性を保つために、
Outlook の .msg ファイルから、
情報交換に最低限必要な情報のみ抜き出して .eml ファイルに変換する。

* Microsoft Outlook 環境必須。
* 日本標準時(JST)のみ対応。
* ファイル名は ディレクトリ内序数・date・sender・subject 情報から自動生成。
  * 書き換える必要がある場合はソースコード要変更。
  * (余裕があればフォーマッタを実装するかも。)

# インストール

```
pip install pywin32 pathlib
pip install --upgrade PySimpleGUIQt
```

# 使い方

以下のいずれかを実行する。

* `python main.py` を実行する(GUIモードで起動)。
  * 表示される白色領域に *.msg ファイルを D&D して使用。
* `python main.py <email.msg>` を実行する(CUIモードで起動)。
  * 一度に一つのファイルのみ処理が可能。
* (Windows 環境) msg2eml.bat に msg ファイルをドラックアンドドロップする。
  * 一度に複数のファイルの処理が可能。
  * パスに半角空白文字を含む場合はエラーとなるので注意。

実行時に Outlook でデータ読み込みを承認するよう要求されるので注意。
