# msg2eml

複数環境でのファイル互換性を保つために、
Outlook の .msg ファイルから、
情報交換に最低限必要な情報のみ抜き出して .eml ファイルに変換する。

* Microsoft Outlook 環境必須。
* 日本標準時(JST)のみ対応。
* ファイル名は ディレクトリ内序数・date・sender・subject 情報から自動生成。
  * 書き換える必要がある場合はソースコード要変更。

# 使い方

以下のいずれかを実行する。

* `python main.py <email.msg>`
* (Windows 環境) msg2eml.bat に msg ファイルをドラックアンドドロップする。

実行時に Outlook でデータ読み込みを承認するよう要求されるので注意。
