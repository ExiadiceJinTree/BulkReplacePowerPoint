- [In English](./README.md)

# Overview
指定のフォルダと下位フォルダ内の、指定の置換元文字列正規表現を含むPower Pointファイル内の対象文字列を、指定の置換先文字列に置換します。

# Description
- Power Pointファイルとして、.pptと.pptxの両方に対応しています。
- 検索(事前確認)と置換
  - 検索(事前確認):  
    置換を実行する前に、置換元文字列を含むPower Pointファイルのフォルダパス、ファイル名、及び、ファイル内の置換元文字列を含むテキストを、TSVファイルとして出力して、置換されるファイルとテキストを事前に確認できます。
    - 出力されるTSVファイル
      - 出力先フォルダパス: 実行スクリプトが存在するフォルダ。
      - ファイル名パターン: SearchResult_YYYYMMDDhhmmss.tsv
  - 検索(事前確認)処理と置換処理は、それぞれ任意で実行するかしないかを選べます。
- 検索/置換元とする文字列の指定について
  - 正規表現として扱われますので注意してください。
  - 大文字小文字の区別をしません。

# Usage
BulkReplacePowerPoint.vbsファイルを実行して、表示されるプロンプトダイアログの指示に従ってください。
