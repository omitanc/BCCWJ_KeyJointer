# BCCWJ_KeyJointer
NINJAL 現代日本語書き言葉均衡コーパス（BCCWJ）のキー列の内容を、サンプルID列の数字を基準に結合するVBAです。

<br>

※このVBAは、Mac OS版のExcelでは動作しません。

<br>

# 使用方法

1. [中納言 BCCWJ](https://chunagon.ninjal.ac.jp/bccwj-nt/search)より、下記の検索条件で検索結果をダウンロード。
2. (任意の名前).csvにして任意のディレクトリに保存。
3. csvファイルをExcelで開く。
4. csvファイルのシートの名前を original に変更。
5. Excelの開発タブからVisual Basicを開き、ダウンロードした.basファイルをインポート
6. Ctrl+Sで保存の際、「いいえ」を選択して、.xlsmの拡張子で保存
7. VBAのスクリプトを実行。
8. Xボタンを押して保存せずにExcelを終了。
9. 同じディレクトリ内に作成されるoutputsフォルダ内に、データ整形後のcsvが生成されていれば完了。

<br>

# BCCWJでの検索条件式
検索には、「短単位検索」の「検索条件式で検索」を使用します
``` IN (registerName="出版・雑誌" GENRE GENRE1="総合" GENRE2="一般" AND (core="true" OR core="false"))```の部分には、任意のレジスター名とジャンル名、コア/非コアを指定してください。

<br>
例）「特定目的・法律」レジスターの「憲法」ジャンルの「非コア」についてクエリを実行する場合

```
キー: (語種="和" OR 語種="漢" OR 語種="外" OR 語種="混" OR 語種="固" OR 語種="記号")
  IN (registerName="特定目的・法律" GENRE GENRE1="01 憲法" AND core="false")
  WITH OPTIONS tglKugiri="|" AND tglBunKugiri="#" AND limitToSelfSentence="1" AND tglFixVariable="2" AND tglWords="20" AND unit="1" AND encoding="UTF-16LE" AND endOfLine="CRLF";
```
