# DataTable

`DataTable` は、Excel の表データを「列名つきのデータ」として扱うための VBA クラスモジュールです。  
表の読み込み、条件抽出、列の追加・更新、並べ替え、シートへの書き戻しを、Worksheet のセル操作よりわかりやすく書けます。

[詳しい説明は GitHub Pages の DataTable ページに整理しています。](https://kanko-tech.github.io/vba-tools/datatable.html)

## これは誰向けか

- Excel の表を VBA で扱う処理を効率的に書きたい人
- 行列ベースの定型処理を短く書きたい人

## 何ができるか

- シート上の表をヘッダー付きデータとして読み込める
- 条件に合う行だけを抽出できる
- 列の追加、名前変更、値の更新ができる
- 指定列で並べ替えできる
- 加工した結果をシートへ書き戻せる

## 使うメリット

セル位置（`A1:D8`など）をハードコードする処理を減らし、「どの列に対して何をするか」をコード上で明確にできます。
そのため、転記ミスや条件分岐の書き間違いを減らしつつ、保守しやすい表処理を書きやすくなります。

## 最低限の使い方

1. `src/ExcelVBA/DataTable/DataTable.cls` を VBA プロジェクトへインポートします
2. ヘッダー行を含む表を `read_range` で読み込みます
3. 抽出や更新を行い、`to_range` でシートへ書き戻します

## 最小例

以下は、`A1:D8` の表をヘッダー行付きで読み込み、`status = "OK"` の行だけを抽出して `A10` から書き出す例です。

```vb
Sub Sample_DataTable_QuickStart()
    Dim tbl As New DataTable
    Dim okRows As DataTable

    tbl.read_range Sheet1.Range("A1:D8"), hasHeader:=True
    Set okRows = tbl.filter_by_equals("status", "OK")

    okRows.to_range Sheet1.Range("A10"), includeHeader:=True
End Sub
```

## 詳細情報

- [DataTable](https://kanko-tech.github.io/vba-tools/datatable.html)
- [DataTable 導入方法](https://kanko-tech.github.io/vba-tools/datatable-setup.html)
- [DataTable 使い方](https://kanko-tech.github.io/vba-tools/datatable-examples.html)
- [DataTable リファレンス](https://kanko-tech.github.io/vba-tools/datatable-reference.html)
- [vba-tools](https://kanko-tech.github.io/vba-tools/)
- ソースコード: `src/ExcelVBA/DataTable`
- ライセンス: `LICENSE`

## 依存関係

`DataTable` を利用するには、`Vector` と `Matrix` が必要です。  
あわせて VBA プロジェクトへインポートして使用してください。
