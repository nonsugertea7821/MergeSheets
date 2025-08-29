# 内部設計書 - Excelシート統合マクロ

## 1. モジュール構成

| モジュール名     | 概要                                           | 主な処理内容 |
|-----------------|----------------------------------------------|--------------|
| MergeSheets      | メイン制御                                     | 設定取得、結果シート準備、ロック確認、日付フォルダ作成、統合元フォルダ処理、統合先保存、エラー処理 |
| ReadConfig       | 設定取得                                     | MergeSheetsシートのB4～B7セルからConfig構造体に値を格納 |
| PrepareResultSheet | 統合結果シート準備                          | 「Results」シート作成、ヘッダ初期化、既存データクリア |
| CheckLock        | ロック確認                                   | 統合先ブックのロックファイルを確認し、エラー発生 |
| MergeFolder      | フォルダ統合（再帰）                          | フォルダ内のExcelファイルを取得しMergeSheetを実行。サブフォルダも再帰処理 |
| MergeSheet       | シート統合                                   | 同名シートチェック、上書き処理、コピー、統合結果記録 |
| UpdateResult     | 統合結果更新                                 | Resultsシートに統合結果（シート名、状態、統合元パス）を追加または更新 |

---

## 2. データ構造

```vba
Type Config
    TargetBook As String     ' 統合先ブックフルパス
    RootFolder As String     ' 統合元ルートフォルダパス
    Overwrite As Boolean     ' 同名シート上書き許可フラグ
End Type
````

---

## 3. 処理フロー

### 3.1 MergeSheets (メイン処理)

1. `ReadConfig` で MergeSheetsシートから設定取得
2. `PrepareResultSheet` で統合結果シートを準備
3. `CheckLock` で統合先ブックのロック確認
4. FileSystemObject作成
5. 日付フォルダ作成（マージファイル保存用）
6. 統合先ブックをコピーしてマージ用新規ファイル作成
7. `MergeFolder` を呼び出して統合元フォルダの全Excelファイルを再帰的に統合
8. 統合先ブックを保存して閉じる
9. エラー発生時はマージファイル削除、統合先ブック閉鎖、ユーザーメッセージ表示

### 3.2 MergeFolder

* 指定フォルダ内のExcelファイルをすべて取得
* 各ファイルのWorkbookを開き、各Worksheetに対して `MergeSheet` を実行
* サブフォルダも再帰的に処理

### 3.3 MergeSheet

* 統合先ブックに同名シートが存在するか確認
* `Overwrite = TRUE` の場合：

  * 仮シート名 (`_tmp`) でコピー
  * 元のシートを削除
  * 仮シートの名前を元に戻す
* `Overwrite = FALSE` の場合：

  * ERR\_SHEET\_DUPLICATE を発生
* 統合元の相対パスを取得して `UpdateResult` に記録

### 3.4 UpdateResult

* 統合結果シート（Results）にシート名、状態（追加/上書き）、統合元パスを追加または更新

---

## 4. エラー処理

* メイン処理で `On Error GoTo ErrHandler`
* エラー番号に応じたユーザーメッセージ生成
* 統合先ブック閉鎖、作成済みマージファイル削除
* エラー番号一覧：

  * ERR\_NO\_TARGET\_FILE (統合先ブックが存在しない)
  * ERR\_LOCK\_FILE (統合先ブックがロックされている)
  * ERR\_SHEET\_DUPLICATE (上書き不可の同名シート)
  * その他予期せぬエラー

---

## 5. ファイル・シート管理

* 統合先ブックはコピーして `_merged_yyyymmdd_hhnnss.xlsx` として保存
* 統合元ブックは読み取り専用で開き、非表示で処理
* 統合結果は「Results」シートに記録し、過去の統合結果も更新可能

