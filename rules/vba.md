---
paths:
  - "**/*.bas"
  - "**/*.cls"
  - "**/*.frm"
---

## VBA開発の注意事項

### AscW の Int16 オーバーフロー問題

**適用条件**: VBAで文字のUnicodeコードポイントを数値として扱う処理（文字種判定・範囲チェック等）を書く場合、`AscW` の戻り値に対して必ず以下の補正を行うこと。

**理由**: `AscW` は Int16（符号付き16bit）を返すため、U+8000 以上の文字は負値になる。また `&H9FFF` 等の16進リテラルも Int16 オーバーフローで負値に解釈される。CJK漢字（U+4E00〜U+9FFF）・半角カタカナ（U+FF65〜U+FF9F）の範囲チェックが破綻する。カタカナ（U+30A0〜U+30FF）は値が 32767 以下のため偶然動作するが、CJK漢字のみを含む文字列が「日本語なし」と誤判定される。

```vb
' NG: 補正なし（U+8000以上の文字で誤判定）
If (AscW(c) >= &H4E00 And AscW(c) <= &H9FFF) Then  ' CJK漢字の判定が破綻

' OK: 負値を正しいコードポイントに補正してから判定
Dim cp As Long
cp = AscW(Mid(s, i, 1))
If cp < 0 Then cp = cp + 65536  ' U+8000以上の文字を正しいコードポイントに補正
' 上限値は16進リテラルを使わず10進数リテラルを使うこと（Int16オーバーフロー回避）
If (cp >= &H4E00 And cp <= 40959) Or _   ' CJK漢字（U+4E00〜U+9FFF）
   (cp >= 65381   And cp <= 65439) Then   ' 半角カタカナ（U+FF65〜U+FF9F）
```

### ファイルダイアログの初期ディレクトリ指定

**禁止**: ファイルダイアログの初期ディレクトリ指定に `ChDrive` / `ChDir` を使わないこと。UNCパス（`\\server\share\...`）で実行時エラー'68'が発生する。代わりに `GetOpenFilename` / `GetSaveAsFilename` の `InitialFileName` 引数を使うこと。

```vb
' NG: UNCパスでエラー68・カレントディレクトリへの副作用もある
If initDir <> "" Then
    ChDrive Left(initDir, 1)
    ChDir initDir
End If
filePath = Application.GetOpenFilename(filter, , title)

' OK: UNCパスでも動作・カレントディレクトリへの副作用もない
If Right(initDir, 1) <> "\" Then initDir = initDir & "\"
filePath = Application.GetOpenFilename(filter, , title, initDir)
' GetSaveAsFilename も同様（InitialFileName は第4引数）
filePath = Application.GetSaveAsFilename(, filter, , title, initDir)
```
