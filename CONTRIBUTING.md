# Contributing

## 前提環境

- Windows + Excel（デスクトップ版）
- PowerShell 7 以上
- ビルド・テスト時のみ、Excel の `トラスト センター` > `マクロの設定` > **「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」** を有効にする（`HKCU:\Software\Microsoft\Office\<ver>\Excel\Security\AccessVBOM = 1`）。作業後は元に戻すことを推奨します。

## ソースの編集箇所

VBA のロジックを変更する場合は `src/*.bas` を編集してください。`.xlsm` 内のコードを VBE で直接編集しても、次回のビルドで上書きされます。

`src/*.bas` は UTF-8（BOM なし）・LF 改行が正規の byte 表現です。エディタの改行コード設定に注意してください。cp932 で表現できない文字が含まれている場合、ビルドは中断します。

## candidate-only 経路

配布物への昇格は行わず、常に `-CandidateOnly` 経路で候補を生成してから検証してください。

```powershell
pwsh -NoProfile -File scripts/build-xlsm.ps1 -CandidateOnly -Profile Test -OutputPath .build-work/<candidate>.xlsm
pwsh -NoProfile -File scripts/run-release-validation.ps1
```

`run-release-validation.ps1` は候補ビルド、必須ケース 1〜68、性能ゲートを順に実行します。すべて green になることを確認してから変更を提案してください。

`-Profile Release` でも同じ source からビルドできることを確認してください。release candidate は test 専用 procedure を含まないため、production 側のロジックだけで正常にコンパイルされる必要があります。

## release safety subset

配布候補に対しては、test 専用 procedure を使わず通常の操作画面（`事前確認`）を経由して、次の3項目の拒否動作を検証します。

```powershell
pwsh -NoProfile -File scripts/Invoke-ReleaseSafetySubset.ps1 -CandidatePath <release-candidate>.xlsm
```

- 走査ルート外への移動の拒否
- reparse point 経由でルートを逸脱する移動先の拒否
- 移動先の衝突時に上書きしないこと

このスクリプトは専用の Excel インスタンスを起動し、そのプロセス ID に一致するダイアログだけを閉じます。他の Excel ウィンドウには干渉しませんが、実行中は対象の Excel を手動で操作しないでください。いずれも計画段階で拒否されるため、実ファイルは削除されません。

ゴミ箱が使えない環境での完全削除フォールバック拒否は、動的検証の対象外です（[SECURITY-MODEL.md](SECURITY-MODEL.md#安全性検査の検証範囲) を参照）。`modRecycleBin.bas` のガード節を変更する場合は、この検証が自動テストで守られていないことを踏まえてレビューしてください。

## テスト専用 procedure を追加する場合

新しいテスト専用 procedure（`*Test` で終わる名前など）を追加する場合は、`#If TEST_BUILD Then ... #End If` で囲んでください。この区画は release candidate のビルド時に `#Const TEST_BUILD = False` となり、コンパイルされません。

- production 側のロジックから呼ばれる補助関数（テストフック変数のリセットなど）は区画に含めず、通常の module scope に残してください。
- 新しい procedure には、対応する検証を `scripts/test-xlsm.ps1` の該当ケースへ追加し、実際に呼び出されることを確認してください。呼び出されない test 専用 procedure は残さないでください。

## 署名を壊す変更

VBA プロジェクトを編集・保存すると、（署名済み配布物がある場合）署名は外れます。署名は最終ビルド後、フルテストが green になったあとにのみ行います。

## セキュリティ報告

脆弱性を報告する場合は、公開の Issue ではなく [SECURITY.md](SECURITY.md) の手順に従ってください。
