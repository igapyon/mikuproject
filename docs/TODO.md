# TODO

この文書には、未完了の作業だけを書く。概要説明や仕様判断は `README.md` と `docs/spec.md` に寄せる。

## mikuproject

- 高優先: WBS workbook レイアウト調整用 helper を導入し、`A1 / C17 / S列` のような Excel 記法でセル位置を扱えるようにする
- 高優先: WBS workbook レイアウト調整用 helper から、必要時に `C17` のようなセル参照を `console.log` などで確認できる仕組みを入れる
- `excel-io` の workbook スタイルにフォントサイズ指定を追加し、XLSX 出力で大きい見出し文字を使えるようにする
- WBS workbook と `mikuproject-sample.xlsx` のタイトル行で、フォントサイズ指定をどこまで使うか整理する
- `Mermaid` ランタイムをこのリポジトリ内でどう扱うか決める
- `Mermaid` の SVG プレビューを、独立リポジトリ単体で再現できるようにする
- `local-data/` 配下のファイルを、参照用・検証用・生成物で整理する
- `local-data/` に置くべきでない生成物や一時ファイルがないか見直す
- `npm test` を `scripts/run-tests.mjs` ベースへ切り替えるか検討する
- `main.test.js` に残っている `xlsx import` の file input wiring を、UI 配線確認と import 結果確認へさらに分離する
- `main.ts` の summary / validation / preview 描画を別モジュール化し、DOM テストをさらに軽くする
- `build:xlsx-sample` の所要時間を個別計測し、sample workbook 生成処理の支配要因を確認する
- `main.test.js` の初期化 DOM をケース別に最小化できるか見直す
- CI 向けに `test:fast` と `test:full` のような実行導線を分けるか検討する
- `docs/spec.md` に残っている実装済み前提との差分を定期的に解消する
- `.xlsx import` の次段として、どのシート・列を今後 import 対象に広げるか整理する
- WBS 用の `ステータス` は `Task.ExtendedAttribute` で扱う前提で、`FieldID / FieldName / 値候補` を設計する
- `TaskStatus` 用 `ExtendedAttribute` を `mikuproject-sample.xlsx` と `WBS workbook` のどちらまで見せるか決める
- `TaskStatus` 用 `ExtendedAttribute` の値候補と、`PercentComplete` / `Active` との関係を整理する
- `Calendars` の `WeekDays / Exceptions / WorkWeeks` を今後も非対応で維持するか再判断する
- `Calendar / Baseline / TimephasedData / ExtendedAttributes` をどの順で扱うか優先順位を決める
- `mikuproject-sample.xlsx` の `Project` シートで、構造忠実方針を崩さない範囲の見た目調整を続ける
- `mikuproject-sample.xlsx` の `Resources / Assignments / NonWorkingDays` で、強調色が過剰にならない最終バランスを調整する
- WBS workbook の表示改善を継続する
- WBS workbook の見た目改善と、構造忠実 workbook との責務分離を保つ
