# TODO

## mikuproject

- `Mermaid` ランタイムの更新手順を README と `src/vendor/` コメントのどちらへ寄せるか最終化する
- `Mermaid` の SVG プレビュー内包方針 (`src/vendor/mermaid` -> `build:app` -> single-file HTML) を README と実装で継続的に一致させる
- `build:app` 生成物としての `src/js/` をどこまで Git 管理するか方針を固める
- `src/ts/ -> src/js/ -> mikuproject.html` の生成責務を README と実装で一貫させる
- `build:xlsx-sample` を実行する前提条件と出力物の位置づけを README に追記する
- テストが重い原因を切り分けるため、`MS Project XML` の import/export と `XLSX` の import/export のどちらが支配的に遅いか計測する
- `XMLSerializer` / `DOMParser` と `XLSX` シリアライズ処理の改善余地を調べ、直列化ロジックの見直しで高速化できるか検討する
- `local-data/` 配下のファイルを、参照用・検証用・生成物で整理する
- `local-data/` に置くべきでない生成物や一時ファイルがないか見直す
- `mikuproject-spec.md` と `README.md` と `TODO.md` の役割分担を整理する
- `mikuproject-spec.md` に残っている実装済み前提との差分を定期的に解消する
- `.xlsx import` の次段として、どのシート・列を今後 import 対象に広げるか整理する
- `Calendars` の `WeekDays / Exceptions / WorkWeeks` を今後も非対応で維持するか再判断する
- `Calendar / Baseline / TimephasedData / ExtendedAttributes` をどの順で扱うか優先順位を決める
- WBS workbook の表示改善を継続する
- WBS workbook の見た目改善と、構造忠実 workbook との責務分離を保つ
