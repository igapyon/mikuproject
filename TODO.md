# TODO

## mikuproject

- `Mermaid` ランタイムをこのリポジトリ内でどう扱うか決める
- `Mermaid` の SVG プレビューを、独立リポジトリ単体で再現できるようにする
- `build:app` 生成物としての `src/js/` をどこまで Git 管理するか方針を固める
- `src/ts/ -> src/js/ -> mikuproject.html` の生成責務を README と実装で一貫させる
- `build:xlsx-sample` を実行する前提条件と出力物の位置づけを README に追記する
- `local-data/` 配下のファイルを、参照用・検証用・生成物で整理する
- `local-data/` に置くべきでない生成物や一時ファイルがないか見直す
- `mikuproject-spec.md` と `README.md` と `TODO.md` の役割分担を整理する
- `mikuproject-spec.md` に残っている実装済み前提との差分を定期的に解消する
- `.xlsx import` の次段として、どのシート・列を今後 import 対象に広げるか整理する
- `Calendars` の `WeekDays / Exceptions / WorkWeeks` を今後も非対応で維持するか再判断する
- `Calendar / Baseline / TimephasedData / ExtendedAttributes` をどの順で扱うか優先順位を決める
- WBS workbook の表示改善を継続する
- WBS workbook の見た目改善と、構造忠実 workbook との責務分離を保つ
