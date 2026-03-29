# Contributing to mikuproject

Thank you for contributing to `mikuproject`.

This project accepts bug reports, feature requests, documentation fixes, tests, and pull requests.

## Before You Open an Issue or Pull Request

- Check whether the topic is already covered by an existing issue or pull request.
- For behavior changes, describe the expected behavior and the current behavior clearly.
- For code changes, include or update tests when practical.
- Keep changes focused. Small, reviewable pull requests are preferred.

## Development Notes

- `index.html` and `mikuproject.html` are generated files.
- Edit `index-src.html`, `mikuproject-src.html`, and files under `src/` instead of editing generated output directly unless regeneration is intentionally part of the change.
- Application logic should normally be edited in `src/ts/`.
- `src/js/` is generated from `src/ts/`, but is currently committed to Git. If you change `src/ts/`, regenerate `src/js/` as part of the same change.
- When a behavior change affects project structure, input/output rules, or AI integration, update the relevant documentation as well.

Documentation roles:

- `README.md`: repository entry point and quick start
- `docs/architecture.md`: overall structure, build flow, generated files, and operational rules
- `docs/spec.md`: format and behavior specifications
- `docs/TODO.md`: incomplete work only

## Recommended Checks

Run relevant commands before submitting a pull request when possible.

```bash
npm run build:js
npm run build:html
npm test
```

If your change touches sample workbook generation, also run:

```bash
npm run build:xlsx-sample
```

## Pull Request Guidelines

- Explain what changed and why.
- Mention any user-visible behavior change.
- Mention any specification or documentation updates if they are part of the change.
- If a change is incomplete or intentionally deferred, say so explicitly.

## Contribution License

By submitting an issue, pull request, comment, documentation change, code change, or other material that is intentionally submitted for inclusion in this project, you agree that:

- Your contribution is provided under the Apache License 2.0 used by this repository.
- The project maintainer may use, modify, rewrite, adapt, edit, and redistribute your contribution as part of this project, as permitted by the project license structure.
- You have the right to submit the contribution.
- Unless you explicitly state otherwise, your contribution is treated as a "Contribution" under Section 5 of the Apache License 2.0.

If you do not want a submission to be treated as a contribution for inclusion in the project, mark that clearly and do not open it as a pull request intended to be merged.

## Attribution

Contributors may be acknowledged in project history, release notes, or other project documents at the maintainer's discretion.

## Code of Collaboration

- Be specific.
- Be respectful.
- Prefer concrete repro steps, fixtures, and tests over vague reports.

---

# mikuproject へのコントリビュート

`mikuproject` へのコントリビュートありがとうございます。

このプロジェクトでは、バグ報告、機能提案、ドキュメント修正、テスト追加、Pull Request を受け付けます。

## Issue / Pull Request の前に

- 既存の issue / pull request と重複していないか確認してください。
- 挙動変更を伴う場合は、期待する挙動と現在の挙動を明確に書いてください。
- コード変更では、可能なら対応するテストも追加または更新してください。
- 変更は小さく、レビューしやすい単位が望ましいです。

## 開発メモ

- `index.html` と `mikuproject.html` は生成物です。
- 生成物を直接編集するのではなく、通常は `index-src.html`、`mikuproject-src.html`、`src/` 配下を編集してください。
- アプリロジックの変更は、通常 `src/ts/` で行ってください。
- `src/js/` は `src/ts/` から生成されますが、現状では Git 管理しています。`src/ts/` を変更した場合は、同じ変更で `src/js/` も更新してください。
- project 構造、入出力ルール、生成AI 連携の挙動を変える場合は、関連ドキュメントも更新してください。

ドキュメントの役割:

- `README.md`: リポジトリの入口と quick start
- `docs/architecture.md`: 全体構成、ビルド、生成物、運用ルール
- `docs/spec.md`: 形式仕様と挙動仕様
- `docs/TODO.md`: 未完了作業のみ

## 推奨チェック

可能であれば、Pull Request 前に関連コマンドを実行してください。

```bash
npm run build:js
npm run build:html
npm test
```

sample workbook 生成に関わる変更では、あわせて次も実行してください。

```bash
npm run build:xlsx-sample
```

## Pull Request のガイド

- 何を変えたか、なぜ変えたかを書いてください。
- ユーザーに見える挙動変更があれば明記してください。
- 仕様書やドキュメント更新を含む場合は、その旨も書いてください。
- 未完了部分や意図的に後回しにした点があれば、明示してください。

## コントリビューションのライセンス

このプロジェクトへ取り込みを意図して issue、pull request、コメント、ドキュメント変更、コード変更、その他の素材を提出した場合、次に同意したものとして扱います。

- あなたのコントリビューションは、このリポジトリで採用している Apache License 2.0 の下で提供されます。
- プロジェクト管理者は、そのコントリビューションを本プロジェクトの一部として、プロジェクトのライセンス構造で許容される範囲で、利用、修正、書き換え、調整、編集、再配布できます。
- あなたは、そのコントリビューションを提出する権利を持っています。
- あなたが明示的に別扱いを示さない限り、その提出物は Apache License 2.0 第5条の "Contribution" として扱われます。

取り込みを意図しない連絡については、その旨を明確に示してください。マージを前提としない相談は、その前提が分かるように記述してください。

## 謝辞

コントリビューター名は、必要に応じて git 履歴、リリースノート、その他の文書で言及されることがあります。

## コラボレーション方針

- 具体的に書く
- 相手を尊重する
- 曖昧な説明より、再現手順、fixture、テストを優先する
