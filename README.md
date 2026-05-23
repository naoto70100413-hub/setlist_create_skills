# claude-code-skills

Claude Code 用の自作スキル集。

## スキル一覧

### dance-setlist

ダンスイベントのセットリスト・スケジュール表を自動作成する。

- 掛け持ちメンバーの制約・ブロック割り・休憩時間を考慮して出演順を最適化
- Spotify API で再生時間を自動取得
- 時刻入りのスケジュール表を Excel で出力

### dance-setlist-formatter

ダンスイベントの申込 CSV を dance-setlist スキル用の入力 Excel に整形する。

- 曲情報 CSV とチーム情報 CSV の 2 ファイルを受け取り受付番号をキーに結合
- 掛け持ち列を自動統合して所定の 5 列 Excel を出力

## インストール

```bash
cp -r dance-setlist ~/.claude/skills/
cp -r dance-setlist-formatter ~/.claude/skills/
```
