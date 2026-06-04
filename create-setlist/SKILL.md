---
name: create-setlist
description: >
  ダンスイベントのセットリスト・スケジュール表を自動作成する。出演チームリスト（Excel/CSV）を受け取り、
  掛け持ちメンバーの制約・ブロック割り・休憩時間を考慮して出演順を最適化し、時刻入りのスケジュール表をExcelで出力する。
  「セットリストを組んで」「出演順を決めて」「ダンスイベントのスケジュールを作って」「チームリストからスケジュール表を
  出力して」「出演チームが揃ったのでセットリストを作りたい」などと言われたら必ず使うこと。
  Excel/CSVファイルがアップロードされてダンスイベントの出演チームデータがある場合も必ず使うこと。
---

# ダンスイベント セットリスト作成スキル

## 概要

出演チームリスト（Excel/CSV）を入力として、制約を満たす最適な出演順に並び替え、
時刻入りスケジュール表（Excel）を生成する。
`scripts/generate_setlist.py` を使って一括処理する。

---

## 入力データ形式

| 列名（デフォルト） | 内容 | 例 |
|---|---|---|
| チーム名 | チームの名前 | TEAM_A |
| 曲名 | 演技で使用する曲 | ダンシングクイーン |
| アーティスト名 | 曲のアーティスト | ABBA |
| 掛け持ちチーム | 共通メンバーがいる他チーム（カンマ区切り） | TEAM_C, TEAM_E |
| 希望時間帯 | 出演したい時間帯 | 前半 / 後半 / 指定なし |

列名が異なる場合は `--col-*` オプションで上書きできる。

---

## 実行手順

### Step 1: 入力確認

ユーザーから以下を確認する（まだ得ていない場合）:
1. **入力ファイルのパス**（Excel/CSV）
2. **イベント開始時刻**（例: `13:00`）

### Step 2: スキルのディレクトリを特定する

このSKILL.mdが置かれているディレクトリのパスを `SKILL_DIR` とする。
Claude Codeでは `SKILL.md` の絶対パスから親ディレクトリを取得することで特定できる。

### Step 3: Spotify で再生時間を取得する

`anthropic-skills:spotify-duration` スキルを使って全チームの再生時間を事前に取得する。

**3-1: 曲リスト CSV を作成する**

入力ファイルを読み込み、`/tmp/setlist_songs.csv`（列: チーム名・曲名・アーティスト名）を生成する。
列名が異なる場合は実際の列名に置き換えること。

```bash
python3 -c "
import pandas as pd
from pathlib import Path
path = Path('<入力ファイルパス>')
df = pd.read_excel(path) if path.suffix in ('.xlsx','.xls') else pd.read_csv(path)
df = df.rename(columns={'<チーム名列>': 'チーム名', '<曲名列>': '曲名', '<アーティスト名列>': 'アーティスト名'})
df[['チーム名','曲名','アーティスト名']].to_csv('/tmp/setlist_songs.csv', index=False, encoding='utf-8-sig')
print('Exported', len(df), 'songs to /tmp/setlist_songs.csv')
"
```

**3-2: spotify-duration スキルを呼び出す**

`anthropic-skills:spotify-duration` スキルを呼び出し、`/tmp/setlist_songs.csv` に再生時間を追記してもらう。

**3-3: `--duration-overrides` JSON を組み立てる**

取得後のCSVを読み込んで `{"チーム名": 秒数}` の JSON を生成する。

```bash
python3 -c "
import pandas as pd, json, sys
df = pd.read_csv('/tmp/setlist_songs.csv')
dur_cols = [c for c in df.columns if c not in ('チーム名','曲名','アーティスト名')]
result, missing = {}, []
if dur_cols:
    dc = dur_cols[-1]
    for _, row in df.iterrows():
        val = str(row[dc]).strip()
        if not val or val in ('nan','None',''):
            missing.append(str(row['チーム名'])); continue
        try:
            if ':' in val:
                parts = val.split(':')
                secs = int(parts[-2]) * 60 + int(parts[-1])
            else:
                v = int(float(val))
                secs = v // 1000 if v > 10000 else v
            result[str(row['チーム名'])] = secs
        except:
            missing.append(str(row['チーム名']))
if missing:
    print('MISSING:', missing, file=sys.stderr)
print(json.dumps(result, ensure_ascii=False))
"
```

取得できなかったチームがあれば、ユーザーに再生時間（分:秒）を手動で確認し、秒数に変換して JSON に追加する。

### Step 4: スクリプト実行

```bash
PYTHONIOENCODING=utf-8 python3 "<SKILL_DIR>/scripts/generate_setlist.py" \
  --input "<入力ファイルパス>" \
  --start-time "<開始時刻 例:13:00>" \
  --output "<出力ファイルパス 例:setlist_YYYYMMDD.xlsx>" \
  --skip-spotify \
  --duration-overrides '<Step 3-3 で生成した JSON>'
```

列名が違う場合は追加オプションを付ける:
```bash
  --col-team "Team"       # チーム名列のヘッダー
  --col-song "Song"       # 曲名列のヘッダー
  --col-artist "Artist"   # アーティスト名列のヘッダー
  --col-shared "Shared"   # 掛け持ちチーム列のヘッダー
  --col-pref "Pref"       # 希望時間帯列のヘッダー
```

ブロック数を固定したい場合:
```bash
  --n-blocks 4  # デフォルト: チーム数15以下→4、16以上→5
```

### Step 5: 結果報告

- 生成されたExcelファイルのパスをユーザーに伝える
- スクリプトが出力するサマリー（チーム数・ブロック数・終了時刻・警告）を表示する
- ⚠ 警告がある場合（掛け持ち制約の違反・再生時間未取得の曲）は目立つように伝える

---

## スケジューリング制約（優先順位順）

1. **掛け持ちチーム間隔（最重要）**: 掛け持ちメンバーのいるチーム同士は最低 **1時間** の間隔
   - 物理的に満たせない場合: 複数ペアの間隔をできるだけ均等化（例: A-B 10分・B-C 60分より A-B 35分・B-C 35分を優先）し ⚠ 警告を出力
2. **ブロック均等割り**: 全チームを 4〜5 ブロックに均等分割（端数は先頭ブロックへ）
3. **ブロック間休憩**: 原則15分、次ブロックは **15分刻み（00/15/30/45分）** のキリの良い時刻に開始
4. **チーム間インターバル**: チーム間は **30秒** の入退場インターバル
5. **希望時間帯（ソフト制約）**: 前半・後半の希望を可能な範囲で反映

---

## 出力 Excel の構成

### Sheet1: セットリスト

- **Row 1**: メインヘッダー（濃紺）
- **Block 1**: ヘッダーなしで直接チーム行が始まる
- **Block 2〜**: 先頭に中青のヘッダー行を挿入してブロックの区切りを示す
- 掛け持ち制約を違反しているチームは**黄色背景**で強調

列構成:
| 出演順 | ブロック内番号 | 出演開始時間 | チーム名 | 曲名 | アーティスト名 | 出演時間 | 入りはけ時間 |
|---|---|---|---|---|---|---|---|
| 1 | 1-1 | 13:00:00 | TEAM_A | ... | ... | 3:45 | 0:30 |
| 2 | 1-2 | 13:04:15 | TEAM_B | ... | ... | 4:12 | 0:30 |
| 出演順 | ブロック内番号 | 出演開始時間 | チーム名 | 曲名 | アーティスト名 | 出演時間 | 入りはけ時間 |
| 3 | 2-1 | 13:30:00 | TEAM_C | ... | ... | 3:20 | 0:30 |

### Sheet2: 掛け持ちチェック表（掛け持ちペアがある場合のみ）

| チームA | チームB | 実際の間隔 | 必要間隔 | 判定 |
|---|---|---|---|---|
| TEAM_A | TEAM_C | 1:15:00 | 1時間以上 | OK |
| TEAM_B | TEAM_D | 0:45:00 | 1時間以上 | 要確認 |

---

## エラー処理

| 状況 | 対応 |
|---|---|
| Spotify認証情報（`~/.claude/spotify_credentials.json`）がない | `references/spotify-setup.md` を案内 |
| Spotifyで曲が見つからない | スクリプトが列挙するので、ユーザーに再生時間（分:秒）を手動入力してもらい、`--duration-overrides '{"チーム名": 秒数}'` で再実行 |
| 掛け持ち制約を満たせない | 最善案を採用し ⚠ 警告をExcelに付記（黄色背景） |
| 入力ファイルの列名が違う | `--col-*` オプションで列名を指定する |

---

## 参照

- `scripts/generate_setlist.py` — メインスクリプト（読み込み・Spotify取得・スケジューリング・Excel出力を一括処理）
- `references/spotify-setup.md` — Spotify API 認証情報の取得・設定手順
