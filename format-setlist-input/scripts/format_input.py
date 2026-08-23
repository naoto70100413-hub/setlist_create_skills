#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
dance-setlist-formatter: 申込CSVを dance-setlist 入力Excelに整形するスクリプト
"""
import os, sys
os.environ.setdefault('PYTHONIOENCODING', 'utf-8')

import argparse
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ── Helpers ──────────────────────────────────────────────────────────────────

def read_csv_auto(path):
    """エンコーディングを自動判別して CSV を読み込む。"""
    for enc in ("utf-8-sig", "utf-8", "shift_jis", "cp932"):
        try:
            df = pd.read_csv(path, encoding=enc, dtype=str).fillna("")
            return df.apply(lambda col: col.str.strip() if col.dtype == object else col)
        except (UnicodeDecodeError, Exception):
            continue
    raise ValueError("CSVファイルの読み込みに失敗しました: {}".format(path))


def find_col(df, candidates):
    """候補列名リストから最初にマッチした列名を返す。なければ None。"""
    for c in candidates:
        if c in df.columns:
            return c
    # 部分一致フォールバック
    for c in candidates:
        matches = [col for col in df.columns if c in col]
        if matches:
            return matches[0]
    return None


def dedup_by_latest_timestamp(df, id_col, ts_col, label):
    """受付番号が重複する行がある場合、タイムスタンプが最新の行を残す。
    タイムスタンプが無い/解析できない行は最も古いものとして扱う
    （同一受付番号内で全て解析不能な場合は、元のファイル内で最後に登場した行を残す）。
    """
    if ts_col is None or ts_col not in df.columns:
        return df
    dup_ids = df.loc[df.duplicated(subset=id_col, keep=False), id_col].unique()
    if len(dup_ids) > 0:
        print("[{}] 受付番号の重複を検出: {} 件 → タイムスタンプが新しい行を採用".format(label, len(dup_ids)))
    df = df.copy()
    df["_ts_parsed"] = pd.to_datetime(df[ts_col], errors="coerce")
    # タイムスタンプ未解析(NaT)の行は最も古い扱いとし、先頭に寄せる。
    # 全て未解析の同一受付番号内では、安定ソートにより元ファイルで最後に登場した行が残る。
    df = df.sort_values("_ts_parsed", kind="stable", na_position="first")
    df = df.drop_duplicates(subset=id_col, keep="last")
    return df.drop(columns=["_ts_parsed"]).sort_index()


def merge_shared_teams(row, shared_cols):
    """掛け持ち列群の値を収集してカンマ区切りの1文字列に統合する。"""
    teams = []
    for col in shared_cols:
        val = str(row.get(col, "")).strip()
        if not val:
            continue
        # オーバーフロー列はすでにカンマ・読点・スペース区切りの場合があるので分割
        for sep in (",", "、", "・", "　", " "):
            if sep in val:
                parts = [p.strip() for p in val.split(sep) if p.strip()]
                teams.extend(parts)
                break
        else:
            teams.append(val)
    # 重複除去（順序保持）
    seen = set()
    unique = []
    for t in teams:
        if t and t not in seen:
            seen.add(t)
            unique.append(t)
    return ", ".join(unique)


# ── Excel Output ─────────────────────────────────────────────────────────────

HDR_FILL = PatternFill("solid", start_color="1F3864")
HDR_FONT = Font(color="FFFFFF", bold=True, name="Arial", size=11)
NOTE_FILL = PatternFill("solid", start_color="E8F4FD")
NOTE_FONT = Font(color="1A5276", italic=True, name="Arial", size=9)
ALT_FILL  = PatternFill("solid", start_color="F5F5F5")
THIN = Side(style="thin", color="BBBBBB")

OUTPUT_COLS = ["チーム名", "曲名", "アーティスト名", "掛け持ちチーム", "希望時間帯"]
COL_WIDTHS  = [18, 24, 20, 32, 14]
NOTE_TEXTS  = [
    "※ 必須",
    "※ 必須（Spotify検索）",
    "※ 必須（Spotify検索）",
    "掛け持ちメンバーがいる場合\n複数はカンマ区切り",
    "前半 / 後半 / 指定なし",
]


def bdr():
    return Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def write_excel(output_path, result_df, missing_ids):
    wb = Workbook()
    ws = wb.active
    ws.title = "出演チームリスト"

    # Header row
    for col, h in enumerate(OUTPUT_COLS, 1):
        c = ws.cell(1, col, h)
        c.fill = HDR_FILL
        c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = bdr()
    ws.row_dimensions[1].height = 20

    # Data rows
    for i, row in enumerate(result_df.itertuples(index=False), 2):
        values = [row.チーム名, row.曲名, row.アーティスト名, row.掛け持ちチーム, row.希望時間帯]
        for col, val in enumerate(values, 1):
            c = ws.cell(i, col, val)
            if i % 2 == 0:
                c.fill = ALT_FILL
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(
                horizontal="left" if col in (1, 2, 3, 4) else "center",
                vertical="center"
            )
            c.border = bdr()
        ws.row_dimensions[i].height = 16

    # Notes row
    note_row = ws.max_row + 2
    for col, note in enumerate(NOTE_TEXTS, 1):
        c = ws.cell(note_row, col, note)
        c.fill = NOTE_FILL
        c.font = NOTE_FONT
        c.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)
        c.border = bdr()
    ws.row_dimensions[note_row].height = 48

    # Warning sheet for unmatched IDs
    if missing_ids:
        ws2 = wb.create_sheet("結合エラー")
        ws2.cell(1, 1, "結合できなかった受付番号")
        ws2.cell(1, 1).font = Font(bold=True, name="Arial")
        for i, mid in enumerate(missing_ids, 2):
            ws2.cell(i, 1, mid)
        ws2.column_dimensions["A"].width = 30

    # Column widths (main sheet)
    for i, w in enumerate(COL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    wb.save(output_path)


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    p = argparse.ArgumentParser(description="申込CSV → dance-setlist 入力Excel 整形スクリプト")
    p.add_argument("--songs",  required=True,  help="曲情報CSVのパス（受付番号・曲名・アーティスト名）")
    p.add_argument("--teams",  required=True,  help="チーム情報CSVのパス（受付番号・チーム名・掛け持ち列・希望時間帯）")
    p.add_argument("--output", default="",     help="出力Excelのパス（省略時は自動命名）")
    # Column name overrides
    p.add_argument("--col-id",            default="受付番号")
    p.add_argument("--col-song",          default="曲名")
    p.add_argument("--col-artist",        default="アーティスト名")
    p.add_argument("--col-team",          default="チーム名")
    p.add_argument("--col-pref",          default="希望時間帯")
    p.add_argument("--col-shared-prefix", default="掛け持ち",
                   help="掛け持ち列のプレフィックス（デフォルト: 掛け持ち）")
    p.add_argument("--col-timestamp", default="タイムスタンプ",
                   help="タイムスタンプ列の名前（デフォルト: タイムスタンプ）")
    args = p.parse_args()

    # --- Load ---
    df_songs = read_csv_auto(args.songs)
    df_teams = read_csv_auto(args.teams)

    # --- Detect key columns ---
    id_col_s = find_col(df_songs, [args.col_id, "受付番号", "受付No", "No", "番号", "ID"])
    id_col_t = find_col(df_teams, [args.col_id, "受付番号", "受付No", "No", "番号", "ID"])
    song_col   = find_col(df_songs, [args.col_song,   "曲名", "楽曲名", "曲"])
    artist_col = find_col(df_songs, [args.col_artist, "アーティスト名", "アーティスト", "歌手"])
    team_col   = find_col(df_teams, [args.col_team,   "チーム名", "チーム", "団体名"])
    pref_col   = find_col(df_teams, [args.col_pref,   "希望時間帯", "希望", "時間帯"])
    ts_col_s   = find_col(df_songs, [args.col_timestamp, "タイムスタンプ", "Timestamp"])
    ts_col_t   = find_col(df_teams, [args.col_timestamp, "タイムスタンプ", "Timestamp"])

    for name, val in [("受付番号(曲情報)", id_col_s), ("受付番号(チーム情報)", id_col_t),
                       ("曲名", song_col), ("アーティスト名", artist_col), ("チーム名", team_col)]:
        if val is None:
            sys.exit("[ERROR] 列が見つかりません: {}".format(name))

    # --- 受付番号の重複行はタイムスタンプが新しい方を採用 ---
    df_songs = dedup_by_latest_timestamp(df_songs, id_col_s, ts_col_s, "曲情報")
    df_teams = dedup_by_latest_timestamp(df_teams, id_col_t, ts_col_t, "チーム情報")

    # --- Detect shared-team columns ---
    shared_cols = [col for col in df_teams.columns if args.col_shared_prefix in col]
    if shared_cols:
        print("掛け持ち列を検出: {}".format(shared_cols))
    else:
        print("掛け持ち列なし（掛け持ちチーム列は空欄で出力します）")

    # --- Rename to canonical names ---
    df_songs = df_songs.rename(columns={id_col_s: "__id__", song_col: "曲名", artist_col: "アーティスト名"})
    df_teams = df_teams.rename(columns={id_col_t: "__id__", team_col: "チーム名"})
    if pref_col:
        df_teams = df_teams.rename(columns={pref_col: "希望時間帯"})
    else:
        df_teams["希望時間帯"] = ""

    # --- Merge shared team columns ---
    df_teams["掛け持ちチーム"] = df_teams.apply(
        lambda row: merge_shared_teams(row, shared_cols), axis=1
    )

    # --- Join on __id__ ---
    songs_slim = df_songs[["__id__", "曲名", "アーティスト名"]]
    teams_slim = df_teams[["__id__", "チーム名", "掛け持ちチーム", "希望時間帯"]]

    merged = pd.merge(songs_slim, teams_slim, on="__id__", how="inner")

    # Detect unmatched
    all_ids = set(df_songs["__id__"]) | set(df_teams["__id__"])
    matched_ids = set(merged["__id__"])
    missing_ids = sorted(all_ids - matched_ids)

    # Final column order
    result = merged[["チーム名", "曲名", "アーティスト名", "掛け持ちチーム", "希望時間帯"]].copy()
    result = result.reset_index(drop=True)

    # --- Output ---
    if args.output:
        out_path = args.output
    else:
        out_path = "setlist_input_{}.xlsx".format(datetime.now().strftime("%Y%m%d"))

    write_excel(out_path, result, missing_ids)

    print("FORMATTER COMPLETE: {}".format(out_path))
    print("  {} チームを出力しました".format(len(result)))
    if missing_ids:
        print("WARNING: 結合できなかった受付番号 {} 件 → 「結合エラー」シートを確認してください".format(len(missing_ids)))
        for mid in missing_ids[:5]:
            print("  - {}".format(mid))
        if len(missing_ids) > 5:
            print("  ... 他 {} 件".format(len(missing_ids) - 5))
    print("次のステップ: このファイルを dance-setlist スキルの入力として使用できます")


if __name__ == "__main__":
    main()
