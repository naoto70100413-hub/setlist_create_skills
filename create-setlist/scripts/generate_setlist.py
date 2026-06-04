#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dance-setlist: ダンスイベントセットリスト生成スクリプト"""
import os, sys
os.environ.setdefault('PYTHONIOENCODING', 'utf-8')
import argparse, json, math, time
from datetime import datetime, timedelta
from pathlib import Path
import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Spotify API ──────────────────────────────────────────────────────────────

def load_spotify_credentials():
    p = Path.home() / ".claude" / "spotify_credentials.json"
    if not p.exists(): return None, None
    c = json.load(open(p))
    return c.get("client_id"), c.get("client_secret")

def get_spotify_token(client_id, client_secret):
    r = requests.post("https://accounts.spotify.com/api/token",
        data={"grant_type": "client_credentials"}, auth=(client_id, client_secret), timeout=10)
    r.raise_for_status()
    return r.json()["access_token"]

def search_track_duration(token, track_name, artist_name, retries=3):
    q = "track:{} artist:{}".format(track_name, artist_name)
    for _ in range(retries):
        try:
            r = requests.get("https://api.spotify.com/v1/search",
                headers={"Authorization": "Bearer {}".format(token)},
                params={"q": q, "type": "track", "limit": 1}, timeout=10)
            if r.status_code == 429: time.sleep(1); continue
            r.raise_for_status()
            items = r.json().get("tracks", {}).get("items", [])
            return items[0]["duration_ms"] // 1000 if items else None
        except Exception: time.sleep(1)
    return None

# ── Scheduling ───────────────────────────────────────────────────────────────

def parse_shared(val):
    if not val or (isinstance(val, float) and math.isnan(val)): return []
    return [t.strip() for t in str(val).split(",") if t.strip()]

def pref_half(val):
    if not val or (isinstance(val, float) and math.isnan(val)): return None
    return "first" if "前半" in str(val) else ("second" if "後半" in str(val) else None)

def build_conflicts(teams):
    names = {t["name"] for t in teams}
    pairs = set()
    for t in teams:
        for o in t["shared_teams"]:
            if o in names and o != t["name"]:
                pairs.add(frozenset([t["name"], o]))
    return pairs

def elapsed_gap(ordered, a, b, durations):
    ia, ib = ordered.index(a), ordered.index(b)
    if ia > ib: ia, ib = ib, ia
    return sum(durations.get(ordered[i], 0) + 30 for i in range(ia, ib))

def score_order(order, durations, conflicts, prefs, n):
    s = 0
    for pair in conflicts:
        a, b = list(pair)
        if a in order and b in order:
            g = elapsed_gap(order, a, b, durations)
            # 凹関数(√)でギャップを評価する。
            # 凹関数はジェンセンの不等式により、同じ合計ギャップでも
            # 均等分散のほうが常に高スコアになる。
            # 例: A-B=10分・B-C=60分 → スコア合計 281.7
            #     A-B=35分・B-C=35分 → スコア合計 305.6 ← 均等ケースが勝つ
            s += (min(g, 3600) / 3600) ** 0.5 * 200
    for i, name in enumerate(order):
        p = prefs.get(name)
        if p == "first" and i < n / 2: s += 5
        elif p == "second" and i >= n / 2: s += 5
    return s

def greedy_schedule(teams, durations, conflicts):
    names = [t["name"] for t in teams]
    prefs = {t["name"]: pref_half(t.get("preferred_time")) for t in teams}
    cc = {n: 0 for n in names}
    for pair in conflicts:
        for n in pair: cc[n] += 1
    remaining = sorted(names, key=lambda x: -cc[x])
    ordered = []
    for name in remaining:
        best_pos, best_score = 0, -float("inf")
        for pos in range(len(ordered) + 1):
            cand = ordered[:pos] + [name] + ordered[pos:]
            sc = score_order(cand, durations, conflicts, prefs, len(names))
            if sc > best_score: best_score, best_pos = sc, pos
        ordered.insert(best_pos, name)
    return ordered

def assign_blocks(ordered, n_blocks):
    n = len(ordered)
    base, extra = divmod(n, n_blocks)
    blocks, idx = [], 0
    for i in range(n_blocks):
        size = base + (1 if i < extra else 0)
        blocks.append(ordered[idx:idx+size])
        idx += size
    return blocks

def round_up_15(dt):
    """次の15分刻み (00/15/30/45分) に切り上げる。"""
    if dt.second > 0:
        dt = dt.replace(second=0) + timedelta(minutes=1)
    rem = dt.minute % 15
    return dt if rem == 0 else dt + timedelta(minutes=(15 - rem))

def calc_schedule(blocks, durations, start_time_str):
    current = datetime.strptime(start_time_str, "%H:%M")
    schedule = []
    for block_idx, block in enumerate(blocks):
        if block_idx > 0:
            break_start = current
            break_end = round_up_15(break_start + timedelta(minutes=15))
            if break_end < break_start + timedelta(minutes=15):
                break_end += timedelta(minutes=15)
            schedule.append({"type":"break","block":block_idx+1,
                "start":break_start,"end":break_end,
                "duration_sec":int((break_end-break_start).total_seconds())})
            current = break_end
        for name in block:
            dur = durations.get(name, 0)
            end = current + timedelta(seconds=dur)
            schedule.append({"type":"team","block":block_idx+1,"name":name,
                "start":current,"end":end,"duration_sec":dur})
            current = end + timedelta(seconds=30)
    return schedule

def real_gap(schedule, a, b):
    times = {e["name"]:(e["start"],e["end"]) for e in schedule if e["type"]=="team"}
    if a not in times or b not in times: return 0
    sa, ea = times[a]; sb, eb = times[b]
    return int((sb - ea).total_seconds()) if sa <= sb else int((sa - eb).total_seconds())

def check_violations(ordered, durations, conflicts, schedule=None):
    out = []
    for pair in conflicts:
        a, b = list(pair)
        if a in ordered and b in ordered:
            g = real_gap(schedule, a, b) if schedule else elapsed_gap(ordered, a, b, durations)
            if g < 3600:
                out.append({"team_a":a,"team_b":b,"gap_sec":g,"gap_str":str(timedelta(seconds=g))})
    return out

# ── Excel Output ─────────────────────────────────────────────────────────────

HDR_FILL    = PatternFill("solid", start_color="1F3864")
HDR_FONT    = Font(color="FFFFFF", bold=True, name="Arial")
BLKHDR_FILL = PatternFill("solid", start_color="2E5FA3")
WARN_FILL   = PatternFill("solid", start_color="FFF2CC")
ALT_FILL    = PatternFill("solid", start_color="F3F3F3")
THIN        = Side(style="thin", color="CCCCCC")

BASE_HEADERS    = ["出演順","ブロック内番号","出演開始時間","チーム名","曲名","アーティスト名","再生時間"]
BASE_COL_WIDTHS = [9, 13, 16, 18, 22, 18, 11]

def bdr():
    return Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def fmt_time(dt): return dt.strftime("%H:%M:%S")

def fmt_dur(sec):
    m, s = divmod(int(sec), 60)
    return "{}:{:02d}".format(m, s)

def apply_hdr(ws, row_idx, fill, font, headers):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row_idx, col)
        c.value = h; c.fill = fill; c.font = font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = bdr()

def write_setlist(ws, schedule, df, durations, violations):
    team_lookup = df.set_index("name").to_dict("index")
    viol_teams = {v["team_a"] for v in violations} | {v["team_b"] for v in violations}

    # 掛け持ちチームの最大数を算出してヘッダーを動的に構築
    max_shared = max((len(info.get("shared_teams", [])) for info in team_lookup.values()), default=0)
    if max_shared == 0:
        shared_headers = ["掛け持ちチーム"]
        shared_widths  = [22]
    else:
        shared_headers = ["掛け持ちチーム{}".format(i+1) for i in range(max_shared)]
        shared_widths  = [18] * max_shared
    headers    = BASE_HEADERS + shared_headers
    col_widths = BASE_COL_WIDTHS + shared_widths

    # Row 1: メインヘッダー
    apply_hdr(ws, 1, HDR_FILL, HDR_FONT, headers)

    team_num = 0
    current_block = None
    block_team_num = {}

    for entry in schedule:
        if entry["type"] == "break":
            continue  # 休憩行は出力しない

        block_idx = entry["block"]
        if block_idx != current_block:
            current_block = block_idx
            block_team_num[block_idx] = 0
            # Block 2 以降の先頭にブロックヘッダー行を挿入
            if block_idx >= 2:
                apply_hdr(ws, ws.max_row + 1, BLKHDR_FILL,
                          Font(color="FFFFFF", bold=True, name="Arial"), headers)

        team_num += 1
        block_team_num[block_idx] += 1
        name = entry["name"]
        info = team_lookup.get(name, {})
        block_label = "{}-{}".format(block_idx, block_team_num[block_idx])

        shared_list = info.get("shared_teams", [])
        # 列数に合わせてパディング（掛け持ちなしなら空文字）
        if max_shared == 0:
            shared_cells = [""]
        else:
            shared_cells = shared_list + [""] * (max_shared - len(shared_list))

        row = [team_num, block_label, fmt_time(entry["start"]),
               name, info.get("song",""), info.get("artist",""), fmt_dur(entry["duration_sec"])] + shared_cells
        ws.append(row)
        ri = ws.max_row
        fill = WARN_FILL if name in viol_teams else (ALT_FILL if team_num % 2 == 0 else None)
        for col in range(1, len(headers) + 1):
            c = ws.cell(ri, col)
            if fill: c.fill = fill
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(horizontal="left" if col==4 else "center", vertical="center")
            c.border = bdr()

    for i, w in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 18

def write_check(ws, ordered, durations, conflicts, schedule):
    hdrs = ["チームA","チームB","実際の間隔","必要間隔","判定"]
    ws.append(hdrs)
    for col in range(1, len(hdrs)+1):
        c = ws.cell(1, col)
        c.fill = HDR_FILL; c.font = HDR_FONT
        c.alignment = Alignment(horizontal="center"); c.border = bdr()
    for pair in conflicts:
        a, b = list(pair)
        if a in ordered and b in ordered:
            g = real_gap(schedule, a, b)
            ok = g >= 3600
            row = [a, b, str(timedelta(seconds=g)), "1時間以上", "OK" if ok else "要確認"]
            ws.append(row)
            ri = ws.max_row
            for col in range(1, len(hdrs)+1):
                c = ws.cell(ri, col)
                if not ok: c.fill = WARN_FILL
                c.font = Font(name="Arial", size=10)
                c.alignment = Alignment(horizontal="center"); c.border = bdr()
    for i in range(1, len(hdrs)+1):
        ws.column_dimensions[get_column_letter(i)].width = 16

# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--input", required=True)
    p.add_argument("--start-time", required=True)
    p.add_argument("--output", required=True)
    p.add_argument("--n-blocks", type=int, default=0)
    p.add_argument("--col-team",   default="チーム名")
    p.add_argument("--col-song",   default="曲名")
    p.add_argument("--col-artist", default="アーティスト名")
    p.add_argument("--col-shared", default="掛け持ちチーム")
    p.add_argument("--col-pref",   default="希望時間帯")
    p.add_argument("--duration-overrides", default="")
    args = p.parse_args()

    path = Path(args.input)
    df = pd.read_excel(path) if path.suffix in (".xlsx",".xls") else pd.read_csv(path)
    col_map = {args.col_team:"name", args.col_song:"song", args.col_artist:"artist",
               args.col_shared:"shared_teams_raw", args.col_pref:"preferred_time"}
    df = df.rename(columns={k:v for k,v in col_map.items() if k in df.columns})
    if "shared_teams_raw" in df.columns:
        df["shared_teams"] = df["shared_teams_raw"].apply(parse_shared)
    else:
        df["shared_teams"] = [[] for _ in range(len(df))]
    teams = df.to_dict("records")

    overrides = json.loads(args.duration_overrides) if args.duration_overrides else {}
    client_id, client_secret = load_spotify_credentials()
    durations = {}; missing = []

    if client_id and client_secret:
        try:
            token = get_spotify_token(client_id, client_secret)
            for t in teams:
                if t["name"] in overrides:
                    durations[t["name"]] = int(overrides[t["name"]]); continue
                sec = search_track_duration(token, t["song"], t["artist"])
                if sec: durations[t["name"]] = sec
                else: missing.append(t); durations[t["name"]] = overrides.get(t["name"],0)
        except Exception as e:
            sys.stderr.write("[Spotify Error] {}\n".format(e))
            missing = [t for t in teams if t["name"] not in overrides]
            for t in teams: durations[t["name"]] = overrides.get(t["name"],0)
    else:
        sys.stderr.write("Spotify認証情報なし\n")
        missing = [t for t in teams if t["name"] not in overrides]
        for t in teams: durations[t["name"]] = overrides.get(t["name"],0)

    if missing:
        print("[手動入力が必要な曲]")
        for t in missing:
            print("  {}: {} / {}".format(t["name"], t["song"], t["artist"]))

    conflicts = build_conflicts(teams)
    n_blocks = args.n_blocks if args.n_blocks > 0 else (5 if len(teams) > 15 else 4)
    ordered = greedy_schedule(teams, durations, conflicts)
    blocks = assign_blocks(ordered, n_blocks)
    schedule = calc_schedule(blocks, durations, args.start_time)
    violations = check_violations(ordered, durations, conflicts, schedule=schedule)

    wb = Workbook()
    ws1 = wb.active; ws1.title = "セットリスト"
    write_setlist(ws1, schedule, df, durations, violations)
    if conflicts:
        ws2 = wb.create_sheet("掛け持ちチェック")
        write_check(ws2, ordered, durations, conflicts, schedule)
    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.output)

    team_entries = [e for e in schedule if e["type"]=="team"]
    print("SETLIST COMPLETE: {}".format(args.output))
    print("  Teams:{} Blocks:{}".format(len(teams), n_blocks))
    if team_entries:
        print("  End time: {}".format(fmt_time(team_entries[-1]["end"])))
    if violations:
        print("WARNING: {} conflict(s):".format(len(violations)))
        for v in violations:
            print("  {} <-> {}: {}".format(v["team_a"], v["team_b"], v["gap_str"]))
    if missing:
        print("WARNING: {} song(s) not found".format(len(missing)))

if __name__ == "__main__":
    main()
