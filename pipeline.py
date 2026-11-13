"""
Атакующая статистика футболистов топ-5 лиг (2015/16 – 2024/25).

Конвейер:  сырые данные (FBref .rds + Understat .csv)
        -> SQLite (football.db)
        -> SQL-запросы из queries.sql
        -> графики PNG (charts/)  +  Excel-отчёт (football_attack_stats.xlsx)

Запуск:  python pipeline.py
"""
from __future__ import annotations

import glob
import os
import re
import sqlite3

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import pyreadr
from adjustText import adjust_text
from matplotlib.colors import LinearSegmentedColormap
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.join(BASE, "data", "raw")
DB = os.path.join(BASE, "football.db")
CHARTS = os.path.join(BASE, "charts")
XLSX = os.path.join(BASE, "football_attack_stats.xlsx")
os.makedirs(CHARTS, exist_ok=True)

LEAGUE_RU = {
    "Premier League": "АПЛ", "La Liga": "Ла Лига", "Bundesliga": "Бундеслига",
    "Serie A": "Серия А", "Ligue 1": "Лига 1",
}
US_LEAGUE = {"epl": "АПЛ", "la_liga": "Ла Лига", "bundesliga": "Бундеслига",
             "serie_a": "Серия А", "ligue_1": "Лига 1", "rfpl": "РПЛ"}

# ---------- оформление (проверенная категориальная палитра, светлая тема) ----------
SURFACE, INK, INK2, GRID = "#fcfcfb", "#0b0b0b", "#52514e", "#e4e3de"
SERIES = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4"]
LEAGUE_COLOR = dict(zip(["АПЛ", "Ла Лига", "Бундеслига", "Серия А", "Лига 1"], SERIES))
SEQ = ["#cde2fb", "#9ec5f4", "#6da7ec", "#3987e5", "#256abf", "#184f95", "#0d366b"]

plt.rcParams.update({
    "figure.facecolor": SURFACE, "axes.facecolor": SURFACE, "savefig.facecolor": SURFACE,
    "font.family": "DejaVu Sans", "font.size": 10, "text.color": INK,
    "axes.labelcolor": INK2, "xtick.color": INK2, "ytick.color": INK2,
    "axes.edgecolor": GRID, "axes.spines.top": False, "axes.spines.right": False,
    "axes.titlesize": 14, "axes.titleweight": "bold", "axes.titlelocation": "left",
    "axes.titlepad": 14, "grid.color": GRID, "grid.linewidth": 0.8,
})
SOURCE_FB = "Источник: FBref (через архив worldfootballR_data). Нападающие = позиция содержит FW."
SOURCE_US = "Источник: Understat (датасет douglasbc/scraping-understat-dataset)."


# =============================== 1. Загрузка в SQLite ===============================
def load_fbref() -> pd.DataFrame:
    std = list(pyreadr.read_r(os.path.join(RAW, "big5_player_standard.rds")).values())[0]
    sht = list(pyreadr.read_r(os.path.join(RAW, "big5_player_shooting.rds")).values())[0]
    key = ["Season_End_Year", "Squad", "Comp", "Url"]
    df = std.merge(sht[key + ["Sh_Standard", "SoT_Standard", "Dist_Standard", "FK_Standard"]],
                   on=key, how="left")
    df = df[df.Season_End_Year.between(2016, 2025)].copy()  # 2025/26 в архиве неполный
    out = pd.DataFrame({
        "season_end": df.Season_End_Year.astype(int),
        "season": df.Season_End_Year.map(lambda y: f"{int(y) - 1}/{str(int(y))[2:]}"),
        "league": df.Comp.map(LEAGUE_RU),
        "squad": df.Squad, "player": df.Player,
        "player_id": df.Url.str.extract(r"/players/([0-9a-f]{8})/")[0],
        "nation": df.Nation, "pos": df.Pos, "age": pd.to_numeric(df.Age, errors="coerce"),
        "mp": df.MP_Playing, "starts": df.Starts_Playing, "minutes": df.Min_Playing,
        "goals": df.Gls, "assists": df.Ast, "npg": df.G_minus_PK, "pk": df.PK, "pkatt": df.PKatt,
        "xg": df.xG_Expected, "npxg": df.npxG_Expected, "xag": df.xAG_Expected,
        "prgc": df.PrgC_Progression, "prgp": df.PrgP_Progression, "prgr": df.PrgR_Progression,
        "shots": df.Sh_Standard, "sot": df.SoT_Standard, "avg_shot_dist": df.Dist_Standard,
        "fk_goals": df.FK_Standard, "fbref_url": df.Url,
    })
    return out[out.minutes.fillna(0) > 0]


def load_understat() -> pd.DataFrame:
    frames = []
    for f in sorted(glob.glob(os.path.join(RAW, "understat", "players_*.csv"))):
        m = re.search(r"players_(.+)_(\d\d)-(\d\d)\.csv$", os.path.basename(f))
        lg, y1, y2 = m.group(1), int(m.group(2)), int(m.group(3))
        if not 15 <= y1 <= 20:            # берём только полные сезоны 2015/16–2020/21
            continue
        d = pd.read_csv(f)
        d.insert(0, "season", f"20{y1}/{y2:02d}")
        d.insert(1, "season_end", 2000 + y2)
        d.insert(2, "league", US_LEAGUE[lg])
        frames.append(d)
    us = pd.concat(frames, ignore_index=True)
    return us.rename(columns={"id": "understat_id", "player_name": "player", "time": "minutes",
                              "xG": "xg", "xA": "xa", "npxG": "npxg", "xGChain": "xgchain",
                              "xGBuildup": "xgbuildup", "team_title": "team"})


def build_db() -> sqlite3.Connection:
    if os.path.exists(DB):
        os.remove(DB)
    con = sqlite3.connect(DB)
    fb, us = load_fbref(), load_understat()
    fb.to_sql("fbref_player_season", con, index=False)
    us.to_sql("understat_player_season", con, index=False)
    con.execute("CREATE INDEX ix_fb ON fbref_player_season(player_id, season_end)")
    con.execute("CREATE INDEX ix_us ON understat_player_season(understat_id, season_end)")
    print(f"SQLite: fbref {len(fb):,} строк, understat {len(us):,} строк")
    return con


def read_queries() -> dict[str, str]:
    text = open(os.path.join(BASE, "queries.sql"), encoding="utf-8").read()
    parts = re.split(r"^-- name: (\w+)\s*$", text, flags=re.M)
    return {parts[i]: parts[i + 1].strip() for i in range(1, len(parts), 2)}


# =============================== 2. Графики ===============================
def save(fig, name, source=SOURCE_FB, y=0.01):
    fig.text(0.01, y, source, fontsize=8, color=INK2, ha="left", va="bottom")
    path = os.path.join(CHARTS, name)
    fig.savefig(path, dpi=150, bbox_inches="tight")
    plt.close(fig)
    return path


def hbar_stacked(df, label, parts, colors, legend, title, subtitle, fname, total_col=None, fmt="{:.0f}"):
    df = df.iloc[::-1]
    fig, ax = plt.subplots(figsize=(10, 0.36 * len(df) + 1.8))
    left = np.zeros(len(df))
    for col, c, lg in zip(parts, colors, legend):
        ax.barh(df[label], df[col], left=left, color=c, height=0.62, label=lg,
                edgecolor=SURFACE, linewidth=1.5)
        left += df[col].values
    tot = df[total_col].values if total_col else left
    for y, v in enumerate(tot):
        ax.text(left[y] + left.max() * 0.01, y, fmt.format(v), va="center", fontsize=9, color=INK)
    ax.set_xlim(0, left.max() * 1.1)
    ax.grid(axis="x"); ax.set_axisbelow(True); ax.tick_params(axis="y", length=0)
    ax.spines["left"].set_visible(False)
    ax.set_title(title)
    ax.text(0, 1.01, subtitle, transform=ax.transAxes, fontsize=9.5, color=INK2)
    ax.legend(loc="lower right", frameon=False, ncol=len(parts))
    return save(fig, fname)


def league_lines(tr, metric, title, subtitle, fname, fmt="{:.2f}"):
    piv = tr.pivot(index="Сезон", columns="Лига", values=metric).dropna(how="all")
    fig, ax = plt.subplots(figsize=(11, 5.6))
    ends = []
    for lg, c in LEAGUE_COLOR.items():
        s = piv[lg]
        ax.plot(s.index, s.values, color=c, lw=2, marker="o", ms=5, label=lg,
                markeredgecolor=SURFACE, markeredgewidth=1.2)
        ends.append(ax.text(len(s) - 1 + 0.15, s.values[-1], f"{lg} {fmt.format(s.values[-1])}",
                            va="center", fontsize=9, color=INK))
    adjust_text(ends, ax=ax, only_move={"text": "y"}, autoalign=False)
    ax.set_xlim(-0.3, len(piv) + 0.9)
    ax.grid(axis="y"); ax.set_axisbelow(True)
    ax.set_title(title)
    ax.text(0, 1.01, subtitle, transform=ax.transAxes, fontsize=9.5, color=INK2)
    ax.legend(frameon=False, ncol=5, loc="upper left", bbox_to_anchor=(0, -0.07))
    return save(fig, fname, y=-0.04)


def make_charts(R: dict[str, pd.DataFrame]) -> list[str]:
    paths = []

    # 1. Бомбардиры: голы с игры + пенальти
    t = R["top_scorers"].head(15)
    paths.append(hbar_stacked(t, "Игрок", ["Голы без пен.", "Пенальти"], SERIES[:2],
                              ["Голы с игры", "Пенальти"],
                              "Лучшие бомбардиры топ-5 лиг, 2015/16 – 2024/25",
                              "Сумма голов в чемпионатах за 10 сезонов, подпись — всего голов",
                              "01_top_scorers.png", total_col="Голы"))

    # 2. Гол+пас
    t = R["top_ga"].head(15)
    paths.append(hbar_stacked(t, "Игрок", ["Голы", "Передачи"], SERIES[:2], ["Голы", "Голевые передачи"],
                              "Гол + пас: самые результативные нападающие",
                              "Топ-5 лиг, 2015/16 – 2024/25, подпись — сумма гол+пас",
                              "02_top_goals_assists.png"))

    # 3. Лучший бомбардир каждого сезона
    s = R["season_top_scorers"].groupby("Сезон", sort=False).agg(
        Игрок=("Игрок", " / ".join), Клуб=("Клуб", " / ".join), Голы=("Голы", "first")).reset_index()
    s = s.iloc[::-1]
    fig, ax = plt.subplots(figsize=(11, 5.8))
    ax.barh(s["Сезон"], s["Голы"], color=SERIES[0], height=0.6)
    for y, (_, r) in enumerate(s.iterrows()):
        ax.text(r["Голы"] - 0.6, y, f"{r['Голы']:.0f}", va="center", ha="right", color="white",
                fontsize=10, fontweight="bold")
        ax.text(r["Голы"] + 0.5, y, f"{r['Игрок']}  ({r['Клуб']})", va="center", fontsize=9.5)
    ax.set_xlim(0, s["Голы"].max() * 1.75)
    ax.set_xlabel("Голы в чемпионате"); ax.grid(axis="x"); ax.set_axisbelow(True)
    ax.tick_params(axis="y", length=0); ax.spines["left"].set_visible(False)
    ax.set_title("Лучший бомбардир топ-5 лиг в каждом сезоне")
    paths.append(save(fig, "03_season_top_scorers.png"))

    # 4–6. Динамика по лигам
    tr = R["league_trends"]
    paths.append(league_lines(tr, "Голы за 90", "Голы за 90 минут у нападающих по лигам",
                              "Нападающие с ≥ 900 мин в сезоне; сумма голов × 90 / сумма минут",
                              "04_trend_goals_p90.png"))
    paths.append(league_lines(tr[tr.season_end >= 2017], "Реализация, %",
                              "Реализация ударов нападающими, % голов от ударов",
                              "Нападающие с ≥ 900 мин в сезоне, 2016/17 – 2024/25",
                              "05_trend_conversion.png", fmt="{:.1f}%"))
    paths.append(league_lines(tr[tr.season_end >= 2017], "Удары за 90",
                              "Удары за 90 минут у нападающих по лигам",
                              "Нападающие с ≥ 900 мин в сезоне, 2016/17 – 2024/25",
                              "06_trend_shots_p90.png"))

    # 7. Голы без пенальти vs npxG
    fx = R["finishing_vs_xg"]
    fig, ax = plt.subplots(figsize=(10.5, 8))
    ax.scatter(fx["npxG"], fx["Голы без пен."], s=26, color=SERIES[0], alpha=0.55,
               edgecolor=SURFACE, linewidth=0.8, zorder=3)
    lim = max(fx["npxG"].max(), fx["Голы без пен."].max()) * 1.05
    ax.plot([0, lim], [0, lim], color=INK2, lw=1, ls="--", zorder=2)
    ax.text(lim * 0.56, lim * 0.50, "линия: голы = npxG", color=INK2, fontsize=9, ha="left")
    hi = pd.concat([fx.nlargest(8, "Голы - npxG"), fx.nsmallest(4, "Голы - npxG"),
                    fx.nlargest(5, "npxG")]).drop_duplicates("Игрок")
    ax.scatter(hi["npxG"], hi["Голы без пен."], s=40, color=SERIES[1], edgecolor=SURFACE,
               linewidth=1.2, zorder=4)
    texts = [ax.text(r["npxG"], r["Голы без пен."], f"{r['Игрок']} ({r['Голы - npxG']:+.0f})",
                     fontsize=8.5) for _, r in hi.iterrows()]
    adjust_text(texts, ax=ax, arrowprops=dict(arrowstyle="-", color=INK2, lw=0.6))
    ax.set_xlim(0, lim); ax.set_ylim(0, lim)
    ax.set_xlabel("Ожидаемые голы без пенальти (npxG)"); ax.set_ylabel("Голы без пенальти")
    ax.grid(True); ax.set_axisbelow(True)
    ax.set_title("Реализация моментов: голы против ожидаемых голов")
    ax.text(0, 1.01, "Нападающие с ≥ 6000 мин, 2017/18 – 2024/25. Выше линии — забивают больше, чем «положено»",
            transform=ax.transAxes, fontsize=9.5, color=INK2)
    paths.append(save(fig, "07_goals_vs_npxg.png"))

    # 8. npxG + xAG за 90
    t = R["p90_leaders"].head(15)
    paths.append(hbar_stacked(t, "Игрок", ["npxG за 90", "xAG за 90"], SERIES[:2],
                              ["npxG за 90", "xAG за 90"],
                              "Ожидаемый вклад в атаку за 90 минут",
                              "npxG + xAG за 90, нападающие с ≥ 8000 мин, 2017/18 – 2024/25",
                              "08_p90_npxg_xag.png", fmt="{:.2f}"))

    # 9. Тепловая карта голов
    hm = R["heatmap_goals"].pivot(index="Игрок", columns="Сезон", values="Голы")
    hm = hm.loc[hm.sum(axis=1).sort_values(ascending=False).index]
    cmap = LinearSegmentedColormap.from_list("seq", SEQ)
    fig, ax = plt.subplots(figsize=(12, 6.5))
    cmap.set_bad(SURFACE)
    ax.imshow(np.ma.masked_invalid(hm.values.astype(float)), cmap=cmap, aspect="auto", vmin=0)
    vmax = np.nanmax(hm.values)
    for i in range(hm.shape[0]):
        for j in range(hm.shape[1]):
            v = hm.values[i, j]
            ax.text(j, i, "–" if np.isnan(v) else f"{v:.0f}", ha="center", va="center", fontsize=9,
                    color="white" if (not np.isnan(v) and v > vmax * 0.5) else INK)
    ax.set_xticks(range(hm.shape[1]), hm.columns); ax.set_yticks(range(hm.shape[0]), hm.index)
    ax.tick_params(length=0)
    for sp in ax.spines.values():
        sp.set_visible(False)
    ax.set_xticks(np.arange(-.5, hm.shape[1]), minor=True)
    ax.set_yticks(np.arange(-.5, hm.shape[0]), minor=True)
    ax.grid(which="minor", color=SURFACE, linewidth=2); ax.tick_params(which="minor", length=0)
    ax.set_title("Голы по сезонам: 12 лучших бомбардиров периода")
    ax.text(0, 1.01, "«–» — игрок не выступал в топ-5 лигах как нападающий",
            transform=ax.transAxes, fontsize=9.5, color=INK2)
    paths.append(save(fig, "09_heatmap_goals.png"))

    # 10. Understat: созидатели (xA)
    t = R["understat_creators"].head(15).copy()
    t["xA_label"] = t["xA"]
    p = hbar_stacked(t, "Игрок", ["xA"], [SERIES[0]], ["xA (ожидаемые передачи)"],
                     "Лучшие созидатели среди нападающих по xA",
                     "Understat, топ-5 лиг + РПЛ, 2015/16 – 2020/21; подпись — сумма xA",
                     "10_understat_xa.png", fmt="{:.1f}")
    paths.append(p)

    # 11. РПЛ: голы vs xG
    r = R["understat_rpl"].iloc[::-1]
    fig, ax = plt.subplots(figsize=(10, 6.5))
    y = np.arange(len(r))
    ax.barh(y + 0.19, r["Голы"], height=0.36, color=SERIES[0], label="Голы")
    ax.barh(y - 0.19, r["xG"], height=0.36, color=SERIES[1], label="xG")
    for yy, g, x in zip(y, r["Голы"], r["xG"]):
        ax.text(g + 0.5, yy + 0.19, f"{g:.0f}", va="center", fontsize=8.5)
        ax.text(x + 0.5, yy - 0.19, f"{x:.1f}", va="center", fontsize=8.5, color=INK2)
    ax.set_yticks(y, r["Игрок"]); ax.tick_params(axis="y", length=0)
    ax.spines["left"].set_visible(False); ax.grid(axis="x"); ax.set_axisbelow(True)
    ax.legend(frameon=False, loc="lower right")
    ax.set_title("РПЛ: лучшие бомбардиры и их ожидаемые голы (xG)")
    ax.text(0, 1.01, "Understat, 2015/16 – 2020/21, сумма за 6 сезонов", transform=ax.transAxes,
            fontsize=9.5, color=INK2)
    paths.append(save(fig, "11_rpl_goals_xg.png", SOURCE_US))
    return paths


# =============================== 3. Excel ===============================
README = [
    ("Отчёт", "Атакующая статистика нападающих топ-5 европейских лиг, сезоны 2015/16 – 2024/25"),
    ("Источник 1", "FBref (Opta/StatsPerform) — архив worldfootballR_data, таблицы big5_player_standard / shooting"),
    ("Источник 2", "Understat — датасет douglasbc/scraping-understat-dataset (топ-5 лиг + РПЛ, 2015/16 – 2020/21)"),
    ("Кто «нападающий»", "FBref: в позиции есть FW (FW, FW,MF, MF,FW). Understat: в позиции есть F."),
    ("Ограничения", "xG/npxG/xAG и прогрессивные действия в FBref — с 2017/18; удары — с 2016/17. "
                    "Сезон 2025/26 исключён (в архиве только 3–5 туров). Ligue 1 2019/20 прерван (COVID)."),
    ("Термины", "npxG — ожидаемые голы без пенальти; xAG — ожидаемые голевые передачи (FBref); "
                "xA — ожидаемые передачи (Understat); «за 90» — на 90 игровых минут."),
    ("Инструменты", "Python (pandas, matplotlib, openpyxl) → SQLite (football.db, queries.sql) → Excel"),
]
SHEETS = [  # (ключ запроса, имя листа)
    ("league_summary", "Сводка по лигам"),
    ("top_scorers", "Бомбардиры"),
    ("top_ga", "Гол+пас"),
    ("season_top_scorers", "Лучшие по сезонам"),
    ("league_trends", "Динамика по лигам"),
    ("finishing_vs_xg", "Голы vs npxG"),
    ("p90_leaders", "Лидеры за 90"),
    ("progressive_carries", "Прогрессия"),
    ("understat_creators", "Understat созидатели"),
    ("understat_rpl", "РПЛ (Understat)"),
]


def style_sheet(ws):
    head = PatternFill("solid", fgColor="1C5CAB")
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = head
        c.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[1].height = 32
    for i, col in enumerate(ws.columns, 1):
        w = max(len(str(c.value)) if c.value is not None else 0 for c in list(col)[:200])
        ws.column_dimensions[get_column_letter(i)].width = min(max(10, w + 2), 45)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def write_excel(R, con, chart_paths):
    trends = R["league_trends"]
    with pd.ExcelWriter(XLSX, engine="openpyxl") as xw:
        pd.DataFrame(README, columns=["Раздел", "Описание"]).to_excel(xw, sheet_name="О данных", index=False)
        for key, name in SHEETS:
            R[key].drop(columns=["season_end"], errors="ignore").to_excel(xw, sheet_name=name, index=False)
        for metric, name in [("Голы за 90", "Сводная голы за 90"), ("Реализация, %", "Сводная реализация")]:
            trends.pivot(index="Сезон", columns="Лига", values=metric)[list(LEAGUE_COLOR)] \
                  .dropna(how="all").reset_index().to_excel(xw, sheet_name=name, index=False)
        pd.read_sql("SELECT * FROM fb_att ORDER BY season_end, league, goals DESC", con) \
          .drop(columns=["season_end"]).to_excel(xw, sheet_name="Данные FBref (нап.)", index=False)
        pd.read_sql("SELECT * FROM us_att ORDER BY season_end, league, goals DESC", con) \
          .drop(columns=["season_end"]).to_excel(xw, sheet_name="Данные Understat (нап.)", index=False)

    wb = load_workbook(XLSX)
    for ws in wb.worksheets:
        style_sheet(ws)
    wb["О данных"].column_dimensions["B"].width = 120

    # Нативные диаграммы Excel
    for name, title in [("Сводная голы за 90", "Голы за 90 минут у нападающих"),
                        ("Сводная реализация", "Реализация ударов, %")]:
        ws = wb[name]
        ch = LineChart(); ch.title = title; ch.height, ch.width = 9, 18
        ch.add_data(Reference(ws, min_col=2, max_col=6, min_row=1, max_row=ws.max_row), titles_from_data=True)
        ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ws.max_row))
        for s, c in zip(ch.series, SERIES):
            s.graphicalProperties.line.solidFill = c[1:]; s.graphicalProperties.line.width = 22000
            s.smooth = False
        ws.add_chart(ch, "H2")
    ws = wb["Бомбардиры"]
    ch = BarChart(); ch.type = "bar"; ch.title = "Топ-20 бомбардиров (голы)"; ch.height, ch.width = 12, 18
    ch.add_data(Reference(ws, min_col=5, min_row=1, max_row=21), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=2, max_row=21))
    ch.y_axis.scaling.orientation = "minMax"; ch.x_axis.scaling.orientation = "maxMin"
    ch.series[0].graphicalProperties.solidFill = SERIES[0][1:]; ch.legend = None
    ws.add_chart(ch, "L2")

    # Лист с картинками matplotlib
    ws = wb.create_sheet("Графики", 1)
    row = 1
    for p in chart_paths:
        img = XLImage(p); scale = 900 / img.width
        img.width, img.height = img.width * scale, img.height * scale
        ws.add_image(img, f"A{row}")
        row += int(img.height / 20) + 3
    wb.save(XLSX)


def main():
    con = build_db()
    Q = read_queries()
    con.executescript(Q.pop("create_views"))
    R = {k: pd.read_sql(q, con) for k, q in Q.items()}
    # клубы у Understat бывают склеены ("Zenit,Arsenal Tula") — убираем повторы
    R["understat_rpl"]["Клубы"] = R["understat_rpl"]["Клубы"].map(
        lambda v: ", ".join(dict.fromkeys(t.strip() for t in v.split(","))))
    for k, d in R.items():
        print(f"  {k:22s} {len(d):4d} строк")
    paths = make_charts(R)
    write_excel(R, con, paths)
    con.close()
    print("Готово:", XLSX, "| графиков:", len(paths))


if __name__ == "__main__":
    main()
