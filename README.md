# Атакующая статистика нападающих топ-5 лиг (2015/16 – 2024/25)

Конвейер: **Python → SQLite (SQL) → Excel + графики**.

| Файл | Что это |
|---|---|
| `pipeline.py` | загрузка сырых данных в SQLite, выполнение SQL, графики, Excel |
| `queries.sql` | все SQL-запросы (представления, агрегаты, оконные функции) |
| `football.db` | база SQLite: `fbref_player_season`, `understat_player_season`, представления `fb_att`, `us_att` |
| `football_attack_stats.xlsx` | отчёт: 16 листов, нативные диаграммы Excel + все графики |
| `charts/*.png` | 11 графиков |
| `data/raw/` | исходные данные |

## Источники
- **FBref** (данные Opta) — архив `JaseZiv/worldfootballR_data`, таблицы big5_player_standard / shooting.
  Голы, передачи, пенальти, удары, удары в створ, xG / npxG / xAG, прогрессивные действия.
- **Understat** — `douglasbc/scraping-understat-dataset`: топ-5 лиг + РПЛ, 2015/16 – 2020/21; xG, xA, ключевые передачи, xGChain.

## Ограничения
- xG и прогрессивные действия в FBref есть с 2017/18, удары — с 2016/17.
- 2025/26 исключён: в архиве только первые 3–5 туров.
- «Нападающий» = позиция содержит FW (FBref) или F (Understat).
- Лига 1 2019/20 прервана из-за COVID (28 туров).

## Запуск
```bash
pip install pandas matplotlib openpyxl pyreadr adjustText
python pipeline.py
```

Пример запроса к базе:
```bash
sqlite3 football.db "SELECT season, player, goals FROM fb_att WHERE league='АПЛ' ORDER BY goals DESC LIMIT 10;"
```
