-- =====================================================================
--  Атакующая статистика футболистов топ-5 лиг, сезоны 2015/16 – 2024/25
--  СУБД: SQLite (football.db). Каждый запрос помечен "-- name: <id>",
--  pipeline.py выполняет их по имени и выгружает в Excel / графики.
--
--  Таблицы (создаются в pipeline.py):
--    fbref_player_season     – FBref: игрок × клуб × сезон (топ-5 лиг)
--    understat_player_season – Understat: игрок × сезон (топ-5 лиг + РПЛ)
-- =====================================================================

-- name: create_views
-- Атакующие игроки = в позиции есть FW (FW, FW,MF, MF,FW, DF,FW ...).
DROP VIEW IF EXISTS fb_att;
CREATE VIEW fb_att AS
SELECT *
FROM fbref_player_season
WHERE pos LIKE '%FW%'
  AND season_end BETWEEN 2016 AND 2025;

-- Understat: позиция содержит F (форвард)
DROP VIEW IF EXISTS us_att;
CREATE VIEW us_att AS
SELECT *
FROM understat_player_season
WHERE position LIKE '%F%';

-- name: league_summary
-- Общая сводка по лигам за весь период
SELECT league                                   AS "Лига",
       COUNT(DISTINCT player_id)                AS "Нападающих",
       SUM(goals)                               AS "Голы",
       SUM(assists)                             AS "Передачи",
       SUM(pk)                                  AS "Пенальти забито",
       ROUND(SUM(goals) * 90.0 / SUM(minutes), 3) AS "Голы за 90",
       ROUND(100.0 * SUM(CASE WHEN season_end >= 2017 THEN goals END)
                   / SUM(CASE WHEN season_end >= 2017 THEN shots END), 1) AS "Реализация, % (с 16/17)"
FROM fb_att
GROUP BY league
ORDER BY "Голы" DESC;

-- name: top_scorers
-- Топ-20 бомбардиров за 10 сезонов (сумма по всем клубам игрока)
SELECT player                                   AS "Игрок",
       GROUP_CONCAT(DISTINCT league)            AS "Лиги",
       COUNT(DISTINCT season_end)               AS "Сезонов",
       SUM(minutes)                             AS "Минуты",
       SUM(goals)                               AS "Голы",
       SUM(npg)                                 AS "Голы без пен.",
       SUM(pk)                                  AS "Пенальти",
       SUM(assists)                             AS "Передачи",
       ROUND(SUM(goals) * 90.0 / SUM(minutes), 2) AS "Голы за 90"
FROM fb_att
GROUP BY player_id
ORDER BY "Голы" DESC
LIMIT 20;

-- name: top_ga
-- Топ-20 по системе гол+пас
SELECT player                                   AS "Игрок",
       SUM(minutes)                             AS "Минуты",
       SUM(goals)                               AS "Голы",
       SUM(assists)                             AS "Передачи",
       SUM(goals) + SUM(assists)                AS "Гол+пас",
       ROUND((SUM(goals) + SUM(assists)) * 90.0 / SUM(minutes), 2) AS "Гол+пас за 90"
FROM fb_att
GROUP BY player_id
ORDER BY "Гол+пас" DESC
LIMIT 20;

-- name: season_top_scorers
-- Лучший бомбардир топ-5 лиг в каждом сезоне (оконная функция RANK)
WITH per_season AS (
    SELECT season_end, season, player_id, player,
           GROUP_CONCAT(squad, ' / ')  AS squads,
           GROUP_CONCAT(DISTINCT league) AS league,
           SUM(goals) AS goals, SUM(assists) AS assists, SUM(minutes) AS minutes
    FROM fb_att
    GROUP BY season_end, player_id
), ranked AS (
    SELECT *, RANK() OVER (PARTITION BY season_end ORDER BY goals DESC) AS rk
    FROM per_season
)
SELECT season  AS "Сезон", player AS "Игрок", squads AS "Клуб", league AS "Лига",
       goals   AS "Голы", assists AS "Передачи", minutes AS "Минуты"
FROM ranked
WHERE rk = 1
ORDER BY season_end;

-- name: league_trends
-- Динамика по лигам: нападающие с ≥ 900 мин в сезоне
SELECT season_end, season AS "Сезон", league AS "Лига",
       COUNT(*)                                        AS "Игроков",
       ROUND(SUM(goals)   * 90.0 / SUM(minutes), 3)    AS "Голы за 90",
       ROUND(SUM(npg)     * 90.0 / SUM(minutes), 3)    AS "Голы без пен. за 90",
       ROUND(SUM(assists) * 90.0 / SUM(minutes), 3)    AS "Передачи за 90",
       ROUND(SUM(npxg)    * 90.0 / SUM(minutes), 3)    AS "npxG за 90",
       CASE WHEN season_end >= 2017 THEN ROUND(SUM(shots) * 90.0 / SUM(minutes), 2) END AS "Удары за 90",
       CASE WHEN season_end >= 2017 THEN ROUND(100.0 * SUM(goals) / SUM(shots), 1)  END AS "Реализация, %",
       CASE WHEN season_end >= 2017 THEN ROUND(100.0 * SUM(sot)   / SUM(shots), 1)  END AS "В створ, %"
FROM fb_att
WHERE minutes >= 900
GROUP BY season_end, league
ORDER BY season_end, league;

-- name: finishing_vs_xg
-- Реализация относительно ожидаемых голов (2017/18–2024/25), ≥ 6000 мин
SELECT player                         AS "Игрок",
       SUM(minutes)                   AS "Минуты",
       SUM(shots)                     AS "Удары",
       SUM(npg)                       AS "Голы без пен.",
       ROUND(SUM(npxg), 1)            AS "npxG",
       ROUND(SUM(npg) - SUM(npxg), 1) AS "Голы - npxG",
       ROUND(SUM(npxg) / SUM(shots), 3) AS "npxG за удар"
FROM fb_att
WHERE season_end >= 2018
GROUP BY player_id
HAVING SUM(minutes) >= 6000
ORDER BY "Голы - npxG" DESC;

-- name: p90_leaders
-- Самые продуктивные за 90 минут по ожидаемым показателям (2017/18–2024/25), ≥ 8000 мин
SELECT player                                      AS "Игрок",
       SUM(minutes)                                AS "Минуты",
       ROUND(SUM(npxg) * 90.0 / SUM(minutes), 3)   AS "npxG за 90",
       ROUND(SUM(xag)  * 90.0 / SUM(minutes), 3)   AS "xAG за 90",
       ROUND((SUM(npxg) + SUM(xag)) * 90.0 / SUM(minutes), 3) AS "npxG+xAG за 90",
       ROUND((SUM(npg) + SUM(assists)) * 90.0 / SUM(minutes), 3) AS "Голы без пен.+пас за 90"
FROM fb_att
WHERE season_end >= 2018
GROUP BY player_id
HAVING SUM(minutes) >= 8000
ORDER BY "npxG+xAG за 90" DESC
LIMIT 20;

-- name: progressive_carries
-- Прогрессивные проносы и приёмы мяча за 90 (2017/18–2024/25), ≥ 8000 мин
SELECT player                                   AS "Игрок",
       SUM(minutes)                             AS "Минуты",
       SUM(prgc)                                AS "Прогр. проносы",
       SUM(prgr)                                AS "Прогр. приёмы",
       ROUND(SUM(prgc) * 90.0 / SUM(minutes), 2) AS "Прогр. проносы за 90",
       ROUND(SUM(prgr) * 90.0 / SUM(minutes), 2) AS "Прогр. приёмы за 90"
FROM fb_att
WHERE season_end >= 2018
GROUP BY player_id
HAVING SUM(minutes) >= 8000
ORDER BY "Прогр. проносы за 90" DESC
LIMIT 20;

-- name: heatmap_goals
-- Голы по сезонам для 12 лучших бомбардиров периода
WITH top AS (
    SELECT player_id FROM fb_att
    GROUP BY player_id ORDER BY SUM(goals) DESC LIMIT 12
)
SELECT f.season_end, f.season AS "Сезон", f.player AS "Игрок", SUM(f.goals) AS "Голы"
FROM fb_att f JOIN top USING (player_id)
GROUP BY f.season_end, f.player_id
ORDER BY f.season_end;

-- name: understat_creators
-- Understat (топ-5 + РПЛ, 2015/16–2020/21): лучшие созидатели среди нападающих
SELECT player                                   AS "Игрок",
       GROUP_CONCAT(DISTINCT league)            AS "Лиги",
       SUM(minutes)                             AS "Минуты",
       SUM(assists)                             AS "Передачи",
       ROUND(SUM(xa), 1)                        AS "xA",
       SUM(key_passes)                          AS "Ключевые передачи",
       ROUND(SUM(key_passes) * 90.0 / SUM(minutes), 2) AS "Ключевые за 90",
       ROUND(SUM(xgchain), 1)                   AS "xGChain"
FROM us_att
GROUP BY understat_id
ORDER BY "xA" DESC
LIMIT 20;

-- name: understat_rpl
-- РПЛ (Understat, 2015/16–2020/21): лучшие бомбардиры и их xG
SELECT player                                   AS "Игрок",
       GROUP_CONCAT(DISTINCT team)              AS "Клубы",
       COUNT(DISTINCT season)                   AS "Сезонов",
       SUM(minutes)                             AS "Минуты",
       SUM(goals)                               AS "Голы",
       ROUND(SUM(xg), 1)                        AS "xG",
       SUM(assists)                             AS "Передачи",
       ROUND(SUM(xa), 1)                        AS "xA",
       SUM(shots)                               AS "Удары"
FROM understat_player_season
WHERE league = 'РПЛ'
GROUP BY understat_id
ORDER BY "Голы" DESC
LIMIT 15;
