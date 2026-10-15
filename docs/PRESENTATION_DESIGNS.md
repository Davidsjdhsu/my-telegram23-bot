# Presentation design catalogue

Designer mode selects one of 50 named kits. Each kit supports 12 semantic compositions:
two covers, four photo arrangements, process, sections, comparison table, conclusion,
statement and native chart. The renderer shares tested primitives and layout systems;
these are not 600 independently authored slide files.

Selection is deterministic. An explicit “дизайн №6” / “template 6” or catalogue name
wins, followed by an explicit legacy family and subject signals. Audience alone does not
force the educational palette. Unknown subjects use neutral kits. No model or network
request is made for selection. The confirmed plan stores template_id and variant, and the
draft shows the chosen design number/name. Old four-family plans still render.

Explicit colours and light/dark background override defaults. Fonts remain editable Office
font names; viewers can substitute unavailable fonts. User photographs retain their full
frame. AI illustrations can be cropped into wide layouts. Content that cannot fit at a
readable size raises the existing validation error for bounded correction; it is not truncated.

Generation, image review, prices, billing and ordinary presentation styles are unchanged.

## Catalogue

| Number | Name | Title font | Layout system |
| --- | --- | --- | --- |
| 01 | Медовый атлас | Georgia | field |
| 02 | Ботанический дневник | Georgia | margin |
| 03 | Океаническая экспедиция | Trebuchet MS | panorama |
| 04 | Лесной маршрут | Verdana | split |
| 05 | Мир животных | Trebuchet MS | gallery |
| 06 | Космическая обсерватория | Trebuchet MS | cinema |
| 07 | Лабораторный журнал | Verdana | report |
| 08 | Физика света | Arial | poster |
| 09 | Исторический архив | Times New Roman | margin |
| 10 | Географический атлас | Trebuchet MS | field |
| 11 | Тосканский журнал | Georgia | editorial |
| 12 | Тихая роскошь | Georgia | gallery |
| 13 | Скандинавский дом | Trebuchet MS | split |
| 14 | Городская архитектура | Arial | poster |
| 15 | Путевой альбом | Georgia | panorama |
| 16 | Кофейная редакция | Georgia | margin |
| 17 | Ресторанная история | Times New Roman | editorial |
| 18 | Пекарня утром | Georgia | gallery |
| 19 | Фермерский рынок | Trebuchet MS | field |
| 20 | Чайный ритуал | Georgia | split |
| 21 | Инженерная система | Trebuchet MS | report |
| 22 | Цифровой горизонт | Arial | cinema |
| 23 | Продуктовый интерфейс | Verdana | split |
| 24 | Энергия будущего | Trebuchet MS | poster |
| 25 | Кибербезопасность | Arial | margin |
| 26 | Стратегический отчёт | Calibri | report |
| 27 | Финансовый обзор | Georgia | margin |
| 28 | Стартап на сцене | Trebuchet MS | poster |
| 29 | Командная работа | Verdana | report |
| 30 | Портрет бренда | Georgia | editorial |
| 31 | Клинический обзор | Arial | report |
| 32 | Движение и спорт | Verdana | poster |
| 33 | Психология общения | Georgia | margin |
| 34 | Баланс и здоровье | Trebuchet MS | split |
| 35 | Нейронаука | Trebuchet MS | cinema |
| 36 | Галерея искусства | Georgia | gallery |
| 37 | Модная редакция | Times New Roman | gallery |
| 38 | Музыкальная сцена | Trebuchet MS | cinema |
| 39 | Кинопанорама | Georgia | panorama |
| 40 | Книжная полка | Georgia | margin |
| 41 | Мастерская предметов | Trebuchet MS | gallery |
| 42 | Автомобильный обзор | Arial | panorama |
| 43 | Городская мобильность | Verdana | field |
| 44 | Мир игр | Trebuchet MS | cinema |
| 45 | Вещи и материалы | Georgia | field |
| 46 | Экологический обзор | Trebuchet MS | report |
| 47 | Социальная инициатива | Verdana | field |
| 48 | Образовательный курс | Verdana | report |
| 49 | Данные и исследования | Arial | report |
| 50 | Событие и фестиваль | Trebuchet MS | poster |

## Verification

`python -m unittest discover -s tests` includes all 600 template/composition combinations,
text retention, distinct geometry, bounds/overlaps, explicit selection, source photo handling,
PPTX reopen/notes-master compatibility and the actual bot builder with stubbed paid providers.

For offline visual QA, install LibreOffice, Poppler and the test Python dependencies, then run:

```
python tools/check_design_catalog.py --source-pptx /path/to/existing-deck.pptx
```

The source must contain portrait and landscape pictures. They are reused only as offline
fixtures. The script uses the production scene builder and chart renderer, exports 600 slides,
checks every slide's text through PDF extraction and creates contact sheets. It makes no paid
calls and is not a live Telegram test. Inspect the rendered images as well as automated results.
