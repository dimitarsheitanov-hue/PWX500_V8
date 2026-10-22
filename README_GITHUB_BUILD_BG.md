## V5.1.38 BUILD FIX6

Add programs: disabled Confirm recognition, robust virtualized center-list folder search, and batched PowerSet UI polling. No production-worker changes. See `BUILD_FIX6_README_BG.txt`.

# CAM Agent V5.1.38 — ПЪЛЕН GitHub пакет за реален EXE

**Тестова версия.** Работещата V5.1.35 / V5.1.37 остава като резервно копие. Първият реален тест на Maestro и PowerSet с **общ пакет от няколко дати** трябва да е под контрол, защото графичните диалози на Windows не могат да се симулират в тази среда.

## Компилация без ръчни редакции

1. Разархивирай ZIP-а. Качи **всичко вътре** в ново празно GitHub repository: `src/`, `build/`, `.github/workflows/build-exe.yml` и README файловете. GitHub web upload понякога не качва скритата `.github`; провери изрично.
2. GitHub → **Actions** → **Build CAM Agent V5.1.38 EXE** → **Run workflow** (main).
3. Изчакай зелени проверки: график, регистри за дни, дублиране на PGMX, разрешени папки, PS parser и PS2EXE.
4. Свали Artifact **CAM_Agent_V5_1_38_EXE_PACKAGE** и разархивирай **`CAM_Agent_V5_1_38_COMPILED_PACKAGE.zip`**, а не само `.exe` файла.
5. Стартирай `CAM_Agent_V5_1_38.exe` от пълната папка. Не отделяй EXE-то от PS1 worker-ите, converter EXE, новия BatchEngine модул, таблиците и шаблоните.

## Три режима в оператора

**`1. ПОДГОТОВКА + PGMX`** – CSV, изделия, папки, HOP, потвърдени изделия, създаване на PGMX; копия в `Основна папка\dd.MM`, запис на `PWX500_DayPGMX_Manifest.json`. След това спира. Maestro/PowerSet не се стартират.

**`2. ОБЩ MIXX + PowerSet + XLS`** – ръчно. Намира чакащите дневни регистри, прави безопасни копия в `Основна папка\Общи_пакети\<диапазон_и_час>`, един Mixx и един PowerSet. Изходните PGMX не се изтриват; проблемните и тези с нулево реално време се МЕСТЯТ в `Проблемни_PGMX` и `Без_времена` в съответните първоначални дневни папки. Поддържа resume на незавършения общ batch по JSON статус, без повторна CSV/PGMX подготовка.

**`ЦЕЛИЯТ ПРОЦЕС 1 → 5`** – досегашният пълен дневен поток; изборът се прави изрично от отделния зелен бутон.

В **Автоматичен старт** избери `Само PGMX (без MIXX)` или `Целият процес 1–5`, час и **Запази**. Избраният режим + `LastRunDate` остават в локалния `PWX500_Agent_Schedule.ini`. Режимът по подразбиране е **само PGMX**, за да не стартира Maestro неочаквано.

## Важни правила

- `Без_времена`: `0` без предишна реална стойност → **5,44 s MEDIAN_FALLBACK**. Повторен ръчен тест и импорт на XLS с времена през бутона. Ново измерено положително време заменя медианата, дори да е по-малко.
- `Проблемни_PGMX`: при известен Maestro warning оригиналът се отделя във **собствената дневна папка**. Без повторен автоматичен опит.
- Тези изключвания се запазват в `PWX500_Agent_AutoExcludedPgmx.json` и работят след рестарт. Съществуващите файлове в изолирани папки от стари версии също се разпознават по име.
- Ако един ProgramCodeF се среща с **различно съдържание** в различни дни, общият пакет СПИРА преди Maestro. Няма презаписване или преименуване на производствени програми.
- Работната изходна папка не може да се сменя, докато има чакащи общи PGMX пакети.
- В PGMX worker-а остават задължителни проверките за разрешена папка, HOP и файловете. Само списъкът на разрешените корени вече е единен с главния агент: основни 8 + допълнително добавените.

## При преминаване от предишната инсталация

**Не изтривай** старата работеща папка. Когато първо използваш новия компилиран пакет в отделна папка, не забравяй, че локалните регистри за изделия, времена и автоматичния график стоят **до EXE-то**. За продължаване на старата история е нужно да се пренесат само локалните данни/настройки (не PS1 и не EXE); направи резервно копие преди прехвърлянето. Не смесвай двата едновременно работещи агента на едни и същи CSV входове.

## Компилация ≠ реален Windows тест

GitHub checks симулират безопасната файловата логика, но не могат да заместят Maestro UI/PowerSet и мрежовите дискове. Не приемай общия много-дневен режим за производствено потвърден преди проба с няколко контролни PGMX.


---

**BUILD FIX1:** GitHub Windows checkout may change LF/CRLF and cause a false SHA256 error. The build integrity checker now accepts only newline-style differences and rejects real edits. A `.gitattributes` file is included to preserve runtime source bytes on Windows. The agent and runtime files have NOT changed from the initial V5.1.38 package.


### BUILD FIX2 / 2026-10-22
The GitHub compiler now embeds the batch engine functions directly into the compiled script, fixing missing-function errors when starting the EXE. No changes to src/*.ps1. See BUILD_FIX2_README_BG.txt.


### BUILD FIX3 / 2026-10-22
Stage 2 incorrectly rejected an exited TimeImportProcess handle. The worker guard now checks HasExited, while active workers and ongoing time imports remain blocked. A new GitHub regression covers the exited/live worker conditions. Only src/CAM_Agent_BatchEngine_V5_1_38.ps1 and build/Test-BatchEngine.ps1 changed compared with BUILD FIX2. The main agent, converters, times worker, and other PGMX/CSV/HOP workers are unchanged.


## BUILD FIX5 – еквивалентни разрешени пътища Y:/V:/Z:

Същите продуктови подпапки на `Y:\1_Izpolzvani`, `V:\1_Izpolzvani` и `Z:\1_Izpolzvani` се приемат за еквивалентни **само ако съответният относителен път е разрешен**. Подробности и тест: `BUILD_FIX5_README_BG.txt` и `build/Test-RootDriveAliases.ps1`. Всички предишни FIX4 корекции са запазени.
