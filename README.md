# Мониторинг sellers.json рекламных сетей

Раз в день скачивает sellers.json 10 сетей, сравнивает со вчерашним списком и сохраняет:
- **Новые** сайты: дата, сеть, сайт, владелец;
- **Удалённые** сайты;
- полный список отдельно (`sellers.db` → `all_sites`, выгрузка `all_sites.xlsx`).

Результаты: `sellers_changes.xlsx` (Excel) и `docs/index.html` (веб-панель с фильтрами).
Первый запуск — базовая загрузка: все сайты попадают только в полный список.

---

## Установка на GitHub (≈10 минут, без сервера и бесплатно)

### 1. Аккаунт и репозиторий
1. Зарегистрируйся на github.com (если нет аккаунта).
2. Справа вверху «+» → **New repository**.
3. Имя, например `sellers-monitor`.
4. Видимость:
   - **Public** — веб-панель будет доступна по ссылке бесплатно (данные и так публичные, это открытые sellers.json);
   - **Private** — панель через GitHub Pages только на платном GitHub Pro; на бесплатном можно скачивать `docs/index.html` и открывать у себя.
5. **Create repository**.

### 2. Загрузить файлы
1. Распакуй архив.
2. В репозитории: **Add file → Upload files**, перетащи всё содержимое папки `sellers-monitor`.
3. ⚠️ Папка `.github` скрытая — на Mac нажми `Cmd+Shift+.`, в Windows включи «Показывать скрытые элементы». Если загрузить её не получилось — создай файл вручную: **Add file → Create new file**, имя `.github/workflows/daily.yml`, вставь содержимое из архива.
4. **Commit changes**.

### 3. Разрешить роботу сохранять результаты
**Settings → Actions → General → Workflow permissions → Read and write permissions → Save.**

### 4. Первый запуск
1. Вкладка **Actions** (если просит — нажми «I understand my workflows, go ahead and enable them»).
2. Слева **Daily sellers.json check** → **Run workflow** → **Run workflow**.
3. Через 1–3 минуты появится зелёная галочка. Это базовая загрузка.
4. Открой лог запуска (шаг «python monitor.py») — там видно, какие сети скачались (✓), а какие нет (✗). Для ✗ поправь URL в `networks.json` (карандаш ✏️ при просмотре файла).

Дальше всё идёт само: каждый день в 06:00 UTC (09:00 МСК). Время меняется в `.github/workflows/daily.yml`, строка `cron`.

### 5. Включить веб-панель (GitHub Pages)
1. **Settings → Pages**.
2. Source: **Deploy from a branch**, Branch: **main**, папка **/docs** → **Save**.
3. Через минуту появится ссылка вида `https://ТВОЙ_ЛОГИН.github.io/sellers-monitor/`.

В панели: новые и удалённые сайты, фильтр по сети и периоду (последний запуск / 7 / 30 / 90 дней / всё), поиск, сортировка, сводка по сетям, статус каждой сети и выгрузка CSV.

### Где смотреть Excel
В репозитории файл `sellers_changes.xlsx` → кнопка **Download** (⬇). Полный список — `all_sites.xlsx`.

---

## Запуск на своём компьютере/сервере
    pip install -r requirements.txt
    python monitor.py --check        # какие URL отвечают
    python monitor.py --export-all   # запуск
Cron:  `0 9 * * * cd /path/to/sellers-monitor && python3 monitor.py`

## Настройки (networks.json)
- `urls` — несколько вариантов, берётся первый рабочий.
- `include_seller_types` — по умолчанию PUBLISHER и BOTH (посредники/SSP отсеиваются); `[]` — брать всё.
- `max_removal_ratio` — если за день пропало больше 30% сайтов сети, это считается сбоем файла, удаления не пишутся.
- Новая сеть — добавь строку в `networks`.
