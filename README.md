# Публичный мониторинг отзывов Wildberries и Ozon

Агент собирает отзывы без доступа к личным кабинетам продавцов — через публичные страницы магазинов и товаров в обычном Chromium-браузере Playwright.

Настроенные источники:

- Ozon: `https://www.ozon.ru/seller/gk-yunikom/`
- Wildberries: `https://www.wildberries.ru/seller/38336`
- бренды: `GIGAS`, `SOFT99`

## Возможности

- ручной запуск по запросу;
- произвольный диапазон дат;
- поиск товаров заданных брендов на публичных страницах продавцов;
- сбор даты, оценки, текста, достоинств и недостатков;
- определение наличия фото и видео;
- накопительная SQLite-база с дедупликацией;
- рейтинг отзывов и доля негатива;
- выявление болей покупателей;
- фильтрация отчёта по бренду и магазину продавца;
- Excel, HTML и Markdown-отчёты.

## Важное ограничение

Сборщик не использует кабинет продавца и не требует API-токенов. Он также не обходит CAPTCHA, авторизацию, региональные ограничения и антибот-защиту. Если маркетплейс ограничивает публичный доступ, запуск завершится с понятным сообщением. Для локальной диагностики можно временно установить `headless: false`.

Разметка публичных сайтов может меняться. Сборщик использует два способа чтения отзывов:

1. анализ публичных JSON-ответов, которые загружает сама страница;
2. резервное чтение видимых карточек отзывов из DOM.

## Установка

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -e .
python -m playwright install chromium
```

Секреты Wildberries и Ozon не нужны.

## Ручной локальный запуск

Последние 7 дней:

```bash
review-monitor run --config config/targets.yml
```

Заданный диапазон:

```bash
review-monitor run \
  --config config/targets.yml \
  --date-from 2026-06-01 \
  --date-to 2026-06-30
```

Дата в формате `YYYY-MM-DD` включает весь календарный день по часовому поясу из `config/targets.yml`.

## Настройка диапазона в YAML

```yaml
date_range:
  date_from: "2026-06-01"
  date_to: "2026-06-30"
  default_lookback_days: 7
```

Аргументы командной строки имеют приоритет над YAML.

## Настроенные магазины и бренды

```yaml
sources:
  - id: ozon-gk-yunikom-public
    collector: public_marketplace_browser
    platform: ozon
    seller_url: https://www.ozon.ru/seller/gk-yunikom/
    seller_name: ГК Юником
    brands: [GIGAS, SOFT99]

  - id: wildberries-seller-38336-public
    collector: public_marketplace_browser
    platform: wildberries
    seller_url: https://www.wildberries.ru/seller/38336
    seller_name: "Wildberries seller 38336"
    brands: [GIGAS, SOFT99]
```

## Ручной запуск в GitHub Actions

Workflow больше не имеет расписания. Для запуска:

1. Откройте репозиторий GitHub.
2. Перейдите в **Actions**.
3. Выберите **Manual marketplace review monitoring**.
4. Нажмите **Run workflow**.
5. Укажите `date_from` и `date_to` или оставьте поля пустыми для последних 7 дней.

После выполнения скачайте artifact `marketplace-review-report-*`.

## Результаты

- `reports/latest.xlsx` — таблицы и аналитика;
- `reports/dashboard.html` — панель с фильтрами бренд/магазин;
- `reports/latest.md` — краткий отчёт;
- `data/reviews.sqlite3` — накопительная база.

Отчёт формируется только по выбранному диапазону дат, а база продолжает хранить историю всех запусков.

## Листы Excel

- «Сводка»;
- «Все отзывы»;
- «Негатив»;
- «Боли»;
- «Отзывы с медиа»;
- «Недельная динамика».

## Категории болей

- упаковка, повреждения и протечки;
- доставка;
- качество и подлинность;
- эффект и эксплуатационные свойства;
- совместимость и характеристики;
- цена;
- запах, цвет и консистенция;
- сервис и коммуникация;
- отсутствие инструкции или информации.

## Локальная диагностика блокировки

В `config/targets.yml` временно установите:

```yaml
headless: false
```

Браузер откроется на экране. Если сайт показывает CAPTCHA или блокировку, проект не пытается её обходить.
