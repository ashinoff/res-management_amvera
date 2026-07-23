# CLAUDE.md — РЭС-менеджмент (res-management)

Контекст проекта для Claude Code. Читается автоматически в начале сеанса.
Карта проекта, задачи переезда/интеграции и уже известные «грабли». Текущее
состояние файла всегда сверяй с кодом — этот файл карта, а не снимок.

> ВАЖНО: полностью перешли на Amvera (2026-07-15). Render-зеркало
> `github.com/ashinoff/res-management` БОЛЬШЕ НЕ ИСПОЛЬЗУЕТСЯ — работаем только
> с этим repo (`res-management_amvera`), синхронизировать в Render не нужно.

## Что это

Система управления РЭС (районные электрические сети): учёт приборов учёта (ПУ)
по структуре сети (ТП → ВЛ → позиции начало/середина/конец), загрузка и анализ
выгрузок со счётчиков (РиМ, Нартис, Энергомера), уведомления об ошибках,
мероприятия РЭС с фотоотчётами, повторные проверки АСКУЭ, проблемные ВЛ,
отчёты и аналитика. Пользователи: админ (АСКУЭ), загрузчики, ответственные РЭС.

Интерфейс на русском. Роли внутри ПО: `admin`, `uploader`, `res_responsible`.

## Стек

- **Бэкенд:** Node.js + Express, Sequelize 6 + PostgreSQL, JWT (jsonwebtoken),
  bcryptjs, multer (память + диск для Excel), xlsx, nodemailer (почта РЭС/ПЭС),
  **Cloudinary** (все вложения: фото/PDF — во внешнем хранилище, не в ФС).
- **Анализаторы:** Python 3 (pandas, openpyxl, xlrd) — `backend/analyzers/*.py`,
  вызываются из Node через `spawn('python3', ...)`. **Контейнер обязан содержать
  и Node, и Python с зависимостями из `backend/requirements.txt`.**
- **Фронтенд:** React 18 + Vite 4, axios, xlsx. SPA без роутера — навигация
  состоянием. Монолит: почти всё в `frontend/src/App.jsx`.
- **Монолит из двух файлов** (как «Светлячок»): `backend/server.js` (~5500 строк),
  `frontend/src/App.jsx` (~6800 строк). Осознанно; правки точечные, не дробить
  без явной просьбы.

## Структура репозитория

```
backend/
  server.js               ВСЁ: модели, роуты, email, python-вызовы, init БД
  analyzers/              nartis_analyzer.py, energomera_analyzer.py,
                          rim_converter_csv.py (pandas/openpyxl)
  requirements.txt        python-зависимости анализаторов
  fix-old-pdfs.js         разовый скрипт (legacy)
  package.json
frontend/
  src/App.jsx             ВСЁ: страницы, поллинг уведомлений, формы
  src/App.css
  public/icons, images
  vite.config.js          dev-proxy /api → VITE_API_URL
```

### Карта `backend/server.js` (порядок в файле)
конфиг/Cloudinary → подключение БД (pool настроен) → модели → хуки User
(bcrypt) → JWT middleware (`authenticateToken`, `checkRole`) → multer →
email-сервис → роуты `/api/*` → `initializeDatabase()` (sync + индексы + seed
РЭСов и админа) → `app.listen`.

### Модели (Sequelize, таблицы во множественном числе: "Notifications" и т.п.)
ResUnit, User (login/password/role/resId/**email — уже есть, NOT NULL**),
NetworkStructure (ТП/ВЛ, startPu/middlePu/endPu), PuStatus, Notification
(type: error/success/info/pending_check/pending_askue/problem_vl; payload —
JSON-строка в `message`), UploadHistory, CheckHistory (вложения — JSON со
ссылками Cloudinary), ProblemVL, NotificationRead, PuUploadHistory.

### Ключевые роуты
`/api/auth/login|me`, `/api/network/structure*`, `/api/upload/analyze`
(Excel → python-анализатор → статусы ПУ + уведомления),
`/api/notifications*` (+ `/counts` — поллится фронтом каждые 30 с),
`/api/notifications/:id/complete-work` (мероприятия + фото в Cloudinary),
`/api/reports/*`, `/api/history/*`, `/api/problem-vl/*`, `/api/users/*`,
`/api/admin/*` (health/cleanup/diagnose), `/api/download/:public_id`.

## Дизайн-инварианты (зафиксировано 2026-07-11, НЕ нарушать)

- **Единственная заставка загрузки — «РОССЕТИ»** (`RossetiLoader`, анимация
  `electric-bulb`). Никаких других: кружки-спиннеры, точки, пульсирующие
  кольца, мигающий текст — всё удалено из CSS и JSX. Не возвращать. Внутри
  кнопок при операции — только текст состояния («Загрузка…», «Проверка…»)
  + `disabled`, без графики.
- **Никаких бесконечных CSS-анимаций в контенте** (`animation: ... infinite`).
  Они были причиной тормозов подсветки у админа (сотни элементов
  анимировались каждый кадр). Удалены: greenGlow/redPulse/phaseError на
  индикаторах фаз, pulse на статус-боксах и critical-icon, вращение/прыжки
  декоративных SVG (`db-header-icon`, `no-issues-icon`). Исключение — только
  `electric-bulb` заставки на экранах загрузки.
- **Единый hover:** карточки (`.notification-compact`, `.problem-card`) —
  рамка `--navy` + фон `--surface-2` + `--shadow-sm`; строки всех таблиц —
  фон `--surface-2`. Переходы только адресные и быстрые
  (`border-color/background/box-shadow .12s`), НЕ `transition: all 0.3s`.
- **Единое выделение (чекбокс):** везде целиком фон `--accent-soft`; у
  карточек ещё рамка `--accent`; выделение не сбрасывается при hover.
  Канонический блок — в конце `App.css` («ЕДИНЫЙ СТИЛЬ ВЫДЕЛЕНИЯ...»).
- **Уведомления кликабельны целиком:** клик по карточке = «Детали»
  (error/pending_askue → модалка деталей, problem_vl → переход к разделу).
  Кнопок «Детали» больше нет. Остальные кнопки («Завершить», «Загрузить»,
  «К проблемным ВЛ») и чекбоксы живут внутри карточки с `stopPropagation`.
  Новые кнопки внутри карточек — тоже обязательно с `stopPropagation`.
- Рамка карточек — единая `1px solid var(--border)`, без цветных полос слева.

## Что уже сделано (перф-фиксы, база для этого репо)

Стартовый код репозитория — УЖЕ с фиксами производительности (не откатывать):
- `/api/notifications`: includes с ограниченными `attributes` (пароль наружу
  не отдаётся), `NotificationRead` — только текущего пользователя.
- `/api/notifications/counts`: фильтр `type IN ('error','pending_askue')` в БД.
- 14 индексов через `CREATE INDEX IF NOT EXISTS` в `initializeDatabase()`.
- `sequelize.sync({ alter: ... })` — только при env `DB_ALTER=true` (разово
  после изменения моделей, потом убрать). Обычный старт — быстрый `sync()`.
- Явный pool (max 10). Фронт: троттлинг inactivity-таймера (30 с),
  поллинг останавливается на скрытой вкладке.

---

# ЗАДАЧИ 1–3 (✅ все выполнены, см. журнал; оставлено как справка)
# Актуальная работа — раздел «ТЕКУЩАЯ ЗАДАЧА» ниже.

## Задача 1. Переезд Render → Amvera (один Docker-контейнер)

Повторить схему «Светлячка»/«СИЗ-контроля»:

1. **Dockerfile мультистейдж:**
   - Stage 1 `node:20-alpine`: `npm ci` + `npm run build` фронта → `frontend/dist`.
   - Stage 2 — runtime с Node **и Python**: удобнее `node:20-slim` +
     `apt-get install python3 python3-pip` + `pip3 install -r
     backend/requirements.txt --break-system-packages`. Скопировать `backend/`,
     `frontend/dist`. `ENV TZ=Europe/Moscow`. CMD `node server.js`, порт 8000
     (`PORT=8000`).
   - Убрать `postinstall` из `backend/package.json` (pip ставится в Dockerfile).
2. **Единый origin.** Express дополнительно раздаёт `frontend/dist` как статику
   + SPA-fallback на `index.html`. **Fallback — ПОСЛЕДНИМ middleware**, после
   всех `/api/...`, иначе перехватит API. Текущий `GET /` (health JSON) →
   перенести на `GET /api/health` (на него же смотрит Amvera).
3. **Фронт на относительный API.** Сейчас `const API_URL =
   import.meta.env.VITE_API_URL || 'http://localhost:3000'` и абсолютные ссылки
   (`/api/download/...` на строке ~4693). Сделать `API_URL = import.meta.env.
   VITE_API_URL || ''` → axios ходит на свой origin (`/api/...`). Dev-режим
   не ломать: vite-proxy уже настроен.
4. **`amvera.yml`:** environment docker, containerPort 8000. persistenceMount
   `/data` не обязателен (файлы в Cloudinary, БД в Postgres), но заведи на
   будущее.
5. **Гонка DNS при старте** (грабля Amvera: «Temporary failure in name
   resolution»): перед `sequelize.authenticate()` — retry-цикл (до 15 попыток,
   пауза 3 с). Не убирать.
6. **Переменные окружения Amvera:** `DATABASE_URL` (managed Postgres Amvera;
   в проде SSL — сейчас включается по `NODE_ENV=production`, проверить, что на
   Amvera работает; возможно понадобится `ssl:false` — у Amvera БД во внутренней
   сети), `JWT_SECRET` (новый — разлогинит старые сессии, это ок),
   `DELETE_PASSWORD`, `CLOUDINARY_*` (3 шт., переносятся как есть),
   `MAIL_HOST/PORT/USER/PASS`, `PORT=8000`, разово `DB_ALTER=true` на первый
   деплой (создание схемы), потом убрать.
7. **Эфемерная ФС:** папка `uploads/` — только временные Excel для
   анализаторов, терять не жалко. Ничего постоянного в ФС не писать.
8. Деплой НЕ автоматический: после `git push` — вручную «Пересобрать» в панели
   Amvera. Напоминать пользователю после каждого пуша.

## Задача 2. Бэкап/восстановление (перенос данных Render → Amvera)

Сейчас бэкапа НЕТ вообще. Сделать и учесть порядок действий:

1. **`GET /api/admin/backup`** (только admin): JSON-дамп ВСЕХ таблиц:
   `{ format: "full", version: 1, exportedAt, tables: { ResUnits: [...],
   Users: [...], NetworkStructures: [...], PuStatuses: [...],
   Notifications: [...], NotificationReads: [...], UploadHistories: [...],
   CheckHistories: [...], PuUploadHistories: [...], ProblemVLs: [...] } }`.
   Пароли пользователей выгружать КАК ЕСТЬ (bcrypt-хэши) — чтобы логины
   пережили переезд; следить, чтобы hook повторного хэширования не сработал
   при restore (в модели уже есть защита `startsWith('$2a$')` — проверить, что
   покрывает и `$2b$`). Ссылки Cloudinary едут внутри JSON — сами файлы
   переносить не нужно, хранилище общее.
2. **`POST /api/admin/restore`** (только admin, принимает этот JSON):
   восстановление в порядке FK-зависимостей: ResUnits → Users →
   NetworkStructures → PuStatuses → Notifications → NotificationReads →
   CheckHistories → UploadHistories → PuUploadHistories → ProblemVLs.
   Схема одна и та же → **id сохраняем как есть** (вставка с явными id),
   после каждой таблицы — `SELECT setval(pg_get_serial_sequence(...), max(id))`.
   Ошибки копить и возвращать списком (первые 20), не падать целиком.
   Восстановление — только в пустую/очищенную БД (проверять и сообщать).
3. **Грабля Amvera:** пользователь БД НЕ суперпользователь — никаких
   `session_replication_role`/отключения триггеров. Только правильный порядок
   вставки.
4. **Кнопки в админке** (фронт): «Скачать бэкап» (blob → файл
   `res-backup-YYYY-MM-DD.json`) и «Восстановить из файла» с подтверждением.
5. **Порядок переезда (важно, объяснять пользователю):** endpoint выгрузки
   должен попасть и на СТАРЫЙ Render-деплой → задеплоить туда → скачать бэкап →
   поднять Amvera с пустой БД → restore. Т.е. коммит с задачей 2 нужен ДО
   отключения Render.

## Задача 3. Интеграция с платформой SUE_system (Keycloak SSO + бейдж)

Эталон — «Светлячок» (проверено в проде 2026-07-08), но он на FastAPI/Python.
Здесь **Express/Node** — паттерн тот же, реализация своя. Полный контракт —
`PLATFORM_INTEGRATION.md` в репо Светлячка и CLAUDE.md платформы.

**Ключевая идея (не отступать):** Keycloak решает только «кто ты» (email) и
«пускать ли» (ОДНА realm-роль доступа). Функциональная роль
(admin/uploader/res_responsible) и `resId` берутся из СВОЕЙ БД по email.
Никакого маппинга ролей из токена и авто-создания учёток.

1. **Фиче-флаг `PLATFORM_SSO`** (env, default OFF). При OFF поведение не
   меняется вообще, обычный логин/пароль всегда остаётся как fallback.
2. **`backend/keycloakPlatform.js`** — проверка Keycloak-JWT по JWKS. В Node:
   пакет `jose` (`createRemoteJWKSet` — кэширует сам) или `jwks-rsa` +
   `jsonwebtoken`. Проверять: подпись, `iss` (`KEYCLOAK_URL` +
   `/realms/KEYCLOAK_REALM`), `exp`, `azp === 'web-desktop'`
   (`KEYCLOAK_AZP`). **aud НЕ проверять** (public-клиент). Требовать realm-роль
   доступа из `realm_access.roles`: env `ACCESS_ROLE`, предложение —
   **`resm-user`** (согласовать с пользователем; паттерн платформы
   `<app>-user`). Нет роли → **403** (личность есть, доступа нет); невалидный
   токен/SSO выключен → 401. **Токен не логировать и не сохранять** — только
   причины отказа.
3. **`POST /api/auth/platform`** — обмен: Keycloak-токен в
   `Authorization: Bearer` → проверка (п.2) → поиск пользователя: сначала по
   `keycloakId`, затем разово по email (регистронезависимо, `LOWER()`), при
   успехе — записать `keycloakId` (привязка). Не нашли → 401. Выдать ОБЫЧНЫЙ
   JWT этого ПО (тот же `jwt.sign`, что в `/api/auth/login`) + тот же формат
   ответа `{ token, user }`, что ждёт фронт.
4. **Колонка `Users.keycloakId`** (STRING(64), nullable, unique) — добавить в
   модель; доедет через разовый `DB_ALTER=true` (или добавить
   `ALTER TABLE ... ADD COLUMN IF NOT EXISTS` к блоку индексов —
   Postgres-safe). `email` в модели уже есть и обязателен — удобно: заполнять
   реальными адресами через экран «Пользователи» (там форма уже с email).
5. **CSP-middleware:** на КАЖДЫЙ ответ (не только /api) —
   `Content-Security-Policy: frame-ancestors 'self' <PLATFORM_ORIGIN>`
   (только эта директива!) и удалить `X-Frame-Options`, если есть.
   `PLATFORM_ORIGIN` — env, default `https://sue-system-ashinoff.amvera.io`.
6. **Фронт (App.jsx), контракт `platform-auth`/`app-ready` НЕ МЕНЯТЬ:**
   - `EMBEDDED = window.self !== window.top`; в iframe старому токену из
     localStorage не доверять.
   - После маунта послать родителю `{type:'app-ready'}`, слушать `message`,
     принимать ТОЛЬКО с `event.origin === VITE_PLATFORM_ORIGIN` и
     `type === 'platform-auth'`; токен обменять через `POST /api/auth/platform`
     чистым `fetch` (не через axios-инстанс, чтобы интерсепторы/редиректы не
     мешали), результат положить как обычный токен.
   - Пока ждём — лоадер «Вход через платформу…» (`ssoPending`); таймаут 5 с
     или неуспех → обычная форма логина (fallback).
   - 401-обработку внутри iframe не превращать в редирект на логин платформы.
7. **`GET /api/platform/badge`** — счётчик для бейджа иконки на рабочем столе.
   Контракт платформы: ответ `{"count": N}`, авторизация ТЕМ ЖЕ
   Keycloak-токеном (Bearer), БЕЗ создания сессии, только чтение. Учётка по
   `keycloakId` → email; не найден/SSO OFF → `{"count":0}`/401 (тихо — платформа
   любые ошибки глотает). `count` — «требует действия» по роли, переиспользовать
   логику `/api/notifications/counts`: для admin — tech_pending + askue_pending
   + активные problem_vl; для res_responsible — свои tech_pending; для uploader —
   свои askue_pending. **CORS:** origin платформы должен быть разрешён для этого
   роута (сейчас `cors()` открыт всем — тогда ничего не менять, но не сужать,
   не добавив `PLATFORM_ORIGIN`).
8. **На стороне платформы (репо SUE_system, отдельно):** запись в
   `src/config/apps.js` (id, `iconUrl: '/apps/<id>.png'`, `badge: true`,
   `roles: ['resm-user','admin']`, URL из `VITE_APP_RESM_URL`), картинка в
   `public/apps/`, в Keycloak — realm-роль `resm-user` и email у пользователей.
   Это НЕ в этом репозитории — только напомнить пользователю.
9. **env для SSO:** `PLATFORM_SSO=true`, `KEYCLOAK_URL`,
   `KEYCLOAK_REALM=platform`, `KEYCLOAK_AZP=web-desktop`,
   `ACCESS_ROLE=resm-user`, `PLATFORM_ORIGIN=...`; фронту при сборке —
   `VITE_PLATFORM_ORIGIN` (в Dockerfile через build-arg или захардкоженный
   default как у СИЗ).

---

# ТЕКУЩАЯ ЗАДАЧА: включение интеграции с платформой SUE_system

Код интеграции в ЭТОМ репо уже написан и проверен (задача 3, коммит `805af90`,
за флагом `PLATFORM_SSO`, default OFF). «Начать интеграцию» = включить её:
env + сторона платформы + Keycloak + сквозная проверка. Порядок:

## Шаг 1 — этот репозиторий (кода не требуется)
1. Убедиться, что в проде живёт актуальный main (включая рестайл 2026-07-11),
   `npx vite build` зелёный.
2. В панели Amvera (проект res-management) добавить env:
   `PLATFORM_SSO=true`, `KEYCLOAK_URL=<url кейклока платформы>`,
   `KEYCLOAK_REALM=platform`, `KEYCLOAK_AZP=web-desktop`,
   `ACCESS_ROLE=resm-user`, `PLATFORM_ORIGIN=https://sue-system-ashinoff.amvera.io`.
   Значения KEYCLOAK_* сверить с работающим «Светлячком» — берём те же.
3. Вручную «Пересобрать». Проверить: обычный вход по паролю НЕ сломался
   (SSO — только дополнительный путь, fallback обязателен).

## Шаг 2 — учётки (данные, не код)
- У каждого пользователя этого ПО, который должен входить через платформу,
  поле `email` должно СОВПАДАТЬ с email в Keycloak (регистр не важен,
  сравнение через LOWER). Заполнять через экран «Пользователи».
- Матчинг только по email → привязка `keycloakId` происходит при первом входе
  автоматически. Авто-создания учёток НЕТ — нет в БД по email = 401.

## Шаг 3 — репозиторий SUE_system (отдельная сессия Claude Code там)
1. `src/config/apps.js`: запись приложения — id (например `resm`),
   `iconUrl: '/apps/resm.png'`, `badge: true`, `roles: ['resm-user','admin']`,
   URL из env `VITE_APP_RESM_URL` (= адрес этого ПО на Amvera).
2. Иконка в `public/apps/resm.png` (попросить у пользователя или сгенерить
   строгую в стиле остальных иконок платформы).
3. Keycloak (realm `platform`): создать realm-роль `resm-user`, выдать её
   нужным пользователям; проверить, что у них заполнен email.
4. Env платформы: `VITE_APP_RESM_URL=<адрес res-management>`. Пересобрать
   платформу вручную.

## Шаг 4 — сквозная проверка (по образцу «Светлячка»)
- Иконка появилась на рабочем столе только у пользователей с ролью `resm-user`.
- Клик по иконке: iframe шлёт `app-ready`, платформа отвечает `platform-auth`,
  ПО входит без формы логина (лоадер РОССЕТИ «Вход через платформу…»).
- Пользователь без роли → 403; email не найден в БД ПО → 401 + обычная форма.
- Бейдж на иконке = `GET /api/platform/badge` → `{count}` (по роли).
- Прямой заход на адрес ПО (не через платформу) → обычный логин работает.
- В логах нет Keycloak-токенов.

Контракты `platform-auth`/`app-ready` и `{count}` зафиксированы платформой —
формат НЕ менять. Эталон живой интеграции — «Светлячок» (прод 2026-07-08).

---

## Грабли (уже наступали в соседних проектах — НЕ повторять)

- **Контейнер Amvera эфемерный.** Постоянные файлы — только Cloudinary/Postgres.
- **Пользователь БД Amvera — не суперпользователь.** Никаких суперюзер-операций.
- **Гонка DNS БД на старте** — обязателен retry-цикл подключения.
- **`sequelize.sync()` без alter не добавляет колонки** в существующие таблицы.
  Новые колонки: `DB_ALTER=true` разово ЛИБО `ADD COLUMN IF NOT EXISTS` рядом с
  индексами. Колонки — nullable/с дефолтом.
- **ENUM в Postgres:** новые значения enum `sync` не добавит — только
  `ALTER TYPE ... ADD VALUE IF NOT EXISTS`.
- **SPA catch-all — строго последний** middleware, иначе съест `/api`.
- **Контракт `platform-auth`/`app-ready` и `/api/platform/badge` → `{count}`**
  зафиксирован на платформе — не менять формат.
- **Keycloak-токен не логировать** нигде (в т.ч. в console.log ошибок).
- **Деплой ручной:** после пуша — «Пересобрать» в панели Amvera.
- **Windows bash-tool:** рабочая директория сбрасывается между вызовами —
  абсолютные пути.

## Проверки перед коммитом

- `node --check backend/server.js`.
- `cd frontend && npm run build` (ловит JSX/импорты).
- Анализаторы: `python3 -m py_compile backend/analyzers/*.py`.
- При `PLATFORM_SSO=false` старый вход по паролю не изменился (регресс-минимум:
  login → me → notifications/counts).
- grep: Keycloak-токен нигде не логируется/не сохраняется.

## Журнал изменений (Claude Code ведёт сам)
- **2026-07-24** — «Загрузить файлы» приведена к формату остальных страниц:
  `.file-upload-container` — карточка на всю ширину (убраны `max-width:1200/margin:auto`),
  заголовок H2 слева с SVG-логотипом (как у других страниц). Красная рамка-предупреждение
  убрана: «Имя файла должно совпадать с номером ПУ» — просто красная строка,
  ниже muted-подзаголовок «Загружайте Excel…», всё по левому краю. 4 карточки типов —
  в ОДНУ строку (`grid repeat(4,1fr)`, на ≤900px → 2), тонкая рамка 1px, БЕЗ hover-
  transform (не дёргаются), шрифты тоньше (`.ft-label` 16px/600, `.ft-sub` 12px/400).
  Большая drop-зона заменена компактной строкой: кнопка «Выбрать файл» + подсказка
  форматов + выбранные файлы как чипы (`.file-pick-*`, `.file-chip*`). Прогресс/кнопка
  загрузки/результаты не тронуты. npm run build — ОК.
- **2026-07-24** — Пакет правок (App.jsx/App.css): (1) нижняя граница волны поднята
  выше разделителя (`.osc-bg bottom: -7 → 4px`). (2) ВСЕ плавающие кнопки структуры
  теперь показываются только при активном скролле-вверх (`{showScrollTop && …}`),
  внутри — удалить/очистить по чекбоксам. (3) Номера секций шин — РИМСКИЕ везде
  (`toRoman`, 1..5→I..V; отображение `СШ-{toRoman(...)}`); в форме секции поле →
  выпадающий список I–V (до 5). (4) Подпись секции крупнее (`.section-title` 15.5px,
  `.section-peak` 13px); число «пик» цветом по уровню (`.peak-num` red>100/amber
  90-100/green<90); процент рядом с Pmax/Пик в модалках (сделано ранее). (5)
  Редизайн «Загрузить файлы»: убрана карточка «Текущий РЭС»; «ВАЖНО» → компактное
  предупреждение с красным светящимся контуром (`.upload-warning`); заголовок
  «Загрузка файлов…» по левому краю с SVG-логотипом (`.upload-header.left`); «1.
  Выберите тип счётчика» → «Тип загрузки»; карточки типов без иконок, крупная метка
  (`.ft-label` 20px 800) + подпись (`.ft-sub`): РИМ/Нартис/Энергомера «журнал
  напряжения», Профиль мощности «Пирамида сети»; «2.» убрано. npm run build — ОК.
- **2026-07-24** — Мелкие правки: (1) нижняя граница волны заголовка поднята на
  ~2 мм (`.osc-bg bottom: -14px → -7px`). (2) Плавающие кнопки структуры: «Обновить
  структуру» и «Выгрузка в Excel» теперь показываются ВМЕСТЕ с кнопкой «наверх»
  (обёрнуты в `{showScrollTop && ...}`), а не всегда; «Очистить»/«Удалить» — как и
  было, по чекбоксам. Цвета: выгрузка Excel — зелёная (`.fab-green`), обновить —
  синяя (`.fab-blue`), очистить историю — оранжевая (`.fab-warn`), удалить — красная
  (без изменений). npm run build — ОК.
- **2026-07-24** — Осциллограф заголовка, финал: (1) искра проходит РОВНО ОДИН раз —
  раньше dash-паттерн тайлился (`dasharray 0.13 1`, период 1.13 ≈ pathLength) и в
  кадре появлялся второй сегмент («дорисовывал хвост в начале названия»). Сделал
  `dasharray 0.13 2` (период 2.13) + offset `0.13 → -1`: соседние копии dash всегда
  за краями [0,1], виден только один бегущий сегмент. (2) Волна во всю ширину меню
  (`.osc-bg left/right: -24px` — компенсация padding 24) и выходит выше/ниже букв
  (`top:-18 / bottom:-14`, ~0.5 см). Цвет/скорость/толщина не менялись. npm run
  build — ОК.
- **2026-07-23** — Осциллограмма заголовка, переделка под «осциллограф»:
  постоянной линии больше нет (`.osc-base` убран из JSX). `OSC_PATH` — высокочастотная
  (8 периодов) глубокая (амплитуда почти во всю высоту, viewBox 0 0 120 44) волна.
  Эффект: тонкая (1.2px) холодно-белая (#f2fbff) бегущая искра с коротким видимым
  хвостом (~2 колебания, `stroke-dasharray: 0.13 1`), проходит слева-направо и
  исчезает (конечная `osc-spark .85s forwards`, re-mount по `key={sweepTick}` на
  каждый успешный полл counts). Свечение (drop-shadow) тонировано `color`:
  энерго-голубой #8fe3ff (норма) / тёплый #ff9aa2 (перегруз). `mix-blend-mode:
  screen` — светится над тёмным фоном, белые буквы не темнит; тонко, не замазывает
  текст. npm run build — ОК, без infinite.
- **2026-07-23** — Осциллограмма заголовка: волна яркая «энергетическая» и ПОВЕРХ
  текста (`.osc-bg` z-index 2 + `mix-blend-mode: screen` — светится над тёмным
  фоном, белые буквы не темнит), неон-цвета (#2bff88 норма / #ff4d5e перегруз),
  база opacity .85 + drop-shadow glow, `.osc-spark` белый с усиленным свечением.
  Плюс в модалках техучёта/случая рядом с «Pmax N кВт»/«Пик N кВт» — процент
  загрузки (`.tech-pmax-pct`) цветом: >100% красный, 90–100% оранжевый, <90%
  зелёный (пороги = как у шкалы `barCls`). npm run build — ОК, без infinite.
- **2026-07-23** — Заголовок сайдбара, доработка: встроен логотип «Мониторинг
  напряжения» с платформы (`RESM_LOGO` — инлайн ResmTile, зелёно-бирюзовая плитка
  с пульс-линией), текст мельче — «МОНИТОРИНГ» капсом (15px) + «напряжения» мелко
  снизу (12px, `.mt-1/.mt-2`). Синусоида (`OSC_PATH`) сделана **глубокой** (amp 15,
  3 периода, viewBox 0 0 120 44) и перенесена **за текст** (`.osc-bg` absolute,
  z-index 0; текст `.monitor-title` z-index 1) — отдельная линия под названием
  убрана. База волны приглушена (opacity .14); по каждому успешному поллу counts —
  электрический «пробег» со свечением (`.osc-spark`, drop-shadow, конечная
  `osc-spark .95s forwards`, re-mount по `key={sweepTick}`), затем гаснет. Цвет
  green/red по counts (transition .12s). prefers-reduced-motion: пробег off, база
  чуть заметнее. Инвариант «без infinite» соблюдён. npm run build — ОК.
- **2026-07-23** — Заголовок сайдбара «Меню» → «МОНИТОРИНГ НАПРЯЖЕНИЯ» (2 строки,
  `.monitor-title`, uppercase, letter-spacing 0.06em) + инлайн-SVG осциллограмма
  (`OSC_PATH` — синусоида ~2.25 периода, viewBox 0 0 120 14, stroke 2px). Анимации
  КОНЕЧНЫЕ (инвариант «без infinite» соблюдён): (1) прочерчивание слева-направо при
  монтировании `osc-draw .9s forwards` (dasharray/dashoffset, pathLength=1); (2)
  «пробег» блика на КАЖДЫЙ успешный полл `/api/notifications/counts` — оверлей-path
  с `key={sweepTick}` (re-mount перезапускает `osc-sweep .8s forwards`), tick
  инкрементится в `loadNotificationCounts` при успехе (без таймеров-циклов); (3)
  цвет линии: `--green` если нет активных error/power_overload (tech_pending+
  powerOverload+power_overload==0), иначе `--red`, переход `transition color .12s`;
  (4) `prefers-reduced-motion: reduce` — анимации off, линия статична. Поллер, пункты
  меню, бейджи не тронуты (только подписка на успех). Проверено: grep — новых
  `animation: infinite` нет; npm run build — ОК.
- **2026-07-23** — UI-правки (App.jsx/App.css): (1) кнопка «Ограничение по АСКУЭ
  выполнено» убрана из тела карточки «Превышение Pном» — только в футере модалки
  деталей (для admin=АСКУЭ + askue_limit); `submitAskue` спрашивает
  подтверждение (`window.confirm`). (2) Хронология в модалке случая: убрана
  толстая серая кайма слева у `.po-step`; статус-слова цветом — «ожидается»
  оранжевый (`.po-wait`), «выполнено/завершён/устранён» зелёный (`.po-done-word`),
  «повторный перегруз» красный (`.po-fail-word`). (3) Главный «Пик N кВт» в
  модалке случая — красный (`.po-details-modal .tech-pmax-value`). (4) «Проблемные
  ВЛ»: описание сжато в строку `.problem-info-row`, справа счётчик активных
  (`.problem-active-counter`), «Всего зарегистрировано» убран. (5) В «Аналитику»
  добавлен отчёт «ТП с перегрузом» (ниже остальных): `/api/reports/overload`,
  фильтр ratioPct≥100, таблица РЭС/ТП/СШ/Sном/лимит/пик/%/статус/циклы. npm run
  build — ОК.
- **2026-07-23** — UI: модалки техучёта/перегруза сделаны содержательными, карточки
  «Превышение Pном» — в стиле «Ожидающие мероприятий» (только App.jsx/App.css).
  Модалка «Сведения о техучёте» (`.tech-details-modal`, шире, +верхний
  `.modal-info` ТП/СШ/№ПУ, крупный Pmax 34px, шкала загрузки, сетка 2 колонки с
  бОльшими отступами, кнопка перехода в «Превышение Pном», развёрнутое пустое
  состояние). Карточки PowerOverload переведены на `.notification-compact
  power_overload` + `.notification-narrow-content` (индикатор-квадрат + info
  ТП·СШ/РЭС/пик/лимит/% + actions), клик→детали. Модалка деталей кейса
  (`.po-details-modal`) — статус-пилюля, крупный пик/лимит, сетка, **хронология**
  (`.po-timeline`/`.po-step` done/fail/pending: АСКУЭ→РЭС→перепроверка→закрытие с
  ФИО/датами/комментариями/фото); кнопки действий вынесены в футер модалки.
  Модалка действия (АСКУЭ/РЭС) — +`.modal-info` со сводкой, счётчик слов для РЭС,
  плейсхолдеры. npm run build — ОК.
- **2026-07-23** — КОРЕНЬ бага «профиль не выявляет перегруз» (подтверждён живым
  тестом на Postgres): `UploadHistory.create` при type='profile' падал —
  `invalid input value for enum "enum_UploadHistories_fileType": "profile"`, до
  анализатора дело не доходило. Фикс: (1) в ОБЕ модели с fileType-enum
  (`UploadHistory`, `PuUploadHistory`) добавлено значение `'profile'`. (2) В
  `initializeDatabase` (Postgres-блок, по образцу power_overload) для таблиц
  `UploadHistories`/`PuUploadHistories` имя enum-типа берётся из pg_catalog по
  колонке fileType → `ALTER TYPE "<имя>" ADD VALUE IF NOT EXISTS 'profile'`
  (идемпотентно, try/catch, двойной старт не падает). (3) В профиль-ветке
  `UploadHistory.create` обёрнут в try/catch — при падении 400 с текстом +
  `console.error('[PROFILE] …')`, больше не выглядит как «ничего не нашлось».
  Логика анализатора/матчинга/кейсов верна (проверено заказчиком: лимит 225→
  пик 328.4 overload+кейс; 360→212.2 ok). node --check — ОК.
- **2026-07-23** — КРИТ. фикс скорости анализатора + таймаут + редизайн модалки
  техучёта. (1) `profile_analyzer.py`: убран `read_only=True`, `_read_sheet`
  переписан на ОДИН проход `iter_rows(values_only=True)` (строки 5/6 — шапка,
  с 9-й — данные до «Итого»). Причина зависания: read-only + случайный
  `ws.cell(row,col)` перепарсивал лист на каждый вызов → квадратично на 1449×8.
  Теперь 1440 строк × 6 ПУ = 0.54 с (было — минуты/зависание). Методику НЕ трогал
  (проверено: 30-мин 1.059→211.8, 60-мин как есть). Плюс `sys.stdout.reconfigure
  (utf-8)` — защита от POSIX/C-локали (en-dash в period не уронит вывод).
  (2) `runProfileAnalyzer` (server.js): таймаут 120 с → `python.kill('SIGKILL')`
  + `{success:false,error:'Анализатор не уложился в 120 с'}`; stderr питона в
  `console.error('[PROFILE] stderr…')`. (3) Модалка «Сведения о техучёте»
  переработана: статус-пилюля в шапке (Норма/Перегруз/Нет данных цветами
  статусов), крупный «Pmax N кВт» + дата·период, горизонтальная шкала загрузки
  (0..120%→0..100% ширины, зелёная<90/оранж90-100/красная>100, статический div
  без transition), сетка 2 колонки (№ПУ/Sном/cosφ/лимит/источник/дата загрузки),
  строка кейса + ссылка, пустое состояние с иконкой. Для «источник» и «дата
  загрузки» добавлены nullable-колонки `TpSection.lastProfileSource`/
  `lastProfileAt` (ALTER IF NOT EXISTS), заполняются в профиль-ветке. node --check,
  py_compile, npm run build — ОК.
- **2026-07-23** — Фикс профиля: перегруз не выявлялся при заниженном Sном.
  ГЛАВНЫЙ БАГ (b): openpyxl отдаёт номера ПУ как float → `str()` давал
  «1294249.0», не матчилось с techPuNumber «1294249» → секция не обновлялась,
  overload не срабатывал. Анализатор: `_pu_str` нормализует номер (int/float→
  без «.0»), matched-нормализация и на сервере (`normPu`, обе стороны). (a)
  Подтверждено: анализатор умножает peakKw=peakRaw×Кт один раз, сервер сравнивает
  и пишет `lastPeakKw` по peakKw (не peakRaw) — двойного/нулевого умножения нет.
  (c) Секции POST/PUT принимают запятую (`parseFloat(replace(',','.'))`) +
  `Number.isFinite` (иначе tnKva→null, cosPhi→0.9); в анализе `hasLimit` требует
  `isFinite(tnKva)&&>0` → unknown, не 0. Диагностика (штатная): ответ
  `/api/upload/analyze` для profile содержит `details[]` по каждому ПУ
  {puNumber,matched,sectionId,tpSection,peakRaw,kt,peakKw,peakAt,tnKva,cosPhi,
  limitKw,decision}; `console.log('[PROFILE]...')`; фронт — свёрнутая таблица
  «Детали расчёта» (1 знак, decision цветом). Зелёный статус (`status-ok`)
  теперь отображается (был скрыт из-за бага матчинга). Модалка «Сведения о
  техучёте» по клику на квадрат секции (stopPropagation): №ПУ/Sном/cosφ/лимит,
  Pmax+дата+период, статус цветом, активный кейс + ссылка «Перейти в Превышение
  Pном» (проброшен `onSectionChange` в NetworkStructure). node --check,
  py_compile, npm run build — ОК; юнит анализатора peakKw=peakRaw×200 сходится.
- **2026-07-23** — ЭТАП 3, Блок Г (server.js/App.jsx): отчёты по перегрузу.
  Бэкенд: `GET /api/reports/overload` (все секции с заданным Sном + данные
  последнего кейса: РЭС/ТП/СШ/Sном/cosφ/лимит/пик/дата/%/статус случая/даты
  АСКУЭ и РЭС/результат перепроверки/циклы; res_responsible — свой РЭС). В
  `/api/analytics/summary` добавлены per-РЭС `overloadSections` (секций
  overloadStatus='overload') и `activeOverloadCases` (кейсы != completed) + в
  totals. Фронт: в отчётах новый тип «Превышение Pном» (опция селекта, ветки
  loadReports/exportToExcel/getReportTitle + отдельная таблица для этого типа,
  т.к. колонки другие); в сводном отчёте Analytics — 2 новые колонки в таблицу
  и Excel (в конец, существующие не тронуты). Проверки: node --check, npm run
  build — ОК. ЭТАП 3 (А+Б+В+Г) завершён.
- **2026-07-23** — ЭТАП 3, Блок В (App.jsx/App.css): плавающие кнопки экрана
  «Структура сети». Вертикальный стек круглых SVG-кнопок (`.structure-fab-stack`,
  fixed справа, над кружком «наверх», gap 10px, 50px): всегда «Выгрузка в Excel»
  (`exportStructureToExcel`, IconDownload) и «Обновить структуру»
  (`loadNetworkStructure`, IconRefresh); при выбранных чекбоксах (admin) — «Очистить
  историю» (IconBroom, `handleClearTpHistory`) и «Удалить выбранные» (IconTrash,
  `setShowDeleteModal`) с бейджем-счётчиком (`.fab-badge`). Те же обработчики, что
  у верхней панели (её не трогал). title-hover, transition только адресные, без
  infinite. `npm run build` — ОК. Блок Г — следующий.
- **2026-07-23** — ЭТАП 3, Блок Б (App.jsx): меню «Превышение Pном» (иконка
  IconZap, роли admin/res_responsible, бейдж `counts.powerOverload`) между
  «Проблемные ВЛ» и «Загруженные документы»; роут `case 'power_overload'` →
  `<PowerOverload>`. Компонент: вкладки Активные/Завершённые, карточки в стиле
  `.notification-compact.problem-card` (клик=детали-модалка с хронологией
  АСКУЭ/РЭС/фото/перепроверка); действия по роли/этапу — admin+askue_limit →
  «Ограничение по АСКУЭ выполнено» (комментарий необязателен), res_responsible+
  res_work → «Мероприятия выполнены» (модалка комментарий≥5 слов + фото,
  multipart на `/api/overload/:id/res-complete`), awaiting_recheck → плашка
  «Ожидает перепроверки». Бейдж цикла «повтор N» при cycles>1. CSS `.po-*` без
  анимаций. `npm run build` — ОК. Блоки В/Г — следующими.
- **2026-07-23** — ЭТАП 3, Блок А (server.js): workflow перегруза. Модель
  **`OverloadCase`** (sectionId FK, resId FK, stage askue_limit/res_work/
  awaiting_recheck/completed, cycles, снимок peakKw/peakAt/tnKva/cosPhi/limitKw/
  ratio/period, askue*/res* поля, attachments JSON, recheck*, closedAt) +
  ассоциации + индексы IF NOT EXISTS (sectionId; resId,stage). Профиль-ветка
  переписана на кейсы: overload+нет кейса→создать askue_limit+уведомление;
  overload+кейс askue_limit/res_work→обновить цифры (без дубля уведомления);
  overload+awaiting_recheck→перепроверка провалена (still_overload, stage→
  askue_limit, cycles+1, новое уведомление «повторный перегруз, цикл N»);
  ok+awaiting_recheck→успех (completed, closedAt, уведомление power_overload
  удалить, создать success в РЭС); ok+askue_limit/res_work→закрыть кейс. Хелпер
  `removeSectionOverloadNotifs` чистит уведомления + NotificationRead. Эндпоинты:
  `GET /api/overload?stage=&resId=` (res_responsible — свой РЭС), `POST
  /api/overload/:id/askue-complete` (admin, askue_limit→res_work + уведомление
  РЭС), `POST /api/overload/:id/res-complete` (res/admin, multipart комментарий+
  фото Cloudinary, res_work→awaiting_recheck + уведомление АСКУЭ). Защита секций:
  DELETE→400 при активном кейсе (завершённые кейсы удаляются вместе с секцией),
  PUT techPuNumber→warning при активном кейсе. `getNotificationCounts` +поле
  `powerOverload` (admin=askue_limit+awaiting_recheck, res=res_work своего РЭСа),
  существующие поля не тронуты. node --check — ОК. Блоки Б/В/Г — следующими.
- **2026-07-23** — «Структура сети»: перепривязка/отвязка ВЛ к секции. Раньше
  селект «Секция…» был только у ВЛ в блоке «ВЛ без секции» — после первой привязки
  сменить/снять секцию было нельзя. Теперь в колонке секции у КАЖДОЙ строки ВЛ
  (привязанной и нет) — селект: value = `item.sectionId ?? ''`; опция value=""
  динамическая («Секция…» у непривязанной / «— Без секции» у привязанной) + все
  секции ТП. `assignSection` переписан: оптимистично двигает ВЛ между группами
  локально через `setNetworkData` (без перезагрузки; группировка/счётчики «(N)»
  пересчитываются из networkData), при ошибке — откат + alert. `stopPropagation`
  на onClick/onChange селекта. Блок «ВЛ без секции» уже скрыт при пустоте
  (`unassigned.length > 0`). Ширину колонки секции не менял (сетка из прошлого
  коммита). Бэкенд PUT `/api/network/structure/:id` НЕ менялся — он уже различает
  отсутствие `sectionId` в body (не трогать) vs `null` (отвязать) vs новое
  значение (валидация: та же ТП+РЭС, иначе 400). Только фронт. `npm run build` — ОК.
- **2026-07-23** — «Структура сети»: выравнивание квадратов ПУ в строгую сетку
  (визуальная правка до этапа 3, только `App.jsx`/`App.css`, бэкенд не тронут).
  Введена единая CSS-grid `.net-grid` (фикс-колонки: `28px minmax(160px,1fr)
  96px 96px 96px 140px 72px`) для ВСЕХ строк ВЛ, заголовков секций и блока «ВЛ без
  секции» → квадраты начало/середина/конец совпадают в три вертикальные колонки
  независимо от длины наименования и наличия номера. `renderPuCell` теперь всегда
  рендерит ровно один квадрат (в т.ч. серый «X» при пустом ПУ) + строку номера
  фиксированной высоты (`.pu-num-line` min-height 18px, ellipsis, letter-spacing
  -0.2px для длинных номеров вроде 12733192358417 — без расширения колонки).
  Заголовок секции — та же сетка: в колонке наименования индикатор техучёта
  (`.status-box--sm`) + текст секции + пик/лимит; над колонками ПУ подписи
  «Начало/Середина/Конец» (`.pu-col-label`, 12px, `--text-muted`, по центру);
  в колонке действий — иконки ред./удал. В «ВЛ без секции» подписи колонок
  показаны один раз (строка-заголовок `colHeader`). Мобилка: `.tp-card
  overflow-x:auto` + `.net-grid min-width:680px` — сетка не ломается, карточка
  скроллится. Хук привязки секций, селекты, чекбоксы, hover/выделение, цвета
  статусов — не тронуты. `--text-secondary` в проекте нет → `--text-muted`.
  `.status-box` (глобальный) не менял. Проверка: `npm run build` — ОК
  (визуально на Amvera сверить).
- **2026-07-23** — Профиль мощности, ЭТАП 2/3 (workflow мероприятий — этап 3, НЕ
  здесь). Анализатор `backend/analyzers/profile_analyzer.py` (openpyxl): выгрузка
  «ПРОФИЛЬ МОЩНОСТИ ДЛЯ 1С», листы «30 мин»/«60 мин»; ПУ по колонкам с C (строка 5),
  kt=Ктт×Ктн (строка 6), данные с 9 до «Итого»; «24:00» → 00:00 след. суток.
  Методика (зафиксирована): 60-мин ряд как есть, иначе 30-мин час H:00=(H:30+
  (H+1):00)/2 (пропуск=0); пик=max, peakKw=max×kt, energy=sum×kt. JSON:
  results[{puNumber,kt,peakRaw,peakKw,peakAt,energyKwh,source,period}]+warnings.
  Проверено: ПУ 1294249 → peakRaw 1.059=(2.118+0)/2 @01.06 01:00, peakKw ×200=211.8;
  60-мин «как есть» + 24:00. server.js: enum Notification +'power_overload' (модель
  + `ALTER TYPE ... ADD VALUE IF NOT EXISTS` в initializeDatabase, имя типа из
  pg_catalog, только Postgres, autocommit — повторный старт не падает).
  `/api/upload/analyze` type='profile' (resId не нужен): `runProfileAnalyzer` →
  матчинг puNumber↔`TpSection.techPuNumber` (trim); limitKw=tnKva×cosPhi,
  overloadStatus=tnKva? (peakKw≥limit?overload:ok):unknown; обновляет lastPeakKw/
  lastPeakAt(`parsePeakAt`)/lastProfilePeriod/overloadStatus; при overload —
  Notification power_overload (resId секции, toUserId=null, errorData с sectionId/
  peakKw/limitKw/ratio/period), дедуп по sectionId (cast errorData::text LIKE);
  ok→старую удаляем; ПУ без секции+warnings → `unmatched` (не ошибка). UploadHistory
  fileType 'profile', processedCount=секций, errorCount=перегрузов.
  `getNotificationCounts`/badge: +power_overload для admin. Фронт: тип «Профиль
  мощности (Пирамида)» (resId не шлётся), сводка «секций N · перегрузов M · не
  привязано K (список)»; живой индикатор секции (red/green/gray) + подпись «пик X
  кВт · время · лимит Y кВт» (1 знак); карточка+модалка power_overload у админа
  (клик=детали, без кнопок действий — этап 3). Проверки: node --check, py_compile,
  npm run build — ОК (живой старт с БД локально не гонялся — нет драйвера; ALTER
  TYPE идемпотентен и под try/catch).
- **2026-07-23** — Секции шин ТП, ЭТАП 1/3 (только модель + экраны структуры; БЕЗ
  анализатора профилей и workflow — следующие коммиты). Бэкенд (`server.js`):
  новая модель **`TpSection`** (resId FK, tpName, sectionNumber, tnKva=Sном кВА,
  cosPhi=0.9, techPuNumber, overloadStatus ENUM ok/overload/unknown, lastPeakKw/
  lastPeakAt/lastProfilePeriod); у `NetworkStructure` — колонка **`sectionId`**
  (nullable; NULL = «ВЛ без секции»). Ассоциации: `NetworkStructure.belongsTo
  TpSection as 'section'`, `TpSection.hasMany NetworkStructure as 'lines'`.
  Идемпотентно в `initializeDatabase` (после `sync()`): `ALTER TABLE
  NetworkStructures ADD COLUMN IF NOT EXISTS sectionId`, индексы IF NOT EXISTS
  `idx_netstruct_section`, `idx_tpsection_res`, уникальный `idx_tpsection_unique
  (resId,tpName,sectionNumber)` — DB_ALTER не нужен (новую таблицу создаёт sync).
  API: `GET /api/network/sections?resId=` (секции + `linesCount`),
  POST/PUT/DELETE `/api/network/sections[/:id]` (admin, как редактирование
  структуры; DELETE→400 если привязаны ВЛ), `PUT /api/network/structure/:id`
  принимает `sectionId` с валидацией (та же ТП и РЭС, иначе 400); GET структуры
  включает `section` (ограниченные attributes). Фронт (`App.jsx`, `NetworkStructure`):
  плоская таблица заменена на **группировку ТП → секции → ВЛ** (карточка ТП →
  блоки секций «СШ-N · кВА · тех.учёт №…» с квадратом-индикатором техучёта тем же
  стилем `.status-box` (unknown=серый/ok=зелёный/overload=красный, данные пока
  всегда unknown) + блок «ВЛ без секции» с селектом привязки). Форма секции
  (номер, «Sном тр-ра, кВА», cosφ=0.9, № ПУ техучёта). Логика начало/середина/
  конец (`renderPuCell`), чекбоксы/удаление/фильтры/экспорт — не тронуты, только
  переиспользованы. БЕЗ новых анимаций (дизайн-инварианты). CSS новых классов —
  в конце `App.css`. Проверки: `node --check`, `npm run build` — ОК (живой старт
  с БД локально не гонялся — нет драйвера; на Amvera sync создаёт таблицу,
  ALTER идемпотентный).
- **2026-07-17** — UX модалки «Отметить выполнение мероприятий» (роль РЭС,
  `Notifications` в `frontend/src/App.jsx`). Раньше кнопка «Подтвердить
  выполнение» была `disabled` при <5 слов в комментарии — клик ничего не давал,
  РЭС думали, что ПО не работает. Теперь: кнопка кликабельна (в `disabled`
  осталось только `submitting`), `handleCompleteWork` при <5 словах не шлёт alert,
  а ставит `commentError=true` → у textarea зажигается красная рамка со свечением
  (`boxShadow rgba(220,38,38,.2)`), а label и счётчик слов краснеют + подсказка
  «нужно не менее 5 слов, чтобы завершить». Подсветка гаснет, как только набрано
  5+ слов (в `onChange`); сбрасывается при открытии модалки. Только фронт.
- **2026-07-11** — Фикс бейджа «Ожидающие проверки АСКУЭ» у загрузчика: в
  `getNotificationCounts` выборка для роли `uploader` была `{toUserId: user.id}`,
  но `pending_askue` создаются с `toUserId=null, resId=РЭС` (broadcast) — счётчик
  их не видел, бейдж не появлялся, хотя список показывал 5. Привёл выборку к той
  же, что в `GET /api/notifications` (Op.or: свои личные + общие pending_askue по
  РЭС). Дополняет предыдущий фикс (тот убрал зависимость от «прочитано» для РЭС).
- **2026-07-11** — Счётчики уведомлений (`getNotificationCounts`, питает и меню-
  бейджи, и `/api/platform/badge`) считают ВСЕ ожидающие `error`/`pending_askue`,
  а НЕ только непрочитанные. Убран фильтр `if (!readIds.has(...))` и лишний
  запрос к `NotificationRead`. Причина: раньше открыл раздел → пометилось
  «прочитано» → пропало из счётчика, хотя ПУ не отработан. Эти уведомления и так
  удаляются при загрузке проверки по ПУ (`Notification.destroy` для `pending_askue`
  ~стр. 3113; `error` ~3292/3304), поэтому счётчик падает при реальной отработке.
  Дедуп по «ПУ+фазы» сохранён; список (`GET /api/notifications`) и так отдавал
  прочитанные с `isRead` → бейдж совпадает со списком. Нужно, чтобы уведомления
  дотягивались в бейдж платформы. Коммит `8638837`.
- **2026-07-11** — Админ может менять `login` существующей учётки: `PUT
  /api/users/:id` принимает `login` с проверкой уникальности (`Op.ne`), во фронте
  снят `disabled` с поля «Логин» в модале редактирования. Цель — унификация
  учёток (одинаковый логин у человека во всех приложениях; роль/доступ платформа
  определяет по email). Коммит `951067d`.
- **2026-07-11** — Удаление пользователя (`DELETE /api/users/:id`) больше не падает
  на FK. Всё в одной `sequelize.transaction`: `NotificationRead` пользователя —
  удаляются (мусор); `Notification.fromUserId/toUserId → NULL` (уведомления
  сохраняются, `toUserId=null` делает личное общим по РЭС — подхватят другие);
  `UploadHistory.userId`/`PuUploadHistory.uploadedBy → NULL` (аудит сохраняем);
  затем `user.destroy`. 5 FK на Users (все в этих таблицах), других нет. Коммит
  `bf1f090`.
- **2026-07-11** — Карточка «Очистка данных системы» (Настройки → Обслуживание,
  `MaintenanceSettings`) видна только под учёткой `user.login === 'admin'` (не
  всей роли admin). Вкладку «База данных» не трогали. Коммит `1b4fc30`.
- **2026-07-11** — В iframe платформы при неуспешном SSO-входе показывается
  «Нет доступа. Обратитесь к администратору» (ветка `EMBEDDED` в `App.jsx`),
  а не `LoginForm`. Вне iframe (standalone) — обычный логин как прежде.
- **2026-07-11** — Новая проверка №11 `irrelevant_problem_vl` («актуальность
  проблемных ВЛ») в health check + автоисправление. Не путать с №8
  `stale_problem_vl` (та — только по возрасту, 90 дней без активности).
  Логика: для каждой активной `ProblemVL` ищем `PuStatus` по `puNumber`;
  если ПУ не найден в структуре ИЛИ `status === 'checked_ok'` — запись
  потеряла актуальность (severity warning, items с reason, первые 10).
  Auto-fix (case в `database-cleanup`): такие ВЛ переводятся в **resolved**
  (как авто-resolve при чистой проверке; НЕ dismissed). В stats добавлен
  `irrelevantProblemVL`. Фронт: запись в `getCleanupDescription`, тип добавлен
  в whitelist кнопки «Очистить», свой рендер примеров (генерик не подходит —
  ждёт поле `count`), исключён из генерик-фолбэка. Ловит записи, оставшиеся
  active с времён до внедрения авто-resolve. `node --check` + `vite build` — ОК.
- **2026-07-11** — Файлы Cloudinary теперь ТОЛЬКО через свой прокси
  (`/api/download`). Причина: `res.cloudinary.com` — общий CDN, попадает в
  базы угроз → Яндекс.Браузер блокирует прямые ссылки («Сайт заблокирован»).
  Бэкенд: роут переписан с redirect 302 на **стриминг** (сервер сам fetch-ит
  signed URL и отдаёт поток со своего домена; `?inline=1` — показать в
  браузере, без — скачать; Content-Type/Length с апстрима, filename в
  UTF-8). Фронт: helper `fileProxyUrl(file, inline)` рядом с `API_URL`;
  все `src=`/`href={file.url}` заменены (миниатюры, «Открыть», просмотр
  картинок, «Открыть в новой вкладке» PDF, fallback «Скачать файл»).
  Старые записи без `public_id` — fallback на прямой url.
  В просмотрщике у изображений добавлена кнопка «Скачать» (раньше была
  только у PDF) — тоже через прокси.
  **Инвариант: прямые ссылки на res.cloudinary.com в разметку не вставлять.**
  `node --check` + `vite build` — ОК.
- **2026-07-11** — Хотфикс прод-ошибки 42883 в проверке целостности БД
  (`backend/server.js`). Check 6 `broken_file_references` (и одноимённый
  auto-fix) падал: `attachments != '[]'` — колонка `DataTypes.JSON`, у типа
  `json` в Postgres НЕТ операторов сравнения. Оба места переведены на
  `Sequelize.where(Sequelize.cast(Sequelize.col('attachments'),'text'),
  {[Op.ne]:'[]'})`. Ошибка глоталась try/catch — «Total issues: 0» мог врать,
  т.к. проверка 6 реально не выполнялась. `node --check` — ОК.
  **Грабля на будущее:** json-колонки (attachments и др.) нельзя сравнивать
  `=`/`!=` напрямую — только через `::text` или перевод колонки в JSONB.
- **2026-07-11** — Перф-фикс подсветки + единый стиль уведомлений/выделения +
  клик-по-карточке = «Детали» + только заставка РОССЕТИ. Логика не тронута,
  только `App.jsx`/`App.css`, `vite build` — ОК. Подробности:
  - **Причина тормозов hover у админа:** бесконечные CSS-анимации на каждом
    элементе списков (`greenGlow`/`redPulse` анимировали ПРОЗРАЧНЫЕ
    псевдоэлементы — невидимая трата ресурсов; `pulse` на красных
    статус-боксах структуры сети — сотни штук у админа; `phaseError` на
    индикаторах фаз; pulse на `critical-icon`). Все удалены. Плюс
    `transition: all 0.3s` в сайдбаре и на уведомлениях → заменены на
    адресные `.12s` (как было у `.problem-card`).
  - `.notification-compact` приведён 1:1 к `.problem-card`: рамка 1px
    (`border-left: 4px` убран), паддинги 14/18, hover navy+surface-2+shadow-sm.
  - Единый блок выделения в конце `App.css`: selected = `--accent-soft`
    целиком (+ рамка `--accent` у карточек), сохраняется при hover; в таблице
    документов был `--blue-bg` → приведён к общему.
  - Кнопки «Детали» удалены (error и pending_askue); клик по всей карточке
    открывает детали (problem_vl → переход к разделу). «Завершить»,
    «Загрузить», «К проблемным ВЛ» — с `stopPropagation`. success/info —
    `cursor: default`.
  - Декоративные анимации SVG остановлены: вращение `db-header-icon`
    (проверка целостности + бэкап), прыжки `no-issues-icon`.
  - **Все альтернативные загрузчики удалены** (JSX + CSS, ~270 строк):
    кружки в 6 кнопках (диагностика, масс-фикс, загрузка структуры, детальный
    отчёт, проверка БД, очистка) — теперь только текст состояния; из CSS —
    `.spinner`, `.spinner-small`, `.loading-spinner-small`×2,
    `.loading-spinner-large`×2, точки `bounce-dots`, кольца `pulse-ring`,
    `pulse-text`, `button-pulse`, `.btn-loading::after` и 4 дубля
    `@keyframes spin`. Пункт из записи `650be55` «мелкие спиннеры оставлены»
    БОЛЬШЕ НЕ АКТУАЛЕН. См. раздел «Дизайн-инварианты».
- **2026-07-11** — «Проблемные ВЛ» + рестайл уведомлений. Коммит `d6f50e7`.
  - Карточка проблемной ВЛ компактная и кликабельная целиком (открывает
    подробности); убраны кнопки «Написать письмо» и «Рассмотреть без
    объяснительной» (модалки/хендлеры оставлены дремать — вернуть позже).
    Мета в одну строку (`.problem-meta`), строка ошибки компактно.
  - Раздел теперь реагирует на основной фильтр РЭС: `<ProblemVL selectedRes>`,
    `loadProblemVLs` шлёт `?resId`, рефетч по смене `selectedRes`; бэкенд
    `/api/problem-vl/list` принимает `resId`.
  - **Логика авто-resolve:** в обычном пути проверки при `!result.has_errors`
    активная `ProblemVL` (по `puNumber`) переводится в `resolved` — ВЛ уходит
    из проблемных (раньше это было только в recheck-ветке).
  - Все уведомления (`.notification-compact`, `.problem-card`): убраны боковые
    цветные полосы (`border-left`), строгий hover тёмно-синей рамкой
    (`--navy`) + `--surface-2`, выбранное — акцентная рамка (`--accent`);
    вся карточка — `cursor: pointer`. Без «прыжка» transform.
- **2026-07-11** — SVG-логотипы вместо PNG + иконки меню/заголовков + чистый
  лоадер. Коммит `f708e4e`. Все декоративные картинки (`/icons/important.png`,
  `PU.png`, `place.png`, `ok.png`) заменены на инлайн-SVG в «синей рамке»
  (класс `.svg-frame` — `--accent-soft` фон, синий border + glow): памятка
  структуры (IconEdit), «Текущий РЭС» (IconMapPin), «имя файла=номер ПУ»
  (IconAlertTriangle), выбор типа счётчика ×3 (новый `IconMeter`), «проблемных
  ВЛ нет» (IconCheck). Левому меню (массив `menuItems`) присвоены SVG-иконки
  (рендер `.menu-ico`); те же иконки в `.svg-frame` добавлены в заголовки h2
  всех разделов. Заставка `RossetiLoader` — убран текст «что грузим»
  (`LoadingSpinner`/`db-loading`/`.loading` рисуют только анимацию). Превью
  загруженных пользователем файлов (`file.url`, вложения) НЕ трогали. Логику не меняли.
- **2026-07-11** — Заставка загрузки «РОССЕТИ» (буквы загораются по очереди),
  перенос 1:1 из «Учёта ПУ». Коммит `650be55`. Новый `frontend/src/RossetiLoader.jsx`
  (7 букв Р-О-С-С-Е-Т-И, `animationDelay: idx*0.3s`) + CSS `@keyframes electric-bulb`
  (серый `#cbd5e1` → электрик-синий `#2563eb` с glow) в `App.css`. Центральный
  `LoadingSpinner` (все типы default/dots/pulse/inline/overlay), блок `db-loading`,
  текстовые `.loading` и экран ожидания входа через платформу теперь рисуют РОССЕТИ.
  Мелкие спиннеры-«кругляшки» внутри кнопок (`.loading-spinner-small`,
  `.spinner-small`) оставлены — РОССЕТИ в кнопку не помещается. Логику не трогали.
- **2026-07-11** — Редизайн фронта в стиле «СИЗ Контроль» (строгий, без эмодзи),
  3 этапа, логика НЕ тронута:
  - Этап 1 (`be1b600`): дизайн-токены `:root` (navy/accent/surface/border/
    семантика) в начале `App.css`, сплошная замена хардкод-цветов на токены —
    убраны 67 фиолетовых градиентов и glow-тени, сайдбар → navy, палитра как в СИЗ.
  - Этап 2 (`ed31b86`): новый `frontend/src/icons.jsx` (инлайн SVG, lucide-стиль,
    `currentColor`). Все ~258 эмодзи в `App.jsx` заменены: в JSX — на иконки
    (семантика цвета green/red/amber через `currentColor`), в строках и
    комментариях — эмодзи удалены (текст сохранён). Класс `.ico` для выравнивания.
    Знак «№» (U+2116) — не эмодзи, оставлен.
  - Этап 3 (`1b4b480`): полировка `App.css` — таблицы (шапка `--surface-2`,
    12-13px muted, border-bottom, hover, без зебры), модалки (overlay
    `rgba(15,23,42,.45)`, `--surface`+`--radius`+`--shadow`), формы (border
    `--border`, focus-кольцо `--accent-soft`), бейджи статусов (`*-bg` + цвет,
    radius 999px), нейтральные спиннеры, заголовки 650/`-0.01em`.
  - Проверки: `npx vite build` после каждого этапа — ОК; сплошной скан — эмодзи
    не осталось. Только `App.css`, `App.jsx`, новый `icons.jsx`.
- **2026-07-11** — Бэкап/restore через **gzip** (обход HTTP 413: прокси Amvera
  режет большие тела до приложения, бэкап ~27МБ). Коммит `319cd70`. Добавлен
  `const zlib = require('zlib')`. Фронт `handleRestore`: сжимает файл через
  `CompressionStream('gzip')` перед отправкой (fallback без сжатия, если API нет).
  Бэкенд `restore`: детект gzip по магическим байтам `1f 8b` → `zlib.gunzipSync`
  перед `JSON.parse` (несжатый файл — как раньше). Бэкенд `backup`: отдаём
  `zlib.gzipSync` + `Content-Encoding: gzip` (браузер распакует прозрачно).
  Зависимостей не добавляли. Формат JSON бэкапа не менялся.
- **2026-07-09** — репозиторий заведён; в базе — перф-фиксы уведомлений,
  индексы, DB_ALTER, троттлинг фронта. Задачи 1–3 описаны выше, не начаты.
- **2026-07-09** — ✅ **Задача 1 сделана** (переезд Render → Amvera, коммит
  `9e78f77`). Что сделано:
  - `Dockerfile` мультистейдж: stage1 `node:20-alpine` собирает фронт
    (`npm install` + `npm run build`, lock-файлов в репо нет — не `npm ci`);
    stage2 `node:20-slim` + `python3/pip` + `requirements.txt`
    (`--break-system-packages`), копирует `backend/` и `frontend/dist`,
    `TZ=Europe/Moscow`, `PORT=8000`, `CMD node server.js` из `/app/backend`.
  - `server.js`: health `GET /` → **`GET /api/health`**; добавлена раздача
    `frontend/dist` (`express.static`) + **SPA-fallback `app.get('*')`
    последним** (пропускает `/api` и `/uploads` через `next()`); подключение к
    БД через `connectWithRetry()` (15×3с, грабля DNS Amvera) вместо прямого
    `authenticate()`.
  - Фронт `App.jsx`: `API_URL` по умолчанию **`''`** (относительный, свой
    origin). Vite-proxy для dev не трогали.
  - `backend/package.json`: убран `postinstall` (pip — в Dockerfile).
  - Добавлены `amvera.yml` (docker, порт 8000), `.dockerignore`, `.gitignore`.
  - Проверки прошли: `node --check`, `py_compile analyzers`, `npm run build`.
  - **Осталось (ручное, на стороне пользователя):** env на Amvera
    (`DATABASE_URL`, `JWT_SECRET`, `DELETE_PASSWORD`, `CLOUDINARY_*`,
    `MAIL_*`, `PORT=8000`, разово `DB_ALTER=true`), при необходимости проверить
    SSL к managed-Postgres, и **вручную «Пересобрать» в панели Amvera**.
    Задачи 2–3 — не начаты.
  - **SSL БД** теперь через env `DB_SSL` (true/false); по умолчанию — включён в
    production. У Amvera Postgres во внутренней сети SSL часто НЕ поддерживается
    → при ошибке подключения на первом деплое выставить `DB_SSL=false`.
- **2026-07-10** — ✅ **Задача 2 сделана** (бэкап/восстановление, коммит
  `60b4c02`). Реализация:
  - `GET /api/admin/backup` (admin) — JSON-дамп всех 10 таблиц через
    `SELECT *` (raw), `{format:"full",version:1,exportedAt,tables:{...}}`,
    отдаётся как файл `res-backup-YYYY-MM-DD.json`. Пароли — как есть.
  - `POST /api/admin/restore` (admin) — файл через **multer memoryStorage**
    (обход лимита `express.json`), требует поле `confirm='true'`. **Решение:
    restore ПОЛНОСТЬЮ заменяет данные** (очистка `DELETE` в обратном FK-порядке
    → вставка raw-SQL `INSERT` с явными id → `setval`), а не «только в пустую
    БД» — так корректно переживаются сеяные на старте ResUnits/admin и нет
    коллизий id. Вставка raw-SQL идёт **мимо Sequelize-hooks** → пароли, id и
    даты сохраняются 1:1. JSON/JSONB-значения при вставке сериализуются обратно в
    строку (pg приводит к jsonb). Ошибки копятся (первые 20), не падает целиком.
  - Список/порядок таблиц — `BACKUP_TABLES` в `server.js` (родители→дети).
  - Фикс hook `User.beforeUpdate`: не перехэшировать хэш и с префиксом `$2b$`
    (раньше только `$2a$`).
  - Фронт: раздел «Обслуживание» (`DatabaseMaintenance`) → «Скачать бэкап»
    (blob→файл) и «Восстановить из файла» (input→`window.confirm`→fetch+FormData).
  - Проверки: `node --check`, `npx vite build` — ОК.
  - **Порядок переезда (объяснять пользователю):** endpoint бэкапа должен быть
    и на СТАРОМ Render → задеплоить туда коммит `60b4c02` → скачать бэкап →
    поднять Amvera (пустая БД, `DB_ALTER=true` разово) → «Восстановить из файла».
    Т.е. этот коммит нужен на Render ДО отключения.
- **2026-07-10** — ✅ **Задача 3 сделана** (интеграция с платформой, коммит
  `805af90`). Реализация:
  - `backend/keycloakPlatform.js` — проверка Keycloak-JWT по JWKS на
    **jwks-rsa + jsonwebtoken** (CommonJS, без ESM `jose`): подпись, `iss`,
    `exp`, `azp==web-desktop`; aud НЕ требуем. Роль доступа `ACCESS_ROLE`
    (`resm-user`). Токен нигде не логируется. Добавлена зависимость `jwks-rsa`.
  - `POST /api/auth/platform` — обмен токена платформы на обычный JWT (тот же
    `{token,user:{id,fio,role,resId,resName}}`, что `/api/auth/login`). Юзер по
    `keycloakId`, затем по `email` (`LOWER`, разовая привязка `keycloakId`). SSO
    off/невалид → 401, нет роли `resm-user` → 403, не найден → 401.
  - `GET /api/platform/badge` — счётчик для бейджа (без сессии, только чтение).
    `count` по роли: admin = tech_pending+askue_pending+problem_vl,
    res_responsible = tech_pending, uploader = askue_pending. Логика вынесена в
    `getNotificationCounts()` (переиспользуется и `/api/notifications/counts`).
  - `Users.keycloakId` (VARCHAR(64) unique) — в модели + `ALTER TABLE ADD
    COLUMN IF NOT EXISTS` в `initializeDatabase` (работает без `DB_ALTER`).
  - CSP `frame-ancestors 'self' <PLATFORM_ORIGIN>` на каждый ответ + снятие
    `X-Frame-Options` (не за флагом — только заголовок).
  - Фронт: `EMBEDDED`-детект, `app-ready`/`platform-auth` через чистый `fetch`,
    лоадер «Вход через платформу…» + 5с fallback на обычный логин, 401 в iframe
    без редиректа. Всё за флагом `PLATFORM_SSO` (default OFF).
  - Проверки: `node --check`, `py_compile`, `require('./keycloakPlatform')`,
    `vite build` — ОК.
  - **Осталось (ручное):** env на Amvera (`PLATFORM_SSO=true`, `KEYCLOAK_URL`,
    `KEYCLOAK_REALM=platform`, `KEYCLOAK_AZP=web-desktop`, `ACCESS_ROLE=resm-user`,
    `PLATFORM_ORIGIN`); в учётках проставить email = email в Keycloak; **на
    стороне платформы SUE_system** (отдельный репо): в `apps.js` запись приложения
    (`badge:true`, `roles:['resm-user','admin']`, `iconUrl`, URL из
    `VITE_APP_RESM_URL`), картинка в `public/apps/`, в Keycloak — realm-роль
    `resm-user` + email пользователям. Пересобрать оба вручную.
