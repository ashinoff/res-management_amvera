# CLAUDE.md — РЭС-менеджмент (res-management)

Контекст проекта для Claude Code. Читается автоматически в начале сеанса.
Карта проекта, задачи переезда/интеграции и уже известные «грабли». Текущее
состояние файла всегда сверяй с кодом — этот файл карта, а не снимок.

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
