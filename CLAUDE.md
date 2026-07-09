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

# ЗАДАЧИ (по порядку, коммитить поэтапно)

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
- **2026-07-09** — репозиторий заведён; в базе — перф-фиксы уведомлений,
  индексы, DB_ALTER, троттлинг фронта. Задачи 1–3 описаны выше, не начаты.
