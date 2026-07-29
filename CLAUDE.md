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

## Система прав (суперадмин + гранулярные права для ЛЮБОЙ роли)

Роли ПО прежние (`admin`/`uploader`/`res_responsible`); поверх них — **суперадмин** и
**гранулярные права**. Права раздаются **любой роли** через раздел «Права доступа»
(с 2026-07-28; раньше — только админам). Роль остаётся отдельным гейтом видимости
меню; право — гейтом конкретного опасного действия.

- **Суперадмин** — учётка `login='admin'` (флаг `User.isSuper`, сид в
  `initializeDatabase()`). Видит и может ВСЁ, правами не ограничивается; в неё же
  подтягивается вход через Платформу (Keycloak) по email. Самозащита: суперадмина
  нельзя удалить/понизить/переименовать НИКОМУ (включая его самого); его учётку
  правит только он сам (смена своего пароля/ФИО/email — можно).
- **Каталог прав** — единственный источник **`PERMISSIONS`** в `backend/server.js`
  (ключ → русское название). Права хранятся в `User.permissions` (JSONB) у любого
  пользователя, в JWT НЕ вшиваются (смена без перелогина; истина — БД + кеш
  `accessCache`, TTL 60с). Есть пары «полное/узкое» право (напр. `structure_edit`
  полностью vs `structure_edit_pu` — только номера ПУ своего РЭС).
- **Как закрыть роут правом:** `requirePerm('<ключ>')` (или `requirePerm(['a','b'])`
  — достаточно любого из ключей). Логика: суперадмин→пропуск; иначе нужен один из
  прав в `permissions`, нет → 403 `{error,permission,title}`. Проверяет ВСЕ роли
  (не только админов). Если роут admin-only — ставь `checkRole(['admin'])` ПЕРЕД
  `requirePerm`; если открыт по праву любой роли — `requirePerm` без checkRole (как
  `PUT /api/network/structure/:id`). Только суперадмин: `requireSuper`.
  `DELETE_PASSWORD` — второй фактор ПОВЕРХ права.
- **Как добавить НОВОЕ право:** (1) ключ+название в `PERMISSIONS`; (2)
  `requirePerm('ключ')` на опасный роут; (3) во фронте скрыть/задизейблить элемент
  через `hasPerm(user,'ключ')` (модульный хелпер в `App.jsx`, работает для любой
  роли). Управление правами: раздел «Права доступа» (только суперадмин, пункт меню
  `superOnly`; в списке — ВСЕ не-суперы: админы/загрузчики/РЭС), роуты
  `GET/PUT /api/admin/permissions[/:userId]`.
- **ИНВАРИАНТ (роль ≠ право):** право проверяется как `isSuper || permissions[key]`
  на бэке (`requirePerm`) И `hasPerm(key)` на фронте — **независимо от роли**. **Роль
  определяет область видимости** (какие РЭС/разделы видны; для не-админа — свой РЭС),
  **право — разрешённые действия**. НИКОГДА не гейтить действие по `role==='admin'`
  поверх права (частая ловушка: скрытый `if (role!=='admin') return` в обработчике —
  так был сломан двойной клик правки ПУ). Право НЕ расширяет географию: не-админ с
  `structure_edit` правит только свой РЭС (проверка `resId` в роуте).
- **ИНВАРИАНТ:** любой НОВЫЙ опасный роут (создание/изменение/удаление данных,
  сервисные операции) по умолчанию закрывается `requirePerm` с ключом каталога.
  Просмотровые (чтение/списки/отчёты/экспорты) — оставлять открытыми. Полная таблица
  роут→ключ — в журнале (запись «Права доступа, коммит 1/3»).

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
   `DELETE_PASSWORD`, `CLOUDINARY_*` (3 шт., переносятся как есть).
   **`JWT_SECRET` и `DELETE_PASSWORD` ОБЯЗАТЕЛЬНЫ** — без любой из них сервер
   падает на старте (`process.exit(1)`), небезопасных дефолтов больше нет.
   Далее — **почта (Яндекс)** — `MAIL_USER` (полный адрес ящика), `MAIL_PASS`
   (**пароль приложения**, не основной пароль ящика), `MAIL_HOST=smtp.yandex.ru`,
   `MAIL_PORT=465`, `MAIL_IMAP_HOST=imap.yandex.ru`, `MAIL_IMAP_PORT=993`;
   приёмник за флагом — `MAIL_INTAKE=true` (default OFF = приёмник не
   запускается, поведение прежнее), опц. `MAIL_INTAKE_ALLOWED` (разрешённые
   отправители через запятую; если пусто — берутся email пользователей из БД),
   `MAIL_INTAKE_INTERVAL_MS` (по умолчанию 60000), `MAIL_FOLDER_PROCESSED/
   ERRORS/REJECTED` (по умолчанию Processed/Errors/Rejected). `PORT=8000`,
   разово `DB_ALTER=true` на первый деплой (создание схемы), потом убрать.
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
- **Почта Яндекс:** `MAIL_PASS` — это **пароль приложения**, не основной пароль:
  id.yandex.ru → Безопасность → Пароли приложений (нужна включённая 2FA). В
  настройках Почты → «Почтовые программы» включить IMAP и вход по паролям
  приложений. **From обязан равняться `MAIL_USER`** — иначе Яндекс режет отправку
  (553 Sender address rejected); все письма шлём через `sendMailAs()` (from
  = `mailFrom()` = `MAIL_USER`, display-name допустим).

## Проверки перед коммитом

- `node --check backend/server.js`.
- `cd frontend && npm run build` (ловит JSX/импорты).
- Анализаторы: `python3 -m py_compile backend/analyzers/*.py`.
- При `PLATFORM_SSO=false` старый вход по паролю не изменился (регресс-минимум:
  login → me → notifications/counts).
- grep: Keycloak-токен нигде не логируется/не сохраняется.

## Журнал изменений (Claude Code ведёт сам)
- **2026-07-29** — Баг прав: uploader с правом не мог открыть модалку правки ПУ.
  **Причина — FRONTEND (слой 2):** в `NetworkStructure.startEdit` был скрытый
  ролевой гейт `if (user.role !== 'admin') return;` ПОВЕРХ права — двойной клик
  вызывал `startEdit`, тот молча выходил для не-админа, что бы ни стояло в правах.
  Фикс: `if (!canEditPu) return;` (`canEditPu = structure_edit || structure_edit_pu`).
  Остальные слои проверены и уже были верны из прошлого коммита: `requirePerm` =
  `isSuper || permissions[key]` (любая роль), кеш прав корректен, `/me`+login+platform
  отдают `permissions` всем ролям, роут `PUT /api/network/structure/:id` — только
  `requirePerm`, без `checkRole`, с ограничением «свой РЭС» для узкого права.
  Заодно (п.3) со смежных structure_edit-роутов снят жёсткий `checkRole(['admin'])`:
  секции `POST/PUT/DELETE /api/network/sections[/:id]` и `POST /api/network/delete-selected`
  теперь гейтятся только `requirePerm('structure_edit')` + область «свой РЭС» для
  не-админа (`resId`-проверка; право не расширяет географию). `clear-all` оставлен
  admin-only (сетевой wipe в admin-only разделе «Обслуживание», не-админу недоступен).
  `saveEdit` показывает реальную ошибку (403 с названием права) тостом. Инвариант
  «роль ≠ право» зафиксирован в разделе «Система прав». node --check / npm run build — ОК.
- **2026-07-28** — Права доступа — для ВСЕХ ролей (не только админов). Спец-костыль
  «загрузчик правит ПУ по умолчанию» убран; вместо него — управляемое право. Backend:
  `requirePerm` теперь проверяет права у ВСЕХ (кроме суперадмина; раньше не-админ
  проходил насквозь) и принимает массив ключей (любой из). Каталог += `structure_edit_pu`
  («Изменение номеров ПУ… без секций/удаления»). `PUT /api/network/structure/:id` —
  без `checkRole`, гейт `requirePerm(['structure_edit','structure_edit_pu'])`; в
  хендлере `fullEdit` (structure_edit/супер) даёт полную правку (вкл. привязку секций),
  иначе (`structure_edit_pu`) — только номера ПУ и только своего РЭС. `GET
  /api/admin/permissions` отдаёт ВСЕХ не-суперов (любой роли) + `role`/`resName`; `PUT`
  снял ограничение «только админам». Frontend: `hasPerm` без привязки к роли
  (`isSuper || permissions[key]`); в «Структуре» `canEditPu = structure_edit ||
  structure_edit_pu`; раздел «Права доступа» показывает роль/РЭС у каждого и
  переименован под «пользователей». Плюс в «Карте опроса» текст «Срез не загружен» →
  «Срез загружается, это может занять несколько секунд, подождите…» (при loading),
  иначе дата или «Срез не синхронизирован». node --check / npm run build — ОК.
- **2026-07-28** — Загрузчики (АСКУЭ) правят структуру в части ПУ — ПО УМОЛЧАНИЮ.
  Backend `PUT /api/network/structure/:id`: `checkRole` расширен до
  `['admin','uploader']`; `requirePerm('structure_edit')` пропускает не-админа
  насквозь → загрузчик по умолчанию (без выдачи права). В хендлере `puOnly` (не
  админ): разрешены ТОЛЬКО номера ПУ и ТОЛЬКО свой РЭС (`structure.resId ===
  req.user.resId`, иначе 403), блок привязки секции `sectionId` пропускается —
  секции/ВЛ/удаление остаются за админом. Frontend: новый флаг `canEditPu =
  canEditStructure || role==='uploader'` — на нём двойной клик правки ПУ и подсказка;
  секции (add/edit/delete/привязка), «Удалить выбранные», FAB и чекбоксы строк
  остаются на `canEditStructure`/`canClearChecks` (загрузчику не видны). node --check
  / npm run build — ОК.
- **2026-07-28** — Карта опроса: правка области видимости (фикс к коммиту 2/3).
  Уточнение: uploader должен видеть карту опроса ТОЛЬКО своего РЭС (а не всю).
  Смысл открытия карты «всем ролям» был лишь в видимости пункта меню без права
  синхронизации, не в расширении области данных. `GET /api/poll-map` scoping
  возвращён к простому правилу: **админ — вся карта (фильтр `?resId`), остальные
  (uploader/res_responsible/uec) — только свой `resId`**. Пункт меню (uploader
  включён) и логика синхронизации не менялись. node --check — ОК.
- **2026-07-28** — Коммит 3/3: анализатор профиля — «дырявый» профиль (ложные пики).
  `backend/analyzers/profile_analyzer.py`: при дыре (один+ ПОДРЯД отсутствующих
  часов на равномерной сетке — пропущенный интервал ИЛИ пустое значение) счётчик
  сваливает потерянный объём в граничные отсчёты → ложные пики. Новые
  `_hour_grid(all_dts)` (полная часовая сетка [min..max] из ВСЕХ строк листа, вкл.
  строки с пустыми значениями — ловит дыры в начале/конце) и `_peak_excluding_gaps`
  (из выбора МАКСИМУМА исключаем ровно 1 заполненный отсчёт до и 1 после каждой
  дыры; общий для двух смежных — раз; дыра в начале → только «после», в конце →
  «до»; всё исключено/пусто → пика нет как при отсутствии данных). `_read_sheet`
  теперь возвращает 4-е значение `all_dts`. Исключение действует ТОЛЬКО на пик
  (peakKw/peakAt и всё производное: overloadStatus/OverloadCase/SectionMonthlyPeak);
  энергия/полнота/формат results не тронуты. В результат по секции добавлены `gaps`
  и `excludedPoints`; одна строка на файл в stderr «[profile] …: секций с дырами K,
  дыр N, исключено M отсчётов». Юнит-прогон `analyzers/test_profile_gaps.py` (без БД):
  пример со спайками, регресс без дыр, дыра в начале/конце, две смежные, сплошные
  дыры, сетка — все зелёные. py_compile / node --check / npm run build — ОК.
- **2026-07-28** — Коммит 2/3: «Карта опроса» доступна всем ролям. Пункт меню
  `poll_map` += `uploader` (роли admin/uploader/res_responsible/uec_responsible).
  `GET /api/poll-map` (и так `authenticateToken`) — scoping изменён: принудительный
  свой РЭС только для `res_responsible`; **admin/uploader/uec_responsible видят всю
  карту** (admin фильтрует через `?resId`). Выбор: карта опроса — сетевой аналитический
  разрез покрытия, uploader (АСКУЭ) работает по всей сети → показываем всё (раньше
  любой не-admin принудительно сужался до своего РЭС). Синхронизация без изменений:
  `POST /api/poll-map/sync` под `pollmap_sync`; `canSync = hasPerm(pollmap_sync) ||
  uec` — у uploader кнопки нет, а в noData-баннере остаётся текст «Дождитесь
  синхронизации администратором» без кнопки. node --check / npm run build — ОК.
- **2026-07-28** — Права/пользователи, коммит 1/3: выпадашка РЭС + защита admin-учёток.
  Диагностика: `GET /api/res/list` открыт всем авторизованным (не gated), а селект РЭС
  в форме пользователя рендерится по `role !== 'admin'` (НЕ по isSuper) — т.е. путь к
  списку РЭС для users_manage-админа уже корректен (работает как у суперадмина). Реальная
  «молчаливо сломанная форма» — редактирование/удаление учёток с ролью **admin**
  не-суперадмином (роль-селект без опции admin → пустой). Фикс: (frontend) `startEdit`
  для admin-строки при не-супере → тост «Учётки администраторов может изменять только
  суперадмин» вместо открытия формы; кнопки ред./удаления в таких строках
  задизейблены с тем же tooltip; alert'ы create/update/delete переведены в `showToast`
  (403 виден внятно). (backend) унифицировано: не-супер НЕ может править ЛЮБУЮ
  admin-учётку (`PUT /api/users/:id` → 403 до изменения полей, закрыт сброс пароля
  чужого админа) и НЕ может удалять admin-учётки (`DELETE` → 403). Создание/назначение
  роли admin — по-прежнему только суперадмин. node --check / npm run build — ОК.
- **2026-07-28** — «Карта опроса»: журнал напряжений = СПОДЭС ИЛИ РиМ (нужен тип ПУ).
  Опрос теперь отдаёт 6-е поле — тип ПУ (`format:3`, коммит в Opros_Piramida
  `85a2b32`). Backend: `PolledMeter += puType` (`ALTER TABLE ADD COLUMN IF NOT
  EXISTS`), sync читает `row[5]`; хелперы `isRimType` (regex `рим|rim`) и
  `hasVoltageJournal(isSpodes,puType)=isSpodes||isRim`. В `/api/poll-map` в Map
  добавлен `puType`, у каждой poll-ячейки новое поле `journal` (доступен ли журнал).
  Кандидаты теперь фильтруются по `hasVoltageJournal` (не только СПОДЭС), в объект
  добавлены `isSpodes/isRim/puType`. Frontend `pollVerdict`: правило «нет журнала»
  теперь по `cell.journal` (не `cell.spodes`) — «Не СПОДЭС и не РиМ → журнал
  напряжений недоступен, плановая замена»; restore/replace-тексты упоминают
  СПОДЭС/РиМ. Модалка кандидатов: заголовок «Кандидаты для контроля (СПОДЭС/РиМ)»,
  колонка «Тип» (бейдж СПОДЭС или `РиМ` `.mini-badge.rim`), пусто →
  «Свободных ПУ с журналом (СПОДЭС/РиМ) нет». Обратная совместимость: пока Опрос на
  5-польном формате, `puType=null` → journal=isSpodes (прежнее поведение), после
  пересборки Опроса РиМ подхватятся. node --check / npm run build — ОК. ⚠️ Порядок:
  «Пересобрать» Опрос → «Синхронизировать» в Мониторинге → «Пересобрать» Мониторинг.
- **2026-07-27** — «Карта опроса», модалка рекомендаций: просторнее + кандидаты
  таблицей. Ширина `.poll-rec-modal` 760→**980px**, паддинги крупнее. Кандидаты
  СПОДЭС теперь **всегда открыты** (сворачивание/`showCandidates` убраны) и выведены
  структурированной таблицей `№ ПУ | Адрес (точка учёта) | Опрос` (`.poll-cand-table`
  в карточке; серийник — кнопка-копирование, столбец «Опрос» = бейдж СПОДЭС +
  собирается/не собирается). Футер сводки починен: элементы стали чипами-пилюлями с
  `gap`/`white-space:nowrap` (были слипшиеся «К замене:2» без пробелов). Только
  фронт (App.jsx/App.css). npm run build — ОК.
- **2026-07-27** — Права доступа, коммит 3/3: самозащита + документация. Закрыт
  обход: не-суперадмин с `users_manage` мог сбросить ПАРОЛЬ суперадмина через
  `PUT /api/users/:id` (эскалация) → добавлена проверка «учётку `isSuper` правит
  только суперадмин». Итог инвариантов самозащиты: суперадмина нельзя удалить
  (isSuper→403 + self-delete→400), понизить/переименовать (никому, включая себя),
  редактировать не-суперадмину; смена собственного пароля/ФИО/email — можно;
  purge/db_tools у не-супер-админа без права → 403 (requirePerm) и скрыты в UI. В
  CLAUDE.md добавлен раздел «Система прав» (где каталог `PERMISSIONS`, как закрыть
  роут `requirePerm`, как добавить новое право = ключ+requirePerm+`hasPerm` в UI,
  инвариант «новые опасные роуты закрываются по умолчанию»). node --check — ОК.
- **2026-07-27** — Права доступа, коммит 2/3 (frontend): раздел «Права доступа» +
  скрытие кнопок. Модульный `hasPerm(user,key)` (isSuper→всё; admin→permissions[key];
  прочие роли не расширяются). Тост `showToast` + интерцептор axios: 403 с полем
  `permission` → жёлто-красный тост «Недостаточно прав: <title>». Пункт меню «Права
  доступа» (`IconLock`, `superOnly`) — только суперадмину (`isSuper` проброшен в
  `MainMenu`). Компонент `PermissionsAdmin`: матрица админы×права из каталога
  (чекбоксы, «Сохранить» на строку, подсказка «без перелогина, в течение минуты»).
  Скрытие кнопок по правам: NetworkStructure (structure_edit — правка ПУ/секций/
  назначение секции/удаление выбранных/FAB; checks_delete — очистка истории ТП),
  Notifications (notifications_delete — чекбоксы/«Удалить выбранные»), UploadedDocuments
  (files_manage — чекбоксы/удаление), PollMap (pollmap_sync — кнопка синхронизации;
  uec остаётся), ExtendedPuModal (checks_delete — «Очистить историю ПУ»). Настройки:
  вкладки гейтятся по праву (structure→structure_upload|edit, diagnose→db_tools,
  maintenance→structure_edit, files→files_manage, database→db_tools|history_purge;
  «Пользователи» всегда — список read-only), default-вкладка = первая доступная;
  UserSettings (users_manage — создать/редакт./удалить; удаление суперадмина скрыто;
  роль «Администратор» в форме — только суперадмину); DatabaseMaintenance (db_tools —
  проверка/бэкап/восстановление; history_purge — очистка до даты). Логика операций не
  менялась; бэкенд остаётся источником истины (403 при прямом вызове). npm run build — ОК.
- **2026-07-27** — Права доступа, коммит 1/3 (backend): суперадмин + гранулярные
  права админов. `User` += `isSuper` (BOOLEAN default false), `permissions` (JSONB
  default `{}`) через `ALTER TABLE ADD COLUMN IF NOT EXISTS` в `initializeDatabase()`
  + идемпотентный сид `UPDATE Users SET isSuper=true WHERE login='admin'`. Каталог
  `PERMISSIONS` (ключ→рус. название) в одном месте: structure_upload, structure_edit,
  checks_delete, notifications_delete, files_manage, history_purge, pollmap_sync,
  users_manage, db_tools. Middleware `requirePerm(key)`: кеш `accessCache`
  (userId→{isSuper,role,permissions,ts}, TTL 60с, `invalidateAccess` при
  сохранении/изменении); суперадмин→пропуск, admin+permissions[key]→пропуск, admin
  без права→403 `{error,permission,title}`, НЕ-админ→насквозь (роли res/uploader/uec
  не ограничиваются — requirePerm стоит ПОСЛЕ checkRole). `requireSuper` — только
  суперадмин. login/platform/me += `{isSuper, permissions}` (в JWT права НЕ вшиты).
  Новые роуты (только isSuper): `GET /api/admin/permissions` (список обычных админов
  + права + каталог), `PUT /api/admin/permissions/:userId` (валидация ключей по
  каталогу, сброс кеша; суперадмина не трогать). users_manage-самозащита: суперадмина
  нельзя удалить/понизить/переименовать; создавать/назначать роль admin может только
  суперадмин. **Инвентаризация роут→право** (опасные/сервисные закрыты, просмотровые
  открыты):
  structure_edit — PUT /network/structure/:id, POST|PUT|DELETE /network/sections[/:id],
  DELETE /network/clear-all, POST /network/delete-selected;
  structure_upload — POST /network/upload-full-structure;
  notifications_delete — DELETE /notifications/:id, POST /notifications/delete-bulk;
  checks_delete — DELETE /history/clear-pu/:pu, POST /history/clear-tp, DELETE
  /history/clear-all; history_purge — POST /admin/purge;
  files_manage — POST /documents/delete-bulk, DELETE /documents/record/:id, DELETE
  /admin/files/:public_id, GET /admin/files/diag/:token;
  pollmap_sync — POST /poll-map/sync (uec_responsible проходит насквозь);
  users_manage — POST /users/create, PUT|DELETE /users/:id;
  db_tools — POST /admin/database-cleanup, GET /admin/backup, POST /admin/restore,
  GET /admin/diagnose/:resId, PUT /admin/fix-notification/:id, POST
  /admin/auto-fix-notification/:id, POST /admin/auto-fix-all/:resId.
  Оставлены ОТКРЫТЫМИ (просмотр/операционные, не в каталоге): GET /users/list,
  /admin/files, /admin/database-health, POST /admin/purge-preview, problem-vl
  list/dismiss/send-email, overload askue/res-complete, notifications mark-read,
  reports/analytics/history GET, upload/analyze (shared с uploader). DELETE_PASSWORD
  сохранён как второй фактор поверх права. Новых ключей каталога не потребовалось.
  node --check — ОК.
- **2026-07-27** — «Уведомления» (Ожидающие мероприятий / Ожидающие АСКУЭ): кнопка
  выгрузки в Excel в строке тулбара (`.notif-toolbar`, справа `margin-left:auto`,
  стиль `.pm-btn--excel`). `exportNotificationsExcel` выгружает именно
  `filteredNotifications` — то, что сейчас на листе (учёт РЭС через resId-загрузку +
  период `periodFilter` + поиск по ТП). Колонки зависят от `filterType`: error —
  РЭС/ТП/ВЛ/Позиция/№ПУ/Фазы с ошибкой/Появилось/Ошибка (фазы через локальный
  `getPhaseErrors`); pending_askue — РЭС/ТП/ВЛ/Позиция/№ПУ/Журнал с/Появилось. Имя
  файла содержит РЭС и период. Только frontend (App.jsx/App.css), бэкенд/логика
  фильтрации не менялись. npm run build — ОК.
- **2026-07-27** — «Карта опроса»: выравнивание квадратов ПУ + модалка/заголовок ТП
  (только вёрстка, логика не тронута). (1) БАГ: квадрат с бейджем «СПОДЭС» был выше
  соседних без бейджа — ряд «плясал». Фикс: `puCell` ВСЕГДА рендерит зону бейджа
  `.pu-badge-zone` фикс. высоты 16px (пусто без СПОДЭС); `.poll-grid` →
  `align-items:start`, `.poll-pu-cell` → `justify-content:flex-start; min-height:0` —
  верх всех ячеек ряда на одной линии, квадраты выровнены независимо от бейджа
  (и «СПОДЭС», и узкий «СПДС» дают одинаковую высоту; `overflow:visible` — glow не
  режется). То же для строк тех.учётов секций. (2) Модалка рекомендаций сделана
  просторнее (как остальные модалки): ширина `min(760px,96vw)`, паддинги 20-22px,
  таблица в карточке (`.poll-rec-tablewrap` — рамка/скругление/тень), ячейки
  12-14px, шапка uppercase на подложке, hover строк, футер — плашка с цветными
  точками-легендой. (3) Заголовок ТП обособлен: `.poll-tp-head` (лёгкая подложка
  surface-2, скругление сверху) + `.poll-tp-name` (акцентный `--navy`, крупнее,
  bold). Мини-сводка ВЛ и вся логика не менялись. Удалён неиспользуемый
  `.spodes-badge` (угловой). npm run build — ОК.
- **2026-07-27** — «Карта опроса», развитие (коммит 4/4): явный СПОДЭС. Компонент
  `SpodesBadge` — слово «СПОДЭС» вместо буквы «С», насыщенный синий (#1d4ed8) +
  двухслойное СТАТИЧНОЕ неоновое свечение (`.spodes-neon`, box-shadow в стиле иконок
  PageHeader, усиление на hover строки/чипа, transition .12s, БЕЗ пульсаций —
  инвариант). Применён везде: квадраты ПУ ВЛ (бейдж перенесён из угла квадрата под
  номер), тех.учёты, таблица модалки рекомендаций, кандидаты, список «В опросе, но
  нет в структуре». На узких экранах (≤640px) сокращается до «СПДС» (два span
  `sb-full`/`sb-short`, переключение media-query). Только frontend. npm run build — ОК.
- **2026-07-27** — «Карта опроса», развитие (коммит 3/4): мини-сводка ВЛ + модалка
  рекомендаций + кандидаты СПОДЭС. Backend `GET /api/poll-map` += `tps`
  (по каждой уникальной ТП структуры: `tpMatched` — нашлась ли в срезе по
  `normTpName`, `candidates` — ПУ среза этой ТП, СПОДЭС, чей serialNorm НЕ в
  `structSerials`; поля serial/tuPath/isCollected; один проход `Map tpNorm→список`)
  и флаг `tpDataAvailable` (срез отдал поле tp). Frontend `PollMap`: (1) под именем
  ВЛ мелкая строка «опрос: X/Y · СПОДЭС: K из Y» (Y — ПУ с номерами, X — собирается,
  K — СПОДЭС; X/Y зел/янт/красн, Y=0 «ПУ не заданы»); при noData скрыта. (2) Строка
  ВЛ/секции кликабельна (`.poll-row-click`, hover .12s) → `ModalShell` (read-only,
  док/fullscreen) «<ТП> — <ВЛ/СШ> · рекомендации»: таблица позиция/№ПУ/статус/Вывод.
  Выводы — константы `POLL_VERDICTS` (`pollVerdict`, первое правило): absent→красн
  «заменить» (пр.1); найден не-СПОДЭС при ЛЮБОМ сборе→янт «плановая замена» (пр.2);
  СПОДЭС+не собирается→янт «восстановить опрос» (пр.3); no_pu→серый «заполнить»
  (пр.4); СПОДЭС+собирается→зел «в порядке». Сортировка по приоритету; футер «К
  замене (replace+planReplace) · Восстановить · Заполнить · В порядке». При noData —
  состав без выводов + подсказка синхронизировать. (3) В модалке collapsible
  «Кандидаты СПОДЭС на этой ТП (N)»: серийник (клик копирует), tuPath, бейджи; пусто
  → tpMatched? «Свободных СПОДЭС нет» : «ТП не найдена в срезе по имени». Блок скрыт,
  если `tpDataAvailable=false` (Опрос ещё на 3-польном формате). node --check /
  npm run build — ОК.
- **2026-07-27** — «Карта опроса», развитие (коммит 2/4): приём расширенного среза
  + хранение ТП (ОБРАТНО СОВМЕСТИМО, Опрос задеплоят позже). Элемент `meters`
  интеграции теперь массив из 3 ИЛИ 5 полей `[serial, spodes01, collected01, tp,
  tu_path]`; валидация `bad_payload` = «длина >= 3» (уже была `< 3`), поля 4-5
  читаются при наличии/непустоте, иначе null. `PolledMeter` += `tp` (STRING),
  `tpNorm` (STRING, индекс `idx_polledmeter_tpnorm`), `tuPath` (TEXT) — через явные
  `ALTER TABLE ... ADD COLUMN IF NOT EXISTS` в `initializeDatabase()` (`sequelize.sync`
  колонки не добавляет). Общая `normTpName()` (верхний регистр, без пробелов/дефисов/
  подчёркиваний, срез ведущего `ТП`/`TP`) — применяется одинаково к именам ТП
  структуры и к `tp` из среза. `sync` пишет новые поля (null при старом формате;
  при дубле serialNorm дозаполняет пустые tp/tuPath). Только backend. node --check — ОК.
- **2026-07-27** — «Карта опроса», развитие (коммит 1/4): тех.учёты секций в карте.
  Backend `GET /api/poll-map`: третьим `findAll` тянется `TpSection` (тот же
  `where` по resId, `include ResUnit`); `sectionRows` = {id, resId, resName, tpName,
  sectionNumber, tnKva, techPuNumber, poll:statusOf(techPuNumber)} — статус
  тех.учёта сверяется тем же Map по serialNorm, без запросов в циклах; серийники
  тех.учётов добавлены в `structSerials` (исключаются из orphans). Тех.учёты входят
  в сводку покрытия по РЭС и ИТОГО (`ensureAgg` + `acc` по section.poll). В ответ
  добавлено поле `sections`. Frontend `PollMap`: группировка РЭС→ТП теперь
  `{secs, vls}` (ТП появляется даже если у неё только секции); секции рендерятся в
  tp-card строками «СШ-<рим> · <кВА> · тех.учёт № …» с тем же квадратом-индикатором
  (`.poll-section-line`), фильтры-легенда действуют и на секции (`secFiltered`).
  Excel-лист «Карта»: колонка **«Тип»** (ПУ ВЛ / тех.учёт), тех.учёты строками
  (ВЛ=«СШ-N», Позиция=«ввод»); Excel/пустой-фильтр учитывают и секции. node --check
  / npm run build — ОК.
- **2026-07-27** — «Карта опроса», структура видна ДО синхронизации (коммит 2).
  Проверено: при пустой `PolledMeter` backend уже отдавал дерево, НО статус ПУ был
  `absent` (красный) и сводка — реальными нулями; поэтому коммит нужен. Backend
  `GET /api/poll-map`: при `noData` каждый ПУ получает нейтральный статус `no_data`
  (не `absent` = «есть срез, ПУ в нём нет»); `acc` игнорирует `no_data`; сводка
  (`byRes`+`total`) отдаётся «прочерками» (все числовые поля + `coveragePct` → `null`),
  список РЭС и дерево структуры сохраняются; `noData:true`. Frontend `PollMap`:
  полноэкранная заглушка убрана — дерево видно всегда; при `noData` сверху жёлтый
  баннер «Срез не синхронизирован…» с кнопкой «Синхронизировать» (только admin/uec),
  иначе подсказка «Дождитесь синхронизации администратором». `no_data` → нейтральный
  контурный индикатор (`.status-nodata`, пунктирная рамка) с tooltip «Нет данных
  среза»; легенда-фильтры при `noData` полупрозрачны (`.status-legend.is-muted`);
  числа сводки через `fmtN` (null→«—»), `% сбора` при null→«—» без цветокласса; Excel
  `stRu`/`statusColor` знают `no_data` (серый «Нет данных»). После успешного sync
  данные перезагружаются `load()` без перезахода (уже было). node --check /
  npm run build — ОК.
- **2026-07-27** — «Карта опроса», диагностика ошибок sync (коммит 1). Backend
  `POST /api/poll-map/sync` вместо общих 502/503 отдаёт структурированный
  `{ error, code, hint, detail }`, различая стадии: `not_configured` (нет
  OPROS_URL/OPROS_API_KEY — hint перечисляет какой именно, 503),
  `dns_or_network` (fetch упал до ответа — ENOTFOUND/ECONNREFUSED/таймаут-abort
  60с; hint c текущим OPROS_URL, detail=код Node, 502), `upstream_401`
  (ключи не совпадают — подсказка про равенство OPROS_API_KEY↔INTEGRATION_API_KEY
  и пересборку), `upstream_503` (в Опросе не задан INTEGRATION_API_KEY),
  `upstream_404` (старая версия Опроса без endpoint), `upstream_other` (detail =
  статус + первые 200 симв. тела), `bad_payload` (200, но JSON не распарсился/нет
  массива meters/элемент не тройка), `empty_snapshot` (**200 c warning**, НЕ
  ошибка: существующий непустой PolledMeter НЕ затирается — `count()` до записи,
  ранний выход без destroy; hint зависит от наличия прежнего среза). Ключ
  OPROS_API_KEY нигде не светится (ответ/лог); OPROS_URL — можно. В лог сервера —
  стадия+статус одной строкой (`[poll-map/sync] <code> …`). Успех теперь
  `{ code:'ok', total, collected, spodes, snapshotAt }`, db-ошибка → `code:'db_error'`
  (прежний срез сохранён транзакцией). Frontend `PollMap.sync`: успех — баннер
  «Синхронизировано: N ПУ, из них собирается M, СПОДЭС K, срез от <дата>»;
  `empty_snapshot` → жёлтый баннер (`.poll-notice.warn`); ошибка — hint человеку +
  мелким `code · detail` (`.poll-notice-tech`) для передачи программисту. Баннер
  стал колоночным (`.poll-notice-main`/`-tech`). node --check / npm run build — ОК.
  ⚠️ Полная эмуляция стадий на живом сервере не гонялась локально (нет Postgres) —
  ветвление проверено ревью + сборкой; прогнать на Amvera после деплоя.
- **2026-07-26** — «Карта опроса», коммит 2 (frontend). Пункт меню «Карта опроса»
  (`IconMapPin`, роли admin/res_responsible/uec_responsible), роут `poll_map`,
  компонент `PollMap`: PageHeader (неоновая иконка, без внешней рамки). Верхняя
  панель: «Срез Пирамиды от <дата>» + поиск по ТП + «Синхронизировать» (только
  admin/uec, спиннер, итог/ошибка баннером `.poll-notice`) + Excel; фильтр РЭС —
  через глобальный selectedRes. `noData` → заглушка. Сводка по РЭС + ИТОГО
  (таблица `.pm-matrix`, % сбора цветом ≥90 зел/≥70 янт/иначе красн). Дерево
  РЭС→ТП→ВЛ (read-only) с 3 квадратами ПУ: collected→зел, not_collected→янт,
  absent→красн, no_pu→серый X; бейдж «С» (СПОДЭС) на зел/янт. Легенда-фильтры
  (4 состояния) + паттерн «полоса→кружки при скролле» (реализован инлайн, как в
  «Структуре»). Collapsible «В опросе, но нет в структуре (N)» (свёрнут). Excel —
  `styleExportSheet`: листы «Карта»/«Сводка»/«Не в структуре». npm run build — ОК.
  **Напоминание:** в env Мониторинга на Amvera добавить `OPROS_URL`/`OPROS_API_KEY`,
  пересобрать оба приложения.
- **2026-07-26** — «Карта опроса», коммит 1 (backend). Env `OPROS_URL`/`OPROS_API_KEY`
  (без них `/api/poll-map/sync` → 503, остальное работает). Модель `PolledMeter`
  (serialRaw, serialNorm[индекс `idx_polledmeter_norm`], isSpodes, isCollected,
  snapshotAt) — таблица через `sync()`. `normSerial` = trim/без пробелов/без ведущих
  нулей. `POST /api/poll-map/sync` (admin+uec_responsible): fetch
  `${OPROS_URL}/api/integration/meters` (X-Api-Key, таймаут 60с), формат
  `{snapshot_at,count,meters:[[serial,spodes01,collected01],…]}`; дедуп по serialNorm
  (collected>spodes), в транзакции очистка+`bulkCreate` батчами 1000 (ошибка апстрима
  502 понятным текстом, старый срез не затирается); ответ `{total,collected,spodes,
  snapshotAt}`. `GET /api/poll-map` (auth; res_responsible — свой resId): структура
  (`NetworkStructure`+ResUnit) × срез одним findAll в Map по serialNorm → per-ПУ
  `status` collected/not_collected/absent/no_pu + spodes; сводка по РЭС и итого
  (totalPu/collected/notCollected/absent/spodes/noPu/coveragePct); `orphans` (серийники
  среза не в структуре); `snapshotAt`, `noData`. Без запросов в циклах. Роль
  `uec_responsible` в enum нет — в checkRole указана на будущее. node --check — ОК.
- **2026-07-26** — «Уведомления» (Ожидающие мероприятий/проверки): в строку
  «Выбрать все» (`.notif-toolbar`, теперь рендерится всегда — чекбокс только у
  админа) добавлены: счётчик «Показано: N [из M]», фильтр по периоду (месяц
  появления уведомления по `createdAt`, YYYY-MM) и инфо «В наличии: mm.yyyy·кол-во»
  по всем текущим уведомлениям. `filteredNotifications` доп. фильтруется по
  `periodFilter` (только этот раздел). CSS `.notif-toolbar/.notif-count/.notif-
  period/.notif-periods-info`. npm run build — ОК.
- **2026-07-26** — UI + статистика БД. (1) Кнопки под файлами (документы и
  «Управление файлами») — фон прозрачный, светится сам SVG (`.btn-view/.btn-icon`
  → transparent, иконки `.ico-glow-blue/-red`); «посмотреть» больше не синяя.
  (2) Статус документа — само СЛОВО цветом (Завершен→зелёный, На проверке→оранжевый
  `#d97706`), фон прозрачный (пилюля убрана). (3) Mac-светофор на ВСЕХ модалках ~в 2
  раза крупнее (`.tl` 13→24px, иконки 9→15, gap 8→11). (4) «Управление файлами»:
  фильтры по РЭС (select из файлов) и по периоду загрузки (date from/to) + кнопка
  «наверх» (скролл `.content`). (5) Настройки → Обслуживание → «Статистика базы
  данных»: новый блок **«Занятое место в базе»** — размер БД (`pg_database_size`),
  новых записей за 30 дней, прирост/мес (≈ по байт-на-строку), топ-8 таблиц по
  размеру (`pg_total_relation_size`), и оценка «хватит ещё ~N мес» если задан env
  `DB_QUOTA_MB` (квота БД в МБ). node --check / npm run build — ОК.
- **2026-07-26** — UI-пакет правок (frontend + мелкий backend). (1) Убрана кнопка
  «Выйти» из шапки (вход через SUE-платформу). (2) «Структура»: селект секции
  менее выделен (`.vl-section-select` прозрачный, проявляется на hover/focus);
  иконки ред./удал. секции — светящиеся (`.ico-glow-blue/.ico-glow-red` — свечение
  самого SVG через drop-shadow); квадратик техучёта красится по загрузке
  (оранжевая зона 85–100% → оранжевый `status-pending`, не зелёный); «Обновить
  структуру» → стиль `.pm-btn--refresh` (как в Анализе мощности) + IconRefresh;
  кружки-фильтры ярче, без рамки/панели, по оси нижних FAB (right 30, ширина 50,
  центр). (3) Тех-модалка: убраны слово Перегруз/Норма и цвет шапки; заголовок =
  ТП · РЭС + цветной %; в теле после ТП — строка РЭС · загрузка %. Для РЭС в
  `/api/network/sections` добавлен `include ResUnit`. (4) Анализ мощности/напряжения:
  больше отступ после «ТП с перегрузом» и «ВЛ в работе у РЭС» (`marginTop 18px` у
  подписи). (5) «Загруженные документы»: статус «На проверке» → оранжевый
  (мак-жёлтый `#febc2e`), глазик — синее свечение, удалить — красное; те же
  glow-классы в «Управлении файлами». (6) Настройки → База данных: карточки
  прозрачные (`.db-header` bg transparent + рамка, тёмный текст), иконки по центру
  и синие со свечением, кнопки `.btn-check-db` в стиле `.pm-btn--refresh` (синий
  контур→заливка, .12s). node --check / npm run build — ОК.
- **2026-07-26** — «Структура сети»: липкая полоса легенды-фильтров → адаптив при
  скролле (только этот раздел, логика фильтрации не менялась). Убран `position:
  sticky` у `.status-legend` (из-за него между шапкой и полосой была щель). При
  скролле `.content` за порог `COLLAPSE_AT=130` полоса плавно растворяется
  (`.is-collapsed` opacity+translateY), а справа-сверху контента появляется
  фиксированная колонка кружков `.legend-dots` (top 84/right 22, z-index 1200 —
  ниже дока 1400 и модалок 1500). Кружки — тот же `statusFilter` (один источник):
  3 кликабельных (зел/красн/янтарь) с активным кольцом цвета статуса + приглушением
  неактивных, 2 индикатора (серый/«X») как disabled-строки полосы (не кликабельны,
  чтобы не менять поведение фильтрации), tooltip через `title`. Обратно вверх —
  гистерезис `EXPAND_AT=80` (~50px, без мигания). Слушатель на `.content` (как у
  кнопки «наверх»), `passive:true` + `requestAnimationFrame`, setState только при
  пересечении порога (функц.-апдейт, React бэйлит). Переход `.25s` — **осознанное
  исключение** из инвариантных .12s для крупного transition (зафиксировано). Мобилка:
  кружки меньше (18px), ближе к краю. npm run build — ОК.
- **2026-07-26** — Редактирование ПУ в «Структуре» — из инлайн-кнопок (галочка/
  крестик, стали неактивны/без иконок) в mac-модалку `ModalShell`. Двойной клик по
  номеру ПУ открывает `puEdit` ({item,position,oldValue}): «Старый номер ПУ»
  (readonly) / «Новый номер ПУ» (input, Enter=сохранить) + Отмена/Сохранить
  (Сохранить disabled, пока не изменено; пусто → убрать ПУ из позиции). Логика
  сохранения та же (`PUT /api/network/structure/:id`), переписана под модалку;
  инлайн-блок `edit-cell/save-btn/editingCell` удалён. CSS `.pu-edit-grid`
  (2 колонки, на узком → 1). npm run build — ОК.
- **2026-07-26** — Оконная оболочка модалок, ФАЗА 1 (движок + ключевые модалки;
  согласовано с владельцем — остальные переводятся тем же паттерном следующими
  заходами). Добавлены: `ModalShell` (backdrop+окно+строка заголовка с mac-
  «светофором»: 🔴 закрыть=onClose, 🟡 свернуть в док, 🟢 fullscreen
  `calc(100vw/vh−24px)` c переходом .12s; иконки в кнопках проявляются на hover
  группы; confirm-вариант — только 🔴), `ModalDockProvider`/`ModalDockCtx`/
  `ModalDockBar` (реестр свёрнутых окон, док внизу, политика «одно активное окно»:
  открытие/разворот сворачивает предыдущее; confirm не участвует; unmount
  снимает плашку → смена раздела чистит док), Esc закрывает активную. `PdfCanvas`
  получил `ResizeObserver` → в fullscreen канвас честно перерисовывается под
  ширину (не CSS-растяжение). Дерево обёрнуто в `ModalDockProvider`. CSS
  `.modal-shell/.traffic-lights/.tl-*/.modal-dock*`. **Аудит модалок (~30, все
  по паттерну `modal-backdrop→modal-content→modal-header`):** confirm (только
  🔴) — удаления с паролем/подтверждения (NetworkStructure/Reports/ProblemVL/
  FileManagement/purge/cleanup/clear-history ≈13 шт); обычные — тех-учёт, секция,
  ErrorDetailsModal, детали/мероприятия перегруза, комментарий, детали/e-mail
  ВЛ, mass-fix/fix, создание/редакт. пользователя, просмотр документов ≈17 шт.
  **Переведены в этой фазе:** тех-учёт, секция (add/edit), удаление выбранных
  (confirm), диагностика файла, **FileViewer** (fullscreen-PDF, используется в
  Reports/FileManagement/UploadedDocuments/PowerOverload — покрыт разом). Остальные
  ⬜ — следующими заходами. npm run build — ОК.
- **2026-07-26** — Оконная оболочка, ЗАВЕРШЕНИЕ: переведены ВСЕ оставшиеся модалки
  на `ModalShell` (31 использование; `modal-backdrop`/`close-btn`/`modal-header` в
  App.jsx = 0). Обычные (полный светофор): ErrorDetailsModal, детали проверки/
  подробная информация, мероприятия, комментарий, детали случая/действие перегруза
  (со статус-пиллой в `titleExtra`), детали/e-mail проблемной ВЛ, mass-fix/fix,
  создание/редакт. пользователя, extended-pu, просмотр документов (FileViewer).
  Confirm (только 🔴, не в доке): удаления с паролем во всех разделах,
  очистка истории/данных, purge, cleanup, удаление файлов/записей (bulk и по
  одному). Идентичные подвалы FileManagement/UploadedDocuments-delete закрыты одним
  `replace_all`. Пересборка после каждой партии — зелёная. npm run build — ОК.
- **2026-07-26** — Убран дубль SheetJS из бандла (коммит 2). Во фронте были и
  `xlsx`, и `xlsx-js-style` (drop-in форк с тем же API). Заменил
  `import * as XLSX from 'xlsx'` → из `xlsx-js-style`, слил `XLSXStyle`→`XLSX`
  (один импорт, одно имя, 13 использований), удалил `xlsx` из package.json
  (обновил lock). Чтения xlsx во фронте нет (только запись/стили) — регрессов
  формата не ждём. Осн. JS-чанк: **1 475.82 → 1 042.07 kB** (−433 kB минифиц.;
  gzip 585.91 → 442.23, −143 kB). Backend `xlsx` не тронут. npm run build — ОК.
- **2026-07-26** — Аутентификация файлового прокси (коммит 1). `/api/f` и
  `/api/download` были ОТКРЫТЫ — закрыты. Backend: `makeFileToken(user)` —
  отдельный JWT `{id,role,resId,scope:'files'}` на 24ч; возвращается в
  `/api/auth/login`, `/api/auth/platform`, `/api/auth/me` как `fileToken` (в URL
  не светим сессионный JWT). Middleware `authenticateFileAccess` на обоих
  прокси-роутах: принимает `Authorization: Bearer` ИЛИ `?t=<jwt>`, verify тем же
  секретом (годится и сессионный, и scope 'files'); нет/протух → 401. diag-роут
  (admin-only) не тронут. Frontend: `fileToken` сохраняется рядом с `token` при
  login/platform/me и чистится везде, где чистится `token` (логаут, 401-интерсептор,
  inactivity, EMBEDDED). `fileProxyUrl` добавляет `?t=<fileToken>` ко всем ссылкам
  (t в query, в JSON-токен не вшит — ссылки протухают с сессией); просмотр через
  axios не менялся (Authorization уже уходит). Ранее разосланные прямые ссылки
  без входа перестанут открываться — это цель фикса. Осн. JS-чанк ДО дедупа
  SheetJS: **1 475.82 kB** (gzip 585.91). node --check / npm run build — ОК.
- **2026-07-26** — PDF-предпросмотр без фрейма: pdf.js в `<canvas>`. Причина: во
  iframe Платформы Яндекс Protect режет ЛЮБУЮ навигацию вложенных фреймов, включая
  `<iframe src=blob:>` (ERR_BLOCKED_BY_CLIENT). Зависимость `pdfjs-dist@4.10.38`,
  подключается **лениво** (`await import('pdfjs-dist')` при первом открытии PDF) —
  ушла в отдельные чанки `pdf-*.js`/`pdf.worker.min-*.mjs`, основной бандл не вырос.
  `workerSrc = new URL('pdfjs-dist/build/pdf.worker.min.mjs', import.meta.url)`.
  Новый компонент `PdfCanvas`: из уже загруженного blob → `arrayBuffer` →
  `getDocument({data})`, рендер ВСЕХ страниц вертикальной лентой (canvas), масштаб
  по ширине × `devicePixelRatio` (текст чёткий); тулбар «стр. N из M» + зум ±
  (пересчёт рендера); лоадер парсинга; ошибка → сообщение + «Скачать»; при
  закрытии/смене файла `doc.destroy()` + revoke. В `FileViewer` PDF-ветка
  `<iframe>`→`<PdfCanvas>`; картинки уже на blob-`<img>`. Грепом: `<iframe>/<object>/
  <embed>` в App.jsx нет; `console.log` из вьюера убран. CSS `.pdf-canvas-viewer/
  .pdf-toolbar/.pdf-pages/.pdf-page-canvas`. npm run build — ОК, pdfjs в ленивом чанке.
- **2026-07-26** — Файловый токен без query-строки, коммит 2. Backend `parseFileToken`
  (в `handleFileProxy`): (1) base64url(JSON `{p,n,i}`) — новый формат без query;
  (2) base64url(строка `res-management/…`) — прежний, name/inline из query;
  (3) сырой public_id (старые ссылки/`/api/download`). Все форматы живы (проверено
  юнит-парсером). Frontend `fileProxyUrl` собирает `toBase64Url(JSON.stringify(
  {p:public_id,n:name,i:0|1}))` без query. Diag не затронут (шлёт base64url строки,
  разбирается `resolvePublicId`). node --check / npm run build — ОК.
- **2026-07-26** — Просмотр вложений БЕЗ навигации (Яндекс Protect режет
  навигацию на `/api/f/...&inline=1`, фоновые XHR — нет), коммит 1. `FileViewer`
  переписан: файл тянется `api.get(fileProxyUrl(f,true),{responseType:'blob'})`,
  `URL.createObjectURL` → картинка в `<img>`, PDF в `<iframe class=pdf-frame>`
  ВНУТРИ модалки; revoke при закрытии/смене файла (effect по `public_id`), лоадер,
  ошибка + бейдж «файл утерян» на 404. Убраны `console.log` и «Открыть в новой
  вкладке»; «Скачать» — обычные `<a download>` через `fileProxyUrl(currentFile)`
  (без inline, ручная сборка ссылки убрана). Новый `BlobImage` (тот же blob-XHR)
  заменил subresource-`<img src={fileProxyUrl}>` в превью «Управления файлами».
  Навигации-ссылки на просмотр переведены на открытие FileViewer: «Открыть» в
  управлении файлами и вложения в хронологии перегруза (`PowerOverload` получил
  собственный FileViewer + `.po-attach-link`). Грепом подтверждено: `target=_blank`
  +fileProxyUrl / window.open / location.href на `/api/f` — нет; остались только
  blob-XHR и `<a download>`. CSS `.pdf-frame/.pdf-viewer-blob/.blob-img-fallback`.
  npm run build — ОК.
- **2026-07-26** — Файловый прокси, «доставка», коммит 2: `handleFileProxy` берёт
  URL у Cloudinary, а не угадывает. Порядок: (0) кешированный рабочий URL → (1)
  базовая `signed(primary/upload)` → (2) **`api.resource(secure_url)` для primary,
  затем alt** (+ `signed(...version)` из того же ответа) → (3) старая цепочка:
  alt/upload, image/raw authenticated, image no-ext+format. Первый `resp.ok`
  стримится. Кеш стал `publicId → рабочий URL` (строка): новые шаги в него
  попадают. Лог неудачи — одной строкой со `status` и `x-cld-error` каждого
  варианта. Убраны неиспользуемые `buildFileVariants/fileVariantUrl/sameVariant`.
  Замечание: успешный `secure_url` может отдавать файл, пока старый угаданный URL
  ещё возвращает закешированную CDN-ошибку — это норма. node --check / npm run
  build — ОК.
- **2026-07-26** — Файловый прокси, «доставка», коммит 1: diag читает причину
  отказа доставки. `cloudinaryProbe` для НАЙДЕННОЙ через `api.resource` комбинации
  дополнительно fetch'ит 4 delivery-URL и возвращает `deliveryTests[{label,status,
  xCldError}]`: (a) текущий прокси-URL `cloudinary.url(...sign_url:true)`, (b) он же
  + `version` из ответа resource, (c) `resource.secure_url` как есть, (d) без
  `sign_url`. Читается заголовок **`x-cld-error`** (официальная причина отказа
  Cloudinary). Фронт «Управление файлами» → модалка диагностики: таблица
  «вариант/HTTP/x-cld-error» + примечание, что CDN кеширует 40x-ошибки до ~24 ч
  (asset может существовать, а старый delivery-URL ещё отдавать закешированный
  отказ). node --check / npm run build — ОК. **TODO:** прогнать diag по
  `res-management/attachment_1784876684686__824_*.pdf`, записать статусы/x-cld-error.
- **2026-07-26** — «Настройки → Управление файлами»: чекбоксы + удаление выбранных
  (только фронт, App.jsx/App.css). Отдельное состояние `checkedIds` (не путать с
  `selectedFiles` просмотрщика), чекбокс на карточке (`.file-check`, выделение
  `.file-card.checked`), панель `.file-bulk-bar` с «Выбрать все показанные»
  (по текущему фильтру) и «Удалить выбранные (N)»; модалка с паролем →
  `handleBulkDelete` циклом по существующему `DELETE /api/admin/files/:public_id`
  (пароль = DELETE_PASSWORD; 403 → «Неверный пароль», модалка не закрывается),
  затем `loadFiles()` и сводка. Бэкенд не менялся. npm run build — ОК.
- **2026-07-26** — Файловый прокси, коммит 3: устойчивый `handleFileProxy`. Вместо
  одной попытки — цепочка: (а) resource_type по расширению + upload → (б) альт.
  resource_type → (в) authenticated для image и raw → (г) для image: public_id без
  расширения + format из расширения; все — signed URL, первая успешная стримится.
  Найденная комбинация кешируется `fileVariantCache` (publicId→variant, без TTL) —
  кешированная пробуется первой (дедуп), чтобы не бить 4-5 запросов на каждое
  открытие. Все 404 → `404 {error:'Файл отсутствует в хранилище', publicId}` + лог
  одной строкой со списком испробованного. Frontend «Управление файлами»: `onError`
  превью помечает файл утерянным → бейдж «Файл утерян — удалите запись» (вместо
  алерта), удаление существующей кнопкой; листинг `/api/admin/files` не менялся.
  CSS `.file-lost-badge`. node --check / npm run build — ОК.
- **2026-07-26** — Файловый прокси, коммит 2: диагностика. Backend
  `GET /api/admin/files/diag/:token` (admin, token — base64url/сырой через
  `resolvePublicId`): `cloudinaryProbe` проверяет ресурс во ВСЕХ комбинациях
  resource_type(image|raw) × type(upload|authenticated) × public_id(как есть|без
  расширения) → JSON `{publicId, ext, found:{resource_type,type,idVariant,bytes,
  format,created_at}|null, results:[…]}`. Frontend «Управление файлами»: кнопка
  «Диагностика» (IconSearch) у каждого файла → модалка с найденной комбинацией и
  матрицей результатов (OK/HTTP-код). CSS `.diag-table`. node --check / npm run
  build — ОК. **TODO (после деплоя):** прогнать 2-3 проблемных файла диагностикой,
  записать сюда рабочую комбинацию/отсутствие — это первопричина 404.
- **2026-07-26** — Файловый прокси, коммит 1: base64url-ссылки + защита. Проблема:
  блокировщики матчились на 'attachment_' внутри public_id, видимого в
  `/api/f/res-management%2Fattachment_…`. Backend: `resolvePublicId(raw)` —
  `Buffer.from(raw,'base64url')`; принимаем только если декодировалось в строку с
  префиксом `res-management/`, иначе трактуем raw как сырой public_id (старые
  ссылки живут). В `handleFileProxy` убран лишний `decodeURIComponent` (Express уже
  декодирует; падал на литеральном '%'); добавлена 403-защита от открытого прокси
  (publicId обязан начинаться с `res-management/`) — на обоих путях (`/api/f`,
  `/api/download`). Frontend: `toBase64Url(str)` (TextEncoder→btoa, `+/`→`-_`, без
  `=`), `fileProxyUrl` и прямая ссылка «Скачать» строят `/api/f/<base64url>`.
  Проверено: roundtrip ок; сырой public_id уходит в fallback (не ложно-base64url).
  Тонкий белый кантик активного пункта левого меню (`inset ring`). node --check /
  npm run build — ОК.
- **2026-07-26** — Унификация шаблона страниц (App.jsx/App.css, только вёрстка/
  стили, логика не тронута). **Аудит внешних рамок:** эталонные (контент на фоне,
  без рамки) — «Структура сети» (`.network-structure`), «Отчёты» (`.reports`),
  «Превышение Pном» (`.power-overload-page`), «Анализ мощности» (`.analytics`),
  «Уведомления» (`.notifications`); с лишней внешней «формой» (surface+radius16+
  shadow[/padding]) — «Загрузить файлы» (`.file-upload-container`), «Проблемные ВЛ»
  (`.problem-vl-container`), «Настройки» (`.settings-container`), «Загруженные
  документы» (`.uploaded-documents`), «Анализ напряжения» (`.analytics-container`),
  «История системы» (`.system-history`, фон surface-2+padding). **Сделано:**
  (1) новый компонент `PageHeader({icon,title})` — иконка в квадратной плашке +
  H-заголовок, применён на ВСЕХ страницах; иконки = как в левом меню (Структура→
  Layers, Загрузка→Upload, Ожидающие мероприятий→Wrench, АСКУЭ→Clipboard, Проблемные
  ВЛ→AlertTriangle, Превышение Pном→Zap, Анализ мощности/напряжения→Chart,
  Документы→Folder, История→Clock, Отчёты→FileText, Настройки→Settings). (2) Иконка
  шапки — **статичная неоново-синяя подсветка** (`.page-header-icon`: цвет #2563eb,
  box-shadow 2 слоя `0 0 8px/.55` + `0 0 20px/.30`, тонкая синяя рамка; на hover
  чуть ярче, transition .12s; БЕЗ infinite/pulse). (3) Единые отступы шапки
  (margin-bottom 24px) + `.content padding-top 22px` — на всех страницах одинаково.
  (4) Внешние рамки сняты правилом в конце App.css (перебивает исходные): у 6
  классов-обёрток `background:transparent; box-shadow:none; border-radius:0;
  padding:0; overflow:visible` — блоки легли на фон страницы, внутренние карточки
  не тронуты. Красные рамки-иконки двух разделов заменены единой неоновой (по
  задаче унификации). Мобилка: `.page-header` flex-wrap. npm run build — ОК.
- **2026-07-26** — UI-пакет (App.jsx/App.css): (1) «Аналитика»→«**Анализ
  напряжения**», пункт меню перенесён вплотную к «Анализ мощности» (два отчёта
  подряд); заголовок страницы тоже «Анализ напряжения». (2) Рамки-иконки в
  заголовках обоих разделов — красные (`.svg-frame--red`). (3) **Все иконки
  левого меню — неоново-синие** со свечением (`.main-menu .menu-ico svg`
  color #38bdf8 + drop-shadow, ярче на hover/active). (4) Фон сайдбара —
  диагональный градиент синий→чёрный (`linear-gradient(150deg, navy-500 → navy →
  #0a1524 → #05070c)`). (5) Верхний отступ страниц восстановлен
  (`.content padding-top 0→22px`) — иконки страниц не липнут к шапке. (6)
  Страница «Анализ мощности»: кнопки приведены к единому формату/размеру
  (`.pm-btn`, высота 38px), «Выгрузка в Excel» — всегда зелёная
  (`.pm-btn--excel`), «Обновить» — синяя-контур→залив; пресеты/инпуты периода
  выровнены по высоте 38px, отступы крупнее. Инварианты соблюдены (переходы
  .12s, без infinite, без hover-transform). npm run build — ОК.
- **2026-07-26** — ЗАДАЧА Б: очистка истории до даты (admin). Два эндпоинта:
  `POST /api/admin/purge-preview {before:'YYYY-MM-DD'}` — только COUNT
  ({uploadHistory, puUploadHistory, checkHistory, notifications, cloudinaryFiles});
  `POST /api/admin/purge {before, password}` (password=DELETE_PASSWORD) — в
  транзакции: (a) собрать public_id из CheckHistory.attachments старше before;
  (b) удалить NotificationRead(по id уведомлений)→Notification(createdAt<before, все
  типы)→CheckHistory→UploadHistory→PuUploadHistory(uploadedAt<before) — все через
  `model.destroy` (хук инвалидации кэша counts срабатывает); (c) commit, ПОСЛЕ —
  `purgeCloudinary` (батчи по 100, отдельно image/raw, not_found не ошибка), ошибки
  копятся счётчиком, БД НЕ откатывается. Ответ `{deleted, cloudinaryDeleted,
  cloudinaryErrors}`. **НИКОГДА не трогаются** NetworkStructure, PuStatus, TpSection,
  **SectionMonthlyPeak**, OverloadCase, Users, ResUnits (подтверждено: в purge только
  5 моделей историй/уведомлений). Фронт: в «Настройки → Обслуживание»
  (`DatabaseMaintenance`) блок «Очистка истории» — date picker, «Показать что будет
  удалено» (preview-счётчики), пароль + «Удалить безвозвратно» с confirm-модалкой
  (перечислено что удаляется/сохраняется, вкл. помесячные пики), сводка результата.
  CSS `.purge-*`. node --check / npm run build / py_compile — ОК.
- **2026-07-26** — ЗАДАЧА А: помесячные пики + раздел «Анализ мощности». Новая
  модель **`SectionMonthlyPeak`** (sectionId FK, year, month 1-12, peakKw, peakAt,
  source; уник. индекс `idx_smp_unique (sectionId,year,month)`; таблица создаётся
  обычным `sync()` без DB_ALTER). В пайплайне профиля (`processProfileFile`, после
  `section.update`) — upsert помесячного пика: месяц по `parsePeakAt(peakAt)`, иначе
  по периоду, иначе пропуск; «максимум побеждает» одним `INSERT … ON CONFLICT
  (sectionId,year,month) DO UPDATE … WHERE EXCLUDED.peakKw > …` (без гонки).
  Существующее поведение (lastPeak*/overloadStatus/OverloadCase/уведомления) не
  тронуто. Бэкфилл при старте (`initializeDatabase`, postgres): из TpSection с
  непустыми lastPeakKw+lastPeakAt тем же ON CONFLICT (идемпотентно). API
  **`GET /api/power/monthly-peaks`** (from/to=YYYY-MM, default 12 мес, resId опц.,
  res_responsible — свой) → `{ months, rows:[{…, peaks:{'YYYY-MM':{peakKw,peakAt,
  ratioPct}|null}}] }`, двумя запросами. Фронт: новый пункт меню «Анализ мощности»
  (admin/res_responsible), компонент `PowerAnalysis` — блок «ТП с перегрузом»
  ПЕРЕНЕСЁН as-is из «Аналитики» (оттуда удалён) + матрица помесячных пиков
  (пресеты 6/12 мес или произвольный from/to, ячейка peakKw·%, цвет текста >100
  red/≥85 amber/иначе green, пусто — серым, фильтр РЭС для админа через
  глобальный selectedRes) + выгрузка Excel матрицы через `styleExportSheet`
  (`cellColor` — цвет по %). CSS `.pm-*` (без infinite, переходы .12s, hover
  строк→surface-2). node --check / npm run build / py_compile — ОК.
- **2026-07-26** — Обход блокировщиков рекламы для вложений. Adblock/uBlock/Яндекс
  Protect резали `/api/download/...` и `attachment_`-public_id
  (`ERR_BLOCKED_BY_CLIENT`) — PDF/фото не открывались. (backend) логика скачивания
  вынесена в общий `handleFileProxy`, повешена на ДВА пути: `/api/download` (legacy,
  не удалять — вдруг где сохранены ссылки) и новый `/api/f/:public_id`;
  `uploadToCloudinary` для НОВЫХ загрузок использует нейтральный префикс public_id
  (`attachment_`→`doc_`; `type==='attachment'`|пусто → `doc`), старые public_id в БД
  не трогаем — тот же обработчик их отдаёт. (frontend) `fileProxyUrl` и единственная
  прямая ссылка → `/api/f/`. Регресс: старое вложение (`attachment_...`) открывается
  через `/api/f/`. `node --check`, `npm run build` — ОК.
- **2026-07-26** — Оформление Excel-выгрузки структуры (App.jsx + зависимость
  `xlsx-js-style@1.2.0`). Community `xlsx` не умеет стили — добавлен drop-in
  `xlsx-js-style` (импорт `XLSXStyle`, используется ТОЛЬКО в выгрузке структуры,
  остальные экспорты на обычном `xlsx`). Модуль-хелпер `styleExportSheet` +
  константы: тёмно-синяя шапка `#25476A` (белый жирный текст, wrap, центр), тонкие
  светлые рамки `#D9DEE5`, лёгкая зебра `#F5F8FB`, автофильтр, высота шапки 22pt.
  Цветной ЖИРНЫЙ текст только у статусов (`xlsStatusColor`: Перегруз/Ошибка→red,
  Норма/Проверен→green, Ожидает→amber, Нет данных/Не проверен/Пусто/—→gray) и у
  «Загрузка, %» (`xlsLoadColor`: >100 red, ≥85 amber, иначе green). Проверено:
  стили реально пишутся в `xl/styles.xml`. `npm run build` — ОК.
- **2026-07-26** — «Структура сети» (только App.jsx/App.css): (1) Excel-выгрузка —
  в лист «Структура» добавлены колонки «Секция», «№ ПУ тех.учёта», «Статус секции»
  (привязка ВЛ→секция видна в общей выгрузке); + второй лист «Секции (техучёт)»
  по секциям (РЭС/ТП/секция/№ПУ техучёта/привязано ВЛ/Sном/cosφ/лимит/пик/загрузка
  %/статус/дата профиля). Лист структуры получил имя «Структура» (был пустой).
  (2) Модалка «Сведения о техучёте»: весь заголовок (ТП·СШ + статус) подсвечивается
  цветом по загрузке секции — класс `tech-modal-header {barCls}` + фон `*-bg` и
  цвет `h3`/border (red/amber/green; «нет данных» — нейтральный). (3) Порог
  «оранжевого» сдвинут 90%→**85%** (`pkCls` у подписи секции и `barCls` модалки):
  красный >100%, оранжевый 85–100%, зелёный <85%. Без infinite/hover-transform.
  npm run build — ОК.
- **2026-07-25** — БАГ расчёта мощности профиля (подтверждён живым письмом): для
  ПУ, присутствующего на ОБОИХ листах «30 мин»/«60 мин» с разной раскладкой
  колонок, `kt` брался из «последнего победившего» листа (общий `pu_info`,
  meta30 перезаписывал meta60), а пик — из приоритетного листа (60) → множитель
  рассинхронизировался с пиком (пик одного столбца × Ктт другого). Пример: ПУ
  1222704 (Ктт 400/1) считался ×200. Фикс в `profile_analyzer.py`: `kt`/`ktRaw`
  держим по КАЖДОМУ листу (`kt60/ktRaw60/kt30/ktRaw30`) и выбираем строго по
  тому листу, откуда взяты данные (`source`). `pu_info` убран, список ПУ —
  стабильное объединение колонок обоих листов. Проверено синтетикой (разный
  порядок колонок на листах): 1222704 → kt=400, peakKw 1.271×400=508.4 (было
  ×200=254.2). py_compile — ОК.
- **2026-07-25** — Код-ревью: баги + перф (server.js/App.jsx, монолит не дроблён,
  правки точечные, каждый пункт — отдельный коммит). **Баги:** (1) секреты —
  `JWT_SECRET`/`DELETE_PASSWORD` теперь ОБЯЗАТЕЛЬНЫ, при отсутствии — `process.exit(1)`
  на старте (до listen), небезопасные дефолты `'secret-key'`/`'1191'` убраны, введена
  константа `JWT_SECRET`. (2) `unhandledRejection` был зарегистрирован дважды с
  `exit(1)` — оставлен ОДИН обработчик, только `console.error`, без exit
  (`uncaughtException` с exit не тронут). (3) Python-fallback: `spawn` не бросает
  синхронно → старый `try/catch` не работал; `runProfileAnalyzer`/`analyzeFile`
  переписаны — `python3`, при событии `'error'` один раз `python` с теми же
  обработчиками (`removeAllListeners('close')` чтобы 'close' первой попытки не
  перебил fallback), ОДИН таймаут 120с на обе попытки (в `analyzeFile` его раньше
  не было). (4) `extractPeriodFromError`: год брался всегда текущий → в январе
  декабрьские данные уезжали на год вперёд; теперь месяц, опережающий текущий
  более чем на 1, относится к прошлому году. (5) Фронт: два `useEffect` со
  скролл-слушателем `.content` и пустыми deps не привязывались, если `.content`
  ещё не в DOM → кнопка «наверх» не работала; переписаны на retry через
  `requestAnimationFrame` с корректным cleanup (дизайн не тронут). **Перф:**
  (6) двойной `JSON.parse` в `getNotificationCounts` и `GET /api/notifications` —
  `getPhaseSignature` получает уже распарсенный объект (ключи дедупликации не
  изменились). (7) `/api/notifications/counts` (+ бейдж платформы) — in-memory
  кэш `getNotificationCounts`, TTL 15с, ключ `role_resId_userId` (Map+timestamp,
  без Redis); полная инвалидация через хуки `afterCreate/Update/Destroy` (+bulk)
  моделей `Notification`/`OverloadCase` — ловят все пути создания/удаления, счётчики
  не залипают дольше 15с и падают сразу после загрузки анализа. (8) Импорт
  структуры Excel — убран N+1: все `ResUnit` одним `findAll` в Map, `upsert(returning)`
  вместо повторного `findOne`, `PuStatus` недостающие — одним `bulkCreate` после
  цикла (семантика `findOrCreate` сохранена: первое вхождение, существующие не
  трогаем). (9) `/api/reports/overload` — последний кейс на секцию одним
  `SELECT DISTINCT ON ("sectionId") … ORDER BY "sectionId","id" DESC` в Map вместо
  `OverloadCase.findOne` в `Promise.all`. (10) Отчёт vl_workload — количество ВЛ
  одним `findAll(COUNT group by resId)` в Map вместо `count` на каждый РЭС. (11)
  Логи: в `analyzeFile`/профиле убраны пер-строка/пер-ПУ `console.log` (~25 шт +
  banner'ы `=== UPLOAD/ANALYZE … ===`), оставлены сводка «Analysis complete» и
  `console.error` реальных ошибок; профиль — одна сводная строка (детализация
  Ктт/Ктн осталась в `details[]` ответа и в письме-автоответе). Форматы ответов
  роутов и контракты платформы не менялись. Проверки: `node --check`,
  `npm run build`, `py_compile` — ОК.
- **2026-07-24** — Заставка-логотип сайдбара «Мониторинг напряжения»: осциллограф
  отцентрован по зазору между двумя словами. Добавлена обёртка `.monitor-brand`
  (высоту задаёт текст), синусоида и сетка центрируются `top:50%/translateY(-50%)`
  по её середине. Синусоида «уже»: амплитуда `amp 19→7`, `periods 8→9` (не
  вываливается волнами вниз). Новая сетка осциллографа `.osc-grid` — тонкие белые
  квадраты (два `repeating-linear-gradient` в real-px, `--cell 5.5px`), база едва
  видна; на пробеге вспыхивает копия `.osc-grid-flash` (keyed by `sweepTick`,
  один прогон 0.85s синхронно с искрой). Инварианты соблюдены (без infinite —
  обе анимации `forwards` по событию; без hover-transform), `mix-blend:screen`
  чтобы белые буквы не темнели. `npm run build` — ОК. Только фронт.
- **2026-07-24** — Почтовый приёмник (Пирамида → профиль мощности) переведён на
  **Яндекс** и реализован. `createEmailTransporter` → `smtp.yandex.ru:465` (secure),
  добавлены `mailFrom()` (from = `MAIL_USER`, display-name «РЭС-менеджмент») и
  `sendMailAs()` — ВСЯ отправка только через них (Яндекс требует From==login).
  Пайплайн профиля вынесен в `processProfileFile(filePath, userId)` (общий для
  HTTP-роута `/api/upload/analyze` и приёмника, поведение не изменилось). Новый
  `startMailIntake()` за флагом `MAIL_INTAKE` (default OFF = прежнее поведение):
  imapflow → `imap.yandex.ru:993`, опрос непрочитанных INBOX каждые
  `MAIL_INTAKE_INTERVAL_MS`; allowlist из `MAIL_INTAKE_ALLOWED` или email
  пользователей БД; парсинг вложения через mailparser; разрешённый+вложение →
  обработка, сводка-ответ, письмо → Processed; чужой адрес → Rejected без ответа;
  разрешённый без вложения → ответ + Errors. Папки создаются с учётом
  `client.namespace` (fallback на плоские имена и INBOX-подпапку), письма реально
  переезжают через `messageMove`. Зависимости `imapflow`/`mailparser`/`dotenv`
  в package.json. `node --check` — ОК. Env для Amvera: `MAIL_USER`, `MAIL_PASS`
  (пароль приложения), `MAIL_HOST=smtp.yandex.ru`, `MAIL_PORT=465`,
  `MAIL_IMAP_HOST=imap.yandex.ru`, `MAIL_IMAP_PORT=993`, `MAIL_INTAKE=true`.
- **2026-07-24** — «Загрузить файлы», редизайн (дизайн-агент, только эта страница,
  логику не трогал): шаги с номерами-бейджами (`.fu-h3`/`.fu-step` 1→тип, 2→файл);
  карточки типов с иконкой `IconMeter` + галочкой `IconCheck` у выбранной, равная
  высота, тонкая рамка, 4 в ряд (≤900px→2), без hover-transform; при выборе типа
  появляется СИНЯЯ кнопка «Выбрать файл» (`.fu-pick-btn` — синяя-контурная
  `--blue-bg/--blue`), финальная «Загрузить и анализировать» (`.fu-submit-btn`) —
  сплошная синяя, hover→navy (два акцента разведены заливка/контур); файлы — чипы-
  пилюли. Инварианты соблюдены (переходы .12s, без infinite/transform). npm run
  build — ОК.
- **2026-07-24** — «Загрузить файлы»: кнопка загрузки — активная **синяя**
  (`.upload-submit`, `background: var(--blue)`, инлайн слева, тень, hover .12s без
  transform), вместо блёклой `.btn-default` (var(--accent), во всю ширину).
  Единый ритм отступов секций (type-selection/progress/result). Только эта страница,
  логику не трогал. npm run build — ОК.
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
