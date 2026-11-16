// =====================================================
// УЛУЧШЕННЫЙ BACKEND ДЛЯ СИСТЕМЫ УПРАВЛЕНИЯ РЭС
// Версия с исправленной логикой уведомлений
// =====================================================

// Устанавливаем кодировку
process.env.LANG = 'ru_RU.UTF-8';
process.env.LC_ALL = 'ru_RU.UTF-8';
process.env.NODE_OPTIONS = '--encoding=utf-8';

console.log('Server starting...');

process.on('uncaughtException', (err) => {
  console.error('Uncaught Exception:', err);
  process.exit(1);
});

// Единственный обработчик: логируем, НО не роняем сервер (одиночный необработанный
// reject не должен ронять прод для всех). uncaughtException выше — там exit оправдан.
process.on('unhandledRejection', (err) => {
  console.error('Unhandled Rejection:', err);
});

const express = require('express');
console.log('Express loaded');
const cors = require('cors');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const multer = require('multer');
const XLSX = require('xlsx');
const nodemailer = require('nodemailer');
const { spawn } = require('child_process');
const path = require('path');
const fs = require('fs');
const zlib = require('zlib');
const { Sequelize, DataTypes, Op } = require('sequelize');
const cloudinary = require('cloudinary').v2;
const { CloudinaryStorage } = require('multer-storage-cloudinary');
const kc = require('./keycloakPlatform');


// =====================================================
// КОНФИГУРАЦИЯ И ИНИЦИАЛИЗАЦИЯ
// =====================================================

require('dotenv').config();

// СЕКРЕТЫ ОБЯЗАТЕЛЬНЫ: без них не стартуем (иначе небезопасные дефолты в проде).
// Проверка до app.listen — при отсутствии переменной сервер падает на старте.
if (!process.env.JWT_SECRET || !process.env.DELETE_PASSWORD) {
  console.error('FATAL: не заданы обязательные переменные окружения:' +
    (!process.env.JWT_SECRET ? ' JWT_SECRET' : '') +
    (!process.env.DELETE_PASSWORD ? ' DELETE_PASSWORD' : '') +
    '. Сервер остановлен.');
  process.exit(1);
}
const JWT_SECRET = process.env.JWT_SECRET;

const app = express();
app.use((req, res, next) => {
  if (req.path.startsWith('/api')) {
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
  }
  next();
});
const PORT = process.env.PORT || 3000;

// ВАЖНО: Пароль для удаления — только из переменной окружения (проверен выше).
const DELETE_PASSWORD = process.env.DELETE_PASSWORD;

// Middleware
app.use(cors());
// Платформа: разрешаем встраивание в iframe ТОЛЬКО платформе (CSP
// frame-ancestors), снимаем легаси X-Frame-Options. Только заголовок — на
// авторизацию не влияет, не за фиче-флагом.
app.use((req, res, next) => {
  res.setHeader('Content-Security-Policy', `frame-ancestors 'self' ${kc.PLATFORM_ORIGIN}`);
  res.removeHeader('X-Frame-Options');
  next();
});
app.use(express.json());
app.use('/uploads', express.static('uploads'));

// Health check (Amvera смотрит на /api/health). Корень / отдаёт SPA (см. ниже).
app.get('/api/health', (req, res) => {
  res.json({
    status: 'OK',
    message: 'RES Management Backend is running',
    version: '2.0.1',
    features: ['user-management', 'phase-detection', 'auto-updates', 'auto-hide-notifications']
  });
});

// Создаем папку uploads если её нет
if (!fs.existsSync('uploads')) {
  fs.mkdirSync('uploads');
}
// =====================================================
// CLOUDINARY - ПРЯМАЯ ЗАГРУЗКА БЕЗ MULTER-STORAGE
// =====================================================

cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key: process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET
});

// Используем обычный multer с памятью
const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB
    files: 5
  },
  fileFilter: (req, file, cb) => {
    const allowedMimes = [
      'image/jpeg',
      'image/jpg',
      'image/png',
      'image/gif',
      'application/pdf'
    ];
    
    if (allowedMimes.includes(file.mimetype)) {
      cb(null, true);
    } else {
      cb(new Error(`Тип файла ${file.mimetype} не поддерживается`));
    }
  }
});

// Перцептивный хеш (pHash) буфера через analyzers/phash.py — для поиска
// ВИЗУАЛЬНО одинаковых фото/сканов (разные имена/размеры, кадр внутри PDF).
// Полностью защищено: любая ошибка → [] (загрузка/анализ не ломаются).
function computePhashes(buffer, mime) {
  return new Promise((resolve) => {
    try {
      if (!buffer || !buffer.length) return resolve([]);
      const os = require('os');
      const tmp = path.join(os.tmpdir(), `ph_${Date.now()}_${Math.round(Math.random() * 1e9)}`);
      fs.writeFile(tmp, buffer, (werr) => {
        if (werr) { console.error('[phash] tmp write:', werr.message); return resolve([]); }
        const scriptPath = path.join(__dirname, 'analyzers', 'phash.py');
        let out = '', done = false;
        const finish = (arr) => { if (done) return; done = true; fs.unlink(tmp, () => {}); resolve(arr); };
        const timer = setTimeout(() => finish([]), 20000); // предохранитель
        const tryRun = (cmd, isFallback) => {
          const child = spawn(cmd, [scriptPath, tmp, mime || '']);
          child.stdout.on('data', d => { out += d.toString(); });
          child.stderr.on('data', () => {});
          child.on('error', () => { if (!isFallback) tryRun('python', true); else { clearTimeout(timer); finish([]); } });
          child.on('close', () => {
            clearTimeout(timer);
            try { const j = JSON.parse(out.trim()); finish(j && j.ok && Array.isArray(j.phashes) ? j.phashes : []); }
            catch { finish([]); }
          });
        };
        tryRun('python3', false);
      });
    } catch (e) { console.error('[phash]', e.message); resolve([]); }
  });
}

// Функция загрузки в Cloudinary
async function uploadToCloudinary(file, type = 'doc') {
  const isPdf = file.mimetype === 'application/pdf';
  // Считаем pHash параллельно с заливкой — не удлиняем загрузку последовательно.
  const phashesPromise = computePhashes(file.buffer, file.mimetype);

  const uploaded = await new Promise((resolve, reject) => {
    const timestamp = Date.now();
    
    // ИСПРАВЛЕНО: сохраняем оригинальное имя в UTF-8 СРАЗУ
    const originalNameUtf8 = Buffer.from(file.originalname, 'latin1').toString('utf8');
    
    // Извлекаем расширение
    const ext = path.extname(originalNameUtf8);
    const nameWithoutExt = originalNameUtf8.replace(/\.[^/.]+$/, '');
    
    // Безопасное имя для public_id (только латиница)
    const safeName = nameWithoutExt
      .replace(/[^a-zA-Z0-9_-]/g, '_')
      .substring(0, 50);
    
    // Нейтральный префикс public_id: слово-триггер "attachment_" режется
    // блокировщиками рекламы. Старые public_id в БД не трогаем — они работают.
    const prefix = (!type || type === 'attachment') ? 'doc' : type;
    const finalPublicId = `${prefix}_${timestamp}_${safeName}${ext}`;
    
    const uploadOptions = {
      folder: 'res-management',
      resource_type: isPdf ? 'raw' : 'image',
      public_id: finalPublicId,
      type: 'upload',  // ← ВАЖНО! Вместо дефолтного
      access_mode: 'public',  // ← Делаем публичным
      use_filename: false,
      unique_filename: false,
      overwrite: false
    };
    
    if (!isPdf) {
      uploadOptions.transformation = [
        { width: 1920, height: 1920, crop: 'limit', quality: 'auto' }
      ];
    }
    
    console.log('=== UPLOADING TO CLOUDINARY ===');
    console.log('Original name (UTF-8):', originalNameUtf8);
    console.log('Type:', isPdf ? 'PDF (raw)' : 'Image');
    
    const uploadStream = cloudinary.uploader.upload_stream(
      uploadOptions,
      (error, result) => {
        if (error) {
          console.error('❌ Cloudinary error:', error);
          reject(error);
        } else {
          console.log('✅ Uploaded:', result.secure_url);
          resolve({
            url: result.secure_url,
            public_id: result.public_id,
            original_name: originalNameUtf8,  // ← Сохраняем UTF-8 имя!
            mime_type: file.mimetype,
            size: file.size
          });
        }
      }
    );
    
    const bufferStream = require('stream').Readable.from(file.buffer);
    bufferStream.pipe(uploadStream);
  });

  let phashes = [];
  try { phashes = await phashesPromise; } catch { phashes = []; }
  return { ...uploaded, phashes };
}

// Хелпер для очистки уже загруженных файлов из Cloudinary при ошибке
async function cleanupCloudinary(publicIds) {
  if (!publicIds || publicIds.length === 0) return;
  console.log(`🧹 Cleaning up ${publicIds.length} Cloudinary files...`);
  await Promise.allSettled(publicIds.map(id =>
    cloudinary.uploader.destroy(id, {
      resource_type: id.toLowerCase().endsWith('.pdf') ? 'raw' : 'image'
    })
  ));
}

// Пакетное удаление файлов из Cloudinary (для purge истории). Батчи по 100,
// отдельно image и raw (pdf) — как в существующем коде. Ошибки НЕ бросаем, а
// копим счётчиком (сирота в Cloudinary лучше битой ссылки в БД). not_found — не ошибка.
async function purgeCloudinary(publicIds) {
  const result = { deleted: 0, errors: 0 };
  if (!publicIds || !publicIds.length) return result;
  const uniq = [...new Set(publicIds)];
  const groups = [
    ['image', uniq.filter(id => !id.toLowerCase().endsWith('.pdf'))],
    ['raw', uniq.filter(id => id.toLowerCase().endsWith('.pdf'))]
  ];
  for (const [rtype, ids] of groups) {
    for (let i = 0; i < ids.length; i += 100) {
      const batch = ids.slice(i, i + 100);
      if (!batch.length) continue;
      try {
        const resp = await cloudinary.api.delete_resources(batch, { resource_type: rtype, type: 'upload' });
        const map = (resp && resp.deleted) || {};
        for (const k of Object.keys(map)) {
          if (map[k] === 'deleted') result.deleted++;
          else if (map[k] !== 'not_found') result.errors++;
        }
      } catch (e) {
        console.error('Cloudinary purge batch error:', e.message);
        result.errors += batch.length;
      }
    }
  }
  return result;
}


// =====================================================
// ПОДКЛЮЧЕНИЕ К БД (PostgreSQL на Render)
// =====================================================

const sequelize = new Sequelize(process.env.DATABASE_URL || 'postgres://user:pass@localhost:5432/dbname', {
  dialect: 'postgres',
  protocol: 'postgres',
  dialectOptions: {
    // SSL управляется env DB_SSL (true/false). По умолчанию — как раньше:
    // включён в production. У managed-Postgres Amvera (внутренняя сеть) SSL
    // часто НЕ поддерживается — тогда выставить DB_SSL=false, без правки кода.
    ssl: (process.env.DB_SSL !== undefined
      ? process.env.DB_SSL === 'true'
      : process.env.NODE_ENV === 'production')
      ? { require: true, rejectUnauthorized: false }
      : false,
    charset: 'utf8',
    client_encoding: 'UTF8'
  },
  // ✅ PERF: явный пул соединений — быстрее отдаём коннекты, не копим "висяки"
  pool: {
    max: 10,
    min: 1,
    acquire: 30000,
    idle: 10000
  },
  logging: false
});

// =====================================================
// МОДЕЛИ ДАННЫХ
// =====================================================

// 1. Модель РЭС (районов)
const ResUnit = sequelize.define('ResUnit', {
  id: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true
  },
  name: {
    type: DataTypes.STRING,
    allowNull: false,
    unique: true
  }
});

// 2. Модель пользователей
const User = sequelize.define('User', {
  id: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true
  },
  fio: {
    type: DataTypes.STRING,
    allowNull: false
  },
  login: {
    type: DataTypes.STRING,
    allowNull: false,
    unique: true
  },
  password: {
    type: DataTypes.STRING,
    allowNull: false
  },
  // Должность — ключ каталога ROLES (ниже). Была ENUM admin/uploader/res_responsible.
  role: {
    type: DataTypes.STRING(40),
    allowNull: false
  },
  resId: {
    type: DataTypes.INTEGER,
    allowNull: true,
    references: {
      model: ResUnit,
      key: 'id'
    }
  },
  email: {
    type: DataTypes.STRING,
    allowNull: false
  },
  // Единый вход через платформу: id пользователя в Keycloak (claim sub).
  keycloakId: {
    type: DataTypes.STRING(64),
    allowNull: true,
    unique: true
  },
  // Суперадмин (учётка login='admin'): видит и может всё, правами не ограничивается.
  isSuper: {
    type: DataTypes.BOOLEAN,
    allowNull: false,
    defaultValue: false
  },
  // Гранулярные права обычного админа: { <ключ каталога PERMISSIONS>: true }.
  // Источник истины — БД (+кеш); в JWT НЕ вшивается (смена прав без перелогина).
  permissions: {
    type: DataTypes.JSONB,
    allowNull: false,
    defaultValue: {}
  }
});

// 3. Модель структуры сети (ТП и ВЛ)
const NetworkStructure = sequelize.define('NetworkStructure', {
  id: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true
  },
  resId: {
    type: DataTypes.INTEGER,
    allowNull: false,
    references: {
      model: ResUnit,
      key: 'id'
    }
  },
  tpName: {
    type: DataTypes.STRING,
    allowNull: false
  },
  vlName: {
    type: DataTypes.STRING,
    allowNull: false
  },
  startPu: {
    type: DataTypes.STRING,
    allowNull: true
  },
  endPu: {
    type: DataTypes.STRING,
    allowNull: true
  },
  middlePu: {
    type: DataTypes.STRING,
    allowNull: true
  },
  // Привязка ВЛ к секции шин ТП (NULL = «ВЛ без секции», старые записи).
  sectionId: {
    type: DataTypes.INTEGER,
    allowNull: true
  },
  // Тип линии: ВЛ (воздушная), КЛ (кабельная), КВЛ (комбинированная), ВУ (вводное
  // устройство). Старые записи и файлы без типа → 'ВЛ'. Валидируется на сервере.
  lineType: {
    type: DataTypes.STRING,
    allowNull: false,
    defaultValue: 'ВЛ',
    validate: {
      isIn: [['ВЛ', 'КЛ', 'КВЛ', 'ВУ']]
    }
  },
  lastUpdate: {
    type: DataTypes.DATE,
    defaultValue: DataTypes.NOW
  }
});

// 3b. Модель секции шин ТП (СШ). Одна ТП может иметь несколько секций;
// у секции — силовой трансформатор (Sном) и ПУ технического учёта на вводе.
const TpSection = sequelize.define('TpSection', {
  id: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true
  },
  resId: {
    type: DataTypes.INTEGER,
    allowNull: false,
    references: { model: ResUnit, key: 'id' }
  },
  tpName: {
    type: DataTypes.STRING,
    allowNull: false
  },
  sectionNumber: {
    type: DataTypes.INTEGER,
    allowNull: false          // номер секции шин
  },
  tnKva: {
    type: DataTypes.FLOAT,
    allowNull: true           // Sном силового трансформатора, кВА
  },
  cosPhi: {
    type: DataTypes.FLOAT,
    defaultValue: 0.9
  },
  techPuNumber: {
    type: DataTypes.STRING,
    allowNull: true           // № ПУ технического учёта на вводе секции
  },
  overloadStatus: {
    type: DataTypes.ENUM('ok', 'overload', 'unknown'),
    defaultValue: 'unknown'
  },
  lastPeakKw: {
    type: DataTypes.FLOAT,
    allowNull: true
  },
  lastPeakAt: {
    type: DataTypes.DATE,
    allowNull: true
  },
  lastProfilePeriod: {
    type: DataTypes.STRING,
    allowNull: true
  },
  lastProfileSource: {
    type: DataTypes.STRING,      // '60' | '30' — источник часового ряда
    allowNull: true
  },
  lastProfileAt: {
    type: DataTypes.DATE,        // когда профиль последний раз обновил секцию
    allowNull: true
  },
  // Данные АСКУЭ (последний снимок при завершении этапа АСКУЭ по кейсу перегруза):
  // { consumers, limitedKw, unpolledKw, totalKw, at }. Для колонок значков в карточке
  // секции («Структура сети»). Колонка/миграция — в initializeDatabase.
  askueData: {
    type: DataTypes.JSON,
    allowNull: true
  }
});

// 3b. Помесячный максимум пика по секции (раздел «Анализ мощности»). Живёт ОТДЕЛЬНО
// от истории загрузок: истории могут стираться (задача purge), пики — навсегда.
// Уникальность (sectionId, year, month) обеспечивает индекс в initializeDatabase().
const SectionMonthlyPeak = sequelize.define('SectionMonthlyPeak', {
  id: { type: DataTypes.INTEGER, primaryKey: true, autoIncrement: true },
  sectionId: {
    type: DataTypes.INTEGER,
    allowNull: false,
    references: { model: TpSection, key: 'id' }
  },
  year: { type: DataTypes.INTEGER, allowNull: false },
  month: { type: DataTypes.INTEGER, allowNull: false },   // 1-12
  peakKw: { type: DataTypes.FLOAT, allowNull: false },
  peakAt: { type: DataTypes.DATE, allowNull: true },      // момент пика (может быть null, если только по периоду)
  source: { type: DataTypes.STRING, allowNull: true }     // '30' | '60'
});

// 3c. Случай перегрузки секции (workflow мероприятий по превышению Pном, этап 3).
// Живёт на секции (sectionId), не на ВЛ. Стадии: askue_limit → res_work →
// awaiting_recheck → completed. cycles растёт при повторном перегрузе.
const OverloadCase = sequelize.define('OverloadCase', {
  id: { type: DataTypes.INTEGER, primaryKey: true, autoIncrement: true },
  sectionId: {
    type: DataTypes.INTEGER,
    allowNull: false,
    references: { model: TpSection, key: 'id' }
  },
  resId: {
    type: DataTypes.INTEGER,
    allowNull: false,
    references: { model: ResUnit, key: 'id' }
  },
  stage: {
    type: DataTypes.ENUM('askue_limit', 'res_work', 'awaiting_recheck', 'completed'),
    defaultValue: 'askue_limit'
  },
  cycles: { type: DataTypes.INTEGER, defaultValue: 1 },
  // Снимок цифр на момент открытия / последнего перегруза
  peakKw: { type: DataTypes.FLOAT, allowNull: true },
  peakAt: { type: DataTypes.STRING, allowNull: true },
  tnKva: { type: DataTypes.FLOAT, allowNull: true },
  cosPhi: { type: DataTypes.FLOAT, allowNull: true },
  limitKw: { type: DataTypes.FLOAT, allowNull: true },
  ratio: { type: DataTypes.FLOAT, allowNull: true },
  period: { type: DataTypes.STRING, allowNull: true },
  // Этап АСКУЭ (ограничение)
  askueUserId: { type: DataTypes.INTEGER, allowNull: true },
  askueCompletedAt: { type: DataTypes.DATE, allowNull: true },
  askueComment: { type: DataTypes.TEXT, allowNull: true },
  // Снимок цифр АСКУЭ этого кейса: { consumers, limitedKw, unpolledKw, totalKw, at }.
  askueData: { type: DataTypes.JSON, allowNull: true },
  // Этап РЭС (мероприятия + фото)
  resUserId: { type: DataTypes.INTEGER, allowNull: true },
  resCompletedAt: { type: DataTypes.DATE, allowNull: true },
  resComment: { type: DataTypes.TEXT, allowNull: true },
  attachments: { type: DataTypes.JSON, allowNull: true },
  // Перепроверка профилем
  recheckAt: { type: DataTypes.DATE, allowNull: true },
  recheckPeakKw: { type: DataTypes.FLOAT, allowNull: true },
  recheckResult: { type: DataTypes.ENUM('ok', 'still_overload'), allowNull: true },
  closedAt: { type: DataTypes.DATE, allowNull: true }
});

// 4. Модель статусов ПУ (приборов учета)
const PuStatus = sequelize.define('PuStatus', {
  id: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true
  },
  puNumber: {
    type: DataTypes.STRING,
    allowNull: false,
    unique: true
  },
  networkStructureId: {
    type: DataTypes.INTEGER,
    allowNull: true,
    references: {
      model: NetworkStructure,
      key: 'id'
    }
  },
  position: {
    type: DataTypes.ENUM('start', 'end', 'middle'),
    allowNull: false
  },
  status: {
    type: DataTypes.ENUM('not_checked', 'checked_ok', 'checked_error', 'pending_recheck', 'empty'),
    defaultValue: 'not_checked'
  },
  errorDetails: {
    type: DataTypes.TEXT,
    allowNull: true
  },
  lastCheck: {
    type: DataTypes.DATE,
    allowNull: true
  }
});

// 5. Модель уведомлений
const Notification = sequelize.define('Notification', {
  id: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true
  },
  fromUserId: {
    type: DataTypes.INTEGER,
    references: {
      model: User,
      key: 'id'
    }
  },
  toUserId: {
    type: DataTypes.INTEGER,
    references: {
      model: User,
      key: 'id'
    }
  },
  resId: {
    type: DataTypes.INTEGER,
    references: {
      model: ResUnit,
      key: 'id'
    }
  },
  networkStructureId: {
    type: DataTypes.INTEGER,
    references: {
      model: NetworkStructure,
      key: 'id'
    }
  },
  puStatusId: {
    type: DataTypes.INTEGER,
    references: {
      model: PuStatus,
      key: 'id'
    }
  },
  type: {
    type: DataTypes.ENUM('error', 'success', 'info', 'pending_check', 'pending_askue', 'problem_vl', 'power_overload'),
    allowNull: false
  },
  message: {
    type: DataTypes.TEXT,
    allowNull: false
  },
  errorData: {
    type: DataTypes.JSON,
    allowNull: true
  },
  comment: {
    type: DataTypes.TEXT,
    allowNull: true
  },
  checkFromDate: {
    type: DataTypes.DATE,
    allowNull: true
  },
  isRead: {
    type: DataTypes.BOOLEAN,
    defaultValue: false
  }
});

// 6. Модель истории загрузок
const UploadHistory = sequelize.define('UploadHistory', {
  id: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true
  },
  userId: {
    type: DataTypes.INTEGER,
    references: {
      model: User,
      key: 'id'
    }
  },
  resId: {
    type: DataTypes.INTEGER,
    references: {
      model: ResUnit,
      key: 'id'
    }
  },
  fileName: {
    type: DataTypes.STRING,
    allowNull: false
  },
  fileType: {
    type: DataTypes.ENUM('rim_single', 'rim_mass', 'nartis', 'energomera', 'profile'),
    allowNull: false
  },
  processedCount: {
    type: DataTypes.INTEGER,
    defaultValue: 0
  },
  errorCount: {
    type: DataTypes.INTEGER,
    defaultValue: 0
  },
  status: {
    type: DataTypes.ENUM('processing', 'completed', 'failed'),
    defaultValue: 'processing'
  }
});

// 7. Модель истории проверок
const CheckHistory = sequelize.define('CheckHistory', {
  id: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true
  },
  resId: {
    type: DataTypes.INTEGER,
    references: {
      model: ResUnit,
      key: 'id'
    }
  },
  networkStructureId: {
    type: DataTypes.INTEGER,
    references: {
      model: NetworkStructure,
      key: 'id'
    }
  },
  puNumber: {
    type: DataTypes.STRING,
    allowNull: false
  },
  tpName: {
    type: DataTypes.STRING,
    allowNull: false
  },
  vlName: {
    type: DataTypes.STRING,
    allowNull: false
  },
  position: {
    type: DataTypes.ENUM('start', 'end', 'middle'),
    allowNull: false
  },
  initialError: {
    type: DataTypes.TEXT,
    allowNull: false
  },
  initialCheckDate: {
    type: DataTypes.DATE,
    allowNull: false
  },
  resComment: {
    type: DataTypes.TEXT,
    allowNull: true
  },
  workCompletedDate: {
    type: DataTypes.DATE,
    allowNull: true
  },
  recheckDate: {
    type: DataTypes.DATE,
    allowNull: true
  },
  recheckResult: {
    type: DataTypes.ENUM('pending', 'ok', 'error'),
    defaultValue: 'pending'
  },
  status: {
    type: DataTypes.ENUM('awaiting_work', 'awaiting_recheck', 'completed'),
    defaultValue: 'awaiting_work'
  },
   // НОВОЕ ПОЛЕ для хранения прикрепленных файлов
    attachments: {
      type: DataTypes.JSON,  // Будем хранить массив объектов с url и public_id
      defaultValue: []
  },
  failureCount: {
    type: DataTypes.INTEGER,
    defaultValue: 1  // Первая ошибка уже считается
  }

});



// 8. Модель проблемных ВЛ (2+ неудачных проверки)
const ProblemVL = sequelize.define('ProblemVL', {
  id: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true
  },
  networkStructureId: {
    type: DataTypes.INTEGER,
    references: {
      model: NetworkStructure,
      key: 'id'
    }
  },
  resId: {
    type: DataTypes.INTEGER,
    references: {
      model: ResUnit,
      key: 'id'
    }
  },
  tpName: {
    type: DataTypes.STRING,
    allowNull: false
  },
  vlName: {
    type: DataTypes.STRING,
    allowNull: false
  },
  position: {
    type: DataTypes.ENUM('start', 'middle', 'end'),
    allowNull: false
  },
  puNumber: {
    type: DataTypes.STRING,
    allowNull: false
  },
  failureCount: {
    type: DataTypes.INTEGER,
    defaultValue: 2
  },
  lastErrorDate: {
    type: DataTypes.DATE,
    allowNull: false
  },
  lastErrorDetails: {
    type: DataTypes.TEXT,
    allowNull: true
  },
  firstReportDate: {
    type: DataTypes.DATE,
    allowNull: false
  },
  resComment: {
    type: DataTypes.TEXT,
    allowNull: true
  },
  status: {
    type: DataTypes.ENUM('active', 'resolved', 'dismissed'),
    defaultValue: 'active'
  }
});

// Вопрос техблоку: РЭС отметил, что проблему ПУ своими силами не решить и нужны
// дополнительные мероприятия. Пока вопрос открыт, линия (networkStructureId)
// выведена из обычного мониторинга: по её ПУ не создаются уведомления об ошибке
// и проблемные ВЛ — линия не гоняется между АСКУЭ и РЭС. Решение принимает
// руководство техблока (право techblock_complete) кнопкой «Завершить».
const Escalation = sequelize.define('Escalation', {
  id: { type: DataTypes.INTEGER, primaryKey: true, autoIncrement: true },
  resId: { type: DataTypes.INTEGER, allowNull: false, references: { model: ResUnit, key: 'id' } },
  networkStructureId: { type: DataTypes.INTEGER, allowNull: true, references: { model: NetworkStructure, key: 'id' } },
  puNumber: { type: DataTypes.STRING, allowNull: false },
  tpName: { type: DataTypes.STRING, allowNull: false },
  vlName: { type: DataTypes.STRING, allowNull: false },
  position: { type: DataTypes.STRING(10), allowNull: true },
  errorDetails: { type: DataTypes.TEXT, allowNull: true },   // ошибка ПУ, с которой пришли
  errorData: { type: DataTypes.JSONB, allowNull: true },     // исходное уведомление целиком
  description: { type: DataTypes.TEXT, allowNull: false },   // развёрнутое описание РЭС
  attachments: { type: DataTypes.JSONB, allowNull: false, defaultValue: [] },
  createdBy: { type: DataTypes.INTEGER, allowNull: true, references: { model: User, key: 'id' } },
  status: { type: DataTypes.STRING(10), allowNull: false, defaultValue: 'open' }, // open | closed
  decision: { type: DataTypes.TEXT, allowNull: true },
  closedBy: { type: DataTypes.INTEGER, allowNull: true, references: { model: User, key: 'id' } },
  closedAt: { type: DataTypes.DATE, allowNull: true },
});

// 9. Модель прочтений уведомлений (НОВАЯ)
const NotificationRead = sequelize.define('NotificationRead', {
  id: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true
  },
  notificationId: {
    type: DataTypes.INTEGER,
    allowNull: false,
    references: {
      model: Notification,
      key: 'id'
    }
  },
  userId: {
    type: DataTypes.INTEGER,
    allowNull: false,
    references: {
      model: User,
      key: 'id'
    }
  },
  readAt: {
    type: DataTypes.DATE,
    defaultValue: DataTypes.NOW
  }
}, {
  indexes: [
    {
      unique: true,
      fields: ['notificationId', 'userId'] // один пользователь может прочитать одно уведомление только раз
    }
  ]
});

// 10. Модель истории загрузок для каждого ПУ (НОВАЯ)
const PuUploadHistory = sequelize.define('PuUploadHistory', {
  id: {
    type: DataTypes.INTEGER,
    primaryKey: true,
    autoIncrement: true
  },
  puNumber: {
    type: DataTypes.STRING,
    allowNull: false,
    index: true
  },
  uploadedBy: {
    type: DataTypes.INTEGER,
    references: {
      model: User,
      key: 'id'
    }
  },
  fileName: {
    type: DataTypes.STRING,
    allowNull: false
  },
  fileType: {
    type: DataTypes.ENUM('rim_single', 'rim_mass', 'nartis', 'energomera', 'profile'),
    allowNull: false
  },
  periodStart: {
    type: DataTypes.DATE,
    allowNull: true
  },
  periodEnd: {
    type: DataTypes.DATE,
    allowNull: true
  },
  hasErrors: {
    type: DataTypes.BOOLEAN,
    defaultValue: false
  },
  errorSummary: {
    type: DataTypes.TEXT,
    allowNull: true
  },
  errorDetails: {
    type: DataTypes.JSON,
    allowNull: true
  },
  uploadStatus: {
    type: DataTypes.ENUM('success', 'duplicate', 'wrong_period', 'error'),
    defaultValue: 'success'
  },
  uploadedAt: {
    type: DataTypes.DATE,
    defaultValue: DataTypes.NOW
  }
});



// =====================================================
// СВЯЗИ МЕЖДУ МОДЕЛЯМИ
// =====================================================

User.belongsTo(ResUnit, { foreignKey: 'resId' });
ResUnit.hasMany(User, { foreignKey: 'resId' });
NetworkStructure.belongsTo(ResUnit, { foreignKey: 'resId' });
NetworkStructure.hasMany(PuStatus, { foreignKey: 'networkStructureId' });
PuStatus.belongsTo(NetworkStructure, { foreignKey: 'networkStructureId' });
// Секции шин ТП
TpSection.belongsTo(ResUnit, { foreignKey: 'resId' });
ResUnit.hasMany(TpSection, { foreignKey: 'resId' });
NetworkStructure.belongsTo(TpSection, { foreignKey: 'sectionId', as: 'section' });
TpSection.hasMany(NetworkStructure, { foreignKey: 'sectionId', as: 'lines' });
// Случаи перегрузки секции
OverloadCase.belongsTo(TpSection, { foreignKey: 'sectionId', as: 'section' });
TpSection.hasMany(OverloadCase, { foreignKey: 'sectionId', as: 'cases' });
SectionMonthlyPeak.belongsTo(TpSection, { foreignKey: 'sectionId', as: 'section' });
TpSection.hasMany(SectionMonthlyPeak, { foreignKey: 'sectionId', as: 'monthlyPeaks' });
OverloadCase.belongsTo(ResUnit, { foreignKey: 'resId' });
OverloadCase.belongsTo(User, { as: 'askueUser', foreignKey: 'askueUserId' });
OverloadCase.belongsTo(User, { as: 'resUser', foreignKey: 'resUserId' });
Notification.belongsTo(User, { as: 'fromUser', foreignKey: 'fromUserId' });
Notification.belongsTo(User, { as: 'toUser', foreignKey: 'toUserId' });
Notification.belongsTo(ResUnit, { foreignKey: 'resId' });
Notification.belongsTo(NetworkStructure, { foreignKey: 'networkStructureId' });
Notification.belongsTo(PuStatus, { foreignKey: 'puStatusId' });
UploadHistory.belongsTo(User, { foreignKey: 'userId' });
UploadHistory.belongsTo(ResUnit, { foreignKey: 'resId' });
CheckHistory.belongsTo(ResUnit, { foreignKey: 'resId' });
CheckHistory.belongsTo(NetworkStructure, { foreignKey: 'networkStructureId' });
ProblemVL.belongsTo(ResUnit, { foreignKey: 'resId' });
ProblemVL.belongsTo(NetworkStructure, { foreignKey: 'networkStructureId' });
Notification.hasMany(NotificationRead, { foreignKey: 'notificationId' });
NotificationRead.belongsTo(Notification, { foreignKey: 'notificationId' });
NotificationRead.belongsTo(User, { foreignKey: 'userId' });
PuUploadHistory.belongsTo(User, { foreignKey: 'uploadedBy' });
Escalation.belongsTo(ResUnit, { foreignKey: 'resId' });
Escalation.belongsTo(NetworkStructure, { foreignKey: 'networkStructureId' });
Escalation.belongsTo(User, { as: 'author', foreignKey: 'createdBy' });
Escalation.belongsTo(User, { as: 'closer', foreignKey: 'closedBy' });

// =====================================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// =====================================================

// ✅ Извлечение фаз из уведомления для дедупликации
// Возвращает строку-ключ вида "A", "A_B", "A_B_C" и т.д.
function getPhaseSignature(notifMessage) {
  try {
    const data = typeof notifMessage === 'string' ? JSON.parse(notifMessage) : notifMessage;
    const phases = [];
    
    const details = data.details;
    const errorText = data.errorDetails || '';
    
    // 1. Проверяем структурированные данные
    if (details && typeof details === 'object') {
      if (details.overvoltage) {
        if (details.overvoltage.phase_A && details.overvoltage.phase_A.count > 0) phases.push('A');
        if (details.overvoltage.phase_B && details.overvoltage.phase_B.count > 0) phases.push('B');
        if (details.overvoltage.phase_C && details.overvoltage.phase_C.count > 0) phases.push('C');
      }
      if (details.undervoltage) {
        if (details.undervoltage.phase_A && details.undervoltage.phase_A.count > 0 && !phases.includes('A')) phases.push('A');
        if (details.undervoltage.phase_B && details.undervoltage.phase_B.count > 0 && !phases.includes('B')) phases.push('B');
        if (details.undervoltage.phase_C && details.undervoltage.phase_C.count > 0 && !phases.includes('C')) phases.push('C');
      }
    }
    
    // 2. Если не нашли в структуре — ищем в тексте
    if (phases.length === 0 && errorText) {
      if (errorText.indexOf('Фаза A') !== -1 || errorText.indexOf('phase_A') !== -1) phases.push('A');
      if (errorText.indexOf('Фаза B') !== -1 || errorText.indexOf('phase_B') !== -1) phases.push('B');
      if (errorText.indexOf('Фаза C') !== -1 || errorText.indexOf('phase_C') !== -1) phases.push('C');
    }
    
    // 3. Если фазы не определены — используем "ALL" как общий ключ
    return phases.length > 0 ? phases.sort().join('_') : 'ALL';
  } catch {
    return 'ALL';
  }
}

// Хеширование паролей
User.beforeCreate(async (user) => {
  user.password = await bcrypt.hash(user.password, 10);
});

User.beforeUpdate(async (user) => {
  // Не перехэшировать уже готовый bcrypt-хэш (важно при restore). Покрываем
  // оба префикса: $2a$ и $2b$.
  if (user.changed('password') && user.password &&
      !user.password.startsWith('$2a$') && !user.password.startsWith('$2b$')) {
    user.password = await bcrypt.hash(user.password, 10);
  }
});

// Валидация пароля
User.prototype.validatePassword = async function(password) {
  return bcrypt.compare(password, this.password);
};

// =====================================================
// ДОЛЖНОСТИ И ПРАВА (как в ЭнергоСфере)
// =====================================================
// Должность = область видимости (scope: 'all' — все РЭС, 'res' — свой РЭС),
// право = действие или раздел. Права настраиваются НЕ персонально, а матрицей
// «должность × право» (таблица RolePermissions, раздел «Права доступа»).
// Стартовые права выдаются один раз на должность (SeededRoles) — дальше
// матрицу ведёт администратор. Суперадмин (login='admin') может всё.
const ROLES = {
  admin:           { title: 'Администратор',      scope: 'all' },
  nach_sue:        { title: 'Начальник СУЭ',       scope: 'all' },
  eng_sue:         { title: 'Инженер СУЭ',         scope: 'all' },
  nach_res:        { title: 'Начальник РЭС',       scope: 'res' },
  master_res:      { title: 'Мастер РЭС',          scope: 'res' },
  nach_riru:       { title: 'Начальник РИРУ',      scope: 'res' },
  eng_realization: { title: 'Инженер реализации',  scope: 'res' },
  eng_development: { title: 'Инженер развития',    scope: 'res' },
  monter:          { title: 'Монтёр',              scope: 'res' },
  glav_ing:        { title: 'Главный инженер',     scope: 'all' },
  nach_srs:        { title: 'Начальник СРС',       scope: 'all' },
  eng_srs:         { title: 'Инженер СРС',         scope: 'all' },
  nach_pts:        { title: 'Начальник ПТС',       scope: 'all' },
  eng_pts:         { title: 'Инженер ПТС',         scope: 'all' },
  nach_stp:        { title: 'Начальник СТП',       scope: 'all' },
  eng_stp:         { title: 'Инженер СТП',         scope: 'all' },
};
// Прежние роли ПО → должности (переезд учёток при первом старте новой версии).
const LEGACY_ROLES = { uploader: 'eng_sue', res_responsible: 'master_res' };

// Каталог прав — ЕДИНЫЙ источник (ключ → группа и русское название для UI и 403).
// Новое право: ключ здесь → requirePerm(ключ) на роут → hasPerm(user, ключ) в UI.
const PERMISSION_GROUPS = [
  { id: 'sections', title: 'Разделы' },
  { id: 'askue', title: 'АСКУЭ' },
  { id: 'res', title: 'РЭС' },
  { id: 'tech', title: 'Техблок' },
  { id: 'structure', title: 'Структура сети' },
  { id: 'admin', title: 'Администрирование' },
];
const PERMISSION_CATALOG = {
  dashboard_view:        { group: 'sections', title: 'Сводка' },
  structure_view:        { group: 'sections', title: 'Структура сети' },
  tech_pending_view:     { group: 'sections', title: 'Ожидающие мероприятий (просмотр)' },
  askue_pending_view:    { group: 'sections', title: 'Ожидающие проверки АСКУЭ (просмотр)' },
  problem_vl_view:       { group: 'sections', title: 'Проблемные ВЛ и ТП' },
  overload_view:         { group: 'sections', title: 'Превышение Pном' },
  power_analysis_view:   { group: 'sections', title: 'Анализ мощности' },
  voltage_analysis_view: { group: 'sections', title: 'Анализ напряжения' },
  deep_analysis_view:    { group: 'sections', title: 'Глубокий анализ работ' },
  documents_view:        { group: 'sections', title: 'Загруженные документы' },
  history_view:          { group: 'sections', title: 'История системы' },
  reports_view:          { group: 'sections', title: 'Отчёты' },
  upload_files:          { group: 'askue', title: 'Загружать журналы и профили мощности (проверка и перепроверка)' },
  overload_askue:        { group: 'askue', title: 'Завершать ограничения по перегрузу (этап АСКУЭ)' },
  problem_vl_manage:     { group: 'askue', title: 'Проблемные ВЛ: письма РЭС и «рассмотрено»' },
  res_work:              { group: 'res', title: 'Отмечать выполнение мероприятий по ошибкам ПУ' },
  overload_res:          { group: 'res', title: 'Отмечать мероприятия по перегрузу ТП' },
  techblock_view:        { group: 'tech', title: 'Видеть «Требуют решения техблока»' },
  techblock_complete:    { group: 'tech', title: 'Завершать вопросы техблока (решение принято)' },
  structure_upload:      { group: 'structure', title: 'Загрузка/обновление структуры сети' },
  structure_edit:        { group: 'structure', title: 'Редактирование структуры (ПУ, ВЛ/КЛ/КВЛ/ВУ, секции, тех.учёты, добавление/удаление)' },
  structure_edit_pu:     { group: 'structure', title: 'Изменение номеров ПУ своего РЭС (без секций/удаления)' },
  checks_delete:         { group: 'structure', title: 'Удаление карточек проверок и истории' },
  settings_view:         { group: 'admin', title: 'Открывать «Настройки»' },
  notifications_delete:  { group: 'admin', title: 'Удаление уведомлений' },
  files_manage:          { group: 'admin', title: 'Управление файлами (удаление, диагностика)' },
  history_purge:         { group: 'admin', title: 'Очистка истории до даты' },
  users_manage:          { group: 'admin', title: 'Управление пользователями' },
  permissions_manage:    { group: 'admin', title: 'Настройка прав должностей' },
  db_tools:              { group: 'admin', title: 'Сервисные операции с базой (проверка, бэкап, восстановление)' },
};
const PERMISSIONS = Object.fromEntries(Object.entries(PERMISSION_CATALOG).map(([k, v]) => [k, v.title]));

// Стартовая матрица. Выдаётся один раз на должность — потом её правит администратор.
const VIEW_COMMON = ['dashboard_view', 'structure_view', 'power_analysis_view', 'voltage_analysis_view', 'documents_view', 'history_view', 'reports_view'];
const TECH_VIEW = [...VIEW_COMMON, 'tech_pending_view', 'problem_vl_view', 'overload_view', 'techblock_view'];
const DEFAULT_ROLE_PERMISSIONS = {
  admin: Object.keys(PERMISSION_CATALOG),
  nach_sue: [...VIEW_COMMON, 'tech_pending_view', 'askue_pending_view', 'problem_vl_view', 'overload_view', 'deep_analysis_view',
    'upload_files', 'overload_askue', 'problem_vl_manage', 'structure_edit_pu', 'techblock_view'],
  eng_sue: [...VIEW_COMMON, 'tech_pending_view', 'askue_pending_view', 'problem_vl_view', 'overload_view',
    'upload_files', 'overload_askue', 'structure_edit_pu'],
  nach_res: [...VIEW_COMMON, 'tech_pending_view', 'overload_view', 'res_work', 'overload_res'],
  master_res: [...VIEW_COMMON, 'tech_pending_view', 'overload_view', 'res_work', 'overload_res'],
  nach_riru: [...VIEW_COMMON, 'tech_pending_view', 'overload_view'],
  eng_realization: [...VIEW_COMMON],
  eng_development: [...VIEW_COMMON],
  monter: ['dashboard_view', 'structure_view'],
  glav_ing: [...TECH_VIEW, 'deep_analysis_view', 'techblock_complete'],
  nach_srs: [...TECH_VIEW, 'techblock_complete'],
  eng_srs: [...TECH_VIEW],
  nach_pts: [...TECH_VIEW, 'techblock_complete'],
  eng_pts: [...TECH_VIEW],
  nach_stp: [...VIEW_COMMON, 'overload_view'],
  eng_stp: [...VIEW_COMMON, 'overload_view'],
};

const RolePermission = sequelize.define('RolePermission', {
  role: { type: DataTypes.STRING(40), allowNull: false },
  permission: { type: DataTypes.STRING(60), allowNull: false },
}, { indexes: [{ unique: true, fields: ['role', 'permission'] }] });
const SeededRole = sequelize.define('SeededRole', {
  role: { type: DataTypes.STRING(40), primaryKey: true },
});

// Матрица в памяти: role → Set(прав). Перечитывается при правке и раз в минуту.
let roleMatrix = new Map();
let roleMatrixAt = 0;
async function loadRoleMatrix(force = false) {
  if (!force && Date.now() - roleMatrixAt < 60 * 1000) return roleMatrix;
  const rows = await RolePermission.findAll({ attributes: ['role', 'permission'], raw: true });
  const m = new Map();
  for (const { role, permission } of rows) {
    if (!PERMISSION_CATALOG[permission]) continue;
    if (!m.has(role)) m.set(role, new Set());
    m.get(role).add(permission);
  }
  roleMatrix = m;
  roleMatrixAt = Date.now();
  return m;
}
// Выдать стартовые права должностям, которым их ещё не выдавали.
async function seedRolePermissions() {
  const seeded = new Set((await SeededRole.findAll({ raw: true })).map(r => r.role));
  for (const [role, perms] of Object.entries(DEFAULT_ROLE_PERMISSIONS)) {
    if (seeded.has(role)) continue;
    await RolePermission.bulkCreate(perms.map(permission => ({ role, permission })), { ignoreDuplicates: true });
    await SeededRole.create({ role });
  }
  await loadRoleMatrix(true);
}

// Порог «проблемной ТП»: кейс перегруза с cycles >= этого и stage != completed —
// мероприятия не дали результата, ТП попадает в «Проблемные ТП».
const PROBLEM_TP_CYCLES = 2;

// id учётки для системных уведомлений (fromUserId). Проставляется в
// initializeDatabase() из суперадмина; null — безопасный fallback (поле nullable).
let SYSTEM_USER_ID = null;

// Кеш учёток: userId → { isSuper, role, resId, ts }. TTL 60с; сбрасывается при
// изменении пользователя. В JWT доверяем только id: должность и РЭС — из БД.
const ACCESS_TTL_MS = 60 * 1000;
const accessCache = new Map();
const invalidateAccess = (userId) => { accessCache.delete(Number(userId)); };
async function getUserAccess(userId) {
  const key = Number(userId);
  const hit = accessCache.get(key);
  let base = hit && (Date.now() - hit.ts) < ACCESS_TTL_MS ? hit : null;
  if (!base) {
    const u = await User.findByPk(key, { attributes: ['id', 'role', 'resId', 'isSuper'] });
    base = { exists: !!u, isSuper: !!(u && u.isSuper), role: u ? u.role : null, resId: u ? u.resId : null, ts: Date.now() };
    accessCache.set(key, base);
  }
  const matrix = await loadRoleMatrix();
  const perms = (base.isSuper || base.role === 'admin') ? new Set(Object.keys(PERMISSION_CATALOG)) : (matrix.get(base.role) || new Set());
  const allRes = base.isSuper || ROLES[base.role]?.scope === 'all';
  return { ...base, perms, allRes };
}
// Учётки, чьей должности выдано право (для адресных уведомлений).
async function usersWithPerm(key, where = {}) {
  const matrix = await loadRoleMatrix();
  const roles = ['admin', ...[...matrix.entries()].filter(([, set]) => set.has(key)).map(([role]) => role)];
  return User.findAll({ where: { ...where, [Op.or]: [{ role: { [Op.in]: roles } }, { isSuper: true }] } });
}
const userCan = (user, key) => !!user && (user.isSuper || (user.perms && user.perms.has(key)));
// Учётка с правами — для мест, где под рукой модель User, а не req.user (бейдж платформы).
async function accessOf(user) {
  const acc = await getUserAccess(user.id);
  return { id: user.id, role: acc.role, resId: acc.resId, isSuper: acc.isSuper, allRes: acc.allRes, perms: acc.perms };
}
// Что отдаём интерфейсу о текущем пользователе (login, platform, me).
async function publicUser(user) {
  const acc = await getUserAccess(user.id);
  return {
    id: user.id,
    fio: user.fio,
    role: acc.role,
    roleTitle: acc.isSuper ? 'Суперадмин' : (ROLES[acc.role]?.title || acc.role),
    resId: acc.resId,
    resName: user.ResUnit?.name,
    isSuper: acc.isSuper,
    allRes: acc.allRes,
    permissions: Object.fromEntries([...acc.perms].map(k => [k, true])),
  };
}

// Middleware для проверки JWT токена. Должность, РЭС и права — свежие из БД (кеш),
// поэтому смена должности или матрицы действует без перелогина.
const authenticateToken = (req, res, next) => {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1];

  if (!token) {
    return res.status(401).json({ error: 'Access denied' });
  }

  jwt.verify(token, JWT_SECRET, (err, payload) => {
    if (err) {
      return res.status(403).json({ error: 'Invalid token' });
    }
    getUserAccess(payload.id).then(acc => {
      if (!acc.exists) return res.status(401).json({ error: 'Учётная запись удалена' });
      req.user = { id: payload.id, role: acc.role, resId: acc.resId, isSuper: acc.isSuper, allRes: acc.allRes, perms: acc.perms };
      next();
    }, e => {
      console.error('auth access error:', e.message);
      res.status(500).json({ error: 'Ошибка проверки доступа' });
    });
  });
};

// Право на действие или раздел. `key` — ключ каталога ИЛИ массив ключей (достаточно
// любого). Суперадмин — всегда. Право не расширяет географию: свой РЭС проверяет роут.
const requirePerm = (key) => (req, res, next) => {
  const list = Array.isArray(key) ? key : [key];
  if (req.user && list.some(k => userCan(req.user, k))) return next();
  return res.status(403).json({ error: 'Недостаточно прав', permission: list[0], title: PERMISSIONS[list[0]] || list[0] });
};

// Только суперадмин.
const requireSuper = (req, res, next) => {
  if (req.user?.isSuper) return next();
  return res.status(403).json({ error: 'Доступно только суперадмину' });
};

// Файловый скоуп-токен — отдельный от сессионного, кладётся в URL (?t=) для
// скачивания вложений обычными <a>-ссылками. В URL не светим полноценный
// сессионный JWT (URL попадают в историю/логи). Живёт 24ч, как и сессия.
function makeFileToken(user) {
  return jwt.sign(
    { id: user.id, role: user.role, resId: user.resId, scope: 'files' },
    JWT_SECRET,
    { expiresIn: '24h' }
  );
}

// Доступ к файловому прокси (/api/f, /api/download): принимаем ЛИБО
// Authorization: Bearer (просмотр через axios), ЛИБО ?t=<jwt> (скачивание
// ссылкой). Годится и обычный сессионный токен, и файловый scope 'files'.
const authenticateFileAccess = (req, res, next) => {
  const authHeader = req.headers['authorization'];
  const token = (authHeader && authHeader.split(' ')[1]) || req.query.t;
  if (!token) return res.status(401).json({ error: 'Access denied' });
  jwt.verify(token, JWT_SECRET, (err, user) => {
    if (err) return res.status(401).json({ error: 'Invalid token' });
    req.user = user;
    next();
  });
};

// Настройка multer для загрузки файлов с ограничением размера
const storageExcel = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, 'uploads/');
  },
  filename: (req, file, cb) => {
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
  }
});

const uploadExcel = multer({ 
  storage: storageExcel,
  limits: {
    fileSize: 10 * 1024 * 1024 // 10MB максимум
  },
  fileFilter: (req, file, cb) => {
    const allowedExtensions = ['.xlsx', '.xls', '.csv'];
    const ext = path.extname(file.originalname).toLowerCase();
    if (allowedExtensions.includes(ext)) {
      cb(null, true);
    } else {
      cb(new Error('Разрешены только Excel и CSV файлы'));
    }
  }
});

// Email сервис (SMTP — Яндекс по умолчанию, всё переопределяется env).
const createEmailTransporter = () => {
  return nodemailer.createTransport({
    host: process.env.MAIL_HOST || 'smtp.yandex.ru',
    port: Number(process.env.MAIL_PORT) || 465,
    secure: true,
    auth: {
      user: process.env.MAIL_USER,
      pass: process.env.MAIL_PASS
    }
  });
};

// From ОБЯЗАН совпадать с логином (Яндекс: 553 Sender address rejected иначе).
// Всегда шлём с MAIL_USER (с display name). Централизованный отправитель.
const mailFrom = () => `"РЭС-менеджмент" <${process.env.MAIL_USER}>`;
async function sendMailAs(to, subject, text) {
  const transporter = createEmailTransporter();
  return transporter.sendMail({ from: mailFrom(), to, subject, text });
}

// =====================================================
// API РОУТЫ
// =====================================================

// 1. АВТОРИЗАЦИЯ
app.post('/api/auth/login', async (req, res) => {
  try {
    const { login, password } = req.body;
    
    const user = await User.findOne({ 
      where: { login },
      include: [ResUnit]
    });
    
    if (!user || !(await user.validatePassword(password))) {
      return res.status(401).json({ error: 'Invalid credentials' });
    }
    
    const token = jwt.sign(
      { 
        id: user.id, 
        role: user.role, 
        resId: user.resId,
        
      },
      JWT_SECRET,
      { expiresIn: '24h' }
    );
    
    res.json({
      token,
      fileToken: makeFileToken(user),
      user: await publicUser(user)
    });
  } catch (error) {
    console.error('Login error:', error);
    res.status(500).json({ error: error.message });
  }
});

// =====================================================
// ЕДИНЫЙ ВХОД ЧЕРЕЗ ПЛАТФОРМУ (Keycloak SSO)
// =====================================================

// Достать Bearer-токен из заголовка Authorization.
function getBearer(req) {
  const h = req.headers['authorization'] || '';
  return h.toLowerCase().startsWith('bearer ') ? h.slice(7).trim() : null;
}

// Обмен Keycloak-токена платформы на обычную сессию приложения.
// Keycloak решает «кто ты» (email) + «пускать ли» (роль ACCESS_ROLE), а
// функциональная роль и resId берутся из СВОЕЙ БД по email.
app.post('/api/auth/platform', async (req, res) => {
  try {
    if (!kc.PLATFORM_SSO) return res.status(401).json({ error: 'Platform SSO disabled' });
    const token = getBearer(req);
    if (!token) return res.status(401).json({ error: 'No token' });

    let claims;
    try {
      claims = await kc.verifyToken(token);
    } catch (e) {
      console.warn('Platform SSO 401:', e.message); // причина — да, токен — никогда
      return res.status(401).json({ error: 'Invalid platform token' });
    }
    const ident = kc.identityFromClaims(claims);
    if (!ident.keycloakId) return res.status(401).json({ error: 'Invalid platform token' });
    if (!kc.hasAccess(ident.roles)) return res.status(403).json({ error: 'Нет доступа к приложению' });

    // Поиск учётки: по keycloakId, затем разово по email (регистронезависимо).
    let user = await User.findOne({ where: { keycloakId: ident.keycloakId }, include: [ResUnit] });
    if (!user && ident.email) {
      user = await User.findOne({
        where: sequelize.where(sequelize.fn('lower', sequelize.col('email')), ident.email.toLowerCase()),
        include: [ResUnit]
      });
      if (user && !user.keycloakId) {
        user.keycloakId = ident.keycloakId; // разовая привязка
        await user.save();
      }
    }
    if (!user) return res.status(401).json({ error: 'Пользователь не найден' });

    // Обычный JWT приложения — тот же формат, что в /api/auth/login.
    const appToken = jwt.sign(
      { id: user.id, role: user.role, resId: user.resId },
      JWT_SECRET,
      { expiresIn: '24h' }
    );
    res.json({
      token: appToken,
      fileToken: makeFileToken(user),
      user: await publicUser(user)
    });
  } catch (error) {
    console.error('Platform auth error:', error.message);
    res.status(500).json({ error: error.message });
  }
});

// Счётчик уведомлений для бейджа иконки на рабочем столе платформы.
// Проверка тем же Keycloak-токеном, БЕЗ создания сессии, только чтение.
app.get('/api/platform/badge', async (req, res) => {
  try {
    if (!kc.PLATFORM_SSO) return res.status(401).json({ error: 'Platform SSO disabled' });
    const token = getBearer(req);
    if (!token) return res.status(401).json({ error: 'No token' });

    let claims;
    try {
      claims = await kc.verifyToken(token);
    } catch (e) {
      return res.status(401).json({ error: 'Invalid platform token' });
    }
    const ident = kc.identityFromClaims(claims);
    if (!ident.keycloakId) return res.status(401).json({ error: 'Invalid platform token' });

    // Только чтение: по keycloakId, затем по email. Не найден → count 0.
    let user = await User.findOne({ where: { keycloakId: ident.keycloakId } });
    if (!user && ident.email) {
      user = await User.findOne({
        where: sequelize.where(sequelize.fn('lower', sequelize.col('email')), ident.email.toLowerCase())
      });
    }
    if (!user) return res.json({ count: 0 });

    // «Требует действия» — то, что человек по своей должности может отработать.
    const acc = await accessOf(user);
    const counts = await getNotificationCountsCached(acc);
    let count = 0;
    if (userCan(acc, 'res_work')) count += counts.tech_pending || 0;
    if (userCan(acc, 'upload_files')) count += counts.askue_pending || 0;
    if (userCan(acc, 'problem_vl_manage')) count += counts.problem_vl || 0;
    if (userCan(acc, 'techblock_complete')) count += counts.techblock || 0;
    count += counts.powerOverload || 0;
    res.json({ count });
  } catch (error) {
    console.error('Platform badge error:', error.message);
    res.status(500).json({ error: error.message });
  }
});

// 1.1 ПОЛУЧЕНИЕ ТЕКУЩЕГО ПОЛЬЗОВАТЕЛЯ
app.get('/api/auth/me', authenticateToken, async (req, res) => {
  try {
    const user = await User.findByPk(req.user.id, {
      attributes: ['id', 'fio', 'login', 'role', 'resId', 'email', 'isSuper', 'permissions'],
      include: [ResUnit]
    });

    if (!user) {
      return res.status(404).json({ error: 'User not found' });
    }

    res.json({
      fileToken: makeFileToken(user),
      user: await publicUser(user)
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 2. ПОЛУЧЕНИЕ СПИСКА РЭС
app.get('/api/res/list', authenticateToken, async (req, res) => {
  try {
    const resList = await ResUnit.findAll({
      order: [['name', 'ASC']]
    });
    res.json(resList);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 3. ПОЛУЧЕНИЕ СТРУКТУРЫ СЕТИ
app.get('/api/network/structure/:resId?', authenticateToken, async (req, res) => {
  try {
    const resId = req.params.resId || req.user.resId;
    
    // Если не админ, может видеть только свой РЭС
    if (!req.user.allRes && resId != req.user.resId) {
      return res.status(403).json({ error: 'Access denied' });
    }
    
    let whereClause = {};
    if (resId) {
      whereClause = { resId };
    }
    
    const structures = await NetworkStructure.findAll({
      where: whereClause,
      include: [
        {
          model: PuStatus,
          required: false,
          attributes: ['id', 'puNumber', 'position', 'status', 'errorDetails', 'lastCheck']
        },
        {
          model: TpSection,
          as: 'section',
          required: false,
          attributes: ['id', 'sectionNumber', 'tnKva', 'techPuNumber', 'overloadStatus']
        },
        ResUnit
      ],
      order: [['tpName', 'ASC'], ['vlName', 'ASC']]
    });

    res.json(structures);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 4. ОБНОВЛЕНИЕ структуры сети. Доступ по праву (любая роль через «Права доступа»):
// structure_edit — полностью (включая привязку секций); structure_edit_pu — ТОЛЬКО
// номера ПУ и ТОЛЬКО своего РЭС.
app.put('/api/network/structure/:id',
  authenticateToken,
  requirePerm(['structure_edit', 'structure_edit_pu']),
  async (req, res) => {
    try {
      const { startPu, middlePu, endPu } = req.body;

      const structure = await NetworkStructure.findByPk(req.params.id);
      if (!structure) {
        return res.status(404).json({ error: 'ВЛ не найдена' });
      }

      // Полная правка — у кого есть structure_edit (или суперадмин); остальные
      // (право structure_edit_pu) — только ПУ и только в своём РЭС.
      const acc = await getUserAccess(req.user.id);
      const fullEdit = acc.isSuper || !!(acc.permissions && acc.permissions.structure_edit);
      if (!fullEdit && Number(structure.resId) !== Number(req.user.resId)) {
        return res.status(403).json({ error: 'Редактировать ПУ можно только в своём РЭС' });
      }

      const updates = {
        startPu: startPu || null,
        middlePu: middlePu || null,
        endPu: endPu || null,
        lastUpdate: new Date()
      };

      // Привязка/перепривязка ВЛ к секции шин (sectionId) — только полная правка.
      // Секция должна быть того же РЭСа и той же ТП, что и ВЛ, иначе 400. null = снять.
      if (fullEdit && 'sectionId' in req.body) {
        const sectionId = req.body.sectionId;
        if (sectionId === null || sectionId === '') {
          updates.sectionId = null;
        } else {
          const section = await TpSection.findByPk(sectionId);
          if (!section) {
            return res.status(400).json({ error: 'Секция не найдена' });
          }
          if (section.resId !== structure.resId || section.tpName !== structure.tpName) {
            return res.status(400).json({ error: 'Секция принадлежит другой ТП или РЭС' });
          }
          updates.sectionId = section.id;
        }
      }

      // Тип линии (ВЛ/КЛ/КВЛ/ВУ) — менять может ТОЛЬКО полная правка
      // (structure_edit / суперадмин). Обладателю structure_edit_pu — запрещено
      // (проверка на СЕРВЕРЕ, не только скрытие в UI).
      if ('lineType' in req.body) {
        if (!fullEdit) {
          return res.status(403).json({ error: 'Менять тип линии может только полная правка структуры (structure_edit)' });
        }
        const LINE_TYPES = ['ВЛ', 'КЛ', 'КВЛ', 'ВУ'];
        if (!LINE_TYPES.includes(req.body.lineType)) {
          return res.status(400).json({ error: 'Недопустимый тип линии (ВЛ/КЛ/КВЛ/ВУ)' });
        }
        updates.lineType = req.body.lineType;
      }

      await structure.update(updates);

      res.json({ success: true, message: 'Структура обновлена' });
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
});

// 4a. СЕКЦИИ ШИН ТП — список секций РЭСа со счётчиком привязанных ВЛ
app.get('/api/network/sections', authenticateToken, async (req, res) => {
  try {
    const resId = req.query.resId || req.user.resId;
    if (!req.user.allRes && resId != req.user.resId) {
      return res.status(403).json({ error: 'Access denied' });
    }

    const where = {};
    if (resId) where.resId = resId;

    const sections = await TpSection.findAll({
      where,
      include: [{ model: ResUnit, attributes: ['name'] }],
      order: [['tpName', 'ASC'], ['sectionNumber', 'ASC']]
    });

    // Счётчик привязанных ВЛ по секциям (одним запросом).
    const counts = await NetworkStructure.findAll({
      attributes: ['sectionId', [sequelize.fn('COUNT', sequelize.col('id')), 'cnt']],
      where: { sectionId: { [Op.ne]: null }, ...(resId ? { resId } : {}) },
      group: ['sectionId'],
      raw: true
    });
    const cntBySection = {};
    counts.forEach(c => { cntBySection[c.sectionId] = parseInt(c.cnt, 10); });

    const result = sections.map(s => ({
      ...s.toJSON(),
      linesCount: cntBySection[s.id] || 0
    }));

    res.json(result);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 4b. Создание секции — право structure_edit (любая роль); область — свой РЭС.
app.post('/api/network/sections',
  authenticateToken,
  requirePerm('structure_edit'),
  async (req, res) => {
    try {
      const { resId, tpName, sectionNumber, tnKva, cosPhi, techPuNumber } = req.body;
      if (!resId || !tpName || sectionNumber === undefined || sectionNumber === null || sectionNumber === '') {
        return res.status(400).json({ error: 'Обязательны РЭС, ТП и номер секции' });
      }
      // Право не расширяет географию: не-админ создаёт секции только своего РЭС.
      if (!req.user.allRes && Number(resId) !== Number(req.user.resId)) {
        return res.status(403).json({ error: 'Создавать секции можно только в своём РЭС' });
      }

      // Десятичный разделитель — запятая или точка; невалид → null (Sном) / 0.9 (cosφ).
      const pfKva = (tnKva === '' || tnKva == null) ? null : parseFloat(String(tnKva).replace(',', '.'));
      const pfCos = (cosPhi === '' || cosPhi == null) ? 0.9 : parseFloat(String(cosPhi).replace(',', '.'));

      const section = await TpSection.create({
        resId,
        tpName: String(tpName).trim(),
        sectionNumber: parseInt(sectionNumber, 10),
        tnKva: Number.isFinite(pfKva) ? pfKva : null,
        cosPhi: Number.isFinite(pfCos) ? pfCos : 0.9,
        techPuNumber: techPuNumber ? String(techPuNumber).trim() : null
      });

      res.json({ success: true, section });
    } catch (error) {
      if (error.name === 'SequelizeUniqueConstraintError') {
        return res.status(400).json({ error: 'Секция с таким номером уже есть у этой ТП' });
      }
      res.status(500).json({ error: error.message });
    }
});

// 4c. Обновление секции — право structure_edit; область — свой РЭС.
app.put('/api/network/sections/:id',
  authenticateToken,
  requirePerm('structure_edit'),
  async (req, res) => {
    try {
      const section = await TpSection.findByPk(req.params.id);
      if (!section) return res.status(404).json({ error: 'Секция не найдена' });
      if (!req.user.allRes && Number(section.resId) !== Number(req.user.resId)) {
        return res.status(403).json({ error: 'Редактировать секции можно только в своём РЭС' });
      }

      const { sectionNumber, tnKva, cosPhi, techPuNumber } = req.body;
      const updates = {};
      if (sectionNumber !== undefined) updates.sectionNumber = parseInt(sectionNumber, 10);
      if (tnKva !== undefined) {
        const v = (tnKva === '' || tnKva === null) ? null : parseFloat(String(tnKva).replace(',', '.'));
        updates.tnKva = Number.isFinite(v) ? v : null;
      }
      if (cosPhi !== undefined) {
        const v = (cosPhi === '' || cosPhi === null) ? 0.9 : parseFloat(String(cosPhi).replace(',', '.'));
        updates.cosPhi = Number.isFinite(v) ? v : 0.9;
      }
      const techPuChanged = techPuNumber !== undefined &&
        (techPuNumber ? String(techPuNumber).trim() : null) !== (section.techPuNumber || null);
      if (techPuNumber !== undefined) updates.techPuNumber = techPuNumber ? String(techPuNumber).trim() : null;

      await section.update(updates);

      // Предупреждение: смена тех.ПУ при открытом случае перегруза.
      let warning = null;
      if (techPuChanged) {
        const activeCase = await OverloadCase.count({
          where: { sectionId: section.id, stage: { [Op.ne]: 'completed' } }
        });
        if (activeCase > 0) warning = 'По секции открыт случай перегруза';
      }
      res.json({ success: true, section, ...(warning ? { warning } : {}) });
    } catch (error) {
      if (error.name === 'SequelizeUniqueConstraintError') {
        return res.status(400).json({ error: 'Секция с таким номером уже есть у этой ТП' });
      }
      res.status(500).json({ error: error.message });
    }
});

// 4d. Удаление секции — право structure_edit; область — свой РЭС. Запрещено, если к ней привязаны ВЛ
app.delete('/api/network/sections/:id',
  authenticateToken,
  requirePerm('structure_edit'),
  async (req, res) => {
    try {
      const section = await TpSection.findByPk(req.params.id);
      if (!section) return res.status(404).json({ error: 'Секция не найдена' });
      if (!req.user.allRes && Number(section.resId) !== Number(req.user.resId)) {
        return res.status(403).json({ error: 'Удалять секции можно только в своём РЭС' });
      }

      const linked = await NetworkStructure.count({ where: { sectionId: section.id } });
      if (linked > 0) {
        return res.status(400).json({
          error: `Нельзя удалить секцию: к ней привязано ВЛ — ${linked}. Сначала отвяжите линии.`
        });
      }

      // Нельзя удалять секцию с незакрытым случаем перегруза.
      const activeCase = await OverloadCase.count({
        where: { sectionId: section.id, stage: { [Op.ne]: 'completed' } }
      });
      if (activeCase > 0) {
        return res.status(400).json({ error: 'Сначала закройте случай перегруза по секции' });
      }

      // Завершённые кейсы отвязываем (сохраняем историю без FK-сирот)? Нет — оставляем
      // как есть; удаление секции с завершёнными кейсами оставит их без секции —
      // но FK NOT NULL. Поэтому удаляем завершённые кейсы этой секции.
      await OverloadCase.destroy({ where: { sectionId: section.id } });

      await section.destroy();
      res.json({ success: true, message: 'Секция удалена' });
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
});

// ===== Случаи перегрузки секций (workflow, этап 3) =====

// Список кейсов (res_responsible — только свой РЭС)
app.get('/api/overload', authenticateToken, async (req, res) => {
  try {
    const { stage, resId } = req.query;
    // Каждый видит свою работу: АСКУЭ — этапы ограничения, РЭС — мероприятия.
    // Все этапы — суперадмин, Администратор и те, кто только смотрит (техблок и др.).
    const canAskue = userCan(req.user, 'overload_askue');
    const isRes = userCan(req.user, 'overload_res');
    if (!userCan(req.user, 'overload_view') && !canAskue && !isRes) return res.json([]);
    const allStages = req.user.isSuper || req.user.role === 'admin' || (!canAskue && !isRes);
    let allowedStages = null; // null = все этапы
    if (!allStages) {
      const set = new Set();
      if (canAskue) { set.add('askue_limit'); set.add('awaiting_recheck'); }
      if (isRes) { set.add('res_work'); set.add('awaiting_recheck'); }
      allowedStages = [...set];
    }

    const where = {};
    // Область: должность «свой РЭС» — свой; «все РЭС» — ?resId либо все.
    if (!req.user.allRes) where.resId = req.user.resId;
    else if (resId) where.resId = parseInt(resId, 10);

    // Этап: явный ?stage — только из своей области (чужой этап → пусто).
    if (stage) {
      if (allowedStages === null) where.stage = stage;
      else if ((allowedStages || []).includes(stage)) where.stage = stage;
      else return res.json([]);
    } else if (allowedStages !== null) {
      if (allowedStages.length === 0) return res.json([]);
      where.stage = { [Op.in]: allowedStages };
    }

    const cases = await OverloadCase.findAll({
      where,
      include: [
        { model: TpSection, as: 'section', attributes: ['id', 'tpName', 'sectionNumber', 'tnKva'] },
        { model: ResUnit, attributes: ['id', 'name'] },
        { model: User, as: 'askueUser', attributes: ['id', 'fio'] },
        { model: User, as: 'resUser', attributes: ['id', 'fio'] }
      ],
      order: [['updatedAt', 'DESC']],
      limit: 2000
    });
    res.json(cases);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// АСКУЭ выполнила ограничение → этап РЭС
app.post('/api/overload/:caseId/askue-complete',
  authenticateToken,
  requirePerm('overload_askue'),
  async (req, res) => {
    try {
      const { comment, consumers, limitedKw, unpolledKw } = req.body;
      // Комментарий обязателен (что именно ограничили) — как у РЭС, мин. 5 слов.
      if (!comment || typeof comment !== 'string') {
        return res.status(400).json({ error: 'Комментарий обязателен' });
      }
      if (comment.trim().split(/\s+/).filter(w => w.length > 0).length < 5) {
        return res.status(400).json({ error: 'Комментарий должен содержать не менее 5 слов' });
      }

      // Числа АСКУЭ: мощности — целые 0–9999, потребители — целые 1–999.
      const intPower = (v, name) => {
        const n = Math.round(Number(v));
        if (!Number.isFinite(n) || n < 0 || n > 9999) throw new Error(`${name}: допустимо целое 0–9999 кВт`);
        return n;
      };
      let consumersN, limitedN, unpolledN, totalN;
      try {
        limitedN = intPower(limitedKw, 'Введённые ограничения');
        unpolledN = intPower(unpolledKw, 'Не введённые (нет опроса)');
        const c = Math.round(Number(consumers));
        if (!Number.isFinite(c) || c < 1 || c > 999) throw new Error('Потребители: допустимо целое 1–999');
        consumersN = c;
        totalN = Math.round(limitedN + unpolledN);   // общая разрешённая мощность ТП
      } catch (ve) {
        return res.status(400).json({ error: ve.message });
      }

      const oc = await OverloadCase.findByPk(req.params.caseId, {
        include: [{ model: TpSection, as: 'section', attributes: ['tpName', 'sectionNumber'] }]
      });
      if (!oc) return res.status(404).json({ error: 'Случай не найден' });
      if (oc.stage !== 'askue_limit') {
        return res.status(400).json({ error: 'Случай не на этапе АСКУЭ' });
      }
      if (!req.user.allRes && oc.resId !== req.user.resId) {
        return res.status(403).json({ error: 'Чужой РЭС' });
      }
      const askueData = {
        consumers: consumersN, limitedKw: limitedN, unpolledKw: unpolledN,
        totalKw: totalN, at: new Date().toISOString()
      };
      await oc.update({
        stage: 'res_work',
        askueUserId: req.user.id,
        askueCompletedAt: new Date(),
        askueComment: comment.trim(),
        askueData
      });
      // Снимок АСКУЭ дублируем на секцию — для колонок значков в «Структуре сети».
      await TpSection.update({ askueData }, { where: { id: oc.sectionId } });

      // Уведомление стадии АСКУЭ убрать, создать для РЭС
      await removeSectionOverloadNotifs(oc.sectionId);
      const s = oc.section || {};
      await Notification.create({
        type: 'power_overload', resId: oc.resId, toUserId: null, fromUserId: req.user.id,
        message: `Требуются мероприятия РЭС по перегрузу ${s.tpName} СШ-${s.sectionNumber}`,
        errorData: {
          sectionId: oc.sectionId, caseId: oc.id, stage: 'res_work', audience: 'res',
          tpName: s.tpName, sectionNumber: s.sectionNumber,
          peakKw: oc.peakKw, peakAt: oc.peakAt, limitKw: oc.limitKw, ratio: oc.ratio,
          tnKva: oc.tnKva, cosPhi: oc.cosPhi, period: oc.period
        }
      });
      res.json({ success: true });
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
});

// РЭС выполнил мероприятия (комментарий + фото) → ожидает перепроверки
app.post('/api/overload/:caseId/res-complete',
  authenticateToken,
  requirePerm('overload_res'),
  upload.array('attachments', 5),
  async (req, res) => {
    const uploadedFiles = [];
    try {
      const { comment } = req.body;
      if (!comment || typeof comment !== 'string') {
        return res.status(400).json({ error: 'Комментарий обязателен' });
      }
      if (comment.trim().split(/\s+/).filter(w => w.length > 0).length < 5) {
        return res.status(400).json({ error: 'Комментарий должен содержать не менее 5 слов' });
      }

      const oc = await OverloadCase.findByPk(req.params.caseId, {
        include: [{ model: TpSection, as: 'section', attributes: ['tpName', 'sectionNumber'] }]
      });
      if (!oc) return res.status(404).json({ error: 'Случай не найден' });
      if (oc.stage !== 'res_work') return res.status(400).json({ error: 'Случай не на этапе мероприятий РЭС' });
      if (!req.user.allRes && oc.resId !== req.user.resId) {
        return res.status(403).json({ error: 'Чужой РЭС' });
      }

      let attachments = [];
      if (req.files && req.files.length > 0) {
        try {
          attachments = await Promise.all(req.files.map(file => uploadToCloudinary(file, req.body.type)));
          uploadedFiles.push(...attachments.map(a => a.public_id));
        } catch (uploadError) {
          await cleanupCloudinary(uploadedFiles);
          return res.status(502).json({ error: 'Не удалось загрузить файлы на сервер.' });
        }
      }

      await oc.update({
        stage: 'awaiting_recheck',
        resUserId: req.user.id,
        resCompletedAt: new Date(),
        resComment: comment,
        attachments
      });

      // Уведомление РЭС убрать, создать для АСКУЭ «ожидает перепроверки»
      await removeSectionOverloadNotifs(oc.sectionId);
      const s = oc.section || {};
      await Notification.create({
        type: 'power_overload', resId: oc.resId, toUserId: null, fromUserId: req.user.id,
        message: `Мероприятия РЭС выполнены, ожидает перепроверки профилем: ${s.tpName} СШ-${s.sectionNumber}`,
        errorData: {
          sectionId: oc.sectionId, caseId: oc.id, stage: 'awaiting_recheck', audience: 'askue',
          tpName: s.tpName, sectionNumber: s.sectionNumber,
          peakKw: oc.peakKw, peakAt: oc.peakAt, limitKw: oc.limitKw, ratio: oc.ratio,
          tnKva: oc.tnKva, cosPhi: oc.cosPhi, period: oc.period
        }
      });
      res.json({ success: true });
    } catch (error) {
      await cleanupCloudinary(uploadedFiles);
      res.status(500).json({ error: error.message });
    }
});

// Удаление кейсов перегруза (админ, право notifications_delete — то же, что у
// удаления уведомлений). Чистит и висящие power_overload-уведомления секций.
app.post('/api/overload/delete-bulk',
  authenticateToken,
  requirePerm('notifications_delete'),
  async (req, res) => {
    try {
      const ids = Array.isArray(req.body?.ids) ? req.body.ids.map(Number).filter(Boolean) : [];
      if (ids.length === 0) return res.status(400).json({ error: 'Не выбраны записи для удаления' });
      const cases = await OverloadCase.findAll({ where: { id: { [Op.in]: ids } }, attributes: ['id', 'sectionId'] });
      const sectionIds = [...new Set(cases.map(c => c.sectionId).filter(Boolean))];
      const deleted = await OverloadCase.destroy({ where: { id: { [Op.in]: ids } } });
      // Убрать возможные висящие уведомления перегруза по затронутым секциям.
      for (const sid of sectionIds) {
        try { await removeSectionOverloadNotifs(sid); } catch { /* noop */ }
      }
      res.json({ success: true, deleted });
    } catch (error) {
      console.error('overload delete-bulk error:', error.message);
      res.status(500).json({ error: error.message });
    }
});

// Обработка файла профиля мощности (Пирамида): анализатор → матчинг секций по
// техучёту → обновление секций/кейсов/уведомлений. Переиспользуется роутом
// /api/upload/analyze и почтовым приёмником. Файл НЕ удаляет (это делает вызывающий).
async function processProfileFile(filePath, userId) {
  const analysis = await runProfileAnalyzer(filePath);
  if (!analysis || !analysis.success) {
    return { success: false, error: analysis?.error || 'Ошибка анализа профиля' };
  }

  const results = analysis.results || [];
  const analyzerWarnings = analysis.warnings || [];

  const normPu = (v) => {
    let s = String(v == null ? '' : v).trim();
    if (s.endsWith('.0') && /^\d+\.0$/.test(s)) s = s.slice(0, -2);
    return s;
  };

  const allSections = await TpSection.findAll({ where: { techPuNumber: { [Op.ne]: null } } });
  const sectionByTechPu = {};
  allSections.forEach(s => {
    const key = normPu(s.techPuNumber);
    if (key) sectionByTechPu[key] = s;
  });

  let sectionsUpdated = 0;
  let overloadCount = 0;
  const unmatched = [];
  const details = [];

  for (const r of results) {
    const section = sectionByTechPu[normPu(r.puNumber)];
    if (!section) {
      unmatched.push(r.puNumber);
      details.push({
        puNumber: r.puNumber, matched: false, sectionId: null, tpSection: null,
        peakRaw: r.peakRaw, kt: r.kt, ktRaw: r.ktRaw || null, peakKw: r.peakKw, peakAt: r.peakAt,
        tnKva: null, cosPhi: null, limitKw: null, decision: 'not_matched'
      });
      continue;
    }

    const cosPhi = Number.isFinite(section.cosPhi) ? section.cosPhi : 0.9;
    const hasLimit = Number.isFinite(section.tnKva) && section.tnKva > 0;
    const limitKw = hasLimit ? section.tnKva * cosPhi : null;
    let overloadStatus = 'unknown';
    if (hasLimit) overloadStatus = (r.peakKw >= limitKw) ? 'overload' : 'ok';

    const period = r.period || analysis.period || null;
    const ratio = limitKw ? (r.peakKw / limitKw) : null;

    details.push({
      puNumber: r.puNumber, matched: true, sectionId: section.id,
      tpSection: `${section.tpName} СШ-${section.sectionNumber}`,
      peakRaw: r.peakRaw, kt: r.kt, ktRaw: r.ktRaw || null, peakKw: r.peakKw, peakAt: r.peakAt,
      tnKva: hasLimit ? section.tnKva : null, cosPhi: hasLimit ? cosPhi : null,
      limitKw, decision: overloadStatus
    });

    const peakAtDate = parsePeakAt(r.peakAt);

    await section.update({
      lastPeakKw: r.peakKw,
      lastPeakAt: peakAtDate,
      lastProfilePeriod: period,
      lastProfileSource: r.source || null,
      lastProfileAt: new Date(),
      overloadStatus
    });
    sectionsUpdated++;

    // Помесячный пик (раздел «Анализ мощности»): месяц по peakAt, иначе по периоду.
    // «Максимум побеждает» одним запросом ON CONFLICT — без гонки read-then-write.
    let ym = null;
    if (peakAtDate) {
      ym = { year: peakAtDate.getFullYear(), month: peakAtDate.getMonth() + 1 };
    } else if (period) {
      const pm = String(period).match(/(\d{2})\.(\d{2})\.(\d{4})/);
      if (pm) ym = { year: Number(pm[3]), month: Number(pm[2]) };
    }
    if (ym) {
      try {
        await sequelize.query(
          `INSERT INTO "SectionMonthlyPeaks"
             ("sectionId","year","month","peakKw","peakAt","source","createdAt","updatedAt")
           VALUES (:sectionId,:year,:month,:peakKw,:peakAt,:source, NOW(), NOW())
           ON CONFLICT ("sectionId","year","month") DO UPDATE
             SET "peakKw"=EXCLUDED."peakKw", "peakAt"=EXCLUDED."peakAt",
                 "source"=EXCLUDED."source", "updatedAt"=NOW()
           WHERE EXCLUDED."peakKw" > "SectionMonthlyPeaks"."peakKw"`,
          { replacements: {
              sectionId: section.id, year: ym.year, month: ym.month,
              peakKw: r.peakKw, peakAt: peakAtDate, source: r.source || null
          } }
        );
      } catch (e) {
        console.error('[PROFILE] monthly peak upsert error:', e.message);
      }
    }

    if (!hasLimit) continue;

    const activeCase = await OverloadCase.findOne({
      where: { sectionId: section.id, stage: { [Op.ne]: 'completed' } },
      order: [['id', 'DESC']]
    });

    const snapshot = () => ({
      peakKw: r.peakKw, peakAt: r.peakAt, tnKva: section.tnKva,
      cosPhi, limitKw, ratio, period
    });
    const notifData = (caseId, stage) => ({
      sectionId: section.id, caseId, stage,
      tpName: section.tpName, sectionNumber: section.sectionNumber,
      techPuNumber: section.techPuNumber,
      peakKw: r.peakKw, peakAt: r.peakAt, tnKva: section.tnKva,
      cosPhi, limitKw, ratio, period
    });

    if (overloadStatus === 'overload') {
      overloadCount++;
      if (!activeCase) {
        const oc = await OverloadCase.create({
          sectionId: section.id, resId: section.resId,
          stage: 'askue_limit', cycles: 1, ...snapshot()
        });
        await removeSectionOverloadNotifs(section.id);
        await Notification.create({
          type: 'power_overload', resId: section.resId, toUserId: null, fromUserId: userId,
          message: `Превышение Pном: ${section.tpName} СШ-${section.sectionNumber} — пик ${r.peakKw} кВт при лимите ${limitKw} кВт`,
          errorData: { ...notifData(oc.id, 'askue_limit'), audience: 'askue' }
        });
      } else if (activeCase.stage === 'awaiting_recheck') {
        const newCycles = activeCase.cycles + 1;
        await activeCase.update({
          recheckAt: new Date(), recheckPeakKw: r.peakKw, recheckResult: 'still_overload',
          stage: 'askue_limit', cycles: newCycles, ...snapshot()
        });
        await removeSectionOverloadNotifs(section.id);
        await Notification.create({
          type: 'power_overload', resId: section.resId, toUserId: null, fromUserId: userId,
          message: `Повторный перегруз после мероприятий (цикл ${newCycles}): ${section.tpName} СШ-${section.sectionNumber} — пик ${r.peakKw} кВт`,
          errorData: { ...notifData(activeCase.id, 'askue_limit'), cycle: newCycles, audience: 'askue' }
        });
        // Эскалация в «Проблемные ТП» — ОДИН раз при достижении порога (cycles 1→2).
        if (newCycles === PROBLEM_TP_CYCLES) {
          await Notification.create({
            type: 'power_overload', resId: section.resId, toUserId: null, fromUserId: userId,
            message: `ТП ${section.tpName} СШ-${section.sectionNumber} переведена в проблемные: мероприятия не дали результата (цикл ${newCycles})`,
            errorData: { ...notifData(activeCase.id, 'askue_limit'), cycle: newCycles, problemTp: true }
          });
        }
      } else {
        await activeCase.update(snapshot());
      }
    } else {
      if (activeCase && activeCase.stage === 'awaiting_recheck') {
        await activeCase.update({
          recheckAt: new Date(), recheckPeakKw: r.peakKw, recheckResult: 'ok',
          stage: 'completed', closedAt: new Date()
        });
        await removeSectionOverloadNotifs(section.id);
        await Notification.create({
          type: 'success', resId: section.resId, toUserId: null, fromUserId: userId,
          message: `Перегруз ${section.tpName} СШ-${section.sectionNumber} устранён, пик ${r.peakKw} кВт при лимите ${limitKw} кВт`
        });
      } else if (activeCase) {
        await activeCase.update({ stage: 'completed', recheckResult: 'ok', closedAt: new Date() });
        await removeSectionOverloadNotifs(section.id);
      }
    }
  }

  // Сводный лог (детализация по каждому ПУ — в details[] ответа и в письме-ответе).
  const allUnmatched = [...unmatched, ...analyzerWarnings];
  console.log(`[PROFILE] обработано секций: ${sectionsUpdated}, перегрузов: ${overloadCount}, ` +
    `не сопоставлено: ${allUnmatched.length}${allUnmatched.length ? ' (' + allUnmatched.join(', ') + ')' : ''}`);

  return {
    success: true,
    sectionsUpdated,
    overloadCount,
    unmatched: allUnmatched,
    details,
    period: analysis.period || null
  };
}

// 5. ЗАГРУЗКА ФАЙЛОВ ДЛЯ АНАЛИЗА
app.post('/api/upload/analyze',
  authenticateToken,
  requirePerm('upload_files'),
  uploadExcel.single('file'),
  async (req, res) => {
    let uploadRecord;
    try {
      const { type, requiredPeriod } = req.body;
      const userId = req.user.id;
      
      // Берем resId из body (если есть) или из токена пользователя
      const resId = req.body.resId || req.user.resId;

      // ── Профиль мощности (Пирамида) — матчинг по ПУ техучёта, resId не нужен ──
      if (type === 'profile') {
        try {
          uploadRecord = await UploadHistory.create({
            userId,
            resId: resId || null,
            fileName: req.file.originalname,
            fileType: 'profile',
            status: 'processing'
          });
        } catch (createErr) {
          // Не глотаем: раньше это выглядело как «профиль ничего не нашёл».
          console.error('[PROFILE] UploadHistory.create failed:', createErr.message);
          try { fs.unlinkSync(req.file.path); } catch (e) {}
          return res.status(400).json({ success: false, error: `Не удалось создать запись загрузки: ${createErr.message}` });
        }

        const result = await processProfileFile(req.file.path, userId);
        try { fs.unlinkSync(req.file.path); } catch (e) {}

        if (!result.success) {
          await uploadRecord.update({ status: 'failed' });
          return res.status(400).json({ success: false, error: result.error });
        }

        await uploadRecord.update({
          processedCount: result.sectionsUpdated,
          errorCount: result.overloadCount,
          status: 'completed'
        });

        return res.json({
          success: true,
          message: 'Профиль обработан',
          type: 'profile',
          sectionsUpdated: result.sectionsUpdated,
          overloadCount: result.overloadCount,
          unmatched: result.unmatched,
          details: result.details
        });
      }

      if (!resId) {
        return res.status(400).json({ error: 'Не выбран РЭС для загрузки' });
      }
      
      // Создаем запись в истории
      uploadRecord = await UploadHistory.create({
        userId,
        resId,
        fileName: req.file.originalname,
        fileType: type,
        status: 'processing'
      });
      
      console.log('Upload record created:', uploadRecord.id);
      
      // Запускаем анализ с передачей оригинального имени файла
      console.log('Starting analysis...');
      // Гарантированное удаление temp-файла на ЛЮБОМ исходе analyzeFile
      // (wrong_period, code!==0, таймаут, «Python недоступен», ошибка JSON) — finally.
      try {
      const analysisResult = await analyzeFile(
        req.file.path,
        type,
        req.file.originalname, // передаем оригинальное имя
        requiredPeriod,
        userId
      );

      console.log('Analysis result:', {
        processed: analysisResult.processed.length,
        errors: analysisResult.errors.length
      });

      // Обновляем историю
      await uploadRecord.update({
        processedCount: analysisResult.processed.length,
        errorCount: analysisResult.errors.length,
        status: 'completed'
      });

      // Отправляем уведомления если есть ошибки
      if (analysisResult.errors.length > 0) {
        console.log(`Creating notifications for ${analysisResult.errors.length} errors`);
        try {
          await createNotifications(userId, resId, analysisResult.errors);
          console.log('Notifications created successfully');
        } catch (notifError) {
          console.error('Error creating notifications:', notifError);
          // НЕ падаем, продолжаем работу!
        }
      }

      // Возвращаем результат
      res.json({
        success: true,
        message: 'Файл обработан успешно',
        processed: analysisResult.processed.length,
        errors: analysisResult.errors.length,
        details: analysisResult.processed
      });
      } finally {
        try { fs.unlinkSync(req.file.path); } catch (e) {}
      }

    } catch (error) {
      console.error('Upload analyze error:', error);
      
      // Обновляем статус в истории
      if (uploadRecord) {
        await uploadRecord.update({ status: 'failed' });
      }
      
      res.status(500).json({ 
        success: false,
        error: error.message 
      });
    }
});

// 6. ЗАГРУЗКА ПОЛНОЙ СТРУКТУРЫ СЕТИ
app.post('/api/network/upload-full-structure',
  authenticateToken,
  requirePerm('structure_upload'),
  uploadExcel.single('file'),
  async (req, res) => {
    const transaction = await sequelize.transaction();
    
    try {
      const workbook = XLSX.readFile(req.file.path);
      const sheetName = workbook.SheetNames[0];
      const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
      
      let processed = 0;
      let errors = [];

      // PERF: все РЭС одним запросом в Map по имени (было — findOne на строку).
      const resUnits = await ResUnit.findAll({ transaction });
      const resByName = new Map(resUnits.map(r => [r.name, r]));

      // PuStatus для новых ПУ копим и создаём одним bulkCreate после цикла (было —
      // findOrCreate на каждую позицию). Первое вхождение ПУ выигрывает, как и у
      // findOrCreate; существующие в БД не трогаем.
      const puToCreate = new Map(); // puNumber -> { networkStructureId, position, status }

      // Кэш секций ТП в рамках загрузки: key "resId|tpName|sectionNumber" -> instance.
      const sectionCache = new Map();
      // Тех.учёт, заявленный в файле по секции (key секции -> номер ПУ). Нужен для
      // проверки: у одной секции шин тех.учёт должен быть ОДИН, даже если линий много.
      const techBySection = new Map();
      const LINE_TYPES = ['ВЛ', 'КЛ', 'КВЛ', 'ВУ'];
      const has = (v) => v !== undefined && v !== null && String(v).trim() !== '';

      // MERGE: предзагружаем существующие линии РЭСов из файла. Пустые ячейки НЕ
      // затирают существующие значения — обновляется только заполненное. Новые
      // линии создаются; пропущенные поля → дефолты (тип 'ВЛ', ПУ/секция пусто).
      const fileResIds = [...new Set(data.map(r => resByName.get(r['РЭС'])?.id).filter(Boolean))];
      const existingLines = fileResIds.length
        ? await NetworkStructure.findAll({ where: { resId: fileResIds }, transaction })
        : [];
      const lineByKey = new Map(existingLines.map(l => [`${l.resId}|${l.tpName}|${l.vlName}`, l]));

      // Обрабатываем каждую строку
      for (const row of data) {
        try {
          const resName = row['РЭС'];
          const resUnit = resByName.get(resName);
          if (!resUnit) {
            errors.push(`Неизвестный РЭС: ${resName}`);
            continue;
          }

          const tpName = row['ТП'] || '';
          const vlName = row['Фидер'] || '';

          // Тип линии: задан → проверяем (недопустимый → в errors, строку пропуск);
          // пусто/нет колонки → НЕ трогаем (у существующей сохранится, у новой 'ВЛ').
          let lineType = null, lineTypeProvided = false;
          if (has(row['Тип'])) {
            const t = String(row['Тип']).trim();
            if (!LINE_TYPES.includes(t)) {
              errors.push(`Недопустимый тип линии "${t}" (${tpName} · ${vlName}) — строка пропущена`);
              continue;
            }
            lineType = t; lineTypeProvided = true;
          }

          // Секция шин + номинал трансформатора (Sном, кВА). Пусто → секцию не
          // трогаем. Задана → привязать/создать; при наличии Sном — записать на секцию.
          let sectionId = null, sectionProvided = false;
          if (has(row['Секция'])) {
            const sectionNumber = parseInt(String(row['Секция']).trim(), 10);
            if (Number.isNaN(sectionNumber)) {
              errors.push(`Некорректный номер секции "${row['Секция']}" (${tpName} · ${vlName}) — секция не изменена`);
            } else {
              const cacheKey = `${resUnit.id}|${tpName}|${sectionNumber}`;
              let section = sectionCache.get(cacheKey);
              if (!section) {
                section = await TpSection.findOne({ where: { resId: resUnit.id, tpName, sectionNumber }, transaction });
                if (!section) {
                  section = await TpSection.create({ resId: resUnit.id, tpName, sectionNumber }, { transaction });
                }
                sectionCache.set(cacheKey, section);
              }
              sectionId = section.id; sectionProvided = true;

              // Параметры секции — обновляем ТОЛЬКО заполненные (merge). Пусто → не трогаем.
              const secUpd = {};
              const kvaRaw = row['Sном, кВА'] ?? row['Sном'];
              if (has(kvaRaw)) {
                const kva = parseFloat(String(kvaRaw).replace(',', '.'));
                if (!Number.isNaN(kva) && section.tnKva !== kva) secUpd.tnKva = kva;
              }
              const techRaw = row['Тех.учёт ПУ'];
              if (has(techRaw)) {
                const tech = String(techRaw).trim();
                // Одна секция шин = один тех.учёт. Если по этой же секции в файле
                // уже встречался ДРУГОЙ тех.учёт (много линий на секции) — это ошибка,
                // такой тех.учёт не пишем. Разные тех.учёты допустимы лишь на разных СШ.
                const prevTech = techBySection.get(cacheKey);
                if (prevTech !== undefined && prevTech !== tech) {
                  errors.push(`Разные тех.учёты на одной секции СШ-${sectionNumber} (${tpName}): «${prevTech}» и «${tech}» — тех.учёт не изменён`);
                } else {
                  techBySection.set(cacheKey, tech);
                  if (section.techPuNumber !== tech) secUpd.techPuNumber = tech;
                }
              }
              if (Object.keys(secUpd).length > 0) await section.update(secUpd, { transaction });
            }
          }

          // Merge линии: у существующей обновляем ТОЛЬКО заполненные поля; новую создаём.
          const key = `${resUnit.id}|${tpName}|${vlName}`;
          let structureId;
          const existing = lineByKey.get(key);
          if (existing) {
            const upd = {};
            if (has(row['Начало'])) upd.startPu = String(row['Начало']);
            if (has(row['Середина'])) upd.middlePu = String(row['Середина']);
            if (has(row['Конец'])) upd.endPu = String(row['Конец']);
            if (lineTypeProvided) upd.lineType = lineType;
            if (sectionProvided) upd.sectionId = sectionId;
            if (Object.keys(upd).length > 0) {
              upd.lastUpdate = new Date();
              await existing.update(upd, { transaction });
            }
            structureId = existing.id;
          } else {
            const created = await NetworkStructure.create({
              resId: resUnit.id, tpName, vlName,
              startPu: has(row['Начало']) ? String(row['Начало']) : null,
              middlePu: has(row['Середина']) ? String(row['Середина']) : null,
              endPu: has(row['Конец']) ? String(row['Конец']) : null,
              lineType: lineTypeProvided ? lineType : 'ВЛ',
              sectionId: sectionProvided ? sectionId : null
            }, { transaction });
            lineByKey.set(key, created);
            structureId = created.id;
          }

          processed++;

          // Статусы только для ЗАПОЛНЕННЫХ ПУ (первое вхождение выигрывает).
          const positions = [
            { pu: row['Начало'], pos: 'start' },
            { pu: row['Конец'], pos: 'end' },
            { pu: row['Середина'], pos: 'middle' }
          ];
          for (const { pu, pos } of positions) {
            if (has(pu)) {
              const k = String(pu);
              if (!puToCreate.has(k)) {
                puToCreate.set(k, { puNumber: k, networkStructureId: structureId, position: pos, status: 'not_checked' });
              }
            }
          }
        } catch (err) {
          errors.push(`Ошибка в строке ${row['ТП']}-${row['Фидер']}: ${err.message}`);
        }
      }

      // Недостающие PuStatus — одним запросом; существующие не трогаем (как findOrCreate).
      if (puToCreate.size > 0) {
        const puNumbers = [...puToCreate.keys()];
        const existing = await PuStatus.findAll({
          where: { puNumber: puNumbers },
          attributes: ['puNumber'],
          transaction
        });
        const existingSet = new Set(existing.map(e => e.puNumber));
        const toInsert = [...puToCreate.values()].filter(p => !existingSet.has(p.puNumber));
        if (toInsert.length > 0) {
          await PuStatus.bulkCreate(toInsert, { transaction });
        }
      }

      await transaction.commit();
      
      // Удаляем файл
      fs.unlinkSync(req.file.path);
      
      res.json({
        success: true,
        message: `Загружено ${processed} записей из ${data.length}`,
        processed,
        total: data.length,
        errors: errors.length > 0 ? errors.slice(0, 10) : []
      });
      
    } catch (error) {
      await transaction.rollback();
      res.status(500).json({ error: error.message });
    }
});

// Единый порядок колонок формата структуры (шаблон = выгрузка = формат загрузки).
const STRUCTURE_COLS = ['РЭС', 'ТП', 'Секция', 'Sном, кВА', 'Тех.учёт ПУ', 'Тип', 'Фидер', 'Начало', 'Середина', 'Конец'];
const STRUCTURE_COL_WIDTHS = [{ wch: 22 }, { wch: 12 }, { wch: 8 }, { wch: 10 }, { wch: 14 }, { wch: 8 }, { wch: 20 }, { wch: 14 }, { wch: 14 }, { wch: 14 }];

// Шаблон Excel для загрузки структуры (формат с типом линии, секцией и Sном
// трансформатора). Отдаётся через уже подключённый XLSX. Право — structure_upload.
app.get('/api/network/structure-template',
  authenticateToken,
  requirePerm('structure_upload'),
  (req, res) => {
    try {
      const rows = [
        STRUCTURE_COLS,
        ['Центральный РЭС', 'ТП-101', 1, 400, '99887766', 'ВЛ', 'Фабрициуса', '12345678', '', '23456789'],
        ['Центральный РЭС', 'ТП-101', 2, 250, '', 'КЛ', 'Ленина', '34567890', '45678901', ''],
        ['Центральный РЭС', 'ТП-205', '', '', '', 'ВУ', 'Ввод-1', '56789012', '', ''],
      ];
      const ws = XLSX.utils.aoa_to_sheet(rows);
      ws['!cols'] = STRUCTURE_COL_WIDTHS;
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Структура');
      const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.setHeader('Content-Disposition', 'attachment; filename="structure-template.xlsx"');
      res.send(buf);
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
});

// Выгрузка ТЕКУЩЕЙ структуры сети в том же формате, что и шаблон загрузки.
// Правишь в Excel и заливаешь обратно — загрузка merge'ит: пустые ячейки не
// затирают, обновляется только заполненное. Просмотровый роут (по роли/РЭС).
app.get('/api/network/structure-export', authenticateToken, async (req, res) => {
  try {
    const resId = req.user.allRes ? (req.query.resId || null) : req.user.resId;
    if (!req.user.allRes && req.query.resId && req.query.resId != req.user.resId) {
      return res.status(403).json({ error: 'Access denied' });
    }
    const where = {};
    if (resId) where.resId = resId;
    const lines = await NetworkStructure.findAll({
      where,
      include: [
        { model: TpSection, as: 'section', required: false, attributes: ['sectionNumber', 'tnKva', 'techPuNumber'] },
        { model: ResUnit, attributes: ['name'] }
      ],
      order: [['tpName', 'ASC'], ['vlName', 'ASC']]
    });
    const rows = [STRUCTURE_COLS];
    for (const l of lines) {
      rows.push([
        l.ResUnit?.name || '',
        l.tpName || '',
        l.section?.sectionNumber ?? '',
        l.section?.tnKva ?? '',
        l.section?.techPuNumber ?? '',
        l.lineType || 'ВЛ',
        l.vlName || '',
        l.startPu || '',
        l.middlePu || '',
        l.endPu || ''
      ]);
    }
    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws['!cols'] = STRUCTURE_COL_WIDTHS;
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Структура');
    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="structure-export.xlsx"');
    res.send(buf);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 7. ПОЛУЧЕНИЕ УВЕДОМЛЕНИЙ

app.get('/api/notifications', authenticateToken, async (req, res) => {
  try {
    const { resId, type } = req.query;   // ← добавили type
    let whereClause = {};

    // Что видно — по правам должности, где — по её области (все РЭС или свой).
    const u = req.user;
    const canAskue = userCan(u, 'overload_askue');
    const isRes = userCan(u, 'overload_res');
    const types = ['success', 'info'];
    if (userCan(u, 'tech_pending_view') || userCan(u, 'res_work')) types.push('error');
    if (userCan(u, 'askue_pending_view') || userCan(u, 'upload_files')) types.push('pending_askue');
    if (userCan(u, 'overload_view') || canAskue || isRes) types.push('power_overload');
    if (userCan(u, 'problem_vl_view')) types.push('problem_vl');

    if (u.allRes) {
      whereClause = resId ? { resId: parseInt(resId) } : {};
      whereClause.type = { [Op.in]: types };
    } else {
      whereClause = {
        [Op.or]: [
          { toUserId: u.id },
          { toUserId: null, resId: u.resId ?? -1, type: { [Op.in]: types } }
        ]
      };
    }

    // ← НОВОЕ: фильтруем по типу прямо в БД, чтобы error не вытеснялись
    //   свежими уведомлениями других типов
    if (type) {
      if (!types.includes(type)) return res.json([]);
      if (u.allRes) whereClause.type = type;
      else whereClause = { [Op.and]: [whereClause, { type }] };
    }
    
    const notifications = await Notification.findAll({
      where: whereClause,
      include: [
        // ✅ PERF+SECURITY: только нужные поля, без хэша пароля
        { model: User, as: 'fromUser', attributes: ['id', 'fio', 'role'] },
        { model: User, as: 'toUser', attributes: ['id', 'fio', 'role'] },
        { model: ResUnit, attributes: ['id', 'name'] },
        { model: NetworkStructure, attributes: ['id', 'tpName', 'vlName'] },
        // ✅ PERF: отметки прочтения только ТЕКУЩЕГО пользователя,
        // а не всех пользователей системы (раньше — взрыв payload'а)
        {
          model: NotificationRead,
          required: false,
          where: { userId: req.user.id },
          attributes: ['id', 'userId']
        }
      ],
      order: [['createdAt', 'DESC']],
      limit: 2000   // ← было 100; подняли, чтобы влезли все мероприятия
    });
    
    const notificationsWithReadStatus = notifications.map(notif => {
      const isRead = notif.NotificationReads?.some(read => 
        read.userId === req.user.id
      ) || false;
      
      return {
        ...notif.toJSON(),
        isRead: isRead
      };
    });
    
    // Адресация power_overload по этапу (audience): 'askue' — пользователям с правом
    // overload_askue (и админам/суперу); 'res'/без метки — РЭС по resId (уже в
    // whereClause) + админам. Так этап АСКУЭ не сыпется РЭСу и наоборот.
    const seesAll = u.isSuper || u.role === 'admin' || (!canAskue && !isRes);
    const audienceFiltered = notificationsWithReadStatus.filter(notif => {
      if (notif.type !== 'power_overload') return true;
      if (seesAll) return true;
      const aud = notif.errorData && notif.errorData.audience;
      if (aud === 'askue') return canAskue;
      return isRes;   // 'res'/без метки → мероприятия РЭС
    });

    // ✅ Дедупликация: для error и pending_askue — 1 ПУ + 1 набор фаз = 1 запись (последняя)
    const seenKeys = new Set();
    const deduplicatedNotifications = audienceFiltered.filter(notif => {
      if (notif.type !== 'error' && notif.type !== 'pending_askue') {
        return true; // Остальные типы не трогаем
      }
      
      try {
        const data = JSON.parse(notif.message);
        if (!data.puNumber) return true;

        const phaseKey = getPhaseSignature(data); // уже распарсено — не парсим повторно
        const key = `${notif.type}_${data.puNumber}_${phaseKey}`;
        
        if (seenKeys.has(key)) {
          return false; // Дубликат — уже видели более свежее (список отсортирован DESC)
        }
        
        seenKeys.add(key);
        return true;
      } catch {
        return true;
      }
    });
    
    res.json(deduplicatedNotifications);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 8. ВЫПОЛНЕНИЕ МЕРОПРИЯТИЙ
const ESCALATION_MIN_WORDS = 15;

app.post('/api/notifications/:id/complete-work', 
  authenticateToken, 
  requirePerm('res_work'),
  upload.array('attachments', 5),
  async (req, res) => {
    const startTime = Date.now();
    console.log('\n=== COMPLETE WORK START ===');
    console.log('Files received:', req.files?.length || 0);
    
    const uploadedFiles = []; // public_id'ы для cleanup при ошибке
    
    try {
      const { comment, checkFromDate } = req.body;
      // «Проблему решить невозможно, нужны доп. мероприятия» — вопрос уходит техблоку.
      const escalate = req.body.escalate === 'true' || req.body.escalate === true;
      
      // 1) ВАЛИДАЦИЯ ДО ВСЕГО ОСТАЛЬНОГО (быстрый отказ без транзакции и без загрузок)
      if (!comment || typeof comment !== 'string') {
        return res.status(400).json({ error: escalate ? 'Опишите проблему для техблока' : 'Комментарий обязателен' });
      }
      const wordCount = comment.trim().split(/\s+/).filter(w => w.length > 0).length;
      if (escalate && wordCount < ESCALATION_MIN_WORDS) {
        return res.status(400).json({ error: `Опишите проблему подробнее — не менее ${ESCALATION_MIN_WORDS} слов: что мешает, что сделано, что нужно` });
      }
      if (wordCount < 5) {
        return res.status(400).json({ error: 'Комментарий должен содержать не менее 5 слов' });
      }
      
      // 2) ЗАГРУЗКА ФАЙЛОВ — ВНЕ ТРАНЗАКЦИИ И ПАРАЛЛЕЛЬНО
      // Раньше файлы лились последовательно внутри транзакции,
      // из-за чего БД-коннект висел открытым 1-2 минуты.
      let attachments = [];
      if (req.files && req.files.length > 0) {
        console.log(`Uploading ${req.files.length} files to Cloudinary in parallel...`);
        try {
          attachments = await Promise.all(
            req.files.map(file => uploadToCloudinary(file, req.body.type))
          );
          uploadedFiles.push(...attachments.map(a => a.public_id));
          console.log(`✅ All ${attachments.length} files uploaded in ${Date.now() - startTime}ms`);
        } catch (uploadError) {
          console.error('❌ Cloudinary upload failed:', uploadError);
          await cleanupCloudinary(uploadedFiles);
          return res.status(502).json({ 
            error: 'Не удалось загрузить файлы на сервер. Попробуйте файлы меньшего размера или проверьте соединение.' 
          });
        }
      }
      
      // 3) КОРОТКАЯ ТРАНЗАКЦИЯ БД — открывается только сейчас
      const transaction = await sequelize.transaction();
      try {
        const notification = await Notification.findByPk(req.params.id, { transaction });
        if (!notification) {
          await transaction.rollback();
          await cleanupCloudinary(uploadedFiles);
          return res.status(404).json({ error: 'Уведомление не найдено' });
        }
        
        const errorData = JSON.parse(notification.message);
        if (!req.user.allRes && notification.resId !== req.user.resId) {
          await transaction.rollback();
          await cleanupCloudinary(uploadedFiles);
          return res.status(403).json({ error: 'Чужой РЭС' });
        }

        if (escalate) {
          // Линию — техблоку: перепроверки АСКУЭ нет, уведомления по линии сняты,
          // пока вопрос открыт — по ней не создаются новые (см. анализ съёмов).
          const esc = await Escalation.create({
            resId: notification.resId,
            networkStructureId: notification.networkStructureId,
            puNumber: errorData.puNumber,
            tpName: errorData.tpName,
            vlName: errorData.vlName,
            position: errorData.position,
            errorDetails: errorData.errorDetails,
            errorData,
            description: comment.trim(),
            attachments,
            createdBy: req.user.id,
          }, { transaction });
          await Notification.destroy({
            where: { type: 'error', networkStructureId: notification.networkStructureId, resId: notification.resId },
            transaction
          });
          await notification.destroy({ transaction });
          await transaction.commit();
          invalidateEscalated();
          return res.json({ success: true, escalated: true, id: esc.id, message: 'Передано техблоку', filesUploaded: attachments.length });
        }
        
        // Создаем запись в CheckHistory
        await CheckHistory.create({
          resId: notification.resId,
          networkStructureId: notification.networkStructureId,
          puNumber: errorData.puNumber,
          tpName: errorData.tpName,
          vlName: errorData.vlName,
          position: errorData.position,
          initialError: errorData.errorDetails,
          initialCheckDate: notification.createdAt,
          resComment: comment,
          workCompletedDate: new Date(),
          checkFromDate: checkFromDate ? new Date(checkFromDate) : new Date(),
          status: 'awaiting_recheck',
          // Сколько раз ошибка приходила ДО отработки (повторные съёмы без
          // рассмотрения РЭСом) — для признака «отработали только после N-го съёма».
          failureCount: errorData.repeatCount || 1,
          attachments: attachments
        }, { transaction });
        
        // Обновляем статус ПУ
        await PuStatus.update(
          { status: 'pending_recheck' },
          { where: { puNumber: errorData.puNumber }, transaction }
        );
        
        // Удаляем старое уведомление
        await notification.destroy({ transaction });
        
        // Создаем уведомление для АСКУЭ
        const askueMessage = {
          puNumber: errorData.puNumber,
          position: errorData.position,
          tpName: errorData.tpName,
          vlName: errorData.vlName,
          resName: errorData.resName,
          errorDetails: errorData.errorDetails,
          checkFromDate: checkFromDate || new Date().toISOString().split('T')[0],
          completedComment: comment,
          completedBy: req.user.id,
          completedAt: new Date()
        };
        
        await Notification.create({
          fromUserId: req.user.id,
          toUserId: null,
          resId: notification.resId,
          networkStructureId: notification.networkStructureId,
          type: 'pending_askue',
          message: JSON.stringify(askueMessage),
          isRead: false
        }, { transaction });
        
        await transaction.commit();
        console.log(`✅ Complete work finished successfully in ${Date.now() - startTime}ms`);
        
        return res.json({ 
          success: true, 
          message: 'Мероприятия отмечены как выполненные',
          filesUploaded: attachments.length
        });
        
      } catch (dbError) {
        await transaction.rollback();
        console.error('❌ DB error in complete-work:', dbError);
        await cleanupCloudinary(uploadedFiles);
        return res.status(500).json({ 
          error: 'Ошибка сохранения данных: ' + (dbError.message || 'неизвестная ошибка БД')
        });
      }
      
    } catch (error) {
      console.error('❌ Complete work error:', error);
      await cleanupCloudinary(uploadedFiles);
      return res.status(500).json({ 
        error: error.message || 'Внутренняя ошибка сервера' 
      });
    }
});

// =====================================================
// ТРЕБУЮТ РЕШЕНИЯ ТЕХБЛОКА
// =====================================================
// РЭС отметил «проблему решить невозможно, нужны доп. мероприятия» — вопрос
// здесь, пока руководство техблока его не завершит. Область — по должности.
const ESCALATION_DECISION_MIN_WORDS = 5;
const escalationOut = (e) => ({
  id: e.id, status: e.status,
  resId: e.resId, resName: e.ResUnit?.name || '',
  networkStructureId: e.networkStructureId,
  puNumber: e.puNumber, tpName: e.tpName, vlName: e.vlName, position: e.position,
  errorDetails: e.errorDetails, details: e.errorData?.details || null,
  firstReportDate: e.errorData?.firstReportDate || null, repeatCount: e.errorData?.repeatCount || 1,
  description: e.description, attachments: e.attachments || [],
  author: e.author ? e.author.fio : null, createdAt: e.createdAt,
  decision: e.decision, closer: e.closer ? e.closer.fio : null, closedAt: e.closedAt,
});

app.get('/api/techblock', authenticateToken, requirePerm('techblock_view'), async (req, res) => {
  try {
    const status = req.query.status === 'closed' ? 'closed' : 'open';
    const where = { status };
    if (!req.user.allRes) where.resId = req.user.resId;
    else if (req.query.resId) where.resId = parseInt(req.query.resId, 10);
    const rows = await Escalation.findAll({
      where,
      include: [
        { model: ResUnit, attributes: ['id', 'name'] },
        { model: User, as: 'author', attributes: ['id', 'fio'] },
        { model: User, as: 'closer', attributes: ['id', 'fio'] },
      ],
      order: status === 'open' ? [['createdAt', 'ASC']] : [['closedAt', 'DESC']],
      limit: 500,
    });
    res.json(rows.map(escalationOut));
  } catch (error) {
    console.error('techblock list error:', error);
    res.status(500).json({ error: error.message });
  }
});

// «Завершить»: решение принято — линия возвращается в обычный мониторинг
// (следующий съём с ошибкой снова даст уведомление РЭС). РЭС получает решение.
app.post('/api/techblock/:id/complete', authenticateToken, requirePerm('techblock_complete'), async (req, res) => {
  try {
    const decision = String(req.body?.decision || '').trim();
    if (decision.split(/\s+/).filter(Boolean).length < ESCALATION_DECISION_MIN_WORDS) {
      return res.status(400).json({ error: `Запишите решение техблока — не менее ${ESCALATION_DECISION_MIN_WORDS} слов` });
    }
    const esc = await Escalation.findByPk(req.params.id);
    if (!esc) return res.status(404).json({ error: 'Вопрос не найден' });
    if (!req.user.allRes && esc.resId !== req.user.resId) return res.status(403).json({ error: 'Чужой РЭС' });
    if (esc.status !== 'open') return res.status(400).json({ error: 'Вопрос уже завершён' });

    await sequelize.transaction(async (t) => {
      await esc.update({ status: 'closed', decision, closedBy: req.user.id, closedAt: new Date() }, { transaction: t });
      // Проблемность линии до решения не копим: прежние активные записи снимаем.
      if (esc.networkStructureId) {
        await ProblemVL.update({ status: 'dismissed' }, {
          where: { networkStructureId: esc.networkStructureId, status: 'active' }, transaction: t
        });
      }
      const closer = await User.findByPk(req.user.id, { attributes: ['fio'], transaction: t });
      await Notification.create({
        fromUserId: req.user.id,
        toUserId: null,
        resId: esc.resId,
        networkStructureId: esc.networkStructureId,
        type: 'info',
        message: `Техблок завершил вопрос по ${esc.tpName} — ${esc.vlName} (ПУ ${esc.puNumber}). Решение: ${decision}${closer ? ` (${closer.fio})` : ''}`,
        isRead: false
      }, { transaction: t });
    });
    invalidateEscalated();
    res.json({ success: true });
  } catch (error) {
    console.error('techblock complete error:', error);
    res.status(500).json({ error: error.message });
  }
});

// =====================================================
// СВОДКА (главная страница)
// =====================================================
// Одним запросом: что ждёт действия, состояние ПУ по РЭС и динамика за 12 недель
// (выявлено ошибок съёмами / отработано РЭС / передано техблоку). Область — по
// должности; «все РЭС» могут сузить ?resId.
app.get('/api/dashboard', authenticateToken, requirePerm('dashboard_view'), async (req, res) => {
  try {
    const resId = !req.user.allRes ? req.user.resId : (req.query.resId ? parseInt(req.query.resId, 10) : null);
    const rf = (col) => (resId ? `AND ${col} = :resId` : '');
    const q = (sql) => sequelize.query(sql, { replacements: { resId }, type: Sequelize.QueryTypes.SELECT });

    const [byRes, weekly, escalationsByWeek, counts] = await Promise.all([
      q(`SELECT r.id, r.name,
            COUNT(DISTINCT ns."tpName")::int AS tp,
            COUNT(DISTINCT ns.id)::int AS vl,
            COUNT(ps.id)::int AS pu,
            COUNT(ps.id) FILTER (WHERE ps.status = 'checked_ok')::int AS ok,
            COUNT(ps.id) FILTER (WHERE ps.status = 'checked_error')::int AS error,
            COUNT(ps.id) FILTER (WHERE ps.status = 'pending_recheck')::int AS pending,
            COUNT(ps.id) FILTER (WHERE ps.status NOT IN ('checked_ok','checked_error','pending_recheck'))::int AS unchecked
         FROM "ResUnits" r
         LEFT JOIN "NetworkStructures" ns ON ns."resId" = r.id
         LEFT JOIN "PuStatuses" ps ON ps."networkStructureId" = ns.id
         WHERE TRUE ${rf('r.id')}
         GROUP BY r.id, r.name ORDER BY r.name`),
      q(`WITH weeks AS (
           SELECT generate_series(date_trunc('week', now()) - interval '11 weeks', date_trunc('week', now()), interval '1 week') AS w
         ),
         found AS (
           SELECT date_trunc('week', h."createdAt") AS w, COUNT(DISTINCT h."puNumber")::int AS n
           FROM "PuUploadHistories" h
           JOIN "PuStatuses" ps ON ps."puNumber" = h."puNumber"
           JOIN "NetworkStructures" ns ON ns.id = ps."networkStructureId"
           WHERE h."hasErrors" = true AND h."createdAt" >= now() - interval '12 weeks' ${rf('ns."resId"')}
           GROUP BY 1
         ),
         closed AS (
           SELECT date_trunc('week', c."workCompletedDate") AS w, COUNT(*)::int AS n,
                  COUNT(*) FILTER (WHERE c."recheckResult" = 'ok')::int AS confirmed
           FROM "CheckHistories" c
           WHERE c."workCompletedDate" >= now() - interval '12 weeks' ${rf('c."resId"')}
           GROUP BY 1
         )
         SELECT to_char(weeks.w, 'YYYY-MM-DD') AS week,
                COALESCE(found.n, 0) AS found, COALESCE(closed.n, 0) AS closed, COALESCE(closed.confirmed, 0) AS confirmed
         FROM weeks LEFT JOIN found ON found.w = weeks.w LEFT JOIN closed ON closed.w = weeks.w
         ORDER BY weeks.w`),
      q(`SELECT COUNT(*) FILTER (WHERE status = 'open')::int AS open,
                COUNT(*) FILTER (WHERE status = 'closed' AND "closedAt" >= now() - interval '30 days')::int AS closed30
         FROM "Escalations" WHERE TRUE ${rf('"resId"')}`),
      getNotificationCountsCached(req.user),
    ]);

    const totals = byRes.reduce((t, r) => {
      for (const k of ['tp', 'vl', 'pu', 'ok', 'error', 'pending', 'unchecked']) t[k] = (t[k] || 0) + r[k];
      return t;
    }, {});
    res.json({
      resId, totals, byRes, weekly,
      waiting: {
        techPending: counts.tech_pending || 0,
        askuePending: counts.askue_pending || 0,
        problemVl: counts.problem_vl || 0,
        overload: counts.powerOverload || 0,
        techblock: escalationsByWeek[0]?.open || 0,
        techblockClosed30: escalationsByWeek[0]?.closed30 || 0,
      },
    });
  } catch (error) {
    console.error('dashboard error:', error);
    res.status(500).json({ error: error.message });
  }
});

// =====================================================
// КАРТОЧКА ПУ (единое окно вместо россыпи модалок)
// =====================================================
// Всё о ПУ и его ТП одним запросом: состояние ПУ, открытое уведомление, история
// загрузок и проверок; по ТП — секции (Sном, пик, данные АСКУЭ), все линии с ПУ и
// статусами, активные перегрузы, проблемные ВЛ и вопросы техблоку. Область — по
// должности: «свой РЭС» видит только свой.
app.get('/api/pu-card/:puNumber', authenticateToken, async (req, res) => {
  try {
    const puNumber = String(req.params.puNumber || '').trim();
    if (!puNumber) return res.status(400).json({ error: 'Не указан номер ПУ' });

    const line = await NetworkStructure.findOne({
      where: { [Op.or]: [{ startPu: puNumber }, { middlePu: puNumber }, { endPu: puNumber }] },
      include: [{ model: ResUnit, attributes: ['id', 'name'] }],
    });
    if (!line) return res.status(404).json({ error: `ПУ ${puNumber} нет в структуре сети` });
    if (!req.user.allRes && line.resId !== req.user.resId) return res.status(403).json({ error: 'ПУ другого РЭС' });
    const position = line.startPu === puNumber ? 'start' : line.middlePu === puNumber ? 'middle' : 'end';

    const tpLines = await NetworkStructure.findAll({
      where: { resId: line.resId, tpName: line.tpName },
      attributes: ['id', 'vlName', 'lineType', 'sectionId', 'startPu', 'middlePu', 'endPu'],
      order: [['vlName', 'ASC']],
    });
    const lineIds = tpLines.map(l => l.id);
    const allPus = tpLines.flatMap(l => [l.startPu, l.middlePu, l.endPu]).filter(Boolean);

    const [statuses, sections, uploads, checks, notif, problems, escalations] = await Promise.all([
      PuStatus.findAll({ where: { puNumber: { [Op.in]: allPus } }, attributes: ['puNumber', 'position', 'status', 'errorDetails', 'lastCheck'], raw: true }),
      TpSection.findAll({ where: { resId: line.resId, tpName: line.tpName }, order: [['sectionNumber', 'ASC']] }),
      PuUploadHistory.findAll({ where: { puNumber }, include: [{ model: User, attributes: ['fio'] }], order: [['uploadedAt', 'DESC']], limit: 30 }),
      CheckHistory.findAll({ where: { puNumber }, order: [['createdAt', 'DESC']], limit: 30 }),
      Notification.findOne({
        where: { type: { [Op.in]: ['error', 'pending_askue'] }, resId: line.resId, message: { [Op.like]: `%"puNumber":"${puNumber.replace(/["%_\\]/g, '')}"%` } },
        order: [['createdAt', 'DESC']],
      }),
      ProblemVL.findAll({ where: { networkStructureId: { [Op.in]: lineIds }, status: 'active' }, raw: true }),
      Escalation.findAll({ where: { networkStructureId: { [Op.in]: lineIds } }, order: [['createdAt', 'DESC']], limit: 10, raw: true }),
    ]);
    const sectionIds = sections.map(s => s.id);
    const overloads = sectionIds.length
      ? await OverloadCase.findAll({ where: { sectionId: { [Op.in]: sectionIds }, stage: { [Op.ne]: 'completed' } }, raw: true })
      : [];

    const statusOf = (pu) => statuses.find(s => s.puNumber === pu) || null;
    const summary = (txt) => { try { const p = JSON.parse(txt); return p.summary || txt; } catch { return txt; } };
    const me = statusOf(puNumber);
    let open = null;
    if (notif) { try { open = { id: notif.id, type: notif.type, createdAt: notif.createdAt, ...JSON.parse(notif.message) }; } catch { open = null; } }

    res.json({
      pu: {
        puNumber, position,
        status: me?.status || 'not_checked',
        errorDetails: me?.errorDetails ? summary(me.errorDetails) : null,
        lastCheck: me?.lastCheck || null,
        openNotification: open,
      },
      line: { id: line.id, vlName: line.vlName, lineType: line.lineType, sectionId: line.sectionId },
      tp: {
        name: line.tpName, resId: line.resId, resName: line.ResUnit?.name || '',
        sections: sections.map(s => {
          const limitKw = s.tnKva ? Math.round(s.tnKva * (s.cosPhi || 0.9) * 10) / 10 : null;
          return {
            id: s.id, number: s.sectionNumber, tnKva: s.tnKva, cosPhi: s.cosPhi, limitKw,
            techPuNumber: s.techPuNumber, overloadStatus: s.overloadStatus,
            lastPeakKw: s.lastPeakKw, lastPeakAt: s.lastPeakAt,
            loadPct: limitKw && s.lastPeakKw ? Math.round((s.lastPeakKw / limitKw) * 100) : null,
            askue: s.askueData || null,
            overload: overloads.find(o => o.sectionId === s.id) || null,
          };
        }),
        lines: tpLines.map(l => ({
          id: l.id, vlName: l.vlName, lineType: l.lineType, sectionId: l.sectionId,
          problem: problems.some(p => p.networkStructureId === l.id),
          escalated: escalations.some(e => e.networkStructureId === l.id && e.status === 'open'),
          pus: ['start', 'middle', 'end'].map(pos => {
            const n = l[`${pos}Pu`];
            return { position: pos, puNumber: n || null, status: n ? (statusOf(n)?.status || 'not_checked') : 'empty' };
          }),
        })),
        escalations: escalations.map(e => ({
          id: e.id, vlName: e.vlName, puNumber: e.puNumber, status: e.status, description: e.description,
          decision: e.decision, createdAt: e.createdAt, closedAt: e.closedAt,
        })),
      },
      uploads: uploads.map(u => ({
        id: u.id, at: u.uploadedAt || u.createdAt, by: u.User?.fio || null, fileName: u.fileName, fileType: u.fileType,
        status: u.uploadStatus, hasErrors: u.hasErrors, errorSummary: u.errorSummary,
        period: u.periodStart ? [u.periodStart, u.periodEnd] : null,
      })),
      checks: checks.map(c => ({
        id: c.id, initialCheckDate: c.initialCheckDate, initialError: c.initialError, resComment: c.resComment,
        workCompletedDate: c.workCompletedDate, recheckDate: c.recheckDate, recheckResult: c.recheckResult,
        status: c.status, failureCount: c.failureCount, attachments: c.attachments || [],
      })),
    });
  } catch (error) {
    console.error('pu-card error:', error);
    res.status(500).json({ error: error.message });
  }
});

// 9. ОЧИСТКА ВСЕХ ДАННЫХ
app.delete('/api/network/clear-all',
  authenticateToken,
  requirePerm('structure_edit'),
  async (req, res) => {
    const { password, beforeDate } = req.body;  // ДОБАВИЛИ beforeDate

    if (password !== DELETE_PASSWORD) {
      return res.status(403).json({ error: 'Неверный пароль' });
    }

    const transaction = await sequelize.transaction();

    try {
      console.log('Starting data cleanup...');
      console.log('Before date:', beforeDate);
      
      // Формируем условие для удаления
      let whereClause = {};
      if (beforeDate) {
        whereClause.createdAt = { [Op.lt]: new Date(beforeDate) };
      }
      
      // 1. NotificationRead
      const notificationReadsDeleted = await NotificationRead.destroy({ 
        where: whereClause, 
        transaction 
      });
      console.log(`Deleted ${notificationReadsDeleted} notification read records`);
      
      // 2. PuUploadHistory (используем uploadedAt)
      let uploadWhereClause = {};
      if (beforeDate) {
        uploadWhereClause.uploadedAt = { [Op.lt]: new Date(beforeDate) };
      }
      
      const puUploadHistoryDeleted = await PuUploadHistory.destroy({ 
        where: uploadWhereClause, 
        transaction 
      });
      console.log(`Deleted ${puUploadHistoryDeleted} PU upload history records`);
      
      // 3. ProblemVL - удаляем только неактивные или старые
      let problemWhereClause = {};
      if (beforeDate) {
        problemWhereClause.createdAt = { [Op.lt]: new Date(beforeDate) };
        problemWhereClause.status = { [Op.ne]: 'active' }; // не трогаем активные
      }
      
      const problemVLDeleted = await ProblemVL.destroy({ 
        where: problemWhereClause, 
        transaction 
      });
      console.log(`Deleted ${problemVLDeleted} problem VL records`);
      
      // 4. CheckHistory  
      const checkHistoryDeleted = await CheckHistory.destroy({ 
        where: whereClause, 
        transaction 
      });
      console.log(`Deleted ${checkHistoryDeleted} check history records`);
      
      // 5. История загрузок
      const uploadsDeleted = await UploadHistory.destroy({ 
        where: whereClause, 
        transaction 
      });
      console.log(`Deleted ${uploadsDeleted} upload records`);
      
      // 6. Уведомления
      const notificationsDeleted = await Notification.destroy({ 
        where: whereClause, 
        transaction 
      });
      console.log(`Deleted ${notificationsDeleted} notifications`);
      
      // 7. Статусы ПУ - НЕ УДАЛЯЕМ, только сбрасываем старые
      if (beforeDate) {
        await PuStatus.update(
          { 
            status: 'not_checked',
            errorDetails: null,
            lastCheck: null
          },
          { 
            where: {
              lastCheck: { [Op.lt]: new Date(beforeDate) }
            },
            transaction 
          }
        );
      } else {
        // Если нет даты - сбрасываем все
        await PuStatus.update(
          { 
            status: 'not_checked',
            errorDetails: null,
            lastCheck: null
          },
          { 
            where: {},
            transaction 
          }
        );
      }
      
      // 8. Теперь можем удалить структуру сети
      const structuresDeleted = await NetworkStructure.destroy({ 
        where: {}, 
        transaction 
      });
      console.log(`Deleted ${structuresDeleted} network structures`);
      
      await transaction.commit();
      
      res.json({
        success: true,
        message: beforeDate 
          ? `Данные до ${new Date(beforeDate).toLocaleDateString('ru-RU')} успешно удалены`
          : 'История и статусы успешно очищены (структура сохранена)',
        deleted: {
          notificationReads: notificationReadsDeleted,
          puUploadHistory: puUploadHistoryDeleted,
          problemVL: problemVLDeleted,
          checkHistory: checkHistoryDeleted,
          uploads: uploadsDeleted,
          notifications: notificationsDeleted,
          // structures: 0  // Структура не удалялась
        }
      });
      
    } catch (error) {
      await transaction.rollback();
      console.error('Clear data error:', error);
      res.status(500).json({ 
        error: 'Ошибка при удалении данных: ' + error.message 
      });
    }
});

// 10. УДАЛЕНИЕ ВЫБРАННЫХ СТРУКТУР
app.post('/api/network/delete-selected',
  authenticateToken,
  requirePerm('structure_edit'),
  async (req, res) => {
    const { ids, password } = req.body;

    // Проверка пароля
    if (password !== DELETE_PASSWORD) {
      return res.status(403).json({ error: 'Неверный пароль' });
    }

    if (!ids || !Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: 'Не выбраны записи для удаления' });
    }

    const transaction = await sequelize.transaction();

    try {
      // Право не расширяет географию: не-админ удаляет только ВЛ своего РЭС.
      if (!req.user.allRes) {
        const foreign = await NetworkStructure.count({
          where: { id: { [Op.in]: ids }, resId: { [Op.ne]: req.user.resId } }, transaction
        });
        if (foreign > 0) {
          await transaction.rollback();
          return res.status(403).json({ error: 'Удалять можно только ВЛ своего РЭС' });
        }
      }

      console.log(`Deleting network structures: ${ids.join(', ')}`);
      
      // ВАЖНО: правильный порядок удаления!
      
      // 1. Сначала удаляем CheckHistory (новое!)
      const checkHistoryDeleted = await CheckHistory.destroy({
        where: {
          networkStructureId: { [Op.in]: ids }
        },
        transaction
      });
      
      // 2. Удаляем уведомления
      const notificationsDeleted = await Notification.destroy({
        where: {
          networkStructureId: { [Op.in]: ids }
        },
        transaction
      });
      
      // 3. Удаляем статусы ПУ
      const puStatusesDeleted = await PuStatus.destroy({
        where: {
          networkStructureId: { [Op.in]: ids }
        },
        transaction
      });
      
      // 4. Теперь можем удалить сами структуры
      const structuresDeleted = await NetworkStructure.destroy({
        where: {
          id: { [Op.in]: ids }
        },
        transaction
      });
      
      await transaction.commit();
      
      res.json({
        success: true,
        message: `Удалено ${structuresDeleted} записей`,
        deleted: {
          structures: structuresDeleted,
          checkHistory: checkHistoryDeleted,
          notifications: notificationsDeleted,
          puStatuses: puStatusesDeleted
        }
      });
      
    } catch (error) {
      await transaction.rollback();
      console.error('Delete selected error:', error);
      res.status(500).json({ error: error.message });
    }
});

// 11. ДЕТАЛЬНЫЕ ОТЧЕТЫ
app.get('/api/reports/detailed', authenticateToken, async (req, res) => {
  try {
    const { type, dateFrom, dateTo } = req.query;
    
    let whereClause = {};
    if (dateFrom || dateTo) {
      whereClause.createdAt = {};
      if (dateFrom) whereClause.createdAt[Op.gte] = new Date(dateFrom);
      if (dateTo) {
        const endDate = new Date(dateTo);
        endDate.setHours(23, 59, 59, 999);
        whereClause.createdAt[Op.lte] = endDate;
      }
    }
    
    // Добавляем фильтр по РЭС для не-админов
    if (req.user.allRes && req.query.resId) {
      whereClause.resId = parseInt(req.query.resId);
    } else if (!req.user.allRes) {
      whereClause.resId = req.user.resId;
    }
    
    let reportData = [];
    
    switch (type) {
      case 'pending_work':
        // Ожидающие мероприятий
        const pendingWork = await Notification.findAll({
          where: {
            ...whereClause,
            type: 'error'
          },
          include: [
            { model: ResUnit },
            { model: NetworkStructure }
          ],
          order: [['createdAt', 'DESC']]
        });
        
        // ✅ Дедупликация: ПУ + набор фаз = 1 запись (последняя)
        const seenKeysWork = new Set();
        reportData = pendingWork
          .filter(n => {
            try {
              const data = JSON.parse(n.message);
              if (!data.puNumber) return true;
              const phaseKey = getPhaseSignature(n.message);
              const key = `${data.puNumber}_${phaseKey}`;
              if (seenKeysWork.has(key)) return false;
              seenKeysWork.add(key);
              return true;
            } catch { return true; }
          })
          .map(n => {
            const data = JSON.parse(n.message);
            return {
              resName: n.ResUnit?.name,
              tpName: data.tpName,
              vlName: data.vlName,
              position: data.position,
              puNumber: data.puNumber,
              errorDetails: data.errorDetails,
              errorDate: n.createdAt
            };
          });
        break;
        
      case 'pending_askue':
        // Ожидающие проверки АСКУЭ
        const pendingAskue = await Notification.findAll({
          where: {
            ...whereClause,
            type: 'pending_askue'
          },
          include: [
            { model: ResUnit },
            { model: NetworkStructure }
          ],
          order: [['createdAt', 'DESC']]
        });
        
        // ✅ Дедупликация: ПУ + набор фаз = 1 запись (последняя)
        const seenKeysAskue = new Set();
        reportData = pendingAskue
          .filter(n => {
            try {
              const data = JSON.parse(n.message);
              if (!data.puNumber) return true;
              const phaseKey = getPhaseSignature(n.message);
              const key = `${data.puNumber}_${phaseKey}`;
              if (seenKeysAskue.has(key)) return false;
              seenKeysAskue.add(key);
              return true;
            } catch { return true; }
          })
          .map(n => {
            const data = JSON.parse(n.message);
            return {
              resName: n.ResUnit?.name,
              tpName: data.tpName,
              vlName: data.vlName,
              position: data.position,
              puNumber: data.puNumber,
              errorDetails: data.errorDetails || 'Требуется перепроверка',
              errorDate: n.createdAt,
              resComment: data.completedComment,
              workCompletedDate: data.completedAt
            };
          });
        break;
        
      case 'completed':
        // Завершенные проверки
        const completed = await CheckHistory.findAll({
          where: {
            ...whereClause,
            status: 'completed'
          },
          include: [ResUnit]
        });
        
        reportData = completed.map(h => ({
          resName: h.ResUnit?.name,
          tpName: h.tpName,
          vlName: h.vlName,
          position: h.position,
          puNumber: h.puNumber,
          errorDetails: h.initialError,
          errorDate: h.initialCheckDate,
          resComment: h.resComment,
          workCompletedDate: h.workCompletedDate,
          recheckDate: h.recheckDate,
          recheckResult: h.recheckResult,
          attachments: h.attachments || []
        }));
        break;
      
      case 'vl_workload':
        // ВЛ в работе у РЭС — аналитика
        let resWorkloadCondition = {};
        if (req.user.allRes && req.query.resId) {
          resWorkloadCondition.id = parseInt(req.query.resId);
        } else if (!req.user.allRes) {
          resWorkloadCondition.id = req.user.resId;
        }
        
        const resListForWorkload = await ResUnit.findAll({
          where: resWorkloadCondition,
          order: [['name', 'ASC']]
        });
        
        // Получаем все error-уведомления
        let errorWhereClause = { type: 'error' };
        if (req.user.allRes && req.query.resId) {
          errorWhereClause.resId = parseInt(req.query.resId);
        } else if (!req.user.allRes) {
          errorWhereClause.resId = req.user.resId;
        }
        
        const allErrors = await Notification.findAll({
          where: errorWhereClause,
          attributes: ['id', 'message', 'networkStructureId', 'resId', 'createdAt'],
          order: [['createdAt', 'DESC']]
        });
        
        // Дедупликация по ПУ + фаза
        const seenKeysVl = new Set();
        const uniqueErrors = allErrors.filter(notif => {
          try {
            const data = JSON.parse(notif.message);
            if (!data.puNumber) return true;
            const phaseKey = getPhaseSignature(notif.message);
            const key = `${data.puNumber}_${phaseKey}`;
            if (seenKeysVl.has(key)) return false;
            seenKeysVl.add(key);
            return true;
          } catch { return true; }
        });
        
        // Группируем проблемные ВЛ (networkStructureId) по РЭС
        const problemVlByRes = {};
        uniqueErrors.forEach(notif => {
          if (!notif.resId || !notif.networkStructureId) return;
          if (!problemVlByRes[notif.resId]) {
            problemVlByRes[notif.resId] = new Set();
          }
          problemVlByRes[notif.resId].add(notif.networkStructureId);
        });
        
        // PERF: количество ВЛ по всем РЭС — ОДНИМ group-by запросом, в Map по resId
        // (было — NetworkStructure.count на каждый РЭС в цикле).
        const vlCountRows = await NetworkStructure.findAll({
          attributes: ['resId', [sequelize.fn('COUNT', sequelize.col('id')), 'cnt']],
          where: { resId: resListForWorkload.map(r => r.id) },
          group: ['resId'],
          raw: true
        });
        const vlCountByRes = new Map(vlCountRows.map(r => [r.resId, parseInt(r.cnt, 10)]));

        // Собираем аналитику
        let totalAllVl = 0;
        let totalProblemVl = 0;

        for (const resUnit of resListForWorkload) {
          const totalVl = vlCountByRes.get(resUnit.id) || 0;

          const problemVlCount = problemVlByRes[resUnit.id]
            ? problemVlByRes[resUnit.id].size 
            : 0;
          
          totalAllVl += totalVl;
          totalProblemVl += problemVlCount;
          
          reportData.push({
            resName: resUnit.name,
            totalVl,
            problemVl: problemVlCount,
            okVl: totalVl - problemVlCount,
            problemPercent: totalVl > 0 
              ? Math.round((problemVlCount / totalVl) * 100) 
              : 0
          });
        }
        
        // Добавляем итоговую строку
        reportData.push({
          resName: 'ИТОГО',
          totalVl: totalAllVl,
          problemVl: totalProblemVl,
          okVl: totalAllVl - totalProblemVl,
          problemPercent: totalAllVl > 0 
            ? Math.round((totalProblemVl / totalAllVl) * 100) 
            : 0,
          isTotal: true
        });
        break;
    }
    
    res.json(reportData);
  } catch (error) {
    console.error('Detailed reports error:', error);
    res.status(500).json({ error: error.message });
  }
});

// 12. УДАЛЕНИЕ УВЕДОМЛЕНИЙ
app.delete('/api/notifications/:id',
  authenticateToken,
  requirePerm('notifications_delete'),
  async (req, res) => {
    try {
      const { password } = req.body;
      
      // Проверка пароля через переменную окружения
      if (password !== DELETE_PASSWORD) {
        return res.status(403).json({ error: 'Неверный пароль' });
      }
      
      const notification = await Notification.findByPk(req.params.id);
      if (!notification) {
        return res.status(404).json({ error: 'Уведомление не найдено' });
      }
      
      await notification.destroy();
      
      res.json({ 
        success: true, 
        message: 'Уведомление удалено' 
      });
      
    } catch (error) {
      console.error('Delete notification error:', error);
      res.status(500).json({ error: error.message });
    }
});
// API для массового удаления уведомлений
app.post('/api/notifications/delete-bulk',
  authenticateToken,
  requirePerm('notifications_delete'),
  async (req, res) => {
    const { ids, password, deleteDocuments } = req.body; // ✅ ДОБАВИЛИ deleteDocuments

    if (password !== DELETE_PASSWORD) {
      return res.status(403).json({ error: 'Неверный пароль' });
    }

    if (!ids || !Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: 'Не выбраны записи для удаления' });
    }

    console.log('=== BULK DELETE NOTIFICATIONS ===');
    console.log('Notification IDs:', ids);
    console.log('Delete documents?', deleteDocuments);

    const transaction = await sequelize.transaction();
    const fileIds = [];           // public_id для удаления ПОСЛЕ commit
    let deletedCount = 0;
    let deletedDocumentsCount = 0;

    try {
      // ✅ НОВОЕ: Если нужно удалить связанные документы
      if (deleteDocuments) {
        // Находим все уведомления
        const notifications = await Notification.findAll({
          where: { id: { [Op.in]: ids } },
          transaction
        });

        console.log(`Found ${notifications.length} notifications to process`);

        // Собираем уникальные номера ПУ из уведомлений
        const puNumbers = new Set();
        notifications.forEach(notif => {
          try {
            const data = JSON.parse(notif.message);
            if (data.puNumber) {
              puNumbers.add(data.puNumber);
            }
          } catch (e) {
            console.error('Error parsing notification message:', e);
          }
        });

        console.log('PU numbers to check:', Array.from(puNumbers));

        // Находим все связанные документы в CheckHistory
        if (puNumbers.size > 0) {
          const relatedDocs = await CheckHistory.findAll({
            where: {
              puNumber: { [Op.in]: Array.from(puNumbers) }
            },
            transaction
          });

          console.log(`Found ${relatedDocs.length} related documents`);

          // Собираем public_id (файлы удалим ПОСЛЕ commit — не держим коннект БД)
          for (const doc of relatedDocs) {
            if (Array.isArray(doc.attachments)) {
              for (const file of doc.attachments) {
                if (file && file.public_id) fileIds.push(file.public_id);
              }
            }
          }

          // Удаляем записи из CheckHistory
          deletedDocumentsCount = await CheckHistory.destroy({
            where: {
              puNumber: { [Op.in]: Array.from(puNumbers) }
            },
            transaction
          });

          console.log(`✅ Deleted ${deletedDocumentsCount} documents`);
        }
      }

      // Удаляем уведомления
      deletedCount = await Notification.destroy({
        where: { id: { [Op.in]: ids } },
        transaction
      });

      await transaction.commit();
    } catch (error) {
      await transaction.rollback();
      console.error('Bulk delete notifications error:', error);
      return res.status(500).json({ error: error.message });
    }

    // После commit — удаляем файлы (purgeCloudinary батчит и различает raw/image по .pdf).
    // Ошибки Cloudinary НЕ откатывают БД — только счётчик.
    const cl = await purgeCloudinary(fileIds);

    console.log('=== DELETE COMPLETE ===');
    console.log(`Notifications: ${deletedCount}`);
    console.log(`Documents: ${deletedDocumentsCount}`);
    console.log(`Files: ${cl.deleted}`);

    res.json({
      success: true,
      message: deleteDocuments
        ? `Удалено:\n• Уведомлений: ${deletedCount}\n• Документов: ${deletedDocumentsCount}\n• Файлов: ${cl.deleted}`
        : `Удалено уведомлений: ${deletedCount}`,
      deletedCount,
      deletedDocumentsCount,
      deletedFilesCount: cl.deleted,
      cloudinaryErrors: cl.errors
    });
});

// роут ДЛЯ ОТЧЕТОВ эксель
app.get('/api/reports/export-history', 
  authenticateToken, 
  requirePerm('history_view'), 
  async (req, res) => {
    try {
      const history = await CheckHistory.findAll({
        include: [ResUnit, NetworkStructure],
        order: [['createdAt', 'DESC']]
      });
      
      const data = history.map(h => ({
        'ID': h.id,
        'РЭС': h.ResUnit?.name,
        'ТП': h.tpName,
        'ВЛ': h.vlName,
        'ПУ': h.puNumber,
        'Позиция': h.position === 'start' ? 'Начало' : h.position === 'middle' ? 'Середина' : 'Конец',
        'Первоначальная ошибка': h.initialError,
        'Дата обнаружения': h.initialCheckDate,
        'Комментарий РЭС': h.resComment || '-',
        'Дата выполнения работ': h.workCompletedDate || '-',
        'Дата перепроверки': h.recheckDate || '-',
        'Результат': h.recheckResult === 'ok' ? 'Исправлено' : h.recheckResult === 'error' ? 'Не исправлено' : 'Ожидает',
        'Статус': h.status === 'completed' ? 'Завершено' : h.status === 'awaiting_recheck' ? 'Ожидает перепроверки' : 'Ожидает работ'
      }));
      
      res.json({
        success: true,
        data: data,
        count: data.length
      });
      
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
});

// =====================================================
// УПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯМИ
// =====================================================

// 13. ПОЛУЧЕНИЕ СПИСКА ПОЛЬЗОВАТЕЛЕЙ
app.get('/api/users/list', authenticateToken, requirePerm(['users_manage', 'settings_view']), async (req, res) => {
  try {
    const users = await User.findAll({
      attributes: ['id', 'fio', 'login', 'role', 'resId', 'email', 'isSuper', 'createdAt'],
      include: [ResUnit],
      order: [['createdAt', 'DESC']]
    });
    
    res.json(users);
  } catch (error) {
    console.error('Get users error:', error);
    res.status(500).json({ error: error.message });
  }
});

// 14. СОЗДАНИЕ ПОЛЬЗОВАТЕЛЯ
app.post('/api/users/create', authenticateToken, requirePerm('users_manage'), async (req, res) => {
  try {
    const { fio, login, password, email, role, resId } = req.body;

    // Валидация
    if (!fio || !login || !password || !email || !role) {
      return res.status(400).json({ error: 'Все поля обязательны' });
    }

    // Только суперадмин может создавать админов (иначе выдача прав обходится
    // созданием нового админа с нужными правами).
    if (role === 'admin') {
      const me = await getUserAccess(req.user.id);
      if (!me.isSuper) return res.status(403).json({ error: 'Создавать администраторов может только суперадмин' });
    }
    
    if (password.length < 6) {
      return res.status(400).json({ error: 'Пароль должен содержать минимум 6 символов' });
    }
    
    if (!ROLES[role]) {
      return res.status(400).json({ error: 'Выберите должность из списка' });
    }
    if (ROLES[role].scope === 'res' && !resId) {
      return res.status(400).json({ error: `Для должности «${ROLES[role].title}» выберите РЭС` });
    }
    
    // Проверка уникальности логина
    const existingUser = await User.findOne({ where: { login } });
    if (existingUser) {
      return res.status(400).json({ error: 'Пользователь с таким логином уже существует' });
    }
    
    // Создание пользователя
    const user = await User.create({
      fio,
      login,
      password,
      email,
      role,
      resId: ROLES[role].scope === 'all' ? null : resId
    });
    
    const createdUser = await User.findByPk(user.id, {
      attributes: ['id', 'fio', 'login', 'role', 'resId', 'email'],
      include: [ResUnit]
    });
    
    res.json({
      success: true,
      message: 'Пользователь создан успешно',
      user: createdUser
    });
    
  } catch (error) {
    console.error('Create user error:', error);
    res.status(500).json({ error: error.message });
  }
});

// 15. ОБНОВЛЕНИЕ ПОЛЬЗОВАТЕЛЯ
app.put('/api/users/:id', authenticateToken, requirePerm('users_manage'), async (req, res) => {
  try {
    const userId = req.params.id;
    const { fio, login, password, email, role, resId } = req.body;

    const user = await User.findByPk(userId);
    if (!user) {
      return res.status(404).json({ error: 'Пользователь не найден' });
    }

    const me = await getUserAccess(req.user.id);
    // Учётки администраторов (в т.ч. суперадмина) может править ТОЛЬКО суперадмин:
    // иначе не-супер с users_manage сбросил бы пароль/поменял поля чужого админа
    // = эскалация. Своих res_responsible/uploader он редактирует свободно.
    if (!me.isSuper && user.role === 'admin') {
      return res.status(403).json({ error: 'Учётки администраторов может изменять только суперадмин' });
    }
    // Суперадмина (login='admin') нельзя понижать/переименовывать НИКОМУ, включая его
    // самого (защита от выстрела в ногу). Смена его пароля/ФИО/email — можно.
    if (user.isSuper) {
      if (role && role !== 'admin') return res.status(403).json({ error: 'Нельзя изменить роль суперадмина' });
      if (login && login !== user.login) return res.status(403).json({ error: 'Нельзя переименовать логин суперадмина' });
    }
    // Назначать роль admin (создавать нового админа сменой роли) может только суперадмин.
    if (!me.isSuper && role === 'admin') {
      return res.status(403).json({ error: 'Назначать роль администратора может только суперадмин' });
    }

    // Обновляем только переданные поля
    const updateData = {};
    if (fio) updateData.fio = fio;
    // Логин можно менять, но он уникален — проверяем занятость другим пользователем.
    if (login && login !== user.login) {
      const taken = await User.findOne({ where: { login, id: { [Op.ne]: userId } } });
      if (taken) {
        return res.status(400).json({ error: 'Логин уже занят' });
      }
      updateData.login = login;
    }
    if (email) updateData.email = email;
    if (role && !ROLES[role]) {
      return res.status(400).json({ error: 'Выберите должность из списка' });
    }
    if (role) updateData.role = role;
    const nextRole = role || user.role;
    if (ROLES[nextRole]?.scope === 'all') {
      updateData.resId = null;
    } else {
      const nextRes = resId !== undefined ? resId : user.resId;
      if (!nextRes) return res.status(400).json({ error: `Для должности «${ROLES[nextRole]?.title || nextRole}» выберите РЭС` });
      updateData.resId = nextRes;
    }
    
    // Если передан новый пароль
    if (password && password.length > 0) {
      if (password.length < 6) {
        return res.status(400).json({ error: 'Пароль должен содержать минимум 6 символов' });
      }
      updateData.password = password;
    }
    
    await user.update(updateData);
    invalidateAccess(userId);   // роль/учётка могли измениться — сбросить кеш доступа

    const updatedUser = await User.findByPk(userId, {
      attributes: ['id', 'fio', 'login', 'role', 'resId', 'email'],
      include: [ResUnit]
    });
    
    res.json({
      success: true,
      message: 'Пользователь обновлен успешно',
      user: updatedUser
    });
    
  } catch (error) {
    console.error('Update user error:', error);
    res.status(500).json({ error: error.message });
  }
});

// 16. УДАЛЕНИЕ ПОЛЬЗОВАТЕЛЯ
app.delete('/api/users/:id', authenticateToken, requirePerm('users_manage'), async (req, res) => {
  try {
    const userId = req.params.id;
    const { password } = req.body;

    // Проверка пароля
    if (password !== DELETE_PASSWORD) {
      return res.status(403).json({ error: 'Неверный пароль' });
    }

    // Нельзя удалить себя
    if (userId == req.user.id) {
      return res.status(400).json({ error: 'Нельзя удалить свой аккаунт' });
    }
    
    const user = await User.findByPk(userId);
    if (!user) {
      return res.status(404).json({ error: 'Пользователь не найден' });
    }

    // Суперадмина (login='admin') удалить нельзя никому.
    if (user.isSuper) {
      return res.status(403).json({ error: 'Нельзя удалить суперадмина' });
    }

    // Учётки администраторов удаляет только суперадмин (иначе не-супер с users_manage
    // мог бы вычистить остальных админов).
    const meDel = await getUserAccess(req.user.id);
    if (user.role === 'admin' && !meDel.isSuper) {
      return res.status(403).json({ error: 'Учётки администраторов может изменять только суперадмин' });
    }

    // Проверяем, не последний ли это админ
    if (user.role === 'admin') {
      const adminCount = await User.count({ where: { role: 'admin' } });
      if (adminCount <= 1) {
        return res.status(400).json({ error: 'Нельзя удалить последнего администратора' });
      }
    }

    // Чистим зависимые записи в одной транзакции, чтобы удаление не падало на
    // FK и при этом сохранялись уведомления (их отработают с других учёток).
    await sequelize.transaction(async (t) => {
      // Служебное: отметки «прочитано» этого пользователя — просто мусор.
      await NotificationRead.destroy({ where: { userId }, transaction: t });
      // Уведомления сохраняем, только обнуляем ссылки на пользователя:
      //  - fromUserId → NULL: уведомление остаётся, теряет лишь имя отправителя;
      //  - toUserId → NULL: личное уведомление становится общим по РЭС и его
      //    подхватят другие ответственные учётки.
      await Notification.update({ fromUserId: null }, { where: { fromUserId: userId }, transaction: t });
      await Notification.update({ toUserId: null }, { where: { toUserId: userId }, transaction: t });
      // Содержательное (аудит загрузок) сохраняем, обнуляя авторство.
      await UploadHistory.update({ userId: null }, { where: { userId }, transaction: t });
      await PuUploadHistory.update({ uploadedBy: null }, { where: { uploadedBy: userId }, transaction: t });
      await user.destroy({ transaction: t });
    });
    invalidateAccess(userId);

    res.json({
      success: true,
      message: 'Пользователь удален'
    });

  } catch (error) {
    console.error('Delete user error:', error);
    res.status(500).json({ error: error.message });
  }
});

// =====================================================
// ПРАВА ДОЛЖНОСТЕЙ (матрица «должность × право»)
// =====================================================
// Каталог должностей и прав — для форм (пользователи) и раздела «Права доступа».
app.get('/api/roles', authenticateToken, (req, res) => {
  res.json({
    roles: Object.entries(ROLES).map(([key, r]) => ({ key, title: r.title, scope: r.scope })),
    groups: PERMISSION_GROUPS,
    catalog: PERMISSION_CATALOG,
    lockedRole: 'admin',
  });
});

// Матрица: по каждой должности — список прав и сколько людей на ней.
app.get('/api/admin/permissions', authenticateToken, requirePerm('permissions_manage'), async (req, res) => {
  try {
    const matrix = await loadRoleMatrix(true);
    const counts = await User.findAll({
      attributes: ['role', [sequelize.fn('COUNT', sequelize.col('id')), 'n']],
      where: { isSuper: false }, group: ['role'], raw: true,
    });
    const people = Object.fromEntries(counts.map(c => [c.role, Number(c.n)]));
    res.json({
      groups: PERMISSION_GROUPS,
      catalog: PERMISSION_CATALOG,
      lockedRole: 'admin',
      roles: Object.entries(ROLES).map(([key, r]) => ({
        key, title: r.title, scope: r.scope, people: people[key] || 0,
        permissions: key === 'admin' ? Object.keys(PERMISSION_CATALOG) : [...(matrix.get(key) || [])],
      })),
    });
  } catch (error) {
    console.error('Get permissions error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Сохранить права должности. «Администратор» не правится — у него всегда всё
// (иначе можно отнять у себя же право настраивать права).
app.put('/api/admin/permissions/:role', authenticateToken, requirePerm('permissions_manage'), async (req, res) => {
  try {
    const role = req.params.role;
    if (!ROLES[role]) return res.status(404).json({ error: 'Нет такой должности' });
    if (role === 'admin') return res.status(400).json({ error: 'У «Администратора» всегда все права' });
    const incoming = Array.isArray(req.body?.permissions) ? req.body.permissions : [];
    const clean = [...new Set(incoming.filter(k => PERMISSION_CATALOG[k]))];
    await sequelize.transaction(async (t) => {
      await RolePermission.destroy({ where: { role }, transaction: t });
      if (clean.length) await RolePermission.bulkCreate(clean.map(permission => ({ role, permission })), { transaction: t });
      await SeededRole.findOrCreate({ where: { role }, transaction: t });
    });
    await loadRoleMatrix(true);   // применяется сразу, без перелогина
    invalidateNotifCountsCache();
    res.json({ success: true, permissions: clean });
  } catch (error) {
    console.error('Save permissions error:', error);
    res.status(500).json({ error: error.message });
  }
});

// =====================================================
// API ДЛЯ ПРОБЛЕМНЫХ ВЛ
// =====================================================

// Получение списка проблемных ВЛ
app.get('/api/problem-vl/list', 
  authenticateToken, 
  requirePerm('problem_vl_view'), 
  async (req, res) => {
    try {
      const where = { status: 'active' };
      if (!req.user.allRes) where.resId = req.user.resId;
      else if (req.query.resId) where.resId = parseInt(req.query.resId);
      // Линии у техблока не показываем проблемными, пока вопрос не завершён.
      const escalated = await escalatedLineIds();
      if (escalated.length) where.networkStructureId = { [Op.or]: [{ [Op.notIn]: escalated }, null] };
      const problemVLs = await ProblemVL.findAll({
        where,
        include: [ResUnit, NetworkStructure],
        order: [['failureCount', 'DESC'], ['lastErrorDate', 'DESC']]
      });
      res.json(problemVLs);
    } catch (error) {
      console.error('Get problem VLs error:', error);
      res.status(500).json({ error: error.message });
    }
});

// Проблемные ТП (перегруз): кейсы, где мероприятия не дали результата — cycles >=
// PROBLEM_TP_CYCLES и ещё не завершён. Выход из списка — когда кейс дошёл до completed.
app.get('/api/problem-vl/overload-tp',
  authenticateToken,
  requirePerm('problem_vl_view'),
  async (req, res) => {
    try {
      const where = { cycles: { [Op.gte]: PROBLEM_TP_CYCLES }, stage: { [Op.ne]: 'completed' } };
      if (!req.user.allRes) where.resId = req.user.resId;
      else if (req.query.resId) where.resId = parseInt(req.query.resId);
      const cases = await OverloadCase.findAll({
        where,
        include: [
          { model: TpSection, as: 'section', attributes: ['id', 'tpName', 'sectionNumber', 'tnKva', 'techPuNumber'] },
          { model: ResUnit, attributes: ['id', 'name'] },
          { model: User, as: 'askueUser', attributes: ['id', 'fio'] },
          { model: User, as: 'resUser', attributes: ['id', 'fio'] }
        ],
        order: [['cycles', 'DESC'], ['id', 'DESC']]
      });
      const list = cases.map(c => ({
        id: c.id, sectionId: c.sectionId, resId: c.resId,
        resName: c.ResUnit?.name || '',
        tpName: c.section?.tpName || '', sectionNumber: c.section?.sectionNumber ?? null,
        techPuNumber: c.section?.techPuNumber || null,
        cycles: c.cycles, stage: c.stage,
        peakKw: c.peakKw, peakAt: c.peakAt, limitKw: c.limitKw, ratio: c.ratio,
        tnKva: c.tnKva, cosPhi: c.cosPhi, period: c.period,
        recheckPeakKw: c.recheckPeakKw, recheckAt: c.recheckAt, recheckResult: c.recheckResult,
        askueComment: c.askueComment, askueCompletedAt: c.askueCompletedAt,
        askueUser: c.askueUser?.fio || null,
        resComment: c.resComment, resCompletedAt: c.resCompletedAt,
        resUser: c.resUser?.fio || null,
        attachments: c.attachments || []
      }));
      res.json(list);
    } catch (error) {
      console.error('Get problem TPs error:', error);
      res.status(500).json({ error: error.message });
    }
});

// Отклонение проблемы
app.put('/api/problem-vl/:id/dismiss', 
  authenticateToken, 
  requirePerm('problem_vl_manage'), 
  async (req, res) => {
    try {
      const { password } = req.body;
      
      if (password !== DELETE_PASSWORD) {
        return res.status(403).json({ error: 'Неверный пароль' });
      }
      
      await ProblemVL.update(
        { status: 'dismissed' },
        { where: { id: req.params.id } }
      );
      
      // Удаляем связанные уведомления
      await Notification.destroy({
        where: {
          type: 'problem_vl',
          message: {
            [Op.like]: `%"puNumber":"${req.params.id}"%`
          }
        }
      });
      
      res.json({ success: true, message: 'Проблема отклонена' });
    } catch (error) {
      console.error('Dismiss problem VL error:', error);
      res.status(500).json({ error: error.message });
    }
});
// отправка писем ENDPOINT :
app.post('/api/problem-vl/:id/send-email', 
  authenticateToken, 
  requirePerm('problem_vl_manage'), 
  async (req, res) => {
    try {
      const problemVL = await ProblemVL.findByPk(req.params.id, {
        include: [ResUnit]
      });
      
      if (!problemVL) {
        return res.status(404).json({ error: 'Проблема не найдена' });
      }
      
      // Находим ответственных за РЭС
      const responsibleUsers = (await usersWithPerm('res_work', { resId: problemVL.resId })).filter(u => !u.isSuper);
      
      if (responsibleUsers.length === 0) {
        return res.status(400).json({ error: 'Не найден ответственный для этого РЭС' });
      }
      
      // Здесь можно добавить отправку реального email через nodemailer
      // Пока просто создадим уведомление
      
      for (const user of responsibleUsers) {
        await Notification.create({
          fromUserId: req.user.id,
          toUserId: user.id,
          resId: problemVL.resId,
          type: 'info',
          message: `⚠️ Требуется объяснительная записка по проблемному ПУ №${problemVL.puNumber} (${problemVL.tpName} - ${problemVL.vlName}). Количество неудачных проверок: ${problemVL.failureCount}`,
          isRead: false
        });
      }
      
      res.json({ success: true, message: 'Уведомление отправлено' });
      
    } catch (error) {
      console.error('Send email error:', error);
      res.status(500).json({ error: error.message });
    }
});
// =====================================================
// Отчет по проблемным ВЛ
// =====================================================
app.get('/api/reports/problem-vl', 
  authenticateToken, 
  async (req, res) => {
    try {
      const { dateFrom, dateTo } = req.query;
      
      let whereClause = {};
      if (dateFrom || dateTo) {
        whereClause.lastErrorDate = {};
        if (dateFrom) whereClause.lastErrorDate[Op.gte] = new Date(dateFrom);
        if (dateTo) {
          const endDate = new Date(dateTo);
          endDate.setHours(23, 59, 59, 999);
          whereClause.lastErrorDate[Op.lte] = endDate;
        }
      }
      
      // Добавляем фильтр по РЭС для не-админов
      // Добавляем фильтр по РЭС
    if (req.user.allRes && req.query.resId) {
      whereClause.resId = parseInt(req.query.resId);
    } else if (!req.user.allRes) {
      whereClause.resId = req.user.resId;
    }
      
      const problemVLs = await ProblemVL.findAll({
        where: whereClause,
        include: [ResUnit],
        order: [['failureCount', 'DESC']]
      });
      
      const reportData = problemVLs.map(p => ({
        resName: p.ResUnit?.name,
        tpName: p.tpName,
        vlName: p.vlName,
        position: p.position === 'start' ? 'Начало' : p.position === 'middle' ? 'Середина' : 'Конец',
        puNumber: p.puNumber,
        failureCount: p.failureCount,
        firstReportDate: p.firstReportDate,
        lastErrorDate: p.lastErrorDate,
        lastErrorDetails: p.lastErrorDetails,
        status: p.status === 'active' ? 'Активная' : p.status === 'resolved' ? 'Решена' : 'Отклонена'
      }));
      
      res.json(reportData);
    } catch (error) {
      console.error('Problem VL report error:', error);
      res.status(500).json({ error: error.message });
    }
});

// Отчёт «Превышение Pном»: все секции с заданным Sном + данные последнего кейса.
app.get('/api/reports/overload', authenticateToken, async (req, res) => {
  try {
    const where = { tnKva: { [Op.ne]: null } };
    if (req.user.allRes && req.query.resId) {
      where.resId = parseInt(req.query.resId, 10);
    } else if (!req.user.allRes) {
      where.resId = req.user.resId;
    }

    const sections = await TpSection.findAll({
      where,
      include: [{ model: ResUnit, attributes: ['name'] }],
      order: [['tpName', 'ASC'], ['sectionNumber', 'ASC']]
    });

    const STAGE_RU = {
      askue_limit: 'Ограничение по АСКУЭ',
      res_work: 'Мероприятия РЭС',
      awaiting_recheck: 'Ожидает перепроверки',
      completed: 'Завершён'
    };

    // PERF: последний кейс на секцию — ОДНИМ запросом (DISTINCT ON), в Map по
    // sectionId (было — OverloadCase.findOne на каждую секцию внутри Promise.all).
    const latestCases = await sequelize.query(
      `SELECT DISTINCT ON ("sectionId") "sectionId", "stage", "askueCompletedAt", "resCompletedAt", "recheckResult", "cycles"
         FROM "OverloadCases"
        ORDER BY "sectionId", "id" DESC`,
      { type: Sequelize.QueryTypes.SELECT }
    );
    const caseBySection = new Map();
    for (const c of latestCases) caseBySection.set(c.sectionId, c);

    const rows = sections.map((s) => {
      const cosPhi = s.cosPhi != null ? s.cosPhi : 0.9;
      const limitKw = s.tnKva * cosPhi;
      const ratio = (s.lastPeakKw != null && limitKw) ? s.lastPeakKw / limitKw : null;
      const oc = caseBySection.get(s.id);
      return {
        resName: s.ResUnit?.name || '',
        tpName: s.tpName,
        sectionNumber: s.sectionNumber,
        tnKva: s.tnKva,
        cosPhi,
        limitKw: Math.round(limitKw * 10) / 10,
        lastPeakKw: s.lastPeakKw != null ? Math.round(s.lastPeakKw * 10) / 10 : null,
        lastPeakAt: s.lastPeakAt,
        ratioPct: ratio != null ? Math.round(ratio * 100) : null,
        caseStage: oc ? STAGE_RU[oc.stage] : '—',
        askueCompletedAt: oc?.askueCompletedAt || null,
        resCompletedAt: oc?.resCompletedAt || null,
        recheckResult: oc?.recheckResult === 'ok' ? 'устранён' : oc?.recheckResult === 'still_overload' ? 'повторный перегруз' : '—',
        cycles: oc?.cycles || 0
      };
    });

    res.json(rows);
  } catch (error) {
    console.error('Overload report error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Матрица помесячных пиков «секции × месяцы» (раздел «Анализ мощности»).
// from/to = YYYY-MM (default: последние 12 месяцев), resId опц. (res_responsible —
// принудительно свой). Всё двумя запросами (секции + пики за диапазон), без циклов.
app.get('/api/power/monthly-peaks', authenticateToken, async (req, res) => {
  try {
    const now = new Date();
    const fmt = (y, m) => `${y}-${String(m).padStart(2, '0')}`;
    const parseYM = (s) => { const m = String(s || '').match(/^(\d{4})-(\d{1,2})$/); return m ? { y: +m[1], mo: +m[2] } : null; };
    const toP = parseYM(req.query.to) || { y: now.getFullYear(), mo: now.getMonth() + 1 };
    let fromP = parseYM(req.query.from);
    if (!fromP) { let y = toP.y, mo = toP.mo - 11; while (mo <= 0) { mo += 12; y--; } fromP = { y, mo }; }

    // Список месяцев from..to включительно (страховка от кривого диапазона).
    const months = [];
    let cy = fromP.y, cm = fromP.mo;
    while ((cy < toP.y || (cy === toP.y && cm <= toP.mo)) && months.length <= 60) {
      months.push(fmt(cy, cm));
      cm++; if (cm > 12) { cm = 1; cy++; }
    }

    const where = {};
    if (!req.user.allRes) where.resId = req.user.resId;
    else if (req.query.resId) where.resId = parseInt(req.query.resId, 10);

    const sections = await TpSection.findAll({
      where,
      include: [{ model: ResUnit, attributes: ['name'] }],
      order: [['tpName', 'ASC'], ['sectionNumber', 'ASC']]
    });
    const sectionIds = sections.map(s => s.id);

    let peaks = [];
    if (sectionIds.length) {
      peaks = await sequelize.query(
        `SELECT "sectionId","year","month","peakKw","peakAt"
           FROM "SectionMonthlyPeaks"
          WHERE "sectionId" IN (:ids)
            AND ("year" * 100 + "month") BETWEEN :fromYM AND :toYM`,
        {
          replacements: { ids: sectionIds, fromYM: fromP.y * 100 + fromP.mo, toYM: toP.y * 100 + toP.mo },
          type: Sequelize.QueryTypes.SELECT
        }
      );
    }
    const bySection = new Map();
    for (const p of peaks) {
      if (!bySection.has(p.sectionId)) bySection.set(p.sectionId, {});
      bySection.get(p.sectionId)[fmt(p.year, p.month)] = p;
    }

    const rows = sections.map(s => {
      const cosPhi = s.cosPhi != null ? s.cosPhi : 0.9;
      const limitKw = s.tnKva != null ? s.tnKva * cosPhi : null;
      const pmap = bySection.get(s.id) || {};
      const peaksOut = {};
      for (const mo of months) {
        const p = pmap[mo];
        peaksOut[mo] = p
          ? { peakKw: Math.round(p.peakKw * 10) / 10, peakAt: p.peakAt, ratioPct: limitKw ? Math.round((p.peakKw / limitKw) * 100) : null }
          : null;
      }
      return {
        sectionId: s.id,
        resName: s.ResUnit?.name || '',
        tpName: s.tpName,
        sectionNumber: s.sectionNumber,
        tnKva: s.tnKva,
        cosPhi,
        limitKw: limitKw != null ? Math.round(limitKw * 10) / 10 : null,
        peaks: peaksOut
      };
    });

    res.json({ months, rows });
  } catch (error) {
    console.error('monthly-peaks error:', error);
    res.status(500).json({ error: error.message });
  }
});
// ── Кэш счётчиков /counts (TTL 15с) ──────────────────────────────────────────
// Роут поллится каждым клиентом раз в 30с + бейдж платформы. Кэшируем результат
// getNotificationCounts. Ключ — role_resId_userId: именно от них зависит выборка
// (res_responsible/uploader фильтруют по resId и toUserId=user.id, admin — глобально).
const notifCountsCache = new Map(); // key -> { ts, data }
const NOTIF_COUNTS_TTL_MS = 15000;
function invalidateNotifCountsCache() { notifCountsCache.clear(); }

// Любое изменение Notification/OverloadCase полностью чистит кэш. Хуки модели
// ловят ВСЕ пути create/update/destroy (и bulk) — надёжнее ручных вызовов по
// десяткам мест создания/удаления, ни одно не пропустим.
[Notification, OverloadCase, Escalation].forEach((M) => {
  M.addHook('afterCreate', invalidateNotifCountsCache);
  M.addHook('afterUpdate', invalidateNotifCountsCache);
  M.addHook('afterDestroy', invalidateNotifCountsCache);
  M.addHook('afterBulkCreate', invalidateNotifCountsCache);
  M.addHook('afterBulkUpdate', invalidateNotifCountsCache);
  M.addHook('afterBulkDestroy', invalidateNotifCountsCache);
});

async function getNotificationCountsCached(user) {
  const key = `${user.role}_${user.isSuper ? 's' : ''}_${user.resId ?? ''}_${user.id ?? ''}`;
  const now = Date.now();
  const hit = notifCountsCache.get(key);
  if (hit && (now - hit.ts) < NOTIF_COUNTS_TTL_MS) return hit.data;
  const data = await getNotificationCounts(user);
  notifCountsCache.set(key, { ts: now, data });
  return data;
}

// Счётчики уведомлений (меню-бейджи и бейдж платформы /api/platform/badge).
// Область — по должности (все РЭС или свой), что считать — по правам.
// user — { id, role, resId, isSuper, allRes, perms } (req.user или accessOf()).
// Возвращает { tech_pending, askue_pending, problem_vl, power_overload,
// powerOverload, techblock }.
async function getNotificationCounts(user) {
  const scope = user.allRes ? {} : { resId: user.resId ?? -1 };
  let whereClause = user.allRes ? {} : {
    resId: user.resId ?? -1,
    [Op.or]: [{ toUserId: null }, { toUserId: user.id }]
  };

  // Проблемные ВЛ (без линий, отданных техблоку) + проблемные ТП (перегруз).
  let problemVLCount = 0;
  let powerOverloadCount = 0;
  if (userCan(user, 'problem_vl_view')) {
    const escalated = await escalatedLineIds();
    problemVLCount = await ProblemVL.count({
      where: { ...scope, status: 'active', ...(escalated.length ? { networkStructureId: { [Op.or]: [{ [Op.notIn]: escalated }, null] } } : {}) }
    });
    problemVLCount += await OverloadCase.count({
      where: { ...scope, cycles: { [Op.gte]: PROBLEM_TP_CYCLES }, stage: { [Op.ne]: 'completed' } }
    });
  }
  if (userCan(user, 'overload_view')) {
    powerOverloadCount = await Notification.count({ where: { ...scope, type: 'power_overload' } });
  }

  // «Превышение Pном» — ТОЛЬКО рабочие кейсы, где человек может действовать.
  // awaiting_recheck в бейдж не входит (там ждём профиль, действия нет).
  const stages = [];
  if (userCan(user, 'overload_askue')) stages.push('askue_limit');
  if (userCan(user, 'overload_res')) stages.push('res_work');
  const powerOverloadCases = stages.length
    ? await OverloadCase.count({ where: { ...scope, stage: { [Op.in]: stages } } })
    : 0;

  const techblock = userCan(user, 'techblock_view')
    ? await Escalation.count({ where: { ...scope, status: 'open' } })
    : 0;

  const types = [];
  if (userCan(user, 'tech_pending_view') || userCan(user, 'res_work')) types.push('error');
  if (userCan(user, 'askue_pending_view') || userCan(user, 'upload_files')) types.push('pending_askue');
  if (!types.length) {
    return { tech_pending: 0, askue_pending: 0, problem_vl: problemVLCount, power_overload: powerOverloadCount, powerOverload: powerOverloadCases, techblock };
  }

  // ✅ PERF: считаем только те типы, которые реально участвуют в счётчиках
  whereClause.type = { [Op.in]: types };

  const allNotifications = await Notification.findAll({
    where: whereClause,
    attributes: ['id', 'type', 'message']
  });

  // Считаем ВСЕ ожидающие уведомления, а не только непрочитанные: бейдж должен
  // висеть, пока ПУ реально не отработан. Эти уведомления удаляются при загрузке
  // проверки по ПУ (см. Notification.destroy для pending_askue и error), поэтому
  // счётчик падает именно при отработке, а не при просмотре. Пометка «прочитано»
  // на счётчик больше не влияет — это нужно и для бейджа приложения на платформе.
  // ✅ Считаем уникальные ПУ+фаза по типам (не общее количество уведомлений!)
  const techKeys = new Set();
  const askueKeys = new Set();

  allNotifications.forEach(notif => {
    try {
      const data = JSON.parse(notif.message);
      const puNumber = data.puNumber;

      if (puNumber) {
        const phaseKey = getPhaseSignature(data); // уже распарсено — не парсим повторно
        const key = `${puNumber}_${phaseKey}`;

        if (notif.type === 'error') techKeys.add(key);
        else if (notif.type === 'pending_askue') askueKeys.add(key);
      } else {
        if (notif.type === 'error') techKeys.add(`id_${notif.id}`);
        else if (notif.type === 'pending_askue') askueKeys.add(`id_${notif.id}`);
      }
    } catch {
      if (notif.type === 'error') techKeys.add(`id_${notif.id}`);
      else if (notif.type === 'pending_askue') askueKeys.add(`id_${notif.id}`);
    }
  });

  return {
    tech_pending: techKeys.size,
    askue_pending: askueKeys.size,
    problem_vl: problemVLCount,
    power_overload: powerOverloadCount,
    powerOverload: powerOverloadCases,
    techblock
  };
}

// Линии, отданные техблоку (открытые вопросы): по ним не ведём «проблемность».
// Кеш 15 с — сбрасывается при создании/закрытии вопроса.
let escalatedCache = { ts: 0, ids: [] };
async function escalatedLineIds() {
  if (Date.now() - escalatedCache.ts < 15000) return escalatedCache.ids;
  const rows = await Escalation.findAll({ where: { status: 'open' }, attributes: ['networkStructureId'], raw: true });
  escalatedCache = { ts: Date.now(), ids: [...new Set(rows.map(r => r.networkStructureId).filter(Boolean))] };
  return escalatedCache.ids;
}
const invalidateEscalated = () => { escalatedCache.ts = 0; };
['afterCreate', 'afterUpdate', 'afterDestroy', 'afterBulkUpdate'].forEach(h => Escalation.addHook(h, invalidateEscalated));

app.get('/api/notifications/counts', authenticateToken, async (req, res) => {
  try {
    res.json(await getNotificationCountsCached(req.user));
  } catch (error) {
    console.error('Error counting notifications:', error);
    res.status(500).json({ error: error.message });
  }
});

// роут Отметить уведомления как прочитанные
app.put('/api/notifications/mark-read', authenticateToken, async (req, res) => {
  try {
    const { type } = req.body;
    let whereClause = {};
    
    // Определяем какие уведомления должны быть помечены
    if (req.user.allRes) {
      whereClause = type === 'all' ? {} : { type };
    } else {
      whereClause = {
        resId: req.user.resId,
        [Op.or]: [
          { toUserId: null },
          { toUserId: req.user.id }
        ]
      };
      if (type !== 'all') whereClause.type = type;
    }
    
    // Получаем все уведомления для пометки
    const notificationsToMark = await Notification.findAll({
      where: whereClause,
      attributes: ['id']
    });
    
    // Создаем записи о прочтении
    const readRecords = notificationsToMark.map(notif => ({
      notificationId: notif.id,
      userId: req.user.id,
      readAt: new Date()
    }));
    
    // Массовое создание с игнорированием дубликатов
    await NotificationRead.bulkCreate(readRecords, {
      ignoreDuplicates: true // игнорируем если уже прочитано
    });
    
    res.json({ success: true });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// API для массового удаления записей
app.post('/api/documents/delete-bulk',
  authenticateToken,
  requirePerm('files_manage'),
  async (req, res) => {
    const { ids, password } = req.body;

    if (password !== DELETE_PASSWORD) {
      return res.status(403).json({ error: 'Неверный пароль' });
    }

    if (!ids || !Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: 'Не выбраны записи для удаления' });
    }

    const transaction = await sequelize.transaction();
    const fileIds = [];
    let deletedCount = 0;

    try {
      // Получаем все записи для удаления, СОБИРАЕМ public_id (файлы удалим после commit)
      const records = await CheckHistory.findAll({
        where: { id: { [Op.in]: ids } },
        transaction
      });
      for (const record of records) {
        if (Array.isArray(record.attachments)) {
          for (const file of record.attachments) {
            if (file && file.public_id) fileIds.push(file.public_id);
          }
        }
      }

      // Удаляем записи из БД
      deletedCount = await CheckHistory.destroy({
        where: { id: { [Op.in]: ids } },
        transaction
      });

      await transaction.commit();
    } catch (error) {
      await transaction.rollback();
      console.error('Bulk delete error:', error);
      return res.status(500).json({ error: error.message });
    }

    // После commit — удаляем файлы (purgeCloudinary сам батчит и различает raw/image
    // по .pdf). Ошибки Cloudinary НЕ откатывают БД — только счётчик.
    const cl = await purgeCloudinary(fileIds);
    res.json({
      success: true,
      message: `Удалено записей: ${deletedCount}`,
      deletedCount,
      cloudinaryErrors: cl.errors
    });
});

// =====================================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ АНАЛИЗА
// =====================================================

// Запуск анализатора профиля мощности (Пирамида). Контракт как у остальных:
// spawn python, JSON в stdout. Возвращает распарсенный объект {success,results,warnings}.
function runProfileAnalyzer(filePath) {
  return new Promise((resolve) => {
    const scriptPath = path.join(process.cwd(), 'analyzers', 'profile_analyzer.py');
    if (!fs.existsSync(scriptPath)) {
      return resolve({ success: false, error: 'profile_analyzer.py не найден' });
    }
    let done = false, current = null;
    const finish = (v) => { if (done) return; done = true; clearTimeout(timer); resolve(v); };

    // Один таймаут на ОБЕ попытки (python3, затем python) — не утекает: убивает
    // текущий процесс. Зависание больше не выглядит как «молча ничего не нашлось».
    const timer = setTimeout(() => {
      try { if (current) current.kill('SIGKILL'); } catch (e) {}
      console.error('[PROFILE] анализатор не уложился в 120 с — SIGKILL');
      finish({ success: false, error: 'Анализатор не уложился в 120 с' });
    }, 120000);

    // spawn НЕ бросает синхронно (ENOENT приходит событием 'error'). Поэтому
    // пробуем 'python3', и в on('error') один раз падаем на 'python' с теми же
    // обработчиками. Если и он дал 'error' — Python недоступен.
    const run = (cmd, isFallback) => {
      let out = '', err = '';
      const child = spawn(cmd, [scriptPath, filePath]);
      current = child;
      child.stdout.on('data', d => { out += d.toString(); });
      child.stderr.on('data', d => { err += d.toString(); });
      child.on('error', () => {
        child.removeAllListeners('close');   // не дать 'close' первой попытки перебить fallback
        if (!isFallback) return run('python', true);
        finish({ success: false, error: 'Python недоступен' });
      });
      child.on('close', (code) => {
        if (done) return;
        if (err) console.error('[PROFILE] stderr:', err);
        if (code !== 0) return finish({ success: false, error: `Ошибка анализа (код ${code}): ${err}` });
        try { finish(JSON.parse(out)); }
        catch (e) { finish({ success: false, error: 'Не удалось разобрать результат анализатора' }); }
      });
    };
    run('python3', false);
  });
}

// Удаляет уведомления power_overload по секции + их отметки прочтения (без хвостов).
async function removeSectionOverloadNotifs(sectionId) {
  const notifs = await Notification.findAll({
    where: {
      type: 'power_overload',
      [Op.and]: [sequelize.where(sequelize.cast(sequelize.col('errorData'), 'text'), { [Op.like]: `%"sectionId":${sectionId}%` })]
    },
    attributes: ['id']
  });
  const ids = notifs.map(n => n.id);
  if (ids.length) {
    await NotificationRead.destroy({ where: { notificationId: ids } }).catch(() => {});
    await Notification.destroy({ where: { id: ids } }).catch(() => {});
  }
}

// 'dd.mm.yyyy HH:MM' → Date (для lastPeakAt). Некорректная строка → null.
function parsePeakAt(s) {
  if (!s) return null;
  const m = String(s).trim().match(/^(\d{2})\.(\d{2})\.(\d{4})\s+(\d{1,2}):(\d{2})$/);
  if (!m) return null;
  const [, dd, mm, yyyy, hh, mi] = m;
  const dt = new Date(Number(yyyy), Number(mm) - 1, Number(dd), Number(hh), Number(mi));
  return isNaN(dt.getTime()) ? null : dt;
}

async function analyzeFile(filePath, type, originalFileName = null, requiredPeriod = null, userId = null) {
  return new Promise((resolve, reject) => {

    
    // Вспомогательная функция для получения названия месяца
    function getMonthName(monthNum) {
      const months = ['', 'января', 'февраля', 'марта', 'апреля', 'мая', 'июня', 
                      'июля', 'августа', 'сентября', 'октября', 'ноября', 'декабря'];
      return months[monthNum] || '';
    }
    
    // Функция для извлечения периода из текста ошибки
    function extractPeriodFromError(errorText) {
      const monthMap = {
        'Янв': 1, 'Фев': 2, 'Мар': 3, 'Апр': 4, 'Май': 5, 'Июн': 6,
        'Июл': 7, 'Авг': 8, 'Сен': 9, 'Окт': 10, 'Ноя': 11, 'Дек': 12
      };
      
      const monthPattern = /(Янв|Фев|Мар|Апр|Май|Июн|Июл|Авг|Сен|Окт|Ноя|Дек)/g;
      const foundMonths = errorText.match(monthPattern);
      
      if (foundMonths && foundMonths.length > 0) {
        const firstMonth = monthMap[foundMonths[0]];
        const lastMonth = monthMap[foundMonths[foundMonths.length - 1]];
        const now = new Date();
        const currentYear = now.getFullYear();
        const currentMonth = now.getMonth() + 1; // 1..12
        // Данные — про прошлое: месяц, опережающий текущий более чем на 1
        // (напр. сейчас январь, а в журнале декабрь), относится к прошлому году.
        const yearFor = (m) => (m > currentMonth + 1) ? currentYear - 1 : currentYear;

        return {
          start: new Date(yearFor(firstMonth), firstMonth - 1, 1),
          end: new Date(yearFor(lastMonth), lastMonth - 1, 28)
        };
      }
      
      return null;
    }
    
    let scriptPath;
    const analyzersDir = path.join(process.cwd(), 'analyzers');
    
    switch(type) {
      case 'rim_single':
        scriptPath = path.join(analyzersDir, 'rim_converter_csv.py');
        break;
      case 'rim_mass':
        scriptPath = path.join(analyzersDir, 'rim_mass_analyzer.py');
        break;
      case 'nartis':
        scriptPath = path.join(analyzersDir, 'nartis_analyzer.py');
        break;
      case 'energomera':
        scriptPath = path.join(analyzersDir, 'energomera_analyzer.py');
        break;
      default:
        return resolve({
          processed: [],
          errors: ['Неизвестный тип анализатора']
        });
    }
    
    // Проверки существования директории и скрипта
    if (!fs.existsSync(analyzersDir)) {
      console.error('Analyzers directory not found:', analyzersDir);
      return resolve({
        processed: [],
        errors: [`Директория analyzers не найдена`]
      });
    }
    
    if (!fs.existsSync(scriptPath)) {
      console.error('Python script not found:', scriptPath);
      return resolve({
        processed: [],
        errors: [`Python скрипт не найден: ${scriptPath}`]
      });
    }
    
    // Запуск Python. spawn НЕ бросает синхронно (ENOENT приходит событием
    // 'error'), поэтому пробуем 'python3', а в on('error') один раз падаем на
    // 'python' с теми же обработчиками stdout/stderr/close.
    let output = '';
    let errorOutput = '';
    let done = false, current = null;

    // Оборачиваем resolve: единожды + гасим таймаут (в теле onClose много resolve).
    const _resolve = resolve;
    resolve = (v) => { if (done) return; done = true; clearTimeout(timer); _resolve(v); };

    // Один таймаут на ОБЕ попытки — покрывает оба процесса и не утекает.
    const timer = setTimeout(() => {
      try { if (current) current.kill('SIGKILL'); } catch (e) {}
      console.error('Python analyzer timed out (120s) — SIGKILL');
      resolve({ processed: [], errors: ['Анализатор не уложился в 120 с'] });
    }, 120000);

    const onClose = async (code) => {
      
      if (code !== 0) {
        return resolve({
          processed: [],
          errors: [`Ошибка анализа (код ${code}): ${errorOutput}`]
        });
      }
      
      try {
        // Парсим результат от Python
        const result = JSON.parse(output);
        
        if (result.success) {
          const processed = [];
          const errors = [];
          
          // Извлекаем номер ПУ из имени файла
          const fileName = originalFileName 
            ? path.basename(originalFileName, path.extname(originalFileName))
            : path.basename(filePath, path.extname(filePath));
          

          if (!fileName || fileName === 'undefined' || fileName === '') {
  console.error('ERROR: Invalid PU number');
  try {
    fs.unlinkSync(filePath);
  } catch (err) {
    console.error('Error deleting file:', err);
  }
  
  return resolve({
    processed: [],
    errors: [],
    success: false,
    message: 'Ошибка: не удалось определить номер ПУ из имени файла'
  });
}
          
          // НОВАЯ ПРОВЕРКА: История загрузок
          const recentUploads = await PuUploadHistory.findAll({
  where: {
    puNumber: fileName,
    uploadedAt: {
      [Op.gte]: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000) // последние 30 дней
    }
  },
  order: [['uploadedAt', 'DESC']]
});

// Извлекаем период из текущего файла
const currentPeriod = result.has_errors ? extractPeriodFromError(result.summary) : null;

// Проверяем дубликаты ТОЛЬКО если текущий файл с ошибкой
if (result.has_errors) {
  // Ищем такую же ошибку в истории
  const sameErrorUpload = recentUploads.find(upload => 
    upload.errorSummary === result.summary && upload.hasErrors
  );
  
  if (sameErrorUpload) {
    
    // Проверяем, была ли загрузка БЕЗ ошибок после этой ошибки
    const successfulUploadAfter = recentUploads.find(upload => 
      !upload.hasErrors && 
      new Date(upload.uploadedAt) > new Date(sameErrorUpload.uploadedAt)
    );
    
    if (successfulUploadAfter) {
      // Ошибка была исправлена, но теперь появилась снова
      // Разрешаем загрузку - это повторное появление ошибки
    } else {
      // Ошибка не была исправлена - это дубликат
      
      // Записываем попытку загрузки дубликата
      if (userId) {
        await PuUploadHistory.create({
          puNumber: fileName,
          uploadedBy: userId,
          fileName: originalFileName || 'unknown',
          fileType: type,
          periodStart: currentPeriod?.start,
          periodEnd: currentPeriod?.end,
          hasErrors: result.has_errors,
          errorSummary: result.summary,
          errorDetails: result.details,
          uploadStatus: 'duplicate'
        });
      }
      
      try {
        fs.unlinkSync(filePath);
      } catch (err) {
        console.error('Error deleting file:', err);
      }
      
      return resolve({
        processed: [{
          puNumber: fileName,
          status: 'duplicate_error',
          error: `❌ Данная ошибка уже была загружена ${new Date(sameErrorUpload.uploadedAt).toLocaleDateString('ru-RU')}! Проверьте статус обработки.`
        }],
        errors: []
      });
    }
  }
  
  // Дополнительная проверка: есть ли активные уведомления или записи в CheckHistory
  const activeCheckHistory = await CheckHistory.findOne({
    where: { 
      puNumber: fileName,
      initialError: result.summary,
      status: ['awaiting_work', 'awaiting_recheck']
    }
  });
  
  if (activeCheckHistory) {
    
    if (userId) {
      await PuUploadHistory.create({
        puNumber: fileName,
        uploadedBy: userId,
        fileName: originalFileName || 'unknown',
        fileType: type,
        periodStart: currentPeriod?.start,
        periodEnd: currentPeriod?.end,
        hasErrors: result.has_errors,
        errorSummary: result.summary,
        errorDetails: result.details,
        uploadStatus: 'duplicate'
      });
    }
    
    try {
      fs.unlinkSync(filePath);
    } catch (err) {
      console.error('Error deleting file:', err);
    }
    
    return resolve({
      processed: [{
        puNumber: fileName,
        status: 'duplicate_error',
        error: '❌ Данная ошибка уже находится в обработке!'
      }],
      errors: []
    });
  }
}
          
          // Ищем ПУ в структуре сети
          const networkStructure = await NetworkStructure.findOne({
            where: {
              [Op.or]: [
                { startPu: fileName },
                { endPu: fileName },
                { middlePu: fileName }
              ]
            },
            include: [ResUnit]
          });
          
          if (networkStructure) {
            
            // Определяем позицию
            let position = 'start';
            if (networkStructure.endPu === fileName) position = 'end';
            else if (networkStructure.middlePu === fileName) position = 'middle';
            
            
            // Проверяем последнюю запись в истории
            const lastCheckHistory = await CheckHistory.findOne({
              where: { 
                puNumber: fileName,
                [Op.or]: [
                  { status: 'awaiting_work' },
                  { status: 'awaiting_recheck' },
                  { status: 'completed' }
                ]
              },
              order: [['createdAt', 'DESC']]
            });
            
            // ПРОВЕРКА 1: Это перепроверка?
            if (lastCheckHistory && lastCheckHistory.status === 'awaiting_recheck') {
              
              // ПРОВЕРКА ПЕРИОДА при перепроверке
              if (result.has_errors) {
                // Сначала проверим дату
                const checkFromDate = lastCheckHistory.workCompletedDate;
                if (checkFromDate) {
                  const requiredDate = new Date(checkFromDate);
                  const requiredMonth = requiredDate.getMonth() + 1;
                  const requiredYear = requiredDate.getFullYear();
                  
                  const errorText = result.summary;
                  const monthMap = {
                    'Янв': 1, 'Фев': 2, 'Мар': 3, 'Апр': 4, 'Май': 5, 'Июн': 6,
                    'Июл': 7, 'Авг': 8, 'Сен': 9, 'Окт': 10, 'Ноя': 11, 'Дек': 12
                  };
                  
                  const monthPattern = /(Янв|Фев|Мар|Апр|Май|Июн|Июл|Авг|Сен|Окт|Ноя|Дек)/g;
                  const foundMonths = errorText.match(monthPattern);
                  
                  if (foundMonths && foundMonths.length > 0) {
                    const lastErrorMonth = foundMonths[foundMonths.length - 1];  // последний месяц в журнале
                    const lastErrorMonthNum = monthMap[lastErrorMonth];
  
                    // Журнал должен включать данные ПОСЛЕ месяца выполнения работ
                    if (lastErrorMonthNum < requiredMonth) {  // если последний месяц раньше требуемого
    
                      return resolve({
                        processed: [{
                          puNumber: fileName,
                          status: 'wrong_period',
                          error: `❌ Неверный период! Требуется журнал событий с ${requiredDate.toLocaleDateString('ru-RU')} по текущую дату. Журнал должен включать данные после ${getMonthName(requiredMonth)} ${requiredYear}!`
                        }],
                        errors: []
                      });
                    }
                  }
                }
              }
              
              // Удаляем уведомления АСКУЭ в любом случае
              await Notification.destroy({
                where: {
                  type: 'pending_askue',
                  message: {
                    [Op.like]: `%"puNumber":"${fileName}"%`
                  }
                }
              });
              
              // ОБРАБОТКА РЕЗУЛЬТАТА ПЕРЕПРОВЕРКИ
              if (!result.has_errors) {
                // УСПЕШНАЯ перепроверка
                
                // Обновляем историю
                await CheckHistory.update({
                  recheckDate: new Date(),
                  recheckResult: 'ok',
                  status: 'completed'
                }, {
                  where: {
                    puNumber: fileName,
                    status: 'awaiting_recheck'
                  }
                });
                
                // Обновляем статус ПУ
                await PuStatus.update({
                  status: 'checked_ok',
                  errorDetails: null,
                  lastCheck: new Date()
                }, {
                  where: { puNumber: fileName }
                });
                
                // Решаем проблемную ВЛ если была
                await ProblemVL.update(
                  { status: 'resolved' },
                  { where: { puNumber: fileName, status: 'active' } }
                );
                
                // Создаем успешное уведомление для всех ответственных РЭС
                await Notification.create({
                  fromUserId: SYSTEM_USER_ID,
                  toUserId: null, // для всех ответственных РЭС
                  resId: networkStructure.resId,
                  networkStructureId: networkStructure.id,
                  type: 'success',
                  message: `✅ Проблема с ПУ ${fileName} (${networkStructure.tpName} - ${networkStructure.vlName}) успешно устранена!`,
                  isRead: false
                });
                
              } else {
                // НЕУСПЕШНАЯ перепроверка
                
                const newFailureCount = (lastCheckHistory.failureCount || 1) + 1;
                
                // Обновляем историю
                await CheckHistory.update({
                  recheckDate: new Date(),
                  recheckResult: 'error',
                  status: 'awaiting_work', // Возвращаем в работу!
                  failureCount: newFailureCount
                }, {
                  where: {
                    puNumber: fileName,
                    status: 'awaiting_recheck'
                  }
                });
                
                // Обновляем статус ПУ
                await PuStatus.update({
                  status: 'checked_error',
                  errorDetails: result.summary,
                  lastCheck: new Date()
                }, {
                  where: { puNumber: fileName }
                });
                
                // Создаем уведомление об ошибке для РЭС
                await Notification.create({
                  fromUserId: SYSTEM_USER_ID,
                  toUserId: null, // для всех ответственных РЭС
                  resId: networkStructure.resId,
                  networkStructureId: networkStructure.id,
                  type: 'error',
                  message: JSON.stringify({
                    puNumber: fileName,
                    position: position,
                    tpName: networkStructure.tpName,
                    vlName: networkStructure.vlName,
                    resName: networkStructure.ResUnit.name,
                    errorDetails: result.summary,
                    details: result.details
                  }),
                  isRead: false
                });
                
                // Проверяем на проблемную ВЛ (2+ ошибки)
                if (newFailureCount >= 2) {
                  const [problemVL, created] = await ProblemVL.findOrCreate({
                    where: { 
                      puNumber: fileName,
                      status: 'active'
                    },
                    defaults: {
                      networkStructureId: networkStructure.id,
                      resId: networkStructure.resId,
                      tpName: networkStructure.tpName,
                      vlName: networkStructure.vlName,
                      position: position,
                      puNumber: fileName,
                      failureCount: newFailureCount,
                      lastErrorDate: new Date(),
                      lastErrorDetails: result.summary,
                      firstReportDate: lastCheckHistory.initialCheckDate,
                      resComment: lastCheckHistory.resComment
                    }
                  });
                  
                  if (!created) {
                    await problemVL.update({
                      failureCount: newFailureCount,
                      lastErrorDate: new Date(),
                      lastErrorDetails: result.summary
                    });
                  }
                  
                  // Уведомление админам о проблемной ВЛ
                  const admins = await usersWithPerm('problem_vl_view');
                  for (const admin of admins) {
                    await Notification.create({
                      fromUserId: SYSTEM_USER_ID,
                      toUserId: admin.id,
                      resId: networkStructure.resId,
                      networkStructureId: networkStructure.id,
                      type: 'problem_vl',
                      message: JSON.stringify({
                        tpName: networkStructure.tpName,
                        vlName: networkStructure.vlName,
                        puNumber: fileName,
                        position: position,
                        failureCount: newFailureCount,
                        errorDetails: result.summary,
                        resComment: lastCheckHistory.resComment,
                        resName: networkStructure.ResUnit.name
                      }),
                      isRead: false
                    });
                  }
                }
              }
              
            } else {
              // НЕ ПЕРЕПРОВЕРКА - обычная проверка или повторная проверка
              
              // ЕДИНОЕ открытое error-уведомление на ПУ (тот же номер = та же ВЛ и
              // позиция). Пока РЭС не рассмотрел ошибку, повторные съёмы НЕ плодят
              // новые уведомления с разными датами (их одновременное закрытие давало
              // «пакет» и дубли в истории), а ОБНОВЛЯЮТ существующее: свежие детали,
              // счётчик repeatCount, исходная дата firstReportDate сохраняется.
              const openErrorNotifs = await Notification.findAll({
                where: { type: 'error', message: { [Op.like]: `%"puNumber":"${fileName}"%` } },
                order: [['createdAt', 'ASC']]
              });

              // Обновляем статус ПУ
              await PuStatus.upsert({
                puNumber: fileName,
                networkStructureId: networkStructure.id,
                position: position,
                status: result.has_errors ? 'checked_error' : 'checked_ok',
                errorDetails: result.has_errors ? result.summary : null,
                lastCheck: new Date()
              });

              // Линия у техблока (РЭС: «решить невозможно, нужны доп. мероприятия»):
              // статус ПУ обновляем, но уведомлений об ошибке и проблемной ВЛ не
              // создаём, пока техблок не завершит вопрос.
              const lineEscalated = result.has_errors && (await escalatedLineIds()).includes(networkStructure.id);
              if (lineEscalated) {
                console.log(`⏸ ПУ ${fileName}: линия у техблока — уведомление не создаём`);
              } else if (!result.has_errors) {
                // Ошибок нет — снимаем все открытые error-уведомления этого ПУ и
                // автоматически решаем проблемную ВЛ (даже если было 2+ ошибки).
                if (openErrorNotifs.length > 0) {
                  await Notification.destroy({
                    where: { type: 'error', message: { [Op.like]: `%"puNumber":"${fileName}"%` } }
                  });
                }
                await ProblemVL.update(
                  { status: 'resolved' },
                  { where: { puNumber: fileName, status: 'active' } }
                );
              } else if (openErrorNotifs.length > 0) {
                // Ошибка по этому ПУ уже висит — ОБНОВЛЯЕМ её (та же ВЛ/номер),
                // не создаём новое; лишние дубликаты (если накопились ранее) удаляем.
                const keep = openErrorNotifs[0];
                let prev = {};
                try { prev = JSON.parse(keep.message) || {}; } catch { prev = {}; }
                const repeatCount = (prev.repeatCount || 1) + 1;
                const firstReportDate = prev.firstReportDate || keep.createdAt;
                await keep.update({
                  message: JSON.stringify({
                    puNumber: fileName, position,
                    tpName: networkStructure.tpName, vlName: networkStructure.vlName,
                    resName: networkStructure.ResUnit.name,
                    errorDetails: result.summary, details: result.details,
                    repeatCount, firstReportDate
                  }),
                  isRead: false
                });
                for (const dup of openErrorNotifs.slice(1)) await dup.destroy();
                console.log(`↻ ПУ ${fileName}: обновлена открытая ошибка (повтор №${repeatCount}), новое уведомление не создаём`);
              } else {
                // Первая ошибка по этому ПУ — уведомление создаст createNotifications.
                errors.push({
                  puNumber: fileName,
                  error: result.summary,
                  details: result.details,
                  networkStructureId: networkStructure.id,
                  resId: networkStructure.resId
                });
              }
            }
            
            // Записываем успешную загрузку в историю
            if (userId) {
                try {
                  await PuUploadHistory.create({
                    puNumber: fileName,
                    uploadedBy: userId,
                    fileName: originalFileName || 'unknown',
                    fileType: type,
                    periodStart: currentPeriod?.start,
                    periodEnd: currentPeriod?.end,
                    hasErrors: result.has_errors,
                    errorSummary: result.has_errors ? result.summary : null,
                    errorDetails: result.has_errors ? result.details : null,
                    uploadStatus: 'success'
                  });
                } catch (error) {
                  console.error('❌ Error creating PuUploadHistory:', error);
                }
              }
            
            // Добавляем в processed
            processed.push({
              puNumber: fileName,
              status: result.has_errors ? 'checked_error' : 'checked_ok',
              error: result.has_errors ? result.summary : null
            });
            
          } else {
            // ПУ не найден в структуре сети
            processed.push({
              puNumber: fileName,
              status: 'not_in_structure',
              error: 'ПУ не найден в структуре сети'
            });
          }
          
          // Удаляем временный файл
          try {
            fs.unlinkSync(filePath);
          } catch (err) {
            console.error('Error deleting file:', err);
          }
          
          console.log(`Analysis complete: processed=${processed.length}, errors=${errors.length}`);
          resolve({ processed, errors });
          
        } else {
          console.error('Python script returned success=false:', result.error);
          resolve({
            processed: [],
            errors: [result.error || 'Неизвестная ошибка Python скрипта']
          });
        }
      } catch (e) {
        console.error('Failed to parse Python output:', e);
        console.error('Raw output was:', output);
        resolve({
          processed: [],
          errors: [`Ошибка парсинга результата: ${e.message}`]
        });
      }
    };

    const run = (cmd, isFallback) => {
      const child = spawn(cmd, [scriptPath, filePath]);
      current = child;
      child.stdout.on('data', (data) => {
        output += data.toString();
      });
      child.stderr.on('data', (data) => {
        errorOutput += data.toString();
        console.error('Python stderr:', data.toString());
      });
      child.on('error', (error) => {
        child.removeAllListeners('close');   // 'close' первой попытки не должен перебить fallback
        if (!isFallback) return run('python', true);
        console.error('Python process error:', error);
        resolve({ processed: [], errors: ['Python недоступен'] });
      });
      child.on('close', onClose);
    };
    run('python3', false);
  });
}

// Создание уведомлений об ошибках
async function createNotifications(fromUserId, resId, errors) {
  console.log('Creating notifications for errors:', errors);
  
  for (const errorInfo of errors) {
    console.log(`Processing error for PU: ${errorInfo.puNumber}`);
    
    // Находим структуру сети для этого ПУ
    const networkStructure = await NetworkStructure.findOne({
      where: {
        [Op.or]: [
          { startPu: errorInfo.puNumber },
          { middlePu: errorInfo.puNumber },
          { endPu: errorInfo.puNumber }
        ]
      },
      include: [ResUnit]
    });
    
    if (!networkStructure) {
      console.log(`WARNING: No network structure found for PU ${errorInfo.puNumber}`);
      continue;
    }
    
    console.log(`NetworkStructure found: TP=${networkStructure.tpName}, VL=${networkStructure.vlName}`);
    
    // Определяем позицию ПУ
    let position = 'start';
    if (networkStructure.middlePu === errorInfo.puNumber) position = 'middle';
    else if (networkStructure.endPu === errorInfo.puNumber) position = 'end';
    
    // Формируем данные для уведомления с полными деталями
    const errorData = {
      puNumber: errorInfo.puNumber,
      position: position,
      tpName: networkStructure.tpName,
      vlName: networkStructure.vlName,
      resName: networkStructure.ResUnit.name,
      errorDetails: errorInfo.error,
      details: errorInfo.details
    };
    
    console.log('Creating notification with data:', errorData);
    
    // ✅ ИСПРАВЛЕНО: используем resId ИЗ СТРУКТУРЫ СЕТИ!
    try {
      const notification = await Notification.create({
        fromUserId,
        toUserId: null,
        resId: networkStructure.resId,  // ✅ ИЗ СТРУКТУРЫ!
        networkStructureId: networkStructure.id,
        type: 'error',
        message: JSON.stringify(errorData),
        isRead: false
      });
      console.log(`✅ Notification created for RES ${networkStructure.resId} (from structure)`);
    } catch (err) {
      console.error(`Failed to create notification:`, err);
    }
  }
  
  console.log('All notifications created');
}

// =====================================================
// API ДЛЯ ПРОВЕРКИ ЦЕЛОСТНОСТИ БАЗЫ ДАННЫХ
// =====================================================

// Проверка целостности базы данных
app.get('/api/admin/database-health', 
  authenticateToken, 
  requirePerm('db_tools'), 
  async (req, res) => {
    try {
      const issues = [];
      
      console.log('Starting database health check...');
      
      // 1. Проверка ПУ без структуры сети
      try {
        const orphanedPuStatuses = await PuStatus.findAll({
          where: {
            networkStructureId: null
          }
        });
        
        if (orphanedPuStatuses.length > 0) {
          issues.push({
            type: 'orphaned_pu_status',
            severity: 'warning',
            count: orphanedPuStatuses.length,
            description: 'Найдены статусы ПУ без привязки к структуре сети',
            items: orphanedPuStatuses.map(p => p.puNumber).slice(0, 10)
          });
        }
        console.log('Check 1: orphaned_pu_status - OK');
      } catch (err) {
        console.error('Error in check 1:', err);
      }
      
      // 2. Проверка дублирующихся статусов ПУ
      try {
        const [duplicates] = await sequelize.query(`
          SELECT "puNumber", COUNT(*) as count
          FROM "PuStatuses"
          GROUP BY "puNumber"
          HAVING COUNT(*) > 1
        `);
        
        if (duplicates.length > 0) {
          issues.push({
            type: 'duplicate_pu_statuses',
            severity: 'warning',
            count: duplicates.length,
            description: 'Найдены дублирующиеся статусы ПУ',
            items: duplicates.slice(0, 10).map(d => ({
              puNumber: d.puNumber,
              count: parseInt(d.count)
            }))
          });
        }
        console.log('Check 2: duplicate_pu_statuses - OK');
      } catch (err) {
        console.error('Error in check 2:', err);
      }
      
      // 3. Проверка уведомлений без связей
      try {
        const orphanedNotifications = await Notification.findAll({
          where: {
            networkStructureId: {
              [Op.not]: null
            }
          },
          include: [{
            model: NetworkStructure,
            required: false
          }]
        });
        
        const orphaned = orphanedNotifications.filter(n => !n.NetworkStructure);
        
        if (orphaned.length > 0) {
          issues.push({
            type: 'orphaned_notifications',
            severity: 'warning',
            count: orphaned.length,
            description: 'Найдены уведомления со ссылками на несуществующие структуры'
          });
        }
        console.log('Check 3: orphaned_notifications - OK');
      } catch (err) {
        console.error('Error in check 3:', err);
      }
      
      // 4. Проверка истории проверок без РЭС
      try {
        const checksWithoutRes = await CheckHistory.count({
          where: {
            resId: null
          }
        });
        
        if (checksWithoutRes > 0) {
          issues.push({
            type: 'checks_without_res',
            severity: 'warning',
            count: checksWithoutRes,
            description: 'Найдены записи истории без привязки к РЭС'
          });
        }
        console.log('Check 4: checks_without_res - OK');
      } catch (err) {
        console.error('Error in check 4:', err);
      }
      
      // 5. Проверка старых неактивных данных
      try {
        const oldDate = new Date();
        oldDate.setFullYear(oldDate.getFullYear() - 1); // Год назад
        
        const oldNotifications = await Notification.count({
          where: {
            createdAt: {
              [Op.lt]: oldDate
            }
          }
        });
        
        if (oldNotifications > 0) {
          issues.push({
            type: 'old_unread_notifications',
            severity: 'info',
            count: oldNotifications,
            description: 'Найдены уведомления старше года'
          });
        }
        console.log('Check 5: old_unread_notifications - OK');
      } catch (err) {
        console.error('Error in check 5:', err);
      }
      
      // 6. Проверка битых файлов в CheckHistory
      try {
        // Колонка attachments имеет тип JSON — у него в Postgres нет оператора
        // сравнения (ошибка 42883), поэтому сравниваем через каст в text
        const checksWithFiles = await CheckHistory.findAll({
          where: {
            [Op.and]: [
              { attachments: { [Op.not]: null } },
              Sequelize.where(
                Sequelize.cast(Sequelize.col('attachments'), 'text'),
                { [Op.ne]: '[]' }
              )
            ]
          }
        });
        
        let brokenFilesCount = 0;
        for (const check of checksWithFiles) {
          if (check.attachments && Array.isArray(check.attachments)) {
            for (const file of check.attachments) {
              if (!file.url || !file.public_id) {
                brokenFilesCount++;
              }
            }
          }
        }
        
        if (brokenFilesCount > 0) {
          issues.push({
            type: 'broken_file_references',
            severity: 'warning',
            count: brokenFilesCount,
            description: 'Найдены записи с некорректными ссылками на файлы'
          });
        }
        console.log('Check 6: broken_file_references - OK');
      } catch (err) {
        console.error('Error in check 6:', err);
      }
      
      // 7. Проверка пользователей без РЭС (кроме админов)
      try {
        const usersWithoutRes = await User.count({
          where: {
            role: {
              [Op.ne]: 'admin'
            },
            resId: null
          }
        });
        
        if (usersWithoutRes > 0) {
          issues.push({
            type: 'users_without_res',
            severity: 'error',
            count: usersWithoutRes,
            description: 'Найдены не-админы без привязки к РЭС'
          });
        }
        console.log('Check 7: users_without_res - OK');
      } catch (err) {
        console.error('Error in check 7:', err);
      }
      
      // 8. Проверка проблемных ВЛ со статусом active но без recent activity
      try {
        const oldProblemVLs = await ProblemVL.count({
          where: {
            status: 'active',
            lastErrorDate: {
              [Op.lt]: new Date(Date.now() - 90 * 24 * 60 * 60 * 1000) // 90 дней
            }
          }
        });
        
        if (oldProblemVLs > 0) {
          issues.push({
            type: 'stale_problem_vl',
            severity: 'info',
            count: oldProblemVLs,
            description: 'Найдены активные проблемные ВЛ без активности более 90 дней'
          });
        }
        console.log('Check 8: stale_problem_vl - OK');
      } catch (err) {
        console.error('Error in check 8:', err);
      }

      // 9. ✅ НОВОЕ: Проверка неактуальных уведомлений (ПУ зеленые, но уведомления висят)
try {
  console.log('Check 9: stale_notifications - START');
  
  // Получаем все уведомления об ошибках и ожидающие АСКУЭ
  const allErrorNotifs = await Notification.findAll({
    where: {
      type: ['error', 'pending_askue']
    }
  });
  
  console.log(`Found ${allErrorNotifs.length} error/askue notifications to check`);
  
  const staleNotifs = [];
  
  // Проверяем каждое уведомление
  for (const notif of allErrorNotifs) {
    try {
      const data = JSON.parse(notif.message);
      const puNumber = data.puNumber;
      
      if (!puNumber) continue;
      
      // Ищем статус этого ПУ в структуре
      const puStatus = await PuStatus.findOne({
        where: { puNumber }
      });
      
      // Проверяем актуальность уведомления
      if (!puStatus) {
        // Случай 1: ПУ вообще не существует в структуре
        staleNotifs.push({
          notificationId: notif.id,
          type: notif.type,
          puNumber,
          tpName: data.tpName,
          vlName: data.vlName,
          resName: data.resName,
          currentStatus: 'not_found',
          notifCreated: notif.createdAt,
          reason: 'ПУ не найден в структуре сети'
        });
      } else if (puStatus.status === 'checked_ok') {
        // Случай 2: ПУ зеленый, а уведомление висит (самый важный!)
        staleNotifs.push({
          notificationId: notif.id,
          type: notif.type,
          puNumber,
          tpName: data.tpName,
          vlName: data.vlName,
          resName: data.resName,
          currentStatus: 'checked_ok',
          notifCreated: notif.createdAt,
          lastCheck: puStatus.lastCheck,
          reason: 'ПУ уже проверен без ошибок'
        });
      }
    } catch (e) {
      console.error('Error checking notification:', e);
    }
  }
  
  if (staleNotifs.length > 0) {
    issues.push({
      type: 'stale_notifications',
      severity: 'warning',
      count: staleNotifs.length,
      description: 'Найдены неактуальные уведомления (ПУ уже зеленые или не существуют)',
      items: staleNotifs.slice(0, 10) // Показываем первые 10 для предпросмотра
    });
  }
  
  console.log('Check 9: stale_notifications - OK, found:', staleNotifs.length);
} catch (err) {
  console.error('Error in check 9:', err);
}

      // 10. ✅ НОВОЕ: Проверка отсутствующих уведомлений
try {
  console.log('Check 10: missing_notifications - START');
  
  // Находим все ПУ с ошибками
  const errorPuStatuses = await PuStatus.findAll({
    where: {
      status: 'checked_error',
      errorDetails: {
        [Op.not]: null
      }
    },
    include: [{
      model: NetworkStructure,
      include: [ResUnit]
    }]
  });
  
  console.log(`Found ${errorPuStatuses.length} PU with errors`);
  
  const missingNotifications = [];
  
  for (const puStatus of errorPuStatuses) {
    if (!puStatus.NetworkStructure) {
      console.log(`Skipping PU ${puStatus.puNumber} - no structure`);
      continue;
    }
    
    // Проверяем есть ли уведомление для этого ПУ
    const existingNotification = await Notification.findOne({
      where: {
        type: 'error',
        message: {
          [Op.like]: `%"puNumber":"${puStatus.puNumber}"%`
        }
      }
    });
    
    if (!existingNotification) {
      // Уведомления нет - это проблема!
      missingNotifications.push({
        puNumber: puStatus.puNumber,
        puStatusId: puStatus.id,
        networkStructureId: puStatus.networkStructureId,
        position: puStatus.position,
        status: puStatus.status,
        errorDetails: puStatus.errorDetails,
        lastCheck: puStatus.lastCheck,
        tpName: puStatus.NetworkStructure.tpName,
        vlName: puStatus.NetworkStructure.vlName,
        resId: puStatus.NetworkStructure.resId,
        resName: puStatus.NetworkStructure.ResUnit?.name
      });
      
      console.log(`⚠️ Missing notification for PU ${puStatus.puNumber}`);
    }
  }
  
  if (missingNotifications.length > 0) {
    issues.push({
      type: 'missing_notifications',
      severity: 'error',
      count: missingNotifications.length,
      description: 'Найдены ПУ с ошибками без уведомлений',
      items: missingNotifications
    });
  }
  
  console.log('Check 10: missing_notifications - OK, found:', missingNotifications.length);
} catch (err) {
  console.error('Error in check 10:', err);
}

      // 11. Проверка АКТУАЛЬНОСТИ проблемных ВЛ: активная запись, но её ПУ
      // уже проверен без ошибок или удалён из структуры сети. Такие записи
      // могли остаться с времён до внедрения авто-resolve при чистой проверке.
try {
  console.log('Check 11: irrelevant_problem_vl - START');

  const activeProblemVLs = await ProblemVL.findAll({
    where: { status: 'active' }
  });

  console.log(`Found ${activeProblemVLs.length} active problem VLs to check`);

  const irrelevantProblems = [];

  for (const problem of activeProblemVLs) {
    try {
      if (!problem.puNumber) continue;

      const puStatus = await PuStatus.findOne({
        where: { puNumber: problem.puNumber }
      });

      if (!puStatus) {
        // ПУ удалён из структуры — проблема потеряла смысл
        irrelevantProblems.push({
          problemId: problem.id,
          puNumber: problem.puNumber,
          tpName: problem.tpName,
          vlName: problem.vlName,
          failureCount: problem.failureCount,
          lastErrorDate: problem.lastErrorDate,
          currentStatus: 'not_found',
          reason: 'ПУ не найден в структуре сети'
        });
      } else if (puStatus.status === 'checked_ok') {
        // ПУ уже зелёный, а ВЛ всё ещё числится проблемной
        irrelevantProblems.push({
          problemId: problem.id,
          puNumber: problem.puNumber,
          tpName: problem.tpName,
          vlName: problem.vlName,
          failureCount: problem.failureCount,
          lastErrorDate: problem.lastErrorDate,
          lastCheck: puStatus.lastCheck,
          currentStatus: 'checked_ok',
          reason: 'ПУ уже проверен без ошибок'
        });
      }
    } catch (e) {
      console.error('Error checking problem VL:', e);
    }
  }

  if (irrelevantProblems.length > 0) {
    issues.push({
      type: 'irrelevant_problem_vl',
      severity: 'warning',
      count: irrelevantProblems.length,
      description: 'Найдены проблемные ВЛ, потерявшие актуальность (ПУ уже зелёные или удалены из структуры)',
      items: irrelevantProblems.slice(0, 10)
    });
  }

  console.log('Check 11: irrelevant_problem_vl - OK, found:', irrelevantProblems.length);
} catch (err) {
  console.error('Error in check 11:', err);
}
      
      // Итоговая статистика
      // Размер БД + топ-таблицы + оценка роста (для планирования места).
      let dbSize = null;
      if (sequelize.getDialect() === 'postgres') {
        try {
          const rows = await sequelize.query(
            `SELECT pg_database_size(current_database()) AS bytes, current_database() AS name`,
            { type: Sequelize.QueryTypes.SELECT });
          const tbls = await sequelize.query(
            `SELECT c.relname AS "table", pg_total_relation_size(c.oid) AS bytes, c.reltuples::bigint AS rows
               FROM pg_class c JOIN pg_namespace n ON n.oid = c.relnamespace
              WHERE c.relkind = 'r' AND n.nspname = 'public'
              ORDER BY pg_total_relation_size(c.oid) DESC LIMIT 8`,
            { type: Sequelize.QueryTypes.SELECT });
          const since = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);
          const g = await sequelize.query(
            `SELECT
               (SELECT COUNT(*) FROM "CheckHistories" WHERE "createdAt" >= :since) +
               (SELECT COUNT(*) FROM "UploadHistories" WHERE "createdAt" >= :since) +
               (SELECT COUNT(*) FROM "PuUploadHistories" WHERE "uploadedAt" >= :since) +
               (SELECT COUNT(*) FROM "Notifications" WHERE "createdAt" >= :since) AS rows30`,
            { replacements: { since }, type: Sequelize.QueryTypes.SELECT });
          const bytes = Number(rows[0].bytes);
          const totalRows = tbls.reduce((s, t) => s + Number(t.rows || 0), 0);
          const bytesPerRow = totalRows > 0 ? bytes / totalRows : 0;
          const rows30 = Number(g[0].rows30 || 0);
          const bytesPerMonth = Math.round(bytesPerRow * rows30);
          const quotaMb = Number(process.env.DB_QUOTA_MB) || null;
          let monthsLeft = null;
          if (quotaMb && bytesPerMonth > 0) {
            const free = quotaMb * 1024 * 1024 - bytes;
            monthsLeft = free > 0 ? Math.floor(free / bytesPerMonth) : 0;
          }
          dbSize = {
            bytes, name: rows[0].name,
            tables: tbls.map(t => ({ table: t.table, bytes: Number(t.bytes), rows: Number(t.rows) })),
            rows30, bytesPerMonth, quotaMb, monthsLeft
          };
        } catch (e) { console.warn('dbSize failed:', e.message); }
      }

      const stats = {
  totalIssues: issues.length,
  dbSize,
  byType: {
    error: issues.filter(i => i.severity === 'error').length,
    warning: issues.filter(i => i.severity === 'warning').length,
    info: issues.filter(i => i.severity === 'info').length
  },
  staleNotifications: issues.find(i => i.type === 'stale_notifications')?.count || 0, // ✅ ДОБАВЛЕНО
  missingNotifications: issues.find(i => i.type === 'missing_notifications')?.count || 0,
  irrelevantProblemVL: issues.find(i => i.type === 'irrelevant_problem_vl')?.count || 0,
  totalRecords: {
    networkStructures: await NetworkStructure.count(),
    puStatuses: await PuStatus.count(),
    notifications: await Notification.count(),
    checkHistory: await CheckHistory.count(),
    uploadHistory: await UploadHistory.count(),
    users: await User.count()
  }
};
      
      console.log('Database health check completed!');
      console.log('Total issues found:', issues.length);
      
      res.json({
        success: true,
        stats,
        issues,
        checkedAt: new Date()
      });
      
    } catch (error) {
      console.error('Database health check error:', error);
      res.status(500).json({ 
        success: false,
        error: error.message,
        stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
      });
    }
});

// Очистка определенного типа проблем
app.post('/api/admin/database-cleanup',
  authenticateToken,
  requirePerm('db_tools'),
  async (req, res) => {
    const transaction = await sequelize.transaction();
    
    try {
      const { cleanupType, password } = req.body;
      
      console.log('Cleanup request:', cleanupType);
      
      if (password !== DELETE_PASSWORD) {
        await transaction.rollback();
        return res.status(403).json({ error: 'Неверный пароль' });
      }
      
      let cleaned = 0;
      
      switch (cleanupType) {
        case 'orphaned_pu_status':
          // Удаляем статусы ПУ без структуры
          cleaned = await PuStatus.destroy({
            where: {
              networkStructureId: null
            },
            transaction
          });
          console.log(`Cleaned orphaned_pu_status: ${cleaned}`);
          break;
          
        case 'duplicate_pu_statuses':
          // Удаляем дубликаты, оставляя последний
          const [duplicates] = await sequelize.query(`
            SELECT "puNumber", COUNT(*) as count
            FROM "PuStatuses"
            GROUP BY "puNumber"
            HAVING COUNT(*) > 1
          `, { transaction });
          
          for (const dup of duplicates) {
            // Оставляем самый новый
            const allStatuses = await PuStatus.findAll({
              where: { puNumber: dup.puNumber },
              order: [['updatedAt', 'DESC']],
              transaction
            });
            
            // Удаляем все кроме первого
            for (let i = 1; i < allStatuses.length; i++) {
              await allStatuses[i].destroy({ transaction });
              cleaned++;
            }
          }
          console.log(`Cleaned duplicate_pu_statuses: ${cleaned}`);
          break;
          
        case 'old_unread_notifications':
          // Удаляем старые уведомления
          const oldDate = new Date();
          oldDate.setFullYear(oldDate.getFullYear() - 1);
          
          cleaned = await Notification.destroy({
            where: {
              createdAt: {
                [Op.lt]: oldDate
              }
            },
            transaction
          });
          console.log(`Cleaned old_unread_notifications: ${cleaned}`);
          break;
          
        case 'orphaned_notifications':
          // Находим и удаляем уведомления с несуществующими связями
          const orphanedNotifs = await Notification.findAll({
            where: {
              networkStructureId: {
                [Op.not]: null
              }
            },
            include: [{
              model: NetworkStructure,
              required: false
            }],
            transaction
          });
          
          for (const notif of orphanedNotifs) {
            if (!notif.NetworkStructure) {
              await notif.destroy({ transaction });
              cleaned++;
            }
          }
          console.log(`Cleaned orphaned_notifications: ${cleaned}`);
          break;
          
        case 'checks_without_res':
          // Удаляем историю проверок без РЭС
          cleaned = await CheckHistory.destroy({
            where: {
              resId: null
            },
            transaction
          });
          console.log(`Cleaned checks_without_res: ${cleaned}`);
          break;
          
        case 'broken_file_references':
          // Чистим битые ссылки на файлы (JSON-колонка — сравнение через ::text)
          const checksWithFiles = await CheckHistory.findAll({
            where: {
              [Op.and]: [
                { attachments: { [Op.not]: null } },
                Sequelize.where(
                  Sequelize.cast(Sequelize.col('attachments'), 'text'),
                  { [Op.ne]: '[]' }
                )
              ]
            },
            transaction
          });
          
          for (const check of checksWithFiles) {
            if (check.attachments && Array.isArray(check.attachments)) {
              const validFiles = check.attachments.filter(
                file => file.url && file.public_id
              );
              
              if (validFiles.length !== check.attachments.length) {
                await check.update({ attachments: validFiles }, { transaction });
                cleaned++;
              }
            }
          }
          console.log(`Cleaned broken_file_references: ${cleaned}`);
          break;
          
        case 'stale_problem_vl':
          // Закрываем старые проблемные ВЛ
          const oldProblemDate = new Date(Date.now() - 90 * 24 * 60 * 60 * 1000);
          
          cleaned = await ProblemVL.update(
            { status: 'dismissed' },
            {
              where: {
                status: 'active',
                lastErrorDate: {
                  [Op.lt]: oldProblemDate
                }
              },
              transaction
            }
          );
          console.log(`Cleaned stale_problem_vl: ${cleaned}`);
          break;

        case 'irrelevant_problem_vl': {
          // Закрываем проблемные ВЛ, потерявшие актуальность: ПУ уже
          // проверен без ошибок или удалён из структуры. Статус — resolved
          // (как в авто-resolve при чистой проверке), НЕ dismissed.
          const activeProblems = await ProblemVL.findAll({
            where: { status: 'active' },
            transaction
          });

          for (const problem of activeProblems) {
            if (!problem.puNumber) continue;

            const puStatus = await PuStatus.findOne({
              where: { puNumber: problem.puNumber },
              transaction
            });

            if (!puStatus || puStatus.status === 'checked_ok') {
              await problem.update({ status: 'resolved' }, { transaction });
              cleaned++;
            }
          }
          console.log(`Cleaned irrelevant_problem_vl: ${cleaned}`);
          break;
        }

          case 'stale_notifications':
  // ✅ НОВОЕ: Очистка неактуальных уведомлений
  console.log('Cleaning stale notifications...');
  
  const allErrorNotifs = await Notification.findAll({
    where: {
      type: ['error', 'pending_askue']
    },
    transaction
  });
  
  console.log(`Found ${allErrorNotifs.length} error/askue notifications to check`);
  
  for (const notif of allErrorNotifs) {
    try {
      const data = JSON.parse(notif.message);
      const puNumber = data.puNumber;
      
      if (!puNumber) continue;
      
      const puStatus = await PuStatus.findOne({
        where: { puNumber },
        transaction
      });
      
      // Удаляем если ПУ зеленый или не существует
      if (!puStatus || puStatus.status === 'checked_ok') {
        await notif.destroy({ transaction });
        cleaned++;
        console.log(`Deleted stale notification for PU ${puNumber} (${puStatus ? 'checked_ok' : 'not_found'})`);
      }
    } catch (e) {
      console.error('Error cleaning stale notification:', e);
    }
  }
  
  console.log(`Cleaned stale_notifications: ${cleaned}`);
  break;

    case 'missing_notifications':
  // ✅ НОВОЕ: Создаем отсутствующие уведомления
  console.log('Creating missing notifications...');
  
  const errorPuStatuses = await PuStatus.findAll({
    where: {
      status: 'checked_error',
      errorDetails: {
        [Op.not]: null
      }
    },
    include: [{
      model: NetworkStructure,
      include: [ResUnit]
    }],
    transaction
  });
  
  console.log(`Found ${errorPuStatuses.length} PU with errors`);
  
  for (const puStatus of errorPuStatuses) {
    if (!puStatus.NetworkStructure) continue;
    
    // Проверяем есть ли уже уведомление
    const existingNotification = await Notification.findOne({
      where: {
        type: 'error',
        message: {
          [Op.like]: `%"puNumber":"${puStatus.puNumber}"%`
        }
      },
      transaction
    });
    
    if (!existingNotification) {
      // Создаем недостающее уведомление
      const errorData = {
        puNumber: puStatus.puNumber,
        position: puStatus.position,
        tpName: puStatus.NetworkStructure.tpName,
        vlName: puStatus.NetworkStructure.vlName,
        resName: puStatus.NetworkStructure.ResUnit?.name,
        errorDetails: puStatus.errorDetails,
        details: null,
        restoredFrom: 'database_health_check', // Метка что восстановлено
        restoredAt: new Date()
      };
      
      await Notification.create({
        fromUserId: SYSTEM_USER_ID, // Системное уведомление
        toUserId: null,
        resId: puStatus.NetworkStructure.resId,
        networkStructureId: puStatus.networkStructureId,
        puStatusId: puStatus.id,
        type: 'error',
        message: JSON.stringify(errorData),
        isRead: false
      }, { transaction });
      
      cleaned++;
      console.log(`✅ Created missing notification for PU ${puStatus.puNumber}`);
    }
  }
  
  console.log(`Created missing_notifications: ${cleaned}`);
  break;
          
        default:
          await transaction.rollback();
          return res.status(400).json({ error: 'Неизвестный тип очистки' });
      }
      
      await transaction.commit();
      
      console.log(`Cleanup completed! Type: ${cleanupType}, Cleaned: ${cleaned}`);
      
      res.json({
        success: true,
        message: `Очищено записей: ${cleaned}`,
        cleanupType,
        cleaned
      });
      
    } catch (error) {
      await transaction.rollback();
      console.error('Database cleanup error:', error);
      res.status(500).json({ 
        success: false,
        error: error.message,
        stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
      });
    }
});

// =====================================================
// ENDPOINT ДЛЯ СКАЧИВАНИЯ ФАЙЛОВ С ПРАВИЛЬНЫМ ИМЕНЕМ
// =====================================================

// Прокси файлов из Cloudinary. ВАЖНО: не редиректим браузер на
// res.cloudinary.com (общий CDN-домен, его блокируют Яндекс.Браузер/фильтры) —
// сервер сам скачивает файл и отдаёт его со своего домена.
// ?inline=1 — показать в браузере (картинки/просмотр), без — скачать файлом.
// Прокси файлов Cloudinary через свой бэкенд. ОДИН обработчик на два пути:
// /api/download — старый (обратная совместимость, вдруг где-то сохранены ссылки),
// /api/f — новый, т.к. слова "download"/"attachment_" в URL режут блокировщики
// рекламы (ERR_BLOCKED_BY_CLIENT: Яндекс Protect, AdGuard, uBlock).
// Идентификатор из URL → public_id. Новый /api/f шлёт base64url (прячем
// 'attachment_'/'download' от блокировщиков рекламы). Если это не наш base64url —
// трактуем как сырой public_id (старые ссылки /api/f/res-management%2F… и /api/download).
function resolvePublicId(raw) {
  if (!raw) return null;
  try {
    const decoded = Buffer.from(String(raw), 'base64url').toString('utf8');
    if (decoded && decoded.startsWith('res-management/')) return decoded;
  } catch (e) {}
  return raw;
}

// Токен /api/f → { publicId, name, inline }. Порядок разбора:
//  1) base64url(JSON {p,n,i}) — новый формат без query-строки;
//  2) base64url(строка res-management/…) — прежний формат, name/inline из query;
//  3) сырой public_id (старые ссылки /api/f/res-management%2F… и /api/download).
function parseFileToken(raw, query) {
  const q = query || {};
  const name = q.name || 'file';
  const inline = q.inline === '1';
  if (!raw) return { publicId: null, name, inline };
  try {
    const decoded = Buffer.from(String(raw), 'base64url').toString('utf8');
    if (decoded && decoded[0] === '{') {
      try {
        const obj = JSON.parse(decoded);
        if (obj && typeof obj.p === 'string' && obj.p.startsWith('res-management/')) {
          return { publicId: obj.p, name: obj.n || name, inline: obj.i === 1 || obj.i === '1' };
        }
      } catch (e) {}
    }
    if (decoded && decoded.startsWith('res-management/')) {
      return { publicId: decoded, name, inline };
    }
  } catch (e) {}
  return { publicId: raw, name, inline };
}

// Кэш РАБОЧЕГО delivery-URL: publicId -> url (без TTL), чтобы не делать по 4-6
// запросов к Cloudinary на каждое открытие файла.
const fileVariantCache = new Map();

async function handleFileProxy(req, res) {
  try {
    // Express УЖЕ декодировал параметр — второй decodeURIComponent падал на литеральном '%'.
    const parsed = parseFileToken(req.params.public_id, req.query);
    const publicId = parsed.publicId;
    // Защита от открытого прокси: отдаём только файлы нашего аккаунта.
    if (!publicId || !publicId.startsWith('res-management/')) {
      return res.status(403).json({ error: 'Недопустимый идентификатор файла' });
    }
    const originalName = parsed.name || 'file';
    const inline = parsed.inline;

    const isPdf = publicId.toLowerCase().endsWith('.pdf');
    const primary = isPdf ? 'raw' : 'image';
    const alt = isPdf ? 'image' : 'raw';
    const withoutExt = publicId.replace(/\.[^/.]+$/, '');
    const ext = (publicId.match(/\.([^/.]+)$/) || [])[1] || null;
    const signed = (id, rt, type, extra = {}) =>
      cloudinary.url(id, { resource_type: rt, type, secure: true, sign_url: true, ...extra });

    let upstream = null, winnerUrl = null;
    const tried = [];
    const tryUrl = async (label, url) => {
      if (!url || upstream) return;
      const rec = { label, status: null, xCldError: null };
      tried.push(rec);
      let r;
      try { r = await fetch(url); }
      catch (e) { rec.xCldError = e.message; return; }
      rec.status = r.status;
      rec.xCldError = r.headers.get('x-cld-error') || null;
      if (r.ok && r.body) { upstream = r; winnerUrl = url; rec.ok = true; return; }
      try { r.body && r.body.cancel && r.body.cancel(); } catch (e) {}
    };
    // Берём secure_url прямо из api.resource (не угадываем URL). Возвращает список URL.
    const resourceUrls = async (rt) => {
      try {
        const r = await cloudinary.api.resource(publicId, { resource_type: rt });
        const out = [];
        if (r.secure_url) out.push(r.secure_url);
        if (r.version) out.push(signed(publicId, rt, 'upload', { version: r.version }));
        return out;
      } catch (e) { return []; }
    };

    // 0) Кешированный рабочий URL — первым (обычно 1 запрос).
    await tryUrl('cached', fileVariantCache.get(publicId));

    // 1) Базовая попытка (как раньше): resource_type по расширению + upload, signed.
    if (!upstream) await tryUrl(`${primary}/upload signed`, signed(publicId, primary, 'upload'));

    // 2) ПЕРВЫМ после базовой — URL У CLOUDINARY: api.resource(secure_url), raw и image.
    if (!upstream) for (const u of await resourceUrls(primary)) { await tryUrl(`api.resource(${primary})`, u); if (upstream) break; }
    if (!upstream) for (const u of await resourceUrls(alt)) { await tryUrl(`api.resource(${alt})`, u); if (upstream) break; }

    // 3) Остальная существующая цепочка: альт. resource_type → authenticated → no-ext.
    if (!upstream) await tryUrl(`${alt}/upload signed`, signed(publicId, alt, 'upload'));
    if (!upstream) await tryUrl('image/authenticated signed', signed(publicId, 'image', 'authenticated'));
    if (!upstream) await tryUrl('raw/authenticated signed', signed(publicId, 'raw', 'authenticated'));
    if (!upstream && ext && withoutExt !== publicId) await tryUrl('image no-ext +format', signed(withoutExt, 'image', 'upload', { format: ext }));

    if (!upstream) {
      const summary = tried.map(t => `${t.label}=${t.status || 'ERR'}${t.xCldError ? `(${t.xCldError})` : ''}`).join(' | ');
      console.error(`File proxy 404 for ${publicId}; tried: ${summary}`);
      return res.status(404).json({ error: 'Файл отсутствует в хранилище', publicId });
    }
    fileVariantCache.set(publicId, winnerUrl); // запомнить рабочий URL

    res.setHeader('Content-Type',
      upstream.headers.get('content-type') || 'application/octet-stream');
    const contentLength = upstream.headers.get('content-length');
    if (contentLength) res.setHeader('Content-Length', contentLength);
    res.setHeader('Content-Disposition',
      `${inline ? 'inline' : 'attachment'}; filename*=UTF-8''${encodeURIComponent(originalName)}`);
    res.setHeader('Cache-Control', 'private, max-age=3600');

    const { Readable } = require('stream');
    Readable.fromWeb(upstream.body).pipe(res);

  } catch (error) {
    console.error('❌ Download error:', error);
    if (!res.headersSent) {
      res.status(500).json({ error: 'Ошибка скачивания файла: ' + error.message });
    }
  }
}
app.get('/api/download/:public_id', authenticateFileAccess, handleFileProxy); // legacy — не удалять
app.get('/api/f/:public_id', authenticateFileAccess, handleFileProxy);        // новый (обход блокировщиков)

// Диагностика: проверяем существование ресурса во ВСЕХ комбинациях Cloudinary —
// resource_type(image|raw) × type(upload|authenticated) × public_id(как есть|без
// расширения). Возвращаем матрицу: где нашёлся (bytes/format/created_at), где 404.
async function cloudinaryProbe(publicId) {
  const withoutExt = publicId.replace(/\.[^/.]+$/, '');
  const ext = (publicId.match(/\.([^/.]+)$/) || [])[1] || null;
  const idVariants = [{ label: 'as-is', id: publicId }];
  if (withoutExt !== publicId) idVariants.push({ label: 'no-ext', id: withoutExt });
  const results = [];
  let found = null, foundResource = null;
  for (const rtype of ['image', 'raw']) {
    for (const type of ['upload', 'authenticated']) {
      for (const v of idVariants) {
        const combo = { resource_type: rtype, type, idVariant: v.label, publicId: v.id };
        try {
          const r = await cloudinary.api.resource(v.id, { resource_type: rtype, type });
          const ok = { ...combo, ok: true, bytes: r.bytes, format: r.format, created_at: r.created_at, version: r.version };
          results.push(ok);
          if (!found) { found = ok; foundResource = r; }
        } catch (e) {
          results.push({ ...combo, ok: false, http: (e && (e.error?.http_code || e.http_code)) || null });
        }
      }
    }
  }

  // Для НАЙДЕННОЙ комбинации проверяем реальную ДОСТАВКУ по нескольким delivery-URL
  // и читаем заголовок x-cld-error (официальная причина отказа Cloudinary).
  let deliveryTests = null;
  if (found && foundResource) {
    const rt = found.resource_type, type = found.type, id = found.publicId, r = foundResource;
    const variants = [
      { label: 'proxy (как сейчас, signed)', url: cloudinary.url(id, { resource_type: rt, type, secure: true, sign_url: true }) },
      { label: 'signed + version', url: cloudinary.url(id, { resource_type: rt, type, secure: true, sign_url: true, version: r.version }) },
      { label: 'resource.secure_url', url: r.secure_url },
      { label: 'unsigned', url: cloudinary.url(id, { resource_type: rt, type, secure: true }) },
    ];
    deliveryTests = [];
    for (const v of variants) {
      const t = { label: v.label };
      try {
        const resp = await fetch(v.url);
        t.status = resp.status;
        t.xCldError = resp.headers.get('x-cld-error') || null;
        try { resp.body && resp.body.cancel && resp.body.cancel(); } catch (e) {}
      } catch (e) {
        t.status = null; t.xCldError = e.message;
      }
      deliveryTests.push(t);
    }
  }

  return { publicId, ext, found, results, deliveryTests };
}

app.get('/api/admin/files/diag/:token', authenticateToken, requirePerm('files_manage'), async (req, res) => {
  try {
    const publicId = resolvePublicId(req.params.token);
    if (!publicId || !publicId.startsWith('res-management/')) {
      return res.status(400).json({ error: 'Некорректный идентификатор файла' });
    }
    res.json(await cloudinaryProbe(publicId));
  } catch (error) {
    console.error('files diag error:', error.message);
    res.status(500).json({ error: error.message });
  }
});

// =====================================================
// ОЧИСТКА ИСТОРИИ ДО ДАТЫ (admin). Удаляются истории/уведомления/файлы СТАРШЕ
// даты. НИКОГДА не трогаются: NetworkStructure, PuStatus, TpSection,
// SectionMonthlyPeak, OverloadCase, Users, ResUnits.
// =====================================================

// Валидация before → Date (начало суток UTC) или null.
function parseBeforeDate(before) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(before || ''))) return null;
  const d = new Date(before + 'T00:00:00Z');
  return isNaN(d.getTime()) ? null : d;
}

// Собрать Cloudinary public_id из CheckHistory старше before (в транзакции, если есть).
async function collectPurgeFileIds(beforeDate, transaction) {
  const checks = await CheckHistory.findAll({
    where: { createdAt: { [Op.lt]: beforeDate } },
    attributes: ['attachments'],
    transaction
  });
  const ids = [];
  for (const ch of checks) {
    const atts = Array.isArray(ch.attachments) ? ch.attachments : [];
    for (const a of atts) if (a && a.public_id) ids.push(a.public_id);
  }
  return ids;
}

// Предпросмотр: только COUNT, ничего не удаляет.
app.post('/api/admin/purge-preview', authenticateToken, requirePerm('history_purge'), async (req, res) => {
  try {
    const beforeDate = parseBeforeDate(req.body.before);
    if (!beforeDate) return res.status(400).json({ error: 'Некорректная дата (нужен формат YYYY-MM-DD)' });
    const [uploadHistory, puUploadHistory, checkHistory, notifications, fileIds] = await Promise.all([
      UploadHistory.count({ where: { createdAt: { [Op.lt]: beforeDate } } }),
      PuUploadHistory.count({ where: { uploadedAt: { [Op.lt]: beforeDate } } }),
      CheckHistory.count({ where: { createdAt: { [Op.lt]: beforeDate } } }),
      Notification.count({ where: { createdAt: { [Op.lt]: beforeDate } } }),
      collectPurgeFileIds(beforeDate, null)
    ]);
    res.json({ uploadHistory, puUploadHistory, checkHistory, notifications, cloudinaryFiles: fileIds.length });
  } catch (error) {
    console.error('purge-preview error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Удаление. Транзакция БД, затем (после commit) — файлы Cloudinary батчами.
app.post('/api/admin/purge', authenticateToken, requirePerm('history_purge'), async (req, res) => {
  const { before, password } = req.body;
  if (password !== DELETE_PASSWORD) return res.status(403).json({ error: 'Неверный пароль' });
  const beforeDate = parseBeforeDate(before);
  if (!beforeDate) return res.status(400).json({ error: 'Некорректная дата (нужен формат YYYY-MM-DD)' });

  const transaction = await sequelize.transaction();
  let fileIds = [];
  const deleted = {};
  try {
    // a) собрать файлы из CheckHistory ДО удаления строк
    fileIds = await collectPurgeFileIds(beforeDate, transaction);

    // b) порядок: NotificationRead (по удаляемым уведомлениям) → Notification →
    // CheckHistory → UploadHistory → PuUploadHistory. Notification.destroy идёт
    // ЧЕРЕЗ МОДЕЛЬ (не raw SQL), чтобы сработал хук инвалидации кэша counts.
    const notifIds = (await Notification.findAll({
      where: { createdAt: { [Op.lt]: beforeDate } }, attributes: ['id'], transaction
    })).map(n => n.id);
    if (notifIds.length) {
      await NotificationRead.destroy({ where: { notificationId: notifIds }, transaction });
    }
    deleted.notifications = await Notification.destroy({ where: { createdAt: { [Op.lt]: beforeDate } }, transaction });
    deleted.checkHistory = await CheckHistory.destroy({ where: { createdAt: { [Op.lt]: beforeDate } }, transaction });
    deleted.uploadHistory = await UploadHistory.destroy({ where: { createdAt: { [Op.lt]: beforeDate } }, transaction });
    deleted.puUploadHistory = await PuUploadHistory.destroy({ where: { uploadedAt: { [Op.lt]: beforeDate } }, transaction });

    // c) commit ДО работы с Cloudinary
    await transaction.commit();
  } catch (error) {
    await transaction.rollback();
    console.error('purge error:', error);
    return res.status(500).json({ error: error.message });
  }

  // После commit — удаляем файлы; ошибки Cloudinary НЕ откатывают БД.
  const cl = await purgeCloudinary(fileIds);
  invalidateNotifCountsCache(); // на всякий случай (destroy уже дёрнул хук)
  res.json({ deleted, cloudinaryDeleted: cl.deleted, cloudinaryErrors: cl.errors });
});

// =====================================================
// БЭКАП / ВОССТАНОВЛЕНИЕ (перенос данных Render → Amvera)
// =====================================================

// Порядок таблиц: родители → дети. Вставка идёт в этом порядке, очистка — в
// обратном (чтобы не нарушать FK).
const BACKUP_TABLES = [
  { key: 'ResUnits', table: 'ResUnits' },
  { key: 'Users', table: 'Users' },
  { key: 'NetworkStructures', table: 'NetworkStructures' },
  { key: 'PuStatuses', table: 'PuStatuses' },
  { key: 'UploadHistories', table: 'UploadHistories' },
  { key: 'CheckHistories', table: 'CheckHistories' },
  { key: 'ProblemVLs', table: 'ProblemVLs' },
  { key: 'Notifications', table: 'Notifications' },
  { key: 'NotificationReads', table: 'NotificationReads' },
  { key: 'PuUploadHistories', table: 'PuUploadHistories' },
  { key: 'RolePermissions', table: 'RolePermissions' },
  { key: 'SeededRoles', table: 'SeededRoles' },
  { key: 'Escalations', table: 'Escalations' },
];

// Полный JSON-дамп всех таблиц. Пароли выгружаются как есть (bcrypt-хэши),
// ссылки Cloudinary едут внутри JSON — сами файлы переносить не нужно.
app.get('/api/admin/backup', authenticateToken, requirePerm('db_tools'), async (req, res) => {
  try {
    const tables = {};
    for (const t of BACKUP_TABLES) {
      const [rows] = await sequelize.query(`SELECT * FROM "${t.table}" ORDER BY id`);
      tables[t.key] = rows;
    }
    const backup = {
      format: 'full',
      version: 1,
      exportedAt: new Date().toISOString(),
      tables,
    };
    const dateStr = new Date().toISOString().slice(0, 10);
    // Отдаём gzip-ом (Content-Encoding: gzip) — браузер распакует прозрачно,
    // пользователь получает обычный .json. Обходит лимиты/таймауты прокси Amvera.
    const gz = zlib.gzipSync(Buffer.from(JSON.stringify(backup), 'utf8'));
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
    res.setHeader('Content-Encoding', 'gzip');
    res.setHeader('Content-Disposition', `attachment; filename="res-backup-${dateStr}.json"`);
    res.send(gz);
  } catch (error) {
    console.error('Backup error:', error.message);
    res.status(500).json({ error: 'Ошибка бэкапа: ' + error.message });
  }
});

// Восстановление из JSON-файла. ПОЛНОСТЬЮ заменяет данные (чистит → вставляет с
// явными id → сбрасывает sequence). Файл принимаем через multer (в обход лимита
// express.json). Вставка raw-SQL — мимо hooks, пароли остаются как есть. Ошибки
// копим (первые 20), не падаем целиком. Пользователь БД не суперюзер — только
// правильный порядок вставки, без отключения триггеров.
const backupUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 100 * 1024 * 1024 } });

app.post('/api/admin/restore', authenticateToken, requirePerm('db_tools'), backupUpload.single('file'), async (req, res) => {
  try {
    if (req.body.confirm !== 'true') {
      return res.status(400).json({ error: 'Требуется подтверждение: восстановление заменит ВСЕ данные в базе.' });
    }
    if (!req.file) {
      return res.status(400).json({ error: 'Файл бэкапа не передан' });
    }
    let backup;
    try {
      // Клиент сжимает файл gzip (прокси Amvera режет большие тела запросов).
      // Определяем gzip по магическим байтам 1f 8b и распаковываем; несжатый
      // файл обрабатывается как раньше.
      let buf = req.file.buffer;
      if (buf.length > 2 && buf[0] === 0x1f && buf[1] === 0x8b) {
        buf = zlib.gunzipSync(buf);
      }
      backup = JSON.parse(buf.toString('utf8'));
    } catch (e) {
      return res.status(400).json({ error: 'Файл не является корректным JSON' });
    }
    if (!backup || backup.format !== 'full' || !backup.tables) {
      return res.status(400).json({ error: 'Неверный формат бэкапа (ожидается format: "full")' });
    }

    const errors = [];
    let inserted = 0;

    // 1) Очистка в обратном порядке (дети → родители)
    for (const t of [...BACKUP_TABLES].reverse()) {
      await sequelize.query(`DELETE FROM "${t.table}"`);
    }

    // 2) Вставка в прямом порядке с явными id
    for (const t of BACKUP_TABLES) {
      const rows = backup.tables[t.key] || [];
      for (const row of rows) {
        try {
          const cols = Object.keys(row);
          if (cols.length === 0) continue;
          const colList = cols.map((c) => `"${c}"`).join(', ');
          const placeholders = cols.map((_, i) => `$${i + 1}`).join(', ');
          const values = cols.map((c) => {
            const v = row[c];
            // JSON/JSONB-объекты из бэкапа — обратно в строку (pg приведёт к jsonb)
            if (v !== null && typeof v === 'object' && !(v instanceof Date)) return JSON.stringify(v);
            return v;
          });
          await sequelize.query(
            `INSERT INTO "${t.table}" (${colList}) VALUES (${placeholders})`,
            { bind: values }
          );
          inserted++;
        } catch (e) {
          if (errors.length < 20) errors.push(`${t.table} id=${row.id}: ${e.message}`);
        }
      }
      // 3) Сброс sequence под текущий max(id)
      try {
        await sequelize.query(
          `SELECT setval(pg_get_serial_sequence('"${t.table}"', 'id'),
             COALESCE((SELECT MAX(id) FROM "${t.table}"), 1),
             (SELECT MAX(id) IS NOT NULL FROM "${t.table}"))`
        );
      } catch (e) {
        if (errors.length < 20) errors.push(`setval ${t.table}: ${e.message}`);
      }
    }

    res.json({ ok: true, inserted, errorsCount: errors.length, errors });
  } catch (error) {
    console.error('Restore error:', error.message);
    res.status(500).json({ error: 'Ошибка восстановления: ' + error.message });
  }
});

// =====================================================
// РАЗДАЧА ФРОНТЕНДА (единый origin) — ПОСЛЕ всех /api-роутов
// =====================================================
const FRONTEND_DIST = path.join(__dirname, '..', 'frontend', 'dist');
if (fs.existsSync(FRONTEND_DIST)) {
  app.use(express.static(FRONTEND_DIST));
  // SPA-fallback: любой НЕ-API путь → index.html (клиентский роутинг).
  // ДОЛЖЕН быть последним middleware, после всех /api-роутов, иначе перехватит API.
  app.get('*', (req, res, next) => {
    if (req.path.startsWith('/api') || req.path.startsWith('/uploads')) return next();
    res.sendFile(path.join(FRONTEND_DIST, 'index.html'));
  });
}

// Обработчик ошибок multer — ДОЛЖЕН стоять ПОСЛЕ всех роутов и SPA-fallback,
// иначе Express не отдаст ему ошибки роутов, объявленных ниже. На не-multer
// ошибки не влияет (next(error)).
app.use((error, req, res, next) => {
  if (error instanceof multer.MulterError) {
    if (error.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({
        error: 'Файл слишком большой. Максимальный размер: 10MB'
      });
    }
  }
  next(error);
});

// =====================================================
// ИНИЦИАЛИЗАЦИЯ БД И ЗАПУСК СЕРВЕРА
// =====================================================

// Amvera: DNS-имя БД может быть ещё не резолвимо в момент старта контейнера
// ("Temporary failure in name resolution"). Ретраим подключение, не падаем сразу.
async function connectWithRetry(maxAttempts = 15, delayMs = 3000) {
  let lastError;
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      await sequelize.authenticate();
      if (attempt > 1) console.log(`Database connected on attempt ${attempt}/${maxAttempts}`);
      return;
    } catch (err) {
      lastError = err;
      console.warn(`Database not ready (attempt ${attempt}/${maxAttempts}): ${err.code || err.name}. Retry in ${delayMs}ms...`);
      if (attempt < maxAttempts) await new Promise((r) => setTimeout(r, delayMs));
    }
  }
  throw lastError;
}

async function initializeDatabase() {
  try {
    await connectWithRetry();
    console.log('Database connected successfully');
    
    // ✅ PERF: alter:true при КАЖДОМ старте — медленно и рискованно на Postgres.
    // Теперь alter включается разово через переменную окружения DB_ALTER=true
    // (после изменения схемы моделей), в остальное время — быстрый sync.
    await sequelize.sync({ alter: process.env.DB_ALTER === 'true' });
    console.log('All models synchronized');
    
    // ✅ PERF: индексы на "горячих" полях. CREATE INDEX IF NOT EXISTS — идемпотентно,
    // при повторных стартах выполняется мгновенно.
    const indexStatements = [
      `CREATE INDEX IF NOT EXISTS idx_notifications_type ON "Notifications" ("type")`,
      `CREATE INDEX IF NOT EXISTS idx_notifications_res ON "Notifications" ("resId")`,
      `CREATE INDEX IF NOT EXISTS idx_notifications_touser ON "Notifications" ("toUserId")`,
      `CREATE INDEX IF NOT EXISTS idx_notifications_created ON "Notifications" ("createdAt" DESC)`,
      `CREATE INDEX IF NOT EXISTS idx_notifreads_user ON "NotificationReads" ("userId")`,
      `CREATE INDEX IF NOT EXISTS idx_checkhist_pu ON "CheckHistories" ("puNumber")`,
      `CREATE INDEX IF NOT EXISTS idx_checkhist_res ON "CheckHistories" ("resId")`,
      `CREATE INDEX IF NOT EXISTS idx_checkhist_status ON "CheckHistories" ("status")`,
      `CREATE INDEX IF NOT EXISTS idx_checkhist_created ON "CheckHistories" ("createdAt" DESC)`,
      `CREATE INDEX IF NOT EXISTS idx_puuploadhist_pu ON "PuUploadHistories" ("puNumber")`,
      `CREATE INDEX IF NOT EXISTS idx_netstruct_res ON "NetworkStructures" ("resId")`,
      `CREATE INDEX IF NOT EXISTS idx_pustatus_struct ON "PuStatuses" ("networkStructureId")`,
      `CREATE INDEX IF NOT EXISTS idx_problemvl_status ON "ProblemVLs" ("status")`,
      `CREATE INDEX IF NOT EXISTS idx_problemvl_res ON "ProblemVLs" ("resId")`,
      // Единый вход через платформу: колонка keycloakId + уникальный индекс.
      // ADD COLUMN IF NOT EXISTS — доезжает без DB_ALTER (Postgres-safe,
      // уникальный индекс допускает много NULL).
      `ALTER TABLE "Users" ADD COLUMN IF NOT EXISTS "keycloakId" VARCHAR(64)`,
      `CREATE UNIQUE INDEX IF NOT EXISTS idx_users_keycloak ON "Users" ("keycloakId")`,
      // Суперадмин + гранулярные права (sequelize.sync колонки не добавляет).
      `ALTER TABLE "Users" ADD COLUMN IF NOT EXISTS "isSuper" BOOLEAN NOT NULL DEFAULT false`,
      `ALTER TABLE "Users" ADD COLUMN IF NOT EXISTS "permissions" JSONB NOT NULL DEFAULT '{}'::jsonb`,
      // Идемпотентный сид: учётка login='admin' — суперадмин.
      `UPDATE "Users" SET "isSuper"=true WHERE "login"='admin'`,
      // Должности вместо ролей: ENUM → строка, прежние роли → должности (LEGACY_ROLES).
      // Повторный старт ничего не меняет: столбец уже строка, старых значений нет.
      `ALTER TABLE "Users" ALTER COLUMN "role" TYPE VARCHAR(40) USING "role"::text`,
      ...Object.entries(LEGACY_ROLES).map(([from, to]) => `UPDATE "Users" SET "role"='${to}' WHERE "role"='${from}'`),
      `DROP TYPE IF EXISTS "enum_Users_role"`,
      // Должности «все РЭС» к РЭС не привязаны.
      `UPDATE "Users" SET "resId"=NULL WHERE "resId" IS NOT NULL AND "role" IN (${Object.entries(ROLES).filter(([, r]) => r.scope === 'all').map(([k]) => `'${k}'`).join(',')})`,
      // Секции шин ТП: привязка ВЛ к секции + индексы. ADD COLUMN IF NOT EXISTS
      // доезжает без DB_ALTER (Postgres-safe). Старые ВЛ остаются с NULL.
      `ALTER TABLE "NetworkStructures" ADD COLUMN IF NOT EXISTS "sectionId" INTEGER`,
      `CREATE INDEX IF NOT EXISTS idx_netstruct_section ON "NetworkStructures" ("sectionId")`,
      // Тип линии (ВЛ/КЛ/КВЛ/ВУ). NOT NULL DEFAULT 'ВЛ' → все существующие
      // записи автоматически получают 'ВЛ', данные не теряются (Postgres-safe).
      `ALTER TABLE "NetworkStructures" ADD COLUMN IF NOT EXISTS "lineType" VARCHAR(8) NOT NULL DEFAULT 'ВЛ'`,
      `CREATE INDEX IF NOT EXISTS idx_tpsection_res ON "TpSections" ("resId")`,
      `CREATE UNIQUE INDEX IF NOT EXISTS idx_tpsection_unique ON "TpSections" ("resId", "tpName", "sectionNumber")`,
      // Случаи перегрузки секции (этап 3)
      `CREATE INDEX IF NOT EXISTS idx_overloadcase_section ON "OverloadCases" ("sectionId")`,
      `CREATE INDEX IF NOT EXISTS idx_overloadcase_res_stage ON "OverloadCases" ("resId", "stage")`,
      // Помесячные пики секции: уникальность (sectionId, year, month) — нужна и для
      // ON CONFLICT «максимум побеждает» в пайплайне профиля/бэкфилле.
      `CREATE UNIQUE INDEX IF NOT EXISTS idx_smp_unique ON "SectionMonthlyPeaks" ("sectionId", "year", "month")`,
      // «Карта опроса» убрана (2026-11): её срез больше не нужен.
      `DROP TABLE IF EXISTS "PolledMeters"`,
      // Вопросы техблоку: открытые по линии и по РЭС.
      `CREATE INDEX IF NOT EXISTS idx_escalation_status ON "Escalations" ("status", "resId")`,
      `CREATE INDEX IF NOT EXISTS idx_escalation_line ON "Escalations" ("networkStructureId")`,
      // Источник ряда и дата обновления профиля на секции (для модалки техучёта)
      `ALTER TABLE "TpSections" ADD COLUMN IF NOT EXISTS "lastProfileSource" VARCHAR(10)`,
      `ALTER TABLE "TpSections" ADD COLUMN IF NOT EXISTS "lastProfileAt" TIMESTAMP WITH TIME ZONE`,
      // Данные АСКУЭ (значки в карточке секции) и снимок на кейсе перегруза.
      `ALTER TABLE "TpSections" ADD COLUMN IF NOT EXISTS "askueData" JSONB`,
      `ALTER TABLE "OverloadCases" ADD COLUMN IF NOT EXISTS "askueData" JSONB`,
      // Дедуп структуры перед уникальным индексом (resId,tpName,vlName): оставляем
      // строку с МИН id (на неё указывают старые входящие FK), но СНАЧАЛА переносим
      // на неё актуальные данные из самой свежей строки группы (МАКС id), затем
      // перевешиваем входящие FK, удаляем дубли, ставим уникальный индекс.
      // Идемпотентно: HAVING count(*)>1 → после чистки update/delete затрагивают 0 строк.
      // IS NOT DISTINCT FROM — чтобы NULL-ключи группировались так же, как в GROUP BY.
      // (1) Перенос данных МАКС id → строку МИН id. startPu/middlePu/endPu/lastUpdate —
      //     значения самой свежей строки как есть; sectionId — первый НЕ-NULL от новых
      //     к старым (COALESCE по группе), чтобы не потерять привязку ВЛ к секции.
      `UPDATE "NetworkStructures" keep SET
         "startPu"    = fresh."startPu",
         "middlePu"   = fresh."middlePu",
         "endPu"      = fresh."endPu",
         "lastUpdate" = fresh."lastUpdate",
         "sectionId"  = fresh."sectionId"
       FROM (
         SELECT
           (array_agg(id ORDER BY id ASC))[1]                                      AS min_id,
           (array_agg("startPu"    ORDER BY id DESC))[1]                           AS "startPu",
           (array_agg("middlePu"   ORDER BY id DESC))[1]                           AS "middlePu",
           (array_agg("endPu"      ORDER BY id DESC))[1]                           AS "endPu",
           (array_agg("lastUpdate" ORDER BY id DESC))[1]                           AS "lastUpdate",
           (array_agg("sectionId"  ORDER BY id DESC) FILTER (WHERE "sectionId" IS NOT NULL))[1] AS "sectionId"
         FROM "NetworkStructures"
         GROUP BY "resId","tpName","vlName"
         HAVING count(*) > 1
       ) fresh
       WHERE keep.id = fresh.min_id`,
      // (2) Перевес входящих FK на строку МИН id.
      `UPDATE "PuStatuses" t SET "networkStructureId" = k.min_id
         FROM "NetworkStructures" ns
         JOIN (SELECT "resId","tpName","vlName", MIN(id) AS min_id FROM "NetworkStructures"
               GROUP BY "resId","tpName","vlName") k
           ON ns."resId" IS NOT DISTINCT FROM k."resId"
          AND ns."tpName" IS NOT DISTINCT FROM k."tpName"
          AND ns."vlName" IS NOT DISTINCT FROM k."vlName"
        WHERE t."networkStructureId" = ns.id AND ns.id <> k.min_id`,
      `UPDATE "Notifications" t SET "networkStructureId" = k.min_id
         FROM "NetworkStructures" ns
         JOIN (SELECT "resId","tpName","vlName", MIN(id) AS min_id FROM "NetworkStructures"
               GROUP BY "resId","tpName","vlName") k
           ON ns."resId" IS NOT DISTINCT FROM k."resId"
          AND ns."tpName" IS NOT DISTINCT FROM k."tpName"
          AND ns."vlName" IS NOT DISTINCT FROM k."vlName"
        WHERE t."networkStructureId" = ns.id AND ns.id <> k.min_id`,
      `UPDATE "CheckHistories" t SET "networkStructureId" = k.min_id
         FROM "NetworkStructures" ns
         JOIN (SELECT "resId","tpName","vlName", MIN(id) AS min_id FROM "NetworkStructures"
               GROUP BY "resId","tpName","vlName") k
           ON ns."resId" IS NOT DISTINCT FROM k."resId"
          AND ns."tpName" IS NOT DISTINCT FROM k."tpName"
          AND ns."vlName" IS NOT DISTINCT FROM k."vlName"
        WHERE t."networkStructureId" = ns.id AND ns.id <> k.min_id`,
      `UPDATE "ProblemVLs" t SET "networkStructureId" = k.min_id
         FROM "NetworkStructures" ns
         JOIN (SELECT "resId","tpName","vlName", MIN(id) AS min_id FROM "NetworkStructures"
               GROUP BY "resId","tpName","vlName") k
           ON ns."resId" IS NOT DISTINCT FROM k."resId"
          AND ns."tpName" IS NOT DISTINCT FROM k."tpName"
          AND ns."vlName" IS NOT DISTINCT FROM k."vlName"
        WHERE t."networkStructureId" = ns.id AND ns.id <> k.min_id`,
      `DELETE FROM "NetworkStructures" ns
         USING (SELECT "resId","tpName","vlName", MIN(id) AS min_id FROM "NetworkStructures"
                GROUP BY "resId","tpName","vlName") k
        WHERE ns."resId" IS NOT DISTINCT FROM k."resId"
          AND ns."tpName" IS NOT DISTINCT FROM k."tpName"
          AND ns."vlName" IS NOT DISTINCT FROM k."vlName"
          AND ns.id <> k.min_id`,
      `CREATE UNIQUE INDEX IF NOT EXISTS idx_netstruct_unique ON "NetworkStructures" ("resId","tpName","vlName")`
    ];

    // Лог дублей структуры ДО чистки (виден в логах Amvera). Дедуп идёт в цикле ниже.
    const dupCountSql = `
      SELECT COALESCE(SUM(cnt - 1), 0)::int AS to_delete, COUNT(*)::int AS groups
      FROM (SELECT count(*) AS cnt FROM "NetworkStructures"
            GROUP BY "resId","tpName","vlName" HAVING count(*) > 1) g`;
    if (sequelize.getDialect() === 'postgres') {
      try {
        const [rows] = await sequelize.query(dupCountSql);
        const r = rows[0] || {};
        console.log(`[dedup NetworkStructures] ДО чистки: групп-дублей ${r.groups || 0}, строк к удалению ${r.to_delete || 0}`);
      } catch (e) { console.warn('[dedup NetworkStructures] count before failed:', e.message); }
    }

    for (const stmt of indexStatements) {
      try {
        await sequelize.query(stmt);
      } catch (idxErr) {
        console.warn('Index creation skipped:', idxErr.message);
      }
    }
    console.log('Performance indexes ensured');

    // Стартовые права должностей (один раз на должность) и матрица в память.
    await seedRolePermissions();

    // Лог дублей структуры ПОСЛЕ чистки — должно быть 0/0 (иначе смотреть warn выше).
    if (sequelize.getDialect() === 'postgres') {
      try {
        const [rows] = await sequelize.query(dupCountSql);
        const r = rows[0] || {};
        console.log(`[dedup NetworkStructures] ПОСЛЕ чистки: групп-дублей ${r.groups || 0}, строк к удалению ${r.to_delete || 0}`);
      } catch (e) { console.warn('[dedup NetworkStructures] count after failed:', e.message); }
    }

    // Backfill помесячных пиков из текущих lastPeakKw/lastPeakAt секций (однократно,
    // идемпотентно через ON CONFLICT «максимум побеждает»). Повторный старт не ломает.
    if (sequelize.getDialect() === 'postgres') {
      try {
        await sequelize.query(`
          INSERT INTO "SectionMonthlyPeaks" ("sectionId","year","month","peakKw","peakAt","source","createdAt","updatedAt")
          SELECT id,
                 EXTRACT(YEAR FROM "lastPeakAt")::int,
                 EXTRACT(MONTH FROM "lastPeakAt")::int,
                 "lastPeakKw", "lastPeakAt", "lastProfileSource", NOW(), NOW()
            FROM "TpSections"
           WHERE "lastPeakKw" IS NOT NULL AND "lastPeakAt" IS NOT NULL
          ON CONFLICT ("sectionId","year","month") DO UPDATE
            SET "peakKw"=EXCLUDED."peakKw", "peakAt"=EXCLUDED."peakAt",
                "source"=EXCLUDED."source", "updatedAt"=NOW()
           WHERE EXCLUDED."peakKw" > "SectionMonthlyPeaks"."peakKw"
        `);
        console.log('SectionMonthlyPeaks backfill ensured');
      } catch (bfErr) {
        console.warn('SectionMonthlyPeaks backfill skipped:', bfErr.message);
      }
    }

    // Новое значение enum уведомлений 'power_overload' (перегруз секции по профилю).
    // sequelize.sync ENUM не расширяет → ALTER TYPE ... ADD VALUE IF NOT EXISTS.
    // Точное имя типа берём из pg_catalog (обычно enum_Notifications_type).
    // Только Postgres; ADD VALUE не выполняется внутри транзакции — sequelize.query
    // работает в autocommit, поэтому ок. Повторный старт не падает (IF NOT EXISTS).
    if (sequelize.getDialect() === 'postgres') {
      try {
        const [typeRows] = await sequelize.query(`
          SELECT t.typname FROM pg_type t
          JOIN pg_attribute a ON a.atttypid = t.oid
          JOIN pg_class c ON c.oid = a.attrelid
          WHERE c.relname = 'Notifications' AND a.attname = 'type'
          LIMIT 1
        `);
        const typeName = typeRows && typeRows[0] && typeRows[0].typname;
        if (typeName) {
          await sequelize.query(
            `ALTER TYPE "${typeName}" ADD VALUE IF NOT EXISTS 'power_overload'`
          );
          console.log(`Enum ${typeName} ensured 'power_overload'`);
        }
      } catch (enumErr) {
        console.warn('Enum power_overload ensure skipped:', enumErr.message);
      }

      // Значение 'profile' в enum колонок fileType (UploadHistories и
      // PuUploadHistories). Без него загрузка type='profile' падала на
      // UploadHistory.create («invalid input value for enum ... "profile"»).
      // Имена enum-типов берём из pg_catalog по колонке fileType каждой таблицы.
      for (const table of ['UploadHistories', 'PuUploadHistories']) {
        try {
          const [rows] = await sequelize.query(`
            SELECT t.typname FROM pg_type t
            JOIN pg_attribute a ON a.atttypid = t.oid
            JOIN pg_class c ON c.oid = a.attrelid
            WHERE c.relname = '${table}' AND a.attname = 'fileType'
            LIMIT 1
          `);
          const typeName = rows && rows[0] && rows[0].typname;
          if (typeName) {
            await sequelize.query(
              `ALTER TYPE "${typeName}" ADD VALUE IF NOT EXISTS 'profile'`
            );
            console.log(`Enum ${typeName} ensured 'profile'`);
          }
        } catch (enumErr) {
          console.warn(`Enum profile ensure skipped for ${table}:`, enumErr.message);
        }
      }
    }

    // Создаем РЭСы если их нет
    const resCount = await ResUnit.count();
    if (resCount === 0) {
      const resList = [
        'Краснополянский РЭС',
        'Адлерский РЭС',
        'Хостинский РЭС',
        'Сочинский РЭС',
        'Дагомысский РЭС',
        'Лазаревский РЭС',
        'Туапсинский РЭС'
      ];
      
      for (const resName of resList) {
        await ResUnit.create({ name: resName });
      }
      console.log('RES units created');
    }
    
    console.log('Database initialization complete');
    
    // ✅ Удаляем СИРИУС из БД если он существует
    try {
      const sirius = await ResUnit.findOne({ where: { name: 'СИРИУС' } });
      if (sirius) {
        const siriusId = sirius.id;
        console.log(`🧹 Found СИРИУС (id=${siriusId}), removing...`);
        
        // Находим ID уведомлений СИРИУС для удаления NotificationRead
        const siriusNotifs = await Notification.findAll({ 
          where: { resId: siriusId }, 
          attributes: ['id'] 
        });
        const siriusNotifIds = siriusNotifs.map(n => n.id);
        
        // Удаляем связанные данные
        if (siriusNotifIds.length > 0) {
          await NotificationRead.destroy({ where: { notificationId: { [Op.in]: siriusNotifIds } } });
        }
        const deletedNotifs = await Notification.destroy({ where: { resId: siriusId } });
        const deletedHistory = await CheckHistory.destroy({ where: { resId: siriusId } });
        const deletedProblemVL = await ProblemVL.destroy({ where: { resId: siriusId } });
        
        // Удаляем PuStatus привязанные к структурам СИРИУС
        const siriusStructures = await NetworkStructure.findAll({ where: { resId: siriusId }, attributes: ['id'] });
        const siriusStructureIds = siriusStructures.map(s => s.id);
        if (siriusStructureIds.length > 0) {
          await PuStatus.destroy({ where: { networkStructureId: { [Op.in]: siriusStructureIds } } });
        }
        
        const deletedStructures = await NetworkStructure.destroy({ where: { resId: siriusId } });
        const deletedUsers = await User.destroy({ where: { resId: siriusId } });
        await sirius.destroy();
        
        console.log(`✅ СИРИУС removed: ${deletedNotifs} notifs, ${deletedStructures} structures, ${deletedHistory} history, ${deletedProblemVL} problemVL, ${deletedUsers} users`);
      }
    } catch (err) {
      console.error('Error removing СИРИУС:', err.message);
    }
    
    // Создаем админа если его нет
    const adminCount = await User.count({ where: { role: 'admin' } });
    if (adminCount === 0) {
      await User.create({
        fio: 'Администратор',
        login: 'admin',
        password: 'admin123',
        role: 'admin',
        email: 'admin@res.ru'
      });
      console.log('Admin user created');
    }

    // id системной учётки для fromUserId системных уведомлений (суперадмин).
    // Если суперадмин ещё не помечен (самый первый старт) — null (поле nullable).
    const superUser = await User.findOne({ where: { isSuper: true }, attributes: ['id'] });
    SYSTEM_USER_ID = superUser?.id ?? null;

  } catch (error) {
    console.error('Database initialization error:', error);
    process.exit(1);
  }
}


// API для получения документов
app.get('/api/documents/list', authenticateToken, async (req, res) => {
  try {
    const { resId } = req.query; // ДОБАВИТЬ
    let whereClause = {};
    
    // Фильтрация по РЭС
    if (req.user.allRes && resId) {
      whereClause.resId = parseInt(resId);
    } else if (!req.user.allRes) {
      whereClause.resId = req.user.resId;
    }
    
    const documents = await CheckHistory.findAll({
      where: whereClause,
      include: [ResUnit],
      order: [['workCompletedDate', 'DESC']]
    });
    
    // Фильтруем только записи с файлами
    const documentsWithFiles = documents.filter(doc => 
      doc.attachments && 
      Array.isArray(doc.attachments) && 
      doc.attachments.length > 0
    );
    
    const formattedDocs = documentsWithFiles.map(doc => ({
      id: doc.id,
      tpName: doc.tpName,
      vlName: doc.vlName,
      puNumber: doc.puNumber,
      uploadedBy: doc.ResUnit?.name || 'Неизвестно',
      workCompletedDate: doc.workCompletedDate,
      resComment: doc.resComment,
      status: doc.status,
      attachments: doc.attachments || []
    }));
    
    res.json(formattedDocs);
  } catch (error) {
    console.error('Get documents error:', error);
    res.status(500).json({ error: error.message });
  }
});

// API для управления файлами 
app.get('/api/admin/files', 
  authenticateToken, 
  requirePerm(['files_manage', 'settings_view']), 
  async (req, res) => {
    try {
      // Получаем ВСЕ записи без проблемного where
      const records = await CheckHistory.findAll({
        include: [ResUnit],
        order: [['createdAt', 'DESC']]
      });
      
      // Фильтруем в JavaScript
      const recordsWithFiles = records.filter(record => 
        record.attachments && 
        Array.isArray(record.attachments) && 
        record.attachments.length > 0
      );
      
      // Собираем все файлы С СТАТУСОМ ЗАПИСИ
      const files = [];
      recordsWithFiles.forEach(record => {
        if (record.attachments && Array.isArray(record.attachments)) {
          record.attachments.forEach(file => {
            files.push({
              ...file,
              recordId: record.id,
              resName: record.ResUnit?.name,
              tpName: record.tpName,
              vlName: record.vlName,  // ✅ ДОБАВЛЕНО
              puNumber: record.puNumber,
              uploadDate: record.workCompletedDate || record.createdAt,
              status: record.status,  // ✅ ДОБАВЛЕНО - статус записи
              resComment: record.resComment  // ✅ ДОБАВЛЕНО - для tooltip
            });
          });
        }
      });
      
      res.json({ files, total: files.length });
    } catch (error) {
      console.error('Error in /api/admin/files:', error);
      res.status(500).json({ error: error.message });
    }
  });

// API для удаления файла из документа
app.delete('/api/documents/record/:recordId',
  authenticateToken,
  requirePerm('files_manage'),
  async (req, res) => {
    try {
      const { password } = req.body;
      const { recordId } = req.params;
      
      if (password !== DELETE_PASSWORD) {
        return res.status(403).json({ error: 'Неверный пароль' });
      }
      
      const record = await CheckHistory.findByPk(recordId);
      if (!record) {
        return res.status(404).json({ error: 'Запись не найдена' });
      }
      
      // Собираем public_id и удаляем через purgeCloudinary — она различает
      // raw/image по .pdf (иначе PDF оставались сиротами в Cloudinary).
      const fileIds = Array.isArray(record.attachments)
        ? record.attachments.filter(f => f && f.public_id).map(f => f.public_id)
        : [];

      // Удаляем запись из БД
      await record.destroy();

      const cl = await purgeCloudinary(fileIds);

      res.json({
        success: true,
        message: 'Запись и все связанные файлы удалены',
        cloudinaryErrors: cl.errors
      });

    } catch (error) {
      console.error('Delete record error:', error);
      res.status(500).json({ error: error.message });
    }
});

// ИСПРАВЛЕННЫЙ эндпоинт для управления файлами в настройках
 
app.delete('/api/admin/files/:public_id',
  authenticateToken,
  requirePerm('files_manage'),
  async (req, res) => {
    try {
      const { password } = req.body;
      const publicId = decodeURIComponent(req.params.public_id);
      
      if (password !== DELETE_PASSWORD) {
        return res.status(403).json({ error: 'Неверный пароль' });
      }
      
      console.log(`🗑️ Attempting to delete file: ${publicId}`);
      
      // ИСПРАВЛЕНО: определяем resource_type по расширению файла
      const isPdf = publicId.toLowerCase().endsWith('.pdf');
      const resourceType = isPdf ? 'raw' : 'image';
      
      console.log(`Resource type: ${resourceType}`);
      
      // Удаляем из Cloudinary с правильным типом
      try {
        await cloudinary.uploader.destroy(publicId, {
          resource_type: resourceType,
          invalidate: true  // очищаем CDN кеш
        });
        console.log(`✅ File deleted from Cloudinary: ${publicId}`);
      } catch (cloudinaryError) {
        console.error('⚠️ Cloudinary deletion warning:', cloudinaryError.message);
        // Продолжаем даже если файл уже удален
      }
      
      // Находим все записи в CheckHistory с этим файлом
      const records = await CheckHistory.findAll();
      let updatedCount = 0;
      
      for (const record of records) {
        if (record.attachments && Array.isArray(record.attachments)) {
          const originalLength = record.attachments.length;
          const newAttachments = record.attachments.filter(
            file => file.public_id !== publicId
          );
          
          if (newAttachments.length < originalLength) {
            await record.update({ attachments: newAttachments });
            updatedCount++;
            console.log(`✅ Updated record ${record.id}`);
          }
        }
      }
      
      console.log(`📊 Total records updated: ${updatedCount}`);
      
      res.json({ 
        success: true, 
        message: 'Файл успешно удален',
        updatedRecords: updatedCount
      });
      
    } catch (error) {
      console.error('❌ Delete file error:', error);
      res.status(500).json({ 
        error: error.message,
        details: 'Ошибка при удалении файла. Проверьте консоль сервера.'
      });
    }
});


// Запуск сервера
// ==================== Почтовый приёмник (Пирамида → профиль мощности) ==========
// За флагом MAIL_INTAKE (default OFF): при OFF ничего не запускается, поведение
// прежнее. Читает IMAP (Яндекс), по разрешённым отправителям обрабатывает
// вложение-профиль, отвечает сводкой и раскладывает письма по папкам
// Processed / Errors / Rejected.
let intakeRunning = false;

// Создать/найти папку с учётом namespace Яндекса; fallback на плоское имя и
// на подпапку INBOX. Идемпотентно (ALREADYEXISTS не роняет).
async function ensureImapFolder(client, name) {
  let prefix = '', delim = '/';
  try {
    const ns = client.namespace;
    const personal = ns && ns.personal && ns.personal[0];
    if (personal) { prefix = personal.prefix || ''; delim = personal.delimiter || '/'; }
  } catch (e) {}
  const candidates = [...new Set([prefix + name, name, 'INBOX' + delim + name])];
  for (const p of candidates) {
    try { await client.mailboxCreate(p); return p; }
    catch (e) {
      const m = String(e && (e.responseText || e.message || e.code) || '');
      if (/alreadyexists|already exists/i.test(m)) return p;
    }
  }
  return candidates[0];
}

async function runMailIntakeOnce() {
  if (intakeRunning) return;
  intakeRunning = true;
  let ImapFlow, simpleParser;
  try {
    ({ ImapFlow } = require('imapflow'));
    ({ simpleParser } = require('mailparser'));
  } catch (e) {
    console.error('[INTAKE] Нет модулей imapflow/mailparser:', e.message);
    intakeRunning = false; return;
  }
  const client = new ImapFlow({
    host: process.env.MAIL_IMAP_HOST || 'imap.yandex.ru',
    port: Number(process.env.MAIL_IMAP_PORT) || 993,
    secure: true,
    auth: { user: process.env.MAIL_USER, pass: process.env.MAIL_PASS },
    logger: false
  });
  try {
    await client.connect();

    // Разрешённые отправители: env MAIL_INTAKE_ALLOWED (через запятую) либо все
    // email пользователей из БД.
    const envAllowed = (process.env.MAIL_INTAKE_ALLOWED || '')
      .split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
    const allowed = new Set(envAllowed);
    const allowSource = envAllowed.length ? 'MAIL_INTAKE_ALLOWED' : 'email пользователей БД';
    if (allowed.size === 0) {
      const users = await User.findAll({ attributes: ['email'] });
      users.forEach(u => { if (u.email) allowed.add(String(u.email).trim().toLowerCase()); });
    }
    console.log(`[INTAKE] Подключено. Разрешённых отправителей: ${allowed.size} (источник: ${allowSource})` +
      (allowed.size ? ` — ${[...allowed].slice(0, 8).join(', ')}` : ' — ПУСТО: все письма будут отклонены!'));

    const fProcessed = await ensureImapFolder(client, process.env.MAIL_FOLDER_PROCESSED || 'Processed');
    const fErrors = await ensureImapFolder(client, process.env.MAIL_FOLDER_ERRORS || 'Errors');
    const fRejected = await ensureImapFolder(client, process.env.MAIL_FOLDER_REJECTED || 'Rejected');

    // Собираем непрочитанные письма INBOX.
    const jobs = [];
    const lock = await client.getMailboxLock('INBOX');
    try {
      for await (const msg of client.fetch({ seen: false }, { uid: true, source: true })) {
        jobs.push({ uid: msg.uid, source: msg.source });
      }
    } finally { lock.release(); }

    console.log(`[INTAKE] Непрочитанных писем в INBOX: ${jobs.length}`);
    for (const job of jobs) {
      let parsed = null;
      try { parsed = await simpleParser(job.source); }
      catch (e) { console.error('[INTAKE] parse error:', e.message); }
      const sender = parsed?.from?.value?.[0]?.address
        ? String(parsed.from.value[0].address).trim().toLowerCase() : '';
      const subject = parsed?.subject || '(без темы)';

      let target = fErrors;
      let reply = null;

      if (!sender || !allowed.has(sender)) {
        target = fRejected; // без ответа
        console.log(`[INTAKE] Отклонено (адрес не разрешён): ${sender || '—'}`);
      } else {
        const att = (parsed.attachments || []).find(a =>
          /\.(xlsx|xls)$/i.test(a.filename || '') || /spreadsheet|excel/i.test(a.contentType || ''));
        if (!att) {
          target = fErrors;
          reply = { subject: 'Ошибка обработки: нет вложения',
            text: `Здравствуйте!\n\nВ письме «${subject}» не найдено вложение с профилем мощности (.xlsx). Пришлите выгрузку «Профиль мощности» из Пирамиды файлом .xlsx.\n\n— РЭС-менеджмент (автоответ).` };
          console.log(`[INTAKE] Нет вложения от ${sender}`);
        } else {
          const tmpDir = path.join(process.cwd(), 'uploads');
          try { fs.mkdirSync(tmpDir, { recursive: true }); } catch (e) {}
          const tmpPath = path.join(tmpDir, `intake_${Date.now()}_${Math.random().toString(36).slice(2)}.xlsx`);
          try {
            fs.writeFileSync(tmpPath, att.content);
            const result = await processProfileFile(tmpPath, null);
            try { fs.unlinkSync(tmpPath); } catch (e) {}
            if (!result.success) {
              target = fErrors;
              reply = { subject: 'Ошибка обработки профиля',
                text: `Здравствуйте!\n\nНе удалось обработать профиль из письма «${subject}»: ${result.error}\n\n— РЭС-менеджмент (автоответ).` };
            } else {
              target = fProcessed;
              const un = result.unmatched.length
                ? `\nНе привязано к структуре ПУ: ${result.unmatched.length} (${result.unmatched.join(', ')})` : '';
              const matched = (result.details || []).filter(d => d.matched);
              const matchedTxt = matched.length
                ? `\n\nОбновлены секции:\n` + matched.map(d =>
                    `• ПУ ${d.puNumber} (Ктт/Ктн ${d.ktRaw || '—'}, kt ${d.kt}) → ${d.tpSection}: пик-пром ${d.peakRaw} × ${d.kt} = ${d.peakKw} кВт${d.decision === 'overload' ? ' (ПЕРЕГРУЗ)' : ''}`
                  ).join('\n')
                : '';
              reply = { subject: 'Профиль мощности обработан',
                text: `Здравствуйте!\n\nПрофиль из письма «${subject}» обработан.\nСекций обновлено: ${result.sectionsUpdated}\nПерегрузов: ${result.overloadCount}${un}${matchedTxt}\n\n— РЭС-менеджмент (автоответ).` };
              console.log(`[INTAKE] Обработано от ${sender}: секций ${result.sectionsUpdated}, перегрузов ${result.overloadCount}`);
            }
          } catch (e) {
            try { fs.unlinkSync(tmpPath); } catch (e2) {}
            target = fErrors;
            reply = { subject: 'Ошибка обработки профиля',
              text: `Здравствуйте!\n\nОшибка обработки профиля из письма «${subject}»: ${e.message}\n\n— РЭС-менеджмент (автоответ).` };
            console.error('[INTAKE] processing error:', e.message);
          }
        }
      }

      if (reply && sender) {
        try { await sendMailAs(sender, reply.subject, reply.text); }
        catch (e) { console.error('[INTAKE] reply send error:', e.message); }
      }

      // Помечаем прочитанным и переносим (если move упадёт — хотя бы не зациклит).
      const mlock = await client.getMailboxLock('INBOX');
      try {
        try { await client.messageFlagsAdd(String(job.uid), ['\\Seen'], { uid: true }); } catch (e) {}
        try { await client.messageMove(String(job.uid), target, { uid: true }); }
        catch (e) { console.error('[INTAKE] move error:', e.message); }
      } finally { mlock.release(); }
    }
  } catch (e) {
    console.error('[INTAKE] error:', e.message);
  } finally {
    try { await client.logout(); } catch (e) {}
    intakeRunning = false;
  }
}

function startMailIntake() {
  if (process.env.MAIL_INTAKE !== 'true') return;
  if (!process.env.MAIL_USER || !process.env.MAIL_PASS) {
    console.warn('[INTAKE] MAIL_INTAKE=true, но MAIL_USER/MAIL_PASS не заданы — приёмник не запущен.');
    return;
  }
  // Период опроса: MAIL_POLL_SECONDS (в секундах) либо MAIL_INTAKE_INTERVAL_MS (в мс).
  const pollSec = Number(process.env.MAIL_POLL_SECONDS);
  const interval = (Number.isFinite(pollSec) && pollSec > 0)
    ? pollSec * 1000
    : (Number(process.env.MAIL_INTAKE_INTERVAL_MS) || 60000);
  console.log(`[INTAKE] Почтовый приёмник включён (Яндекс IMAP, каждые ${Math.round(interval / 1000)} с).`);
  runMailIntakeOnce().catch(e => console.error('[INTAKE] first run:', e.message));
  setInterval(() => runMailIntakeOnce().catch(e => console.error('[INTAKE] tick:', e.message)), interval);
}

initializeDatabase().then(() => {
  app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
    console.log(`Environment: ${process.env.NODE_ENV || 'development'}`);
    console.log('Features enabled:');
    console.log('- User Management ✓');
    console.log('- Phase Detection ✓'); 
    console.log('- Auto Updates ✓');
    console.log('- Auto Hide Notifications ✓');
    startMailIntake();
  });
});

// API для получения истории загрузок по конкретному ПУ
// Текущий статус ПУ (для модалки «Структуры сети», вызываемой из глубокого
// анализа) — фактическая ошибка по состоянию: status + errorDetails + lastCheck.
app.get('/api/pu-status/:puNumber',
  authenticateToken,
  async (req, res) => {
    try {
      const { puNumber } = req.params;
      const st = await PuStatus.findOne({ where: { puNumber } });
      if (!st) return res.json(null);
      res.json({
        puNumber: st.puNumber, position: st.position, status: st.status,
        errorDetails: st.errorDetails, lastCheck: st.lastCheck,
        networkStructureId: st.networkStructureId,
      });
    } catch (error) {
      console.error('Get pu-status error:', error);
      res.status(500).json({ error: error.message });
    }
});

app.get('/api/history/uploads/:puNumber',
  authenticateToken,
  async (req, res) => {
    try {
      const { puNumber } = req.params;

      const uploads = await PuUploadHistory.findAll({
        where: { puNumber },
        include: [{
          model: User,
          attributes: ['fio', 'login']
        }],
        order: [['uploadedAt', 'DESC']],
        limit: 20
      });
      
      res.json(uploads);
    } catch (error) {
      console.error('Get upload history error:', error);
      res.status(500).json({ error: error.message });
    }
});

// API для получения истории проверок по конкретному ПУ
app.get('/api/history/checks/:puNumber', 
  authenticateToken, 
  async (req, res) => {
    try {
      const { puNumber } = req.params;
      
      const checks = await CheckHistory.findAll({
        where: { puNumber },
        include: [ResUnit],
        order: [['createdAt', 'DESC']]
      });
      
      res.json(checks);
    } catch (error) {
      console.error('Get check history error:', error);
      res.status(500).json({ error: error.message });
    }
});

// API для получения всей истории загрузок с фильтрами
app.get('/api/history/uploads', 
  authenticateToken, 
  async (req, res) => {
    try {
      const { resId, tpName, puNumber, dateFrom, dateTo, fileType, status, page = 1, limit = 50 } = req.query;
      
      // Сначала получаем все PU для нужного РЭС
      let structureWhere = {};
      
      // Фильтр по РЭС
      if (req.user.allRes && resId) {
        structureWhere.resId = resId;
      } else if (!req.user.allRes) {
        structureWhere.resId = req.user.resId;
      }
      
      // Фильтр по ТП если указан
      if (tpName) {
        structureWhere.tpName = { [Op.like]: `%${tpName}%` };
      }
      
      // Получаем все номера ПУ из структуры с нужными фильтрами
      const structures = await NetworkStructure.findAll({
        where: structureWhere,
        attributes: ['startPu', 'middlePu', 'endPu', 'tpName', 'vlName', 'resId'],
        include: [ResUnit]
      });
      
      // Собираем все номера ПУ
      const puNumbers = new Set();
      const puToStructureMap = {};
      
      structures.forEach(s => {
        if (s.startPu) {
          puNumbers.add(s.startPu);
          puToStructureMap[s.startPu] = s;
        }
        if (s.middlePu) {
          puNumbers.add(s.middlePu);
          puToStructureMap[s.middlePu] = s;
        }
        if (s.endPu) {
          puNumbers.add(s.endPu);
          puToStructureMap[s.endPu] = s;
        }
      });
      
      // Если нет ПУ для этого РЭС - возвращаем пустой результат
      if (puNumbers.size === 0) {
        return res.json({
          uploads: [],
          total: 0,
          page: parseInt(page),
          totalPages: 0
        });
      }
      
      // Теперь ищем загрузки только для этих ПУ
      let uploadWhere = {
        puNumber: Array.from(puNumbers)
      };
      
      // Дополнительные фильтры
      if (puNumber) {
        uploadWhere.puNumber = { 
          [Op.and]: [
            { [Op.in]: Array.from(puNumbers) },
            { [Op.like]: `%${puNumber}%` }
          ]
        };
      }
      if (fileType) uploadWhere.fileType = fileType;
      if (status) uploadWhere.uploadStatus = status;
      
      if (dateFrom || dateTo) {
        uploadWhere.uploadedAt = {};
        if (dateFrom) uploadWhere.uploadedAt[Op.gte] = new Date(dateFrom);
        if (dateTo) {
          const endDate = new Date(dateTo);
          endDate.setHours(23, 59, 59, 999);
          uploadWhere.uploadedAt[Op.lte] = endDate;
        }
      }
      
      const offset = (page - 1) * limit;
      
      const { count, rows } = await PuUploadHistory.findAndCountAll({
        where: uploadWhere,
        include: [{
          model: User,
          attributes: ['fio', 'login', 'resId'],
          include: [ResUnit]
        }],
        order: [['uploadedAt', 'DESC']],
        limit: parseInt(limit),
        offset
      });
      
      // Добавляем информацию о структуре
      const uploadsWithStructure = rows.map(upload => {
        const structure = puToStructureMap[upload.puNumber];
        return {
          ...upload.toJSON(),
          tpName: structure?.tpName,
          vlName: structure?.vlName,
          resName: structure?.ResUnit?.name,
          resId: structure?.resId
        };
      });
      
      res.json({
        uploads: uploadsWithStructure,
        total: count,
        page: parseInt(page),
        totalPages: Math.ceil(count / limit)
      });
      
    } catch (error) {
      console.error('Get all uploads history error:', error);
      res.status(500).json({ error: error.message });
    }
});

// API для получения всей истории проверок
app.get('/api/history/checks', 
  authenticateToken, 
  async (req, res) => {
    try {
      const { resId, tpName, puNumber, dateFrom, dateTo, status, page = 1, limit = 50 } = req.query;
      
      let whereClause = {};
      
      // Фильтры
      if (status) whereClause.status = status;
      if (puNumber) whereClause.puNumber = { [Op.like]: `%${puNumber}%` };
      if (tpName) whereClause.tpName = { [Op.like]: `%${tpName}%` };
      
      // ИСПРАВЛЕНО: Фильтр по РЭС
      if (req.user.allRes && resId) {
        whereClause.resId = resId;
      } else if (!req.user.allRes) {
        whereClause.resId = req.user.resId;
      }
      
      if (dateFrom || dateTo) {
        whereClause.createdAt = {};
        if (dateFrom) whereClause.createdAt[Op.gte] = new Date(dateFrom);
        if (dateTo) {
          const endDate = new Date(dateTo);
          endDate.setHours(23, 59, 59, 999);
          whereClause.createdAt[Op.lte] = endDate;
        }
      }
      
      const offset = (page - 1) * limit;
      
      const { count, rows } = await CheckHistory.findAndCountAll({
        where: whereClause,
        include: [ResUnit],
        order: [['createdAt', 'DESC']],
        limit: parseInt(limit),
        offset
      });
      
      res.json({
        checks: rows,
        total: count,
        page: parseInt(page),
        totalPages: Math.ceil(count / limit)
      });
      
    } catch (error) {
      console.error('Get all checks history error:', error);
      res.status(500).json({ error: error.message });
    }
});
// Очистка истории по конкретному ПУ
app.delete('/api/history/clear-pu/:puNumber',
  authenticateToken,
  requirePerm('checks_delete'),
  async (req, res) => {
    const { password } = req.body;
    const { puNumber } = req.params;

    if (password !== DELETE_PASSWORD) {
      return res.status(403).json({ error: 'Неверный пароль' });
    }

    const transaction = await sequelize.transaction();

    try {
      // Удаляем из PuUploadHistory
      const uploadsDeleted = await PuUploadHistory.destroy({
        where: { puNumber },
        transaction
      });
      
      // Удаляем из CheckHistory
      const checksDeleted = await CheckHistory.destroy({
        where: { puNumber },
        transaction
      });
      
      // Удаляем из PuStatus
      await PuStatus.update(
        { 
          status: 'not_checked',
          errorDetails: null,
          lastCheck: null
        },
        { 
          where: { puNumber },
          transaction
        }
      );
      
      await transaction.commit();
      
      res.json({
        success: true,
        message: `История ПУ ${puNumber} очищена`,
        deleted: {
          uploads: uploadsDeleted,
          checks: checksDeleted
        }
      });
      
    } catch (error) {
      await transaction.rollback();
      res.status(500).json({ error: error.message });
    }
});

// Очистка истории по ТП
app.post('/api/history/clear-tp',
  authenticateToken,
  requirePerm('checks_delete'),
  async (req, res) => {
    const { password, tpNames, resId } = req.body;

    if (password !== DELETE_PASSWORD) {
      return res.status(403).json({ error: 'Неверный пароль' });
    }

    // Находим все ПУ для выбранных ТП
    const structures = await NetworkStructure.findAll({
      where: {
        tpName: tpNames,
        resId: resId
      }
    });

    const puNumbers = [];
    structures.forEach(s => {
      if (s.startPu) puNumbers.push(s.startPu);
      if (s.middlePu) puNumbers.push(s.middlePu);
      if (s.endPu) puNumbers.push(s.endPu);
    });

    if (puNumbers.length === 0) {
      return res.json({
        success: true,
        message: 'Нет ПУ для очистки',
        deleted: { uploads: 0, checks: 0 }
      });
    }

    const transaction = await sequelize.transaction();

    try {
      // Удаляем историю
      const uploadsDeleted = await PuUploadHistory.destroy({
        where: { puNumber: puNumbers },
        transaction
      });
      
      const checksDeleted = await CheckHistory.destroy({
        where: { puNumber: puNumbers },
        transaction
      });
      
      // Сбрасываем статусы
      await PuStatus.update(
        { 
          status: 'not_checked',
          errorDetails: null,
          lastCheck: null
        },
        { 
          where: { puNumber: puNumbers },
          transaction
        }
      );
      
      await transaction.commit();
      
      res.json({
        success: true,
        message: `История для ${tpNames.length} ТП очищена`,
        deleted: {
          uploads: uploadsDeleted,
          checks: checksDeleted,
          puCount: puNumbers.length
        }
      });
      
    } catch (error) {
      await transaction.rollback();
      res.status(500).json({ error: error.message });
    }
});

// Очистка всей истории
app.delete('/api/history/clear-all',
  authenticateToken,
  requirePerm('checks_delete'),
  async (req, res) => {
    const { password } = req.body;

    if (password !== DELETE_PASSWORD) {
      return res.status(403).json({ error: 'Неверный пароль' });
    }

    const transaction = await sequelize.transaction();

    try {
      // Удаляем всю историю
      const uploadsDeleted = await PuUploadHistory.destroy({
        where: {},
        transaction
      });
      
      const checksDeleted = await CheckHistory.destroy({
        where: {},
        transaction
      });
      
      // Сбрасываем все статусы
      await PuStatus.update(
        { 
          status: 'not_checked',
          errorDetails: null,
          lastCheck: null
        },
        { 
          where: {},
          transaction
        }
      );
      
      await transaction.commit();
      
      res.json({
        success: true,
        message: 'Вся история очищена',
        deleted: {
          uploads: uploadsDeleted,
          checks: checksDeleted
        }
      });
      
    } catch (error) {
      await transaction.rollback();
      res.status(500).json({ error: error.message });
    }
});

// API для детального отчета аналитики
app.get('/api/analytics/detailed', 
  authenticateToken, 
  async (req, res) => {
    try {
      const { dateFrom, dateTo } = req.query;
      
      // Условие по РЭС
      let resCondition = {};
      if (!req.user.allRes) {
        resCondition.id = req.user.resId;
      }
      
      // Получаем РЭСы
      const resList = await ResUnit.findAll({
        where: resCondition,
        order: [['name', 'ASC']]
      });

      // Условие по дате
      let dateCondition = {};
      if (dateFrom || dateTo) {
        dateCondition.uploadedAt = {};
        if (dateFrom) dateCondition.uploadedAt[Op.gte] = new Date(dateFrom);
        if (dateTo) {
          const endDate = new Date(dateTo);
          endDate.setHours(23, 59, 59, 999);
          dateCondition.uploadedAt[Op.lte] = endDate;
        }
      }

      const detailedData = [];

      // Для каждого РЭС
      for (const res of resList) {
        const structures = await NetworkStructure.findAll({
          where: { resId: res.id },
          order: [['tpName', 'ASC'], ['vlName', 'ASC']]
        });

        // Собираем все номера ПУ для этого РЭС
        const allPuNumbers = new Set();
        structures.forEach(s => {
          if (s.startPu) allPuNumbers.add(s.startPu);
          if (s.middlePu) allPuNumbers.add(s.middlePu);
          if (s.endPu) allPuNumbers.add(s.endPu);
        });

        // Получаем историю загрузок для всех ПУ в периоде
        const uploads = await PuUploadHistory.findAll({
          where: {
            puNumber: Array.from(allPuNumbers),
            ...dateCondition
          },
          include: [{
            model: User,
            attributes: ['fio']
          }],
          order: [['uploadedAt', 'DESC']]
        });

        // Создаем Map для быстрого доступа к последней загрузке каждого ПУ
        const latestUploads = new Map();
        uploads.forEach(upload => {
          if (!latestUploads.has(upload.puNumber)) {
            latestUploads.set(upload.puNumber, upload);
          }
        });

        // Обрабатываем каждую ВЛ
        for (const structure of structures) {
          const puPositions = [
            { pu: structure.startPu, position: 'start' },
            { pu: structure.middlePu, position: 'middle' },
            { pu: structure.endPu, position: 'end' }
          ];

          // Собираем информацию о ПУ
          const puData = {};
          let checkedCount = 0;

          for (const { pu, position } of puPositions) {
            if (pu) {
              const upload = latestUploads.get(pu);
              
              puData[position] = {
                number: pu,
                status: upload ? (upload.hasErrors ? 'Ошибка ✗' : 'Проверен ✓') : 'Не проверен',
                error: upload && upload.hasErrors ? upload.errorSummary : '—',
                uploadedBy: upload ? upload.User?.fio || 'Неизвестно' : '—',
                uploadDate: upload ? new Date(upload.uploadedAt).toLocaleString('ru-RU') : '—'
              };
              
              if (upload) checkedCount++;
            } else {
              puData[position] = {
                number: '—',
                status: 'Пусто',
                error: '—',
                uploadedBy: '—',
                uploadDate: '—'
              };
            }
          }

          // Определяем статус ВЛ
          const totalPu = [structure.startPu, structure.middlePu, structure.endPu]
            .filter(Boolean).length;
          
          let vlStatus;
          if (checkedCount === 0) {
            vlStatus = '❌ Не проверена';
          } else if (checkedCount === totalPu) {
            vlStatus = '✅ Проверена';
          } else {
            vlStatus = '⚠️ Частично проверена';
          }

          // Добавляем строку в отчет
          detailedData.push({
            resName: res.name,
            tpName: structure.tpName,
            vlName: structure.vlName,
            vlStatus,
            checkedPuCount: checkedCount,
            totalPuCount: totalPu,
            startPu: puData.start,
            middlePu: puData.middle,
            endPu: puData.end
          });
        }
      }

      res.json({
        success: true,
        data: detailedData,
        period: {
          from: dateFrom,
          to: dateTo
        },
        totalRows: detailedData.length
      });

    } catch (error) {
      console.error('Detailed analytics error:', error);
      res.status(500).json({ error: error.message });
    }
});

// =====================================================
// API ДЛЯ ДИАГНОСТИКИ ДАННЫХ (НОВОЕ)
// =====================================================

// Получить полную диагностику по РЭС
app.get('/api/admin/diagnose/:resId',
  authenticateToken,
  requirePerm('db_tools'),
  async (req, res) => {
    try {
      const resId = parseInt(req.params.resId);
      
      // 1. Структура сети
      const structures = await NetworkStructure.findAll({
        where: { resId },
        include: [
          ResUnit,
          {
            model: PuStatus,
            required: false
          }
        ],
        order: [['tpName', 'ASC'], ['vlName', 'ASC']]
      });
      
      // 2. Уведомления
      const notifications = await Notification.findAll({
        where: {
          networkStructureId: {
            [Op.in]: structures.map(s => s.id)
          }
        },
        include: [
          {
            model: NetworkStructure,
            include: [ResUnit]
          },
          { model: User, as: 'fromUser', attributes: ['id', 'fio', 'role'] },
          { model: User, as: 'toUser', attributes: ['id', 'fio', 'role'] }
        ],
        order: [['createdAt', 'DESC']]
      });
      
      // 3. Находим несоответствия
      const mismatches = [];
      
      notifications.forEach(notif => {
        if (notif.NetworkStructure) {
          const structureResId = notif.NetworkStructure.resId;
          const notifResId = notif.resId;
          
          if (structureResId !== notifResId) {
            mismatches.push({
              notificationId: notif.id,
              type: notif.type,
              message: notif.message,
              notifResId: notifResId,
              structureResId: structureResId,
              notifResName: notif.ResUnit?.name,
              structureResName: notif.NetworkStructure.ResUnit?.name,
              tpName: notif.NetworkStructure.tpName,
              vlName: notif.NetworkStructure.vlName,
              createdAt: notif.createdAt
            });
          }
        }
      });
      
      // ✅ 4. НОВОЕ: Проверка неактуальных уведомлений
      const staleNotifications = [];
      
      const errorNotifications = await Notification.findAll({
        where: {
          type: ['error', 'pending_askue']
        }
      });
      
      for (const notif of errorNotifications) {
        try {
          const data = JSON.parse(notif.message);
          const puNumber = data.puNumber;
          
          if (!puNumber) continue;
          
          // Проверяем статус ПУ
          const puStatus = await PuStatus.findOne({
            where: { puNumber }
          });
          
          if (puStatus) {
            // Если ПУ зеленый, а уведомление висит - это устаревшее
            if (puStatus.status === 'checked_ok') {
              staleNotifications.push({
                notificationId: notif.id,
                type: notif.type,
                puNumber,
                tpName: data.tpName,
                vlName: data.vlName,
                resName: data.resName,
                currentStatus: 'checked_ok',
                notifCreated: notif.createdAt,
                lastCheck: puStatus.lastCheck,
                reason: 'ПУ уже проверен без ошибок'
              });
            }
          } else {
            // ПУ вообще не существует
            staleNotifications.push({
              notificationId: notif.id,
              type: notif.type,
              puNumber,
              tpName: data.tpName,
              vlName: data.vlName,
              resName: data.resName,
              currentStatus: 'not_found',
              notifCreated: notif.createdAt,
              reason: 'ПУ не найден в структуре'
            });
          }
        } catch (e) {
          console.error('Error checking notification:', e);
        }
      }
      
      // 5. Формируем список проблем
      const issues = [];
      
      // Несоответствия resId
      if (mismatches.length > 0) {
        issues.push({
          type: 'resid_mismatch',
          severity: 'error',
          count: mismatches.length,
          description: 'Несоответствия resId между уведомлениями и структурой',
          items: mismatches
        });
      }
      
      // ✅ НОВОЕ: Неактуальные уведомления
      if (staleNotifications.length > 0) {
        issues.push({
          type: 'stale_notifications',
          severity: 'warning',
          count: staleNotifications.length,
          description: 'Найдены неактуальные уведомления',
          items: staleNotifications
        });
      }
      
      // 6. Итоговая статистика
      const stats = {
        totalStructures: structures.length,
        totalNotifications: notifications.length,
        mismatches: mismatches.length,
        staleNotifications: staleNotifications.length, // ✅ НОВОЕ
        notificationsByType: {
          error: notifications.filter(n => n.type === 'error').length,
          pending_askue: notifications.filter(n => n.type === 'pending_askue').length,
          success: notifications.filter(n => n.type === 'success').length,
          problem_vl: notifications.filter(n => n.type === 'problem_vl').length
        }
      };
      
      res.json({
        structures,
        notifications,
        mismatches,
        staleNotifications, // ✅ НОВОЕ
        issues,
        stats
      });
      
    } catch (error) {
      console.error('Diagnose error:', error);
      res.status(500).json({ error: error.message });
    }
});

// Исправить конкретное уведомление
app.put('/api/admin/fix-notification/:id',
  authenticateToken,
  requirePerm('db_tools'),
  async (req, res) => {
    try {
      const { newResId, password } = req.body;
      
      console.log('=== FIX NOTIFICATION ===');
      console.log('Notification ID:', req.params.id);
      console.log('New ResId:', newResId);
      console.log('Password provided:', !!password);
      
      if (password !== DELETE_PASSWORD) {
        console.log('❌ Wrong password');
        return res.status(403).json({ error: 'Неверный пароль' });
      }
      
      const notification = await Notification.findByPk(req.params.id);
      
      if (!notification) {
        console.log('❌ Notification not found');
        return res.status(404).json({ error: 'Уведомление не найдено' });
      }
      
      const oldResId = notification.resId;
      console.log(`Changing resId: ${oldResId} → ${newResId}`);
      
      await notification.update({ resId: parseInt(newResId) });
      
      // Проверяем что обновилось
      await notification.reload();
      console.log(`✅ Updated! New resId in DB: ${notification.resId}`);
      
      res.json({
        success: true,
        message: `ResId изменен: ${oldResId} → ${notification.resId}`,
        oldResId,
        newResId: notification.resId
      });
      
    } catch (error) {
      console.error('❌ Fix notification error:', error);
      res.status(500).json({ error: error.message });
    }
});

// Автоисправление по структуре (для одного уведомления)
app.post('/api/admin/auto-fix-notification/:id',
  authenticateToken,
  requirePerm('db_tools'),
  async (req, res) => {
    try {
      const { password } = req.body;
      
      console.log('=== AUTO-FIX NOTIFICATION ===');
      console.log('Notification ID:', req.params.id);
      console.log('Password provided:', !!password);
      
      if (password !== DELETE_PASSWORD) {
        console.log('❌ Wrong password');
        return res.status(403).json({ error: 'Неверный пароль' });
      }
      
      // Сначала находим уведомление
      const notification = await Notification.findByPk(req.params.id);
      
      if (!notification) {
        console.log('❌ Notification not found');
        return res.status(404).json({ error: 'Уведомление не найдено' });
      }
      
      console.log('Found notification with networkStructureId:', notification.networkStructureId);
      
      if (!notification.networkStructureId) {
        console.log('❌ No networkStructureId');
        return res.status(400).json({ error: 'Нет привязки к структуре' });
      }
      
      // Теперь находим структуру отдельно
      const structure = await NetworkStructure.findByPk(notification.networkStructureId, {
        include: [ResUnit]
      });
      
      if (!structure) {
        console.log('❌ Structure not found');
        return res.status(400).json({ error: 'Структура не найдена' });
      }
      
      const correctResId = structure.resId;
      const oldResId = notification.resId;
      
      console.log(`Auto-fixing: ${oldResId} → ${correctResId} (from structure)`);
      
      await notification.update({ resId: correctResId });
      
      // Проверяем что обновилось
      await notification.reload();
      console.log(`✅ Auto-fixed! New resId in DB: ${notification.resId}`);
      
      res.json({
        success: true,
        message: `Автоисправление: ${oldResId} → ${notification.resId}`,
        oldResId,
        newResId: notification.resId
      });
      
    } catch (error) {
      console.error('❌ Auto-fix error:', error);
      res.status(500).json({ error: error.message });
    }
});

// Массовое автоисправление ВСЕХ несоответствий для РЭС
app.post('/api/admin/auto-fix-all/:resId',
  authenticateToken,
  requirePerm('db_tools'),
  async (req, res) => {
    const transaction = await sequelize.transaction();
    
    try {
      const { password } = req.body;
      const resId = parseInt(req.params.resId);
      
      console.log('=== MASS AUTO-FIX ===');
      console.log('ResId:', resId);
      console.log('Password provided:', !!password);
      
      if (password !== DELETE_PASSWORD) {
        await transaction.rollback();
        return res.status(403).json({ error: 'Неверный пароль' });
      }
      
      // Находим все структуры этого РЭС
      const structures = await NetworkStructure.findAll({
        where: { resId },
        transaction
      });
      
      console.log(`Found ${structures.length} structures for RES ${resId}`);
      
      // Находим все уведомления с привязкой к этим структурам
      const notifications = await Notification.findAll({
        where: {
          networkStructureId: {
            [Op.in]: structures.map(s => s.id)
          }
        },
        include: [{
          model: NetworkStructure,
          required: true
        }],
        transaction
      });
      
      console.log(`Found ${notifications.length} notifications`);
      
      let fixed = 0;
      let alreadyCorrect = 0;
      const fixedDetails = [];
      
      for (const notif of notifications) {
        const correctResId = notif.NetworkStructure.resId;
        
        if (notif.resId !== correctResId) {
          const oldResId = notif.resId;
          
          await notif.update({
            resId: correctResId
          }, { transaction });
          
          fixedDetails.push({
            notificationId: notif.id,
            oldResId,
            newResId: correctResId,
            tpName: notif.NetworkStructure.tpName,
            vlName: notif.NetworkStructure.vlName
          });
          
          fixed++;
          console.log(`✅ Fixed notification ${notif.id}: ${oldResId} → ${correctResId}`);
        } else {
          alreadyCorrect++;
        }
      }
      
      await transaction.commit();
      
      console.log(`=== MASS AUTO-FIX COMPLETE ===`);
      console.log(`Fixed: ${fixed}, Already correct: ${alreadyCorrect}`);
      
      res.json({
        success: true,
        message: `Массовое исправление завершено!`,
        stats: {
          total: notifications.length,
          fixed,
          alreadyCorrect
        },
        fixedDetails
      });
      
    } catch (error) {
      await transaction.rollback();
      console.error('❌ Mass auto-fix error:', error);
      res.status(500).json({ error: error.message });
    }
});

// Новый эндпоинт для аналитики
app.get('/api/analytics/summary', 
  authenticateToken, 
  async (req, res) => {
    try {
      const { dateFrom, dateTo } = req.query;
      
      // Условие по РЭС
      let resCondition = {};
      if (!req.user.allRes) {
        resCondition.id = req.user.resId;
      }
      
      // Получаем все РЭС
      const resList = await ResUnit.findAll({
        where: resCondition,
        order: [['name', 'ASC']]
      });

      // Условие по дате
      let dateCondition = {};
      if (dateFrom || dateTo) {
        dateCondition.uploadedAt = {};
        if (dateFrom) dateCondition.uploadedAt[Op.gte] = new Date(dateFrom);
        if (dateTo) {
          const endDate = new Date(dateTo);
          endDate.setHours(23, 59, 59, 999);
          dateCondition.uploadedAt[Op.lte] = endDate;
        }
      }

      // Собираем статистику для каждого РЭС
      const analytics = await Promise.all(
        resList.map(async (res) => {
          const structures = await NetworkStructure.findAll({
            where: { resId: res.id }
          });
          
          const tpCount = new Set(structures.map(s => s.tpName)).size;
          const vlCount = structures.length;
          
          let totalPuCount = 0;
          const allPuNumbers = new Set();
          
          structures.forEach(s => {
            if (s.startPu) { totalPuCount++; allPuNumbers.add(s.startPu); }
            if (s.middlePu) { totalPuCount++; allPuNumbers.add(s.middlePu); }
            if (s.endPu) { totalPuCount++; allPuNumbers.add(s.endPu); }
          });
          
          // 2. Получаем ВСЕ загрузки в периоде для этого РЭС
          const uploads = await PuUploadHistory.findAll({
            where: {
              puNumber: Array.from(allPuNumbers),
              ...dateCondition
            },
            order: [['uploadedAt', 'DESC']]
          });
          
          // 3. НОВАЯ ЛОГИКА: уникальные ПУ (берем только последнюю загрузку каждого ПУ)
          const uniquePuUploads = new Map();
          uploads.forEach(upload => {
            if (!uniquePuUploads.has(upload.puNumber)) {
              uniquePuUploads.set(upload.puNumber, upload);
            }
          });
          
          const uniqueCheckedPuCount = uniquePuUploads.size;
          
          // 4. НОВАЯ ЛОГИКА: проверенные ВЛ (где хотя бы один ПУ проверен)
          const checkedVLCount = structures.filter(structure => {
            const puNumbers = [structure.startPu, structure.middlePu, structure.endPu]
              .filter(Boolean);
            
            // Проверяем есть ли хотя бы одна загрузка по ПУ этой ВЛ
            return puNumbers.some(puNumber => uniquePuUploads.has(puNumber));
          }).length;
          
          // 5. Проценты
          const vlCoveragePercent = vlCount > 0 
            ? Math.round((checkedVLCount / vlCount) * 100) 
            : 0;
          
          const puCoveragePercent = totalPuCount > 0
            ? Math.round((uniqueCheckedPuCount / totalPuCount) * 100)
            : 0;

          // Перегруз секций (этап 3)
          const overloadSections = await TpSection.count({
            where: { resId: res.id, overloadStatus: 'overload' }
          });
          const activeOverloadCases = await OverloadCase.count({
            where: { resId: res.id, stage: { [Op.ne]: 'completed' } }
          });

          return {
            resId: res.id,
            resName: res.name,
            tpCount,
            vlCount,
            vlCoveragePercent,
            totalPuCount,
            uniqueCheckedPuCount,
            puCoveragePercent,
            overloadSections,
            activeOverloadCases
          };
        })
      );

      // Итоги
      const totals = analytics.reduce((acc, curr) => ({
        tpCount: acc.tpCount + curr.tpCount,
        vlCount: acc.vlCount + curr.vlCount,
        totalPuCount: acc.totalPuCount + curr.totalPuCount,
        uniqueCheckedPuCount: acc.uniqueCheckedPuCount + curr.uniqueCheckedPuCount,
        overloadSections: acc.overloadSections + curr.overloadSections,
        activeOverloadCases: acc.activeOverloadCases + curr.activeOverloadCases
      }), {
        tpCount: 0,
        vlCount: 0,
        totalPuCount: 0,
        uniqueCheckedPuCount: 0,
        overloadSections: 0,
        activeOverloadCases: 0
      });
      
      // Добавляем проценты к итогам
      totals.vlCoveragePercent = totals.vlCount > 0
        ? Math.round((analytics.reduce((sum, a) => sum + (a.vlCoveragePercent * a.vlCount / 100), 0) / totals.vlCount) * 100)
        : 0;
      
      totals.puCoveragePercent = totals.totalPuCount > 0
        ? Math.round((totals.uniqueCheckedPuCount / totals.totalPuCount) * 100)
        : 0;
      
      res.json({
        analytics,
        totals,
        period: {
          from: dateFrom,
          to: dateTo
        }
      });
      
    } catch (error) {
      console.error('Analytics error:', error);
      res.status(500).json({ error: error.message });
    }
});

// «Глубокий анализ работ» — детектор формально закрытых мероприятий РЭС.
// Просмотровый роут (admin), без requirePerm. Всё считается из CheckHistory на лету,
// БЕЗ новых таблиц и БЕЗ запросов в цикле — один findAll, дальше математика в JS.
app.get('/api/analytics/work-quality', authenticateToken, requirePerm('deep_analysis_view'), async (req, res) => {
  try {
    const numParam = (v, def, min, max) => {
      const n = Number(v);
      if (!Number.isFinite(n)) return def;
      return Math.min(max, Math.max(min, Math.round(n)));
    };
    const windowDays  = numParam(req.query.windowDays, 4, 1, 60);
    const brigades    = numParam(req.query.brigades, 3, 1, 100);
    const perBrigade  = numParam(req.query.perBrigade, 5, 1, 100);
    const batchMinutes = numParam(req.query.batchMinutes, 60, 1, 1440);
    const batchCount   = numParam(req.query.batchCount, 10, 2, 1000);
    // fastDays — за сколько дней после проверки закрытие считаем «мгновенным»
    // (для выездных мероприятий физически невыполнимо), дефолт 1.
    const fastDays     = numParam(req.query.fastDays, 1, 0, 60);
    const resIdFilter  = req.query.resId ? parseInt(req.query.resId, 10) : null;

    // Период по workCompletedDate — дефолт последние 30 дней.
    let dTo = req.query.dateTo ? new Date(req.query.dateTo) : new Date();
    let dFrom = req.query.dateFrom ? new Date(req.query.dateFrom) : new Date(dTo);
    if (!req.query.dateFrom) dFrom.setDate(dFrom.getDate() - 30);
    const from = new Date(dFrom); from.setHours(0, 0, 0, 0);
    const to = new Date(dTo); to.setHours(23, 59, 59, 999);

    const where = { workCompletedDate: { [Op.ne]: null, [Op.gte]: from, [Op.lte]: to } };
    if (resIdFilter) where.resId = resIdFilter;

    const records = await CheckHistory.findAll({
      where,
      attributes: ['id', 'resId', 'puNumber', 'tpName', 'vlName', 'resComment',
        'attachments', 'initialCheckDate', 'workCompletedDate', 'recheckDate',
        'recheckResult', 'status', 'failureCount'],
      include: [{ model: ResUnit, attributes: ['id', 'name'] }]
    });

    const params = {
      resId: resIdFilter, dateFrom: from.toISOString(), dateTo: to.toISOString(),
      windowDays, brigades, perBrigade, batchMinutes, batchCount, fastDays
    };
    if (records.length === 0) return res.json({ params, rows: [] });

    // Группировка по РЭС
    const byRes = new Map();
    for (const r of records) {
      if (!byRes.has(r.resId)) byRes.set(r.resId, { resId: r.resId, resName: r.ResUnit?.name || '—', items: [] });
      byRes.get(r.resId).items.push(r);
    }

    const periodDays = Math.max(1, Math.round((to - from) / 86400000));
    const capacityLimit = brigades * perBrigade * windowDays;

    const parseAtt = (a) => {
      if (!a) return [];
      if (Array.isArray(a)) return a;
      if (typeof a === 'string') { try { const p = JSON.parse(a); return Array.isArray(p) ? p : []; } catch { return []; } }
      return [];
    };
    const normComment = (c) => String(c || '').trim().toLowerCase().replace(/\s+/g, ' ');
    const dayKey = (dt) => {
      const d = new Date(dt);
      return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    };
    const dayDiff2 = (a, b) => (a && b) ? Math.round((new Date(b) - new Date(a)) / 86400000) : null;
    // Единая строка реестра для всех блоков проблем: ТП, ВЛ, ПУ, позиция, текст
    // ошибки, дата ошибки → дата закрытия, дней, текущее состояние (status + перепроверка).
    const recOf = (x) => ({
      tpName: x.tpName, vlName: x.vlName, puNumber: x.puNumber, position: x.position,
      error: String(x.initialError || '').trim().slice(0, 300),
      checkDate: x.initialCheckDate, closeDate: x.workCompletedDate,
      days: dayDiff2(x.initialCheckDate, x.workCompletedDate),
      status: x.status, recheck: x.recheckResult,
      failureCount: x.failureCount || 1,
    });

    const rows = [];
    for (const g of byRes.values()) {
      const items = g.items
        .filter(x => x.workCompletedDate)
        .sort((a, b) => new Date(a.workCompletedDate) - new Date(b.workCompletedDate));
      const totalClosed = items.length;
      const uniqueTp = new Set(items.map(x => x.tpName)).size;
      const closes = items.map(x => ({ t: new Date(x.workCompletedDate), tp: x.tpName }));

      // 1) capacity — скользящее окно windowDays календарных дней, макс уникальных ТП.
      let maxTpInWindow = 0, windowStart = null;
      const dayStarts = [...new Set(closes.map(c => dayKey(c.t)))].sort();
      for (const ds of dayStarts) {
        const start = new Date(`${ds}T00:00:00`);
        const end = new Date(start); end.setDate(end.getDate() + windowDays);
        const tps = new Set();
        for (const c of closes) if (c.t >= start && c.t < end) tps.add(c.tp);
        if (tps.size > maxTpInWindow) { maxTpInWindow = tps.size; windowStart = start; }
      }
      const ratio = capacityLimit > 0 ? maxTpInWindow / capacityLimit : 0;

      // 2) batches — ПАКЕТ = КАЛЕНДАРНЫЙ ДЕНЬ закрытия. Пакетным (подозрительным)
      //    считаем день, в который закрыто больше дневной ёмкости РЭС
      //    (dailyCap = бригады × план ТП/бригаду). Внутри — реестр ТП с ВЛ,
      //    датой ошибки, датой закрытия и текущим состоянием.
      const dailyCap = Math.max(1, brigades * perBrigade);
      const byDay = new Map(); // dayKey -> item[]
      for (const x of items) {
        const key = dayKey(x.workCompletedDate);
        if (!byDay.has(key)) byDay.set(key, []);
        byDay.get(key).push(x);
      }
      const batches = [];
      const batchTpCount = new Map(); // tpName -> сколько закрытий этого ТП попало в пакетные дни
      for (const [day, dayItems] of [...byDay.entries()].sort((a, b) => a[0].localeCompare(b[0]))) {
        if (dayItems.length <= dailyCap) continue; // норма для дня — не пакет
        for (const x of dayItems) batchTpCount.set(x.tpName, (batchTpCount.get(x.tpName) || 0) + 1);
        const recs = dayItems.map(recOf).sort((a, b) => a.tpName.localeCompare(b.tpName, 'ru'));
        batches.push({
          day,
          count: dayItems.length,
          uniqueTp: new Set(dayItems.map(x => x.tpName)).size,
          dailyCap,
          records: recs,
        });
      }
      batches.sort((a, b) => b.count - a.count);

      // 3) afterHours — доля закрытий с локальным часом <7 или >=21 (TZ Europe/Moscow).
      let afterCnt = 0;
      for (const c of closes) { const h = c.t.getHours(); if (h < 7 || h >= 21) afterCnt++; }
      const afterHoursShare = totalClosed > 0 ? afterCnt / totalClosed : 0;

      // 4) templateComments — одинаковый нормализованный resComment, count>=5 на РАЗНЫХ ТП.
      const commentMap = new Map();
      for (const x of items) {
        const norm = normComment(x.resComment);
        if (!norm) continue;
        if (!commentMap.has(norm)) commentMap.set(norm, { text: String(x.resComment || '').trim(), tps: new Set(), count: 0, records: [] });
        const e = commentMap.get(norm); e.count++; e.tps.add(x.tpName); e.records.push(recOf(x));
      }
      // Шаблонный коммент группируем по ГРУППЕ ТП: несём реестр записей (ТП/ВЛ/
      // позиция/даты/состояние) — фронт объединит по ТП, ВЛ покажет раздельно.
      const templateComments = [];
      for (const e of commentMap.values()) {
        if (e.count >= 5 && e.tps.size >= 2) {
          const tps = [...e.tps].sort((a, b) => a.localeCompare(b, 'ru'));
          const recs = e.records.sort((a, b) => a.tpName.localeCompare(b.tpName, 'ru'));
          templateComments.push({ text: e.text.slice(0, 200), count: e.count, tpCount: tps.length, tps: tps.slice(0, 50), records: recs.slice(0, 100) });
        }
      }
      templateComments.sort((a, b) => b.count - a.count);

      // 5) duplicatePhotos — ключ original_name|size, встречается на РАЗНЫХ ТП.
      //    Несём пример вложения (для превью через FileViewer) и реестр записей
      //    (ТП/ВЛ/даты/состояние), где этот файл приложен.
      const photoMap = new Map();
      for (const x of items) {
        for (const att of parseAtt(x.attachments)) {
          const name = att && att.original_name;
          const size = att && att.size;
          if (!name || size == null) continue;
          const key = `${name}|${size}`;
          if (!photoMap.has(key)) photoMap.set(key, { name, size, tps: new Set(), records: [], sample: null });
          const e = photoMap.get(key);
          e.tps.add(x.tpName);
          e.records.push(recOf(x));
          if (!e.sample && att.url && att.public_id) {
            e.sample = { url: att.url, public_id: att.public_id, original_name: att.original_name, mime_type: att.mime_type };
          }
        }
      }
      const duplicatePhotos = [];
      for (const e of photoMap.values()) {
        if (e.tps.size >= 2) {
          const recs = e.records.sort((a, b) => a.tpName.localeCompare(b.tpName, 'ru'));
          duplicatePhotos.push({
            name: e.name,
            tpCount: e.tps.size,
            isPdf: /\.pdf$/i.test(e.name || '') || (e.sample && /\.pdf$/i.test(e.sample.public_id || '')),
            sample: e.sample,
            records: recs.slice(0, 100),
          });
        }
      }
      duplicatePhotos.sort((a, b) => b.tpCount - a.tpCount);

      // 5b) visualDuplicates — ВИЗУАЛЬНЫЕ дубли по перцептивному хешу (pHash):
      //     тот же кадр даже при другом имени/размере или внутри PDF. Кластеризуем
      //     по расстоянию Хэмминга; группа на ≥2 РАЗНЫХ ТП — подозрительна.
      const PHASH_THRESHOLD = 6; // из 64 бит
      const hamHex = (a, b) => {
        let x;
        try { x = BigInt('0x' + a) ^ BigInt('0x' + b); } catch { return 64; }
        let c = 0; while (x) { c += Number(x & 1n); x >>= 1n; }
        return c;
      };
      const phEntries = [];
      for (const x of items) {
        for (const att of parseAtt(x.attachments)) {
          const phs = Array.isArray(att.phashes) ? att.phashes : [];
          for (const ph of phs) {
            if (typeof ph === 'string' && ph.length >= 8) phEntries.push({ ph, tpName: x.tpName, rec: recOf(x), att });
          }
        }
        if (phEntries.length > 4000) break; // предохранитель от тяжёлой кластеризации
      }
      const phClusters = [];
      for (const e of phEntries) {
        let cl = null;
        for (const c of phClusters) { if (hamHex(c.rep, e.ph) <= PHASH_THRESHOLD) { cl = c; break; } }
        if (!cl) { cl = { rep: e.ph, tps: new Set(), records: [], sample: null, isPdf: false }; phClusters.push(cl); }
        cl.tps.add(e.tpName);
        cl.records.push(e.rec);
        if (!cl.sample && e.att && e.att.url && e.att.public_id) {
          cl.sample = { url: e.att.url, public_id: e.att.public_id, original_name: e.att.original_name, mime_type: e.att.mime_type };
          cl.isPdf = /\.pdf$/i.test(e.att.original_name || '') || /\.pdf$/i.test(e.att.public_id || '');
        }
      }
      const visualDuplicates = [];
      for (const c of phClusters) {
        if (c.tps.size >= 2) {
          const seen = new Set();
          const recs = [];
          for (const r of c.records) {
            const k = `${r.tpName}|${r.puNumber}|${r.closeDate}`;
            if (!seen.has(k)) { seen.add(k); recs.push(r); }
          }
          recs.sort((a, b) => a.tpName.localeCompare(b.tpName, 'ru'));
          visualDuplicates.push({ tpCount: c.tps.size, count: recs.length, isPdf: c.isPdf, sample: c.sample, records: recs.slice(0, 100) });
        }
      }
      visualDuplicates.sort((a, b) => b.tpCount - a.tpCount);

      // 6) recheckFail — доля error среди перепроверенных (ok/error).
      const rechecked = items.filter(x => x.recheckResult === 'ok' || x.recheckResult === 'error');
      const checked = rechecked.length;
      const failed = rechecked.filter(x => x.recheckResult === 'error').length;
      const failShare = checked > 0 ? failed / checked : 0;
      const recheckFailFlag = failShare > 0.4 && checked >= 5;

      // 7) fast — «мгновенное» закрытие: дней между проверкой и мероприятиями <= fastDays.
      //    Это ядро признака фикции: выездные работы не делаются в день/следующий день.
      const dayDiff = (a, b) => (a && b) ? Math.round((new Date(b) - new Date(a)) / 86400000) : null;
      let fastCount = 0;
      for (const x of items) {
        const d = dayDiff(x.initialCheckDate, x.workCompletedDate);
        if (d != null && d >= 0 && d <= fastDays) fastCount++;
      }
      const fastShare = totalClosed > 0 ? fastCount / totalClosed : 0;
      const fastFlag = fastShare > 0.5 && totalClosed >= 5;

      // Перечень подозрительных ТП для ручного разбора: по каждому ТП — в чём проблема
      // и по каждой записи — дата проверки → дата мероприятий и сколько дней прошло.
      const tpAgg = new Map();
      const tpEntry = (tp) => {
        if (!tpAgg.has(tp)) tpAgg.set(tp, { tpName: tp, closedCount: 0, batch: 0, templates: [], photos: [], night: 0, recheckError: 0, records: [], fast: 0 });
        return tpAgg.get(tp);
      };
      for (const x of items) {
        const e = tpEntry(x.tpName);
        e.closedCount++;
        const days = dayDiff(x.initialCheckDate, x.workCompletedDate);
        if (days != null && days >= 0 && days <= fastDays) e.fast++;
        e.records.push({ puNumber: x.puNumber, checkDate: x.initialCheckDate, workDate: x.workCompletedDate, days });
      }
      for (const [tp, c] of batchTpCount) tpEntry(tp).batch += c;
      for (const e of commentMap.values()) {
        if (e.count >= 5 && e.tps.size >= 2) for (const tp of e.tps) tpEntry(tp).templates.push(e.text.slice(0, 120));
      }
      for (const e of photoMap.values()) {
        if (e.tps.size >= 2) for (const tp of e.tps) tpEntry(tp).photos.push(e.name);
      }
      for (const c of closes) { const h = c.t.getHours(); if (h < 7 || h >= 21) tpEntry(c.tp).night++; }
      for (const x of items) if (x.recheckResult === 'error') tpEntry(x.tpName).recheckError++;

      const suspectTp = [];
      for (const e of tpAgg.values()) {
        const problems = [];
        const daysVals = e.records.map(r => r.days).filter(d => d != null && d >= 0);
        const minDays = daysVals.length ? Math.min(...daysVals) : null;
        if (e.fast > 0) problems.push(`Мгновенное закрытие: ${minDays === 0 ? 'в день проверки' : `через ${minDays} дн. после проверки`} (${e.fast} из ${e.closedCount}) — выездные работы за такой срок невыполнимы`);
        if (e.batch > 0) problems.push(`Пакетное закрытие (${e.batch} шт.) — закрыто «пачкой» вместе с другими ТП за минуты`);
        if (e.templates.length) problems.push(`Шаблонный комментарий: «${e.templates[0]}»${e.templates.length > 1 ? ` и ещё ${e.templates.length - 1}` : ''} — тот же текст на других ТП`);
        if (e.photos.length) problems.push(`Дубль фото: ${e.photos.length === 1 ? e.photos[0] : `${e.photos.length} файл(ов)`} — то же фото приложено к другим ТП`);
        if (e.night > 0) problems.push(`Закрытий в нерабочее время: ${e.night} (до 7:00 / после 21:00)`);
        if (e.recheckError > 0) problems.push(`Провал перепроверки: ${e.recheckError} — после «мероприятий» ошибка повторилась`);
        if (problems.length) {
          const recs = e.records.slice().sort((a, b) => new Date(a.workDate) - new Date(b.workDate));
          suspectTp.push({
            tpName: e.tpName, closedCount: e.closedCount, minDays, fast: e.fast,
            batch: e.batch, night: e.night, recheckError: e.recheckError,
            templates: e.templates.slice(0, 3), photos: e.photos.slice(0, 3),
            problems, records: recs
          });
        }
      }
      suspectTp.sort((a, b) => b.problems.length - a.problems.length || b.closedCount - a.closedCount || a.tpName.localeCompare(b.tpName, 'ru'));

      // Скоринг 0..100
      let score = 0;
      const reasons = [];
      if (ratio > 2) {
        score += 45;
        reasons.push(`Пик закрытий ${maxTpInWindow} ТП за ${windowDays} дн. при ёмкости ${capacityLimit} — превышение в ${ratio.toFixed(1)} раза`);
      } else if (ratio > 1) {
        score += 30;
        reasons.push(`Пик закрытий ${maxTpInWindow} ТП за ${windowDays} дн. при ёмкости ${capacityLimit} — превышение в ${ratio.toFixed(1)} раза`);
      }
      if (batches.length > 0) {
        score += Math.min(30, batches.length * 10);
        reasons.push(`Пакетные закрытия: ${batches.length} дн., где закрыто больше дневной ёмкости (>${dailyCap} ТП/день)`);
      }
      if (afterHoursShare > 0.3) {
        score += 10;
        reasons.push(`Доля ночных закрытий (до 7:00 / после 21:00): ${Math.round(afterHoursShare * 100)}%`);
      }
      if (templateComments.length > 0) {
        score += Math.min(15, templateComments.length * 5);
        reasons.push(`Шаблонные комментарии: ${templateComments.length} (одинаковый текст на разных ТП)`);
      }
      if (duplicatePhotos.length > 0) {
        score += Math.min(30, duplicatePhotos.length * 10);
        reasons.push(`Дубли фото: ${duplicatePhotos.length} файл(ов) встречаются на разных ТП`);
      }
      if (visualDuplicates.length > 0) {
        score += Math.min(20, visualDuplicates.length * 10);
        reasons.push(`Визуально одинаковые фото/сканы: ${visualDuplicates.length} групп(ы) на разных ТП`);
      }
      if (recheckFailFlag) {
        score += 20;
        reasons.push(`Провалы перепроверки: ${Math.round(failShare * 100)}% из ${checked} перепроверенных`);
      }
      if (fastFlag) {
        score += 15;
        reasons.push(`Мгновенное закрытие: ${fastCount} из ${totalClosed} (${Math.round(fastShare * 100)}%) закрыты в течение ≤${fastDays} дн. после проверки — для выездных работ физически невыполнимо`);
      }
      score = Math.min(100, score);
      const level = score >= 60 ? 'red' : score >= 30 ? 'yellow' : 'green';

      // Реестр «Мгновенное закрытие»: все закрытия с зазором проверка→закрытие
      // 0..30 дн. (фронт фильтрует ползунками «от/до дней», дефолт 0..fastDays).
      const fastList = items
        .map(recOf)
        .filter(r => r.days != null && r.days >= 0 && r.days <= 30)
        .sort((a, b) => a.days - b.days || a.tpName.localeCompare(b.tpName, 'ru'));
      // Реестр «Провал перепроверки»: записи с recheckResult === 'error'.
      const recheckFailList = items
        .filter(x => x.recheckResult === 'error')
        .map(recOf)
        .sort((a, b) => a.tpName.localeCompare(b.tpName, 'ru'));
      // Реестр «Отработка только после повторных съёмов»: ошибка приходила ≥2 раз
      // до отработки (repeatCount перенесён в failureCount при закрытии).
      const repeatList = items
        .filter(x => (x.failureCount || 1) >= 2)
        .map(recOf)
        .sort((a, b) => b.failureCount - a.failureCount || a.tpName.localeCompare(b.tpName, 'ru'));
      // Реестр «Закрытия в нерабочее время»: локальный час закрытия <7 или ≥21.
      const nightList = items
        .filter(x => { const h = new Date(x.workCompletedDate).getHours(); return h < 7 || h >= 21; })
        .map(recOf)
        .sort((a, b) => a.tpName.localeCompare(b.tpName, 'ru'));

      rows.push({
        resId: g.resId, resName: g.resName,
        totalClosed, uniqueTp, periodDays,
        capacity: { maxTpInWindow, windowStart: windowStart ? windowStart.toISOString() : null, capacityLimit, ratio },
        batches, afterHoursShare, templateComments, duplicatePhotos, visualDuplicates,
        recheck: { checked, failed, failShare },
        fast: { count: fastCount, share: fastShare, days: fastDays },
        fastList, recheckFailList, repeatList, nightList,
        suspectTp,
        score, level, reasons
      });
    }

    rows.sort((a, b) => b.score - a.score);
    res.json({ params, rows });
  } catch (error) {
    console.error('work-quality error:', error);
    res.status(500).json({ error: error.message });
  }
});

// Бэкфилл pHash для УЖЕ загруженных вложений (визуальные дубли в анализе видят
// только вложения с phashes). Обрабатываем партиями — админ вызывает повторно,
// пока remaining>0. Best-effort: ошибка по файлу пропускается.
app.post('/api/admin/backfill-phash', authenticateToken, requirePerm('db_tools'), async (req, res) => {
  const https = require('https');
  const downloadBuffer = (url) => new Promise((resolve, reject) => {
    https.get(url, (r) => {
      if (r.statusCode !== 200) { r.resume(); return reject(new Error('HTTP ' + r.statusCode)); }
      const chunks = [];
      r.on('data', c => chunks.push(c));
      r.on('end', () => resolve(Buffer.concat(chunks)));
    }).on('error', reject);
  });
  try {
    const limit = Math.min(100, Math.max(1, parseInt(req.body?.limit, 10) || 30));
    // Кандидаты: записи с непустыми attachments. Фильтр «нужен phash» — в коде.
    const records = await CheckHistory.findAll({
      where: { attachments: { [Op.ne]: null } },
      order: [['id', 'ASC']],
      limit: 3000,
    });
    let scanned = 0, updated = 0, filesHashed = 0, remaining = 0;
    for (const rec of records) {
      const atts = Array.isArray(rec.attachments) ? rec.attachments : [];
      const needs = atts.some(a => a && a.url && !Array.isArray(a.phashes));
      if (!needs) continue;
      remaining++;
      if (scanned >= limit) continue; // за этот вызов не обрабатываем — вернём в remaining
      scanned++;
      let changed = false;
      for (const a of atts) {
        if (!a || !a.url || Array.isArray(a.phashes)) continue;
        try {
          const buf = await downloadBuffer(a.url);
          const mime = a.mime_type || (/\.pdf$/i.test(a.url) ? 'application/pdf' : 'image/jpeg');
          a.phashes = await computePhashes(buf, mime);
          filesHashed++;
          changed = true;
        } catch (e) {
          a.phashes = []; // помечаем обработанным, чтобы не зациклиться на битом файле
          changed = true;
          console.error('[backfill-phash] file:', e.message);
        }
      }
      if (changed) {
        rec.set('attachments', atts);
        rec.changed('attachments', true);
        await rec.save();
        updated++;
      }
    }
    remaining = Math.max(0, remaining - scanned);
    res.json({ scannedRecords: scanned, updatedRecords: updated, filesHashed, remaining });
  } catch (error) {
    console.error('backfill-phash error:', error);
    res.status(500).json({ error: error.message });
  }
});
