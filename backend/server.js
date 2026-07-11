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

process.on('unhandledRejection', (err) => {
  console.error('Unhandled Rejection:', err);
  process.exit(1);
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
const app = express();
app.use((req, res, next) => {
  if (req.path.startsWith('/api')) {
    res.setHeader('Content-Type', 'application/json; charset=utf-8');
  }
  next();
});
const PORT = process.env.PORT || 3000;

// ВАЖНО: Пароль для удаления из переменной окружения
const DELETE_PASSWORD = process.env.DELETE_PASSWORD || '1191';

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

// Функция загрузки в Cloudinary
async function uploadToCloudinary(file, type = 'attachment') {
  const isPdf = file.mimetype === 'application/pdf';
  
  return new Promise((resolve, reject) => {
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
    
    const finalPublicId = `${type}_${timestamp}_${safeName}${ext}`;
    
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
  role: {
    type: DataTypes.ENUM('admin', 'uploader', 'res_responsible'),
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
  lastUpdate: {
    type: DataTypes.DATE,
    defaultValue: DataTypes.NOW
  }
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
    type: DataTypes.ENUM('error', 'success', 'info', 'pending_check', 'pending_askue', 'problem_vl'), // ← ЗДЕСЬ ДОБАВИТЬ 'problem_vl'
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
    type: DataTypes.ENUM('rim_single', 'rim_mass', 'nartis', 'energomera'),
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
    type: DataTypes.ENUM('rim_single', 'rim_mass', 'nartis', 'energomera'),
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

// Middleware для проверки JWT токена
const authenticateToken = (req, res, next) => {
  const authHeader = req.headers['authorization'];
  const token = authHeader && authHeader.split(' ')[1];

  if (!token) {
    return res.status(401).json({ error: 'Access denied' });
  }

  jwt.verify(token, process.env.JWT_SECRET || 'secret-key', (err, user) => {
    if (err) {
      return res.status(403).json({ error: 'Invalid token' });
    }
    req.user = user;
    next();
  });
};

// Middleware для проверки ролей
const checkRole = (roles) => {
  return (req, res, next) => {
    if (!roles.includes(req.user.role)) {
      return res.status(403).json({ error: 'Access denied for your role' });
    }
    next();
  };
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

// Обработчик ошибок multer
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

// Email сервис
const createEmailTransporter = () => {
  return nodemailer.createTransport({
    host: process.env.MAIL_HOST || 'smtp.mail.ru',
    port: process.env.MAIL_PORT || 465,
    secure: true,
    auth: {
      user: process.env.MAIL_USER,
      pass: process.env.MAIL_PASS
    }
  });
};

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
      process.env.JWT_SECRET || 'secret-key',
      { expiresIn: '24h' }
    );
    
    res.json({
      token,
      user: {
        id: user.id,
        fio: user.fio,
        role: user.role,
        resId: user.resId,
        resName: user.ResUnit?.name
      }
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
      process.env.JWT_SECRET || 'secret-key',
      { expiresIn: '24h' }
    );
    res.json({
      token: appToken,
      user: {
        id: user.id,
        fio: user.fio,
        role: user.role,
        resId: user.resId,
        resName: user.ResUnit?.name
      }
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

    const counts = await getNotificationCounts(user);
    let count = 0;
    if (user.role === 'admin') count = counts.tech_pending + counts.askue_pending + counts.problem_vl;
    else if (user.role === 'res_responsible') count = counts.tech_pending;
    else if (user.role === 'uploader') count = counts.askue_pending;
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
      attributes: ['id', 'fio', 'login', 'role', 'resId', 'email'],
      include: [ResUnit]
    });
    
    if (!user) {
      return res.status(404).json({ error: 'User not found' });
    }
    
    res.json({
      user: {
        id: user.id,
        fio: user.fio,
        role: user.role,
        resId: user.resId,
        resName: user.ResUnit?.name
      }
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
    if (req.user.role !== 'admin' && resId != req.user.resId) {
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
        ResUnit
      ],
      order: [['tpName', 'ASC'], ['vlName', 'ASC']]
    });
    
    res.json(structures);
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// 4. ОБНОВЛЕНИЕ структуры сети (только админ)
app.put('/api/network/structure/:id', 
  authenticateToken, 
  checkRole(['admin']), 
  async (req, res) => {
    try {
      const { startPu, middlePu, endPu } = req.body;
      
      await NetworkStructure.update({
        startPu: startPu || null,
        middlePu: middlePu || null,
        endPu: endPu || null,
        lastUpdate: new Date()
      }, {
        where: { id: req.params.id }
      });
      
      res.json({ success: true, message: 'Структура обновлена' });
    } catch (error) {
      res.status(500).json({ error: error.message });
    }
});

// 5. ЗАГРУЗКА ФАЙЛОВ ДЛЯ АНАЛИЗА
app.post('/api/upload/analyze',
  authenticateToken,
  checkRole(['admin', 'uploader']),
  uploadExcel.single('file'),
  async (req, res) => {
    let uploadRecord;
    try {
      const { type, requiredPeriod } = req.body;
      const userId = req.user.id;
      
      // Берем resId из body (если есть) или из токена пользователя
      const resId = req.body.resId || req.user.resId;

      console.log('=== UPLOAD DEBUG ===');
      console.log('userId from token:', userId);
      console.log('req.user:', req.user);

      console.log('=== UPLOAD ANALYZE START ===');
      console.log('User:', req.user);
      console.log('Request body:', req.body);
      console.log('Final resId:', resId);
      console.log('File:', req.file?.originalName);

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
      
      console.log('=== UPLOAD ANALYZE COMPLETE ===');
      
      // Возвращаем результат
      res.json({
        success: true,
        message: 'Файл обработан успешно',
        processed: analysisResult.processed.length,
        errors: analysisResult.errors.length,
        details: analysisResult.processed
      });
      
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
  checkRole(['admin']), 
  uploadExcel.single('file'), 
  async (req, res) => {
    const transaction = await sequelize.transaction();
    
    try {
      const workbook = XLSX.readFile(req.file.path);
      const sheetName = workbook.SheetNames[0];
      const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
      
      let processed = 0;
      let errors = [];
      
           
      // Обрабатываем каждую строку
      for (const row of data) {
        try {
          // Ищем РЭС по полному имени из Excel
          const resName = row['РЭС'];
          const res = await ResUnit.findOne({ 
            where: { name: resName },
            transaction 
          });
          
          if (!res) {
            errors.push(`Неизвестный РЭС: ${resName}`);
            continue;
          }
          
          // Создаем или обновляем запись
          await NetworkStructure.upsert({
            resId: res.id,
            tpName: row['ТП'] || '',
            vlName: row['Фидер'] || '',
            startPu: row['Начало'] ? String(row['Начало']) : null,
            endPu: row['Конец'] ? String(row['Конец']) : null,
            middlePu: row['Середина'] ? String(row['Середина']) : null
          }, {
            transaction
          });
          
          processed++;
          
          // ✅ Находим ID структуры для привязки PuStatus
          const structure = await NetworkStructure.findOne({
            where: { 
              resId: res.id, 
              tpName: row['ТП'] || '', 
              vlName: row['Фидер'] || '' 
            },
            transaction
          });
          const structureId = structure?.id;
          
          // Создаем статусы для новых ПУ
          const positions = [
            { pu: row['Начало'], pos: 'start' },
            { pu: row['Конец'], pos: 'end' },
            { pu: row['Середина'], pos: 'middle' }
          ];
          
          for (const { pu, pos } of positions) {
            if (pu) {
              await PuStatus.findOrCreate({
                where: { puNumber: String(pu) },
                defaults: {
                  networkStructureId: structureId,
                  position: pos,
                  status: 'not_checked'
                },
                transaction
              });
            }
          }
          
        } catch (err) {
          errors.push(`Ошибка в строке ${row['ТП']}-${row['Фидер']}: ${err.message}`);
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

// 7. ПОЛУЧЕНИЕ УВЕДОМЛЕНИЙ

app.get('/api/notifications', authenticateToken, async (req, res) => {
  try {
    const { resId, type } = req.query;   // ← добавили type
    let whereClause = {};
    
    if (req.user.role === 'admin') {
      // Админ видит ВСЕ уведомления нужного РЭС
      if (resId) {
        whereClause = { resId: parseInt(resId) };
      } else {
        whereClause = {};  // Все уведомления
      }
      
    } else if (req.user.role === 'res_responsible') {
      whereClause = {
        resId: req.user.resId,
        [Op.or]: [
          { toUserId: null },
          { toUserId: req.user.id }
        ]
      };
      
    } else if (req.user.role === 'uploader') {
      whereClause = {
        [Op.or]: [
          { toUserId: req.user.id },
          { 
            toUserId: null,
            resId: req.user.resId,
            type: 'pending_askue'
          }
        ]
      };
      
    } else {
      whereClause = { toUserId: req.user.id };
    }
    
    // ← НОВОЕ: фильтруем по типу прямо в БД, чтобы error не вытеснялись
    //   свежими уведомлениями других типов
    if (type) {
      whereClause.type = type;
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
    
    // ✅ Дедупликация: для error и pending_askue — 1 ПУ + 1 набор фаз = 1 запись (последняя)
    const seenKeys = new Set();
    const deduplicatedNotifications = notificationsWithReadStatus.filter(notif => {
      if (notif.type !== 'error' && notif.type !== 'pending_askue') {
        return true; // Остальные типы не трогаем
      }
      
      try {
        const data = JSON.parse(notif.message);
        if (!data.puNumber) return true;
        
        const phaseKey = getPhaseSignature(notif.message);
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

app.post('/api/notifications/:id/complete-work', 
  authenticateToken, 
  checkRole(['res_responsible']),
  upload.array('attachments', 5),
  async (req, res) => {
    const startTime = Date.now();
    console.log('\n=== COMPLETE WORK START ===');
    console.log('Files received:', req.files?.length || 0);
    
    const uploadedFiles = []; // public_id'ы для cleanup при ошибке
    
    try {
      const { comment, checkFromDate } = req.body;
      
      // 1) ВАЛИДАЦИЯ ДО ВСЕГО ОСТАЛЬНОГО (быстрый отказ без транзакции и без загрузок)
      if (!comment || typeof comment !== 'string') {
        return res.status(400).json({ error: 'Комментарий обязателен' });
      }
      const wordCount = comment.trim().split(/\s+/).filter(w => w.length > 0).length;
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

// 9. ОЧИСТКА ВСЕХ ДАННЫХ
app.delete('/api/network/clear-all', 
  authenticateToken, 
  checkRole(['admin']), 
  async (req, res) => {
    const transaction = await sequelize.transaction();
    
    try {
      const { password, beforeDate } = req.body;  // ДОБАВИЛИ beforeDate
      
      if (password !== DELETE_PASSWORD) {
        return res.status(403).json({ error: 'Неверный пароль' });
      }
      
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
  checkRole(['admin']), 
  async (req, res) => {
    const transaction = await sequelize.transaction();
    
    try {
      const { ids, password } = req.body;
      
      // Проверка пароля
      if (password !== DELETE_PASSWORD) {
        return res.status(403).json({ error: 'Неверный пароль' });
      }
      
      if (!ids || !Array.isArray(ids) || ids.length === 0) {
        return res.status(400).json({ error: 'Не выбраны записи для удаления' });
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
    if (req.user.role === 'admin' && req.query.resId) {
      whereClause.resId = parseInt(req.query.resId);
    } else if (req.user.role !== 'admin') {
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
        if (req.user.role === 'admin' && req.query.resId) {
          resWorkloadCondition.id = parseInt(req.query.resId);
        } else if (req.user.role !== 'admin') {
          resWorkloadCondition.id = req.user.resId;
        }
        
        const resListForWorkload = await ResUnit.findAll({
          where: resWorkloadCondition,
          order: [['name', 'ASC']]
        });
        
        // Получаем все error-уведомления
        let errorWhereClause = { type: 'error' };
        if (req.user.role === 'admin' && req.query.resId) {
          errorWhereClause.resId = parseInt(req.query.resId);
        } else if (req.user.role !== 'admin') {
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
        
        // Собираем аналитику
        let totalAllVl = 0;
        let totalProblemVl = 0;
        
        for (const resUnit of resListForWorkload) {
          const totalVl = await NetworkStructure.count({
            where: { resId: resUnit.id }
          });
          
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
  checkRole(['admin']), 
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
  checkRole(['admin']), 
  async (req, res) => {
    const transaction = await sequelize.transaction();
    
    try {
      const { ids, password, deleteDocuments } = req.body; // ✅ ДОБАВИЛИ deleteDocuments
      
      if (password !== DELETE_PASSWORD) {
        await transaction.rollback();
        return res.status(403).json({ error: 'Неверный пароль' });
      }
      
      if (!ids || !Array.isArray(ids) || ids.length === 0) {
        await transaction.rollback();
        return res.status(400).json({ error: 'Не выбраны записи для удаления' });
      }
      
      console.log('=== BULK DELETE NOTIFICATIONS ===');
      console.log('Notification IDs:', ids);
      console.log('Delete documents?', deleteDocuments);
      
      let deletedDocumentsCount = 0;
      let deletedFilesCount = 0;
      
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
          
          // Удаляем файлы из Cloudinary
          for (const doc of relatedDocs) {
            if (doc.attachments && doc.attachments.length > 0) {
              for (const file of doc.attachments) {
                try {
                  const isPdf = file.public_id.toLowerCase().endsWith('.pdf');
                  await cloudinary.uploader.destroy(file.public_id, {
                    resource_type: isPdf ? 'raw' : 'image'
                  });
                  deletedFilesCount++;
                  console.log(`✅ Deleted file: ${file.public_id}`);
                } catch (err) {
                  console.error(`❌ Error deleting file ${file.public_id}:`, err);
                }
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
      const deletedCount = await Notification.destroy({
        where: { id: { [Op.in]: ids } },
        transaction
      });
      
      await transaction.commit();
      
      console.log('=== DELETE COMPLETE ===');
      console.log(`Notifications: ${deletedCount}`);
      console.log(`Documents: ${deletedDocumentsCount}`);
      console.log(`Files: ${deletedFilesCount}`);
      
      res.json({
        success: true,
        message: deleteDocuments 
          ? `Удалено:\n• Уведомлений: ${deletedCount}\n• Документов: ${deletedDocumentsCount}\n• Файлов: ${deletedFilesCount}`
          : `Удалено уведомлений: ${deletedCount}`,
        deletedCount,
        deletedDocumentsCount,
        deletedFilesCount
      });
      
    } catch (error) {
      await transaction.rollback();
      console.error('Bulk delete notifications error:', error);
      res.status(500).json({ error: error.message });
    }
});

// роут ДЛЯ ОТЧЕТОВ эксель
app.get('/api/reports/export-history', 
  authenticateToken, 
  checkRole(['admin']), 
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
app.get('/api/users/list', authenticateToken, checkRole(['admin']), async (req, res) => {
  try {
    const users = await User.findAll({
      attributes: ['id', 'fio', 'login', 'role', 'resId', 'email', 'createdAt'],
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
app.post('/api/users/create', authenticateToken, checkRole(['admin']), async (req, res) => {
  try {
    const { fio, login, password, email, role, resId } = req.body;
    
    // Валидация
    if (!fio || !login || !password || !email || !role) {
      return res.status(400).json({ error: 'Все поля обязательны' });
    }
    
    if (password.length < 6) {
      return res.status(400).json({ error: 'Пароль должен содержать минимум 6 символов' });
    }
    
    if (role !== 'admin' && !resId) {
      return res.status(400).json({ error: 'Для не-админов необходимо указать РЭС' });
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
      resId: role === 'admin' ? null : resId
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
app.put('/api/users/:id', authenticateToken, checkRole(['admin']), async (req, res) => {
  try {
    const userId = req.params.id;
    const { fio, password, email, role, resId } = req.body;
    
    const user = await User.findByPk(userId);
    if (!user) {
      return res.status(404).json({ error: 'Пользователь не найден' });
    }
    
    // Обновляем только переданные поля
    const updateData = {};
    if (fio) updateData.fio = fio;
    if (email) updateData.email = email;
    if (role) updateData.role = role;
    if (role === 'admin') {
      updateData.resId = null;
    } else if (resId !== undefined) {
      updateData.resId = resId;
    }
    
    // Если передан новый пароль
    if (password && password.length > 0) {
      if (password.length < 6) {
        return res.status(400).json({ error: 'Пароль должен содержать минимум 6 символов' });
      }
      updateData.password = password;
    }
    
    await user.update(updateData);
    
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
app.delete('/api/users/:id', authenticateToken, checkRole(['admin']), async (req, res) => {
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
    
    // Проверяем, не последний ли это админ
    if (user.role === 'admin') {
      const adminCount = await User.count({ where: { role: 'admin' } });
      if (adminCount <= 1) {
        return res.status(400).json({ error: 'Нельзя удалить последнего администратора' });
      }
    }
    
    await user.destroy();
    
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
// API ДЛЯ ПРОБЛЕМНЫХ ВЛ
// =====================================================

// Получение списка проблемных ВЛ
app.get('/api/problem-vl/list', 
  authenticateToken, 
  checkRole(['admin']), 
  async (req, res) => {
    try {
      const where = { status: 'active' };
      if (req.query.resId) where.resId = parseInt(req.query.resId);
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

// Отклонение проблемы
app.put('/api/problem-vl/:id/dismiss', 
  authenticateToken, 
  checkRole(['admin']), 
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
  checkRole(['admin']), 
  async (req, res) => {
    try {
      const problemVL = await ProblemVL.findByPk(req.params.id, {
        include: [ResUnit]
      });
      
      if (!problemVL) {
        return res.status(404).json({ error: 'Проблема не найдена' });
      }
      
      // Находим ответственных за РЭС
      const responsibleUsers = await User.findAll({
        where: {
          resId: problemVL.resId,
          role: 'res_responsible'
        }
      });
      
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
    if (req.user.role === 'admin' && req.query.resId) {
      whereClause.resId = parseInt(req.query.resId);
    } else if (req.user.role !== 'admin') {
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
// Получение количества непрочитанных уведомлений
// Счётчики уведомлений по роли (переиспользуются эндпоинтом ниже и бейджем
// платформы /api/platform/badge). Возвращает { tech_pending, askue_pending,
// problem_vl }. user — { id, role, resId }.
async function getNotificationCounts(user) {
  let whereClause = {};

  if (user.role === 'admin') {
    whereClause = {};
  } else if (user.role === 'res_responsible') {
    whereClause = {
      resId: user.resId,
      [Op.or]: [
        { toUserId: null },
        { toUserId: user.id }
      ]
    };
  } else {
    whereClause = {
      toUserId: user.id
    };
  }

  // Для проблемных ВЛ считаем только активные (только для admin)
  let problemVLCount = 0;
  if (user.role === 'admin') {
    problemVLCount = await ProblemVL.count({
      where: { status: 'active' }  // Только активные!
    });
  }

  // ✅ PERF: считаем только те типы, которые реально участвуют в счётчиках
  whereClause.type = { [Op.in]: ['error', 'pending_askue'] };

  const allNotifications = await Notification.findAll({
    where: whereClause,
    attributes: ['id', 'type', 'message']
  });

  // ✅ PERF: один индексированный запрос по userId вместо гигантского IN(...)
  const readNotifications = await NotificationRead.findAll({
    where: { userId: user.id },
    attributes: ['notificationId']
  });

  const readIds = new Set(readNotifications.map(r => r.notificationId));

  // ✅ Считаем уникальные ПУ+фаза по типам (не общее количество уведомлений!)
  const techKeys = new Set();
  const askueKeys = new Set();

  allNotifications.forEach(notif => {
    if (!readIds.has(notif.id)) {
      try {
        const data = JSON.parse(notif.message);
        const puNumber = data.puNumber;

        if (puNumber) {
          const phaseKey = getPhaseSignature(notif.message);
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
    }
  });

  return {
    tech_pending: techKeys.size,
    askue_pending: askueKeys.size,
    problem_vl: problemVLCount
  };
}

app.get('/api/notifications/counts', authenticateToken, async (req, res) => {
  try {
    res.json(await getNotificationCounts(req.user));
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
    if (req.user.role === 'admin') {
      whereClause = type === 'all' ? {} : { type };
    } else if (req.user.role === 'res_responsible') {
      whereClause = {
        resId: req.user.resId,
        [Op.or]: [
          { toUserId: null },
          { toUserId: req.user.id }
        ]
      };
      if (type !== 'all') whereClause.type = type;
    } else {
      whereClause = { toUserId: req.user.id };
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
  checkRole(['admin']), 
  async (req, res) => {
    const transaction = await sequelize.transaction();
    
    try {
      const { ids, password } = req.body;
      
      if (password !== DELETE_PASSWORD) {
        return res.status(403).json({ error: 'Неверный пароль' });
      }
      
      if (!ids || !Array.isArray(ids) || ids.length === 0) {
        return res.status(400).json({ error: 'Не выбраны записи для удаления' });
      }
      
      // Получаем все записи для удаления
      const records = await CheckHistory.findAll({
        where: { id: { [Op.in]: ids } },
        transaction
      });
      
      // Удаляем все файлы из Cloudinary
      for (const record of records) {
        if (record.attachments && record.attachments.length > 0) {
          for (const file of record.attachments) {
            try {
              await cloudinary.uploader.destroy(file.public_id);
              console.log(`Deleted file from Cloudinary: ${file.public_id}`);
            } catch (err) {
              console.error('Error deleting file from Cloudinary:', err);
            }
          }
        }
      }
      
      // Удаляем записи из БД
      const deletedCount = await CheckHistory.destroy({
        where: { id: { [Op.in]: ids } },
        transaction
      });
      
      await transaction.commit();
      
      res.json({
        success: true,
        message: `Удалено записей: ${deletedCount}`,
        deletedCount
      });
      
    } catch (error) {
      await transaction.rollback();
      console.error('Bulk delete error:', error);
      res.status(500).json({ error: error.message });
    }
});

// =====================================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ АНАЛИЗА
// =====================================================

async function analyzeFile(filePath, type, originalFileName = null, requiredPeriod = null, userId = null) {
  return new Promise((resolve, reject) => {

    console.log('=== ANALYZE FILE DEBUG ===');
    console.log('Received userId:', userId);
    console.log('All params:', { filePath, type, originalFileName, requiredPeriod, userId });
    
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
        const currentYear = new Date().getFullYear();
        
        return {
          start: new Date(currentYear, firstMonth - 1, 1),
          end: new Date(currentYear, lastMonth - 1, 28)
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
    
    // Запуск Python
    let python;
    try {
      python = spawn('python3', [scriptPath, filePath]);
      console.log('Python3 spawn created successfully');
    } catch (err) {
      console.error('Failed to spawn python3, trying python:', err);
      try {
        python = spawn('python', [scriptPath, filePath]);
        console.log('Python spawn created successfully');
      } catch (err2) {
        console.error('Both python3 and python failed:', err2);
        return resolve({
          processed: [],
          errors: ['Python не установлен на сервере. Убедитесь что в Build Command есть: npm install && pip install xlrd']
        });
      }
    }

    console.log('Running Python script:', scriptPath);
    console.log('Analyzing file:', filePath);

    let output = '';
    let errorOutput = '';

    python.stdout.on('data', (data) => {
      output += data.toString();
      console.log('Python stdout chunk:', data.toString());
    });

    python.stderr.on('data', (data) => {
      errorOutput += data.toString();
      console.error('Python stderr:', data.toString());
    });

    python.on('error', (error) => {
      console.error('Python process error:', error);
      return resolve({
        processed: [],
        errors: [`Python не установлен или недоступен. Убедитесь что в Build Command на Render есть: npm install && pip install xlrd`]
      });
    });

    python.on('close', async (code) => {
      console.log('Python process closed with code:', code);
      
      if (code !== 0) {
        return resolve({
          processed: [],
          errors: [`Ошибка анализа (код ${code}): ${errorOutput}`]
        });
      }
      
      try {
        // Парсим результат от Python
        const result = JSON.parse(output);
        console.log('Parsed result:', JSON.stringify(result));
        
        if (result.success) {
          const processed = [];
          const errors = [];
          
          // Извлекаем номер ПУ из имени файла
          const fileName = originalFileName 
            ? path.basename(originalFileName, path.extname(originalFileName))
            : path.basename(filePath, path.extname(filePath));
          
          console.log('Extracted PU number from filename:', fileName);

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
    console.log(`Found same error in history for PU ${fileName}, checking if it was fixed...`);
    
    // Проверяем, была ли загрузка БЕЗ ошибок после этой ошибки
    const successfulUploadAfter = recentUploads.find(upload => 
      !upload.hasErrors && 
      new Date(upload.uploadedAt) > new Date(sameErrorUpload.uploadedAt)
    );
    
    if (successfulUploadAfter) {
      // Ошибка была исправлена, но теперь появилась снова
      console.log(`Error was fixed on ${successfulUploadAfter.uploadedAt} but now appeared again`);
      // Разрешаем загрузку - это повторное появление ошибки
    } else {
      // Ошибка не была исправлена - это дубликат
      console.log(`DUPLICATE: Same error still not fixed for PU ${fileName}`);
      
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
    console.log(`DUPLICATE: Active CheckHistory record exists for this error`);
    
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
            console.log(`Found network structure for PU ${fileName}: TP=${networkStructure.tpName}, VL=${networkStructure.vlName}`);
            
            // Определяем позицию
            let position = 'start';
            if (networkStructure.endPu === fileName) position = 'end';
            else if (networkStructure.middlePu === fileName) position = 'middle';
            
            console.log(`PU position: ${position}`);
            
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
              console.log(`This is a recheck for PU ${fileName}`);
              
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
                      console.log(`PERIOD MISMATCH: Required from month ${requiredMonth}, but journal ends at month ${lastErrorMonthNum}`);
    
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
              console.log('Deleted ASKUE notifications');
              
              // ОБРАБОТКА РЕЗУЛЬТАТА ПЕРЕПРОВЕРКИ
              if (!result.has_errors) {
                // УСПЕШНАЯ перепроверка
                console.log(`Recheck successful - errors fixed for PU ${fileName}`);
                
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
                  fromUserId: 1,
                  toUserId: null, // для всех ответственных РЭС
                  resId: networkStructure.resId,
                  networkStructureId: networkStructure.id,
                  type: 'success',
                  message: `✅ Проблема с ПУ ${fileName} (${networkStructure.tpName} - ${networkStructure.vlName}) успешно устранена!`,
                  isRead: false
                });
                
              } else {
                // НЕУСПЕШНАЯ перепроверка
                console.log(`Recheck failed - errors still present for PU ${fileName}`);
                
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
                  fromUserId: 1,
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
                console.log('Created error notification for RES after failed recheck');
                
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
                  const admins = await User.findAll({ where: { role: 'admin' } });
                  for (const admin of admins) {
                    await Notification.create({
                      fromUserId: 1,
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
                  console.log('Created problem VL notification for admins');
                }
              }
              
            } else {
              // НЕ ПЕРЕПРОВЕРКА - обычная проверка или повторная проверка
              
              // ✅ ИСПРАВЛЕНИЕ: Удаляем старые уведомления для этого ПУ с теми же фазами
              const newPhaseKey = getPhaseSignature(JSON.stringify({
                errorDetails: result.summary,
                details: result.details
              }));
              
              const oldNotifs = await Notification.findAll({
                where: {
                  type: 'error',
                  message: {
                    [Op.like]: `%"puNumber":"${fileName}"%`
                  }
                }
              });
              
              let deletedOldNotifs = 0;
              for (const oldNotif of oldNotifs) {
                const oldPhaseKey = getPhaseSignature(oldNotif.message);
                if (oldPhaseKey === newPhaseKey) {
                  await oldNotif.destroy();
                  deletedOldNotifs++;
                }
              }
              
              if (deletedOldNotifs > 0) {
                console.log(`🧹 Cleaned up ${deletedOldNotifs} old error notifications for PU ${fileName} (phases: ${newPhaseKey})`);
              }
              
              // Если ошибок НЕТ — удаляем ВСЕ старые error-уведомления для этого ПУ
              if (!result.has_errors && oldNotifs.length > deletedOldNotifs) {
                const remaining = oldNotifs.length - deletedOldNotifs;
                await Notification.destroy({
                  where: {
                    type: 'error',
                    message: { [Op.like]: `%"puNumber":"${fileName}"%` }
                  }
                });
                console.log(`🧹 PU ${fileName} clean — removed ${remaining} remaining error notifications`);
              }
              
              // Обновляем статус ПУ
              await PuStatus.upsert({
                puNumber: fileName,
                networkStructureId: networkStructure.id,
                position: position,
                status: result.has_errors ? 'checked_error' : 'checked_ok',
                errorDetails: result.has_errors ? result.summary : null,
                lastCheck: new Date()
              });

              // Успешная проверка (без ошибок) — автоматически снимаем ВЛ с
              // проблемных (даже если ранее было 2+ неудачных проверок).
              if (!result.has_errors) {
                await ProblemVL.update(
                  { status: 'resolved' },
                  { where: { puNumber: fileName, status: 'active' } }
                );
              }

              // Если есть ошибки - добавляем для создания уведомлений
              if (result.has_errors) {
                errors.push({
                  puNumber: fileName,
                  error: result.summary,
                  details: result.details,
                  networkStructureId: networkStructure.id,
                  resId: networkStructure.resId
                });
                console.log('Added error for notification creation');
              }
            }
            
            // Записываем успешную загрузку в историю
            if (userId) {
                console.log('=== CREATING PuUploadHistory ===');
                console.log('userId:', userId);
                console.log('Data to save:', {
                  puNumber: fileName,
                  uploadedBy: userId,
                  fileName: originalFileName || 'unknown',
                  fileType: type,
                  uploadStatus: 'success'
                });
  
                try {
                  const record = await PuUploadHistory.create({
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
                  console.log('✅ PuUploadHistory created:', record.id);
                } catch (error) {
                  console.error('❌ Error creating PuUploadHistory:', error);
                }
              } else {
                console.log('⚠️ No userId provided, skipping history save');
              }
            
            // Добавляем в processed
            processed.push({
              puNumber: fileName,
              status: result.has_errors ? 'checked_error' : 'checked_ok',
              error: result.has_errors ? result.summary : null
            });
            
          } else {
            // ПУ не найден в структуре сети
            console.log(`WARNING: NetworkStructure not found for PU: ${fileName}`);
            processed.push({
              puNumber: fileName,
              status: 'not_in_structure',
              error: 'ПУ не найден в структуре сети'
            });
          }
          
          // Удаляем временный файл
          try {
            fs.unlinkSync(filePath);
            console.log('Temporary file deleted');
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
    });
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
  checkRole(['admin']), 
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
      const stats = {
  totalIssues: issues.length,
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
  checkRole(['admin']), 
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
        fromUserId: 1, // Системное уведомление
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
app.get('/api/download/:public_id', async (req, res) => {
  try {
    const publicId = decodeURIComponent(req.params.public_id);
    const originalName = req.query.name || 'file';
    const inline = req.query.inline === '1';

    // Определяем тип ресурса (PDF хранятся как raw, остальное — image)
    const isPdf = publicId.toLowerCase().endsWith('.pdf');
    const resourceType = isPdf ? 'raw' : 'image';

    // Signed URL для доступа к приватным файлам — но идёт по нему СЕРВЕР
    const fileUrl = cloudinary.url(publicId, {
      resource_type: resourceType,
      type: 'upload',
      secure: true,
      sign_url: true
    });

    const upstream = await fetch(fileUrl);
    if (!upstream.ok || !upstream.body) {
      console.error('Download proxy: upstream status', upstream.status, 'for', publicId);
      return res.status(upstream.status === 404 ? 404 : 502)
        .json({ error: 'Файл недоступен в хранилище' });
    }

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
];

// Полный JSON-дамп всех таблиц. Пароли выгружаются как есть (bcrypt-хэши),
// ссылки Cloudinary едут внутри JSON — сами файлы переносить не нужно.
app.get('/api/admin/backup', authenticateToken, checkRole(['admin']), async (req, res) => {
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

app.post('/api/admin/restore', authenticateToken, checkRole(['admin']), backupUpload.single('file'), async (req, res) => {
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
      `CREATE UNIQUE INDEX IF NOT EXISTS idx_users_keycloak ON "Users" ("keycloakId")`
    ];
    for (const stmt of indexStatements) {
      try {
        await sequelize.query(stmt);
      } catch (idxErr) {
        console.warn('Index creation skipped:', idxErr.message);
      }
    }
    console.log('Performance indexes ensured');
    
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
    if (req.user.role === 'admin' && resId) {
      whereClause.resId = parseInt(resId);
    } else if (req.user.role !== 'admin') {
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
  checkRole(['admin']), 
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
  checkRole(['admin']), 
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
      
      // Удаляем все файлы из Cloudinary
      if (record.attachments && record.attachments.length > 0) {
        for (const file of record.attachments) {
          try {
            await cloudinary.uploader.destroy(file.public_id);
            console.log(`Deleted file from Cloudinary: ${file.public_id}`);
          } catch (err) {
            console.error('Error deleting file from Cloudinary:', err);
          }
        }
      }
      
      // Удаляем запись из БД
      await record.destroy();
      
      res.json({ 
        success: true, 
        message: 'Запись и все связанные файлы удалены' 
      });
      
    } catch (error) {
      console.error('Delete record error:', error);
      res.status(500).json({ error: error.message });
    }
});

// ИСПРАВЛЕННЫЙ эндпоинт для управления файлами в настройках
 
app.delete('/api/admin/files/:public_id', 
  authenticateToken, 
  checkRole(['admin']), 
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
initializeDatabase().then(() => {
  app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
    console.log(`Environment: ${process.env.NODE_ENV || 'development'}`);
    console.log('Features enabled:');
    console.log('- User Management ✓');
    console.log('- Phase Detection ✓'); 
    console.log('- Auto Updates ✓');
    console.log('- Auto Hide Notifications ✓');
  });
});

// Обработка ошибок
process.on('unhandledRejection', (err) => {
  console.error('Unhandled rejection:', err);
  process.exit(1);
});

// API для получения истории загрузок по конкретному ПУ
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
      if (req.user.role === 'admin' && resId) {
        structureWhere.resId = resId;
      } else if (req.user.role !== 'admin') {
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
      if (req.user.role === 'admin' && resId) {
        whereClause.resId = resId;
      } else if (req.user.role !== 'admin') {
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
  checkRole(['admin']), 
  async (req, res) => {
    const transaction = await sequelize.transaction();
    
    try {
      const { password } = req.body;
      const { puNumber } = req.params;
      
      if (password !== DELETE_PASSWORD) {
        return res.status(403).json({ error: 'Неверный пароль' });
      }
      
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
  checkRole(['admin']), 
  async (req, res) => {
    const transaction = await sequelize.transaction();
    
    try {
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
  checkRole(['admin']), 
  async (req, res) => {
    const transaction = await sequelize.transaction();
    
    try {
      const { password } = req.body;
      
      if (password !== DELETE_PASSWORD) {
        return res.status(403).json({ error: 'Неверный пароль' });
      }
      
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
      if (req.user.role !== 'admin') {
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
  checkRole(['admin']), 
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
  checkRole(['admin']), 
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
  checkRole(['admin']), 
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
  checkRole(['admin']), 
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
      if (req.user.role !== 'admin') {
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
          
          return {
            resId: res.id,
            resName: res.name,
            tpCount,
            vlCount,
            vlCoveragePercent,
            totalPuCount,
            uniqueCheckedPuCount,
            puCoveragePercent
          };
        })
      );
      
      // Итоги
      const totals = analytics.reduce((acc, curr) => ({
        tpCount: acc.tpCount + curr.tpCount,
        vlCount: acc.vlCount + curr.vlCount,
        totalPuCount: acc.totalPuCount + curr.totalPuCount,
        uniqueCheckedPuCount: acc.uniqueCheckedPuCount + curr.uniqueCheckedPuCount
      }), {
        tpCount: 0,
        vlCount: 0,
        totalPuCount: 0,
        uniqueCheckedPuCount: 0
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
