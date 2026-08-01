// =====================================================
// УЛУЧШЕННЫЙ FRONTEND ДЛЯ СИСТЕМЫ УПРАВЛЕНИЯ РЭС
// Файл: src/App.jsx
// Версия с исправленными фазами и загрузкой из АСКУЭ
// =====================================================

import React, { useState, useEffect, createContext, useContext, useCallback, useMemo, useRef } from 'react';
import axios from 'axios';
import './App.css';
// xlsx-js-style — drop-in форк SheetJS (тот же API, включая utils/read), плюс стили
// ячеек. Один импорт вместо двух ('xlsx' удалён из зависимостей — был дубль).
import * as XLSX from 'xlsx-js-style';
import { IconCheck, IconX, IconAlertTriangle, IconAlertCircle, IconZap, IconChart, IconSearch, IconEye, IconTrash, IconCalendar, IconWrench, IconClock, IconRefresh, IconClipboard, IconArrowRight, IconArrowLeft, IconArrowUp, IconArrowDown, IconEdit, IconMapPin, IconFileText, IconFolder, IconPaperclip, IconRocket, IconHelp, IconBell, IconLightbulb, IconLock, IconMegaphone, IconMail, IconBuilding, IconBroom, IconMessage, IconLayers, IconDownload, IconPlug, IconLink, IconDatabase, IconInfo, IconMeter, IconUpload, IconSettings, IconUsers } from './icons.jsx';
import RossetiLoader from './RossetiLoader.jsx';

// =====================================================
// НАСТРОЙКА API КЛИЕНТА
// =====================================================

const API_URL = import.meta.env.VITE_API_URL || '';

// base64url от строки (юникод-безопасно: btoa не умеет не-latin1). +/→-_, без '='.
// Прячем public_id (со словом 'attachment_') из URL — иначе блокировщики режут запрос.
const toBase64Url = (str) => {
  const bytes = new TextEncoder().encode(String(str));
  let bin = '';
  for (const b of bytes) bin += String.fromCharCode(b);
  return btoa(bin).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
};

// Все файлы (фото/PDF из Cloudinary) показываем и скачиваем ТОЛЬКО через свой
// бэкенд (/api/f) — прямые ссылки на res.cloudinary.com блокируются браузерами,
// а слова download/attachment_ и query-строка в пути режут блокировщики
// (ERR_BLOCKED_BY_CLIENT). Поэтому public_id+name+inline прячем в ОДИН base64url
// JSON-токен {p,n,i} — без query-строки. inline=true — открыть в браузере.
const fileProxyUrl = (file, inline = false) => {
  if (!file) return '';
  if (!file.public_id) return file.url; // старые записи без public_id — как было
  const token = toBase64Url(JSON.stringify({ p: file.public_id, n: file.original_name || 'file', i: inline ? 1 : 0 }));
  // Файловый токен в query (?t=): нужен для скачивания обычными <a>-ссылками (у них
  // нет Authorization). В сам JSON-токен НЕ вшиваем — так ссылки протухают с сессией.
  const ft = localStorage.getItem('fileToken');
  return `${API_URL}/api/f/${token}${ft ? `?t=${encodeURIComponent(ft)}` : ''}`;
};
// ── Оформление Excel-выгрузок (xlsx-js-style) ────────────────────────────────
// Аккуратно, «как в качественном ПО»: тёмно-синяя шапка с белым жирным текстом,
// тонкие светлые рамки, лёгкая зебра, цветной ТЕКСТ только у статусов/процентов.
const XLS_HEADER_FILL = '25476A';   // сине-стальной, не кричащий
const XLS_ZEBRA = 'F5F8FB';
const XLS_BORDER = { style: 'thin', color: { rgb: 'D9DEE5' } };
const XLS_ALL_BORDERS = { top: XLS_BORDER, bottom: XLS_BORDER, left: XLS_BORDER, right: XLS_BORDER };
const XLS_COLORS = { red: 'B91C1C', green: '15803D', amber: 'B45309', gray: '6B7280' };
const XLS_HEADER_STYLE = {
  font: { bold: true, sz: 11, color: { rgb: 'FFFFFF' } },
  fill: { fgColor: { rgb: XLS_HEADER_FILL } },
  alignment: { horizontal: 'center', vertical: 'center', wrapText: true },
  border: XLS_ALL_BORDERS
};
const xlsStatusColor = (val) => {
  const v = String(val == null ? '' : val).trim();
  if (v === 'Перегруз' || v === 'Ошибка') return XLS_COLORS.red;
  if (v === 'Норма' || v === 'Проверен') return XLS_COLORS.green;
  if (v === 'Ожидает перепроверки') return XLS_COLORS.amber;
  if (v === 'Нет данных' || v === 'Не проверен' || v === 'Пусто' || v === '—') return XLS_COLORS.gray;
  return null;
};
const xlsLoadColor = (val) => {
  const n = Number(val);
  if (!isFinite(n)) return XLS_COLORS.gray;
  if (n > 100) return XLS_COLORS.red;
  if (n >= 85) return XLS_COLORS.amber;
  return XLS_COLORS.green;
};
// Применяет оформление ко всему листу. colorFns: имя колонки → (значение)→rgb|null
// (цветной жирный текст). cellColor(dataRowIdx, colName, value)→rgb|null — для
// матриц, где цвет зависит не от значения ячейки (напр. % загрузки, а не сам кВт).
// Плюс автофильтр и высота шапки.
const styleExportSheet = (ws, colorFns = {}, cellColor = null) => {
  if (!ws['!ref']) return;
  const range = XLSX.utils.decode_range(ws['!ref']);
  const headerName = {};
  for (let c = range.s.c; c <= range.e.c; c++) {
    const cell = ws[XLSX.utils.encode_cell({ r: range.s.r, c })];
    if (cell) { cell.s = XLS_HEADER_STYLE; headerName[c] = String(cell.v); }
  }
  for (let r = range.s.r + 1; r <= range.e.r; r++) {
    const zebra = (r - range.s.r) % 2 === 0;
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r, c })];
      if (!cell) continue;
      const st = {
        font: { sz: 10.5, color: { rgb: '1F2937' } },
        alignment: { vertical: 'center' },
        border: XLS_ALL_BORDERS
      };
      if (zebra) st.fill = { fgColor: { rgb: XLS_ZEBRA } };
      const fn = colorFns[headerName[c]];
      if (fn) {
        const col = fn(cell.v);
        if (col) st.font = { sz: 10.5, bold: true, color: { rgb: col } };
        st.alignment = { vertical: 'center', horizontal: 'center' };
      }
      if (cellColor) {
        const col2 = cellColor(r - range.s.r - 1, headerName[c], cell.v);
        if (col2) {
          st.font = { sz: 10.5, bold: true, color: { rgb: col2 } };
          st.alignment = { vertical: 'center', horizontal: 'center' };
        }
      }
      cell.s = st;
    }
  }
  ws['!autofilter'] = { ref: ws['!ref'] };
  ws['!rows'] = [{ hpt: 22 }];
};

// Единый вход через платформу (SSO): origin платформы и признак встраивания в iframe.
const PLATFORM_ORIGIN = import.meta.env.VITE_PLATFORM_ORIGIN || 'https://sue-system-ashinoff.amvera.io';
const EMBEDDED = typeof window !== 'undefined' && window.self !== window.top;

const api = axios.create({
  baseURL: API_URL,
  headers: {
    'Content-Type': 'application/json'
  },
  timeout: 120000
});

// Добавляем токен к каждому запросу
api.interceptors.request.use(config => {
  const token = localStorage.getItem('token');
  if (token) {
    config.headers.Authorization = `Bearer ${token}`;
  }
  return config;
});

// Право на действие в UI: суперадмин видит всё; остальным (любой роли) право
// раздаётся через «Права доступа» и лежит в user.permissions.
const hasPerm = (user, key) => !!(user && (user.isSuper || (user.permissions && user.permissions[key])));

// Человеческий тост (для 403 «нет права» и т.п.) без внешних зависимостей.
function showToast(msg, kind = 'error') {
  try {
    const el = document.createElement('div');
    el.className = `app-toast app-toast--${kind}`;
    el.textContent = msg;
    document.body.appendChild(el);
    setTimeout(() => { el.classList.add('app-toast--out'); }, 2600);
    setTimeout(() => { el.remove(); }, 3050);
  } catch { /* noop */ }
}

// Обработка ошибок авторизации
api.interceptors.response.use(
  response => response,
  error => {
    if (error.response?.status === 401) {
      localStorage.removeItem('token');
      localStorage.removeItem('fileToken');
      // В iframe платформы не редиректим на логин (иначе перезагрузка окна на 401).
      if (!EMBEDDED) window.location.href = '/';
    }
    // 403 из-за отсутствия права — человеческий тост с названием права.
    if (error.response?.status === 403 && error.response.data?.permission) {
      showToast('Недостаточно прав: ' + (error.response.data.title || error.response.data.permission));
    }
    return Promise.reject(error);
  }
);

// =====================================================
// КОНТЕКСТ АВТОРИЗАЦИИ
// =====================================================

const AuthContext = createContext(null);

// =====================================================
// УНИВЕРСАЛЬНЫЙ КОМПОНЕНТ ЗАГРУЗКИ
// =====================================================

function LoadingSpinner({ type = 'default' }) {
  // Заставка загрузки — только анимация «РОССЕТИ» (буквы по очереди), без текста.
  if (type === 'inline') {
    return <div className="loading-inline"><RossetiLoader size="small" /></div>;
  }
  const body = <div className="loading-container"><RossetiLoader /></div>;
  return type === 'overlay' ? <div className="loading-overlay">{body}</div> : body;
}


// =====================================================
// КОМПОНЕНТ АВТОРИЗАЦИИ
// =====================================================

function LoginForm({ onLogin }) {
  const [credentials, setCredentials] = useState({ login: '', password: '' });
  const [error, setError] = useState('');
  const [loading, setLoading] = useState(false);

  const handleSubmit = async (e) => {
    e.preventDefault();
    setLoading(true);
    setError('');
    
    try {
      const response = await api.post('/api/auth/login', credentials);
      localStorage.setItem('token', response.data.token);
      if (response.data.fileToken) localStorage.setItem('fileToken', response.data.fileToken);
      onLogin(response.data.user);
    } catch (err) {
      setError(err.response?.data?.error || 'Ошибка входа');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="login-container">
      <div className="login-box">
        <h2>Вход в систему контроля уровня напряжения</h2>
        <form onSubmit={handleSubmit}>
          <div className="form-group">
            <label>Логин</label>
            <input
              type="text"
              value={credentials.login}
              onChange={(e) => setCredentials({...credentials, login: e.target.value})}
              required
            />
          </div>
          <div className="form-group">
            <label>Пароль</label>
            <input
              type="password"
              value={credentials.password}
              onChange={(e) => setCredentials({...credentials, password: e.target.value})}
              required
            />
          </div>
          {error && <div className="error-message">{error}</div>}
          <button type="submit" className="btn-default" disabled={loading}>
            {loading ? 'Вход...' : 'Войти'}
          </button>
        </form>
      </div>
    </div>
  );
}

// =====================================================
// ГЛАВНОЕ МЕНЮ
// =====================================================

// Номер секции шин римскими (1..5 → I..V); прочее — как есть.
const ROMAN = { 1: 'I', 2: 'II', 3: 'III', 4: 'IV', 5: 'V' };
const toRoman = (n) => ROMAN[n] || String(n);

// Высокочастотная синусоида (осциллограф), центрированная по средней линии
// (viewBox 0 0 120 44, mid=22). Узкая амплитуда — колебания идут ровно в зазоре
// между словами «МОНИТОРИНГ» и «напряжения», не вываливаясь волнами вниз.
const OSC_PATH = (() => {
  const w = 120, mid = 22, amp = 14, periods = 9, n = 300;
  let d = '';
  for (let i = 0; i <= n; i++) {
    const x = (w * i / n).toFixed(2);
    const y = (mid - amp * Math.sin(periods * 2 * Math.PI * i / n)).toFixed(2);
    d += (i === 0 ? 'M' : 'L') + x + ' ' + y + ' ';
  }
  return d.trim();
})();

// Логотип «Мониторинг напряжения» с платформы (ResmTile: зелёно-бирюзовая плитка,
// белая «пульс»-линия). Инлайн-SVG.
const RESM_LOGO = (
  <svg viewBox="0 0 512 512" width="34" height="34" aria-hidden="true">
    <defs>
      <linearGradient id="resmGradSb" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0" stopColor="#22c55e" />
        <stop offset="1" stopColor="#0e7490" />
      </linearGradient>
    </defs>
    <rect width="512" height="512" rx="112" fill="url(#resmGradSb)" />
    <g transform="translate(88 88) scale(14)" fill="none" stroke="#fff" strokeWidth="2.1" strokeLinecap="round" strokeLinejoin="round">
      <path d="M22 12h-4l-3 9L9 3l-3 9H2" />
    </g>
  </svg>
);

function MainMenu({ activeSection, onSectionChange, userRole, isSuper }) {
  const [notificationCounts, setNotificationCounts] = useState({
    tech_pending: 0,
    askue_pending: 0,
    problem_vl: 0
  });
  // Тик «пробега» осциллограммы — инкремент по каждому успешному поллу counts.
  const [sweepTick, setSweepTick] = useState(0);

  // Загружаем количество уведомлений
  useEffect(() => {
    loadNotificationCounts();
    
    // PERF: не опрашиваем сервер, пока вкладка скрыта
    const interval = setInterval(() => {
      if (!document.hidden) loadNotificationCounts();
    }, 30000); // Обновляем каждые 30 сек
    
    // При возврате на вкладку — обновляем сразу
    const handleVisibility = () => {
      if (!document.hidden) loadNotificationCounts();
    };
    document.addEventListener('visibilitychange', handleVisibility);
    
    // Слушаем события обновления
    const handleUpdate = () => loadNotificationCounts();
    window.addEventListener('notificationsUpdated', handleUpdate);
    
    return () => {
      clearInterval(interval);
      document.removeEventListener('visibilitychange', handleVisibility);
      window.removeEventListener('notificationsUpdated', handleUpdate);
    };
  }, []);

  const loadNotificationCounts = async () => {
    try {
      const response = await api.get('/api/notifications/counts');
      setNotificationCounts(response.data);
      // Однократный «пробег» осциллограммы на факт успешного ответа (без таймеров).
      setSweepTick(t => t + 1);
    } catch (error) {
      console.error('Error loading notification counts:', error);
    }
  };

  // Красный, если есть активные error (tech_pending) или перегрузы Pном; иначе зелёный.
  const oscDanger =
    (notificationCounts.tech_pending || 0) +
    (notificationCounts.powerOverload || 0) +
    (notificationCounts.power_overload || 0) > 0;

  const menuItems = [
    { id: 'structure', label: 'Структура сети', icon: <IconLayers size={18} />, roles: ['admin', 'uploader', 'res_responsible'] },
    { id: 'poll_map', label: 'Карта опроса', icon: <IconMapPin size={18} />, roles: ['admin', 'uploader', 'res_responsible', 'uec_responsible'] },
    { id: 'upload', label: 'Загрузить файлы', icon: <IconUpload size={18} />, roles: ['admin', 'uploader'] },
    { id: 'tech_pending', label: 'Ожидающие мероприятий', icon: <IconWrench size={18} />, roles: ['admin', 'res_responsible'], badge: notificationCounts.tech_pending },
    { id: 'askue_pending', label: 'Ожидающие проверки АСКУЭ', icon: <IconClipboard size={18} />, roles: ['admin', 'uploader'], badge: notificationCounts.askue_pending },
    { id: 'problem_vl', label: 'Проблемные ВЛ', icon: <IconAlertTriangle size={18} />, roles: ['admin'], badge: notificationCounts.problem_vl },
    { id: 'power_overload', label: 'Превышение Pном', icon: <IconZap size={18} />, roles: ['admin', 'uploader', 'res_responsible'], badge: notificationCounts.powerOverload },
    { id: 'power_analysis', label: 'Анализ мощности', icon: <IconChart size={18} />, roles: ['admin', 'res_responsible'] },
    { id: 'analytics', label: 'Анализ напряжения', icon: <IconChart size={18} />, roles: ['admin', 'uploader', 'res_responsible'] },
    { id: 'deep_analysis', label: 'Глубокий анализ работ', icon: <IconSearch size={18} />, roles: ['admin'] },
    { id: 'documents', label: 'Загруженные документы', icon: <IconFolder size={18} />, roles: ['admin', 'uploader', 'res_responsible'] },
    { id: 'history', label: 'История системы', icon: <IconClock size={18} />, roles: ['admin', 'uploader', 'res_responsible'] },
    { id: 'reports', label: 'Отчеты', icon: <IconFileText size={18} />, roles: ['admin', 'uploader', 'res_responsible'] },
    { id: 'settings', label: 'Настройки', icon: <IconSettings size={18} />, roles: ['admin'] },
    { id: 'permissions', label: 'Права доступа', icon: <IconLock size={18} />, roles: ['admin'], superOnly: true }
  ];

  const visibleItems = menuItems.filter(item => item.roles.includes(userRole) && (!item.superOnly || isSuper));

  return (
    <nav className="main-menu">
      <div className="monitor-head">
        <div className="monitor-brand">
          {/* Сетка осциллографа (тонкие белые квадраты), отцентрована по средней
              линии между словами. Базовая — едва заметная; на пробеге вспыхивает. */}
          <div className={`osc-grid ${oscDanger ? 'danger' : 'ok'}`} aria-hidden="true" />
          {sweepTick > 0 && (
            <div key={`g${sweepTick}`} className={`osc-grid-flash ${oscDanger ? 'danger' : 'ok'}`}
                 aria-hidden="true" />
          )}
          <svg className={`osc-bg ${oscDanger ? 'danger' : 'ok'}`} viewBox="0 0 120 44"
               preserveAspectRatio="none" aria-hidden="true">
            {sweepTick > 0 && (
              <path key={sweepTick} className="osc-spark" d={OSC_PATH} fill="none" pathLength="1" />
            )}
          </svg>
          <div className="monitor-title">
            <span className="monitor-logo">{RESM_LOGO}</span>
            <span className="monitor-text">
              <span className="mt-1">МОНИТОРИНГ</span>
              <span className="mt-2">напряжения</span>
            </span>
          </div>
        </div>
      </div>
      {visibleItems.map(item => (
        <button
          key={item.id}
          onClick={() => onSectionChange(item.id)}
          className={`menu-item ${activeSection === item.id ? 'active' : ''}`}
        >
          <span className="menu-ico">{item.icon}</span>
          <span className="menu-label">{item.label}</span>
          {item.badge > 0 && (
            <span className="notification-badge">{item.badge > 99 ? '99+' : item.badge}</span>
          )}
        </button>
      ))}
    </nav>
  );
}

// Единая шапка страницы: SVG-иконка в квадратной плашке (статичная неоново-синяя
// подсветка, БЕЗ анимаций — инвариант) + заголовок. Одинаковые отступы на всех страницах.
function PageHeader({ icon, title }) {
  return (
    <div className="page-header">
      <span className="page-header-icon">{icon}</span>
      <h2 className="page-header-title">{title}</h2>
    </div>
  );
}

// ===== Оконная оболочка модалок: mac-«светофор» + сворачивание в док + «одно активное окно» =====
const ModalDockCtx = createContext(null);

function ModalDockProvider({ children }) {
  const [chips, setChips] = useState([]);          // видимые плашки дока [{id,title,icon}]
  const regs = useRef(new Map());                  // id -> { minimize, restore, close }
  const activeId = useRef(null);                    // текущее развёрнутое обычное окно

  const api = useMemo(() => ({
    register(id, handlers) { regs.current.set(id, handlers); },
    unregister(id) {
      regs.current.delete(id);
      if (activeId.current === id) activeId.current = null;
      setChips(cs => cs.filter(c => c.id !== id));   // мёртвых плашек не оставляем
    },
    // Окно стало активным (открыто/развёрнуто): свернуть предыдущее активное.
    setActive(id) {
      const prev = activeId.current;
      if (prev && prev !== id) { const h = regs.current.get(prev); if (h) h.minimize(); }
      activeId.current = id;
      setChips(cs => cs.filter(c => c.id !== id));
    },
    clearActive(id) { if (activeId.current === id) activeId.current = null; },
    toDock(id, chip) { setChips(cs => cs.some(c => c.id === id) ? cs : [...cs, chip]); },
  }), []);

  return (
    <ModalDockCtx.Provider value={api}>
      {children}
      <ModalDockBar chips={chips} regs={regs} />
    </ModalDockCtx.Provider>
  );
}

function ModalDockBar({ chips, regs }) {
  if (!chips.length) return null;
  const call = (id, k) => { const h = regs.current.get(id); if (h && h[k]) h[k](); };
  return (
    <div className="modal-dock">
      {chips.map(c => (
        <div className="modal-dock-chip" key={c.id}>
          <span className="modal-dock-ico">{c.icon}</span>
          <span className="modal-dock-title" title={c.title}>{c.title}</span>
          <button className="modal-dock-btn" title="Развернуть" onClick={() => call(c.id, 'restore')}><IconArrowUp className="ico" /></button>
          <button className="modal-dock-btn" title="Закрыть" onClick={() => call(c.id, 'close')}><IconX className="ico" /></button>
        </div>
      ))}
    </div>
  );
}

// Контекст оболочки — детям (напр. FileViewer) для реакции на fullscreen.
const ModalShellCtx = createContext({ fullscreen: false });

let __modalSeq = 0;
// title — узел заголовка; dockTitle — строка для плашки дока (если title не строка).
function ModalShell({ title, dockTitle, titleExtra, icon, onClose, onBackdrop, variant = 'normal', className = '', headerClassName = '', children }) {
  const dock = useContext(ModalDockCtx);
  const id = useRef('m' + (++__modalSeq)).current;
  const isConfirm = variant === 'confirm';
  const [minimized, setMinimized] = useState(false);
  const [fullscreen, setFullscreen] = useState(false);
  const onCloseRef = useRef(onClose); onCloseRef.current = onClose;
  const chipTitle = dockTitle || (typeof title === 'string' ? title : 'Окно');

  const minimize = () => {
    if (isConfirm || !dock) return;
    setMinimized(true);
    dock.toDock(id, { id, title: chipTitle, icon });
    dock.clearActive(id);
  };

  // Регистрация + политика «одно окно» (только обычные модалки, не confirm).
  useEffect(() => {
    if (isConfirm || !dock) return;
    dock.register(id, {
      minimize: () => { setMinimized(true); dock.toDock(id, { id, title: chipTitle, icon }); },
      restore: () => { setMinimized(false); dock.setActive(id); },
      close: () => { onCloseRef.current && onCloseRef.current(); },
    });
    dock.setActive(id);            // при открытии стать активной (свернёт предыдущую)
    return () => dock.unregister(id);
    // eslint-disable-next-line
  }, []);

  // Esc закрывает активную (не свёрнутую) модалку.
  useEffect(() => {
    if (minimized) return;
    const h = (e) => { if (e.key === 'Escape') { e.stopPropagation(); onCloseRef.current && onCloseRef.current(); } };
    window.addEventListener('keydown', h);
    return () => window.removeEventListener('keydown', h);
  }, [minimized]);

  if (minimized) return null;

  const handleBackdrop = () => {
    if (onBackdrop === null) return;                  // явно отключено
    const fn = onBackdrop || onClose;
    if (fn) fn();
  };

  return (
    <ModalShellCtx.Provider value={{ fullscreen }}>
      <div className={`modal-backdrop ${isConfirm ? 'modal-backdrop--confirm' : ''}`} onClick={handleBackdrop}>
        <div className={`modal-content modal-shell ${fullscreen ? 'modal-shell--fs' : ''} ${className}`} onClick={(e) => e.stopPropagation()}>
          <div className={`modal-shell-bar ${headerClassName}`}>
            <div className="traffic-lights">
              <button className="tl tl-red" title="Закрыть" onClick={() => onCloseRef.current && onCloseRef.current()}><IconX className="tl-ico" /></button>
              {!isConfirm && <button className="tl tl-yellow" title="Свернуть в док" onClick={minimize}><span className="tl-min" /></button>}
              {!isConfirm && <button className="tl tl-green" title={fullscreen ? 'Обычный размер' : 'Во весь экран'} onClick={() => setFullscreen(f => !f)}><span className="tl-fs" /></button>}
            </div>
            <div className="modal-shell-title">{title}</div>
            {titleExtra && <div className="modal-shell-extra">{titleExtra}</div>}
          </div>
          <div className="modal-shell-body">{children}</div>
        </div>
      </div>
    </ModalShellCtx.Provider>
  );
}

// ── SVG-значки колонок АСКУЭ в карточке секции («Структура сети») ──
// Свечение — через класс .ico-glow-blue (как у «редактировать»); wrapper .askue-ico.
const IconAskueLimited = ({ size = 18 }) => (   // «P» + wi-fi (введённые ограничения)
  <svg viewBox="0 0 24 24" width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <text x="1.5" y="17" fontSize="13" fontWeight="700" fill="currentColor" stroke="none">P</text>
    <path d="M12.5 8.5a7 7 0 0 1 9 0" />
    <path d="M14.5 11.5a4 4 0 0 1 5 0" />
    <circle cx="17" cy="15" r="1" fill="currentColor" stroke="none" />
  </svg>
);
const IconAskueUnpolled = ({ size = 18 }) => (  // «P» + перечёркнутый wi-fi (нет опроса)
  <svg viewBox="0 0 24 24" width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <text x="1.5" y="17" fontSize="13" fontWeight="700" fill="currentColor" stroke="none">P</text>
    <path d="M12.5 8.5a7 7 0 0 1 9 0" />
    <path d="M14.5 11.5a4 4 0 0 1 5 0" />
    <circle cx="17" cy="15" r="1" fill="currentColor" stroke="none" />
    <line x1="11.5" y1="6" x2="22.5" y2="17" />
  </svg>
);
const IconAskueSum = ({ size = 18 }) => (       // «ΣP» (общая разрешённая мощность)
  <svg viewBox="0 0 24 24" width={size} height={size} fill="none" stroke="none">
    <text x="12" y="17" fontSize="12" fontWeight="700" fill="currentColor" textAnchor="middle">ΣP</text>
  </svg>
);

// ── Значки этапов кейса перегруза (тело карточки «Превышение Pном») ──
const IconStageLimit = ({ size = 24 }) => (     // манометр — ввод ограничения мощности (АСКУЭ)
  <svg viewBox="0 0 24 24" width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M3.5 17a8.5 8.5 0 0 1 17 0" />
    <line x1="12" y1="17" x2="16.5" y2="11" />
    <circle cx="12" cy="17" r="1.7" fill="currentColor" stroke="none" />
    <line x1="12" y1="8.5" x2="12" y2="10" /><line x1="6" y1="17" x2="7.3" y2="17" /><line x1="18" y1="17" x2="16.7" y2="17" />
  </svg>
);
const IconStageRecheck = ({ size = 24 }) => (   // столбцы профиля + галочка — перепроверка профилем
  <svg viewBox="0 0 24 24" width={size} height={size} fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <line x1="4" y1="20" x2="4" y2="13" /><line x1="9" y1="20" x2="9" y2="6" /><line x1="14" y1="20" x2="14" y2="15" />
    <path d="M15.5 10.5l2.2 2.2L22 8.4" />
  </svg>
);

// =====================================================
// КОМПОНЕНТ СТРУКТУРЫ СЕТИ
// =====================================================

function NetworkStructure({ onSectionChange } = {}) {
  const [networkData, setNetworkData] = useState([]);
  const [techModal, setTechModal] = useState(null); // секция для «Сведения о техучёте»
  const [techCase, setTechCase] = useState(null);   // активный кейс секции (если есть)
  const [loading, setLoading] = useState(true);
  const [searchTp, setSearchTp] = useState('');
  const { user, selectedRes } = useContext(AuthContext);
  const canEditStructure = hasPerm(user, 'structure_edit');   // секции/ВЛ/удаление/привязка
  const canClearChecks = hasPerm(user, 'checks_delete');      // очистка истории по ТП
  // Правка номеров ПУ: полное право structure_edit ИЛИ узкое structure_edit_pu
  // (раздаётся любой роли через «Права доступа»). Секции — только canEditStructure.
  const canEditPu = canEditStructure || hasPerm(user, 'structure_edit_pu');
  const [modalOpen, setModalOpen] = useState(false);
  const [selectedDetails, setSelectedDetails] = useState(null);
  const [selectedItem, setSelectedItem] = useState(null);
  const [selectedPosition, setSelectedPosition] = useState(null);
  const [showExtendedModal, setShowExtendedModal] = useState(false);
  const [selectedPuData, setSelectedPuData] = useState(null);
  const [activeTab, setActiveTab] = useState('current'); // current, uploads, checks
  const [uploadHistory, setUploadHistory] = useState([]);
  const [checkHistory, setCheckHistory] = useState([]);
  const [historyLoading, setHistoryLoading] = useState(false);
  const [showClearHistoryModal, setShowClearHistoryModal] = useState(false);
  const [clearHistoryPassword, setClearHistoryPassword] = useState('');
  const [clearHistoryType, setClearHistoryType] = useState(''); // 'pu', 'tp', 'all'
  const [clearHistoryPu, setClearHistoryPu] = useState('');
  const [showScrollTop, setShowScrollTop] = useState(false);
  const [legendCollapsed, setLegendCollapsed] = useState(false); // полоса фильтров → кружки при скролле
  const [statusFilter, setStatusFilter] = useState(null);
  
  // Для редактирования ПУ (mac-модалка: { item, position, oldValue })
  const [puEdit, setPuEdit] = useState(null);
  const [editValue, setEditValue] = useState('');
  
  // Для выбора и удаления
  const [selectedIds, setSelectedIds] = useState([]);
  const [showDeleteModal, setShowDeleteModal] = useState(false);
  const [deletePassword, setDeletePassword] = useState('');

  // Секции шин ТП (этап 1)
  const [sections, setSections] = useState([]);
  const [sectionModal, setSectionModal] = useState(null); // { tpName, resId, id?, ... } | null
  const [sectionForm, setSectionForm] = useState({ sectionNumber: '', tnKva: '', cosPhi: '0.9', techPuNumber: '' });

  // Используем переданный selectedRes, если нет - берем из контекста
  
  
  // Оптимизированная функция загрузки
  const loadNetworkStructure = useCallback(async () => {
    try {
      const response = await api.get(`/api/network/structure/${selectedRes || ''}`);
      setNetworkData(response.data);
    } catch (error) {
      console.error('Error loading network structure:', error);
    } finally {
      setLoading(false);
    }
  }, [selectedRes]);

  // Загрузка секций шин выбранного РЭСа
  const loadSections = useCallback(async () => {
    try {
      const response = await api.get('/api/network/sections', { params: { resId: selectedRes || undefined } });
      setSections(response.data);
    } catch (error) {
      console.error('Error loading sections:', error);
      setSections([]);
    }
  }, [selectedRes]);

  useEffect(() => {
    loadSections();
  }, [loadSections]);

  // Открыть форму секции (создание/редактирование)
  const openSectionModal = (tpName, resId, section = null) => {
    if (section) {
      setSectionForm({
        sectionNumber: String(section.sectionNumber ?? ''),
        tnKva: section.tnKva != null ? String(section.tnKva) : '',
        cosPhi: section.cosPhi != null ? String(section.cosPhi) : '0.9',
        techPuNumber: section.techPuNumber || ''
      });
      setSectionModal({ ...section, tpName, resId });
    } else {
      setSectionForm({ sectionNumber: '', tnKva: '', cosPhi: '0.9', techPuNumber: '' });
      setSectionModal({ tpName, resId });
    }
  };

  const saveSection = async () => {
    try {
      const payload = {
        sectionNumber: sectionForm.sectionNumber,
        tnKva: sectionForm.tnKva,
        cosPhi: sectionForm.cosPhi,
        techPuNumber: sectionForm.techPuNumber
      };
      if (sectionModal.id) {
        await api.put(`/api/network/sections/${sectionModal.id}`, payload);
      } else {
        await api.post('/api/network/sections', {
          ...payload,
          resId: sectionModal.resId,
          tpName: sectionModal.tpName
        });
      }
      setSectionModal(null);
      await loadSections();
    } catch (error) {
      alert('Ошибка: ' + (error.response?.data?.error || error.message));
    }
  };

  const deleteSection = async (section) => {
    if (!window.confirm(`Удалить секцию СШ-${toRoman(section.sectionNumber)}?`)) return;
    try {
      await api.delete(`/api/network/sections/${section.id}`);
      await loadSections();
    } catch (error) {
      alert('Ошибка: ' + (error.response?.data?.error || error.message));
    }
  };

  // Привязка/отвязка ВЛ к секции
  // Привязка/перепривязка/отвязка ВЛ к секции. Оптимистично двигаем ВЛ между
  // группами локально (без перезагрузки), при ошибке — откат + сообщение.
  const assignSection = async (item, rawSectionId) => {
    const newSectionId = (rawSectionId === '' || rawSectionId == null) ? null : parseInt(rawSectionId, 10);
    const prev = item.sectionId ?? null;
    if (newSectionId === prev) return;

    setNetworkData(data => data.map(d => d.id === item.id ? { ...d, sectionId: newSectionId } : d));
    try {
      await api.put(`/api/network/structure/${item.id}`, {
        startPu: item.startPu,
        middlePu: item.middlePu,
        endPu: item.endPu,
        sectionId: newSectionId
      });
    } catch (error) {
      // Откат локального перемещения.
      setNetworkData(data => data.map(d => d.id === item.id ? { ...d, sectionId: prev } : d));
      alert('Ошибка: ' + (error.response?.data?.error || error.message));
    }
  };

  // Клик по квадрату техучёта секции → модалка сведений + подтягиваем активный кейс.
  const openTechModal = async (section) => {
    setTechModal(section);
    setTechCase(null);
    try {
      const { data } = await api.get('/api/overload', { params: { resId: section.resId } });
      const active = (data || []).find(c => c.sectionId === section.id && c.stage !== 'completed');
      if (active) setTechCase(active);
    } catch (e) { /* без кейса — не критично */ }
  };

  useEffect(() => {
    loadNetworkStructure();

    // Слушаем события обновления
    const handleUpdate = () => loadNetworkStructure();

    window.addEventListener('structureUpdated', handleUpdate);
    window.addEventListener('dataCleared', handleUpdate);
    window.addEventListener('structureDeleted', handleUpdate);

    // Тихое автообновление раз в 60 сек — чтобы изменения из почтового приёмника
    // (пик/перегруз секций) появлялись без ручной перезагрузки. Без спиннера
    // (loadNetworkStructure/loadSections не трогают loading), и только когда
    // вкладка активна. При возврате на вкладку — перезапрос сразу.
    const refresh = () => { loadNetworkStructure(); loadSections(); };
    const interval = setInterval(() => { if (!document.hidden) refresh(); }, 60000);
    const handleVisibility = () => { if (!document.hidden) refresh(); };
    document.addEventListener('visibilitychange', handleVisibility);

    return () => {
      window.removeEventListener('structureUpdated', handleUpdate);
      window.removeEventListener('dataCleared', handleUpdate);
      window.removeEventListener('structureDeleted', handleUpdate);
      clearInterval(interval);
      document.removeEventListener('visibilitychange', handleVisibility);
    };
  }, [loadNetworkStructure, loadSections]);

  useEffect(() => {
    // .content может ещё не быть в DOM на момент маунта — ждём через rAF,
    // иначе слушатель не привяжется. Здесь же — сворачивание полосы фильтров в кружки.
    let contentElement = null;
    let rafId = null;
    let ticking = false;
    const COLLAPSE_AT = 130; // ≈ высота полосы + запас
    const EXPAND_AT = 80;    // гистерезис ~50px — на границе не мигает
    const evaluate = () => {
      ticking = false;
      if (!contentElement) return;
      const top = contentElement.scrollTop;
      setShowScrollTop(top > 300);
      // Меняем состояние ТОЛЬКО при пересечении порога (функц. апдейт — React
      // бэйлит при том же значении → без ре-рендеров на каждый пиксель).
      setLegendCollapsed(prev => (!prev && top > COLLAPSE_AT) ? true : (prev && top < EXPAND_AT) ? false : prev);
    };
    const handleScroll = () => {
      if (!ticking) { ticking = true; requestAnimationFrame(evaluate); }
    };
    const attach = () => {
      contentElement = document.querySelector('.content');
      if (contentElement) {
        contentElement.addEventListener('scroll', handleScroll, { passive: true });
        evaluate();
      } else {
        rafId = requestAnimationFrame(attach);
      }
    };
    attach();
    return () => {
      if (rafId) cancelAnimationFrame(rafId);
      if (contentElement) contentElement.removeEventListener('scroll', handleScroll);
    };
  }, []);

  const getStatusColor = (status) => {
    switch(status) {
      case 'checked_ok': return 'status-ok';
      case 'checked_error': return 'status-error';
      case 'not_checked': return 'status-unchecked';
      case 'pending_recheck': return 'status-pending';
      case 'empty': return 'status-empty';
      default: return 'status-empty';
    }
  };

  const handleCellClick = (item, position) => {
  const puNumber = position === 'start' ? item.startPu : 
                   position === 'middle' ? item.middlePu : 
                   item.endPu;
  
  if (puNumber && item.PuStatuses) {
    const status = item.PuStatuses.find(s => 
      s.puNumber === puNumber && s.position === position
    );
    
    // НОВАЯ ЛОГИКА - открываем расширенное окно для всех
    setSelectedPuData({
      puNumber,
      position,
      tpName: item.tpName,
      vlName: item.vlName,
      resName: item.ResUnit?.name,
      status: status || { status: 'not_checked' },
      item
    });
    setShowExtendedModal(true);
    setActiveTab('current');
    loadPuHistory(puNumber);
  }
};
  
  // Функция загрузки истории
  const loadPuHistory = async (puNumber) => {
    setHistoryLoading(true);
    try {
      const [uploadsRes, checksRes] = await Promise.all([
        api.get(`/api/history/uploads/${puNumber}`),
        api.get(`/api/history/checks/${puNumber}`)
      ]);
      
      setUploadHistory(uploadsRes.data);
      setCheckHistory(checksRes.data);
    } catch (error) {
      console.error('Error loading PU history:', error);
    } finally {
      setHistoryLoading(false);
    }
  };

      
      
  
  // Начать редактирование — открываем mac-модалку (старый ПУ / новый ПУ).
  const startEdit = (item, position) => {
    if (!canEditPu) return;   // гейт по ПРАВУ (structure_edit / structure_edit_pu), не по роли
    const currentValue = position === 'start' ? item.startPu :
                        position === 'middle' ? item.middlePu :
                        item.endPu;
    setPuEdit({ item, position, oldValue: currentValue || '' });
    setEditValue(currentValue || '');
  };

  // Сохранить изменения ПУ из модалки
  const saveEdit = async () => {
    if (!puEdit) return;
    const { item, position } = puEdit;
    try {
      const updateData = {
        startPu: item.startPu,
        middlePu: item.middlePu,
        endPu: item.endPu
      };
      if (position === 'start') updateData.startPu = editValue || null;
      if (position === 'middle') updateData.middlePu = editValue || null;
      if (position === 'end') updateData.endPu = editValue || null;

      await api.put(`/api/network/structure/${item.id}`, updateData);
      await loadNetworkStructure();
      setPuEdit(null);
      setEditValue('');
    } catch (error) {
      showToast(error.response?.data?.error || 'Ошибка при сохранении');
    }
  };

  const cancelEdit = () => {
    setPuEdit(null);
    setEditValue('');
  };
  
  // Обработка выбора строк
  const handleSelectRow = (id) => {
    setSelectedIds(prev => {
      if (prev.includes(id)) {
        return prev.filter(i => i !== id);
      } else {
        return [...prev, id];
      }
    });
  };
  
  const handleSelectAll = () => {
    if (selectedIds.length === filteredData.length) {
      setSelectedIds([]);
    } else {
      setSelectedIds(filteredData.map(item => item.id));
    }
  };
  
  // Удаление выбранных с автообновлением
  const handleDeleteSelected = async () => {
    try {
      const response = await api.post('/api/network/delete-selected', {
        ids: selectedIds,
        password: deletePassword
      });
    
      alert(response.data.message);
      setShowDeleteModal(false);
      setDeletePassword('');
      setSelectedIds([]);
      setSearchTp(''); // Очищаем поле поиска!
    
      // Автообновление
      await loadNetworkStructure();
    
    } catch (error) {
      alert('Ошибка удаления: ' + (error.response?.data?.error || error.message));
    }
  };

// Функция очистки истории ПУ
const handleClearPuHistory = async (puNumber) => {
  setClearHistoryPu(puNumber);
  setClearHistoryType('pu');
  setShowClearHistoryModal(true);
};

// Функция очистки истории по ТП
const handleClearTpHistory = async () => {
  if (selectedIds.length === 0) {
    alert('Выберите строки для очистки истории');
    return;
  }
  setClearHistoryType('tp');
  setShowClearHistoryModal(true);
};
  
// Функция выполнения очистки
const executeClearHistory = async () => {
    try {
      let response;
      
      if (clearHistoryType === 'pu') {
        response = await api.delete(`/api/history/clear-pu/${clearHistoryPu}`, {
          data: { password: clearHistoryPassword }
        });
      } else if (clearHistoryType === 'tp') {
        console.log('=== DEBUG CLEAR TP ===');
        console.log('selectedIds:', selectedIds);
        console.log('networkData:', networkData);
        console.log('filteredData:', filteredData);
        
        // Используйте правильный массив данных
        const dataToUse = searchTp ? filteredData : networkData;
        console.log('Using data:', dataToUse);
        
        // Собираем выбранные строки
        const selectedRows = dataToUse.filter(item => selectedIds.includes(item.id));
        console.log('Selected rows:', selectedRows);
        
        // Собираем уникальные ТП
        const selectedTps = [...new Set(selectedRows.map(item => item.tpName))];
        console.log('Selected TPs:', selectedTps);
        
        // ИСПРАВЛЕНИЕ: Объявляем resIdToUse И проверяем значения
        const resIdToUse = selectedRes || user?.resId || 1; // Добавляем fallback на 1
        console.log('selectedRes:', selectedRes);
        console.log('user:', user);
        console.log('user.resId:', user?.resId);
        console.log('RES ID to use:', resIdToUse);
        
        if (!resIdToUse) {
          alert('Ошибка: не определен РЭС для очистки');
          return;
        }
        
        response = await api.post('/api/history/clear-tp', {
          password: clearHistoryPassword,
          tpNames: selectedTps,
          resId: resIdToUse
        });
      }
      
      alert(response.data.message);
      setShowClearHistoryModal(false);
      setClearHistoryPassword('');
      setSelectedIds([]);
      
      // Обновляем структуру
      await loadNetworkStructure();
      
    } catch (error) {
      alert('Ошибка: ' + (error.response?.data?.error || error.message));
    }
  };

  
  const renderPuCell = (item, position) => {
    const puNumber = position === 'start' ? item.startPu :
                     position === 'middle' ? item.middlePu :
                     item.endPu;

    // Всегда рендерим ровно один квадрат + строку номера фиксированной высоты,
    // чтобы во всех строках было ровно 3 квадрата в идеальных вертикальных колонках.
    return (
      <div
        className="pu-cell"
        onDoubleClick={() => { if (canEditPu) startEdit(item, position); }}
        title={canEditPu ? 'Двойной клик для редактирования' : ''}
      >
        {puNumber ? (
          <div
            className={`status-box ${getStatusColor(
              item.PuStatuses?.find(s => s.puNumber === puNumber && s.position === position)?.status || 'not_checked'
            )}`}
            onClick={() => handleCellClick(item, position)}
          />
        ) : (
          <div className="status-box status-empty">X</div>
        )}
        <span className="pu-number pu-num-line" title={puNumber || ''}>{puNumber || ''}</span>
      </div>
    );
  };
  
  if (loading) return <LoadingSpinner message="Загрузка структуры сети..." submessage="Это может занять несколько секунд" />;
  
  const filteredData = networkData.filter(item => {
  // Фильтр по ТП
  if (searchTp && !item.tpName.toLowerCase().includes(searchTp.toLowerCase())) {
    return false;
  }
  
  // Фильтр по статусу
  if (statusFilter) {
    // Проверяем есть ли хотя бы один ПУ с нужным статусом
    const hasStatus = item.PuStatuses?.some(status => {
      if (statusFilter === 'empty') {
        // Проверяем пустые ячейки
        const hasStart = item.startPu;
        const hasMiddle = item.middlePu;
        const hasEnd = item.endPu;
        return !hasStart || !hasMiddle || !hasEnd;
      }
      return status.status === statusFilter;
    });
    
    // Также проверяем пустые ячейки если нет статусов
    if (!hasStatus && statusFilter === 'empty') {
      return !item.startPu || !item.middlePu || !item.endPu;
    }
    
    return hasStatus;
  }
  
  return true;
});
  const uniqueTps = [...new Set(filteredData.map(item => item.tpName))];
  
  // Функция экспорта в Excel
  const exportStructureToExcel = () => {
    if (filteredData.length === 0) {
      alert('Нет данных для экспорта');
      return;
    }

    // Хелперы секций/техучёта (те же формулы, что и в UI).
    const sectionsById = new Map(sections.map(s => [s.id, s]));
    const resNameById = new Map(filteredData.map(i => [i.resId, i.ResUnit?.name]).filter(([, n]) => n));
    const round1 = (v) => (v == null ? null : Math.round(Number(v) * 10) / 10);
    const sectionLimit = (sec) => (sec && sec.tnKva != null)
      ? sec.tnKva * (sec.cosPhi != null ? sec.cosPhi : 0.9) : null;
    const sectionPct = (sec) => {
      const lim = sectionLimit(sec);
      return (lim && sec.lastPeakKw != null) ? (sec.lastPeakKw / lim) * 100 : null;
    };
    const statusRu = (sec) => sec?.overloadStatus === 'overload' ? 'Перегруз'
      : sec?.overloadStatus === 'ok' ? 'Норма' : 'Нет данных';

    // Подготавливаем данные (лист «Структура»)
    const exportData = filteredData.map(item => {
      // Находим статусы для каждого ПУ
      const getStatus = (puNumber, position) => {
        if (!puNumber) return 'Пусто';
        const status = item.PuStatuses?.find(s => s.puNumber === puNumber && s.position === position);

        switch(status?.status) {
          case 'checked_ok': return 'Проверен';
          case 'checked_error': return 'Ошибка';
          case 'pending_recheck': return 'Ожидает перепроверки';
          case 'not_checked': return 'Не проверен';
          default: return 'Не проверен';
        }
      };

      const sec = item.sectionId != null ? sectionsById.get(item.sectionId) : null;

      return {
        'РЭС': item.ResUnit?.name || '',
        'ТП': item.tpName || '',
        'ВЛ': item.vlName || '',
        'Секция': sec ? `СШ-${toRoman(sec.sectionNumber)}` : '—',
        '№ ПУ тех.учёта': sec?.techPuNumber || '—',
        'Статус секции': sec ? statusRu(sec) : '—',
        'ПУ Начало': item.startPu || '-',
        'Статус начала': getStatus(item.startPu, 'start'),
        'ПУ Середина': item.middlePu || '-',
        'Статус середины': getStatus(item.middlePu, 'middle'),
        'ПУ Конец': item.endPu || '-',
        'Статус конца': getStatus(item.endPu, 'end'),
        'Последнее обновление': new Date(item.lastUpdate).toLocaleDateString('ru-RU')
      };
    });

    // Создаем Excel файл (xlsx-js-style — со стилями ячеек)
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportData);

    // Устанавливаем ширину колонок
    ws['!cols'] = [
      { wch: 20 }, // РЭС
      { wch: 15 }, // ТП
      { wch: 15 }, // ВЛ
      { wch: 10 }, // Секция
      { wch: 18 }, // № ПУ тех.учёта
      { wch: 14 }, // Статус секции
      { wch: 15 }, // ПУ Начало
      { wch: 20 }, // Статус начала
      { wch: 15 }, // ПУ Середина
      { wch: 20 }, // Статус середины
      { wch: 15 }, // ПУ Конец
      { wch: 20 }, // Статус конца
      { wch: 20 }  // Последнее обновление
    ];

    styleExportSheet(ws, {
      'Статус секции': xlsStatusColor,
      'Статус начала': xlsStatusColor,
      'Статус середины': xlsStatusColor,
      'Статус конца': xlsStatusColor
    });
    XLSX.utils.book_append_sheet(wb, ws, 'Структура');

    // Второй лист «Секции (техучёт)» — по секциям: привязка + техучёт + перегруз.
    const vlCountBySection = filteredData.reduce((acc, i) => {
      if (i.sectionId != null) acc[i.sectionId] = (acc[i.sectionId] || 0) + 1;
      return acc;
    }, {});
    const sectionsData = sections
      .slice()
      .sort((a, b) => (a.tpName || '').localeCompare(b.tpName || '', 'ru') || (a.sectionNumber - b.sectionNumber))
      .map(sec => {
        const pct = sectionPct(sec);
        return {
          'РЭС': resNameById.get(sec.resId) || '',
          'ТП': sec.tpName || '',
          'Секция': `СШ-${toRoman(sec.sectionNumber)}`,
          '№ ПУ тех.учёта': sec.techPuNumber || '—',
          'Привязано ВЛ': vlCountBySection[sec.id] || 0,
          'Sном, кВА': sec.tnKva != null ? sec.tnKva : '—',
          'cosφ': sec.cosPhi != null ? sec.cosPhi : 0.9,
          'Лимит, кВт': round1(sectionLimit(sec)) ?? '—',
          'Пик, кВт': sec.lastPeakKw != null ? round1(sec.lastPeakKw) : '—',
          'Загрузка, %': pct != null ? Math.round(pct) : '—',
          'Статус': statusRu(sec),
          'Дата профиля': sec.lastProfileAt ? new Date(sec.lastProfileAt).toLocaleDateString('ru-RU') : '—'
        };
      });
    if (sectionsData.length > 0) {
      const wsSec = XLSX.utils.json_to_sheet(sectionsData);
      wsSec['!cols'] = [
        { wch: 20 }, { wch: 15 }, { wch: 10 }, { wch: 18 }, { wch: 12 },
        { wch: 10 }, { wch: 8 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 14 }
      ];
      styleExportSheet(wsSec, { 'Статус': xlsStatusColor, 'Загрузка, %': xlsLoadColor });
      XLSX.utils.book_append_sheet(wb, wsSec, 'Секции (техучёт)');
    }

    const fileName = `Структура_сети_${selectedRes ? `РЭС_${selectedRes}_` : ''}${new Date().toLocaleDateString('ru-RU').split('.').join('-')}.xlsx`;
    XLSX.writeFile(wb, fileName);

    alert(` экспортирована в файл: ${fileName}`);
  };
  
  return (
    <div className="network-structure">
      <PageHeader icon={<IconLayers size={24} />} title="Структура сети" />
      {canEditPu && (
  <p className="edit-hint">
    <span className="svg-frame svg-frame--sm" style={{marginRight: 8}}><IconEdit size={20} /></span>
      Двойной клик по номеру счетчика для редактирования
  </p>
)}
      

<div className="structure-controls">
  <div className="search-box">
    <input 
      type="text"
      placeholder="Поиск по ТП..."
      value={searchTp}
      onChange={(e) => setSearchTp(e.target.value)}
      className="search-input"
    />
  </div>
  
  <div className="action-buttons-group">
    {selectedIds.length > 0 && (
  <>
    {canClearChecks && (
      <button
        className="clear-history-btn"
        onClick={handleClearTpHistory}
      >
        Очистить историю ({selectedIds.length})
      </button>
    )}

    {canEditStructure && (
      <button
        className="delete-selected-btn"
        onClick={() => setShowDeleteModal(true)}
      >
        Удалить выбранные ({selectedIds.length})
      </button>
    )}
  </>
)}
    
    <button
      className="pm-btn pm-btn--refresh"
       onClick={() => {
        setLoading(true);  // Показать загрузку
        loadNetworkStructure();
      }}
      disabled={loading}
    >
      <IconRefresh className="ico" /> {loading ? 'Обновление...' : 'Обновить структуру'}
    </button>
    
    <button 
      className="export-btn" 
      onClick={exportStructureToExcel}
    >
      <IconChart className="ico" /> Экспорт в Excel
    </button>
  </div>
</div>

     
      
      <div className={`status-legend ${legendCollapsed ? 'is-collapsed' : ''}`}>
  <div
    className={`legend-item ${statusFilter === 'checked_ok' ? 'active' : ''}`}
    onClick={() => setStatusFilter(statusFilter === 'checked_ok' ? null : 'checked_ok')}
  >
    <span className="status-box status-ok"></span> Проверен без отклонений
  </div>
  <div
    className={`legend-item ${statusFilter === 'checked_error' ? 'active' : ''}`}
    onClick={() => setStatusFilter(statusFilter === 'checked_error' ? null : 'checked_error')}
  >
    <span className="status-box status-error"></span> Проверен с отклонениями
  </div>
  <div
    className={`legend-item ${statusFilter === 'pending_recheck' ? 'active' : ''}`}
    onClick={() => setStatusFilter(statusFilter === 'pending_recheck' ? null : 'pending_recheck')}
  >
    <span className="status-box status-pending"></span> Ожидает проверки
  </div>
  <div className="legend-item disabled">
    <span className="status-box status-unchecked"></span> Не проверен
  </div>
  <div className="legend-item disabled">
    <span className="status-box status-empty">X</span> ПУ не задан
  </div>
</div>

{/* Компактные кружки-фильтры при скролле (тот же statusFilter — один источник правды).
    Три кликабельных (как в полосе) + два индикатора (как disabled-строки полосы). */}
<div className={`legend-dots ${legendCollapsed ? 'show' : ''}`} aria-hidden={!legendCollapsed}>
  <button type="button" title="Проверен без отклонений"
    className={`legend-dot status-ok ${statusFilter === 'checked_ok' ? 'active' : ''}`}
    onClick={() => setStatusFilter(statusFilter === 'checked_ok' ? null : 'checked_ok')} />
  <button type="button" title="Проверен с отклонениями"
    className={`legend-dot status-error ${statusFilter === 'checked_error' ? 'active' : ''}`}
    onClick={() => setStatusFilter(statusFilter === 'checked_error' ? null : 'checked_error')} />
  <button type="button" title="Ожидает проверки"
    className={`legend-dot status-pending ${statusFilter === 'pending_recheck' ? 'active' : ''}`}
    onClick={() => setStatusFilter(statusFilter === 'pending_recheck' ? null : 'pending_recheck')} />
  <span className="legend-dot status-unchecked is-indicator" title="Не проверен" />
  <span className="legend-dot status-empty is-indicator" title="ПУ не задан">X</span>
</div>
      
      <div className="structure-grouped">
        {(() => {
          // Класс индикатора техучёта секции по overloadStatus (данные придут в этапе 2).
          const overloadClass = (st) =>
            st === 'ok' ? 'status-ok' : st === 'overload' ? 'status-error' : 'status-unchecked';

          // Строка одной ВЛ в единой сетке net-grid: [чекбокс][наименование]
          // [начало][середина][конец][селект секции][действия]. Колонки фиксированы
          // → квадраты всех строк совпадают по вертикали.
          const renderVlRow = (item, showSectionSelect, askueCells = null) => {
            const tpSections = sections.filter(s => s.tpName === item.tpName);
            return (
              <div key={item.id} className={`net-grid vl-row ${selectedIds.includes(item.id) ? 'selected' : ''}`}>
                <div className="col-check">
                  {(canEditStructure || canClearChecks) && (
                    <input
                      type="checkbox"
                      className="vl-check"
                      checked={selectedIds.includes(item.id)}
                      onChange={() => handleSelectRow(item.id)}
                    />
                  )}
                </div>
                <span className="vl-name" title={item.vlName}>{item.vlName}</span>
                {askueCells || <><span className="askue-cell" /><span className="askue-cell" /><span className="askue-cell" /><span className="askue-cell" /></>}
                {renderPuCell(item, 'start')}
                {renderPuCell(item, 'middle')}
                {renderPuCell(item, 'end')}
                <div className="col-section">
                  {canEditStructure && (
                    <select
                      className="vl-section-select"
                      value={item.sectionId ?? ''}
                      onClick={(e) => e.stopPropagation()}
                      onChange={(e) => { e.stopPropagation(); assignSection(item, e.target.value); }}
                      disabled={tpSections.length === 0}
                      title={tpSections.length === 0 ? 'Сначала добавьте секцию этой ТП' : 'Секция ВЛ'}
                    >
                      {/* value="" — плейсхолдер «Секция…» у непривязанной ВЛ,
                          либо пункт «— Без секции» (отвязать) у привязанной. */}
                      <option value="">{item.sectionId ? '— Без секции' : 'Секция…'}</option>
                      {tpSections.map(s => (
                        <option key={s.id} value={s.id}>СШ-{toRoman(s.sectionNumber)}</option>
                      ))}
                    </select>
                  )}
                </div>
                <div className="col-actions"></div>
              </div>
            );
          };

          // Строка-заголовок колонок ПУ (подписи Начало/Середина/Конец) + значки АСКУЭ.
          const colHeader = (nameCol, actionsCol, askueCells = null) => (
            <div className="net-grid net-grid-head">
              <div className="col-check"></div>
              <div className="vl-name section-title-cell">{nameCol}</div>
              {askueCells || <><span className="askue-cell" /><span className="askue-cell" /><span className="askue-cell" /><span className="askue-cell" /></>}
              <div className="pu-col-label">Начало</div>
              <div className="pu-col-label">Середина</div>
              <div className="pu-col-label">Конец</div>
              <div className="col-section"></div>
              <div className="col-actions">{actionsCol}</div>
            </div>
          );

          // Значки/числа АСКУЭ по секции. Ширины колонок резервируются всегда;
          // без данных — «—». Дата — в tooltip.
          const askueDate = (ad) => ad && ad.at ? new Date(ad.at).toLocaleString('ru-RU') : null;
          const askueTip = (name, ad) => `${name} · ${askueDate(ad) ? `данные АСКУЭ от ${askueDate(ad)}` : 'данных АСКУЭ нет'}`;
          const askueHeadCells = (section) => {
            const ad = section.askueData || null;
            return (
              <>
                <span className="askue-cell askue-head" title={askueTip('Потребители', ad)}><span className="askue-ico ico-glow-blue"><IconUsers size={24} /></span></span>
                <span className="askue-cell askue-head" title={askueTip('Опрашиваемые (кВт)', ad)}><span className="askue-ico ico-glow-green"><IconAskueLimited size={24} /></span></span>
                <span className="askue-cell askue-head" title={askueTip('Не опрашиваемые (кВт)', ad)}><span className="askue-ico ico-glow-red"><IconAskueUnpolled size={24} /></span></span>
                <span className="askue-cell askue-head" title={askueTip('Общая разрешённая мощность ТП (кВт)', ad)}><span className="askue-ico ico-glow-blue"><IconAskueSum size={24} /></span></span>
              </>
            );
          };
          // 4 ячейки-числа АСКУЭ (вставляются в первую строку ВЛ секции — на линию
          // цветных квадратов ПУ). Без данных — приглушённый «—».
          const askueNumberCells = (section) => {
            const ad = section.askueData || null;
            const num = (v, unit) => (v == null
              ? <span className="askue-empty">—</span>
              : <span className="askue-val"><b className="askue-num">{v}</b><i className="askue-unit">{unit}</i></span>);
            return (
              <>
                <span className="askue-cell" title={askueTip('Потребители', ad)}>{num(ad?.consumers, 'шт')}</span>
                <span className="askue-cell" title={askueTip('Опрашиваемые (кВт)', ad)}>{num(ad?.limitedKw, 'кВт')}</span>
                <span className="askue-cell" title={askueTip('Не опрашиваемые (кВт)', ad)}>{num(ad?.unpolledKw, 'кВт')}</span>
                <span className="askue-cell" title={askueTip('Общая разрешённая мощность ТП (кВт)', ad)}>{num(ad?.totalKw, 'кВт')}</span>
              </>
            );
          };
          // Отдельная строка чисел — только для секции без привязанных ВЛ (некуда вставить).
          const askueNumRow = (section) => (
            <div className="net-grid section-askue-row">
              <div className="col-check"></div>
              <div className="vl-name"></div>
              {askueNumberCells(section)}
              <div className="pu-col-label"></div><div className="pu-col-label"></div><div className="pu-col-label"></div>
              <div className="col-section"></div>
              <div className="col-actions"></div>
            </div>
          );

          return uniqueTps.map(tp => {
            const tpItems = filteredData.filter(i => i.tpName === tp);
            const tpSections = sections.filter(s => s.tpName === tp);
            const resId = tpItems[0]?.resId || tpItems[0]?.ResUnit?.id;
            const unassigned = tpItems.filter(i => !i.sectionId);

            return (
              <div key={tp} className="tp-card">
                <div className="tp-card-head">
                  <span className="tp-card-title">{tpItems[0]?.ResUnit?.name} · ТП {tp}</span>
                  {canEditStructure && (
                    <button className="section-add-btn" onClick={() => openSectionModal(tp, resId)}>
                      <IconLayers className="ico" /> Добавить секцию
                    </button>
                  )}
                </div>

                {/* Секции ТП */}
                {tpSections.map(section => {
                  const lines = tpItems.filter(i => i.sectionId === section.id);
                  const parts = [`СШ-${toRoman(section.sectionNumber)}`];
                  if (section.tnKva != null) parts.push(`${section.tnKva} кВА`);
                  if (section.techPuNumber) parts.push(`тех.учёт № ${section.techPuNumber}`);
                  // Живая подпись перегруза (этап 2): пик · время · лимит (1 знак).
                  const cosPhi = section.cosPhi != null ? section.cosPhi : 0.9;
                  const limitKw = section.tnKva != null ? section.tnKva * cosPhi : null;
                  const fmt1 = (v) => (v == null ? '—' : Number(v).toFixed(1));
                  const hasPeak = section.lastPeakKw != null;
                  const pkPct = (hasPeak && limitKw) ? (section.lastPeakKw / limitKw) * 100 : null;
                  const pkCls = pkPct == null ? '' : pkPct > 100 ? 'red' : pkPct >= 85 ? 'amber' : 'green';
                  let peakDate = '';
                  if (section.lastPeakAt) {
                    const d = new Date(section.lastPeakAt);
                    if (!isNaN(d.getTime())) peakDate = d.toLocaleString('ru-RU', { day: '2-digit', month: '2-digit', year: 'numeric', hour: '2-digit', minute: '2-digit' });
                  }
                  // Квадратик техучёта: по уровню загрузки (оранжевая зона 85–100% → оранжевый),
                  // иначе по overloadStatus.
                  const boxCls = pkCls === 'red' ? 'status-error' : pkCls === 'amber' ? 'status-pending'
                    : pkCls === 'green' ? 'status-ok' : overloadClass(section.overloadStatus);
                  const sectionNameCol = (
                    <span className="section-title-inline">
                      <span className={`status-box status-box--sm ${boxCls}`}
                            style={{ cursor: 'pointer' }}
                            title="Сведения о техучёте"
                            onClick={(e) => { e.stopPropagation(); openTechModal(section); }}></span>
                      <span className="section-title">{parts.join(' · ')}</span>
                      {hasPeak && (
                        <span className="section-peak muted">
                          пик <span className={`peak-num ${pkCls}`}>{fmt1(section.lastPeakKw)}</span> кВт
                          {peakDate ? ` · ${peakDate}` : ''}
                          {limitKw != null ? ` · лимит ${fmt1(limitKw)} кВт` : ''}
                        </span>
                      )}
                    </span>
                  );
                  const sectionActions = canEditStructure ? (
                    <>
                      <button className="link-btn" title="Редактировать секцию" onClick={() => openSectionModal(tp, resId, section)}>
                        <IconEdit className="ico ico-glow-blue" />
                      </button>
                      <button className="link-btn" title="Удалить секцию" onClick={() => deleteSection(section)}>
                        <IconTrash className="ico ico-glow-red" />
                      </button>
                    </>
                  ) : null;
                  return (
                    <div key={section.id} className="section-block">
                      {colHeader(sectionNameCol, sectionActions, askueHeadCells(section))}
                      <div className="section-lines">
                        {lines.length === 0
                          ? <>{askueNumRow(section)}<div className="section-empty">Нет привязанных ВЛ</div></>
                          : lines.map((item, i) => renderVlRow(item, false, i === 0 ? askueNumberCells(section) : null))}
                      </div>
                    </div>
                  );
                })}

                {/* ВЛ без секции */}
                {unassigned.length > 0 && (
                  <div className="section-block section-block--none">
                    {colHeader(<span className="section-title muted">ВЛ без секции ({unassigned.length})</span>, null)}
                    <div className="section-lines">
                      {unassigned.map(item => renderVlRow(item, true))}
                    </div>
                  </div>
                )}
              </div>
            );
          });
        })()}
      </div>

      {/* Сведения о техучёте секции */}
      {techModal && (() => {
        const s = techModal;
        const cosPhi = s.cosPhi != null ? s.cosPhi : 0.9;
        const limitKw = s.tnKva != null ? s.tnKva * cosPhi : null;
        const f1 = (v) => (v == null ? '—' : Number(v).toFixed(1));
        const hasData = s.lastPeakKw != null;
        const status = s.overloadStatus === 'overload'
          ? { t: 'Перегруз', cls: 'status-error' }
          : s.overloadStatus === 'ok'
          ? { t: 'Норма', cls: 'status-ok' }
          : { t: 'Нет данных', cls: 'status-unchecked' };
        const STAGE_RU = { askue_limit: 'Ограничение по АСКУЭ', res_work: 'Мероприятия РЭС', awaiting_recheck: 'Ожидает перепроверки' };
        // Шкала загрузки: 0..120% лимита → 0..100% ширины полосы.
        const pct = (limitKw && hasData) ? (s.lastPeakKw / limitKw) * 100 : null;
        const barW = pct != null ? Math.min(pct, 120) * 100 / 120 : 0;
        const barCls = pct == null ? 'gray' : pct > 100 ? 'red' : pct >= 85 ? 'amber' : 'green';
        const srcRu = s.lastProfileSource === '60' ? '60 мин' : s.lastProfileSource === '30' ? '30 мин' : '—';
        return (
          <ModalShell
            title={<>{s.tpName} · {s.ResUnit?.name || '—'}{pct != null && <> · <span className={`tech-title-pct ${barCls}`}>{Math.round(pct)}%</span></>}</>}
            dockTitle={`${s.tpName} · ${s.ResUnit?.name || ''}`}
            className="tech-details-modal"
            icon={<IconLayers size={16} />}
            onClose={() => setTechModal(null)}
          >
              <div className="modal-body">
                <div className="modal-info">
                  <p><strong>ТП:</strong> {s.tpName}</p>
                  <p><strong>РЭС:</strong> {s.ResUnit?.name || '—'}{pct != null && <> · загрузка <span className={`tech-pmax-pct ${barCls}`}>{Math.round(pct)}%</span></>}</p>
                  <p><strong>Секция шин:</strong> СШ-{toRoman(s.sectionNumber)}</p>
                  <p><strong>№ ПУ техучёта:</strong> {s.techPuNumber || '—'}</p>
                </div>

                {hasData ? (
                  <>
                    <div className="tech-pmax">
                      <div className="tech-pmax-value">
                        Pmax {f1(s.lastPeakKw)} кВт
                        {pct != null && <span className={`tech-pmax-pct ${barCls}`}>{Math.round(pct)}%</span>}
                      </div>
                      <div className="tech-pmax-sub">
                        {s.lastPeakAt ? new Date(s.lastPeakAt).toLocaleString('ru-RU') : '—'}
                        {s.lastProfilePeriod ? ` · период ${s.lastProfilePeriod}` : ''}
                      </div>
                    </div>

                    <div className="tech-bar">
                      <div className={`tech-bar-fill ${barCls}`} style={{ width: `${barW}%` }} />
                    </div>
                    <div className="tech-bar-label">
                      {pct != null ? `${Math.round(pct)}% от лимита ${f1(limitKw)} кВт` : `лимит ${f1(limitKw)} кВт`}
                    </div>

                    <div className="tech-grid">
                      <div><span className="k">Sном тр-ра, кВА</span><span className="v">{s.tnKva != null ? s.tnKva : '—'}</span></div>
                      <div><span className="k">cosφ</span><span className="v">{cosPhi}</span></div>
                      <div><span className="k">Лимит, кВт</span><span className="v">{f1(limitKw)}</span></div>
                      <div><span className="k">Источник ряда</span><span className="v">{srcRu}</span></div>
                      <div><span className="k">Дата загрузки профиля</span><span className="v">{s.lastProfileAt ? new Date(s.lastProfileAt).toLocaleString('ru-RU') : '—'}</span></div>
                    </div>

                    {techCase && (
                      <div className="highlight-box">
                        <strong>Открыт случай перегруза:</strong> {STAGE_RU[techCase.stage] || techCase.stage}
                        {typeof onSectionChange === 'function' && (
                          <div style={{ marginTop: 8 }}>
                            <button className="action-btn" onClick={() => { setTechModal(null); onSectionChange('power_overload'); }}>
                              Перейти в «Превышение Pном»
                            </button>
                          </div>
                        )}
                      </div>
                    )}
                  </>
                ) : (
                  <div className="tech-empty">
                    <span className="svg-frame"><IconZap size={24} /></span>
                    <p>Профиль мощности не загружался</p>
                    <p className="muted" style={{ fontSize: 12.5 }}>Загрузите «Профиль мощности (Пирамида)» в разделе «Загрузить файлы», чтобы увидеть Pmax и статус загрузки секции.</p>
                  </div>
                )}
              </div>
              <div className="modal-footer">
                <button className="action-btn" onClick={() => setTechModal(null)}>Закрыть</button>
              </div>
          </ModalShell>
        );
      })()}

      {/* Форма секции шин ТП */}
      {sectionModal && (
        <ModalShell
          title={`${sectionModal.id ? 'Редактировать секцию' : 'Новая секция'} — ТП ${sectionModal.tpName}`}
          icon={<IconLayers size={16} />}
          onClose={() => setSectionModal(null)}
        >
            <div className="modal-body">
              <div className="form-group">
                <label>Номер секции шин</label>
                <select value={sectionForm.sectionNumber}
                  onChange={(e) => setSectionForm({ ...sectionForm, sectionNumber: e.target.value })}>
                  <option value="">— выберите —</option>
                  {[1, 2, 3, 4, 5].map(n => (
                    <option key={n} value={n}>СШ-{toRoman(n)}</option>
                  ))}
                </select>
              </div>
              <div className="form-group">
                <label>Sном тр-ра, кВА</label>
                <input type="number" step="any" value={sectionForm.tnKva}
                  onChange={(e) => setSectionForm({ ...sectionForm, tnKva: e.target.value })} />
              </div>
              <div className="form-group">
                <label>cosφ</label>
                <input type="number" step="any" value={sectionForm.cosPhi}
                  onChange={(e) => setSectionForm({ ...sectionForm, cosPhi: e.target.value })} />
              </div>
              <div className="form-group">
                <label>№ ПУ технического учёта</label>
                <input type="text" value={sectionForm.techPuNumber}
                  onChange={(e) => setSectionForm({ ...sectionForm, techPuNumber: e.target.value })} />
              </div>
            </div>
            <div className="modal-footer">
              <button className="cancel-btn" onClick={() => setSectionModal(null)}>Отмена</button>
              <button className="confirm-btn" onClick={saveSection}
                disabled={sectionForm.sectionNumber === ''}>Сохранить</button>
            </div>
        </ModalShell>
      )}

      {puEdit && (
        <ModalShell
          title="Редактирование ПУ"
          dockTitle={`ПУ · ${puEdit.item.tpName}`}
          icon={<IconEdit size={16} />}
          className="pu-edit-modal"
          onClose={cancelEdit}
        >
            <div className="modal-body">
              <div className="modal-info">
                <p><strong>ТП:</strong> {puEdit.item.tpName} · <strong>ВЛ:</strong> {puEdit.item.vlName}</p>
                <p><strong>Позиция:</strong> {puEdit.position === 'start' ? 'Начало' : puEdit.position === 'middle' ? 'Середина' : 'Конец'}</p>
              </div>
              <div className="pu-edit-grid">
                <div className="form-group">
                  <label>Старый номер ПУ</label>
                  <input type="text" value={puEdit.oldValue || '—'} readOnly className="pu-edit-old" />
                </div>
                <div className="form-group">
                  <label>Новый номер ПУ</label>
                  <input type="text" value={editValue} autoFocus
                    onChange={(e) => setEditValue(e.target.value)}
                    onKeyDown={(e) => { if (e.key === 'Enter') saveEdit(); }}
                    placeholder="Пусто — убрать ПУ из позиции" />
                </div>
              </div>
            </div>
            <div className="modal-footer">
              <button className="cancel-btn" onClick={cancelEdit}>Отмена</button>
              <button className="confirm-btn" onClick={saveEdit}
                disabled={editValue === (puEdit.oldValue || '')}>Сохранить</button>
            </div>
        </ModalShell>
      )}

      <ErrorDetailsModal
        isOpen={modalOpen}
        onClose={() => setModalOpen(false)}
        details={selectedDetails}
        tpName={selectedItem?.tpName}
        vlName={selectedItem?.vlName}
        position={selectedPosition}
      />

{showExtendedModal && selectedPuData && (
        <ExtendedPuModal
          isOpen={showExtendedModal}
          onClose={() => {
            setShowExtendedModal(false);
            setSelectedPuData(null);
            setUploadHistory([]);
            setCheckHistory([]);
          }}
          puData={selectedPuData}
          activeTab={activeTab}
          setActiveTab={setActiveTab}
          uploadHistory={uploadHistory}
          checkHistory={checkHistory}
          loading={historyLoading}
          handleClearPuHistory={handleClearPuHistory}
        />
      )}
      
      {/* Модальное окно для удаления */}
      {showDeleteModal && (
        <ModalShell
          variant="confirm"
          title="Подтверждение удаления"
          className="delete-modal"
          onClose={() => {setShowDeleteModal(false); setDeletePassword('');}}
        >
            <div className="modal-body">
              <p>Вы собираетесь удалить {selectedIds.length} записей.</p>
              <p className="warning"><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /> Это действие нельзя отменить!</p>
              <div className="form-group">
                <label>Введите пароль администратора:</label>
                <input
                  type="password"
                  value={deletePassword}
                  onChange={(e) => setDeletePassword(e.target.value)}
                  placeholder="Пароль"
                  autoFocus
                  autoComplete="new-password"    
                  name="delete-notification-password"  
                />
              </div>
            </div>
            <div className="modal-footer">
              <button className="cancel-btn" onClick={() => {setShowDeleteModal(false); setDeletePassword('');}}>
                Отмена
              </button>
              <button 
                className="danger-btn" 
                onClick={handleDeleteSelected}
                disabled={!deletePassword}
              >
                Удалить
              </button>
            </div>
        </ModalShell>
      )}

{/* Модальное окно очистки истории */}
{showClearHistoryModal && (
  <ModalShell variant="confirm" title="Подтверждение очистки истории" className="delete-modal"
    onClose={() => setShowClearHistoryModal(false)}>
      <div className="modal-body">
        <p>
          {clearHistoryType === 'pu' && `Вы собираетесь очистить всю историю для ПУ ${clearHistoryPu}`}
          {clearHistoryType === 'tp' && `Вы собираетесь очистить историю для выбранных строк (${selectedIds.length} записей)`}
          {clearHistoryType === 'all' && 'Вы собираетесь очистить ВСЮ историю системы'}
        </p>
        <p className="warning"><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /> Будут удалены все записи о загрузках и проверках!</p>
        <div className="form-group">
          <label>Введите пароль администратора:</label>
          <input
            type="password"
            value={clearHistoryPassword}
            onChange={(e) => setClearHistoryPassword(e.target.value)}
            placeholder="Пароль"
            autoFocus
          />
        </div>
      </div>
      <div className="modal-footer">
        <button className="cancel-btn" onClick={() => setShowClearHistoryModal(false)}>
          Отмена
        </button>
        <button 
          className="danger-btn" 
          onClick={executeClearHistory}
          disabled={!clearHistoryPassword}
        >
          Очистить историю
        </button>
      </div>
  </ModalShell>
)}

{/* Плавающие кнопки действий структуры (этап 3, блок В) */}
      {/* Все плавающие кнопки появляются вместе с кнопкой «наверх» (при прокрутке). */}
      {showScrollTop && (
        <div className="structure-fab-stack">
          {selectedIds.length > 0 && (
            <>
              {canEditStructure && (
                <button className="fab-btn fab-danger" title="Удалить выбранные" onClick={() => setShowDeleteModal(true)}>
                  <IconTrash className="ico" />
                  <span className="fab-badge">{selectedIds.length}</span>
                </button>
              )}
              {canClearChecks && (
                <button className="fab-btn fab-warn" title="Очистить историю выбранных" onClick={handleClearTpHistory}>
                  <IconBroom className="ico" />
                  <span className="fab-badge">{selectedIds.length}</span>
                </button>
              )}
            </>
          )}
          <button className="fab-btn fab-blue" title="Обновить структуру" disabled={loading}
            onClick={() => { setLoading(true); loadNetworkStructure(); }}>
            <IconRefresh className="ico" />
          </button>
          <button className="fab-btn fab-green" title="Выгрузка в Excel" onClick={exportStructureToExcel}>
            <IconDownload className="ico" />
          </button>
        </div>
      )}

      {showScrollTop && (
        <button
    className="scroll-to-top"
    onClick={() => {
      const contentElement = document.querySelector('.content');
      if (contentElement) {
        contentElement.scrollTo({ top: 0, behavior: 'smooth' });
      }
    }}
    title="Наверх"
  >
    <IconArrowUp className="ico" />
  </button>
)}
      
    </div>
  );
}

// Модальное окно с деталями ошибки
function ErrorDetailsModal({ isOpen, onClose, details, tpName, vlName, position }) {
  if (!isOpen || !details) return null;
  
  // Парсим детали если они в формате JSON строки
  let errorSummary = '';
  let parsedDetails = null;
  
  try {
    if (details?.errorDetails) {
      const parsed = JSON.parse(details.errorDetails);
      errorSummary = parsed.summary || details.errorDetails;
      parsedDetails = parsed.details;
    }
  } catch (e) {
    errorSummary = details?.errorDetails || 'Нет данных';
  }
  
 
// Парсим фазы из деталей - красим ТОЛЬКО явно указанные!
const getPhaseErrors = () => {
  const phases = { A: false, B: false, C: false };
  
  if (parsedDetails) {
    // Проверяем только конкретные фазы
    if (parsedDetails.overvoltage) {
      if (parsedDetails.overvoltage.phase_A && parsedDetails.overvoltage.phase_A.count > 0) phases.A = true;
      if (parsedDetails.overvoltage.phase_B && parsedDetails.overvoltage.phase_B.count > 0) phases.B = true;
      if (parsedDetails.overvoltage.phase_C && parsedDetails.overvoltage.phase_C.count > 0) phases.C = true;
    }
    
    if (parsedDetails.undervoltage) {
      if (parsedDetails.undervoltage.phase_A && parsedDetails.undervoltage.phase_A.count > 0) phases.A = true;
      if (parsedDetails.undervoltage.phase_B && parsedDetails.undervoltage.phase_B.count > 0) phases.B = true;
      if (parsedDetails.undervoltage.phase_C && parsedDetails.undervoltage.phase_C.count > 0) phases.C = true;
    }
  }
  
  // Проверяем текст только на явные упоминания
  if (errorSummary) {
    if (errorSummary.indexOf('Фаза A') !== -1 || errorSummary.indexOf('phase_A') !== -1) phases.A = true;
    if (errorSummary.indexOf('Фаза B') !== -1 || errorSummary.indexOf('phase_B') !== -1) phases.B = true;
    if (errorSummary.indexOf('Фаза C') !== -1 || errorSummary.indexOf('phase_C') !== -1) phases.C = true;
  }
  
  return phases;
};
  
  const phaseErrors = getPhaseErrors();
  
  return (
    <ModalShell title={`Детали проверки ПУ #${details?.puNumber}`} className="error-details-modal"
      icon={<IconClipboard size={16} />} onClose={onClose}>
        <div className="modal-body">
          <div className="modal-info">
            <p><strong>ТП:</strong> {tpName}</p>
            <p><strong>Фидер:</strong> {vlName}</p>
            <p><strong>Позиция:</strong> {position === 'start' ? 'Начало' : position === 'middle' ? 'Середина' : 'Конец'}</p>
          </div>
          
          <div className="phase-indicators-large">
            <div className={`phase-indicator ${phaseErrors.A ? 'phase-error' : ''}`}>A</div>
            <div className={`phase-indicator ${phaseErrors.B ? 'phase-error' : ''}`}>B</div>
            <div className={`phase-indicator ${phaseErrors.C ? 'phase-error' : ''}`}>C</div>
          </div>
          
          <div className="error-summary">
            <h4>Обнаруженные отклонения:</h4>
            <div className="error-text">{errorSummary}</div>
          </div>
        </div>
        
        <div className="modal-footer">
          <button className="action-btn" onClick={onClose}>Закрыть</button>
        </div>
    </ModalShell>
  );
}


// =====================================================
// КОМПОНЕНТ ЗАГРУЗКИ ФАЙЛОВ
// =====================================================

function FileUpload() {
  const [selectedType, setSelectedType] = useState('');
  const [files, setFiles] = useState([]);
  const [uploadProgress, setUploadProgress] = useState({ current: 0, total: 0 });
  const [uploading, setUploading] = useState(false);
  const [uploadResult, setUploadResult] = useState(null);
  const { user } = useContext(AuthContext);
  const [dragActive, setDragActive] = useState(false);

  const fileTypes = [
    { id: 'rim_single', label: 'РИМ', description: 'журнал напряжения' },
    { id: 'nartis', label: 'Нартис', description: 'журнал напряжения' },
    { id: 'energomera', label: 'Энергомера', description: 'журнал напряжения' },
    { id: 'profile', label: 'Профиль мощности', description: 'Пирамида сети' }
  ];

  const handleFileSelect = (e) => {
    setFiles(Array.from(e.target.files));
    setUploadResult(null);
  };

  const handleDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (e.type === "dragenter" || e.type === "dragover") {
      setDragActive(true);
    } else if (e.type === "dragleave") {
      setDragActive(false);
    }
  };

  const handleDrop = (e) => {
    e.preventDefault();
    e.stopPropagation();
    setDragActive(false);
    
    if (e.dataTransfer.files && e.dataTransfer.files[0]) {
      setFiles(Array.from(e.dataTransfer.files));
      setUploadResult(null);
    }
  };

  const removeFile = (index) => {
    setFiles(files.filter((_, i) => i !== index));
  };

  const handleUpload = async () => {
  if (!files.length || !selectedType) {
    alert('Выберите тип файла и файлы для загрузки');
    return;
  }

  // ── Профиль мощности (Пирамида): resId не нужен, матчинг по ПУ техучёта ──
  if (selectedType === 'profile') {
    setUploading(true);
    setUploadResult(null);
    setUploadProgress({ current: 0, total: files.length });
    let sectionsUpdated = 0, overloadCount = 0;
    const unmatched = [];
    const fileErrors = [];
    const details = [];
    for (let i = 0; i < files.length; i++) {
      setUploadProgress({ current: i + 1, total: files.length });
      const formData = new FormData();
      formData.append('file', files[i]);
      formData.append('type', 'profile');
      try {
        const response = await api.post('/api/upload/analyze', formData, {
          headers: { 'Content-Type': 'multipart/form-data' }
        });
        sectionsUpdated += response.data.sectionsUpdated || 0;
        overloadCount += response.data.overloadCount || 0;
        (response.data.unmatched || []).forEach(pu => unmatched.push(pu));
        (response.data.details || []).forEach(d => details.push(d));
      } catch (error) {
        fileErrors.push({ fileName: files[i].name, error: error.response?.data?.error || 'Ошибка загрузки' });
      }
    }
    setUploadResult({
      profile: true,
      sectionsUpdated,
      overloadCount,
      unmatched,
      details,
      errors: fileErrors
    });
    setFiles([]);
    setSelectedType('');
    setUploading(false);
    window.dispatchEvent(new CustomEvent('structureUpdated'));
    window.dispatchEvent(new CustomEvent('notificationsUpdated'));
    return;
  }

  // Определяем resId
  let resIdToUse;
  if (user.role === 'admin') {
    resIdToUse = user.resId || 1;
  } else {
    resIdToUse = user.resId;
  }
  
  if (!resIdToUse) {
    alert('Ошибка: не определен РЭС для загрузки');
    return;
  }
  
  setUploading(true);
  setUploadResult(null);
  setUploadProgress({ current: 0, total: files.length });
  
  const results = [];
  const errors = [];
  let duplicatesCount = 0;
  let successCount = 0;
  let problemsCount = 0;
  let wrongPeriodCount = 0;
  
  // Обрабатываем каждый файл
for (let i = 0; i < files.length; i++) {
  const file = files[i];
  setUploadProgress({ current: i + 1, total: files.length });
  
  const formData = new FormData();
  formData.append('file', file);
  formData.append('type', selectedType);
  formData.append('resId', resIdToUse);
  
  try {
    const response = await api.post('/api/upload/analyze', formData, {
      headers: { 'Content-Type': 'multipart/form-data' }
    });
    
    // Проверка на разные статусы
    const firstDetail = response.data.details?.[0];
    
    if (firstDetail) {
      if (firstDetail.status === 'duplicate_error') {
        duplicatesCount++;
        results.push({
          fileName: file.name,
          status: 'duplicate',
          message: firstDetail.error
        });
      } else if (firstDetail.status === 'wrong_period') {
        wrongPeriodCount++;
        results.push({
          fileName: file.name,
          status: 'wrong_period',
          message: firstDetail.error
        });
      } else if (firstDetail.status === 'not_in_structure') {
        results.push({
          fileName: file.name,
          status: 'not_found',
          message: 'ПУ не найден в структуре сети'
        });
      } else {
        // Обычная обработка
        if (response.data.errors > 0) {
          problemsCount += response.data.errors;
        } else {
          successCount++;
        }
        
        results.push({
          fileName: file.name,
          status: 'processed',
          ...response.data
        });
      }
    }
    
  } catch (error) {
    errors.push({
      fileName: file.name,
      error: error.response?.data?.error || 'Ошибка загрузки'
    });
  }
}
  
  // Показываем итоговый результат
  setUploadResult({
    success: errors.length === 0,
    totalFiles: files.length,
    successCount,
    problemsCount,
    duplicatesCount,
    wrongPeriodCount,
    errorCount: errors.length,
    results,
    errors
  });
  
  // Формируем итоговое сообщение
  let message = `Обработано файлов: ${files.length}\n`;
  if (successCount > 0) message += `Отклонений по напряжению не найдено: ${successCount}\n`;
  if (problemsCount > 0) message += `Отклонения по напряжению найдены: ${problemsCount}\n`;
  if (duplicatesCount > 0) message += `Загружен ранее использованный файл: ${duplicatesCount}\n`;
  if (wrongPeriodCount > 0) message += `Неверный период загруженного файла: ${wrongPeriodCount}\n`;
  if (errors.length > 0) message += `Ошибок загрузки: ${errors.length}`;
  
  alert(message);
  
  // Сбрасываем форму
  setFiles([]);
  setSelectedType('');
  setUploading(false);
  
  // Создаем событие для обновления структуры
  window.dispatchEvent(new CustomEvent('structureUpdated'));
  window.dispatchEvent(new CustomEvent('notificationsUpdated'));
};

  return (
    <div className="file-upload-container">
      <PageHeader icon={<IconUpload size={24} />} title="Загрузка файлов для анализа" />
      <p className="upload-note">Имя файла должно совпадать с номером ПУ</p>
      <p className="upload-subtitle">Загружайте Excel файлы с данными счётчиков для автоматической проверки</p>

      {/* Шаг 1 — тип загрузки */}
      <div className="file-type-selection">
        <h3 className="fu-h3"><span className="fu-step">1</span> Тип загрузки</h3>
        <div className="file-types-grid">
          {fileTypes.map(type => (
            <div
              key={type.id}
              className={`file-type-card ${selectedType === type.id ? 'selected' : ''}`}
              onClick={() => setSelectedType(type.id)}
            >
              <span className="ft-check"><IconCheck className="ico" /></span>
              <span className="ft-icon"><IconMeter className="ico" /></span>
              <span className="ft-label">{type.label}</span>
              <span className="ft-sub">{type.description}</span>
            </div>
          ))}
        </div>
      </div>

      {/* Шаг 2 — файл (появляется после выбора типа) */}
      {selectedType && (
        <div className="file-pick-section">
          <h3 className="fu-h3"><span className="fu-step">2</span> Файл</h3>
          <div className="file-pick-row">
            <input
              type="file"
              id="file-input"
              accept=".xlsx,.xls,.csv"
              multiple
              onChange={handleFileSelect}
              style={{ display: 'none' }}
            />
            <label htmlFor="file-input" className="fu-pick-btn">
              <IconFolder className="ico" /> Выбрать файл
            </label>
            <span className="file-pick-hint">.xlsx, .xls, .csv — можно несколько</span>
          </div>
          {files.length > 0 && (
            <div className="file-chips">
              {files.map((file, idx) => (
                <span key={idx} className="file-chip">
                  <IconFileText className="ico" />
                  <span className="file-chip-name" title={file.name}>{file.name}</span>
                  <button className="file-chip-x" onClick={() => removeFile(idx)}><IconX className="ico" /></button>
                </span>
              ))}
            </div>
          )}
        </div>
      )}

      {/* Прогресс загрузки */}
      {uploading && (
        <div className="upload-progress-section">
          <LoadingSpinner
            type="pulse"
            message={`Обработка файлов: ${uploadProgress.current} из ${uploadProgress.total}`}
            submessage="Пожалуйста, не закрывайте страницу"
          />
        </div>
      )}

      {/* Финальное действие — сплошная синяя */}
      {files.length > 0 && !uploading && (
        <div className="upload-actions">
          <button
            onClick={handleUpload}
            disabled={!selectedType}
            className="fu-submit-btn"
          >
            <IconRocket className="ico" />
            Загрузить и анализировать ({files.length})
          </button>
        </div>
      )}

      {/* Результат загрузки профиля мощности */}
      {uploadResult && uploadResult.profile && (
        <div className="upload-result success">
          <h3>Профиль мощности обработан</h3>
          <p>Секций обновлено: <strong>{uploadResult.sectionsUpdated}</strong></p>
          <p>Перегрузов: <strong style={{ color: uploadResult.overloadCount > 0 ? 'var(--red)' : 'inherit' }}>{uploadResult.overloadCount}</strong></p>
          <p>Не привязано к структуре ПУ: <strong>{uploadResult.unmatched.length}</strong></p>
          {uploadResult.unmatched.length > 0 && (
            <p className="pu-number" style={{ wordBreak: 'break-all' }}>{uploadResult.unmatched.join(', ')}</p>
          )}
          {uploadResult.errors.length > 0 && (
            <div style={{ color: 'var(--red)', marginTop: 8 }}>
              {uploadResult.errors.map((e, i) => <div key={i}>{e.fileName}: {e.error}</div>)}
            </div>
          )}

          {uploadResult.details && uploadResult.details.length > 0 && (
            <details className="profile-details">
              <summary>Детали расчёта ({uploadResult.details.length})</summary>
              <div className="scroll" style={{ maxHeight: 360, overflow: 'auto', marginTop: 8 }}>
                <table className="data">
                  <thead>
                    <tr>
                      <th>ПУ</th><th>Секция</th><th>peakRaw</th><th>Кт</th><th>peakKw</th>
                      <th>Дата пика</th><th>Sном</th><th>cosφ</th><th>Лимит</th><th>Решение</th>
                    </tr>
                  </thead>
                  <tbody>
                    {uploadResult.details.map((d, i) => {
                      const f1 = (v) => (v == null ? '—' : Number(v).toFixed(1));
                      const cls = d.decision === 'overload' ? 'var(--red)' : d.decision === 'ok' ? 'var(--green)' : 'var(--text-muted)';
                      const RU = { overload: 'Перегруз', ok: 'Норма', unknown: 'Нет Sном', not_matched: 'Не сопоставлен' };
                      return (
                        <tr key={i}>
                          <td>{d.puNumber}</td>
                          <td>{d.tpSection || '—'}</td>
                          <td>{f1(d.peakRaw)}</td>
                          <td>{d.kt ?? '—'}</td>
                          <td><strong>{f1(d.peakKw)}</strong></td>
                          <td>{d.peakAt || '—'}</td>
                          <td>{d.tnKva != null ? f1(d.tnKva) : '—'}</td>
                          <td>{d.cosPhi != null ? d.cosPhi : '—'}</td>
                          <td>{f1(d.limitKw)}</td>
                          <td style={{ color: cls, fontWeight: 600 }}>{RU[d.decision] || d.decision}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </details>
          )}
        </div>
      )}
    </div>
  );
}

// =====================================================
// КОМПОНЕНТ УВЕДОМЛЕНИЙ (ИСПРАВЛЕННЫЙ!)
// =====================================================

function Notifications({ filterType, onSectionChange, selectedRes }) {
  const [notifications, setNotifications] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selectedNotification, setSelectedNotification] = useState(null);
  const [showCompleteModal, setShowCompleteModal] = useState(false);
  const [comment, setComment] = useState('');
  const [checkFromDate, setCheckFromDate] = useState(new Date().toISOString().split('T')[0]);
  const [searchTp, setSearchTp] = useState('');
  const [periodFilter, setPeriodFilter] = useState(''); // фильтр по месяцу появления уведомления
  const { user } = useContext(AuthContext);
  const canDeleteNotifs = hasPerm(user, 'notifications_delete');
  const [showDeleteModal, setShowDeleteModal] = useState(false);
  const [deleteNotificationId, setDeleteNotificationId] = useState(null);
  const [deletePassword, setDeletePassword] = useState('');
  const [showDetailsModal, setShowDetailsModal] = useState(false);
  const [detailsNotification, setDetailsNotification] = useState(null);
  const [uploadingPu, setUploadingPu] = useState(null);
  const [attachedFiles, setAttachedFiles] = useState([]); // ДОБАВЛЕНО!
  const [submitting, setSubmitting] = useState(false);
  const [commentError, setCommentError] = useState(false); // подсветка «минимум 5 слов»
  const [selectedNotificationIds, setSelectedNotificationIds] = useState([]);
  const [showBulkDeleteModal, setShowBulkDeleteModal] = useState(false);
  const [bulkDeletePassword, setBulkDeletePassword] = useState('');
  const [showScrollTop, setShowScrollTop] = useState(false);
  const [deleteRelatedDocs, setDeleteRelatedDocs] = useState(false);
  
 // Оптимизированная функция загрузки
const loadNotifications = useCallback(async () => {
  try {
    const params = new URLSearchParams();
    if (selectedRes) params.set('resId', selectedRes);
    if (filterType) params.set('type', filterType);   // шлём тип на сервер
    const qs = params.toString();
    const response = await api.get(`/api/notifications${qs ? `?${qs}` : ''}`);
    const filtered = response.data.filter(n => (filterType ? n.type === filterType : true));
    setNotifications(filtered);
  } catch (error) {
    console.error('Error loading notifications:', error);
  } finally {
    setLoading(false);
  }
}, [filterType, selectedRes]);

  useEffect(() => {
    loadNotifications();
     markAsRead();
    
    // Слушаем события обновления
    const handleUpdate = () => loadNotifications();
    
    window.addEventListener('structureUpdated', handleUpdate);
    window.addEventListener('notificationsUpdated', handleUpdate);
    window.addEventListener('dataCleared', handleUpdate);
    
    // PERF: автообновление каждые 30 секунд, но только на видимой вкладке
    const interval = setInterval(() => {
      if (!document.hidden) loadNotifications();
    }, 30000);
    
    return () => {
      window.removeEventListener('structureUpdated', handleUpdate);
      window.removeEventListener('notificationsUpdated', handleUpdate);
      window.removeEventListener('dataCleared', handleUpdate);
      clearInterval(interval);
    };
  }, [loadNotifications]);

  useEffect(() => {
    // .content может ещё не быть в DOM на момент маунта — ждём через rAF,
    // иначе слушатель не привяжется и кнопка «наверх» не появится.
    let contentElement = null;
    let rafId = null;
    const handleScroll = () => {
      if (contentElement) setShowScrollTop(contentElement.scrollTop > 300);
    };
    const attach = () => {
      contentElement = document.querySelector('.content');
      if (contentElement) {
        contentElement.addEventListener('scroll', handleScroll);
        handleScroll();
      } else {
        rafId = requestAnimationFrame(attach);
      }
    };
    attach();
    return () => {
      if (rafId) cancelAnimationFrame(rafId);
      if (contentElement) contentElement.removeEventListener('scroll', handleScroll);
    };
  }, []);

  const markAsRead = async () => {
  try {
    // Отмечаем уведомления как прочитанные при открытии
    await api.put('/api/notifications/mark-read', { 
      type: filterType === 'error' ? 'error' : 
            filterType === 'pending_askue' ? 'pending_askue' : 
            'all'
    });
    
    // Обновляем счетчики
    window.dispatchEvent(new CustomEvent('notificationsUpdated'));
  } catch (error) {
    console.error('Error marking as read:', error);
  }
};

  const handleCompleteWork = async () => {
    const wordCount = comment.trim().split(' ').filter(word => word.length > 0).length;
    if (wordCount < 5) {
      // Не молча блокируем, а зажигаем красную рамку и фразу — чтобы было понятно,
      // почему кнопка не срабатывает (раньше клик по disabled ничего не давал).
      setCommentError(true);
      return;
    }
    setCommentError(false);

     setSubmitting(true);
    
    try {
      const formData = new FormData();
      formData.append('comment', comment);
      formData.append('checkFromDate', checkFromDate);
      
      // Добавляем файлы
      attachedFiles.forEach(file => {
        formData.append('attachments', file);
      });
      
      await api.post(`/api/notifications/${selectedNotification.id}/complete-work`, formData, {
        headers: { 'Content-Type': 'multipart/form-data' }
      });

      // Закрываем модальное окно сразу
    setShowCompleteModal(false);
    setComment('');
    setAttachedFiles([]);
    setSelectedNotification(null);

      
      alert('Мероприятия отмечены как выполненные');
      setShowCompleteModal(false);
      setComment('');
      setAttachedFiles([]);
      setSelectedNotification(null);
      
      await loadNotifications();
      
    } catch (error) {
      console.error('Complete work error:', error);
      let msg;
      if (error.code === 'ECONNABORTED' || (error.message && error.message.toLowerCase().includes('timeout'))) {
        msg = 'Превышено время ожидания. Возможно, файлы слишком большие — попробуйте уменьшить их размер или загрузить меньше файлов.';
      } else if (!error.response) {
        msg = 'Сервер не отвечает. Проверьте интернет-соединение и попробуйте снова.';
      } else if (error.response.data?.error) {
        msg = error.response.data.error;
      } else {
        msg = `Ошибка сервера (код ${error.response.status})`;
      }
      alert('Ошибка: ' + msg);
    } finally {
    setSubmitting(false); // ДОБАВИТЬ - разблокируем кнопку в любом случае
  }
};

  const handleDeleteNotification = async () => {
    try {
      await api.delete(`/api/notifications/${deleteNotificationId}`, {
        data: { password: deletePassword }
      });
     
      alert('Уведомление удалено');
      setShowDeleteModal(false);
      setDeletePassword('');
      setDeleteNotificationId(null);
      
      // ВАЖНО: Автообновление после удаления!
      await loadNotifications();
      
    } catch (error) {
      alert('Ошибка удаления: ' + (error.response?.data?.error || error.message));
    }
  };

  const handleSelectNotification = (id) => {
    setSelectedNotificationIds(prev => {
      if (prev.includes(id)) {
        return prev.filter(i => i !== id);
      } else {
        return [...prev, id];
      }
    });
  };

  const handleSelectAll = () => {
    if (selectedNotificationIds.length === filteredNotifications.length) {
      setSelectedNotificationIds([]);
    } else {
      setSelectedNotificationIds(filteredNotifications.map(n => n.id));
    }
  };

  const handleBulkDelete = async () => {
  try {
    const response = await api.post('/api/notifications/delete-bulk', {
      ids: selectedNotificationIds,
      password: bulkDeletePassword,
      deleteDocuments: deleteRelatedDocs // Передаём опцию
    });
    
    // Показываем детальный результат
    alert(response.data.message);
    
    setShowBulkDeleteModal(false);
    setBulkDeletePassword('');
    setDeleteRelatedDocs(false); // Сбрасываем
    setSelectedNotificationIds([]);
    setSearchTp('');
    
    await loadNotifications();
    
    // Если удаляли документы - обновляем и их
    if (deleteRelatedDocs) {
      window.dispatchEvent(new CustomEvent('documentsUpdated'));
    }
    
  } catch (error) {
    alert('Ошибка удаления: ' + (error.response?.data?.error || error.message));
  }
};

  // Функция загрузки файла прямо из уведомления АСКУЭ
  const handleFileUpload = async (puNumber, notificationData) => {
  const input = document.createElement('input');
  input.type = 'file';
  input.accept = '.xlsx,.xls,.csv';
  
  input.onchange = async (e) => {
    const file = e.target.files[0];
    if (!file) return;
    
    // Проверяем имя файла
    const fileName = file.name.split('.')[0];
    if (fileName !== puNumber) {
      alert(`Имя файла должно быть ${puNumber}.xls или ${puNumber}.xlsx`);
      return;
    }
    
    setUploadingPu(puNumber);
    
    const formData = new FormData();
    formData.append('file', file);
    formData.append('type', 'rim_single');
    formData.append('resId', user.resId);
    formData.append('requiredPeriod', notificationData.checkFromDate);
    
    try {
      const response = await api.post('/api/upload/analyze', formData, {
        headers: { 'Content-Type': 'multipart/form-data' }
      });
      
      // ПРОВЕРЯЕМ РЕЗУЛЬТАТ!
      if (response.data.details && response.data.details.length > 0) {
        const firstResult = response.data.details[0];
        
        // Проверяем статус
        if (firstResult.status === 'wrong_period') {
          // Показываем ошибку периода
          alert(firstResult.error);
          // НЕ обновляем уведомления, чтобы можно было попробовать снова
          return;
        } else if (firstResult.status === 'duplicate_error') {
          // Показываем ошибку дубликата
          alert(firstResult.error);
          return;
        }
      }
      
      // Если все ок
      alert('Файл успешно загружен и обработан!');
      await loadNotifications();
      window.dispatchEvent(new CustomEvent('structureUpdated'));
      
    } catch (error) {
      alert('Ошибка загрузки: ' + (error.response?.data?.error || error.message));
    } finally {
      setUploadingPu(null);
    }
  };
  
  input.click();
};

  // ИСПРАВЛЕННАЯ функция определения фаз - без регулярных выражений!
  const getPhaseErrors = useCallback((errorDetails) => {
    const phases = { A: false, B: false, C: false };
    
    if (!errorDetails) return phases;
    
    try {
      let data = null;
      let textToAnalyze = '';
      
      // Пытаемся распарсить JSON
      if (typeof errorDetails === 'string') {
        try {
          const parsed = JSON.parse(errorDetails);
          data = parsed.details || parsed;
          textToAnalyze = parsed.summary || errorDetails;
        } catch {
          textToAnalyze = errorDetails;
        }
      } else if (typeof errorDetails === 'object') {
        data = errorDetails.details || errorDetails;
        textToAnalyze = errorDetails.summary || JSON.stringify(errorDetails);
      }
      
      // Проверяем структурированные данные ТОЛЬКО если есть конкретные фазы
      if (data && typeof data === 'object') {
        if (data.overvoltage) {
          if (data.overvoltage.phase_A && data.overvoltage.phase_A.count > 0) phases.A = true;
          if (data.overvoltage.phase_B && data.overvoltage.phase_B.count > 0) phases.B = true;
          if (data.overvoltage.phase_C && data.overvoltage.phase_C.count > 0) phases.C = true;
        }
        
        if (data.undervoltage) {
          if (data.undervoltage.phase_A && data.undervoltage.phase_A.count > 0) phases.A = true;
          if (data.undervoltage.phase_B && data.undervoltage.phase_B.count > 0) phases.B = true;
          if (data.undervoltage.phase_C && data.undervoltage.phase_C.count > 0) phases.C = true;
        }
      }
      
      // Проверяем текст ТОЛЬКО на явные упоминания конкретных фаз
      if (textToAnalyze) {
        // Только если явно написано "Фаза A" или "phase_A"
        if (textToAnalyze.indexOf('Фаза A') !== -1 || textToAnalyze.indexOf('phase_A') !== -1) phases.A = true;
        if (textToAnalyze.indexOf('Фаза B') !== -1 || textToAnalyze.indexOf('phase_B') !== -1) phases.B = true;
        if (textToAnalyze.indexOf('Фаза C') !== -1 || textToAnalyze.indexOf('phase_C') !== -1) phases.C = true;
      }
    } catch (e) {
      console.error('Error parsing phase errors:', e);
    }
    
    return phases;
  }, []);

  if (loading) return <LoadingSpinner type="dots" message="Загрузка уведомлений..." />;

  const title = filterType === 'error' ? 'Ожидающие мероприятий' : 
                filterType === 'pending_askue' ? 'Ожидающие проверки АСКУЭ' : 
                'Все уведомления';

  // Период уведомления — месяц его появления (createdAt), YYYY-MM.
  const periodKey = (n) => {
    if (!n.createdAt) return null;
    const d = new Date(n.createdAt);
    if (isNaN(d.getTime())) return null;
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
  };
  // Доступные периоды в текущих уведомлениях (с количеством), новые сверху.
  const periodCounts = {};
  notifications.forEach(n => { const k = periodKey(n); if (k) periodCounts[k] = (periodCounts[k] || 0) + 1; });
  const availablePeriods = Object.keys(periodCounts).sort().reverse().map(k => {
    const [y, m] = k.split('-'); return { key: k, label: `${m}.${y}`, count: periodCounts[k] };
  });

  // Фильтрация по ТП + периоду
  const filteredNotifications = notifications.filter(notif => {
    if (periodFilter && periodKey(notif) !== periodFilter) return false;
    if (!searchTp) return true;
    try {
      const data = JSON.parse(notif.message);
      return data.tpName?.toLowerCase().includes(searchTp.toLowerCase());
    } catch {
      return true;
    }
  });

  // Экспорт в Excel того, что СЕЙЧАС на листе (учитывает РЭС/период/поиск).
  const exportNotificationsExcel = () => {
    if (!filteredNotifications.length) { alert('Нет данных для экспорта'); return; }
    const fmtD = (v) => { if (!v) return ''; const d = new Date(v); return isNaN(d.getTime()) ? '' : d.toLocaleDateString('ru-RU'); };
    const posRu = (p) => p === 'start' ? 'Начало' : p === 'middle' ? 'Середина' : p === 'end' ? 'Конец' : (p || '');
    const isAskue = filterType === 'pending_askue';
    const rowsOut = filteredNotifications.map(n => {
      let d = {}; try { d = JSON.parse(n.message); } catch { d = {}; }
      if (isAskue) {
        return {
          'РЭС': d.resName || '', 'ТП': d.tpName || '', 'ВЛ': d.vlName || '',
          'Позиция': posRu(d.position), '№ ПУ': d.puNumber || '',
          'Журнал с': fmtD(d.checkFromDate), 'Появилось': fmtD(n.createdAt),
        };
      }
      const ph = getPhaseErrors(d.details || d.errorDetails) || {};
      return {
        'РЭС': d.resName || '', 'ТП': d.tpName || '', 'ВЛ': d.vlName || '',
        'Позиция': posRu(d.position), '№ ПУ': d.puNumber || '',
        'Фазы с ошибкой': ['A', 'B', 'C'].filter(k => ph[k]).join(', '),
        'Появилось': fmtD(n.createdAt), 'Ошибка': d.errorDetails || d.details || '',
      };
    });
    const sheetName = isAskue ? 'Ожидающие АСКУЭ' : filterType === 'error' ? 'Ожидающие мероприятий' : 'Уведомления';
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rowsOut);
    ws['!cols'] = isAskue
      ? [{ wch: 18 }, { wch: 18 }, { wch: 16 }, { wch: 10 }, { wch: 16 }, { wch: 14 }, { wch: 14 }]
      : [{ wch: 18 }, { wch: 18 }, { wch: 16 }, { wch: 10 }, { wch: 16 }, { wch: 14 }, { wch: 14 }, { wch: 44 }];
    styleExportSheet(ws);
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    const resTag = (user.role === 'admin' && selectedRes) ? `РЭС_${selectedRes}_` : '';
    const perTag = periodFilter ? `${periodFilter}_` : '';
    XLSX.writeFile(wb, `${sheetName}_${resTag}${perTag}${new Date().toLocaleDateString('ru-RU').split('.').join('-')}.xlsx`);
  };

  const headerIcon = filterType === 'error'
    ? <IconWrench size={24} />
    : filterType === 'pending_askue'
    ? <IconClipboard size={24} />
    : <IconBell size={24} />;

  return (
    <div className="notifications">
  <PageHeader icon={headerIcon} title={title} />

  <div className="notifications-controls">
    <div className="search-box">
      <input
        type="text"
        placeholder="Поиск по ТП..."
        value={searchTp}
        onChange={(e) => setSearchTp(e.target.value)}
        className="search-input"
        autoComplete="new-password"
        autoCorrect="off"
        autoCapitalize="off"
        spellCheck="false"
        name={`search-tp-${Date.now()}`}  // Динамическое имя
        id={`search-tp-${Date.now()}`}
      />
    </div>
    
    {canDeleteNotifs && selectedNotificationIds.length > 0 && (
      <button
        className="delete-selected-btn"
        onClick={() => setShowBulkDeleteModal(true)}
      >
        Удалить выбранные ({selectedNotificationIds.length})
      </button>
    )}
  </div>
  
  <div className="select-all-wrapper notif-toolbar">
    {canDeleteNotifs && (
      <label className="notif-selectall">
        <input
          type="checkbox"
          checked={selectedNotificationIds.length === filteredNotifications.length && filteredNotifications.length > 0}
          onChange={handleSelectAll}
        />
        <span>Выбрать все</span>
      </label>
    )}
    <span className="notif-count">
      Показано: <strong>{filteredNotifications.length}</strong>
      {filteredNotifications.length !== notifications.length ? ` из ${notifications.length}` : ''}
    </span>
    {availablePeriods.length > 0 && (
      <label className="notif-period">
        <IconCalendar className="ico" /> Период:
        <select value={periodFilter} onChange={(e) => setPeriodFilter(e.target.value)}>
          <option value="">Все ({notifications.length})</option>
          {availablePeriods.map(p => (
            <option key={p.key} value={p.key}>{p.label} ({p.count})</option>
          ))}
        </select>
      </label>
    )}
    {availablePeriods.length > 0 && (
      <span className="notif-periods-info" title="Периоды (месяцы) в текущих уведомлениях">
        В наличии: {availablePeriods.map(p => `${p.label}·${p.count}`).join(', ')}
      </span>
    )}
    <button
      className="pm-btn pm-btn--excel notif-export-btn"
      onClick={exportNotificationsExcel}
      disabled={!filteredNotifications.length}
      title="Выгрузить в Excel то, что отображается (с учётом РЭС/периода/поиска)"
    >
      <IconDownload className="ico" /> Excel
    </button>
  </div>
      
      <div className="notifications-list">
  {filteredNotifications.map(notif => (
    <div 
      key={notif.id} 
      className={`notification-compact ${notif.type} ${!notif.isRead ? 'unread' : ''} ${selectedNotificationIds.includes(notif.id) ? 'selected' : ''}`}
      onClick={() => {
        // Клик по всему уведомлению = "Детали"
        if (notif.type === 'error' || notif.type === 'pending_askue') {
          try {
            const data = JSON.parse(notif.message);
            setDetailsNotification({ ...notif, data });
            setShowDetailsModal(true);
          } catch (e) { /* некорректный формат — игнорируем */ }
        } else if (notif.type === 'power_overload') {
          setDetailsNotification({ ...notif, data: notif.errorData || {} });
          setShowDetailsModal(true);
        } else if (notif.type === 'problem_vl') {
          if (typeof onSectionChange === 'function') onSectionChange('problem_vl');
        }
      }}
      title={notif.type === 'error' || notif.type === 'pending_askue' || notif.type === 'power_overload' ? 'Открыть подробности' : notif.type === 'problem_vl' ? 'Перейти к проблемным ВЛ' : undefined}
    >
      {/* ЧЕКБОКС ТЕПЕРЬ СНАРУЖИ И СЛЕВА */}
      {canDeleteNotifs && (
        <input
          type="checkbox"
          className="notification-checkbox-left"
          checked={selectedNotificationIds.includes(notif.id)}
          onChange={() => handleSelectNotification(notif.id)}
          onClick={(e) => e.stopPropagation()}
        />
      )}
      
      {/* КОМПАКТНЫЕ УВЕДОМЛЕНИЯ ОБ ОШИБКАХ */}
      {notif.type === 'error' && (() => {
        try {
          const data = JSON.parse(notif.message);
          const phaseErrors = getPhaseErrors(data.details || data.errorDetails);
          
          return (
            <div className="notification-narrow-content">
              {/* УБИРАЕМ ЧЕКБОКС ОТСЮДА */}
              
              <div className="notification-phases">
                <div className={`phase-indicator ${phaseErrors.A ? 'phase-error' : ''}`}>A</div>
                <div className={`phase-indicator ${phaseErrors.B ? 'phase-error' : ''}`}>B</div>
                <div className={`phase-indicator ${phaseErrors.C ? 'phase-error' : ''}`}>C</div>
              </div>
              
              <div className="notification-narrow-info">
                <div className="notification-tp">{data.tpName}</div>
                <div className="notification-narrow-details">
                  <span className="label">РЭС:</span> {data.resName} | 
                  <span className="label"> ТП:</span> {data.tpName} | 
                  <span className="label"> ВЛ:</span> {data.vlName} | 
                  <span className="label"> Позиция:</span> {
                    data.position === 'start' ? 'Начало' : 
                    data.position === 'middle' ? 'Середина' : 'Конец'
                  }
                </div>
                <div className="notification-pu-number">
                  ПУ №: <strong>{data.puNumber}</strong>
                </div>
              </div>
              
              <div className="notification-narrow-actions">
                {/* Кнопка "Детали" убрана — детали открываются кликом по самому уведомлению */}
                {user.role === 'res_responsible' && (
                  <button 
                    className="btn-complete-green"
                    onClick={(e) => {
                      e.stopPropagation();
                      setSelectedNotification({ id: notif.id, data });
                      setComment('');
                      setCommentError(false);
                      setShowCompleteModal(true);
                    }}
                    title="Выполнить мероприятия"
                  >
                    Завершить
                  </button>
                )}
                
                {/* УБИРАЕМ КНОПКУ УДАЛЕНИЯ */}
              </div>
            </div>
          );
        } catch (e) {
          return <div className="error-text">Ошибка отображения</div>;
        }
      })()}
      
      {/* КОМПАКТНЫЕ УВЕДОМЛЕНИЯ АСКУЭ */}
      {notif.type === 'pending_askue' && (() => {
        try {
          const data = JSON.parse(notif.message);
          return (
            <div className="notification-compact-content askue">
              {/* УБИРАЕМ ЧЕКБОКС ОТСЮДА */}
              
              <div className="notification-main-info">
                <div className="notification-location">
                  <span className="label">ТП:</span> {data.tpName} | 
                  <span className="label"> ПУ №:</span> <strong>{data.puNumber}</strong> | 
                  <span className="label"> Журнал с:</span> <strong>{new Date(data.checkFromDate).toLocaleDateString('ru-RU')}</strong>
                </div>
              </div>
              
              <div className="notification-actions-row">
                <div className="notification-buttons">
                  <button 
                    className="btn-upload-orange"  // Изменили класс
                    onClick={(e) => { e.stopPropagation(); handleFileUpload(data.puNumber, data); }}
                    disabled={uploadingPu === data.puNumber}
                    title="Загрузить файл"
                  >
                    {uploadingPu === data.puNumber ? 'Загрузка...' : 'Загрузить'}
                  </button>
                  
                  {/* Кнопка "Детали" убрана — детали открываются кликом по самому уведомлению */}
                  {/* УБИРАЕМ КНОПКУ УДАЛЕНИЯ */}
                </div>
              </div>
            </div>
          );
        } catch (e) {
          return <div className="error-text">Ошибка отображения</div>;
        }
      })()}

      {/* УВЕДОМЛЕНИЯ О ПРОБЛЕМНЫХ ВЛ */}
      {notif.type === 'problem_vl' && (() => {
        try {
          const data = JSON.parse(notif.message);
          return (
            <div className="notification-compact-content problem-vl">
              {/* УБИРАЕМ ЧЕКБОКС ОТСЮДА */}
              
              <div className="problem-vl-alert">
                <span className="critical-icon"><IconAlertCircle className="ico" style={{color:'var(--red)'}} /></span>
                <div className="problem-vl-header">
                  <h4>Критическая проблема!</h4>
                  <span className="failure-count">{data.failureCount} неудачных проверок</span>
                </div>
              </div>
              
              <div className="notification-main-info">
                <div className="notification-location">
                  <span className="label">РЭС:</span> {data.resName} | 
                  <span className="label"> ТП:</span> {data.tpName} | 
                  <span className="label"> ВЛ:</span> {data.vlName}
                </div>
                <div className="notification-pu">
                  <span className="label">ПУ №:</span> <strong>{data.puNumber}</strong> | 
                  <span className="label"> Позиция:</span> {
                    data.position === 'start' ? 'Начало' :
                    data.position === 'middle' ? 'Середина' : 'Конец'
                  }
                </div>
              </div>
              
              <div className="problem-error-details">
                <p className="error-label">Последняя ошибка:</p>
                <p className="error-text">{data.errorDetails}</p>
              </div>
              
              {data.resComment && (
                <div className="problem-res-comment">
                  <p className="comment-label">Комментарий РЭС:</p>
                  <p className="comment-text">{data.resComment}</p>
                </div>
              )}
              
              <div className="notification-actions-row">
                <div className="notification-buttons">
                  <button 
                    className="btn-view-problem"
                    onClick={(e) => {
                      e.stopPropagation();
                      if (typeof onSectionChange === 'function') {
                        onSectionChange('problem_vl');
                      }
                    }}
                    title="Перейти к проблемным ВЛ"
                  >
                    <IconChart className="ico" /> К проблемным ВЛ
                  </button>
                  
                  {/* УБИРАЕМ КНОПКУ УДАЛЕНИЯ */}
                </div>
              </div>
            </div>
          );
        } catch (e) {
          console.error('Error parsing problem VL notification:', e);
          return <div className="error-text">Ошибка отображения уведомления</div>;
        }
      })()}
            
            {/* ПЕРЕГРУЗ СЕКЦИИ ПО ПРОФИЛЮ МОЩНОСТИ (этап 2) */}
            {notif.type === 'power_overload' && (() => {
              const data = notif.errorData || {};
              const f1 = (v) => (v == null ? '—' : Number(v).toFixed(1));
              const ratioPct = data.ratio != null ? Math.round(data.ratio * 100) : null;
              return (
                <div className="notification-compact-content problem-vl">
                  <div className="problem-vl-alert">
                    <span className="critical-icon"><IconAlertTriangle className="ico" style={{ color: 'var(--red)' }} /></span>
                    <div className="problem-vl-header">
                      <h4>Превышение Pном</h4>
                      {ratioPct != null && <span className="failure-count">{ratioPct}% от лимита</span>}
                    </div>
                  </div>
                  <div className="notification-main-info">
                    <div className="notification-location">
                      <span className="label">ТП:</span> {data.tpName} · <strong>СШ-{toRoman(data.sectionNumber)}</strong>
                    </div>
                    <div className="notification-pu">
                      пик <strong>{f1(data.peakKw)} кВт</strong> при лимите <strong>{f1(data.limitKw)} кВт</strong>
                      {data.peakAt ? <> · {data.peakAt}</> : null}
                    </div>
                  </div>
                </div>
              );
            })()}

            {/* УСПЕШНЫЕ УВЕДОМЛЕНИЯ */}
            {notif.type === 'success' && (
              <div className="notification-compact-content success">
                <div className="success-icon"><IconCheck className="ico" style={{color:'var(--green)'}} /></div>
                <div className="success-text">{notif.message}</div>
              </div>
            )}

            {/* ИНФОРМАЦИОННЫЕ УВЕДОМЛЕНИЯ */}
            {notif.type === 'info' && (
              <div className="notification-compact-content info">
                <div className="info-icon"><IconInfo className="ico" style={{color:'var(--blue)'}} /></div>
                <div className="info-text">{notif.message}</div>
              </div>
            )}
          </div>
        ))}
      </div>

      {/* Модальное окно деталей */}
      {showDetailsModal && detailsNotification && (
        <ModalShell title="Подробная информация" className="details-modal"
          icon={<IconFileText size={16} />} onClose={() => setShowDetailsModal(false)}>
            <div className="modal-body">
              {detailsNotification.type === 'error' && (
                <>
                  {/* Показываем фазы в детальном окне */}
                  <div className="phase-indicators-large">
                    {(() => {
                      const phases = { A: false, B: false, C: false };
                      
                      // Проверяем только явные упоминания фаз
                      if (detailsNotification.data.details && typeof detailsNotification.data.details === 'object') {
                        const details = detailsNotification.data.details;
                        if (details.overvoltage) {
                          if (details.overvoltage.phase_A && details.overvoltage.phase_A.count > 0) phases.A = true;
                          if (details.overvoltage.phase_B && details.overvoltage.phase_B.count > 0) phases.B = true;
                          if (details.overvoltage.phase_C && details.overvoltage.phase_C.count > 0) phases.C = true;
                        }
                        if (details.undervoltage) {
                          if (details.undervoltage.phase_A && details.undervoltage.phase_A.count > 0) phases.A = true;
                          if (details.undervoltage.phase_B && details.undervoltage.phase_B.count > 0) phases.B = true;
                          if (details.undervoltage.phase_C && details.undervoltage.phase_C.count > 0) phases.C = true;
                        }
                      }
                      
                      const errorText = detailsNotification.data.errorDetails || '';
                      if (errorText.indexOf('Фаза A') !== -1 || errorText.indexOf('phase_A') !== -1) phases.A = true;
                      if (errorText.indexOf('Фаза B') !== -1 || errorText.indexOf('phase_B') !== -1) phases.B = true;
                      if (errorText.indexOf('Фаза C') !== -1 || errorText.indexOf('phase_C') !== -1) phases.C = true;
                      
                      return (
                        <>
                          <div className={`phase-indicator ${phases.A ? 'phase-error' : ''}`}>A</div>
                          <div className={`phase-indicator ${phases.B ? 'phase-error' : ''}`}>B</div>
                          <div className={`phase-indicator ${phases.C ? 'phase-error' : ''}`}>C</div>
                        </>
                      );
                    })()}
                  </div>
                  
                  <div className="detail-row">
                    <strong>РЭС:</strong> {detailsNotification.data.resName}
                  </div>
                  <div className="detail-row">
                    <strong>ТП:</strong> {detailsNotification.data.tpName}
                  </div>
                  <div className="detail-row">
                    <strong>Фидер:</strong> {detailsNotification.data.vlName}
                  </div>
                  <div className="detail-row">
                    <strong>ПУ №:</strong> {detailsNotification.data.puNumber}
                  </div>
                  <div className="detail-row">
                    <strong>Позиция:</strong> {
                      detailsNotification.data.position === 'start' ? 'Начало' :
                      detailsNotification.data.position === 'middle' ? 'Середина' : 'Конец'
                    }
                  </div>
                  <div className="error-details-box">
                    <strong>Детали ошибки:</strong>
                    <p>{detailsNotification.data.errorDetails}</p>
                  </div>
                </>
              )}
              
              {detailsNotification.type === 'pending_askue' && (
                <>
                  <div className="askue-details-content">
                    <h4><IconZap className="ico" style={{color:'var(--amber)'}} /> Требуется снять журнал событий</h4>
                    <div className="detail-row">
                      <strong>ПУ №:</strong> {detailsNotification.data.puNumber}
                    </div>
                    <div className="detail-row">
                      <strong>ТП:</strong> {detailsNotification.data.tpName}
                    </div>
                    <div className="detail-row">
                      <strong>Фидер:</strong> {detailsNotification.data.vlName}
                    </div>
                    <div className="highlight-box">
                      <strong><IconCalendar className="ico" /> Журнал событий с даты:</strong>
                      <p>{new Date(detailsNotification.data.checkFromDate).toLocaleDateString('ru-RU')}</p>
                    </div>
                    <div className="highlight-box">
                      <strong><IconMessage className="ico" /> Комментарий РЭС:</strong>
                      <p>{detailsNotification.data.completedComment}</p>
                    </div>
                    <div className="detail-row">
                      <strong>Мероприятия выполнены:</strong> {new Date(detailsNotification.data.completedAt).toLocaleString('ru-RU')}
                    </div>
                  </div>
                </>
              )}

              {detailsNotification.type === 'power_overload' && (() => {
                const d = detailsNotification.data || {};
                const f1 = (v) => (v == null ? '—' : Number(v).toFixed(1));
                const ratioPct = d.ratio != null ? Math.round(d.ratio * 100) : null;
                return (
                  <div className="askue-details-content">
                    <h4><IconAlertTriangle className="ico" style={{ color: 'var(--red)' }} /> Превышение номинальной мощности</h4>
                    <div className="detail-row"><strong>ТП:</strong> {d.tpName} · СШ-{toRoman(d.sectionNumber)}</div>
                    <div className="detail-row"><strong>ПУ техучёта:</strong> {d.techPuNumber || '—'}</div>
                    <div className="detail-row"><strong>Пик:</strong> {f1(d.peakKw)} кВт{d.peakAt ? ` (${d.peakAt})` : ''}</div>
                    <div className="detail-row"><strong>Лимит:</strong> {f1(d.limitKw)} кВт {ratioPct != null ? `(${ratioPct}% от лимита)` : ''}</div>
                    <div className="detail-row"><strong>Sном тр-ра:</strong> {d.tnKva != null ? `${d.tnKva} кВА` : '—'} · cosφ {d.cosPhi ?? '—'}</div>
                    {d.period && <div className="detail-row"><strong>Период:</strong> {d.period}</div>}
                  </div>
                );
              })()}
            </div>

            <div className="modal-footer">
              <button className="action-btn" onClick={() => setShowDetailsModal(false)}>
                Закрыть
              </button>
            </div>
        </ModalShell>
      )}

      {/* Модальное окно для выполнения мероприятий */}
      {showCompleteModal && selectedNotification && (
        <ModalShell title="Отметить выполнение мероприятий" className="complete-work-modal"
          icon={<IconWrench size={16} />} onClose={() => setShowCompleteModal(false)}>
            <div className="modal-body">
              <div className="work-info">
                <p><strong>ТП:</strong> {selectedNotification.data.tpName}</p>
                <p><strong>ВЛ:</strong> {selectedNotification.data.vlName}</p>
                <p><strong>ПУ №:</strong> {selectedNotification.data.puNumber}</p>
              </div>
              
              <div className="form-group">
                <label style={commentError ? { color: '#dc2626', fontWeight: 600 } : undefined}>
                  Что было выполнено? (минимум 5 слов)
                </label>
                <textarea
                  value={comment}
                  onChange={(e) => {
                    const v = e.target.value;
                    setComment(v);
                    // как только набрано 5+ слов — гасим красную подсветку
                    if (commentError && v.trim().split(' ').filter(w => w.length > 0).length >= 5) {
                      setCommentError(false);
                    }
                  }}
                  placeholder="Опишите выполненные работы..."
                  rows={4}
                  style={commentError ? {
                    borderColor: '#dc2626',
                    boxShadow: '0 0 0 3px rgba(220,38,38,0.2)',
                    outline: 'none'
                  } : undefined}
                />
                <small className="word-count" style={commentError ? { color: '#dc2626', fontWeight: 600 } : undefined}>
                  Слов: {comment.trim().split(' ').filter(w => w.length > 0).length} из 5
                  {commentError ? ' — нужно не менее 5 слов, чтобы завершить' : ''}
                </small>
              </div>
              
              <div className="form-group">
                <label>Журнал событий требуется с даты:</label>
                <input
                  type="date"
                  value={checkFromDate}
                  onChange={(e) => setCheckFromDate(e.target.value)}
                />
              </div>
              
              <div className="form-group">
                <label>Прикрепить фото/документы (макс. 5 файлов по 10MB)</label>
                <input
                  type="file"
                  multiple
                  accept="image/*,application/pdf"
                  onChange={(e) => {
                    const files = Array.from(e.target.files).slice(0, 5);
                    setAttachedFiles(files);
                  }}
                />
                {attachedFiles.length > 0 && (
                  <div className="attached-files-list">
                    <p>Выбрано файлов: {attachedFiles.length}</p>
                    {attachedFiles.map((file, idx) => (
                      <div key={idx} className="attached-file-item">
                        {file.name} ({(file.size / 1024 / 1024).toFixed(2)} MB)
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
            
            <div className="modal-footer">
              <button className="cancel-btn" onClick={() => setShowCompleteModal(false)}>
                Отмена
                </button>
              <button
                className="confirm-btn"
                onClick={handleCompleteWork}
                disabled={submitting}
              >
                {submitting ? 'Отправка...' : 'Подтвердить выполнение'}
                </button>
            </div>
        </ModalShell>
      )}

      {/* Модальное окно для удаления */}
      {showDeleteModal && (
        <ModalShell variant="confirm" title="Подтверждение удаления" className="delete-modal"
          onClose={() => {setShowDeleteModal(false); setDeletePassword('');}}>
            <div className="modal-body">
              <p>Вы собираетесь удалить это уведомление.</p>
              <p className="warning"><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /> Это действие нельзя отменить!</p>
              <div className="form-group">
                <label>Введите пароль администратора:</label>
                <input
                  type="password"
                  value={deletePassword}
                  onChange={(e) => setDeletePassword(e.target.value)}
                  placeholder="Пароль"
                  autoFocus
                  autoComplete="new-password"    
                  name="delete-notification-password"  
                />
              </div>
            </div>
            <div className="modal-footer">
              <button className="cancel-btn" onClick={() => {setShowDeleteModal(false); setDeletePassword('');}}>
                Отмена
              </button>
              <button 
                className="danger-btn" 
                onClick={handleDeleteNotification}
                disabled={!deletePassword}
              >
                Удалить
              </button>
            </div>
        </ModalShell>
      )}

      {/* ДОБАВЬТЕ ЭТО МОДАЛЬНОЕ ОКНО: */}
      {showBulkDeleteModal && (
        <ModalShell variant="confirm" title="Подтверждение удаления" className="delete-modal"
          onClose={() => {setShowBulkDeleteModal(false); setBulkDeletePassword('');}}>
            <div className="modal-body">
              <p>Вы собираетесь удалить {selectedNotificationIds.length} уведомлений.</p>
              <p className="warning"><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /> Это действие нельзя отменить!</p>
              <div className="form-group">
                <label>Введите пароль администратора:</label>
                <input
                  type="password"
                  value={bulkDeletePassword}
                  onChange={(e) => setBulkDeletePassword(e.target.value)}
                  placeholder="Пароль"
                  autoFocus
                  autoComplete="new-password"  // Добавить
                  name={`delete-password-${Date.now()}`}
                />
              </div>
            </div>
            <div className="modal-footer">
              <button className="cancel-btn" onClick={() => {setShowBulkDeleteModal(false); setBulkDeletePassword('');}}>
                Отмена
              </button>
              <button 
                className="danger-btn" 
                onClick={handleBulkDelete}
                disabled={!bulkDeletePassword}
              >
                Удалить выбранные
              </button>
            </div>
        </ModalShell>
      )}

      {/* ДОБАВЬТЕ КНОПКУ ПРОКРУТКИ: */}
      {showScrollTop && (
        <button 
          className="scroll-to-top"
          onClick={() => {
            const contentElement = document.querySelector('.content');
            if (contentElement) {
              contentElement.scrollTo({ top: 0, behavior: 'smooth' });
            }
          }}
          title="Наверх"
        >
          <IconArrowUp className="ico" />
        </button>
      )}
    </div>
  );
}

    

// =====================================================
// КОМПОНЕНТ ОТЧЕТОВ
// =====================================================

function Reports() {
  const [showCommentModal, setShowCommentModal] = useState(false);
  const [selectedComment, setSelectedComment] = useState(null);
  const { user, selectedRes } = useContext(AuthContext);

  const [reportType, setReportType] = useState('pending_work');
  const [reportData, setReportData] = useState([]);
  
  const [loading, setLoading] = useState(false);
  const [dateFrom, setDateFrom] = useState(
    new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]
  );
  const [dateTo, setDateTo] = useState(new Date().toISOString().split('T')[0]);
  const [searchTp, setSearchTp] = useState('');
  
  const [showFileViewer, setShowFileViewer] = useState(false);
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [currentFileIndex, setCurrentFileIndex] = useState(0);
  
  // НОВОЕ: определяем нужны ли даты для этого типа отчета
  const needsDateFilter = () => {
    return reportType === 'completed' || reportType === 'problem_vl';
  };
  
  useEffect(() => {
    loadReports();
  }, [reportType, dateFrom, dateTo, selectedRes]);

  const loadReports = async () => {
    setLoading(true);
    try {
      let response;
      
      // ИСПРАВЛЕНО: передаем даты только для completed и problem_vl
      const params = {
        type: reportType,
        resId: user.role === 'admin' ? selectedRes : user.resId
      };
      
      // Добавляем даты только если нужно
      if (needsDateFilter()) {
        params.dateFrom = dateFrom;
        params.dateTo = dateTo;
      }
      
      if (reportType === 'problem_vl') {
        response = await api.get('/api/reports/problem-vl', { params });
      } else if (reportType === 'power_overload') {
        response = await api.get('/api/reports/overload', { params });
      } else {
        response = await api.get('/api/reports/detailed', { params });
      }

      setReportData(response.data);
    } catch (error) {
      console.error('Error loading reports:', error);
      setReportData([]);
    } finally {
      setLoading(false);
    }
  };
  
  // Функция для открытия просмотра файлов
  const viewAttachments = (attachments) => {
    console.log('Viewing attachments:', attachments);
    
    if (attachments && attachments.length > 0) {
      setSelectedFiles(attachments);
      setCurrentFileIndex(0);
      setShowFileViewer(true);
    }
  };
  
  // ИСПРАВЛЕННАЯ функция exportToExcel
  const exportToExcel = () => {
  try {
    console.log('=== EXPORT START ===');
    console.log('Report type:', reportType);
    console.log('Filtered data length:', filteredData.length);
    console.log('First item:', filteredData[0]);
    
    if (filteredData.length === 0) {
      alert('Нет данных для экспорта');
      return;
    }

    // Подготавливаем данные для экспорта
    const exportData = filteredData.map((item, index) => {
      console.log(`Processing item ${index}:`, item);
      
      if (reportType === 'power_overload') {
        return {
          'РЭС': item.resName || '',
          'ТП': item.tpName || '',
          'СШ': item.sectionNumber,
          'Sном, кВА': item.tnKva,
          'cosφ': item.cosPhi,
          'Лимит, кВт': item.limitKw,
          'Последний пик, кВт': item.lastPeakKw ?? '',
          'Дата пика': item.lastPeakAt ? new Date(item.lastPeakAt).toLocaleString('ru-RU') : '',
          '%': item.ratioPct != null ? item.ratioPct + '%' : '',
          'Статус случая': item.caseStage || '—',
          'Дата АСКУЭ': item.askueCompletedAt ? new Date(item.askueCompletedAt).toLocaleDateString('ru-RU') : '',
          'Дата РЭС': item.resCompletedAt ? new Date(item.resCompletedAt).toLocaleDateString('ru-RU') : '',
          'Результат перепроверки': item.recheckResult || '—',
          'Циклы': item.cycles || 0
        };
      } else if (reportType === 'problem_vl') {
        return {
          'РЭС': item.resName || '',
          'ТП': item.tpName || '',
          'ВЛ': item.vlName || '',
          'Позиция': item.position || '', // УЖЕ преобразовано на бэкенде
          'Номер ПУ': item.puNumber || '',
          'Количество неудачных проверок': item.failureCount || 0,
          'Дата первого обращения': item.firstReportDate ?
            new Date(item.firstReportDate).toLocaleDateString('ru-RU') : '',
          'Дата последней проверки': item.lastErrorDate ?
            new Date(item.lastErrorDate).toLocaleDateString('ru-RU') : '',
          'Последняя ошибка': item.lastErrorDetails || '',
          'Статус проблемы': item.status || ''
        };
      } else if (reportType === 'pending_work') {
        return {
          'РЭС': item.resName || '',
          'ТП': item.tpName || '',
          'ВЛ': item.vlName || '',
          'Позиция': item.position === 'start' ? 'Начало' : 
                     item.position === 'middle' ? 'Середина' : 'Конец',
          'Номер ПУ': item.puNumber || '',
          'Ошибка': item.errorDetails || '',
          'Дата обнаружения': formatDate(item.errorDate)
        };
      } else if (reportType === 'pending_askue') {
        return {
          'РЭС': item.resName || '',
          'ТП': item.tpName || '',
          'ВЛ': item.vlName || '',
          'Позиция': item.position === 'start' ? 'Начало' : 
                     item.position === 'middle' ? 'Середина' : 'Конец',
          'Номер ПУ': item.puNumber || '',
          'Ошибка': item.errorDetails || '',
          'Дата обнаружения': formatDate(item.errorDate),
          'Комментарий РЭС': item.resComment || '',
          'Дата завершения мероприятий': formatDate(item.workCompletedDate)
        };
      } else if (reportType === 'completed') {
        return {
          'РЭС': item.resName || '',
          'ТП': item.tpName || '',
          'ВЛ': item.vlName || '',
          'Позиция': item.position === 'start' ? 'Начало' : 
                     item.position === 'middle' ? 'Середина' : 'Конец',
          'Номер ПУ': item.puNumber || '',
          'Ошибка': item.errorDetails || '',
          'Дата обнаружения': formatDate(item.errorDate),
          'Комментарий РЭС': item.resComment || '',
          'Дата завершения мероприятий': formatDate(item.workCompletedDate),
          'Дата перепроверки': formatDate(item.recheckDate),
          'Результат': item.recheckResult === 'ok' ? 'Исправлено' : 'Не исправлено'
        };
      }
    });

    console.log('Export data prepared:', exportData);
    console.log('Export data length:', exportData.length);

    // Проверяем наличие библиотеки XLSX
    if (typeof XLSX === 'undefined') {
      console.error('XLSX library not found!');
      alert('Ошибка: библиотека XLSX не загружена. Обновите страницу.');
      return;
    }

    // Создаем новую книгу Excel
    const wb = XLSX.utils.book_new();
    
    // Создаем лист с данными
    const ws = XLSX.utils.json_to_sheet(exportData);
    
    // Устанавливаем ширину колонок в зависимости от типа отчета
    let columnWidths;
    
    if (reportType === 'problem_vl') {
      columnWidths = [
        { wch: 20 }, // РЭС
        { wch: 15 }, // ТП
        { wch: 15 }, // ВЛ
        { wch: 12 }, // Позиция
        { wch: 15 }, // Номер ПУ
        { wch: 30 }, // Количество неудачных проверок
        { wch: 20 }, // Дата первого обращения
        { wch: 20 }, // Дата последней проверки
        { wch: 60 }, // Последняя ошибка
        { wch: 15 }  // Статус проблемы
      ];
    } else if (reportType === 'pending_work') {
      columnWidths = [
        { wch: 20 }, // РЭС
        { wch: 15 }, // ТП
        { wch: 15 }, // ВЛ
        { wch: 12 }, // Позиция
        { wch: 15 }, // Номер ПУ
        { wch: 50 }, // Ошибка
        { wch: 18 }  // Дата обнаружения
      ];
    } else if (reportType === 'pending_askue') {
      columnWidths = [
        { wch: 20 }, { wch: 15 }, { wch: 15 }, { wch: 12 }, { wch: 15 },
        { wch: 50 }, { wch: 18 }, { wch: 40 }, { wch: 25 }
      ];
    } else if (reportType === 'completed') {
      columnWidths = [
        { wch: 20 }, { wch: 15 }, { wch: 15 }, { wch: 12 }, { wch: 15 },
        { wch: 50 }, { wch: 18 }, { wch: 40 }, { wch: 25 }, { wch: 18 }, { wch: 15 }
      ];
    }
    
    ws['!cols'] = columnWidths;
    
    // Добавляем лист в книгу
    const sheetName = getReportTitle();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);
    
    // Генерируем имя файла
    const fileName = `Отчет_${sheetName}_${new Date().toLocaleDateString('ru-RU').split('.').join('-')}.xlsx`;
    
    console.log('Saving file:', fileName);
    
    // Сохраняем файл
    XLSX.writeFile(wb, fileName);
    
    console.log('=== EXPORT SUCCESS ===');
    
    // Показываем уведомление
    alert(`Отчет успешно экспортирован!\n\nФайл: ${fileName}\nЗаписей: ${exportData.length}`);
    
  } catch (error) {
    console.error('EXPORT ERROR:', error);
    console.error('Error stack:', error.stack);
    alert(`Ошибка экспорта: ${error.message}\n\nПроверьте консоль браузера (F12) для деталей.`);
  }
};

  // Вспомогательная функция для форматирования даты
  const formatDate = (dateString) => {
    if (!dateString) return '';
    return new Date(dateString).toLocaleString('ru-RU', {
      day: '2-digit',
      month: '2-digit',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
  };

  const getReportTitle = () => {
    switch (reportType) {
      case 'pending_work':
        return 'Ожидающие мероприятий';
      case 'pending_askue':
        return 'Ожидающие проверки АСКУЭ';
      case 'completed':
        return 'Завершенные проверки';
      case 'problem_vl':
        return 'Проблемные ВЛ';
      case 'power_overload':
        return 'Превышение Pном';
      default:
        return 'Отчет';
    }
      return title.length > 31 ? title.substring(0, 28) + '...' : title;
  };

  // Фильтрация по ТП с мемоизацией
  const filteredData = useMemo(() => 
    reportData.filter(item => 
      !searchTp || item.tpName?.toLowerCase().includes(searchTp.toLowerCase())
    ), [reportData, searchTp]
  );

  if (loading) return <LoadingSpinner message="Формирование отчета..." submessage="Собираем данные из базы" />;

  return (
    <div className="reports">
      <PageHeader icon={<IconFileText size={24} />} title="Отчеты по проверкам" />

      {user.role !== 'admin' && (
        <div className="res-indicator">
          <span>Показаны данные для: <strong>{user.resName}</strong></span>
        </div>
      )}
      
      <div className="report-controls">
        <div className="control-group">
          <label>Тип отчета:</label>
          <select value={reportType} onChange={(e) => setReportType(e.target.value)}>
            <option value="pending_work">Ожидающие мероприятий</option>
            <option value="pending_askue">Ожидающие проверки АСКУЭ</option>
            <option value="completed">Завершенные проверки</option>
            <option value="problem_vl">Проблемные ВЛ (2+ ошибки)</option>
            <option value="power_overload">Превышение Pном</option>
          </select>
        </div>
        
        {/* ИСПРАВЛЕНО: Показываем даты только для completed и problem_vl */}
        {needsDateFilter() && (
          <>
            <div className="control-group">
              <label>Период с:</label>
              <input 
                type="date" 
                value={dateFrom}
                onChange={(e) => setDateFrom(e.target.value)}
              />
            </div>
            
            <div className="control-group">
              <label>по:</label>
              <input 
                type="date" 
                value={dateTo}
                onChange={(e) => setDateTo(e.target.value)}
              />
            </div>
          </>
        )}
        
        <div className="control-group">
          <input 
            type="text"
            placeholder="Поиск по ТП..."
            value={searchTp}
            onChange={(e) => setSearchTp(e.target.value)}
            className="search-input"
          />
        </div>
        
        <button className="export-btn" onClick={exportToExcel}>
          <IconChart className="ico" /> Экспорт в Excel
        </button>
      </div>
      
      <div className="report-summary">
        <h3>{getReportTitle()}</h3>
        <p>Найдено записей: {filteredData.length}</p>
        {/* ДОБАВЛЕНО: Подсказка для пользователя */}
        {!needsDateFilter() && (
          <p className="info-hint">
            <IconInfo className="ico" style={{color:'var(--blue)'}} /> Отчет показывает текущее состояние на момент формирования
          </p>
        )}
      </div>
      
      <div className="report-table-wrapper" style={{ position: 'relative' }}>
        {loading && <LoadingSpinner type="overlay" message="Обновление данных..." />}
  
        <div className={`report-table ${loading ? 'loading' : ''}`}>
          {reportType === 'power_overload' ? (
            <table>
              <thead>
                <tr>
                  <th>РЭС</th><th>ТП</th><th>СШ</th><th>Sном, кВА</th><th>cosφ</th>
                  <th>Лимит, кВт</th><th>Пик, кВт</th><th>Дата пика</th><th>%</th>
                  <th>Статус случая</th><th>АСКУЭ</th><th>РЭС</th><th>Перепроверка</th><th>Циклы</th>
                </tr>
              </thead>
              <tbody>
                {filteredData.map((item, idx) => (
                  <tr key={idx}>
                    <td>{item.resName}</td>
                    <td>{item.tpName}</td>
                    <td>СШ-{toRoman(item.sectionNumber)}</td>
                    <td>{item.tnKva}</td>
                    <td>{item.cosPhi}</td>
                    <td>{item.limitKw}</td>
                    <td>{item.lastPeakKw ?? '—'}</td>
                    <td>{item.lastPeakAt ? new Date(item.lastPeakAt).toLocaleString('ru-RU') : '—'}</td>
                    <td>{item.ratioPct != null ? item.ratioPct + '%' : '—'}</td>
                    <td>{item.caseStage}</td>
                    <td>{item.askueCompletedAt ? new Date(item.askueCompletedAt).toLocaleDateString('ru-RU') : '—'}</td>
                    <td>{item.resCompletedAt ? new Date(item.resCompletedAt).toLocaleDateString('ru-RU') : '—'}</td>
                    <td>{item.recheckResult}</td>
                    <td>{item.cycles}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          ) : (
          <table>
            <thead>
              <tr>
                <th>РЭС</th>
                <th>ТП</th>
                <th>ВЛ</th>
                <th>Позиция</th>
                <th>Номер ПУ</th>
                
                {/* Разные колонки для разных типов отчетов */}
                {reportType === 'problem_vl' ? (
                  <>
                    <th>Кол-во ошибок</th>
                    <th>Первое обращение</th>
                    <th>Последняя проверка</th>
                    <th>Последняя ошибка</th>
                    <th>Статус</th>
                  </>
                ) : reportType === 'pending_work' ? (
                  <>
                    <th>Ошибка</th>
                    <th>Дата обнаружения</th>
                  </>
                ) : reportType === 'pending_askue' ? (
                  <>
                    <th>Ошибка</th>
                    <th>Дата обнаружения</th>
                    <th>Комментарий РЭС</th>
                    <th>Дата завершения мероприятий</th>
                  </>
                ) : reportType === 'completed' ? (
                  <>
                    <th>Ошибка</th>
                    <th>Дата обнаружения</th>
                    <th>Комментарий РЭС</th>
                    <th>Дата завершения мероприятий</th>
                    <th>Дата перепроверки</th>
                    <th>Результат</th>
                    <th>Файлы</th>
                  </>
                ) : null}
              </tr>
            </thead>
            <tbody>
              {filteredData.map((item, idx) => (
                <tr key={idx}>
                  <td>{item.resName}</td>
                  <td>{item.tpName}</td>
                  <td>{item.vlName}</td>
                  <td>{item.position === 'start' ? 'Начало' : item.position === 'middle' ? 'Середина' : 'Конец'}</td>
                  <td>{item.puNumber}</td>
                  
                  {/* Данные для проблемных ВЛ */}
                  {reportType === 'problem_vl' ? (
                    <>
                      <td>
                        <span className="failure-count-badge">{item.failureCount}</span>
                      </td>
                      <td>{new Date(item.firstReportDate).toLocaleDateString('ru-RU')}</td>
                      <td>{new Date(item.lastErrorDate).toLocaleDateString('ru-RU')}</td>
                      <td className="error-cell">{item.lastErrorDetails}</td>
                      <td>
                        <span className={`status-badge ${
                          item.status === 'Активная' ? 'status-active' : 
                          item.status === 'Решена' ? 'status-resolved' : 
                          'status-dismissed'
                        }`}>
                          {item.status}
                        </span>
                      </td>
                    </>
                  
                  /* Данные для ожидающих мероприятий */
                  ) : reportType === 'pending_work' ? (
                    <>
                      <td className="error-cell">{item.errorDetails}</td>
                      <td>{new Date(item.errorDate).toLocaleDateString('ru-RU')}</td>
                    </>
                  
                  /* Данные для ожидающих АСКУЭ */
                  ) : reportType === 'pending_askue' ? (
                    <>
                      <td className="error-cell">{item.errorDetails}</td>
                      <td>{new Date(item.errorDate).toLocaleDateString('ru-RU')}</td>
                      <td>{item.resComment}</td>
                      <td>{new Date(item.workCompletedDate).toLocaleDateString('ru-RU')}</td>
                    </>
                  
                  /* Данные для завершенных проверок */
                  ) : reportType === 'completed' ? (
                    <>
                      <td className="error-cell">{item.errorDetails}</td>
                      <td>{new Date(item.errorDate).toLocaleDateString('ru-RU')}</td>
                      <td>{item.resComment}</td>
                      <td>{new Date(item.workCompletedDate).toLocaleDateString('ru-RU')}</td>
                      <td>{new Date(item.recheckDate).toLocaleDateString('ru-RU')}</td>
                      <td className="status-cell">
                        <span 
                          className={item.recheckResult === 'ok' ? 'status-ok clickable' : 'status-error clickable'}
                          onClick={() => {
                            setSelectedComment({
                              comment: item.resComment,
                              tpName: item.tpName,
                              vlName: item.vlName,
                              puNumber: item.puNumber,
                              result: item.recheckResult
                            });
                            setShowCommentModal(true);
                          }}
                          style={{ cursor: 'pointer' }}
                          title="Нажмите для просмотра комментария"
                        >
                          {item.recheckResult === 'ok' ? 'Исправлено' : 'Не исправлено'}
                        </span>
                      </td>
                      <td>
                        {item.attachments && item.attachments.length > 0 ? (
                          <button 
                            className="btn-view-files"
                            onClick={() => viewAttachments(item.attachments)}
                          >
                            <IconPaperclip className="ico" /> {item.attachments.length} файл(ов)
                          </button>
                        ) : (
                          <span className="no-files">—</span>
                        )}
                      </td>
                    </>
                  ) : null}
                </tr>
              ))}
            </tbody>
          </table>
          )}
        </div>
      </div>

      {filteredData.length === 0 && (
        <div className="no-data">
          <p>Нет данных для отображения {needsDateFilter() ? 'за выбранный период' : 'на данный момент'}</p>
        </div>
      )}
      
      {showFileViewer && (
        <FileViewer 
          files={selectedFiles}
          currentIndex={currentFileIndex}
          onClose={() => setShowFileViewer(false)}
          onNext={() => setCurrentFileIndex((prev) => (prev + 1) % selectedFiles.length)}
          onPrev={() => setCurrentFileIndex((prev) => (prev - 1 + selectedFiles.length) % selectedFiles.length)}
        />
      )}
    
      {/* Модальное окно для комментария */}
      {showCommentModal && selectedComment && (
        <ModalShell title="Комментарий РЭС" className="comment-modal"
          icon={<IconMessage size={16} />} onClose={() => setShowCommentModal(false)}>
            <div className="modal-body">
              <div className="comment-info">
                <p><strong>ТП:</strong> {selectedComment.tpName}</p>
                <p><strong>ВЛ:</strong> {selectedComment.vlName}</p>
                <p><strong>ПУ №:</strong> {selectedComment.puNumber}</p>
                <p><strong>Результат:</strong> 
                  <span className={selectedComment.result === 'ok' ? 'status-ok' : 'status-error'}>
                    {selectedComment.result === 'ok' ? 'Исправлено' : 'Не исправлено'}
                  </span>
                </p>
              </div>
              
              <div className="comment-content">
                <h4>Выполненные работы:</h4>
                <p>{selectedComment.comment}</p>
              </div>
            </div>
            
            <div className="modal-footer">
              <button className="action-btn" onClick={() => setShowCommentModal(false)}>
                Закрыть
              </button>
            </div>
        </ModalShell>
      )}
    </div>
  );
}

// =====================================================
// КОМПОНЕНТ ПРОБЛЕМНЫХ ВЛ (2+ НЕУДАЧНЫХ ПРОВЕРКИ)
// =====================================================

// Меню «Превышение Pном» — список случаев перегруза секций (этап 3, блок Б).
function PowerOverload({ selectedRes }) {
  const { user } = useContext(AuthContext);
  const [cases, setCases] = useState([]);
  const [loading, setLoading] = useState(true);
  const [tab, setTab] = useState(''); // id этапа-вкладки; по умолчанию — первая доступная
  const [detailsCase, setDetailsCase] = useState(null);
  const [selectedIds, setSelectedIds] = useState([]);   // выбор кейсов для удаления (админ)
  const [delIds, setDelIds] = useState(null);            // id для подтверждения удаления
  const [deleting, setDeleting] = useState(false);
  const [actionModal, setActionModal] = useState(null); // { c, mode: 'askue'|'res' }
  const [comment, setComment] = useState('');
  const [commentError, setCommentError] = useState(false);
  const [files, setFiles] = useState([]);
  const [submitting, setSubmitting] = useState(false);
  // Числа этапа АСКУЭ (строки для контролируемого ввода)
  const [askueForm, setAskueForm] = useState({ consumers: '', limitedKw: '', unpolledKw: '' });
  const [askueError, setAskueError] = useState('');
  // Просмотрщик вложений (без навигации)
  const [showFileViewer, setShowFileViewer] = useState(false);
  const [viewerFiles, setViewerFiles] = useState([]);
  const [viewerIndex, setViewerIndex] = useState(0);
  const openViewer = (attachments, i) => { setViewerFiles(attachments); setViewerIndex(i); setShowFileViewer(true); };

  const f1 = (v) => (v == null ? '—' : Number(v).toFixed(1));
  const STAGE_RU = {
    askue_limit: 'Ограничение по АСКУЭ',
    res_work: 'Мероприятия РЭС',
    awaiting_recheck: 'Ожидает перепроверки',
    completed: 'Завершён'
  };

  // Состояние этапа кейса (idx: 0 ограничение, 1 мероприятия РЭС, 2 перепроверка).
  const stageOrder = { askue_limit: 0, res_work: 1, awaiting_recheck: 2, completed: 3 };
  const stageState = (c, idx) => {
    const cur = stageOrder[c.stage] ?? 0;
    return idx < cur ? 'done' : idx === cur ? 'current' : 'todo';
  };
  const STAGE_GLOW = { done: 'ico-glow-green', current: 'ico-glow-amber', todo: 'ico-dim' };
  const STAGE_WORD = { done: 'выполнено', current: 'ожидание', todo: 'ещё не начато' };
  const poStages = (c) => {
    const items = [
      { ico: <IconStageLimit size={34} />, name: 'Ввод ограничения по АСКУЭ' },
      { ico: <IconWrench size={30} />, name: 'Мероприятия РЭС' },
      { ico: <IconStageRecheck size={34} />, name: 'Перепроверка профилем' },
    ];
    return (
      <div className="po-stages">
        {items.map((it, i) => {
          const st = stageState(c, i);
          return (
            <span key={i} className={`po-stage-ico ${STAGE_GLOW[st]}`} title={`${it.name} — ${STAGE_WORD[st]}`}>
              {it.ico}
            </span>
          );
        })}
      </div>
    );
  };

  // showSpinner=false — тихий фон-рефреш (автообновление/возврат на вкладку), без мигания спиннера.
  const load = async (showSpinner = true) => {
    if (showSpinner) setLoading(true);
    try {
      const params = {};
      if (user.role === 'admin' && selectedRes) params.resId = selectedRes;
      const { data } = await api.get('/api/overload', { params });
      setCases(data);
    } catch (e) {
      console.error('overload load', e);
    } finally {
      if (showSpinner) setLoading(false);
    }
  };

  useEffect(() => {
    load();
    const h = () => load(false);
    window.addEventListener('notificationsUpdated', h);
    // Тихое автообновление раз в 60 сек (изменения из почтового приёмника — новые
    // перегрузы) — только когда вкладка активна; при возврате на вкладку — сразу.
    const interval = setInterval(() => { if (!document.hidden) load(false); }, 60000);
    const handleVisibility = () => { if (!document.hidden) load(false); };
    document.addEventListener('visibilitychange', handleVisibility);
    return () => {
      window.removeEventListener('notificationsUpdated', h);
      clearInterval(interval);
      document.removeEventListener('visibilitychange', handleVisibility);
    };
    // eslint-disable-next-line
  }, [selectedRes]);

  // Вкладки «своей работы» по роли/праву. Backend уже отдаёт только доступную область,
  // здесь — разбивка по этапам с бейджами-счётчиками.
  const isAdmin = user.role === 'admin' || user.isSuper;
  const askueRight = hasPerm(user, 'overload_askue');
  const isRes = user.role === 'res_responsible';
  const canDeleteCases = isAdmin && hasPerm(user, 'notifications_delete'); // удаление кейсов = право удаления уведомлений
  const cnt = (stg) => cases.filter(c => c.stage === stg).length;
  // Значок + цвет свечения/счётчика по вкладке-этапу.
  const TAB_META = {
    askue_limit:      { icon: <IconStageLimit size={22} />, glow: 'ico-glow-red', color: 'var(--red)' },
    res_work:         { icon: <IconWrench size={20} />, glow: 'ico-glow-amber', color: 'var(--amber)' },
    awaiting_recheck: { icon: <IconStageRecheck size={22} />, glow: 'ico-glow-blue', color: 'var(--blue, #2563eb)' },
    completed:        { icon: <IconCheck size={20} />, glow: 'ico-glow-green', color: 'var(--green)' },
  };
  const toggleSel = (id) => setSelectedIds(s => s.includes(id) ? s.filter(x => x !== id) : [...s, id]);
  const doDeleteCases = async () => {
    if (!delIds || !delIds.length) return;
    setDeleting(true);
    try {
      await api.post('/api/overload/delete-bulk', { ids: delIds });
      setSelectedIds(s => s.filter(id => !delIds.includes(id)));
      setDelIds(null);
      window.dispatchEvent(new CustomEvent('notificationsUpdated'));
      await load();
    } catch (e) {
      showToast(e.response?.data?.error || 'Ошибка удаления');
    } finally { setDeleting(false); }
  };
  const stageTabs = [];
  if (isAdmin) {
    stageTabs.push({ id: 'askue_limit', label: 'К ограничению' });
    stageTabs.push({ id: 'res_work', label: 'Мероприятия РЭС' });
    stageTabs.push({ id: 'awaiting_recheck', label: 'Ожидают перепроверки' });
    stageTabs.push({ id: 'completed', label: 'Завершённые' });
  } else {
    if (askueRight) stageTabs.push({ id: 'askue_limit', label: 'К ограничению' });
    if (isRes) stageTabs.push({ id: 'res_work', label: 'Мероприятия РЭС' });
    if (askueRight || isRes) stageTabs.push({ id: 'awaiting_recheck', label: 'Ожидают перепроверки' });
  }
  const curTab = (tab && stageTabs.some(t => t.id === tab)) ? tab : (stageTabs[0]?.id || 'askue_limit');
  const shown = cases.filter(c => c.stage === curTab);

  const openAskue = (c) => { setComment(''); setCommentError(false); setAskueForm({ consumers: '', limitedKw: '', unpolledKw: '' }); setAskueError(''); setActionModal({ c, mode: 'askue' }); };
  const openRes = (c) => { setComment(''); setCommentError(false); setFiles([]); setActionModal({ c, mode: 'res' }); };

  const submitAskue = async () => {
    const consumers = Math.round(Number(askueForm.consumers));
    const limitedKw = Math.round(Number(askueForm.limitedKw));
    const unpolledKw = Math.round(Number(askueForm.unpolledKw));
    if (!Number.isFinite(consumers) || consumers < 1 || consumers > 999) { setAskueError('Потребители: целое 1–999'); return; }
    if (!Number.isFinite(limitedKw) || limitedKw < 0 || limitedKw > 9999) { setAskueError('Введённые ограничения: целое 0–9999 кВт'); return; }
    if (!Number.isFinite(unpolledKw) || unpolledKw < 0 || unpolledKw > 9999) { setAskueError('Не введённые (нет опроса): целое 0–9999 кВт'); return; }
    setAskueError('');
    if (comment.trim().split(/\s+/).filter(w => w.length > 0).length < 5) { setCommentError(true); return; }
    if (!window.confirm('Подтвердите: ограничение по АСКУЭ выполнено?')) return;
    setSubmitting(true);
    try {
      await api.post(`/api/overload/${actionModal.c.id}/askue-complete`, { comment, consumers, limitedKw, unpolledKw });
      setActionModal(null);
      window.dispatchEvent(new CustomEvent('notificationsUpdated'));
      await load();
    } catch (e) {
      alert('Ошибка: ' + (e.response?.data?.error || e.message));
    } finally { setSubmitting(false); }
  };

  const submitRes = async () => {
    if (comment.trim().split(/\s+/).filter(w => w.length > 0).length < 5) { setCommentError(true); return; }
    setSubmitting(true);
    try {
      const fd = new FormData();
      fd.append('comment', comment);
      files.forEach(f => fd.append('attachments', f));
      await api.post(`/api/overload/${actionModal.c.id}/res-complete`, fd, {
        headers: { 'Content-Type': 'multipart/form-data' }
      });
      setActionModal(null);
      window.dispatchEvent(new CustomEvent('notificationsUpdated'));
      await load();
    } catch (e) {
      alert('Ошибка: ' + (e.response?.data?.error || e.message));
    } finally { setSubmitting(false); }
  };

  if (loading) return <LoadingSpinner message="Загрузка случаев перегруза..." />;

  const cardHeader = (c) => {
    const s = c.section || {};
    const ratioPct = c.ratio != null ? Math.round(c.ratio * 100) : null;
    return `${s.tpName} · СШ-${s.sectionNumber} · пик ${f1(c.peakKw)} кВт / лимит ${f1(c.limitKw)} кВт${ratioPct != null ? ` (${ratioPct}%)` : ''}${c.peakAt ? ` · ${c.peakAt}` : ''}`;
  };

  return (
    <div className="power-overload-page">
      <div className="section-header">
        <PageHeader icon={<IconZap size={24} />} title="Превышение Pном" />
      </div>

      <div className="po-tabs">
        {stageTabs.map(t => {
          const m = TAB_META[t.id] || {};
          return (
            <button key={t.id} className={`po-tab po-tab--rich ${curTab === t.id ? 'active' : ''}`} onClick={() => setTab(t.id)}>
              <span className={`po-tab-ico ${m.glow || ''}`}>{m.icon}</span>
              <span className="po-tab-label">{t.label}</span>
              <span className="po-tab-count" style={{ color: m.color }}>{cnt(t.id)}</span>
            </button>
          );
        })}
      </div>

      {canDeleteCases && selectedIds.length > 0 && (
        <div className="po-bulk-bar">
          <span>Выбрано: <b>{selectedIds.length}</b></span>
          <button className="delete-selected-btn" onClick={() => setDelIds(selectedIds)}>
            <IconTrash className="ico" /> Удалить выбранные ({selectedIds.length})
          </button>
        </div>
      )}

      {shown.length === 0 ? (
        <div className="no-issues"><IconCheck className="ico" style={{ color: 'var(--green)' }} /> Случаев нет</div>
      ) : (
        <div className="notifications-list">
          {shown.map(c => {
            const s = c.section || {};
            const ratioPct = c.ratio != null ? Math.round(c.ratio * 100) : null;
            const indClass = c.stage === 'completed' ? 'status-ok'
              : c.stage === 'awaiting_recheck' ? 'status-pending' : 'status-error';
            return (
              <div key={c.id} className="notification-compact power_overload"
                   onClick={() => setDetailsCase(c)} title="Открыть подробности">
                <div className="notification-narrow-content">
                  {canDeleteCases && (
                    <input type="checkbox" className="notification-checkbox-left"
                      checked={selectedIds.includes(c.id)}
                      onClick={(e) => e.stopPropagation()}
                      onChange={() => toggleSel(c.id)} />
                  )}
                  <div className="po-indicator">
                    <span className={`status-box ${indClass}`}></span>
                  </div>
                  <div className="notification-narrow-info">
                    <div className="notification-tp">{s.tpName} · СШ-{toRoman(s.sectionNumber)}
                      {c.cycles > 1 && <span className="po-cycle-badge">повтор {c.cycles}</span>}
                    </div>
                    <div className="notification-narrow-details">
                      <span className="label">РЭС:</span> {c.ResUnit?.name || '—'} |
                      <span className="label"> Пик:</span> {f1(c.peakKw)} кВт |
                      <span className="label"> Лимит:</span> {f1(c.limitKw)} кВт
                      {ratioPct != null ? ` (${ratioPct}%)` : ''}
                    </div>
                    <div className="notification-pu-number">
                      Этап: <strong>{STAGE_RU[c.stage]}</strong>
                    </div>
                    {c.stage === 'awaiting_recheck' && (
                      <div className="po-await-hint">
                        Ожидает загрузки профиля мощности{c.resCompletedAt ? ` после отчёта РЭС от ${new Date(c.resCompletedAt).toLocaleDateString('ru-RU')}` : ''}
                      </div>
                    )}
                  </div>
                  {poStages(c)}
                  <div className="notification-narrow-actions">
                    {c.stage === 'awaiting_recheck' && (
                      <span className="po-plaque">Ожидает перепроверки</span>
                    )}
                    {canDeleteCases && (
                      <button className="btn-icon danger po-del-btn" title="Удалить кейс"
                        onClick={(e) => { e.stopPropagation(); setDelIds([c.id]); }}>
                        <IconTrash className="ico ico-glow-red" />
                      </button>
                    )}
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {/* Подтверждение удаления кейса(ов) перегруза */}
      {delIds && (
        <ModalShell variant="confirm" title="Удаление кейса перегруза" className="delete-modal"
          onClose={() => !deleting && setDelIds(null)}>
          <div className="modal-body">
            <p>Удалить {delIds.length > 1 ? `выбранные кейсы (${delIds.length})` : 'этот кейс'} перегруза? История этапов будет потеряна безвозвратно.</p>
          </div>
          <div className="modal-footer">
            <button className="cancel-btn" onClick={() => setDelIds(null)} disabled={deleting}>Отмена</button>
            <button className="danger-btn" onClick={doDeleteCases} disabled={deleting}>
              {deleting ? 'Удаление...' : 'Удалить'}
            </button>
          </div>
        </ModalShell>
      )}

      {/* Детали кейса — крупная содержательная модалка */}
      {detailsCase && (() => {
        const c = detailsCase; const s = c.section || {};
        const ratioPct = c.ratio != null ? Math.round(c.ratio * 100) : null;
        const status = c.stage === 'completed'
          ? { t: c.recheckResult === 'ok' ? 'Устранён' : 'Завершён', cls: 'status-ok' }
          : c.stage === 'awaiting_recheck'
          ? { t: 'Ожидает перепроверки', cls: 'status-pending' }
          : { t: STAGE_RU[c.stage], cls: 'status-error' };
        return (
          <ModalShell
            title={<>{s.tpName} · СШ-{toRoman(s.sectionNumber)}</>}
            dockTitle={`${s.tpName} СШ-${toRoman(s.sectionNumber)}`}
            titleExtra={<span className={`tech-pill ${status.cls}`}>{status.t}</span>}
            className="po-details-modal" icon={<IconZap size={16} />}
            onClose={() => setDetailsCase(null)}
          >
              <div className="modal-body">
                <div className="tech-pmax">
                  <div className="tech-pmax-value">Пик {f1(c.peakKw)} кВт
                    {ratioPct != null && <span className={`tech-pmax-pct ${ratioPct > 100 ? 'red' : ratioPct >= 90 ? 'amber' : 'green'}`}>{ratioPct}%</span>}
                  </div>
                  <div className="tech-pmax-sub">
                    лимит {f1(c.limitKw)} кВт{ratioPct != null ? ` · ${ratioPct}% от лимита` : ''}
                    {c.peakAt ? ` · ${c.peakAt}` : ''}
                  </div>
                </div>

                <div className="tech-grid">
                  <div><span className="k">Этап</span><span className="v">{STAGE_RU[c.stage]}{c.cycles > 1 ? ` · повтор ${c.cycles}` : ''}</span></div>
                  <div><span className="k">РЭС</span><span className="v">{c.ResUnit?.name || '—'}</span></div>
                  <div><span className="k">Sном, кВА</span><span className="v">{c.tnKva != null ? c.tnKva : '—'}</span></div>
                  <div><span className="k">cosφ</span><span className="v">{c.cosPhi ?? '—'}</span></div>
                  {c.period && <div><span className="k">Период выгрузки</span><span className="v">{c.period}</span></div>}
                </div>

                <div className="po-timeline">
                  <h4>Хронология</h4>
                  {c.askueCompletedAt ? (
                    <div className="po-step done">
                      <strong>Ограничение по АСКУЭ</strong> — <span className="po-done-word">выполнено</span> · {c.askueUser?.fio || '—'}, {new Date(c.askueCompletedAt).toLocaleString('ru-RU')}
                      {c.askueComment ? <div className="po-step-comment">{c.askueComment}</div> : null}
                    </div>
                  ) : <div className="po-step pending"><strong>Ограничение по АСКУЭ</strong> — <span className="po-wait">ожидается</span></div>}

                  {c.resCompletedAt ? (
                    <div className="po-step done">
                      <strong>Мероприятия РЭС</strong> — <span className="po-done-word">выполнено</span> · {c.resUser?.fio || '—'}, {new Date(c.resCompletedAt).toLocaleString('ru-RU')}
                      {c.resComment ? <div className="po-step-comment">{c.resComment}</div> : null}
                      {Array.isArray(c.attachments) && c.attachments.length > 0 && (
                        <div className="po-attachments">
                          {c.attachments.map((a, i) => (
                            <button key={i} type="button" className="po-attach-link" onClick={() => openViewer(c.attachments, i)}>Файл {i + 1}</button>
                          ))}
                        </div>
                      )}
                    </div>
                  ) : <div className="po-step pending"><strong>Мероприятия РЭС</strong> — <span className="po-wait">ожидается</span></div>}

                  {c.recheckAt ? (
                    <div className={`po-step ${c.recheckResult === 'ok' ? 'done' : 'fail'}`}>
                      <strong>Перепроверка профилем</strong> — {new Date(c.recheckAt).toLocaleString('ru-RU')}: {c.recheckResult === 'ok' ? <span className="po-done-word">перегруз устранён</span> : <span className="po-fail-word">повторный перегруз</span>}
                      {c.recheckPeakKw != null ? ` (пик ${f1(c.recheckPeakKw)} кВт)` : ''}
                    </div>
                  ) : <div className="po-step pending"><strong>Перепроверка профилем</strong> — <span className="po-wait">ожидается</span></div>}

                  {c.closedAt && <div className="po-step done"><strong>Случай закрыт</strong> — <span className="po-done-word">завершён</span> · {new Date(c.closedAt).toLocaleString('ru-RU')}</div>}
                </div>
              </div>
              <div className="modal-footer">
                {hasPerm(user, 'overload_askue') && c.stage === 'askue_limit' && (
                  <button className="confirm-btn" onClick={() => openAskue(c)}>Ограничение по АСКУЭ выполнено</button>
                )}
                {(user.role === 'res_responsible' || isAdmin) && c.stage === 'res_work' && (
                  <button className="confirm-btn" onClick={() => openRes(c)}>Мероприятия выполнены</button>
                )}
                <button className="action-btn" onClick={() => setDetailsCase(null)}>Закрыть</button>
              </div>
          </ModalShell>
        );
      })()}

      {/* Действие: АСКУЭ (комментарий) / РЭС (комментарий + фото) */}
      {actionModal && (() => {
        const s = actionModal.c.section || {};
        const isRes = actionModal.mode === 'res';
        return (
          <ModalShell
            title={isRes ? 'Мероприятия РЭС выполнены' : 'Ограничение по АСКУЭ выполнено'}
            icon={<IconWrench size={16} />} onClose={() => setActionModal(null)}
          >
              <div className="modal-body">
                <div className="modal-info">
                  <p><strong>ТП:</strong> {s.tpName} · СШ-{toRoman(s.sectionNumber)}</p>
                  <p><strong>Пик:</strong> {f1(actionModal.c.peakKw)} кВт · лимит {f1(actionModal.c.limitKw)} кВт</p>
                </div>
                {!isRes && (() => {
                  const lim = Math.round(Number(askueForm.limitedKw)) || 0;
                  const unp = Math.round(Number(askueForm.unpolledKw)) || 0;
                  const total = (askueForm.limitedKw === '' && askueForm.unpolledKw === '') ? '' : lim + unp;
                  const numInput = (key, max) => (
                    <input type="number" min="0" max={max} step="1" value={askueForm[key]}
                      onChange={(e) => { setAskueForm(f => ({ ...f, [key]: e.target.value })); if (askueError) setAskueError(''); }} />
                  );
                  return (
                    <>
                      <div className="askue-grid">
                        <div className="form-group">
                          <label>Потребители (шт.)</label>
                          {numInput('consumers', 999)}
                        </div>
                        <div className="form-group">
                          <label>Опрашиваемые, P (кВт)</label>
                          {numInput('limitedKw', 9999)}
                        </div>
                        <div className="form-group">
                          <label>Не опрашиваемые (кВт)</label>
                          {numInput('unpolledKw', 9999)}
                        </div>
                        <div className="form-group">
                          <label>Общая разрешённая мощность ТП (кВт)</label>
                          <input type="text" value={total} readOnly className="askue-total" title="Сумма: введённые + не введённые" />
                        </div>
                      </div>
                      {askueError && <p className="askue-error" style={{ color: '#dc2626', fontWeight: 600, margin: '0 0 10px' }}>{askueError}</p>}
                    </>
                  );
                })()}
                <div className="form-group">
                  <label style={commentError ? { color: '#dc2626', fontWeight: 600 } : undefined}>
                    {isRes ? 'Что выполнено? (минимум 5 слов)' : 'Что именно ограничили? (минимум 5 слов)'}
                  </label>
                  <textarea rows={5} value={comment}
                    placeholder={isRes ? 'Опишите выполненные мероприятия…' : 'Например: введено ограничение мощности по АСКУЭ для абонентов…'}
                    onChange={(e) => { setComment(e.target.value); if (commentError && e.target.value.trim().split(/\s+/).filter(w => w.length > 0).length >= 5) setCommentError(false); }}
                    style={commentError ? { borderColor: '#dc2626', boxShadow: '0 0 0 3px rgba(220,38,38,0.2)', outline: 'none' } : undefined} />
                  <small className="muted">Слов: {comment.trim().split(/\s+/).filter(w => w.length > 0).length} из 5</small>
                </div>
                {isRes && (
                  <div className="form-group">
                    <label>Фото/документы (макс. 5)</label>
                    <input type="file" multiple accept="image/*,application/pdf"
                      onChange={(e) => setFiles(Array.from(e.target.files).slice(0, 5))} />
                    {files.length > 0 && <p className="muted">Выбрано файлов: {files.length}</p>}
                  </div>
                )}
              </div>
              <div className="modal-footer">
                <button className="cancel-btn" onClick={() => setActionModal(null)}>Отмена</button>
                <button className="confirm-btn" disabled={submitting}
                  onClick={isRes ? submitRes : submitAskue}>
                  {submitting ? 'Отправка...' : 'Подтвердить'}
                </button>
              </div>
          </ModalShell>
        );
      })()}

      {showFileViewer && (
        <FileViewer
          files={viewerFiles}
          currentIndex={viewerIndex}
          onClose={() => setShowFileViewer(false)}
          onNext={() => setViewerIndex((p) => (p + 1) % viewerFiles.length)}
          onPrev={() => setViewerIndex((p) => (p - 1 + viewerFiles.length) % viewerFiles.length)}
        />
      )}
    </div>
  );
}

// Ярлык/цвет этапа кейса перегруза (для «Проблемных ТП»).
const OVERLOAD_STAGE = {
  askue_limit:      { label: 'Ожидает ограничения АСКУЭ', cls: 'stage-red' },
  res_work:         { label: 'Мероприятия РЭС',            cls: 'stage-amber' },
  awaiting_recheck: { label: 'Ожидает перепроверки',       cls: 'stage-blue' },
  completed:        { label: 'Завершён',                   cls: 'stage-green' },
};
const f1 = (v) => (v == null ? '—' : Number(v).toFixed(1));

// Список карточек «Проблемные ТП» (стиль как у «Проблемных ВЛ»).
function ProblemTpList({ tps, onOpen }) {
  if (!tps.length) {
    return (
      <div className="no-data">
        <p><span className="svg-frame" style={{ marginRight: 8 }}><IconCheck size={26} /></span>
          Проблемных ТП нет</p>
      </div>
    );
  }
  return (
    <div className="problem-list">
      {tps.map(tp => {
        const st = OVERLOAD_STAGE[tp.stage] || { label: tp.stage, cls: '' };
        return (
          <div key={tp.id} className="problem-card" onClick={() => onOpen(tp)} title="Открыть историю">
            <div className="problem-header">
              <div className="problem-title">
                <h3>{tp.tpName} · СШ-{toRoman(tp.sectionNumber)}</h3>
                <span className="res-badge">{tp.resName}</span>
              </div>
              <span className="failure-badge critical">
                <IconX className="ico" style={{ color: 'var(--red)' }} /> {tp.cycles} циклов
              </span>
            </div>
            <div className="problem-meta">
              <span className={`po-stage-badge ${st.cls}`}>{st.label}</span>
              <span><b>Пик:</b> {f1(tp.peakKw)} кВт / лимит {f1(tp.limitKw)} кВт</span>
              {tp.recheckPeakKw != null && <span><b>После мероприятий:</b> {f1(tp.recheckPeakKw)} кВт</span>}
              {tp.techPuNumber && <span><b>Тех.учёт:</b> {tp.techPuNumber}</span>}
            </div>
          </div>
        );
      })}
    </div>
  );
}

// Модалка истории проблемной ТП: этапы, пики, комментарии АСКУЭ/РЭС, фото.
function ProblemTpModal({ tp, onClose, onViewFiles }) {
  const st = OVERLOAD_STAGE[tp.stage] || { label: tp.stage, cls: '' };
  const dt = (v) => v ? new Date(v).toLocaleString('ru-RU') : '—';
  return (
    <ModalShell title={`${tp.tpName} · СШ-${toRoman(tp.sectionNumber)} — история`}
      icon={<IconAlertTriangle size={16} />} className="details-modal" onClose={onClose}>
      <div className="modal-body">
        <div className="detail-section">
          <div className="detail-row"><strong>РЭС:</strong> {tp.resName}</div>
          <div className="detail-row"><strong>Циклов мероприятий:</strong> <span className="failure-count">{tp.cycles}</span></div>
          <div className="detail-row"><strong>Текущий этап:</strong> <span className={`po-stage-badge ${st.cls}`}>{st.label}</span></div>
          <div className="detail-row"><strong>Пик:</strong> {f1(tp.peakKw)} кВт · лимит {f1(tp.limitKw)} кВт{tp.ratio != null ? ` (${Math.round(tp.ratio * 100)}%)` : ''}{tp.peakAt ? ` · ${tp.peakAt}` : ''}</div>
          {tp.recheckPeakKw != null && (
            <div className="detail-row"><strong>Пик после мероприятий:</strong> {f1(tp.recheckPeakKw)} кВт{tp.recheckAt ? ` · ${dt(tp.recheckAt)}` : ''} — <span className="po-fail-word">повторный перегруз</span></div>
          )}
        </div>

        <div className="detail-section">
          <h5>Этап АСКУЭ (ограничение):</h5>
          {tp.askueCompletedAt ? (
            <>
              <div className="detail-row"><strong>Завершён:</strong> {dt(tp.askueCompletedAt)}{tp.askueUser ? ` · ${tp.askueUser}` : ''}</div>
              {tp.askueComment && <div className="comment-box"><p>{tp.askueComment}</p></div>}
            </>
          ) : <div className="detail-row po-wait">ожидается</div>}
        </div>

        <div className="detail-section">
          <h5>Этап РЭС (мероприятия):</h5>
          {tp.resCompletedAt ? (
            <>
              <div className="detail-row"><strong>Завершён:</strong> {dt(tp.resCompletedAt)}{tp.resUser ? ` · ${tp.resUser}` : ''}</div>
              {tp.resComment && <div className="comment-box"><p>{tp.resComment}</p></div>}
              {Array.isArray(tp.attachments) && tp.attachments.length > 0 && (
                <button className="pm-btn pm-btn--refresh" style={{ marginTop: 8 }}
                  onClick={() => onViewFiles(tp.attachments, 0)}>
                  <IconPaperclip className="ico" /> Фото/документы ({tp.attachments.length})
                </button>
              )}
            </>
          ) : <div className="detail-row po-wait">ожидается</div>}
        </div>
      </div>
      <div className="modal-footer">
        <button className="action-btn" onClick={onClose}>Закрыть</button>
      </div>
    </ModalShell>
  );
}

function ProblemVL({ selectedRes }) {
  const [problemVLs, setProblemVLs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showDeleteModal, setShowDeleteModal] = useState(false);
  const [selectedProblem, setSelectedProblem] = useState(null);
  const [deletePassword, setDeletePassword] = useState('');
  const [showDetailsModal, setShowDetailsModal] = useState(false);
  const [detailsProblem, setDetailsProblem] = useState(null);
  const [showEmailModal, setShowEmailModal] = useState(false);
  const [emailProblem, setEmailProblem] = useState(null);
  const [tab, setTab] = useState('vl');            // 'vl' | 'tp'
  const [overloadTps, setOverloadTps] = useState([]);
  const [tpDetails, setTpDetails] = useState(null); // раскрытая карточка проблемной ТП
  const [tpViewer, setTpViewer] = useState(null);   // { files, index } — просмотр фото РЭС

  useEffect(() => {
    loadProblemVLs();
    loadOverloadTps();

    const handleUpdate = () => { loadProblemVLs(); loadOverloadTps(); };
    window.addEventListener('problemVLUpdated', handleUpdate);
    window.addEventListener('notificationsUpdated', handleUpdate);

    return () => {
      window.removeEventListener('problemVLUpdated', handleUpdate);
      window.removeEventListener('notificationsUpdated', handleUpdate);
    };
  }, [selectedRes]);

  const loadOverloadTps = async () => {
    try {
      const { data } = await api.get('/api/problem-vl/overload-tp', {
        params: selectedRes ? { resId: selectedRes } : {}
      });
      setOverloadTps(data);
    } catch (error) {
      console.error('Error loading problem TPs:', error);
    }
  };

const handleSendEmail = async () => {
    try {
      await api.post(`/api/problem-vl/${emailProblem.id}/send-email`);
      alert('Письмо отправлено ответственному РЭС');
      setShowEmailModal(false);
      setEmailProblem(null);
    } catch (error) {
      alert('Ошибка отправки письма: ' + error.response?.data?.error || error.message);
    }
  };
  
  const loadProblemVLs = async () => {
    try {
      const response = await api.get('/api/problem-vl/list', {
        params: selectedRes ? { resId: selectedRes } : {}
      });
      setProblemVLs(response.data);
    } catch (error) {
      console.error('Error loading problem VLs:', error);
    } finally {
      setLoading(false);
    }
  };

  const handleDismiss = async () => {
    try {
      await api.put(`/api/problem-vl/${selectedProblem.id}/dismiss`, {
        password: deletePassword
      });
      
      alert('Проблема отклонена');
      setShowDeleteModal(false);
      setDeletePassword('');
      setSelectedProblem(null);
      loadProblemVLs();
    } catch (error) {
      alert('Ошибка: ' + (error.response?.data?.error || error.message));
    }
  };

  if (loading) return <LoadingSpinner type="pulse" message="Загрузка проблемных ВЛ..." submessage="Анализируем критические проблемы" />;
  return (
    <div className="problem-vl-container">
      <PageHeader icon={<IconAlertTriangle size={24} />} title="Проблемные ВЛ" />

      <div className="po-tabs">
        <button className={`po-tab ${tab === 'vl' ? 'active' : ''}`} onClick={() => setTab('vl')}>
          Проблемные ВЛ ({problemVLs.filter(p => p.status === 'active').length})
        </button>
        <button className={`po-tab ${tab === 'tp' ? 'active' : ''}`} onClick={() => setTab('tp')}>
          Проблемные ТП (перегруз) ({overloadTps.length})
        </button>
      </div>

      {tab === 'tp' ? (
        <ProblemTpList tps={overloadTps} onOpen={setTpDetails} />
      ) : (<>
      <div className="problem-info-row">
        <div className="problem-info-text">
          <p>В этом разделе отображаются ВЛ, которые не прошли проверку 2 и более раз после выполнения мероприятий РЭС.</p>
          <p>Это требует особого внимания и возможного выезда на место.</p>
        </div>
        <div className="problem-active-counter">
          <div className="pac-value">{problemVLs.filter(p => p.status === 'active').length}</div>
          <div className="pac-label">Активных проблем</div>
        </div>
      </div>

      {problemVLs.length === 0 ? (
  <div className="no-data">
    <p>
      <span className="svg-frame" style={{marginRight: 8}}><IconCheck size={26} /></span>
      Проблемных ВЛ нет
    </p>
  </div>
      ) : (
        <div className="problem-list">
          {problemVLs.map(problem => (
            <div
              key={problem.id}
              className="problem-card"
              onClick={() => { setDetailsProblem(problem); setShowDetailsModal(true); }}
              title="Открыть подробности"
            >
              <div className="problem-header">
                <div className="problem-title">
                  <h3>{problem.tpName} — {problem.vlName}</h3>
                  <span className="res-badge">{problem.ResUnit?.name}</span>
                </div>
                <span className="failure-badge critical">
                  <IconX className="ico" style={{color:'var(--red)'}} /> {problem.failureCount} неудачных
                </span>
              </div>

              <div className="problem-meta">
                <span><b>ПУ №:</b> {problem.puNumber}</span>
                <span><b>Позиция:</b> {problem.position === 'start' ? 'Начало' : problem.position === 'middle' ? 'Середина' : 'Конец'}</span>
                <span><b>Первое обращение:</b> {new Date(problem.firstReportDate).toLocaleDateString('ru-RU')}</span>
                <span><b>Последняя проверка:</b> {new Date(problem.lastErrorDate).toLocaleDateString('ru-RU')}</span>
              </div>

              <div className="problem-error-line"><b>Ошибка:</b> {problem.lastErrorDetails}</div>
              {problem.resComment && (
                <div className="problem-error-line"><b>Комментарий РЭС:</b> {problem.resComment}</div>
              )}
            </div>
          ))}
        </div>
      )}
      </>)}

      {/* Подробности проблемной ТП (перегруз) */}
      {tpDetails && (
        <ProblemTpModal tp={tpDetails} onClose={() => setTpDetails(null)}
          onViewFiles={(files, index) => setTpViewer({ files, index })} />
      )}
      {tpViewer && (
        <FileViewer
          files={tpViewer.files}
          currentIndex={tpViewer.index}
          onClose={() => setTpViewer(null)}
          onNext={() => setTpViewer(v => ({ ...v, index: (v.index + 1) % v.files.length }))}
          onPrev={() => setTpViewer(v => ({ ...v, index: (v.index - 1 + v.files.length) % v.files.length }))}
        />
      )}

      {/* Модальное окно подтверждения отклонения */}
      {showDeleteModal && (
        <ModalShell variant="confirm" title="Рассмотреть без объяснительной" className="delete-modal"
          onClose={() => setShowDeleteModal(false)}>
            <div className="modal-body">
              <p>Вы уверены, что хотите закрыть эту проблему без объяснительной записки?</p>
              <div className="problem-summary">
                <p><strong>{selectedProblem?.tpName} - {selectedProblem?.vlName}</strong></p>
                <p>ПУ №{selectedProblem?.puNumber} ({selectedProblem?.failureCount} ошибок)</p>
              </div>
              <p className="warning"><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /> Проблема будет закрыта без дальнейших действий!</p>
              <div className="form-group">
                <label>Введите пароль администратора:</label>
                <input
                  type="password"
                  value={deletePassword}
                  onChange={(e) => setDeletePassword(e.target.value)}
                  placeholder="Пароль"
                  autoFocus
                />
              </div>
            </div>
            <div className="modal-footer">
              <button className="cancel-btn" onClick={() => setShowDeleteModal(false)}>
                Отмена
              </button>
              <button 
                className="danger-btn" 
                onClick={handleDismiss}
                disabled={!deletePassword}
              >
                Закрыть без объяснительной
              </button>
            </div>
        </ModalShell>
      )}

      {/* Модальное окно с подробностями */}
      {showDetailsModal && detailsProblem && (
        <ModalShell title="Подробная информация о проблемной ВЛ" className="details-modal"
          icon={<IconAlertTriangle size={16} />} onClose={() => setShowDetailsModal(false)}>
            <div className="modal-body">
              <h4>{detailsProblem.tpName} - {detailsProblem.vlName}</h4>
              
              <div className="detail-section">
                <h5>Общая информация:</h5>
                <div className="detail-row">
                  <strong>РЭС:</strong> {detailsProblem.ResUnit?.name}
                </div>
                <div className="detail-row">
                  <strong>ПУ №:</strong> {detailsProblem.puNumber}
                </div>
                <div className="detail-row">
                  <strong>Позиция:</strong> {
                    detailsProblem.position === 'start' ? 'Начало' :
                    detailsProblem.position === 'middle' ? 'Середина' : 'Конец'
                  }
                </div>
              </div>
              
              <div className="detail-section">
                <h5>История проблемы:</h5>
                <div className="detail-row">
                  <strong>Первое обращение:</strong> {new Date(detailsProblem.firstReportDate).toLocaleString('ru-RU')}
                </div>
                <div className="detail-row">
                  <strong>Последняя проверка:</strong> {new Date(detailsProblem.lastErrorDate).toLocaleString('ru-RU')}
                </div>
                <div className="detail-row">
                  <strong>Количество неудачных проверок:</strong> <span className="failure-count">{detailsProblem.failureCount}</span>
                </div>
              </div>
              
              <div className="error-details-box">
                <strong>Последняя ошибка:</strong>
                <p>{detailsProblem.lastErrorDetails}</p>
              </div>
              
              {detailsProblem.resComment && (
                <div className="comment-box">
                  <strong>Последний комментарий РЭС:</strong>
                  <p>{detailsProblem.resComment}</p>
                </div>
              )}
            </div>
            <div className="modal-footer">
              <button className="action-btn" onClick={() => setShowDetailsModal(false)}>
                Закрыть
              </button>
            </div>
        </ModalShell>
      )}

    {/* НОВОЕ: Модальное окно отправки письма */}
      {showEmailModal && emailProblem && (
        <ModalShell title="Направить письмо исполнителю" className="email-modal"
          icon={<IconMail size={16} />} onClose={() => setShowEmailModal(false)}>
            <div className="modal-body">
              <p>Будет отправлено уведомление ответственному РЭС с требованием предоставить объяснительную записку.</p>
              <div className="problem-summary">
                <p><strong>РЭС:</strong> {emailProblem.ResUnit?.name}</p>
                <p><strong>ТП:</strong> {emailProblem.tpName}</p>
                <p><strong>ВЛ:</strong> {emailProblem.vlName}</p>
                <p><strong>ПУ №:</strong> {emailProblem.puNumber}</p>
                <p><strong>Количество неудачных проверок:</strong> {emailProblem.failureCount}</p>
              </div>
            </div>
            <div className="modal-footer">
              <button className="cancel-btn" onClick={() => setShowEmailModal(false)}>
                Отмена
              </button>
              <button 
                className="primary-btn" 
                onClick={handleSendEmail}
              >
                <IconMail className="ico" /> Отправить уведомление
              </button>
            </div>
        </ModalShell>
      )}
    </div>
  );
}



// =====================================================
// КОМПОНЕНТ НАСТРОЕК С УПРАВЛЕНИЕМ ПОЛЬЗОВАТЕЛЯМИ
// =====================================================

function Settings() {
  const { user } = useContext(AuthContext);
  // Вкладки с опасными действиями видны только при соответствующем праве
  // (суперадмин видит все). «Пользователи» — список read-only, действия гейтятся внутри.
  const tabs = [
    { id: 'structure', label: 'Структура сети', show: hasPerm(user, 'structure_upload') || hasPerm(user, 'structure_edit') },
    { id: 'users', label: 'Пользователи', show: true },
    { id: 'diagnose', label: <><IconSearch className="ico" /> Диагностика данных</>, show: hasPerm(user, 'db_tools') },
    { id: 'maintenance', label: 'Обслуживание', show: hasPerm(user, 'structure_edit') },
    { id: 'files', label: 'Управление файлами', show: hasPerm(user, 'files_manage') },
    { id: 'database', label: 'База данных', show: hasPerm(user, 'db_tools') || hasPerm(user, 'history_purge') },
  ].filter(t => t.show);
  const [activeTab, setActiveTab] = useState(tabs[0]?.id || 'users');

  return (
    <div className="settings-container">
      <PageHeader icon={<IconSettings size={24} />} title="Настройки системы" />

      <div className="settings-tabs">
        {tabs.map(t => (
          <button key={t.id} className={activeTab === t.id ? 'active' : ''} onClick={() => setActiveTab(t.id)}>
            {t.label}
          </button>
        ))}
      </div>

      <div className="settings-content">
        {activeTab === 'structure' && <StructureSettings />}
        {activeTab === 'users' && <UserSettings />}
        {activeTab === 'diagnose' && <DiagnoseData />}  {/* НОВЫЙ */}
        {activeTab === 'maintenance' && <MaintenanceSettings />}
        {activeTab === 'files' && <FileManagement />}
        {activeTab === 'database' && <DatabaseMaintenance />}
      </div>
    </div>
  );
}
// Новый подкомпонент управления файлами
function FileManagement() {
  const [files, setFiles] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selectedFile, setSelectedFile] = useState(null);
  const [showDeleteModal, setShowDeleteModal] = useState(false);
  const [deletePassword, setDeletePassword] = useState('');
  const [showFileViewer, setShowFileViewer] = useState(false);
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [currentFileIndex, setCurrentFileIndex] = useState(0);
  
  // НОВОЕ: Фильтры
  const [searchTp, setSearchTp] = useState('');
  const [statusFilter, setStatusFilter] = useState('');
  const [resFilter, setResFilter] = useState('');
  const [dateFrom, setDateFrom] = useState('');
  const [dateTo, setDateTo] = useState('');
  const [fmShowTop, setFmShowTop] = useState(false);

  // Кнопка «наверх» для списка файлов (скролл у .content).
  useEffect(() => {
    let el = null, raf = null;
    const onScroll = () => { if (el) setFmShowTop(el.scrollTop > 300); };
    const attach = () => {
      el = document.querySelector('.content');
      if (el) { el.addEventListener('scroll', onScroll, { passive: true }); onScroll(); }
      else raf = requestAnimationFrame(attach);
    };
    attach();
    return () => { if (raf) cancelAnimationFrame(raf); if (el) el.removeEventListener('scroll', onScroll); };
  }, []);
  // Диагностика файла в Cloudinary
  const [diag, setDiag] = useState(null);        // { file, data|null }
  const [diagLoading, setDiagLoading] = useState(false);
  // Файлы, отдавшие 404 (утеряны в хранилище) — по public_id
  const [lostFiles, setLostFiles] = useState(new Set());
  const markLost = (id) => setLostFiles(prev => { const n = new Set(prev); n.add(id); return n; });
  // Множественный выбор + удаление выбранных
  const [checkedIds, setChecked] = useState(new Set());
  const [showBulkModal, setShowBulkModal] = useState(false);
  const [bulkPassword, setBulkPassword] = useState('');
  const [bulkDeleting, setBulkDeleting] = useState(false);
  const toggleCheck = (id) => setChecked(prev => {
    const n = new Set(prev);
    n.has(id) ? n.delete(id) : n.add(id);
    return n;
  });

  const handleBulkDelete = async () => {
    setBulkDeleting(true);
    const ids = [...checkedIds];
    let ok = 0, fail = 0, badPass = false;
    for (const id of ids) {
      try {
        await api.delete(`/api/admin/files/${encodeURIComponent(id)}`, { data: { password: bulkPassword } });
        ok++;
      } catch (e) {
        if (e.response?.status === 403) { badPass = true; break; }
        fail++;
      }
    }
    setBulkDeleting(false);
    if (badPass) { alert('Неверный пароль'); return; }
    setShowBulkModal(false);
    setBulkPassword('');
    setChecked(new Set());
    alert(`Удалено файлов: ${ok}${fail ? `, ошибок: ${fail}` : ''}`);
    loadFiles();
  };

  useEffect(() => {
    loadFiles();
  }, []);

  const runDiag = async (file) => {
    setDiag({ file, data: null }); setDiagLoading(true);
    try {
      const token = toBase64Url(file.public_id || file.filename);
      const { data } = await api.get(`/api/admin/files/diag/${token}`);
      setDiag({ file, data });
    } catch (e) {
      setDiag({ file, data: { error: e.response?.data?.error || e.message } });
    } finally { setDiagLoading(false); }
  };
  
  const loadFiles = async () => {
    try {
      console.log('Loading files...');
      const response = await api.get('/api/admin/files');
      console.log('Files response:', response.data);
      setFiles(response.data.files);
    } catch (error) {
      console.error('Error loading files:', error);
    } finally {
      setLoading(false);
    }
  };
  
  const handleDeleteFile = async () => {
    try {
      const publicId = selectedFile.public_id || selectedFile.filename;
      
      await api.delete(`/api/admin/files/${encodeURIComponent(publicId)}`, {
        data: { password: deletePassword }
      });
      
      alert('Файл удален успешно');
      setShowDeleteModal(false);
      setDeletePassword('');
      setSelectedFile(null);
      loadFiles();
      
    } catch (error) {
      console.error('Delete error:', error);
      alert('Ошибка удаления: ' + (error.response?.data?.error || error.message));
    }
  };
  
  const getTotalSize = () => {
    const totalBytes = files.reduce((sum, file) => sum + (file.size || 0), 0);
    return (totalBytes / 1024 / 1024).toFixed(2);
  };
  
  // НОВОЕ: Функция получения статуса
  const getStatusInfo = (status) => {
    switch(status) {
      case 'completed':
        return { text: 'Завершено', class: 'status-completed', icon: <IconCheck className="ico" style={{color:'var(--green)'}} /> };
      case 'awaiting_recheck':
        return { text: 'Ожидает перепроверки', class: 'status-awaiting', icon: <IconClock className="ico" /> };
      case 'awaiting_work':
        return { text: 'Ожидает мероприятий', class: 'status-work', icon: <IconWrench className="ico" /> };
      default:
        return { text: 'Неизвестно', class: 'status-unknown', icon: <IconHelp className="ico" /> };
    }
  };
  
  // НОВОЕ: Фильтрация файлов
  // Список РЭС для фильтра (из файлов).
  const resOptions = [...new Set(files.map(f => f.resName).filter(Boolean))].sort((a, b) => a.localeCompare(b, 'ru'));

  const filteredFiles = files.filter(file => {
    // Фильтр по ТП
    if (searchTp && !file.tpName?.toLowerCase().includes(searchTp.toLowerCase())) {
      return false;
    }
    // Фильтр по статусу
    if (statusFilter && file.status !== statusFilter) {
      return false;
    }
    // Фильтр по РЭС
    if (resFilter && file.resName !== resFilter) {
      return false;
    }
    // Фильтр по периоду загрузки (uploadDate)
    if (file.uploadDate) {
      const d = new Date(file.uploadDate);
      if (dateFrom && d < new Date(dateFrom + 'T00:00:00')) return false;
      if (dateTo && d > new Date(dateTo + 'T23:59:59')) return false;
    }
    return true;
  });

  // «Выбрать все показанные» — работает по текущему фильтру.
  const shownIds = filteredFiles.map(f => f.public_id).filter(Boolean);
  const allShownChecked = shownIds.length > 0 && shownIds.every(id => checkedIds.has(id));
  const toggleSelectAll = () => setChecked(prev => {
    const n = new Set(prev);
    if (allShownChecked) shownIds.forEach(id => n.delete(id));
    else shownIds.forEach(id => n.add(id));
    return n;
  });

  if (loading) return <div className="loading"><RossetiLoader /></div>;

  return (
    <div className="settings-section">
      <h3>Управление загруженными файлами</h3>
      
      {/* НОВОЕ: Фильтры */}
      <div className="file-filters">
        <div className="filter-group">
          <label><IconSearch className="ico" /> Поиск по ТП:</label>
          <input
            type="text"
            value={searchTp}
            onChange={(e) => setSearchTp(e.target.value)}
            placeholder="Введите название ТП..."
            className="search-input"
          />
        </div>
        
        <div className="filter-group">
          <label><IconChart className="ico" /> Статус:</label>
          <select
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
            className="status-filter"
          >
            <option value="">Все статусы</option>
            <option value="completed">Завершено</option>
            <option value="awaiting_recheck">Ожидает перепроверки</option>
            <option value="awaiting_work">Ожидает мероприятий</option>
          </select>
        </div>

        <div className="filter-group">
          <label><IconBuilding className="ico" /> РЭС:</label>
          <select value={resFilter} onChange={(e) => setResFilter(e.target.value)} className="status-filter">
            <option value="">Все РЭС</option>
            {resOptions.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
        </div>

        <div className="filter-group">
          <label><IconCalendar className="ico" /> Период:</label>
          <div style={{ display: 'inline-flex', gap: 6, alignItems: 'center' }}>
            <input type="date" value={dateFrom} onChange={(e) => setDateFrom(e.target.value)} className="search-input" />
            <span>—</span>
            <input type="date" value={dateTo} onChange={(e) => setDateTo(e.target.value)} className="search-input" />
          </div>
        </div>

        {(searchTp || statusFilter || resFilter || dateFrom || dateTo) && (
          <button
            className="clear-filters-btn"
            onClick={() => { setSearchTp(''); setStatusFilter(''); setResFilter(''); setDateFrom(''); setDateTo(''); }}
          >
            <IconX className="ico" /> Очистить фильтры
          </button>
        )}
      </div>
      
      <div className="file-stats">
        <div className="stat-card">
          <h4>Всего файлов</h4>
          <p className="stat-value">{files.length}</p>
        </div>
        <div className="stat-card">
          <h4>Показано</h4>
          <p className="stat-value">{filteredFiles.length}</p>
        </div>
        <div className="stat-card">
          <h4>Общий размер</h4>
          <p className="stat-value">{getTotalSize()} MB</p>
        </div>
      </div>

      {filteredFiles.length > 0 && (
        <div className="file-bulk-bar">
          <label className="file-selectall">
            <input type="checkbox" checked={allShownChecked} onChange={toggleSelectAll} />
            Выбрать все показанные
          </label>
          <span className="muted">Выбрано: {checkedIds.size}</span>
          <button className="delete-selected-btn" disabled={checkedIds.size === 0}
            onClick={() => setShowBulkModal(true)}>
            <IconTrash className="ico" /> Удалить выбранные ({checkedIds.size})
          </button>
        </div>
      )}

      {filteredFiles.length === 0 ? (
        <div className="no-data">
          <p>
            {searchTp || statusFilter 
              ? 'По вашим фильтрам ничего не найдено'
              : 'Нет загруженных файлов'}
          </p>
        </div>
      ) : (
        <div className="files-grid">
          {filteredFiles.map((file, idx) => {
            const statusInfo = getStatusInfo(file.status);
            
            return (
              <div key={idx} className={`file-card ${checkedIds.has(file.public_id) ? 'checked' : ''}`}>
                <input type="checkbox" className="file-check"
                  checked={checkedIds.has(file.public_id)}
                  onChange={() => toggleCheck(file.public_id)}
                  title="Выбрать файл" />
                {(file.url.toLowerCase().endsWith('.jpg') ||
                  file.url.toLowerCase().endsWith('.jpeg') || 
                  file.url.toLowerCase().endsWith('.png') || 
                  file.url.toLowerCase().endsWith('.gif')) ? (
                  <BlobImage file={file} alt={file.original_name} className="file-thumbnail" onLost={markLost} />
                ) : (
                  <div className="file-icon"><IconFileText className="ico" /></div>
                )}
                {lostFiles.has(file.public_id) && (
                  <div className="file-lost-badge" title="Ресурс не найден в Cloudinary">
                    <IconAlertTriangle className="ico" /> Файл утерян — удалите запись
                  </div>
                )}
                
                <div className="file-info">
                  <p className="file-name">{file.original_name}</p>
                  <p className="file-meta">
                    <strong>РЭС:</strong> {file.resName}<br/>
                    <strong>ТП:</strong> {file.tpName}<br/>
                    <strong>ВЛ:</strong> {file.vlName}<br/>
                    <strong>ПУ:</strong> {file.puNumber}<br/>
                    <strong>Дата:</strong> {new Date(file.uploadDate).toLocaleDateString('ru-RU')}
                  </p>
                  
                  {/* НОВОЕ: Статус документа */}
                  <div 
                    className={`file-status ${statusInfo.class}`}
                    title={file.resComment || 'Нет комментария'}
                  >
                    <span className="status-icon">{statusInfo.icon}</span>
                    <span className="status-text">{statusInfo.text}</span>
                  </div>
                </div>
                
                <div className="file-actions">
                  <button
                    onClick={() => { setSelectedFiles([file]); setCurrentFileIndex(0); setShowFileViewer(true); }}
                    className="btn-icon"
                    title="Открыть"
                  >
                    <IconEye className="ico ico-glow-blue" />
                  </button>
                  <button
                    onClick={() => runDiag(file)}
                    className="btn-icon"
                    title="Диагностика файла в хранилище"
                  >
                    <IconSearch className="ico ico-glow-blue" />
                  </button>
                  <button
                    onClick={() => {
                      setSelectedFile(file);
                      setShowDeleteModal(true);
                    }}
                    className="btn-icon danger"
                    title="Удалить"
                  >
                    <IconTrash className="ico ico-glow-red" />
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      )}

      {fmShowTop && (
        <button className="scroll-to-top" title="Наверх"
          onClick={() => { const c = document.querySelector('.content'); if (c) c.scrollTo({ top: 0, behavior: 'smooth' }); }}>
          <IconArrowUp className="ico" />
        </button>
      )}

      {/* Диагностика файла: матрица комбинаций Cloudinary */}
      {diag && (
        <ModalShell
          title="Диагностика файла"
          dockTitle={`Диагностика: ${diag.file.original_name}`}
          icon={<IconSearch size={16} />}
          onClose={() => setDiag(null)}
        >
            <div className="modal-body">
              <p className="file-name">{diag.file.original_name}</p>
              <p className="muted" style={{ fontSize: 12, wordBreak: 'break-all', marginBottom: 12 }}>{diag.file.public_id}</p>
              {diagLoading ? (
                <div className="loading"><RossetiLoader /></div>
              ) : diag.data?.error ? (
                <p style={{ color: 'var(--red)' }}>{diag.data.error}</p>
              ) : diag.data ? (
                <>
                  <p style={{ marginBottom: 10 }}>
                    {diag.data.found ? (
                      <span style={{ color: 'var(--green)', fontWeight: 700 }}>
                        Найден: {diag.data.found.resource_type}/{diag.data.found.type} ({diag.data.found.idVariant}) · {diag.data.found.bytes} байт · {diag.data.found.format || '—'}
                      </span>
                    ) : (
                      <span style={{ color: 'var(--red)', fontWeight: 700 }}>Файл не найден ни в одной комбинации — утерян в хранилище</span>
                    )}
                  </p>
                  <table className="diag-table">
                    <thead><tr><th>resource_type</th><th>type</th><th>public_id</th><th>Результат</th></tr></thead>
                    <tbody>
                      {diag.data.results.map((r, i) => (
                        <tr key={i}>
                          <td>{r.resource_type}</td><td>{r.type}</td><td>{r.idVariant}</td>
                          <td>{r.ok ? <span style={{ color: 'var(--green)', fontWeight: 700 }}>OK</span> : <span style={{ color: 'var(--red)' }}>{r.http || 404}</span>}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>

                  {diag.data.deliveryTests && diag.data.deliveryTests.length > 0 && (
                    <>
                      <p style={{ margin: '14px 0 6px', fontWeight: 600 }}>Доставка (delivery-URL) найденной комбинации:</p>
                      <table className="diag-table">
                        <thead><tr><th>Вариант URL</th><th>HTTP</th><th>x-cld-error</th></tr></thead>
                        <tbody>
                          {diag.data.deliveryTests.map((t, i) => (
                            <tr key={i}>
                              <td style={{ textAlign: 'left' }}>{t.label}</td>
                              <td>{t.status == null ? '—' : <span style={{ color: t.status >= 200 && t.status < 300 ? 'var(--green)' : 'var(--red)', fontWeight: 700 }}>{t.status}</span>}</td>
                              <td style={{ textAlign: 'left', color: t.xCldError ? 'var(--red)' : 'var(--text-muted)' }}>{t.xCldError || '—'}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                      <p className="muted" style={{ fontSize: 12, marginTop: 8 }}>
                        <IconInfo className="ico" style={{ color: 'var(--blue)' }} /> Cloudinary CDN кеширует ошибки доставки (40x) до ~24 ч (экспоненциально): если комбинация в первой таблице найдена (asset существует), но delivery-URL всё ещё отдаёт ошибку — это может быть закешированный отказ, который сам истечёт. Рабочий URL из `resource.secure_url` при этом уже отдаёт файл.
                      </p>
                    </>
                  )}
                </>
              ) : null}
            </div>
            <div className="modal-footer">
              <button className="action-btn" onClick={() => setDiag(null)}>Закрыть</button>
            </div>
        </ModalShell>
      )}

      {/* Массовое удаление выбранных файлов */}
      {showBulkModal && (
        <ModalShell variant="confirm" title="Удаление выбранных файлов" className="delete-modal"
          onClose={() => setShowBulkModal(false)}
          onBackdrop={() => !bulkDeleting && setShowBulkModal(false)}>
            <div className="modal-body">
              <p>Будет удалено файлов: <strong>{checkedIds.size}</strong>. Файлы удаляются из хранилища и отвязываются от историй проверок. Действие необратимо.</p>
              <input type="password" placeholder="Пароль удаления" value={bulkPassword}
                onChange={e => setBulkPassword(e.target.value)} className="purge-pass"
                style={{ marginTop: 10, width: '100%' }} />
            </div>
            <div className="modal-footer">
              <button className="action-btn" onClick={() => setShowBulkModal(false)} disabled={bulkDeleting}>Отмена</button>
              <button className="delete-selected-btn" onClick={handleBulkDelete} disabled={bulkDeleting || !bulkPassword}>
                {bulkDeleting ? 'Удаление...' : 'Удалить'}
              </button>
            </div>
        </ModalShell>
      )}

      {/* Модальное окно удаления */}
      {showDeleteModal && (
        <ModalShell variant="confirm" title="Подтверждение удаления файла" className="delete-modal"
          onClose={() => setShowDeleteModal(false)}>
            <div className="modal-body">
              <p>Вы собираетесь удалить файл:</p>
              <p><strong>{selectedFile?.original_name}</strong></p>
              <p className="warning"><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /> Это действие нельзя отменить!</p>
              <div className="form-group">
                <label>Введите пароль администратора:</label>
                <input
                  type="password"
                  value={deletePassword}
                  onChange={(e) => setDeletePassword(e.target.value)}
                  placeholder="Пароль"
                  autoFocus
                />
              </div>
            </div>
            <div className="modal-footer">
              <button className="cancel-btn" onClick={() => setShowDeleteModal(false)}>
                Отмена
              </button>
              <button 
                className="danger-btn" 
                onClick={handleDeleteFile}
                disabled={!deletePassword}
              >
                Удалить
              </button>
            </div>
        </ModalShell>
      )}

      {/* Просмотрщик файлов */}
      {showFileViewer && (
        <FileViewer 
          files={selectedFiles}
          currentIndex={currentFileIndex}
          onClose={() => setShowFileViewer(false)}
          onNext={() => setCurrentFileIndex((prev) => (prev + 1) % selectedFiles.length)}
          onPrev={() => setCurrentFileIndex((prev) => (prev - 1 + selectedFiles.length) % selectedFiles.length)}
        />
      )}
    </div>
  );
}

function DiagnoseData() {
  const [selectedRes, setSelectedRes] = useState('');
  const [resList, setResList] = useState([]);
  const [diagData, setDiagData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [showFixModal, setShowFixModal] = useState(false);
  const [fixingNotif, setFixingNotif] = useState(null);
  const [fixPassword, setFixPassword] = useState('');
  const [newResId, setNewResId] = useState('');
  const [activeView, setActiveView] = useState('mismatches');
  
  // НОВОЕ: для массового исправления
  const [showMassFixModal, setShowMassFixModal] = useState(false);
  const [massFixPassword, setMassFixPassword] = useState('');
  const [massFixing, setMassFixing] = useState(false);
  
  useEffect(() => {
    loadResList();
  }, []);
  
  const loadResList = async () => {
    try {
      const response = await api.get('/api/res/list');
      setResList(response.data);
    } catch (error) {
      console.error('Error loading RES list:', error);
    }
  };
  
  const loadDiagnostics = async () => {
    if (!selectedRes) {
      alert('Выберите РЭС');
      return;
    }
    
    setLoading(true);
    try {
      const response = await api.get(`/api/admin/diagnose/${selectedRes}`);
      setDiagData(response.data);
    } catch (error) {
      alert('Ошибка диагностики: ' + error.message);
    } finally {
      setLoading(false);
    }
  };
  
  // НОВАЯ ФУНКЦИЯ: массовое исправление
  const handleMassAutoFix = async () => {
    setMassFixing(true);
    try {
      console.log('Mass auto-fixing all mismatches for RES:', selectedRes);
      
      const response = await api.post(`/api/admin/auto-fix-all/${selectedRes}`, {
        password: massFixPassword
      });
      
      console.log('Mass fix response:', response.data);
      
      alert(
        `${response.data.message}\n\n` +
        `Всего проверено: ${response.data.stats.total}\n` +
        `Исправлено: ${response.data.stats.fixed}\n` +
        `Уже корректных: ${response.data.stats.alreadyCorrect}`
      );
      
      setShowMassFixModal(false);
      setMassFixPassword('');
      
      // Перезагружаем диагностику
      await loadDiagnostics();
      
    } catch (error) {
      console.error('Mass fix error:', error);
      alert('Ошибка: ' + (error.response?.data?.error || error.message));
    } finally {
      setMassFixing(false);
    }
  };
  
  const handleManualFix = async () => {
    try {
      console.log('Manual fixing notification:', fixingNotif.notificationId);
      console.log('New resId:', newResId);
      
      const response = await api.put(`/api/admin/fix-notification/${fixingNotif.notificationId}`, {
        newResId: parseInt(newResId),
        password: fixPassword
      });
      
      console.log('Response:', response.data);
      
      alert(`${response.data.message}`);
      
      setShowFixModal(false);
      setFixPassword('');
      setNewResId('');
      setFixingNotif(null);
      
      await loadDiagnostics();
      
    } catch (error) {
      console.error('Manual fix error:', error);
      alert('Ошибка: ' + (error.response?.data?.error || error.message));
    }
  };
  
  return (
    <div className="diagnose-container">
      <div className="diagnose-header">
        <div className="diagnose-title">
          <div className="diagnose-icon"><IconSearch className="ico" /></div>
          <div>
            <h3>Диагностика и исправление данных</h3>
            <p>Проверка соответствия resId в уведомлениях и структуре сети</p>
          </div>
        </div>
      </div>
      
      <div className="diagnose-controls">
        <div className="control-group">
          <label>Выберите РЭС для диагностики:</label>
          <select 
            value={selectedRes}
            onChange={(e) => setSelectedRes(e.target.value)}
            className="diagnose-select"
          >
            <option value="">-- Выберите РЭС --</option>
            {resList.map(res => (
              <option key={res.id} value={res.id}>{res.name}</option>
            ))}
          </select>
        </div>
        
        <button 
          onClick={loadDiagnostics}
          disabled={!selectedRes || loading}
          className="btn-diagnose"
        >
          {loading ? (
            <>
                            Диагностика...
            </>
          ) : (
            <>
              <span><IconSearch className="ico" /></span>
              Запустить диагностику
            </>
          )}
        </button>
        
        {/* НОВАЯ КНОПКА: Массовое исправление */}
        {diagData && diagData.stats.mismatches > 0 && (
          <button 
            onClick={() => setShowMassFixModal(true)}
            className="btn-mass-fix"
            disabled={loading}
          >
            <span><IconWrench className="ico" /></span>
            Исправить ВСЕ ({diagData.stats.mismatches})
          </button>
        )}
      </div>
      
      {diagData && (
        <>
          {/* Статистика */}
          <div className="diagnose-stats">
            <div className="stat-card">
              <div className="stat-icon"><IconChart className="ico" /></div>
              <div className="stat-content">
                <h4>Структуры</h4>
                <p className="stat-value">{diagData.stats.totalStructures}</p>
              </div>
            </div>
            
            <div className="stat-card">
              <div className="stat-icon"><IconBell className="ico" /></div>
              <div className="stat-content">
                <h4>Уведомления</h4>
                <p className="stat-value">{diagData.stats.totalNotifications}</p>
              </div>
            </div>
            
            <div className={`stat-card ${diagData.stats.mismatches > 0 ? 'error' : 'success'}`}>
              <div className="stat-icon">{diagData.stats.mismatches > 0 ? <IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /> : <IconCheck className="ico" style={{color:'var(--green)'}} />}</div>
              <div className="stat-content">
                <h4>Несоответствия</h4>
                <p className="stat-value">{diagData.stats.mismatches}</p>
              </div>
            </div>
          </div>
          
          {/* Вкладки просмотра */}
          <div className="diagnose-tabs">
            <button 
              className={`diagnose-tab ${activeView === 'mismatches' ? 'active' : ''}`}
              onClick={() => setActiveView('mismatches')}
            >
              <span className="tab-icon"><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /></span>
              Несоответствия
              {diagData.stats.mismatches > 0 && (
                <span className="tab-badge">{diagData.stats.mismatches}</span>
              )}
            </button>
            <button 
              className={`diagnose-tab ${activeView === 'all' ? 'active' : ''}`}
              onClick={() => setActiveView('all')}
            >
              <span className="tab-icon"><IconClipboard className="ico" /></span>
              Все уведомления
              <span className="tab-badge">{diagData.stats.totalNotifications}</span>
            </button>
            <button 
              className={`diagnose-tab ${activeView === 'structures' ? 'active' : ''}`}
              onClick={() => setActiveView('structures')}
            >
              <span className="tab-icon"><IconLayers className="ico" /></span>
              Структура сети
              <span className="tab-badge">{diagData.stats.totalStructures}</span>
            </button>
          </div>
          
          {/* Контент */}
          <div className="diagnose-content">
            {/* НЕСООТВЕТСТВИЯ */}
            {activeView === 'mismatches' && (
              <div className="mismatches-view">
                {diagData.mismatches.length === 0 ? (
                  <div className="no-issues">
                    <div className="no-issues-icon"><IconCheck className="ico" style={{color:'var(--green)'}} /></div>
                    <h4>Несоответствий не найдено!</h4>
                    <p>Все resId в уведомлениях соответствуют структуре сети</p>
                  </div>
                ) : (
                  <div className="mismatches-list">
                    {diagData.mismatches.map((mismatch, idx) => (
                      <div key={idx} className="mismatch-card">
                        <div className="mismatch-header">
                          <div className="mismatch-type">
                            <span className={`type-badge ${mismatch.type}`}>{mismatch.type}</span>
                            <span className="date-badge">
                              {new Date(mismatch.createdAt).toLocaleDateString('ru-RU')}
                            </span>
                          </div>
                          <div className="mismatch-location">
                            <strong>{mismatch.tpName}</strong> - {mismatch.vlName}
                          </div>
                        </div>
                        
                        <div className="mismatch-comparison">
                          <div className="comparison-item wrong">
                            <div className="comparison-label"><IconX className="ico" style={{color:'var(--red)'}} /> ResId в уведомлении:</div>
                            <div className="comparison-value">
                              {mismatch.notifResId} ({mismatch.notifResName})
                            </div>
                          </div>
                          
                          <div className="comparison-arrow"><IconArrowRight className="ico" /></div>
                          
                          <div className="comparison-item correct">
                            <div className="comparison-label"><IconCheck className="ico" style={{color:'var(--green)'}} /> ResId в структуре:</div>
                            <div className="comparison-value">
                              {mismatch.structureResId} ({mismatch.structureResName})
                            </div>
                          </div>
                        </div>
                        
                        <div className="mismatch-actions">
                          <button 
                            className="btn-manual-fix"
                            onClick={() => {
                              setFixingNotif(mismatch);
                              setNewResId(mismatch.structureResId);
                              setShowFixModal(true);
                            }}
                          >
                            <span><IconEdit className="ico" /></span>
                            Исправить вручную
                          </button>
                        </div>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}
            
            {/* ВСЕ УВЕДОМЛЕНИЯ */}
            {activeView === 'all' && (
              <div className="all-notifications-view">
                <div className="table-container">
                  <table className="diagnose-table">
                    <thead>
                      <tr>
                        <th>ID</th>
                        <th>Тип</th>
                        <th>ResId</th>
                        <th>РЭС</th>
                        <th>ТП - ВЛ</th>
                        <th>Дата</th>
                        <th>Статус</th>
                      </tr>
                    </thead>
                    <tbody>
                      {diagData.notifications.map(notif => {
                        const isCorrect = !notif.NetworkStructure || 
                          (notif.resId === notif.NetworkStructure.resId);
                        
                        return (
                          <tr key={notif.id} className={isCorrect ? '' : 'mismatch-row'}>
                            <td>{notif.id}</td>
                            <td>
                              <span className={`type-badge ${notif.type}`}>
                                {notif.type}
                              </span>
                            </td>
                            <td>
                              <strong>{notif.resId}</strong>
                              {!isCorrect && (
                                <span className="wrong-icon" title="Несоответствие!"><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /></span>
                              )}
                            </td>
                            <td>{notif.ResUnit?.name || '—'}</td>
                            <td>
                              {notif.NetworkStructure ? 
                                `${notif.NetworkStructure.tpName} - ${notif.NetworkStructure.vlName}` : 
                                '—'}
                            </td>
                            <td>{new Date(notif.createdAt).toLocaleDateString('ru-RU')}</td>
                            <td>
                              {isCorrect ? (
                                <span className="status-ok"><IconCheck className="ico" style={{color:'var(--green)'}} /> OK</span>
                              ) : (
                                <span className="status-error"><IconX className="ico" style={{color:'var(--red)'}} /> Ошибка</span>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
            
            {/* СТРУКТУРА СЕТИ */}
            {activeView === 'structures' && (
              <div className="structures-view">
                <div className="table-container">
                  <table className="diagnose-table">
                    <thead>
                      <tr>
                        <th>ID</th>
                        <th>ResId</th>
                        <th>РЭС</th>
                        <th>ТП</th>
                        <th>ВЛ</th>
                        <th>Начало</th>
                        <th>Середина</th>
                        <th>Конец</th>
                        <th>Уведомления</th>
                      </tr>
                    </thead>
                    <tbody>
                      {diagData.structures.map(struct => {
                        const notifCount = diagData.notifications.filter(
                          n => n.networkStructureId === struct.id
                        ).length;
                        
                        return (
                          <tr key={struct.id}>
                            <td>{struct.id}</td>
                            <td><strong>{struct.resId}</strong></td>
                            <td>{struct.ResUnit?.name}</td>
                            <td>{struct.tpName}</td>
                            <td>{struct.vlName}</td>
                            <td>{struct.startPu || '—'}</td>
                            <td>{struct.middlePu || '—'}</td>
                            <td>{struct.endPu || '—'}</td>
                            <td>
                              {notifCount > 0 ? (
                                <span className="notif-count-badge">{notifCount}</span>
                              ) : (
                                '—'
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
          </div>
        </>
      )}
      
      {/* НОВОЕ: Модальное окно массового исправления */}
      {showMassFixModal && (
        <ModalShell title="Массовое автоисправление" className="mass-fix-modal"
          icon={<IconWrench size={16} />} onClose={() => setShowMassFixModal(false)}>
            <div className="modal-body">
              <div className="mass-fix-info">
                <div className="info-icon"><IconLightbulb className="ico" /></div>
                <div>
                  <p><strong>Будет исправлено несоответствий: {diagData.stats.mismatches}</strong></p>
                  <p>Все resId в уведомлениях будут автоматически установлены согласно структуре сети.</p>
                </div>
              </div>
              
              <div className="mass-fix-preview">
                <h4>Предпросмотр изменений:</h4>
                <div className="preview-list">
                  {diagData.mismatches.slice(0, 5).map((m, idx) => (
                    <div key={idx} className="preview-item">
                      <span className="preview-location">{m.tpName} - {m.vlName}</span>
                      <span className="preview-change">
                        {m.notifResId} <IconArrowRight className="ico" /> {m.structureResId}
                      </span>
                    </div>
                  ))}
                  {diagData.mismatches.length > 5 && (
                    <div className="preview-more">
                      ... и еще {diagData.mismatches.length - 5} изменений
                    </div>
                  )}
                </div>
              </div>
              
              <div className="warning-box">
                <span className="warning-icon"><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /></span>
                <div>
                  <strong>Внимание!</strong>
                  <p>Это действие изменит все несоответствия автоматически. Убедитесь что вы понимаете последствия.</p>
                </div>
              </div>
              
              <div className="form-group">
                <label>
                  <span className="label-icon"><IconLock className="ico" /></span>
                  Введите пароль администратора:
                </label>
                <input
                  type="password"
                  value={massFixPassword}
                  onChange={(e) => setMassFixPassword(e.target.value)}
                  placeholder="Введите пароль"
                  autoFocus
                  className="password-input"
                />
              </div>
            </div>
            
            <div className="modal-footer">
              <button 
                className="cancel-btn" 
                onClick={() => setShowMassFixModal(false)}
              >
                Отмена
              </button>
              <button 
                className="btn-mass-fix-confirm" 
                onClick={handleMassAutoFix}
                disabled={!massFixPassword || massFixing}
              >
                {massFixing ? (
                  <>
                                        Исправление...
                  </>
                ) : (
                  <>
                    <span><IconRocket className="ico" /></span>
                    Исправить ВСЕ ({diagData.stats.mismatches})
                  </>
                )}
              </button>
            </div>
        </ModalShell>
      )}

      {/* Модальное окно ручного исправления */}
      {showFixModal && fixingNotif && (
        <ModalShell title="Ручное исправление ResId" className="fix-modal"
          icon={<IconEdit size={16} />} onClose={() => setShowFixModal(false)}>

            <div className="modal-body">
              <div className="fix-info">
                <p><strong>Уведомление #{fixingNotif.notificationId}</strong></p>
                <p><strong>ТП - ВЛ:</strong> {fixingNotif.tpName} - {fixingNotif.vlName}</p>
              </div>
              
              <div className="fix-comparison">
                <div className="form-group">
                  <label>Текущий ResId:</label>
                  <input 
                    type="text" 
                    value={`${fixingNotif.notifResId} (${fixingNotif.notifResName})`}
                    disabled 
                    className="disabled-input"
                  />
                </div>
                
                <div className="arrow-down"><IconArrowDown className="ico" /></div>
                
                <div className="form-group">
                  <label>Новый ResId:</label>
                  <select 
                    value={newResId}
                    onChange={(e) => setNewResId(e.target.value)}
                    className="select-input"
                  >
                    {resList.map(res => (
                      <option key={res.id} value={res.id}>
                        {res.id} - {res.name}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
              
              <div className="form-group">
                <label>
                  <span className="label-icon"><IconLock className="ico" /></span>
                  Пароль администратора:
                </label>
                <input
                  type="password"
                  value={fixPassword}
                  onChange={(e) => setFixPassword(e.target.value)}
                  placeholder="Введите пароль"
                  className="password-input"
                />
              </div>
            </div>
            
            <div className="modal-footer">
              <button 
                className="cancel-btn" 
                onClick={() => setShowFixModal(false)}
              >
                Отмена
              </button>
              <button 
                className="primary-btn" 
                onClick={handleManualFix}
                disabled={!fixPassword || !newResId}
              >
                <span><IconCheck className="ico" style={{color:'var(--green)'}} /></span>
                Применить
              </button>
            </div>
        </ModalShell>
      )}
    </div>
  );
}

// Подкомпонент настроек структуры
function StructureSettings() {
  const [file, setFile] = useState(null);
  const [uploading, setUploading] = useState(false);
  const [message, setMessage] = useState('');
  const [uploadStats, setUploadStats] = useState(null);
  
  
  const handleFileSelect = (e) => {
    setFile(e.target.files[0]);
    setMessage('');
    setUploadStats(null);
  };

  const handleUploadStructure = async () => {
    if (!file) {
      alert('Выберите файл');
      return;
    }

    setUploading(true);
    const formData = new FormData();
    formData.append('file', file);
    try {
      const response = await api.post('/api/network/upload-full-structure', formData, {
        headers: { 'Content-Type': 'multipart/form-data' }
      });
      
      setMessage('Структура сети успешно загружена!');
      setUploadStats(response.data);
      setFile(null);
      
      // Создаем событие для обновления структуры
      window.dispatchEvent(new CustomEvent('structureUpdated'));
      
    } catch (error) {
      console.error('Upload error:', error);
      setMessage('Ошибка загрузки: ' + (error.response?.data?.error || 'Неизвестная ошибка'));
      setUploadStats(null);
    } finally {
      setUploading(false);
    }
  };
  
  return (
    <div className="settings-section">
      <h3>Загрузка структуры сети</h3>
      <p className="section-description">
        Загрузите Excel файл со структурой сети. Формат: РЭС | ТП | Фидер | Начало | Середина | Конец
      </p>
      
      <div className="upload-area">
        <input 
          type="file" 
          accept=".xlsx,.xls"
          onChange={handleFileSelect}
          id="structure-file"
        />
        <label htmlFor="structure-file" className="file-label">
          {file ? file.name : 'Выберите файл Excel'}
        </label>
      </div>
      
      <button 
        onClick={handleUploadStructure} 
        disabled={uploading || !file}
        className={`primary-btn ${uploading ? 'btn-loading' : ''}`}
      >
        {uploading ? (
          <>
                        Загрузка...
          </>
        ) : (
          'Загрузить структуру'
        )}
      </button>
      
      {message && (
        <div className={message.includes('успешно') ? 'success-message' : 'error-message'}>
          {message}
        </div>
      )}
      
      {uploadStats && (
        <div className="upload-stats">
          <h4>Результаты загрузки:</h4>
          <div className="stats-grid">
            <div className="stat-item">
              <span className="stat-label">Обработано:</span>
              <span className="stat-value">{uploadStats.processed}</span>
            </div>
            <div className="stat-item">
              <span className="stat-label">Всего записей:</span>
              <span className="stat-value">{uploadStats.total}</span>
            </div>
          </div>
          {uploadStats.errors && uploadStats.errors.length > 0 && (
            <div className="errors-list">
              <p><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /> Ошибки при загрузке:</p>
              <ul>
                {uploadStats.errors.slice(0, 5).map((err, idx) => (
                  <li key={idx}>{err}</li>
                ))}
              </ul>
              {uploadStats.errors.length > 5 && (
                <p>... и еще {uploadStats.errors.length - 5} ошибок</p>
              )}
            </div>
          )}
        </div>
      )}
    </div>
  );
}



// Подкомпонент управления пользователями
function UserSettings() {
  const { user: authUser } = useContext(AuthContext);
  const canManageUsers = hasPerm(authUser, 'users_manage');
  const amSuper = !!authUser?.isSuper;   // роль admin может назначать только суперадмин
  const [users, setUsers] = useState([]);
  const [loading, setLoading] = useState(false);
  const [showCreateModal, setShowCreateModal] = useState(false);
  const [showEditModal, setShowEditModal] = useState(false);
  const [editingUser, setEditingUser] = useState(null);
  const [resList, setResList] = useState([]);
  
  // Форма для создания/редактирования
  const [userForm, setUserForm] = useState({
    fio: '',
    login: '',
    password: '',
    email: '',
    role: 'uploader',
    resId: ''
  });
  
  useEffect(() => {
    loadUsers();
    loadResList();
  }, []);
  
  const loadUsers = async () => {
    setLoading(true);
    try {
      const response = await api.get('/api/users/list');
      setUsers(response.data);
    } catch (error) {
      console.error('Error loading users:', error);
    } finally {
      setLoading(false);
    }
  };
  
  const loadResList = async () => {
    try {
      const response = await api.get('/api/res/list');
      setResList(response.data);
    } catch (error) {
      console.error('Error loading RES list:', error);
    }
  };
  
  const handleCreateUser = async () => {
    try {
      await api.post('/api/users/create', userForm);
      showToast('Пользователь создан успешно', 'ok');
      setShowCreateModal(false);
      setUserForm({
        fio: '',
        login: '',
        password: '',
        email: '',
        role: 'uploader',
        resId: ''
      });
      loadUsers();
    } catch (error) {
      showToast(error.response?.data?.error || ('Ошибка создания: ' + error.message));
    }
  };
  
  const handleUpdateUser = async () => {
    try {
      await api.put(`/api/users/${editingUser.id}`, userForm);
      showToast('Пользователь обновлён успешно', 'ok');
      setShowEditModal(false);
      setEditingUser(null);
      loadUsers();
    } catch (error) {
      showToast(error.response?.data?.error || ('Ошибка обновления: ' + error.message));
    }
  };
  
  const handleDeleteUser = async (userId) => {
    if (!confirm('Удалить пользователя?')) return;
    
    const password = prompt('Введите пароль администратора:');
    if (!password) return;
    
    try {
      await api.delete(`/api/users/${userId}`, { data: { password } });
      showToast('Пользователь удалён', 'ok');
      loadUsers();
    } catch (error) {
      showToast(error.response?.data?.error || ('Ошибка удаления: ' + error.message));
    }
  };
  
  const startEdit = (user) => {
    // Учётки администраторов правит только суперадмин — не открываем «сломанную» форму.
    if (user.role === 'admin' && !amSuper) {
      showToast('Учётки администраторов может изменять только суперадмин');
      return;
    }
    setEditingUser(user);
    setUserForm({
      fio: user.fio,
      login: user.login,
      password: '',
      email: user.email,
      role: user.role,
      resId: user.resId || ''
    });
    setShowEditModal(true);
  };
  
  
  return (
    <div className="settings-section">
      <div className="section-header">
        <h3>Управление пользователями</h3>
        <div className="header-actions">
          {canManageUsers && (
            <button onClick={() => setShowCreateModal(true)} className="primary-btn">
              Новый пользователь
            </button>
          )}
        </div>
      </div>
      
      <div className="users-table-container">
        {loading ? (
          <LoadingSpinner type="inline" message="Загрузка пользователей..." />
        ) : (
          <table className="users-table">
            <thead>
              <tr>
                <th>ФИО</th>
                <th>Логин</th>
                <th>Роль</th>
                <th>РЭС</th>
                <th>Email</th>
                <th>Действия</th>
              </tr>
            </thead>
            <tbody>
              {[...users].sort((a, b) => (a.fio || '').localeCompare(b.fio || '', 'ru')).map(user => (
                <tr key={user.id}>
                  <td>{user.fio}</td>
                  <td><strong>{user.login}</strong></td>
                  <td>
                    <span className={`role-badge role-${user.role}`}>
                      {user.role === 'admin' ? 'Админ' : 
                       user.role === 'uploader' ? 'АСКУЭ' : 
                       'ТЕХБЛОК'}
                    </span>
                  </td>
                  <td>{user.ResUnit?.name || '-'}</td>
                  <td>{user.email}</td>
                  <td>
                    <div className="action-buttons">
                      {canManageUsers ? (() => {
                        // Учётки администраторов правит только суперадмин.
                        const adminLock = user.role === 'admin' && !amSuper;
                        const lockTitle = 'Учётки администраторов может изменять только суперадмин';
                        return (
                          <>
                            <button
                              onClick={() => startEdit(user)}
                              className="btn-icon"
                              disabled={adminLock}
                              title={adminLock ? lockTitle : 'Редактировать'}
                            >
                              <IconEdit className="ico" />
                            </button>
                            {!user.isSuper && (
                              <button
                                onClick={() => handleDeleteUser(user.id)}
                                className="btn-icon danger"
                                disabled={adminLock}
                                title={adminLock ? lockTitle : 'Удалить'}
                              >
                                <IconTrash className="ico" />
                              </button>
                            )}
                          </>
                        );
                      })() : <span className="muted">—</span>}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>
      
      {/* Модальное окно создания пользователя */}
      {showCreateModal && (
        <ModalShell title="Создание пользователя" className="user-modal"
          icon={<IconSettings size={16} />} onClose={() => setShowCreateModal(false)}>
            <div className="modal-body">
              <div className="form-group">
                <label>ФИО</label>
                <input
                  type="text"
                  value={userForm.fio}
                  onChange={(e) => setUserForm({...userForm, fio: e.target.value})}
                  placeholder="Иванов Иван Иванович"
                />
              </div>
              
              <div className="form-group">
                <label>Логин</label>
                <input
                  type="text"
                  value={userForm.login}
                  onChange={(e) => setUserForm({...userForm, login: e.target.value})}
                  placeholder="ivanov"
                />
              </div>
              
              <div className="form-group">
                <label>Пароль</label>
                <input
                  type="password"
                  value={userForm.password}
                  onChange={(e) => setUserForm({...userForm, password: e.target.value})}
                  placeholder="Минимум 6 символов"
                />
              </div>
              
              <div className="form-group">
                <label>Email</label>
                <input
                  type="email"
                  value={userForm.email}
                  onChange={(e) => setUserForm({...userForm, email: e.target.value})}
                  placeholder="ivanov@res.ru"
                />
              </div>
              
              <div className="form-group">
                <label>Роль</label>
                <select
                  value={userForm.role}
                  onChange={(e) => setUserForm({...userForm, role: e.target.value})}
                >
                  {amSuper && <option value="admin">Администратор</option>}
                  <option value="uploader">Загрузчик АСКУЭ</option>
                  <option value="res_responsible">Ответственный РЭС</option>
                </select>
              </div>
              
              {userForm.role !== 'admin' && (
                <div className="form-group">
                  <label>РЭС</label>
                  <select
                    value={userForm.resId}
                    onChange={(e) => setUserForm({...userForm, resId: e.target.value})}
                  >
                    <option value="">Выберите РЭС</option>
                    {resList.map(res => (
                      <option key={res.id} value={res.id}>{res.name}</option>
                    ))}
                  </select>
                </div>
              )}
            </div>
            
            <div className="modal-footer">
              <button className="cancel-btn" onClick={() => setShowCreateModal(false)}>
                Отмена
              </button>
              <button 
                className="primary-btn" 
                onClick={handleCreateUser}
                disabled={!userForm.fio || !userForm.login || !userForm.password || !userForm.email}
              >
                Создать
              </button>
            </div>
        </ModalShell>
      )}

      {/* Модальное окно редактирования (аналогично создания) */}
      {showEditModal && (
        <ModalShell title="Редактирование пользователя" className="user-modal"
          icon={<IconSettings size={16} />} onClose={() => setShowEditModal(false)}>
            <div className="modal-body">
              <div className="form-group">
                <label>ФИО</label>
                <input
                  type="text"
                  value={userForm.fio}
                  onChange={(e) => setUserForm({...userForm, fio: e.target.value})}
                />
              </div>
              
              <div className="form-group">
                <label>Логин</label>
                <input
                  type="text"
                  value={userForm.login}
                  onChange={(e) => setUserForm({...userForm, login: e.target.value})}
                />
              </div>
              
              <div className="form-group">
                <label>Новый пароль (оставьте пустым чтобы не менять)</label>
                <input
                  type="password"
                  value={userForm.password}
                  onChange={(e) => setUserForm({...userForm, password: e.target.value})}
                  placeholder="Оставьте пустым"
                />
              </div>
              
              <div className="form-group">
                <label>Email</label>
                <input
                  type="email"
                  value={userForm.email}
                  onChange={(e) => setUserForm({...userForm, email: e.target.value})}
                />
              </div>
              
              <div className="form-group">
                <label>Роль</label>
                <select
                  value={userForm.role}
                  onChange={(e) => setUserForm({...userForm, role: e.target.value})}
                >
                  {amSuper && <option value="admin">Администратор</option>}
                  <option value="uploader">Загрузчик АСКУЭ</option>
                  <option value="res_responsible">Ответственный РЭС</option>
                </select>
              </div>
              
              {userForm.role !== 'admin' && (
                <div className="form-group">
                  <label>РЭС</label>
                  <select
                    value={userForm.resId}
                    onChange={(e) => setUserForm({...userForm, resId: e.target.value})}
                  >
                    <option value="">Выберите РЭС</option>
                    {resList.map(res => (
                      <option key={res.id} value={res.id}>{res.name}</option>
                    ))}
                  </select>
                </div>
              )}
            </div>
            
            <div className="modal-footer">
              <button className="cancel-btn" onClick={() => setShowEditModal(false)}>
                Отмена
              </button>
              <button 
                className="primary-btn" 
                onClick={handleUpdateUser}
              >
                Сохранить
              </button>
            </div>
        </ModalShell>
      )}
    </div>
  );
}

// Подкомпонент обслуживания системы
function MaintenanceSettings() {
  const { user } = useContext(AuthContext);
  const [clearing, setClearing] = useState(false);
  const [showClearModal, setShowClearModal] = useState(false);
  const [clearPassword, setClearPassword] = useState('');
  const [clearBeforeDate, setClearBeforeDate] = useState(''); // НОВОЕ
  
  const handleClearAll = async () => {
    setClearing(true);
    try {
      const response = await api.delete('/api/network/clear-all', {
        data: { 
          password: clearPassword,
          beforeDate: clearBeforeDate  // ДОБАВИЛИ
        }
      });
      
      alert(response.data.message);
      setShowClearModal(false);
      setClearPassword('');
      setClearBeforeDate(''); // ДОБАВИЛИ
      
      window.dispatchEvent(new CustomEvent('dataCleared'));
      
    } catch (error) {
      alert('Ошибка: ' + (error.response?.data?.error || 'Неизвестная ошибка'));
    } finally {
      setClearing(false);
    }
  };
  
  return (
    <div className="settings-section">
      <h3>Обслуживание системы</h3>
      
      {/* Очистка данных — только для учётки admin */}
      {user?.login === 'admin' && (
      <div className="maintenance-card danger">
        <h4><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /> Очистка данных системы</h4>
        <p>Удаляет историю, статусы проверок и уведомления.</p>
        <p className="info-text"><IconCheck className="ico" style={{color:'var(--green)'}} /> Структура сети НЕ удаляется!</p>
        <button
          onClick={() => setShowClearModal(true)}
          disabled={clearing}
          className="danger-btn"
        >
          {clearing ? 'Удаление...' : 'Очистить данные'}
        </button>
      </div>
      )}
      
      <div className="maintenance-card">
        <h4><IconChart className="ico" /> Статистика системы</h4>
        <div className="stats-grid">
          <div className="stat-item">
            <span className="stat-label">Версия системы:</span>
            <span className="stat-value">2.0.1</span>
          </div>
          <div className="stat-item">
            <span className="stat-label">База данных:</span>
            <span className="stat-value">PostgreSQL</span>
          </div>
        </div>
      </div>
      
      {/* Модифицированное модальное окно */}
      {showClearModal && (
        <ModalShell variant="confirm" title="Очистка данных системы" className="delete-modal"
          onClose={() => setShowClearModal(false)}>
            <div className="modal-body">
              {/* НОВОЕ: выбор периода */}
              <div className="form-group">
                <label>Удалить данные до (необязательно):</label>
                <input
                  type="date"
                  value={clearBeforeDate}
                  onChange={(e) => setClearBeforeDate(e.target.value)}
                  max={new Date().toISOString().split('T')[0]}
                />
                {clearBeforeDate && (
                  <p className="info">
                    <IconInfo className="ico" style={{color:'var(--blue)'}} /> Будут удалены данные до {new Date(clearBeforeDate).toLocaleDateString('ru-RU')}
                  </p>
                )}
                {!clearBeforeDate && (
                  <p className="info">
                    <IconInfo className="ico" style={{color:'var(--blue)'}} /> Если дата не указана - будет удалена ВСЯ история
                  </p>
                )}
              </div>
              
              <p className="warning"><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /> ВНИМАНИЕ! Будут удалены:</p>
              <ul>
                <li><IconX className="ico" style={{color:'var(--red)'}} /> <s>Структура сети</s> <span style={{color: 'green'}}>НЕ УДАЛЯЕТСЯ</span></li>
                <li>Все статусы проверок {clearBeforeDate && 'за указанный период'}</li>
                <li>Все уведомления {clearBeforeDate && 'за указанный период'}</li>
                <li>Вся история загрузок {clearBeforeDate && 'за указанный период'}</li>
                <li>Вся история проверок {clearBeforeDate && 'за указанный период'}</li>
              </ul>
              <p className="warning">Это действие НЕЛЬЗЯ отменить!</p>
              <div className="form-group">
                <label>Введите пароль администратора:</label>
                <input
                  type="password"
                  value={clearPassword}
                  onChange={(e) => setClearPassword(e.target.value)}
                  placeholder="Пароль"
                  autoFocus
                />
              </div>
            </div>
            <div className="modal-footer">
              <button className="cancel-btn" onClick={() => setShowClearModal(false)}>
                Отмена
              </button>
              <button 
                className="danger-btn" 
                onClick={handleClearAll}
                disabled={!clearPassword || clearing}
              >
                {clearing ? 'Удаление...' : clearBeforeDate ? 'Удалить старые данные' : 'Удалить всё'}
              </button>
            </div>
        </ModalShell>
      )}
    </div>
  );
}

// =====================================================
// ОСНОВНОЕ ПРИЛОЖЕНИЕ
// =====================================================

// =====================================================
// Компонент для просмотра файлов
// =====================================================
// Загружает файл фоновым XHR (blob:) и показывает <img> — БЕЗ subresource-запроса
// на /api/f (их режут блокировщики Яндекс Protect). onLost — колбэк при 404.
function BlobImage({ file, alt, className, onLost }) {
  const [src, setSrc] = useState(null);
  const [err, setErr] = useState(false);
  useEffect(() => {
    let dead = false, obj = null;
    setSrc(null); setErr(false);
    api.get(fileProxyUrl(file, true), { responseType: 'blob' })
      .then(r => { if (dead) return; obj = URL.createObjectURL(r.data); setSrc(obj); })
      .catch(e => { if (dead) return; setErr(true); if (e.response?.status === 404 && onLost) onLost(file.public_id); });
    return () => { dead = true; if (obj) URL.revokeObjectURL(obj); };
  }, [file.public_id]);
  if (err) return <div className={`blob-img-fallback ${className || ''}`}><IconAlertTriangle className="ico" /></div>;
  if (!src) return <div className={`blob-img-fallback ${className || ''}`}><RossetiLoader size="small" /></div>;
  return <img src={src} alt={alt} className={className} />;
}

// Просмотр вложений БЕЗ навигации: файл тянется фоновым XHR в blob:, картинка —
// PDF в <canvas> через pdf.js (БЕЗ <iframe> — Яндекс Protect режет навигацию
// вложенных фреймов, включая blob:). blob уже загружен XHR-ом. pdfjs-dist грузим
// ЛЕНИВО динамическим import — отдельным чанком, не раздувая основной бандл.
function PdfCanvas({ blob, downloadHref, downloadName }) {
  const containerRef = useRef(null);
  const docRef = useRef(null);
  const [numPages, setNumPages] = useState(0);
  const [pageNow, setPageNow] = useState(1);
  const [scale, setScale] = useState(1);       // множитель поверх fit-по-ширине
  const [status, setStatus] = useState('loading'); // loading | ready | error
  const [renderTick, setRenderTick] = useState(0); // перерисовка при смене ширины (fullscreen/resize)

  // Перерисовать канвас при заметной смене ширины контейнера (fullscreen, ресайз окна) —
  // НЕ CSS-растяжение готового канваса (мыло), а честный ре-рендер под новую ширину.
  useEffect(() => {
    const el = containerRef.current;
    if (!el || typeof ResizeObserver === 'undefined') return;
    let lastW = el.clientWidth, t = null;
    const ro = new ResizeObserver(() => {
      const w = el.clientWidth;
      if (Math.abs(w - lastW) < 4) return;
      lastW = w;
      clearTimeout(t);
      t = setTimeout(() => setRenderTick(x => x + 1), 120);
    });
    ro.observe(el);
    return () => { ro.disconnect(); clearTimeout(t); };
  }, []);

  // Загрузка документа (ленивый import pdfjs)
  useEffect(() => {
    let dead = false;
    setStatus('loading'); setNumPages(0); setPageNow(1);
    (async () => {
      try {
        const pdfjsLib = await import('pdfjs-dist');
        pdfjsLib.GlobalWorkerOptions.workerSrc =
          new URL('pdfjs-dist/build/pdf.worker.min.mjs', import.meta.url).toString();
        const data = await blob.arrayBuffer();
        if (dead) return;
        const doc = await pdfjsLib.getDocument({ data }).promise;
        if (dead) { try { doc.destroy(); } catch (e) {} return; }
        docRef.current = doc;
        setNumPages(doc.numPages);
        setStatus('ready');
      } catch (e) {
        console.error('pdf.js error:', e);
        if (!dead) setStatus('error');
      }
    })();
    return () => {
      dead = true;
      if (docRef.current) { try { docRef.current.destroy(); } catch (e) {} docRef.current = null; }
    };
  }, [blob]);

  // Рендер всех страниц вертикальной лентой (учёт devicePixelRatio — текст чёткий)
  useEffect(() => {
    if (status !== 'ready' || !docRef.current || !containerRef.current) return;
    let cancelled = false;
    const doc = docRef.current;
    const container = containerRef.current;
    (async () => {
      container.innerHTML = '';
      const dpr = window.devicePixelRatio || 1;
      const cw = (container.clientWidth || 800) - 8;
      for (let p = 1; p <= doc.numPages; p++) {
        if (cancelled) return;
        const page = await doc.getPage(p);
        const base = page.getViewport({ scale: 1 });
        const fit = cw / base.width;
        const viewport = page.getViewport({ scale: fit * scale * dpr });
        const canvas = document.createElement('canvas');
        canvas.className = 'pdf-page-canvas';
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        canvas.style.width = (viewport.width / dpr) + 'px';
        canvas.style.height = (viewport.height / dpr) + 'px';
        container.appendChild(canvas);
        await page.render({ canvasContext: canvas.getContext('2d'), viewport }).promise;
      }
    })();
    return () => { cancelled = true; };
  }, [status, scale, renderTick]);

  const onScroll = () => {
    const c = containerRef.current; if (!c) return;
    const canvases = c.querySelectorAll('canvas');
    let cur = 1;
    for (let i = 0; i < canvases.length; i++) {
      if (canvases[i].offsetTop - c.scrollTop <= c.clientHeight / 3) cur = i + 1;
    }
    setPageNow(cur);
  };

  if (status === 'error') {
    return (
      <div className="file-not-supported">
        <p style={{ color: 'var(--red)' }}>Не удалось отобразить PDF</p>
        <a href={downloadHref} download={downloadName} className="download-link">Скачать файл</a>
      </div>
    );
  }

  return (
    <div className="pdf-canvas-viewer">
      <div className="pdf-toolbar">
        <span className="pdf-pageinfo">стр. {pageNow} из {numPages || '…'}</span>
        <div className="pdf-zoom">
          <button className="pdf-zoom-btn" onClick={() => setScale(s => Math.max(0.5, +(s - 0.2).toFixed(2)))} disabled={status !== 'ready'}>−</button>
          <span className="pdf-zoom-val">{Math.round(scale * 100)}%</span>
          <button className="pdf-zoom-btn" onClick={() => setScale(s => Math.min(3, +(s + 0.2).toFixed(2)))} disabled={status !== 'ready'}>+</button>
        </div>
        <a href={downloadHref} download={downloadName} className="btn-download-pdf pdf-toolbar-dl">
          <IconDownload className="ico" /> Скачать
        </a>
      </div>
      {status === 'loading' && <div className="file-viewer-status"><RossetiLoader /></div>}
      <div className="pdf-pages" ref={containerRef} onScroll={onScroll} />
    </div>
  );
}

// Просмотр вложений БЕЗ навигации: файл тянется фоновым XHR в blob:, картинка —
// в <img>, PDF — в <canvas> через pdf.js (никаких iframe/object/embed).
function FileViewer({ files, currentIndex, onClose, onNext, onPrev }) {
  const currentFile = files[currentIndex];
  const url = (currentFile.url || '').toLowerCase();
  const isImage = url.endsWith('.jpg') || url.endsWith('.jpeg') || url.endsWith('.png') || url.endsWith('.gif');
  const isPdf = url.endsWith('.pdf');

  const [blobUrl, setBlobUrl] = useState(null); // objectURL для <img>
  const [pdfBlob, setPdfBlob] = useState(null); // сырой blob для pdf.js
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null); // { lost: bool }

  useEffect(() => {
    let dead = false, obj = null;
    setBlobUrl(null); setPdfBlob(null); setError(null); setLoading(true);
    if (!isImage && !isPdf) { setLoading(false); return; }
    api.get(fileProxyUrl(currentFile, true), { responseType: 'blob' })
      .then(r => {
        if (dead) return;
        if (isImage) { obj = URL.createObjectURL(r.data); setBlobUrl(obj); }
        else { setPdfBlob(r.data); }
      })
      .catch(e => { if (dead) return; setError({ lost: e.response?.status === 404 }); })
      .finally(() => { if (!dead) setLoading(false); });
    return () => { dead = true; if (obj) URL.revokeObjectURL(obj); };
    // eslint-disable-next-line
  }, [currentFile.public_id]);

  return (
    <ModalShell
      title={`Просмотр файлов (${currentIndex + 1} из ${files.length})`}
      dockTitle={currentFile.original_name}
      icon={<IconFolder size={16} />}
      className="file-viewer-container"
      onClose={onClose}
    >
        <div className="file-viewer-content">
          {loading ? (
            <div className="file-viewer-status"><RossetiLoader /></div>
          ) : error ? (
            <div className="file-not-supported">
              {error.lost
                ? <div className="file-lost-badge"><IconAlertTriangle className="ico" /> Файл утерян в хранилище</div>
                : <p style={{ color: 'var(--red)' }}>Не удалось загрузить файл</p>}
              <a href={fileProxyUrl(currentFile)} download={currentFile.original_name} className="download-link">
                Попробовать скачать
              </a>
            </div>
          ) : isImage ? (
            <div className="image-viewer-wrap">
              <img src={blobUrl} alt={currentFile.original_name} className="file-viewer-image" />
              <a href={fileProxyUrl(currentFile)} download={currentFile.original_name} className="btn-download-pdf image-download-btn">
                <span><IconDownload className="ico" /></span>
                Скачать {currentFile.original_name}
              </a>
            </div>
          ) : isPdf && pdfBlob ? (
            <PdfCanvas blob={pdfBlob} downloadHref={fileProxyUrl(currentFile)} downloadName={currentFile.original_name} />
          ) : (
            <div className="file-not-supported">
              <p>Предпросмотр недоступен</p>
              <a href={fileProxyUrl(currentFile)} download={currentFile.original_name} className="download-link">
                Скачать файл
              </a>
            </div>
          )}
        </div>

        <div className="file-viewer-info">
          <p><strong>Имя файла:</strong> {currentFile.original_name}</p>
          <p><strong>Загружен:</strong> {currentFile.uploaded_at ? new Date(currentFile.uploaded_at).toLocaleString('ru-RU') : '—'}</p>
        </div>

        {files.length > 1 && (
          <div className="file-viewer-navigation">
            <button onClick={onPrev} className="nav-btn">
              <IconArrowLeft className="ico" /> Предыдущий
            </button>
            <button onClick={onNext} className="nav-btn">
              Следующий <IconArrowRight className="ico" />
            </button>
          </div>
        )}
    </ModalShell>
  );
}

// =====================================================
// КОМПОНЕНТ ЗАГРУЖЕННЫХ ДОКУМЕНТОВ
// =====================================================

function UploadedDocuments() {
  const [documents, setDocuments] = useState([]);
  const [loading, setLoading] = useState(true);
  const [selectedFile, setSelectedFile] = useState(null);
  const [showDeleteModal, setShowDeleteModal] = useState(false);
  const [deletePassword, setDeletePassword] = useState('');
  const [showFileViewer, setShowFileViewer] = useState(false);
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [currentFileIndex, setCurrentFileIndex] = useState(0);
  const { user, selectedRes } = useContext(AuthContext);
  const canManageFiles = hasPerm(user, 'files_manage');

  const [deleteRecordId, setDeleteRecordId] = useState(null);
  const [showDeleteRecordModal, setShowDeleteRecordModal] = useState(false);
  const [selectedIds, setSelectedIds] = useState([]);
  const [showBulkDeleteModal, setShowBulkDeleteModal] = useState(false);
  
  useEffect(() => {
    loadDocuments();
  }, [selectedRes]);
  
  const loadDocuments = async () => {
    try {
      // ИСПРАВЛЕНО: используем selectedRes для фильтрации
      const params = new URLSearchParams();
      
      if (user.role === 'admin' && selectedRes) {
        params.append('resId', selectedRes);
      }
      
      const queryString = params.toString();
      const url = `/api/documents/list${queryString ? `?${queryString}` : ''}`;
      
      console.log('Loading documents:', { url, selectedRes });
      
      const response = await api.get(url);
      setDocuments(response.data);
    } catch (error) {
      console.error('Error loading documents:', error);
    } finally {
      setLoading(false);
    }
  };
  
  const handleViewFile = (files) => {
    setSelectedFiles(files);
    setCurrentFileIndex(0);
    setShowFileViewer(true);
  };
  
  const handleDeleteFile = async () => {
    try {
      await api.delete(`/api/documents/${selectedFile.recordId}/${selectedFile.fileIndex}`, {
        data: { password: deletePassword }
      });
      
      alert('Файл удален успешно');
      setShowDeleteModal(false);
      setDeletePassword('');
      setSelectedFile(null);
      loadDocuments();
      
    } catch (error) {
      alert('Ошибка удаления: ' + (error.response?.data?.error || error.message));
    }
  };
  
  const handleDeleteRecord = async () => {
    try {
      await api.delete(`/api/documents/record/${deleteRecordId}`, {
        data: { password: deletePassword }
      });
      
      alert('Запись удалена успешно');
      setShowDeleteRecordModal(false);
      setDeletePassword('');
      setDeleteRecordId(null);
      loadDocuments();
      
    } catch (error) {
      alert('Ошибка удаления: ' + (error.response?.data?.error || error.message));
    }
  };

  const handleSelectRecord = (id) => {
    setSelectedIds(prev => {
      if (prev.includes(id)) {
        return prev.filter(i => i !== id);
      } else {
        return [...prev, id];
      }
    });
  };

  const handleSelectAll = () => {
    if (selectedIds.length === documents.length) {
      setSelectedIds([]);
    } else {
      setSelectedIds(documents.map(doc => doc.id));
    }
  };

  const handleBulkDelete = async () => {
    try {
      await api.post('/api/documents/delete-bulk', {
        ids: selectedIds,
        password: deletePassword
      });
      
      alert(`Удалено записей: ${selectedIds.length}`);
      setShowBulkDeleteModal(false);
      setDeletePassword('');
      setSelectedIds([]);
      loadDocuments();
      
    } catch (error) {
      alert('Ошибка удаления: ' + (error.response?.data?.error || error.message));
    }
  };
  
  if (loading) return <LoadingSpinner type="dots" message="Загрузка документов..." />;
  
  return (
    <div className="uploaded-documents">
      <PageHeader icon={<IconFolder size={24} />} title="Загруженные документы" />
      
      <div className="documents-controls">
        <div className="documents-info">
          <p>Всего документов: <strong>{documents.reduce((sum, doc) => sum + (doc.attachments?.length || 0), 0)}</strong></p>
        </div>
        
        {canManageFiles && selectedIds.length > 0 && (
  <button
    className="delete-selected-btn"
    onClick={() => setShowBulkDeleteModal(true)}
  >
    <IconTrash className="ico" /> Удалить выбранные ({selectedIds.length})
  </button>
)}
</div>
      
      <div className="documents-table">
        <table>
          <thead>
            <tr>
              {canManageFiles && (
                <th className="checkbox-column">
                  <input
                    type="checkbox"
                    checked={selectedIds.length === documents.length && documents.length > 0}
                    onChange={handleSelectAll}
                  />
                </th>
              )}
              <th>ТП</th>
              <th>ВЛ</th>
              <th>ПУ №</th>
              <th>Загрузил</th>
              <th>Дата загрузки</th>
              <th>Комментарий</th>
              <th>Статус</th>
              <th>Файлы</th>
              <th>Действия</th>
            </tr>
          </thead>
          <tbody>
            {documents.map((doc) => (
              <tr key={doc.id} className={selectedIds.includes(doc.id) ? 'selected' : ''}>
                {canManageFiles && (
                  <td className="checkbox-column">
                    <input
                      type="checkbox"
                      checked={selectedIds.includes(doc.id)}
                      onChange={() => handleSelectRecord(doc.id)}
                    />
                  </td>
                )}
                <td>{doc.tpName}</td>
                <td>{doc.vlName}</td>
                <td><strong>{doc.puNumber}</strong></td>
                <td>{doc.uploadedBy}</td>
                <td>{new Date(doc.workCompletedDate).toLocaleDateString('ru-RU')}</td>
                <td className="comment-cell">{doc.resComment}</td>
                <td>
                  <span className={`status-badge status-${doc.status}`}>
                    {doc.status === 'completed' ? 'Завершен' : 'На проверке'}
                  </span>
                </td>
                <td>
                  <span className="file-count">{doc.attachments?.length || 0} файл(ов)</span>
                </td>
                <td>
                  <div className="action-buttons">
                    {doc.attachments && doc.attachments.length > 0 && (
                      <button 
                        className="btn-view"
                        onClick={() => handleViewFile(doc.attachments)}
                        title="Просмотреть"
                      >
                        <IconEye className="ico ico-glow-blue" />
                      </button>
                    )}
                      {canManageFiles && (
      <button
        className="btn-icon danger"
        onClick={() => {
          setDeleteRecordId(doc.id);
          setShowDeleteRecordModal(true);
        }}
        title="Удалить запись"
      >
        <IconTrash className="ico ico-glow-red" />
      </button>
    )}
                  </div>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      
      {documents.length === 0 && (
        <div className="no-data">
          <p>Пока нет загруженных документов</p>
        </div>
      )}

      {/* Модальное окно массового удаления */}
{showBulkDeleteModal && (
  <ModalShell variant="confirm" title="Подтверждение удаления" className="delete-modal"
    onClose={() => { setShowBulkDeleteModal(false); setDeletePassword(''); }}>
      <div className="modal-body">
        <div className="delete-summary">
          <div className="delete-icon"><IconTrash className="ico" /></div>
          <div>
            {/* ИСПРАВЛЕНО: selectedIds вместо selectedNotificationIds */}
            <p className="delete-title">Вы собираетесь удалить <strong>{selectedIds.length}</strong> записей с документами</p>
          </div>
        </div>
        
        <p className="warning"><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /> Это действие нельзя отменить!</p>
        <p>Будут удалены все файлы и записи истории для выбранных документов.</p>
        
        <div className="form-group">
          <label>Введите пароль администратора:</label>
          <input
            type="password"
            value={deletePassword}
            onChange={(e) => setDeletePassword(e.target.value)}
            placeholder="Пароль"
            autoFocus
            autoComplete="new-password"
            name={`delete-password-${Date.now()}`}
          />
        </div>
      </div>
      
      <div className="modal-footer">
        <button className="cancel-btn" onClick={() => {
          setShowBulkDeleteModal(false); 
          setDeletePassword('');
        }}>
          Отмена
        </button>
        <button 
          className="danger-btn" 
          onClick={handleBulkDelete}
          disabled={!deletePassword}
        >
          <IconTrash className="ico" /> Удалить записи
        </button>
      </div>
  </ModalShell>
)}

      {/* Модальное окно удаления записи */}
      {showDeleteRecordModal && (
        <ModalShell variant="confirm" title="Подтверждение удаления записи" className="delete-modal"
          onClose={() => setShowDeleteRecordModal(false)}>
            <div className="modal-body">
              <p>Вы собираетесь удалить всю запись вместе со всеми файлами.</p>
              <p className="warning"><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /> Это действие нельзя отменить!</p>
              <div className="form-group">
                <label>Введите пароль администратора:</label>
                <input
                  type="password"
                  value={deletePassword}
                  onChange={(e) => setDeletePassword(e.target.value)}
                  placeholder="Пароль"
                  autoFocus
                />
              </div>
            </div>
            <div className="modal-footer">
              <button className="cancel-btn" onClick={() => setShowDeleteRecordModal(false)}>
                Отмена
              </button>
              <button 
                className="danger-btn" 
                onClick={handleDeleteRecord}
                disabled={!deletePassword}
              >
                Удалить запись
              </button>
            </div>
        </ModalShell>
      )}
      
      {/* Модальное окно удаления файла */}
      {showDeleteModal && (
        <ModalShell variant="confirm" title="Подтверждение удаления файла" className="delete-modal"
          onClose={() => setShowDeleteModal(false)}>
            <div className="modal-body">
              <p>Вы собираетесь удалить файл:</p>
              <p><strong>{selectedFile?.original_name}</strong></p>
              <p className="warning"><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /> Это действие нельзя отменить!</p>
              <div className="form-group">
                <label>Введите пароль администратора:</label>
                <input
                  type="password"
                  value={deletePassword}
                  onChange={(e) => setDeletePassword(e.target.value)}
                  placeholder="Пароль"
                  autoFocus
                />
              </div>
            </div>
            <div className="modal-footer">
              <button className="cancel-btn" onClick={() => setShowDeleteModal(false)}>
                Отмена
              </button>
              <button 
                className="danger-btn" 
                onClick={handleDeleteFile}
                disabled={!deletePassword}
              >
                Удалить
              </button>
            </div>
        </ModalShell>
      )}

      {/* Просмотрщик файлов */}
      {showFileViewer && (
        <FileViewer 
          files={selectedFiles}
          currentIndex={currentFileIndex}
          onClose={() => setShowFileViewer(false)}
          onNext={() => setCurrentFileIndex((prev) => (prev + 1) % selectedFiles.length)}
          onPrev={() => setCurrentFileIndex((prev) => (prev - 1 + selectedFiles.length) % selectedFiles.length)}
        />
      )}
    </div>
  );
}


// Новый компонент расширенного модального окна
function ExtendedPuModal({ 
  isOpen, 
  onClose, 
  puData, 
  activeTab, 
  setActiveTab, 
  uploadHistory, 
  checkHistory, 
  loading,
  handleClearPuHistory
}) {
  const { user } = useContext(AuthContext);
  if (!isOpen || !puData) return null;
  
  // Парсим детали ошибки для отображения фаз
  const getPhaseErrors = () => {
    const phases = { A: false, B: false, C: false };
    
    if (puData.status.errorDetails) {
      try {
        const parsed = JSON.parse(puData.status.errorDetails);
        const errorSummary = parsed.summary || puData.status.errorDetails;
        
        if (errorSummary.indexOf('Фаза A') !== -1) phases.A = true;
        if (errorSummary.indexOf('Фаза B') !== -1) phases.B = true;
        if (errorSummary.indexOf('Фаза C') !== -1) phases.C = true;
      } catch (e) {
        // Если не JSON, проверяем как текст
        const errorText = puData.status.errorDetails;
        if (errorText.indexOf('Фаза A') !== -1) phases.A = true;
        if (errorText.indexOf('Фаза B') !== -1) phases.B = true;
        if (errorText.indexOf('Фаза C') !== -1) phases.C = true;
      }
    }
    
    return phases;
  };
  
  const phaseErrors = getPhaseErrors();
  
  return (
    <ModalShell title={`ПУ #${puData.puNumber} — Детальная информация`} className="extended-pu-modal"
      dockTitle={`ПУ #${puData.puNumber}`} icon={<IconMeter size={16} />} onClose={onClose}>
        {/* Информация о местоположении */}
        <div className="pu-location-info">
          <p><strong>РЭС:</strong> {puData.resName}</p>
          <p><strong>ТП:</strong> {puData.tpName}</p>
          <p><strong>Фидер:</strong> {puData.vlName}</p>
          <p><strong>Позиция:</strong> {
            puData.position === 'start' ? 'Начало' : 
            puData.position === 'middle' ? 'Середина' : 'Конец'
          }</p>
        </div>
        
        {/* Вкладки */}
        <div className="modal-tabs">
          <button 
            className={`tab-btn ${activeTab === 'current' ? 'active' : ''}`}
            onClick={() => setActiveTab('current')}
          >
            Текущее состояние
          </button>
          <button 
            className={`tab-btn ${activeTab === 'uploads' ? 'active' : ''}`}
            onClick={() => setActiveTab('uploads')}
          >
            История загрузок ({uploadHistory.length})
          </button>
          <button 
            className={`tab-btn ${activeTab === 'checks' ? 'active' : ''}`}
            onClick={() => setActiveTab('checks')}
          >
            История проверок ({checkHistory.length})
          </button>
        </div>
        
        <div className="modal-body">
          {loading ? (
            <div className="loading"><RossetiLoader /></div>
          ) : (
            <>
              {/* Вкладка текущего состояния */}
              {activeTab === 'current' && (
                <div className="tab-content">
                  {puData.status.status === 'checked_error' ? (
                    <>
                      <div className="phase-indicators-large">
                        <div className={`phase-indicator ${phaseErrors.A ? 'phase-error' : ''}`}>A</div>
                        <div className={`phase-indicator ${phaseErrors.B ? 'phase-error' : ''}`}>B</div>
                        <div className={`phase-indicator ${phaseErrors.C ? 'phase-error' : ''}`}>C</div>
                      </div>
                      
                      <div className="error-details-box">
                        <h4>Обнаруженные отклонения:</h4>
                        <div className="error-text">
                          {(() => {
                            try {
                              const parsed = JSON.parse(puData.status.errorDetails);
                              return parsed.summary || puData.status.errorDetails;
                            } catch {
                              return puData.status.errorDetails;
                            }
                          })()}
                        </div>
                      </div>
                      
                      <div className="error-meta">
                        <p><strong>Последняя проверка:</strong> {
                          puData.status.lastCheck 
                            ? new Date(puData.status.lastCheck).toLocaleString('ru-RU')
                            : 'Неизвестно'
                        }</p>
                      </div>
                    </>
                  ) : puData.status.status === 'checked_ok' ? (
                    <div className="success-state">
                      <div className="success-icon"><IconCheck className="ico" style={{color:'var(--green)'}} /></div>
                      <h4>Проверен без ошибок</h4>
                      <p>Последняя проверка: {
                        puData.status.lastCheck 
                          ? new Date(puData.status.lastCheck).toLocaleString('ru-RU')
                          : 'Неизвестно'
                      }</p>
                    </div>
                  ) : puData.status.status === 'pending_recheck' ? (
                    <div className="pending-state">
                      <div className="pending-icon"><IconClock className="ico" /></div>
                      <h4>Ожидает перепроверки АСКУЭ</h4>
                      <p>Мероприятия выполнены РЭС, требуется загрузить новый файл для проверки</p>
                    </div>
                  ) : (
                    <div className="not-checked-state">
                      <div className="not-checked-icon"><IconHelp className="ico" /></div>
                      <h4>Не проверялся</h4>
                      <p>Для этого ПУ еще не загружались файлы для анализа</p>
                    </div>
                  )}
                </div>
              )}
              
              {/* Вкладка истории загрузок */}
              {activeTab === 'uploads' && (
                <div className="tab-content">
                  {uploadHistory.length === 0 ? (
                    <p className="no-data">Нет истории загрузок</p>
                  ) : (
                    <div className="history-table">
                      <table>
                        <thead>
                          <tr>
                            <th>Дата</th>
                            <th>Загрузил</th>
                            <th>Файл</th>
                            <th>Статус</th>
                            <th>Ошибка</th>
                          </tr>
                        </thead>
                        <tbody>
                          {uploadHistory.map((upload, idx) => (
                            <tr key={idx} className={upload.uploadStatus}>
                              <td>{new Date(upload.uploadedAt).toLocaleString('ru-RU')}</td>
                              <td>{upload.User?.fio || 'Неизвестно'}</td>
                              <td>{upload.fileName}</td>
                              <td>
                                <span className={`status-badge status-${upload.uploadStatus}`}>
                                  {upload.uploadStatus === 'success' ? 'Успешно' :
                                   upload.uploadStatus === 'duplicate' ? 'Дубликат' :
                                   upload.uploadStatus === 'wrong_period' ? 'Неверный период' :
                                   'Ошибка'}
                                </span>
                              </td>
                              <td className="error-cell">
                                {upload.hasErrors ? (
                                  <details>
                                    <summary>Показать ошибку</summary>
                                    <pre>{upload.errorSummary}</pre>
                                  </details>
                                ) : (
                                  '—'
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              )}
              
              {/* Вкладка истории проверок */}
              {activeTab === 'checks' && (
                <div className="tab-content">
                  {checkHistory.length === 0 ? (
                    <p className="no-data">Нет истории проверок</p>
                  ) : (
                    <div className="history-timeline">
                      {checkHistory.map((check, idx) => (
                        <div key={idx} className="timeline-item">
                          <div className="timeline-date">
                            {new Date(check.initialCheckDate).toLocaleDateString('ru-RU')}
                          </div>
                          
                          <div className="timeline-content">
                            <div className="timeline-step error">
                              <h5><IconAlertCircle className="ico" style={{color:'var(--red)'}} /> Обнаружена ошибка</h5>
                              <p>{check.initialError}</p>
                            </div>
                            
                            {check.workCompletedDate && (
                              <div className="timeline-step work">
                                <h5><IconWrench className="ico" /> Мероприятия выполнены</h5>
                                <p><strong>Дата:</strong> {new Date(check.workCompletedDate).toLocaleDateString('ru-RU')}</p>
                                <p><strong>Комментарий:</strong> {check.resComment}</p>
                                {check.attachments && check.attachments.length > 0 && (
                                  <p><strong>Файлов:</strong> {check.attachments.length}</p>
                                )}
                              </div>
                            )}
                            
                            {check.recheckDate && (
                              <div className={`timeline-step recheck ${check.recheckResult}`}>
                                <h5>{check.recheckResult === 'ok' ? 'Перепроверка успешна' : 'Ошибка не устранена'}</h5>
                                <p><strong>Дата:</strong> {new Date(check.recheckDate).toLocaleDateString('ru-RU')}</p>
                              </div>
                            )}
                            
                            <div className="timeline-status">
                              <strong>Текущий статус:</strong> {
                                check.status === 'awaiting_work' ? 'Ожидает мероприятий' :
                                check.status === 'awaiting_recheck' ? 'Ожидает перепроверки' :
                                'Завершено'
                              }
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  )}
                </div>
              )}
            </>
          )}
        </div>
        
        <div className="modal-footer">
          <button className="action-btn" onClick={onClose}>Закрыть</button>
          {hasPerm(user, 'checks_delete') && (
            <button
              className="danger-btn"
              onClick={() => {
                onClose();
                handleClearPuHistory(puData.puNumber);
              }}
            >
              <IconTrash className="ico" /> Очистить историю
            </button>
          )}
        </div>
    </ModalShell>
  );
}

// =====================================================
// КОМПОНЕНТ ИСТОРИИ СИСТЕМЫ
// =====================================================

function SystemHistory() {
  const [activeTab, setActiveTab] = useState('uploads'); // uploads или checks
  const [uploads, setUploads] = useState([]);
  const [checks, setChecks] = useState([]);
  const [loading, setLoading] = useState(false);
  const [currentPage, setCurrentPage] = useState(1);
  const [totalPages, setTotalPages] = useState(1);
  const { user } = useContext(AuthContext);
  const [resList, setResList] = useState([]);
  
  // Фильтры
  const [filters, setFilters] = useState({
  puNumber: '',
  tpName: '',
  resId: user.role === 'admin' ? '' : user.resId, // ИЗМЕНЕНО - не-админы видят только свой РЭС
  dateFrom: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000).toISOString().split('T')[0],
  dateTo: new Date().toISOString().split('T')[0],
  fileType: '',
  status: ''
});

  useEffect(() => {
    loadResList();
  }, []);
  
  const loadResList = async () => {
    try {
      const response = await api.get('/api/res/list');
      setResList(response.data);
    } catch (error) {
      console.error('Error loading RES list:', error);
    }
  };
  
  useEffect(() => {
    loadData();
  }, [activeTab, currentPage, filters]);
  
  const loadData = async () => {
  setLoading(true);
  try {
    if (activeTab === 'uploads') {
      const params = new URLSearchParams({
        page: currentPage,
        limit: 100,
        ...filters,
        resId: user.role === 'admin' ? filters.resId : user.resId // Принудительно для не-админов
      });
      
      const response = await api.get(`/api/history/uploads?${params}`);
      setUploads(response.data.uploads);
      setTotalPages(response.data.totalPages);
    } else {
      const params = new URLSearchParams({
        page: currentPage,
        limit: 100,
        puNumber: filters.puNumber,
        resId: user.role === 'admin' ? filters.resId : user.resId, // Принудительно для не-админов
        tpName: filters.tpName,
        dateFrom: filters.dateFrom,
        dateTo: filters.dateTo,
        status: filters.status
      });
      
      const response = await api.get(`/api/history/checks?${params}`);
      setChecks(response.data.checks);
      setTotalPages(response.data.totalPages);
    }
  } catch (error) {
    console.error('Error loading history:', error);
  } finally {
    setLoading(false);
  }
};
  
  const handleFilterChange = (key, value) => {
    setFilters(prev => ({ ...prev, [key]: value }));
    setCurrentPage(1);
  };
  
  const exportToExcel = () => {
    const data = activeTab === 'uploads' ? uploads : checks;
    if (data.length === 0) {
      alert('Нет данных для экспорта');
      return;
    }
    
    let exportData;
    if (activeTab === 'uploads') {
      exportData = uploads.map(upload => ({
        'Дата загрузки': new Date(upload.uploadedAt).toLocaleString('ru-RU'),
        'РЭС': upload.resName || '',
        'ТП': upload.tpName || '',
        'ВЛ': upload.vlName || '',
        'Номер ПУ': upload.puNumber,
        'Загрузил': upload.User?.fio || '',
        'Тип файла': upload.fileType,
        'Статус': upload.uploadStatus === 'success' ? 'Успешно' : 
                  upload.uploadStatus === 'duplicate' ? 'Дубликат' : 
                  upload.uploadStatus === 'wrong_period' ? 'Неверный период' : 'Ошибка',
        'Есть ошибки': upload.hasErrors ? 'Да' : 'Нет',
        'Текст ошибки': upload.errorSummary || ''
      }));
    } else {
      exportData = checks.map(check => ({
        'РЭС': check.ResUnit?.name || '',
        'ТП': check.tpName,
        'ВЛ': check.vlName,
        'Номер ПУ': check.puNumber,
        'Позиция': check.position === 'start' ? 'Начало' : 
                   check.position === 'middle' ? 'Середина' : 'Конец',
        'Дата обнаружения': new Date(check.initialCheckDate).toLocaleString('ru-RU'),
        'Первоначальная ошибка': check.initialError,
        'Дата выполнения работ': check.workCompletedDate ? 
          new Date(check.workCompletedDate).toLocaleString('ru-RU') : '',
        'Комментарий РЭС': check.resComment || '',
        'Дата перепроверки': check.recheckDate ? 
          new Date(check.recheckDate).toLocaleString('ru-RU') : '',
        'Результат': check.recheckResult === 'ok' ? 'Исправлено' : 
                     check.recheckResult === 'error' ? 'Не исправлено' : 'Ожидает',
        'Статус': check.status === 'awaiting_work' ? 'Ожидает мероприятий' :
                  check.status === 'awaiting_recheck' ? 'Ожидает перепроверки' : 'Завершено'
      }));
    }
    
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(exportData);
    
    // Устанавливаем ширину колонок
    const maxWidth = 50;
    const cols = Object.keys(exportData[0] || {}).map(() => ({ wch: maxWidth }));
    ws['!cols'] = cols;
    
    XLSX.utils.book_append_sheet(wb, ws, activeTab === 'uploads' ? 'История загрузок' : 'История проверок');
    
    const fileName = `История_${activeTab === 'uploads' ? 'загрузок' : 'проверок'}_${new Date().toLocaleDateString('ru-RU').split('.').join('-')}.xlsx`;
    XLSX.writeFile(wb, fileName);
  };
  
  return (
    <div className="system-history">
      <PageHeader icon={<IconClock size={24} />} title="История системы" />

      {user.role !== 'admin' && (
    <div className="res-indicator">
      <span>Показаны данные для: <strong>{user.resName}</strong></span>
    </div>
  )}

      
      <div className="history-tabs">
        <button 
          className={`tab-btn ${activeTab === 'uploads' ? 'active' : ''}`}
          onClick={() => {
            setActiveTab('uploads');
            setCurrentPage(1);
          }}
        >
          История загрузок
        </button>
        <button 
          className={`tab-btn ${activeTab === 'checks' ? 'active' : ''}`}
          onClick={() => {
            setActiveTab('checks');
            setCurrentPage(1);
          }}
        >
          История проверок
        </button>
      </div>
      
      <div className="history-filters">
        <div className="filter-row">
          <div className="filter-group">
            <label>Номер ПУ:</label>
            <input 
              type="text"
              value={filters.puNumber}
              onChange={(e) => handleFilterChange('puNumber', e.target.value)}
              placeholder="Поиск по ПУ"
            />
          </div>

          {/* ДОБАВИТЬ выбор РЭС */}
          <div className="filter-group">
            <label>РЭС:</label>
            <select 
              value={filters.resId}
              onChange={(e) => handleFilterChange('resId', e.target.value)}
              disabled={user.role !== 'admin'} // Не-админы видят только свой РЭС
            >
              <option value="">Все РЭС</option>
              {resList.map(res => (
                <option key={res.id} value={res.id}>{res.name}</option>
              ))}
            </select>
          </div>
          
          {activeTab === 'uploads' && (
            <>
              <div className="filter-group">
                <label>ТП:</label>
                <input 
                  type="text"
                  value={filters.tpName}
                  onChange={(e) => handleFilterChange('tpName', e.target.value)}
                  placeholder="Поиск по ТП"
                />
              </div>
              
              <div className="filter-group">
                <label>Тип файла:</label>
                <select 
                  value={filters.fileType}
                  onChange={(e) => handleFilterChange('fileType', e.target.value)}
                >
                  <option value="">Все типы</option>
                  <option value="rim_single">Счетчик РИМ</option>
                  <option value="nartis">Счетчик Нартис</option>
                  <option value="energomera">Счетчик Энергомера</option>
                </select>
              </div>
              
              <div className="filter-group">
                <label>Статус:</label>
                <select 
                  value={filters.status}
                  onChange={(e) => handleFilterChange('status', e.target.value)}
                >
                  <option value="">Все статусы</option>
                  <option value="success">Успешно</option>
                  <option value="duplicate">Дубликат</option>
                  <option value="wrong_period">Неверный период</option>
                  <option value="error">Ошибка</option>
                </select>
              </div>
            </>
          )}
          
          {activeTab === 'checks' && (
  <>
    {/* ДОБАВИТЬ поле ТП для истории проверок */}
    <div className="filter-group">
      <label>ТП:</label>
      <input 
        type="text"
        value={filters.tpName}
        onChange={(e) => handleFilterChange('tpName', e.target.value)}
        placeholder="Поиск по ТП"
      />
    </div>

    <div className="filter-group">
      <label>Статус:</label>
      <select 
        value={filters.status}
        onChange={(e) => handleFilterChange('status', e.target.value)}
      >
        <option value="">Все статусы</option>
        <option value="awaiting_work">Ожидает мероприятий</option>
        <option value="awaiting_recheck">Ожидает перепроверки</option>
        <option value="completed">Завершено</option>
      </select>
    </div>
  </>
)}
</div>    
        
        <div className="filter-row">
          <div className="filter-group">
            <label>Период с:</label>
            <input 
              type="date"
              value={filters.dateFrom}
              onChange={(e) => handleFilterChange('dateFrom', e.target.value)}
            />
          </div>
          
          <div className="filter-group">
            <label>по:</label>
            <input 
              type="date"
              value={filters.dateTo}
              onChange={(e) => handleFilterChange('dateTo', e.target.value)}
            />
          </div>
          
          <button className="export-btn" onClick={exportToExcel}>
            <IconChart className="ico" /> Экспорт в Excel
          </button>
        </div>
      </div>
      
      <div className="history-content">
        {loading ? (
          <LoadingSpinner 
            type="dots"
            message="Загрузка истории системы..." 
            submessage="Обрабатываем записи за выбранный период" 
          />
        ) : (
          <>
            {activeTab === 'uploads' && (
              <div className="history-table">
                <table>
                  <thead>
                    <tr>
                      <th>Дата загрузки</th>
                      <th>РЭС</th>
                      <th>ТП</th>
                      <th>ВЛ</th>
                      <th>ПУ №</th>
                      <th>Загрузил</th>
                      <th>Тип</th>
                      <th>Статус</th>
                      <th>Ошибка</th>
                    </tr>
                  </thead>
                  <tbody>
                    {uploads.map((upload, idx) => (
                      <tr key={idx} className={upload.uploadStatus}>
                        <td>{new Date(upload.uploadedAt).toLocaleString('ru-RU')}</td>
                        <td>{upload.resName || '—'}</td>
                        <td>{upload.tpName || '—'}</td>
                        <td>{upload.vlName || '—'}</td>
                        <td><strong>{upload.puNumber}</strong></td>
                        <td>{upload.User?.fio || 'Неизвестно'}</td>
                        <td>{upload.fileType}</td>
                        <td>
                          <span className={`status-badge status-${upload.uploadStatus}`}>
                            {upload.uploadStatus === 'success' ? <IconCheck className="ico" style={{color:'var(--green)'}} /> :
                             upload.uploadStatus === 'duplicate' ? <IconRefresh className="ico" /> :
                             upload.uploadStatus === 'wrong_period' ? <IconCalendar className="ico" /> : <IconX className="ico" style={{color:'var(--red)'}} />}
                          </span>
                        </td>
                        <td className="error-cell">
                          {upload.hasErrors && (
                            <details>
                              <summary>Показать</summary>
                              <pre>{upload.errorSummary}</pre>
                            </details>
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
            
            {activeTab === 'checks' && (
              <div className="history-table">
                <table>
                  <thead>
                    <tr>
                      <th>РЭС</th>
                      <th>ТП</th>
                      <th>ВЛ</th>
                      <th>ПУ №</th>
                      <th>Дата ошибки</th>
                      <th>Ошибка</th>
                      <th>Мероприятия</th>
                      <th>Перепроверка</th>
                      <th>Статус</th>
                    </tr>
                  </thead>
                  <tbody>
                    {checks.map((check, idx) => (
                      <tr key={idx}>
                        <td>{check.ResUnit?.name}</td>
                        <td>{check.tpName}</td>
                        <td>{check.vlName}</td>
                        <td><strong>{check.puNumber}</strong></td>
                        <td>{new Date(check.initialCheckDate).toLocaleDateString('ru-RU')}</td>
                        <td className="error-cell">
                          <details>
                            <summary>Показать</summary>
                            <pre>{check.initialError}</pre>
                          </details>
                        </td>
                        <td>
                          {check.workCompletedDate ? (
                            <>
                              <div>{new Date(check.workCompletedDate).toLocaleDateString('ru-RU')}</div>
                              <small>{check.resComment}</small>
                            </>
                          ) : '—'}
                        </td>
                        <td>
                          {check.recheckDate ? (
                            <span className={check.recheckResult === 'ok' ? 'status-ok' : 'status-error'}>
                              {check.recheckResult === 'ok' ? <IconCheck className="ico" style={{color:'var(--green)'}} /> : <IconX className="ico" style={{color:'var(--red)'}} />}
                              {' ' + new Date(check.recheckDate).toLocaleDateString('ru-RU')}
                            </span>
                          ) : '—'}
                        </td>
                        <td>
                          <span className={`status-badge status-${check.status}`}>
                            {check.status === 'awaiting_work' ? 'Ожидает работ' :
                             check.status === 'awaiting_recheck' ? 'Ожидает проверки' :
                             'Завершено'}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
            
            {/* Пагинация */}
            {totalPages > 1 && (
              <div className="pagination">
                <button 
                  onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                  disabled={currentPage === 1}
                >
                  <IconArrowLeft className="ico" /> Назад
                </button>
                <span>Страница {currentPage} из {totalPages}</span>
                <button 
                  onClick={() => setCurrentPage(p => Math.min(totalPages, p + 1))}
                  disabled={currentPage === totalPages}
                >
                  Вперед <IconArrowRight className="ico" />
                </button>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}

function Analytics() {
  const [analytics, setAnalytics] = useState([]);
  const [totals, setTotals] = useState({});
  const [vlWorkload, setVlWorkload] = useState([]);
  const [loading, setLoading] = useState(false);
  const [loadingDetailed, setLoadingDetailed] = useState(false);
  const [dateFrom, setDateFrom] = useState(
    new Date(Date.now() - 30 * 24 * 60 * 60 * 1000).toISOString().split('T')[0]
  );
  const [dateTo, setDateTo] = useState(new Date().toISOString().split('T')[0]);
  const { user } = useContext(AuthContext);
  
  useEffect(() => {
    loadAnalytics();
    loadVlWorkload();
  }, [dateFrom, dateTo]);

  const loadVlWorkload = async () => {
    try {
      const response = await api.get('/api/reports/detailed', {
        params: { type: 'vl_workload' }
      });
      setVlWorkload(response.data);
    } catch (error) {
      console.error('Error loading VL workload:', error);
      setVlWorkload([]);
    }
  };
  
  const loadAnalytics = async () => {
    setLoading(true);
    try {
      const response = await api.get('/api/analytics/summary', {
        params: { dateFrom, dateTo }
      });
      setAnalytics(response.data.analytics);
      setTotals(response.data.totals);
    } catch (error) {
      console.error('Error loading analytics:', error);
    } finally {
      setLoading(false);
    }
  };
  
  const exportToExcel = () => {
    const data = analytics.map(row => ({
      'РЭС': row.resName,
      'Всего ТП': row.tpCount,
      'Всего ВЛ': row.vlCount,
      '% охвата ВЛ': row.vlCoveragePercent + '%',
      'Всего ПУ': row.totalPuCount,
      'Проверено ПУ': row.uniqueCheckedPuCount,
      '% охвата ПУ': row.puCoveragePercent + '%',
      'Секций с перегрузом': row.overloadSections ?? 0,
      'Активных случаев Pном': row.activeOverloadCases ?? 0
    }));
    
    // Добавляем итоги
    if (user.role === 'admin') {
      data.push({
        'РЭС': 'ИТОГО',
        'Всего ТП': totals.tpCount,
        'Всего ВЛ': totals.vlCount,
        '% охвата ВЛ': totals.vlCoveragePercent + '%',
        'Всего ПУ': totals.totalPuCount,
        'Проверено ПУ': totals.uniqueCheckedPuCount,
        '% охвата ПУ': totals.puCoveragePercent + '%',
        'Секций с перегрузом': totals.overloadSections ?? 0,
        'Активных случаев Pном': totals.activeOverloadCases ?? 0
      });
    }
    
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Аналитика');
    
    // УЛУЧШЕНО: добавляем РЭС в имя файла для не-админов
    const fileName = user.role !== 'admin' 
      ? `Аналитика_${user.resName}_${new Date().toLocaleDateString('ru-RU').split('.').join('-')}.xlsx`
      : `Аналитика_${new Date().toLocaleDateString('ru-RU').split('.').join('-')}.xlsx`;
    
    XLSX.writeFile(wb, fileName);
  };
  
  const exportDetailedReport = async () => {
    setLoadingDetailed(true);
    
    try {
      const response = await api.get('/api/analytics/detailed', {
        params: { dateFrom, dateTo }
      });
      
      const detailedData = response.data.data;
      
      if (detailedData.length === 0) {
        alert('Нет данных для детального отчета');
        return;
      }
      
      const excelData = detailedData.map(row => ({
        'РЭС': row.resName,
        'ТП': row.tpName,
        'ВЛ': row.vlName,
        'Статус ВЛ': row.vlStatus,
        'Проверено/Всего ПУ': `${row.checkedPuCount}/${row.totalPuCount}`,
        
        'ПУ Начало': row.startPu.number,
        'Статус ПУ начала': row.startPu.status,
        'Ошибка ПУ начала': row.startPu.error,
        'Кто проверил ПУ начало': row.startPu.uploadedBy,
        'Дата проверки ПУ начала': row.startPu.uploadDate,
        
        'ПУ Середина': row.middlePu.number,
        'Статус ПУ середины': row.middlePu.status,
        'Ошибка ПУ середины': row.middlePu.error,
        'Кто проверил ПУ середину': row.middlePu.uploadedBy,
        'Дата проверки ПУ середины': row.middlePu.uploadDate,
        
        'ПУ Конец': row.endPu.number,
        'Статус ПУ конца': row.endPu.status,
        'Ошибка ПУ конца': row.endPu.error,
        'Кто проверил ПУ конец': row.endPu.uploadedBy,
        'Дата проверки ПУ конца': row.endPu.uploadDate
      }));
      
      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(excelData);
      
      ws['!cols'] = [
        { wch: 20 }, { wch: 15 }, { wch: 15 }, { wch: 20 }, { wch: 15 },
        { wch: 15 }, { wch: 15 }, { wch: 50 }, { wch: 20 }, { wch: 18 },
        { wch: 15 }, { wch: 15 }, { wch: 50 }, { wch: 20 }, { wch: 18 },
        { wch: 15 }, { wch: 15 }, { wch: 50 }, { wch: 20 }, { wch: 18 }
      ];
      
      XLSX.utils.book_append_sheet(wb, ws, 'Детальный отчет');
      
      // УЛУЧШЕНО: добавляем РЭС в имя файла для не-админов
      const fileName = user.role !== 'admin'
        ? `Детальный_отчет_${user.resName}_${new Date().toLocaleDateString('ru-RU').split('.').join('-')}.xlsx`
        : `Детальный_отчет_${new Date().toLocaleDateString('ru-RU').split('.').join('-')}.xlsx`;
      
      XLSX.writeFile(wb, fileName);
      
      alert(`Детальный отчет успешно выгружен!\n\nВсего строк: ${detailedData.length}\nФайл: ${fileName}`);
      
    } catch (error) {
      console.error('Error exporting detailed report:', error);
      alert('Ошибка при выгрузке детального отчета: ' + error.message);
    } finally {
      setLoadingDetailed(false);
    }
  };
  
  if (loading) return <LoadingSpinner type="pulse" message="Загрузка аналитики..." submessage="Подсчитываем статистику" />;
  
  return (
    <div className="analytics-container">
      <PageHeader icon={<IconChart size={24} />} title="Анализ напряжения" />
      
      {/* НОВОЕ: Индикатор для не-админов */}
      {user.role !== 'admin' && (
        <div className="res-indicator">
          <span><IconMapPin className="ico" /> Показаны данные для: <strong>{user.resName}</strong></span>
        </div>
      )}
      
      <div className="analytics-controls">
        <div className="control-group">
          <label>Период с:</label>
          <input 
            type="date" 
            value={dateFrom}
            onChange={(e) => setDateFrom(e.target.value)}
          />
        </div>
        <div className="control-group">
          <label>по:</label>
          <input 
            type="date" 
            value={dateTo}
            onChange={(e) => setDateTo(e.target.value)}
          />
        </div>
        
        <button onClick={exportToExcel} className="export-btn">
          Экспорт сводного отчета
        </button>
        <button 
          onClick={exportDetailedReport} 
          className="export-btn detailed"
          disabled={loadingDetailed}
        >
          {loadingDetailed ? (
            <>
                            Формирование...
            </>
          ) : (
            <>
              <span><IconClipboard className="ico" /></span>
              Экспорт детального отчета
            </>
          )}
        </button>
      </div>
      
      <div className="analytics-table">
        <table>
          <thead>
            <tr>
              <th>РЭС</th>
              <th>Всего ТП</th>
              <th>Всего ВЛ</th>
              <th>% охвата ВЛ</th>
              <th>Всего ПУ</th>
              <th>Проверено ПУ</th>
              <th>% охвата ПУ</th>
              <th>Секций с перегрузом</th>
              <th>Активных случаев Pном</th>
            </tr>
          </thead>
          <tbody>
            {analytics.map(row => (
              <tr key={row.resId}>
                <td>{row.resName}</td>
                <td>{row.tpCount}</td>
                <td>{row.vlCount}</td>
                <td>
                  <div className="progress-cell">
                    <div className="progress-bar-small">
                      <div 
                        className="progress-fill-small"
                        style={{ width: `${row.vlCoveragePercent}%` }}
                      />
                    </div>
                    <span>{row.vlCoveragePercent}%</span>
                  </div>
                </td>
                <td>{row.totalPuCount}</td>
                <td><strong>{row.uniqueCheckedPuCount}</strong></td>
                <td>
                  <div className="progress-cell">
                    <div className="progress-bar-small">
                      <div 
                        className="progress-fill-small"
                        style={{ width: `${row.puCoveragePercent}%` }}
                      />
                    </div>
                    <span>{row.puCoveragePercent}%</span>
                  </div>
                </td>
                <td>{row.overloadSections ?? 0}</td>
                <td>{row.activeOverloadCases ?? 0}</td>
              </tr>
            ))}
            {user.role === 'admin' && (
              <tr className="totals-row">
                <td><strong>ИТОГО</strong></td>
                <td><strong>{totals.tpCount}</strong></td>
                <td><strong>{totals.vlCount}</strong></td>
                <td>
                  <strong>{totals.vlCoveragePercent}%</strong>
                </td>
                <td><strong>{totals.totalPuCount}</strong></td>
                <td><strong>{totals.uniqueCheckedPuCount}</strong></td>
                <td>
                  <strong>{totals.puCoveragePercent}%</strong>
                </td>
                <td><strong>{totals.overloadSections ?? 0}</strong></td>
                <td><strong>{totals.activeOverloadCases ?? 0}</strong></td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      
      {/* ВЛ в работе у РЭС */}
      <h2 style={{ marginTop: '32px' }}><span className="svg-frame"><IconZap size={24} /></span> ВЛ в работе у РЭС</h2>
      <p className="info-hint" style={{ marginTop: '18px', marginBottom: '12px' }}>
        <IconInfo className="ico" style={{color:'var(--blue)'}} /> Текущее состояние: количество ВЛ с нерешёнными проблемами на данный момент
      </p>
      
      <div className="analytics-table">
        <table>
          <thead>
            <tr>
              <th>РЭС</th>
              <th>Всего ВЛ</th>
              <th>ВЛ с проблемами</th>
              <th>Без проблем</th>
              <th>% проблемных</th>
            </tr>
          </thead>
          <tbody>
            {vlWorkload.map((item, idx) => (
              <tr key={idx} className={item.isTotal ? 'totals-row' : ''}>
                <td>{item.isTotal ? <strong>{item.resName}</strong> : item.resName}</td>
                <td>{item.isTotal ? <strong>{item.totalVl}</strong> : item.totalVl}</td>
                <td style={{ color: item.problemVl > 0 ? '#e53e3e' : 'inherit', fontWeight: item.problemVl > 0 ? 'bold' : 'normal' }}>
                  {item.isTotal ? <strong>{item.problemVl}</strong> : item.problemVl}
                </td>
                <td style={{ color: '#38a169' }}>
                  {item.isTotal ? <strong>{item.okVl}</strong> : item.okVl}
                </td>
                <td>
                  {item.isTotal ? (
                    <strong>{item.problemPercent}%</strong>
                  ) : (
                    <div className="progress-cell">
                      <div className="progress-bar-small">
                        <div 
                          className="progress-fill-small"
                          style={{ 
                            width: `${item.problemPercent}%`,
                            backgroundColor: item.problemPercent > 50 ? '#e53e3e' : item.problemPercent > 20 ? '#ed8936' : '#38a169'
                          }}
                        />
                      </div>
                      <span>{item.problemPercent}%</span>
                    </div>
                  )}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ═══ Раздел «Анализ мощности»: перенесённый блок «ТП с перегрузом» + матрица
// помесячных пиков (секции × месяцы). Роли admin / res_responsible. ═══
function PowerAnalysis({ selectedRes }) {
  const { user } = useContext(AuthContext);
  const [overloadRows, setOverloadRows] = useState([]);
  const [monthly, setMonthly] = useState({ months: [], rows: [] });
  const [loading, setLoading] = useState(false);
  const [preset, setPreset] = useState('12'); // '6' | '12' | 'custom'
  const ymNow = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`; };
  const ymShift = (months) => { const d = new Date(); d.setMonth(d.getMonth() - months); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`; };
  const [from, setFrom] = useState(ymShift(11));
  const [to, setTo] = useState(ymNow());

  const fmt1 = (v) => (v == null ? '—' : Number(v).toFixed(1));
  const loadCls = (pct) => pct == null ? '' : pct > 100 ? 'pm-red' : pct >= 85 ? 'pm-amber' : 'pm-green';
  const monLabel = (ym) => { const [y, m] = ym.split('-'); return `${m}.${y}`; };

  const range = () => {
    if (preset === '6') return { from: ymShift(5), to: ymNow() };
    if (preset === '12') return { from: ymShift(11), to: ymNow() };
    return { from, to };
  };

  const load = async () => {
    setLoading(true);
    try {
      const { from: f, to: t } = range();
      const params = { from: f, to: t };
      if (user.role === 'admin' && selectedRes) params.resId = selectedRes;
      const [ov, mp] = await Promise.all([
        api.get('/api/reports/overload', { params: user.role === 'admin' && selectedRes ? { resId: selectedRes } : {} }),
        api.get('/api/power/monthly-peaks', { params })
      ]);
      const rows = (ov.data || []).filter(r => r.ratioPct != null && r.ratioPct >= 100).sort((a, b) => (b.ratioPct || 0) - (a.ratioPct || 0));
      setOverloadRows(rows);
      setMonthly(mp.data || { months: [], rows: [] });
    } catch (e) {
      console.error('PowerAnalysis load error:', e);
      setOverloadRows([]);
      setMonthly({ months: [], rows: [] });
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); /* eslint-disable-next-line */ }, [preset, from, to, selectedRes]);

  const exportMatrix = () => {
    const { months, rows } = monthly;
    if (!rows.length) { alert('Нет данных для экспорта'); return; }
    const data = rows.map(r => {
      const base = {
        'РЭС': r.resName || '', 'ТП': r.tpName || '', 'СШ': `СШ-${toRoman(r.sectionNumber)}`,
        'Sном, кВА': r.tnKva != null ? r.tnKva : '—', 'Лимит, кВт': r.limitKw != null ? r.limitKw : '—'
      };
      months.forEach(m => {
        const p = r.peaks[m];
        base[monLabel(m)] = p ? p.peakKw : '';
      });
      return base;
    });
    // Цвет ячеек месяцев — по % загрузки соответствующего пика.
    const monthColor = (ri, colName) => {
      const m = months.find(mm => monLabel(mm) === colName);
      if (!m) return null;
      const p = monthly.rows[ri]?.peaks?.[m];
      if (!p || p.ratioPct == null) return null;
      return p.ratioPct > 100 ? XLS_COLORS.red : p.ratioPct >= 85 ? XLS_COLORS.amber : XLS_COLORS.green;
    };
    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [{ wch: 18 }, { wch: 15 }, { wch: 8 }, { wch: 10 }, { wch: 11 }, ...months.map(() => ({ wch: 10 }))];
    styleExportSheet(ws, {}, monthColor);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Помесячные пики');
    XLSX.writeFile(wb, `Помесячные_пики_${new Date().toLocaleDateString('ru-RU').split('.').join('-')}.xlsx`);
  };

  return (
    <div className="analytics">
      <PageHeader icon={<IconChart size={24} />} title="Анализ мощности" />

      {/* Перенесено из «Аналитики»: ТП с перегрузом (последний пик ≥ лимита). */}
      <h2 style={{ marginTop: '8px', fontSize: 20 }}>ТП с перегрузом</h2>
      <p className="info-hint" style={{ marginTop: '18px', marginBottom: '12px' }}>
        <IconInfo className="ico" style={{ color: 'var(--blue)' }} /> Секции шин, где последний пик мощности достиг или превысил лимит Sном·cosφ
      </p>
      <div className="analytics-table">
        {overloadRows.length === 0 ? (
          <div className="no-data" style={{ padding: '16px' }}>
            <span className="svg-frame" style={{ marginRight: 8 }}><IconCheck size={22} /></span>
            Секций с перегрузом нет
          </div>
        ) : (
          <table>
            <thead>
              <tr>
                <th>РЭС</th><th>ТП</th><th>СШ</th><th>Sном, кВА</th><th>Лимит, кВт</th>
                <th>Пик, кВт</th><th>%</th><th>Статус случая</th><th>Циклы</th>
              </tr>
            </thead>
            <tbody>
              {overloadRows.map((r, idx) => (
                <tr key={idx}>
                  <td>{r.resName}</td>
                  <td>{r.tpName}</td>
                  <td>СШ-{toRoman(r.sectionNumber)}</td>
                  <td>{r.tnKva}</td>
                  <td>{r.limitKw}</td>
                  <td style={{ color: 'var(--red)', fontWeight: 700 }}>{r.lastPeakKw ?? '—'}</td>
                  <td style={{ color: 'var(--red)', fontWeight: 700 }}>{r.ratioPct != null ? r.ratioPct + '%' : '—'}</td>
                  <td>{r.caseStage}</td>
                  <td>{r.cycles || 0}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </div>

      {/* Матрица помесячных пиков: строки — секции, колонки — месяцы периода. */}
      <h2 style={{ marginTop: '32px', fontSize: 20 }}>Помесячные пики мощности</h2>
      <div className="pm-controls">
        <div className="pm-presets">
          <button className={`pm-preset ${preset === '6' ? 'active' : ''}`} onClick={() => setPreset('6')}>6 мес</button>
          <button className={`pm-preset ${preset === '12' ? 'active' : ''}`} onClick={() => setPreset('12')}>12 мес</button>
          <button className={`pm-preset ${preset === 'custom' ? 'active' : ''}`} onClick={() => setPreset('custom')}>Период</button>
        </div>
        {preset === 'custom' && (
          <div className="pm-custom">
            <input type="month" value={from} onChange={(e) => setFrom(e.target.value)} />
            <span>—</span>
            <input type="month" value={to} onChange={(e) => setTo(e.target.value)} />
          </div>
        )}
        <div className="pm-actions">
          <button className="pm-btn pm-btn--refresh" onClick={load} disabled={loading}>
            <IconRefresh className="ico" /> {loading ? 'Обновление...' : 'Обновить'}
          </button>
          <button className="pm-btn pm-btn--excel" onClick={exportMatrix}>
            <IconDownload className="ico" /> Выгрузка в Excel
          </button>
        </div>
      </div>

      {loading ? (
        <div className="loading"><RossetiLoader /></div>
      ) : monthly.rows.length === 0 ? (
        <div className="no-data" style={{ padding: '16px' }}>Нет секций для отображения</div>
      ) : (
        <div className="pm-table-wrap">
          <table className="pm-matrix">
            <thead>
              <tr>
                <th className="pm-sticky">ТП · СШ</th>
                {user.role === 'admin' && <th>РЭС</th>}
                <th>Лимит, кВт</th>
                {monthly.months.map(m => <th key={m}>{monLabel(m)}</th>)}
              </tr>
            </thead>
            <tbody>
              {monthly.rows.map(r => (
                <tr key={r.sectionId}>
                  <td className="pm-sticky"><strong>{r.tpName}</strong> · СШ-{toRoman(r.sectionNumber)}</td>
                  {user.role === 'admin' && <td>{r.resName}</td>}
                  <td>{r.limitKw != null ? fmt1(r.limitKw) : '—'}</td>
                  {monthly.months.map(m => {
                    const p = r.peaks[m];
                    if (!p) return <td key={m} className="pm-empty">нет данных</td>;
                    return (
                      <td key={m} className={loadCls(p.ratioPct)}>
                        <span className="pm-kw">{fmt1(p.peakKw)}</span>
                        {p.ratioPct != null && <span className="pm-pct"> · {p.ratioPct}%</span>}
                      </td>
                    );
                  })}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function DatabaseMaintenance() {
  const { user } = useContext(AuthContext);
  const canDbTools = hasPerm(user, 'db_tools');       // проверка/бэкап/восстановление/очистка
  const canPurge = hasPerm(user, 'history_purge');    // очистка истории до даты
  const [healthCheck, setHealthCheck] = useState(null);
  const [loading, setLoading] = useState(false);
  const [showCleanupModal, setShowCleanupModal] = useState(false);
  const [cleanupType, setCleanupType] = useState('');
  const [cleanupPassword, setCleanupPassword] = useState('');
  const [cleaning, setCleaning] = useState(false);
  const [backuping, setBackuping] = useState(false);
  const [restoring, setRestoring] = useState(false);
  // Очистка истории до даты
  const [purgeBefore, setPurgeBefore] = useState('');
  const [purgePreview, setPurgePreview] = useState(null);
  const [purgePreviewing, setPurgePreviewing] = useState(false);
  const [purgeConfirm, setPurgeConfirm] = useState(false);
  const [purgePassword, setPurgePassword] = useState('');
  const [purging, setPurging] = useState(false);
  const [purgeResult, setPurgeResult] = useState(null);

  const handlePurgePreview = async () => {
    if (!purgeBefore) { alert('Выберите дату'); return; }
    setPurgePreviewing(true); setPurgeResult(null);
    try {
      const { data } = await api.post('/api/admin/purge-preview', { before: purgeBefore });
      setPurgePreview(data);
    } catch (e) {
      alert('Ошибка предпросмотра: ' + (e.response?.data?.error || e.message));
      setPurgePreview(null);
    } finally { setPurgePreviewing(false); }
  };

  const handlePurge = async () => {
    setPurging(true);
    try {
      const { data } = await api.post('/api/admin/purge', { before: purgeBefore, password: purgePassword });
      setPurgeResult(data);
      setPurgePreview(null);
      setPurgeConfirm(false);
      setPurgePassword('');
    } catch (e) {
      alert('Ошибка очистки: ' + (e.response?.data?.error || e.message));
    } finally { setPurging(false); }
  };

    const runHealthCheck = async () => {
    setLoading(true);
    try {
      const response = await api.get('/api/admin/database-health');
  
  // ДОБАВЬ ЭТО
    console.log('=== DATABASE HEALTH CHECK RESPONSE ===');
    console.log('Full response:', response.data);
    console.log('Stats:', response.data.stats);
    console.log('staleNotifications:', response.data.stats?.staleNotifications);
    console.log('missingNotifications:', response.data.stats?.missingNotifications);
    console.log('======================================');
    
    setHealthCheck(response.data);
  } catch (error) {
    alert('Ошибка проверки БД: ' + error.message);
  } finally {
    setLoading(false);
  }
};
  
  const handleCleanup = async () => {
    setCleaning(true);
    try {
      const response = await api.post('/api/admin/database-cleanup', {
        cleanupType,
        password: cleanupPassword
      });
      
      alert(`Очистка завершена!\n\nУдалено записей: ${response.data.cleaned}`);
      setShowCleanupModal(false);
      setCleanupPassword('');
      
      // Запускаем проверку заново
      runHealthCheck();
      
    } catch (error) {
      alert('Ошибка очистки: ' + (error.response?.data?.error || error.message));
    } finally {
      setCleaning(false);
    }
  };

  // Скачать полный бэкап БД (JSON-файл)
  const handleBackup = async () => {
    setBackuping(true);
    try {
      const response = await api.get('/api/admin/backup', { responseType: 'blob' });
      const url = window.URL.createObjectURL(new Blob([response.data], { type: 'application/json' }));
      const a = document.createElement('a');
      a.href = url;
      a.download = `res-backup-${new Date().toISOString().slice(0, 10)}.json`;
      document.body.appendChild(a);
      a.click();
      a.remove();
      window.URL.revokeObjectURL(url);
    } catch (error) {
      alert('Ошибка бэкапа: ' + (error.response?.data?.error || error.message));
    } finally {
      setBackuping(false);
    }
  };

  // Восстановить БД из файла бэкапа. Через fetch + FormData (multipart), чтобы
  // не упереться в лимит express.json и не спорить с Content-Type axios-инстанса.
  const handleRestore = async (e) => {
    const file = e.target.files?.[0];
    e.target.value = ''; // сброс — чтобы можно было выбрать тот же файл повторно
    if (!file) return;
    if (!window.confirm('ВНИМАНИЕ! Восстановление ПОЛНОСТЬЮ заменит все данные в базе (текущие будут удалены). Продолжить?')) return;
    setRestoring(true);
    try {
      const fd = new FormData();
      // Сжимаем файл gzip перед отправкой (прокси Amvera режет большие тела —
      // HTTP 413). JSON жмётся в 10-20 раз. Сервер распознаёт gzip по магическим
      // байтам. Если CompressionStream недоступен (старый браузер) — шлём как есть.
      if (window.CompressionStream) {
        const gzBlob = await new Response(
          file.stream().pipeThrough(new CompressionStream('gzip'))
        ).blob();
        fd.append('file', gzBlob, file.name + '.gz');
      } else {
        fd.append('file', file);
      }
      fd.append('confirm', 'true');
      const token = localStorage.getItem('token');
      const resp = await fetch(`${API_URL}/api/admin/restore`, {
        method: 'POST',
        headers: { Authorization: `Bearer ${token}` },
        body: fd,
      });
      const r = await resp.json().catch(() => ({}));
      if (!resp.ok) throw new Error(r.error || `HTTP ${resp.status}`);
      let msg = `Восстановление завершено.\nВставлено записей: ${r.inserted}`;
      if (r.errorsCount) msg += `\n\nОшибок: ${r.errorsCount}\n• ` + (r.errors || []).join('\n• ');
      alert(msg);
    } catch (error) {
      alert('Ошибка восстановления: ' + error.message);
    } finally {
      setRestoring(false);
    }
  };

  const getSeverityColor = (severity) => {
    switch(severity) {
      case 'error': return '#ff4d4f';
      case 'warning': return '#faad14';
      case 'info': return '#1890ff';
      default: return '#52c41a';
    }
  };
  
  const getSeverityIcon = (severity) => {
    switch(severity) {
      case 'error': return <IconAlertCircle className="ico" style={{color:'var(--red)'}} />;
      case 'warning': return <IconAlertTriangle className="ico" style={{color:'var(--amber)'}} />;
      case 'info': return <IconInfo className="ico" style={{color:'var(--blue)'}} />;
      default: return <IconCheck className="ico" style={{color:'var(--green)'}} />;
    }
  };
  
  const getCleanupDescription = (type) => {
    const descriptions = {
      'orphaned_pu_status': {
        title: 'Статусы ПУ без структуры',
        desc: 'Удалить статусы приборов учета, которые не привязаны к структуре сети',
        icon: <IconPlug className="ico" />
      },
      'duplicate_pu_statuses': {
        title: 'Дубликаты статусов',
        desc: 'Удалить дублирующиеся статусы ПУ, оставив только последние',
        icon: <IconClipboard className="ico" />
      },
      'old_unread_notifications': {
        title: 'Старые уведомления',
        desc: 'Удалить уведомления старше одного года',
        icon: <IconCalendar className="ico" />
      },
      'orphaned_notifications': {
        title: 'Уведомления без связей',
        desc: 'Удалить уведомления, ссылающиеся на несуществующие объекты',
        icon: <IconLink className="ico" />
      },
      'checks_without_res': {
        title: 'Проверки без РЭС',
        desc: 'Удалить записи истории проверок без привязки к РЭС',
        icon: <IconMapPin className="ico" />
      },
      'broken_file_references': {
        title: 'Битые ссылки на файлы',
        desc: 'Очистить некорректные ссылки на файлы в истории проверок',
        icon: <IconPaperclip className="ico" />
      },
      'stale_problem_vl': {
        title: 'Старые проблемные ВЛ',
        desc: 'Закрыть проблемные ВЛ без активности более 90 дней',
        icon: <IconZap className="ico" style={{color:'var(--amber)'}} />
      },
      'irrelevant_problem_vl': {
        title: 'Неактуальные проблемные ВЛ',
        desc: 'Закрыть проблемные ВЛ, у которых ПУ уже проверен без ошибок или удалён из структуры',
        icon: <IconAlertTriangle className="ico" style={{color:'var(--amber)'}} />
      },
        'stale_notifications': {
      title: 'Неактуальные уведомления',
      desc: 'Удалить уведомления для ПУ, которые уже проверены без ошибок',
      icon: <IconBell className="ico" />
    },
        'missing_notifications': {
      title: 'Отсутствующие уведомления',
      desc: 'Создать уведомления для ПУ с ошибками, у которых нет уведомлений',
      icon: <IconMegaphone className="ico" />
    }
    };
    return descriptions[type] || { title: 'Неизвестная операция', desc: '', icon: <IconHelp className="ico" /> };
  };
  
  return (
    <div className="database-maintenance">
      {/* Проверка целостности (db_tools) */}
      {canDbTools && (
      <div className="db-header">
        <div className="db-header-content">
          <div className="db-header-icon"><IconWrench className="ico" /></div>
          <div className="db-header-text">
            <h3>Проверка целостности базы данных</h3>
            <p>Диагностика и устранение проблем в структуре данных</p>
          </div>
        </div>
        <button
          className="btn-check-db"
          onClick={runHealthCheck}
          disabled={loading}
        >
          {loading ? (
            <>
                            Проверка...
            </>
          ) : (
            <>
              <span><IconSearch className="ico" /></span>
              Запустить проверку
            </>
          )}
        </button>
      </div>
      )}

      {/* Резервная копия базы: скачать бэкап / восстановить из файла (db_tools) */}
      {canDbTools && (
      <div className="db-header" style={{ marginTop: 12 }}>
        <div className="db-header-content">
          <div className="db-header-icon"><IconDatabase className="ico" /></div>
          <div className="db-header-text">
            <h3>Резервная копия базы</h3>
            <p>Скачать полный бэкап (JSON) или восстановить данные из файла. Восстановление заменяет все данные.</p>
          </div>
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
          <button className="btn-check-db" onClick={handleBackup} disabled={backuping || restoring}>
            {backuping ? 'Скачивание...' : 'Скачать бэкап'}
          </button>
          <label className="btn-check-db" style={{ cursor: restoring ? 'default' : 'pointer', opacity: restoring ? 0.6 : 1 }}>
            {restoring ? 'Восстановление...' : 'Восстановить из файла'}
            <input type="file" accept=".json,application/json" onChange={handleRestore} disabled={restoring || backuping} style={{ display: 'none' }} />
          </label>
        </div>
      </div>
      )}

      {/* Очистка истории до даты (history_purge) */}
      {canPurge && (
      <div className="db-header" style={{ marginTop: 12 }}>
        <div className="db-header-content">
          <div className="db-header-icon"><IconBroom className="ico" /></div>
          <div className="db-header-text">
            <h3>Очистка истории</h3>
            <p>Удаляет истории загрузок/проверок, уведомления и связанные файлы <strong>старше выбранной даты</strong>. Структура сети, тех-учёт, помесячные пики, случаи перегруза и пользователи — сохраняются.</p>
          </div>
        </div>
        <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap', alignItems: 'center' }}>
          <input type="date" value={purgeBefore} className="purge-date"
            onChange={(e) => { setPurgeBefore(e.target.value); setPurgePreview(null); setPurgeResult(null); }} />
          <button className="btn-check-db" onClick={handlePurgePreview} disabled={purgePreviewing || !purgeBefore}>
            {purgePreviewing ? 'Подсчёт...' : 'Показать что будет удалено'}
          </button>
        </div>
      </div>
      )}

      {purgePreview && (
        <div className="purge-preview">
          <div className="purge-preview-title">Будет удалено (старше {purgeBefore}):</div>
          <div className="purge-preview-grid">
            <div><span>Истории загрузок</span><b>{purgePreview.uploadHistory}</b></div>
            <div><span>Истории по ПУ</span><b>{purgePreview.puUploadHistory}</b></div>
            <div><span>Истории проверок</span><b>{purgePreview.checkHistory}</b></div>
            <div><span>Уведомления</span><b>{purgePreview.notifications}</b></div>
            <div><span>Файлы (Cloudinary)</span><b>{purgePreview.cloudinaryFiles}</b></div>
          </div>
          <div className="purge-preview-actions">
            <input type="password" placeholder="Пароль удаления" value={purgePassword}
              onChange={(e) => setPurgePassword(e.target.value)} className="purge-pass" />
            <button className="delete-selected-btn" disabled={!purgePassword}
              onClick={() => setPurgeConfirm(true)}>Удалить безвозвратно</button>
          </div>
        </div>
      )}

      {purgeResult && (
        <div className="purge-result">
          <IconCheck className="ico" style={{ color: 'var(--green)' }} /> Очистка завершена. Удалено:
          {' '}уведомлений {purgeResult.deleted.notifications ?? 0}, проверок {purgeResult.deleted.checkHistory ?? 0},
          {' '}загрузок {purgeResult.deleted.uploadHistory ?? 0}, по ПУ {purgeResult.deleted.puUploadHistory ?? 0}.
          {' '}Файлов Cloudinary: {purgeResult.cloudinaryDeleted}{purgeResult.cloudinaryErrors ? `, ошибок файлов: ${purgeResult.cloudinaryErrors}` : ''}.
        </div>
      )}

      {purgeConfirm && (
        <ModalShell variant="confirm" title="Подтверждение очистки"
          onClose={() => setPurgeConfirm(false)}
          onBackdrop={() => !purging && setPurgeConfirm(false)}>
            <div className="modal-body">
              <p>Будут <strong>безвозвратно</strong> удалены данные старше <strong>{purgeBefore}</strong>:</p>
              <ul style={{ margin: '8px 0 12px 18px' }}>
                <li>Истории загрузок: {purgePreview?.uploadHistory ?? 0}</li>
                <li>Истории по ПУ: {purgePreview?.puUploadHistory ?? 0}</li>
                <li>Истории проверок: {purgePreview?.checkHistory ?? 0}</li>
                <li>Уведомления: {purgePreview?.notifications ?? 0}</li>
                <li>Файлы Cloudinary: {purgePreview?.cloudinaryFiles ?? 0}</li>
              </ul>
              <p style={{ color: 'var(--green)' }}>Сохраняются: структура сети, тех-учёт секций, <strong>помесячные пики мощности</strong>, случаи перегруза, пользователи и справочники.</p>
            </div>
            <div className="modal-footer">
              <button className="action-btn" onClick={() => setPurgeConfirm(false)} disabled={purging}>Отмена</button>
              <button className="delete-selected-btn" onClick={handlePurge} disabled={purging}>
                {purging ? 'Удаление...' : 'Удалить безвозвратно'}
              </button>
            </div>
        </ModalShell>
      )}

      {loading && (
        <div className="db-loading">
          <div className="loading-animation">
            <RossetiLoader />
          </div>
        </div>
      )}
      
      {healthCheck && !loading && (
  <>
    {/* Общая статистика */}
    <div className="db-summary-grid">
      <div className="db-summary-card total">
        <div className="summary-icon"><IconChart className="ico" /></div>
        <div className="summary-content">
          <h4>Всего проблем</h4>
          <p className="summary-value">{healthCheck.stats.totalIssues}</p>
        </div>
      </div>
      
      <div className="db-summary-card error">
        <div className="summary-icon"><IconAlertCircle className="ico" style={{color:'var(--red)'}} /></div>
        <div className="summary-content">
          <h4>Критических</h4>
          <p className="summary-value">{healthCheck.stats.byType.error}</p>
        </div>
      </div>
      
      <div className="db-summary-card warning">
        <div className="summary-icon"><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /></div>
        <div className="summary-content">
          <h4>Предупреждений</h4>
          <p className="summary-value">{healthCheck.stats.byType.warning}</p>
        </div>
      </div>
      
      <div className="db-summary-card info">
        <div className="summary-icon"><IconInfo className="ico" style={{color:'var(--blue)'}} /></div>
        <div className="summary-content">
          <h4>Информация</h4>
          <p className="summary-value">{healthCheck.stats.byType.info}</p>
        </div>
      </div>
      
      {/* ВСЕГДА показываем - зеленые если 0, красные если > 0 */}
      <div className={`db-summary-card ${
        (healthCheck.stats.staleNotifications || 0) > 0 ? 'warning' : 'success'
      }`}>
        <div className="summary-icon">
          {(healthCheck.stats.staleNotifications || 0) > 0 ? <IconBell className="ico" /> : <IconCheck className="ico" style={{color:'var(--green)'}} />}
        </div>
        <div className="summary-content">
          <h4>Неактуальных уведомлений</h4>
          <p className="summary-value">{healthCheck.stats.staleNotifications || 0}</p>
        </div>
      </div>
      
      <div className={`db-summary-card ${
        (healthCheck.stats.missingNotifications || 0) > 0 ? 'error' : 'success'
      }`}>
        <div className="summary-icon">
          {(healthCheck.stats.missingNotifications || 0) > 0 ? <IconMegaphone className="ico" /> : <IconCheck className="ico" style={{color:'var(--green)'}} />}
        </div>
        <div className="summary-content">
          <h4>Отсутствующих уведомлений</h4>
          <p className="summary-value">{healthCheck.stats.missingNotifications || 0}</p>
        </div>
      </div>
            </div>
          
          
          {/* Статистика записей */}
          <div className="db-records-section">
            <h4><IconChart className="ico" /> Статистика базы данных</h4>
            <div className="db-records-grid">
              <div className="record-stat">
                <span className="record-label">Структура сети:</span>
                <span className="record-value">{healthCheck.stats.totalRecords.networkStructures}</span>
              </div>
              <div className="record-stat">
                <span className="record-label">Статусы ПУ:</span>
                <span className="record-value">{healthCheck.stats.totalRecords.puStatuses}</span>
              </div>
              <div className="record-stat">
                <span className="record-label">Уведомления:</span>
                <span className="record-value">{healthCheck.stats.totalRecords.notifications}</span>
              </div>
              <div className="record-stat">
                <span className="record-label">История проверок:</span>
                <span className="record-value">{healthCheck.stats.totalRecords.checkHistory}</span>
              </div>
              <div className="record-stat">
                <span className="record-label">История загрузок:</span>
                <span className="record-value">{healthCheck.stats.totalRecords.uploadHistory}</span>
              </div>
              <div className="record-stat">
                <span className="record-label">Пользователи:</span>
                <span className="record-value">{healthCheck.stats.totalRecords.users}</span>
              </div>
            </div>
          </div>

          {/* Занятое место в базе + оценка запаса */}
          {healthCheck.stats.dbSize && (() => {
            const s = healthCheck.stats.dbSize;
            const mb = (b) => (b / 1024 / 1024).toFixed(1) + ' МБ';
            return (
              <div className="db-records-section">
                <h4><IconDatabase className="ico" /> Занятое место в базе</h4>
                <div className="db-records-grid">
                  <div className="record-stat"><span className="record-label">Размер БД:</span><span className="record-value">{mb(s.bytes)}</span></div>
                  <div className="record-stat"><span className="record-label">Новых записей за 30 дней:</span><span className="record-value">{s.rows30}</span></div>
                  <div className="record-stat"><span className="record-label">Прирост в месяц (≈):</span><span className="record-value">{s.bytesPerMonth ? mb(s.bytesPerMonth) : '—'}</span></div>
                  {s.quotaMb ? (
                    <>
                      <div className="record-stat"><span className="record-label">Квота (DB_QUOTA_MB):</span><span className="record-value">{s.quotaMb} МБ</span></div>
                      <div className="record-stat"><span className="record-label">Хватит ещё (≈):</span><span className="record-value">{s.monthsLeft != null ? `~${s.monthsLeft} мес` : '—'}</span></div>
                    </>
                  ) : (
                    <div className="record-stat" style={{ gridColumn: '1 / -1' }}>
                      <span className="record-label">Оценка «хватит на N месяцев»:</span>
                      <span className="record-value" style={{ fontSize: 12, color: 'var(--text-muted)' }}>задайте env DB_QUOTA_MB (квота БД в МБ)</span>
                    </div>
                  )}
                </div>
                <div style={{ marginTop: 10 }}>
                  <div className="record-label" style={{ marginBottom: 6 }}>Крупнейшие таблицы:</div>
                  <table className="diag-table">
                    <thead><tr><th>Таблица</th><th>Размер</th><th>Записей (≈)</th></tr></thead>
                    <tbody>
                      {s.tables.map((t, i) => (
                        <tr key={i}><td style={{ textAlign: 'left' }}>{t.table}</td><td>{mb(t.bytes)}</td><td>{Number(t.rows).toLocaleString('ru-RU')}</td></tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </div>
            );
          })()}

          {/* Обнаруженные проблемы */}
          {healthCheck.issues.length === 0 ? (
            <div className="db-no-issues">
              <div className="no-issues-icon"><IconCheck className="ico" style={{color:'var(--green)'}} /></div>
              <h4>База данных в отличном состоянии!</h4>
              <p>Проблем не обнаружено. Все работает корректно.</p>
            </div>
          ) : (
            <div className="db-issues-section">
              <h4><IconSearch className="ico" /> Обнаруженные проблемы</h4>
              <div className="db-issues-list">
  {healthCheck.issues.map((issue, idx) => {
    const cleanupInfo = getCleanupDescription(issue.type);
    return (
      <div 
        key={idx} 
        className={`db-issue-card severity-${issue.severity}`}
        style={{borderLeftColor: getSeverityColor(issue.severity)}}
      >
        <div className="issue-header">
          <div className="issue-title-row">
            <span className="issue-type-icon">{cleanupInfo.icon}</span>
            <span className="issue-severity-icon">{getSeverityIcon(issue.severity)}</span>
            <h5>{issue.description}</h5>
          </div>
          <span className="issue-count-badge">{issue.count} записей</span>
        </div>
        
        {/* ДОБАВЬ ЭТО ПОСЛЕ issue-header, ПЕРЕД старым details: */}
        
        {/* Специальное отображение для неактуальных уведомлений */}
        {issue.type === 'stale_notifications' && issue.items && issue.items.length > 0 && (
          <details className="issue-details">
            <summary>
              <span><IconClipboard className="ico" /></span>
              Показать неактуальные уведомления (первые 10)
            </summary>
            <div className="stale-notifs-list">
              {issue.items.slice(0, 10).map((item, i) => (
                <div key={i} className="stale-notif-card">
                  <div className="stale-notif-header">
                    <div>
                      <strong>ПУ #{item.puNumber}</strong>
                      <span className="notif-type-badge">{item.type === 'error' ? 'Ошибка' : 'АСКУЭ'}</span>
                    </div>
                    <span className={`stale-status ${item.currentStatus === 'checked_ok' ? 'ok' : 'not-found'}`}>
                      {item.currentStatus === 'checked_ok' ? 'Проверен' : 'Не найден'}
                    </span>
                  </div>
                  <div className="stale-notif-info">
                    <p><strong><IconMapPin className="ico" /> Местоположение:</strong> {item.tpName} - {item.vlName}</p>
                    <p><strong><IconBuilding className="ico" /> РЭС:</strong> {item.resName}</p>
                    <p><strong><IconCalendar className="ico" /> Уведомление создано:</strong> {new Date(item.notifCreated).toLocaleString('ru-RU')}</p>
                    {item.lastCheck && (
                      <p><strong><IconCheck className="ico" style={{color:'var(--green)'}} /> Последняя успешная проверка:</strong> {new Date(item.lastCheck).toLocaleString('ru-RU')}</p>
                    )}
                  </div>
                  <div className="stale-reason">
                    <IconLightbulb className="ico" /> <strong>Причина:</strong> {item.reason}
                  </div>
                </div>
              ))}
              {issue.items.length > 10 && (
                <div className="more-items">
                  ... и еще {issue.items.length - 10} неактуальных уведомлений
                </div>
              )}
            </div>
          </details>
        )}

        {/* Специальное отображение для отсутствующих уведомлений */}
{issue.type === 'missing_notifications' && issue.items && issue.items.length > 0 && (
  <details className="issue-details">
    <summary>
      <span><IconClipboard className="ico" /></span>
      Показать ПУ без уведомлений (первые 10)
    </summary>
    <div className="missing-notifs-list">
      {issue.items.slice(0, 10).map((item, i) => (
        <div key={i} className="missing-notif-card">
          <div className="missing-notif-header">
            <div>
              <strong>ПУ #{item.puNumber}</strong>
              <span className="status-badge status-error"><IconX className="ico" style={{color:'var(--red)'}} /> Ошибка</span>
            </div>
            <span className="missing-badge">
              <IconMegaphone className="ico" /> Нет уведомления
            </span>
          </div>
          <div className="missing-notif-info">
            <p><strong><IconMapPin className="ico" /> Местоположение:</strong> {item.tpName} - {item.vlName}</p>
            <p><strong><IconBuilding className="ico" /> РЭС:</strong> {item.resName}</p>
            <p><strong><IconAlertCircle className="ico" style={{color:'var(--red)'}} /> Статус ПУ:</strong> {item.status}</p>
            <p><strong><IconCalendar className="ico" /> Последняя проверка:</strong> {
              item.lastCheck 
                ? new Date(item.lastCheck).toLocaleString('ru-RU')
                : 'Неизвестно'
            }</p>
          </div>
          <div className="missing-error-preview">
            <strong><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /> Ошибка в статусе:</strong>
            <pre>{item.errorDetails?.substring(0, 200)}...</pre>
          </div>
        </div>
      ))}
      {issue.items.length > 10 && (
        <div className="more-items">
          ... и еще {issue.items.length - 10} ПУ без уведомлений
        </div>
      )}
    </div>
  </details>
)}
          
        {/* Неактуальные проблемные ВЛ — свой список примеров */}
        {issue.type === 'irrelevant_problem_vl' && issue.items && issue.items.length > 0 && (
          <details className="issue-details">
            <summary>
              <span><IconClipboard className="ico" /></span>
              Показать неактуальные проблемные ВЛ (первые 10)
            </summary>
            <ul className="issue-items-list">
              {issue.items.slice(0, 10).map((item, i) => (
                <li key={i}>
                  <strong>ПУ {item.puNumber}</strong> — {item.tpName} / {item.vlName}: {item.reason}
                </li>
              ))}
              {issue.count > 10 && (
                <li className="more-items">... и еще {issue.count - 10}</li>
              )}
            </ul>
          </details>
        )}

        {/* Старое отображение для остальных типов */}
        {!['stale_notifications', 'irrelevant_problem_vl'].includes(issue.type) && issue.items && issue.items.length > 0 && (
          <details className="issue-details">
            <summary>
              <span><IconClipboard className="ico" /></span>
              Показать примеры (первые 10)
            </summary>
            <ul className="issue-items-list">
              {issue.items.slice(0, 10).map((item, i) => (
                <li key={i}>
                  {typeof item === 'object' ? 
                    `ПУ: ${item.puNumber} (встречается ${item.count} раз)` : 
                    item
                  }
                </li>
              ))}
              {issue.items.length > 10 && (
                <li className="more-items">... и еще {issue.items.length - 10}</li>
              )}
            </ul>
          </details>
        )}
        
        {/* Кнопка очистки */}
        {['orphaned_pu_status', 'duplicate_pu_statuses', 'old_unread_notifications', 
          'orphaned_notifications', 'checks_without_res', 'broken_file_references', 
          'stale_problem_vl', 'irrelevant_problem_vl', 'stale_notifications', 'missing_notifications'].includes(issue.type) && (
          <button 
            className="btn-cleanup"
            onClick={() => {
              setCleanupType(issue.type);
              setShowCleanupModal(true);
            }}
          >
            <span><IconBroom className="ico" /></span>
            Очистить
          </button>
        )}
      </div>
    );
  })}
</div>
            </div>
          )}
          
          <div className="db-check-time">
            <span><IconClock className="ico" /></span>
            Последняя проверка: {new Date(healthCheck.checkedAt).toLocaleString('ru-RU')}
          </div>
        </>
      )}
      
      {/* Модальное окно очистки */}
      {showCleanupModal && (
        <ModalShell variant="confirm" title={getCleanupDescription(cleanupType).title}
          className="cleanup-modal-modern" onClose={() => setShowCleanupModal(false)}>
            <div className="modal-body">
              <p className="cleanup-description">{getCleanupDescription(cleanupType).desc}</p>
              <div className="warning-box">
                <span className="warning-icon"><IconAlertTriangle className="ico" style={{color:'var(--amber)'}} /></span>
                <div>
                  <strong>Внимание!</strong>
                  <p>Это действие нельзя отменить. Все удаленные данные будут потеряны безвозвратно.</p>
                </div>
              </div>
              <div className="form-group">
                <label>
                  <span className="label-icon"><IconLock className="ico" /></span>
                  Введите пароль администратора:
                </label>
                <input
                  type="password"
                  value={cleanupPassword}
                  onChange={(e) => setCleanupPassword(e.target.value)}
                  placeholder="••••••••"
                  autoFocus
                  className="password-input-modern"
                />
              </div>
            </div>
            <div className="modal-footer-modern">
              <button 
                className="btn-cancel-modern" 
                onClick={() => setShowCleanupModal(false)}
              >
                Отмена
              </button>
              <button 
                className="btn-danger-modern" 
                onClick={handleCleanup}
                disabled={!cleanupPassword || cleaning}
              >
                {cleaning ? (
                  <>
                                        Очистка...
                  </>
                ) : (
                  <>
                    <span><IconBroom className="ico" /></span>
                    Очистить
                  </>
                )}
              </button>
            </div>
        </ModalShell>
      )}
    </div>
  );
}

// Выводы по точке мониторинга (модалка рекомендаций «Карты опроса»). Тексты — в
// одном месте; правило — первое совпавшее. В футере replace+planReplace = «К замене».
const POLL_VERDICTS = {
  replace:     { priority: 1, color: 'red',   bucket: 'toReplace', text: 'Не опрашивается — точка бесполезна для мониторинга. Заменить на СПОДЭС/РиМ / включить в опрос' },
  planReplace: { priority: 2, color: 'amber', bucket: 'toReplace', text: 'Журнал напряжений недоступен — необходимо заменить ПУ в структуре сети' },
  restore:     { priority: 3, color: 'amber', bucket: 'toRestore', text: 'Журнал доступен (СПОДЭС/РиМ), но сбор не идёт — восстановить опрос (связь/маршрут), замена не требуется' },
  fill:        { priority: 4, color: 'gray',  bucket: 'toFill',    text: 'ПУ не задан — заполнить в структуре' },
  ok:          { priority: 5, color: 'green', bucket: 'ok',        text: 'В порядке' },
};
const pollVerdict = (cell) => {
  if (cell.status === 'no_pu') return POLL_VERDICTS.fill;
  if (cell.status === 'absent') return POLL_VERDICTS.replace;
  // Журнал напряжений доступен для СПОДЭС ИЛИ РиМ. Нет журнала → плановая замена.
  if (!cell.journal) return POLL_VERDICTS.planReplace;          // найден, но не СПОДЭС и не РиМ
  if (cell.status === 'not_collected') return POLL_VERDICTS.restore;  // журнал есть, но сбор не идёт
  return POLL_VERDICTS.ok;                                       // журнал есть + собирается
};

// Явный бейдж СПОДЭС: полное слово, неоново-синее СТАТИЧНОЕ свечение (стиль иконок
// PageHeader — без пульсаций). На узких экранах сокращается до «СПДС» через CSS.
const SpodesBadge = () => (
  <span className="spodes-neon"><span className="sb-full">СПОДЭС</span><span className="sb-short">СПДС</span></span>
);

// ═══ Раздел «Карта опроса структуры сети»: сверка ПУ структуры со срезом «Опрос ПУ» ═══
function PollMap({ selectedRes }) {
  const { user } = useContext(AuthContext);
  const canSync = hasPerm(user, 'pollmap_sync') || user.role === 'uec_responsible';
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [syncing, setSyncing] = useState(false);
  const [notice, setNotice] = useState(null);   // { type:'ok'|'err', text }
  const [searchTp, setSearchTp] = useState('');
  const [statusFilter, setStatusFilter] = useState(null); // collected|not_collected|absent|no_pu
  const [legendCollapsed, setLegendCollapsed] = useState(false);
  const [showOrphans, setShowOrphans] = useState(false);
  const [recModal, setRecModal] = useState(null); // { title, dockTitle, resName, tpName, items:[{position,pu,cell}] }
  const [copiedSerial, setCopiedSerial] = useState(null);

  const load = async () => {
    setLoading(true);
    try {
      const params = (user.role === 'admin' && selectedRes) ? { resId: selectedRes } : {};
      const { data } = await api.get('/api/poll-map', { params });
      setData(data);
    } catch (e) {
      setNotice({ type: 'err', text: e.response?.data?.error || e.message });
    } finally { setLoading(false); }
  };
  useEffect(() => { load(); /* eslint-disable-next-line */ }, [selectedRes]);

  const sync = async () => {
    setSyncing(true); setNotice(null);
    try {
      const { data } = await api.post('/api/poll-map/sync');
      if (data.warning || data.code === 'empty_snapshot') {
        // Пустой срез — не ошибка, но и не успех: жёлтый баннер, прежние данные сохранены.
        setNotice({ type: 'warn', text: data.hint, code: data.code });
        await load();
      } else {
        const dt = data.snapshotAt ? new Date(data.snapshotAt).toLocaleString('ru-RU') : '—';
        setNotice({ type: 'ok', text: `Синхронизировано: ${data.total} ПУ, из них собирается ${data.collected}, СПОДЭС ${data.spodes}, срез от ${dt}` });
        await load();
      }
    } catch (e) {
      const d = e.response?.data || {};
      setNotice({ type: 'err', text: d.hint || d.error || 'Ошибка синхронизации', code: d.code, detail: d.detail });
    } finally { setSyncing(false); }
  };

  // Полоса легенды → кружки при скролле (паттерн из «Структуры»).
  useEffect(() => {
    let el = null, raf = null, ticking = false;
    const evaluate = () => {
      ticking = false; if (!el) return;
      const top = el.scrollTop;
      setLegendCollapsed(prev => (!prev && top > 130) ? true : (prev && top < 80) ? false : prev);
    };
    const onScroll = () => { if (!ticking) { ticking = true; requestAnimationFrame(evaluate); } };
    const attach = () => { el = document.querySelector('.content'); if (el) { el.addEventListener('scroll', onScroll, { passive: true }); evaluate(); } else raf = requestAnimationFrame(attach); };
    attach();
    return () => { if (raf) cancelAnimationFrame(raf); if (el) el.removeEventListener('scroll', onScroll); };
  }, []);

  const pctCls = (p) => p >= 90 ? 'green' : p >= 70 ? 'amber' : 'red';
  const boxCls = (st) => st === 'collected' ? 'status-ok' : st === 'not_collected' ? 'status-pending'
    : st === 'absent' ? 'status-error' : st === 'no_data' ? 'status-nodata' : 'status-empty';
  const fmtDate = (d) => d ? new Date(d).toLocaleString('ru-RU') : '—';
  const fmtN = (v) => v == null ? '—' : v;   // прочерк, когда среза нет (noData)

  const rows = data?.rows || [];
  const sections = data?.sections || [];
  const filtered = rows.filter(r => {
    if (searchTp && !r.tpName?.toLowerCase().includes(searchTp.toLowerCase())) return false;
    if (statusFilter) {
      const has = [r.poll.start, r.poll.middle, r.poll.end].some(c => c.status === statusFilter);
      if (!has) return false;
    }
    return true;
  });
  const secFiltered = sections.filter(s => {
    if (searchTp && !s.tpName?.toLowerCase().includes(searchTp.toLowerCase())) return false;
    if (statusFilter && s.poll.status !== statusFilter) return false;
    return true;
  });
  // Группировка РЭС → ТП → { secs (тех.учёты), vls }. ТП появляется, даже если у неё
  // только секции без ВЛ (или наоборот).
  const byRes = {};
  const bucket = (resName, tpName) => {
    (byRes[resName] = byRes[resName] || {});
    return (byRes[resName][tpName] = byRes[resName][tpName] || { secs: [], vls: [] });
  };
  secFiltered.forEach(s => bucket(s.resName, s.tpName).secs.push(s));
  filtered.forEach(r => bucket(r.resName, r.tpName).vls.push(r));

  const puCell = (pu, cell) => {
    // Зона бейджа резервируется ВСЕГДА (фикс. высота) — квадраты всех ПУ ряда
    // стоят на одной линии независимо от наличия «СПОДЭС».
    const showSp = cell.spodes && (cell.status === 'collected' || cell.status === 'not_collected');
    return (
      <div className="pu-cell poll-pu-cell" title={cell.status === 'no_data' ? 'Нет данных среза' : (pu || 'ПУ не задан')}>
        <div className={`status-box ${boxCls(cell.status)}`}>{cell.status === 'no_pu' ? 'X' : ''}</div>
        <span className="pu-num-line">{pu || ''}</span>
        <span className="pu-badge-zone">{showSp && <SpodesBadge />}</span>
      </div>
    );
  };

  // Лукап кандидатов/совпадения ТП по «resName||tpName».
  const tpInfo = {};
  (data?.tps || []).forEach(t => { tpInfo[t.resName + '||' + t.tpName] = t; });
  const tpDataAvailable = !!data?.tpDataAvailable;

  // Мини-сводка ВЛ: Y — ПУ с заданными номерами, X — собирается, K — СПОДЭС.
  const vlMini = (r) => {
    const assigned = [r.startPu, r.middlePu, r.endPu];
    const cells = [r.poll.start, r.poll.middle, r.poll.end];
    const idx = assigned.map((p, i) => (p ? i : -1)).filter(i => i >= 0);
    const Y = idx.length;
    const X = idx.filter(i => cells[i].status === 'collected').length;
    const K = idx.filter(i => cells[i].spodes).length;
    if (Y === 0) return <span className="poll-vl-mini muted">ПУ не заданы</span>;
    const cls = X === Y ? 'green' : X === 0 ? 'red' : 'amber';
    return <span className="poll-vl-mini">опрос: <b className={`mini-${cls}`}>{X}/{Y}</b> · СПОДЭС: {K} из {Y}</span>;
  };

  // Открыть модалку рекомендаций по ВЛ / по секции (тех.учёту).
  const openVlRec = (r) => setRecModal({
    resName: r.resName, tpName: r.tpName,
    title: `${r.tpName} — ${r.vlName} · рекомендации`,
    dockTitle: `${r.tpName} — ${r.vlName}`,
    items: [
      { position: 'Начало', pu: r.startPu, cell: r.poll.start },
      { position: 'Середина', pu: r.middlePu, cell: r.poll.middle },
      { position: 'Конец', pu: r.endPu, cell: r.poll.end },
    ],
  });
  const openSecRec = (sec) => setRecModal({
    resName: sec.resName, tpName: sec.tpName,
    title: `${sec.tpName} — СШ-${toRoman(sec.sectionNumber)} · рекомендации`,
    dockTitle: `${sec.tpName} — СШ-${toRoman(sec.sectionNumber)}`,
    items: [{ position: 'Ввод (тех.учёт)', pu: sec.techPuNumber, cell: sec.poll }],
  });
  const copySerial = (s) => {
    try { navigator.clipboard?.writeText(s); setCopiedSerial(s); setTimeout(() => setCopiedSerial(null), 1200); } catch { /* noop */ }
  };

  const exportExcel = () => {
    if (!data || (!rows.length && !sections.length)) { alert('Нет данных для экспорта'); return; }
    const stRu = (s) => s === 'collected' ? 'Собирается' : s === 'not_collected' ? 'Не собирается' : s === 'absent' ? 'Отсутствует' : s === 'no_data' ? 'Нет данных' : 'ПУ не задан';
    const mapRows = [];
    rows.forEach(r => {
      [['Начало', r.startPu, r.poll.start], ['Середина', r.middlePu, r.poll.middle], ['Конец', r.endPu, r.poll.end]].forEach(([pos, pu, c]) => {
        if (!pu) return;
        mapRows.push({ 'Тип': 'ПУ ВЛ', 'РЭС': r.resName, 'ТП': r.tpName, 'ВЛ': r.vlName, 'Позиция': pos, '№ ПУ': pu, 'Статус': stRu(c.status), 'СПОДЭС': c.spodes ? 'да' : '' });
      });
    });
    sections.forEach(sec => {
      if (!sec.techPuNumber) return;
      mapRows.push({ 'Тип': 'тех.учёт', 'РЭС': sec.resName, 'ТП': sec.tpName, 'ВЛ': `СШ-${toRoman(sec.sectionNumber)}`, 'Позиция': 'ввод', '№ ПУ': sec.techPuNumber, 'Статус': stRu(sec.poll.status), 'СПОДЭС': sec.poll.spodes ? 'да' : '' });
    });
    const sumRows = [...(data.summary?.byRes || []), { ...data.summary?.total, resName: 'ИТОГО' }].map(a => ({
      'РЭС': a.resName, 'Всего ПУ': a.totalPu, 'Собирается': a.collected, 'Не собирается': a.notCollected,
      'Отсутствует': a.absent, 'СПОДЭС': a.spodes, 'ПУ не задан': a.noPu, '% сбора': a.coveragePct
    }));
    const orphRows = (data.orphans?.list || []).map(o => ({ '№ ПУ': o.serial, 'СПОДЭС': o.isSpodes ? 'да' : '', 'Собирается': o.isCollected ? 'да' : '' }));

    const wb = XLSX.utils.book_new();
    const statusColor = (v) => v === 'Собирается' ? XLS_COLORS.green : v === 'Не собирается' ? XLS_COLORS.amber : v === 'Отсутствует' ? XLS_COLORS.red : XLS_COLORS.gray; // 'Нет данных'/'ПУ не задан' → серый
    const ws1 = XLSX.utils.json_to_sheet(mapRows);
    ws1['!cols'] = [{ wch: 10 }, { wch: 18 }, { wch: 16 }, { wch: 16 }, { wch: 10 }, { wch: 16 }, { wch: 14 }, { wch: 8 }];
    styleExportSheet(ws1, { 'Статус': statusColor });
    XLSX.utils.book_append_sheet(wb, ws1, 'Карта');
    const ws2 = XLSX.utils.json_to_sheet(sumRows);
    ws2['!cols'] = [{ wch: 18 }, { wch: 9 }, { wch: 11 }, { wch: 13 }, { wch: 12 }, { wch: 8 }, { wch: 11 }, { wch: 9 }];
    styleExportSheet(ws2, { '% сбора': (v) => { const n = Number(v); return n >= 90 ? XLS_COLORS.green : n >= 70 ? XLS_COLORS.amber : XLS_COLORS.red; } });
    XLSX.utils.book_append_sheet(wb, ws2, 'Сводка');
    if (orphRows.length) {
      const ws3 = XLSX.utils.json_to_sheet(orphRows);
      ws3['!cols'] = [{ wch: 18 }, { wch: 8 }, { wch: 11 }];
      styleExportSheet(ws3);
      XLSX.utils.book_append_sheet(wb, ws3, 'Не в структуре');
    }
    XLSX.writeFile(wb, `Карта_опроса_${new Date().toLocaleDateString('ru-RU').split('.').join('-')}.xlsx`);
  };

  return (
    <div className="poll-map">
      <PageHeader icon={<IconMapPin size={24} />} title="Карта опроса структуры сети" />

      <div className="poll-topbar">
        <div className="poll-snapshot">
          {loading
            ? 'Срез загружается, это может занять несколько секунд, подождите…'
            : data?.snapshotAt
              ? <>Срез Пирамиды от <strong>{fmtDate(data.snapshotAt)}</strong></>
              : 'Срез не синхронизирован'}
        </div>
        <div className="poll-controls">
          <input type="text" className="search-input" placeholder="Поиск по ТП..." value={searchTp} onChange={(e) => setSearchTp(e.target.value)} />
          {canSync && (
            <button className="pm-btn pm-btn--refresh" onClick={sync} disabled={syncing}>
              <IconRefresh className="ico" /> {syncing ? 'Синхронизация...' : 'Синхронизировать'}
            </button>
          )}
          <button className="pm-btn pm-btn--excel" onClick={exportExcel} disabled={!rows.length && !sections.length}>
            <IconDownload className="ico" /> Excel
          </button>
        </div>
      </div>

      {notice && (
        <div className={`poll-notice ${notice.type}`}>
          <div className="poll-notice-main">
            {notice.type === 'ok' ? <IconCheck className="ico" /> : <IconAlertTriangle className="ico" />} {notice.text}
          </div>
          {(notice.code || notice.detail) && (
            <div className="poll-notice-tech">
              {notice.code && <code>{notice.code}</code>}{notice.detail ? ` · ${notice.detail}` : ''}
            </div>
          )}
        </div>
      )}

      {loading ? (
        <div className="loading"><RossetiLoader /></div>
      ) : (
        <>
          {/* Срез не синхронизирован — структура всё равно видна, данные опроса «прочерками» */}
          {data?.noData && (
            <div className="poll-notice warn poll-nosync">
              <div className="poll-notice-main">
                <IconAlertTriangle className="ico" /> Срез не синхронизирован — показана структура сети без данных опроса.
              </div>
              {canSync
                ? <button className="pm-btn pm-btn--refresh poll-nosync-btn" onClick={sync} disabled={syncing}>
                    <IconRefresh className="ico" /> {syncing ? 'Синхронизация...' : 'Синхронизировать'}
                  </button>
                : <div className="poll-notice-tech">Дождитесь синхронизации администратором.</div>}
            </div>
          )}
          {/* Сводка покрытия */}
          <div className="pm-table-wrap" style={{ marginBottom: 18 }}>
            <table className="pm-matrix">
              <thead>
                <tr><th className="pm-sticky">РЭС</th><th>Всего</th><th>Собирается</th><th>Не собирается</th><th>Отсутствует</th><th>СПОДЭС</th><th>ПУ не задан</th><th>% сбора</th></tr>
              </thead>
              <tbody>
                {(data.summary?.byRes || []).map(a => (
                  <tr key={a.resId}>
                    <td className="pm-sticky"><strong>{a.resName}</strong></td>
                    <td>{fmtN(a.totalPu)}</td><td>{fmtN(a.collected)}</td><td>{fmtN(a.notCollected)}</td>
                    <td>{fmtN(a.absent)}</td><td>{fmtN(a.spodes)}</td><td>{fmtN(a.noPu)}</td>
                    <td className={a.coveragePct == null ? '' : `pm-${pctCls(a.coveragePct)}`}><strong>{a.coveragePct == null ? '—' : `${a.coveragePct}%`}</strong></td>
                  </tr>
                ))}
                {data.summary?.total && (
                  <tr className="poll-total-row">
                    <td className="pm-sticky"><strong>ИТОГО</strong></td>
                    <td>{fmtN(data.summary.total.totalPu)}</td><td>{fmtN(data.summary.total.collected)}</td><td>{fmtN(data.summary.total.notCollected)}</td>
                    <td>{fmtN(data.summary.total.absent)}</td><td>{fmtN(data.summary.total.spodes)}</td><td>{fmtN(data.summary.total.noPu)}</td>
                    <td className={data.summary.total.coveragePct == null ? '' : `pm-${pctCls(data.summary.total.coveragePct)}`}><strong>{data.summary.total.coveragePct == null ? '—' : `${data.summary.total.coveragePct}%`}</strong></td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          {/* Легенда-фильтры */}
          <div className={`status-legend ${legendCollapsed ? 'is-collapsed' : ''} ${data?.noData ? 'is-muted' : ''}`}>
            {[['collected', 'status-ok', 'Собирается'], ['not_collected', 'status-pending', 'В опросе, не собирается'], ['absent', 'status-error', 'В опросе отсутствует'], ['no_pu', 'status-empty', 'ПУ не задан']].map(([k, cls, lbl]) => (
              <div key={k} className={`legend-item ${statusFilter === k ? 'active' : ''}`} onClick={() => setStatusFilter(statusFilter === k ? null : k)}>
                <span className={`status-box ${cls}`}>{k === 'no_pu' ? 'X' : ''}</span> {lbl}
              </div>
            ))}
          </div>
          <div className={`legend-dots ${legendCollapsed ? 'show' : ''}`} aria-hidden={!legendCollapsed}>
            {[['collected', 'status-ok', 'Собирается'], ['not_collected', 'status-pending', 'В опросе, не собирается'], ['absent', 'status-error', 'В опросе отсутствует'], ['no_pu', 'status-empty', 'ПУ не задан']].map(([k, cls, lbl]) => (
              <button key={k} type="button" title={lbl} className={`legend-dot ${cls} ${statusFilter === k ? 'active' : ''}`} onClick={() => setStatusFilter(statusFilter === k ? null : k)}>{k === 'no_pu' ? 'X' : ''}</button>
            ))}
          </div>

          {/* Дерево */}
          <div className="poll-tree">
            {Object.keys(byRes).sort((a, b) => a.localeCompare(b, 'ru')).map(resName => (
              <div key={resName} className="poll-res">
                <h3 className="poll-res-title">{resName}</h3>
                {Object.keys(byRes[resName]).sort((a, b) => a.localeCompare(b, 'ru')).map(tpName => (
                  <div key={tpName} className="tp-card">
                    <div className="poll-grid poll-grid--head poll-tp-head">
                      <span className="poll-vl-name poll-tp-name">{tpName}</span>
                      <span className="pu-col-label">Начало</span><span className="pu-col-label">Середина</span><span className="pu-col-label">Конец</span>
                    </div>
                    {/* Секции шин (тех.учёт на вводе) — клик открывает рекомендации */}
                    {byRes[resName][tpName].secs.map(sec => (
                      <div key={'s' + sec.id} className="poll-grid poll-section-line poll-row-click" onClick={() => openSecRec(sec)} title="Рекомендации по тех.учёту">
                        <span className="poll-vl-name section-line-name">
                          СШ-{toRoman(sec.sectionNumber)}{sec.tnKva != null ? ` · ${sec.tnKva} кВА` : ''}{sec.techPuNumber ? ` · тех.учёт № ${sec.techPuNumber}` : ''}
                        </span>
                        {puCell(sec.techPuNumber, sec.poll)}
                        <span className="pu-cell-empty" /><span className="pu-cell-empty" />
                      </div>
                    ))}
                    {byRes[resName][tpName].vls.map(r => (
                      <div key={r.id} className="poll-grid poll-row-click" onClick={() => openVlRec(r)} title="Рекомендации по ВЛ">
                        <span className="poll-vl-name poll-vl-name--stack">
                          <span className="poll-vl-title">{r.vlName}</span>
                          {!data.noData && vlMini(r)}
                        </span>
                        {puCell(r.startPu, r.poll.start)}
                        {puCell(r.middlePu, r.poll.middle)}
                        {puCell(r.endPu, r.poll.end)}
                      </div>
                    ))}
                  </div>
                ))}
              </div>
            ))}
            {filtered.length === 0 && secFiltered.length === 0 && <div className="no-data" style={{ padding: 16 }}>Нет объектов по фильтру</div>}
          </div>

          {/* В опросе, но нет в структуре */}
          {data.orphans?.count > 0 && (
            <div className="poll-orphans">
              <button className="poll-orphans-toggle" onClick={() => setShowOrphans(v => !v)}>
                {showOrphans ? <IconArrowDown className="ico" /> : <IconArrowRight className="ico" />} В опросе, но нет в структуре ({data.orphans.count})
              </button>
              {showOrphans && (
                <div className="poll-orphans-list">
                  {data.orphans.list.map((o, i) => (
                    <span key={i} className="poll-orphan-chip">
                      {o.serial}
                      {o.isCollected && <span className="mini-badge green">собирается</span>}
                      {o.isSpodes && <SpodesBadge />}
                    </span>
                  ))}
                </div>
              )}
            </div>
          )}
        </>
      )}

      {/* Модалка рекомендаций по ВЛ / тех.учёту */}
      {recModal && (
        <ModalShell
          title={recModal.title}
          dockTitle={recModal.dockTitle}
          icon={<IconMapPin size={16} />}
          className="poll-rec-modal"
          onClose={() => setRecModal(null)}
        >
          {(() => {
            const nd = !!data?.noData;
            const enriched = recModal.items.map((it, i) => ({ ...it, order: i, verdict: nd ? null : pollVerdict(it.cell) }));
            const sorted = nd ? enriched : [...enriched].sort((a, b) => (a.verdict.priority - b.verdict.priority) || (a.order - b.order));
            const buckets = { toReplace: 0, toRestore: 0, toFill: 0, ok: 0 };
            if (!nd) enriched.forEach(e => { buckets[e.verdict.bucket]++; });
            const tp = tpInfo[recModal.resName + '||' + recModal.tpName];
            const cands = tp?.candidates || [];
            const stLabel = (st) => st === 'collected' ? 'Собирается' : st === 'not_collected' ? 'Не собирается' : st === 'absent' ? 'Отсутствует' : st === 'no_pu' ? 'ПУ не задан' : 'Нет данных';
            return (
              <div className="poll-rec">
                {nd && (
                  <div className="poll-notice warn" style={{ marginBottom: 12 }}>
                    <div className="poll-notice-main"><IconAlertTriangle className="ico" /> Срез не синхронизирован — показан только состав. Выполните синхронизацию для рекомендаций.</div>
                  </div>
                )}
                <div className="poll-rec-tablewrap">
                <table className="poll-rec-table">
                  <thead><tr><th>Позиция</th><th>№ ПУ</th><th>Статус</th><th>Вывод</th></tr></thead>
                  <tbody>
                    {sorted.map((it, i) => (
                      <tr key={i}>
                        <td>{it.position}</td>
                        <td className="poll-rec-pu">{it.pu || '—'}{it.cell.spodes && <SpodesBadge />}</td>
                        <td><span className={`rec-status rec-st-${it.cell.status}`}>{stLabel(it.cell.status)}</span></td>
                        <td>{it.verdict ? <span className={`rec-verdict rec-v-${it.verdict.color}`}>{it.verdict.text}</span> : <span className="muted">—</span>}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                </div>
                {!nd && (
                  <div className="poll-rec-footer">
                    <span className="poll-rec-foot-item"><span className="rec-dot rec-v-red" />К замене: <b>{buckets.toReplace}</b></span>
                    <span className="poll-rec-foot-item"><span className="rec-dot rec-v-amber" />Восстановить опрос: <b>{buckets.toRestore}</b></span>
                    <span className="poll-rec-foot-item"><span className="rec-dot rec-v-gray" />Заполнить: <b>{buckets.toFill}</b></span>
                    <span className="poll-rec-foot-item"><span className="rec-dot rec-v-green" />В порядке: <b>{buckets.ok}</b></span>
                  </div>
                )}
                {tpDataAvailable && (
                  <div className="poll-cand">
                    <div className="poll-cand-head">Кандидаты для контроля (СПОДЭС/РиМ) на этой ТП ({cands.length})</div>
                    {cands.length === 0 ? (
                      <div className="muted" style={{ padding: '6px 0' }}>
                        {tp?.tpMatched ? 'Свободных ПУ с журналом напряжений (СПОДЭС/РиМ) на этой ТП нет.' : 'ТП не найдена в срезе Пирамиды по имени (сверьте написание).'}
                      </div>
                    ) : (
                      <>
                        <div className="poll-cand-hint muted">Эти ПУ уже опрашиваются и имеют журнал напряжений (СПОДЭС или РиМ) на этой ТП — можно рассмотреть как контрольные точки взамен проблемных. Клик по номеру копирует серийник.</div>
                        <div className="poll-cand-tablewrap">
                          <table className="poll-cand-table">
                            <thead><tr><th>№ ПУ</th><th>Адрес (точка учёта)</th><th>Тип</th><th>Опрос</th></tr></thead>
                            <tbody>
                              {cands.map((c, i) => (
                                <tr key={i}>
                                  <td>
                                    <button className="poll-cand-serial" title="Скопировать серийник" onClick={() => copySerial(c.serial)}>
                                      {c.serial}{copiedSerial === c.serial && <span className="poll-copied">скопировано</span>}
                                    </button>
                                  </td>
                                  <td className="poll-cand-addr">{c.tuPath || '—'}</td>
                                  <td className="poll-cand-poll">
                                    {c.isSpodes ? <SpodesBadge /> : <span className="mini-badge rim">РиМ</span>}
                                  </td>
                                  <td className="poll-cand-poll">
                                    {c.isCollected
                                      ? <span className="mini-badge green">собирается</span>
                                      : <span className="mini-badge">не собирается</span>}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </>
                    )}
                  </div>
                )}
              </div>
            );
          })()}
        </ModalShell>
      )}
    </div>
  );
}

// ═══ Раздел «Права доступа» (только суперадмин): матрица админы × права ═══
function PermissionsAdmin() {
  const [catalog, setCatalog] = useState({});
  const [admins, setAdmins] = useState([]);
  const [loading, setLoading] = useState(true);
  const [savingId, setSavingId] = useState(null);
  const [dirty, setDirty] = useState({}); // userId -> локально отредактированные права

  const load = async () => {
    setLoading(true);
    try {
      const { data } = await api.get('/api/admin/permissions');
      setCatalog(data.catalog || {});
      setAdmins(data.admins || []);
      setDirty({});
    } catch (e) {
      showToast(e.response?.data?.error || 'Ошибка загрузки прав', 'error');
    } finally { setLoading(false); }
  };
  useEffect(() => { load(); /* eslint-disable-next-line */ }, []);

  const permsOf = (a) => dirty[a.id] || a.permissions || {};
  const isDirty = (a) => !!dirty[a.id];
  const toggle = (a, key) => {
    const cur = { ...permsOf(a) };
    if (cur[key]) delete cur[key]; else cur[key] = true;
    setDirty(d => ({ ...d, [a.id]: cur }));
  };
  const save = async (a) => {
    setSavingId(a.id);
    try {
      const { data } = await api.put(`/api/admin/permissions/${a.id}`, { permissions: permsOf(a) });
      setAdmins(list => list.map(x => x.id === a.id ? { ...x, permissions: data.permissions } : x));
      setDirty(d => { const n = { ...d }; delete n[a.id]; return n; });
      showToast('Права сохранены — вступят в силу в течение минуты', 'ok');
    } catch (e) {
      showToast(e.response?.data?.error || 'Ошибка сохранения прав', 'error');
    } finally { setSavingId(null); }
  };

  const catKeys = Object.keys(catalog);

  return (
    <div className="permissions-admin">
      <PageHeader icon={<IconLock size={24} />} title="Права доступа" />
      <p className="perm-hint muted">
        Отметьте разрешённые действия для каждого пользователя (админы, загрузчики, ответственные РЭС).
        Изменения применяются <strong>без перелогина</strong> и вступают в силу в течение минуты.
        Суперадмин может всё и в списке не показан.
      </p>
      {loading ? (
        <div className="loading"><RossetiLoader /></div>
      ) : admins.length === 0 ? (
        <div className="no-data" style={{ padding: 16 }}>Пользователей нет. Создайте их в «Настройках → Пользователи».</div>
      ) : (
        <div className="perm-table-wrap">
          <table className="perm-table">
            <thead>
              <tr>
                <th className="perm-sticky">Пользователь</th>
                {catKeys.map(k => <th key={k} title={catalog[k]}>{catalog[k]}</th>)}
                <th></th>
              </tr>
            </thead>
            <tbody>
              {admins.map(a => (
                <tr key={a.id}>
                  <td className="perm-sticky">
                    <div className="perm-admin-name">{a.fio}</div>
                    <div className="perm-admin-login muted">
                      {a.login}
                      {' · '}{a.role === 'admin' ? 'Администратор' : a.role === 'uploader' ? 'Загрузчик АСКУЭ' : a.role === 'res_responsible' ? 'Ответственный РЭС' : a.role}
                      {a.resName ? ` · ${a.resName}` : ''}
                    </div>
                  </td>
                  {catKeys.map(k => (
                    <td key={k} className="perm-check-cell">
                      <input type="checkbox" checked={!!permsOf(a)[k]} onChange={() => toggle(a, k)} />
                    </td>
                  ))}
                  <td>
                    <button className="pm-btn pm-btn--refresh perm-save-btn" disabled={!isDirty(a) || savingId === a.id} onClick={() => save(a)}>
                      {savingId === a.id ? 'Сохранение...' : 'Сохранить'}
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

// «Глубокий анализ работ» — детектор формально закрытых мероприятий РЭС.
// Вся математика на бэкенде (GET /api/analytics/work-quality) из CheckHistory;
// здесь только UI. Роль admin, работает по кнопке (без поллинга).
const DA_LEVEL_COLOR = { red: '#B91C1C', yellow: '#B45309', green: '#15803D' };
const DA_LEVEL_LABEL = { red: 'Высокий риск', yellow: 'Средний риск', green: 'Норма' };
const DA_DEFAULTS = { resId: '', dateFrom: '', dateTo: '', windowDays: 4, brigades: 3, perBrigade: 5 };

function DeepWorkAnalysis() {
  const [resList, setResList] = useState([]);
  const [settings, setSettings] = useState(() => {
    try {
      const saved = JSON.parse(localStorage.getItem('deepAnalysisSettings') || '{}');
      return { ...DA_DEFAULTS, ...saved };
    } catch { return { ...DA_DEFAULTS }; }
  });
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [expanded, setExpanded] = useState({});

  useEffect(() => {
    api.get('/api/res/list').then(r => setResList(r.data)).catch(e => console.error('Error loading RES list:', e));
  }, []);

  useEffect(() => {
    localStorage.setItem('deepAnalysisSettings', JSON.stringify(settings));
  }, [settings]);

  const setF = (key, value) => setSettings(prev => ({ ...prev, [key]: value }));

  const analyze = async () => {
    setLoading(true); setError('');
    try {
      const params = new URLSearchParams();
      if (settings.resId) params.set('resId', settings.resId);
      if (settings.dateFrom) params.set('dateFrom', settings.dateFrom);
      if (settings.dateTo) params.set('dateTo', settings.dateTo);
      params.set('windowDays', settings.windowDays);
      params.set('brigades', settings.brigades);
      params.set('perBrigade', settings.perBrigade);
      const resp = await api.get(`/api/analytics/work-quality?${params}`);
      setData(resp.data);
      setExpanded({});
    } catch (e) {
      console.error('Deep analysis error:', e);
      setError(e.response?.data?.error || 'Ошибка анализа');
      setData(null);
    } finally {
      setLoading(false);
    }
  };

  const dt = (v) => v ? new Date(v).toLocaleString('ru-RU') : '—';
  const pct = (v) => `${Math.round((v || 0) * 100)}%`;

  const exportExcel = () => {
    if (!data || !data.rows.length) { alert('Нет данных для экспорта'); return; }
    const sumRows = data.rows.map(r => ({
      'РЭС': r.resName,
      'Балл': r.score,
      'Закрыто': r.totalClosed,
      'Уникальных ТП': r.uniqueTp,
      'Пик ТП/окно': r.capacity.maxTpInWindow,
      'Ёмкость': r.capacity.capacityLimit,
      'Превышение, раз': Number(r.capacity.ratio.toFixed(2)),
      '% ночных': Math.round(r.afterHoursShare * 100),
      '% провалов перепроверки': Math.round(r.recheck.failShare * 100)
    }));
    const reasonRows = [];
    data.rows.forEach(r => {
      if (!r.reasons.length) reasonRows.push({ 'РЭС': r.resName, 'Причина': '—' });
      else r.reasons.forEach(reason => reasonRows.push({ 'РЭС': r.resName, 'Причина': reason }));
    });
    const scoreColor = (v) => { const n = Number(v); return n >= 60 ? XLS_COLORS.red : n >= 30 ? XLS_COLORS.amber : XLS_COLORS.green; };
    const wb = XLSX.utils.book_new();
    const ws1 = XLSX.utils.json_to_sheet(sumRows);
    ws1['!cols'] = [{ wch: 18 }, { wch: 7 }, { wch: 9 }, { wch: 14 }, { wch: 12 }, { wch: 10 }, { wch: 15 }, { wch: 10 }, { wch: 22 }];
    styleExportSheet(ws1, { 'Балл': scoreColor });
    XLSX.utils.book_append_sheet(wb, ws1, 'Сводка');
    const ws2 = XLSX.utils.json_to_sheet(reasonRows);
    ws2['!cols'] = [{ wch: 18 }, { wch: 80 }];
    styleExportSheet(ws2);
    XLSX.utils.book_append_sheet(wb, ws2, 'Причины');
    XLSX.writeFile(wb, `Глубокий_анализ_работ_${new Date().toLocaleDateString('ru-RU').split('.').join('-')}.xlsx`);
  };

  return (
    <div className="deep-analysis">
      <PageHeader icon={<IconSearch size={24} />} title="Глубокий анализ работ" />

      <div className="history-filters">
        <div className="filter-row">
          <div className="filter-group">
            <label>РЭС</label>
            <select value={settings.resId} onChange={e => setF('resId', e.target.value)}>
              <option value="">Все РЭС</option>
              {resList.map(r => <option key={r.id} value={r.id}>{r.name}</option>)}
            </select>
          </div>
          <div className="filter-group">
            <label>Период с</label>
            <input type="date" value={settings.dateFrom} onChange={e => setF('dateFrom', e.target.value)} />
          </div>
          <div className="filter-group">
            <label>по</label>
            <input type="date" value={settings.dateTo} onChange={e => setF('dateTo', e.target.value)} />
          </div>
          <div className="filter-group">
            <label>Окно, дней</label>
            <input type="number" min="1" max="60" value={settings.windowDays} onChange={e => setF('windowDays', e.target.value)} />
          </div>
          <div className="filter-group">
            <label>Бригад ОВБ</label>
            <input type="number" min="1" max="100" value={settings.brigades} onChange={e => setF('brigades', e.target.value)} />
          </div>
          <div className="filter-group">
            <label>План ТП/бригаду/день</label>
            <input type="number" min="1" max="100" value={settings.perBrigade} onChange={e => setF('perBrigade', e.target.value)} />
          </div>
          <div className="filter-group">
            <button className="export-btn" onClick={analyze} disabled={loading}>
              <IconSearch size={16} /> Проанализировать
            </button>
          </div>
          {data && data.rows.length > 0 && (
            <div className="filter-group">
              <button className="export-btn" onClick={exportExcel}>
                <IconDownload size={16} /> Экспорт в Excel
              </button>
            </div>
          )}
        </div>
      </div>

      {loading ? (
        <div className="loading"><RossetiLoader /></div>
      ) : error ? (
        <div className="no-data"><p><span className="svg-frame" style={{ marginRight: 8 }}><IconAlertTriangle size={26} /></span>{error}</p></div>
      ) : !data ? (
        <div className="no-data"><p><span className="svg-frame" style={{ marginRight: 8 }}><IconSearch size={26} /></span>Задайте параметры и нажмите «Проанализировать»</p></div>
      ) : data.rows.length === 0 ? (
        <div className="no-data"><p><span className="svg-frame" style={{ marginRight: 8 }}><IconCheck size={26} /></span>За выбранный период закрытых мероприятий нет</p></div>
      ) : (
        <div className="problem-list">
          {data.rows.map(r => {
            const open = !!expanded[r.resId];
            const ratioTxt = r.capacity.ratio > 1 ? ` — превышение в ${r.capacity.ratio.toFixed(1)} раза` : '';
            return (
              <div key={r.resId} className="problem-card">
                <div className="problem-header">
                  <div className="problem-title">
                    <h3>{r.resName}</h3>
                    <span className="res-badge">{DA_LEVEL_LABEL[r.level]}</span>
                  </div>
                  <span className="failure-badge" style={{ background: DA_LEVEL_COLOR[r.level], color: '#fff' }}>
                    Балл {r.score}
                  </span>
                </div>
                <div className="problem-meta">
                  <span>Закрыто <b>{r.totalClosed}</b> мероприятий по <b>{r.uniqueTp}</b> ТП; пик: <b>{r.capacity.maxTpInWindow}</b> ТП за {settings.windowDays} дн. при ёмкости {r.capacity.capacityLimit}{ratioTxt}</span>
                </div>
                <button className="da-toggle" onClick={() => setExpanded(prev => ({ ...prev, [r.resId]: !open }))}>
                  {open ? 'Свернуть детали' : 'Показать детали'}
                  <IconArrowDown size={14} style={{ transform: open ? 'rotate(180deg)' : 'none', transition: 'transform .12s ease' }} />
                </button>
                {open && (
                  <div className="da-details">
                    <div className="detail-section">
                      <h5>Причины оценки</h5>
                      {r.reasons.length ? (
                        <ul className="da-reasons">{r.reasons.map((reason, i) => <li key={i}>{reason}</li>)}</ul>
                      ) : <div className="detail-row">Признаков формального закрытия не обнаружено.</div>}
                    </div>

                    <div className="detail-section">
                      <h5>Пакетные закрытия ({r.batches.length})</h5>
                      {r.batches.length ? (
                        <table className="diag-table">
                          <thead><tr><th>Начало</th><th>Конец</th><th>Закрытий</th><th>Уник. ТП</th><th>Длит., мин</th></tr></thead>
                          <tbody>
                            {r.batches.map((b, i) => (
                              <tr key={i}><td>{dt(b.start)}</td><td>{dt(b.end)}</td><td>{b.count}</td><td>{b.uniqueTp}</td><td>{b.durationMin}</td></tr>
                            ))}
                          </tbody>
                        </table>
                      ) : <div className="detail-row">Нет.</div>}
                    </div>

                    <div className="detail-section">
                      <h5>Шаблонные комментарии ({r.templateComments.length})</h5>
                      {r.templateComments.length ? (
                        <table className="diag-table">
                          <thead><tr><th>Текст</th><th>Повторов</th><th>ТП</th></tr></thead>
                          <tbody>
                            {r.templateComments.map((c, i) => (
                              <tr key={i}><td style={{ textAlign: 'left' }}>{c.text}</td><td>{c.count}</td><td>{c.tpCount}</td></tr>
                            ))}
                          </tbody>
                        </table>
                      ) : <div className="detail-row">Нет.</div>}
                    </div>

                    <div className="detail-section">
                      <h5>Дубли фото ({r.duplicatePhotos.length})</h5>
                      {r.duplicatePhotos.length ? (
                        <table className="diag-table">
                          <thead><tr><th>Файл</th><th>Встречен на ТП</th></tr></thead>
                          <tbody>
                            {r.duplicatePhotos.map((p, i) => (
                              <tr key={i}><td style={{ textAlign: 'left' }}>{p.name}</td><td>{p.tpCount}</td></tr>
                            ))}
                          </tbody>
                        </table>
                      ) : <div className="detail-row">Нет.</div>}
                    </div>

                    <div className="detail-section">
                      <h5>Перепроверки</h5>
                      <div className="detail-row">
                        Перепроверено: <b>{r.recheck.checked}</b> · провалов: <b>{r.recheck.failed}</b> ({pct(r.recheck.failShare)}) · ночных закрытий: <b>{pct(r.afterHoursShare)}</b>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [user, setUser] = useState(null);
  // Пока в iframe ждём/меняем токен платформы — показываем лоадер, а не форму логина.
  const [ssoPending, setSsoPending] = useState(EMBEDDED);
  const [activeSection, setActiveSection] = useState('structure');
  const [selectedRes, setSelectedRes] = useState(null);
  const [resList, setResList] = useState([]);

  // Оптимизированная проверка токена
  useEffect(() => {
    // В iframe платформы не доверяем старому токену из localStorage — ждём
    // свежий токен платформы (иначе мигнёт предыдущий пользователь).
    if (EMBEDDED) {
      localStorage.removeItem('token');
      localStorage.removeItem('fileToken');
      return;
    }
    const token = localStorage.getItem('token');
    if (token) {
      api.get('/api/auth/me')
        .then(response => {
          if (response.data.fileToken) localStorage.setItem('fileToken', response.data.fileToken);
          setUser(response.data.user);
          setSelectedRes(response.data.user.resId);
        })
        .catch(() => {
          try {
            const payload = JSON.parse(atob(token.split('.')[1]));
            setUser({
              id: payload.id,
              role: payload.role,
              resId: payload.resId
            });
            setSelectedRes(payload.resId);
          } catch (error) {
            localStorage.removeItem('token');
            localStorage.removeItem('fileToken');
          }
        });
    }
  }, []);

  // Единый вход через платформу: слушаем токен из iframe и меняем на свою сессию.
  // Контракт platform-auth/app-ready фиксирован платформой — не менять.
  useEffect(() => {
    if (!EMBEDDED) return;
    const exchange = async (kcToken) => {
      try {
        // Чистый fetch (не axios), чтобы 401-интерсептор не мешал.
        const resp = await fetch(`${API_URL}/api/auth/platform`, {
          method: 'POST',
          headers: { Authorization: `Bearer ${kcToken}` }
        });
        if (!resp.ok) throw new Error('sso failed');
        const data = await resp.json();
        localStorage.setItem('token', data.token);
        if (data.fileToken) localStorage.setItem('fileToken', data.fileToken);
        setUser(data.user);
        if (data.user?.resId) setSelectedRes(data.user.resId);
      } catch {
        localStorage.removeItem('token');
        localStorage.removeItem('fileToken');
        setUser(null); // упадём на обычную форму логина
      } finally {
        setSsoPending(false);
      }
    };
    const onMessage = (event) => {
      if (event.origin !== PLATFORM_ORIGIN) return; // доверяем только платформе
      const d = event.data;
      if (d && d.type === 'platform-auth' && d.token) exchange(d.token);
    };
    window.addEventListener('message', onMessage);
    // Сообщаем платформе, что готовы принять токен (она ответит platform-auth).
    window.parent.postMessage({ type: 'app-ready' }, PLATFORM_ORIGIN);
    // Токен так и не пришёл за 5с обычная форма логина (fallback).
    const timer = setTimeout(() => setSsoPending(false), 5000);
    return () => { window.removeEventListener('message', onMessage); clearTimeout(timer); };
  }, []);

  useEffect(() => {
    if (user && user.role === 'admin') {
      loadResList();
    }
  }, [user]);

  const loadResList = async () => {
    try {
      const response = await api.get('/api/res/list');
      setResList(response.data);
    } catch (error) {
      console.error('Error loading RES list:', error);
    }
  };

  const handleLogin = useCallback((userData) => {
    setUser({
      id: userData.id,
      fio: userData.fio,
      role: userData.role,
      resId: userData.resId,
      resName: userData.resName
    });
    if (userData.resId) {
      setSelectedRes(userData.resId);
    }
  }, []);

  useEffect(() => {
  let inactivityTimer;
  let lastReset = 0;
  const INACTIVITY_TIME = 2 * 60 * 60 * 1000; // 2 часа
  const RESET_THROTTLE = 30 * 1000; // PERF: сбрасываем таймер не чаще раза в 30 сек
  
  const resetTimer = () => {
    // PERF: раньше clearTimeout/setTimeout дёргались на КАЖДЫЙ mousemove/scroll
    // (тысячи раз в секунду) — вкладка подтормаживала. Теперь троттлинг.
    const now = Date.now();
    if (now - lastReset < RESET_THROTTLE) return;
    lastReset = now;
    
    clearTimeout(inactivityTimer);
    inactivityTimer = setTimeout(() => {
      localStorage.removeItem('token');
      localStorage.removeItem('fileToken');
      setUser(null);
      alert('Сессия истекла из-за неактивности. Пожалуйста, войдите снова.');
    }, INACTIVITY_TIME);
  };
  
  const events = ['mousedown', 'mousemove', 'keypress', 'scroll', 'touchstart', 'click'];
  
  if (user) {
    lastReset = 0; // первый сброс — сразу
    resetTimer();
    events.forEach(event => {
      document.addEventListener(event, resetTimer, { passive: true });
    });
  }
  
  return () => {
    clearTimeout(inactivityTimer);
    events.forEach(event => {
      document.removeEventListener(event, resetTimer);
    });
  };
}, [user]);

  const handleLogout = () => {
    localStorage.removeItem('token');
    localStorage.removeItem('fileToken');
    setUser(null);
    setSelectedRes(null);
  };

  if (!user) {
    // Внутри iframe платформы НЕ показываем свою форму логина: пока ждём токен —
    // заставка; если вход не удался (нет доступа/учётки) — сообщение, а не логин.
    if (EMBEDDED) {
      if (ssoPending) {
        return (
          <div className="login-container">
            <div className="login-box" style={{ textAlign: 'center' }}>
              <RossetiLoader />
              <h2 style={{ marginTop: 16 }}>Вход через платформу…</h2>
            </div>
          </div>
        );
      }
      return (
        <div className="login-container">
          <div className="login-box" style={{ textAlign: 'center' }}>
            <h2>Нет доступа</h2>
            <p style={{ marginTop: 10, color: '#5b6b82' }}>
              У вашей учётной записи нет доступа к этому приложению.<br />
              Обратитесь к администратору.
            </p>
          </div>
        </div>
      );
    }
    return <LoginForm onLogin={handleLogin} />;
  }

const renderContent = () => {
  switch (activeSection) {
    case 'structure':
      return <NetworkStructure onSectionChange={setActiveSection} />;
    case 'poll_map':
      return <PollMap selectedRes={selectedRes} />;
    case 'upload':
      return <FileUpload />;
    case 'tech_pending':
      return <Notifications filterType="error" onSectionChange={setActiveSection} selectedRes={selectedRes} />;
    case 'askue_pending':
      return <Notifications filterType="pending_askue" onSectionChange={setActiveSection} selectedRes={selectedRes} />;
    case 'problem_vl':
      return <ProblemVL selectedRes={selectedRes} />;
    case 'power_overload':
      return <PowerOverload selectedRes={selectedRes} />;
    case 'power_analysis':
      return <PowerAnalysis selectedRes={selectedRes} />;
    case 'documents':
      return <UploadedDocuments />;
    case 'reports':
      return <Reports />;
    case 'settings':
      return <Settings />;
    case 'permissions':
      return <PermissionsAdmin />;
    case 'history':
      return <SystemHistory />;
    case 'analytics':
      return <Analytics />;
    case 'deep_analysis':
      return <DeepWorkAnalysis />;
    default:
      return <NetworkStructure />;
  }
};

  return (
    <AuthContext.Provider value={{ user, selectedRes }}>
      <ModalDockProvider>
      <div className="app">
        <MainMenu
          activeSection={activeSection}
          onSectionChange={setActiveSection}
          userRole={user.role}
          isSuper={user.isSuper}
        />
        
        <div className="main-content">
          <header className="app-header">
            <div className="header-left">
              <h1>Система контроля уровня напряжения в сетях 0,4 кВ</h1>
              {user.role === 'admin' && activeSection !== 'history' && activeSection !== 'analytics' && activeSection !== 'deep_analysis' && (
                <select 
                  value={selectedRes || ''}
                  onChange={(e) => setSelectedRes(e.target.value ? parseInt(e.target.value) : null)}
                >
                  <option value="">Все РЭСы</option>
                  {resList.map(res => (
                    <option key={res.id} value={res.id}>{res.name}</option>
                  ))}
                </select>
              )}
              {user.resId && (
                <span className="res-name">
                  {resList.find(r => r.id === user.resId)?.name || user.resName}
                </span>
              )}
            </div>
            
            <div className="header-right">
              <span>{user.fio}</span>
              <span className="user-role">
                ({user.role === 'admin' ? 'Администратор' :
                  user.role === 'uploader' ? 'Загрузчик' : 'Ответственный'})
              </span>
            </div>
          </header>
          
          <main className="content">
            {renderContent()}
          </main>
        </div>
      </div>
      </ModalDockProvider>
    </AuthContext.Provider>
  );
}
