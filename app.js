require('dotenv').config();
const express = require('express');
const session = require('express-session');
const sqlite3 = require('sqlite3').verbose();
const multer = require('multer');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;
const LOG_ARCHIVE_DAYS = parseInt(process.env.LOG_ARCHIVE_DAYS || '180', 10) || 180;
const LOG_ARCHIVE_ENABLED = String(process.env.LOG_ARCHIVE_ENABLED || 'true').toLowerCase() !== 'false';
const LOG_ARCHIVE_INTERVAL_MS = 24 * 60 * 60 * 1000;
const EXPORT_DIR = path.join(__dirname, 'exports');

// Database setup - persistent disk on Render, local in development
const DB_PATH = process.env.NODE_ENV === 'production' ? '/data/publications.db' : './publications.db';
const db = new sqlite3.Database(DB_PATH, (err) => {
  if (err) {
    console.error('Error opening database:', err.message);
  } else {
    console.log('Connected to SQLite database.');
    initDatabase();
  }
});

// Initialize database tables
function initDatabase() {
  db.serialize(() => {
    db.run(`CREATE TABLE IF NOT EXISTS customers (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      contact_name TEXT NOT NULL,
      company TEXT NOT NULL,
      customer_id TEXT NOT NULL,
      email TEXT NOT NULL,
      cc_emails TEXT,
      products TEXT,
      markets TEXT,
      content_types TEXT,
      regions TEXT,
      customer_type TEXT,
      subscription_tier TEXT,
      preferred_frequency TEXT DEFAULT 'Immediate',
      status TEXT DEFAULT 'Active',
      date_added DATETIME DEFAULT CURRENT_TIMESTAMP,
      last_notified DATETIME
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS publications (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      title TEXT NOT NULL,
      publication_number TEXT NOT NULL,
      products TEXT,
      markets TEXT,
      content_type TEXT,
      regions TEXT,
      urgency TEXT,
      summary TEXT,
      action_required TEXT,
      author_name TEXT,
      reviewer TEXT,
      distribution_status TEXT DEFAULT 'Draft',
      date_published DATETIME,
      recipients_count INTEGER DEFAULT 0,
      file_path TEXT,
      file_name TEXT,
      is_archived INTEGER DEFAULT 0,
      archived_at DATETIME
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS distribution_logs (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      publication_number TEXT,
      publication_title TEXT,
      content_type TEXT,
      urgency TEXT,
      recipient_name TEXT,
      recipient_company TEXT,
      recipient_email TEXT,
      sent_date DATETIME DEFAULT CURRENT_TIMESTAMP,
      delivery_status TEXT DEFAULT 'Sent',
      acknowledged INTEGER DEFAULT 0,
      acknowledgment_date DATETIME,
      match_reason TEXT,
      is_archived INTEGER DEFAULT 0,
      archived_at DATETIME
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS metadata (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      category TEXT NOT NULL,
      value TEXT NOT NULL,
      sort_order INTEGER DEFAULT 0,
      is_active INTEGER DEFAULT 1,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )`);

    function addColumnIfMissing(table, column, definition) {
      db.all(`PRAGMA table_info(${table})`, [], (err, rows) => {
        if (err) return;
        const hasColumn = (rows || []).some(row => row.name === column);
        if (!hasColumn) {
          db.run(`ALTER TABLE ${table} ADD COLUMN ${column} ${definition}`);
        }
      });
    }

    addColumnIfMissing('publications', 'is_archived', 'INTEGER DEFAULT 0');
    addColumnIfMissing('publications', 'archived_at', 'DATETIME');
    addColumnIfMissing('distribution_logs', 'is_archived', 'INTEGER DEFAULT 0');
    addColumnIfMissing('distribution_logs', 'archived_at', 'DATETIME');

    // Migrate old "Comprehensive" tier to "All Announcements"
    db.run("UPDATE customers SET subscription_tier = 'All Announcements' WHERE subscription_tier = 'Comprehensive'");

    // Seed metadata if empty
    db.get('SELECT COUNT(*) as count FROM metadata', [], (err, row) => {
      if (err || (row && row.count === 0)) {
        const seeds = [
          ['product', 'All Products', 0],
          ['product', '8.8L GSI', 1],
          ['product', '8.8L DSI', 2],
          ['product', '22L DSI', 3],
          ['product', '4.3L GSI', 4],
          ['product', '6.0L GSI', 5],
          ['product', '3.0L GSI', 6],
          ['product', '2.4L GSI', 7],
          ['product', '8.8L LPG', 8],
          ['market', 'All Markets', 0],
          ['market', 'Power Systems', 1],
          ['market', 'Industrial', 2],
          ['market', 'On-Road', 3],
          ['market', 'Material Handling', 4],
          ['market', 'Specialty', 5],
          ['market', 'Marine', 6],
          ['market', 'Oil & Gas', 7],
          ['market', 'Agriculture', 8],
          ['content_type', 'All Content Types', 0],
          ['content_type', 'Service Bulletin', 1],
          ['content_type', 'Notice of Change', 2],
          ['content_type', 'Manual Update', 3],
          ['content_type', 'Safety Notice', 4],
          ['content_type', 'Product Alert', 5],
          ['content_type', 'Recall Notice', 6],
          ['content_type', 'Technical Tip', 7],
          ['content_type', 'Training Notice', 8],
          ['content_type', 'Product Announcement', 9],
          ['region', 'Global', 0],
          ['region', 'North America', 1],
          ['region', 'EMEA', 2],
          ['region', 'APAC', 3],
          ['region', 'LATAM', 4],
        ];
        const stmt = db.prepare('INSERT INTO metadata (category, value, sort_order) VALUES (?, ?, ?)');
        seeds.forEach(s => stmt.run(s));
        stmt.finalize();
        console.log('Seeded default metadata values.');
      }
    });
  });
  scheduleLogArchive();
}

// Helper: fetch active metadata grouped by category
function getMetadata() {
  return new Promise((resolve, reject) => {
    db.all('SELECT * FROM metadata WHERE is_active = 1 ORDER BY category, sort_order, value', [], (err, rows) => {
      if (err) return reject(err);
      const grouped = { products: [], markets: [], content_types: [], regions: [] };
      (rows || []).forEach(row => {
        if (row.category === 'product') grouped.products.push(row.value);
        else if (row.category === 'market') grouped.markets.push(row.value);
        else if (row.category === 'content_type') grouped.content_types.push(row.value);
        else if (row.category === 'region') grouped.regions.push(row.value);
      });
      resolve(grouped);
    });
  });
}

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

// Helper: promisified db queries
function dbAll(sql, params) {
  return new Promise((resolve, reject) => {
    db.all(sql, params || [], (err, rows) => err ? reject(err) : resolve(rows || []));
  });
}
function dbGet(sql, params) {
  return new Promise((resolve, reject) => {
    db.get(sql, params || [], (err, row) => err ? reject(err) : resolve(row));
  });
}
function dbRun(sql, params) {
  return new Promise((resolve, reject) => {
    db.run(sql, params || [], function(err) { err ? reject(err) : resolve(this); });
  });
}

function toSortDir(value, fallback = 'asc') {
  const normalized = String(value || '').toLowerCase();
  if (normalized === 'asc' || normalized === 'desc') return normalized;
  return fallback;
}

function parseBoolean(value) {
  return ['1', 'true', 'yes', 'on'].includes(String(value || '').toLowerCase());
}

function clampNumber(value, min, max, fallback) {
  const parsed = parseInt(value, 10);
  if (!Number.isFinite(parsed)) return fallback;
  return Math.min(Math.max(parsed, min), max);
}

function buildPublicationFilters(filters) {
  const whereParams = [];
  let whereSql = 'WHERE 1=1';
  if (filters.search) {
    whereSql += ' AND (title LIKE ? OR publication_number LIKE ?)';
    const s = `%${filters.search}%`;
    whereParams.push(s, s);
  }
  if (filters.status) {
    whereSql += ' AND distribution_status = ?';
    whereParams.push(filters.status);
  }
  if (filters.urgency) {
    whereSql += ' AND urgency = ?';
    whereParams.push(filters.urgency);
  }
  return { whereSql, whereParams };
}

function buildLogFilters(filters) {
  const whereParams = [];
  let whereSql = 'WHERE 1=1';
  if (filters.search) {
    whereSql += ' AND (publication_number LIKE ? OR publication_title LIKE ? OR recipient_name LIKE ? OR recipient_company LIKE ? OR recipient_email LIKE ?)';
    const s = `%${filters.search}%`;
    whereParams.push(s, s, s, s, s);
  }
  if (filters.urgency) {
    whereSql += ' AND urgency = ?';
    whereParams.push(filters.urgency);
  }
  return { whereSql, whereParams };
}

async function archiveOldLogs() {
  if (!LOG_ARCHIVE_ENABLED) return;
  try {
    ensureDir(EXPORT_DIR);
    const cutoff = `-${LOG_ARCHIVE_DAYS} days`;
    const logs = await dbAll(
      `SELECT * FROM distribution_logs
       WHERE (is_archived IS NULL OR is_archived = 0)
         AND sent_date <= datetime('now', ?)`,
      [cutoff]
    );
    if (!logs.length) return;

    const fileDate = new Date().toISOString().slice(0, 10);
    const fileName = `distribution-logs-archive-${fileDate}.xlsx`;
    const filePath = path.join(EXPORT_DIR, fileName);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Archived Logs');
    sheet.columns = [
      { header: 'Sent Date', key: 'sent_date', width: 22 },
      { header: 'Publication #', key: 'publication_number', width: 16 },
      { header: 'Title', key: 'publication_title', width: 40 },
      { header: 'Type', key: 'content_type', width: 18 },
      { header: 'Urgency', key: 'urgency', width: 16 },
      { header: 'Recipient', key: 'recipient_name', width: 22 },
      { header: 'Company', key: 'recipient_company', width: 24 },
      { header: 'Email', key: 'recipient_email', width: 28 },
      { header: 'Status', key: 'delivery_status', width: 14 },
      { header: 'Match Reason', key: 'match_reason', width: 40 }
    ];
    logs.forEach(log => {
      sheet.addRow({
        sent_date: log.sent_date ? new Date(log.sent_date).toLocaleString() : '',
        publication_number: log.publication_number,
        publication_title: log.publication_title,
        content_type: log.content_type,
        urgency: log.urgency,
        recipient_name: log.recipient_name,
        recipient_company: log.recipient_company,
        recipient_email: log.recipient_email,
        delivery_status: log.delivery_status,
        match_reason: log.match_reason || ''
      });
    });
    await workbook.xlsx.writeFile(filePath);

    await dbRun(
      `UPDATE distribution_logs
       SET is_archived = 1, archived_at = datetime('now')
       WHERE (is_archived IS NULL OR is_archived = 0)
         AND sent_date <= datetime('now', ?)`,
      [cutoff]
    );
    console.log(`Archived ${logs.length} log(s) to ${filePath}`);
  } catch (err) {
    console.error('Archive old logs error:', err);
  }
}

function scheduleLogArchive() {
  if (!LOG_ARCHIVE_ENABLED) return;
  ensureDir(EXPORT_DIR);
  archiveOldLogs();
  setInterval(archiveOldLogs, LOG_ARCHIVE_INTERVAL_MS);
}

// Middleware
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static('public'));
app.use('/bootstrap', express.static(path.join(__dirname, 'node_modules/bootstrap/dist')));
app.use('/bootstrap-icons', express.static(path.join(__dirname, 'node_modules/bootstrap-icons')));
app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));

// Trust proxy (Render, Heroku, etc.) so secure cookies work behind HTTPS reverse proxy
app.set('trust proxy', 1);

// Session setup
app.use(session({
  secret: process.env.SESSION_SECRET || crypto.randomBytes(32).toString('hex'),
  resave: false,
  saveUninitialized: false,
  cookie: {
    secure: process.env.NODE_ENV === 'production',
    httpOnly: true,
    maxAge: 24 * 60 * 60 * 1000 // 24 hours
  }
}));

function getBaseUrl() {
  return process.env.BASE_URL || process.env.RENDER_EXTERNAL_URL || 'http://localhost:3000';
}

// Auth middleware — protects all admin routes
function requireAuth(req, res, next) {
  if (req.session && req.session.authenticated) {
    return next();
  }
  res.redirect('/login');
}

// Login routes (public)
app.get('/login', (req, res) => {
  res.send(`<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Login - PSI Publications</title>
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="/bootstrap/css/bootstrap.min.css">
  <link rel="stylesheet" href="/bootstrap-icons/font/bootstrap-icons.min.css">
  <link rel="stylesheet" href="/styles.css">
</head>
<body style="background: var(--gray-50); display: flex; align-items: center; justify-content: center; min-height: 100vh;">
  <div style="width: 100%; max-width: 400px; padding: 20px;">
    <div style="text-align: center; margin-bottom: 24px;">
      <img src="/images/psi-logo.svg" alt="PSI" style="height: 48px; margin-bottom: 12px;">
      <h1 style="font-size: 20px; font-weight: 700; color: var(--gray-900); margin: 0;">Admin Login</h1>
      <p style="font-size: 13px; color: var(--gray-500); margin: 6px 0 0;">Publication Distribution System</p>
    </div>
    ${req.query.error ? '<div style="background: var(--danger-light); color: #721c24; padding: 10px 16px; border-radius: 8px; margin-bottom: 16px; font-size: 13px; text-align: center;"><i class="bi bi-exclamation-triangle"></i> Invalid username or password</div>' : ''}
    <div class="form-container">
      <form action="/login" method="POST" style="padding: 24px;">
        <div class="form-group">
          <label><i class="bi bi-person"></i> Username</label>
          <input type="text" name="username" required autocomplete="username" placeholder="Enter username">
        </div>
        <div class="form-group">
          <label><i class="bi bi-lock"></i> Password</label>
          <input type="password" name="password" required autocomplete="current-password" placeholder="Enter password">
        </div>
        <button type="submit" class="btn btn-primary" style="width: 100%; margin-top: 8px;"><i class="bi bi-box-arrow-in-right"></i> Sign In</button>
      </form>
    </div>
  </div>
</body>
</html>`);
});

app.post('/login', (req, res) => {
  const { username, password } = req.body;
  const adminUser = process.env.ADMIN_USER || 'admin';
  const adminPass = process.env.ADMIN_PASS || 'changeme';

  if (username === adminUser && password === adminPass) {
    req.session.authenticated = true;
    res.redirect('/');
  } else {
    res.redirect('/login?error=1');
  }
});

app.get('/logout', (req, res) => {
  req.session.destroy();
  res.redirect('/login');
});

// Protect all admin routes
app.use((req, res, next) => {
  // Public routes that don't need auth
  const publicPaths = ['/login', '/subscribe', '/unsubscribe', '/health'];
  if (publicPaths.some(p => req.path === p || req.path.startsWith(p + '?'))) {
    return next();
  }
  // Static assets are already served above
  requireAuth(req, res, next);
});

// File upload setup - persistent disk on Render, local in development
const UPLOAD_DIR = process.env.NODE_ENV === 'production' ? '/data/uploads' : 'uploads';
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    if (!fs.existsSync(UPLOAD_DIR)) {
      fs.mkdirSync(UPLOAD_DIR, { recursive: true });
    }
    cb(null, UPLOAD_DIR);
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});
const upload = multer({ storage: storage });

// Email transporter - configured from environment variables
let transporter = null;

function isSmtpConfigured() {
  return !!(process.env.SMTP_HOST && process.env.SMTP_USER && process.env.SMTP_PASS);
}

if (isSmtpConfigured()) {
  transporter = nodemailer.createTransport({
    host: process.env.SMTP_HOST,
    port: parseInt(process.env.SMTP_PORT) || 587,
    secure: parseInt(process.env.SMTP_PORT) === 465,
    auth: {
      user: process.env.SMTP_USER,
      pass: process.env.SMTP_PASS
    }
  });
  console.log(`SMTP configured: ${process.env.SMTP_HOST}:${process.env.SMTP_PORT || 587}`);
} else {
  console.log('SMTP not configured. Emails will be logged to console only.');
}

// ============================================================
// ROUTES
// ============================================================

// Dashboard
app.get('/', async (req, res) => {
  try {
    const totalCustomers = (await dbGet('SELECT COUNT(*) as c FROM customers')).c;
    const activeCustomers = (await dbGet("SELECT COUNT(*) as c FROM customers WHERE status = 'Active'")).c;
    const totalPublications = (await dbGet('SELECT COUNT(*) as c FROM publications')).c;
    const distributedPublications = (await dbGet("SELECT COUNT(*) as c FROM publications WHERE distribution_status = 'Distributed'")).c;
    const totalLogs = (await dbGet('SELECT COUNT(*) as c FROM distribution_logs')).c;
    const recentPublications = await dbAll('SELECT * FROM publications ORDER BY id DESC LIMIT 5');
    const recentLogs = await dbAll('SELECT * FROM distribution_logs ORDER BY sent_date DESC LIMIT 10');

    res.render('index', {
      stats: { totalCustomers, activeCustomers, totalPublications, distributedPublications, totalLogs },
      recentPublications,
      recentLogs
    });
  } catch (err) {
    console.error('Dashboard error:', err);
    res.status(500).send('Server error');
  }
});

// ============================================================
// CUSTOMER MANAGEMENT
// ============================================================

app.get('/customers', async (req, res) => {
  try {
    const { search, status, tier, sort, dir } = req.query;
    const customerSortColumns = {
      contact_name: 'contact_name',
      company: 'company',
      customer_id: 'customer_id',
      email: 'email',
      customer_type: 'customer_type',
      subscription_tier: 'subscription_tier',
      status: 'status',
      date_added: 'date_added'
    };
    const sortBy = customerSortColumns[sort] ? sort : 'company';
    const sortDir = toSortDir(dir, 'asc');

    let sql = 'SELECT * FROM customers WHERE 1=1';
    const params = [];

    if (search) {
      sql += ' AND (contact_name LIKE ? OR company LIKE ? OR email LIKE ? OR customer_id LIKE ?)';
      const s = `%${search}%`;
      params.push(s, s, s, s);
    }
    if (status) {
      sql += ' AND status = ?';
      params.push(status);
    }
    if (tier) {
      sql += ' AND subscription_tier = ?';
      params.push(tier);
    }
    sql += ` ORDER BY ${customerSortColumns[sortBy]} ${sortDir.toUpperCase()}, contact_name ASC`;

    const customers = await dbAll(sql, params);
    res.render('customers', {
      customers,
      search: search || '',
      statusFilter: status || '',
      tierFilter: tier || '',
      sortBy,
      sortDir,
      success: req.query.success || '',
      error: req.query.error || ''
    });
  } catch (err) {
    console.error('Customers error:', err);
    res.status(500).send('Server error');
  }
});

app.get('/customers/export', async (req, res) => {
  try {
    const { search, status, tier, sort, dir } = req.query;
    const customerSortColumns = {
      contact_name: 'contact_name',
      company: 'company',
      customer_id: 'customer_id',
      email: 'email',
      customer_type: 'customer_type',
      subscription_tier: 'subscription_tier',
      status: 'status',
      date_added: 'date_added'
    };
    const sortBy = customerSortColumns[sort] ? sort : 'company';
    const sortDir = toSortDir(dir, 'asc');

    let sql = 'SELECT * FROM customers WHERE 1=1';
    const params = [];
    if (search) {
      sql += ' AND (contact_name LIKE ? OR company LIKE ? OR email LIKE ? OR customer_id LIKE ?)';
      const s = `%${search}%`;
      params.push(s, s, s, s);
    }
    if (status) { sql += ' AND status = ?'; params.push(status); }
    if (tier) { sql += ' AND subscription_tier = ?'; params.push(tier); }
    sql += ` ORDER BY ${customerSortColumns[sortBy]} ${sortDir.toUpperCase()}, contact_name ASC`;
    const customers = await dbAll(sql, params);

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Customers');

    sheet.columns = [
      { header: 'Customer ID', key: 'customer_id', width: 16 },
      { header: 'Contact Name', key: 'contact_name', width: 22 },
      { header: 'Company', key: 'company', width: 24 },
      { header: 'Email', key: 'email', width: 28 },
      { header: 'CC Emails', key: 'cc_emails', width: 28 },
      { header: 'Type', key: 'customer_type', width: 14 },
      { header: 'Subscription Tier', key: 'subscription_tier', width: 18 },
      { header: 'Status', key: 'status', width: 10 },
      { header: 'Products', key: 'products', width: 30 },
      { header: 'Markets', key: 'markets', width: 24 },
      { header: 'Content Types', key: 'content_types', width: 24 },
      { header: 'Regions', key: 'regions', width: 20 },
      { header: 'Date Added', key: 'date_added', width: 18 },
      { header: 'Last Notified', key: 'last_notified', width: 18 }
    ];

    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D32' } };

    customers.forEach(c => {
      sheet.addRow({
        customer_id: c.customer_id || ((c.customer_type || '').toLowerCase() === 'internal' ? 'Internal Employee' : ''),
        contact_name: c.contact_name,
        company: c.company,
        email: c.email,
        cc_emails: c.cc_emails || '',
        customer_type: c.customer_type || '',
        subscription_tier: c.subscription_tier || '',
        status: c.status || '',
        products: c.products || '',
        markets: c.markets || '',
        content_types: c.content_types || '',
        regions: c.regions || '',
        date_added: c.date_added ? new Date(c.date_added).toLocaleDateString() : '',
        last_notified: c.last_notified ? new Date(c.last_notified).toLocaleDateString() : ''
      });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=customers-' + new Date().toISOString().slice(0, 10) + '.xlsx');
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('Customer export error:', err);
    res.status(500).send('Export failed');
  }
});

app.post('/customers/import', upload.single('importFile'), async (req, res) => {
  try {
    if (!req.file) {
      return res.redirect('/customers?error=' + encodeURIComponent('No import file selected'));
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(req.file.path);
    const sheet = workbook.worksheets[0];
    if (!sheet) {
      fs.unlinkSync(req.file.path);
      return res.redirect('/customers?error=' + encodeURIComponent('Import file is empty'));
    }

    const headerMap = {};
    const headerAliases = {
      'customer id': 'customer_id',
      'contact name': 'contact_name',
      'company': 'company',
      'email': 'email',
      'cc emails': 'cc_emails',
      'cc email': 'cc_emails',
      'type': 'customer_type',
      'customer type': 'customer_type',
      'subscription tier': 'subscription_tier',
      'status': 'status',
      'products': 'products',
      'markets': 'markets',
      'content types': 'content_types',
      'content type': 'content_types',
      'regions': 'regions'
    };

    const headerRow = sheet.getRow(1);
    headerRow.eachCell((cell, colNumber) => {
      const text = String(cell.text || cell.value || '').trim().toLowerCase();
      if (!text) return;
      const key = headerAliases[text] || text;
      headerMap[key] = colNumber;
    });

    const getCell = (row, key) => {
      const col = headerMap[key];
      if (!col) return '';
      const cell = row.getCell(col);
      return String(cell.text || cell.value || '').trim();
    };

    const normalizeList = (value) => {
      if (!value) return '';
      const parts = String(value)
        .split(/[;,]+/)
        .map(v => v.trim())
        .filter(Boolean);
      return parts.join('; ');
    };

    const normalizeStatus = (value) => {
      const raw = String(value || '').trim().toLowerCase();
      if (!raw) return 'Active';
      if (raw.startsWith('inact')) return 'Inactive';
      if (raw.startsWith('susp')) return 'Suspended';
      return 'Active';
    };

    const normalizeCustomerType = (value) => {
      const raw = String(value || '').trim().toLowerCase();
      if (!raw) return 'End User';
      if (raw === 'internal') return 'Internal';
      if (raw === 'oem') return 'OEM';
      if (raw === 'dealer') return 'Dealer';
      if (raw === 'distributor') return 'Distributor';
      if (raw === 'end user' || raw === 'enduser') return 'End User';
      return value.trim();
    };

    const normalizeTier = (value) => {
      const raw = String(value || '').trim().toLowerCase();
      if (!raw) return 'All Announcements';
      if (raw.startsWith('essential')) return 'Essential';
      if (raw.startsWith('standard')) return 'Standard';
      if (raw.includes('all')) return 'All Announcements';
      if (raw.includes('comprehensive')) return 'All Announcements';
      return value.trim();
    };

    let inserted = 0;
    let updated = 0;
    let skipped = 0;

    await dbRun('BEGIN TRANSACTION');
    for (let i = 2; i <= sheet.rowCount; i += 1) {
      const row = sheet.getRow(i);
      if (!row || row.actualCellCount === 0) continue;

      const contact_name = getCell(row, 'contact_name');
      const company = getCell(row, 'company');
      const email = getCell(row, 'email');
      const rawCustomerId = getCell(row, 'customer_id');
      const customer_type = normalizeCustomerType(getCell(row, 'customer_type'));
      const isInternal = customer_type.toLowerCase() === 'internal';
      const customer_id = isInternal ? '' : (rawCustomerId && rawCustomerId.toLowerCase() === 'internal employee' ? '' : rawCustomerId);

      if (!contact_name || !company || !email) {
        skipped += 1;
        console.warn(`Import row ${i} skipped: missing contact name, company, or email`);
        continue;
      }
      if (!isInternal && !customer_id) {
        skipped += 1;
        console.warn(`Import row ${i} skipped: customer ID required for non-internal customers`);
        continue;
      }

      const cc_emails = getCell(row, 'cc_emails');
      const subscription_tier = normalizeTier(getCell(row, 'subscription_tier'));
      const status = normalizeStatus(getCell(row, 'status'));
      const products = normalizeList(getCell(row, 'products'));
      const markets = normalizeList(getCell(row, 'markets'));
      const content_types = normalizeList(getCell(row, 'content_types'));
      const regions = normalizeList(getCell(row, 'regions'));

      let existing = null;
      if (customer_id && !isInternal) {
        existing = await dbGet('SELECT * FROM customers WHERE customer_id = ?', [customer_id]);
      }
      if (!existing) {
        existing = await dbGet('SELECT * FROM customers WHERE email = ? AND company = ?', [email, company]);
      }

      if (existing) {
        await dbRun(
          `UPDATE customers SET contact_name=?, company=?, customer_id=?, email=?, cc_emails=?, products=?, markets=?, content_types=?, regions=?, customer_type=?, subscription_tier=?, status=? WHERE id=?`,
          [
            contact_name || existing.contact_name,
            company || existing.company,
            isInternal ? '' : (customer_id || existing.customer_id),
            email || existing.email,
            cc_emails || existing.cc_emails || '',
            products || existing.products || '',
            markets || existing.markets || '',
            content_types || existing.content_types || '',
            regions || existing.regions || '',
            customer_type || existing.customer_type,
            subscription_tier || existing.subscription_tier,
            status || existing.status,
            existing.id
          ]
        );
        updated += 1;
      } else {
        await dbRun(
          `INSERT INTO customers (contact_name, company, customer_id, email, cc_emails, products, markets, content_types, regions, customer_type, subscription_tier, status)
           VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
          [
            contact_name,
            company,
            isInternal ? '' : customer_id,
            email,
            cc_emails || '',
            products || '',
            markets || '',
            content_types || '',
            regions || '',
            customer_type,
            subscription_tier,
            status
          ]
        );
        inserted += 1;
      }
    }
    await dbRun('COMMIT');

    fs.unlinkSync(req.file.path);
    const message = `Import complete: ${inserted} added, ${updated} updated, ${skipped} skipped.`;
    res.redirect('/customers?success=' + encodeURIComponent(message));
  } catch (err) {
    try {
      await dbRun('ROLLBACK');
    } catch (rollbackErr) {
      console.error('Import rollback error:', rollbackErr);
    }
    if (req.file && req.file.path && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path);
    }
    console.error('Customer import error:', err);
    res.redirect('/customers?error=' + encodeURIComponent('Customer import failed'));
  }
});

app.get('/customers/new', async (req, res) => {
  try {
    const metadata = await getMetadata();
    res.render('customer_form', { customer: null, metadata });
  } catch (err) {
    res.status(500).send('Server error');
  }
});

app.post('/customers', (req, res) => {
  const { contact_name, company, customer_id, email, cc_emails, customer_type, subscription_tier, preferred_frequency, status } = req.body;
  const normalizedCustomerType = (customer_type || 'End User').trim() || 'End User';
  const isInternalEmployee = normalizedCustomerType.toLowerCase() === 'internal';
  const normalizedCustomerId = isInternalEmployee ? '' : String(customer_id || '').trim();
  if (!isInternalEmployee && !normalizedCustomerId) {
    return res.redirect('/customers?error=' + encodeURIComponent('Customer ID is required unless customer type is Internal'));
  }

  const products = Array.isArray(req.body.products) ? req.body.products.join('; ') : req.body.products || '';
  const markets = Array.isArray(req.body.markets) ? req.body.markets.join('; ') : req.body.markets || '';
  const content_types = Array.isArray(req.body.content_types) ? req.body.content_types.join('; ') : req.body.content_types || '';
  const regions = Array.isArray(req.body.regions) ? req.body.regions.join('; ') : req.body.regions || '';

  db.run(`INSERT INTO customers (contact_name, company, customer_id, email, cc_emails, products, markets, content_types, regions, customer_type, subscription_tier, preferred_frequency, status)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [contact_name, company, normalizedCustomerId, email, cc_emails, products, markets, content_types, regions, normalizedCustomerType, subscription_tier, preferred_frequency, status],
    function(err) {
      if (err) {
        console.error('Create customer error:', err);
        return res.redirect('/customers?error=' + encodeURIComponent('Error creating customer'));
      }
      res.redirect('/customers?success=' + encodeURIComponent('Customer created successfully'));
    });
});

app.post('/customers/delete-bulk', express.json(), async (req, res) => {
  try {
    const ids = Array.isArray(req.body.ids)
      ? req.body.ids.map(id => parseInt(id, 10)).filter(Number.isInteger)
      : [];
    if (!ids.length) {
      return res.status(400).json({ error: 'No customers selected' });
    }
    const placeholders = ids.map(() => '?').join(',');
    const rows = await dbAll(`SELECT id, customer_type FROM customers WHERE id IN (${placeholders})`, ids);
    const deletable = rows.filter(row => String(row.customer_type || '').toLowerCase() !== 'internal');
    const skipped = rows.length - deletable.length;
    if (!deletable.length) {
      return res.status(400).json({ error: 'Internal customers cannot be deleted' });
    }
    const deleteIds = deletable.map(row => row.id);
    const deletePlaceholders = deleteIds.map(() => '?').join(',');
    const result = await dbRun(`DELETE FROM customers WHERE id IN (${deletePlaceholders})`, deleteIds);
    res.json({ success: true, count: result.changes || deleteIds.length, skipped });
  } catch (err) {
    console.error('Delete customers error:', err);
    res.status(500).json({ error: 'Failed to delete customers' });
  }
});

app.get('/customers/:id/edit', async (req, res) => {
  try {
    const metadata = await getMetadata();
    const customer = await dbGet('SELECT * FROM customers WHERE id = ?', [req.params.id]);
    if (!customer) return res.status(404).send('Customer not found');
    res.render('customer_form', { customer, metadata });
  } catch (err) {
    res.status(500).send('Server error');
  }
});

app.post('/customers/:id', (req, res) => {
  const id = req.params.id;
  const { contact_name, company, customer_id, email, cc_emails, customer_type, subscription_tier, preferred_frequency, status } = req.body;
  const normalizedCustomerType = (customer_type || 'End User').trim() || 'End User';
  const isInternalEmployee = normalizedCustomerType.toLowerCase() === 'internal';
  const normalizedCustomerId = isInternalEmployee ? '' : String(customer_id || '').trim();
  if (!isInternalEmployee && !normalizedCustomerId) {
    return res.redirect('/customers?error=' + encodeURIComponent('Customer ID is required unless customer type is Internal'));
  }

  const products = Array.isArray(req.body.products) ? req.body.products.join('; ') : req.body.products || '';
  const markets = Array.isArray(req.body.markets) ? req.body.markets.join('; ') : req.body.markets || '';
  const content_types = Array.isArray(req.body.content_types) ? req.body.content_types.join('; ') : req.body.content_types || '';
  const regions = Array.isArray(req.body.regions) ? req.body.regions.join('; ') : req.body.regions || '';

  db.run(`UPDATE customers SET contact_name=?, company=?, customer_id=?, email=?, cc_emails=?, products=?, markets=?, content_types=?, regions=?, customer_type=?, subscription_tier=?, preferred_frequency=?, status=? WHERE id=?`,
    [contact_name, company, normalizedCustomerId, email, cc_emails, products, markets, content_types, regions, normalizedCustomerType, subscription_tier, preferred_frequency, status, id],
    function(err) {
      if (err) {
        console.error('Update customer error:', err);
        return res.redirect('/customers?error=' + encodeURIComponent('Error updating customer'));
      }
      res.redirect('/customers?success=' + encodeURIComponent('Customer updated successfully'));
    });
});

app.post('/customers/:id/delete', (req, res) => {
  db.run('DELETE FROM customers WHERE id = ?', [req.params.id], function(err) {
    if (err) return res.redirect('/customers?error=' + encodeURIComponent('Error deleting customer'));
    res.redirect('/customers?success=' + encodeURIComponent('Customer deleted successfully'));
  });
});

// ============================================================
// PUBLICATION MANAGEMENT
// ============================================================

app.get('/publications', async (req, res) => {
  try {
    const { search, status, urgency, archived, page, pageSize } = req.query;
    const showArchived = parseBoolean(archived);
    const filterParts = buildPublicationFilters({ search, status, urgency });
    let whereSql = filterParts.whereSql;
    const whereParams = filterParts.whereParams;
    if (showArchived) {
      whereSql += ' AND is_archived = 1';
    } else {
      whereSql += ' AND (is_archived IS NULL OR is_archived = 0)';
    }

    const totalRow = await dbGet(`SELECT COUNT(*) as c FROM publications ${whereSql}`, whereParams);
    const totalPublications = totalRow ? totalRow.c : 0;
    const perPage = clampNumber(pageSize, 10, 100, 20);
    const totalPages = Math.max(1, Math.ceil(totalPublications / perPage));
    const currentPage = clampNumber(page, 1, totalPages, 1);
    const offset = (currentPage - 1) * perPage;

    const listParams = whereParams.slice();
    listParams.push(perPage, offset);
    const publications = await dbAll(`SELECT * FROM publications ${whereSql} ORDER BY id DESC LIMIT ? OFFSET ?`, listParams);
    res.render('publications', {
      publications,
      search: search || '',
      statusFilter: status || '',
      urgencyFilter: urgency || '',
      archivedFilter: showArchived ? '1' : '',
      page: currentPage,
      pageSize: perPage,
      totalPublications,
      totalPages,
      pageStart: totalPublications ? offset + 1 : 0,
      pageEnd: Math.min(offset + perPage, totalPublications),
      success: req.query.success || '',
      error: req.query.error || ''
    });
  } catch (err) {
    console.error('Publications error:', err);
    res.status(500).send('Server error');
  }
});

app.get('/publications/new', async (req, res) => {
  try {
    const metadata = await getMetadata();
    res.render('publication_form', { publication: null, metadata });
  } catch (err) {
    res.status(500).send('Server error');
  }
});

app.post('/publications', upload.single('document'), (req, res) => {
  const { title, publication_number, content_type, urgency, summary, action_required, author_name, reviewer } = req.body;
  const products = Array.isArray(req.body.products) ? req.body.products.join('; ') : req.body.products || '';
  const markets = Array.isArray(req.body.markets) ? req.body.markets.join('; ') : req.body.markets || '';
  const regions = Array.isArray(req.body.regions) ? req.body.regions.join('; ') : req.body.regions || '';
  const file_path = req.file ? req.file.path : null;
  const file_name = req.file ? req.file.originalname : null;

  db.run(`INSERT INTO publications (title, publication_number, products, markets, content_type, regions, urgency, summary, action_required, author_name, reviewer, file_path, file_name)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [title, publication_number, products, markets, content_type, regions, urgency, summary, action_required, author_name, reviewer, file_path, file_name],
    function(err) {
      if (err) {
        console.error('Create publication error:', err);
        return res.redirect('/publications?error=' + encodeURIComponent('Error creating publication'));
      }
      res.redirect('/publications?success=' + encodeURIComponent('Publication uploaded successfully'));
    });
});

app.get('/publications/:id/edit', async (req, res) => {
  try {
    const metadata = await getMetadata();
    const publication = await dbGet('SELECT * FROM publications WHERE id = ?', [req.params.id]);
    if (!publication) return res.status(404).send('Publication not found');
    res.render('publication_form', { publication, metadata });
  } catch (err) {
    res.status(500).send('Server error');
  }
});

app.post('/publications/archive-bulk', express.json(), async (req, res) => {
  try {
    const { ids, apply_all, filters } = req.body || {};
    if (parseBoolean(apply_all)) {
      const { whereSql, whereParams } = buildPublicationFilters({
        search: filters && filters.search,
        status: filters && filters.status,
        urgency: filters && filters.urgency
      });
      const result = await dbRun(
        `UPDATE publications
         SET is_archived = 1, archived_at = datetime("now")
         ${whereSql} AND (is_archived IS NULL OR is_archived = 0)`,
        whereParams
      );
      return res.json({ success: true, count: result.changes || 0 });
    }

    const parsedIds = Array.isArray(ids)
      ? ids.map(id => parseInt(id, 10)).filter(Number.isInteger)
      : [];
    if (!parsedIds.length) {
      return res.status(400).json({ error: 'No publications selected' });
    }
    const placeholders = parsedIds.map(() => '?').join(',');
    const result = await dbRun(
      `UPDATE publications
       SET is_archived = 1, archived_at = datetime("now")
       WHERE id IN (${placeholders})`,
      parsedIds
    );
    res.json({ success: true, count: result.changes || parsedIds.length });
  } catch (err) {
    console.error('Archive publications error:', err);
    res.status(500).json({ error: 'Failed to archive publications' });
  }
});

app.post('/publications/restore-bulk', express.json(), async (req, res) => {
  try {
    const { ids, apply_all, filters } = req.body || {};
    if (parseBoolean(apply_all)) {
      const { whereSql, whereParams } = buildPublicationFilters({
        search: filters && filters.search,
        status: filters && filters.status,
        urgency: filters && filters.urgency
      });
      const result = await dbRun(
        `UPDATE publications
         SET is_archived = 0, archived_at = NULL
         ${whereSql} AND is_archived = 1`,
        whereParams
      );
      return res.json({ success: true, count: result.changes || 0 });
    }

    const parsedIds = Array.isArray(ids)
      ? ids.map(id => parseInt(id, 10)).filter(Number.isInteger)
      : [];
    if (!parsedIds.length) {
      return res.status(400).json({ error: 'No publications selected' });
    }
    const placeholders = parsedIds.map(() => '?').join(',');
    const result = await dbRun(
      `UPDATE publications SET is_archived = 0, archived_at = NULL WHERE id IN (${placeholders})`,
      parsedIds
    );
    res.json({ success: true, count: result.changes || parsedIds.length });
  } catch (err) {
    console.error('Restore publications error:', err);
    res.status(500).json({ error: 'Failed to restore publications' });
  }
});

app.post('/publications/delete-bulk', express.json(), async (req, res) => {
  try {
    const permanent = parseBoolean(req.body.permanent);
    if (!permanent) {
      return res.status(400).json({ error: 'Permanent delete requires confirmation' });
    }

    const ids = Array.isArray(req.body.ids)
      ? req.body.ids.map(id => parseInt(id, 10)).filter(Number.isInteger)
      : [];
    const applyAll = parseBoolean(req.body.apply_all);

    if (applyAll) {
      const filters = req.body.filters || {};
      const { whereSql, whereParams } = buildPublicationFilters({
        search: filters.search,
        status: filters.status,
        urgency: filters.urgency
      });
      const rows = await dbAll(
        `SELECT id, file_path FROM publications ${whereSql} AND is_archived = 1`,
        whereParams
      );
      if (!rows.length) {
        return res.status(400).json({ error: 'No archived publications selected' });
      }
      rows.forEach(pub => {
        if (pub.file_path && fs.existsSync(pub.file_path)) {
          fs.unlinkSync(pub.file_path);
        }
      });
      const deleteIds = rows.map(pub => pub.id);
      const deletePlaceholders = deleteIds.map(() => '?').join(',');
      const result = await dbRun(`DELETE FROM publications WHERE id IN (${deletePlaceholders})`, deleteIds);
      return res.json({ success: true, count: result.changes || deleteIds.length });
    }

    if (!ids.length) {
      return res.status(400).json({ error: 'No publications selected' });
    }

    const placeholders = ids.map(() => '?').join(',');
    const rows = await dbAll(`SELECT id, is_archived, file_path FROM publications WHERE id IN (${placeholders})`, ids);
    const deletable = rows.filter(row => row.is_archived === 1);

    if (!deletable.length) {
      return res.status(400).json({ error: 'Only archived publications can be permanently deleted' });
    }

    deletable.forEach(pub => {
      if (pub.file_path && fs.existsSync(pub.file_path)) {
        fs.unlinkSync(pub.file_path);
      }
    });

    const deleteIds = deletable.map(pub => pub.id);
    const deletePlaceholders = deleteIds.map(() => '?').join(',');
    const result = await dbRun(`DELETE FROM publications WHERE id IN (${deletePlaceholders})`, deleteIds);
    res.json({ success: true, count: result.changes || deleteIds.length });
  } catch (err) {
    console.error('Delete publications error:', err);
    res.status(500).json({ error: 'Failed to delete publications' });
  }
});

app.post('/publications/:id', upload.single('document'), async (req, res) => {
  try {
    const id = req.params.id;
    const existing = await dbGet('SELECT * FROM publications WHERE id = ?', [id]);
    if (!existing) return res.redirect('/publications?error=' + encodeURIComponent('Publication not found'));

    const { title, publication_number, content_type, urgency, summary, action_required, author_name, reviewer } = req.body;
    const products = Array.isArray(req.body.products) ? req.body.products.join('; ') : req.body.products || '';
    const markets = Array.isArray(req.body.markets) ? req.body.markets.join('; ') : req.body.markets || '';
    const regions = Array.isArray(req.body.regions) ? req.body.regions.join('; ') : req.body.regions || '';
    const file_path = req.file ? req.file.path : existing.file_path;
    const file_name = req.file ? req.file.originalname : existing.file_name;

    await dbRun(`UPDATE publications SET title=?, publication_number=?, products=?, markets=?, content_type=?, regions=?, urgency=?, summary=?, action_required=?, author_name=?, reviewer=?, file_path=?, file_name=? WHERE id=?`,
      [title, publication_number, products, markets, content_type, regions, urgency, summary, action_required, author_name, reviewer, file_path, file_name, id]);

    res.redirect('/publications?success=' + encodeURIComponent('Publication updated successfully'));
  } catch (err) {
    console.error('Update publication error:', err);
    res.redirect('/publications?error=' + encodeURIComponent('Error updating publication'));
  }
});

app.post('/publications/:id/delete', async (req, res) => {
  try {
    await dbRun('UPDATE publications SET is_archived = 1, archived_at = datetime("now") WHERE id = ?', [req.params.id]);
    res.redirect('/publications?success=' + encodeURIComponent('Publication archived successfully'));
  } catch (err) {
    res.redirect('/publications?error=' + encodeURIComponent('Error archiving publication'));
  }
});

// Legacy approve route — redirect to new preview flow
app.post('/publications/:id/approve', (req, res) => {
  res.redirect('/publications/' + req.params.id + '/distribute');
});

// Distribution preview — show matching customers with checkboxes
app.get('/publications/:id/distribute', async (req, res) => {
  try {
    const pub = await dbGet('SELECT * FROM publications WHERE id = ?', [req.params.id]);
    if (!pub) return res.redirect('/publications?error=' + encodeURIComponent('Publication not found'));

    const allCustomers = await dbAll("SELECT * FROM customers WHERE status = 'Active' ORDER BY company, contact_name");

    const matched = [];
    const unmatched = [];
    allCustomers.forEach(c => {
      if (matches(pub, c)) {
        matched.push(c);
      } else {
        unmatched.push(c);
      }
    });

    res.render('distribute_preview', { publication: pub, matched, unmatched });
  } catch (err) {
    console.error('Distribution preview error:', err);
    res.status(500).send('Server error');
  }
});

// Execute distribution to selected customers
app.post('/publications/:id/distribute', express.json(), async (req, res) => {
  try {
    const id = req.params.id;
    const { customerIds } = req.body;

    if (!Array.isArray(customerIds) || customerIds.length === 0) {
      return res.status(400).json({ error: 'No customers selected' });
    }

    const pub = await dbGet('SELECT * FROM publications WHERE id = ?', [id]);
    if (!pub) return res.status(404).json({ error: 'Publication not found' });

    let recipientsCount = 0;
    for (const custId of customerIds) {
      const customer = await dbGet('SELECT * FROM customers WHERE id = ?', [custId]);
      if (!customer) continue;
      sendEmail(pub, customer);
      logDistribution(pub, customer);
      recipientsCount++;
    }

    await dbRun('UPDATE publications SET distribution_status = ?, date_published = datetime("now"), recipients_count = ? WHERE id = ?',
      ['Distributed', recipientsCount, id]);

    res.json({ success: true, count: recipientsCount });
  } catch (err) {
    console.error('Distribution error:', err);
    res.status(500).json({ error: 'Distribution failed' });
  }
});

// Email preview for a publication (local browser render)
app.get('/publications/:id/email-preview', async (req, res) => {
  try {
    const pub = await dbGet('SELECT * FROM publications WHERE id = ?', [req.params.id]);
    if (!pub) return res.status(404).send('Publication not found');

    const baseUrl = getBaseUrl();
    const logoAsset = resolveLogoAsset(baseUrl);
    const previewEmail = (req.query.email || 'preview@example.com').toString();
    const html = generateEmailHTML(pub, previewEmail, {
      baseUrl,
      logoSrc: logoAsset ? logoAsset.publicUrl : ''
    });
    res.set('Content-Type', 'text/html; charset=utf-8');
    res.send(html);
  } catch (err) {
    console.error('Email preview error:', err);
    res.status(500).send('Server error');
  }
});

// Download publication
app.get('/publications/:id/download', (req, res) => {
  db.get('SELECT file_path, file_name FROM publications WHERE id = ?', [req.params.id], (err, row) => {
    if (err || !row || !row.file_path) {
      return res.status(404).send('File not found');
    }
    res.download(row.file_path, row.file_name);
  });
});

// ============================================================
// DISTRIBUTION LOGIC
// ============================================================

// Matching logic
function matches(publication, customer) {
  const pubProducts = publication.products ? publication.products.split(';').map(p => p.trim()) : [];
  const custProducts = customer.products ? customer.products.split(';').map(p => p.trim()) : [];
  const productMatch = custProducts.includes('All Products') || pubProducts.some(p => custProducts.includes(p));

  const pubMarkets = publication.markets ? publication.markets.split(';').map(m => m.trim()) : [];
  const custMarkets = customer.markets ? customer.markets.split(';').map(m => m.trim()) : [];
  const marketMatch = custMarkets.includes('All Markets') || pubMarkets.some(m => custMarkets.includes(m));

  const custContentTypes = customer.content_types ? customer.content_types.split(';').map(c => c.trim()) : [];
  const contentTypeMatch = custContentTypes.includes('All Content Types') || custContentTypes.includes(publication.content_type);

  const pubRegions = publication.regions ? publication.regions.split(';').map(r => r.trim()) : [];
  const custRegions = customer.regions ? customer.regions.split(';').map(r => r.trim()) : [];
  const regionMatch = custRegions.includes('Global') || pubRegions.some(r => custRegions.includes(r));

  let tierMatch = true;
  if (publication.urgency === 'Critical/Safety') {
    // All tiers
  } else if (publication.urgency === 'High' || publication.urgency === 'Standard') {
    tierMatch = customer.subscription_tier === 'Standard' || customer.subscription_tier === 'All Announcements';
  } else if (publication.urgency === 'Informational') {
    tierMatch = customer.subscription_tier === 'All Announcements';
  }

  return productMatch && marketMatch && contentTypeMatch && regionMatch && tierMatch;
}

function resolveLogoAsset(baseUrl) {
  const logoCandidates = ['PSI FLAG LOGO.jpg', 'psi-logo.svg'];
  for (const logoName of logoCandidates) {
    const filePath = path.join(__dirname, 'public', 'images', logoName);
    if (fs.existsSync(filePath)) {
      return {
        fileName: logoName,
        filePath,
        publicUrl: `${baseUrl}/images/${encodeURIComponent(logoName)}`
      };
    }
  }
  return null;
}

// Send email with document attached
function sendEmail(publication, customer) {
  const shortTitle = publication.title.length > 50 ? publication.title.substring(0, 50).trim() : publication.title;
  const subject = `PSI ${publication.content_type} ${publication.publication_number} – ${shortTitle}`;
  const baseUrl = getBaseUrl();
  const fromAddress = process.env.SMTP_FROM || 'publications@psi.com';
  const logoAsset = resolveLogoAsset(baseUrl);
  const logoCid = 'psi-logo@psi-publications';

  console.log(`EMAIL: To=${customer.email} Subject="${subject}"`);

  const mailOptions = {
    from: fromAddress,
    to: customer.email,
    cc: customer.cc_emails || undefined,
    subject: subject,
    html: generateEmailHTML(publication, customer.email, {
      baseUrl,
      logoSrc: logoAsset ? `cid:${logoCid}` : ''
    })
  };

  const attachments = [];

  // Attach the document if a file exists
  if (publication.file_path && fs.existsSync(publication.file_path)) {
    attachments.push({
      filename: publication.file_name || path.basename(publication.file_path),
      path: publication.file_path
    });
    console.log(`  Attaching document: ${attachments[attachments.length - 1].filename}`);
  }

  // Attach logo inline so it renders in email clients
  if (logoAsset) {
    attachments.push({
      filename: logoAsset.fileName,
      path: logoAsset.filePath,
      cid: logoCid,
      contentDisposition: 'inline'
    });
    console.log(`  Attaching inline logo: ${logoAsset.fileName}`);
  }

  if (attachments.length) {
    mailOptions.attachments = attachments;
  }

  if (transporter) {
    transporter.sendMail(mailOptions, (error, info) => {
      if (error) {
        console.error('Email send error:', error.message);
      } else {
        console.log('Email sent successfully:', info.response);
      }
    });
  } else {
    console.log('  (SMTP not configured - email logged only)');
  }
}

// Generate email HTML
function generateEmailHTML(publication, customerEmail, options = {}) {
  const baseUrl = options.baseUrl || getBaseUrl();
  const releaseDate = new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
  const fileName = publication.file_name || '';
  const unsubscribeUrl = `${baseUrl}/unsubscribe?email=${encodeURIComponent(customerEmail || '')}`;
  const logoAsset = resolveLogoAsset(baseUrl);
  const logoSrc = options.logoSrc || (logoAsset ? logoAsset.publicUrl : '');

  const escapeHtml = (value) => String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
  const formatText = (value) => escapeHtml(value).replace(/\r?\n/g, '<br>');

  // Build structured "Applies To" lines
  const productsList = publication.products ? publication.products.split(';').map(p => p.trim()).filter(Boolean) : [];
  const marketsList = publication.markets ? publication.markets.split(';').map(m => m.trim()).filter(Boolean) : [];
  const regionsList = publication.regions ? publication.regions.split(';').map(r => r.trim()).filter(Boolean) : [];

  const urgency = publication.urgency || 'Standard';
  const urgencyStyleMap = {
    'Critical/Safety': { bg: '#fdeaea', fg: '#9f1d1d', border: '#f3c5c5', accent: '#c62828', actionBg: '#fff2f0' },
    'High': { bg: '#fff4e5', fg: '#9a4d00', border: '#f4d7b0', accent: '#d9822b', actionBg: '#fff8ef' },
    'Standard': { bg: '#e8f5e9', fg: '#1b5e20', border: '#c8e6c9', accent: '#2f7d32', actionBg: '#f3faf4' },
    'Informational': { bg: '#e9f2ff', fg: '#0b4ea2', border: '#c9dcff', accent: '#1976d2', actionBg: '#f3f8ff' }
  };
  const urgencyStyle = urgencyStyleMap[urgency] || urgencyStyleMap.Standard;

  const shortenLabel = (item) => {
    const raw = String(item || '').trim();
    if (!raw) return '';
    const map = {
      'North America': 'NA',
      'Latin America': 'LATAM',
      'Europe, Middle East and Africa': 'EMEA',
      'Asia Pacific': 'APAC',
      'All Products': 'All',
      'All Markets': 'All',
      'Global': 'Global'
    };
    if (map[raw]) return map[raw];
    return raw.length > 24 ? `${raw.slice(0, 21).trim()}...` : raw;
  };

  const compactList = (items, maxItems) => {
    if (!items.length) return 'Not specified';
    const shortItems = items.map(shortenLabel).filter(Boolean);
    if (shortItems.includes('All')) return 'All';
    const shown = shortItems.slice(0, maxItems).join(', ');
    const remaining = shortItems.length - maxItems;
    return remaining > 0 ? `${shown} +${remaining} more` : shown;
  };

  const productsCompact = compactList(productsList, 4);
  const marketsCompact = compactList(marketsList, 4);
  const regionsCompact = compactList(regionsList, 4);
  const pubLabel = `${publication.content_type || ''} ${publication.publication_number || ''}`.trim();
  const summaryText = publication.summary ? formatText(publication.summary) : 'No summary provided.';
  const titleText = formatText(publication.title || 'Untitled Publication');
  const fileLabel = fileName ? formatText(fileName) : '';

  const actionBlock = publication.action_required ? `
                <div style="border-left:4px solid ${urgencyStyle.accent}; background-color:${urgencyStyle.actionBg}; padding:14px 16px; margin:18px 0 16px 0; border-radius:8px;">
                  <div style="font-size:14px; line-height:1.65; color:#334155;">
                    <strong style="color:#111827;">Action Required</strong><br>
                    ${formatText(publication.action_required)}
                  </div>
                </div>` : '';

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="x-apple-disable-message-reformatting">
    <title>PSI ${escapeHtml(pubLabel)}</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&amp;display=swap" rel="stylesheet">
    <style>
      body {
        margin: 0;
        padding: 0;
        background-color: #f3f4f6;
        font-family: Inter, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif;
      }
      .email-shell {
        width: 1120px;
        max-width: 1120px;
        border: 1px solid #dbe4ed;
        border-radius: 14px;
        overflow: hidden;
        background: #ffffff;
      }
      @media only screen and (max-width: 1140px) {
        .email-shell {
          width: 100% !important;
          max-width: 100% !important;
        }
        .main-col, .side-col {
          display: block !important;
          width: 100% !important;
        }
        .main-col {
          border-right: 0 !important;
          border-bottom: 1px solid #e2e8f0 !important;
        }
      }
    </style>
  </head>
  <body>
    <div style="display:none; max-height:0; overflow:hidden; opacity:0; color:transparent;">
      ${escapeHtml(pubLabel)}. ${escapeHtml(publication.title || '')}.
    </div>
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse; background-color:#f3f4f6;">
      <tr>
        <td align="center" style="padding:28px 16px;">
          <table role="presentation" cellpadding="0" cellspacing="0" width="1120" class="email-shell" style="border-collapse:collapse;">
            <tr>
              <td style="padding:22px 28px; background-color:#43a047; color:#ffffff;">
                <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;">
                  <tr>
                    <td style="vertical-align:middle; padding-right:12px;">
                      <div style="font-size:22px; line-height:1.2; font-weight:700; letter-spacing:0.2px; text-transform:uppercase;">
                        POWER SOLUTIONS INTERNATIONAL
                      </div>
                      <div style="font-size:13px; opacity:0.95; margin-top:4px; font-weight:500;">
                        Publication Notification
                      </div>
                    </td>
                    <td align="right" style="vertical-align:middle; white-space:nowrap;">
                      ${logoSrc ? `<img src="${logoSrc}" alt="PSI logo" width="148" style="display:block; width:148px; max-width:148px; height:auto; border:0; outline:none; text-decoration:none;">` : ''}
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
            <tr>
              <td style="padding:0;">
                <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;">
                  <tr>
                    <td class="main-col" width="69%" style="padding:30px; border-right:1px solid #e2e8f0; vertical-align:top;">
                      <div style="margin-bottom:22px;">
                        <div style="font-size:12px; letter-spacing:1.3px; text-transform:uppercase; color:#43a047; font-weight:700; margin-bottom:8px;">
                          Bulletin ID: ${escapeHtml(pubLabel)}
                        </div>
                        <div style="font-size:34px; line-height:1.25; font-weight:700; color:#0f172a;">
                          ${titleText}
                        </div>
                      </div>

                      <div style="font-size:19px; line-height:1.3; font-weight:700; color:#1f2937; margin:0 0 10px 0;">Summary</div>
                      <div style="font-size:15px; line-height:1.75; color:#475569; margin:0 0 6px 0;">
                        ${summaryText}
                      </div>

                      ${actionBlock}

                      ${fileLabel ? `<div style="margin-top:20px; border:1px solid #e2e8f0; background-color:#f8fafc; border-radius:10px; padding:14px 16px;">
                        <div style="font-size:11px; text-transform:uppercase; letter-spacing:1.2px; color:#64748b; font-weight:700; margin-bottom:8px;">Attached Documentation</div>
                        <div style="font-size:14px; color:#0f172a; font-weight:600; line-height:1.45;">
                          ${fileLabel}
                        </div>
                      </div>` : ''}

                    </td>
                    <td class="side-col" width="31%" style="padding:30px; background-color:#f8fafc; vertical-align:top;">
                      <div style="margin-bottom:20px;">
                        <div style="font-size:11px; font-weight:700; letter-spacing:1.2px; text-transform:uppercase; color:#94a3b8; margin-bottom:10px;">Service Details</div>
                        <div style="padding:0 0 12px 0;">
                          <div style="font-size:10px; text-transform:uppercase; letter-spacing:1px; color:#64748b; font-weight:700; margin-bottom:4px;">Release Date</div>
                          <div style="font-size:14px; color:#0f172a; font-weight:600;">${escapeHtml(releaseDate)}</div>
                        </div>
                        <div style="font-size:10px; text-transform:uppercase; letter-spacing:1px; color:#64748b; font-weight:700; margin-bottom:6px;">Priority</div>
                        <span style="display:inline-block; padding:5px 10px; border-radius:999px; font-size:12px; font-weight:700; background-color:${urgencyStyle.bg}; color:${urgencyStyle.fg}; border:1px solid ${urgencyStyle.border};">
                          ${escapeHtml(urgency)}
                        </span>
                      </div>

                      <div style="height:1px; background:#e2e8f0; margin:0 0 20px 0;"></div>

                      <div style="font-size:11px; font-weight:700; letter-spacing:1.2px; text-transform:uppercase; color:#94a3b8; margin-bottom:10px;">Applies To</div>
                      <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;">
                        <tr>
                          <td style="width:62px; padding:0 8px 8px 0; font-size:10px; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.8px;">Prod</td>
                          <td style="padding:0 0 8px 0; font-size:13px; color:#334155; line-height:1.45;">${escapeHtml(productsCompact)}</td>
                        </tr>
                        <tr>
                          <td style="width:62px; padding:0 8px 8px 0; font-size:10px; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.8px;">Mkt</td>
                          <td style="padding:0 0 8px 0; font-size:13px; color:#334155; line-height:1.45;">${escapeHtml(marketsCompact)}</td>
                        </tr>
                        <tr>
                          <td style="width:62px; padding:0 8px 0 0; font-size:10px; color:#64748b; font-weight:700; text-transform:uppercase; letter-spacing:0.8px;">Reg</td>
                          <td style="padding:0; font-size:13px; color:#334155; line-height:1.45;">${escapeHtml(regionsCompact)}</td>
                        </tr>
                      </table>

                      <div style="margin-top:20px; border-top:1px solid #e2e8f0; padding-top:16px;">
                        <div style="font-size:11px; font-weight:700; letter-spacing:1px; text-transform:uppercase; color:#64748b; margin-bottom:8px;">Need Help?</div>
                        <div style="font-size:12px; color:#475569; line-height:1.55;">
                          Contact Technical Support or reply directly to this notification.
                        </div>
                      </div>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
            <tr>
              <td style="padding:16px 28px; background-color:#f1f5f9; border-top:1px solid #e2e8f0;">
                <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse;">
                  <tr>
                    <td style="font-size:11px; color:#64748b; line-height:1.55;">
                      You received this notification based on your PSI distribution profile.<br>
                      &copy; ${new Date().getFullYear()} Power Solutions International. All rights reserved.
                    </td>
                    <td align="right" style="white-space:nowrap; font-size:11px; color:#64748b;">
                      <a href="${unsubscribeUrl}" style="color:#4b647d; text-decoration:underline; font-weight:600;">Unsubscribe</a>
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
          </table>
        </td>
      </tr>
    </table>
  </body>
</html>`;
}

// Log distribution
function logDistribution(publication, customer) {
  db.run(`INSERT INTO distribution_logs (publication_number, publication_title, content_type, urgency, recipient_name, recipient_company, recipient_email, match_reason)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?)`,
    [publication.publication_number, publication.title, publication.content_type, publication.urgency, customer.contact_name, customer.company, customer.email, `Products: ${customer.products}, Markets: ${customer.markets}`]);
}

// ============================================================
// DISTRIBUTION LOGS
// ============================================================

app.get('/logs', async (req, res) => {
  try {
    const { search, urgency, sort, dir, archived, page, pageSize } = req.query;
    const showArchived = parseBoolean(archived);
    const logSortColumns = {
      sent_date: 'sent_date',
      publication_number: 'publication_number',
      publication_title: 'publication_title',
      content_type: 'content_type',
      urgency: 'urgency',
      recipient_name: 'recipient_name',
      recipient_company: 'recipient_company',
      recipient_email: 'recipient_email',
      delivery_status: 'delivery_status'
    };
    const sortBy = logSortColumns[sort] ? sort : 'sent_date';
    const sortDir = toSortDir(dir, 'desc');

    const filterParts = buildLogFilters({ search, urgency });
    let whereSql = filterParts.whereSql;
    const whereParams = filterParts.whereParams;
    if (showArchived) {
      whereSql += ' AND is_archived = 1';
    } else {
      whereSql += ' AND (is_archived IS NULL OR is_archived = 0)';
    }

    const totalRow = await dbGet(`SELECT COUNT(*) as c FROM distribution_logs ${whereSql}`, whereParams);
    const totalLogs = totalRow ? totalRow.c : 0;
    const perPage = clampNumber(pageSize, 10, 200, 25);
    const totalPages = Math.max(1, Math.ceil(totalLogs / perPage));
    const currentPage = clampNumber(page, 1, totalPages, 1);
    const offset = (currentPage - 1) * perPage;

    const listParams = whereParams.slice();
    listParams.push(perPage, offset);
    const logs = await dbAll(
      `SELECT * FROM distribution_logs ${whereSql}
       ORDER BY ${logSortColumns[sortBy]} ${sortDir.toUpperCase()}, id DESC
       LIMIT ? OFFSET ?`,
      listParams
    );

    // Group logs by distribution event (same publication_number + same sent_date rounded to the minute)
    const groups = [];
    const groupMap = {};
    logs.forEach(log => {
      const dateKey = log.sent_date ? log.sent_date.substring(0, 16) : 'unknown'; // YYYY-MM-DDTHH:MM
      const key = log.publication_number + '|' + dateKey;
      if (!groupMap[key]) {
        groupMap[key] = {
          key,
          publication_number: log.publication_number,
          publication_title: log.publication_title,
          content_type: log.content_type,
          urgency: log.urgency,
          sent_date: log.sent_date,
          entries: []
        };
        groups.push(groupMap[key]);
      }
      groupMap[key].entries.push(log);
    });

    res.render('logs', {
      logs,
      groups,
      search: search || '',
      urgencyFilter: urgency || '',
      archivedFilter: showArchived ? '1' : '',
      sortBy,
      sortDir,
      page: currentPage,
      pageSize: perPage,
      totalLogs,
      totalPages,
      pageStart: totalLogs ? offset + 1 : 0,
      pageEnd: Math.min(offset + perPage, totalLogs),
      success: req.query.success || '',
      error: req.query.error || ''
    });
  } catch (err) {
    console.error('Logs error:', err);
    res.status(500).send('Server error');
  }
});

app.post('/logs/archive-bulk', express.json(), async (req, res) => {
  try {
    const { ids, apply_all, filters } = req.body || {};
    if (parseBoolean(apply_all)) {
      const { whereSql, whereParams } = buildLogFilters({
        search: filters && filters.search,
        urgency: filters && filters.urgency
      });
      const result = await dbRun(
        `UPDATE distribution_logs
         SET is_archived = 1, archived_at = datetime("now")
         ${whereSql} AND (is_archived IS NULL OR is_archived = 0)`,
        whereParams
      );
      return res.json({ success: true, count: result.changes || 0 });
    }

    const parsedIds = Array.isArray(ids)
      ? ids.map(id => parseInt(id, 10)).filter(Number.isInteger)
      : [];
    if (!parsedIds.length) {
      return res.status(400).json({ error: 'No log entries selected' });
    }
    const placeholders = parsedIds.map(() => '?').join(',');
    const result = await dbRun(
      `UPDATE distribution_logs
       SET is_archived = 1, archived_at = datetime("now")
       WHERE id IN (${placeholders})`,
      parsedIds
    );
    res.json({ success: true, count: result.changes || parsedIds.length });
  } catch (err) {
    console.error('Archive logs error:', err);
    res.status(500).json({ error: 'Failed to archive logs' });
  }
});

app.post('/logs/restore-bulk', express.json(), async (req, res) => {
  try {
    const { ids, apply_all, filters } = req.body || {};
    if (parseBoolean(apply_all)) {
      const { whereSql, whereParams } = buildLogFilters({
        search: filters && filters.search,
        urgency: filters && filters.urgency
      });
      const result = await dbRun(
        `UPDATE distribution_logs
         SET is_archived = 0, archived_at = NULL
         ${whereSql} AND is_archived = 1`,
        whereParams
      );
      return res.json({ success: true, count: result.changes || 0 });
    }

    const parsedIds = Array.isArray(ids)
      ? ids.map(id => parseInt(id, 10)).filter(Number.isInteger)
      : [];
    if (!parsedIds.length) {
      return res.status(400).json({ error: 'No log entries selected' });
    }
    const placeholders = parsedIds.map(() => '?').join(',');
    const result = await dbRun(
      `UPDATE distribution_logs SET is_archived = 0, archived_at = NULL WHERE id IN (${placeholders})`,
      parsedIds
    );
    res.json({ success: true, count: result.changes || parsedIds.length });
  } catch (err) {
    console.error('Restore logs error:', err);
    res.status(500).json({ error: 'Failed to restore logs' });
  }
});

// Delete selected log entries (permanent)
app.post('/logs/delete-bulk', express.json(), async (req, res) => {
  try {
    const permanent = parseBoolean(req.body.permanent);
    if (!permanent) {
      return res.status(400).json({ error: 'Permanent delete requires confirmation' });
    }
    const applyAll = parseBoolean(req.body.apply_all);
    if (applyAll) {
      const filters = req.body.filters || {};
      const { whereSql, whereParams } = buildLogFilters({
        search: filters.search,
        urgency: filters.urgency
      });
      const rows = await dbAll(
        `SELECT id FROM distribution_logs ${whereSql} AND is_archived = 1`,
        whereParams
      );
      if (!rows.length) {
        return res.status(400).json({ error: 'No archived log entries selected' });
      }
      const ids = rows.map(row => row.id);
      const placeholders = ids.map(() => '?').join(',');
      await dbRun(`DELETE FROM distribution_logs WHERE id IN (${placeholders})`, ids);
      return res.json({ success: true, count: ids.length });
    }

    const ids = Array.isArray(req.body.ids)
      ? req.body.ids.map(id => parseInt(id, 10)).filter(Number.isInteger)
      : [];
    if (!ids.length) {
      return res.status(400).json({ error: 'No log entries selected' });
    }
    const placeholders = ids.map(() => '?').join(',');
    const rows = await dbAll(`SELECT id, is_archived FROM distribution_logs WHERE id IN (${placeholders})`, ids);
    const deletable = rows.filter(row => row.is_archived === 1);
    if (!deletable.length) {
      return res.status(400).json({ error: 'Only archived logs can be permanently deleted' });
    }
    const deleteIds = deletable.map(row => row.id);
    const deletePlaceholders = deleteIds.map(() => '?').join(',');
    await dbRun(`DELETE FROM distribution_logs WHERE id IN (${deletePlaceholders})`, deleteIds);
    res.json({ success: true, count: deleteIds.length });
  } catch (err) {
    console.error('Delete logs error:', err);
    res.status(500).json({ error: 'Failed to delete logs' });
  }
});

// Export logs to Excel
app.get('/logs/export', async (req, res) => {
  try {
    const { search, urgency, sort, dir, archived } = req.query;
    const showArchived = parseBoolean(archived);
    const logSortColumns = {
      sent_date: 'sent_date',
      publication_number: 'publication_number',
      publication_title: 'publication_title',
      content_type: 'content_type',
      urgency: 'urgency',
      recipient_name: 'recipient_name',
      recipient_company: 'recipient_company',
      recipient_email: 'recipient_email',
      delivery_status: 'delivery_status'
    };
    const sortBy = logSortColumns[sort] ? sort : 'sent_date';
    const sortDir = toSortDir(dir, 'desc');

    const filterParts = buildLogFilters({ search, urgency });
    let whereSql = filterParts.whereSql;
    const whereParams = filterParts.whereParams;
    if (showArchived) {
      whereSql += ' AND is_archived = 1';
    } else {
      whereSql += ' AND (is_archived IS NULL OR is_archived = 0)';
    }
    const logs = await dbAll(
      `SELECT * FROM distribution_logs ${whereSql}
       ORDER BY ${logSortColumns[sortBy]} ${sortDir.toUpperCase()}, id DESC`,
      whereParams
    );

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Distribution Logs');

    sheet.columns = [
      { header: 'Date', key: 'sent_date', width: 20 },
      { header: 'Publication #', key: 'publication_number', width: 16 },
      { header: 'Title', key: 'publication_title', width: 35 },
      { header: 'Type', key: 'content_type', width: 18 },
      { header: 'Urgency', key: 'urgency', width: 14 },
      { header: 'Recipient', key: 'recipient_name', width: 22 },
      { header: 'Company', key: 'recipient_company', width: 22 },
      { header: 'Email', key: 'recipient_email', width: 28 },
      { header: 'Status', key: 'delivery_status', width: 12 },
      { header: 'Match Reason', key: 'match_reason', width: 20 }
    ];

    // Style header row
    sheet.getRow(1).font = { bold: true };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D32' } };
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };

    logs.forEach(log => {
      sheet.addRow({
        sent_date: log.sent_date ? new Date(log.sent_date).toLocaleString() : '',
        publication_number: log.publication_number,
        publication_title: log.publication_title,
        content_type: log.content_type,
        urgency: log.urgency,
        recipient_name: log.recipient_name,
        recipient_company: log.recipient_company,
        recipient_email: log.recipient_email,
        delivery_status: log.delivery_status,
        match_reason: log.match_reason || ''
      });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=distribution-logs-' + new Date().toISOString().slice(0, 10) + '.xlsx');
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('Export error:', err);
    res.status(500).send('Export failed');
  }
});

app.get('/logs/export-selected', async (req, res) => {
  try {
    const applyAll = parseBoolean(req.query.apply_all);
    let logs = [];

    if (applyAll) {
      const { search, urgency, archived } = req.query;
      const showArchived = parseBoolean(archived);
      const filterParts = buildLogFilters({ search, urgency });
      let whereSql = filterParts.whereSql;
      const whereParams = filterParts.whereParams;
      if (showArchived) {
        whereSql += ' AND is_archived = 1';
      } else {
        whereSql += ' AND (is_archived IS NULL OR is_archived = 0)';
      }
      logs = await dbAll(
        `SELECT * FROM distribution_logs ${whereSql} ORDER BY sent_date DESC, id DESC`,
        whereParams
      );
    } else {
      const ids = String(req.query.ids || '')
        .split(',')
        .map(id => parseInt(id, 10))
        .filter(Number.isInteger);
      if (!ids.length) {
        return res.status(400).send('No log entries selected');
      }
      const placeholders = ids.map(() => '?').join(',');
      logs = await dbAll(`SELECT * FROM distribution_logs WHERE id IN (${placeholders})`, ids);
    }

    if (!logs.length) {
      return res.status(400).send('No log entries selected');
    }

    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Selected Logs');
    sheet.columns = [
      { header: 'Date', key: 'sent_date', width: 20 },
      { header: 'Publication #', key: 'publication_number', width: 16 },
      { header: 'Title', key: 'publication_title', width: 35 },
      { header: 'Type', key: 'content_type', width: 18 },
      { header: 'Urgency', key: 'urgency', width: 14 },
      { header: 'Recipient', key: 'recipient_name', width: 22 },
      { header: 'Company', key: 'recipient_company', width: 22 },
      { header: 'Email', key: 'recipient_email', width: 28 },
      { header: 'Status', key: 'delivery_status', width: 12 },
      { header: 'Match Reason', key: 'match_reason', width: 20 }
    ];

    sheet.getRow(1).font = { bold: true };
    sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF2E7D32' } };
    sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };

    logs.forEach(log => {
      sheet.addRow({
        sent_date: log.sent_date ? new Date(log.sent_date).toLocaleString() : '',
        publication_number: log.publication_number,
        publication_title: log.publication_title,
        content_type: log.content_type,
        urgency: log.urgency,
        recipient_name: log.recipient_name,
        recipient_company: log.recipient_company,
        recipient_email: log.recipient_email,
        delivery_status: log.delivery_status,
        match_reason: log.match_reason || ''
      });
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename=distribution-logs-selected-' + new Date().toISOString().slice(0, 10) + '.xlsx');
    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('Export selected logs error:', err);
    res.status(500).send('Export failed');
  }
});

// Resend a distribution email
app.post('/logs/:id/resend', async (req, res) => {
  try {
    const logEntry = await dbGet('SELECT * FROM distribution_logs WHERE id = ?', [req.params.id]);
    if (!logEntry) {
      return res.redirect('/logs?error=' + encodeURIComponent('Log entry not found'));
    }

    const publication = await dbGet('SELECT * FROM publications WHERE publication_number = ?', [logEntry.publication_number]);
    const customer = await dbGet('SELECT * FROM customers WHERE email = ? AND company = ?', [logEntry.recipient_email, logEntry.recipient_company]);

    if (!publication || !customer) {
      return res.redirect('/logs?error=' + encodeURIComponent('Could not find original publication or customer record'));
    }

    sendEmail(publication, customer);

    await dbRun(`INSERT INTO distribution_logs (publication_number, publication_title, content_type, urgency, recipient_name, recipient_company, recipient_email, delivery_status, match_reason)
                 VALUES (?, ?, ?, ?, ?, ?, ?, 'Resent', ?)`,
      [logEntry.publication_number, logEntry.publication_title, logEntry.content_type, logEntry.urgency, logEntry.recipient_name, logEntry.recipient_company, logEntry.recipient_email, 'Manual resend']);

    res.redirect('/logs?success=' + encodeURIComponent(`Resent notification to ${logEntry.recipient_name} (${logEntry.recipient_email})`));
  } catch (err) {
    console.error('Resend error:', err);
    res.redirect('/logs?error=' + encodeURIComponent('Error resending notification'));
  }
});

// Bulk resend multiple log entries
app.post('/logs/resend-bulk', express.json(), async (req, res) => {
  try {
    if (parseBoolean(req.body.apply_all)) {
      return res.status(400).json({ error: 'Bulk resend requires explicit selection' });
    }
    const { ids } = req.body;
    if (!Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: 'No log entries selected' });
    }

    let count = 0;
    for (const id of ids) {
      const logEntry = await dbGet('SELECT * FROM distribution_logs WHERE id = ?', [id]);
      if (!logEntry) continue;

      const publication = await dbGet('SELECT * FROM publications WHERE publication_number = ?', [logEntry.publication_number]);
      const customer = await dbGet('SELECT * FROM customers WHERE email = ? AND company = ?', [logEntry.recipient_email, logEntry.recipient_company]);

      if (!publication || !customer) continue;

      sendEmail(publication, customer);

      await dbRun(`INSERT INTO distribution_logs (publication_number, publication_title, content_type, urgency, recipient_name, recipient_company, recipient_email, delivery_status, match_reason)
                   VALUES (?, ?, ?, ?, ?, ?, ?, 'Resent', ?)`,
        [logEntry.publication_number, logEntry.publication_title, logEntry.content_type, logEntry.urgency, logEntry.recipient_name, logEntry.recipient_company, logEntry.recipient_email, 'Bulk resend']);
      count++;
    }

    res.json({ success: true, count });
  } catch (err) {
    console.error('Bulk resend error:', err);
    res.status(500).json({ error: 'Bulk resend failed' });
  }
});

// ============================================================
// SETTINGS / METADATA MANAGEMENT
// ============================================================

app.get('/settings', (req, res) => {
  db.all('SELECT * FROM metadata ORDER BY category, sort_order, value', [], (err, rows) => {
    if (err) return res.status(500).send('Database error');
    const grouped = { products: [], markets: [], content_types: [], regions: [] };
    (rows || []).forEach(row => {
      const key = row.category === 'product' ? 'products' : row.category === 'market' ? 'markets' : row.category === 'content_type' ? 'content_types' : 'regions';
      grouped[key].push(row);
    });
    res.render('settings', {
      metadata: grouped,
      success: req.query.success || '',
      error: req.query.error || ''
    });
  });
});

app.post('/settings/metadata', async (req, res) => {
  const { category, value } = req.body;
  if (!category || !value || !value.trim()) {
    return res.redirect('/settings?error=' + encodeURIComponent('Category and value are required'));
  }
  try {
    const existing = await dbGet('SELECT id FROM metadata WHERE category = ? AND value = ?', [category, value.trim()]);
    if (existing) {
      return res.redirect('/settings?error=' + encodeURIComponent('This value already exists'));
    }
    const maxOrder = await dbGet('SELECT MAX(sort_order) as m FROM metadata WHERE category = ?', [category]);
    const nextOrder = (maxOrder && maxOrder.m !== null) ? maxOrder.m + 1 : 0;
    await dbRun('INSERT INTO metadata (category, value, sort_order) VALUES (?, ?, ?)', [category, value.trim(), nextOrder]);
    res.redirect('/settings?success=' + encodeURIComponent(`Added "${value.trim()}" successfully`));
  } catch (err) {
    console.error('Add metadata error:', err);
    res.redirect('/settings?error=' + encodeURIComponent('Error adding metadata'));
  }
});

app.post('/settings/metadata/:id/delete', (req, res) => {
  db.run('DELETE FROM metadata WHERE id = ?', [req.params.id], function(err) {
    if (err) return res.redirect('/settings?error=' + encodeURIComponent('Error removing metadata'));
    res.redirect('/settings?success=' + encodeURIComponent('Metadata value removed'));
  });
});

// Reorder metadata item (up/down buttons)
app.post('/settings/metadata/:id/reorder', async (req, res) => {
  try {
    const id = parseInt(req.params.id);
    const direction = req.body.direction;

    const item = await dbGet('SELECT * FROM metadata WHERE id = ?', [id]);
    if (!item) return res.redirect('/settings?error=' + encodeURIComponent('Item not found'));

    // Get all items in the same category, ordered
    const siblings = await dbAll('SELECT * FROM metadata WHERE category = ? ORDER BY sort_order, value', [item.category]);
    const currentIndex = siblings.findIndex(s => s.id === id);
    if (currentIndex === -1) return res.redirect('/settings');

    const swapIndex = direction === 'up' ? currentIndex - 1 : currentIndex + 1;
    if (swapIndex < 0 || swapIndex >= siblings.length) return res.redirect('/settings');

    // Re-number the whole category sequentially first
    for (let i = 0; i < siblings.length; i++) {
      await dbRun('UPDATE metadata SET sort_order = ? WHERE id = ?', [i, siblings[i].id]);
    }

    // Now swap the two items
    const newIdx = direction === 'up' ? currentIndex - 1 : currentIndex + 1;
    await dbRun('UPDATE metadata SET sort_order = ? WHERE id = ?', [newIdx, siblings[currentIndex].id]);
    await dbRun('UPDATE metadata SET sort_order = ? WHERE id = ?', [currentIndex, siblings[newIdx].id]);

    res.redirect('/settings');
  } catch (err) {
    console.error('Reorder error:', err);
    res.redirect('/settings?error=' + encodeURIComponent('Error reordering item'));
  }
});

// Drag-and-drop reorder (AJAX)
app.post('/settings/metadata/reorder-bulk', express.json(), async (req, res) => {
  try {
    const { ids } = req.body;
    if (!Array.isArray(ids) || ids.length === 0) {
      return res.status(400).json({ error: 'Invalid ids array' });
    }
    for (let i = 0; i < ids.length; i++) {
      await dbRun('UPDATE metadata SET sort_order = ? WHERE id = ?', [i, parseInt(ids[i])]);
    }
    res.json({ success: true });
  } catch (err) {
    console.error('Bulk reorder error:', err);
    res.status(500).json({ error: 'Reorder failed' });
  }
});

// Edit metadata value
app.post('/settings/metadata/:id/edit', async (req, res) => {
  try {
    const id = parseInt(req.params.id);
    const newValue = (req.body.value || '').trim();

    if (!newValue) {
      return res.redirect('/settings?error=' + encodeURIComponent('Value cannot be empty'));
    }

    const item = await dbGet('SELECT * FROM metadata WHERE id = ?', [id]);
    if (!item) return res.redirect('/settings?error=' + encodeURIComponent('Item not found'));

    const existing = await dbGet('SELECT id FROM metadata WHERE category = ? AND value = ? AND id != ?', [item.category, newValue, id]);
    if (existing) {
      return res.redirect('/settings?error=' + encodeURIComponent('A value with that name already exists'));
    }

    await dbRun('UPDATE metadata SET value = ? WHERE id = ?', [newValue, id]);
    res.redirect('/settings?success=' + encodeURIComponent(`Renamed to "${newValue}" successfully`));
  } catch (err) {
    console.error('Edit metadata error:', err);
    res.redirect('/settings?error=' + encodeURIComponent('Error editing metadata'));
  }
});

// ============================================================
// UNSUBSCRIBE
// ============================================================

app.get('/unsubscribe', async (req, res) => {
  const email = req.query.email || '';
  let message = '';
  let done = false;

  if (req.query.confirmed === 'true' && email) {
    try {
      const customer = await dbGet('SELECT * FROM customers WHERE email = ?', [email]);
      if (customer) {
        await dbRun("UPDATE customers SET status = 'Inactive' WHERE email = ?", [email]);
        message = 'You have been unsubscribed. You will no longer receive publication notifications.';
        done = true;
      } else {
        message = 'Email address not found in our system.';
      }
    } catch (err) {
      console.error('Unsubscribe error:', err);
      message = 'An error occurred. Please contact support.';
    }
  }

  res.send(`<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Unsubscribe - PSI Publications</title>
  <style>
    body { font-family: Arial, sans-serif; background: #f6f7f9; margin: 0; padding: 40px 20px; color: #222; }
    .container { max-width: 500px; margin: 0 auto; background: #fff; border-radius: 10px; border: 1px solid #e6e8eb; padding: 40px; text-align: center; }
    h1 { font-size: 20px; color: #43A047; margin: 0 0 12px; }
    p { font-size: 14px; line-height: 1.6; color: #444; }
    .btn { display: inline-block; padding: 10px 24px; border-radius: 6px; font-size: 14px; font-weight: 600; text-decoration: none; margin: 8px 4px; cursor: pointer; border: none; }
    .btn-danger { background: #dc3545; color: #fff; }
    .btn-secondary { background: #e2e4e8; color: #444; }
    .footer { margin-top: 20px; font-size: 11px; color: #999; }
  </style>
</head>
<body>
  <div class="container">
    <h1>PSI Publication Notifications</h1>
    ${done || message ? `<p>${message}</p>` : `
      <p>You are requesting to unsubscribe <strong>${email}</strong> from PSI publication notifications.</p>
      <p>Are you sure? You will no longer receive service bulletins, safety notices, or other publication alerts.</p>
      <a href="/unsubscribe?email=${encodeURIComponent(email)}&confirmed=true" class="btn btn-danger">Yes, Unsubscribe</a>
      <a href="javascript:window.close()" class="btn btn-secondary">Cancel</a>
    `}
    <div class="footer">&copy; ${new Date().getFullYear()} Power Solutions International</div>
  </div>
</body>
</html>`);
});

// ============================================================
// CUSTOMER SELF-SERVICE PORTAL
// ============================================================

app.get('/subscribe', async (req, res) => {
  try {
    const metadata = await getMetadata();
    res.render('subscribe', {
      metadata,
      success: req.query.success || '',
      error: req.query.error || ''
    });
  } catch (err) {
    console.error('Subscribe page error:', err);
    res.status(500).send('Server error');
  }
});

app.post('/subscribe', async (req, res) => {
  try {
    const { contact_name, company, email, cc_emails, products, markets, content_types, regions, customer_type, subscription_tier } = req.body;
    const normalizedCustomerType = (customer_type || 'End User').trim() || 'End User';
    const isInternalEmployee = normalizedCustomerType.toLowerCase() === 'internal';

    if (!contact_name || !company || !email) {
      return res.redirect('/subscribe?error=' + encodeURIComponent('Name, company, and email are required'));
    }

    // Check for existing subscription
    const existing = await dbGet('SELECT * FROM customers WHERE email = ?', [email]);
    if (existing) {
      if (existing.status === 'Inactive') {
        // Reactivate
        const existingCustomerId = String(existing.customer_id || '').trim();
        let reactivatedCustomerId = '';
        if (isInternalEmployee) {
          reactivatedCustomerId = '';
        } else if (existingCustomerId) {
          reactivatedCustomerId = existingCustomerId;
        } else {
          const count = await dbGet('SELECT COUNT(*) as c FROM customers');
          reactivatedCustomerId = 'SELF-' + String(count.c + 1).padStart(5, '0');
        }
        await dbRun("UPDATE customers SET status = 'Active', contact_name = ?, company = ?, customer_id = ?, products = ?, markets = ?, content_types = ?, regions = ?, customer_type = ?, subscription_tier = ?, cc_emails = ? WHERE id = ?",
          [contact_name, company,
           reactivatedCustomerId,
           Array.isArray(products) ? products.join('; ') : (products || ''),
           Array.isArray(markets) ? markets.join('; ') : (markets || ''),
           Array.isArray(content_types) ? content_types.join('; ') : (content_types || ''),
           Array.isArray(regions) ? regions.join('; ') : (regions || ''),
           normalizedCustomerType, subscription_tier || 'All Announcements', cc_emails || '', existing.id]);
        return res.redirect('/subscribe?success=' + encodeURIComponent('Welcome back! Your subscription has been reactivated.'));
      }
      return res.redirect('/subscribe?error=' + encodeURIComponent('This email is already subscribed. Contact support to update your profile.'));
    }

    let customerId = '';
    if (!isInternalEmployee) {
      // Generate customer ID for external self-signups
      const count = await dbGet('SELECT COUNT(*) as c FROM customers');
      customerId = 'SELF-' + String(count.c + 1).padStart(5, '0');
    }

    await dbRun(`INSERT INTO customers (contact_name, company, customer_id, email, cc_emails, products, markets, content_types, regions, customer_type, subscription_tier, status)
                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Active')`,
      [contact_name, company, customerId, email, cc_emails || '',
       Array.isArray(products) ? products.join('; ') : (products || ''),
       Array.isArray(markets) ? markets.join('; ') : (markets || ''),
       Array.isArray(content_types) ? content_types.join('; ') : (content_types || ''),
       Array.isArray(regions) ? regions.join('; ') : (regions || ''),
       normalizedCustomerType, subscription_tier || 'All Announcements']);

    res.redirect('/subscribe?success=' + encodeURIComponent('Thank you! You are now subscribed to PSI publication notifications.'));
  } catch (err) {
    console.error('Subscribe error:', err);
    res.redirect('/subscribe?error=' + encodeURIComponent('An error occurred. Please try again.'));
  }
});

// ============================================================
// HEALTH CHECK
// ============================================================

app.get('/health', (req, res) => {
  res.status(200).json({ status: 'ok', timestamp: new Date().toISOString() });
});

// ============================================================
// START SERVER
// ============================================================

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
