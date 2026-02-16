require('dotenv').config();
const express = require('express');
const session = require('express-session');
const sqlite3 = require('sqlite3').verbose();
const multer = require('multer');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

const app = express();
const PORT = process.env.PORT || 3000;

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
      file_name TEXT
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
      match_reason TEXT
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS metadata (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      category TEXT NOT NULL,
      value TEXT NOT NULL,
      sort_order INTEGER DEFAULT 0,
      is_active INTEGER DEFAULT 1,
      created_at DATETIME DEFAULT CURRENT_TIMESTAMP
    )`);

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

// Middleware
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static('public'));
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
  <link rel="stylesheet" href="/styles.css">
</head>
<body style="background: var(--gray-50); display: flex; align-items: center; justify-content: center; min-height: 100vh;">
  <div style="width: 100%; max-width: 400px; padding: 20px;">
    <div style="text-align: center; margin-bottom: 24px;">
      <img src="/images/psi-logo.svg" alt="PSI" style="height: 48px; margin-bottom: 12px;">
      <h1 style="font-size: 20px; font-weight: 700; color: var(--gray-900); margin: 0;">Admin Login</h1>
      <p style="font-size: 13px; color: var(--gray-500); margin: 6px 0 0;">Publication Distribution System</p>
    </div>
    ${req.query.error ? '<div style="background: var(--danger-light); color: #721c24; padding: 10px 16px; border-radius: 8px; margin-bottom: 16px; font-size: 13px; text-align: center;">Invalid username or password</div>' : ''}
    <div class="form-container">
      <form action="/login" method="POST" style="padding: 24px;">
        <div class="form-group">
          <label>Username</label>
          <input type="text" name="username" required autocomplete="username" placeholder="Enter username">
        </div>
        <div class="form-group">
          <label>Password</label>
          <input type="password" name="password" required autocomplete="current-password" placeholder="Enter password">
        </div>
        <button type="submit" class="btn btn-primary" style="width: 100%; margin-top: 8px;">Sign In</button>
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
    const { search, status, tier } = req.query;
    let sql = 'SELECT * FROM customers WHERE 1=1';
    const params = [];

    if (search) {
      sql += ' AND (contact_name LIKE ? OR company LIKE ? OR email LIKE ?)';
      const s = `%${search}%`;
      params.push(s, s, s);
    }
    if (status) {
      sql += ' AND status = ?';
      params.push(status);
    }
    if (tier) {
      sql += ' AND subscription_tier = ?';
      params.push(tier);
    }
    sql += ' ORDER BY company, contact_name';

    const customers = await dbAll(sql, params);
    res.render('customers', {
      customers,
      search: search || '',
      statusFilter: status || '',
      tierFilter: tier || '',
      success: req.query.success || '',
      error: req.query.error || ''
    });
  } catch (err) {
    console.error('Customers error:', err);
    res.status(500).send('Server error');
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
  const products = Array.isArray(req.body.products) ? req.body.products.join('; ') : req.body.products || '';
  const markets = Array.isArray(req.body.markets) ? req.body.markets.join('; ') : req.body.markets || '';
  const content_types = Array.isArray(req.body.content_types) ? req.body.content_types.join('; ') : req.body.content_types || '';
  const regions = Array.isArray(req.body.regions) ? req.body.regions.join('; ') : req.body.regions || '';

  db.run(`INSERT INTO customers (contact_name, company, customer_id, email, cc_emails, products, markets, content_types, regions, customer_type, subscription_tier, preferred_frequency, status)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
    [contact_name, company, customer_id, email, cc_emails, products, markets, content_types, regions, customer_type, subscription_tier, preferred_frequency, status],
    function(err) {
      if (err) {
        console.error('Create customer error:', err);
        return res.redirect('/customers?error=' + encodeURIComponent('Error creating customer'));
      }
      res.redirect('/customers?success=' + encodeURIComponent('Customer created successfully'));
    });
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
  const products = Array.isArray(req.body.products) ? req.body.products.join('; ') : req.body.products || '';
  const markets = Array.isArray(req.body.markets) ? req.body.markets.join('; ') : req.body.markets || '';
  const content_types = Array.isArray(req.body.content_types) ? req.body.content_types.join('; ') : req.body.content_types || '';
  const regions = Array.isArray(req.body.regions) ? req.body.regions.join('; ') : req.body.regions || '';

  db.run(`UPDATE customers SET contact_name=?, company=?, customer_id=?, email=?, cc_emails=?, products=?, markets=?, content_types=?, regions=?, customer_type=?, subscription_tier=?, preferred_frequency=?, status=? WHERE id=?`,
    [contact_name, company, customer_id, email, cc_emails, products, markets, content_types, regions, customer_type, subscription_tier, preferred_frequency, status, id],
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
    const { search, status, urgency } = req.query;
    let sql = 'SELECT * FROM publications WHERE 1=1';
    const params = [];

    if (search) {
      sql += ' AND (title LIKE ? OR publication_number LIKE ?)';
      const s = `%${search}%`;
      params.push(s, s);
    }
    if (status) {
      sql += ' AND distribution_status = ?';
      params.push(status);
    }
    if (urgency) {
      sql += ' AND urgency = ?';
      params.push(urgency);
    }
    sql += ' ORDER BY id DESC';

    const publications = await dbAll(sql, params);
    res.render('publications', {
      publications,
      search: search || '',
      statusFilter: status || '',
      urgencyFilter: urgency || '',
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
    const pub = await dbGet('SELECT * FROM publications WHERE id = ?', [req.params.id]);
    if (pub && pub.file_path && fs.existsSync(pub.file_path)) {
      fs.unlinkSync(pub.file_path);
    }
    await dbRun('DELETE FROM publications WHERE id = ?', [req.params.id]);
    res.redirect('/publications?success=' + encodeURIComponent('Publication deleted successfully'));
  } catch (err) {
    res.redirect('/publications?error=' + encodeURIComponent('Error deleting publication'));
  }
});

// Approve and distribute
app.post('/publications/:id/approve', (req, res) => {
  const id = req.params.id;
  db.run('UPDATE publications SET distribution_status = ? WHERE id = ?', ['Approved', id], (err) => {
    if (err) {
      return res.redirect('/publications?error=' + encodeURIComponent('Error approving publication'));
    }
    distributePublication(id);
    res.redirect('/publications?success=' + encodeURIComponent('Publication approved and distribution started'));
  });
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

function distributePublication(publicationId) {
  console.log(`Starting distribution for publication ID: ${publicationId}`);

  db.get('SELECT * FROM publications WHERE id = ?', [publicationId], (err, pub) => {
    if (err || !pub) {
      console.log('Error getting publication:', err);
      return;
    }

    console.log(`Distributing: ${pub.title} (${pub.publication_number})`);

    db.all('SELECT * FROM customers WHERE status = ?', ['Active'], (err, customers) => {
      if (err) {
        console.log('Error getting customers:', err);
        return;
      }

      console.log(`Found ${customers.length} active customers`);
      let recipientsCount = 0;

      customers.forEach(customer => {
        const match = matches(pub, customer);
        console.log(`Checking ${customer.contact_name} (${customer.company}): ${match ? 'MATCH' : 'NO MATCH'}`);
        if (match) {
          sendEmail(pub, customer);
          logDistribution(pub, customer);
          recipientsCount++;
        }
      });

      console.log(`Distribution complete: ${recipientsCount} recipients`);

      db.run('UPDATE publications SET distribution_status = ?, date_published = datetime("now"), recipients_count = ? WHERE id = ?',
        ['Distributed', recipientsCount, publicationId]);
    });
  });
}

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

// Send email with document attached
function sendEmail(publication, customer) {
  const shortTitle = publication.title.length > 50 ? publication.title.substring(0, 50).trim() : publication.title;
  const subject = `PSI ${publication.content_type} ${publication.publication_number} – ${shortTitle}`;
  const fromAddress = process.env.SMTP_FROM || 'publications@psi.com';

  console.log(`EMAIL: To=${customer.email} Subject="${subject}"`);

  const mailOptions = {
    from: fromAddress,
    to: customer.email,
    cc: customer.cc_emails || undefined,
    subject: subject,
    html: generateEmailHTML(publication, customer.email)
  };

  // Attach the document if a file exists
  if (publication.file_path && fs.existsSync(publication.file_path)) {
    mailOptions.attachments = [{
      filename: publication.file_name || path.basename(publication.file_path),
      path: publication.file_path
    }];
    console.log(`  Attaching: ${mailOptions.attachments[0].filename}`);
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
function generateEmailHTML(publication, customerEmail) {
  const baseUrl = process.env.BASE_URL || 'http://localhost:3000';
  const releaseDate = new Date().toLocaleDateString('en-US', { month: 'long', day: 'numeric', year: 'numeric' });
  const fileName = publication.file_name || '';
  const unsubscribeUrl = `${baseUrl}/unsubscribe?email=${encodeURIComponent(customerEmail || '')}`;

  // Build structured "Applies To" lines
  const productsList = publication.products ? publication.products.split(';').map(p => p.trim()).filter(Boolean) : [];
  const marketsList = publication.markets ? publication.markets.split(';').map(m => m.trim()).filter(Boolean) : [];
  const regionsList = publication.regions ? publication.regions.split(';').map(r => r.trim()).filter(Boolean) : [];

  const actionBlock = publication.action_required ? `
                <div style="border-left:4px solid #e67700; background-color:#fff7e6; padding:12px 14px; margin:0 0 16px 0;">
                  <div style="font-size:14px; line-height:1.55;">
                    <strong>Action Required</strong><br>
                    ${publication.action_required}
                  </div>
                </div>` : '';

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>PSI ${publication.content_type} ${publication.publication_number}</title>
  </head>
  <body style="margin:0; padding:0; background-color:#f6f7f9;">
    <div style="display:none; max-height:0; overflow:hidden; opacity:0; color:transparent;">
      ${publication.content_type} ${publication.publication_number}. ${publication.title}.
    </div>
    <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border-collapse:collapse; background-color:#f6f7f9;">
      <tr>
        <td align="center" style="padding:24px 12px;">
          <table role="presentation" cellpadding="0" cellspacing="0" width="650" style="max-width:650px; width:100%; border-collapse:collapse; background-color:#ffffff; border:1px solid #e6e8eb; border-radius:10px; overflow:hidden;">
            <tr>
              <td style="padding:18px 22px; background-color:#43A047; color:#ffffff; font-family:Arial, sans-serif;">
                <div style="font-size:16px; font-weight:700; letter-spacing:0.2px;">
                  POWER SOLUTIONS INTERNATIONAL
                </div>
                <div style="font-size:12px; opacity:0.9; margin-top:4px;">
                  Publication Notification
                </div>
              </td>
            </tr>
            <tr>
              <td style="padding:22px; font-family:Arial, sans-serif; color:#222;">
                <div style="font-size:16px; font-weight:700; margin:0 0 10px 0;">
                  ${publication.content_type} ${publication.publication_number}
                </div>
                <div style="font-size:14px; line-height:1.55; margin:0 0 14px 0;">
                  <strong>Title:</strong> ${publication.title}<br>
                  <strong>Release Date:</strong> ${releaseDate}<br>
                  <strong>Priority:</strong> ${publication.urgency}
                </div>
                <table role="presentation" cellpadding="0" cellspacing="0" width="100%" style="border:1px solid #e6e8eb; background-color:#fbfbfc; border-radius:8px; border-collapse:collapse; margin:0 0 16px 0;">
                  <tr>
                    <td style="padding:14px; font-family:Arial, sans-serif;">
                      <div style="font-size:13px; font-weight:700; color:#222; margin:0 0 8px 0;">Applies To</div>
                      ${productsList.length ? '<div style="font-size:13px; color:#444; line-height:1.55; margin:0 0 4px 0;"><strong>Products:</strong> ' + productsList.join(', ') + '</div>' : ''}
                      ${marketsList.length ? '<div style="font-size:13px; color:#444; line-height:1.55; margin:0 0 4px 0;"><strong>Markets:</strong> ' + marketsList.join(', ') + '</div>' : ''}
                      ${regionsList.length ? '<div style="font-size:13px; color:#444; line-height:1.55; margin:0;"><strong>Regions:</strong> ' + regionsList.join(', ') + '</div>' : ''}
                    </td>
                  </tr>
                </table>
                ${publication.summary ? `<div style="font-size:14px; line-height:1.65; margin:0 0 12px 0;">
                  <strong>Summary</strong><br>
                  ${publication.summary}
                </div>` : ''}
                ${actionBlock}
                ${fileName ? `<div style="border:1px solid #e6e8eb; background-color:#fbfbfc; border-radius:8px; padding:14px; margin:0 0 16px 0;">
                  <div style="font-size:13px; color:#444; line-height:1.55;">
                    <strong>&#128206; Attached Document</strong><br>
                    ${fileName}
                  </div>
                </div>` : ''}
                <div style="font-size:13px; color:#444; line-height:1.6; margin-top:14px;">
                  Questions: reply to this email or contact Technical Support.
                </div>
              </td>
            </tr>
            <tr>
              <td style="padding:16px 22px; background-color:#f3f4f6; font-family:Arial, sans-serif; color:#666; font-size:11px; line-height:1.5;">
                You received this notification based on your PSI distribution profile.<br>
                &copy; ${new Date().getFullYear()} Power Solutions International. All rights reserved.<br>
                <a href="${unsubscribeUrl}" style="color:#888; text-decoration:underline;">Unsubscribe</a> from future notifications.
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
    const { search, urgency } = req.query;
    let sql = 'SELECT * FROM distribution_logs WHERE 1=1';
    const params = [];

    if (search) {
      sql += ' AND (publication_number LIKE ? OR publication_title LIKE ? OR recipient_name LIKE ? OR recipient_company LIKE ? OR recipient_email LIKE ?)';
      const s = `%${search}%`;
      params.push(s, s, s, s, s);
    }
    if (urgency) {
      sql += ' AND urgency = ?';
      params.push(urgency);
    }
    sql += ' ORDER BY sent_date DESC';

    const logs = await dbAll(sql, params);
    res.render('logs', {
      logs,
      search: search || '',
      urgencyFilter: urgency || '',
      success: req.query.success || '',
      error: req.query.error || ''
    });
  } catch (err) {
    console.error('Logs error:', err);
    res.status(500).send('Server error');
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

    if (!contact_name || !company || !email) {
      return res.redirect('/subscribe?error=' + encodeURIComponent('Name, company, and email are required'));
    }

    // Check for existing subscription
    const existing = await dbGet('SELECT * FROM customers WHERE email = ?', [email]);
    if (existing) {
      if (existing.status === 'Inactive') {
        // Reactivate
        await dbRun("UPDATE customers SET status = 'Active', contact_name = ?, company = ?, products = ?, markets = ?, content_types = ?, regions = ?, customer_type = ?, subscription_tier = ?, cc_emails = ? WHERE id = ?",
          [contact_name, company,
           Array.isArray(products) ? products.join('; ') : (products || ''),
           Array.isArray(markets) ? markets.join('; ') : (markets || ''),
           Array.isArray(content_types) ? content_types.join('; ') : (content_types || ''),
           Array.isArray(regions) ? regions.join('; ') : (regions || ''),
           customer_type || 'End User', subscription_tier || 'All Announcements', cc_emails || '', existing.id]);
        return res.redirect('/subscribe?success=' + encodeURIComponent('Welcome back! Your subscription has been reactivated.'));
      }
      return res.redirect('/subscribe?error=' + encodeURIComponent('This email is already subscribed. Contact support to update your profile.'));
    }

    // Generate customer ID
    const count = await dbGet('SELECT COUNT(*) as c FROM customers');
    const customerId = 'SELF-' + String(count.c + 1).padStart(5, '0');

    await dbRun(`INSERT INTO customers (contact_name, company, customer_id, email, cc_emails, products, markets, content_types, regions, customer_type, subscription_tier, status)
                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Active')`,
      [contact_name, company, customerId, email, cc_emails || '',
       Array.isArray(products) ? products.join('; ') : (products || ''),
       Array.isArray(markets) ? markets.join('; ') : (markets || ''),
       Array.isArray(content_types) ? content_types.join('; ') : (content_types || ''),
       Array.isArray(regions) ? regions.join('; ') : (regions || ''),
       customer_type || 'End User', subscription_tier || 'All Announcements']);

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
