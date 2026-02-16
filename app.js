require('dotenv').config();
const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const multer = require('multer');
const nodemailer = require('nodemailer');
const path = require('path');
const fs = require('fs');

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
    tierMatch = customer.subscription_tier === 'Standard' || customer.subscription_tier === 'Comprehensive';
  } else if (publication.urgency === 'Informational') {
    tierMatch = customer.subscription_tier === 'Comprehensive';
  }

  return productMatch && marketMatch && contentTypeMatch && regionMatch && tierMatch;
}

// Send email
function sendEmail(publication, customer) {
  const subject = `[PSI ${publication.content_type}] ${publication.urgency} — ${publication.title} — ${publication.products}`;
  const fromAddress = process.env.SMTP_FROM || 'publications@psi.com';

  console.log(`EMAIL: To=${customer.email} Subject="${subject}"`);

  const mailOptions = {
    from: fromAddress,
    to: customer.email,
    cc: customer.cc_emails || undefined,
    subject: subject,
    html: generateEmailHTML(publication, customer)
  };

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
function generateEmailHTML(publication, customer) {
  const baseUrl = process.env.BASE_URL || 'http://localhost:3000';
  const urgencyColor = publication.urgency === 'Critical/Safety' ? '#cc0000' : publication.urgency === 'High' ? '#e67700' : '#43A047';
  return `
    <div style="font-family: Arial, sans-serif; max-width: 650px; margin: 0 auto; color: #333;">
      <div style="background-color: #43A047; padding: 20px 30px; color: white;">
        <h2 style="margin: 0; font-size: 18px;">POWER SOLUTIONS INTERNATIONAL</h2>
        <p style="margin: 5px 0 0; font-size: 12px; color: #a5d6a7;">Publication Notification</p>
      </div>
      ${publication.urgency !== 'Informational' ? `<div style="background-color: ${urgencyColor}; color: white; padding: 10px 30px; font-weight: bold; font-size: 14px;">&#9632; ${publication.urgency} PRIORITY</div>` : ''}
      <div style="padding: 30px; background-color: #f9f9f9; border: 1px solid #ddd;">
        <table style="width: 100%; font-size: 14px; margin-bottom: 20px;">
          <tr><td style="color: #666; width: 140px; padding: 5px 0;">Publication:</td><td style="font-weight: bold;">${publication.publication_number}</td></tr>
          <tr><td style="color: #666; padding: 5px 0;">Type:</td><td>${publication.content_type}</td></tr>
          <tr><td style="color: #666; padding: 5px 0;">Products:</td><td>${publication.products}</td></tr>
          <tr><td style="color: #666; padding: 5px 0;">Markets:</td><td>${publication.markets}</td></tr>
          <tr><td style="color: #666; padding: 5px 0;">Date:</td><td>${new Date().toLocaleDateString()}</td></tr>
        </table>
        <hr style="border: none; border-top: 1px solid #ddd;">
        <h3 style="color: #43A047; margin-top: 20px;">${publication.title}</h3>
        <p style="line-height: 1.6;">${publication.summary}</p>
        ${publication.action_required ? `<div style="background-color: #fff3cd; border-left: 4px solid #e67700; padding: 15px; margin: 20px 0;"><strong>Action Required:</strong><br>${publication.action_required}</div>` : ''}
        <div style="text-align: center; margin: 30px 0;">
          <a href="${baseUrl}/publications/${publication.id}/download" style="background-color: #43A047; color: white; padding: 14px 40px; text-decoration: none; border-radius: 4px; font-weight: bold; font-size: 15px; display: inline-block;">View Full Document</a>
        </div>
      </div>
      <div style="padding: 20px 30px; font-size: 11px; color: #999; background-color: #f0f0f0;">
        <p>You received this notification because your PSI distribution profile includes: ${customer.products}, ${customer.markets}.</p>
        <p>&copy; 2026 Power Solutions International. All rights reserved.</p>
      </div>
    </div>
  `;
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

// Reorder metadata item
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
