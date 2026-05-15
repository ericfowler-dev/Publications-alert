# PSI Automated Publication Distribution System

A Node.js web application that automates the distribution of PSI publications to customers based on their subscription profiles.

## Features

- **Customer Profile Management**: Create and manage customer subscription preferences
- **Publication Upload**: Upload documents with comprehensive metadata tagging
- **Automated Matching**: Intelligent tag-based matching for products, markets, content types, regions, and subscription tiers
- **Email Distribution**: Professional HTML email notifications
- **Distribution Logging**: Complete audit trail of all notifications
- **Web Interface**: Clean, responsive UI for all operations

## Setup

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Start the server:**
   ```bash
   npm start
   ```

3. **Access the application:**
   Open http://localhost:3000 in your browser

## Testing the Distribution System

### 1. Create Test Customers

1. Go to "Manage Customers" → "Add New Customer"
2. Fill in customer details with subscription preferences:
   - **Products**: Select relevant products (e.g., "8.8L GSI", "22L DSI")
   - **Markets**: Select markets (e.g., "Power Systems", "Industrial")
   - **Content Types**: Select "Service Bulletin", "Safety Notice", etc.
   - **Regions**: Select "North America", "EMEA", etc.
   - **Subscription Tier**: Choose "Essential", "Standard", or "All Announcements"
   - **Status**: Set to "Active"

### 2. Upload a Test Publication

1. Go to "Upload New Publication"
2. Fill in publication details:
   - **Products**: Select products that match your test customers
   - **Markets**: Select markets that match your test customers
   - **Content Type**: Choose a type your customers are subscribed to
   - **Regions**: Select regions that match your customers
   - **Urgency**: Choose urgency level
3. Optionally upload a document file
4. Click "Upload Publication"

### 3. Distribute the Publication

1. In the Publications list, find your uploaded publication
2. Click "Approve & Distribute"
3. Check the server console for detailed logging:
   - Publication details and tags
   - Number of active customers found
   - Matching results for each customer
   - Email sending logs
   - Final recipient count

### 4. Verify Results

1. **Check the Publications page**: Status should change to "Distributed" with recipient count
2. **Check the Logs page**: Should show entries for each matched customer
3. **Check server console**: Detailed matching and distribution logs

## Email Configuration

For production email sending:

1. Update the SMTP settings in `app.js` (lines 58-63)
2. Uncomment the `transporter.sendMail()` code in the `sendEmail()` function
3. Configure with your email provider (Gmail, Outlook, etc.)

Currently, emails are logged to console for testing.

## Database

- SQLite database (`publications.db`) is created automatically
- Tables: `customers`, `publications`, `distribution_logs`
- Data persists between server restarts

## Production Data Retention

The application uses SQLite. In local development, `publications.db`, `uploads/`, and `exports/` live in the project folder. In production on Render, the app is configured to use the attached persistent disk at `/data`.

Production stores:

- `/data/publications.db`: customers, publications, distribution logs, distribution jobs, metadata, and session records.
- `/data/uploads`: uploaded publication documents and customer import session files.
- `/data/exports`: generated distribution-log archive spreadsheets.

Deployment and migration policy:

- Keep the same Render service and persistent disk attached during revisions. Recreating the service or disk starts with a new empty data volume.
- Startup migrations are additive only: tables are created with `CREATE TABLE IF NOT EXISTS`, columns are added with guarded `ALTER TABLE ... ADD COLUMN`, and indexes are created with `CREATE INDEX IF NOT EXISTS`.
- Do not ship destructive schema migrations without first copying `/data/publications.db`.
- Set `SESSION_SECRET` in production so admin sessions survive deploys/restarts.
- The app does not currently create its own automated database backups. Backups should be handled by Render disk snapshots or a scheduled copy of `/data/publications.db` before releases and at the business cadence PSI chooses.

Distribution log retention:

- Distribution events are appended to `distribution_logs`; manual and bulk resends create new log rows.
- Logs can be archived and restored in the UI, but hard delete is disabled so the audit trail remains retained in the database.
- `LOG_ARCHIVE_DAYS` controls when old active log rows are moved to the archived view. The default production value is 180 days. Archiving does not delete the rows.
- Distribution sends are queued in the database and processed in controlled batches using `EMAIL_SEND_CONCURRENCY` and `EMAIL_SEND_BATCH_DELAY_MS`.

## Matching Logic

Publications are distributed to customers when ALL of these conditions are met:

1. **Product Match**: Customer products include "All Products" OR overlap with publication products
2. **Market Match**: Customer markets include "All Markets" OR overlap with publication markets
3. **Content Type Match**: Customer content types include the publication's content type OR "All Content Types"
4. **Region Match**: Customer regions include "Global" OR overlap with publication regions
5. **Tier Match**: Based on urgency and subscription tier

## Troubleshooting

- **No emails sent**: Check customer status is "Active" and tags match publication
- **Logs not appearing**: Check database connection and table creation
- **Matching issues**: Review console logs for detailed matching information

## Architecture

- **Backend**: Node.js + Express
- **Database**: SQLite
- **Frontend**: EJS templates
- **Email**: Nodemailer
- **File Upload**: Multer
