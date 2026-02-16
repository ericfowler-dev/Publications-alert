# PSI Automated Publication Distribution System
## Portal-Driven Architecture

---

## 1. The Concept

The system works like this:

1. An author **uploads a document to the portal** and fills in metadata (product, market, content type, urgency)
2. The portal **automatically scans all customer profiles** against the document's metadata
3. Every customer whose profile matches **receives an automated email** with a link to the document (and optionally the attachment)
4. Everything is logged — who was notified, when, delivery status

No manual distribution lists. No exports. No mail merges. Upload, tag, publish — done.

```
┌─────────────────────────────────────────────────────────┐
│                                                         │
│   AUTHOR uploads document + assigns metadata            │
│                          │                              │
│                          ▼                              │
│              ┌───────────────────────┐                  │
│              │                       │                  │
│              │    DOCUMENT PORTAL    │                  │
│              │                       │                  │
│              │  1. Stores document   │                  │
│              │  2. Reads metadata    │                  │
│              │  3. Scans customer    │                  │
│              │     profiles          │                  │
│              │  4. Matches metadata  │                  │
│              │     → profiles        │                  │
│              │  5. Sends emails      │                  │
│              │  6. Logs everything   │                  │
│              │                       │                  │
│              └───────────────────────┘                  │
│                          │                              │
│                          ▼                              │
│        ┌─────────────────┼─────────────────┐           │
│        │                 │                 │            │
│    Customer A        Customer B       Customer C       │
│    (matched)         (matched)        (not matched)    │
│    ✅ gets email     ✅ gets email    ❌ no email      │
│                                                         │
└─────────────────────────────────────────────────────────┘
```

---

## 2. Platform Options

There are several ways to build this portal. Here are the realistic options ranked by fit for PSI:

### Option A: SharePoint + Power Automate (Recommended Starting Point)

PSI already has M365. SharePoint can serve as both the document repository and the customer profile database. Power Automate provides the automation engine.

| Component | M365 Tool | Role |
|---|---|---|
| Document portal / library | SharePoint Document Library | Store publications with metadata columns |
| Customer profile database | SharePoint List | Customer records with subscription tags |
| Matching engine + email dispatch | Power Automate | Triggered on document upload, queries profiles, sends emails |
| Email delivery | Outlook / Exchange Online | Sends through PSI's existing email infrastructure |
| Audit log | SharePoint List | Logs every notification sent |
| Customer-facing access | SharePoint site (external sharing) | Customers can browse/download their documents |

**Pros:** No new platforms. No incremental cost. Familiar tools. Enterprise security built in.
**Cons:** Power Automate has some complexity for tag-matching logic. SharePoint external sharing requires configuration. Email volume subject to Exchange limits.

### Option B: Custom Web Portal (Best Long-Term)

A purpose-built web application — could be a lightweight app built on Azure, AWS, or even a low-code platform like PowerApps or Retool.

**Pros:** Exactly what you need. Best UX. Full control. Scalable.
**Cons:** Development time and cost. Needs hosting and maintenance.

### Option C: Monday.com + Make.com + SendGrid (Previous Design)

Monday.com acts as the portal, Make.com does matching, SendGrid delivers.

**Pros:** Leverages existing Monday.com. Good automation.
**Cons:** Monday.com isn't a great document portal. Adds Make.com and SendGrid costs. More duct tape than Option A or B.

### Option D: Third-Party Document Distribution Platform

Platforms like Manula, Document360, Confluence + automation, or industry-specific solutions.

**Pros:** Built for this purpose.
**Cons:** New platform to learn and pay for. May not fit PSI's specific taxonomy needs.

**Recommendation:** Start with **Option A (SharePoint + Power Automate)** — it's free with your existing M365 licenses, your team already uses SharePoint, and Power Automate is powerful enough to handle the matching and email logic. If it outgrows SharePoint's capabilities, the taxonomy and data model transfer directly to a custom portal (Option B) later.

---

## 3. Architecture: SharePoint + Power Automate

### 3.1 High-Level System Diagram

```
┌─────────────────────────────────────────────────────────────────┐
│                     SHAREPOINT SITE                             │
│              "PSI Publication Distribution"                     │
│                                                                 │
│  ┌────────────────────┐      ┌────────────────────────────┐    │
│  │  DOCUMENT LIBRARY  │      │  CUSTOMER PROFILES LIST    │    │
│  │                    │      │                            │    │
│  │  • Upload docs     │      │  • Company, contact info   │    │
│  │  • Metadata columns│      │  • Product tags            │    │
│  │  • Version control │      │  • Market tags             │    │
│  │  • Approval flow   │      │  • Content type prefs      │    │
│  │                    │      │  • Region tags             │    │
│  └────────┬───────────┘      │  • Subscription tier       │    │
│           │                  └─────────────┬──────────────┘    │
│           │                                │                    │
│           └──────────┬─────────────────────┘                   │
│                      │                                          │
│                      ▼                                          │
│           ┌──────────────────┐      ┌──────────────────────┐   │
│           │  POWER AUTOMATE  │      │  DISTRIBUTION LOG    │   │
│           │                  │      │  (SharePoint List)   │   │
│           │  • Triggered on  │─────▶│                      │   │
│           │    publish       │      │  • Who was notified  │   │
│           │  • Matches tags  │      │  • When              │   │
│           │  • Sends emails  │      │  • What publication  │   │
│           │                  │      │  • Delivery status   │   │
│           └──────────────────┘      └──────────────────────┘   │
│                      │                                          │
│                      ▼                                          │
│              ┌──────────────┐                                   │
│              │   OUTLOOK    │                                   │
│              │   (M365)     │                                   │
│              │              │                                   │
│              │  Personalized│                                   │
│              │  emails to   │                                   │
│              │  matched     │                                   │
│              │  customers   │                                   │
│              └──────────────┘                                   │
└─────────────────────────────────────────────────────────────────┘
```

---

## 4. SharePoint Document Library: Publication Portal

This is where authors upload documents. The metadata columns on the library drive the entire matching system.

### 4.1 Document Library Configuration

**Library Name:** `Publications`
**Location:** SharePoint site → "PSI Publication Distribution"

### 4.2 Metadata Columns

These columns are filled in by the author at the time of upload. They define who should receive the document.

| Column Name | Column Type | Required | Values / Configuration |
|---|---|---|---|
| **Document Title** | Single line of text | Yes | Title of the publication |
| **Publication Number** | Single line of text | Yes | e.g., SB-2026-0058 |
| **Products** | Managed Metadata (multi-value) OR Choice (multi-select) | Yes | 8.8L GSI, 8.8L DSI, 22L DSI, 4.3L GSI, 6.0L GSI, 3.0L GSI, 2.4L GSI, 8.8L LPG, All Products |
| **Markets** | Managed Metadata (multi-value) OR Choice (multi-select) | Yes | Power Systems, Industrial, On-Road, Material Handling, Specialty, Marine, Oil & Gas, Agriculture, All Markets |
| **Content Type** | Choice (single-select) | Yes | Service Bulletin, Notice of Change, Manual Update, Safety Notice, Product Alert, Recall Notice, Technical Tip, Training Notice, Product Announcement |
| **Regions** | Choice (multi-select) | Yes | North America, EMEA, APAC, LATAM, Global |
| **Urgency** | Choice (single-select) | Yes | Critical/Safety, High, Standard, Informational |
| **Summary** | Multiple lines of text | Yes | 2-3 sentence description for the notification email |
| **Action Required** | Multiple lines of text | No | What the customer needs to do |
| **Author Name** | Person | Yes | PSI team member who created the document |
| **Reviewer** | Person | No | Approver (if approval workflow is enabled) |
| **Distribution Status** | Choice (single-select) | Auto | Draft, Pending Approval, Approved, Distributed |
| **Date Published** | Date | Auto | Set by Power Automate when distributed |
| **Recipients Count** | Number | Auto | Set by Power Automate after distribution |

### 4.3 Managed Metadata vs. Choice Columns

**Recommended: Managed Metadata** (via the SharePoint Term Store) for Products and Markets. This gives you:
- A centralized, governed taxonomy shared across all SharePoint sites and libraries
- Hierarchical terms (e.g., Products → Gasoline → 8.8L GSI)
- Easier to maintain and extend as new products launch
- Better search and filtering

**Alternative: Choice columns** (multi-select) are simpler to set up and work fine if your taxonomy is relatively flat and stable. Start here if you want speed; migrate to Managed Metadata later if needed.

### 4.4 Content Approval (Optional but Recommended)

Enable SharePoint's built-in content approval on the library:

1. Library Settings → Versioning Settings → Require content approval = **Yes**
2. When an author uploads a document, its status is "Pending"
3. A reviewer approves it → status changes to "Approved"
4. Power Automate triggers on approval status change → begins distribution

This adds a quality gate before anything goes out.

### 4.5 Library Views

| View Name | Filters / Grouping | Purpose |
|---|---|---|
| **All Publications** | Sort by Date Published descending | Master view |
| **Pending Approval** | Distribution Status = Pending Approval | Reviewer's queue |
| **Ready to Distribute** | Distribution Status = Approved | Admin oversight |
| **Distributed** | Distribution Status = Distributed | History |
| **By Product** | Group by Products column | Browse by product platform |
| **By Content Type** | Group by Content Type column | Browse by type |
| **Critical / Safety** | Filter: Urgency = Critical/Safety | Quick access to safety items |

---

## 5. SharePoint List: Customer Profiles

This is the customer subscription database. Each item is one customer contact with their tag profile.

### 5.1 List Configuration

**List Name:** `Customer Profiles`
**Location:** Same SharePoint site

### 5.2 Columns

| Column Name | Column Type | Required | Values / Configuration |
|---|---|---|---|
| **Contact Name** | Single line of text (Title column) | Yes | Full name of the contact |
| **Company** | Single line of text | Yes | Company / organization |
| **Customer ID** | Single line of text | Yes | PSI ERP/CRM customer number |
| **Email** | Single line of text | Yes | Primary notification email |
| **CC Emails** | Multiple lines of text (plain) | No | Additional emails, comma-separated |
| **Products** | Managed Metadata (multi-value) OR Choice (multi-select) | Yes | **Same values as Document Library** |
| **Markets** | Managed Metadata (multi-value) OR Choice (multi-select) | Yes | **Same values as Document Library** |
| **Content Types** | Choice (multi-select) | Yes | **Same values as Document Library** + "All Content Types" |
| **Regions** | Choice (multi-select) | Yes | **Same values as Document Library** + "Global" |
| **Customer Type** | Choice (single-select) | Yes | OEM, Distributor, Dealer, End User, Internal |
| **Subscription Tier** | Choice (single-select) | Yes | Essential, Standard, Comprehensive |
| **Preferred Frequency** | Choice (single-select) | No | Immediate, Daily Digest, Weekly Digest (default: Immediate) |
| **Status** | Choice (single-select) | Yes | Active, Inactive, Suspended |
| **Date Added** | Date | Auto | When the record was created |
| **Last Notified** | Date | Auto | Updated by Power Automate each time a notification is sent |
| **Notes** | Multiple lines of text | No | Special instructions |

### 5.3 Critical Design Rule

> **The Product, Market, Content Type, and Region values in the Customer Profiles list MUST be identical to the values in the Document Library metadata columns.** This is what makes the automated matching work. If the Document Library has "8.8L GSI" and the Customer Profile has "8.8 GSI" (missing the L), the match will fail silently.

Use Managed Metadata from the Term Store to enforce this — both the library and the list pull from the same term set, eliminating spelling/formatting mismatches.

### 5.4 List Views

| View Name | Filters / Grouping | Purpose |
|---|---|---|
| **All Active** | Status = Active | Default working view |
| **By Company** | Group by Company | Company-centric view |
| **By Product** | Group by Products | See who's subscribed to what |
| **By Market** | Group by Markets | Market-based view |
| **Inactive / Suspended** | Status ≠ Active | Cleanup review |
| **OEMs** | Customer Type = OEM | Filtered by type |
| **Distributors** | Customer Type = Distributor | Filtered by type |

---

## 6. SharePoint List: Distribution Log

Audit trail of every notification dispatched.

### 6.1 List Configuration

**List Name:** `Distribution Log`
**Location:** Same SharePoint site

### 6.2 Columns

| Column Name | Column Type | Purpose |
|---|---|---|
| **Log Entry** | Single line of text (Title) | Auto: "{Publication Number} → {Contact Name}" |
| **Publication Number** | Single line of text | Which publication was distributed |
| **Publication Title** | Single line of text | Title of the document |
| **Content Type** | Choice | Type of publication |
| **Urgency** | Choice | Urgency level |
| **Recipient Name** | Single line of text | Who was notified |
| **Recipient Company** | Single line of text | Their company |
| **Recipient Email** | Single line of text | Email address used |
| **Sent Date** | Date and Time | When the notification was dispatched |
| **Delivery Status** | Choice | Sent, Failed |
| **Acknowledged** | Yes/No | For Critical/Safety: has the customer confirmed receipt? |
| **Acknowledgment Date** | Date | When they acknowledged |
| **Match Reason** | Multiple lines of text | Which tags caused the match (for troubleshooting) |

---

## 7. Power Automate: The Matching & Distribution Engine

This is the core automation. When a document is published (approved) in the SharePoint library, Power Automate runs the matching logic and sends emails.

### 7.1 Flow 1: Automatic Distribution on Publish

**Trigger:** When a file is created or modified in the Publications library AND the Distribution Status column = "Approved" (or when the Approval Status changes to Approved if using content approval).

**Detailed Flow:**

```
TRIGGER: When a file property is modified
  → Condition: Distribution Status = "Approved"
  → (Prevents re-triggering on other edits)
         │
         ▼
STEP 1: GET DOCUMENT METADATA
  → Action: Get file properties
  → Retrieve: Products, Markets, Content Type, Regions, 
     Urgency, Summary, Action Required, Publication Number,
     Document Title, file link (URL)
         │
         ▼
STEP 2: GET ALL ACTIVE CUSTOMER PROFILES
  → Action: Get items from "Customer Profiles" list
  → Filter: Status eq 'Active'
  → Note: SharePoint "Get items" returns up to 5,000 items 
     with pagination enabled. For larger lists, use "Send 
     an HTTP request to SharePoint" with $top and $skip.
         │
         ▼
STEP 3: LOOP THROUGH EACH CUSTOMER — "Apply to each"
  │
  │  For each customer profile:
  │       │
  │       ▼
  │  STEP 3a: TAG MATCHING LOGIC (Condition block)
  │  
  │  ┌─────────────────────────────────────────────────┐
  │  │  MATCH = TRUE if ALL of the following are met:  │
  │  │                                                 │
  │  │  PRODUCT MATCH:                                 │
  │  │    Customer Products contains ANY value from    │
  │  │    Document Products                            │
  │  │    OR Customer Products contains "All Products" │
  │  │                                                 │
  │  │  MARKET MATCH:                                  │
  │  │    Customer Markets contains ANY value from     │
  │  │    Document Markets                             │
  │  │    OR Customer Markets contains "All Markets"   │
  │  │                                                 │
  │  │  CONTENT TYPE MATCH:                            │
  │  │    Customer Content Types contains the          │
  │  │    Document's Content Type                      │
  │  │    OR Customer Content Types contains           │
  │  │    "All Content Types"                          │
  │  │                                                 │
  │  │  REGION MATCH:                                  │
  │  │    Customer Regions contains ANY value from     │
  │  │    Document Regions                             │
  │  │    OR Customer Regions contains "Global"        │
  │  │                                                 │
  │  │  TIER MATCH:                                    │
  │  │    If Urgency = Critical/Safety → all tiers     │
  │  │    If Urgency = High or Standard → Standard     │
  │  │    and Comprehensive only                       │
  │  │    If Urgency = Informational → Comprehensive   │
  │  │    only                                         │
  │  └─────────────────────────────────────────────────┘
  │       │
  │       ├── NO MATCH → Skip (do nothing, next customer)
  │       │
  │       └── MATCH → Continue to Step 3b
  │              │
  │              ▼
  │  STEP 3b: CHECK PREFERRED FREQUENCY
  │       │
  │       ├── "Immediate" OR Urgency = Critical/Safety
  │       │    → Go to Step 3c (send now)
  │       │
  │       ├── "Daily Digest"
  │       │    → Add to digest queue (separate SharePoint 
  │       │      list item) → skip email for now
  │       │
  │       └── "Weekly Digest"
  │            → Add to digest queue → skip email for now
  │              │
  │              ▼
  │  STEP 3c: SEND EMAIL (Outlook - Send an email V2)
  │       │
  │       │  To: Customer Email
  │       │  CC: CC Emails (if populated)
  │       │  Subject: [PSI {Content Type}] {Urgency} — 
  │       │           {Document Title} — {Products}
  │       │  Body: HTML template (see Section 8)
  │       │    • Publication Number + Title
  │       │    • Summary
  │       │    • Affected Products + Markets
  │       │    • Action Required
  │       │    • Link to document in portal
  │       │  Attachment: Optional — include the PDF from
  │       │              the document library
  │       │
  │       │  ⚠️ NOTE ON OUTLOOK SENDING LIMITS:
  │       │  Exchange Online allows ~10,000 recipients/day
  │       │  but rate limits apply. For large distributions
  │       │  (500+), add a "Delay" action (e.g., 5-second
  │       │  pause between emails) to avoid throttling.
  │              │
  │              ▼
  │  STEP 3d: LOG THE NOTIFICATION
  │       │
  │       │  Action: Create item in "Distribution Log" list
  │       │    • Log Entry: "{Pub Number} → {Contact Name}"
  │       │    • Publication Number, Title, Content Type, 
  │       │      Urgency
  │       │    • Recipient Name, Company, Email
  │       │    • Sent Date: utcNow()
  │       │    • Delivery Status: "Sent"
  │       │    • Match Reason: "Products: {matched tags}, 
  │       │      Markets: {matched tags}"
  │       │
  │       └── (Loop continues to next customer)
         │
         ▼
STEP 4: UPDATE DOCUMENT METADATA
  → Action: Update file properties
  → Set Distribution Status = "Distributed"
  → Set Date Published = utcNow()
  → Set Recipients Count = count of emails sent
         │
         ▼
STEP 5: NOTIFY ADMIN
  → Action: Send email or Teams message to admin
  → "{Publication Number} distributed to {X} recipients"
         │
         ▼
END
```

### 7.2 Flow 2: Daily Digest

**Trigger:** Recurrence — Daily at 8:00 AM CT

```
TRIGGER: Recurrence (Daily, 8:00 AM)
         │
         ▼
STEP 1: GET DIGEST QUEUE ITEMS
  → Get items from "Digest Queue" list
  → Filter: Sent = No, Frequency = Daily
         │
         ▼
STEP 2: GROUP BY CUSTOMER
  → Use "Select" and "Filter array" actions to group
     pending publications by recipient email
         │
         ▼
STEP 3: FOR EACH CUSTOMER WITH PENDING ITEMS
  │
  │  STEP 3a: COMPOSE DIGEST EMAIL
  │    → Build HTML body listing all pending publications
  │      for this customer (title, number, summary, link
  │      for each)
  │
  │  STEP 3b: SEND DIGEST EMAIL
  │    → Subject: [PSI Daily Digest] {X} New Publications
  │    → Body: Compiled digest HTML
  │
  │  STEP 3c: LOG EACH PUBLICATION NOTIFICATION
  │    → Create Distribution Log entry for each pub
  │
  │  STEP 3d: MARK QUEUE ITEMS AS SENT
  │    → Update Digest Queue items: Sent = Yes
  │
  └── (Next customer)
         │
         ▼
END
```

### 7.3 Flow 3: Acknowledgment Reminder (Critical/Safety)

**Trigger:** Recurrence — Every 12 hours

```
TRIGGER: Recurrence (Every 12 hours)
         │
         ▼
STEP 1: QUERY DISTRIBUTION LOG
  → Filter: Urgency = "Critical/Safety"
     AND Acknowledged = No
     AND Sent Date < (utcNow minus 48 hours)
         │
         ▼
STEP 2: FOR EACH OVERDUE ITEM
  │
  │  STEP 2a: SEND REMINDER EMAIL
  │    → Subject: [PSI REMINDER] Acknowledgment Required —
  │              {Publication Title}
  │    → Body: "This critical safety publication requires 
  │      your acknowledgment. Please confirm receipt by 
  │      clicking the link below."
  │
  └── (Next overdue item)
         │
         ▼
STEP 3: NOTIFY ADMIN
  → Summary of all overdue acknowledgments
         │
         ▼
END
```

### 7.4 Implementing Tag Matching in Power Automate

The tag matching logic is the most complex part of the flow. Here's how to implement it with Power Automate's expression language:

**For Choice (multi-select) columns:**

SharePoint multi-select choice columns return values as a semicolon-separated string (e.g., `"8.8L GSI; 22L DSI; 4.3L GSI"`). The document's Products column might contain `"8.8L GSI; 8.8L DSI"`.

To check if a customer's products overlap with the document's products:

```
Method: Use "contains()" checks or split into arrays and compare

Expression for Product Match:
  OR(
    contains(customerProducts, 'All Products'),
    // For each document product, check if it appears in customer products
    contains(customerProducts, documentProduct1),
    contains(customerProducts, documentProduct2),
    ... etc.
  )
```

**Practical approach for Power Automate:**

Since the document may have multiple product tags and you need to check if ANY of them match ANY of the customer's tags, the cleanest approach is:

1. **Split** the document's Products string into an array: `split(documentProducts, '; ')`
2. **Split** the customer's Products string into an array: `split(customerProducts, '; ')`
3. Use an **intersection** check: `length(intersection(docProductsArray, custProductsArray))` — if > 0, there's a match
4. OR check if customer contains "All Products"

**Full matching expression (compose action):**

```
// Product Match
@or(
  contains(items('Apply_to_each')?['Products'], 'All Products'),
  greater(
    length(
      intersection(
        split(triggerBody()?['Products'], '; '),
        split(items('Apply_to_each')?['Products'], '; ')
      )
    ),
    0
  )
)

// Repeat similar logic for Markets, Content Types, Regions
// Then combine all four with AND:

@and(productMatch, marketMatch, contentTypeMatch, regionMatch, tierMatch)
```

**For Managed Metadata columns:**

Managed Metadata values are stored differently (as term GUIDs + labels). You'll need to use "Send an HTTP request to SharePoint" to get the raw values, or use the label text with similar string matching.

**Recommendation:** Start with Choice (multi-select) columns for simplicity. The `contains()` and `intersection()` approach works reliably. Migrate to Managed Metadata later if you need centralized term governance.

---

## 8. Email Template (HTML)

Use this HTML template in the Power Automate "Send an email" action body. Dynamic values are shown as `{placeholders}`.

```html
<div style="font-family: Arial, sans-serif; max-width: 650px; margin: 0 auto; color: #333;">

  <!-- Header -->
  <div style="background-color: #003366; padding: 20px 30px; color: white;">
    <h2 style="margin: 0; font-size: 18px;">
      POWER SOLUTIONS INTERNATIONAL
    </h2>
    <p style="margin: 5px 0 0; font-size: 12px; color: #99bbdd;">
      Publication Notification
    </p>
  </div>

  <!-- Urgency Banner (conditional — only for High and Critical) -->
  <!-- Critical/Safety: background #cc0000 -->
  <!-- High: background #e67700 -->
  <div style="background-color: {urgencyColor}; color: white; padding: 10px 30px; font-weight: bold; font-size: 14px;">
    ■ {urgencyLevel} PRIORITY
  </div>

  <!-- Body -->
  <div style="padding: 30px; background-color: #f9f9f9; border: 1px solid #ddd;">

    <table style="width: 100%; font-size: 14px; margin-bottom: 20px;">
      <tr>
        <td style="color: #666; width: 140px; padding: 5px 0;">Publication:</td>
        <td style="font-weight: bold;">{publicationNumber}</td>
      </tr>
      <tr>
        <td style="color: #666; padding: 5px 0;">Type:</td>
        <td>{contentType}</td>
      </tr>
      <tr>
        <td style="color: #666; padding: 5px 0;">Products:</td>
        <td>{products}</td>
      </tr>
      <tr>
        <td style="color: #666; padding: 5px 0;">Markets:</td>
        <td>{markets}</td>
      </tr>
      <tr>
        <td style="color: #666; padding: 5px 0;">Date:</td>
        <td>{datePublished}</td>
      </tr>
    </table>

    <hr style="border: none; border-top: 1px solid #ddd;">

    <h3 style="color: #003366; margin-top: 20px;">
      {documentTitle}
    </h3>

    <p style="line-height: 1.6;">
      {summary}
    </p>

    <!-- Action Required (conditional — only if populated) -->
    <div style="background-color: #fff3cd; border-left: 4px solid #e67700; padding: 15px; margin: 20px 0;">
      <strong>Action Required:</strong><br>
      {actionRequired}
    </div>

    <!-- CTA Button -->
    <div style="text-align: center; margin: 30px 0;">
      <a href="{documentLink}" 
         style="background-color: #003366; color: white; padding: 14px 40px; 
                text-decoration: none; border-radius: 4px; font-weight: bold;
                font-size: 15px; display: inline-block;">
        View Full Document
      </a>
    </div>

  </div>

  <!-- Footer -->
  <div style="padding: 20px 30px; font-size: 11px; color: #999; background-color: #f0f0f0;">
    <p>
      You received this notification because your PSI distribution profile 
      includes: {matchedTags}.
    </p>
    <p>
      To update your notification preferences, contact PSI Customer Care 
      at <a href="mailto:customercare@psiengines.com">customercare@psiengines.com</a> 
      or call 1-800-XXX-XXXX.
    </p>
    <p>© 2026 Power Solutions International. All rights reserved.</p>
  </div>

</div>
```

---

## 9. Customer-Facing Portal Access

Optionally, give customers direct access to browse and download their documents via a SharePoint site with external sharing.

### 9.1 External Access Options

| Approach | How It Works | Pros | Cons |
|---|---|---|---|
| **SharePoint External Sharing (per-user)** | Invite external users to the SharePoint site. They sign in with their email. | Secure. Familiar. Auditable. | Requires Azure AD B2B guest accounts. Each customer needs an invite. |
| **SharePoint External Link (per-document)** | Each email includes a direct link to the document with "Anyone with the link" or "Specific people" access. | Simple. No portal login needed. | Less secure. No browsing capability. |
| **SharePoint Pages + Audience Targeting** | Build a SharePoint site with document web parts that use audience targeting to show only relevant documents. | Good UX. Customers see only their content. | Complex to configure. Requires Azure AD groups per segment. |
| **Separate Public Portal (Future)** | Build a customer-facing web app that authenticates customers and serves documents based on their profile. | Best UX. Full control. | Development effort. |

### 9.2 Recommended Starting Approach

Start with the simplest approach: **the email contains a direct link to the document in SharePoint**, shared via "Specific people" external sharing. No portal login required for the customer — they click the link, verify their email, and access the document.

As the system matures, build out a full customer portal (either SharePoint-based or custom) where customers can log in, browse all their publications, manage preferences, and acknowledge critical items.

---

## 10. Implementation Roadmap

### Phase 1: Build the Foundation (Weeks 1-2)

- [ ] Create SharePoint site: "PSI Publication Distribution"
- [ ] Define the complete tag taxonomy and document it as the canonical reference
- [ ] **If using Managed Metadata:** Configure Term Store with term sets for Products, Markets, Content Types, Regions
- [ ] **If using Choice columns:** Define all values identically on both the library and the list
- [ ] Build the Publications document library with all metadata columns
- [ ] Build the Customer Profiles list with all columns
- [ ] Build the Distribution Log list with all columns
- [ ] Create saved views for each list/library

### Phase 2: Build the Automation (Weeks 2-4)

- [ ] Build Power Automate Flow 1: Automatic Distribution on Publish
  - [ ] Trigger configuration (on file properties modified, Distribution Status = Approved)
  - [ ] Get document metadata
  - [ ] Get all active customer profiles
  - [ ] Implement tag matching logic (product, market, content type, region, tier)
  - [ ] Send email action with HTML template
  - [ ] Create Distribution Log entry
  - [ ] Update document metadata (status, date, count)
  - [ ] Admin notification
- [ ] Build Power Automate Flow 2: Daily Digest (if needed for Phase 1)
- [ ] Build Power Automate Flow 3: Acknowledgment Reminder
- [ ] Create the HTML email template
- [ ] Test all flows with sample data

### Phase 3: Load Data & Test (Weeks 4-5)

- [ ] Enter initial customer profiles (start with 25-50 key accounts)
  - [ ] Assign Product, Market, Content Type, Region tags to each
  - [ ] Set Subscription Tier and Preferred Frequency
- [ ] Upload 3-5 test publications with varying metadata
- [ ] Run end-to-end tests:
  - [ ] Verify tag matching produces correct recipient lists
  - [ ] Verify emails are sent with correct content and formatting
  - [ ] Verify Distribution Log entries are created
  - [ ] Test edge cases: "All Products" wildcard, multiple tag matches, no matches
  - [ ] Test urgency-based tier filtering (Essential only gets Critical/Safety)
- [ ] Send test distributions to internal PSI team acting as "customers"
- [ ] Fix issues

### Phase 4: Pilot (Weeks 5-7)

- [ ] Select 10-15 external pilot customers across different products/markets
- [ ] Communicate the new system to pilot customers
- [ ] Distribute 3-5 real publications through the system
- [ ] Collect feedback on email content, frequency, relevance
- [ ] Iterate on template, tag structure, and workflow

### Phase 5: Full Rollout (Weeks 7-12)

- [ ] Load all remaining customer profiles
- [ ] Train internal authors on: upload document → fill metadata → submit for approval
- [ ] Train reviewers on: approval process
- [ ] Go live — all new publications flow through the portal
- [ ] Monitor delivery, log entries, and customer feedback
- [ ] First quarterly data hygiene review

### Phase 6: Enhancements (Ongoing)

- [ ] Build customer-facing portal for self-service browsing and preference management
- [ ] Add acknowledgment workflow for Critical/Safety items
- [ ] Add weekly digest flow
- [ ] Integrate with PSI's CRM/ERP for automatic customer profile syncing
- [ ] Add Power BI dashboards for advanced analytics
- [ ] Evaluate migration to custom web portal if SharePoint is outgrown

---

## 11. Governance

### 11.1 Taxonomy Management

| Action | Frequency | Owner |
|---|---|---|
| Review and update Product tags | When new products launch or are discontinued | Product Management / Engineering |
| Review and update Market tags | Annually or when entering new markets | Product Management |
| Audit Customer Profiles for accuracy | Quarterly | Customer Care |
| Review bounced/failed notifications | Weekly | System Admin |
| Review unacknowledged Critical items | Within 48 hours of distribution | System Admin |

### 11.2 Data Hygiene

- **New customer onboarding:** Add profile to Customer Profiles list as part of the onboarding process
- **Customer offboarding:** Set Status to Inactive (don't delete — preserves audit trail)
- **Contact changes:** Update email, name, tags when notified by the customer
- **Bounce management:** If a Distribution Log shows "Failed" for a customer, flag their profile for review
- **Quarterly review:** Admin runs a report of all Active profiles and validates with the sales/account management team

### 11.3 Roles

| Role | Responsibilities |
|---|---|
| **System Owner** | Overall accountability. Defines taxonomy. Approves process changes. |
| **System Admin** | Manages Power Automate flows. Troubleshoots issues. Maintains SharePoint site. |
| **Publication Authors** | Upload documents. Fill in metadata correctly. Submit for approval. |
| **Reviewers/Approvers** | Validate metadata and content. Approve for distribution. |
| **Customer Care** | Manage customer profiles. Handle preference change requests. |

---

## 12. Costs

| Item | Cost | Notes |
|---|---|---|
| SharePoint Online | Included in M365 | Already licensed |
| Power Automate | Included in M365 | Standard connectors (SharePoint, Outlook) are included. Premium connectors not needed for this design. |
| Outlook / Exchange Online | Included in M365 | Email delivery |
| Additional tooling | $0 | No external services required |
| **Total incremental cost** | **$0** | Fully built on existing M365 infrastructure |

---

## 13. Limitations & Considerations

| Limitation | Mitigation |
|---|---|
| **Exchange Online sending limits** (~10,000 recipients/day) | Add delays between emails in Power Automate. For very large distributions, consider SendGrid as the email backend. |
| **SharePoint list item limit** (theoretically 30M, practically performant to ~100K items) | Distribution Log will grow over time. Archive old entries annually. |
| **Power Automate "Apply to each" performance** (sequential by default) | For 500+ customers, the flow may take 30-60+ minutes. Use Concurrency Control (parallel branches, up to 50) to speed up. |
| **No built-in open/read tracking** from Outlook | If delivery analytics are critical, add SendGrid as the email service (Power Automate has a SendGrid connector). |
| **Multi-select Choice column matching** requires string manipulation in Power Automate | The `intersection()` approach in Section 7.4 handles this, but it's not as elegant as a database query. Works reliably for hundreds of customers. |
| **External customer portal access** requires Azure AD B2B or anonymous links | Start with email links. Build out portal access as a later enhancement. |

---

## 14. Quick Reference: The Complete Workflow

```
AUTHOR                          SHAREPOINT                      POWER AUTOMATE               CUSTOMER
  │                                 │                                │                           │
  │  1. Upload document             │                                │                           │
  │  2. Fill in metadata            │                                │                           │
  │     (products, markets,         │                                │                           │
  │      content type, urgency,     │                                │                           │
  │      summary, action req'd)     │                                │                           │
  │  3. Submit for approval         │                                │                           │
  │────────────────────────────────▶│                                │                           │
  │                                 │                                │                           │
  │              REVIEWER approves  │                                │                           │
  │                                 │  Status → "Approved"           │                           │
  │                                 │───────────────────────────────▶│                           │
  │                                 │                                │                           │
  │                                 │              Flow triggers:    │                           │
  │                                 │              • Reads metadata  │                           │
  │                                 │              • Gets all active │                           │
  │                                 │                customer profiles                          │
  │                                 │              • Matches tags    │                           │
  │                                 │              • For each match: │                           │
  │                                 │                                │                           │
  │                                 │                                │  Sends personalized email │
  │                                 │                                │─────────────────────────▶│
  │                                 │                                │                           │
  │                                 │                                │  Creates log entry        │
  │                                 │                                │─────────▶ Distribution    │
  │                                 │                                │           Log             │
  │                                 │                                │                           │
  │                                 │  Updates: Status → Distributed │                           │
  │                                 │◀──────────────────────────────│                           │
  │                                 │  Sets Recipients Count + Date  │                           │
  │                                 │                                │                           │
  │                                 │                                │  Notifies admin:          │
  │                                 │                                │  "SB-2026-0058 sent to    │
  │                                 │                                │   47 recipients"          │
```

---

*Document Version: 3.0*
*Created: February 2026*
*Platform: Microsoft 365 (SharePoint + Power Automate + Outlook)*
*Author: PSI Customer Care — Automated Distribution System*
