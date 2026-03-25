# 📚 Automated Web Scraping — Email CSV, Save to Google Sheets & Microsoft Excel

An **n8n automation workflow** that scrapes book data from a website, sorts it by price, exports it as a CSV, and simultaneously saves the data to **Google Sheets**, **Microsoft Excel 365**, and sends the CSV as an **email attachment via Gmail**.

---

## 🔄 Workflow Overview
```
Manual Trigger
     │
     ▼
Fetch Website Content (books.toscrape.com)
     │
     ▼
Extract All Books from Page (HTML → list items)
     │
     ▼
Split Out (one item per book)
     │
     ▼
Extract Individual Book Price & Title
     │
     ▼
Sort by Price (descending)
     │
     ├──▶ Save to Google Sheets
     ├──▶ Save to Microsoft Excel 365
     └──▶ Convert to CSV File ──▶ Send CSV via Gmail
```

---

## ✨ Features

- 🌐 **Web Scraping** — Fetches and parses HTML from any target website using CSS selectors
- 📊 **Google Sheets Integration** — Appends scraped data automatically to a connected spreadsheet
- 📑 **Microsoft Excel 365 Integration** — Writes data to an Excel workbook via Microsoft Graph API
- 📧 **Email Delivery** — Converts data to CSV and sends it as a Gmail attachment
- 🔃 **Sorting** — Books are sorted by price in descending order before export

---

## 🛠️ Setup Instructions

### 1. Import the Workflow
- Open your **n8n** instance
- Go to **Workflows → Import from File**
- Upload the `.json` file from this repo

### 2. Configure the Target Website
- Open the **"Fetch website content"** node and replace the URL with your target site
- Update CSS selectors in the HTML extraction nodes to match the site's structure

### 3. Google Cloud Credentials
Enable these APIs in [Google Cloud Console](https://console.cloud.google.com):
- ✅ Google Drive API
- ✅ Google Sheets API
- ✅ Gmail API

Connect credentials to the **Google Sheets** and **Gmail** nodes.

### 4. Microsoft Azure Credentials
- Register an app in [Azure Portal](https://portal.azure.com) with **Microsoft Graph** permissions
- Connect credentials to the **Microsoft Excel 365** node
- ⚠️ Pre-create column headers (`title`, `price`) in your Excel sheet before running

---

## 📦 Nodes Used

| Node | Purpose |
|------|---------|
| `Manual Trigger` | Starts the workflow on demand |
| `HTTP Request` | Fetches raw HTML from the target website |
| `HTML` (Extract all books) | Parses the page into a list of book elements |
| `Split Out` | Splits the array into individual items |
| `HTML` (Extract price) | Extracts title and price per book |
| `Sort` | Sorts books by price (descending) |
| `Convert to File` | Converts data to a `.csv` binary file |
| `Google Sheets` | Appends data to a Google Sheet |
| `Microsoft Excel 365` | Appends data to an Excel workbook |
| `Gmail` | Sends the CSV as an email attachment |

---

## 🔑 Prerequisites

- [n8n](https://n8n.io) (self-hosted or cloud)
- Google Cloud account with Drive, Sheets & Gmail APIs enabled
- Microsoft Azure account with Microsoft Graph permissions
- Gmail OAuth2 configured in n8n

---

## 📝 Notes

- Swap the **Manual Trigger** for a **Schedule Trigger** to automate runs (e.g., daily)
- Excel requires pre-existing column headers; Google Sheets auto-maps them

---

## 📄 License

Open for personal and educational use. Modify freely for your own data pipelines.
