// ============================================================
//  Ojaa Meats & More — WhatsApp Cloud API Bot
//  Node.js + Express
//  Free hosting: Render.com + cron-job.org
//  Payments: Manual EFT
//  Orders + Menu: Google Sheets via SheetDB
//  Invoices: Generated as PDF and sent via WhatsApp
// ============================================================

const express = require('express');
const axios   = require('axios');
const { execSync } = require('child_process');
const fs      = require('fs');
const path    = require('path');
const { google } = require('googleapis');
const app     = express();
app.use(express.json());

// ─── GOOGLE AUTH ─────────────────────────────────────────────
const SPREADSHEET_ID = '1G7-WkIBOjx99rdD2_v4sBlJx43VBuYRAo5LGcPldY8o';
const DRIVE_FOLDER_ID = '10VmiMRsTl6M4UTurYaZktPihjyUjRu6m';

function getGoogleAuth() {
  const credentials = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT);
  return new google.auth.GoogleAuth({
    credentials,
    scopes: [
      'https://www.googleapis.com/auth/spreadsheets',
      'https://www.googleapis.com/auth/drive',
    ],
  });
}

async function getSheetsClient() {
  const auth = getGoogleAuth();
  return google.sheets({ version: 'v4', auth });
}

async function getDriveClient() {
  const auth = getGoogleAuth();
  return google.drive({ version: 'v3', auth });
}

const CONFIG = {
  VERIFY_TOKEN:    process.env.VERIFY_TOKEN    || 'ojaa_verify_2026',
  WA_TOKEN:        process.env.WA_TOKEN        || 'YOUR_WHATSAPP_ACCESS_TOKEN',
  PHONE_NUMBER_ID: process.env.PHONE_NUMBER_ID || 'YOUR_PHONE_NUMBER_ID',
  BUSINESS_NAME:   process.env.BUSINESS_NAME   || 'Ojaa Meats & More',
  BANK_NAME:       process.env.BANK_NAME       || 'Standard Bank',
  ACCOUNT_NAME:    process.env.ACCOUNT_NAME    || 'CJ Manufacturing CC',
  ACCOUNT_NUMBER:  process.env.ACCOUNT_NUMBER  || '060348526',
  BRANCH_CODE:     process.env.BRANCH_CODE     || '051001',
  PORT: process.env.PORT || 3000,
};

// ─── INVOICE GENERATOR (Python via child process) ───────────
const INVOICE_SCRIPT = `
import sys, json, io, base64
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

def generate_invoice(order):
    buffer = io.BytesIO()
    w, h = A4
    c = canvas.Canvas(buffer, pagesize=A4)

    green = colors.HexColor('#2d6a2d')
    dark  = colors.HexColor('#1a1a1a')
    light = colors.HexColor('#f5f5f5')
    grey  = colors.HexColor('#666666')
    white = colors.white

    # HEADER
    c.setFillColor(green)
    c.rect(0, h - 58*mm, w, 58*mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont('Helvetica-Bold', 22)
    c.drawString(20*mm, h - 18*mm, 'C. J Manufacturing CC')
    c.setFont('Helvetica', 10)
    c.drawString(20*mm, h - 27*mm, 't/a Ojaa Meats & More')
    c.setFont('Helvetica', 9)
    c.drawString(20*mm, h - 36*mm, 'CK No: 91/32575/23')
    c.drawString(20*mm, h - 44*mm, '6 Torwood Drive, Hayfields, Pietermaritzburg, 3201')
    c.drawString(20*mm, h - 52*mm, 'VAT Reg No: 4200130013')
    c.setFont('Helvetica-Bold', 28)
    c.drawRightString(w - 20*mm, h - 20*mm, 'INVOICE')
    c.setFont('Helvetica', 10)
    c.drawRightString(w - 20*mm, h - 30*mm, '# ' + order['orderNum'])
    c.drawRightString(w - 20*mm, h - 40*mm, 'Date: ' + order['date'])

    # BILL TO
    c.setFillColor(light)
    c.rect(20*mm, h - 82*mm, w - 40*mm, 20*mm, fill=1, stroke=0)
    c.setFillColor(grey)
    c.setFont('Helvetica', 9)
    c.drawString(25*mm, h - 65*mm, 'BILL TO')
    c.setFillColor(dark)
    c.setFont('Helvetica-Bold', 11)
    c.drawString(25*mm, h - 73*mm, order['customerName'] or 'Valued Customer')
    c.setFont('Helvetica', 9)
    c.drawString(25*mm, h - 80*mm, order['phone'])

    # TABLE HEADER
    table_top = h - 94*mm
    c.setFillColor(dark)
    c.rect(20*mm, table_top - 8*mm, w - 40*mm, 8*mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont('Helvetica-Bold', 9)
    c.drawString(25*mm, table_top - 5.5*mm, 'DESCRIPTION')
    c.drawRightString(105*mm, table_top - 5.5*mm, 'QTY')
    c.drawRightString(145*mm, table_top - 5.5*mm, 'UNIT PRICE')
    c.drawRightString(w - 20*mm, table_top - 5.5*mm, 'TOTAL')

    # ITEMS
    y = table_top - 8*mm
    subtotal = 0
    for i, item in enumerate(order['items']):
        row_h = 9*mm
        y -= row_h
        c.setFillColor(light if i % 2 == 0 else white)
        c.rect(20*mm, y, w - 40*mm, row_h, fill=1, stroke=0)
        line_total = item['price'] * item['qty']
        subtotal += line_total
        c.setFillColor(dark)
        c.setFont('Helvetica', 9)
        c.drawString(25*mm, y + 3*mm, item['name'] + ' (' + item['unit'] + ')')
        c.drawRightString(105*mm, y + 3*mm, str(item['qty']))
        c.drawRightString(145*mm, y + 3*mm, 'R ' + '{:.2f}'.format(item['price']))
        c.drawRightString(w - 20*mm, y + 3*mm, 'R ' + '{:.2f}'.format(line_total))

    # TOTALS
    y -= 5*mm
    vat_excl   = round(subtotal / 1.15, 2)
    vat_amount = round(subtotal - vat_excl, 2)

    def total_row(label, amount, bold=False, highlight=False):
        nonlocal y
        y -= 8*mm
        if highlight:
            c.setFillColor(green)
            c.rect(105*mm, y, w - 125*mm, 8*mm, fill=1, stroke=0)
            c.setFillColor(white)
        else:
            c.setFillColor(dark)
        c.setFont('Helvetica-Bold' if bold else 'Helvetica', 10)
        c.drawRightString(145*mm, y + 2.5*mm, label)
        c.drawRightString(w - 20*mm, y + 2.5*mm, 'R ' + '{:.2f}'.format(amount))

    total_row('Sub Total (excl. VAT)', vat_excl)
    total_row('VAT 15%', vat_amount)
    total_row('TOTAL DUE (incl. VAT)', subtotal, bold=True, highlight=True)

    # BANKING DETAILS
    y -= 15*mm
    box_h = 42*mm
    c.setFillColor(light)
    c.rect(20*mm, y - box_h + 5*mm, w - 40*mm, box_h, fill=1, stroke=0)
    c.setFillColor(green)
    c.setFont('Helvetica-Bold', 10)
    c.drawString(25*mm, y + 1*mm, 'BANKING DETAILS')
    details = [
        ('Bank', 'Standard Bank'),
        ('Account Name', 'CJ Manufacturing CC'),
        ('Account Number', '060348526'),
        ('Branch Code', '051001'),
        ('Reference', order['orderNum']),
        ('Email POP to', 'bronnie@infin.co.za'),
    ]
    dy = y - 7*mm
    c.setFillColor(dark)
    for label, value in details:
        c.setFont('Helvetica-Bold', 9)
        c.drawString(25*mm, dy, label + ':')
        c.setFont('Helvetica', 9)
        c.drawString(75*mm, dy, value)
        dy -= 5.5*mm

    # FOOTER
    c.setFillColor(green)
    c.rect(0, 0, w, 12*mm, fill=1, stroke=0)
    c.setFillColor(white)
    c.setFont('Helvetica', 8)
    c.drawCentredString(w/2, 4*mm, 'C. J Manufacturing CC t/a Ojaa Meats & More | VAT Reg No: 4200130013 | bronnie@infin.co.za')

    c.save()
    buffer.seek(0)
    return base64.b64encode(buffer.read()).decode()

order = json.loads(sys.argv[1])
print(generate_invoice(order))
`;

async function generateInvoicePDF(orderData) {
  try {
    const scriptPath = '/tmp/gen_invoice.py';
    fs.writeFileSync(scriptPath, INVOICE_SCRIPT);
    const orderJson = JSON.stringify(orderData).replace(/'/g, "\\'");
    const result = execSync(`python3 ${scriptPath} '${orderJson}'`, { timeout: 30000 });
    return Buffer.from(result.toString().trim(), 'base64');
  } catch (err) {
    console.error('Invoice generation error:', err.message);
    return null;
  }
}

async function uploadInvoiceToCloudinary(pdfBuffer, orderNum) {
  try {
    const cloudName = process.env.CLOUDINARY_CLOUD_NAME;
    const apiKey    = process.env.CLOUDINARY_API_KEY;
    const apiSecret = process.env.CLOUDINARY_API_SECRET;

    const base64 = pdfBuffer.toString('base64');
    const dataUri = `data:application/pdf;base64,${base64}`;

    const timestamp = Math.floor(Date.now() / 1000);
    const publicId  = `invoices/${orderNum}.pdf`;
    const crypto    = require('crypto');
    // Params must be alphabetical for Cloudinary signature
    const sigString = `access_mode=public&public_id=${publicId}&timestamp=${timestamp}&type=upload${apiSecret}`;
    const signature = crypto
      .createHash('sha1')
      .update(sigString)
      .digest('hex');

    const formData = new (require('form-data'))();
    formData.append('file', dataUri);
    formData.append('public_id', publicId);
    formData.append('timestamp', timestamp);
    formData.append('api_key', apiKey);
    formData.append('signature', signature);
    formData.append('resource_type', 'raw');
    formData.append('type', 'upload');
    formData.append('access_mode', 'public');

    const res = await axios.post(
      `https://api.cloudinary.com/v1_1/${cloudName}/auto/upload`,
      formData,
      { headers: formData.getHeaders() }
    );

    console.log(`Invoice ${orderNum} uploaded to Cloudinary`);
    return res.data.secure_url;
  } catch (err) {
    console.error('Cloudinary upload error:', err.response?.data || err.message);
    return null;
  }
}

async function getOrCreateFolder(drive, name, parentId) {
  // Check if folder already exists
  const res = await drive.files.list({
    q: `name='${name}' and mimeType='application/vnd.google-apps.folder' and '${parentId}' in parents and trashed=false`,
    fields: 'files(id, name)',
  });
  if (res.data.files.length > 0) {
    return res.data.files[0].id;
  }
  // Create it
  const folder = await drive.files.create({
    requestBody: {
      name,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [parentId],
    },
    fields: 'id',
  });
  return folder.data.id;
}

async function saveInvoiceToDrive(pdfBuffer, orderNum) {
  try {
    const drive = await getDriveClient();
    const now = new Date();
    const month = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

    // Get or create month folder
    const monthFolderId = await getOrCreateFolder(drive, month, DRIVE_FOLDER_ID);

    const { Readable } = require('stream');
    const stream = new Readable();
    stream.push(pdfBuffer);
    stream.push(null);
    await drive.files.create({
      requestBody: {
        name: `Invoice-${orderNum}.pdf`,
        parents: [monthFolderId],
        mimeType: 'application/pdf',
      },
      media: {
        mimeType: 'application/pdf',
        body: stream,
      },
    });
    console.log(`Invoice ${orderNum} saved to Drive under ${day}`);
  } catch (err) {
    console.error('Drive save error:', err.message);
  }
}

async function uploadAndSendInvoice(phone, pdfBuffer, orderNum) {
  try {
    // Upload PDF as media to WhatsApp
    const FormData = require('form-data');
    const form = new FormData();
    form.append('file', pdfBuffer, {
      filename: `Invoice-${orderNum}.pdf`,
      contentType: 'application/pdf',
    });
    form.append('messaging_product', 'whatsapp');
    form.append('type', 'application/pdf');

    const uploadRes = await axios.post(
      `https://graph.facebook.com/v19.0/${CONFIG.PHONE_NUMBER_ID}/media`,
      form,
      {
        headers: {
          Authorization: `Bearer ${CONFIG.WA_TOKEN}`,
          ...form.getHeaders(),
        },
      }
    );

    const mediaId = uploadRes.data.id;

    // Send document message
    await axios.post(
      `https://graph.facebook.com/v19.0/${CONFIG.PHONE_NUMBER_ID}/messages`,
      {
        messaging_product: 'whatsapp',
        to: phone,
        type: 'document',
        document: {
          id: mediaId,
          filename: `Invoice-${orderNum}.pdf`,
          caption: `Your invoice for order ${orderNum}. Please use the order number as your payment reference.`,
        },
      },
      {
        headers: {
          Authorization: `Bearer ${CONFIG.WA_TOKEN}`,
          'Content-Type': 'application/json',
        },
      }
    );
    console.log(`Invoice sent to ${phone} for order ${orderNum}`);
    return true;
  } catch (err) {
    console.error('Invoice send error:', err.response?.data || err.message);
    return false;
  }
}

// ─── MENU CACHE ──────────────────────────────────────────────
let menuCache = [];
let menuCacheTime = 0;
const MENU_CACHE_TTL = 10 * 60 * 1000;

async function getMenu() {
  const now = Date.now();
  if (menuCache.length > 0 && (now - menuCacheTime) < MENU_CACHE_TTL) {
    return menuCache;
  }
  try {
    const sheets = await getSheetsClient();
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Menu',
    });
    const rows = res.data.values || [];
    if (rows.length > 1) {
      const headers = rows[0];
      const nameIdx  = headers.indexOf('Name');
      const unitIdx  = headers.indexOf('Unit');
      const priceIdx = headers.findIndex(h => h.includes('Price'));
      menuCache = rows.slice(1).map((row, index) => ({
        id: index,
        letter: String.fromCharCode(65 + index),
        name: row[nameIdx] || '',
        unit: row[unitIdx] || '',
        price: parseFloat(row[priceIdx]) || 0,
      }));
      menuCacheTime = now;
    }
  } catch (err) {
    console.error('Menu load error:', err.message);
  }
  return menuCache;
}

// ─── STOCK ON HAND ───────────────────────────────────────────
let stockCache = {};
let stockCacheTime = 0;
const STOCK_CACHE_TTL = 60 * 1000; // 1 min cache

async function getStockOnHand() {
  const now = Date.now();
  if (Object.keys(stockCache).length > 0 && (now - stockCacheTime) < STOCK_CACHE_TTL) {
    return stockCache;
  }
  try {
    const sheets = await getSheetsClient();
    const [stockRes, lineItemsRes] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Stock on hand' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Line items' }),
    ]);

    const stockRows = stockRes.data.values || [];
    if (stockRows.length < 4) { stockCache = {}; return stockCache; }

    // Date is in cell C2 (row index 1, col index 2)
    const stockDateStr = stockRows[1]?.[2] || '';
    // Parse DD/MM/YYYY
    let stockDate = null;
    if (stockDateStr) {
      const parts = stockDateStr.split('/');
      if (parts.length === 3) {
        stockDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
      }
    }

    // Read items + starting stock - rows from index 3 onwards (header in row 3, data row 4+)
    // Headers row is at index 3: Item | Starting stock | Orders | Stock on hand
    const headerRow = stockRows[3] || [];
    const itemIdx = headerRow.indexOf('Item');
    const startIdx = headerRow.indexOf('Starting stock');
    const items = {};
    for (let i = 4; i < stockRows.length; i++) {
      const row = stockRows[i];
      const name = row[itemIdx];
      const start = parseFloat(row[startIdx]) || 0;
      if (name) items[name] = { start, sold: 0 };
    }

    // Sum line items since stock date
    const liRows = lineItemsRes.data.values || [];
    if (liRows.length > 1) {
      const liHeaders = liRows[0];
      const tsIdx  = liHeaders.indexOf('Timestamp');
      const prodIdx = liHeaders.indexOf('Product');
      const qtyIdx  = liHeaders.indexOf('Qty');
      for (let i = 1; i < liRows.length; i++) {
        const r = liRows[i];
        const ts = new Date(r[tsIdx]);
        if (stockDate && ts < stockDate) continue;
        const prod = r[prodIdx];
        const qty = parseFloat(r[qtyIdx]) || 0;
        if (items[prod]) items[prod].sold += qty;
      }
    }

    // Build stock-on-hand map
    const result = {};
    for (const [name, v] of Object.entries(items)) {
      result[name] = v.start - v.sold;
    }
    stockCache = result;
    stockCacheTime = now;
  } catch (err) {
    console.error('Stock load error:', err.message);
  }
  return stockCache;
}

async function updateStockSheet() {
  try {
    const sheets = await getSheetsClient();
    const [stockRes, lineItemsRes] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Stock on hand' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Line items' }),
    ]);
    const stockRows = stockRes.data.values || [];
    if (stockRows.length < 5) return;

    // Parse stock date from C2 (DD/MM/YYYY)
    const stockDateStr = stockRows[1]?.[2] || '';
    let stockDate = null;
    if (stockDateStr) {
      const parts = stockDateStr.split('/');
      if (parts.length === 3) {
        stockDate = new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
      }
    }

    const headerRow = stockRows[3] || [];
    const itemIdx   = headerRow.indexOf('Item');
    const startIdx  = headerRow.indexOf('Starting stock');
    const ordersIdx = headerRow.indexOf('Orders');
    const sohIdx    = headerRow.indexOf('Stock on hand');

    // Sum sold quantities by product since stock date
    const soldByProduct = {};
    const liRows = lineItemsRes.data.values || [];
    if (liRows.length > 1) {
      const liHeaders = liRows[0];
      const tsIdx   = liHeaders.indexOf('Timestamp');
      const prodIdx = liHeaders.indexOf('Product');
      const qtyIdx  = liHeaders.indexOf('Qty');
      for (let i = 1; i < liRows.length; i++) {
        const r = liRows[i];
        const ts = new Date(r[tsIdx]);
        if (stockDate && ts < stockDate) continue;
        const prod = r[prodIdx];
        const qty = parseFloat(r[qtyIdx]) || 0;
        soldByProduct[prod] = (soldByProduct[prod] || 0) + qty;
      }
    }

    // Build update batch
    const updates = [];
    for (let i = 4; i < stockRows.length; i++) {
      const row = stockRows[i];
      const name = row[itemIdx];
      if (!name) continue;
      const start  = parseFloat(row[startIdx]) || 0;
      const sold   = soldByProduct[name] || 0;
      const onHand = start - sold;
      // Update Orders column and Stock on hand column for this row
      const rowNum = i + 1; // 1-indexed
      const ordersCol = String.fromCharCode(65 + ordersIdx);
      const sohCol    = String.fromCharCode(65 + sohIdx);
      updates.push({ range: `Stock on hand!${ordersCol}${rowNum}`, values: [[sold]] });
      updates.push({ range: `Stock on hand!${sohCol}${rowNum}`, values: [[onHand]] });
    }

    if (updates.length > 0) {
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { valueInputOption: 'USER_ENTERED', data: updates },
      });
    }

    // Invalidate stock cache so next read is fresh
    stockCacheTime = 0;
  } catch (err) {
    console.error('Stock sheet update error:', err.message);
  }
}

// ─── SESSIONS ────────────────────────────────────────────────
const sessions = {};
function getSession(phone) {
  if (!sessions[phone]) {
    sessions[phone] = { state: 'welcome', cart: {}, orderNum: null, pendingItem: null, deliveryType: null, deliveryAddress: null };
  }
  return sessions[phone];
}

// ─── SHEETDB ─────────────────────────────────────────────────
// ─── BOT STATUS CHECK ────────────────────────────────────────
// Reads Control sheet C5 — if "Off" bot sends away message
let botStatusCache = 'On';
let botStatusCacheTime = 0;
const STATUS_CACHE_TTL = 60 * 1000; // check every 60 seconds

async function getBotStatus() {
  const now = Date.now();
  if (now - botStatusCacheTime < STATUS_CACHE_TTL) return botStatusCache;
  try {
    const sheets = await getSheetsClient();
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: 'Control',
    });
    const rows = res.data.values || [];
    if (rows.length > 0) {
      let found = null;
      for (const row of rows) {
        for (const val of row) {
          if (val === 'On' || val === 'Off') {
            found = val;
            break;
          }
        }
        if (found) break;
      }
      if (found) {
        const previous = botStatusCache;
        botStatusCache = found;
        botStatusCacheTime = now;
        console.log(`Bot status: ${botStatusCache}`);

        // If bot just switched back On — reset any incomplete sessions
        if (previous === 'Off' && botStatusCache === 'On') {
          console.log('Bot switched On — resetting incomplete sessions');
          const incomplete = ['ordering', 'confirm', 'payment'];
          for (const phone of Object.keys(sessions)) {
            if (incomplete.includes(sessions[phone].state)) {
              sessions[phone] = { state: 'welcome', cart: {}, orderNum: null };
              console.log(`Reset session for ${phone}`);
            }
          }
        }
      }
    }
  } catch (err) {
    console.error('Status check error:', err.message);
  }
  return botStatusCache;
}

// ─── GOOGLE SHEETS HELPERS ───────────────────────────────────
async function sheetsGet(sheetName, searchCol, searchVal) {
  try {
    const sheets = await getSheetsClient();
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: sheetName,
    });
    const rows = res.data.values || [];
    if (rows.length === 0) return [];
    const headers = rows[0];
    const colIndex = headers.indexOf(searchCol);
    if (colIndex === -1) return [];
    return rows.slice(1)
      .filter(row => row[colIndex] === searchVal)
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i] || '');
        return obj;
      });
  } catch (err) {
    console.error('Sheets GET error:', err.message);
    return [];
  }
}

async function sheetsPost(sheetName, rowData) {
  try {
    const sheets = await getSheetsClient();
    // Get headers first
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!1:1`,
    });
    const headers = res.data.values?.[0] || Object.keys(rowData);
    const row = headers.map(h => rowData[h] !== undefined ? String(rowData[h]) : '');
    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: sheetName,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [row] },
    });
    return true;
  } catch (err) {
    console.error('Sheets POST error:', err.message);
    return null;
  }
}

async function sheetsUpdate(sheetName, searchCol, searchVal, rowData) {
  try {
    const sheets = await getSheetsClient();
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: sheetName,
    });
    const rows = res.data.values || [];
    if (rows.length === 0) return null;
    const headers = rows[0];
    const colIndex = headers.indexOf(searchCol);
    if (colIndex === -1) return null;
    const rowIndex = rows.findIndex((row, i) => i > 0 && row[colIndex] === searchVal);
    if (rowIndex === -1) return null;
    const updatedRow = headers.map((h, i) => rowData[h] !== undefined ? String(rowData[h]) : (rows[rowIndex][i] || ''));
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${sheetName}!A${rowIndex + 1}`,
      valueInputOption: 'USER_ENTERED',
      requestBody: { values: [updatedRow] },
    });
    return true;
  } catch (err) {
    console.error('Sheets UPDATE error:', err.message);
    return null;
  }
}

async function updateSessionInSheets(phone, session) {
  const rows = await sheetsGet('Sessions', 'Phone', phone);
  const rowData = { State: session.state, Cart: JSON.stringify(session.cart || {}), Order_Num: session.orderNum || '', Last_updated: new Date().toISOString() };
  if (rows && rows.length > 0) {
    await sheetsUpdate('Sessions', 'Phone', phone, rowData);
  } else {
    await sheetsPost('Sessions', { Phone: phone, ...rowData });
  }
}

async function logOrderToSheets(phone, orderNum, cart, total, customerName, deliveryType = '', deliveryAddress = '') {
  const items = Object.values(cart).map(i => `${i.name} x${i.qty}`).join(', ');
  const timestamp = new Date().toISOString();

  // Log summary to Orders sheet
  await sheetsPost('Orders', {
    Timestamp: timestamp,
    Order_num: orderNum,
    Phone: phone,
    Name: customerName || '',
    Items: items,
    VAT: (total - total / 1.15).toFixed(2),
    Total: total,
    Delivery_Type: deliveryType,
    Delivery_Address: deliveryAddress,
    Paid: 'NO',
    POP_Confirmed: 'NO',
    Delivered: 'NO',
    Invoice_URL: '',
  });

  // Log one row per product to Line_Items sheet
  for (const item of Object.values(cart)) {
    const lineTotal = item.price * item.qty;
    await sheetsPost('Line items', {
      Timestamp: timestamp,
      Order_Num: orderNum,
      Phone: phone,
      Name: customerName || '',
      Product: item.name,
      Unit: item.unit,
      Qty: item.qty,
      Unit_Price: item.price,
      Line_Total: lineTotal,
    });
  }
}

// ─── HELPERS ─────────────────────────────────────────────────
async function menuText() {
  const menu = await getMenu();
  let lines = [`*${CONFIG.BUSINESS_NAME} — Today's menu:*\n`];
  menu.forEach(m => lines.push(`${m.letter}. ${m.name} (${m.unit}) — R${m.price}`));
  lines.push('\nReply with the letter and quantity.');
  lines.push('Example: *Ax2, Cx1* means 2x beef mince + 1x boerewors');
  return lines.join('\n');
}

function cartSummary(cart) {
  let lines = ['*Your order:*\n'];
  let total = 0;
  Object.values(cart).forEach(item => {
    const sub = item.price * item.qty;
    total += sub;
    lines.push(`- ${item.name} x${item.qty} — R${sub}`);
  });
  lines.push(`\n*Total: R${total}*`);
  return { text: lines.join('\n'), total };
}

async function parseOrderText(text) {
  const menu = await getMenu();
  const parts = text.split(',');
  const items = {};
  const errors = [];
  parts.forEach(p => {
    const m = p.trim().match(/^([a-zA-Z])x(\d+)$/i);
    if (!m) { errors.push(`"${p.trim()}" not recognised`); return; }
    const letter = m[1].toUpperCase();
    const qty    = parseInt(m[2]);
    const prod   = menu.find(x => x.letter === letter);
    if (!prod)                { errors.push(`Item "${letter}" does not exist`); return; }
    if (qty < 1 || qty > 50) { errors.push(`Quantity for "${letter}" must be between 1 and 50`); return; }
    items[letter] = { ...prod, qty };
  });
  return { items, errors };
}

function genOrderNum() {
  const d = new Date();
  const pad = n => String(n).padStart(2, '0');
  return `ORD-${d.getFullYear()}${pad(d.getMonth()+1)}${pad(d.getDate())}-${String(Math.floor(1000 + Math.random() * 9000))}`;
}

function formatDate(d) {
  const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
  return `${d.getDate()} ${months[d.getMonth()]} ${d.getFullYear()}`;
}

function paymentInstructions(orderNum, total) {
  return `*Payment details:*\n\nBank: ${CONFIG.BANK_NAME}\nAccount Name: ${CONFIG.ACCOUNT_NAME}\nAccount Number: ${CONFIG.ACCOUNT_NUMBER}\nBranch Code: ${CONFIG.BRANCH_CODE}\nAmount: *R${total}*\nReference: *${orderNum}*\n\nEmail proof of payment to: bronnie@infin.co.za\n\nReply *PAID* once payment is done.`;
}

// ─── WHATSAPP SENDERS ────────────────────────────────────────
async function sendMessage(to, text) {
  try {
    await axios.post(
      `https://graph.facebook.com/v19.0/${CONFIG.PHONE_NUMBER_ID}/messages`,
      { messaging_product: 'whatsapp', to, type: 'text', text: { body: text, preview_url: false } },
      { headers: { Authorization: `Bearer ${CONFIG.WA_TOKEN}`, 'Content-Type': 'application/json' } }
    );
  } catch (err) {
    console.error('Send error:', err.response?.data || err.message);
  }
}

async function sendButtons(to, body, buttons) {
  try {
    await axios.post(
      `https://graph.facebook.com/v19.0/${CONFIG.PHONE_NUMBER_ID}/messages`,
      {
        messaging_product: 'whatsapp', to, type: 'interactive',
        interactive: {
          type: 'button', body: { text: body },
          action: { buttons: buttons.map((b, i) => ({ type: 'reply', reply: { id: `btn_${i}`, title: b } })) },
        },
      },
      { headers: { Authorization: `Bearer ${CONFIG.WA_TOKEN}`, 'Content-Type': 'application/json' } }
    );
  } catch (err) {
    console.error('Button error:', err.response?.data || err.message);
    await sendMessage(to, body);
  }
}

async function sendList(to, body, buttonText, sections) {
  try {
    await axios.post(
      `https://graph.facebook.com/v19.0/${CONFIG.PHONE_NUMBER_ID}/messages`,
      {
        messaging_product: 'whatsapp', to, type: 'interactive',
        interactive: {
          type: 'list',
          body: { text: body },
          action: {
            button: buttonText,
            sections,
          },
        },
      },
      { headers: { Authorization: `Bearer ${CONFIG.WA_TOKEN}`, 'Content-Type': 'application/json' } }
    );
  } catch (err) {
    console.error('List error:', err.response?.data || err.message);
    await sendMessage(to, body);
  }
}

async function buildMenuSections(cart = {}) {
  const menu = await getMenu();
  const stock = await getStockOnHand();
  const rows = menu.map(m => {
    const onHand = stock[m.name];
    const isOutOfStock = onHand !== undefined && onHand < 1;
    let title = m.name.length > 24 ? m.name.substring(0, 24) : m.name;
    if (isOutOfStock) {
      title = `(OOS) ${m.name}`.substring(0, 24);
    }
    return {
      id: `item_${m.letter}`,
      title,
      description: isOutOfStock ? 'Out of stock' : `${m.unit} — R${m.price}`,
    };
  });
  const sections = [{ title: 'Our Products', rows }];
  if (Object.keys(cart).length > 0) {
    sections.push({
      title: '─────────────',
      rows: [{
        id: 'finalise_order',
        title: '✅ Finalise order',
        description: 'Review and confirm your current order',
      }]
    });
  }
  return sections;
}

async function sendMenuList(to, cart = {}, body = null) {
  const sections = await buildMenuSections(cart);
  await sendList(
    to,
    body || `Welcome to ${CONFIG.BUSINESS_NAME}! Fresh meat delivered to your door in Pietermaritzburg. 👇 Browse our menu below:`,
    '🛒 View Menu & Order',
    sections
  );
}

async function sendQuantityButtons(to, itemName) {
  await sendList(
    to,
    `How many *${itemName}* would you like?`,
    'Select quantity',
    [{
      title: 'Quantity',
      rows: [
        { id: 'qty_1', title: '1' },
        { id: 'qty_2', title: '2' },
        { id: 'qty_3', title: '3' },
        { id: 'qty_4', title: '4' },
        { id: 'qty_5', title: '5' },
      ],
    }]
  );
}

async function sendMoreQuantityButtons(to, itemName) {
  await sendButtons(to, `How many *${itemName}* would you like?`, ['4', '5', '10']);
}
async function sendInvoiceAndSaveToDrive(phone, orderNum, cart, customerName) {
  const invoiceItems = Object.values(cart).map(item => ({
    name: item.name, unit: item.unit, price: item.price, qty: item.qty
  }));
  const orderData = {
    orderNum,
    date: formatDate(new Date()),
    customerName: customerName || 'Valued Customer',
    phone: `+${phone}`,
    items: invoiceItems,
  };
  const pdfBuffer = await generateInvoicePDF(orderData);
  if (pdfBuffer) {
    await uploadAndSendInvoice(phone, pdfBuffer, orderNum);
    await saveInvoiceToDrive(pdfBuffer, orderNum);
    const invoiceUrl = await uploadInvoiceToCloudinary(pdfBuffer, orderNum);
    if (invoiceUrl) {
      await sheetsUpdate('Orders', 'Order_num', orderNum, { Invoice_URL: invoiceUrl });
    }
  }
}

async function handleMessage(phone, text, customerName, interactiveId = '') {
  // Check if bot is switched off in Control sheet
  const status = await getBotStatus();
  if (status === 'Off') {
    await sendMessage(phone, `Hi there! 👋 Thanks for messaging *${CONFIG.BUSINESS_NAME}*!\n\nWe are currently away making more delicious meat for you! 🥩🔪\n\nWe will be back shortly — please message us again soon and we will get your order sorted. We appreciate your patience! 😊`);
    return;
  }

  const session = getSession(phone);
  const msg   = text.trim();
  const upper = msg.toUpperCase();

  if (['HI', 'HELLO', 'START', 'RESTART', 'MENU'].includes(upper)) {
    session.state = 'welcome';
    session.cart  = {};
  }

  // Handle Contact us and Delivery info from anywhere in the flow
  if (interactiveId === 'btn_contact' || upper === 'CONTACT US') {
    await sendMessage(phone, `You can reach us at:\n\n📞 *+27 82 617 9993 (Ross)*\n\n_or_\n\n📧 *bronnie@infin.co.za*\n\nWe are available Mon-Fri, 8:30am-5pm.`);
    return;
  }
  if (interactiveId === 'btn_delivery' || upper === 'DELIVERY INFO') {
    await sendMessage(phone, `We deliver Mon-Fri, 8:30am-5pm in Pietermaritzburg.\nWe will reach out to you for further details on your delivery.\n\n🚚 Delivery available within 5km of our workshop only.`);
    return;
  }

  if (session.state === 'welcome') {
    session.state = 'ordering';
    await updateSessionInSheets(phone, session);
    await sendMenuList(phone, session.cart);
    await sendButtons(phone, 'More info:', ['Delivery info', 'Contact us']);
    return;
  }

  if (session.state === 'ordering') {
    if (upper === 'CONTACT US') {
      await sendMessage(phone, `You can reach us at:\n\n📞 *+27 82 617 9993 (Ross)*\n\n_or_\n\n📧 *bronnie@infin.co.za*\n\nWe are available Mon-Fri, 8:30am-5pm.`);
      await sendMenuList(phone, session.cart);
      return;
    }
    if (upper === 'DELIVERY INFO') {
      await sendMessage(phone, 'We deliver Mon-Fri, 8:30am-5pm in Pietermaritzburg.\nWe will reach out to you for further details on your delivery.');
      await sendMenuList(phone, session.cart);
      return;
    }
    if (['VIEW MENU', 'PLACE AN ORDER', 'VIEW MENU & ORDER', 'MENU'].includes(upper)) {
      await sendMenuList(phone, session.cart);
      return;
    }

    // Customer tapped Finalise order from the list
    if (interactiveId === 'finalise_order' || ['CONFIRM ORDER', 'CONFIRM', 'FINALISE ORDER', 'FINALISE', '✅ FINALISE ORDER'].includes(upper)) {
      if (Object.keys(session.cart).length === 0) {
        await sendMessage(phone, 'Your cart is empty. Please select at least one item first.');
        await sendMenuList(phone, session.cart);
        return;
      }
      const { text: summary } = cartSummary(session.cart);
      session.state = 'confirm';
      await updateSessionInSheets(phone, session);
      await sendButtons(phone, `Here is your order:\n\n${summary}\n\nShall I confirm this?`, ['Yes, confirm', 'Add more items', 'Cancel']);
      return;
    }

    // Cancel order
    if (['CANCEL ORDER', 'CANCEL'].includes(upper)) {
      session.cart = {};
      session.pendingItem = null;
      await updateSessionInSheets(phone, session);
      await sendMessage(phone, 'Order cancelled. No problem! Let us know when you are ready to order again.');
      await sendMenuList(phone, session.cart);
      return;
    }

    // Customer selected a product from the list
    const menu = await getMenu();
    const selectedByList = menu.find(m => interactiveId === `item_${m.letter}`);
    if (selectedByList) {
      const stock = await getStockOnHand();
      const onHand = stock[selectedByList.name];
      if (onHand !== undefined && onHand < 1) {
        await sendMessage(phone, `Sorry, *${selectedByList.name}* is currently out of stock. Please choose another item.`);
        await sendMenuList(phone, session.cart);
        return;
      }
      session.pendingItem = selectedByList;
      await updateSessionInSheets(phone, session);
      await sendQuantityButtons(phone, selectedByList.name);
      return;
    }

    // Handle response to stock-cap offer
    if (session.pendingStockCap) {
      const cap = session.pendingStockCap;
      if (['YES', 'YES, ORDER ' + cap, 'ORDER ' + cap].includes(upper)) {
        const item = session.pendingItem;
        session.cart[item.letter] = { ...item, qty: cap };
        session.pendingItem = null;
        session.pendingStockCap = null;
        await updateSessionInSheets(phone, session);
        const { text: cartSummaryText } = cartSummary(session.cart);
        await sendList(
          phone,
          `✅ *${item.name} x${cap}* added to your order.\n\n${cartSummaryText}\n\nTap *View Menu & Order* to add more items or *Finalise Order* when done 👇`,
          '🛒 View Menu & Order',
          await buildMenuSections(session.cart)
        );
        await sendButtons(phone, 'Ready to checkout?', ['✅ Finalise Order', 'Cancel']);
        return;
      }
      if (['NO', 'CANCEL'].includes(upper)) {
        session.pendingItem = null;
        session.pendingStockCap = null;
        await updateSessionInSheets(phone, session);
        await sendMessage(phone, 'No problem. Choose another item below.');
        await sendMenuList(phone, session.cart);
        return;
      }
    }

    // Customer selected quantity via buttons (1, 2, 3, 4, 5)
    const qty = parseInt(msg);
    if (!isNaN(qty) && qty > 0 && qty <= 5 && session.pendingItem) {
      const item = session.pendingItem;
      // Check stock
      const stock = await getStockOnHand();
      const onHand = stock[item.name];
      if (onHand !== undefined && qty > onHand) {
        if (onHand < 1) {
          session.pendingItem = null;
          await updateSessionInSheets(phone, session);
          await sendMessage(phone, `Sorry, *${item.name}* is out of stock.`);
          await sendMenuList(phone, session.cart);
          return;
        }
        // Offer to cap qty at available stock
        session.pendingStockCap = onHand;
        await updateSessionInSheets(phone, session);
        await sendButtons(
          phone,
          `Sorry, only *${onHand}* x ${item.name} available. Would you like to order ${onHand} instead?`,
          [`Yes, order ${onHand}`, 'No, cancel']
        );
        return;
      }
      session.cart[item.letter] = { ...item, qty };
      session.pendingItem = null;
      await updateSessionInSheets(phone, session);
      const { text: cartSummaryText } = cartSummary(session.cart);
      await sendList(
        phone,
        `✅ *${item.name} x${qty}* added to your order.\n\n${cartSummaryText}\n\nTap *View Menu & Order* to add more items or *Finalise Order* when done 👇`,
        '🛒 View Menu & Order',
        await buildMenuSections(session.cart)
      );
      await sendButtons(phone, 'Ready to checkout?', ['✅ Finalise Order', 'Cancel']);
      return;
    }

    // Fallback
    await sendMenuList(phone, session.cart);
    return;
  }

  if (session.state === 'confirm') {
    if (['YES, CONFIRM', 'YES', 'CONFIRM', 'CONFIRM ORDER'].includes(upper)) {
      const orderNum = genOrderNum();
      session.orderNum = orderNum;
      const { text: summary, total } = cartSummary(session.cart);
      session.state = 'delivery';
      await updateSessionInSheets(phone, session);
      await sendMessage(phone, `✅ Order *${orderNum}* confirmed!\n\n${summary}`);
      await sendButtons(phone, 'How would you like to receive your order?\n\n🚚 *Delivery available within 5km of our workshop only.*', [
        '🏭 Workshop',
        '🚚 Delivery',
      ]);
      return;
    }
    if (['ADD MORE ITEMS', 'ADD', 'ADD ANOTHER ITEM', 'ADD ANOTHER'].includes(upper)) {
      session.state = 'ordering';
      await updateSessionInSheets(phone, session);
      await sendMenuList(phone, session.cart);
      return;
    }
    if (['CANCEL', 'CANCEL ORDER'].includes(upper)) {
      session.cart  = {};
      session.state = 'ordering';
      session.pendingItem = null;
      await updateSessionInSheets(phone, session);
      await sendMessage(phone, 'Order cancelled. No problem! Let us know when you are ready to order again.');
      await sendMenuList(phone);
      return;
    }
    const { text: summary } = cartSummary(session.cart);
    await sendButtons(phone, `${summary}\n\nWhat would you like to do?`, ['Confirm order', 'Add another item', 'Cancel order']);
    return;
  }

  if (session.state === 'delivery') {
    const { text: summary, total } = cartSummary(session.cart);
    const orderNum = session.orderNum;

    if (['🏭 WORKSHOP', 'WORKSHOP COLLECTION', 'WORKSHOP'].includes(upper) || interactiveId === 'btn_0') {
      session.deliveryType = 'Collection - Workshop';
      session.deliveryAddress = '';
      session.state = 'payment';
      await updateSessionInSheets(phone, session);
      await logOrderToSheets(phone, orderNum, session.cart, total, customerName, session.deliveryType, session.deliveryAddress);
      await updateStockSheet();
      await sendMessage(phone, `Great! Please collect your order from:\n\n📍 *14 Loftus Street, Murrayfield Park, Mkondeni*\n\n${paymentInstructions(orderNum, total)}`);
      await sendInvoiceAndSaveToDrive(phone, orderNum, session.cart, customerName);
      return;
    }

    if (['🚚 DELIVERY', 'DELIVERY'].includes(upper) || interactiveId === 'btn_1') {
      session.deliveryType = 'Delivery';
      session.state = 'awaiting_address';
      await updateSessionInSheets(phone, session);
      await sendMessage(phone, `Please reply with your full delivery address 📍`);
      return;
    }

    // Fallback - re-show options
    await sendButtons(phone, 'How would you like to receive your order?\n\n🚚 *Delivery available within 5km of our workshop only.*', [
      '🏭 Workshop',
      '🚚 Delivery',
    ]);
    return;
  }

  if (session.state === 'awaiting_address') {
    const { text: summary, total } = cartSummary(session.cart);
    const orderNum = session.orderNum;
    session.deliveryAddress = text;
    session.state = 'payment';
    await updateSessionInSheets(phone, session);
    await logOrderToSheets(phone, orderNum, session.cart, total, customerName, session.deliveryType, session.deliveryAddress);
    await updateStockSheet();
    await sendMessage(phone, `Got it! We will deliver to:\n\n📍 *${text}*\n\n${paymentInstructions(orderNum, total)}`);
    await sendInvoiceAndSaveToDrive(phone, orderNum, session.cart, customerName);
    return;
  }

  if (session.state === 'payment') {
    if (['PAID', "I'VE PAID", 'IVE PAID'].includes(upper)) {
      session.state = 'done';
      await updateSessionInSheets(phone, session);
      await sheetsUpdate('Orders', 'Order_num', session.orderNum, { Paid: 'YES' });
      await sendMessage(phone, `Thank you! We will verify your payment and confirm your delivery slot shortly.\n\nOrder: *${session.orderNum}*\n\nIf you have any questions call us on +27 82 617 9993 (Ross).`);
      return;
    }
    if (upper === 'HELP') {
      await sendMessage(phone, `No problem! Call us on +27 82 617 9993 (Ross) or reply here.\n\nOrder: *${session.orderNum}*`);
      return;
    }
    const { total } = cartSummary(session.cart);
    await sendButtons(phone, `Your order *${session.orderNum}* is waiting for payment.\n\n${paymentInstructions(session.orderNum, total)}`, ["I've paid", 'Help']);
    return;
  }

  if (session.state === 'done') {
    session.state = 'ordering';
    session.cart  = {};
    await updateSessionInSheets(phone, session);
    await sendMessage(phone, `Thanks for ordering from ${CONFIG.BUSINESS_NAME}! Reply MENU to place another order anytime.`);
    return;
  }

  await sendMessage(phone, `Hi! Reply MENU to see our products and place an order.`);
}

// ─── ROUTES ──────────────────────────────────────────────────
app.get('/dashboard-data', async (req, res) => {
  try {
    res.setHeader('Access-Control-Allow-Origin', '*');
    const sheets = await getSheetsClient();

    const [ordersRes, lineItemsRes] = await Promise.all([
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Orders' }),
      sheets.spreadsheets.values.get({ spreadsheetId: SPREADSHEET_ID, range: 'Line items' }),
    ]);

    const parseSheet = (res) => {
      const rows = res.data.values || [];
      if (rows.length < 2) return [];
      const headers = rows[0];
      return rows.slice(1).map(row => {
        const obj = {};
        headers.forEach((h, i) => obj[h] = row[i] || '');
        return obj;
      });
    };

    res.json({
      orders: parseSheet(ordersRes),
      lineItems: parseSheet(lineItemsRes),
    });
  } catch (err) {
    console.error('Dashboard data error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

app.get('/register-number', async (req, res) => {
  try {
    const pin = req.query.pin || '111222';
    const result = await axios.post(
      `https://graph.facebook.com/v19.0/${CONFIG.PHONE_NUMBER_ID}/register`,
      { messaging_product: 'whatsapp', pin },
      { headers: { Authorization: `Bearer ${CONFIG.WA_TOKEN}`, 'Content-Type': 'application/json' } }
    );
    res.json({ success: true, pin, data: result.data });
  } catch (err) {
    console.error('Register error:', err.response?.data || err.message);
    res.status(500).json({ error: err.response?.data || err.message });
  }
});

app.get('/setup-sheets', async (req, res) => {
  try {
    const sheets = await getSheetsClient();

    // Get sheet IDs
    const meta = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const ordersSheet = meta.data.sheets.find(s => s.properties.title === 'Orders');
    if (!ordersSheet) return res.status(404).json({ error: 'Orders sheet not found' });
    const sheetId = ordersSheet.properties.sheetId;

    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: {
        requests: [
          // Paid column (G = index 6) - rows 2 to 1000
          {
            setDataValidation: {
              range: { sheetId, startRowIndex: 1, endRowIndex: 1000, startColumnIndex: 6, endColumnIndex: 7 },
              rule: {
                condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'YES' }, { userEnteredValue: 'NO' }] },
                showCustomUi: true,
                strict: false,
              },
            },
          },
          // POP Confirmed column (H = index 7) - rows 2 to 1000
          {
            setDataValidation: {
              range: { sheetId, startRowIndex: 1, endRowIndex: 1000, startColumnIndex: 7, endColumnIndex: 8 },
              rule: {
                condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'YES' }, { userEnteredValue: 'NO' }] },
                showCustomUi: true,
                strict: false,
              },
            },
          },
          // Delivered column (I = index 8) - rows 2 to 1000
          {
            setDataValidation: {
              range: { sheetId, startRowIndex: 1, endRowIndex: 1000, startColumnIndex: 8, endColumnIndex: 9 },
              rule: {
                condition: { type: 'ONE_OF_LIST', values: [{ userEnteredValue: 'YES' }, { userEnteredValue: 'NO' }] },
                showCustomUi: true,
                strict: false,
              },
            },
          },
        ],
      },
    });

    res.json({ success: true, message: 'Dropdowns set for Paid and Delivered columns' });
  } catch (err) {
    console.error('Setup error:', err.message);
    res.status(500).json({ error: err.message });
  }
});

app.get('/health', (req, res) => {
  res.status(200).json({ status: 'ok', bot: CONFIG.BUSINESS_NAME, uptime: Math.floor(process.uptime()) + 's', sessions: Object.keys(sessions).length, menuItems: menuCache.length });
});

app.get('/webhook', (req, res) => {
  const mode = req.query['hub.mode'], token = req.query['hub.verify_token'], challenge = req.query['hub.challenge'];
  if (mode === 'subscribe' && token === CONFIG.VERIFY_TOKEN) { console.log('Webhook verified'); res.status(200).send(challenge); }
  else res.sendStatus(403);
});

app.post('/webhook', async (req, res) => {
  res.sendStatus(200);
  try {
    const value   = req.body?.entry?.[0]?.changes?.[0]?.value;
    const message = value?.messages?.[0];
    if (!message) return;
    const phone        = message.from;
    const customerName = value?.contacts?.[0]?.profile?.name || '';
    let text = '';
    let interactiveId = '';

    if (message.type === 'text') {
      text = message.text?.body || '';
    } else if (message.type === 'interactive') {
      const interactive = message.interactive;
      if (interactive?.type === 'button_reply') {
        text = interactive.button_reply?.title || '';
        interactiveId = interactive.button_reply?.id || '';
      } else if (interactive?.type === 'list_reply') {
        text = interactive.list_reply?.title || '';
        interactiveId = interactive.list_reply?.id || '';
      }
    }

    if (!text && !interactiveId) return;
    console.log(`[${new Date().toISOString()}] [${phone}] text="${text}" id="${interactiveId}"`);
    await handleMessage(phone, text, customerName, interactiveId);
  } catch (err) {
    console.error('Error:', err);
  }
});

app.get('/', (req, res) => res.json({ bot: CONFIG.BUSINESS_NAME, status: 'running' }));

getMenu().then(m => console.log(`Startup: ${m.length} menu items loaded`));

// Background stock refresh every 2 minutes
setInterval(async () => {
  try {
    stockCacheTime = 0; // invalidate cache
    await getStockOnHand();
    console.log('Stock cache refreshed in background');
  } catch (err) {
    console.error('Background stock refresh error:', err.message);
  }
}, 2 * 60 * 1000);

app.listen(CONFIG.PORT, () => console.log(`${CONFIG.BUSINESS_NAME} bot running on port ${CONFIG.PORT}`));
