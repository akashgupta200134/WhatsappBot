'use strict';

const { chromium } = require('playwright');
const XLSX         = require('xlsx');
const nodemailer   = require('nodemailer');
const fs           = require('fs');   
const path         = require('path');

// ═══════════════════════════════════════════════════════════════
//  CONFIGURATION
// ═══════════════════════════════════════════════════════════════

const CONFIG = {

  EXCEL_FILE     : 'C:\\Users\\Akash\\Desktop\\WhatsappBot\\Copy of contactsDemo-Mainsheet_Macro.xlsx',
  IMAGE_FILE     : 'C:\\Users\\Akash\\Desktop\\WhatsappBot\\INDIA BANNER.jpg',
  LOG_FILE       : 'C:\\Users\\Akash\\Desktop\\WhatsappBot\\log.txt',
  CHROME_PROFILE : 'C:\\Users\\Akash\\Desktop\\WhatsappBot\\chrome-profile',

  COL_PHONE    : 'PhoneNumber',
  COL_REMARK   : 'Remark',

  COUNTRY_CODE : '91',
  MIN_DIGITS   : 10,
  MAX_DIGITS   : 15,

  MAX_RETRIES  : 2,
  RETRY_FAILED : true,

  DELAY_BETWEEN_CONTACTS : 5000,
  DELAY_AFTER_IMAGE      : 4000,
  DELAY_AFTER_MESSAGE    : 3000,
  WHATSAPP_LOAD_TIMEOUT  : 25000,
  WA_READY_TIMEOUT       : 90000,

  MESSAGE:`Hello 👋

Greetings from Dextero Business Solutions Pvt. Ltd.
(Proud Member of BNI Amigos, Mumbai)

🍽️ Are You Looking to Grow Your Food & Beverage Business Online?

Join our *Complimentary Webinar* and discover how E-Commerce can help you increase sales, improve customer experience, and streamline operations.

📅 Date: 24 June 2026
🕚 Time: 11:00 AM – 12:00 PM IST

🚀 What You\'ll Learn:
✅ Launch your online store quickly
✅ Manage products, inventory & orders efficiently
✅ Increase direct online sales
✅ Improve customer retention & engagement
✅ Scale your business without complexity
✅ Leverage analytics for better decision-making

🎯 Ideal For:
• Restaurants
• Cafés
• Cloud Kitchens
• Bakeries
• Food Manufacturers
• Distributors & Retailers
• Food & Beverage Entrepreneurs

🎁 Participation is absolutely FREE!

👉 Register Now:
https://dextero.in/webinar/fnb-101

For any queries:
Bhavika Joshi
Business Consultant
Dextero Business Solutions Pvt. Ltd.
📞 +91 99872 97200
🌐 www.dextero.in | www.actifyzone.com
Projects | Products | Staffing | Training

Limited seats available. Register today!`,

  EMAIL: {
    FROM     : 'akash.gupta@dextero.in',
    PASSWORD : 'Akash@30#',
    TO       : 'Akash.gupta@dextero.in',
    SUBJECT  : 'WhatsApp Bot Run Report',
  },
};

// ═══════════════════════════════════════════════════════════════
//  STATUS CONSTANTS
// ═══════════════════════════════════════════════════════════════

const STATUS = {
  DONE        : 'Done',
  FAILED      : 'Failed',
  INVALID     : 'InvalidNumber',
  DUPLICATE   : 'Duplicate',
  NO_WHATSAPP : 'NotOnWhatsApp',
};

// ═══════════════════════════════════════════════════════════════
//  LOGGER
// ═══════════════════════════════════════════════════════════════

const Logger = {
  init() {
    const dir = path.dirname(CONFIG.LOG_FILE);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    fs.writeFileSync(CONFIG.LOG_FILE, '');
  },
  _write(level, msg) {
    const time = new Date().toLocaleTimeString('en-IN', { hour12: false });
    const date = new Date().toLocaleDateString('en-IN');
    const line = `[${date} ${time}] [${level}]  ${msg}`;
    console.log(line);
    fs.appendFileSync(CONFIG.LOG_FILE, line + '\n');
  },
  info   : (msg) => Logger._write('INFO ', msg),
  warn   : (msg) => Logger._write('WARN ', msg),
  error  : (msg) => Logger._write('ERROR', msg),
  success: (msg) => Logger._write('  OK ', msg),
  divider: ()    => Logger._write('─────', '─'.repeat(50)),
  header : (msg) => {
    Logger._write('═════', '═'.repeat(50));
    Logger._write('═════', `  ${msg}`);
    Logger._write('═════', '═'.repeat(50));
  },
};

// ═══════════════════════════════════════════════════════════════
//  VALIDATORS
// ═══════════════════════════════════════════════════════════════

const Validator = {

  excelFileExists() {
    if (!fs.existsSync(CONFIG.EXCEL_FILE)) {
      Logger.error(`Excel file not found at: ${CONFIG.EXCEL_FILE}`);
      return false;
    }
    return true;
  },

  imageFileExists() {
    if (!fs.existsSync(CONFIG.IMAGE_FILE)) {
      Logger.error(`Image file not found at: ${CONFIG.IMAGE_FILE}`);
      return false;
    }
    return true;
  },

  emailConfigValid() {
    if (!CONFIG.EMAIL.FROM || !CONFIG.EMAIL.PASSWORD || !CONFIG.EMAIL.TO) {
      Logger.error('Email config incomplete.');
      return false;
    }
    return true;
  },

  phoneNumber(raw) {
    if (!raw && raw !== 0) return { valid: false, reason: 'empty' };
    let num = String(raw).replace(/\D/g, '').trim();
    if (num.length === 0) return { valid: false, reason: 'empty after cleaning' };
    if (num.startsWith('9191') && num.length === 14) num = num.slice(2);
    if (num.length === CONFIG.MIN_DIGITS) num = CONFIG.COUNTRY_CODE + num;
    if (num.length < CONFIG.MIN_DIGITS || num.length > CONFIG.MAX_DIGITS) {
      return { valid: false, reason: `Invalid length: ${num.length} digits` };
    }
    return { valid: true, number: num };
  },

  findDuplicates(rows) {
    const seen  = new Map();
    const dupes = new Set();
    for (let i = 0; i < rows.length; i++) {
      const raw = String(rows[i][CONFIG.COL_PHONE] || '').trim();
      if (!raw) continue;
      const result = Validator.phoneNumber(raw);
      if (!result.valid) continue;
      const num = result.number;
      if (seen.has(num)) {
        dupes.add(num);
        Logger.warn(`Duplicate: ${num} at rows ${seen.get(num) + 2} and ${i + 2}`);
      } else {
        seen.set(num, i);
      }
    }
    return dupes;
  },
};

// ═══════════════════════════════════════════════════════════════
//  EXCEL MANAGER
// ═══════════════════════════════════════════════════════════════

const ExcelManager = {
  workbook        : null,
  sheetName       : null,
  sheet           : null,
  phoneColLetter  : null,
  remarkColLetter : null,

  load() {
    ExcelManager.workbook  = XLSX.readFile(CONFIG.EXCEL_FILE);
    ExcelManager.sheetName = ExcelManager.workbook.SheetNames[0];
    ExcelManager.sheet     = ExcelManager.workbook.Sheets[ExcelManager.sheetName];
    Logger.info(`Reading sheet: "${ExcelManager.sheetName}"`);

    for (const cellAddr in ExcelManager.sheet) {
      if (cellAddr[0] === '!') continue;
      const match = cellAddr.match(/^([A-Z]+)1$/);
      if (!match) continue;
      const colLetter = match[1];
      const cellValue = String(ExcelManager.sheet[cellAddr].v || '').trim();
      if (cellValue === CONFIG.COL_PHONE)  ExcelManager.phoneColLetter  = colLetter;
      if (cellValue === CONFIG.COL_REMARK) ExcelManager.remarkColLetter = colLetter;
    }

    if (!ExcelManager.phoneColLetter)
      throw new Error(`Column "${CONFIG.COL_PHONE}" not found in row 1.`);
    if (!ExcelManager.remarkColLetter)
      throw new Error(`Column "${CONFIG.COL_REMARK}" not found in row 1.`);

    Logger.info(`Columns → Phone: ${ExcelManager.phoneColLetter} | Remark: ${ExcelManager.remarkColLetter}`);
  },

  getRows() {
    ExcelManager.workbook  = XLSX.readFile(CONFIG.EXCEL_FILE);
    ExcelManager.sheetName = ExcelManager.workbook.SheetNames[0];
    ExcelManager.sheet     = ExcelManager.workbook.Sheets[ExcelManager.sheetName];
    return XLSX.utils.sheet_to_json(ExcelManager.sheet, { defval: '' });
  },

  writeRemark(excelRow, value) {
    const wb    = XLSX.readFile(CONFIG.EXCEL_FILE);
    const sName = wb.SheetNames[0];
    const sh    = wb.Sheets[sName];
    sh[`${ExcelManager.remarkColLetter}${excelRow}`] = { v: value, t: 's' };
    XLSX.writeFile(wb, CONFIG.EXCEL_FILE);
  },
};

// ═══════════════════════════════════════════════════════════════
//  EMAIL SENDER
// ═══════════════════════════════════════════════════════════════

const EmailSender = {
  transporter: null,

 init() {
  EmailSender.transporter = nodemailer.createTransport({
    host   : 'smtp.office365.com',
    port   : 587,
    secure : false,
    auth   : { user: CONFIG.EMAIL.FROM, pass: CONFIG.EMAIL.PASSWORD },
    tls    : { 
      ciphers          : 'SSLv3', 
      rejectUnauthorized: false,
      minVersion       : 'TLSv1.2',
    },
    family            : 4,        // force IPv4
    connectionTimeout : 10000,
    greetingTimeout   : 10000,
    socketTimeout     : 15000,
  });
},

  buildHtml(summary, contacts) {
    const now  = new Date().toLocaleString('en-IN');
    const rate = summary.total > 0 ? ((summary.sent / summary.total) * 100).toFixed(1) : '0.0';
    const contactRows = contacts.map((c, i) => {
      const color =
        c.status === STATUS.DONE        ? '#e6f9ee' :
        c.status === STATUS.FAILED      ? '#fdecea' :
        c.status === STATUS.NO_WHATSAPP ? '#fff3e0' :
        c.status === STATUS.INVALID     ? '#fce4ec' :
        c.status === STATUS.DUPLICATE   ? '#f3e5f5' : '#f5f5f5';
      const icon =
        c.status === STATUS.DONE        ? '✅' :
        c.status === STATUS.FAILED      ? '❌' :
        c.status === STATUS.NO_WHATSAPP ? '⚠️' :
        c.status === STATUS.INVALID     ? '🚫' :
        c.status === STATUS.DUPLICATE   ? '🔁' : '⏭️';
      return `<tr style="background:${color}">
        <td style="padding:8px 12px;border-bottom:1px solid #eee">${i + 1}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #eee">${c.phone}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #eee">${icon} ${c.status}</td>
        <td style="padding:8px 12px;border-bottom:1px solid #eee">${c.note || '—'}</td>
      </tr>`;
    }).join('');

    return `<!DOCTYPE html><html><head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;margin:0;padding:20px;background:#f9f9f9">
<div style="max-width:800px;margin:auto;background:#fff;border-radius:8px;overflow:hidden;box-shadow:0 2px 8px rgba(0,0,0,.1)">
  <div style="background:#075E54;padding:24px 32px">
    <h1 style="color:#fff;margin:0;font-size:22px">WhatsApp Bot — Run Report</h1>
    <p style="color:#B2DFDB;margin:6px 0 0">${now}</p>
  </div>
  <div style="display:flex;flex-wrap:wrap;gap:12px;padding:24px 32px;background:#f4f4f4">
    ${[
      ['#075E54', summary.total,      'Total'],
      ['#4CAF50', summary.sent,       'Sent ✅'],
      ['#f44336', summary.failed,     'Failed ❌'],
      ['#FF9800', summary.noWhatsapp, 'Not on WA ⚠️'],
      ['#9C27B0', summary.skipped,    'Skipped ⏭️'],
      ['#607D8B', summary.invalid,    'Invalid 🚫'],
      ['#607D8B', summary.duplicates, 'Duplicates 🔁'],
    ].map(([color, val, label]) => `
      <div style="flex:1;min-width:110px;background:#fff;border-radius:8px;padding:16px;text-align:center;border-top:4px solid ${color}">
        <div style="font-size:28px;font-weight:bold;color:${color}">${val}</div>
        <div style="color:#666;font-size:13px;margin-top:4px">${label}</div>
      </div>`).join('')}
  </div>
  <div style="padding:0 32px 16px;background:#f4f4f4">
    <div style="background:#fff;border-radius:8px;padding:16px 20px;display:flex;gap:32px;flex-wrap:wrap">
      <div><div style="color:#999;font-size:12px">SUCCESS RATE</div><div style="font-size:18px;font-weight:bold;color:#075E54">${rate}%</div></div>
      <div><div style="color:#999;font-size:12px">DURATION</div><div style="font-size:18px;font-weight:bold;color:#333">${summary.durationMinutes} min</div></div>
      <div><div style="color:#999;font-size:12px">STOP REASON</div><div style="font-size:18px;font-weight:bold;color:#333">${summary.stopReason}</div></div>
    </div>
  </div>
  <div style="padding:24px 32px">
    <h2 style="font-size:16px;color:#333;margin:0 0 12px">Contact Details</h2>
    <table style="width:100%;border-collapse:collapse;font-size:13px">
      <thead><tr style="background:#075E54;color:#fff">
        <th style="padding:10px 12px;text-align:left">#</th>
        <th style="padding:10px 12px;text-align:left">Phone</th>
        <th style="padding:10px 12px;text-align:left">Status</th>
        <th style="padding:10px 12px;text-align:left">Note</th>
      </tr></thead>
      <tbody>${contactRows}</tbody>
    </table>
  </div>
  <div style="padding:16px 32px;background:#f9f9f9;border-top:1px solid #eee;text-align:center">
    <p style="color:#999;font-size:12px;margin:0">Generated by WhatsApp Bot | ${now}</p>
  </div>
</div></body></html>`;
  },

  async send(summary, contacts) {
  try {
    Logger.info('Sending email report...');
    await EmailSender.transporter.verify();   // test connection first
    await EmailSender.transporter.sendMail({
      from    : CONFIG.EMAIL.FROM,
      to      : CONFIG.EMAIL.TO,
      subject : `${CONFIG.EMAIL.SUBJECT} — ${summary.sent}/${summary.total} Sent | ${new Date().toLocaleDateString('en-IN')}`,
      html    : EmailSender.buildHtml(summary, contacts),
    });
    Logger.success('Email sent successfully.');
  } catch (err) {
    Logger.error(`Email failed: ${err.message}`);
    // Log details to help debug
    Logger.warn(`Email config: host=smtp.office365.com port=587 user=${CONFIG.EMAIL.FROM}`);
  }
},
};

// ═══════════════════════════════════════════════════════════════
//  WHATSAPP CORE
// ═══════════════════════════════════════════════════════════════

const WA = {
  page    : null,
  context : null,

  async openChat(phoneNumber) {
    try {
      await WA.page.goto(
        `https://web.whatsapp.com/send?phone=${phoneNumber}`,
        { waitUntil: 'load', timeout: 30000 }
      );
    } catch (e) {
      Logger.warn(`  goto timeout (non-fatal): ${e.message}`);
    }

    const deadline = Date.now() + CONFIG.WHATSAPP_LOAD_TIMEOUT;
    while (Date.now() < deadline) {
      const composeBox = await WA.page.$('[data-testid="conversation-compose-box-input"]').catch(() => null);
      if (composeBox) return 'ready';

      const popupText = await WA.page.evaluate(() => {
        const popup = document.querySelector('[data-testid="popup-contents"]');
        return popup ? popup.innerText.toLowerCase() : '';
      }).catch(() => '');

      if (
        popupText.includes('invalid') ||
        popupText.includes('not on whatsapp') ||
        popupText.includes('phone number shared')
      ) {
        Logger.warn(`  Invalid-number popup for ${phoneNumber}: "${popupText.slice(0, 60)}"`);
        await WA.page.keyboard.press('Escape').catch(() => {});
        return 'not_on_whatsapp';
      }
      await WA.page.waitForTimeout(800);
    }
    return 'timeout';
  },

  async sendImage() {
    const { execSync } = require('child_process');

    Logger.info('  → Copying image to clipboard via PowerShell...');
    const psCmd = `Add-Type -AssemblyName System.Windows.Forms; ` +
      `[System.Windows.Forms.Clipboard]::SetImage(` +
      `[System.Drawing.Image]::FromFile('${CONFIG.IMAGE_FILE.replace(/\\/g, '\\\\')}'))`;
    execSync(`powershell -command "${psCmd}"`, { timeout: 5000 });
    Logger.info('  → Clipboard set.');

    Logger.info('  → Focusing compose box and pasting...');
    await WA.page.click('[data-testid="conversation-compose-box-input"]', { timeout: 8000 });
    await WA.page.waitForTimeout(500);
    await WA.page.keyboard.press('Control+V');
    Logger.info('  → Paste triggered, waiting for preview...');
    await WA.page.waitForTimeout(3000);

    Logger.info('  → Sending image...');
    // Try clicking send button, fall back to Enter
    let sendClicked = false;
    for (const sel of [
      'button[aria-label="Send 1 selected"]',
      'button[aria-label="Send"]',
      '[data-testid="send"]',
    ]) {
      try {
        const btn = await WA.page.$(sel);
        if (btn) {
          await btn.click();
          sendClicked = true;
          Logger.info(`  → Image sent via: ${sel}`);
          break;
        }
      } catch { /* try next */ }
    }
    if (!sendClicked) {
      Logger.warn('  → Send button not found, pressing Enter...');
      await WA.page.keyboard.press('Enter');
    }

    await WA.page.waitForTimeout(CONFIG.DELAY_AFTER_IMAGE);
    Logger.info('  → Image sent.');
  },

  async sendMessage() {
    Logger.info('  → Waiting for compose box...');
    await WA.page.waitForSelector(
      '[data-testid="conversation-compose-box-input"]',
      { timeout: 10000 }
    );

    Logger.info('  → Pasting message...');
    await WA.page.evaluate((msg) => {
      const box = document.querySelector('[data-testid="conversation-compose-box-input"]');
      if (!box) throw new Error('Compose box not found');
      box.focus();
      const dt = new DataTransfer();
      dt.setData('text/plain', msg);
      box.dispatchEvent(new ClipboardEvent('paste', {
        bubbles      : true,
        cancelable   : true,
        clipboardData: dt,
      }));
    }, CONFIG.MESSAGE);

    await WA.page.waitForTimeout(1500);

    Logger.info('  → Sending message...');
    let sendClicked = false;
    for (const sel of [
      'button[aria-label="Send"]',
      '[data-testid="send"]',
      'span[data-icon="send"]',
    ]) {
      try {
        const btn = await WA.page.$(sel);
        if (btn) {
          await btn.click();
          sendClicked = true;
          break;
        }
      } catch { /* try next */ }
    }
    if (!sendClicked) {
      Logger.warn('  → Send button not found, pressing Enter...');
      await WA.page.keyboard.press('Enter');
    }

    await WA.page.waitForTimeout(CONFIG.DELAY_AFTER_MESSAGE);
    Logger.info('  → Message sent.');
  },

  async sendToContact(phoneNumber) {
    let lastError = '';
    for (let attempt = 1; attempt <= CONFIG.MAX_RETRIES + 1; attempt++) {
      try {
        if (attempt > 1) {
          Logger.warn(`  Retry ${attempt - 1}/${CONFIG.MAX_RETRIES} for ${phoneNumber}`);
          await WA.page.waitForTimeout(3000);
        }

        const chatStatus = await WA.openChat(phoneNumber);

        if (chatStatus === 'not_on_whatsapp') {
          return { success: false, status: STATUS.NO_WHATSAPP, note: 'Number not on WhatsApp' };
        }
        if (chatStatus === 'timeout') {
          lastError = 'Chat did not load within timeout';
          Logger.warn(`  Attempt ${attempt}: ${lastError}`);
          continue;
        }

        await WA.sendImage();
        await WA.sendMessage();

        return { success: true, status: STATUS.DONE, note: '' };

      } catch (err) {
        lastError = err.message;
        Logger.warn(`  Attempt ${attempt} failed: ${err.message}`);
        try {
          await WA.page.goto('https://web.whatsapp.com', { waitUntil: 'load', timeout: 15000 });
          await WA.page.waitForTimeout(2000);
        } catch { /* ignore */ }
      }
    }

    return {
      success: false,
      status : STATUS.FAILED,
      note   : `Failed after ${CONFIG.MAX_RETRIES + 1} attempts: ${lastError}`,
    };
  },
}; 



// ═══════════════════════════════════════════════════════════════
//  GRACEFUL SHUTDOWN
// ═══════════════════════════════════════════════════════════════

let isShuttingDown  = false;
let summarySnapshot = null;
let contactSnapshot = [];

async function shutdown(reason) {
  if (isShuttingDown) return;
  isShuttingDown = true;
  Logger.warn(`Shutdown triggered: ${reason}`);
  if (summarySnapshot) {
    summarySnapshot.stopReason      = reason;
    summarySnapshot.durationMinutes = ((Date.now() - summarySnapshot.startTime) / 60000).toFixed(1);
    await EmailSender.send(summarySnapshot, contactSnapshot);
  }
  if (WA.context) { try { await WA.context.close(); } catch {} }
  Logger.info('Bot shut down cleanly.');
  process.exit(0);
}

process.on('SIGINT',            () => shutdown('Manual stop (Ctrl+C)'));
process.on('SIGTERM',           () => shutdown('Process terminated'));
process.on('uncaughtException', async (err) => {
  Logger.error(`Uncaught exception: ${err.message}`);
  await shutdown('Uncaught exception');
});

// ═══════════════════════════════════════════════════════════════
//  MAIN
// ═══════════════════════════════════════════════════════════════

async function main() {
  Logger.init();
  Logger.header('WhatsApp Bot — Production Run');
  const startTime = Date.now();

  // ── PRE-FLIGHT ──────────────────────────────────────────────
  Logger.info('Running pre-flight checks...');
  if (!Validator.excelFileExists())  process.exit(1);
  if (!Validator.imageFileExists())  process.exit(1);
  if (!Validator.emailConfigValid()) process.exit(1);
  Logger.success('All pre-flight checks passed.');
  Logger.divider();

  // ── LOAD EXCEL ──────────────────────────────────────────────
  Logger.info('Loading Excel file...');
  try { ExcelManager.load(); } catch (err) {
    Logger.error(`Failed to load Excel: ${err.message}`); process.exit(1);
  }

  let rows;
  try { rows = ExcelManager.getRows(); } catch (err) {
    Logger.error(`Failed to read rows: ${err.message}`); process.exit(1);
  }

  if (!rows || rows.length === 0) {
    Logger.error('STOPPED: No data found in Excel sheet.'); process.exit(1);
  }
  Logger.info(`Total rows in Excel: ${rows.length}`);

  // ── DUPLICATE CHECK ─────────────────────────────────────────
  Logger.info('Checking for duplicates...');
  const duplicateNumbers = Validator.findDuplicates(rows);
  if (duplicateNumbers.size > 0) {
    Logger.warn(`Found ${duplicateNumbers.size} duplicate(s).`);
  } else {
    Logger.success('No duplicates found.');
  }

  // ── FILTER WHAT NEEDS PROCESSING ────────────────────────────
  const needsProcessing = rows.filter((r) => {
    const phone  = String(r[CONFIG.COL_PHONE]  || '').trim();
    const remark = String(r[CONFIG.COL_REMARK] || '').trim();
    if (!phone)                                           return false;
    if (remark === STATUS.DONE)                           return false;
    if (remark === STATUS.INVALID)                        return false;
    if (remark === STATUS.NO_WHATSAPP)                    return false;
    if (remark === STATUS.DUPLICATE)                      return false;
    if (remark === STATUS.FAILED && !CONFIG.RETRY_FAILED) return false;
    return true;
  });

  if (needsProcessing.length === 0) {
    Logger.warn('STOPPED: All contacts already processed.');
    const contacts = rows.map((r) => ({
      phone  : String(r[CONFIG.COL_PHONE]  || '').trim(),
      status : String(r[CONFIG.COL_REMARK] || 'Unknown').trim(),
      note   : 'Pre-existing status',
    }));
    const summary = {
      total: rows.length,
      sent           : contacts.filter(c => c.status === STATUS.DONE).length,
      failed         : contacts.filter(c => c.status === STATUS.FAILED).length,
      noWhatsapp     : contacts.filter(c => c.status === STATUS.NO_WHATSAPP).length,
      skipped        : 0,
      invalid        : contacts.filter(c => c.status === STATUS.INVALID).length,
      duplicates     : contacts.filter(c => c.status === STATUS.DUPLICATE).length,
      stopReason     : 'All contacts already processed',
      durationMinutes: '0.0',
      startTime,
    };
    EmailSender.init();
    await EmailSender.send(summary, contacts);
    process.exit(0);
  }

  Logger.info(`Contacts needing processing: ${needsProcessing.length}`);
  Logger.divider();
  EmailSender.init();


  // ── LAUNCH BROWSER ──────────────────────────────────────────
  Logger.info('Launching Chrome...');
  if (!fs.existsSync(CONFIG.CHROME_PROFILE)) {
    fs.mkdirSync(CONFIG.CHROME_PROFILE, { recursive: true });
    Logger.warn('First run — scan QR code. You have 90 seconds.');
  } else {
    Logger.info('Existing Chrome profile found — no QR scan needed.');
  }

 // In your launchPersistentContext call, update args:
WA.context = await chromium.launchPersistentContext(
  CONFIG.CHROME_PROFILE,
  {
    channel  : 'chrome',
    headless : false,
    args     : [
      '--start-maximized',
      '--unsafely-treat-insecure-origin-as-secure=https://web.whatsapp.com',
      '--clipboard-write',                     // allow clipboard write
      '--allow-clipboard-read',               // allow clipboard read
    ],
    viewport         : null,
    permissions      : ['clipboard-read', 'clipboard-write'],  // grant permissions
  }
);
  WA.page = await WA.context.newPage();

  // ── LOAD WHATSAPP WEB ───────────────────────────────────────
  Logger.info('Opening WhatsApp Web...');
  try {
    await WA.page.goto('https://web.whatsapp.com', { waitUntil: 'load', timeout: 30000 });
  } catch (e) {
    Logger.warn(`goto timeout (non-fatal): ${e.message}`);
  }

  // Give the page a moment to settle, then dismiss any blocking dialog
  // (WhatsApp sometimes shows a 'Close' / 'Got it' overlay on first paint)
  await WA.page.waitForTimeout(3000);
  for (const label of ['Close', 'Got it', 'OK', 'Continue']) {
    try {
      const btn =
        await WA.page.$(`button[aria-label="${label}"]`).catch(() => null) ||
        await WA.page.$(`button:has-text("${label}")`).catch(() => null);
      if (btn) {
        await btn.click().catch(() => {});
        Logger.info(`  Dismissed "${label}" dialog.`);
        await WA.page.waitForTimeout(1000);
        break;
      }
    } catch { /* ignore */ }
  }

  Logger.info('Waiting for WhatsApp to be ready (up to 90s)...');
  let waReady = false;
  const waDeadline = Date.now() + CONFIG.WA_READY_TIMEOUT;
  while (Date.now() < waDeadline) {
    const found =
      await WA.page.$('[data-testid="chat-list-search"]').catch(() => null) ||   // search bar (older WA)
      await WA.page.$('[data-testid="chat-list"]').catch(() => null) ||           // chat list pane
      await WA.page.$('#pane-side').catch(() => null) ||                          // left sidebar
      await WA.page.$('[data-testid="cell-frame-container"]').catch(() => null) || // any chat row
      await WA.page.$('[aria-label="Search input textbox"]').catch(() => null) || // search box aria
      await WA.page.$('div[contenteditable="true"][data-tab="3"]').catch(() => null); // old compose

    if (found) { waReady = true; break; }
    await WA.page.waitForTimeout(1000);
  }

  if (!waReady) {
    Logger.error('WhatsApp Web did not become ready. Scan QR or check internet.');
    await WA.context.close(); process.exit(1);
  }
  Logger.success('WhatsApp Web is ready.');
  Logger.divider();

  // ── TRACKING ────────────────────────────────────────────────
  const summary = {
    total: rows.length, sent: 0, failed: 0, noWhatsapp: 0,
    skipped: 0, invalid: 0, duplicates: 0,
    stopReason: 'All contacts processed successfully',
    durationMinutes: '0', startTime,
  };
  const contactLog  = [];
  const seenNumbers = new Set();
  summarySnapshot   = summary;
  contactSnapshot   = contactLog;

  // ── MAIN LOOP ───────────────────────────────────────────────
  for (let i = 0; i < rows.length; i++) {
    if (isShuttingDown) break;

    const rawPhone = String(rows[i][CONFIG.COL_PHONE]  || '').trim();
    const remark   = String(rows[i][CONFIG.COL_REMARK] || '').trim();
    const excelRow = i + 2;

    Logger.info(`─── Row ${excelRow} / ${rows.length + 1} ───`);

    if (!rawPhone) {
      Logger.warn(`Row ${excelRow}: Empty phone. Stopping.`);
      summary.stopReason = `Empty phone at row ${excelRow}`; break;
    }

    const phoneResult = Validator.phoneNumber(rawPhone);
    if (!phoneResult.valid) {
      Logger.warn(`${rawPhone} → INVALID: ${phoneResult.reason}`);
      ExcelManager.writeRemark(excelRow, STATUS.INVALID);
      contactLog.push({ phone: rawPhone, status: STATUS.INVALID, note: phoneResult.reason });
      summary.invalid++; continue;
    }

    const phoneNumber = phoneResult.number;

    if (seenNumbers.has(phoneNumber)) {
      Logger.warn(`${phoneNumber} → DUPLICATE`);
      ExcelManager.writeRemark(excelRow, STATUS.DUPLICATE);
      contactLog.push({ phone: phoneNumber, status: STATUS.DUPLICATE, note: 'Same number appeared earlier' });
      summary.duplicates++; continue;
    }
    seenNumbers.add(phoneNumber);

    if (remark === STATUS.DONE) {
      Logger.info(`${phoneNumber} → Already Done. Skipping.`);
      contactLog.push({ phone: phoneNumber, status: STATUS.DONE, note: 'Already sent' });
      summary.skipped++; continue;
    }

    if ([STATUS.INVALID, STATUS.DUPLICATE, STATUS.NO_WHATSAPP].includes(remark)) {
      Logger.info(`${phoneNumber} → Previously "${remark}". Skipping.`);
      contactLog.push({ phone: phoneNumber, status: remark, note: 'Previous run result' });
      summary.skipped++; continue;
    }

    if (remark === STATUS.FAILED && !CONFIG.RETRY_FAILED) {
      Logger.info(`${phoneNumber} → Previously Failed. RETRY_FAILED=false. Skipping.`);
      contactLog.push({ phone: phoneNumber, status: STATUS.FAILED, note: 'RETRY_FAILED disabled' });
      summary.skipped++; continue;
    }

    Logger.info(`${phoneNumber} → Sending...`);
    const result = await WA.sendToContact(phoneNumber);

    ExcelManager.writeRemark(excelRow, result.status);
    contactLog.push({ phone: phoneNumber, status: result.status, note: result.note });

    if (result.status === STATUS.DONE) {
      Logger.success(`${phoneNumber} → ✓ Sent`);
      summary.sent++;
    } else if (result.status === STATUS.NO_WHATSAPP) {
      Logger.warn(`${phoneNumber} → ⚠ Not on WhatsApp`);
      summary.noWhatsapp++;
    } else {
      Logger.error(`${phoneNumber} → ✗ Failed: ${result.note}`);
      summary.failed++;
    }

    await WA.page.waitForTimeout(CONFIG.DELAY_BETWEEN_CONTACTS);
  }

  // ── FINAL ───────────────────────────────────────────────────
  Logger.divider();
  summary.durationMinutes = ((Date.now() - startTime) / 60000).toFixed(1);
  Logger.header('Run Complete');
  Logger.info   (`Total           : ${summary.total}`);
  Logger.success(`Sent            : ${summary.sent}`);
  Logger.error  (`Failed          : ${summary.failed}`);
  Logger.warn   (`Not on WhatsApp : ${summary.noWhatsapp}`);
  Logger.info   (`Skipped         : ${summary.skipped}`);
  Logger.info   (`Invalid         : ${summary.invalid}`);
  Logger.info   (`Duplicates      : ${summary.duplicates}`);
  Logger.info   (`Duration        : ${summary.durationMinutes} min`);
  Logger.info   (`Stop Reason     : ${summary.stopReason}`);
  Logger.divider();

  await EmailSender.send(summary, contactLog);
  await WA.context.close();
  Logger.header('Bot Finished');
}

main();



