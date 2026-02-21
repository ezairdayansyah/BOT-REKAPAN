require('dotenv').config();
const TelegramBot = require('node-telegram-bot-api');
const { google } = require('googleapis');
const fs = require('fs');
const path = require('path');
const PDFDocument = require('pdfkit');

// ===== PARSE FUNCTIONS =====

// Parse aktivasi dengan Object.entries pattern
function parseAktivasi(text, userRow, username) {
  let data = {
    tanggal: new Date().toLocaleDateString('id-ID', {
      weekday: 'long',
      day: 'numeric',
      month: 'long',
      year: 'numeric',
      timeZone: 'Asia/Jakarta',
    }),
    channel: '',
    workorder: '',
    ao: '',
    scOrderNo: '',
    serviceNo: '',
    customerName: '',
    workzone: '',
    contactPhone: '',
    odp: '',
    symptom: '',
    memo: '',
    tikor: '',
    snOnt: '',
    nikOnt: '',
    stbId: '',
    nikStb: '',
    teknisi: (userRow && userRow[8]) ? (userRow[8] || username).replace('@', '') : (username || ''),
  };

  // Regex patterns menggunakan Object.entries
  const patterns = {
    channel: /CHANNEL\s*:\s*(.+?)(?=\n|$)/i,
    workorder: /WORKORDER\s*:\s*(.+?)(?=\n|$)/i,
    ao: /AO\s*:\s*(.+?)(?=\n|$)/i,
    scOrderNo: /SC\s*ORDER\s*NO\s*:\s*(.+?)(?=\n|$)/i,
    serviceNo: /SERVICE\s*NO\s*:\s*(.+?)(?=\n|$)/i,
    customerName: /CUSTOMER\s*NAME\s*:\s*(.+?)(?=\n|$)/i,
    workzone: /WORKZONE\s*:\s*(.+?)(?=\n|$)/i,
    contactPhone: /CONTACT\s*PHONE\s*:\s*(.+?)(?=\n|$)/i,
    odp: /ODP\s*:\s*(.+?)(?=\n|$)/i,
    symptom: /SYMPTOM\s*:\s*(.+?)(?=\n|$)/i,
    memo: /MEMO\s*:\s*(.+?)(?=\n|$)/i,
    tikor: /TIKOR\s*:\s*(.+?)(?=\n|$)/i,
    snOnt: /SN\s*ONT\s*:\s*(.+?)(?=\n|$)/i,
    nikOnt: /NIK\s*ONT\s*:\s*(.+?)(?=\n|$)/i,
    stbId: /STB\s*ID\s*:\s*(.+?)(?=\n|$)/i,
    nikStb: /NIK\s*STB\s*:\s*(.+?)(?=\n|$)/i,
  };

  // Extract semua pattern
  for (const [key, pattern] of Object.entries(patterns)) {
    const match = text.match(pattern);
    if (match && match[1]) {
      data[key] = match[1].trim();
    }
  }

  // Fallback untuk SN ONT brand-specific patterns
  if (!data.snOnt) {
    const snPatterns = [
      /(ZTEG[A-Z0-9]+)/i,
      /(HWTC[A-Z0-9]+)/i,
      /(HUAW[A-Z0-9]+)/i,
      /(FHTT[A-Z0-9]+)/i,
      /(FHTTC[A-Z0-9]+)/i,
      /(FIBR[A-Z0-9]+)/i
    ];
    for (const pattern of snPatterns) {
      const match = text.match(pattern);
      if (match && match[1]) {
        data.snOnt = match[1].trim();
        break;
      }
    }
  }

  // SC ORDER NO fallback
  if (!data.scOrderNo) {
    const scMatch = text.match(/\b(SC\d{6,})\b/i);
    if (scMatch && scMatch[1]) {
      data.scOrderNo = scMatch[1].trim();
    }
  }

  // AO fallback
  if (!data.ao && data.scOrderNo) {
    data.ao = data.scOrderNo;
  }

  // WORKORDER fallback
  if (!data.workorder && data.ao) {
    data.workorder = data.ao;
  }

  return data;
}

// ===== ENVIRONMENT & CONFIG =====

const TOKEN = process.env.TELEGRAM_TOKEN;
const SHEET_ID = process.env.SHEET_ID;
const GOOGLE_SERVICE_ACCOUNT_KEY = process.env.GOOGLE_SERVICE_ACCOUNT_KEY;

if (!TOKEN || !SHEET_ID || !GOOGLE_SERVICE_ACCOUNT_KEY) {
  console.error('ERROR: Missing required environment variables!');
  process.exit(1);
}

const REKAPAN_SHEET = 'REKAPAN QUALITY';
const MASTER_SHEET = 'MASTER';

// ===== GOOGLE SHEETS SETUP =====

let serviceAccount;
try {
  let keyData = GOOGLE_SERVICE_ACCOUNT_KEY;
  if (!keyData.startsWith('{')) {
    try {
      keyData = Buffer.from(keyData, 'base64').toString('utf-8');
    } catch (e) {
      console.log('Not base64 encoded, using as is');
    }
  }
  serviceAccount = JSON.parse(keyData);
  console.log('✓ Google Service Account parsed successfully');
} catch (e) {
  console.error('ERROR parsing Google credentials:', e.message);
  process.exit(1);
}

const auth = new google.auth.GoogleAuth({
  credentials: serviceAccount,
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});
const sheets = google.sheets({ version: 'v4', auth });

// ===== TELEGRAM BOT SETUP =====

let bot;
const PORT = process.env.PORT || 3000;
const RAILWAY_STATIC_URL = process.env.RAILWAY_STATIC_URL;
const USE_WEBHOOK = process.env.USE_WEBHOOK === 'true' || !!RAILWAY_STATIC_URL;

if (USE_WEBHOOK && RAILWAY_STATIC_URL) {
  const express = require('express');
  const app = express();
  app.use(express.json());
  
  bot = new TelegramBot(TOKEN);
  const webhookUrl = `https://${RAILWAY_STATIC_URL}/bot${TOKEN}`;
  
  bot.setWebHook(webhookUrl)
    .then(() => console.log(`✓ Webhook set to: ${webhookUrl}`))
    .catch(err => console.error('Failed to set webhook:', err));
  
  app.post(`/bot${TOKEN}`, (req, res) => {
    bot.processUpdate(req.body);
    res.sendStatus(200);
  });
  
  app.get('/', (req, res) => res.send('Bot is running!'));
  
  app.listen(PORT, () => console.log(`✓ Server listening on port ${PORT}`));
} else {
  bot = new TelegramBot(TOKEN, { polling: true });
  console.log('✓ Bot running in polling mode');
}

// ===== HELPER FUNCTIONS =====

async function getSheetData(sheetName) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: sheetName,
    });
    return res.data.values || [];
  } catch (error) {
    console.error(`Error getting ${sheetName}:`, error.message);
    throw error;
  }
}

async function appendSheetData(sheetName, values) {
  try {
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: sheetName,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [values] },
    });
  } catch (error) {
    console.error(`Error appending to ${sheetName}:`, error.message);
    throw error;
  }
}

async function updateSheetData(sheetName, range, values) {
  try {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `${sheetName}!${range}`,
      valueInputOption: 'USER_ENTERED',
      resource: { values },
    });
  } catch (error) {
    console.error(`Error updating ${sheetName}:`, error.message);
    throw error;
  }
}

async function sendTelegram(chatId, text, options = {}) {
  const maxLength = 4000;
  const maxRetries = 3;
  
  async function sendWithRetry(message, retries = 0) {
    try {
      return await bot.sendMessage(chatId, message, { parse_mode: 'HTML', ...options });
    } catch (error) {
      if (retries < maxRetries) {
        console.log(`Retry ${retries + 1} for chat ${chatId}`);
        await new Promise(resolve => setTimeout(resolve, 1000 * (retries + 1)));
        return sendWithRetry(message, retries + 1);
      }
      throw error;
    }
  }
  
  if (text.length <= maxLength) {
    return sendWithRetry(text);
  } else {
    const lines = text.split('\n');
    let chunk = '';
    let promises = [];
    for (const line of lines) {
      if ((chunk + line + '\n').length > maxLength) {
        promises.push(sendWithRetry(chunk));
        chunk = '';
      }
      chunk += line + '\n';
    }
    if (chunk.trim()) promises.push(sendWithRetry(chunk));
    return Promise.all(promises);
  }
}

async function sendPDFFile(chatId, dataRows, headers, filename, options = {}) {
  try {
    const tempDir = '/tmp';
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }
    
    const filePath = path.join(tempDir, filename);
    const doc = new PDFDocument({ 
      margin: 10,
      size: 'A4',
      layout: 'landscape'
    });
    
    const stream = fs.createWriteStream(filePath);
    doc.pipe(stream);
    
    // Title
    doc.fontSize(16).font('Helvetica-Bold').text('DATA AKTIVASI', { align: 'center' });
    doc.fontSize(9).font('Helvetica').text(`Generated: ${new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' })} WIB`, { align: 'center' });
    doc.moveDown();
    
    // Table header
    const pageWidth = doc.page.width - 20;
    const colWidth = pageWidth / headers.length;
    
    doc.fontSize(8).font('Helvetica-Bold').fillColor('#000');
    doc.rect(10, doc.y, pageWidth, 18).fill('#CCCCCC');
    doc.fillColor('black');
    
    const headerY = doc.y + 3;
    headers.forEach((header, i) => {
      doc.text(header, 10 + (i * colWidth) + 2, headerY, {
        width: colWidth - 4,
        align: 'left',
        fontSize: 7
      });
    });
    
    doc.moveDown(1.3);
    
    // Table rows
    doc.fontSize(7).font('Helvetica');
    dataRows.forEach((row) => {
      if (doc.y > doc.page.height - 30) {
        doc.addPage();
        doc.fontSize(8).font('Helvetica-Bold').fillColor('#000');
        doc.rect(10, doc.y, pageWidth, 18).fill('#CCCCCC');
        doc.fillColor('black');
        const newHeaderY = doc.y + 3;
        headers.forEach((header, i) => {
          doc.text(header, 10 + (i * colWidth) + 2, newHeaderY, {
            width: colWidth - 4,
            align: 'left',
            fontSize: 7
          });
        });
        doc.moveDown(1.3);
        doc.fontSize(7).font('Helvetica');
      }
      
      const rowY = doc.y;
      headers.forEach((_, i) => {
        const cellData = (row[i] || '').toString().substring(0, 50);
        doc.text(cellData, 10 + (i * colWidth) + 2, rowY, {
          width: colWidth - 4,
          align: 'left'
        });
      });
      
      doc.moveDown();
    });
    
    // Footer
    doc.fontSize(8).text(`Total Records: ${dataRows.length}`, { align: 'right' });
    
    doc.end();
    
    await new Promise((resolve, reject) => {
      stream.on('finish', resolve);
      stream.on('error', reject);
    });
    
    await bot.sendDocument(chatId, filePath, {
      caption: `📄 File PDF berhasil digenerate!\nFilename: ${filename}`,
      ...options
    });
    
    fs.unlinkSync(filePath);
  } catch (error) {
    console.error('Error sending PDF:', error);
    throw error;
  }
}

// Cek user dari MASTER sheet (H=ID, I=USERNAME, J=ROLE, K=STATUS)
async function getUserData(username) {
  try {
    const data = await getSheetData(MASTER_SHEET);
    console.log(`[CHECK] Username: @${username}`);
    console.log(`[CHECK] Total users: ${data.length - 1}`);
    
    if (!data || data.length <= 1) {
      console.log('[CHECK] No users in MASTER sheet');
      return null;
    }
    
    for (let i = 1; i < data.length; i++) {
      const sheetUsername = (data[i][8] || '').replace('@', '').toLowerCase().trim();  // I = index 8
      const role = (data[i][9] || '').toUpperCase().trim();                            // J = index 9
      const status = (data[i][10] || '').toUpperCase().trim();                         // K = index 10
      const inputUsername = (username || '').replace('@', '').toLowerCase().trim();
      
      console.log(`[CHECK] Row ${i}: username="${sheetUsername}", role="${role}", status="${status}"`);
      
      if (sheetUsername === inputUsername && status === 'AKTIF') {
        console.log(`[CHECK] ✓ User found and AKTIF!`);
        return data[i];
      }
    }
    
    console.log(`[CHECK] ✗ User not found or not AKTIF`);
    return null;
  } catch (error) {
    console.error('Error checking user:', error);
    return null;
  }
}

async function isAdmin(username) {
  const user = await getUserData(username);
  return user && (user[9] || '').toUpperCase() === 'ADMIN';  // J = index 9 = ROLE
}

function getTodayDateString() {
  const today = new Date();
  return today.toLocaleDateString('id-ID', {
    weekday: 'long',
    day: 'numeric',
    month: 'long',
    year: 'numeric',
    timeZone: 'Asia/Jakarta'
  });
}

function parseIndonesianDate(dateStr) {
  const months = {
    'januari': '01', 'februari': '02', 'maret': '03', 'april': '04',
    'mei': '05', 'juni': '06', 'juli': '07', 'agustus': '08',
    'september': '09', 'oktober': '10', 'november': '11', 'desember': '12'
  };
  
  const parts = dateStr.toLowerCase().split(' ');
  if (parts.length >= 4) {
    const day = parts[1].padStart(2, '0');
    const month = months[parts[2]];
    const year = parts[3];
    if (month) {
      return new Date(`${year}-${month}-${day}`);
    }
  }
  return null;
}

function filterDataByPeriod(data, period, customDate = null) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  let startDate, endDate;
  
  if (customDate) {
    const datePattern = /(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/;
    const match = customDate.match(datePattern);
    if (match) {
      const day = parseInt(match[1]);
      const month = parseInt(match[2]) - 1;
      const year = parseInt(match[3]);
      const targetDate = new Date(year, month, day);
      
      if (period === 'daily') {
        startDate = new Date(targetDate);
        endDate = new Date(targetDate);
        endDate.setHours(23, 59, 59, 999);
      } else if (period === 'weekly') {
        const dayOfWeek = targetDate.getDay();
        const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
        startDate = new Date(targetDate);
        startDate.setDate(targetDate.getDate() + mondayOffset);
        startDate.setHours(0, 0, 0, 0);
        endDate = new Date(startDate);
        endDate.setDate(startDate.getDate() + 6);
        endDate.setHours(23, 59, 59, 999);
      } else if (period === 'monthly') {
        startDate = new Date(targetDate.getFullYear(), targetDate.getMonth(), 1);
        endDate = new Date(targetDate.getFullYear(), targetDate.getMonth() + 1, 0);
        endDate.setHours(23, 59, 59, 999);
      }
    }
  } else {
    switch (period) {
      case 'daily':
        startDate = new Date(today);
        endDate = new Date(today);
        endDate.setHours(23, 59, 59, 999);
        break;
      case 'weekly':
        const dayOfWeek = today.getDay();
        const mondayOffset = dayOfWeek === 0 ? -6 : 1 - dayOfWeek;
        startDate = new Date(today);
        startDate.setDate(today.getDate() + mondayOffset);
        endDate = new Date(startDate);
        endDate.setDate(startDate.getDate() + 6);
        endDate.setHours(23, 59, 59, 999);
        break;
      case 'monthly':
        startDate = new Date(today.getFullYear(), today.getMonth(), 1);
        endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
        endDate.setHours(23, 59, 59, 999);
        break;
      default:
        return data.slice(1);
    }
  }
  
  const filtered = [];
  for (let i = 1; i < data.length; i++) {
    const dateStr = data[i][0];
    if (dateStr) {
      const rowDate = parseIndonesianDate(dateStr);
      if (rowDate && rowDate >= startDate && rowDate <= endDate) {
        filtered.push(data[i]);
      }
    }
  }
  
  return filtered;
}

function generateCSV(data, headers) {
  let csv = headers.join(',') + '\n';
  
  data.forEach(row => {
    const csvRow = row.map(cell => {
      const cellStr = (cell || '').toString();
      if (cellStr.includes(',') || cellStr.includes('"') || cellStr.includes('\n')) {
        return '"' + cellStr.replace(/"/g, '""') + '"';
      }
      return cellStr;
    });
    csv += csvRow.join(',') + '\n';
  });
  
  return csv;
}

// ===== MESSAGE HANDLER =====

bot.on('message', async (msg) => {
  const chatId = msg.chat.id;
  const messageId = msg.message_id;
  const text = (msg.text || '').trim();
  const username = msg.from.username || '';
  const chatType = msg.chat.type;
  
  console.log(`\n[MSG] Chat: ${chatId}, User: @${username}, Type: ${chatType}`);
  console.log(`[MSG] Text: ${text.substring(0, 100)}`);
  
  try {
    // Ignore group messages yang bukan /aktivasi
    if ((chatType === 'group' || chatType === 'supergroup') && !/^\/aktivasi\b/i.test(text)) {
      return;
    }
    
    // === /aktivasi: Input data aktivasi ===
    if (/^\/aktivasi\b/i.test(text)) {
      console.log(`[AKTIVASI] User: @${username}`);
      
      if (!username) {
        return sendTelegram(chatId, '❌ Anda harus memiliki username Telegram.\nSilakan atur username di pengaturan Telegram Anda.', { reply_to_message_id: messageId });
      }
      
      const user = await getUserData(username);
      if (!user) {
        console.log(`[AKTIVASI] User not found for @${username}`);
        return sendTelegram(chatId, `❌ @${username} tidak terdaftar di MASTER sheet.\nSilakan hubungi admin.`, { reply_to_message_id: messageId });
      }
      
      const inputText = text.replace(/^\/aktivasi\s*/i, '').trim();
      if (!inputText) {
        return sendTelegram(chatId, 'Silakan kirim data aktivasi setelah /aktivasi', { reply_to_message_id: messageId });
      }
      
      const parsed = parseAktivasi(inputText, user, username);
      
      if (!parsed.ao) {
        return sendTelegram(chatId, '❌ Field AO wajib diisi.', { reply_to_message_id: messageId });
      }
      
      // Cek duplikat berdasarkan AO
      const data = await getSheetData(REKAPAN_SHEET);
      let isDuplicate = false;
      for (let i = 1; i < data.length; i++) {
        if ((data[i][3] || '').toUpperCase().trim() === parsed.ao.toUpperCase().trim()) {
          isDuplicate = true;
          break;
        }
      }
      
      if (isDuplicate) {
        return sendTelegram(chatId, '❌ Data duplikat. AO sudah diinput sebelumnya.', { reply_to_message_id: messageId });
      }
      
      // Susun row sesuai struktur 18 kolom
      const row = [
        parsed.tanggal,        // A: TANGGAL
        parsed.channel,        // B: CHANNEL
        parsed.workorder,      // C: WORKORDER
        parsed.ao,             // D: AO
        parsed.scOrderNo,      // E: SC_ORDER_NO
        parsed.serviceNo,      // F: SERVICE_NO
        parsed.customerName,   // G: CUSTOMER_NAME
        parsed.workzone,       // H: WORKZONE
        parsed.contactPhone,   // I: CONTACT_PHONE
        parsed.odp,            // J: ODP
        parsed.symptom,        // K: SYMPTOM
        parsed.memo,           // L: MEMO
        parsed.tikor,          // M: TIKOR
        parsed.snOnt,          // N: SN_ONT
        parsed.nikOnt,         // O: NIK_ONT
        parsed.stbId,          // P: STB_ID
        parsed.nikStb,         // Q: NIK_STB
        parsed.teknisi         // R: TEKNISI
      ];
      
      await appendSheetData(REKAPAN_SHEET, row);
      
      let confirmMsg = '✅ Data berhasil disimpan!\n';
      confirmMsg += '<b>Lanjut GROUP FULFILLMENT dan PT1</b> 🚀';
      
      return sendTelegram(chatId, confirmMsg, { reply_to_message_id: messageId });
    }
    
    // === /cari: Statistik aktivasi user ===
    else if (/^\/cari\b/i.test(text)) {
      const user = await getUserData(username);
      if (!user) {
        return sendTelegram(chatId, '❌ Anda tidak terdaftar sebagai user aktif.', { reply_to_message_id: messageId });
      }
      
      const data = await getSheetData(REKAPAN_SHEET);
      const userTeknisi = (user[8] || username).replace('@', '').toLowerCase();
      let count = 0;
      let channelMap = {}, workzoneMap = {};
      
      for (let i = 1; i < data.length; i++) {
        const teknisiData = (data[i][17] || '').replace('@', '').toLowerCase();
        if (teknisiData === userTeknisi) {
          count++;
          const channel = (data[i][1] || '-').toUpperCase();
          const workzone = (data[i][7] || '-').toUpperCase();
          channelMap[channel] = (channelMap[channel] || 0) + 1;
          workzoneMap[workzone] = (workzoneMap[workzone] || 0) + 1;
        }
      }
      
      let msg = `📊 <b>STATISTIK AKTIVASI</b>\n👤 Teknisi: ${user[8] || username}\n📈 Total: ${count} SSL\n\n`;
      
      if (count === 0) {
        msg += '⚠️ Belum ada data aktivasi.\n';
      } else {
        msg += '<b>Per Channel:</b>\n';
        Object.entries(channelMap).sort((a, b) => b[1] - a[1]).forEach(([c, cnt]) => {
          msg += `• ${c}: ${cnt}\n`;
        });
        msg += '\n<b>Per Workzone:</b>\n';
        Object.entries(workzoneMap).sort((a, b) => b[1] - a[1]).forEach(([w, cnt]) => {
          msg += `• ${w}: ${cnt}\n`;
        });
        msg += '\n💾 <i>Gunakan /exportcari untuk download data lengkap</i>';
      }
      
      msg += `\n📅 ${new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' })} WIB`;
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /exportcari: Export aktivasi user ke PDF ===
    else if (/^\/exportcari\b/i.test(text)) {
      const user = await getUserData(username);
      if (!user) {
        return sendTelegram(chatId, '❌ Anda tidak terdaftar sebagai user aktif.', { reply_to_message_id: messageId });
      }
      
      const data = await getSheetData(REKAPAN_SHEET);
      const userTeknisi = (user[8] || username).replace('@', '').toLowerCase();
      const userActivations = [];
      
      // Headers sesuai struktur 18 kolom
      const headers = ['TANGGAL', 'CHANNEL', 'WORKORDER', 'AO', 'SC_ORDER_NO', 'SERVICE_NO', 'CUSTOMER_NAME', 'WORKZONE', 'CONTACT_PHONE', 'ODP', 'SYMPTOM', 'MEMO', 'TIKOR', 'SN_ONT', 'NIK_ONT', 'STB_ID', 'NIK_STB', 'TEKNISI'];
      
      // Filter data untuk user ini
      for (let i = 1; i < data.length; i++) {
        const teknisiData = (data[i][17] || '').replace('@', '').toLowerCase();
        if (teknisiData === userTeknisi) {
          userActivations.push(data[i]);
        }
      }
      
      if (userActivations.length === 0) {
        return sendTelegram(chatId, '❌ Tidak ada data aktivasi untuk diekspor.', { reply_to_message_id: messageId });
      }
      
      // Generate PDF
      const filename = `aktivasi_${userTeknisi}_${new Date().toISOString().split('T')[0]}.pdf`;
      
      await sendPDFFile(chatId, userActivations, headers, filename, { reply_to_message_id: messageId });
    }
    
    // === /ps: Laporan harian (admin only) ===
    else if (/^\/ps\b/i.test(text)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Hanya admin yang bisa menggunakan command ini.', { reply_to_message_id: messageId });
      }
      
      const args = text.split(' ').slice(1);
      const customDate = args.length > 0 ? args[0] : null;
      
      const data = await getSheetData(REKAPAN_SHEET);
      const filteredData = customDate ? 
        filterDataByPeriod(data, 'daily', customDate) : 
        filterDataByPeriod(data, 'daily');
      
      let total = filteredData.length;
      let teknisiMap = {}, workzoneMap = {}, channelMap = {};
      
      filteredData.forEach(row => {
        const teknisi = (row[17] || '-').toUpperCase();
        const workzone = (row[7] || '-').toUpperCase();
        const channel = (row[1] || '-').toUpperCase();
        teknisiMap[teknisi] = (teknisiMap[teknisi] || 0) + 1;
        workzoneMap[workzone] = (workzoneMap[workzone] || 0) + 1;
        channelMap[channel] = (channelMap[channel] || 0) + 1;
      });
      
      const dateLabel = customDate || getTodayDateString();
      let msg = `📊 <b>LAPORAN HARIAN</b>\nTanggal: ${dateLabel}\nTotal: ${total} SSL\n\n`;
      
      if (total === 0) {
        msg += '⚠️ Tidak ada data.\n';
      } else {
        msg += `Teknisi Aktif: ${Object.keys(teknisiMap).length}\n`;
        msg += `Workzone: ${Object.keys(workzoneMap).length}\n`;
        msg += `Channel: ${Object.keys(channelMap).length}\n\n`;
        
        msg += '<b>TOP TEKNISI:</b>\n';
        Object.entries(teknisiMap).sort((a, b) => b[1] - a[1]).slice(0, 10).forEach(([t, c], i) => {
          msg += `${i + 1}. ${t}: ${c} SSL\n`;
        });
        
        msg += '\n<b>PERFORMA WORKZONE:</b>\n';
        Object.entries(workzoneMap).sort((a, b) => b[1] - a[1]).forEach(([w, c], i) => {
          msg += `${i + 1}. ${w}: ${c} SSL\n`;
        });
        
        msg += '\n<b>PERFORMA OWNER:</b>\n';
        Object.entries(channelMap).sort((a, b) => b[1] - a[1]).forEach(([ch, c], i) => {
          msg += `${i + 1}. ${ch}: ${c} SSL\n`;
        });
      }
      
      // Format timestamp dd/m/yyyy hh.mm.ss
      const formatter = new Intl.DateTimeFormat('id-ID', {
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit',
        timeZone: 'Asia/Jakarta',
        hour12: false
      });
      
      const parts = formatter.formatToParts(new Date());
      const partsMap = {};
      parts.forEach(part => {
        if (part.type !== 'literal') {
          partsMap[part.type] = part.value;
        }
      });
      
      const dateStr = partsMap.day + '/' + partsMap.month + '/' + partsMap.year;
      const timeStr = partsMap.hour + '.' + partsMap.minute + '.' + partsMap.second;
      
      msg += `\n⏰ ${dateStr} ${timeStr} WIB`;
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /topteknisi: Ranking teknisi (admin only) ===
    else if (/^\/topteknisi\b/i.test(text)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Hanya admin yang bisa menggunakan command ini.', { reply_to_message_id: messageId });
      }
      
      const args = text.split(' ').slice(1);
      const period = args[0] || 'all';
      const customDate = args[1] || null;
      
      const data = await getSheetData(REKAPAN_SHEET);
      let filteredData;
      
      switch (period.toLowerCase()) {
        case 'daily':
          filteredData = filterDataByPeriod(data, 'daily', customDate);
          break;
        case 'weekly':
          filteredData = filterDataByPeriod(data, 'weekly', customDate);
          break;
        case 'monthly':
          filteredData = filterDataByPeriod(data, 'monthly', customDate);
          break;
        default:
          filteredData = data.slice(1);
      }
      
      let teknisiMap = {};
      filteredData.forEach(row => {
        const teknisi = (row[17] || '-').toUpperCase();
        if (teknisi !== '-') {
          teknisiMap[teknisi] = (teknisiMap[teknisi] || 0) + 1;
        }
      });
      
      const sortedTeknisi = Object.entries(teknisiMap).sort((a, b) => b[1] - a[1]);
      const periodLabel = {
        daily: customDate ? `Harian (${customDate})` : 'Hari ini',
        weekly: customDate ? `Mingguan (${customDate})` : 'Minggu ini',
        monthly: customDate ? `Bulanan (${customDate})` : 'Bulan ini',
        all: 'Keseluruhan'
      };
      
      let msg = `🏆 <b>RANKING TEKNISI</b>\nPeriode: ${periodLabel[period.toLowerCase()] || 'Keseluruhan'}\n\n`;
      
      if (sortedTeknisi.length === 0) {
        msg += '⚠️ Belum ada data.\n';
      } else {
        msg += `Total Teknisi: ${sortedTeknisi.length}\n\n`;
        msg += '<b>TOP 20:</b>\n';
        
        sortedTeknisi.slice(0, 20).forEach(([teknisi, count], index) => {
          let icon = '';
          if (index === 0) icon = '🥇';
          else if (index === 1) icon = '🥈';
          else if (index === 2) icon = '🥉';
          else icon = `${index + 1}.`;
          
          msg += `${icon} ${teknisi}: <b>${count} SSL</b>\n`;
        });
        
        if (sortedTeknisi.length > 20) {
          msg += `\n... dan ${sortedTeknisi.length - 20} teknisi lainnya`;
        }
      }
      
      msg += `\n⏰ ${new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' })} WIB`;
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /allps: Ringkasan total aktivasi (admin only) ===
    else if (/^\/allps\b/i.test(text)) {
      if (!(await isAdmin(username))) {
        return sendTelegram(chatId, '❌ Hanya admin yang bisa menggunakan command ini.', { reply_to_message_id: messageId });
      }
      
      const data = await getSheetData(REKAPAN_SHEET);
      let total = Math.max(0, data.length - 1);
      let channelMap = {}, workzoneMap = {}, teknisiMap = {};
      
      for (let i = 1; i < data.length; i++) {
        const channel = (data[i][1] || '-').toUpperCase();
        const workzone = (data[i][7] || '-').toUpperCase();
        const teknisi = (data[i][17] || '-').toUpperCase();
        channelMap[channel] = (channelMap[channel] || 0) + 1;
        workzoneMap[workzone] = (workzoneMap[workzone] || 0) + 1;
        teknisiMap[teknisi] = (teknisiMap[teknisi] || 0) + 1;
      }
      
      let msg = '📊 <b>RINGKASAN AKTIVASI TOTAL</b>\n';
      msg += `TOTAL KESELURUHAN: ${total} SSL\n\n`;
      
      msg += '<b>BERDASARKAN CHANNEL:</b>\n';
      Object.entries(channelMap).sort((a, b) => b[1] - a[1]).forEach(([ch, c]) => {
        msg += `• ${ch}: ${c}\n`;
      });
      
      msg += '\n<b>BERDASARKAN WORKZONE:</b>\n';
      Object.entries(workzoneMap).sort((a, b) => b[1] - a[1]).forEach(([w, c]) => {
        msg += `• ${w}: ${c}\n`;
      });
      
      let teknisiArr = Object.entries(teknisiMap).map(([name, count]) => ({ name, count }));
      teknisiArr.sort((a, b) => b.count - a.count);
      
      msg += '\n<b>TOP 5 TEKNISI:</b>\n';
      teknisiArr.slice(0, 5).forEach((t, i) => {
        msg += `${i + 1}. ${t.name}: ${t.count}\n`;
      });
      
      msg += `\n⏰ ${new Date().toLocaleString('id-ID', { timeZone: 'Asia/Jakarta' })} WIB`;
      return sendTelegram(chatId, msg, { reply_to_message_id: messageId });
    }
    
    // === /help: Bantuan ===
    else if (/^\/help\b/i.test(text) || /^\/start\b/i.test(text)) {
      let helpMsg = '🤖 <b>Bot Rekapan Quality</b>\n\n';
      
      helpMsg += '<b>📝 Commands User:</b>\n';
      helpMsg += '• <code>/aktivasi [data]</code> - Input aktivasi\n';
      helpMsg += '• <code>/cari</code> - Statistik Anda\n';
      helpMsg += '• <code>/exportcari</code> - Download data aktivasi (CSV)\n';
      helpMsg += '• <code>/help</code> - Bantuan\n\n';
      
      if (await isAdmin(username)) {
        helpMsg += '<b>👑 Admin Commands:</b>\n';
        helpMsg += '• <code>/ps</code> - Laporan harian\n';
        helpMsg += '• <code>/ps [dd/mm/yyyy]</code> - Laporan tanggal custom\n';
        helpMsg += '• <code>/topteknisi [periode] [tanggal]</code> - Ranking teknisi\n';
        helpMsg += '   Periode: all, daily, weekly, monthly\n';
        helpMsg += '• <code>/allps</code> - Ringkasan total\n\n';
      }
      
      helpMsg += '<b>📊 Format Input:</b>\n';
      helpMsg += '<code>AO : [value]\nCHANNEL : [value]\nSERVICE NO : [value]\n... (dan field lainnya)</code>\n\n';
      
      helpMsg += '🚀 Bot siap membantu aktivasi Anda!';
      
      return sendTelegram(chatId, helpMsg, { reply_to_message_id: messageId });
    }
    
    // Default
    else if (text.startsWith('/')) {
      return sendTelegram(chatId, '❓ Command tidak dikenali. Ketik /help untuk bantuan.', { reply_to_message_id: messageId });
    }
    
  } catch (err) {
    console.error('Error:', err);
    return sendTelegram(chatId, '❌ Terjadi kesalahan sistem. Coba lagi nanti.', { reply_to_message_id: messageId });
  }
});

// Error handling
process.on('uncaughtException', (err) => {
  console.error('Uncaught Exception:', err);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('Unhandled Rejection:', reason);
});

console.log('\n✓ Bot Rekapan Quality started!');
console.log(`✓ Mode: ${USE_WEBHOOK ? 'Webhook' : 'Polling'}`);
console.log(`✓ MASTER sheet: ${MASTER_SHEET}`);
console.log(`✓ REKAPAN sheet: ${REKAPAN_SHEET}`);
