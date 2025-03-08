/**
 * Konfigurasi webhook WhatsApp
 */
const WEBHOOK_CONFIG = {
  url: "url webhook whatsapp HelmAi", // Ganti dengan URL server Anda
  timeout: 30000 // Timeout dalam milidetik
};

/**
 * Membuat UI dengan tombol untuk akses cepat ke fungsi-fungsi
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Webhook Tools')
    .addItem('Kirim Pesan WhatsApp', 'showSendMessageDialog')
    .addItem('Test Webhook', 'testWebhook')
    .addSeparator()
    .addItem('Generate Laporan', 'generateReport')
    .addToUi();
}

/**
 * Menampilkan dialog untuk mengirim pesan WhatsApp
 */
function showSendMessageDialog() {
  const html = HtmlService.createHtmlOutput(`
    <html>
      <head>
        <base target="_top">
        <style>
          body { font-family: Arial, sans-serif; margin: 0; padding: 15px; }
          .form-group { margin-bottom: 15px; }
          label { display: block; margin-bottom: 5px; font-weight: bold; }
          input, textarea, select { width: 100%; padding: 8px; box-sizing: border-box; }
          .button-row { display: flex; justify-content: flex-end; margin-top: 20px; }
          button { padding: 8px 15px; background-color: #4285f4; color: white; border: none; border-radius: 4px; }
          .result { margin-top: 20px; padding: 10px; border: 1px solid #ccc; display: none; }
          .success { background-color: #d4edda; }
          .error { background-color: #f8d7da; }
        </style>
      </head>
      <body>
        <h3>Kirim Pesan WhatsApp</h3>
        
        <div class="form-group">
          <label for="phoneNumber">Nomor WhatsApp:</label>
          <input type="text" id="phoneNumber" placeholder="628123456789" />
        </div>
        
        <div class="form-group">
          <label for="name">Nama (Opsional):</label>
          <input type="text" id="name" placeholder="Nama pengirim" />
        </div>
        
        <div class="form-group">
          <label for="messageType">Jenis Pesan:</label>
          <select id="messageType">
            <option value="custom">Pesan Custom</option>
            <option value="mulai">/mulai - Memulai sesi</option>
            <option value="status">/status - Cek status</option>
            <option value="berhenti">/berhenti - Mengakhiri sesi</option>
            <option value="cs">/cs - Mode CS</option>
            <option value="konten">/konten - Mode konten</option>
            <option value="sales">/sales - Mode sales</option>
            <option value="clear">/clear - Bersihkan history</option>
          </select>
        </div>
        
        <div class="form-group" id="customMessageGroup">
          <label for="message">Pesan:</label>
          <textarea id="message" rows="4" placeholder="Ketik pesan Anda di sini..."></textarea>
        </div>
        
        <div class="button-row">
          <button onclick="sendMessage()">Kirim Pesan</button>
        </div>
        
        <div id="resultContainer" class="result">
          <div id="resultContent"></div>
        </div>
        
        <script>
          // Event handler untuk dropdown jenis pesan
          document.getElementById('messageType').addEventListener('change', function() {
            const customGroup = document.getElementById('customMessageGroup');
            const messageInput = document.getElementById('message');
            
            if (this.value === 'custom') {
              customGroup.style.display = 'block';
              messageInput.value = '';
            } else {
              customGroup.style.display = 'none';
              
              // Set nilai pesan berdasarkan jenis pesan yg dipilih
              switch(this.value) {
                case 'mulai': messageInput.value = '/mulai'; break;
                case 'status': messageInput.value = '/status'; break;
                case 'berhenti': messageInput.value = '/berhenti'; break;
                case 'cs': messageInput.value = '/cs'; break;
                case 'konten': messageInput.value = '/konten'; break;
                case 'sales': messageInput.value = '/sales'; break;
                case 'clear': messageInput.value = '/clear'; break;
              }
            }
          });
          
          // Fungsi untuk mengirim pesan
          function sendMessage() {
            const phoneNumber = document.getElementById('phoneNumber').value.trim();
            const name = document.getElementById('name').value.trim();
            const messageType = document.getElementById('messageType').value;
            let message = '';
            
            if (messageType === 'custom') {
              message = document.getElementById('message').value.trim();
            } else {
              message = '/' + messageType;
            }
            
            if (!phoneNumber) {
              showResult(false, 'Nomor WhatsApp harus diisi!');
              return;
            }
            
            if (!message) {
              showResult(false, 'Pesan harus diisi!');
              return;
            }
            
            // Kirim data ke fungsi Apps Script
            google.script.run
              .withSuccessHandler(function(response) {
                if (response && response.success) {
                  showResult(true, 'Pesan berhasil dikirim!<br><br>Respons: ' + 
                    JSON.stringify(response.data, null, 2));
                } else {
                  showResult(false, 'Gagal mengirim pesan!<br><br>Error: ' + 
                    JSON.stringify(response.error || response, null, 2));
                }
              })
              .withFailureHandler(function(error) {
                showResult(false, 'Error: ' + error);
              })
              .sendWebhookMessage(phoneNumber, message, name);
          }
          
          // Tampilkan hasil
          function showResult(success, message) {
            const container = document.getElementById('resultContainer');
            const content = document.getElementById('resultContent');
            
            container.style.display = 'block';
            container.className = 'result ' + (success ? 'success' : 'error');
            content.innerHTML = message;
          }
        </script>
      </body>
    </html>
  `).setWidth(400).setHeight(500).setTitle('Kirim Pesan WhatsApp');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Kirim Pesan WhatsApp');
}

/**
 * Fungsi untuk mengirim pesan melalui webhook
 * Dipanggil dari dialog UI
 */
function sendWebhookMessage(phoneNumber, message, name = "") {
  const formattedPhone = formatPhoneNumber(phoneNumber);
  return sendToWebhook(formattedPhone, message, name);
}

/**
 * Format nomor telepon ke format yang benar
 * @param {string} phone - Nomor telepon input
 * @returns {string} Nomor telepon yang diformat
 */
function formatPhoneNumber(phone) {
  // Hapus semua karakter non-digit
  let cleaned = phone.replace(/\D/g, '');
  
  // Hapus awalan 0
  if (cleaned.startsWith('0')) {
    cleaned = cleaned.substring(1);
  }
  
  // Tambahkan kode negara Indonesia jika belum ada
  if (!cleaned.startsWith('62')) {
    cleaned = '62' + cleaned;
  }
  
  return cleaned;
}

/**
 * Kirim data ke webhook WhatsApp
 * @param {string} phoneNumber - Nomor WhatsApp penerima
 * @param {string} message - Pesan yang akan dikirim
 * @param {string} name - Nama pengirim (opsional)
 * @returns {Object} Respons dari webhook
 */
function sendToWebhook(phoneNumber, message, name = "") {
  const payload = {
    from: phoneNumber,
    message: message,
    name: name,
    device: "Google Sheets App",
    bufferImage: null
  };
  
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    timeout: WEBHOOK_CONFIG.timeout
  };
  
  try {
    const response = UrlFetchApp.fetch(WEBHOOK_CONFIG.url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    // Log aktivitas ke sheet
    logActivity(phoneNumber, message, name, responseCode, responseText);
    
    if (responseCode < 200 || responseCode >= 300) {
      Logger.log("Webhook error: " + responseCode + " - " + responseText);
      return { success: false, code: responseCode, message: responseText };
    }
    
    let responseData;
    try {
      responseData = JSON.parse(responseText);
    } catch (e) {
      responseData = responseText;
    }
    
    return { success: true, data: responseData };
  } catch (error) {
    Logger.log("Error calling webhook: " + error);
    logActivity(phoneNumber, message, name, 500, error.toString());
    return { success: false, error: error.toString() };
  }
}

/**
 * Log aktivitas ke sheet
 */
function logActivity(phoneNumber, message, name, responseCode, responseText) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Activity Log");
  
  if (!sheet) {
    sheet = ss.insertSheet("Activity Log");
    sheet.appendRow([
      "Timestamp",
      "Phone Number",
      "Name",
      "Message",
      "Response Code",
      "Response"
    ]);
    sheet.getRange(1, 1, 1, 6).setFontWeight("bold").setBackground("#f3f3f3");
  }
  
  sheet.appendRow([
    new Date(),
    phoneNumber,
    name,
    message,
    responseCode,
    responseText
  ]);
}

/**
 * Test koneksi webhook
 */
function testWebhook() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const options = {
      method: "get",
      muteHttpExceptions: true,
      timeout: WEBHOOK_CONFIG.timeout
    };
    
    const response = UrlFetchApp.fetch(WEBHOOK_CONFIG.url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    if (responseCode >= 200 && responseCode < 300) {
      ui.alert(
        'Sukses',
        'Koneksi ke webhook berhasil!\n\nResponse Code: ' + responseCode + '\n\nResponse: ' + responseText,
        ui.ButtonSet.OK
      );
    } else {
      ui.alert(
        'Gagal',
        'Koneksi ke webhook gagal!\n\nResponse Code: ' + responseCode + '\n\nResponse: ' + responseText,
        ui.ButtonSet.OK
      );
    }
  } catch (error) {
    ui.alert(
      'Error',
      'Terjadi error saat terhubung ke webhook:\n\n' + error.toString(),
      ui.ButtonSet.OK
    );
  }
}

/**
 * Generate laporan aktivitas webhook yang diperbaiki
 * Dengan perbaikan jumlah kolom yang konsisten
 */
function generateReport() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let logSheet = ss.getSheetByName("Activity Log");
  
  if (!logSheet || logSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("Tidak ada data aktivitas untuk dianalisis. Kirim beberapa pesan terlebih dahulu.");
    return;
  }
  
  // Cek apakah sheet laporan sudah ada, jika ya hapus dan buat baru
  let reportSheet = ss.getSheetByName("Laporan Aktivitas");
  if (reportSheet) {
    ss.deleteSheet(reportSheet);
  }
  
  reportSheet = ss.insertSheet("Laporan Aktivitas");
  
  // Dapatkan semua data log
  const lastRow = logSheet.getLastRow();
  if (lastRow <= 1) {
    SpreadsheetApp.getUi().alert("Tidak ada data aktivitas untuk dianalisis. Sheet Activity Log kosong.");
    return;
  }
  
  const dataRange = logSheet.getRange(2, 1, lastRow - 1, 6).getValues();
  
  // Pastikan dataRange memiliki data
  if (!dataRange || dataRange.length === 0) {
    SpreadsheetApp.getUi().alert("Tidak dapat mengambil data dari Activity Log. Pastikan format data benar.");
    return;
  }
  
  // Statistik dasar
  const totalMessages = dataRange.length;
  let successCount = 0;
  let failCount = 0;
  let commands = {};
  let phoneNumbers = {};
  
  dataRange.forEach(function(row) {
    // Pastikan row valid dan memiliki semua elemen yang dibutuhkan
    if (!row || row.length < 5) return;
    
    const timestamp = row[0] || new Date();
    const phone = row[1] || "Unknown";
    const name = row[2] || "";
    const message = row[3] || "";
    const responseCode = row[4] || 0;
    
    // Hitung sukses/gagal
    if (responseCode >= 200 && responseCode < 300) {
      successCount++;
    } else {
      failCount++;
    }
    
    // Hitung perintah
    if (message && message.toString().startsWith('/')) {
      const command = message.toString().split(' ')[0];
      commands[command] = (commands[command] || 0) + 1;
    }
    
    // Hitung per nomor telepon
    if (phone && phone.toString().trim() !== "") {
      phoneNumbers[phone] = phoneNumbers[phone] || { count: 0, name: name || "Unknown" };
      phoneNumbers[phone].count++;
    }
  });
  
  // Buat setiap bagian laporan secara terpisah
  
  // --- Judul Laporan ---
  reportSheet.getRange("A1").setValue("Laporan Aktivitas Webhook");
  reportSheet.getRange("A1").setFontWeight("bold").setFontSize(14);
  
  reportSheet.getRange("A2").setValue("Dibuat pada:");
  reportSheet.getRange("B2").setValue(new Date());
  
  // --- Ringkasan ---
  reportSheet.getRange("A4").setValue("Ringkasan");
  reportSheet.getRange("A4").setFontWeight("bold").setBackground("#f3f3f3");
  
  reportSheet.getRange("A5").setValue("Total pesan:");
  reportSheet.getRange("B5").setValue(totalMessages);
  
  // Cek untuk division by zero
  const successPercentage = totalMessages > 0 ? (successCount / totalMessages * 100).toFixed(2) + "%" : "0%";
  const failPercentage = totalMessages > 0 ? (failCount / totalMessages * 100).toFixed(2) + "%" : "0%";
  
  reportSheet.getRange("A6").setValue("Pesan sukses:");
  reportSheet.getRange("B6").setValue(successCount);
  reportSheet.getRange("C6").setValue(successPercentage);
  
  reportSheet.getRange("A7").setValue("Pesan gagal:");
  reportSheet.getRange("B7").setValue(failCount);
  reportSheet.getRange("C7").setValue(failPercentage);
  
  // --- Analisis Perintah ---
  let currentRow = 9;
  
  reportSheet.getRange("A" + currentRow).setValue("Analisis Perintah");
  reportSheet.getRange("A" + currentRow).setFontWeight("bold").setBackground("#f3f3f3");
  currentRow++;
  
  // Header analisis perintah
  reportSheet.getRange("A" + currentRow + ":C" + currentRow).setValues([["Perintah", "Jumlah", "Persentase"]]);
  reportSheet.getRange("A" + currentRow + ":C" + currentRow).setFontWeight("bold");
  currentRow++;
  
  // Isi data perintah
  const commandKeys = Object.keys(commands);
  if (commandKeys.length > 0) {
    for (let i = 0; i < commandKeys.length; i++) {
      const command = commandKeys[i];
      const count = commands[command];
      const percentage = totalMessages > 0 ? (count / totalMessages * 100).toFixed(2) + "%" : "0%";
      
      reportSheet.getRange("A" + currentRow + ":C" + currentRow).setValues([[command, count, percentage]]);
      currentRow++;
    }
  } else {
    reportSheet.getRange("A" + currentRow + ":C" + currentRow).setValues([["Tidak ada perintah yang digunakan", "-", "-"]]);
    currentRow++;
  }
  
  // --- Analisis Pengguna ---
  currentRow += 1; // Tambah baris kosong
  
  reportSheet.getRange("A" + currentRow).setValue("Analisis Pengguna");
  reportSheet.getRange("A" + currentRow).setFontWeight("bold").setBackground("#f3f3f3");
  currentRow++;
  
  // Header analisis pengguna
  reportSheet.getRange("A" + currentRow + ":D" + currentRow).setValues([["Nomor Telepon", "Nama", "Jumlah Pesan", "Persentase"]]);
  reportSheet.getRange("A" + currentRow + ":D" + currentRow).setFontWeight("bold");
  currentRow++;
  
  // Isi data pengguna
  const phoneKeys = Object.keys(phoneNumbers);
  if (phoneKeys.length > 0) {
    for (let i = 0; i < phoneKeys.length; i++) {
      const phone = phoneKeys[i];
      const info = phoneNumbers[phone];
      const percentage = totalMessages > 0 ? (info.count / totalMessages * 100).toFixed(2) + "%" : "0%";
      
      reportSheet.getRange("A" + currentRow + ":D" + currentRow).setValues([[phone, info.name, info.count, percentage]]);
      currentRow++;
    }
  } else {
    reportSheet.getRange("A" + currentRow + ":D" + currentRow).setValues([["Tidak ada nomor telepon yang direkam", "-", "-", "-"]]);
    currentRow++;
  }
  
  // Format sheet
  reportSheet.autoResizeColumns(1, 4);
  
  // Tambahkan grafik jika ada data
  if (totalMessages > 0 && successCount + failCount > 0) {
    try {
      // Buat grafik sukses vs gagal
      let chartDataRange = reportSheet.getRange("A6:B7");
      
      let chart = reportSheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(chartDataRange)
        .setPosition(5, 6, 0, 0)
        .setOption('title', 'Sukses vs Gagal')
        .setOption('width', 400)
        .setOption('height', 300)
        .build();
      
      reportSheet.insertChart(chart);
    } catch (error) {
      Logger.log("Error saat membuat grafik: " + error);
      // Lanjutkan meskipun terjadi error pada grafik
    }
  }
  
  // Tampilkan pesan
  SpreadsheetApp.getUi().alert("Laporan berhasil dibuat di sheet 'Laporan Aktivitas'");
  ss.setActiveSheet(reportSheet);
}

/**
 * Fungsi untuk batch mengirim pesan ke beberapa nomor
 * Berguna untuk pengujian massal
 */
function batchSendMessages() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Cek apakah sheet penerima ada
  let recipientSheet = ss.getSheetByName("Recipients");
  if (!recipientSheet) {
    // Tanya user apakah ingin membuat sheet penerima
    const response = ui.alert(
      'Sheet Penerima Tidak Ditemukan',
      'Sheet "Recipients" tidak ditemukan. Apakah Anda ingin membuatnya?',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      recipientSheet = ss.insertSheet("Recipients");
      recipientSheet.appendRow(["Nomor WhatsApp", "Nama", "Pesan"]);
      recipientSheet.appendRow(["628123456789", "John Doe", "/status"]);
      recipientSheet.appendRow(["628987654321", "Jane Smith", "/mulai"]);
      recipientSheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#f3f3f3");
      recipientSheet.autoResizeColumns(1, 3);
      
      ui.alert(
        'Sheet Penerima Dibuat',
        'Sheet "Recipients" telah dibuat. Silakan isi dengan data penerima dan pesan, kemudian jalankan fungsi ini lagi.',
        ui.ButtonSet.OK
      );
      ss.setActiveSheet(recipientSheet);
      return;
    } else {
      return;
    }
  }
  
  // Dapatkan data penerima
  const dataRange = recipientSheet.getDataRange().getValues();
  if (dataRange.length <= 1) {
    ui.alert('Data Kosong', 'Tidak ada data penerima. Silakan isi sheet "Recipients".', ui.ButtonSet.OK);
    return;
  }
  
  // Konfirmasi pengiriman
  const confirmResponse = ui.alert(
    'Konfirmasi Pengiriman',
    'Anda akan mengirim pesan ke ' + (dataRange.length - 1) + ' nomor. Lanjutkan?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResponse !== ui.Button.YES) {
    return;
  }
  
  // Kirim pesan
  let successCount = 0;
  let failCount = 0;
  
  // Skip baris header
  for (let i = 1; i < dataRange.length; i++) {
    const row = dataRange[i];
    const phoneNumber = row[0].toString();
    const name = row[1].toString();
    const message = row[2].toString();
    
    if (!phoneNumber || !message) continue;
    
    try {
      const response = sendToWebhook(formatPhoneNumber(phoneNumber), message, name);
      if (response.success) {
        successCount++;
      } else {
        failCount++;
      }
      
      // Tambahkan delay untuk menghindari rate limiting
      Utilities.sleep(1000);
    } catch (error) {
      failCount++;
      Logger.log("Error sending to " + phoneNumber + ": " + error);
    }
  }
  
  // Tampilkan hasil
  ui.alert(
    'Hasil Pengiriman',
    'Berhasil: ' + successCount + '\nGagal: ' + failCount,
    ui.ButtonSet.OK
  );
}

/**
 * Fungsi untuk membuat form pengiriman WhatsApp yang terhubung ke webhook
 */
function createWhatsAppForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Tanya nama form
  const formNameResponse = ui.prompt(
    'Buat Form WhatsApp',
    'Masukkan nama untuk form:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (formNameResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const formName = formNameResponse.getResponseText() || "Form WhatsApp Toko Helm";
  
  // Buat form baru
  const form = FormApp.create(formName)
    .setTitle(formName)
    .setDescription('Form ini terhubung dengan sistem WhatsApp Toko Helm.');
  
  // Tambahkan pertanyaan
  form.addTextItem()
    .setTitle('Nama')
    .setRequired(true);
    
  form.addTextItem()
    .setTitle('Nomor WhatsApp')
    .setRequired(true)
    .setHelpText('Contoh: 08123456789');
    
  const messageTypeItem = form.addMultipleChoiceItem()
    .setTitle('Jenis Pesan')
    .setChoiceValues([
      'Pertanyaan Produk',
      'Keluhan',
      'Request Katalog',
      'Lainnya'
    ])
    .setRequired(true);
    
  form.addParagraphTextItem()
    .setTitle('Pesan')
    .setRequired(true);
    
  // Tambahkan trigger untuk menangani submission form
  const triggerId = ScriptApp.newTrigger('onFormSubmit')
    .forForm(form)
    .onFormSubmit()
    .create()
    .getUniqueId();
    
  // Simpan ID form dan trigger untuk referensi
  PropertiesService.getScriptProperties().setProperty('WEBHOOK_FORM_ID', form.getId());
  PropertiesService.getScriptProperties().setProperty('WEBHOOK_FORM_TRIGGER_ID', triggerId);
  
  // Tampilkan link form
  ui.alert(
    'Form Berhasil Dibuat',
    'Form telah dibuat dan terhubung ke webhook WhatsApp.\n\nURL Form:\n' + form.getPublishedUrl(),
    ui.ButtonSet.OK
  );
}
