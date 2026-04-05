const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { execSync } = require('child_process');
const sqlite3 = require('sqlite3').verbose();
const Docxtemplater = require('docxtemplater');
const PizZip = require('pizzip');
const ExcelJS = require('exceljs');

let mainWindow;
let db;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1280,
    height: 720,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false,
      enableRemoteModule: true
    },
    icon: path.join(__dirname, 'assets', 'icon.ico')
  });

  mainWindow.loadFile(path.join(__dirname, 'src', 'renderer', 'index.html'));
  
  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

function initDatabase() {
  let dbPath;
  if (app.isPackaged) {
    dbPath = path.join(process.resourcesPath, 'database', 'data.db');
  } else {
    dbPath = path.join(__dirname, 'database', 'data.db');
  }

  db = new sqlite3.Database(dbPath, (err) => {
    if (err) console.error('Database opening error: ', err);
  });

  db.serialize(() => {
    // Bảng Rooms
    db.run(`CREATE TABLE IF NOT EXISTS rooms (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      name TEXT UNIQUE,
      capacity INTEGER DEFAULT 1,
      status INTEGER DEFAULT 0
    )`);
    
    // Bảng Tenants (Người thuê)
    db.run(`CREATE TABLE IF NOT EXISTS tenants (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      room_id INTEGER,
      fullname TEXT,
      phone TEXT,
      cccd TEXT,
      cccd_image TEXT,
      move_in_date DATE,
      status INTEGER DEFAULT 1,
      FOREIGN KEY(room_id) REFERENCES rooms(id)
    )`);

    // Bảng Settings
    db.run(`CREATE TABLE IF NOT EXISTS settings (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      phone TEXT DEFAULT '0901234567',
      room_fee REAL DEFAULT 1500000,
      electricity_price REAL DEFAULT 3500,
      water_price REAL DEFAULT 20000,
      garbage_fee REAL DEFAULT 20000,
      internet_fee REAL DEFAULT 100000,
      electricity_loss REAL DEFAULT 5,
      excluded_rooms TEXT DEFAULT '[]'
    )`);

    // Thêm cột phone nếu chưa có (cho database cũ) - kiểm tra trước để tránh lỗi trùng
    db.all(`PRAGMA table_info(settings)`, (err, rows) => {
      if (!err) {
        const hasPhone = rows.some(row => row.name === 'phone');
        if (!hasPhone) {
          db.run(`ALTER TABLE settings ADD COLUMN phone TEXT DEFAULT '0901234567'`);
        }
      }
    });

    // Bảng Billing
    db.run(`CREATE TABLE IF NOT EXISTS billing (
      id INTEGER PRIMARY KEY AUTOINCREMENT,
      room_id INTEGER,
      month INTEGER,
      year INTEGER,
      electricity_old REAL,
      electricity_new REAL,
      water_old REAL,
      water_new REAL,
      extra_fee REAL DEFAULT 0,
      note TEXT,
      locked INTEGER DEFAULT 0,
      FOREIGN KEY(room_id) REFERENCES rooms(id)
    )`);

    // Insert các phòng mặc định nếu chưa có
    const defaultRooms = ['1A','2A','3A','4A','5A','6A','1B','2B','3B','4B','5B','6B'];
    defaultRooms.forEach(room => {
      db.run(`INSERT OR IGNORE INTO rooms (name) VALUES (?)`, [room]);
    });

    // Insert setting mặc định nếu chưa có
    db.get(`SELECT COUNT(*) as count FROM settings`, (err, row) => {
      if (row.count === 0) {
        db.run(`INSERT INTO settings DEFAULT VALUES`);
      }
    });
  });
}

app.whenReady().then(() => {
  initDatabase();
  createWindow();
});

app.on('window-all-closed', () => {
  if (db) db.close();
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (mainWindow === null) createWindow();
});

// IPC Handlers
ipcMain.handle('get-rooms', async () => {
  return new Promise((resolve, reject) => {
    db.all(`SELECT * FROM rooms ORDER BY name`, (err, rows) => {
      if (err) reject(err);
      resolve(rows);
    });
  });
});

ipcMain.handle('get-settings', async () => {
  return new Promise((resolve, reject) => {
    db.get(`SELECT * FROM settings LIMIT 1`, (err, row) => {
      if (err) reject(err);
      resolve(row);
    });
  });
});

ipcMain.handle('get-billing', async (event, roomId, month, year) => {
  return new Promise((resolve, reject) => {
    db.get(`SELECT * FROM billing WHERE room_id = ? AND month = ? AND year = ?`, [roomId, month, year], (err, row) => {
      if (err) return reject(err);
      
      if(row) {
        resolve(row);
      } else {
        // Tự động lấy dữ liệu tháng trước
        let prevMonth = month - 1;
        let prevYear = year;
        if(prevMonth < 1) {
          prevMonth = 12;
          prevYear = year - 1;
        }
        
        db.get(`SELECT electricity_new, water_new FROM billing WHERE room_id = ? AND month = ? AND year = ?`, 
                [roomId, prevMonth, prevYear], 
                (err, lastMonth) => {
                  if(err) return reject(err);
                  
                  if(lastMonth) {
                    resolve({
                      electricity_old: lastMonth.electricity_new,
                      electricity_new: null,
                      water_old: lastMonth.water_new,
                      water_new: null,
                      locked: 0
                    });
                  } else {
                    resolve(null);
                  }
                });
      }
    });
  });
});

ipcMain.handle('save-billing', async (event, billingData) => {
  return new Promise((resolve, reject) => {
    const { room_id, month, year, electricity_old, electricity_new, water_old, water_new, extra_fee, note, locked } = billingData;

    // Kiểm tra xem bản ghi đã tồn tại chưa
    db.get(`SELECT id FROM billing WHERE room_id = ? AND month = ? AND year = ?`, [room_id, month, year], (err, row) => {
      if (err) return reject(err);

      if (row) {
        // Update nếu đã tồn tại
        db.run(`UPDATE billing SET
                electricity_old = ?, electricity_new = ?, water_old = ?, water_new = ?, extra_fee = ?, note = ?, locked = ?
                WHERE room_id = ? AND month = ? AND year = ?`,
                [electricity_old, electricity_new, water_old, water_new, extra_fee, note, locked, room_id, month, year],
                function(err) {
                  if (err) reject(err);
                  resolve(row.id);
                });
      } else {
        // Insert nếu chưa tồn tại
        db.run(`INSERT INTO billing
                (room_id, month, year, electricity_old, electricity_new, water_old, water_new, extra_fee, note, locked)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`,
                [room_id, month, year, electricity_old, electricity_new, water_old, water_new, extra_fee, note, locked],
                function(err) {
                  if (err) reject(err);
                  resolve(this.lastID);
                });
      }
    });
  });
});

ipcMain.handle('save-settings', async (event, settings) => {
  return new Promise((resolve, reject) => {
    const { phone, room_fee, electricity_price, water_price, garbage_fee, internet_fee, electricity_loss } = settings;

    db.run(`UPDATE settings SET
            phone = ?, room_fee = ?, electricity_price = ?, water_price = ?, garbage_fee = ?, internet_fee = ?, electricity_loss = ?
            WHERE id = 1`,
            [phone, room_fee, electricity_price, water_price, garbage_fee, internet_fee, electricity_loss],
            function(err) {
              if (err) reject(err);
              resolve(this.changes);
            });
  });
});

ipcMain.handle('lock-room', async (event, roomId, month, year) => {
  return new Promise((resolve, reject) => {
    db.run(`UPDATE billing SET locked = 1 WHERE room_id = ? AND month = ? AND year = ?`,
            [roomId, month, year],
            function(err) {
              if (err) reject(err);
              resolve(this.changes);
            });
  });
});

ipcMain.handle('get-tenants', async (event, roomId) => {
  return new Promise((resolve, reject) => {
    db.all(`SELECT * FROM tenants WHERE room_id = ? AND status = 1 ORDER BY id`, [roomId], (err, rows) => {
      if (err) reject(err);
      resolve(rows);
    });
  });
});

ipcMain.handle('save-tenant', async (event, tenant) => {
  return new Promise(async (resolve, reject) => {
    const { room_id, fullname, phone, cccd, imageData } = tenant;

    // Xử lý ảnh nếu có
    let cccd_image = null;
    if (imageData) {
      // imageData có thể là tên file (khi edit) hoặc dữ liệu file mới
      const uploadsDir = path.join(app.isPackaged ? process.resourcesPath : __dirname, 'resources', 'uploads');

      // Đảm bảo thư mục uploads tồn tại
      if (!fs.existsSync(uploadsDir)) {
        fs.mkdirSync(uploadsDir, { recursive: true });
      }

      // Nếu imageData là base64 (file mới được upload từ renderer)
      if (imageData.startsWith('data:')) {
        const base64Data = imageData.replace(/^data:image\/\w+;base64,/, '');
        const ext = imageData.match(/^data:image\/(\w+);base64,/)[1];
        const fileName = `cccd_${Date.now()}.${ext}`;
        const filePath = path.join(uploadsDir, fileName);

        fs.writeFileSync(filePath, Buffer.from(base64Data, 'base64'));
        cccd_image = fileName;
      } else {
        // Giữ nguyên tên file cũ
        cccd_image = imageData;
      }
    }

    // Kiểm tra xem có phải đang sửa (có id trong tenant không)
    if (tenant.id) {
      db.run(`UPDATE tenants SET room_id = ?, fullname = ?, phone = ?, cccd = ?, cccd_image = ?
              WHERE id = ?`,
              [room_id, fullname, phone, cccd, cccd_image, tenant.id],
              function(err) {
                if (err) reject(err);
                resolve(tenant.id);
              });
    } else {
      db.run(`INSERT INTO tenants (room_id, fullname, phone, cccd, cccd_image, move_in_date, status)
              VALUES (?, ?, ?, ?, ?, DATE('now'), 1)`,
              [room_id, fullname, phone, cccd, cccd_image],
              function(err) {
                if (err) reject(err);
                resolve(this.lastID);
              });
    }
  });
});

ipcMain.handle('delete-tenant', async (event, tenantId) => {
  return new Promise((resolve, reject) => {
    db.run(`DELETE FROM tenants WHERE id = ?`, [tenantId], function(err) {
      if (err) reject(err);
      resolve(this.changes);
    });
  });
});

ipcMain.handle('export-invoice', async (event, roomId, month, year) => {
  return new Promise(async (resolve, reject) => {
    try {
      // Lấy thông tin phòng
      const room = await new Promise((res, rej) => {
        db.get(`SELECT r.name FROM rooms r WHERE r.id = ?`, [roomId], (err, row) => {
          if (err) rej(err);
          res(row);
        });
      });

      if (!room) return resolve(false);

      // Lấy thông tin billing
      let billing = await new Promise((res, rej) => {
        db.get(`SELECT electricity_old, electricity_new, water_old, water_new, extra_fee
                FROM billing WHERE room_id = ? AND month = ? AND year = ?`,
                [roomId, month, year],
                (err, row) => {
                  if (err) rej(err);
                  res(row);
                });
      });

      // Nếu chưa có billing, tạo mới với giá trị 0
      if (!billing) {
        billing = { electricity_old: 0, electricity_new: 0, water_old: 0, water_new: 0, extra_fee: 0 };
        await new Promise((res, rej) => {
          db.run(`INSERT INTO billing (room_id, month, year, electricity_old, electricity_new, water_old, water_new, extra_fee)
                  VALUES (?, ?, ?, 0, 0, 0, 0, 0)`,
                  [roomId, month, year],
                  function(err) {
                    if (err) rej(err);
                    res();
                  });
        });
      }

      // Lấy tên người thuê (lấy người đầu tiên nếu có nhiều)
      const tenants = await new Promise((res, rej) => {
        db.all(`SELECT fullname FROM tenants WHERE room_id = ? AND status = 1`, [roomId], (err, rows) => {
          if (err) rej(err);
          res(rows);
        });
      });
      const tenantName = tenants.length > 0 ? tenants[0].fullname : '';

      const settings = await new Promise((res, rej) => {
        db.get(`SELECT * FROM settings LIMIT 1`, (err, row) => {
          if (err) rej(err);
          res(row || {});
        });
      });

      // Tính toán đúng công thức
      // Điện hao tải: Tiền điện hao tải = (Tiền điện của phòng đó) * (% tỷ lệ hao tải)
      const elecUsed = Math.max(0, (billing.electricity_new || 0) - (billing.electricity_old || 0));
      const elecLossPercent = (settings.electricity_loss || 5) / 100;

      // Tiền điện = số kWh * giá điện/kWh
      const elecAmount = elecUsed * (settings.electricity_price || 3500);

      // Tiền điện hao tải = tiền điện của phòng đó * % điện hao tải
      const elecLossAmount = elecAmount * elecLossPercent;

      const waterUsed = Math.max(0, (billing.water_new || 0) - (billing.water_old || 0));
      // Tiền nước = số khối * giá nước/m3
      const waterAmount = waterUsed * (settings.water_price || 20000);

      // Tổng = tiền phòng + tiền điện + tiền nước + tiền internet + tiền điện hao tải + tiền rác
      const total = (settings.room_fee || 0) + elecAmount + waterAmount + (settings.internet_fee || 0) + elecLossAmount + (settings.garbage_fee || 0) + (billing.extra_fee || 0);

      // Load template file
      const templatePath = path.join(app.isPackaged ? process.resourcesPath : __dirname, 'resources', 'PHIẾU THU TIỀN NHÀ.docx');

      if(!fs.existsSync(templatePath)) {
        // Tạo file mẫu nếu chưa tồn tại
        const content = Buffer.from('UEsFBgAAAAAAAAAAAAAAAAAAAAAAAA==', 'base64');
        fs.writeFileSync(templatePath, content);
      }

      const template = fs.readFileSync(templatePath, 'binary');
      const zip = new PizZip(template);

      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true
      });

      // Lấy ngày hiện tại
      const now = new Date();
      const day = now.getDate();

      doc.render({
        phone: settings.phone || '0901234567',
        room_name: room.name,
        day: day,
        month: month,
        year: year,
        room_fee: (settings.room_fee || 0).toLocaleString('vi-VN'),
        electricity_old: billing.electricity_old || 0,
        electricity_new: billing.electricity_new || 0,
        electricity_used: elecUsed,
        electricity_price: (settings.electricity_price || 3500).toLocaleString('vi-VN'),
        electricity_amount: elecAmount.toLocaleString('vi-VN'),
        electricity_loss: elecLossAmount.toLocaleString('vi-VN'),
        water_old: billing.water_old || 0,
        water_new: billing.water_new || 0,
        water_used: waterUsed,
        water_price: (settings.water_price || 20000).toLocaleString('vi-VN'),
        water_amount: waterAmount.toLocaleString('vi-VN'),
        garbage_fee: (settings.garbage_fee || 0).toLocaleString('vi-VN'),
        internet_fee: (settings.internet_fee || 0).toLocaleString('vi-VN'),
        total: total.toLocaleString('vi-VN')
      });

      const output = doc.getZip().generate({ type: 'nodebuffer' });

      const outputPath = path.join(app.getPath('documents'), `Phiếu thu Phòng ${room.name} tháng ${month}-${year}.docx`);
      fs.writeFileSync(outputPath, output);

      resolve(outputPath);
    } catch (error) {
      console.error('export-invoice error:', error);
      reject(error);
    }
  });
});

// Xuất Excel báo cáo tổng hợp 12 phòng
ipcMain.handle('export-excel', async (event, month, year) => {
  return new Promise(async (resolve, reject) => {
    try {
      const settings = await new Promise((res, rej) => {
        db.get(`SELECT * FROM settings LIMIT 1`, (err, row) => {
          if (err) rej(err);
          res(row);
        });
      });

      const rooms = await new Promise((res, rej) => {
        db.all(`SELECT * FROM rooms ORDER BY name`, (err, rows) => {
          if (err) rej(err);
          res(rows);
        });
      });

      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'Quản Lý Phòng Trọ';
      workbook.created = new Date();

      const worksheet = workbook.addWorksheet(`Báo cáo tháng ${month}-${year}`);

      // Tiêu đề
      worksheet.mergeCells('A1:M1');
      worksheet.getCell('A1').value = `BÁO CÁO TIỀN PHÒNG THÁNG ${month} NĂM ${year}`;
      worksheet.getCell('A1').font = { size: 16, bold: true, color: { argb: 'FF667EEA' } };
      worksheet.getCell('A1').alignment = { horizontal: 'center' };

      // Header row
      const headers = ['STT', 'Phòng', 'Điện cũ', 'Điện mới', 'Điện SD', 'Tiền điện', 'Điện hao tải', 'Nước cũ', 'Nước mới', 'Nước SD', 'Tiền nước', 'Phí khác', 'Tổng tiền'];
      worksheet.addRow(headers);

      // Style header
      const headerRow = worksheet.getRow(2);
      headerRow.font = { bold: true, color: { argb: 'FFFFFFFF' } };
      headerRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF667EEA' } };
      headerRow.alignment = { horizontal: 'center' };

      let stt = 0;
      let grandTotal = 0;

      for (const room of rooms) {
        const billing = await new Promise((res, rej) => {
          db.get(`SELECT * FROM billing WHERE room_id = ? AND month = ? AND year = ?`,
            [room.id, month, year], (err, row) => {
              if (err) rej(err);
              res(row);
            });
        });

        stt++;

        // Tính toán
        const elecOld = billing?.electricity_old || 0;
        const elecNew = billing?.electricity_new || 0;
        const waterOld = billing?.water_old || 0;
        const waterNew = billing?.water_new || 0;

        const elecUsed = Math.max(0, elecNew - elecOld);
        const elecLossPercent = (settings.electricity_loss || 5) / 100;

        // Tiền điện = số kWh * giá điện/kWh
        const elecAmount = elecUsed * (settings.electricity_price || 3500);

        // Tiền điện hao tải = tiền điện của phòng đó * % điện hao tải
        const elecLossAmount = elecAmount * elecLossPercent;

        const waterUsed = Math.max(0, waterNew - waterOld);
        // Tiền nước = số khối * giá nước/m3
        const waterAmount = waterUsed * (settings.water_price || 20000);

        // Tổng = tiền phòng + tiền điện + tiền nước + tiền internet + tiền điện hao tải + tiền rác
        const total = (settings.room_fee || 0) + elecAmount + waterAmount + (settings.internet_fee || 0) + elecLossAmount + (settings.garbage_fee || 0) + (billing?.extra_fee || 0);
        grandTotal += total;

        worksheet.addRow([
          stt,
          room.name,
          elecOld,
          elecNew,
          elecUsed,
          elecAmount,
          elecLossAmount,
          waterOld,
          waterNew,
          waterUsed,
          waterAmount,
          (settings.room_fee || 0) + (settings.garbage_fee || 0) + (settings.internet_fee || 0) + (billing?.extra_fee || 0),
          total
        ]);
      }

      // Thêm dòng tổng cộng
      worksheet.addRow([]);
      const totalRow = worksheet.addRow(['', 'TỔNG CỘNG', '', '', '', '', '', '', '', '', '', '', grandTotal]);
      totalRow.font = { bold: true };
      totalRow.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFEBEE' } };

      // Set column widths
      worksheet.getColumn(1).width = 5;
      worksheet.getColumn(2).width = 8;
      worksheet.getColumn(3).width = 10;
      worksheet.getColumn(4).width = 10;
      worksheet.getColumn(5).width = 10;
      worksheet.getColumn(6).width = 12;
      worksheet.getColumn(7).width = 12;
      worksheet.getColumn(8).width = 10;
      worksheet.getColumn(9).width = 10;
      worksheet.getColumn(10).width = 10;
      worksheet.getColumn(11).width = 12;
      worksheet.getColumn(12).width = 12;
      worksheet.getColumn(13).width = 14;

      // Save file
      const outputPath = path.join(app.getPath('documents'), `Báo cáo tháng ${month}-${year}.xlsx`);
      await workbook.xlsx.writeFile(outputPath);

      // Cập nhật master Excel file
      try {
        // Chuẩn bị dữ liệu cho các phòng
        const roomsData = [];
        for (const room of rooms) {
          const billing = await new Promise((res, rej) => {
            db.get(`SELECT * FROM billing WHERE room_id = ? AND month = ? AND year = ?`,
              [room.id, month, year], (err, row) => {
                if (err) rej(err);
                res(row);
              });
          });

          const elecOld = billing?.electricity_old || 0;
          const elecNew = billing?.electricity_new || 0;
          const elecUsed = Math.max(0, elecNew - elecOld);

          // Lấy thông tin người thuê
          const tenants = await new Promise((res, rej) => {
            db.all(`SELECT fullname, cccd FROM tenants WHERE room_id = ? AND status = 1 ORDER BY id`, [room.id], (err, rows) => {
              if (err) rej(err);
              res(rows);
            });
          });

          const fullname = tenants.length > 0 ? tenants[0].fullname : null;
          const cccd = tenants.length > 0 ? tenants[0].cccd : null;

          roomsData.push({
            name: room.name,
            fullname: fullname,
            cccd: cccd,
            elecUsed: elecUsed
          });
        }

        // Gọi Python script để cập nhật master Excel
        const pythonScript = path.join(__dirname, 'export_master_excel.py');
        const jsonData = JSON.stringify({
          month: month,
          year: year,
          rooms: roomsData
        });

        // Escape quotes in JSON for command line
        const escapedJson = jsonData.replace(/"/g, '\\"');
        const cmd = `python "${pythonScript}" --json "{${escapedJson.slice(1, -1)}}"`;

        try {
          const pythonResult = execSync(cmd, { encoding: 'utf8', timeout: 30000 });
          console.log('Master Excel updated:', pythonResult);
        } catch (pythonError) {
          console.warn('Warning: Could not update master Excel:', pythonError.message);
          // Don't fail the export if master Excel update fails
        }
      } catch (masterError) {
        console.warn('Warning: Error updating master Excel:', masterError.message);
        // Don't fail the export if master Excel update fails
      }

      resolve(outputPath);
    } catch (error) {
      reject(error);
    }
  });
});

// Xuất tất cả phiếu thu trong 1 file Word (mỗi phòng một trang)
ipcMain.handle('export-all-invoices', async (event, roomIds, month, year) => {
  return new Promise(async (resolve, reject) => {
    console.log('export-all-invoices called with:', { roomIds, month, year });
    try {
      // Lấy cài đặt
      const settings = await new Promise((res, rej) => {
        db.get(`SELECT * FROM settings LIMIT 1`, (err, row) => {
          if (err) rej(err);
          res(row || {});
        });
      });

      console.log('Settings:', settings);

      const roomsData = [];
      const roomsList = await new Promise((res, rej) => {
        db.all(`SELECT * FROM rooms ORDER BY name`, (err, rows) => {
          if (err) rej(err);
          res(rows);
        });
      });

      // Lấy dữ liệu cho từng phòng
      for (const roomId of roomIds) {
        console.log('Processing roomId:', roomId);

        // Lấy thông tin phòng
        const room = roomsList.find(r => r.id === roomId);
        if (!room) continue;

        // Lấy thông tin billing
        let billing = await new Promise((res, rej) => {
          db.get(`SELECT b.electricity_old, b.electricity_new, b.water_old, b.water_new, b.extra_fee, b.locked
                  FROM billing b
                  WHERE b.room_id = ? AND b.month = ? AND b.year = ?`,
                  [roomId, month, year],
                  (err, row) => {
                    if (err) rej(err);
                    res(row);
                  });
        });

        console.log('Billing for room', roomId, ':', billing);

        // Nếu chưa có billing, tạo mới với giá trị 0
        if (!billing) {
          billing = {
            electricity_old: 0,
            electricity_new: 0,
            water_old: 0,
            water_new: 0,
            extra_fee: 0
          };
          // Insert vào database
          await new Promise((res, rej) => {
            db.run(`INSERT INTO billing (room_id, month, year, electricity_old, electricity_new, water_old, water_new, extra_fee)
                    VALUES (?, ?, ?, 0, 0, 0, 0, 0)`,
                    [roomId, month, year],
                    function(err) {
                      if (err) rej(err);
                      res();
                    });
          });
        }

        // Lấy tên người thuê
        const tenants = await new Promise((res, rej) => {
          db.all(`SELECT fullname FROM tenants WHERE room_id = ? AND status = 1`, [roomId], (err, rows) => {
            if (err) rej(err);
            res(rows);
          });
        });
        const tenantName = tenants.length > 0 ? tenants[0].fullname : '';

        // Tính toán với giá trị mặc định nếu settings null
        const elecUsed = Math.max(0, (billing.electricity_new || 0) - (billing.electricity_old || 0));
        const elecLossPercent = (settings.electricity_loss || 5) / 100;
        const elecAmount = elecUsed * (settings.electricity_price || 3500);
        const elecLossAmount = elecAmount * elecLossPercent;

        const waterUsed = Math.max(0, (billing.water_new || 0) - (billing.water_old || 0));
        const waterAmount = waterUsed * (settings.water_price || 20000);

        const total = (settings.room_fee || 0) + elecAmount + waterAmount + (settings.internet_fee || 0) + elecLossAmount + (settings.garbage_fee || 0) + (billing.extra_fee || 0);

        // Làm tròn lên hàng nghìn
        const roundUp = (num) => Math.ceil(num / 1000) * 1000;

        const now = new Date();
        const dateStr = `${now.getDate()}/${now.getMonth() + 1}/${now.getFullYear()}`;

        roomsData.push({
          room_name: room.name,
          tenant_name: tenantName,
          date: dateStr,
          month: month,
          year: year,
          elec_old: billing.electricity_old || 0,
          elec_new: billing.electricity_new || 0,
          elec_used: elecUsed,
          elec_total: roundUp(elecAmount),
          elec_loss_percent: settings.electricity_loss || 5,
          elec_loss_amount: roundUp(elecLossAmount),
          water_old: billing.water_old || 0,
          water_new: billing.water_new || 0,
          water_used: waterUsed,
          water_total: roundUp(waterAmount),
          room_fee: roundUp(settings.room_fee || 0),
          garbage_fee: roundUp(settings.garbage_fee || 0),
          internet_fee: roundUp(settings.internet_fee || 0),
          extra_fee: roundUp(billing.extra_fee || 0),
          total_amount: roundUp(total),
          total_text: 'Bằng chữ: ' + new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(roundUp(total))
        });
      }

      console.log('roomsData length:', roomsData.length);

      if (roomsData.length === 0) {
        return resolve(false);
      }

      // Load template file
      const templatePath = path.join(app.isPackaged ? process.resourcesPath : __dirname, 'resources', 'PHIẾU THU TIỀN NHÀ.docx');

      console.log('Template path:', templatePath);
      console.log('Template exists:', fs.existsSync(templatePath));

      if (!fs.existsSync(templatePath)) {
        return reject(new Error('Không tìm thấy file mẫu phiếu thu: ' + templatePath));
      }

      const template = fs.readFileSync(templatePath, 'binary');
      const zip = new PizZip(template);

      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true
      });

      // Render dữ liệu với vòng lặp {#rooms}{/rooms}
      doc.render({
        rooms: roomsData
      });

      const output = doc.getZip().generate({ type: 'nodebuffer' });

      // Hiện hộp thoại lưu file
      const { filePath } = await dialog.showSaveDialog(mainWindow, {
        title: 'Lưu phiếu thu',
        defaultPath: `Phiếu thu tháng ${month}-${year}.docx`,
        filters: [{ name: 'Word Documents', extensions: ['docx'] }]
      });

      if (!filePath) {
        return resolve(false); // Người dùng hủy
      }

      fs.writeFileSync(filePath, output);


      require('electron').shell.showItemInFolder(filePath);

      resolve(filePath);
    } catch (error) {
      console.error('export-all-invoices error:', error);
      reject(error);
    }
  });
});


