const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('finance.db');

db.serialize(() => {
    // 1. Tạo bảng users trước nếu chưa tồn tại
    db.run(`CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE,
        password TEXT,
        role TEXT DEFAULT 'staff'
    )`, (err) => {
        if (err) console.error("Lỗi tạo bảng:", err.message);
    });

    // 2. Thêm tài khoản Quản lý (Manager)
    db.run(`INSERT OR IGNORE INTO users (username, password, role) VALUES (?, ?, ?)`,
        ['manager1', '123456', 'manager'], (err) => {
            if (err) console.log("Lỗi insert manager1:", err.message);
            else console.log("Đã tạo/đã có tài khoản Quản lý: manager1 / 123456");
        });

    // 3. Thêm tài khoản Nhân viên (Staff)
    db.run(`INSERT OR IGNORE INTO users (username, password, role) VALUES (?, ?, ?)`,
        ['staff1', '123456', 'staff'], (err) => {
            if (err) console.log("Lỗi insert staff1:", err.message);
            else console.log("Đã tạo/đã có tài khoản Nhân viên: staff1 / 123456");
        });

    // 4. Thêm tài khoản Admin (Phòng hờ)
    db.run(`INSERT OR IGNORE INTO users (username, password, role) VALUES (?, ?, ?)`,
        ['admin', '123456', 'superadmin'], (err) => {
            if (err) console.log("Lỗi insert admin:", err.message);
            else console.log("Đã tạo/đã có tài khoản Admin: admin / 123456");
        });
});

db.close();