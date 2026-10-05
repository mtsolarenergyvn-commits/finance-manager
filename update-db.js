const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('finance.db');

db.run("ALTER TABLE users ADD COLUMN role TEXT DEFAULT 'staff'", (err) => {
    if (err) {
        console.log("Cột role có thể đã tồn tại hoặc có lỗi:", err.message);
    } else {
        console.log("Đã thêm thành công cột 'role' vào bảng users!");
    }
    db.close();
});