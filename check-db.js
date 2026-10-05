const sqlite3 = require('sqlite3').verbose();
const fs = require('fs');

const dbFiles = ['database.db', 'finance_old.db', 'finance.db', 'mt_solar.db'];

dbFiles.forEach(file => {
    if (fs.existsSync(file)) {
        const db = new sqlite3.Database(file, sqlite3.OPEN_READONLY, (err) => {
            if (!err) {
                db.all("SELECT name FROM sqlite_master WHERE type='table'", [], (err, tables) => {
                    if (!err) {
                        const tableNames = tables.map(t => t.name);
                        if (tableNames.includes('users')) {
                            db.all("SELECT username, role FROM users", [], (err, users) => {
                                console.Printf ? null : console.log(`\n📁 File: ${file}`);
                                console.log(`   Các bảng: ${tableNames.join(', ')}`);
                                if (users && users.length > 0) {
                                    console.log(`   👤 Users có trong file này:`, users);
                                } else {
                                    console.log(`   👤 Bảng users trống.`);
                                }
                            });
                        }
                    }
                });
            }
        });
    }
});