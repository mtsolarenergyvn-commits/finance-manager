const express = require('express');
const sqlite3 = require('sqlite3').verbose();
const session = require('express-session');
const path = require('path');
const ExcelJS = require('exceljs');
const XLSX = require('xlsx');

const app = express();
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Server is running on port ${PORT}`);
});
const dict = {
    vi: {
        brand: "MT SOLAR - HỆ THỐNG QUẢN TRỊ DOANH NGHIỆP",
        salesTitle: "Quản Lý Bán Hàng",
        salesDesc: "Lên đơn hàng, nhân viên, sản phẩm.",
        salesBtn: "Vào Bán Hàng",
        revTitle: "Quản Lý Doanh Thu",
        revDesc: "Theo dõi dòng tiền thu vào.",
        revBtn: "Vào Doanh Thu",
        expTitle: "Quản Lý Chi Phí",
        expDesc: "Theo dõi dòng tiền chi ra.",
        expBtn: "Vào Chi Phí",
        profTitle: "Báo Cáo Lợi Nhuận",
        profDesc: "Doanh thu trừ Chi phí.",
        profBtn: "Xem Báo Cáo",
        lblEmployee: "Nhân viên",
        lblProduct: "Sản phẩm",
        lblQuantity: "Số lượng",
        lblPrice: "Đơn giá",
        lblDiscount: "Chiết khấu (%)",
        lblAmount: "Thành tiền",
        lblCustomer: "Khách hàng",
        lblDate: "Ngày",
        lblNote: "Ghi chú",
        btnAddOrder: "Thêm Đơn Hàng",
        btnFilter: "Lọc",
        btnAll: "Tất cả",
        btnExport: "Xuất File Excel",
        tableHeaderList: "Danh Sách Đơn Hàng",
        thActions: "Hành động",
        btnEdit: "Sửa",
        btnDelete: "Xóa"
    },
    en: {
        brand: "MT SOLAR - ENTERPRISE MANAGEMENT SYSTEM",
        salesTitle: "Sales Management",
        salesDesc: "Manage orders, staff, and products.",
        salesBtn: "Go to Sales",
        revTitle: "Revenue Management",
        revDesc: "Track incoming cash flows.",
        revBtn: "Go to Revenues",
        expTitle: "Expense Management",
        expDesc: "Track outgoing expenses.",
        expBtn: "Go to Expenses",
        profTitle: "Profit Report",
        profDesc: "Revenue minus Expenses.",
        profBtn: "View Report",
        lblEmployee: "Employee",
        lblProduct: "Product",
        lblQuantity: "Quantity",
        lblPrice: "Unit Price",
        lblDiscount: "Discount (%)",
        lblAmount: "Amount",
        lblCustomer: "Customer",
        lblDate: "Date",
        lblNote: "Note",
        btnAddOrder: "Add Order",
        btnFilter: "Filter",
        btnAll: "All",
        btnExport: "Export Excel",
        tableHeaderList: "Order List",
        thActions: "Actions",
        btnEdit: "Edit",
        btnDelete: "Delete"
    },
    cn: {
        brand: "MT SOLAR - 企业管理系统",
        salesTitle: "销售管理",
        salesDesc: "管理订单、员工与产品。",
        salesBtn: "进入销售",
        revTitle: "收入管理",
        revDesc: "追踪现金流入情况。",
        revBtn: "进入收入",
        expTitle: "费用管理",
        expDesc: "追踪各项支出开销。",
        expBtn: "进入费用",
        profTitle: "利润报表",
        profDesc: "总收入减去总支出。",
        profBtn: "查看报表",
        lblEmployee: "员工",
        lblProduct: "产品",
        lblQuantity: "数量",
        lblPrice: "单价",
        lblDiscount: "折扣 (%)",
        lblAmount: "总金额",
        lblCustomer: "客户",
        lblDate: "日期",
        lblNote: "备注",
        btnAddOrder: "添加订单",
        btnFilter: "筛选",
        btnAll: "全部",
        btnExport: "导出 Excel",
        tableHeaderList: "订单列表",
        thActions: "操作",
        btnEdit: "编辑",
        btnDelete: "删除"
    }
};

app.set('view engine', 'ejs');
app.use(express.static('views'));
app.use(express.urlencoded({ extended: true }));
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

app.use(session({
    secret: 'khoa-bi-mat-mt-solar',
    resave: false,
    saveUninitialized: true
}));

app.use((req, res, next) => {
    const lang = req.session.lang || 'vi';
    res.locals.currentLang = lang;
    res.locals.t = dict[lang];
    next();
});

const db = new sqlite3.Database(path.join(__dirname, 'finance.db'), (err) => {
    if (err) console.error('Lỗi DB:', err.message);
    else console.log('Đã kết nối SQLite thành công.');
});

db.serialize(() => {
    db.run(`CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE,
        password TEXT,
        role TEXT DEFAULT 'staff'
    )`);

    db.get(`SELECT * FROM users WHERE username = 'admin'`, (err, row) => {
        if (!row) {
            db.run(`INSERT INTO users (username, password, role) VALUES ('admin', '123456', 'admin')`);
        }
    });

    db.run(`CREATE TABLE IF NOT EXISTS revenues (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        amount REAL, 
        category TEXT, 
        note TEXT, 
        date TEXT,
        sale_id TEXT DEFAULT NULL
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS expenses (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        amount REAL, 
        category TEXT, 
        note TEXT, 
        date TEXT
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS sales (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        employee TEXT, 
        product TEXT, 
        quantity INTEGER, 
        unit_price REAL, 
        discount REAL,
        amount REAL, 
        customer TEXT, 
        note TEXT, 
        date TEXT,
        invoice_code TEXT
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS quotes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        code TEXT,
        customer TEXT,
        employee TEXT,
        date TEXT,
        total REAL,
        validUntil TEXT,
        note TEXT,
        vatRate REAL
    )`);

    db.run(`CREATE TABLE IF NOT EXISTS employees (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT,
        phone TEXT,
        position TEXT
    )`);

});
// Danh sách tài khoản mẫu cần khởi tạo
const defaultUsers = [
    ['adminkhang', '123123', 'super admin'],
    ['admin1', '123123', 'admin'],
    ['admin2', '123123', 'admin'],
    ['liyongmng', '123456', 'manager'],
    ['uyensale1', '123456', 'staff'],
    ['khangsale2', '123456', 'staff'],
    ['namsale3', '123456', 'staff'],
    ['thucsale4', '123456', 'staff'],
    ['hungsale5', '123456', 'staff'],
    ['tramsale6', '123456', 'staff'],
    ['trieusale7', '123456', 'staff'],
    ['lanwh1', '123456', 'staff'],
    ['thoanwh2', '123456', 'staff']
];

// Chèn tự động vào database nếu chưa tồn tại
defaultUsers.forEach(([username, password, role]) => {
    db.get(`SELECT * FROM users WHERE username = ?`, [username], (err, row) => {
        if (!row) {
            db.run(`INSERT INTO users (username, password, role) VALUES (?, ?, ?)`, [username, password, role]);
        }
    });
});

const checkAuth = (req, res, next) => req.session && req.session.loggedIn ? next() : res.redirect('/login');

app.get('/lang/:locale', (req, res) => {
    const lang = req.params.locale;
    if (['vi', 'en', 'cn'].includes(lang)) {
        req.session.lang = lang;
    }
    res.redirect(req.get('referer') || '/');
});

function getDateFilterSQL(from, to, dateField = 'date') {
    let conditions = [];
    let params = [];
    if (from) {
        conditions.push(`${dateField} >= ?`);
        params.push(from);
    }
    if (to) {
        conditions.push(`${dateField} <= ?`);
        params.push(to);
    }
    let clause = conditions.length > 0 ? `WHERE ` + conditions.join(' AND ') : '';
    return { clause, params };
}

// Middleware kiểm tra quyền hạn (chỉ cho phép các role trong danh sách đi tiếp)
function checkRole(allowedRoles) {
    return (req, res, next) => {
        const userRole = (req.session.user && req.session.user.role) ? req.session.user.role : 'staff';

        // In ra terminal để debug xem tài khoản nào đang truy cập và role là gì
        console.log(`User đang thao tác: ${req.session.username}, Role thực tế: ${userRole}`);

        if (allowedRoles.includes(userRole)) {
            next();
        } else {
            res.send("<script>alert('Bạn chưa được cấp quyền thực hiện chức năng này!'); window.history.back();</script>");
        }
    };
}
// Hàm sinh mã tự động theo ngày: #MTBG-DDMMYYYY01 hoặc #MTXK-DDMMYYYY01
function generateCode(prefix, count) {
    const now = new Date();
    const day = String(now.getDate()).padStart(2, '0');
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const year = now.getFullYear();
    const seq = String(count || 1).padStart(2, '0');
    return `#${prefix}-${day}${month}${year}${seq}`;
}

// ==================== AUTH & ACCOUNT ROUTES ====================
app.get('/login', (req, res) => res.render('login', { error: null }));

app.post('/login', (req, res) => {
    const { username, password } = req.body;
    db.get(`SELECT * FROM users WHERE username = ? AND password = ?`, [username, password], (err, row) => {
        if (row) {
            req.session.loggedIn = true;
            req.session.username = username;
            req.session.user = {
                id: row.id,
                username: row.username,
                role: row.role || 'staff'
            };
            res.redirect('/');
        } else {
            res.render('login', { error: 'Sai tài khoản hoặc mật khẩu!' });
        }
    });
});

app.get('/logout', (req, res) => {
    req.session.destroy(() => res.redirect('/login'));
});

app.post('/change-password', checkAuth, (req, res) => {
    const { oldPassword, newPassword } = req.body;
    const username = req.session.username || 'admin';

    db.get(`SELECT * FROM users WHERE username = ? AND password = ?`, [username, oldPassword], (err, row) => {
        if (!row) {
            return res.send("<script>alert('Mật khẩu cũ không chính xác!'); window.history.back();</script>");
        }
        db.run(`UPDATE users SET password = ? WHERE username = ?`, [newPassword, username], (err) => {
            if (err) {
                return res.send("<script>alert('Lỗi cập nhật mật khẩu!'); window.history.back();</script>");
            }
            res.send("<script>alert('Đổi mật khẩu thành công!'); window.location.href='/';</script>");
        });
    });
});

app.get('/', checkAuth, (req, res) => {
    res.render('index');
});

app.get('/quote', checkAuth, (req, res) => {
    res.render('quote');
});


// ==================== 1. QUẢN LÝ BÁN HÀNG ====================
app.get('/sales', checkAuth, (req, res) => {
    let { from, to } = req.query;
    let clause = '';
    let params = [];

    if (from && to) {
        clause = `WHERE date BETWEEN ? AND ?`;
        params = [from, to];
    }

    db.all(`SELECT * FROM sales ${clause} ORDER BY date DESC, id DESC`, params, (err, salesRows) => {
        if (err) {
            console.error("Lỗi lấy danh sách sales:", err.message);
            salesRows = [];
        }

        db.all(`SELECT * FROM employees ORDER BY id DESC`, (err2, empRows) => {
            res.render('sales', {
                sales: salesRows || [],
                employees: empRows || [],
                from: from || '',
                to: to || ''
            });
        });
    });
});

app.post('/sales/add', checkAuth, (req, res) => {
    let { employee, customer, date, note, products, quantities, unit_prices, discounts, amounts } = req.body;
    const invoice_code = 'INV-' + Date.now();

    const prods = Array.isArray(products) ? products : [products];
    const quants = Array.isArray(quantities) ? quantities : [quantities];
    const prices = Array.isArray(unit_prices) ? unit_prices : [unit_prices];
    const discs = Array.isArray(discounts) ? discounts : [discounts];
    const amts = Array.isArray(amounts) ? amounts : [amounts];

    let totalOrderAmount = 0;
    amts.forEach(amt => totalOrderAmount += Number(amt || 0));

    let completed = 0;
    const totalItems = prods.length;

    prods.forEach((prod, index) => {
        if (!prod) {
            completed++;
            if (completed === totalItems) saveRevenue();
            return;
        }

        const qty = quants[index] || 1;
        const price = prices[index] || 0;
        const disc = discs[index] || 0;
        const amt = amts[index] || (qty * price);

        const sqlSales = `INSERT INTO sales (employee, product, quantity, unit_price, discount, amount, customer, note, date, invoice_code) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)`;

        db.run(sqlSales, [employee, prod, qty, price, disc || 0, amt, customer, note, date, invoice_code], function (err) {
            if (err) console.error("Lỗi thêm sales:", err.message);
            completed++;
            if (completed === totalItems) {
                saveRevenue();
            }
        });
    });

    function saveRevenue() {
        const revNote = `Đơn hàng (${invoice_code}) từ NV: ${employee}, KH: ${customer || 'Khách lẻ'} (${totalItems} sản phẩm)`;
        const sqlRev = `INSERT INTO revenues (amount, category, note, date, sale_id) VALUES (?, ?, ?, ?, ?)`;

        db.run(sqlRev, [totalOrderAmount, 'Bán hàng', revNote, date, invoice_code], (err) => {
            if (err) console.error("Lỗi thêm revenues:", err.message);
            res.redirect('/sales');
        });
    }
});

app.post('/sales/delete/:id', checkAuth, checkRole(['superadmin', 'manager']), (req, res) => {
    const saleId = req.params.id;
    db.get(`SELECT invoice_code FROM sales WHERE id = ?`, [saleId], (err, row) => {
        if (row && row.invoice_code) {
            db.run(`DELETE FROM revenues WHERE sale_id = ?`, [row.invoice_code], () => {
                db.run(`DELETE FROM sales WHERE invoice_code = ?`, [row.invoice_code], () => {
                    res.redirect('/sales');
                });
            });
        } else {
            db.run(`DELETE FROM revenues WHERE sale_id = ?`, [saleId], () => {
                db.run(`DELETE FROM sales WHERE id = ?`, [saleId], () => {
                    res.redirect('/sales');
                });
            });
        }
    });
});

app.post('/sales/edit/:id', checkAuth, checkRole(['superadmin', 'manager']), (req, res) => {
    const saleId = req.params.id;
    const { employee, product, quantity, unit_price, discount, amount, customer, note, date } = req.body;
    const finalAmount = amount || (parseFloat(quantity) * parseFloat(unit_price) * (1 - (parseFloat(discount) || 0) / 100));

    const updateSalesSQL = `UPDATE sales SET employee = ?, product = ?, quantity = ?, unit_price = ?, discount = ?, amount = ?, customer = ?, note = ?, date = ? WHERE id = ?`;

    db.run(updateSalesSQL, [employee, product, quantity, unit_price, discount || 0, finalAmount, customer, note, date, saleId], (err) => {
        if (err) {
            console.error("Lỗi sửa sales:", err.message);
            return res.redirect('/sales');
        }

        db.get(`SELECT invoice_code FROM sales WHERE id = ?`, [saleId], (err2, row) => {
            const invoiceCode = row ? row.invoice_code : null;

            if (invoiceCode) {
                db.get(`SELECT SUM(amount) as totalOrderAmount FROM sales WHERE invoice_code = ?`, [invoiceCode], (err3, sumRow) => {
                    const newTotal = sumRow ? sumRow.totalOrderAmount : finalAmount;
                    const revNote = `Đơn hàng (${invoiceCode}) từ NV: ${employee}, KH: ${customer || 'Khách lẻ'}`;

                    db.run(`UPDATE revenues SET amount = ?, note = ?, date = ? WHERE sale_id = ?`,
                        [newTotal, revNote, date, invoiceCode], () => {
                            res.redirect('/sales');
                        });
                });
            } else {
                res.redirect('/sales');
            }
        });
    });
});

app.get('/sales/export', checkAuth, (req, res) => {
    let { from, to } = req.query;
    let clause = '';
    let params = [];

    if (from && to) {
        clause = `WHERE date BETWEEN ? AND ?`;
        params = [from, to];
    }

    const sql = `SELECT id, employee AS 'Nhân viên', product AS 'Sản phẩm', quantity AS 'Số lượng', unit_price AS 'Đơn giá', discount AS 'Chiết khấu (%)', amount AS 'Thành tiền', customer AS 'Khách hàng', note AS 'Ghi chú', date AS 'Ngày' FROM sales ${clause} ORDER BY date DESC`;

    db.all(sql, params, (err, rows) => {
        if (err) {
            console.error("Lỗi xuất Excel:", err.message);
            return res.status(500).send("Lỗi xuất file Excel");
        }

        const worksheet = XLSX.utils.json_to_sheet(rows);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, "DanhSachDonHang");

        const buffer = XLSX.write(workbook, { bookType: 'xlsx', type: 'buffer' });

        res.setHeader('Content-Disposition', 'attachment; filename="Danh_Sach_Don_Hang.xlsx"');
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
    });
});

// ==================== ROUTE BÁO GIÁ (PRINT-QUOTE) ====================
app.post('/quote/print', (req, res) => {
    const { employee, customerName, customerTaxCode, customerPhone, quoteDate, validUntil, note, vatRate, products, units, quantities, unitPrices, amounts } = req.body;

    // 1. Gom nhóm danh sách sản phẩm từ form
    let itemsList = [];
    if (products) {
        const prods = Array.isArray(products) ? products : [products];
        const unts = Array.isArray(units) ? units : [units];
        const qtys = Array.isArray(quantities) ? quantities : [quantities];
        const prices = Array.isArray(unitPrices) ? unitPrices : [unitPrices];
        const amts = Array.isArray(amounts) ? amounts : [amounts];

        prods.forEach((p, idx) => {
            if (p && p.trim() !== '') {
                itemsList.push({
                    product: p,
                    unit: unts[idx] || 'Bộ',
                    quantity: Number(qtys[idx] || 1),
                    unit_price: Number(prices[idx] || 0),
                    amount: Number(amts[idx] || 0)
                });
            }
        });
    }

    // 2. Định dạng ngày tháng báo giá và ngày hiệu lực
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, '0');
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const yyyy = today.getFullYear();
    const dateStr = `${dd}${mm}${yyyy}`;

    const formattedDate = quoteDate ? quoteDate.split('-').reverse().join('/') : `${dd}/${mm}/${yyyy}`;
    const formattedValidUntil = validUntil ? validUntil.split('-').reverse().join('/') : '';

    // 3. Truy vấn Database để tự động tăng số thứ tự mã báo giá trong ngày (#MTBG-DDMMYYYY01, 02,...)
    const codePrefix = `#MTBG-${dateStr}`;
    db.get(`SELECT code FROM quotes WHERE code LIKE ? ORDER BY id DESC LIMIT 1`, [`${codePrefix}%`], (err, row) => {
        let runningNum = 1;
        if (row && row.code) {
            const lastNum = parseInt(row.code.slice(-2), 10);
            if (!isNaN(lastNum)) {
                runningNum = lastNum + 1;
            }
        }
        const quoteCode = `${codePrefix}${String(runningNum).padStart(2, '0')}`;

        const quoteData = {
            code: quoteCode,
            employee: employee || 'Hoài Nam',
            customer: customerName || 'Khách lẻ',
            taxCode: customerTaxCode || '',
            phone: customerPhone || '',
            date: formattedDate,
            validUntil: formattedValidUntil,
            note: note || '',
            vatRate: vatRate !== undefined ? Number(vatRate) : 8
        };

        // Lưu vào cơ sở dữ liệu (nếu bảng quotes đã có) để tracking mã tăng dần
        db.run(`INSERT INTO quotes (code, customer, employee, date, total) VALUES (?, ?, ?, ?, ?)`,
            [quoteCode, quoteData.customer, quoteData.employee, quoteData.date, itemsList.reduce((s, i) => s + i.amount, 0)],
            () => {
                res.render('print-quote', {
                    quote: quoteData,
                    items: itemsList
                });
            }
        );
    });
});

// ==================== ROUTE PHIẾU XUẤT KHO (PRINT-TEMPLATE) ====================
app.get('/sales/invoice/:invoice_code', checkAuth, (req, res) => {
    const code = req.params.invoice_code;
    const staffName = (req.session.user && (req.session.user.fullname || req.session.user.username))
        ? (req.session.user.fullname || req.session.user.username)
        : 'Nguyễn Quang Khang';

    db.all(`SELECT * FROM sales WHERE invoice_code = ? OR id = ?`, [code, code], (err, rows) => {
        if (err || !rows || rows.length === 0) {
            return res.status(404).send("Không tìm thấy dữ liệu phiếu xuất kho!");
        }

        const firstItem = rows[0] || {};
        // Tạo mã xuất kho chuẩn format #MTXK-DDMMYYYY0n
        const xkCode = firstItem.invoice_code && firstItem.invoice_code.startsWith('#MTXK')
            ? firstItem.invoice_code
            : generateCode('MTXK', firstItem.id || 1);

        res.render('print-template', {
            items: rows,
            invoiceCode: xkCode,
            type: 'invoice',
            staffName: staffName
        });
    });
});

// ==================== QUẢN LÝ NHÂN VIÊN ====================
app.get('/employees', checkAuth, (req, res) => {
    db.all(`SELECT * FROM employees ORDER BY id DESC`, (err, rows) => {
        res.render('employees', { employees: rows || [] });
    });
});

app.post('/employees/add', checkAuth, (req, res) => {
    const { name, phone, position } = req.body;
    db.run(`INSERT INTO employees (name, phone, position) VALUES (?, ?, ?)`, [name, phone, position], () => {
        res.redirect('/sales');
    });
});

app.post('/employees/edit/:id', checkAuth, (req, res) => {
    const { name, phone, position } = req.body;
    db.run(`UPDATE employees SET name = ?, phone = ?, position = ? WHERE id = ?`, [name, phone, position, req.params.id], () => {
        res.redirect('/sales');
    });
});

app.post('/employees/delete/:id', checkAuth, (req, res) => {
    db.run(`DELETE FROM employees WHERE id = ?`, [req.params.id], () => {
        res.redirect('/sales');
    });
});

// ==================== 2. QUẢN LÝ DOANH THU ====================
app.get('/revenues', checkAuth, (req, res) => {
    const { from, to } = req.query;
    const { clause, params } = getDateFilterSQL(from, to);
    db.all(`SELECT * FROM revenues ${clause} ORDER BY date DESC, id DESC`, params, (err, rows) => {
        res.render('revenues', { revenues: rows || [], from: from || '', to: to || '' });
    });
});

app.post('/revenues/add', checkAuth, (req, res) => {
    const { amount, category, note, date } = req.body;
    db.run(`INSERT INTO revenues (amount, category, note, date) VALUES (?, ?, ?, ?)`, [amount, category, note, date], () => res.redirect('/revenues'));
});

app.post('/revenues/delete/:id', checkAuth, checkRole(['superadmin', 'manager']), (req, res) => {
    const revId = req.params.id;
    db.get(`SELECT sale_id FROM revenues WHERE id = ?`, [revId], (err, row) => {
        if (row && row.sale_id) {
            return res.redirect('/revenues');
        }
        db.run(`DELETE FROM revenues WHERE id = ?`, [revId], () => res.redirect('/revenues'));
    });
});

app.post('/revenues/edit/:id', checkAuth, checkRole(['superadmin', 'manager']), (req, res) => {
    const revId = req.params.id;
    const { amount, category, note, date } = req.body;

    db.get(`SELECT sale_id FROM revenues WHERE id = ?`, [revId], (err, row) => {
        if (row && row.sale_id) {
            return res.redirect('/revenues');
        }
        db.run(`UPDATE revenues SET amount = ?, category = ?, note = ?, date = ? WHERE id = ?`,
            [amount, category, note, date, revId], () => {
                res.redirect('/revenues');
            });
    });
});

app.get('/revenues/export', checkAuth, async (req, res) => {
    try {
        const { from, to } = req.query;
        const { clause, params } = getDateFilterSQL(from, to);
        const query = `SELECT * FROM revenues ${clause} ORDER BY date DESC`;

        db.all(query, params, async (err, revenues) => {
            if (err) {
                console.error(err);
                return res.status(500).send("Lỗi lấy dữ liệu xuất Excel");
            }

            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Doanh Thu');

            worksheet.columns = [
                { header: 'Ngày', key: 'date', width: 15 },
                { header: 'Nguồn Thu / Phân Loại', key: 'category', width: 30 },
                { header: 'Mã Đơn Hàng (nếu có)', key: 'sale_id', width: 25 },
                { header: 'Số Tiền (VNĐ)', key: 'amount', width: 20 },
                { header: 'Ghi Chú', key: 'note', width: 30 }
            ];

            revenues.forEach(item => {
                worksheet.addRow({
                    date: item.date,
                    category: item.category,
                    sale_id: item.sale_id || 'Thủ công',
                    amount: item.amount,
                    note: item.note || ''
                });
            });

            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', `attachment; filename=Bao-Cao-Doanh-Thu-${from || 'tat-ca'}_den_${to || 'nay'}.xlsx`);

            await workbook.xlsx.write(res);
            res.end();
        });
    } catch (error) {
        console.error(error);
        res.status(500).send("Lỗi xuất file Excel");
    }
});

// ==================== 3. QUẢN LÝ CHI PHÍ ====================
app.get('/expenses', checkAuth, (req, res) => {
    const { from, to } = req.query;
    const { clause, params } = getDateFilterSQL(from, to);
    db.all(`SELECT * FROM expenses ${clause} ORDER BY date DESC, id DESC`, params, (err, rows) => {
        res.render('expenses', { expenses: rows || [], from: from || '', to: to || '' });
    });
});

app.post('/expenses/add', checkAuth, checkRole(['superadmin', 'manager']), (req, res) => {
    const { amount, category, note, date } = req.body;
    db.run(`INSERT INTO expenses (amount, category, note, date) VALUES (?, ?, ?, ?)`, [amount, category, note, date], () => res.redirect('/expenses'));
});

app.post('/expenses/delete/:id', checkAuth, checkRole(['superadmin', 'manager']), (req, res) => {
    db.run(`DELETE FROM expenses WHERE id = ?`, [req.params.id], () => res.redirect('/expenses'));
});

app.post('/expenses/edit/:id', checkAuth, checkRole(['superadmin', 'manager']), (req, res) => {
    const { amount, category, note, date } = req.body;
    db.run(`UPDATE expenses SET amount = ?, category = ?, note = ?, date = ? WHERE id = ?`,
        [amount, category, note, date, req.params.id], () => {
            res.redirect('/expenses');
        });
});

app.get('/expenses/export', checkAuth, async (req, res) => {
    try {
        const { from, to } = req.query;
        const { clause, params } = getDateFilterSQL(from, to);
        const query = `SELECT * FROM expenses ${clause} ORDER BY date DESC`;

        db.all(query, params, async (err, expenses) => {
            if (err) {
                console.error(err);
                return res.status(500).send("Lỗi lấy dữ liệu xuất Excel chi phí");
            }

            const workbook = new ExcelJS.Workbook();
            const worksheet = workbook.addWorksheet('Chi Phí');

            worksheet.columns = [
                { header: 'Ngày', key: 'date', width: 15 },
                { header: 'Khoản Chi', key: 'category', width: 30 },
                { header: 'Số Tiền (VNĐ)', key: 'amount', width: 20 },
                { header: 'Ghi Chú', key: 'note', width: 30 }
            ];

            expenses.forEach(item => {
                worksheet.addRow({
                    date: item.date,
                    category: item.category,
                    amount: item.amount,
                    note: item.note || ''
                });
            });

            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
            res.setHeader('Content-Disposition', `attachment; filename=Bao-Cao-Chi-Phi-${from || 'tat-ca'}_den_${to || 'nay'}.xlsx`);

            await workbook.xlsx.write(res);
            res.end();
        });
    } catch (error) {
        console.error(error);
        res.status(500).send("Lỗi xuất file Excel chi phí");
    }
});

// ==================== 4. BÁO CÁO LỢI NHUẬN ====================
app.get('/profit', checkAuth, (req, res) => {
    const { from, to } = req.query;
    let revFilter = getDateFilterSQL(from, to);
    let expFilter = getDateFilterSQL(from, to);

    db.get(`SELECT SUM(amount) as t FROM revenues ${revFilter.clause}`, revFilter.params, (e1, rR) => {
        db.get(`SELECT SUM(amount) as t FROM expenses ${expFilter.clause}`, expFilter.params, (e2, rE) => {
            const rev = rR?.t || 0;
            const exp = rE?.t || 0;
            res.render('profit', {
                totalRevenue: rev,
                totalExpense: exp,
                netProfit: rev - exp,
                from: from || '',
                to: to || ''
            });
        });
    });
});

app.get('/profit/export', checkAuth, (req, res) => {
    const { from, to } = req.query;
    let revFilter = getDateFilterSQL(from, to);
    let expFilter = getDateFilterSQL(from, to);

    db.get(`SELECT SUM(amount) as t FROM revenues ${revFilter.clause}`, revFilter.params, (e1, rR) => {
        db.get(`SELECT SUM(amount) as t FROM expenses ${expFilter.clause}`, expFilter.params, (e2, rE) => {
            const rev = rR?.t || 0;
            const exp = rE?.t || 0;
            const profit = rev - exp;
            const margin = rev > 0 ? ((profit / rev) * 100).toFixed(1) : 0;

            let csv = '\uFEFFBáo Cáo Lợi Nhuận Tổng Hợp\n';
            csv += `Khoảng thời gian:,${from || 'Tất cả'} đến ${to || 'Tất cả'}\n\n`;
            csv += `Chỉ số,Số tiền (VNĐ),Tỷ lệ\n`;
            csv += `Tổng Doanh Thu,${rev},100%\n`;
            csv += `Tổng Chi Phí,${exp},-\n`;
            csv += `Lợi Nhuận Thực Tế,${profit},${margin}%\n`;

            res.setHeader('Content-Type', 'text/csv; charset=utf-8');
            res.setHeader('Content-Disposition', 'attachment; filename=Bao_Cao_Loi_Nhuan.csv');
            res.status(200).send(csv);
        });
    });
});

app.listen(PORT, () => console.log(`Server đang chạy tại: http://localhost:${PORT}`));