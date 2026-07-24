const express = require('express');
const mongoose = require('mongoose');
const session = require('express-session');
const bcrypt = require('bcryptjs');
const ExcelJS = require('exceljs');
const multer = require('multer');
const path = require('path');

if (process.env.NODE_ENV !== 'production') {
    try { require('dotenv').config(); } catch (e) {}
}

const User = require('./models/User');
const Sparepart = require('./models/Sparepart');
const Repair = require('./models/Repair');

const app = express();

// WAJIB UNTUK VERCEL & PROXY
app.set('trust proxy', 1);

// Naikkan limit payload menjadi 50MB untuk mengantisipasi file import besar (20MB+)
app.use(express.urlencoded({ limit: '50mb', extended: true }));
app.use(express.json({ limit: '50mb' }));

const upload = multer({ 
    storage: multer.memoryStorage(),
    limits: { fileSize: 50 * 1024 * 1024 } // Batasan file 50MB di Multer
});

// Pengaturan Path Absolut Views
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');

// --- KONEKSI MONGODB ---
const mongoUri = process.env.MONGODB_URI;

async function connectDB() {
    if (mongoose.connection.readyState >= 1) return;
    try {
        await mongoose.connect(mongoUri, { serverSelectionTimeoutMS: 5000 });
        console.log('MongoDB Connected Successfully');
    } catch (err) {
        console.error('MongoDB Connection Error:', err);
        throw err;
    }
}

// Konfigurasi Session
app.use(session({
    secret: process.env.SESSION_SECRET || 'rahasia_super_aman_123',
    resave: false,
    saveUninitialized: false,
    cookie: { 
        maxAge: 12 * 60 * 60 * 1000,
        secure: process.env.NODE_ENV === 'production',
        httpOnly: true,
        sameSite: 'lax'
    }
}));

// Middleware Cek Koneksi DB
app.use(async (req, res, next) => {
    try {
        await connectDB();
        next();
    } catch (err) {
        res.status(500).send('Koneksi database gagal: ' + err.message);
    }
});

// Middleware Auth
const isAuthenticated = (req, res, next) => {
    if (req.session && req.session.user) return next();
    res.redirect('/login?error=Session+expired.+Please+login+again.');
};

const isAdmin = (req, res, next) => {
    if (req.session && req.session.user && req.session.user.role === 'admin') return next();
    if (!req.session || !req.session.user) return res.redirect('/login');
    res.status(403).send('Akses Ditolak: Khusus Admin');
};

// --- ROUTES AUTH ---
app.get('/', (req, res) => res.redirect('/login'));

app.get('/login', (req, res) => {
    if (req.session && req.session.user) {
        if (req.session.user.role === 'admin') return res.redirect('/admin/sparepart');
        return res.redirect('/teknisi/dashboard');
    }
    res.render('login', { error: req.query.error || null });
});

app.post('/login', async (req, res) => {
    const { username, password } = req.body;
    try {
        const user = await User.findOne({ username });
        if (user && await bcrypt.compare(password, user.password)) {
            req.session.user = user;
            req.session.save(() => {
                if (user.role === 'admin') return res.redirect('/admin/sparepart');
                return res.redirect('/teknisi/dashboard');
            });
            return;
        }
        res.render('login', { error: 'Username atau Password salah!' });
    } catch (err) {
        res.render('login', { error: 'Terjadi kesalahan sistem: ' + err.message });
    }
});

app.get('/logout', (req, res) => {
    req.session.destroy(() => {
        res.clearCookie('connect.sid');
        res.redirect('/login');
    });
});

// --- ADMIN: SPAREPART ---
app.get('/admin/sparepart', isAdmin, async (req, res) => {
    try {
        const { search } = req.query;
        let query = {};
        if (search) {
            const regex = new RegExp(search, 'i');
            query = { $or: [{ namaBarang: regex }, { kodeBarang: regex }, { lokasiBarang: regex }, { typeBarang: regex }] };
        }
        const spareparts = await Sparepart.find(query);
        res.render('admin-sparepart', { user: req.session.user, spareparts, search: search || '', error: null });
    } catch (err) {
        res.status(500).render('admin-sparepart', { user: req.session.user, spareparts: [], search: '', error: err.message });
    }
});

app.post('/admin/sparepart/save', isAdmin, async (req, res) => {
    try {
        const { id, namaBarang, kodeBarang, lokasiBarang, typeBarang, stokBarang, fungsiBarang } = req.body;
        if (id) {
            await Sparepart.findByIdAndUpdate(id, { namaBarang, kodeBarang, lokasiBarang, typeBarang, stokBarang, fungsiBarang });
        } else {
            await Sparepart.create({ namaBarang, kodeBarang, lokasiBarang, typeBarang, stokBarang, barangKeluar: 0, barangMasuk: Number(stokBarang) || 0, fungsiBarang });
        }
        res.redirect('/admin/sparepart');
    } catch (err) {
        res.redirect('/admin/sparepart');
    }
});

app.post('/admin/sparepart/delete/:id', isAdmin, async (req, res) => {
    try {
        await Sparepart.findByIdAndDelete(req.params.id);
        res.redirect('/admin/sparepart');
    } catch (err) {
        res.redirect('/admin/sparepart');
    }
});

app.post('/admin/sparepart/update-stok', isAdmin, async (req, res) => {
    try {
        const { sparepartId, jenis, jumlah } = req.body;
        const qty = parseInt(jumlah) || 0;
        const sp = await Sparepart.findById(sparepartId);
        if (sp) {
            if (jenis === 'tambah') {
                sp.stokBarang += qty;
                sp.barangMasuk += qty;
            } else if (jenis === 'kurang' && sp.stokBarang >= qty) {
                sp.stokBarang -= qty;
                sp.barangKeluar += qty;
            }
            await sp.save();
        }
        res.redirect('/admin/sparepart');
    } catch (err) {
        res.redirect('/admin/sparepart');
    }
});

// EXPORT EXCEL SPAREPART
app.get('/admin/sparepart/export', isAdmin, async (req, res) => {
    try {
        const spareparts = await Sparepart.find();
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Data Sparepart');

        worksheet.columns = [
            { header: 'Kode Barang', key: 'kodeBarang', width: 15 },
            { header: 'Nama Barang', key: 'namaBarang', width: 25 },
            { header: 'Lokasi', key: 'lokasiBarang', width: 15 },
            { header: 'Type', key: 'typeBarang', width: 15 },
            { header: 'Stok', key: 'stokBarang', width: 10 },
            { header: 'Barang Masuk', key: 'barangMasuk', width: 15 },
            { header: 'Barang Keluar', key: 'barangKeluar', width: 15 },
            { header: 'Fungsi Barang', key: 'fungsiBarang', width: 25 }
        ];

        spareparts.forEach(s => worksheet.addRow(s));

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Data_Sparepart.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        res.status(500).send(err.message);
    }
});

// IMPORT EXCEL SPAREPART (Mendukung file besar 20MB+)
app.post('/admin/sparepart/import', isAdmin, upload.single('fileExcel'), async (req, res) => {
    try {
        if (!req.file) return res.redirect('/admin/sparepart');

        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(req.file.buffer);
        const worksheet = workbook.getWorksheet(1);

        const bulkOperations = [];
        worksheet.eachRow((row, rowNumber) => {
            if (rowNumber === 1) return;

            const kodeBarang = row.getCell(1).text ? row.getCell(1).text.trim() : '';
            const namaBarang = row.getCell(2).text ? row.getCell(2).text.trim() : '';
            const lokasiBarang = row.getCell(3).text ? row.getCell(3).text.trim() : '';
            const typeBarang = row.getCell(4).text ? row.getCell(4).text.trim() : '';
            const stokBarang = Number(row.getCell(5).value) || 0;
            const barangMasuk = Number(row.getCell(6).value) || stokBarang;
            const barangKeluar = Number(row.getCell(7).value) || 0;
            const fungsiBarang = row.getCell(8).text ? row.getCell(8).text.trim() : '';

            if (kodeBarang && namaBarang) {
                bulkOperations.push({
                    updateOne: {
                        filter: { kodeBarang },
                        update: { $set: { namaBarang, lokasiBarang, typeBarang, stokBarang, barangMasuk, barangKeluar, fungsiBarang } },
                        upsert: true
                    }
                });
            }
        });

        if (bulkOperations.length > 0) {
            await Sparepart.bulkWrite(bulkOperations);
        }

        res.redirect('/admin/sparepart');
    } catch (err) {
        res.redirect('/admin/sparepart');
    }
});

app.get('/admin/sparepart/repair-history/:id', isAdmin, async (req, res) => {
    try {
        const sparepartId = req.params.id;
        const repairs = await Repair.find({ 'sparepartsUsed.sparepart': sparepartId }).sort({ tanggal: -1 });

        const formattedData = repairs.map(r => {
            const usedItem = r.sparepartsUsed.find(item => item.sparepart && item.sparepart.toString() === sparepartId);
            return {
                tanggal: r.tanggal,
                shift: r.shift,
                namaMesin: r.namaMesin,
                problem: r.problem,
                analisaPenyebab: r.analisaPenyebab,
                caraPerbaikan: r.caraPerbaikan,
                jumlahPakai: usedItem ? usedItem.jumlahPakai : 1,
                pic: r.pic,
                jamMulai: r.jamMulai,
                jamSelesai: r.jamSelesai,
                status: r.status,
                keterangan: r.keterangan
            };
        });
        res.json(formattedData);
    } catch (err) {
        res.status(500).json({ error: "Gagal mengambil data" });
    }
});

// --- ADMIN: REPAIR LOGS & SORTIR ---
app.get('/admin/repair', isAdmin, async (req, res) => {
    try {
        let { page, filterType, month, weekStart, weekEnd, tanggalFilter, year, namaMesin } = req.query;
        page = parseInt(page) || 1;
        const limit = 20;
        const skip = (page - 1) * limit;

        let query = {};
        if (namaMesin) query.namaMesin = new RegExp(namaMesin, 'i');

        if (filterType === 'date' && tanggalFilter) {
            query.tanggal = { $gte: new Date(tanggalFilter), $lte: new Date(tanggalFilter + 'T23:59:59') };
        } else if (filterType === 'month' && month) {
            const targetMonth = new Date(month);
            query.tanggal = { 
                $gte: new Date(targetMonth.getFullYear(), targetMonth.getMonth(), 1), 
                $lte: new Date(targetMonth.getFullYear(), targetMonth.getMonth() + 1, 0, 23, 59, 59) 
            };
        } else if (filterType === 'year' && year) {
            query.tanggal = { $gte: new Date(`${year}-01-01T00:00:00`), $lte: new Date(`${year}-12-31T23:59:59`) };
        } else if (filterType === 'week' && weekStart && weekEnd) {
            query.tanggal = { $gte: new Date(weekStart), $lte: new Date(weekEnd + 'T23:59:59') };
        }

        const totalRepairs = await Repair.countDocuments(query);
        const repairs = await Repair.find(query).populate('sparepartsUsed.sparepart').sort({ tanggal: -1 }).skip(skip).limit(limit);

        res.render('admin-repair', {
            user: req.session.user,
            repairs,
            currentPage: page,
            totalPages: Math.ceil(totalRepairs / limit) || 1,
            filterType: filterType || 'all',
            month: month || '',
            weekStart: weekStart || '',
            weekEnd: weekEnd || '',
            tanggalFilter: tanggalFilter || '',
            year: year || '',
            namaMesin: namaMesin || ''
        });
    } catch (err) {
        res.status(500).send(err.message);
    }
});

app.post('/admin/repair/status/:id', isAdmin, async (req, res) => {
    try {
        await Repair.findByIdAndUpdate(req.params.id, { status: req.body.status });
        res.redirect('/admin/repair');
    } catch (err) {
        res.redirect('/admin/repair');
    }
});

app.get('/admin/repair/export', isAdmin, async (req, res) => {
    try {
        const { startDate, endDate, namaMesin } = req.query;
        let query = {};
        if (startDate && endDate) query.tanggal = { $gte: new Date(startDate), $lte: new Date(endDate + 'T23:59:59') };
        if (namaMesin) query.namaMesin = new RegExp(namaMesin, 'i');

        const repairs = await Repair.find(query).populate('sparepartsUsed.sparepart');
        const workbook = new ExcelJS.Workbook();
        const worksheet = workbook.addWorksheet('Laporan Perbaikan');

        worksheet.columns = [
            { header: 'Tanggal', key: 'tanggal', width: 15 },
            { header: 'Shift', key: 'shift', width: 10 },
            { header: 'Nama Mesin', key: 'namaMesin', width: 20 },
            { header: 'Problem', key: 'problem', width: 25 },
            { header: 'Analisa Penyebab', key: 'analisaPenyebab', width: 25 },
            { header: 'Cara Perbaikan', key: 'caraPerbaikan', width: 25 },
            { header: 'Sparepart Dipakai', key: 'sparepart', width: 25 },
            { header: 'Total QTY', key: 'jumlahPakai', width: 10 },
            { header: 'PIC', key: 'pic', width: 15 },
            { header: 'Jam', key: 'jam', width: 20 },
            { header: 'Status', key: 'status', width: 10 },
            { header: 'Keterangan', key: 'keterangan', width: 20 }
        ];

        repairs.forEach(r => {
            let listSp = r.sparepartsUsed.map(item => `${item.sparepart ? item.sparepart.namaBarang : '-'} (${item.jumlahPakai}x)`).join(', ');
            let totalQty = r.sparepartsUsed.reduce((sum, item) => sum + item.jumlahPakai, 0);

            worksheet.addRow({
                tanggal: r.tanggal ? r.tanggal.toISOString().split('T')[0] : '',
                shift: r.shift,
                namaMesin: r.namaMesin,
                problem: r.problem,
                analisaPenyebab: r.analisaPenyebab,
                caraPerbaikan: r.caraPerbaikan,
                sparepart: listSp || '-',
                jumlahPakai: totalQty,
                pic: r.pic,
                jam: `${r.jamMulai} - ${r.jamSelesai}`,
                status: r.status,
                keterangan: r.keterangan
            });
        });

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader('Content-Disposition', 'attachment; filename=Laporan_Perbaikan.xlsx');
        await workbook.xlsx.write(res);
        res.end();
    } catch (err) {
        res.status(500).send(err.message);
    }
});

// --- ADMIN: USERS ---
app.get('/admin/users', isAdmin, async (req, res) => {
    try {
        const users = await User.find();
        res.render('admin-users', { user: req.session.user, users, error: req.query.error || null });
    } catch (err) {
        res.status(500).send(err.message);
    }
});

app.post('/admin/user/add', isAdmin, async (req, res) => {
    try {
        const { username, password, role } = req.body;
        const hashedPassword = await bcrypt.hash(password, 10);
        await User.create({ username, password: hashedPassword, role });
        res.redirect('/admin/users');
    } catch (err) {
        res.redirect('/admin/users?error=Gagal membuat user');
    }
});

app.post('/admin/user/delete/:id', isAdmin, async (req, res) => {
    try {
        await User.findByIdAndDelete(req.params.id);
        res.redirect('/admin/users');
    } catch (err) {
        res.redirect('/admin/users');
    }
});

// --- TEKNISI DASHBOARD & REPAIR ---
app.get('/teknisi/dashboard', isAuthenticated, async (req, res) => {
    try {
        let { page } = req.query;
        page = parseInt(page) || 1;
        const limit = 20;
        const skip = (page - 1) * limit;

        const totalRepairs = await Repair.countDocuments();
        const repairs = await Repair.find().populate('sparepartsUsed.sparepart').sort({ tanggal: -1 }).skip(skip).limit(limit);
        const spareparts = await Sparepart.find();

        res.render('teknisi-dashboard', {
            user: req.session.user,
            spareparts,
            repairs,
            currentPage: page,
            totalPages: Math.ceil(totalRepairs / limit) || 1
        });
    } catch (err) {
        res.status(500).send(err.message);
    }
});
// --- TEKNISI: SPAREPART (Pencarian Sparepart) ---
app.get('/teknisi/sparepart', isAuthenticated, async (req, res) => {
    try {
        const { search } = req.query;
        let query = {};
        if (search) {
            const regex = new RegExp(search, 'i');
            query = { 
                $or: [
                    { namaBarang: regex }, 
                    { kodeBarang: regex }, 
                    { lokasiBarang: regex }, 
                    { typeBarang: regex },
                    { fungsiBarang: regex }
                ] 
            };
        }
        const spareparts = await Sparepart.find(query);
        res.render('teknisi-sparepart', { 
            user: req.session.user, 
            spareparts, 
            search: search || '', 
            error: null 
        });
    } catch (err) {
        res.status(500).render('teknisi-sparepart', { 
            user: req.session.user, 
            spareparts: [], 
            search: '', 
            error: err.message 
        });
    }
});

app.post('/repair/add', isAuthenticated, async (req, res) => {
    try {
        const { tanggal, shift, namaMesin, problem, analisaPenyebab, caraPerbaikan, sparepartIds, jumlahPakais, pic, jamMulai, jamSelesai, status, keterangan } = req.body;
        
        let sparepartsUsed = [];
        if (sparepartIds) {
            const ids = Array.isArray(sparepartIds) ? sparepartIds : [sparepartIds];
            const qtys = Array.isArray(jumlahPakais) ? jumlahPakais : [jumlahPakais];

            for (let i = 0; i < ids.length; i++) {
                if (ids[i] && qtys[i] > 0) {
                    const spId = ids[i];
                    const qty = Number(qtys[i]);
                    const sp = await Sparepart.findById(spId);
                    if (sp) {
                        sp.stokBarang -= qty;
                        sp.barangKeluar += qty;
                        await sp.save();
                    }
                    sparepartsUsed.push({ sparepart: spId, jumlahPakai: qty });
                }
            }
        }

        await Repair.create({
            tanggal, shift, namaMesin, problem, analisaPenyebab, caraPerbaikan,
            sparepartsUsed, pic, jamMulai, jamSelesai, status, keterangan
        });

        const redirectUrl = req.session.user.role === 'admin' ? '/admin/repair' : '/teknisi/dashboard';
        res.redirect(redirectUrl);
    } catch (err) {
        res.status(500).send('Gagal menyimpan: ' + err.message);
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));

module.exports = app;
