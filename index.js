const express = require('express');
const mongoose = require('mongoose');
const session = require('express-session');
const bcrypt = require('bcryptjs');
const ExcelJS = require('exceljs');
const path = require('path');

if (process.env.NODE_ENV !== 'production') {
    try { require('dotenv').config(); } catch (e) {}
}

const User = require('./models/User');
const Sparepart = require('./models/Sparepart');
const Repair = require('./models/Repair');

const app = express();

app.use(express.urlencoded({ extended: true }));
app.use(express.json());

// Pengaturan Path Absolut Views untuk Vercel
app.set('views', path.join(__dirname, 'views'));
app.set('view engine', 'ejs');

app.use(session({
    secret: process.env.SESSION_SECRET || 'rahasia_super_aman_123',
    resave: false,
    saveUninitialized: false
}));

// --- OPTIMASI KONEKSI MONGODB UNTUK SERVERLESS (VERCEL) ---
const mongoUri = process.env.MONGODB_URI;

async function connectDB() {
    if (mongoose.connection.readyState >= 1) {
        return;
    }
    try {
        await mongoose.connect(mongoUri, {
            serverSelectionTimeoutMS: 5000,
        });
        console.log('MongoDB Connected Successfully');
    } catch (err) {
        console.error('MongoDB Connection Error:', err);
        throw err;
    }
}

// Middleware agar koneksi database selalu dicek di setiap request
app.use(async (req, res, next) => {
    try {
        await connectDB();
        next();
    } catch (err) {
        res.status(500).send('Koneksi database gagal: ' + err.message);
    }
});

const isAuthenticated = (req, res, next) => {
    if (req.session.user) return next();
    res.redirect('/login');
};

const isAdmin = (req, res, next) => {
    if (req.session.user && req.session.user.role === 'admin') return next();
    res.status(403).send('Akses Ditolak: Khusus Admin');
};

// --- ROUTES AUTHENTICATION ---
app.get('/', (req, res) => res.redirect('/login'));
app.get('/login', (req, res) => res.render('login', { error: null }));

app.post('/login', async (req, res) => {
    const { username, password } = req.body;
    try {
        const user = await User.findOne({ username });
        if (user && await bcrypt.compare(password, user.password)) {
            req.session.user = user;
            if (user.role === 'admin') return res.redirect('/admin/sparepart');
            return res.redirect('/teknisi/dashboard');
        }
        res.render('login', { error: 'Username atau Password salah!' });
    } catch (err) {
        res.render('login', { error: 'Terjadi kesalahan sistem: ' + err.message });
    }
});

app.get('/logout', (req, res) => req.session.destroy(() => res.redirect('/login')));

// --- ADMIN: HALAMAN SPAREPART & STOK ---
app.get('/admin/sparepart', isAdmin, async (req, res) => {
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
        res.render('admin-sparepart', { user: req.session.user, spareparts, search: search || '' });
    } catch (err) {
        res.status(500).send(err.message);
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
        res.status(500).send('Gagal menyimpan sparepart: ' + err.message);
    }
});
// ROUTE HAPUS SPAREPART
app.post('/admin/sparepart/delete/:id', isAdmin, async (req, res) => {
    try {
        const { id } = req.params;
        await Sparepart.findByIdAndDelete(id);
        res.redirect('/admin/sparepart');
    } catch (err) {
        res.status(500).send('Gagal menghapus sparepart: ' + err.message);
    }
});

// ROUTE TAMBAH / KURANG QTY STOK SPAREPART
app.post('/admin/sparepart/update-stok', isAdmin, async (req, res) => {
    try {
        const { sparepartId, jenis, jumlah } = req.body;
        const qty = parseInt(jumlah) || 0;

        const sp = await Sparepart.findById(sparepartId);
        if (!sp) {
            return res.status(404).send('Sparepart tidak ditemukan');
        }

        if (jenis === 'tambah') {
            sp.stokBarang += qty;
            sp.barangMasuk += qty;
        } else if (jenis === 'kurang') {
            if (sp.stokBarang < qty) {
                return res.status(400).send('Stok tidak mencukupi untuk dikurangi!');
            }
            sp.stokBarang -= qty;
            sp.barangKeluar += qty;
        }

        await sp.save();
        res.redirect('/admin/sparepart');
    } catch (err) {
        res.status(500).send('Gagal mengupdate stok: ' + err.message);
    }
});

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

// --- ADMIN: HALAMAN REPAIR ---
app.get('/admin/repair', isAdmin, async (req, res) => {
    try {
        let { page, filterType, month, weekStart, weekEnd } = req.query;
        page = parseInt(page) || 1;
        const limit = 20;
        const skip = (page - 1) * limit;

        let query = {};
        const now = new Date();

        if (filterType === 'month') {
            const targetMonth = month ? new Date(month) : new Date(now.getFullYear(), now.getMonth(), 1);
            const startMonth = new Date(targetMonth.getFullYear(), targetMonth.getMonth(), 1);
            const endMonth = new Date(targetMonth.getFullYear(), targetMonth.getMonth() + 1, 0, 23, 59, 59);
            query.tanggal = { $gte: startMonth, $lte: endMonth };
        } else if (filterType === 'week' && weekStart && weekEnd) {
            query.tanggal = { $gte: new Date(weekStart), $lte: new Date(weekEnd + 'T23:59:59') };
        }

        const totalRepairs = await Repair.countDocuments(query);
        const repairs = await Repair.find(query)
            .populate('sparepartsUsed.sparepart')
            .sort({ tanggal: -1 })
            .skip(skip)
            .limit(limit);

        res.render('admin-repair', {
            user: req.session.user,
            repairs,
            currentPage: page,
            totalPages: Math.ceil(totalRepairs / limit) || 1,
            filterType: filterType || 'all',
            month: month || '',
            weekStart: weekStart || '',
            weekEnd: weekEnd || ''
        });
    } catch (err) {
        res.status(500).send(err.message);
    }
});

app.post('/admin/user/add', isAdmin, async (req, res) => {
    try {
        const { username, password, role } = req.body;
        const hashedPassword = await bcrypt.hash(password, 10);
        await User.create({ username, password: hashedPassword, role });
        res.redirect('/admin/repair');
    } catch (err) {
        res.status(500).send('Gagal membuat user: ' + err.message);
    }
});

app.get('/admin/repair/export', isAdmin, async (req, res) => {
    try {
        const { startDate, endDate, namaMesin } = req.query;
        let query = {};
        if (startDate && endDate) {
            query.tanggal = { $gte: new Date(startDate), $lte: new Date(endDate + 'T23:59:59') };
        }
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

// --- TEKNISI DASHBOARD ---
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

// --- TEKNISI: HALAMAN CARI SPAREPART ---
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
        res.render('teknisi-sparepart', { user: req.session.user, spareparts, search: search || '' });
    } catch (err) {
        res.status(500).send(err.message);
    }
});

// Input Perbaikan & Pengurangan Otomatis Stok
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

                    sparepartsUsed.push({
                        sparepart: spId,
                        jumlahPakai: qty
                    });
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
        res.status(500).send('Gagal menyimpan perbaikan: ' + err.message);
    }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));

module.exports = app;
