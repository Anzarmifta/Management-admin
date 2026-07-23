const mongoose = require('mongoose');

const sparepartSchema = new mongoose.Schema({
    namaBarang: { type: String, required: true },
    kodeBarang: { type: String, required: true, unique: true },
    lokasiBarang: { type: String },
    typeBarang: { type: String },
    stokBarang: { type: Number, default: 0 },
    barangKeluar: { type: Number, default: 0 },
    barangMasuk: { type: Number, default: 0 },
    fungsiBarang: { type: String }
});

module.exports = mongoose.model('Sparepart', sparepartSchema);
