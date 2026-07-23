const mongoose = require('mongoose');

const repairSchema = new mongoose.Schema({
    tanggal: { type: Date, required: true },
    shift: { type: String, required: true },
    namaMesin: { type: String, required: true },
    problem: { type: String, required: true },
    analisaPenyebab: { type: String },
    caraPerbaikan: { type: String },
    sparepartsUsed: [{
        sparepart: { type: mongoose.Schema.Types.ObjectId, ref: 'Sparepart' },
        jumlahPakai: { type: Number, default: 0 }
    }],
    pic: { type: String, required: true },
    jamMulai: { type: String, required: true },
    jamSelesai: { type: String, required: true },
    status: { type: String, enum: ['ok', 'open', 'progress'], default: 'open' },
    keterangan: { type: String }
});

module.exports = mongoose.model('Repair', repairSchema);
