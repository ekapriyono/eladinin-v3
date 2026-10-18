const express = require('express');
const path = require('path');
const { google } = require('googleapis');
require('dotenv').config();

const app = express();

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

// Rute Utama: Menampilkan index.html
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Autentikasi Google Sheets API via Service Account
const getSheetsClient = async () => {
    const auth = new google.auth.GoogleAuth({
        credentials: {
            client_email: process.env.GOOGLE_CLIENT_EMAIL,
            private_key: process.env.GOOGLE_PRIVATE_KEY ? process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, '\n') : undefined,
        },
        scopes: ['https://www.googleapis.com/auth/spreadsheets'],
    });
    const client = await auth.getClient();
    return google.sheets({ version: 'v4', auth: client });
};

let visitCount = 1420;

// API Utama: Membaca Layanan, Pengumuman, FAQ, SKM, dan Footer
app.get('/api/data', async (req, res) => {
    try {
        visitCount++;
        const sheets = await getSheetsClient();
        const spreadsheetId = process.env.SPREADSHEET_ID;

        // 1. Ambil data dari sheet 'Layanan'
        let services = [];
        try {
            const resServices = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: 'Layanan!A:G',
            });
            const rows = resServices.data.values || [];
            for (let i = 1; i < rows.length; i++) {
                if (rows[i][1]) {
                    services.push({
                        no: rows[i][0],
                        judul: rows[i][1],
                        deskripsi: rows[i][2],
                        link: rows[i][3],
                        icon: rows[i][4] || "fa-solid fa-gear",
                        kategori: rows[i][5] || "Umum",
                        status: (rows[i][6] || "ONLINE").toUpperCase()
                    });
                }
            }
        } catch (e) { console.error("Error Sheet Layanan:", e.message); }

        // 2. Ambil data dari sheet 'Pengumuman'
        let announcement = null;
        try {
            const resAnnounce = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: 'Pengumuman!A2:C2',
            });
            const aRow = resAnnounce.data.values ? resAnnounce.data.values[0] : null;
            if (aRow && aRow[0] === 'AKTIF') {
                announcement = { pesan: aRow[1], tipe: aRow[2] || 'info' };
            }
        } catch (e) { /* Gunakan null jika sheet belum ada */ }

        // 3. Ambil data dari sheet 'FAQ'
        let faqs = [];
        try {
            const resFaq = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: 'FAQ!A:B',
            });
            const fRows = resFaq.data.values || [];
            for (let i = 1; i < fRows.length; i++) {
                if (fRows[i][0]) {
                    faqs.push({ pertanyaan: fRows[i][0], jawaban: fRows[i][1] });
                }
            }
        } catch (e) { /* Gunakan array kosong jika sheet belum ada */ }

        // 4. Ambil data dari sheet 'SKM'
        let skm = {
            periode: "Triwulan III - 2026",
            nilai: "88.75",
            mutu: "A",
            kinerja: "Sangat Baik",
            catatan: "Berdasarkan hasil survei responden pengguna layanan digital Dinas Pendidikan Kabupaten Bangkalan."
        };
        try {
            const resSKM = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: 'SKM!A1:E5',
            });
            const rows = resSKM.data.values || [];
            if (rows.length >= 2) {
                // Baris ke-2 (indeks 1) berisi data SKM
                const sRow = rows[1];
                skm = {
                    periode: sRow[0] || "Triwulan III - 2026",
                    nilai: sRow[1] || "88.75",
                    mutu: sRow[2] || "A",
                    kinerja: sRow[3] || "Sangat Baik",
                    catatan: sRow[4] || "Berdasarkan hasil survei responden pengguna layanan digital Dinas Pendidikan Kabupaten Bangkalan."
                };
            }
        } catch (e) {
            console.log("Sheet SKM belum dibuat atau error:", e.message);
        }

        // 5. Ambil data dari sheet 'Footer'
        let footer = {
            deskripsi: "Sistem Elektronik Layanan Dinas Pendidikan Kabupaten Bangkalan. Mewujudkan tata kelola pelayanan publik yang efektif, efisien, transparan, dan akuntabel.",
            alamat: "Jl. Soekarno Hatta No. 23, Bangkalan - Madura",
            email: "disdik.bklkab@gmail.com",
            telepon: "+62 852-3613-7770",
            jam_kerja: "Senin - Jumat: 08.00 - 15.00 WIB",
            instagram: "#",
            tiktok: "#",
            facebook: "#",
            youtube: "#",
            copyright: "© 2026 E-Ladinin v3.0. Dinas Pendidikan Kabupaten Bangkalan.",
            pengembang: "Created by TIM IT Sekretariat Disdik"
        };

        try {
            const resFooter = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: 'Footer!A:B',
            });
            const ftRows = resFooter.data.values || [];
            ftRows.forEach(row => {
                if (row[0] && row[1]) {
                    const key = row[0].toLowerCase().trim();
                    if (footer.hasOwnProperty(key)) {
                        footer[key] = row[1];
                    }
                }
            });
        } catch (e) { /* Gunakan nilai default jika sheet belum ada */ }

        res.json({ services, announcement, faqs, skm, footer, visitCount });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// API: Tambah Layanan Baru dari Modal Admin
app.post('/api/add-service', async (req, res) => {
    try {
        const { judul, deskripsi, link, icon, kategori } = req.body;
        const sheets = await getSheetsClient();
        const spreadsheetId = process.env.SPREADSHEET_ID;

        const getRows = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: 'Layanan!A:A',
        });
        const newNo = getRows.data.values ? getRows.data.values.length : 1;

        await sheets.spreadsheets.values.append({
            spreadsheetId,
            range: 'Layanan!A:G',
            valueInputOption: 'USER_ENTERED',
            requestBody: {
                values: [[newNo, judul, deskripsi, link, icon || "fa-solid fa-gear", kategori || "Umum", "ONLINE"]]
            },
        });

        res.json({ status: 'success', message: '✅ Layanan berhasil ditambahkan!' });
    } catch (error) {
        res.json({ status: 'error', message: '❌ Error: ' + error.message });
    }
});

// API: Cek Password Admin
app.post('/api/check-password', (req, res) => {
    const { password } = req.body;
    const SECRET_PASSWORD = process.env.ADMIN_PASSWORD || "admin2026";
    res.json({ success: password === SECRET_PASSWORD });
});

module.exports = app;

if (process.env.NODE_ENV !== 'production') {
    const PORT = process.env.PORT || 3000;
    app.listen(PORT, () => {
        console.log(`Server E-Ladinin v3.0 aktif di http://localhost:${PORT}`);
    });
}