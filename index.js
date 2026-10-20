const express = require('express');
const path = require('path');
const { google } = require('googleapis');
require('dotenv').config();

const app = express();

app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

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

// API Utama: Membaca Data Spreadsheet Realtime (Tanpa Cache)
app.get('/api/data', async (req, res) => {
    // Matikan Cache di Vercel/Browser agar update link di Google Sheet langsung berubah saat itu juga
    res.setHeader('Cache-Control', 'no-store, no-cache, must-revalidate, proxy-revalidate');
    res.setHeader('Pragma', 'no-cache');
    res.setHeader('Expires', '0');

    try {
        const sheets = await getSheetsClient();
        const spreadsheetId = process.env.SPREADSHEET_ID;

        // 1. Sheet Layanan
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
        } catch (e) { 
            console.error("Error Sheet Layanan:", e.message); 
        }

        // 2. Sheet Pengumuman
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
        } catch (e) {}

        // 3. Sheet FAQ
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
        } catch (e) {}

        // 4. Sheet SKM
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
            const rowsSKM = resSKM.data.values || [];
            if (rowsSKM.length >= 2) {
                const sRow = rowsSKM[1];
                skm = {
                    periode: sRow[0] || "Triwulan III - 2026",
                    nilai: sRow[1] || "88.75",
                    mutu: sRow[2] || "A",
                    kinerja: sRow[3] || "Sangat Baik",
                    catatan: sRow[4] || "Berdasarkan hasil survei responden pengguna layanan digital Dinas Pendidikan Kabupaten Bangkalan."
                };
            }
        } catch (e) { 
            console.error("Error Sheet SKM:", e.message); 
        }

        // 5. Sheet Footer & Counter Pengunjung
        let footer = {
            deskripsi: "Sistem Elektronik Layanan Dinas Pendidikan Kabupaten Bangkalan.",
            alamat: "Jl. Soekarno Hatta No. 23, Bangkalan - Madura",
            email: "disdik.bklkab@gmail.com",
            telepon: "+62 852-3613-7770",
            jam_kerja: "Senin - Jumat: 08.00 - 15.00 WIB",
            instagram: "#",
            tiktok: "#",
            facebook: "#",
            youtube: "#",
            copyright: "© 2026 E-Ladinin v3.0. Dinas Pendidikan Kabupaten Bangkalan.",
            pengembang: "Created by TIM IT Sekretariat Disdik",
            link_mekanisme: "#",
            link_pengaduan: "#"
        };

        let currentVisitCount = 6911;
        let visitRowIndex = null;

        try {
            const resFooter = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: 'Footer!A1:B50',
            });
            const ftRows = resFooter.data.values || [];
            ftRows.forEach((row, index) => {
                if (row[0] && row[1]) {
                    const key = row[0].toString().toLowerCase().trim().replace(/\s+/g, '_');
                    const val = row[1].toString().trim();

                    if (key === 'visit_count') {
                        currentVisitCount = parseInt(val, 10) || 6911;
                        visitRowIndex = index + 1;
                    } else {
                        footer[key] = val;
                    }
                }
            });

            currentVisitCount += 1;

            if (visitRowIndex) {
                sheets.spreadsheets.values.update({
                    spreadsheetId,
                    range: `Footer!B${visitRowIndex}`,
                    valueInputOption: 'USER_ENTERED',
                    requestBody: { values: [[currentVisitCount]] }
                }).catch(err => console.error("Gagal update visit_count:", err.message));
            }

        } catch (e) { 
            console.error("Error Sheet Footer/Counter:", e.message); 
        }

        res.json({ services, announcement, faqs, skm, footer, visitCount: currentVisitCount });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

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