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

// Fungsi Bantu: Resolve Thumbnail Instagram yang Lebih Stabil
async function resolveIgThumbnail(url) {
    if (!url || url === '#') return 'https://via.placeholder.com/300';
    
    // Jika admin memasukkan direct link gambar (.jpg, .png, .webp), langsung gunakan
    if (url.match(/\.(jpeg|jpg|gif|png|webp)$/i)) return url;
    
    // Jika memasukkan link postingan IG, gunakan format alternatif atau arahkan ke placeholder/fallback aman
    try {
        if (url.includes('/p/') || url.includes('/reel/')) {
            // Instagram sering memblokir hotlink langsung via tag img standar tanpa token graph api.
            // Solusi paling aman & stabil: Admin dapat menggunakan direct link gambar dari ImgBB/Postimages, 
            // atau kita sediakan fallback ke URL media jika diizinkan browser.
            const cleanUrl = url.split('?')[0];
            return `${cleanUrl.endsWith('/') ? cleanUrl : cleanUrl + '/'}media/?size=l`;
        }
        return url;
    } catch (e) {
        return 'https://via.placeholder.com/300';
    }
}

// API Utama: Membaca Data Spreadsheet Realtime
app.get('/api/data', async (req, res) => {
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

        // 5. Sheet Instagram Feed (Mengambil Baris 2 sampai 5)
        let instagramFeed = [];
        let instagramRaw = [];
        try {
            const resIg = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: 'Instagram!A2:C5',
            });
            const igRows = resIg.data.values || [];
            instagramRaw = igRows; // Untuk dikirim ke admin form
            for (let row of igRows) {
                if (row[0] || row[1] || row[2]) {
                    const postUrl = row[2] || row[1] || '#';
                    const resolvedImg = await resolveIgThumbnail(postUrl);
                    instagramFeed.push({
                        judul: row[0] || 'Postingan Instagram Dinas Pendidikan',
                        gambar: resolvedImg,
                        link: postUrl
                    });
                }
            }
        } catch (e) {
            console.error("Error Sheet Instagram:", e.message);
        }

        // 6. Sheet Footer & Settings Data
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
            link_pengaduan: "#",
            about_versi: "v3.0 (Official Build 2026)",
            about_judul_pencetus: "PENCETUS APLIKASI",
            about_pencetus: "Kepala Dinas Pendidikan Kabupaten Bangkalan",
            about_judul_penggagas: "PENGGAGAS APLIKASI",
            about_penggagas: "Bidang Pembinaan & Sekretariat Disdik Bangkalan",
            about_judul_pembuat: "PEMBUAT & PENGEMBANG APLIKASI",
            about_pembuat: "TIM IT Sekretariat Dinas Pendidikan Kabupaten Bangkalan",
            about_deskripsi: "E-Ladinin (Elektronik Layanan Dinas Pendidikan) merupakan inovasi portal layanan publik terpadu berbasis digital untuk meningkatkan transparansi, efisiensi, dan akuntabilitas pelayanan administrasi pendidikan di Kabupaten Bangkalan."
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
                if (row[0] && row[1] !== undefined) {
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

        res.json({ services, announcement, faqs, skm, instagramFeed, instagramRaw, footer, visitCount: currentVisitCount });
    } catch (error) {
        res.status(500).json({ error: error.message });
    }
});

// Endpoint Tambah Layanan
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

// Endpoint Simpan Pengaturan Footer & About
app.post('/api/update-settings', async (req, res) => {
    try {
        const settings = req.body;
        const sheets = await getSheetsClient();
        const spreadsheetId = process.env.SPREADSHEET_ID;

        const resFooter = await sheets.spreadsheets.values.get({
            spreadsheetId,
            range: 'Footer!A1:B100',
        });
        const rows = resFooter.data.values || [];

        for (const [key, val] of Object.entries(settings)) {
            let foundIndex = -1;
            for (let i = 0; i < rows.length; i++) {
                if (rows[i][0] && rows[i][0].toString().toLowerCase().trim().replace(/\s+/g, '_') === key) {
                    foundIndex = i + 1;
                    break;
                }
            }

            if (foundIndex !== -1) {
                await sheets.spreadsheets.values.update({
                    spreadsheetId,
                    range: `Footer!B${foundIndex}`,
                    valueInputOption: 'USER_ENTERED',
                    requestBody: { values: [[val]] }
                });
            } else {
                await sheets.spreadsheets.values.append({
                    spreadsheetId,
                    range: 'Footer!A:B',
                    valueInputOption: 'USER_ENTERED',
                    requestBody: { values: [[key, val]] }
                });
            }
        }

        res.json({ success: true, message: '✅ Pengaturan berhasil diperbarui!' });
    } catch (error) {
        res.json({ success: false, message: '❌ Gagal memperbarui pengaturan: ' + error.message });
    }
});

// Endpoint Update Feed Instagram dari Panel Admin
app.post('/api/update-instagram', async (req, res) => {
    try {
        const { posts } = req.body; 
        const sheets = await getSheetsClient();
        const spreadsheetId = process.env.SPREADSHEET_ID;

        const values = posts.map(p => [p.judul || '', p.link || '', p.link || '']);

        await sheets.spreadsheets.values.update({
            spreadsheetId,
            range: 'Instagram!A2:C5',
            valueInputOption: 'USER_ENTERED',
            requestBody: { values }
        });

        res.json({ success: true, message: '✅ Feed Instagram berhasil diperbarui!' });
    } catch (error) {
        res.json({ success: false, message: '❌ Gagal memperbarui IG: ' + error.message });
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