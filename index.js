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

async function resolveIgThumbnail(url) {
    if (!url || url === '#') return 'https://via.placeholder.com/300';
    if (url.includes('drive.google.com')) {
        const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
        if (match && match[1]) {
            return `https://lh3.googleusercontent.com/d/${match[1]}`;
        }
    }
    if (url.match(/\.(jpeg|jpg|gif|png|webp)$/i) || url.includes('lh3.googleusercontent.com')) {
        return url;
    }
    try {
        if (url.includes('/p/') || url.includes('/reel/')) {
            const cleanUrl = url.split('?')[0];
            return `${cleanUrl.endsWith('/') ? cleanUrl : cleanUrl + '/'}media/?size=l`;
        }
        return url;
    } catch (e) {
        return 'https://via.placeholder.com/300';
    }
}

function parseNumber(val) {
    if (val === undefined || val === null) return 0;
    if (typeof val === 'number') return Math.round(val);
    const raw = String(val).trim();
    if (!raw) return 0;
    const cleaned = raw.replace(/\s+/g, '').replace(/\./g, '').replace(/,/g, '').replace(/[^0-9-]/g, '');
    const num = parseInt(cleaned, 10);
    return Number.isNaN(num) ? 0 : num;
}

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

        // 5. Sheet Instagram Feed
        let instagramFeed = [];
        let instagramRaw = [];
        try {
            const resIg = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: 'Instagram!A2:C5',
            });
            const igRows = resIg.data.values || [];
            instagramRaw = igRows;
            for (let row of igRows) {
                if (row[0] || row[1] || row[2]) {
                    const imgInput = row[1] || row[2] || '#';
                    const postLink = row[2] || '#';
                    const resolvedImg = await resolveIgThumbnail(imgInput);
                    instagramFeed.push({
                        judul: row[0] || 'Postingan Instagram Dinas Pendidikan',
                        gambar: resolvedImg,
                        link: postLink
                    });
                }
            }
        } catch (e) {
            console.error("Error Sheet Instagram:", e.message);
        }

        // 6. Sheet Rekap & Sheet ASN
        let rekapRows = [];
        let rekapError = null;
        try {
            const resRekap = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: 'Rekap!A1:K100',
            });
            rekapRows = resRekap.data.values || [];
        } catch (e) {
            rekapError = e.message;
            console.error("Error Sheet Rekap:", e.message);
        }

        let explicitTotalRow = null;
        const items = [];
        let sumNegPns = 0, sumNegPppk = 0, sumNegPw = 0, sumNegTot = 0;
        let sumSwsPns = 0, sumSwsPppk = 0, sumSwsPw = 0, sumSwsTot = 0, sumTotAsn = 0;

        for (let i = 0; i < rekapRows.length; i++) {
            const r = rekapRows[i] || [];
            const colA = String(r[0] || '').trim();
            const colB = String(r[1] || '').trim();

            if (colB.toUpperCase() === 'TOTAL' || colA.toUpperCase() === 'TOTAL') {
                explicitTotalRow = r;
                continue;
            }

            if (/^\d+$/.test(colA) && colB) {
                const negeriPns = parseNumber(r[2]);
                const negeriPppk = parseNumber(r[3]);
                const negeriPppkPw = parseNumber(r[4]);
                const negeriTotal = r[5] !== undefined && String(r[5]).trim() !== ''
                    ? parseNumber(r[5])
                    : (negeriPns + negeriPppk + negeriPppkPw);

                const swastaPns = parseNumber(r[6]);
                const swastaPppk = parseNumber(r[7]);
                const swastaPppkPw = parseNumber(r[8]);
                const swastaTotal = r[9] !== undefined && String(r[9]).trim() !== ''
                    ? parseNumber(r[9])
                    : (swastaPns + swastaPppk + swastaPppkPw);

                const totalAsn = r[10] !== undefined && String(r[10]).trim() !== ''
                    ? parseNumber(r[10])
                    : (negeriTotal + swastaTotal);

                items.push({
                    no: parseNumber(colA),
                    jabatan: colB,
                    negeriPns,
                    negeriPppk,
                    negeriPppkPw,
                    negeriTotal,
                    swastaPns,
                    swastaPppk,
                    swastaPppkPw,
                    swastaTotal,
                    totalAsn
                });

                sumNegPns += negeriPns;
                sumNegPppk += negeriPppk;
                sumNegPw += negeriPppkPw;
                sumNegTot += negeriTotal;
                sumSwsPns += swastaPns;
                sumSwsPppk += swastaPppk;
                sumSwsPw += swastaPppkPw;
                sumSwsTot += swastaTotal;
                sumTotAsn += totalAsn;
            }
        }

        let rekapTotalObj = {
            negeriPns: sumNegPns,
            negeriPppk: sumNegPppk,
            negeriPppkPw: sumNegPw,
            negeriTotal: sumNegTot,
            swastaPns: sumSwsPns,
            swastaPppk: sumSwsPppk,
            swastaPppkPw: sumSwsPw,
            swastaTotal: sumSwsTot,
            totalAsn: sumTotAsn
        };

        if (explicitTotalRow) {
            rekapTotalObj = {
                negeriPns: explicitTotalRow[2] !== undefined && String(explicitTotalRow[2]).trim() !== '' ? parseNumber(explicitTotalRow[2]) : sumNegPns,
                negeriPppk: explicitTotalRow[3] !== undefined && String(explicitTotalRow[3]).trim() !== '' ? parseNumber(explicitTotalRow[3]) : sumNegPppk,
                negeriPppkPw: explicitTotalRow[4] !== undefined && String(explicitTotalRow[4]).trim() !== '' ? parseNumber(explicitTotalRow[4]) : sumNegPw,
                negeriTotal: explicitTotalRow[5] !== undefined && String(explicitTotalRow[5]).trim() !== '' ? parseNumber(explicitTotalRow[5]) : sumNegTot,
                swastaPns: explicitTotalRow[6] !== undefined && String(explicitTotalRow[6]).trim() !== '' ? parseNumber(explicitTotalRow[6]) : sumSwsPns,
                swastaPppk: explicitTotalRow[7] !== undefined && String(explicitTotalRow[7]).trim() !== '' ? parseNumber(explicitTotalRow[7]) : sumSwsPppk,
                swastaPppkPw: explicitTotalRow[8] !== undefined && String(explicitTotalRow[8]).trim() !== '' ? parseNumber(explicitTotalRow[8]) : sumSwsPw,
                swastaTotal: explicitTotalRow[9] !== undefined && String(explicitTotalRow[9]).trim() !== '' ? parseNumber(explicitTotalRow[9]) : sumSwsTot,
                totalAsn: explicitTotalRow[10] !== undefined && String(explicitTotalRow[10]).trim() !== '' ? parseNumber(explicitTotalRow[10]) : sumTotAsn
            };
        }

        let pnsVal = rekapTotalObj.negeriPns + rekapTotalObj.swastaPns;
        let pppkVal = rekapTotalObj.negeriPppk + rekapTotalObj.swastaPppk;
        let pppkPwVal = rekapTotalObj.negeriPppkPw + rekapTotalObj.swastaPppkPw;
        let totalVal = rekapTotalObj.totalAsn || (pnsVal + pppkVal + pppkPwVal);
        let ketVal = "REKAP JUMLAH ASN";

        let asnSheetRows = [];
        try {
            const resAsn = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: 'ASN!A1:B20',
            });
            asnSheetRows = resAsn.data.values || [];
            asnSheetRows.forEach(row => {
                const k = String(row[0] || '').toUpperCase().trim();
                const v = row[1];
                if ((k === 'PERIODE' || k === 'KETERANGAN') && v) {
                    ketVal = String(v).trim();
                }
                if (items.length === 0 && pnsVal === 0 && pppkVal === 0 && pppkPwVal === 0) {
                    if (k === 'PNS') pnsVal = parseNumber(v);
                    else if (k === 'PPPK') pppkVal = parseNumber(v);
                    else if (k === 'PPPK PW' || k === 'PPPK_PW' || k === 'PPPKPW') pppkPwVal = parseNumber(v);
                    totalVal = pnsVal + pppkVal + pppkPwVal;
                }
            });
        } catch (_eAsn) {}

        const asn = {
            pns: pnsVal,
            pppk: pppkVal,
            pppkPw: pppkPwVal,
            total: totalVal,
            keterangan: ketVal,
            rekapItems: items,
            rekapTotal: rekapTotalObj,
            debug: {
                rekapRowsCount: rekapRows.length,
                rekapError,
                asnRowsCount: asnSheetRows.length
            }
        };

        // 7. Sheet Footer & Settings Data
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

        res.json({ services, announcement, faqs, skm, asn, instagramFeed, instagramRaw, footer, visitCount: currentVisitCount });
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

app.post('/api/update-instagram', async (req, res) => {
    try {
        const { posts } = req.body;
        const sheets = await getSheetsClient();
        const spreadsheetId = process.env.SPREADSHEET_ID;

        const values = posts.map(p => [p.judul || '', p.gambar || '', p.link || '']);

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