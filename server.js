const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');

const app = express();
const port = 3000;

// ตั้งค่าการอัปโหลดไฟล์ในหน่วยความจำ
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// หน้าแรก
app.get('/', (req, res) => {
    res.send(`
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Upload Excel</title>
            <!-- ใช้ CDN ของ Bootstrap -->
            <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
        </head>
        <body>
            <div class="container mt-5">
                <h1 class="text-center">Upload Excel File</h1>
                <form id="uploadForm" action="/upload" method="post" enctype="multipart/form-data" class="mt-4">
                    <div class="mb-3">
                        <input type="file" name="file" class="form-control" required />
                    </div>
                    <button type="submit" class="btn btn-primary w-100">Upload and Process</button>
                </form>
            </div>
        </body>
        </html>
    `);
});

// API สำหรับอัปโหลดไฟล์
app.post('/upload', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.send('กรุณาอัปโหลดไฟล์ Excel');
    }

    // อ่านไฟล์จาก buffer ในหน่วยความจำ
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });

    /** ประมวลผล Sheet แรก */
    const sheetName1 = workbook.SheetNames[0]; // ดึง Sheet แรก
    const sheet1 = workbook.Sheets[sheetName1];
    const data1 = xlsx.utils.sheet_to_json(sheet1);

    // การประมวลผลข้อมูลจาก Sheet แรก
    let productSet1 = new Set();
    let modelMap1 = new Map();

    data1.forEach(row => {
        if (row['สินค้า'] && row['รุ่นแบบ']) {
            productSet1.add(row['สินค้า']);
            if (!modelMap1.has(row['สินค้า'])) {
                modelMap1.set(row['สินค้า'], new Set());
            }
            modelMap1.get(row['สินค้า']).add(row['รุ่นแบบ']);
        }
    });

    const result1 = Array.from(modelMap1.entries()).map(([product, models]) => ({
        product,
        totalModels: models.size
    })).sort((a, b) => b.totalModels - a.totalModels); // เรียงจากมากไปน้อย

    // คำนวณจำนวนสินค้าทั้งหมดและจำนวนรุ่นแบบทั้งหมด
    const totalProducts = productSet1.size;
    const totalModels = result1.reduce((acc, row) => acc + row.totalModels, 0);

    /** ประมวลผล Sheet ถัดไป (Sheet ที่ 2 เป็นต้นไป) */
    const otherSheets = workbook.SheetNames.slice(1); // ดึง Sheet ทั้งหมดหลัง Sheet แรก
    if (otherSheets.length === 0) {
        return res.send('ไม่พบ Sheet ถัดไปในไฟล์ Excel');
    }

    let allResults = [];

    otherSheets.forEach(sheetName => {
        const sheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(sheet);

        if (data.length > 0) {
            let productCountMap = new Map();
            let rankSetMap = new Map();

            data.forEach(row => {
                const product = row['โปรสินค้า'];
                const rank = row['อันดับ'];

                if (product) {
                    productCountMap.set(product, (productCountMap.get(product) || 0) + 1);
                }

                if (rank) {
                    if (!rankSetMap.has(product)) {
                        rankSetMap.set(product, new Set());
                    }
                    rankSetMap.get(product).add(rank);
                }
            });

            const result = Array.from(productCountMap.entries()).map(([product, count]) => ({
                product,
                duplicateCount: count,
                uniqueRanks: rankSetMap.get(product)?.size || 0
            })).sort((a, b) => b.duplicateCount - a.duplicateCount); // เรียงจากมากไปน้อย

            // คำนวณผลรวมในแถวล่างสุดของตาราง
            const totalDuplicateCount = result.reduce((acc, row) => acc + row.duplicateCount, 0);
            const totalUniqueRanks = result.reduce((acc, row) => acc + row.uniqueRanks, 0);

            // เพิ่มแถวผลรวมสุดท้าย
            result.push({
                product: 'รวมทั้งหมด',
                duplicateCount: totalDuplicateCount,
                uniqueRanks: totalUniqueRanks
            });

            allResults.push({ sheetName, result });
        }
    });

    /** ส่งผลลัพธ์ */
    res.send(`
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Excel Results</title>
            <!-- ใช้ CDN ของ Bootstrap -->
            <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
        </head>
        <body>
            <div class="container mt-5">
                <h1 class="text-center">ผลลัพธ์จากการประมวลผล Excel</h1>

                <div class="mt-4">
                    <h3>ข้อมูลจาก Sheet แรก (${sheetName1})</h3>
                    <ul>
                        <li>จำนวนสินค้าทั้งหมด: <strong>${totalProducts}</strong></li>
                        <li>จำนวนรุ่นแบบทั้งหมด: <strong>${totalModels}</strong></li>
                    </ul>
                    <table class="table table-bordered table-striped">
                        <thead class="table-dark">
                            <tr>
                                <th>สินค้า</th>
                                <th>จำนวนรุ่นแบบ</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${result1.map(row => `
                                <tr>
                                    <td>${row.product}</td>
                                    <td>${row.totalModels}</td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>

                ${allResults.map(sheetResult => `
                    <div class="mt-4">
                        <h3>ข้อมูลจาก Sheet (${sheetResult.sheetName})</h3>
                        <table class="table table-bordered table-striped">
                            <thead class="table-dark">
                                <tr>
                                    <th>โปรสินค้า (รหัสสินค้า)</th>
                                    <th>จำนวนข้อมูลที่ซ้ำ (ครั้ง)</th>
                                    <th>จำนวนอันดับ</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${sheetResult.result.map(row => `
                                    <tr>
                                        <td>${row.product}</td>
                                        <td>${row.duplicateCount}</td>
                                        <td>${row.uniqueRanks}</td>
                                    </tr>
                                `).join('')}
                            </tbody>
                        </table>
                    </div>
                `).join('')}

                <div class="mt-4 text-center">
                    <a href="/" class="btn btn-primary">Upload File ใหม่</a>
                </div>
            </div>
        </body>
        </html>
    `);
});

// เริ่มเซิร์ฟเวอร์
app.listen(port, () => {
    console.log(`Server running at http://localhost:${port}`);
});
