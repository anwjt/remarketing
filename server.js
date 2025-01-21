const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');

const app = express();
const port = 3000;

// ตั้งค่าการจัดเก็บไฟล์ที่อัพโหลด
const storage = multer.diskStorage({
    destination: (req, file, cb) => {
        cb(null, 'uploads/');
    },
    filename: (req, file, cb) => {
        cb(null, Date.now() + path.extname(file.originalname));
    }
});
const upload = multer({ storage: storage });

// ให้บริการไฟล์ CSS และ JS ของ Bootstrap
app.use('/static', express.static(path.join(__dirname, 'node_modules/bootstrap/dist')));

// หน้าแรกของเว็บ
app.get('/', (req, res) => {
    res.send(`
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Upload Excel</title>
            <link rel="stylesheet" href="/static/css/bootstrap.min.css">
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

// API สำหรับอัพโหลดไฟล์ Excel
app.post('/upload', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.send('กรุณาอัพโหลดไฟล์ Excel');
    }

    // อ่านข้อมูลจากไฟล์ Excel
    const workbook = xlsx.readFile(req.file.path);
    const sheetName1 = workbook.SheetNames[0];
    const sheet1 = workbook.Sheets[sheetName1];
    const data1 = xlsx.utils.sheet_to_json(sheet1);

    const sheetName2 = workbook.SheetNames[1];
    const sheet2 = workbook.Sheets[sheetName2];
    const data2 = xlsx.utils.sheet_to_json(sheet2);

    // ประมวลผลข้อมูลใน Sheet1 (ไม่ต้องเปลี่ยนแปลงตามที่เคยทำ)
    let productCount = new Set();
    let modelCount = new Map();

    data1.forEach(row => {
        if (row['สินค้า'] && row['รุ่นแบบ']) {
            productCount.add(row['สินค้า']);
            
            if (!modelCount.has(row['สินค้า'])) {
                modelCount.set(row['สินค้า'], new Set());
            }
            modelCount.get(row['สินค้า']).add(row['รุ่นแบบ']);
        }
    });

    const result1 = {
        totalProducts: productCount.size,
        totalModels: Array.from(modelCount.values()).reduce((acc, models) => acc + models.size, 0),
        productModelDetails: []
    };

    modelCount.forEach((models, product) => {
        result1.productModelDetails.push({
            product,
            totalModels: models.size
        });
    });

    // ประมวลผลข้อมูลใน Sheet2 (ตามที่คุณต้องการ)
    let productCount2 = new Map();

    data2.forEach(row => {
        if (row['โปรสินค้า'] && row['อันดับ']) {
            const productCode = row['โปรสินค้า']; // รหัสสินค้า
            const rank = row['อันดับ'].split('/')[0]; // เลขสองตัวหน้าเป็นอันดับ
            
            // นับจำนวนรหัสสินค้า
            if (!productCount2.has(productCode)) {
                productCount2.set(productCode, { count: 0, ranks: new Set() });
            }
            productCount2.get(productCode).count++;
            productCount2.get(productCode).ranks.add(rank);
        }
    });

    const result2 = Array.from(productCount2.entries()).map(([product, { count, ranks }]) => ({
        product,
        duplicateCount: count,
        uniqueRanks: ranks.size
    }));

    // แสดงผลลัพธ์ในรูปแบบ HTML
    res.send(`
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>Excel Results</title>
            <link rel="stylesheet" href="/static/css/bootstrap.min.css">
        </head>
        <body>
            <div class="container mt-5">
                <h1 class="text-center">ผลลัพธ์จากการประมวลผล Excel</h1>

                <div class="mt-4">
                    <h3>ข้อมูลจาก Sheet 1</h3>
                    <ul>
                        <li>จำนวนสินค้าทั้งหมด: <strong>${result1.totalProducts}</strong></li>
                        <li>จำนวนรุ่นแบบทั้งหมด: <strong>${result1.totalModels}</strong></li>
                    </ul>
                    <table class="table table-bordered">
                        <thead class="table-dark">
                            <tr>
                                <th>สินค้า</th>
                                <th>จำนวนรุ่นแบบ</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${result1.productModelDetails.map(row => `
                                <tr>
                                    <td>${row.product}</td>
                                    <td>${row.totalModels}</td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>

                <div class="mt-4">
                    <h3>ข้อมูลจาก Sheet 2</h3>
                    <table class="table table-bordered">
                        <thead class="table-dark">
                            <tr>
                                <th>โปรสินค้า (รหัสสินค้า)</th>
                                <th>จำนวนข้อมูลที่ซ้ำ (ครั้ง)</th>
                                <th>จำนวนอันดับ</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${result2.map(row => `
                                <tr>
                                    <td>${row.product}</td>
                                    <td>${row.duplicateCount}</td>
                                    <td>${row.uniqueRanks}</td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>

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
