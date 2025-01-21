const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');

const app = express();
const port = 3000;

// ตั้งค่าการจัดเก็บไฟล์ในหน่วยความจำแทนการเก็บในดิสก์
const storage = multer.memoryStorage();
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

// API สำหรับอัปโหลดไฟล์ Excel
app.post('/upload', upload.single('file'), (req, res) => {
    if (!req.file) {
        return res.send('กรุณาอัพโหลดไฟล์ Excel');
    }

    // อ่านข้อมูลจากไฟล์ Excel ที่ถูกเก็บในหน่วยความจำ
    const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
    const sheetName1 = workbook.SheetNames[0];
    const sheet1 = workbook.Sheets[sheetName1];
    const data1 = xlsx.utils.sheet_to_json(sheet1);  // ข้อมูลจาก Sheet1

    const sheetName2 = workbook.SheetNames[1];
    const sheet2 = workbook.Sheets[sheetName2];
    const data2 = xlsx.utils.sheet_to_json(sheet2);  // ข้อมูลจาก Sheet2

    // ประมวลผลข้อมูลใน Sheet1 (หมวดหมู่)
    let categoryCount = new Map();
    data1.forEach(row => {
        const category = row['หมวดหมู่']; // ชื่อหมวดหมู่
        if (category) {
            categoryCount.set(category, (categoryCount.get(category) || 0) + 1);
        }
    });

    // ประมวลผลข้อมูลใน Sheet2 (โปรสินค้าและอันดับ)
    let productCount = new Map();
    let rankSet = new Map();

    data2.forEach(row => {
        const product = row['โปรสินค้า']; // รหัสสินค้า
        const rank = row['อันดับ']; // อันดับ

        // นับจำนวนโปรสินค้า
        if (product) {
            productCount.set(product, (productCount.get(product) || 0) + 1);
        }

        // เก็บชุดอันดับ (2 ตัวแรกเป็นหมวดหมู่)
        if (rank) {
            const categoryCode = rank.toString().slice(0, 2); // ใช้ 2 ตัวแรกของอันดับเป็นรหัสหมวดหมู่

            // ถ้ายังไม่มีหมวดหมู่ใน rankSet ให้เพิ่ม
            if (!rankSet.has(product)) {
                rankSet.set(product, new Set());
            }
            rankSet.get(product).add(rank); // เก็บอันดับทั้งหมดที่ซ้ำกัน
        }
    });

    // ผลลัพธ์จากการคำนวณ Sheet1
    const result1 = Array.from(categoryCount.entries()).map(([category, count]) => {
        return {
            category,
            categoryCount: count
        };
    });

    // ผลลัพธ์จากการคำนวณ Sheet2
    const result2 = Array.from(productCount.entries()).map(([product, count]) => {
        const ranks = rankSet.get(product) || new Set();
        const rankCount = ranks.size; // จำนวนอันดับที่ซ้ำกัน (จาก Set)

        return {
            product,
            productCount: count,
            rankCount: rankCount
        };
    });

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
                    <h3>ข้อมูลจาก Sheet 1 (หมวดหมู่)</h3>
                    <table class="table table-bordered">
                        <thead class="table-dark">
                            <tr>
                                <th>หมวดหมู่</th>
                                <th>จำนวนหมวดหมู่</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${result1.map(category => `
                                <tr>
                                    <td>${category.category}</td>
                                    <td>${category.categoryCount}</td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                </div>

                <div class="mt-4">
                    <h3>ข้อมูลจาก Sheet 2 (โปรสินค้าและอันดับ)</h3>
                    <table class="table table-bordered">
                        <thead class="table-dark">
                            <tr>
                                <th>โปรสินค้า (รหัสสินค้า)</th>
                                <th>จำนวนโปรสินค้า</th>
                                <th>จำนวนอันดับซ้ำ (หมวดหมู่)</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${result2.map(category => `
                                <tr>
                                    <td>${category.product}</td>
                                    <td>${category.productCount}</td>
                                    <td>${category.rankCount}</td>
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
