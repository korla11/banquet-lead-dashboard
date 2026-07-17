const express = require('express');
const cors = require('cors');
const multer = require('multer');
const xlsx = require('xlsx');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;
const ADMIN_PASSWORD = "0808";

// Enable CORS and JSON parsing
app.disable('etag');
app.use((req, res, next) => {
    res.set({
        'Cache-Control': 'no-store, no-cache, must-revalidate, proxy-revalidate, max-age=0',
        'Pragma': 'no-cache',
        'Expires': '0',
        'Surrogate-Control': 'no-store'
    });
    next();
});
app.use(cors());
app.use(express.json());

// Serve static dashboard files from the current folder
app.use(express.static(__dirname, {
    cacheControl: false,
    etag: false,
    lastModified: false,
    maxAge: 0
}));

const DB_FILE = path.join(__dirname, 'db.json');
const HTML_FILE = path.join(__dirname, 'index.html');

// Stable user passwords
const USER_PASSWORDS = {
    "李冰冰": "8076",
    "冯玲": "2418",
    "郭楠": "9724",
    "张议丹": "5679",
    "徐俊强": "7307",
    "统筹": "1025",
    "灯光": "8276",
    "前厅": "2208",
    "后厨": "7721",
    "行政": "9574"
};

// Initialize database
function loadDatabase() {
    if (fs.existsSync(DB_FILE)) {
        try {
            return JSON.parse(fs.readFileSync(DB_FILE, 'utf8'));
        } catch (e) {
            console.error("Error reading db.json, reinitializing...", e);
        }
    }
    
    // Fallback: Recover from index.html BANQUET_DATA
    let raw_records = [];
    if (fs.existsSync(HTML_FILE)) {
        const html = fs.readFileSync(HTML_FILE, 'utf8');
        const match = html.match(/const BANQUET_DATA = (\[.*?\]);/s);
        if (match) {
            try {
                raw_records = JSON.parse(match[1]);
                console.log(`Recovered ${raw_records.length} records from index.html`);
            } catch (e) {
                console.error("Failed to parse BANQUET_DATA from index.html", e);
            }
        }
    }
    
    const db = {
        raw_records: raw_records,
        custom_records: [],
        modified_records: {}
    };
    saveDatabase(db);
    return db;
}

function saveDatabase(db) {
    fs.writeFileSync(DB_FILE, JSON.stringify(db, null, 2), 'utf8');
}

// Merge data helper
function getMergedRecords(db) {
    // Start with raw records
    const recordsMap = new Map();
    db.raw_records.forEach(rec => {
        recordsMap.set(rec.id, { ...rec });
    });
    
    // Apply modifications to raw records
    Object.keys(db.modified_records).forEach(id => {
        if (recordsMap.has(id)) {
            const rawRec = recordsMap.get(id);
            const edits = db.modified_records[id];
            recordsMap.set(id, { ...rawRec, ...edits });
        }
    });
    
    // Add custom records
    const finalRecords = Array.from(recordsMap.values());
    db.custom_records.forEach(rec => {
        // Also apply modifications if any (though usually custom records are updated directly inline)
        const customEdits = db.modified_records[rec.id] || {};
        finalRecords.push({ ...rec, ...customEdits });
    });
    
    // Sort by id descending or default order
    return finalRecords;
}

// Multer setup for Excel uploads in memory
const upload = multer({ storage: multer.memoryStorage() });

// Helper to parse Excel file from buffer
function parseExcelBuffer(buffer) {
    const workbook = xlsx.read(buffer, { type: 'buffer' });
    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    // Convert to 2D array of formatted strings
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, raw: false });
    if (rows.length === 0) return [];
    
    const headers = rows[0];
    const dataRows = rows.slice(1);
    
    const headerMap = {};
    headers.forEach((h, i) => {
        if (h !== null && h !== undefined) {
            headerMap[String(h).trim()] = i;
        }
    });
    
    const clean = (val) => {
        if (val === null || val === undefined) return "";
        return String(val).trim().replace(/\r/g, "").replace(/\n/g, " ").trim();
    };
    
    const parsedRecords = [];
    dataRows.forEach(row => {
        // Skip empty rows
        if (row.length === 0 || row.every(cell => cell === null || cell === undefined || String(cell).trim() === "")) {
            return;
        }
        
        const idIndex = headerMap["自动编号"] !== undefined ? headerMap["自动编号"] : 0;
        const nameIndex = headerMap["客户姓名"] !== undefined ? headerMap["客户姓名"] : 3;
        const ownerIndex = headerMap["负责人"] !== undefined ? headerMap["负责人"] : 2;
        
        const id_val = clean(row[idIndex]);
        const name_val = clean(row[nameIndex]);
        const owner_val = clean(row[ownerIndex]);
        
        if (!id_val && !name_val && !owner_val) return;
        
        const channel_val = clean(row[headerMap["获客渠道"] !== undefined ? headerMap["获客渠道"] : 1]);
        const phone_val = clean(row[headerMap["联系电话"] !== undefined ? headerMap["联系电话"] : 4]);
        const raw_type = clean(row[headerMap["宴会类型"] !== undefined ? headerMap["宴会类型"] : 5]);
        const banquet_date = clean(row[headerMap["宴会日期"] !== undefined ? headerMap["宴会日期"] : 6]);
        const package_val = clean(row[headerMap["餐标套餐"] !== undefined ? headerMap["餐标套餐"] : 7]);
        const card_status = clean(row[headerMap["宴会卡购买状态"] !== undefined ? headerMap["宴会卡购买状态"] : 8]);
        const reg_time = clean(row[headerMap["登记时间"] !== undefined ? headerMap["登记时间"] : 9]);
        const remarks = clean(row[headerMap["备注"] !== undefined ? headerMap["备注"] : 10]);
        
        let card_type = "生日卡";
        if (raw_type.includes("生日") || card_status.includes("生日")) {
            card_type = "生日卡";
        } else if (raw_type.includes("订婚") || raw_type.includes("婚") || card_status.includes("订婚")) {
            card_type = "订婚卡";
        }
        
        parsedRecords.push({
            id: id_val,
            channel: channel_val || "未记录",
            owner: owner_val || "未分配",
            name: name_val || "无姓名",
            phone: phone_val || "未填",
            raw_type: raw_type || "生日宴",
            card_type: card_type,
            banquet_date: banquet_date || "未定",
            package: package_val || "未选定",
            card_status: card_status || "未处理",
            reg_time: reg_time || "2026年7月13日",
            remarks: remarks || "无"
        });
    });
    return parsedRecords;
}

// API: Get complete data state
app.get('/api/data', (req, res) => {
    try {
        const db = loadDatabase();
        const records = getMergedRecords(db);
        res.json({
            success: true,
            records: records,
            passwords: USER_PASSWORDS
        });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

// API: Add new custom record
app.post('/api/data', (req, res) => {
    try {
        const db = loadDatabase();
        const newRecord = req.body;
        
        // Generate stable CUSTOM_ ID using current timestamp
        if (!newRecord.id) {
            newRecord.id = `CUSTOM_${Date.now()}`;
        }
        if (!newRecord.reg_time) {
            const now = new Date();
            newRecord.reg_time = `${now.getFullYear()}年${now.getMonth() + 1}月${now.getDate()}日`;
        }
        
        db.custom_records.push(newRecord);
        saveDatabase(db);
        
        res.json({ success: true, record: newRecord });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

// API: Modify existing record
app.put('/api/data/:id', (req, res) => {
    try {
        const { id } = req.params;
        const db = loadDatabase();
        const edits = req.body;
        
        // If it's a custom record, we can modify it directly in custom_records
        const customIdx = db.custom_records.findIndex(rec => rec.id === id);
        if (customIdx !== -1) {
            db.custom_records[customIdx] = { ...db.custom_records[customIdx], ...edits };
        } else {
            // Otherwise, track it in modified_records map
            db.modified_records[id] = { ...db.modified_records[id], ...edits };
        }
        
        saveDatabase(db);
        res.json({ success: true, id: id });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

// API: Delete record (admin only)
app.delete('/api/data/:id', (req, res) => {
    if (req.get('x-admin-password') !== ADMIN_PASSWORD) {
        return res.status(403).json({ success: false, error: "仅系统管理员可以删除客资" });
    }

    try {
        const { id } = req.params;
        const db = loadDatabase();
        const rawLength = db.raw_records.length;
        const customLength = db.custom_records.length;

        db.raw_records = db.raw_records.filter(rec => String(rec.id) !== String(id));
        db.custom_records = db.custom_records.filter(rec => String(rec.id) !== String(id));
        if (db.modified_records && db.modified_records[id]) {
            delete db.modified_records[id];
        }

        if (db.raw_records.length === rawLength && db.custom_records.length === customLength) {
            return res.status(404).json({ success: false, error: "未找到需要删除的客资" });
        }

        saveDatabase(db);
        res.json({ success: true, id: id });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

// API: Admin Excel Upload
app.post('/api/upload-excel', upload.single('file'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ success: false, error: "No file uploaded" });
        }
        
        const newRawRecords = parseExcelBuffer(req.file.buffer);
        if (newRawRecords.length === 0) {
            return res.status(400).json({ success: false, error: "Excel sheet is empty or invalid structure" });
        }
        
        const db = loadDatabase();
        db.raw_records = newRawRecords;
        saveDatabase(db);
        
        const records = getMergedRecords(db);
        res.json({
            success: true,
            records: records,
            message: `Successfully uploaded Excel sheet. Parsed ${newRawRecords.length} records.`
        });
    } catch (e) {
        res.status(500).json({ success: false, error: e.message });
    }
});

// Start listening
app.listen(PORT, () => {
    console.log(`Banquet Dashboard server running at http://localhost:${PORT}`);
});
