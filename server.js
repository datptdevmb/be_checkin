const express = require("express");
const cors = require("cors");
const xlsx = require("xlsx");
const path = require("path");

const app = express();
const PORT = 3001;

app.use(cors());
app.use(express.json());

const filePath = path.join(__dirname, "data/employees.xlsx");
const outputPath = path.join(__dirname, "data/employees_checkin.xlsx");
const sheetName = "Sheet1";

const workbook = xlsx.readFile(filePath);
const sheet = workbook.Sheets[sheetName];
const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, range: 2 });

const header = rows[0];
const dataRows = rows.slice(1);

const employees = {};
dataRows.forEach((row) => {
    const rowData = {};
    header.forEach((key, index) => {
        if (typeof key === "string") {
            rowData[key.trim()] = row[index];
        }
    });

    const id = rowData["Mã NV"]?.toString().trim();
    if (id) {
        employees[id] = {
            name: rowData["Họ và tên"],
            unit: rowData["Phòng ban"],
            team: rowData["ĐỘI"],
            phone: rowData["Điện thoại"],
            checkedIn: false,
        };
    }
});

// 🔐 Biến cờ an toàn
let isWriting = false;

// ✅ Hàm ghi dữ liệu an toàn
function saveToExcel() {
    if (isWriting) return; // Tránh ghi trùng lặp
    isWriting = true;

    try {
        const updatedRows = [header.concat("Checkin")];
        Object.entries(employees).forEach(([code, info], index) => {
            updatedRows.push([
                index + 1,
                code,
                info.name,
                info.unit,
                info.team,
                info.phone,
                info.checkedIn ? "TRUE" : "FALSE",
            ]);
        });

        const newSheet = xlsx.utils.aoa_to_sheet(updatedRows);
        const newBook = xlsx.utils.book_new();
        xlsx.utils.book_append_sheet(newBook, newSheet, sheetName);
        xlsx.writeFile(newBook, outputPath);
        console.log("💾 Ghi file check-in thành công.");
    } catch (err) {
        console.error("❌ Lỗi khi ghi file:", err.message);
    } finally {
        isWriting = false;
    }
}

// ⏲️ Auto backup mỗi 8 giây
setInterval(saveToExcel, 8000);

// 📌 API tra cứu
app.get("/api/employees/:id", (req, res) => {
    const id = req.params.id.trim().toUpperCase();
    const emp = employees[id];
    if (emp) {
        res.json({ success: true, data: emp });
    } else {
        res.status(404).json({ success: false, message: "Không tìm thấy mã nhân viên." });
    }
});

// 📋 Chưa check-in
app.get("/api/employee/unchecked", (req, res) => {
    const unchecked = Object.entries(employees)
        .filter(([_, emp]) => !emp.checkedIn)
        .map(([id, emp]) => ({ id, name: emp.name, unit: emp.unit, team: emp.team, phone: emp.phone }));

    res.json({
        success: true,
        message: `Còn ${unchecked.length} người chưa check-in.`,
        count: unchecked.length,
        data: unchecked
    });
});

// ✅ Đã check-in
app.get("/api/employee/checked", (req, res) => {
    const checked = Object.entries(employees)
        .filter(([_, emp]) => emp.checkedIn)
        .map(([id, emp]) => ({ id, name: emp.name, unit: emp.unit, team: emp.team, phone: emp.phone }));

    res.json({ success: true, count: checked.length, data: checked });
});

// 🔘 Check-in
app.get("/api/checkin/:id", (req, res) => {
    const id = req.params.id.trim().toUpperCase();

    if (!id || !employees[id]) {
        return res.status(400).json({ success: false, message: "Mã nhân viên không hợp lệ." });
    }

    if (employees[id].checkedIn) {
        return res.json({ success: true, alreadyChecked: true, message: "Đã check-in trước đó." });
    }

    employees[id].checkedIn = true;

    saveToExcel(); // Ghi ngay khi có checkin
    res.json({ success: true, message: "✅ Check-in thành công." });
});

app.listen(PORT, () => {
    console.log(`✅ Server đang chạy tại http://localhost:${PORT}`);
});
