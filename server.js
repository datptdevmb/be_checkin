const express = require("express");
const cors = require("cors");
const xlsx = require("xlsx");
const path = require("path");
const fs = require("fs");


const app = express();
const PORT = 3001;

app.use(cors());
app.use(express.json());

const filePath = path.join(__dirname, "data/employees.xlsx");
const outputPath = path.join(__dirname, "data/employees_checkin.xlsx");
const scoreFilePath = path.join(__dirname, "data/team_scores.xlsx");
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


let isWriting = false;


function saveToExcel() {
    if (isWriting) return;
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


setInterval(saveToExcel, 8000);


app.get("/api/employees/:id", (req, res) => {
    const id = req.params.id.trim().toUpperCase();
    const emp = employees[id];
    if (emp) {
        res.json({ success: true, data: emp });
    } else {
        res.status(404).json({ success: false, message: "Không tìm thấy mã nhân viên." });
    }
});


app.get("/api/employee/unchecked", (req, res) => {
    console.log("🔍 Truy vấn danh sách chưa  check-in");
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

// -------------------- API CHẤM ĐIỂM -------------------- //
function appendScore(teamId, judgeId, scorePart) {
    let rows = [
        [
            "Đội",
            "Giám khảo",
            "📚 Tính mới",
            "📚 Tính khả thi",
            "📚 Tính hiệu quả",
            "📚 Phong cách trình bày",
            "🎯 Phù hợp chủ đề",
            "🎯 Sáng tạo",
            "🎯 Biểu cảm",
            "Tổng điểm",
            "Thời gian"
        ]
    ];

    if (fs.existsSync(scoreFilePath)) {
        const wbOld = xlsx.readFile(scoreFilePath);
        const wsOld = wbOld.Sheets["Scores"];
        const data = xlsx.utils.sheet_to_json(wsOld, { header: 1, defval: "" });

        rows = [data[0], ...data.slice(1).filter(row => !(row[0] == teamId && row[1] == judgeId))];
    }

    const row = [
        teamId,
        judgeId,
        scorePart.part1?.understanding || 0,
        scorePart.part1?.logic || 0,
        scorePart.part1?.expression || 0,
        scorePart.part1?.expression1 || 0,
        scorePart.part2?.teamwork || 0,
        scorePart.part2?.creativity || 0,
        scorePart.part2?.completion || 0,
    ];

    const total = row.slice(2).reduce((a, b) => a + parseFloat(b || 0), 0);
    row.push(total.toFixed(2));
    row.push(new Date().toLocaleString("vi-VN"));

    rows.push(row);

    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.aoa_to_sheet(rows);
    xlsx.utils.book_append_sheet(wb, ws, "Scores");
    xlsx.writeFile(wb, scoreFilePath);

    console.log("✅ Ghi điểm thành công:", { teamId, judgeId });
}

// ✅ Nhận điểm từ client
app.post("/api/score/:teamId/:judgeId", (req, res) => {
    const { teamId, judgeId } = req.params;
    const score = req.body;

    if (!score.part1 && !score.part2) {
        return res.status(400).json({ success: false, message: "Thiếu dữ liệu part1 hoặc part2." });
    }

    try {
        appendScore(teamId, judgeId, score);
        res.json({ success: true, message: "✅ Đã lưu điểm vào Excel." });
    } catch (err) {
        console.error("❌ Lỗi khi ghi điểm:", err.message);
        res.status(500).json({ success: false, message: "Lỗi server khi ghi điểm." });
    }
});

// ✅ API trả danh sách điểm đã chấm
app.get("/api/scores", (req, res) => {
    try {
        if (!fs.existsSync(scoreFilePath)) {
            return res.json({ success: true, data: [] });
        }

        const workbook = xlsx.readFile(scoreFilePath);
        const sheet = workbook.Sheets["Scores"];
        const data = xlsx.utils.sheet_to_json(sheet, { defval: "" });

        res.json({ success: true, data });
    } catch (err) {
        console.error("❌ Lỗi đọc file điểm:", err.message);
        res.status(500).json({ success: false, message: "Lỗi server khi đọc điểm." });
    }
});
app.listen(PORT, () => {
    console.log(`✅ Server đang chạy tại http://localhost:${PORT}`);
});
