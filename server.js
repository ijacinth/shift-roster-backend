import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import OpenAI from "openai";
import ExcelJS from "exceljs";
import multer from "multer";
import XLSX from "xlsx";

dotenv.config();

const app = express();
app.use(cors());
app.use(express.json());

const upload = multer();

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

app.post("/generate", upload.single("file"), async (req, res) => {
  try {
    const message = req.body.message || "";
    let fileContent = "";

    // 🔹 Handle file upload
    if (req.file) {
      const fileName = req.file.originalname;

      // ✅ If Excel file
      if (fileName.endsWith(".xlsx")) {
        const workbook = XLSX.read(req.file.buffer, { type: "buffer" });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];

        const jsonData = XLSX.utils.sheet_to_json(sheet);

        fileContent = JSON.stringify(jsonData, null, 2);
      }

      // ✅ If CSV or text
      else {
        fileContent = req.file.buffer.toString("utf-8");
      }
    }

    const prompt = `
You are a shift roster assistant.

User request:
${message}

Uploaded data:
${fileContent}

Instructions:
- Use uploaded data if provided
- Generate fair and balanced schedules
- Avoid overlapping shifts for same employee

If user asks for Excel:
Return ONLY JSON:
{
  "roster": [
    {"employee": "John", "day": "Monday", "shift": "9AM-5PM"}
  ]
}

Otherwise return formatted text.
`;

    const response = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages: [{ role: "user", content: prompt }]
    });

    const aiText = response.choices[0].message.content;

    let data;

    try {
      data = JSON.parse(aiText);
    } catch {
      data = null;
    }

    // 🔥 Excel output
    if (data && data.roster) {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet("Roster");

      sheet.columns = [
        { header: "Employee", key: "employee", width: 20 },
        { header: "Day", key: "day", width: 15 },
        { header: "Shift", key: "shift", width: 20 }
      ];

      data.roster.forEach(r => sheet.addRow(r));

      const buffer = await workbook.xlsx.writeBuffer();

      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader(
        "Content-Disposition",
        "attachment; filename=roster.xlsx"
      );

      return res.send(buffer);
    }

    res.json({ output: aiText });

  } catch (error) {
    console.error(error);
    res.status(500).json({ error: error.message });
  }
});

const PORT = 5000;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});