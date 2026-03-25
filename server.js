import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import OpenAI from "openai";
import ExcelJS from "exceljs";
import multer from "multer";

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

    // 🔹 Read uploaded file
    if (req.file) {
      fileContent = req.file.buffer.toString("utf-8");
    }

    const prompt = `
You are a shift roster assistant.

User request:
${message}

Uploaded file content:
${fileContent}

If user asks for Excel or download:
Return ONLY JSON in this format:
{
  "roster": [
    {"employee": "John", "day": "Monday", "shift": "9AM-5PM"}
  ]
}

Otherwise return formatted text.
`;

    const response = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      messages: [
        { role: "user", content: prompt }
      ]
    });

    const aiText = response.choices[0].message.content;

    let data;

    try {
      data = JSON.parse(aiText);
    } catch {
      data = null;
    }

    // 🔥 Excel generation
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