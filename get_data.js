import * as XLSX from "xlsx";
import path from "path";
import { promises as fs } from "fs";

export default async function handler(req, res) {
  try {
    const filePath = path.join(process.cwd(), "dbtbkksewa excel.xlsx");
    const fileBuffer = await fs.readFile(filePath);

    const workbook = XLSX.read(fileBuffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

    const formattedData = jsonData.map((row) => {
      const newRow = {};
      for (const key in row) {
        let val = row[key];
        if (key.toLowerCase().includes("tanggal") && typeof val === "number") {
          const date = XLSX.SSF.parse_date_code(val);
          if (date) {
            val = `${String(date.d).padStart(2, "0")}/${String(date.m).padStart(
              2,
              "0"
            )}/${date.y}`;
          }
        }
        if (typeof val === "string" && val.match(/^[0-9,.]+$/)) {
          val = parseFloat(val.replace(/,/g, ""));
        }
        newRow[key.toLowerCase().replace(/\s+/g, "_")] = val;
      }
      return newRow;
    });

    res.status(200).json(formattedData);
  } catch (error) {
    console.error("Error:", error);
    res.status(500).json({
      error: "Gagal membaca file Excel",
      detail: error.message,
    });
  }
}
