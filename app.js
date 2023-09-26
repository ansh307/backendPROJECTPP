const express = require("express");
const app = express();
const dotenv = require("dotenv");
// const mongoose = require("mongoose");
const cors = require("cors");
const ExcelJS = require("exceljs");
const XLSX = require("xlsx");
const fileUpload = require("express-fileupload");

dotenv.config();

// async function connectToDatabase() {
//   try {
//     await mongoose.connect(
//       process.env.DATABASE.replace("<PASSWORD>", process.env.DATABASE_PASSWORD)
//     );
//     console.log("DB connection successful");
//   } catch (error) {
//     console.error("DB connection error:", error);
//   }
// }
// connectToDatabase();

app.set("view engine", "ejs");
app.use(express.json());
app.use(express.urlencoded({ extended: true }));
app.use(express.static("public"));
app.use(cors());
app.use(fileUpload());

let uploadedData = []; // Store the data from the uploaded Excel file
app.post("/api/upload", async (req, res) => {
  try {
    if (!req.files || !req.files.excelFile) {
      return res.status(400).send("No file was uploaded.");
    }

    const excelFile = req.files.excelFile;

    // Detect file format (XLS or XLSX) by reading the file content
    let buffer = excelFile.data;
    const extension = excelFile.name.split(".").pop().toLowerCase();

    if (extension === "xls") {
      // Convert XLS to XLSX
      const xlsWorkbook = XLSX.read(buffer, { type: "buffer" });
      const xlsxData = XLSX.write(xlsWorkbook, { bookType: "xlsx", type: "buffer" });
      buffer = xlsxData;
    }

    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(buffer);

    const worksheet = workbook.worksheets[0];
    const data = [];

    const headerRow = worksheet.getRow(1);
    const headerValues = headerRow.values;

    // Map column names to their respective indices
    const columnIndices = {};

    for (let i = 1; i <= headerValues.length; i++) {
      columnIndices[headerValues[i - 1]] = i;
    }

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber !== 1) {
        const rowData = {};

        // Loop through all columns and include them in rowData
        for (const columnName in columnIndices) {
          rowData[columnName] = row.getCell(columnIndices[columnName]).value;
        }

        data.push(rowData);
      }
    });

    if (data.length === 0) {
      return res.status(400).send("No data found in the Excel file.");
    }

    // Store the data from the uploaded Excel file
    const uploadedData = data;

    // Send a response with JSON content type and status 200
    res.status(200).json(uploadedData);
  } catch (error) {
    console.error("Error handling Excel file:", error);
    res.status(500).send("Internal Server Error");
  }
});


app.post("/api/submit", (req, res) => {
  try {
    const selectedNames = req.body.selectedNames;
    let modifiedSelectedNames = selectedNames;

    // If only one name is selected, convert it to an array
    if (!Array.isArray(modifiedSelectedNames)) {
      modifiedSelectedNames = [modifiedSelectedNames];
    }

    // Check if no names are selected
    if (!modifiedSelectedNames || modifiedSelectedNames.length === 0) {
      return res.status(400).send("Select at least one name to download.");
    }

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Selected Names");

    // Add the column titles to the worksheet
    worksheet.addRow([
      "ID",
      "Name",
      "Enrollment No.",
      "Roll no.",
      "College",
      "Branch",
      "Year",
      "Contact No.",
      "E Mail ID",
    ]);

    // Add the selected data to the worksheet
    modifiedSelectedNames.forEach((nameData) => {
      const rowData = JSON.parse(nameData);
      const row = [
        rowData["ID"],
        rowData["Name"],
        rowData["Enrollment No."],
        rowData["Roll no."],
        rowData["College"],
        rowData["Branch"],
        rowData["Year"],
        rowData["Contact No."],
        rowData["E Mail ID"],
      ];
      worksheet.addRow(row);
    });

    const filename = `selected_names_${Date.now()}.xlsx`;

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);

    workbook.xlsx
      .write(res)
      .then(() => {
        res.end();
      })
      .catch((error) => {
        console.error("Error creating Excel file:", error);
        res.status(500).send("Error creating Excel file.");
      });
  } catch (error) {
    res.end("Atleast select one ,  cannot download empty excel file.");
  }
});

app.get("/api/getData", (req, res) => {
  const searchQuery = req.query.search;

  try {
    const filteredData = searchQuery
      ? uploadedData.filter((item) => {
          return item.Name.includes(searchQuery);
        })
      : uploadedData;
    console.log(filteredData);
    res.json(filteredData);
  } catch (error) {
    console.error("Error in /api/getData:", error);
    res.status(500).send("Internal Server Error");
  }
});

const PORT = process.env.PORT || 30035;

app.listen(PORT, () => {
  console.log("Listening on port 3002....");
});
