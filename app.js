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
app.post("/api/upload", (req, res) => {
  if (!req.files || !req.files.excelFile) {
    return res.status(400).send("No file was uploaded.");
  }

  const excelFile = req.files.excelFile;

  // Check the file extension to determine if it's XLS
  const extension = excelFile.name.split(".").pop().toLowerCase();

  if (extension === "xls") {
    try {
      // Convert XLS to XLSX
      const xlsData = excelFile.data;
      const xlsWorkbook = XLSX.read(xlsData, { type: "buffer" });
      const xlsxData = XLSX.write(xlsWorkbook, {
        bookType: "xlsx",
        type: "buffer",
      });

      // Replace the original file with the XLSX version
      excelFile.data = xlsxData;
      excelFile.name = "converted.xlsx";
    } catch (error) {
      console.error("Error converting XLS to XLSX:", error);
      return res.status(500).send("Error converting XLS to XLSX.");
    }
  }

  // Now excelFile should contain XLSX format, continue processing

  const workbook = new ExcelJS.Workbook();
  workbook.xlsx
    .load(excelFile.data)
    .then((workbook) => {
      const worksheet = workbook.worksheets[0];
      const data = [];

      // Assuming the header row contains the column names "fffs," "Name," "Enrollment No.," "Branch," and "E Mail ID"
      const headerRow = worksheet.getRow(1);
      const nameCol = headerRow.getCell(2); // Adjust the column number as needed
      const enrollmentNoCol = headerRow.getCell(3); // Adjust the column number as needed
      const branchCol = headerRow.getCell(4); // Adjust the column number as needed
      const emailCol = headerRow.getCell(6); // Adjust the column number as needed

      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber !== 1) {
          // Skip the header row
          const rowData = {
            Name: row.getCell(nameCol).value,
            "Enrollment No.": row.getCell(enrollmentNoCol).value,
            Branch: row.getCell(branchCol).value,
            "E Mail ID": row.getCell(emailCol).value,
          };
          data.push(rowData);
        }
      });

      // Store the data from the uploaded Excel file
      uploadedData = data;

      // Send a response
      res.json(data);
    })
    .catch((error) => {
      console.error("Error parsing Excel file:", error);
      res.status(500).send("Error parsing Excel file.");
    });
});


app.post("/api/submit", (req, res) => {
  try{
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
    })
}
catch(error) {
  res.end("Atleast select one ,  cannot download empty excel file.");
}});




app.get("/api/getData", (req, res) => {
  const searchQuery = req.query.search;

  try {
    const filteredData = searchQuery
      ? uploadedData.filter((item) => {
          return (
            item.Name.includes(searchQuery) 
          
          );
        })
      : uploadedData;
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
 