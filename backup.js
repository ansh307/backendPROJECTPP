app.post("/api/submit", (req, res) => {
    const selectedNames = req.body.selectedNames;
    console.log(req.body)
    console.log(selectedNames);
    let modifiedSelectedNames = selectedNames;
    const selectedData = [];
    // If only one name is selected, convert it to an array
    if (typeof selectedNames === "string") {
      modifiedSelectedNames = [selectedNames];
    }
  
    // Check if no names are selected
    if (!modifiedSelectedNames || modifiedSelectedNames.length === 0) {
      return res.status(400).send("Select at least one name to download.");
    }
  
    selectedNames.forEach((name) => {
      const foundData = data.find((item) => item["Name"] === name);
      if (foundData) {
        selectedData.push(foundData);
      }
    });
  
    // Create a new Excel workbook
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Selected Names");
  
    const columnTitles = [
      "Select",
      "ID",
      "Name",
      "Enrollment No.",
      "Roll no.",
      "College",
      "Branch",
      "Year",
      "Contact No.",
      "E Mail ID",
    ];
  
    // Add the selected names to the worksheet
     worksheet.addRow(columnTitles);
  
    // Add the selected data to the worksheet
    selectedData.forEach((dataItem) => {
      const row = [];
      columnTitles.forEach((title) => {
        row.push(dataItem[title]);
      });
      worksheet.addRow(row);
    });
    // Generate a unique filename for the Excel file
    const filename = `selected_names_${Date.now()}.xlsx`;
  
    // Set response headers to make it a downloadable file
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename="${filename}"`);
  
    // Send the Excel file to the client
    workbook.xlsx
      .write(res)
      .then(() => {
        res.end();
      })
      .catch((error) => {
        console.error("Error creating Excel file:", error);
        res.status(500).send("Error creating Excel file.");
      });
  });


  <input
  type="text"
  className="custom-input"
  placeholder="Search by Phone Number"
  value={searchQuery}
  onChange={this.handleSearch}
/>