const path = require("path");
require("dotenv").config({ path: path.resolve(__dirname, "../.env") });
const express = require("express");
const XLSX = require("xlsx");
const fs = require("fs");
const excel4node = require("excel4node");
const dbMysql = require("../src/config/dbMysql");
const dbMssql = require("../src/config/dbMssql");
const app = express();

app.use(express.json());

app.post("/convert-to-excel", (req, res) => {
  const data = req.body;

  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet([[data.value, data.date]]);

  XLSX.utils.book_append_sheet(workbook, worksheet, "Data");

  const fileName = `data-${new Date().toISOString()}.xlsx`;

  XLSX.writeFile(workbook, fileName, { type: "file" });

  res.sendFile(fileName, { root: __dirname }, (error) => {
    if (error) {
      console.error(error);
      res.status(500).send("Error sending Excel file");
    }
  });
});

app.post("/download", async (req, res) => {
  const data = req.body

  res.json(data);


  // const testJsonToXlsx = async () => {
  //   const workbook = new excel4node.Workbook();
  //   const worksheet = workbook.addWorksheet("Relátorio");

  //   const resultSQL = await dbMssql.query(process.env.SELECT_BASE, {
  //     type: dbMssql.QueryTypes.SELECT,
  //   });
  //   const headers = Object.keys(resultSQL[0]);
  //   headers.forEach((header, index) => {
  //     worksheet.cell(1, index + 1).string(header);
  //   });
  //   resultSQL.forEach((data, dataIndex) => {
  //     headers.forEach((header, headerIndex) => {
  //       const value = data[header];
  //       if (typeof value === "number" || typeof value === "object") {
  //         worksheet
  //           .cell(dataIndex + 2, headerIndex + 1)
  //           .string(JSON.stringify(value));
  //       } else {
  //         worksheet.cell(dataIndex + 2, headerIndex + 1).string(value);
  //       }
  //     });
  //   });
  //   res.setHeader(
  //     "Content-Type",
  //     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  //   );
  //   res.setHeader("Content-Disposition", "attachment; filename=data.xlsx");

  //   workbook.write("data.xlsx", res);

  //   console.log(resultSQL);
  // };
  // testJsonToXlsx();
});

app.get("/", (req, res) => {
  res.json("Server DataSql to Excel");
});
app.listen(process.env.PORT, () => {
  console.log(`Server 2 listening at http://localhost:${process.env.PORT}`);
});
