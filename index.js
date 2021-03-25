const fs = require("fs");
const config = require("./config");
const xl = require("excel4node");


const source = fs.readFileSync("./data.json", { encoding: "utf8", flag: "r" });

const wb = new xl.Workbook();
const ws = wb.addWorksheet("Promo");

const headingColumnNames = [
  "Promo ID",
  "Promo Name",
  "Blast Schedule",
  "Requestor",
  "Customer Type",
  "Status",
  "Total Click",
  "Total Customer",
];

try {
  const raw = JSON.parse(source);

  let headingColumnIndex = 1;
  headingColumnNames.forEach((heading) => {
    ws.cell(1, headingColumnIndex++).string(heading);
  });

  let rowIndex = 2;
  raw.data.promo.forEach((record) => {
    const selected = [
      "promo_id",
      "promo_name",
      "start_notification",
      "requestor_name",
      "customer_type",
      "status",
      "total_click_notification",
      "total_customer",
    ];

    let columnIndex = 1;
    selected.forEach((columnName) => {
      if (typeof record[columnName] === "string")
        ws.cell(rowIndex, columnIndex++).string(record[columnName]);
      else ws.cell(rowIndex, columnIndex++).number(record[columnName]);
    });
    rowIndex++;
  });
  wb.write(`Promo_${new Date().toLocaleDateString().replace(/\//g, "-")}.xlsx`);
} catch (error) {
  console.log(error);
}
