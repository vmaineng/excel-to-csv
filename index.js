import { program } from "commander";
import XLSX from "xlsx";
import fs from "fs";
import path from "path";

program
  .requiredOption("-i, --input <file>", "input Excel file (XLSX or XLS)")
  .option("-o, --output <file>", "output CSV file", "output.csv")
  .option("-s, --sheet <name>", "name of the worksheet to use", "Sheet 1")
  .option("--headered", "does the Excel sheet have a header row?", true);

program.parse(process.argv);
const options = program.opts();

(async () => {
  try {
    console.log(`Reading from ${options.input}...`);

    const workbook = XLSX.readFile(options.input);

    console.log(`Available sheets: ${workbook.SheetNames.join(", ")}`);

    if (!workbook.SheetNames.includes(options.sheet)) {
      throw new Error(
        `Sheet ${
          options.sheet
        } not found in workbook.Available sheets: ${workbook.SheetNames.join(
          ", "
        )}`
      );
    }

    const worksheet = workbook.Sheets[options.sheet];

    let data = XLSX.utils.sheet_to_json(worksheet, {
      header: options.headered ? 1 : undefined,
    });

    let csvContent = "";

    if (options.headered) {
      const headers = data[0];
      const nameIndex = headers.findIndex((header) =>
        header.toLowerCase().includes("name")
      );
      const amountIndex = headers.findIndex((header) =>
        header.toLowerCase().includes("amount")
      );

      if (nameIndex === -1 || amountIndex === -1) {
        throw new Error("Could not find 'name' or 'amount' columns in header.");
      }
      csvContent += "name,amount\n";

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (row[nameIndex] !== undefined && row[amountIndex] !== undefined) {
          csvContent += `${row[nameIndex]},${row[amountIndex]}\n`;
        }
      }
    } else {
      csvContent += "name,amount\n";
      for (const row of data) {
        if (row[0] !== undefined && row[1] !== undefined) {
          csvContent += `${row[0]},${row[1]}\n`;
        }
      }
    }

    fs.writeFileSync(options.output, csvContent);
    console.log(`CSV written to ${options.output}`);
  } catch (error) {
    console.error("Error:", error.message);
    process.exit(1);
  }
})();
