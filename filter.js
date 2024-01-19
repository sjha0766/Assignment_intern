const xlsx = require("xlsx");
const moment = require("moment");
const fs = require("fs");
const { log } = require("console");

function formatDateTime(dateTime) {
  const unixTimestamp = (dateTime - 25569) * 86400 * 1000;

  const formattedDateTime = moment(unixTimestamp).format(
    "MM-DD-YYYY hh:mm:ss a"
  );

  return formattedDateTime;
}
function formatDate(date) {
  const unixTimestamp = (date - 25569) * 86400 * 1000; // Convert Excel date to Unix timestamp

  // Format date and time using moment.js
  const formatDate = moment(unixTimestamp).format("MM-DD-YYYY");

  return formatDate;
}
const filterData = (url) => {
  console.log("Reading and displaying Excel data");
  who_Has_Worked_Seven_Days(url);
  who_Have_Less_Than_Ten_Hours(url);
  who_Have_Worked_More_Than_Fourteen_Hours(url);
};

const excelData = (rowData) => {
  const workbook = xlsx.readFile(rowData);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(sheet, { raw: true });

  // Format date and time in the data
  data.forEach((row) => {
    if (row.Time) {
      row.Time = formatDateTime(row.Time);
    }
  });

  data.forEach((row) => {
    if (row["Time Out"]) {
      row["Time Out"] = formatDateTime(row["Time Out"]);
    }
  });

  data.forEach((row) => {
    if (row["Pay Cycle Start Date"]) {
      row["Pay Cycle Start Date"] = formatDate(row["Pay Cycle Start Date"]);
    }
  });

  data.forEach((row) => {
    if (row["Pay Cycle End Date"]) {
      row["Pay Cycle End Date"] = formatDate(row["Pay Cycle End Date"]);
    }
  });

  data.forEach((row) => {
    if (row.Pay) {
      row.Pay = formatDateTime(row.Pay);
    }
  });
  return data;
};

const who_Has_Worked_Seven_Days = (filePath) => {
  const data = excelData(filePath);

  const fData = data.slice(1);

  const employees = {};

  fData.forEach((row) => {
    const employeeName = row["Employee Name"];
    const position = row["Position ID"];

    if (!employees[employeeName]) {
      employees[employeeName] = {
        position: position,
        daysWorked: 1,
      };
    } else {
      // Check if the position is the same as the previous day
      if (employees[employeeName].position === position) {
        employees[employeeName].daysWorked += 1;

        // Check if the employee worked for 7 consecutive days
        if (employees[employeeName].daysWorked === 7) {
          const outputFile = "results.json";
          let results = [];
          try {
            const existingData = fs.readFileSync(outputFile, "utf8");
            results = JSON.parse(existingData);
          } catch (err) {}

          fData.forEach((row) => {
            const employeeName = row["Employee Name"];
            const position = row["Position ID"];

            if (
              !results.some(
                (result) =>
                  result.employeeName === employeeName &&
                  result.position === position
              )
            ) {
              const existingEmployee = results.find(
                (result) => result.employeeName === employeeName
              );

              if (existingEmployee) {
                existingEmployee.daysWorked += 1;
              }
            }
          });
          const outputData = {
            [employeeName]: {
              position: position,
              daysWorked: 7,
            },
          };
          results.push(outputData);
          fs.writeFileSync(outputFile, JSON.stringify(results, null, 2));
        }
      } else {
        employees[employeeName] = {
          position: position,
          daysWorked: 1,
        };
      }
    }
  });
  console.log(`Results saved to results.json`);
};

const who_Have_Less_Than_Ten_Hours = (filePath) => {
  const data = excelData(filePath);
  const fData = data.slice(1);

  // Output array for employees with less than 10 hours between shifts
  const results = [];

  for (let index = 0; index < fData.length - 1; index++) {
    const currentRow = fData[index];
    const nextRow = fData[index + 1];

    const employeeName = currentRow["Employee Name"];
    const position = currentRow["Position ID"];

    // Parse time from 'MM-DD-YYYY hh:mm:ss a' format using moment.js
    const currentTime = moment(currentRow.Time, "MM-DD-YYYY hh:mm:ss a");
    const nextTime = moment(nextRow.Time, "MM-DD-YYYY hh:mm:ss a");

    const timeDifference = nextTime.diff(currentTime, "hours", true); // Difference in hours
    if (timeDifference > 1 && timeDifference < 10) {
      const result = {
        "Position ID": position,
        "Position Status": currentRow["Position Status"],
        "Employee Name": employeeName,
        TimeDiff: timeDifference,
      };
      results.push(result);
    }
  }

  // Write all results to the JSON file
  const outputFile = "shift_results.json";
  fs.writeFileSync(outputFile, JSON.stringify(results, null, 2));
  console.log(`Results saved to ${outputFile}`);
};

const who_Have_Worked_More_Than_Fourteen_Hours = (filePath) => {
  const data = excelData(filePath);
  const fData = data.slice(1);

  // Output array for employees with less than 14 hours between shifts
  const results = [];

  for (let index = 0; index < fData.length - 1; index++) {
    const currentRow = fData[index];
    const nextRow = fData[index + 1];

    const employeeName = currentRow["Employee Name"];
    const position = currentRow["Position ID"];

    // Parse time from  using moment.js
    const currentTime = moment(currentRow.Time, "MM-DD-YYYY hh:mm:ss a");
    const nextTime = moment(nextRow.Time, "MM-DD-YYYY hh:mm:ss a");

    const timeDifference = nextTime.diff(currentTime, "hours", true); // Difference in hours
    if (timeDifference > 14) {
      const result = {
        "Position ID": position,
        "Position Status": currentRow["Position Status"],
        "Employee Name": employeeName,
        TimeDiff: timeDifference,
      };
      results.push(result);
    }
  }

  // Write all results in the JSON file
  const outputFile = "long_shift_results.json";
  fs.writeFileSync(outputFile, JSON.stringify(results, null, 2));
  console.log(`Results saved to ${outputFile}`);
};

module.exports = filterData; //export the filterData method
