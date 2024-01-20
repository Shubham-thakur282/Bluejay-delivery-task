const xlsx = require("xlsx");
const fs = require("fs");

let workBook = xlsx.readFile(__dirname + "/Assignment_Timecard.xlsx"); 
let worksheet = workBook.Sheets[workBook.SheetNames[0]];

// The below two lines are used to find total number of rows in the excel with !ref property
const range = worksheet["!ref"];
const rows = range ? range.split(":")[1].match(/\d+/)[0] : 0;


const filePath ='output.txt';

// Below function is used to change the exceldate to the understandable date and then format that date
function excelDateTimeToJsDate(excelDateTime) {
  const millisecondsPerDay = 24 * 60 * 60 * 1000;
  const daysSinceExcelEpoch = excelDateTime - 25569;
  const millisecondsSinceExcelEpoch = daysSinceExcelEpoch * millisecondsPerDay;
  return new Date(millisecondsSinceExcelEpoch);
}

function formatJsDate(jsDate) {
  const options = {
    year: "numeric",
    month: "numeric",
    day: "numeric",
    hour: "numeric",
    minute: "numeric",
    second: "numeric",
    timeZone: "UTC",
  };
  return jsDate.toLocaleString("en-US", options);
}

// this function used to use above both function at one place to get the readable date from excel
function formatDate(dateString) {
  if (dateString === "") {
    return 0;
  }
  let jsDate = excelDateTimeToJsDate(dateString);

  let readableFormat = formatJsDate(jsDate);
  return readableFormat;
}


// this function is used to get the day from the readable date
function getDay(cellValue) {
  if (cellValue === "") {
    return 0;
  }
  let day;
  let readableFormat = formatDate(cellValue);
  let date = new Date(readableFormat);
  day = date.getDate();

  return day;
}


// this function is used to find the employees who have worked for consecutive 7 days
function getNameConsecutiveDays() {
  let line = "\n\n\nOutPut for the employees who have worked for consecutive 7 days"
  fs.appendFile(filePath, line+'\n\n', 'utf8', err => {
    if (err) {
      throw err;
    }
    console.log('Line appended to file');
  });

  let countDays = 1;
  for (let i = 2; i < rows; i++) {
    if (i === 2) {
      countDays + getDay(worksheet[`C${i}`].v);
    } else if (i === rows) {
      countDays + getDay(worksheet[`C${i}`].v);
    } else {
      if (worksheet[`A${i - 1}`].v === worksheet[`A${i}`].v) {
        let prevDay = getDay(worksheet[`C${i - 1}`].v);
        let presentDay = getDay(worksheet[`C${i}`].v);
        if (prevDay + 1 === presentDay) {
          countDays += 1;
        } else if (prevDay === presentDay) {
          continue;
        } else {
          if (prevDay + 1 !== presentDay && prevDay !== presentDay) {
            countDays = 1;
          }
        }
      } else {
        countDays = 1;
      }
    }
    if (countDays === 7) {
      let output = `The employee with ID: ${worksheet[`A${i}`].v}, Name ${
        worksheet[`H${i}`].v
      } has worked for consecutive & 7 days`
      
      fs.appendFile(filePath, output+"\n", 'utf8', err => {
        if (err) {
          throw err;
        }
        console.log('Line appended to file');
      });
    }
  }
}
getNameConsecutiveDays();

// this gives time in hour and minute so that we can use them to find difference
function getHourAndMinut(dateString) {
  var parts = dateString.split(" ");

  var time = parts[1];

  return time;
}

// this function is used to find diffrence between two times when between the shifts
function getTimeDifference(time1, time2) {
  const [hours1, minutes1] = time1.split(":").map(Number);
  const [hours2, minutes2] = time2.split(":").map(Number);

  let totalHours = hours1 - hours2;
  let totalMinutes = minutes1 - minutes2;

  if (totalMinutes < 0) {
    totalHours -= 1;
    totalMinutes += 60;
  }

  const formattedResult = `${String(Math.abs(totalHours)).padStart(
    2,
    "0"
  )}:${String(totalMinutes).padStart(2, "0")}`;
  return formattedResult;
}


// this function is used to change hour and minute to minutes so that we can use them in checking if that is certain hours
function timeStringToMinutes(timeString) {
  const [hours, minutes] = timeString.split(":").map(Number);
  const totalMinutes = hours * 60 + minutes;

  return totalMinutes;
}


// this function give the list of employess who have less than 10 hours of time between shifts but greater than 1 hour
function shiftA() {
  let shiftTimeGap, timeBefore, timeAfter, hourBefore, hourAfter;
  let line = "\n\n\nOutPut for who have less than 10 hours of time between shifts but greater than 1 hour"
  fs.appendFile(filePath, line+'\n\n', 'utf8', err => {
    if (err) {
      throw err;
    }
    console.log('Line appended to file');
  });

  for (let i = 2; i < rows; i++) {
    timeAfter = formatDate(worksheet[`C${i + 1}`].v);
    timeBefore = formatDate(worksheet[`D${i}`].v);
    if (i === 2) {
      if (worksheet[`D${i}`].v === "") {
        continue;
      }
      timeAfter = formatDate(worksheet[`C${i + 1}`].v);
      timeBefore = formatDate(worksheet[`D${i}`].v);
      if (worksheet[`A${i + 1}`].v === worksheet[`A${i}`].v) {
        if (getDay(worksheet[`C${i + 1}`].v) === getDay(worksheet[`D${i}`].v)) {
          hourBefore = getHourAndMinut(timeBefore);
          hourAfter = getHourAndMinut(timeAfter);
          shiftTimeGap = getTimeDifference(hourAfter, hourBefore);
        }
      }
    } else if (i === rows) {
      break;
    } else {
      if (worksheet[`D${i}`].v === "") {
        continue;
      }
      if (worksheet[`C${i}`].v === "") {
        continue;
      }
      timeAfter = formatDate(worksheet[`C${i + 1}`].v);
      timeBefore = formatDate(worksheet[`D${i}`].v);
      if (worksheet[`A${i + 1}`].v === worksheet[`A${i}`].v) {
        if (getDay(worksheet[`C${i + 1}`].v) === getDay(worksheet[`D${i}`].v)) {
          hourBefore = getHourAndMinut(timeBefore);
          hourAfter = getHourAndMinut(timeAfter);
          shiftTimeGap = getTimeDifference(hourAfter, hourBefore);
        }
      }
    }

    let totalMinutes = timeStringToMinutes(shiftTimeGap);

    if (totalMinutes > 1 * 60 && totalMinutes < 10 * 60) {

      output = `The employee with ID: ${worksheet[`A${i}`].v}, Name ${worksheet[`H${i}`].v} have ${shiftTimeGap} time between the shifts`
      fs.appendFile(filePath, output+'\n', 'utf8', err => {
        if (err) {
          throw err;
        }
        console.log('Line appended to file');
      });
    }
  }
}
shiftA();


// this fucntion give employess Who has worked for more than 14 hours in a single shift
function shiftB() {
  let line = "\n\n\nOutPut for employess Who has worked for more than 14 hours in a single shift"
  fs.appendFile(filePath, line+'\n\n', 'utf8', err => {
    if (err) {
      throw err;
    }
    console.log('Line appended to file');
  });
  for (let i = 2; i < rows; i++) {
    if (worksheet[`E${i}`].v === "") {
      continue;
    }
    if (timeStringToMinutes(worksheet[`E${i}`].v) >= 14 * 60) {
        let output =`The employee with ID: ${worksheet[`A${i}`].v}, Name ${worksheet[`H${i}`].v} has worked ${worksheet[`E${i}`].v}`
        
        fs.appendFile(filePath, output+"\n", 'utf8', err => {
          if (err) {
            throw err;
          }
          console.log('Line appended to file');
        });
    }
  }
}

shiftB();
