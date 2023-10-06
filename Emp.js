const XLSX = require('xlsx');
const fs = require('fs');
const workbook = XLSX.readFile('employee_data.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const employeeData = new Map();
const ONE_HOUR = 60 * 60 * 1000; 
const ONE_DAY = 24 * 60 * 60 * 1000; 
let currentEmployeeName = null; 
for (const cellAddress in sheet) {
  if (cellAddress[0] === '!') continue; 
  const cell = sheet[cellAddress];
  const positionId = sheet['A' + cellAddress.substring(1)].v;
  const positionStatus = sheet['B' + cellAddress.substring(1)].v;
  const timeIn = new Date(sheet['C' + cellAddress.substring(1)].w);
  const timeOut = new Date(sheet['D' + cellAddress.substring(1)].w);
  const timecardHours = parseFloat(sheet['E' + cellAddress.substring(1)].v);
  const employeeName = sheet['G' + cellAddress.substring(1)].v;
  const fileNumber = sheet['H' + cellAddress.substring(1)].v;
  if (!employeeData.has(employeeName)) {
    employeeData.set(employeeName, {
      shifts: [],
      consecutiveDays: 0, 
    });
  }
 if (employeeName !== currentEmployeeName) {
    currentEmployeeName = employeeName;
employeeData.get(employeeName).consecutiveDays = 1;
    if (employeeData.get(employeeName).consecutiveDays === 7) {
      console.log(`Employee ${employeeName} has worked for 7 consecutive days`);
    }
   if (timecardHours > 14) {
      console.log(`Employee ${employeeName} has worked for more than 14 hours in a single shift`);
    }
 if (
    employeeData.get(employeeName).shifts.length > 0 &&
    timeIn - employeeData.get(employeeName).shifts.slice(-1)[0].timeOut <= ONE_DAY
  ) {
    employeeData.get(employeeName).consecutiveDays++;
  }
  if (employeeData.get(employeeName).consecutiveDays === 7) {
    console.log(`Employee ${employeeName} has worked for 7 consecutive days`);
  }

  if (employeeData.get(employeeName).shifts.length > 0) {
    const previousShift = employeeData.get(employeeName).shifts.slice(-1)[0];
    const timeBetweenShifts = timeIn - previousShift.timeOut;

    if (timeBetweenShifts < 10 * ONE_HOUR && timeBetweenShifts > ONE_HOUR) {
      console.log(`Employee ${employeeName} has less than 10 hours between shifts`);
    }
  }

  if (timecardHours > 14) {
    console.log(`Employee ${employeeName} has worked for more than 14 hours in a single shift`);
  }
  employeeData.get(employeeName).shifts.push({
    positionId,
    positionStatus,
    timeIn,
    timeOut,
  });
}
