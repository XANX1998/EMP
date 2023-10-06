const XLSX = require('xlsx');
const fs = require('fs');

// Load the Excel file
const workbook = XLSX.readFile('employee_data.xlsx');
const sheet = workbook.Sheets[workbook.SheetNames[0]];

// Define a map to store employee data
const employeeData = new Map();

// Define constants for time intervals
const ONE_HOUR = 60 * 60 * 1000; // 1 hour in milliseconds
const ONE_DAY = 24 * 60 * 60 * 1000; // 1 day in milliseconds;

let currentEmployeeName = null; // Track the current employee being processed

// Iterate through rows in the sheet
for (const cellAddress in sheet) {
  if (cellAddress[0] === '!') continue; // Skip non-cell entries

  const cell = sheet[cellAddress];

  // Extract relevant cell values
  const positionId = sheet['A' + cellAddress.substring(1)].v;
  const positionStatus = sheet['B' + cellAddress.substring(1)].v;
  const timeIn = new Date(sheet['C' + cellAddress.substring(1)].w);
  const timeOut = new Date(sheet['D' + cellAddress.substring(1)].w);
  const timecardHours = parseFloat(sheet['E' + cellAddress.substring(1)].v);
  const employeeName = sheet['G' + cellAddress.substring(1)].v;
  const fileNumber = sheet['H' + cellAddress.substring(1)].v;

  // Initialize the employee's data if it doesn't exist
  if (!employeeData.has(employeeName)) {
    employeeData.set(employeeName, {
      shifts: [],
      consecutiveDays: 0, // Start with 0 consecutive days
    });
  }

  // Check if a new employee is being processed
  if (employeeName !== currentEmployeeName) {
    currentEmployeeName = employeeName;

    // Reset consecutive days for the new employee
    employeeData.get(employeeName).consecutiveDays = 1;

    // Check conditions and print employee data
    if (employeeData.get(employeeName).consecutiveDays === 7) {
      console.log(`Employee ${employeeName} has worked for 7 consecutive days`);
    }

    if (timecardHours > 14) {
      console.log(`Employee ${employeeName} has worked for more than 14 hours in a single shift`);
    }
  }

  // Check for consecutive days
  if (
    employeeData.get(employeeName).shifts.length > 0 &&
    timeIn - employeeData.get(employeeName).shifts.slice(-1)[0].timeOut <= ONE_DAY
  ) {
    employeeData.get(employeeName).consecutiveDays++;
  }

  // Check conditions and print employee data
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

  // Add the shift to the employee's data
  employeeData.get(employeeName).shifts.push({
    positionId,
    positionStatus,
    timeIn,
    timeOut,
  });
}

// This map contains the data you need, and you can further process it as required.
