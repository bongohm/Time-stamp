let MySheets  = SpreadsheetApp.getActiveSpreadsheet();
let EmployeeSheet  = MySheets.getSheetByName("รายชื่อพนักงาน");  

function doGet(e) {
    return HtmlService.createTemplateFromFile("index").evaluate();
}

function Login() {
    event.preventDefault();  // Prevent default form submission
    var uid = document.getElementById("uid").value;
    var pass = document.getElementById("pass").value;
    google.script.run.withSuccessHandler(ReturnMsg).LoginCheck(uid, pass);
}

function LoginCheck(pUID, pPassword) {
    // Searching for the user ID in column 3
    let ReturnData = EmployeeSheet.getRange("C:C").createTextFinder(pUID).matchEntireCell(true).findAll();
    let StartRow = 0;
    ReturnData.forEach(function(range) {
        StartRow = range.getRow();
    });

    if (StartRow > 0) {
        // Assuming the password is in column 4
        let TmpPass = EmployeeSheet.getRange(StartRow, 4).getValue();
        if (TmpPass == pPassword) {
            var employeeName = EmployeeSheet.getRange(StartRow, 2).getValue(); // Assuming the employee's name is in column 2
            var employeeID = EmployeeSheet.getRange(StartRow, 3).getValue(); // User ID is in column 3
            return { success: true, employeeName: employeeName, employeeID: employeeID };
        }
    }

    return { success: false }; // Return false if no match is found
}

function OpenPage(PageName) {
    return HtmlService.createHtmlOutputFromFile(PageName).getContent();
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index');
}

function getEmployeeData(employeeID) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("รายชื่อพนักงาน");
  var data = sheet.getDataRange().getValues();

  for (var i = 0; i < data.length; i++) {
    // ตรวจสอบว่าข้อมูลในคอลัมน์ที่สาม (index 2) ตรงกับ employeeID ที่ได้รับหรือไม่
    if (data[i][2] == employeeID) { 
      // คืนค่าข้อมูลพนักงานจากแถวนั้นๆ
      return data[i]; // Returns the entire row containing the employee data
    }
  }
  return null; // ถ้าไม่พบ employeeID
}

function registerUser(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("รายชื่อพนักงาน");
  if (!sheet) {
    sheet = ss.insertSheet("รายชื่อพนักงาน");
  }
  
  // บันทึกข้อมูลในแถวใหม่
  var lastRow = sheet.getLastRow(); // หาแถวสุดท้ายที่มีข้อมูล
  var newRow = lastRow + 1; // ตั้งค่าแถวสำหรับข้อมูลใหม่
  sheet.appendRow([data.fullname, data.pin, data.id, data.password]);
  
  // ตั้งค่าขนาดตัวอักษรสำหรับแถวใหม่
  var range = sheet.getRange(newRow, 1, 1, 4); // ระบุช่วงเซลล์ที่จะตั้งค่า (แถวใหม่, คอลัมน์แรก, 1 แถว, 4 คอลัมน์)
  range.setFontSize(12); // ตั้งค่าขนาดตัวอักษรเป็น 12
}

function getemployeePassword() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var employeeSheet = ss.getSheetByName("รายชื่อพนักงาน");
  var getLastRow = employeeSheet.getLastRow();
  return employeeSheet.getRange(2, 2, getLastRow - 1).getValues();
}

function updatePassword(employeeID, newPassword) {
  var ss = SpreadsheetApp.openById('17CftuNYzF0H6TGk-6z5p3rV5BV3j59j2IBI_w8RcIHU');
  var sheet = ss.getSheetByName('รายชื่อพนักงาน');
  var data = sheet.getDataRange().getValues();

  for (var i = 0; i < data.length; i++) {
    if (data[i][2] == employeeID) { // ตรวจสอบ ID ของพนักงานในคอลัม C
      data[i][3] = newPassword; // อัพเดทรหัสผ่านใหม่ในคอลัม D
      break;
    }
  }

  sheet.getRange(1, 1, data.length, data[0].length).setValues(data); // อัพเดทข้อมูลในชีท
}

function clockInWithPin(pin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("ลงเวลาทำงาน"); // "Work Time Log" sheet
  var employeeSheet = ss.getSheetByName("รายชื่อพนักงาน"); // "Employee List" sheet
  var employeeData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow(), 5).getValues();
  var returnDate = '';
  var error = 'SUCCESS';
  var returnArray = [];
  var employee = '';
  var pinMatched = false;

  // Validate PIN and get employee name
  employeeData.forEach(function(row) {
    if (pin === row[1].toString()) { // Change to use column 2 (index 1)
      employee = row[0]; // Employee name is in column 1
      pinMatched = true;
    }
  });

  if (!pinMatched) {
    error = 'PIN ไม่ถูกต้อง';
    returnArray.push([error, '', '']);
    return returnArray;
  }

  // Check if already clocked in without clocking out
  var clockInData = mainSheet.getRange(2, 1, mainSheet.getLastRow(), 3).getValues();
  var alreadyClockedIn = clockInData.some(row => row[0] === employee && row[2] === '');

  if (alreadyClockedIn) {
    error = 'ต้องตอกบัตรออกก่อนตอกบัตรเข้า.';
    returnArray.push([error, '', employee]);
    return returnArray;
  }

  // Record the clock-in
  var newDate = new Date();
  returnDate = Utilities.formatDate(newDate, "GMT+7", "dd/MM/yyyy HH:mm:ss");
  
  // Use returnDate which is the formatted date string instead of newDate
  mainSheet.appendRow([employee, returnDate, '', '']);

  // Notify via LINE
  var message = "พนักงาน " + employee + "\nลงเวลาเข้างาน: " + returnDate;
  sendLineNotify(message); // Make sure sendLineNotify function is defined

  // Format the newly added row
  var lastRow = mainSheet.getLastRow();
  var range = mainSheet.getRange(lastRow, 1, 1, mainSheet.getLastColumn());
  range.setFontSize(12).setHorizontalAlignment('center');

  returnArray.push([error, returnDate, employee]);
  return returnArray;
}

function clockOutWithPin(pin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("ลงเวลาทำงาน"); // "Work Time Log" sheet
  var employeeSheet = ss.getSheetByName("รายชื่อพนักงาน"); // "Employee List" sheet
  var employeeData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow() - 1, 5).getValues();
  var newDate = new Date();
  var returnDate = Utilities.formatDate(newDate, "GMT+7", "dd/MM/yyyy HH:mm:ss");
  var error = 'SUCCESS';
  var foundRecord = false;
  var returnArray = [];
  var employeeName = '';

  // Validate PIN and get employee name
  employeeData.some(function(row) {
    if (pin === row[1].toString()) { // Change to use column 2 (index 1)
      employeeName = row[0]; // Employee name is in column 1
      return true; // Found the employee, stop the search
    }
    return false; // Continue the search
  });

  if (!employeeName) {
    error = 'PIN ไม่ถูกต้อง';
    returnArray.push([error, '', '']);
    return returnArray;
  }

  // Retrieve clock-in records and process clock-out
  var clockInRecords = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 4).getValues();
  clockInRecords.some(function(record, index) {
    if (record[0] === employeeName && record[2] === '') { // Check for clock-in without clock-out
      var rowIndex = index + 2; // Adjusting for header row and 1-indexed
      mainSheet.getRange(rowIndex, 3).setValue(newDate).setNumberFormat("dd/MM/yyyy HH:mm:ss").setHorizontalAlignment("center").setFontSize(12);
      
      var clockInTime = record[1];
      var totalTimeMinutes = Math.floor((newDate - clockInTime) / 60000);
      mainSheet.getRange(rowIndex, 4).setValue(totalTimeMinutes).setHorizontalAlignment("center").setFontSize(12);
      
      foundRecord = true;

      // Notify via LINE
      var message = "พนักงาน " + employeeName + "\nลงเวลาออกงาน: " + returnDate;
      sendLineNotify(message); // Make sure sendLineNotify function is defined

      return true; // Found and processed clock-out, stop the search
    }
    return false; // Continue the search
  });

  if (!foundRecord) {
    error = 'ต้องตอกบัตรเข้าก่อนตอกบัตรออก.';
    returnArray.push([error, '', employeeName]);
    return returnArray;
  }

  // Optionally calculate total hours worked
  TotalHours1(); // Uncomment if TotalHours1() is defined and needed here

  returnArray.push([error, returnDate, employeeName]);
  return returnArray;
}

function break1clockInWithPin(pin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("ลงเวลาพักครั้งที่1"); // "Break 1 Time Log" sheet
  var employeeSheet = ss.getSheetByName("รายชื่อพนักงาน"); // "Employee List" sheet
  var employeeData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow(), 5).getValues();
  var returnDate = '';
  var error = 'SUCCESS';
  var returnArray = [];
  var employee = '';
  var pinMatched = false;

  // Validate PIN and get employee name
  employeeData.forEach(function(row) {
    if (pin === row[1].toString()) { // Use column 2 (index 1) for PIN
      employee = row[0]; // Employee name is in column 1
      pinMatched = true;
    }
  });

  if (!pinMatched) {
    error = 'PIN ไม่ถูกต้อง';
    returnArray.push([error, '', '']);
    return returnArray;
  }

  // Check if already clocked in without clocking out
  var clockInData = mainSheet.getRange(2, 1, mainSheet.getLastRow(), 3).getValues();
  var alreadyClockedIn = clockInData.some(row => row[0] === employee && row[2] === '');

  if (alreadyClockedIn) {
    error = 'ต้องตอกบัตรออกก่อนตอกบัตรเข้า.';
    returnArray.push([error, '', employee]);
    return returnArray;
  }

  // Record the clock-in
  var newDate = new Date();
  returnDate = Utilities.formatDate(newDate, "GMT+7", "dd/MM/yyyy HH:mm:ss");
  mainSheet.appendRow([employee, newDate, '', '']);

  // Notify via LINE
  var message = "พนักงาน " + employee + "\nลงเวลาพักครั้งที่1: " + returnDate;
  sendLineNotify(message);

  // Format the newly added row
  var lastRow = mainSheet.getLastRow();
  var range = mainSheet.getRange(lastRow, 1, 1, mainSheet.getLastColumn());
  range.setFontSize(12).setHorizontalAlignment('center');

  returnArray.push([error, returnDate, employee]);
  return returnArray;
}

function break1clockOutWithPin(pin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("ลงเวลาพักครั้งที่1"); // "Break 1 Time Log" sheet
  var employeeSheet = ss.getSheetByName("รายชื่อพนักงาน"); // "Employee List" sheet
  var employeeData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow() - 1, 5).getValues();
  var newDate = new Date();
  var returnDate = Utilities.formatDate(newDate, "GMT+7", "dd/MM/yyyy HH:mm:ss");
  var error = 'SUCCESS';
  var foundRecord = false;
  var returnArray = [];
  var employeeName = '';

  // Validate PIN and get employee name
  employeeData.some(function(row) {
    if (pin === row[1].toString()) { // Use column 2 (index 1) for PIN
      employeeName = row[0]; // Employee name is in column 1
      return true; // Found the employee, stop the search
    }
    return false; // Continue the search
  });

  if (!employeeName) {
    error = 'PIN ไม่ถูกต้อง';
    returnArray.push([error, '', '']);
    return returnArray;
  }

  // Retrieve clock-in records and process clock-out
  var clockInRecords = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 4).getValues();
  clockInRecords.some(function(record, index) {
    if (record[0] === employeeName && record[2] === '') { // Check for clock-in without clock-out
      var rowIndex = index + 2; // Adjusting for header row and 1-indexed
      mainSheet.getRange(rowIndex, 3).setValue(newDate).setNumberFormat("dd/MM/yyyy HH:mm:ss").setHorizontalAlignment("center").setFontSize(12);
      
      var clockInTime = record[1];
      var totalTimeMinutes = Math.floor((newDate - clockInTime) / 60000);
      mainSheet.getRange(rowIndex, 4).setValue(totalTimeMinutes).setHorizontalAlignment("center").setFontSize(12);
      
      foundRecord = true;

      // Notify via LINE
      var message = "พนักงาน " + employeeName + "\nลงเวลากลับจากพักครั้งที่1: " + returnDate;
      sendLineNotify(message);

      return true; // Found and processed clock-out, stop the search
    }
    return false; // Continue the search
  });

  if (!foundRecord) {
    error = 'ต้องตอกบัตรเข้าก่อนตอกบัตรออก.';
    returnArray.push([error, '', employeeName]);
    return returnArray;
  }

  // Optionally calculate total hours worked
  TotalHours2(); // Uncomment if TotalHours1() is defined and needed here

  returnArray.push([error, returnDate, employeeName]);
  return returnArray;
}


function break2clockInWithPin(pin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("ลงเวลาพักครั้งที่2"); // "Break 2 Time Log" sheet
  var employeeSheet = ss.getSheetByName("รายชื่อพนักงาน"); // "Employee List" sheet
  var employeeData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow(), 5).getValues();
  var returnDate = '';
  var error = 'SUCCESS';
  var returnArray = [];
  var employee = '';
  var pinMatched = false;

  // Validate PIN and get employee name
  employeeData.forEach(function(row) {
    if (pin === row[1].toString()) { // Use column 2 (index 1) for PIN
      employee = row[0]; // Employee name is in column 1
      pinMatched = true;
    }
  });

  if (!pinMatched) {
    error = 'PIN ไม่ถูกต้อง';
    returnArray.push([error, '', '']);
    return returnArray;
  }

  // Check if already clocked in without clocking out
  var clockInData = mainSheet.getRange(2, 1, mainSheet.getLastRow(), 3).getValues();
  var alreadyClockedIn = clockInData.some(row => row[0] === employee && row[2] === '');

  if (alreadyClockedIn) {
    error = 'ต้องตอกบัตรออกก่อนตอกบัตรเข้า.';
    returnArray.push([error, '', employee]);
    return returnArray;
  }

  // Record the clock-in
  var newDate = new Date();
  returnDate = Utilities.formatDate(newDate, "GMT+7", "dd/MM/yyyy HH:mm:ss");
  mainSheet.appendRow([employee, newDate, '', '']);

  // Notify via LINE
  var message = "พนักงาน " + employee + "\nลงเวลาพักครั้งที่2: " + returnDate;
  sendLineNotify(message);

  // Format the newly added row
  var lastRow = mainSheet.getLastRow();
  var range = mainSheet.getRange(lastRow, 1, 1, mainSheet.getLastColumn());
  range.setFontSize(12).setHorizontalAlignment('center');

  returnArray.push([error, returnDate, employee]);
  return returnArray;
}

function break2clockOutWithPin(pin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("ลงเวลาพักครั้งที่2"); // "Break 2 Time Log" sheet
  var employeeSheet = ss.getSheetByName("รายชื่อพนักงาน"); // "Employee List" sheet
  var employeeData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow() - 1, 5).getValues();
  var newDate = new Date();
  var returnDate = Utilities.formatDate(newDate, "GMT+7", "dd/MM/yyyy HH:mm:ss");
  var error = 'SUCCESS';
  var foundRecord = false;
  var returnArray = [];
  var employeeName = '';

  // Validate PIN and get employee name
  employeeData.some(function(row) {
    if (pin === row[1].toString()) { // Use column 2 (index 1) for PIN
      employeeName = row[0]; // Employee name is in column 1
      return true; // Found the employee, stop the search
    }
    return false; // Continue the search
  });

  if (!employeeName) {
    error = 'PIN ไม่ถูกต้อง';
    returnArray.push([error, '', '']);
    return returnArray;
  }

  // Retrieve clock-in records and process clock-out
  var clockInRecords = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 4).getValues();
  clockInRecords.some(function(record, index) {
    if (record[0] === employeeName && record[2] === '') { // Check for clock-in without clock-out
      var rowIndex = index + 2; // Adjusting for header row and 1-indexed
      mainSheet.getRange(rowIndex, 3).setValue(newDate).setNumberFormat("dd/MM/yyyy HH:mm:ss").setHorizontalAlignment("center").setFontSize(12);
      
      var clockInTime = record[1];
      var totalTimeMinutes = Math.floor((newDate - clockInTime) / 60000);
      mainSheet.getRange(rowIndex, 4).setValue(totalTimeMinutes).setHorizontalAlignment("center").setFontSize(12);
      
      foundRecord = true;

      // Notify via LINE
      var message = "พนักงาน " + employeeName + "\nลงเวลากลับจากพักครั้งที่2: " + returnDate;
      sendLineNotify(message);

      return true; // Found and processed clock-out, stop the search
    }
    return false; // Continue the search
  });

  if (!foundRecord) {
    error = 'ต้องตอกบัตรเข้าก่อนตอกบัตรออก.';
    returnArray.push([error, '', employeeName]);
    return returnArray;
  }

  // Optionally calculate total hours worked
  TotalHours3(); // Uncomment if TotalHours1() is defined and needed here

  returnArray.push([error, returnDate, employeeName]);
  return returnArray;
}

function bathroomclockInWithPin(pin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("ลงเวลาเข้าห้องน้ำ"); // "Bathroom Time Log" sheet
  var employeeSheet = ss.getSheetByName("รายชื่อพนักงาน"); // "Employee List" sheet
  var employeeData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow(), 5).getValues();
  var returnDate = '';
  var error = 'SUCCESS';
  var returnArray = [];
  var employee = '';
  var pinMatched = false;

  // Validate PIN and get employee name
  employeeData.forEach(function(row) {
    if (pin === row[1].toString()) { // Use column 2 (index 1) for PIN
      employee = row[0]; // Employee name is in column 1
      pinMatched = true;
    }
  });

  if (!pinMatched) {
    error = 'PIN ไม่ถูกต้อง';
    returnArray.push([error, '', '']);
    return returnArray;
  }

  // Check if already clocked in without clocking out
  var clockInData = mainSheet.getRange(2, 1, mainSheet.getLastRow(), 3).getValues();
  var alreadyClockedIn = clockInData.some(row => row[0] === employee && row[2] === '');

  if (alreadyClockedIn) {
    error = 'ต้องตอกบัตรออกก่อนตอกบัตรเข้า.';
    returnArray.push([error, '', employee]);
    return returnArray;
  }

  // Record the clock-in
  var newDate = new Date();
  returnDate = Utilities.formatDate(newDate, "GMT+7", "dd/MM/yyyy HH:mm:ss");
  mainSheet.appendRow([employee, newDate, '', '']);

  // Notify via LINE
  var message = "พนักงาน " + employee + "\nลงเวลาเข้าห้องน้ำ: " + returnDate;
  sendLineNotify(message);

  // Format the newly added row
  var lastRow = mainSheet.getLastRow();
  var range = mainSheet.getRange(lastRow, 1, 1, mainSheet.getLastColumn());
  range.setFontSize(12).setHorizontalAlignment('center');

  returnArray.push([error, returnDate, employee]);
  return returnArray;
}

function bathroomclockOutWithPin(pin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("ลงเวลาเข้าห้องน้ำ"); // "Bathroom Time Log" sheet
  var employeeSheet = ss.getSheetByName("รายชื่อพนักงาน"); // "Employee List" sheet
  var employeeData = employeeSheet.getRange(2, 1, employeeSheet.getLastRow() - 1, 5).getValues();
  var newDate = new Date();
  var returnDate = Utilities.formatDate(newDate, "GMT+7", "dd/MM/yyyy HH:mm:ss");
  var error = 'SUCCESS';
  var foundRecord = false;
  var returnArray = [];
  var employeeName = '';

  // Validate PIN and get employee name
  employeeData.some(function(row) {
    if (pin === row[1].toString()) { // Use column 2 (index 1) for PIN
      employeeName = row[0]; // Employee name is in column 1
      return true; // Found the employee, stop the search
    }
    return false; // Continue the search
  });

  if (!employeeName) {
    error = 'PIN ไม่ถูกต้อง';
    returnArray.push([error, '', '']);
    return returnArray;
  }

  // Retrieve clock-in records and process clock-out
  var clockInRecords = mainSheet.getRange(2, 1, mainSheet.getLastRow() - 1, 4).getValues();
  clockInRecords.some(function(record, index) {
    if (record[0] === employeeName && record[2] === '') { // Check for clock-in without clock-out
      var rowIndex = index + 2; // Adjusting for header row and 1-indexed
      mainSheet.getRange(rowIndex, 3).setValue(newDate).setNumberFormat("dd/MM/yyyy HH:mm:ss").setHorizontalAlignment("center").setFontSize(12);
      
      var clockInTime = record[1];
      var totalTimeMinutes = Math.floor((newDate - clockInTime) / 60000);
      mainSheet.getRange(rowIndex, 4).setValue(totalTimeMinutes).setHorizontalAlignment("center").setFontSize(12);
      
      foundRecord = true;

      // Notify via LINE
      var message = "พนักงาน " + employeeName + "\nลงเวลากลับจากเข้าห้องน้ำ: " + returnDate;
      sendLineNotify(message);

      return true; // Found and processed clock-out, stop the search
    }
    return false; // Continue the search
  });

  if (!foundRecord) {
    error = 'ต้องตอกบัตรเข้าก่อนตอกบัตรออก.';
    returnArray.push([error, '', employeeName]);
    return returnArray;
  }

  // Optionally calculate total hours worked
  TotalHours4(); // Uncomment if TotalHours1() is defined and needed here

  returnArray.push([error, returnDate, employeeName]);
  return returnArray;
}

function businessclockInWithPinAndNote(pin, note) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("ลงเวลาทำธุระ");
  var employeeSheet = ss.getSheetByName("รายชื่อพนักงาน");
  var lastRow = mainSheet.getLastRow();
  var employeeLastRow = employeeSheet.getLastRow();
  var return_date = '';
  var error = 'SUCCESS';
  var return_array = [];
  var employee = '';

  // Validate PIN and get employee name
  for (var k = 2; k <= employeeLastRow; k++) {
    if (pin == employeeSheet.getRange(k, 2).getValue()) { // Use column 2 (index 1) for PIN
      employee = employeeSheet.getRange(k, 1).getValue(); // Use column 1 (index 0) for employee name
      break;
    }
  }

  if (!employee) {
    error = 'PIN ไม่ถูกต้อง';
    return_array.push([error, '', '']);
    return return_array;
  }

  // Check if the employee has already clocked in without clocking out
  for (var j = 2; j <= lastRow; j++) {
    if (employee == mainSheet.getRange(j, 1).getValue() && mainSheet.getRange(j, 3).getValue() == '') {
      error = 'ต้องตอกบัตรออกก่อนตอกบัตรเข้า';
      return_array.push([error, return_date, employee]);
      return return_array;
    }
  }

  // Clock in the employee
  var new_date = new Date();
  return_date = Utilities.formatDate(new_date, "GMT+7", "dd/MM/yyyy HH:mm:ss");
  mainSheet.appendRow([employee, new_date, '', '', note]); // Append the new clock in record with note in the 5th column

  // Send LINE Notify with the note included
  var message = "พนักงาน " + employee + "\nได้ไปทำธุระ: " + return_date + ". หมายเหตุ: " + note;
  sendLineNotify(message);

  // Update lastRow after appending
  lastRow = mainSheet.getLastRow();

  // Set font size and alignment for the newly added row
  var range = mainSheet.getRange(lastRow, 1, 1, mainSheet.getLastColumn());
  range.setFontSize(12);
  range.setHorizontalAlignment('center');

  return_array.push([error, return_date, employee]);
  return return_array;
}

function businessclockOutWithPin(pin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("ลงเวลาทำธุระ");
  var employeeSheet = ss.getSheetByName("รายชื่อพนักงาน");
  var lastRow = mainSheet.getLastRow();
  var employeeLastRow = employeeSheet.getLastRow();
  var new_date = new Date();
  var return_date = Utilities.formatDate(new_date, "GMT+7", "dd/MM/yyyy HH:mm:ss");
  var error = 'SUCCESS';
  var foundRecord = false;
  var return_array = [];
  var employee = '';

  // Validate PIN and get employee name
  var pinMatched = false;
  for (var k = 2; k <= employeeLastRow; k++) {
    if (pin == employeeSheet.getRange(k, 2).getValue()) { // Use column 2 (index 1) for PIN
      pinMatched = true;
      employee = employeeSheet.getRange(k, 1).getValue(); // Use column 1 (index 0) for employee name
      break;
    }
  }
  
  if (!pinMatched) {
    error = 'PIN ไม่ถูกต้อง';
    return_array.push([error, '', employee]);
    return return_array;
  }

  // Check if the employee has clocked in and not yet clocked out
  for (var j = 2; j <= lastRow; j++) {
    if (employee == mainSheet.getRange(j, 1).getValue() && mainSheet.getRange(j, 3).getValue() == '') {
      mainSheet.getRange(j, 3).setValue(new_date)
               .setNumberFormat("dd/MM/yyyy HH:mm:ss")
               .setHorizontalAlignment("center")
               .setFontSize(12);
      var clockInTime = mainSheet.getRange(j, 2).getValue();
      var totalTimeMinutes = Math.floor((new_date - clockInTime) / 60000);
      mainSheet.getRange(j, 4).setValue(totalTimeMinutes)
               .setHorizontalAlignment("center")
               .setFontSize(12);  
      foundRecord = true;

      // Send LINE Notify
      var message = "พนักงาน " + employee + "\nได้กลับจากทำธุระ: " + return_date;
      sendLineNotify(message);

      break;
    }
  }
  
  if (!foundRecord) {
    error = 'ต้องตอกบัตรเข้าก่อนตอกบัตรออก';
    return_array.push([error, '', employee]);
    return return_array; 
  }

  TotalHours5(); // Uncomment if TotalHours1() is defined and needed here

  return_array.push([error, return_date, employee]);
  return return_array;
}

function addZero(i) {
  if (i < 10) {
    i = "0" + i;
  }
  return i;
}

function getDate(date_in) {
  var currentDate = date_in;
  var currentMonth = addZero(currentDate.getMonth() + 1);
  var currentYear = currentDate.getFullYear();
  var currentHours = addZero(currentDate.getHours());
  var currentMinutes = addZero(currentDate.getMinutes());
  var currentSeconds = addZero(currentDate.getSeconds());

  var date = currentMonth + '/' + currentDate.getDate() + '/' + currentYear + ' ' +
             currentHours + ':' + currentMinutes.toString() + ':' + currentSeconds.toString();
  return date;
}

function TotalHours1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("ลงเวลาทำงาน");
  var lastRow = mainSheet.getLastRow();
  var totals = [];
  var currentDate = new Date(); // ตัวแปรสำหรับวันที่ปัจจุบัน

  for (var j = 2; j <= lastRow; j++) {
    var rate = mainSheet.getRange(j, 4).getValue();
    var name = mainSheet.getRange(j, 1).getValue();
    var foundRecord = false;
    
    for(var i = 0; i < totals.length; i++) {
      if(name == totals[i][0] && rate != '') {
        totals[i][1] = totals[i][1] + rate;
        foundRecord = true;
      }
    }
    
    if(foundRecord == false && rate != '') {
      totals.push([name, rate]);
    }
  }
  
  mainSheet.getRange("F5:H10000").clear(); // ปรับเป็น H1000 เพื่อล้างคอลัมน์วันที่ด้วย
  
  for(var i = 0; i < totals.length; i++) {
    mainSheet.getRange(2+i, 6).setValue(totals[i][0]).setFontSize(12);
    mainSheet.getRange(2+i, 7).setValue(totals[i][1]).setFontSize(12);
    mainSheet.getRange(2+i, 8).setValue(currentDate).setFontSize(12); // บันทึกวันที่ปัจจุบันในคอลัมน์ที่ 8
  }
}

function TotalHours2() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("ลงเวลาพักครั้งที่1");
  var lastRow = mainSheet.getLastRow();
  var totals = [];
  var currentDate = new Date(); // ตัวแปรสำหรับวันที่ปัจจุบัน

  for (var j = 2; j <= lastRow; j++) {
    var rate = mainSheet.getRange(j, 4).getValue();
    var name = mainSheet.getRange(j, 1).getValue();
    var foundRecord = false;
    
    for(var i = 0; i < totals.length; i++) {
      if(name == totals[i][0] && rate != '') {
        totals[i][1] = totals[i][1] + rate;
        foundRecord = true;
      }
    }
    
    if(foundRecord == false && rate != '') {
      totals.push([name, rate]);
    }
  }
  
  mainSheet.getRange("F5:H10000").clear(); // ปรับเป็น H1000 เพื่อล้างคอลัมน์วันที่ด้วย
  
  for(var i = 0; i < totals.length; i++) {
    mainSheet.getRange(2+i, 6).setValue(totals[i][0]).setFontSize(12);
    mainSheet.getRange(2+i, 7).setValue(totals[i][1]).setFontSize(12);
    mainSheet.getRange(2+i, 8).setValue(currentDate).setFontSize(12); // บันทึกวันที่ปัจจุบันในคอลัมน์ที่ 8
  }
}

function TotalHours3() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("ลงเวลาพักครั้งที่2");
  var lastRow = mainSheet.getLastRow();
  var totals = [];
  var currentDate = new Date(); // ตัวแปรสำหรับวันที่ปัจจุบัน

  for (var j = 2; j <= lastRow; j++) {
    var rate = mainSheet.getRange(j, 4).getValue();
    var name = mainSheet.getRange(j, 1).getValue();
    var foundRecord = false;
    
    for(var i = 0; i < totals.length; i++) {
      if(name == totals[i][0] && rate != '') {
        totals[i][1] = totals[i][1] + rate;
        foundRecord = true;
      }
    }
    
    if(foundRecord == false && rate != '') {
      totals.push([name, rate]);
    }
  }
  
  mainSheet.getRange("F5:H10000").clear(); // ปรับเป็น H1000 เพื่อล้างคอลัมน์วันที่ด้วย
  
  for(var i = 0; i < totals.length; i++) {
    mainSheet.getRange(2+i, 6).setValue(totals[i][0]).setFontSize(12);
    mainSheet.getRange(2+i, 7).setValue(totals[i][1]).setFontSize(12);
    mainSheet.getRange(2+i, 8).setValue(currentDate).setFontSize(12); // บันทึกวันที่ปัจจุบันในคอลัมน์ที่ 8
  }
}

function TotalHours4() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("ลงเวลาเข้าห้องน้ำ");
  var lastRow = mainSheet.getLastRow();
  var totals = [];
  var currentDate = new Date(); // ตัวแปรสำหรับวันที่ปัจจุบัน

  for (var j = 2; j <= lastRow; j++) {
    var rate = mainSheet.getRange(j, 4).getValue();
    var name = mainSheet.getRange(j, 1).getValue();
    var foundRecord = false;
    
    for(var i = 0; i < totals.length; i++) {
      if(name == totals[i][0] && rate != '') {
        totals[i][1] = totals[i][1] + rate;
        foundRecord = true;
      }
    }
    
    if(foundRecord == false && rate != '') {
      totals.push([name, rate]);
    }
  }
  
  mainSheet.getRange("F5:H10000").clear(); // ปรับเป็น H1000 เพื่อล้างคอลัมน์วันที่ด้วย
  
  for(var i = 0; i < totals.length; i++) {
    mainSheet.getRange(2+i, 6).setValue(totals[i][0]).setFontSize(12);
    mainSheet.getRange(2+i, 7).setValue(totals[i][1]).setFontSize(12);
    mainSheet.getRange(2+i, 8).setValue(currentDate).setFontSize(12); // บันทึกวันที่ปัจจุบันในคอลัมน์ที่ 8
  }
}

function TotalHours5() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = ss.getSheetByName("ลงเวลาทำธุระ");
  var lastRow = mainSheet.getLastRow();
  var totals = [];
  var currentDate = new Date(); // ตัวแปรสำหรับวันที่ปัจจุบัน

  for (var j = 2; j <= lastRow; j++) {
    var rate = mainSheet.getRange(j, 4).getValue();
    var name = mainSheet.getRange(j, 1).getValue();
    var foundRecord = false;
    
    for(var i = 0; i < totals.length; i++) {
      if(name == totals[i][0] && rate != '') {
        totals[i][1] = totals[i][1] + rate;
        foundRecord = true;
      }
    }
    
    if(foundRecord == false && rate != '') {
      totals.push([name, rate]);
    }
  }
  
  mainSheet.getRange("F5:H10000").clear(); // ปรับเป็น H1000 เพื่อล้างคอลัมน์วันที่ด้วย
  
  for(var i = 0; i < totals.length; i++) {
    mainSheet.getRange(2+i, 6).setValue(totals[i][0]).setFontSize(12);
    mainSheet.getRange(2+i, 7).setValue(totals[i][1]).setFontSize(12);
    mainSheet.getRange(2+i, 8).setValue(currentDate).setFontSize(12); // บันทึกวันที่ปัจจุบันในคอลัมน์ที่ 8
  }
}

function sendLineNotify(message) {
  var token = "o6SwBzH8ENms3Uxkf8rewLpLR3wYKcOMYnDBP9em5OH"; // แทนที่ด้วย Token ของคุณ
  var options = {
    "method": "post",
    "headers": {
      "Authorization": "Bearer " + token
    },
    "payload": {
      "message": message
    },
    "muteHttpExceptions": true
  };
  UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
}


// ไฟลที่จะ Backup
//https://docs.google.com/spreadsheets/d/1ARG2JOl4KKK535Pv6-imOqUk5FCZ-gODPFAthFzxZao/edit#gid=0
const srcGsId = "1ARG2JOl4KKK535Pv6-imOqUk5FCZ-gODPFAthFzxZao"
// โฟล์เดอร์ปลายทาง
const monthFolders = {
  '01': '1FdFkkUS53faJDy4YgjwANaAb8Cib6lsO',
  '02': '1tkiwfXDZFYKnRv1Xfj6vjgS2GYgTUeeF',
  '03': '1v2xqbP6EXZP-YuiHJWhESsXKYs2Fc_pY',
  '04': '1cyOyf4d154JSb0N5sKsyvZu3-6vE0FLv',
  '05': '1zuoaJRDGv2sB6tSFLRARHtlE_VpirS-6',
  '06': '1Bus5hcuytK7h2MI4ZkrDXJcwvYSAI0Sr',
  '07': '1f8mvb69oZP7bV7ZGUyghawbzdgo-dONR',
  '08': '1cLa30IRkqsXJXodaNdScvxax11pvcizd',
  '09': '15K9aoBs21-3v7aHZNZduGACXBnWxF6Oh',
  '10': '12v4F2DEv4qJ1LQ8d9CjXCfdqVxsZMnsf',
  '11': '1hHXfdh5_EM_Ncb9n0MGmrZA7rnP_c2-A',
  '12': '1Ox4kXeVIq_z9OUu1dWekO4jrPwFRsbSP'
};

function backupGsFile() {
  const currentDate = new Date();
  const currentMonth = Utilities.formatDate(currentDate, "GMT+7", "MM"); // Use "MM" for month format
  const dstFolderId = monthFolders[currentMonth]; // Select folder based on current month
  const backupDateTime = Utilities.formatDate(currentDate, "GMT+7", "yyyy-MM-dd_HH:mm:ss");
  const ssSrcToBackup = SpreadsheetApp.openById(srcGsId); 
  const ssCopyOfBackup = ssSrcToBackup.copy(backupDateTime + "_" + ssSrcToBackup.getName());
  
  const driveCopyOfBackup = DriveApp.getFileById(ssCopyOfBackup.getId()); 
  const dstFolder = DriveApp.getFolderById(dstFolderId); 
  driveCopyOfBackup.moveTo(dstFolder);

  // Call to clear the original spreadsheet after backup
  clearOriginalSpreadsheet();
}

//=====================================================
//
function createTimeDrivenTriggers() {
  // Delete existing triggers
  const allTriggers = ScriptApp.getProjectTriggers();
  allTriggers.forEach(tg => ScriptApp.deleteTrigger(tg));

  // Create a new trigger for 08:59 AM
  ScriptApp.newTrigger('backupGsFile')
           .timeBased()
           .everyDays(1)
           .atHour(8)
           .nearMinute(59)
           .create();

  // Create another new trigger for 20:59 PM
  ScriptApp.newTrigger('backupGsFile')
           .timeBased()
           .everyDays(1)
           .atHour(20)
           .nearMinute(59)
           .create();
}

function clearOriginalSpreadsheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets(); // Get all sheets in the spreadsheet

  sheets.forEach(function(sheet) {
    // Skip the sheet named "ลงเวลาเข้างาน"
    if (sheet.getName() !== "รายชื่อพนักงาน") {
      // Clear the entire sheet except A1:G1
      sheet.getRange('A2:H1000').clear(); // Adjust the range as needed
    }
  });
} 
