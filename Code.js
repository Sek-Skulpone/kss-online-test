const SPREADSHEET_ID = '1KZicFPjH37Key3fcFgEVqXZULLar-IcV0j9tg0tWOEo'; 

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('ระบบจัดการสอบออนไลน์ | โรงเรียน')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getScriptUrl() { 
  return ScriptApp.getService().getUrl(); 
}

// ล็อกอินผู้ใช้งาน
function login(username, password) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const user = username.toString().trim();
  const pass = password.toString().trim();

  // ตรวจสอบชีตครู/แอดมิน
  const uSheet = ss.getSheetByName("Users_Teachers");
  if (uSheet) {
    const uData = uSheet.getRange("B2:E" + uSheet.getLastRow()).getValues();
    for (let r of uData) {
      if (r[0].toString() === user && r[1].toString() === pass) {
        const isAdmin = String(r[3]).toLowerCase() === "admin";
        return { 
          status: "success", 
          role: "Teacher", 
          name: r[2], 
          displayRole: isAdmin ? "Admin" : "ครูผู้สอน",
          isAdmin: isAdmin 
        };
      }
    }
  }

  // ตรวจสอบชีตนักเรียนตามระดับชั้น
  const levels = ["ม.1", "ม.2", "ม.3", "ม.4", "ม.5", "ม.6"];
  for (let lv of levels) {
    const sSheet = ss.getSheetByName("Students " + lv);
    if (!sSheet) continue;
    const sData = sSheet.getRange("A2:G" + sSheet.getLastRow()).getValues(); 
    for (let r of sData) {
      if (r[1].toString() === user && r[2].toString() === pass) {
        return { 
          status: "success", role: "Student", sid: r[1].toString(), 
          name: r[3], level: lv, no: r[6], room: r[5], displayRole: "นักเรียน"
        };
      }
    }
  }
  return { status: "fail" };
}

// ดึงข้อมูลตารางสอบ และเช็คว่าทำข้อสอบตัวไหนเสร็จแล้วบ้าง
function getExamData(level, studentId) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const eSheet = ss.getSheetByName(level);
  if (!eSheet) return { exams: [], finished: [] };

  const lastRow = eSheet.getLastRow();
  if (lastRow < 2) return { exams: [], finished: [] };
  
  const eData = eSheet.getRange(2, 1, lastRow - 1, 7).getValues();
  
  const exams = eData.map((r, i) => {
    let startTimeStr = "00:00";
    try {
      if (r[6] instanceof Date) {
        startTimeStr = Utilities.formatDate(r[6], "GMT+7", "HH:mm");
      } else if (typeof r[6] === 'string' && r[6].includes(':')) {
        startTimeStr = r[6].trim();
      }
    } catch (e) { startTimeStr = "00:00"; }

    let thaiDateStr = "";
    let isoDate = "";
    if (r[0] instanceof Date) {
      const months = ["ม.ค.", "ก.พ.", "มี.ค.", "เม.ย.", "พ.ค.", "มิ.ย.", "ก.ค.", "ส.ค.", "ก.ย.", "ต.ค.", "พ.ย.", "ธ.ค."];
      const d = r[0].getDate();
      const m = months[r[0].getMonth()];
      const y = (r[0].getFullYear() + 543).toString().slice(-2);
      thaiDateStr = d + " " + m + " " + y;
      isoDate = Utilities.formatDate(r[0], "GMT+7", "yyyy-MM-dd");
    }

    const examUrl = r[4] || "";
    // ตรวจสอบว่าเป็นลิงก์จริง หรือระบุเป็นรหัสชีตข้อสอบ (เช่น Q_MA11101 หรือ sheet:Q_MA11101)
    let isNativeExam = false;
    let sheetName = "";
    if (examUrl && !examUrl.toLowerCase().startsWith("http")) {
      // ตรวจสอบว่ามีชีตนี้อยู่ในระบบจริงหรือไม่
      const targetSheet = ss.getSheetByName(examUrl);
      if (targetSheet) {
        isNativeExam = true;
        sheetName = examUrl;
      }
    }

    return {
      rowNum: i + 2,
      rawDate: r[0] instanceof Date ? r[0].toISOString() : new Date().toISOString(), 
      isoDate: isoDate,
      date: thaiDateStr,
      timeRange: r[1] || "",
      subjectCode: r[2] ? r[2].toString() : "",
      subjectName: r[3] || "",
      url: examUrl,
      isNative: isNativeExam,
      targetSheetName: sheetName,
      duration: r[5] || 0,
      startTimeG: startTimeStr
    };
  });

  let finished = [];
  try {
    const sSheet = ss.getSheetByName("Students " + level);
    if (sSheet && studentId) {
      const sData = sSheet.getDataRange().getValues();
      const studentRow = sData.find(row => row[1].toString() === studentId.toString());
      if (studentRow) {
        // ดึงรายชื่อวิชาที่สอบไปแล้ว ตั้งแต่คอลัมน์ H เป็นต้นไป
        finished = studentRow.slice(7).map(v => v.toString().trim()).filter(v => v !== "");
      }
    }
  } catch (e) { console.log("Error finding finished exams: " + e.toString()); }

  return { exams, finished };
}

// บันทึกสถานะเพื่อล็อกการเข้าสอบซ้ำ (ล็อกเมื่อนักเรียนกดคลิกเริ่มทำข้อสอบ)
function recordEntry(id, level, code) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Students " + level);
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][1].toString() === id.toString()) {
      let col = 8;
      // วิ่งหาคอลัมน์ว่างในแถวของนักเรียน
      while (data[i][col-1]) { col++; }
      sheet.getRange(i + 1, col).setValue(code);
      return true;
    }
  }
}

// อัปเดตลิงก์วิชาสอบ (สำหรับครู)
function updateLink(level, row, url) {
  SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(level).getRange(row, 5).setValue(url);
  return "อัปเดตลิงก์วิชาเรียบร้อย";
}

// ปลดล็อกนักเรียนเพื่อให้สามารถทำข้อสอบนั้นใหม่ได้อีกครั้ง
function unlockStudent(level, sid, code) {
  const sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName("Students " + level);
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][1].toString() === sid.toString()) {
      for (let j = 7; j < data[i].length; j++) {
        if (data[i][j].toString() === code.toString()) {
          sheet.getRange(i+1, j+1).clearContent();
          return "ปลดล็อกสำเร็จ";
        }
      }
    }
  }
  return "ไม่พบข้อมูลวิชาที่ล็อก";
}

// ดึงโจทย์ข้อสอบจาก Google Sheets แบบปลอดภัย (กรองข้อมูลเฉลยออกเพื่อความปลอดภัยของเฉลย)
function getQuestions(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return null;
    
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];
    
    // ดึงคอลัมน์ A (ข้อที่) ถึง F (ตัวเลือก 4)
    const rawData = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    
    // แปลงให้เป็น Object ส่งไปหน้าบ้าน โดยไม่รวมเฉลย
    return rawData.map(r => {
      return {
        no: r[0],
        question: r[1] ? r[1].toString() : "",
        choices: [
          r[2] ? r[2].toString() : "",
          r[3] ? r[3].toString() : "",
          r[4] ? r[4].toString() : "",
          r[5] ? r[5].toString() : ""
        ]
      };
    });
  } catch (e) {
    return { error: e.toString() };
  }
}

// บันทึกกิจกรรมด้านความปลอดภัย (Security Logger) ป้องกันการสลับหน้าจอหรือทุจริต
function logSecurityEvent(studentId, studentName, level, subjectCode, eventType, details) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName("Security_Logs");
    if (!logSheet) {
      // สร้างชีตใหม่หากยังไม่มี
      logSheet = ss.insertSheet("Security_Logs");
      logSheet.appendRow(["วัน-เวลาที่เกิดเหตุ", "รหัสนักเรียน", "ชื่อ-นามสกุล", "ชั้นเรียน", "รหัสวิชา", "ประเภทเหตุการณ์", "รายละเอียดเหตุการณ์"]);
      logSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#ffebee").setFontColor("#b71c1c");
    }
    
    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    logSheet.appendRow([timestamp, studentId, studentName, level, subjectCode, eventType, details]);
    return true;
  } catch (e) {
    console.log("Error logging security event: " + e.toString());
    return false;
  }
}

// ตรวจคำตอบบนเซิร์ฟเวอร์ คำนวณคะแนน และบันทึกผลการสอบลงชีต Exam_Results
function submitExam(studentId, studentName, level, room, no, subjectCode, sheetName, studentAnswers, status) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const qSheet = ss.getSheetByName(sheetName);
    if (!qSheet) return { status: "error", message: "ไม่พบข้อมูลชีตข้อสอบในระบบ" };

    const lastRow = qSheet.getLastRow();
    if (lastRow < 2) return { status: "error", message: "ชีตข้อสอบไม่มีคำถาม" };

    // ดึงคำถาม คีย์เฉลย (คอลัมน์ G / Index 6) และคะแนนต่อข้อ (คอลัมน์ H / Index 7)
    const qData = qSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    
    let totalScore = 0;
    let maxScore = 0;
    
    // คำนวณคะแนนที่นักเรียนทำได้
    qData.forEach((row, index) => {
      const correctAns = parseInt(row[6]); // ตัวเลือกที่ถูกต้อง 1, 2, 3, 4
      const points = parseFloat(row[7]) || 1; // คะแนนของข้อนี้ (เริ่มต้น 1)
      maxScore += points;
      
      const studentAns = parseInt(studentAnswers[index]); // คำตอบของนักเรียน
      if (studentAns === correctAns) {
        totalScore += points;
      }
    });

    // บันทึกผลสอบในชีต Exam_Results
    let resSheet = ss.getSheetByName("Exam_Results");
    if (!resSheet) {
      resSheet = ss.insertSheet("Exam_Results");
      resSheet.appendRow(["วัน-เวลาส่ง", "รหัสนักเรียน", "ชื่อ-นามสกุล", "ชั้นเรียน", "ห้อง", "เลขที่", "รหัสวิชา", "คะแนนที่ได้", "คะแนนเต็ม", "สถานะการส่ง", "หมายเหตุ"]);
      resSheet.getRange("A1:K1").setFontWeight("bold").setBackground("#e8f5e9").setFontColor("#2e7d32");
    }

    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");
    resSheet.appendRow([timestamp, studentId, studentName, level, room, no, subjectCode, totalScore, maxScore, status, status === "FORCE_LOCKED" ? "ถูกตัดสิทธิ์วิชาเนื่องจากตรวจพบการทุจริตสลับหน้าจอ" : "ส่งข้อสอบตามปกติ"]);

    // บันทึกประวัติความปลอดภัยว่า ส่งข้อสอบเรียบร้อยแล้ว
    logSecurityEvent(studentId, studentName, level, subjectCode, status, `คะแนนที่ได้: ${totalScore}/${maxScore}`);

    return { status: "success", score: totalScore, maxScore: maxScore };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

// นำเข้าข้อสอบจากไฟล์ Word (.docx) ที่แปลงเป็นรายการบนหน้าเว็บเบราว์เซอร์แล้ว
function importExamQuestions(subjectCode, examData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = "Q_" + subjectCode.toString().trim().replace(/[^a-zA-Z0-9_ก-๙]/g, "");
    
    let qSheet = ss.getSheetByName(sheetName);
    if (qSheet) {
      qSheet.clear(); // ล้างชีตเก่าออก
    } else {
      qSheet = ss.insertSheet(sheetName);
    }
    
    // ตั้งค่าหัวตาราง
    qSheet.appendRow(["ข้อที่", "โจทย์คำถาม", "ตัวเลือก 1", "ตัวเลือก 2", "ตัวเลือก 3", "ตัวเลือก 4", "เฉลย (เลข 1-4)", "คะแนน (เช่น 1)"]);
    qSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#e3f2fd").setFontColor("#1565c0");
    
    // เขียนคำถามลงไป
    const rowsToWrite = examData.map(q => {
      return [
        q.no,
        q.question,
        q.choices[0] || "",
        q.choices[1] || "",
        q.choices[2] || "",
        q.choices[3] || "",
        parseInt(q.correctAns) || 1,
        parseFloat(q.points) || 1
      ];
    });
    
    qSheet.getRange(2, 1, rowsToWrite.length, 8).setValues(rowsToWrite);
    
    return { status: "success", message: `นำเข้าข้อสอบวิชา ${subjectCode} เรียบร้อยแล้ว ทั้งหมด ${rowsToWrite.length} ข้อ ชื่อชีต: ${sheetName}`, sheetName: sheetName };
  } catch (e) {
    return { status: "error", message: "ไม่สามารถบันทึกข้อสอบได้: " + e.toString() };
  }
}
