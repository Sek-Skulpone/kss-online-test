const SPREADSHEET_ID = '1KZicFPjH37Key3fcFgEVqXZULLar-IcV0j9tg0tWOEo'; 

// ==============================================
// CONFIGURATION FOR FIREBASE FIRESTORE (OPTIONAL)
// ==============================================
const FIREBASE_PROJECT_ID = ''; // ใส่ Project ID ของ Firebase เช่น 'my-exam-project'
const FIREBASE_API_KEY = '';    // ใส่ Web API Key ของ Firebase

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
      isNativeExam = true;
      sheetName = examUrl;
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

// ดึงโจทย์ข้อสอบแบบปลอดภัย (กรองข้อมูลเฉลยออกก่อนส่งไปหน้านักเรียน)
function getQuestions(sheetName) {
  // ลองดึงจาก Firebase ก่อนหากระบุค่าเชื่อมต่อไว้
  if (FIREBASE_PROJECT_ID && FIREBASE_API_KEY) {
    const fbQuestions = getQuestionsFromFirebase(sheetName);
    if (fbQuestions && fbQuestions.length > 0) {
      console.log("Loaded questions from Firebase Firestore successfully.");
      return fbQuestions;
    }
  }

  // ดึงจาก Google Sheets (Fallback)
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
  const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");

  // บันทึกลง Firebase (ถ้าเปิดใช้)
  if (FIREBASE_PROJECT_ID && FIREBASE_API_KEY) {
    writeDocumentToFirebase("security_logs", studentId + "_" + eventType + "_" + new Date().getTime(), {
      timestamp: timestamp,
      studentId: studentId,
      studentName: studentName,
      level: level,
      subjectCode: subjectCode,
      eventType: eventType,
      details: details
    });
  }

  // บันทึกลง Google Sheet
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let logSheet = ss.getSheetByName("Security_Logs");
    if (!logSheet) {
      logSheet = ss.insertSheet("Security_Logs");
      logSheet.appendRow(["วัน-เวลาที่เกิดเหตุ", "รหัสนักเรียน", "ชื่อ-นามสกุล", "ชั้นเรียน", "รหัสวิชา", "ประเภทเหตุการณ์", "รายละเอียดเหตุการณ์"]);
      logSheet.getRange("A1:G1").setFontWeight("bold").setBackground("#ffebee").setFontColor("#b71c1c");
    }
    logSheet.appendRow([timestamp, studentId, studentName, level, subjectCode, eventType, details]);
    return true;
  } catch (e) {
    console.log("Error logging security event: " + e.toString());
    return false;
  }
}

// ตรวจคำตอบบนเซิร์ฟเวอร์ คำนวณคะแนน และบันทึกผลการสอบลงชีต
function submitExam(studentId, studentName, level, room, no, subjectCode, sheetName, studentAnswers, status) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let qData = [];
    let isFromFirebase = false;

    // ลองตรวจกระดาษคำตอบจากฐานข้อมูล Firebase ก่อน
    if (FIREBASE_PROJECT_ID && FIREBASE_API_KEY) {
      const fbExamDoc = readDocumentFromFirebase("exams", sheetName);
      if (fbExamDoc && fbExamDoc.questions) {
        qData = fbExamDoc.questions;
        isFromFirebase = true;
      }
    }

    // หากไม่มี Firebase หรือไม่พบคลังข้อสอบ ดึงจากชีตตามปกติ
    if (!isFromFirebase) {
      const qSheet = ss.getSheetByName(sheetName);
      if (!qSheet) return { status: "error", message: "ไม่พบข้อมูลชีตข้อสอบในระบบ" };

      const lastRow = qSheet.getLastRow();
      if (lastRow < 2) return { status: "error", message: "ชีตข้อสอบไม่มีคำถาม" };

      // ดึงคอลัมน์ A ถึง H (เฉลยคือ Index 6, คะแนนคือ Index 7)
      const rawSheetData = qSheet.getRange(2, 1, lastRow - 1, 8).getValues();
      qData = rawSheetData.map(r => ({
        correctAns: parseInt(r[6]),
        points: parseFloat(r[7]) || 1
      }));
    }
    
    let totalScore = 0;
    let maxScore = 0;
    
    qData.forEach((row, index) => {
      const correctAns = parseInt(row.correctAns); // 1-4
      const points = parseFloat(row.points) || 1;
      maxScore += points;
      
      const studentAns = parseInt(studentAnswers[index]);
      if (studentAns === correctAns) {
        totalScore += points;
      }
    });

    const timestamp = Utilities.formatDate(new Date(), "GMT+7", "yyyy-MM-dd HH:mm:ss");

    // บันทึกคะแนนลง Firebase
    if (FIREBASE_PROJECT_ID && FIREBASE_API_KEY) {
      writeDocumentToFirebase("exam_results", studentId + "_" + subjectCode, {
        timestamp: timestamp,
        studentId: studentId,
        studentName: studentName,
        level: level,
        room: room,
        no: no,
        subjectCode: subjectCode,
        score: totalScore,
        maxScore: maxScore,
        status: status,
        details: status === "FORCE_LOCKED" ? "ถูกตัดสิทธิ์สอบเนื่องจากตรวจพบการทุจริตสลับจอ" : "ส่งข้อสอบตามปกติ"
      });
    }

    // บันทึกคะแนนลง Google Sheet
    let resSheet = ss.getSheetByName("Exam_Results");
    if (!resSheet) {
      resSheet = ss.insertSheet("Exam_Results");
      resSheet.appendRow(["วัน-เวลาส่ง", "รหัสนักเรียน", "ชื่อ-นามสกุล", "ชั้นเรียน", "ห้อง", "เลขที่", "รหัสวิชา", "คะแนนที่ได้", "คะแนนเต็ม", "สถานะการส่ง", "หมายเหตุ"]);
      resSheet.getRange("A1:K1").setFontWeight("bold").setBackground("#e8f5e9").setFontColor("#2e7d32");
    }
    resSheet.appendRow([timestamp, studentId, studentName, level, room, no, subjectCode, totalScore, maxScore, status, status === "FORCE_LOCKED" ? "ถูกตัดสิทธิ์วิชาเนื่องจากตรวจพบการทุจริตสลับหน้าจอ" : "ส่งข้อสอบตามปกติ"]);

    logSecurityEvent(studentId, studentName, level, subjectCode, status, `คะแนนที่ได้: ${totalScore}/${maxScore}`);

    return { status: "success", score: totalScore, maxScore: maxScore };
  } catch (e) {
    return { status: "error", message: e.toString() };
  }
}

// นำเข้าข้อสอบจากไฟล์ Word (.docx) เขียนลง Google Sheets และ Mirror ขึ้น Firebase
function importExamQuestions(subjectCode, examData) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = "Q_" + subjectCode.toString().trim().replace(/[^a-zA-Z0-9_ก-๙]/g, "");
    
    // 1. บันทึกลง Firebase Firestore (หากมีคีย์ระบุไว้)
    if (FIREBASE_PROJECT_ID && FIREBASE_API_KEY) {
      const fbExamData = {
        subjectCode: subjectCode,
        sheetName: sheetName,
        questions: examData.map(q => ({
          no: parseInt(q.no),
          question: q.question,
          choices: q.choices,
          correctAns: parseInt(q.correctAns),
          points: parseFloat(q.points)
        }))
      };
      
      writeDocumentToFirebase("exams", sheetName, fbExamData);
    }
    
    // 2. บันทึกลง Google Sheet
    let qSheet = ss.getSheetByName(sheetName);
    if (qSheet) {
      qSheet.clear();
    } else {
      qSheet = ss.insertSheet(sheetName);
    }
    
    qSheet.appendRow(["ข้อที่", "โจทย์คำถาม", "ตัวเลือก 1", "ตัวเลือก 2", "ตัวเลือก 3", "ตัวเลือก 4", "เฉลย (เลข 1-4)", "คะแนน (เช่น 1)"]);
    qSheet.getRange("A1:H1").setFontWeight("bold").setBackground("#e3f2fd").setFontColor("#1565c0");
    
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
    
    return { 
      status: "success", 
      message: `นำเข้าข้อสอบวิชา ${subjectCode} เรียบร้อยแล้ว ทั้งหมด ${rowsToWrite.length} ข้อ (บันทึกเสร็จสิ้นทั้งชีตตารางและระบบเก็บคลาวด์)`, 
      sheetName: sheetName 
    };
  } catch (e) {
    return { status: "error", message: "ไม่สามารถบันทึกข้อสอบได้: " + e.toString() };
  }
}

// ==============================================
// FIREBASE FIRESTORE REST CLIENT FOR GOOGLE APPS SCRIPT
// ==============================================

// เขียนเอกสารลง Firestore
function writeDocumentToFirebase(collection, documentId, dataObject) {
  try {
    const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/${collection}/${documentId}?key=${FIREBASE_API_KEY}`;
    
    const payload = toFirestoreDocument(dataObject);
    const options = {
      method: "patch", // patch จะแก้ไขข้อมูลเดิมหรือสร้างใหม่หากไม่มี
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    if (code !== 200) {
      console.log(`Firebase Write Error (${code}): ` + response.getContentText());
      return false;
    }
    return true;
  } catch (e) {
    console.log("Firebase Connection Failed (Write): " + e.toString());
    return false;
  }
}

// อ่านเอกสารจาก Firestore
function readDocumentFromFirebase(collection, documentId) {
  try {
    const url = `https://firestore.googleapis.com/v1/projects/${FIREBASE_PROJECT_ID}/databases/(default)/documents/${collection}/${documentId}?key=${FIREBASE_API_KEY}`;
    const options = {
      method: "get",
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      const resData = JSON.parse(response.getContentText());
      return fromFirestoreDocument(resData);
    }
    return null;
  } catch (e) {
    console.log("Firebase Connection Failed (Read): " + e.toString());
    return null;
  }
}

// ดึงข้อสอบและคำถามเฉพาะส่วนของนักเรียนจาก Firestore (ซ่อนเฉลย)
function getQuestionsFromFirebase(sheetName) {
  const doc = readDocumentFromFirebase("exams", sheetName);
  if (!doc || !doc.questions) return null;
  
  // ซ่อนฟิลด์ correctAns และ points ก่อนส่งออกไปยังหน้าบ้านนักเรียน
  return doc.questions.map(q => {
    return {
      no: q.no,
      question: q.question,
      choices: q.choices
    };
  });
}

// --- ตัวแปลงโครงสร้าง JSON-to-Firestore REST ---

function toFirestoreDocument(jsonObj) {
  const fields = {};
  for (let key in jsonObj) {
    fields[key] = toFirestoreValue(jsonObj[key]);
  }
  return { fields: fields };
}

function toFirestoreValue(val) {
  if (val === null || val === undefined) {
    return { nullValue: null };
  }
  if (typeof val === 'string') {
    return { stringValue: val };
  }
  if (typeof val === 'number') {
    if (Number.isInteger(val)) {
      return { integerValue: val.toString() };
    } else {
      return { doubleValue: val };
    }
  }
  if (typeof val === 'boolean') {
    return { booleanValue: val };
  }
  if (Array.isArray(val)) {
    return { arrayValue: { values: val.map(toFirestoreValue) } };
  }
  if (typeof val === 'object') {
    return { mapValue: toFirestoreDocument(val) };
  }
  return { stringValue: val.toString() };
}

function fromFirestoreDocument(doc) {
  if (!doc || !doc.fields) return {};
  const json = {};
  for (let key in doc.fields) {
    json[key] = fromFirestoreValue(doc.fields[key]);
  }
  return json;
}

function fromFirestoreValue(field) {
  if (!field) return null;
  if ('stringValue' in field) return field.stringValue;
  if ('integerValue' in field) return parseInt(field.integerValue);
  if ('doubleValue' in field) return parseFloat(field.doubleValue);
  if ('booleanValue' in field) return field.booleanValue;
  if ('nullValue' in field) return null;
  if ('arrayValue' in field && field.arrayValue.values) {
    return field.arrayValue.values.map(fromFirestoreValue);
  }
  if ('mapValue' in field) {
    return fromFirestoreDocument(field.mapValue);
  }
  return null;
}
