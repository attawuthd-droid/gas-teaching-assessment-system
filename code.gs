// ====================== SHEET NAMES ======================
const SHEETS = {
  STUDENTS: "Students",
  TEACHERS: "Teachers",
  ASSESSMENT_TYPES: "AssessmentTypes",
  ASSESSMENTS: "Assessments"
};

// ====================== SETUP & ROUTING ======================
function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ระบบประเมินการเรียนการสอน')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('doGet Error:', error);
    return HtmlService.createHtmlOutput('<h1>Error initializing web app</h1><p>' + error.message + '</p>');
  }
}

// ====================== HELPER FUNCTIONS ======================
function createResponse(success, data, message = '') {
  return { success: success, data: data, message: message, timestamp: new Date().toISOString() };
}

function getSheet(sheetName) {
  try { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName) || null; } 
  catch (error) { return null; }
}

function validateSheetSchema(sheet, expectedHeaders) {
  if (!sheet) return false;
  const headers = sheet.getDataRange().getValues()[0];
  return expectedHeaders.every(h => headers.includes(h));
}

function sheetToObjects(sheet, headers = null) {
  try {
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return [];
    const headerRow = headers || data[0];
    const objects = [];
    for (let i = 1; i < data.length; i++) {
      const obj = {};
      for (let j = 0; j < headerRow.length; j++) {
        obj[headerRow[j]] = data[i][j] !== undefined && data[i][j] !== null ? data[i][j] : '';
      }
      objects.push(obj);
    }
    return objects;
  } catch (error) { return []; }
}

const getYearFromRow = (row) => {
    for (let key in row) {
        const cleanKey = key.trim().toLowerCase();
        if (cleanKey === 'year' || cleanKey === 'ชั้นปี') return row[key] ? String(row[key]) : '';
    }
    return '';
};

// ====================== LOGIN FUNCTIONS ======================

function loginUser(username, password) {
  try {
    const teachersSheet = getSheet(SHEETS.TEACHERS);
    if (!teachersSheet || !validateSheetSchema(teachersSheet, ['fullname', 'email', 'password'])) {
      return createResponse(false, null, 'ไม่พบแผ่นงาน Teachers หรือโครงสร้างไม่ถูกต้อง');
    }
    const teachers = sheetToObjects(teachersSheet);
    const teacher = teachers.find(t => String(t.email).trim().toLowerCase() === String(username).trim().toLowerCase() && String(t.password).trim() === String(password).trim());

    if (teacher) return createResponse(true, { user: { username: teacher.email, fullname: teacher.fullname, role: 'teacher' } });
    return createResponse(false, null, 'ชื่อผู้ใช้หรือรหัสผ่านอาจารย์ไม่ถูกต้อง');
  } catch (error) {
    return createResponse(false, null, `เกิดข้อผิดพลาด: ${error.message}`);
  }
}

function loginStudent(studentId, password) {
  try {
    const studentSheet = getSheet(SHEETS.STUDENTS);
    if (!studentSheet || !validateSheetSchema(studentSheet, ['id', 'fullname', 'password'])) {
      return createResponse(false, null, 'ไม่พบแผ่นงาน Students หรือไม่มีคอลัมน์ password');
    }
    const students = sheetToObjects(studentSheet);
    const student = students.find(s => String(s.id).trim() === String(studentId).trim() && String(s.password).trim() === String(password).trim());

    if (student) return createResponse(true, { user: { id: student.id, fullname: student.fullname, role: 'student' } });
    return createResponse(false, null, 'รหัสนิสิตหรือรหัสผ่านไม่ถูกต้อง');
  } catch (error) {
    return createResponse(false, null, `เกิดข้อผิดพลาด: ${error.message}`);
  }
}

// ====================== DATA FUNCTIONS ======================

function getAllData(username) {
  try {
    const teacherSheet = getSheet(SHEETS.TEACHERS);
    const teacher = sheetToObjects(teacherSheet).find(t => String(t.email).trim().toLowerCase() === String(username).trim().toLowerCase());
    if (!teacher) return createResponse(false, null, 'ไม่พบข้อมูลอาจารย์');
    
    const teacherFullname = teacher.fullname || '';
    const assessmentSheet = getSheet(SHEETS.ASSESSMENTS);
    const filteredAssessments = sheetToObjects(assessmentSheet).filter(a => String(a.teacher_name).trim().toLowerCase() === String(teacherFullname).trim().toLowerCase());

    const studentSheet = getSheet(SHEETS.STUDENTS);
    const students = sheetToObjects(studentSheet).map(s => { s.year = getYearFromRow(s); return s; });

    const assessmentTypeSheet = getSheet(SHEETS.ASSESSMENT_TYPES);
    const assessmentTypes = sheetToObjects(assessmentTypeSheet).map(row => ({
      id: row.id?.toString() || '',
      name: row.name?.toString() || '',
      year: getYearFromRow(row),
      question: (() => {
        try { return row.question ? (typeof row.question === 'string' ? JSON.parse(row.question) : row.question) : []; } 
        catch (e) { return []; }
      })()
    }));

    return createResponse(true, { assessments: filteredAssessments, teacherFullname: teacherFullname, students: students, assessmentTypes: assessmentTypes });
  } catch (error) {
    return createResponse(false, null, `เกิดข้อผิดพลาด: ${error.message}`);
  }
}

// ดึงข้อมูลให้นิสิต (ส่ง ID นิสิตมา เพื่อหาชั้นปี)
function getStudentData(studentId) {
  try {
    const studentSheet = getSheet(SHEETS.STUDENTS);
    if (!studentSheet) throw new Error("ไม่พบแผ่นงาน Students");
    
    const students = sheetToObjects(studentSheet);
    const student = students.find(s => String(s.id).trim() === String(studentId).trim());
    if (!student) throw new Error("ไม่พบข้อมูลนิสิตในระบบ");

    const fullname = student.fullname;
    const studentYear = getYearFromRow(student);

    const assessSheet = getSheet(SHEETS.ASSESSMENTS);
    const typeSheet = getSheet(SHEETS.ASSESSMENT_TYPES);
    
    const assessments = assessSheet ? sheetToObjects(assessSheet).filter(a => String(a.student_fullname).trim() === String(fullname).trim()) : [];
    const types = typeSheet ? sheetToObjects(typeSheet).map(row => ({
       name: row.name,
       year: getYearFromRow(row),
       question: typeof row.question === 'string' ? JSON.parse(row.question) : row.question
    })) : [];

    return createResponse(true, { fullname, studentYear, assessments, types });
  } catch(e) {
    return createResponse(false, null, e.message);
  }
}

// ====================== CRUD FUNCTIONS ======================
function saveAssessment(assessments) {
  try {
    const assessmentSheet = getSheet(SHEETS.ASSESSMENTS);
    const headers = assessmentSheet.getDataRange().getValues()[0];
    const newRows = [];
    const savedAssessments = [];
    assessments.forEach(a => {
      const id = Utilities.getUuid();
      const timestamp = new Date().toISOString();
      const row = headers.map(header => {
        if (header === 'id') return id;
        else if (header === 'timestamp') return timestamp;
        else if (header === 'datetime_edit') return '';
        return a[header] !== undefined ? a[header] : '';
      });
      newRows.push(row);
      savedAssessments.push({ id, timestamp, ...a });
    });
    assessmentSheet.getRange(assessmentSheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
    return createResponse(true, { savedAssessments, message: `บันทึก ${newRows.length} การประเมินเรียบร้อย` });
  } catch (error) { return createResponse(false, null, error.message); }
}

function updateAssessment(assessment) {
  try {
    const assessmentSheet = getSheet(SHEETS.ASSESSMENTS);
    const data = assessmentSheet.getDataRange().getValues();
    const headers = data[0];
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][headers.indexOf('id')]).trim() === String(assessment.id).trim()) { rowIndex = i + 1; break; }
    }
    if (rowIndex === -1) return createResponse(false, null, 'ไม่พบการประเมิน');
    const row = headers.map(header => {
      if (header === 'datetime_edit') return new Date().toISOString();
      return assessment[header] !== undefined ? assessment[header] : data[rowIndex - 1][headers.indexOf(header)];
    });
    assessmentSheet.getRange(rowIndex, 1, 1, headers.length).setValues([row]);
    return createResponse(true, { message: 'แก้ไขเรียบร้อย' });
  } catch (error) { return createResponse(false, null, error.message); }
}

function deleteAssessment(assessmentId) {
  try {
    const assessmentSheet = getSheet(SHEETS.ASSESSMENTS);
    const data = assessmentSheet.getDataRange().getValues();
    const headers = data[0];
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][headers.indexOf('id')]).trim() === String(assessmentId).trim()) { rowIndex = i + 1; break; }
    }
    if (rowIndex === -1) return createResponse(false, null, 'ไม่พบการประเมิน');
    assessmentSheet.deleteRow(rowIndex);
    return createResponse(true, { message: 'ลบเรียบร้อย' });
  } catch (error) { return createResponse(false, null, error.message); }
}