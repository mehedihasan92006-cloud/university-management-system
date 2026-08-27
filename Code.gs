/**
 * University Management System – Google Apps Script Backend
 * ---------------------------------------------------------
 * 1. Create a new Google Spreadsheet
 * 2. Extensions → Apps Script → paste this entire file
 * 3. Run the function `setupSheets` once (from the editor)
 * 4. Deploy → New deployment → Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 * 5. Copy the Web App URL and put it in js/config.js
 */

const TABLES = [
  "users", "students", "faculty", "courses", "enrollments",
  "fees", "exams", "results", "books", "issues", "rooms", "allocations",
  "id_cards"
];

const HEADERS = {
  users:        ["id", "username", "email", "password", "full_name", "role", "student_id", "faculty_id", "created_at"],
  students:     ["id", "student_id", "first_name", "last_name", "email", "phone", "department", "year", "gender", "address", "created_at"],
  faculty:      ["id", "faculty_id", "first_name", "last_name", "email", "phone", "department", "designation", "created_at"],
  courses:      ["id", "course_code", "title", "credits", "department", "faculty_id", "semester"],
  enrollments:  ["id", "student_id", "course_id", "enroll_date", "status"],
  fees:         ["id", "student_id", "amount", "fee_type", "due_date", "paid", "paid_date", "created_at"],
  exams:        ["id", "course_id", "exam_name", "exam_date", "max_marks"],
  results:      ["id", "exam_id", "student_id", "marks_obtained", "grade"],
  books:        ["id", "isbn", "title", "author", "category", "total_copies", "available_copies"],
  issues:       ["id", "book_id", "student_id", "issue_date", "due_date", "return_date", "status"],
  rooms:        ["id", "room_number", "block", "capacity", "occupied"],
  allocations:  ["id", "room_id", "student_id", "allocate_date", "status"],
  id_cards:     ["id", "code", "type", "label", "status", "issued_date", "used_by", "used_date"]
};

function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  try {
    let body = {};
    if (e.postData && e.postData.contents) {
      body = JSON.parse(e.postData.contents);
    } else if (e.parameter) {
      body = e.parameter;
      if (body.data && typeof body.data === "string") {
        try { body.data = JSON.parse(body.data); } catch (_) {}
      }
    }

    const action = (body.action || "").toLowerCase();
    let result;

    switch (action) {
      case "list":
        result = listRows(body.table);
        break;
      case "get":
        result = getRow(body.table, Number(body.id));
        break;
      case "insert":
        result = insertRow(body.table, body.data || {});
        break;
      case "update":
        result = updateRow(body.table, Number(body.id), body.data || {});
        break;
      case "delete":
        result = deleteRow(body.table, Number(body.id));
        break;
      case "deletewhere":
        result = deleteWhere(body.table, body.field, body.value);
        break;
      case "login":
        result = loginUser(body.usernameOrEmail, body.password, body.loginType);
        break;
      case "signup":
        result = signupUser(body.data || {});
        break;
      case "issueid":
        result = issueId(body.type, body.label);
        break;
      case "seedadmin":
        result = ensureDefaultAdmin();
        break;
      case "askai":
        result = askAI(body.prompt);
        break;
      case "ping":
        result = { ok: true, message: "UMS Google Sheets API is running" };
        break;
      default:
        result = { ok: false, error: "Unknown action: " + action };
    }

    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message || String(err) });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(table) {
  if (TABLES.indexOf(table) === -1) throw new Error("Invalid table: " + table);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(table);
  if (!sheet) {
    sheet = ss.insertSheet(table);
    const headers = HEADERS[table];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0].map(String);
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const obj = {};
    let empty = true;
    for (let j = 0; j < headers.length; j++) {
      let val = data[i][j];
      if (val !== "" && val !== null && val !== undefined) empty = false;
      // Convert numbers that should stay numbers
      if (headers[j] === "id" || headers[j].endsWith("_id") ||
          headers[j] === "credits" || headers[j] === "amount" ||
          headers[j] === "max_marks" || headers[j] === "marks_obtained" ||
          headers[j] === "total_copies" || headers[j] === "available_copies" ||
          headers[j] === "capacity" || headers[j] === "occupied") {
        val = (val === "" || val === null) ? null : Number(val);
      }
      if (headers[j] === "paid") {
        val = (val === true || val === "TRUE" || val === "true" || val === 1 || val === "1");
      }
      obj[headers[j]] = val === "" ? null : val;
    }
    if (!empty) rows.push(obj);
  }
  return rows;
}

function nextId(sheet) {
  const data = sheet.getDataRange().getValues();
  let max = 0;
  for (let i = 1; i < data.length; i++) {
    const id = Number(data[i][0]);
    if (!isNaN(id) && id > max) max = id;
  }
  return max + 1;
}

function listRows(table) {
  const sheet = getSheet(table);
  return { ok: true, data: sheetToObjects(sheet) };
}

function getRow(table, id) {
  const rows = sheetToObjects(getSheet(table));
  const row = rows.find(r => Number(r.id) === id) || null;
  return { ok: true, data: row };
}

function insertRow(table, data) {
  const sheet = getSheet(table);
  const headers = HEADERS[table];
  const id = nextId(sheet);
  data.id = id;
  if (!data.created_at && headers.indexOf("created_at") !== -1) {
    data.created_at = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  const row = headers.map(h => (data[h] !== undefined && data[h] !== null) ? data[h] : "");
  sheet.appendRow(row);
  return { ok: true, id: id, data: data };
}

function updateRow(table, id, patch) {
  const sheet = getSheet(table);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === id) {
      for (let j = 0; j < headers.length; j++) {
        const h = headers[j];
        if (h !== "id" && patch[h] !== undefined) {
          data[i][j] = patch[h] === null ? "" : patch[h];
        }
      }
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([data[i]]);
      const obj = {};
      headers.forEach((h, j) => obj[h] = data[i][j] === "" ? null : data[i][j]);
      return { ok: true, data: obj };
    }
  }
  return { ok: false, error: "Row not found" };
}

function deleteRow(table, id) {
  const sheet = getSheet(table);
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (Number(data[i][0]) === id) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: "Row not found" };
}

function deleteWhere(table, field, value) {
  const sheet = getSheet(table);
  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(String);
  const col = headers.indexOf(field);
  if (col === -1) return { ok: false, error: "Unknown field" };
  let deleted = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][col]) === String(value)) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return { ok: true, deleted: deleted };
}

function hashPassword(password) {
  // Same simple demo hash as the original static version (not secure, demo only)
  let hash = 0;
  const s = String(password);
  for (let i = 0; i < s.length; i++) {
    hash = ((hash << 5) - hash) + s.charCodeAt(i);
    hash |= 0;
  }
  return "h_" + hash + "_" + s.length;
}

function loginUser(usernameOrEmail, password, loginType) {
  const users = sheetToObjects(getSheet("users"));
  const user = users.find(u =>
    u.username === usernameOrEmail || u.email === usernameOrEmail
  );

  if (!user || user.password !== hashPassword(password)) {
    return { ok: false, error: "Invalid username or password" };
  }

  // Enforce that the selected login tab matches the account's actual role.
  // A Student ID/password can only sign in through Student Login, a Faculty
  // ID/password only through Faculty Login, etc.
  const type = String(loginType || "administrative").toLowerCase();
  const requiredRoleByType = { student: "student", faculty: "faculty", administrative: "admin" };
  const requiredRole = requiredRoleByType[type];

  if (requiredRole && user.role !== requiredRole) {
    const label = { student: "Student", faculty: "Faculty", admin: "Administrative" }[requiredRole] || requiredRole;
    return {
      ok: false,
      error: "This account is not registered for " + label + " Login. Please choose the correct login tab."
    };
  }

  return {
    ok: true,
    user: {
      id: user.id,
      username: user.username,
      full_name: user.full_name || user.username,
      role: user.role,
      email: user.email,
      student_id: user.student_id || null,
      faculty_id: user.faculty_id || null
    }
  };
}

function issueId(type, label) {
  type = String(type || "").trim().toLowerCase();
  if (type !== "student" && type !== "faculty") {
    return { ok: false, error: "Invalid ID type. Choose Student or Faculty." };
  }

  const prefix = type === "student" ? "STU" : "FAC";
  const rows = sheetToObjects(getSheet("id_cards"));
  let maxNum = 0;
  rows.forEach(function(r) {
    if (r.type !== type) return;
    const m = String(r.code || "").match(new RegExp("^" + prefix + "(\\d+)$"));
    if (m) {
      const n = parseInt(m[1], 10);
      if (n > maxNum) maxNum = n;
    }
  });
  const code = prefix + String(maxNum + 1).padStart(3, "0");

  const result = insertRow("id_cards", {
    code: code,
    type: type,
    label: (label || "").trim() || null,
    status: "available",
    issued_date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")
  });
  return { ok: true, code: code, id: result.id };
}

function signupUser(data) {
  const role = String(data.role || "").trim().toLowerCase(); // "student" or "faculty"
  if (role !== "student" && role !== "faculty") {
    return { ok: false, error: "Invalid signup type. Choose Student or Faculty." };
  }

  const identifierField = role === "student" ? "student_id" : "faculty_id";
  const identifierLabel = role === "student" ? "Student ID" : "Faculty ID";
  const identifier = String(data[identifierField] || data.identifier || "").trim();
  const email = (data.email || "").trim();
  const password = data.password || "";
  const full_name = (data.full_name || "").trim();

  if (!identifier || !email || !password) {
    return { ok: false, error: identifierLabel + ", email and password are required." };
  }
  if (password.length < 6) {
    return { ok: false, error: "Password must be at least 6 characters." };
  }

  // The ID must already exist as an "available" card issued by an admin
  // in the ID Cards tab — this stops anyone from typing in a made-up ID.
  const idCards = sheetToObjects(getSheet("id_cards"));
  const card = idCards.find(function(c) { return c.type === role && String(c.code) === identifier; });
  if (!card) {
    return { ok: false, error: "This " + identifierLabel + " has not been issued. Please contact the administrator." };
  }
  if (card.status === "used") {
    return { ok: false, error: "This " + identifierLabel + " has already been used to create an account." };
  }

  const users = sheetToObjects(getSheet("users"));
  const duplicate = users.some(u => u.username === identifier || u.email === email);
  if (duplicate) {
    return { ok: false, error: "An account with this " + identifierLabel + " or email already exists." };
  }

  const payload = {
    username: identifier,
    email: email,
    password: hashPassword(password),
    full_name: full_name || card.label || identifier,
    role: role,
    created_at: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")
  };
  payload[identifierField] = identifier;

  const result = insertRow("users", payload);

  updateRow("id_cards", card.id, {
    status: "used",
    used_by: result.id,
    used_date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")
  });

  return {
    ok: true,
    message: "Account created successfully! Please login using " + identifierLabel + " Login.",
    id: result.id
  };
}

function ensureDefaultAdmin() {
  const users = sheetToObjects(getSheet("users"));
  const admin = users.find(u => u.username === "admin");
  if (admin) {
    return { ok: true, message: "Admin already exists", id: admin.id };
  }
  const result = insertRow("users", {
    username: "admin",
    email: "admin@university.edu",
    password: hashPassword("admin123"),
    full_name: "System Administrator",
    role: "admin",
    created_at: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")
  });
  return { ok: true, message: "Admin created", id: result.id };
}

function askAI(prompt) {
  prompt = String(prompt || "").trim();
  if (!prompt) return { ok: false, error: "Please enter a question." };

  const apiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
  if (!apiKey) {
    return { ok: false, error: "AI is not configured. Add GEMINI_API_KEY in Script Properties." };
  }

  const model = PropertiesService.getScriptProperties().getProperty("GEMINI_MODEL") || "gemini-3.6-flash";
  const response = UrlFetchApp.fetch(
    "https://generativelanguage.googleapis.com/v1beta/models/" + model + ":generateContent?key=" + encodeURIComponent(apiKey),
    {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      systemInstruction: {
        parts: [{
          text: "You are an intelligent assistant integrated into Mehedi's University Management System. Provide concise and helpful support regarding student, course, fee, and hostel records."
        }]
      },
      contents: [{
        role: "user",
        parts: [{ text: prompt }]
      }],
      generationConfig: {
        temperature: 0.7
      }
    }),
    muteHttpExceptions: true
    }
  );

  const status = response.getResponseCode();
  const data = JSON.parse(response.getContentText());
  if (status < 200 || status >= 300) {
    return { ok: false, error: data.error && data.error.message || "Gemini request failed." };
  }
  const reply = data.candidates && data.candidates[0] &&
    data.candidates[0].content && data.candidates[0].content.parts &&
    data.candidates[0].content.parts.map(function(part) { return part.text || ""; }).join("");
  return reply ? { ok: true, reply: reply } : { ok: false, error: "Gemini returned no answer." };
}

/**
 * Run this ONCE from the Apps Script editor
 * (select setupSheets → Run)
 */
function setupSheets() {
  TABLES.forEach(function(t) {
    getSheet(t);
  });
  ensureDefaultAdmin();
  SpreadsheetApp.getUi().alert("All sheets created and default admin ready!\n\nUsername: admin\nPassword: admin123");
}
