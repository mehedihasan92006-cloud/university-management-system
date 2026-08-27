/**
 * University Management System – Google Apps Script Backend
 * Profile photos are stored on Google Drive and the URL is saved on the users sheet.
 */

const TABLES = [
  "users", "students", "faculty", "courses", "enrollments",
  "fees", "exams", "results", "books", "issues", "rooms", "allocations",
  "id_cards"
];

const HEADERS = {
  users:        ["id", "username", "email", "password", "full_name", "role", "student_id", "faculty_id", "created_at", "photo_url", "photo_file_id"],
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
  id_cards:     ["id", "code", "type", "label", "recipient_email", "status", "issued_date", "email_sent_date", "used_by", "used_date"]
};

function doGet(e) {
  return HtmlService.createTemplateFromFile("index").evaluate()
    .setTitle("University Management System")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
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
      case "saveprofile":
        result = saveProfile(Number(body.id || body.userId), body.data || {});
        break;
      case "savephoto":
        result = saveProfilePhoto(Number(body.id || body.userId), body.photo || body.photo_url || "");
        break;
      case "removephoto":
        result = removeProfilePhoto(Number(body.id || body.userId));
        break;
      case "ping":
        result = { ok: true, message: "UMS API is running" };
        break;
      case "issueid":
        result = issueId(body.type, body.label, body.recipientEmail || body.recipient_email);
        break;
      case "sendidemail":
        result = sendIdEmail(Number(body.id));
        break;
      case "askai":
      case "aichat":
        result = { ok: true, reply: "AI assistant is not configured yet. Please contact the administrator." };
        break;
      default:
        result = { ok: false, error: "Unknown action: " + action };
    }

    result = sanitizeData(result);
    return jsonResponse(result);
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message || String(err) });
  }
}

function getGravatarUrl(email) {
  if (!email) return "https://www.gravatar.com/avatar/?d=mp";
  const cleanEmail = String(email).trim().toLowerCase();
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, cleanEmail);
  let hashStr = "";
  for (let i = 0; i < rawHash.length; i++) {
    let byteStr = (rawHash[i] < 0 ? rawHash[i] + 256 : rawHash[i]).toString(16);
    if (byteStr.length === 1) byteStr = "0" + byteStr;
    hashStr += byteStr;
  }
  return "https://www.gravatar.com/avatar/" + hashStr + "?d=mp&s=200";
}

function userPhoto(user) {
  if (user && user.photo_url) return user.photo_url;
  return getGravatarUrl(user && user.email);
}

/** Never send passwords back. Keep student_id — the UI needs it. */
function formatCellValue(val) {
  if (val === "" || val === null || val === undefined) return null;
  // Google Sheets Date objects → "yyyy-MM-dd" (avoids [object Object] in the UI)
  if (Object.prototype.toString.call(val) === "[object Date]" && !isNaN(val.getTime())) {
    try {
      return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } catch (_) {
      return val.toISOString().slice(0, 10);
    }
  }
  // Any other non-primitive object would become "[object Object]" in the UI
  if (typeof val === "object") {
    try {
      if (typeof val.getTime === "function" && !isNaN(val.getTime())) {
        return Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
    } catch (_) {}
    return String(val);
  }
  return val;
}

function sanitizeData(obj) {
  if (obj === null || obj === undefined) return obj;
  if (Object.prototype.toString.call(obj) === "[object Date]") return formatCellValue(obj);
  if (typeof obj !== "object") return obj;
  if (Array.isArray(obj)) return obj.map(item => sanitizeData(item));
  const cleanObj = {};
  for (let key in obj) {
    if (key === "password") continue;
    cleanObj[key] = sanitizeData(obj[key]);
  }
  return cleanObj;
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
  } else {
    ensureHeaders(sheet, HEADERS[table]);
  }
  return sheet;
}

function ensureHeaders(sheet, expected) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = expected.filter(h => current.indexOf(h) === -1);
  if (!missing.length) return;
  const next = current.concat(missing);
  sheet.getRange(1, 1, 1, next.length).setValues([next]);
  sheet.getRange(1, 1, 1, next.length).setFontWeight("bold");
}

function sheetHeaders(sheet) {
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
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
      let val = formatCellValue(data[i][j]);
      if (val !== "" && val !== null && val !== undefined) empty = false;
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

function findRowIndex(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (Number(data[i][0]) === Number(id)) return i + 1;
  }
  return -1;
}

function listRows(table) {
  const data = sheetToObjects(getSheet(table));
  // Back-fill missing Student / Faculty IDs on older user accounts
  if (table === "users") {
    data.forEach(u => {
      const role = String(u.role || "").toLowerCase();
      if (role === "student" && !u.student_id) {
        const code = nextCode("student");
        updateRow("users", u.id, { student_id: code });
        u.student_id = code;
      }
      if (role === "faculty" && !u.faculty_id) {
        const code = nextCode("faculty");
        updateRow("users", u.id, { faculty_id: code });
        u.faculty_id = code;
      }
      // Clean broken created_at values like "[object Object]"
      if (u.created_at && String(u.created_at) === "[object Object]") {
        updateRow("users", u.id, { created_at: "" });
        u.created_at = null;
      }
    });
  }
  return { ok: true, data: data };
}

function getRow(table, id) {
  const rows = sheetToObjects(getSheet(table));
  return { ok: true, data: rows.find(r => Number(r.id) === id) || null };
}

function insertRow(table, data) {
  const sheet = getSheet(table);
  const headers = sheetHeaders(sheet);
  const id = nextId(sheet);
  data.id = id;
  const row = headers.map(h => (data[h] !== undefined && data[h] !== null) ? data[h] : "");
  sheet.appendRow(row);
  return { ok: true, id: id, data: data };
}

function updateRow(table, id, data) {
  const sheet = getSheet(table);
  const headers = sheetHeaders(sheet);
  const rowIndex = findRowIndex(sheet, id);
  if (rowIndex === -1) return { ok: false, error: "Row not found with id: " + id };

  const currentRange = sheet.getRange(rowIndex, 1, 1, headers.length);
  const currentValues = currentRange.getValues()[0];
  const updated = headers.map((h, idx) => {
    if (h === "id") return id;
    if (data[h] !== undefined) return data[h] === null ? "" : data[h];
    return currentValues[idx];
  });
  currentRange.setValues([updated]);
  return { ok: true, id: id, data: data };
}

function deleteRow(table, id) {
  const sheet = getSheet(table);
  const rowIndex = findRowIndex(sheet, id);
  if (rowIndex === -1) return { ok: false, error: "Row not found with id: " + id };
  sheet.deleteRow(rowIndex);
  return { ok: true, id: id };
}

function deleteWhere(table, field, value) {
  const sheet = getSheet(table);
  const headers = sheetHeaders(sheet);
  const colIndex = headers.indexOf(field);
  if (colIndex === -1) return { ok: false, error: "Invalid field: " + field };

  const data = sheet.getDataRange().getValues();
  let deleted = 0;
  for (let i = data.length - 1; i >= 1; i--) {
    const cellVal = data[i][colIndex];
    if (String(cellVal) === String(value) || Number(cellVal) === Number(value)) {
      sheet.deleteRow(i + 1);
      deleted++;
    }
  }
  return { ok: true, deleted: deleted };
}

function hashPassword(password) {
  let hash = 0;
  const s = String(password);
  for (let i = 0; i < s.length; i++) {
    hash = ((hash << 5) - hash) + s.charCodeAt(i);
    hash |= 0;
  }
  return "h_" + hash + "_" + s.length;
}

function publicUser(user) {
  return {
    id: user.id,
    username: user.username,
    full_name: user.full_name || user.username,
    role: user.role,
    email: user.email,
    student_id: user.student_id || null,
    faculty_id: user.faculty_id || null,
    photo_url: userPhoto(user)
  };
}

function loginUser(usernameOrEmail, password, loginType) {
  const users = sheetToObjects(getSheet("users"));
  const user = users.find(u => u.username === usernameOrEmail || u.email === usernameOrEmail);

  if (!user || user.password !== hashPassword(password)) {
    return { ok: false, error: "Invalid username or password" };
  }

  return { ok: true, user: publicUser(user) };
}

function signupUser(data) {
  const email = (data.email || "").trim();
  const password = data.password || "";
  const full_name = (data.full_name || "").trim();
  const role = String(data.role || "student").toLowerCase();

  if (!email || !password) {
    return { ok: false, error: "Email and password are required." };
  }

  const existing = sheetToObjects(getSheet("users"));
  if (existing.some(u => u.email && u.email.toLowerCase() === email.toLowerCase())) {
    return { ok: false, error: "An account with this email already exists." };
  }

  // Prefer a pre-issued ID card matching this email + role
  const cards = sheetToObjects(getSheet("id_cards"));
  const matched = cards.find(c =>
    c.recipient_email &&
    String(c.recipient_email).toLowerCase() === email.toLowerCase() &&
    String(c.type || "").toLowerCase() === role &&
    String(c.status || "available").toLowerCase() !== "used"
  );

  const payload = {
    username: email.split("@")[0],
    email: email,
    password: hashPassword(password),
    full_name: full_name || email.split("@")[0],
    role: role,
    created_at: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")
  };

  // Auto-assign Student ID or Faculty ID
  if (role === "student") {
    payload.student_id = matched ? matched.code : nextCode("student");
  } else if (role === "faculty") {
    payload.faculty_id = matched ? matched.code : nextCode("faculty");
  }

  const result = insertRow("users", payload);

  // Mark pre-issued card as used
  if (matched) {
    updateRow("id_cards", matched.id, {
      status: "used",
      used_by: result.id,
      used_date: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")
    });
  }

  // Create matching profile row in students / faculty sheet
  const nameParts = String(payload.full_name || "").trim().split(/\s+/);
  const first_name = nameParts[0] || payload.username;
  const last_name = nameParts.slice(1).join(" ") || "";
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

  if (role === "student" && payload.student_id) {
    const already = sheetToObjects(getSheet("students")).some(
      s => String(s.student_id) === String(payload.student_id) ||
           (s.email && String(s.email).toLowerCase() === email.toLowerCase())
    );
    if (!already) {
      insertRow("students", {
        student_id: payload.student_id,
        first_name: first_name,
        last_name: last_name,
        email: email,
        phone: "",
        department: "",
        year: "",
        gender: "",
        address: "",
        created_at: today
      });
    }
  }

  if (role === "faculty" && payload.faculty_id) {
    const already = sheetToObjects(getSheet("faculty")).some(
      f => String(f.faculty_id) === String(payload.faculty_id) ||
           (f.email && String(f.email).toLowerCase() === email.toLowerCase())
    );
    if (!already) {
      insertRow("faculty", {
        faculty_id: payload.faculty_id,
        first_name: first_name,
        last_name: last_name,
        email: email,
        phone: "",
        department: "",
        designation: "",
        created_at: today
      });
    }
  }

  return {
    ok: true,
    message: "Account created successfully! Your " +
      (role === "faculty" ? "Faculty ID is " + payload.faculty_id :
       role === "student" ? "Student ID is " + payload.student_id : "account is ready") + ".",
    user: {
      id: result.id,
      email: email,
      full_name: payload.full_name,
      role: role,
      student_id: payload.student_id || null,
      faculty_id: payload.faculty_id || null,
      photo_url: getGravatarUrl(email)
    }
  };
}

/** Generate next Student (STU###) or Faculty (FAC###) ID automatically. */
function nextCode(type) {
  const prefix = type === "faculty" ? "FAC" : "STU";
  let max = 0;

  // From id_cards
  sheetToObjects(getSheet("id_cards")).forEach(c => {
    if (String(c.type || "").toLowerCase() !== type) return;
    const m = String(c.code || "").match(new RegExp("^" + prefix + "(\\d+)$", "i"));
    if (m) {
      const n = parseInt(m[1], 10);
      if (n > max) max = n;
    }
  });

  // From users sheet
  sheetToObjects(getSheet("users")).forEach(u => {
    const code = type === "faculty" ? u.faculty_id : u.student_id;
    const m = String(code || "").match(new RegExp("^" + prefix + "(\\d+)$", "i"));
    if (m) {
      const n = parseInt(m[1], 10);
      if (n > max) max = n;
    }
  });

  // From students / faculty sheets
  if (type === "student") {
    sheetToObjects(getSheet("students")).forEach(s => {
      const m = String(s.student_id || "").match(/^STU(\d+)$/i);
      if (m) {
        const n = parseInt(m[1], 10);
        if (n > max) max = n;
      }
    });
  } else {
    sheetToObjects(getSheet("faculty")).forEach(f => {
      const m = String(f.faculty_id || "").match(/^FAC(\d+)$/i);
      if (m) {
        const n = parseInt(m[1], 10);
        if (n > max) max = n;
      }
    });
  }

  return prefix + String(max + 1).padStart(3, "0");
}

/**
 * Issue a new Student or Faculty ID (auto-generated code).
 * Called from the ID Cards page "Generate ID" button.
 */
function issueId(type, label, recipientEmail) {
  type = String(type || "student").toLowerCase();
  if (type !== "student" && type !== "faculty") {
    return { ok: false, error: "Type must be student or faculty." };
  }
  recipientEmail = String(recipientEmail || "").trim();
  if (!recipientEmail) {
    return { ok: false, error: "Recipient email is required." };
  }

  const cards = sheetToObjects(getSheet("id_cards"));
  const dup = cards.find(c =>
    c.recipient_email &&
    String(c.recipient_email).toLowerCase() === recipientEmail.toLowerCase() &&
    String(c.type || "").toLowerCase() === type &&
    String(c.status || "available").toLowerCase() !== "used"
  );
  if (dup) {
    return { ok: false, error: "An unused " + type + " ID already exists for this email: " + dup.code };
  }

  const code = nextCode(type);
  const issued = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const row = {
    code: code,
    type: type,
    label: String(label || "").trim() || "",
    recipient_email: recipientEmail,
    status: "available",
    issued_date: issued,
    email_sent_date: "",
    used_by: "",
    used_date: ""
  };
  const result = insertRow("id_cards", row);
  return {
    ok: true,
    id: result.id,
    code: code,
    type: type,
    message: (type === "faculty" ? "Faculty" : "Student") + " ID generated: " + code
  };
}

/** Send the issued ID to the recipient by email (optional). */
function sendIdEmail(id) {
  if (!id) return { ok: false, error: "Missing ID card id." };
  const row = getRow("id_cards", id);
  if (!row.data) return { ok: false, error: "ID card not found." };
  const card = row.data;
  if (!card.recipient_email) return { ok: false, error: "No recipient email on this ID." };
  if (String(card.status || "").toLowerCase() === "used") {
    return { ok: false, error: "This ID has already been used." };
  }

  const kind = card.type === "faculty" ? "Faculty" : "Student";
  const subject = "Your " + kind + " ID – Mehedi's University Management System";
  const body =
    "Hello" + (card.label ? " " + card.label : "") + ",\n\n" +
    "Your " + kind + " ID has been issued:\n\n" +
    "  ID Code : " + card.code + "\n" +
    "  Email   : " + card.recipient_email + "\n\n" +
    "Sign up on the University Management System using this exact email address. " +
    "The system will automatically link your account to this ID.\n\n" +
    "— Mehedi's University Management System";

  try {
    MailApp.sendEmail(card.recipient_email, subject, body);
    const sent = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
    updateRow("id_cards", id, { email_sent_date: sent });
    return { ok: true, message: "Email sent to " + card.recipient_email };
  } catch (err) {
    return { ok: false, error: "Could not send email: " + (err.message || String(err)) };
  }
}

function getAvatarFolder() {
  const name = "UMS Profile Photos";
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

function drivePhotoUrl(fileId) {
  return "https://drive.google.com/thumbnail?id=" + fileId + "&sz=w400";
}

function saveBlobPhoto(userId, dataUrl) {
  const match = String(dataUrl).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
  if (!match) throw new Error("Invalid image data. Please choose a JPG or PNG photo.");
  const mime = match[1];
  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 2 * 1024 * 1024) throw new Error("Photo is too large. Use an image under 2 MB.");
  const ext = mime.indexOf("png") !== -1 ? "png" : "jpg";
  const blob = Utilities.newBlob(bytes, mime, "avatar_" + userId + "." + ext);

  const folder = getAvatarFolder();
  const old = folder.getFilesByName("avatar_" + userId + ".jpg");
  while (old.hasNext()) old.next().setTrashed(true);
  const oldPng = folder.getFilesByName("avatar_" + userId + ".png");
  while (oldPng.hasNext()) oldPng.next().setTrashed(true);

  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { fileId: file.getId(), url: drivePhotoUrl(file.getId()) };
}

function saveProfilePhoto(userId, photoData) {
  if (!userId) return { ok: false, error: "Missing user id." };
  if (!photoData) return { ok: false, error: "No photo provided." };

  const row = getRow("users", userId);
  if (!row.data) return { ok: false, error: "User not found." };

  let photo_url = "";
  let photo_file_id = "";

  try {
    const saved = saveBlobPhoto(userId, photoData);
    photo_url = saved.url;
    photo_file_id = saved.fileId;
  } catch (err) {
    // Drive may be blocked; store a compressed data URL in the sheet instead.
    if (String(photoData).length > 45000) {
      return { ok: false, error: "Photo is too large to save. Try a smaller image." };
    }
    photo_url = photoData;
  }

  updateRow("users", userId, { photo_url: photo_url, photo_file_id: photo_file_id });
  return { ok: true, photo_url: photo_url, message: "Profile picture saved." };
}

function removeProfilePhoto(userId) {
  if (!userId) return { ok: false, error: "Missing user id." };
  const row = getRow("users", userId);
  if (!row.data) return { ok: false, error: "User not found." };

  if (row.data.photo_file_id) {
    try {
      const file = DriveApp.getFileById(String(row.data.photo_file_id));
      file.setTrashed(true);
    } catch (_) {}
  }

  updateRow("users", userId, { photo_url: "", photo_file_id: "" });
  return { ok: true, photo_url: getGravatarUrl(row.data.email), message: "Profile picture removed." };
}

function saveProfile(userId, data) {
  if (!userId) return { ok: false, error: "Missing user id." };
  const row = getRow("users", userId);
  if (!row.data) return { ok: false, error: "User not found." };

  const patch = {};
  if (data.full_name !== undefined) patch.full_name = String(data.full_name || "").trim();

  let photo_url = row.data.photo_url || "";
  if (data.remove_photo) {
    const removed = removeProfilePhoto(userId);
    photo_url = removed.photo_url || "";
  } else if (data.photo || data.photo_url) {
    const saved = saveProfilePhoto(userId, data.photo || data.photo_url);
    if (!saved.ok) return saved;
    photo_url = saved.photo_url;
  }

  if (Object.keys(patch).length) updateRow("users", userId, patch);

  const fresh = getRow("users", userId).data;
  return {
    ok: true,
    message: "Profile saved.",
    user: publicUser(Object.assign({}, fresh, { photo_url: photo_url || userPhoto(fresh) }))
  };
}
