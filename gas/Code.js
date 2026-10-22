// ============================================================
// 🧒 Weekling 學齡前玩伴配對 - Google Apps Script
//
// 這份檔案對應 Apps Script 專案裡的「程式碼」。
// 部署方式：貼上覆蓋原本的程式，或用 clasp push。
// 推上去之後要「部署 > 管理部署作業 > 編輯 > 新版本」，
// 網址可以維持不變。執行身分維持「我」，存取權維持「任何人」。
//
// 請不要執行 setup()，它會清空試算表。
// ============================================================

var SHEET_ID = "16VJ9T8aNCdQnXMDboxsm2ecBcGWu0t3gB-YbHN6pyaM";
var LIFF_ID = "2008649114-wxWy3jJs";
var LINE_TOKEN = "";
var TIMEZONE = "Asia/Taipei";

var SHEETS = {
  parents: "家長資料",
  venues: "場所資料",
  matches: "配對紀錄",
  schedule: "可約時段",
  request: "場所申請",
  tags: "標籤設定",
  feedback: "意見回饋"
};

var SLOT_NAMES = ["週六上午", "週六下午", "週日上午", "週日下午"];

// ============================================================
// 工具
// ============================================================

function getSs() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getSheet(name) {
  var sheet = getSs().getSheetByName(name);
  if (!sheet) throw new Error("找不到分頁: " + name);
  return sheet;
}

function getHeaderInfo(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var lastRow = Math.max(sheet.getLastRow(), 1);
  var rows = sheet.getRange(1, 1, Math.min(lastRow, 6), lastCol).getValues();
  var markers = ["LINE UID", "提交時間", "類型", "場所ID", "配對ID", "申請ID"];
  for (var r = 0; r < rows.length; r++) {
    var cells = rows[r].map(function (c) { return String(c).trim(); });
    for (var i = 0; i < markers.length; i++) {
      if (cells.indexOf(markers[i]) !== -1) {
        return { headerRow: r + 1, headers: rows[r] };
      }
    }
  }
  return { headerRow: Math.min(2, rows.length), headers: rows[Math.min(1, rows.length - 1)] || [] };
}

// 同名欄位（兩個都叫「狀態」）讀取時以後面那個為準，寫入也寫後面那個。
function updateCellByHeader(sheet, rowNum, headerName, value) {
  var info = getHeaderInfo(sheet);
  var col = 0;
  for (var i = 0; i < info.headers.length; i++) {
    if (String(info.headers[i]).trim() === headerName) col = i + 1;
  }
  if (!col) {
    console.warn("找不到欄位: " + headerName);
    return;
  }
  sheet.getRange(rowNum, col).setValue(value);
}

function appendByHeaders(sheet, map) {
  var info = getHeaderInfo(sheet);
  var counts = {};
  info.headers.forEach(function (h) {
    var key = String(h).trim();
    if (key) counts[key] = (counts[key] || 0) + 1;
  });
  var seen = {};
  var row = info.headers.map(function (h) {
    var key = String(h).trim();
    if (!key || !Object.prototype.hasOwnProperty.call(map, key)) return "";
    seen[key] = (seen[key] || 0) + 1;
    if (key === "狀態" && seen[key] < counts[key]) return "";
    return map[key];
  });
  sheet.appendRow(row);
}

function getAllData(sheetName) {
  var sheet = getSheet(sheetName);
  var info = getHeaderInfo(sheet);
  var values = sheet.getDataRange().getValues();
  var out = [];
  for (var r = info.headerRow; r < values.length; r++) {
    var row = values[r];
    var blank = true;
    for (var c = 0; c < row.length; c++) {
      if (row[c] !== "" && row[c] != null) { blank = false; break; }
    }
    if (blank) continue;
    var obj = {};
    info.headers.forEach(function (h, i) {
      var key = String(h).trim();
      if (!key) return;
      obj[key] = row[i];
    });
    out.push(obj);
  }
  return out;
}

function findRowByHeader(sheet, headerName, value) {
  if (value == null || String(value).trim() === "") return -1;
  var info = getHeaderInfo(sheet);
  var col = -1;
  for (var i = 0; i < info.headers.length; i++) {
    if (String(info.headers[i]).trim() === headerName) { col = i; break; }
  }
  if (col < 0) return -1;
  var values = sheet.getDataRange().getValues();
  var want = String(value).trim();
  for (var r = info.headerRow; r < values.length; r++) {
    if (String(values[r][col] == null ? "" : values[r][col]).trim() === want) return r + 1;
  }
  return -1;
}

function findByUid(rows, lineUid) {
  var want = String(lineUid || "").trim();
  if (!want) return null;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i]["LINE UID"] || "").trim() === want) return rows[i];
  }
  return null;
}

function todayText() {
  return Utilities.formatDate(new Date(), TIMEZONE, "yyyy/M/d");
}

function formatDate(value) {
  if (value == null || value === "") return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TIMEZONE, "yyyy-MM-dd");
  }
  var s = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  var slash = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})$/);
  if (slash) return slash[1] + "-" + ("0" + slash[2]).slice(-2) + "-" + ("0" + slash[3]).slice(-2);
  var parsed = new Date(s);
  if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, TIMEZONE, "yyyy-MM-dd");
  return s;
}

function parseDateInput(value) {
  if (value == null || value === "") return "";
  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) return value;
  var m = String(value).trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return String(value);
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function ageOf(birthday) {
  var day = formatDate(birthday);
  var born = day.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!born) return null;
  var today = Utilities.formatDate(new Date(), TIMEZONE, "yyyy-MM-dd").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  var age = Number(today[1]) - Number(born[1]);
  if (today[2] + today[3] < born[2] + born[3]) age--;
  if (age < 0 || age > 30) return null;
  return age;
}

function asTagList(value) {
  if (Array.isArray(value)) {
    return value.map(function (v) { return String(v).trim(); }).filter(function (v) { return v; });
  }
  if (value == null || value === "") return [];
  return String(value).split(/[、,，\n]/).map(function (s) { return s.trim(); }).filter(function (s) { return s; });
}

function joinTags() {
  for (var i = 0; i < arguments.length; i++) {
    var list = asTagList(arguments[i]);
    if (list.length) return list.join("、");
  }
  return "";
}

function generateId(prefix) {
  var randomStr = Math.random().toString(36).substring(2, 5).toUpperCase();
  var timeStr = Date.now().toString().slice(-4);
  return prefix + "-" + randomStr + timeStr;
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getTagSheets() {
  var sheets = getSs().getSheets().filter(function (s) {
    return s.getName().trim() === "標籤設定";
  });
  if (!sheets.length) throw new Error("找不到分頁: 標籤設定");
  return sheets;
}

// ============================================================
// 讀取標籤與場所
// ============================================================

function getConfig() {
  var tags = { personality: [], activity: [], emotion: [], energy: [], disliked: [] };
  var seen = { personality: {}, activity: {}, emotion: {}, energy: {}, disliked: {} };
  var typeMap = { "個性": "personality", "活動": "activity", "情緒": "emotion", "體力": "energy", "不喜歡": "disliked" };

  getTagSheets().forEach(function (sheet) {
    var info = getHeaderInfo(sheet);
    var values = sheet.getDataRange().getValues();
    var typeCol = -1, labelCol = -1, statusCol = -1;
    info.headers.forEach(function (h, i) {
      var key = String(h).trim();
      if (key === "類型") typeCol = i;
      if (key === "標籤文字") labelCol = i;
      if (key === "狀態") statusCol = i;
    });
    for (var r = info.headerRow; r < values.length; r++) {
      var type = typeCol >= 0 ? values[r][typeCol] : values[r][0];
      var label = labelCol >= 0 ? values[r][labelCol] : values[r][1];
      var active = statusCol >= 0 ? values[r][statusCol] : values[r][2];
      var bucket = typeMap[String(type || "").trim()];
      var text = String(label || "").trim();
      if (!bucket || !text || String(active || "").trim() === "停用" || seen[bucket][text]) continue;
      seen[bucket][text] = true;
      tags[bucket].push(text);
    }
  });

  var venueSheet = getSheet(SHEETS.venues);
  var venueInfo = getHeaderInfo(venueSheet);
  var venueValues = venueSheet.getDataRange().getValues();
  function venueCol(name) {
    for (var i = 0; i < venueInfo.headers.length; i++) {
      if (String(venueInfo.headers[i]).trim() === name) return i;
    }
    return -1;
  }
  var venues = [];
  for (var r = venueInfo.headerRow; r < venueValues.length; r++) {
    if (String(venueValues[r][venueCol("狀態")] || "").trim() !== "上架") continue;
    venues.push({
      id: venueValues[r][venueCol("場所ID")],
      name: venueValues[r][venueCol("場所名稱")],
      type: venueValues[r][venueCol("類型")],
      area: venueValues[r][venueCol("地區")],
      address: venueValues[r][venueCol("地址")],
      ageRange: venueValues[r][venueCol("適合年齡")]
    });
  }

  return { success: true, tags: tags, venues: venues };
}

// ============================================================
// 家長註冊 / 更新
// ============================================================

function parentFieldMap(data, now, isNew) {
  var map = {
    "LINE UID": data.lineUid,
    "家長暱稱": data.nickname || "",
    "地區": data.area || "",
    "小孩姓名": data.childName || "",
    "小孩生日": parseDateInput(data.childBirthday),
    "小孩性別": data.childGender || "",
    "就讀幼兒園／托嬰": data.kindergarten || "",
    "生肖": data.zodiac || "",
    "社交個性": joinTags(data.personalityTags, data.personality),
    "情緒特質": joinTags(data.emotionTags, data.emotionTag, data.emotion),
    "體力類型": joinTags(data.energyTags, data.energyTag, data.energy),
    "喜歡的活動": joinTags(data.likedActivities),
    "LINE ID": data.lineId || "",
    "不喜歡的活動": joinTags(data.dislikedActivities),
    "家長自介或期許": data.parentIntro || "",
    "最後更新": now
  };
  if (isNew) {
    map["狀態"] = "啟用";
    map["加入日期"] = now;
  }
  return map;
}

function scheduleFieldMap(data) {
  return {
    "LINE UID": data.lineUid,
    "家長暱稱": data.nickname || "",
    "偏好場所": data.preferVenues || "",
    "週六上午": data.satAM ? "✅" : "",
    "週六下午": data.satPM ? "✅" : "",
    "週日上午": data.sunAM ? "✅" : "",
    "週日下午": data.sunPM ? "✅" : ""
  };
}

function registerParent(data) {
  if (!data || !data.lineUid) return { success: false, message: "缺少 LINE UID" };
  var parentSheet = getSheet(SHEETS.parents);
  var scheduleSheet = getSheet(SHEETS.schedule);
  var now = todayText();
  var parentRow = findRowByHeader(parentSheet, "LINE UID", data.lineUid);
  var fields = parentFieldMap(data, now, parentRow < 0);

  if (parentRow < 0) {
    appendByHeaders(parentSheet, fields);
  } else {
    Object.keys(fields).forEach(function (key) {
      if (key === "LINE UID") return;
      updateCellByHeader(parentSheet, parentRow, key, fields[key]);
    });
  }

  var slots = scheduleFieldMap(data);
  var scheduleRow = findRowByHeader(scheduleSheet, "LINE UID", data.lineUid);
  if (scheduleRow < 0) {
    appendByHeaders(scheduleSheet, slots);
  } else {
    Object.keys(slots).forEach(function (key) {
      if (key === "LINE UID") return;
      updateCellByHeader(scheduleSheet, scheduleRow, key, slots[key]);
    });
  }

  return { success: true, message: "資料已同步！" };
}

function getProfile(lineUid) {
  var me = findByUid(getAllData(SHEETS.parents), lineUid);
  if (!me) return { success: false, message: "尚未註冊" };
  var sch = findByUid(getAllData(SHEETS.schedule), lineUid) || {};
  return {
    success: true,
    profile: {
      nickname: me["家長暱稱"] || "",
      area: me["地區"] || "",
      childName: me["小孩姓名"] || "",
      childBirthday: formatDate(me["小孩生日"]),
      childGender: me["小孩性別"] || "",
      zodiac: me["生肖"] || "",
      kindergarten: me["就讀幼兒園／托嬰"] || "",
      personality: me["社交個性"] || "",
      emotion: me["情緒特質"] || "",
      energy: me["體力類型"] || "",
      likedActivities: me["喜歡的活動"] || "",
      dislikedActivities: me["不喜歡的活動"] || "",
      parentIntro: me["家長自介或期許"] || "",
      lineId: me["LINE ID"] || "",
      satAM: sch["週六上午"] === "✅",
      satPM: sch["週六下午"] === "✅",
      sunAM: sch["週日上午"] === "✅",
      sunPM: sch["週日下午"] === "✅",
      status: me["狀態"] || ""
    }
  };
}

function updateStatus(lineUid, status) {
  var sheet = getSheet(SHEETS.parents);
  var row = findRowByHeader(sheet, "LINE UID", lineUid);
  if (row < 0) return { success: false, message: "尚未註冊" };
  updateCellByHeader(sheet, row, "狀態", status || "");
  updateCellByHeader(sheet, row, "最後更新", todayText());
  return { success: true };
}

function confirmActive(lineUid) {
  return updateStatus(lineUid, "啟用");
}

// ============================================================
// 配對
// ============================================================

function findMatches(lineUid) {
  var parents = getAllData(SHEETS.parents);
  var schedules = getAllData(SHEETS.schedule);
  var me = findByUid(parents, lineUid);
  if (!me || me["狀態"] !== "啟用") return { success: false, message: "帳號狀態異常" };
  var mySch = findByUid(schedules, lineUid);
  var matches = [];
  parents.forEach(function (other) {
    var uid = String(other["LINE UID"] || "").trim();
    if (!uid || uid === String(lineUid).trim()) return;
    if (other["狀態"] !== "啟用") return;
    var otherSch = findByUid(schedules, uid);
    var common = [];
    if (mySch && otherSch) {
      SLOT_NAMES.forEach(function (slot) {
        if (mySch[slot] === "✅" && otherSch[slot] === "✅") common.push(slot);
      });
    }
    matches.push({
      lineUid: uid,
      nickname: other["家長暱稱"] || "",
      area: other["地區"] || "",
      childName: other["小孩姓名"] || "",
      childAge: ageOf(other["小孩生日"]),
      childGender: other["小孩性別"] || "",
      zodiac: other["生肖"] || "",
      personality: other["社交個性"] || "",
      emotion: other["情緒特質"] || "",
      energy: other["體力類型"] || "",
      likedActivities: other["喜歡的活動"] || "",
      parentIntro: other["家長自介或期許"] || "",
      lineId: other["LINE ID"] || "",
      commonSlots: common
    });
  });
  return { success: true, count: matches.length, matches: matches };
}

function createMatch(fromUid, toUid, venueId, date, timeSlot) {
  var sheet = getSheet(SHEETS.matches);
  var matchId = generateId("M");
  appendByHeaders(sheet, {
    "配對ID": matchId,
    "發起人UID": fromUid || "",
    "被邀請人UID": toUid || "",
    "場所ID": venueId || "",
    "約定日期": parseDateInput(date),
    "時段": timeSlot || "",
    "狀態": "等待確認",
    "家長A確認": "✅",
    "家長B確認": "⏳",
    "建立日期": todayText()
  });
  return { success: true, matchId: matchId, message: "邀請已送出！" };
}

function confirmMatch(matchId, lineUid) {
  var sheet = getSheet(SHEETS.matches);
  var info = getHeaderInfo(sheet);
  var values = sheet.getDataRange().getValues();
  function col(name) {
    for (var i = 0; i < info.headers.length; i++) {
      if (String(info.headers[i]).trim() === name) return i;
    }
    return -1;
  }
  var idCol = col("配對ID");
  var toCol = col("被邀請人UID");
  for (var r = info.headerRow; r < values.length; r++) {
    if (String(values[r][idCol]) === String(matchId) && String(values[r][toCol]).trim() === String(lineUid || "").trim()) {
      updateCellByHeader(sheet, r + 1, "家長B確認", "✅");
      updateCellByHeader(sheet, r + 1, "狀態", "已確認");
      return { success: true, message: "配對成功！" };
    }
  }
  return { success: false, message: "找不到配對紀錄" };
}

function getMyMatches(lineUid) {
  var matches = getAllData(SHEETS.matches);
  var parents = getAllData(SHEETS.parents);
  var mine = matches.filter(function (m) {
    return String(m["發起人UID"] || "").trim() === String(lineUid || "").trim() ||
      String(m["被邀請人UID"] || "").trim() === String(lineUid || "").trim();
  });
  mine.reverse();
  return {
    success: true,
    matches: mine.map(function (m) {
      var isSender = String(m["發起人UID"] || "").trim() === String(lineUid || "").trim();
      var otherUid = isSender ? m["被邀請人UID"] : m["發起人UID"];
      var other = findByUid(parents, otherUid) || {};
      return {
        matchId: m["配對ID"] || "",
        status: m["狀態"] || "",
        isSender: isSender,
        otherNickname: other["家長暱稱"] || "",
        otherChildName: other["小孩姓名"] || "",
        otherChildAge: ageOf(other["小孩生日"]),
        otherChildGender: other["小孩性別"] || "",
        otherLineId: other["LINE ID"] || "",
        venueId: m["場所ID"] || "",
        date: formatDate(m["約定日期"]),
        timeSlot: m["時段"] || "",
        createdAt: formatDate(m["建立日期"])
      };
    })
  };
}

function getMatchDetail(matchId) {
  var matches = getAllData(SHEETS.matches);
  var found = null;
  for (var i = 0; i < matches.length; i++) {
    if (String(matches[i]["配對ID"]) === String(matchId)) { found = matches[i]; break; }
  }
  if (!found) return { success: false, message: "找不到配對" };
  var from = findByUid(getAllData(SHEETS.parents), found["發起人UID"]) || {};
  return {
    success: true,
    matchId: found["配對ID"] || "",
    status: found["狀態"] || "",
    fromUid: found["發起人UID"] || "",
    fromNickname: from["家長暱稱"] || "",
    fromChildName: from["小孩姓名"] || "",
    fromChildAge: ageOf(from["小孩生日"]),
    fromChildGender: from["小孩性別"] || "",
    fromPersonality: from["社交個性"] || "",
    fromEnergy: from["體力類型"] || "",
    fromLikedActivities: from["喜歡的活動"] || "",
    fromParentIntro: from["家長自介或期許"] || "",
    venueId: found["場所ID"] || "",
    date: formatDate(found["約定日期"]),
    timeSlot: found["時段"] || ""
  };
}

// ============================================================
// 場所申請、意見
// ============================================================

function submitVenueRequest(data) {
  data = data || {};
  var sheet = getSheet(SHEETS.request);
  appendByHeaders(sheet, {
    "申請ID": generateId("R"),
    "申請人UID": data.lineUid || "",
    "場所名稱": data.venueName || "",
    "類型": data.type || "",
    "地區": data.area || "",
    "地址": data.address || "",
    "適合年齡": data.ageRange || "",
    "推薦理由": data.reason || "",
    "申請日期": todayText(),
    "狀態": "待審核"
  });
  return { success: true, message: "申請已送出，審核通過後會出現在場所清單！" };
}

function submitFeedback(body) {
  body = body || {};
  var sheet = getSheet(SHEETS.feedback);
  appendByHeaders(sheet, {
    "提交時間": todayText(),
    "LINE UID": body.lineUid || "",
    "暱稱": body.nickname || "",
    "內容": body.feedback || "",
    "聯絡方式": body.contact || ""
  });
  return { success: true };
}

// ============================================================
// Web App 入口（網址與參數維持原本的用法）
// ============================================================

function doGet(e) {
  try {
    var p = (e && e.parameter) || {};
    var action = p.action || "";
    var result;
    switch (action) {
      case "getProfile": result = getProfile(p.lineUid); break;
      case "findMatches": result = findMatches(p.lineUid); break;
      case "getMyMatches": result = getMyMatches(p.lineUid); break;
      case "getMatchDetail": result = getMatchDetail(p.matchId); break;
      default: result = getConfig();
    }
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ success: false, message: err.toString() });
  }
}

function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var action = body.action;
    var result;
    switch (action) {
      case "register": result = registerParent(body.data); break;
      case "findMatches": result = findMatches(body.lineUid); break;
      case "createMatch": result = createMatch(body.fromUid, body.toUid, body.venueId, body.date, body.timeSlot); break;
      case "confirmMatch": result = confirmMatch(body.matchId, body.lineUid); break;
      case "confirmActive": result = confirmActive(body.lineUid); break;
      case "submitVenue": result = submitVenueRequest(body.data); break;
      case "submitFeedback": result = submitFeedback(body); break;
      case "updateStatus": result = updateStatus(body.lineUid, body.status); break;
      default: result = { success: false, message: "未知的 action: " + action };
    }
    return jsonOut(result);
  } catch (err) {
    return jsonOut({ success: false, message: err.toString() });
  }
}

// ============================================================
// 初始化。只在全新空白試算表使用。執行會清空分頁。
// ============================================================

function setup() {
  throw new Error("setup() 會清空家長、場所、配對等分頁，已停用。真的要重建空白表才呼叫 setupDangerously()。");
}

function setupDangerously() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheetsConfig = [
    {
      name: "家長資料",
      title: "孩子の友情・你來經紀 ── 家長資料",
      headers: ["LINE UID", "家長暱稱", "狀態", "地區", "小孩姓名", "小孩生日", "小孩性別",
        "就讀幼兒園／托嬰", "生肖", "狀態", "加入日期",
        "社交個性", "情緒特質", "體力類型", "喜歡的活動", "LINE ID", "不喜歡的活動",
        "家長自介或期許", "最後更新"]
    },
    {
      name: "場所資料",
      title: "孩子の友情・你來經紀 ── 場所資料",
      headers: ["場所ID", "場所名稱", "類型", "地區", "地址", "適合年齡", "推薦理由", "推薦人UID", "狀態"]
    },
    {
      name: "配對紀錄",
      title: "孩子の友情・你來經紀 ── 配對紀錄",
      headers: ["配對ID", "發起人UID", "被邀請人UID", "場所ID", "約定日期", "時段", "狀態", "家長A確認", "家長B確認", "建立日期"]
    },
    {
      name: "可約時段",
      title: "孩子の友情・你來經紀 ── 可約時段",
      headers: ["LINE UID", "家長暱稱", "偏好場所", "週六上午", "週六下午", "週日上午", "週日下午"]
    },
    {
      name: "場所申請",
      title: "孩子の友情・你來經紀 ── 場所申請",
      headers: ["申請ID", "申請人UID", "場所名稱", "類型", "地區", "地址", "適合年齡", "推薦理由", "申請日期", "狀態"]
    },
    {
      name: "標籤設定",
      title: "孩子の友情・你來經紀 ── 標籤設定",
      headers: ["類型", "標籤文字", "狀態"],
      data: [
        ["個性", "活潑外向", "啟用"], ["個性", "文靜內斂", "啟用"], ["個性", "慢熱但熟了很好", "啟用"],
        ["活動", "跑跳追逐", "啟用"], ["活動", "積木／樂高", "啟用"],
        ["情緒", "情緒穩定", "啟用"], ["情緒", "容易哭鬧", "啟用"],
        ["體力", "高體力停不下來", "啟用"], ["體力", "喜歡靜態活動", "啟用"],
        ["不喜歡", "人多吵雜", "啟用"], ["不喜歡", "激烈衝撞", "啟用"]
      ]
    },
    {
      name: "意見回饋",
      title: "",
      headers: ["提交時間", "LINE UID", "暱稱", "內容", "聯絡方式"],
      headerOnFirstRow: true
    }
  ];

  sheetsConfig.forEach(function (config) {
    var sheet = ss.getSheetByName(config.name);
    if (!sheet) sheet = ss.insertSheet(config.name);
    else sheet.clear();
    var headerRow = config.headerOnFirstRow ? 1 : 2;
    if (!config.headerOnFirstRow) {
      sheet.getRange(1, 1).setValue(config.title);
      sheet.getRange(1, 1, 1, config.headers.length).merge()
        .setBackground("#32E1C8").setFontColor("#ffffff")
        .setFontWeight("bold").setFontSize(13)
        .setHorizontalAlignment("center");
    }
    sheet.getRange(headerRow, 1, 1, config.headers.length)
      .setValues([config.headers])
      .setBackground("#E17D00").setFontColor("#ffffff")
      .setFontWeight("bold").setHorizontalAlignment("center");
    sheet.setFrozenRows(headerRow);
    if (config.data) {
      sheet.getRange(headerRow + 1, 1, config.data.length, config.data[0].length).setValues(config.data);
    }
  });
}
