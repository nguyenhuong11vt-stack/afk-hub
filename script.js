function doGet() {
  initDB(); // Đảm bảo Database luôn sẵn sàng
  return HtmlService
    .createHtmlOutputFromFile("index")
    .setTitle("AFK HUB")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==========================================
// 1. CẤU HÌNH DATABASE
// ==========================================
function getDB() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function initDB() {
  var ss = getDB();
  var sheets = {
    'Users': ['ID', 'Username', 'PasswordHash', 'DisplayName', 'Symbol', 'Role', 'Token', 'CreatedAt'],
    'AFK': ['UserID', 'StartTime'],
    'Sessions': ['SessionID', 'UserID', 'StartTime', 'EndTime', 'Duration'],
    'Scores': ['UserID', 'Game', 'BestScore', 'UpdatedAt']
  };

  for (var sheetName in sheets) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(sheets[sheetName]);
      // Định dạng header
      sheet.getRange(1, 1, 1, sheets[sheetName].length).setFontWeight("bold").setBackground("#d9d9d9");
      sheet.setFrozenRows(1);
    }
  }
}

// ==========================================
// 2. TIỆN ÍCH HỖ TRỢ
// ==========================================
function generateUUID() {
  return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
    var r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
    return v.toString(16);
  });
}

function hashPassword(password) {
  var rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + "AFK_SECRET_SALT");
  var txtHash = '';
  for (var i = 0; i < rawHash.length; i++) {
    var hashVal = rawHash[i];
    if (hashVal < 0) {
      hashVal += 256;
    }
    if (hashVal.toString(16).length == 1) {
      txtHash += '0';
    }
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

// ==========================================
// 3. HỆ THỐNG XÁC THỰC & TÀI KHOẢN
// ==========================================
function registerUser(username, password, display, symbol) {
  try {
    var sheet = getDB().getSheetByName('Users');
    var data = sheet.getDataRange().getValues();
    
    // Kiểm tra trùng lặp
    for (var i = 1; i < data.length; i++) {
      if (data[i][1].toString().toLowerCase() === username.toLowerCase()) {
        return { success: false, message: 'Username đã tồn tại!' };
      }
    }

    var role = (username.toLowerCase() === 'mayo') ? 'OWNER' : 'MEMBER';
    var id = generateUUID();
    var hash = hashPassword(password);
    var date = new Date().getTime();

    sheet.appendRow([id, username, hash, display, symbol, role, '', date]);
    return { success: true, message: 'Tạo tài khoản thành công!' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function loginUser(username, password) {
  try {
    var sheet = getDB().getSheetByName('Users');
    var data = sheet.getDataRange().getValues();
    var hash = hashPassword(password);

    for (var i = 1; i < data.length; i++) {
      if (data[i][1].toString().toLowerCase() === username.toLowerCase() && data[i][2] === hash) {
        var token = generateUUID();
        sheet.getRange(i + 1, 7).setValue(token); // Cập nhật token
        
        return {
          success: true,
          token: token,
          user: {
            id: data[i][0],
            username: data[i][1],
            displayName: data[i][3],
            symbol: data[i][4],
            role: data[i][5]
          }
        };
      }
    }
    return { success: false, message: 'Sai username hoặc mật khẩu!' };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function getUserByToken(token) {
  if (!token) return null;
  var sheet = getDB().getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][6] === token) {
      return {
        id: data[i][0],
        username: data[i][1],
        displayName: data[i][3],
        symbol: data[i][4],
        role: data[i][5],
        row: i + 1
      };
    }
  }
  return null;
}

function checkAuth(token) {
  var user = getUserByToken(token);
  if (user) return { success: true, user: user };
  return { success: false, message: 'Bạn chưa đăng nhập hoặc phiên hết hạn.' };
}

function logoutUser(token) {
  var user = getUserByToken(token);
  if (user) {
    var sheet = getDB().getSheetByName('Users');
    sheet.getRange(user.row, 7).setValue('');
  }
  return { success: true };
}

function updateProfile(token, displayName, symbol) {
  var auth = checkAuth(token);
  if (!auth.success) return auth;
  
  var sheet = getDB().getSheetByName('Users');
  sheet.getRange(auth.user.row, 4).setValue(displayName);
  sheet.getRange(auth.user.row, 5).setValue(symbol);
  
  return { success: true, message: 'Cập nhật thành công!' };
}

// ==========================================
// 4. HỆ THỐNG AFK
// ==========================================
function startAFK(token) {
  var auth = checkAuth(token);
  if (!auth.success) return auth;
  
  var sheet = getDB().getSheetByName('AFK');
  var data = sheet.getDataRange().getValues();
  
  // Kiểm tra đang AFK chưa
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === auth.user.id) {
      return { success: false, message: 'Tài khoản này đang AFK!' };
    }
  }
  
  var startTime = new Date().getTime();
  sheet.appendRow([auth.user.id, startTime]);
  return { success: true, startTime: startTime };
}

function stopAFK(token) {
  var auth = checkAuth(token);
  if (!auth.success) return auth;
  return processStopAFK(auth.user.id);
}

function processStopAFK(userId) {
  var afkSheet = getDB().getSheetByName('AFK');
  var afkData = afkSheet.getDataRange().getValues();
  
  for (var i = 1; i < afkData.length; i++) {
    if (afkData[i][0] === userId) {
      var startTime = afkData[i][1];
      var endTime = new Date().getTime();
      var duration = endTime - startTime;
      
      // Xóa khỏi bảng AFK
      afkSheet.deleteRow(i + 1);
      
      // Lưu vào Sessions
      var sessionSheet = getDB().getSheetByName('Sessions');
      sessionSheet.appendRow([generateUUID(), userId, startTime, endTime, duration]);
      
      return { success: true, message: 'Đã kết thúc AFK' };
    }
  }
  return { success: false, message: 'Không tìm thấy phiên AFK' };
}

function getAFKRanking(token) {
  var usersSheet = getDB().getSheetByName('Users');
  var afkSheet = getDB().getSheetByName('AFK');
  
  var usersData = usersSheet.getDataRange().getValues();
  var afkData = afkSheet.getDataRange().getValues();
  
  var usersMap = {};
  for (var i = 1; i < usersData.length; i++) {
    usersMap[usersData[i][0]] = {
      display: usersData[i][3],
      symbol: usersData[i][4],
      role: usersData[i][5]
    };
  }
  
  var ranking = [];
  var now = new Date().getTime();
  
  for (var j = 1; j < afkData.length; j++) {
    var uid = afkData[j][0];
    if (usersMap[uid]) {
      ranking.push({
        display: usersMap[uid].display,
        symbol: usersMap[uid].symbol,
        role: usersMap[uid].role,
        startTime: afkData[j][1],
        duration: now - afkData[j][1]
      });
    }
  }
  
  // Sắp xếp: Duration giảm dần, nếu bằng thì StartTime nhỏ hơn lên trước
  ranking.sort(function(a, b) {
    if (b.duration === a.duration) {
      return a.startTime - b.startTime;
    }
    return b.duration - a.duration;
  });
  
  return { success: true, ranking: ranking };
}

// ==========================================
// 5. HỆ THỐNG TRÒ CHƠI (SCORES)
// ==========================================
function saveGameScore(token, game, score) {
  var auth = checkAuth(token);
  if (!auth.success) return auth;
  
  var sheet = getDB().getSheetByName('Scores');
  var data = sheet.getDataRange().getValues();
  var now = new Date().getTime();
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === auth.user.id && data[i][1] === game) {
      if (score > data[i][2]) {
        sheet.getRange(i + 1, 3).setValue(score);
        sheet.getRange(i + 1, 4).setValue(now);
        return { success: true, newBest: true, score: score };
      }
      return { success: true, newBest: false, score: data[i][2] };
    }
  }
  
  sheet.appendRow([auth.user.id, game, score, now]);
  return { success: true, newBest: true, score: score };
}

function getProfile(token) {
  var auth = checkAuth(token);
  if (!auth.success) return auth;
  
  var uid = auth.user.id;
  
  // Get Sessions
  var sessionSheet = getDB().getSheetByName('Sessions');
  var sessionData = sessionSheet.getDataRange().getValues();
  var totalSessions = 0;
  var longestAFK = 0;
  var history = [];
  
  for (var i = sessionData.length - 1; i >= 1; i--) {
    if (sessionData[i][1] === uid) {
      totalSessions++;
      if (sessionData[i][4] > longestAFK) longestAFK = sessionData[i][4];
      if (history.length < 5) {
        history.push({
          start: sessionData[i][2],
          duration: sessionData[i][4]
        });
      }
    }
  }
  
  // Get Current AFK
  var afkSheet = getDB().getSheetByName('AFK');
  var afkData = afkSheet.getDataRange().getValues();
  var currentAFK = null;
  for (var j = 1; j < afkData.length; j++) {
    if (afkData[j][0] === uid) {
      currentAFK = afkData[j][1];
      break;
    }
  }
  
  // Get Scores
  var scoreSheet = getDB().getSheetByName('Scores');
  var scoreData = scoreSheet.getDataRange().getValues();
  var flappyScore = 0;
  for (var k = 1; k < scoreData.length; k++) {
    if (scoreData[k][0] === uid && scoreData[k][1] === 'FlappyBird') {
      flappyScore = scoreData[k][2];
      break;
    }
  }
  
  return {
    success: true,
    user: auth.user,
    stats: {
      totalSessions: totalSessions,
      longestAFK: longestAFK,
      currentAFK: currentAFK,
      flappyScore: flappyScore,
      history: history
    }
  };
}

// ==========================================
// 6. ADMIN PANEL (CHỈ OWNER)
// ==========================================
function getAdminStats(token) {
  var auth = checkAuth(token);
  if (!auth.success) return auth;
  if (auth.user.role !== 'OWNER') return { success: false, message: 'Bạn không có quyền OWNER' };
  
  var usersCount = getDB().getSheetByName('Users').getLastRow() - 1;
  var afkCount = getDB().getSheetByName('AFK').getLastRow() - 1;
  var sessionCount = getDB().getSheetByName('Sessions').getLastRow() - 1;
  var gameCount = getDB().getSheetByName('Scores').getLastRow() - 1;
  
  return {
    success: true,
    stats: { users: usersCount, afks: afkCount, sessions: sessionCount, games: gameCount }
  };
}

function getUsers(token) {
  var auth = checkAuth(token);
  if (!auth.success) return auth;
  if (auth.user.role !== 'OWNER') return { success: false, message: 'Bạn không có quyền OWNER' };
  
  var sheet = getDB().getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  var users = [];
  for (var i = 1; i < data.length; i++) {
    users.push({
      id: data[i][0],
      username: data[i][1],
      displayName: data[i][3],
      symbol: data[i][4],
      role: data[i][5]
    });
  }
  return { success: true, users: users };
}

function adminStopAFK(token, targetId) {
  var auth = checkAuth(token);
  if (!auth.success || auth.user.role !== 'OWNER') return { success: false, message: 'Từ chối quyền truy cập' };
  return processStopAFK(targetId);
}

function adminDeleteUser(token, targetId) {
  var auth = checkAuth(token);
  if (!auth.success || auth.user.role !== 'OWNER') return { success: false, message: 'Từ chối quyền truy cập' };
  
  var sheet = getDB().getSheetByName('Users');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === targetId) {
      if (data[i][5] === 'OWNER') {
        return { success: false, message: 'Không thể xóa tài khoản OWNER!' };
      }
      sheet.deleteRow(i + 1);
      return { success: true, message: 'Đã xóa người dùng' };
    }
  }
  return { success: false, message: 'Không tìm thấy người dùng' };
}
