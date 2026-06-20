// ============================================================
// 誰是臥底 - Google Apps Script 後端
// 部署方式：
//   1. 開啟 Google 試算表 → 擴充功能 → Apps Script
//   2. 貼上此程式碼，修改 SPREADSHEET_ID
//   3. 部署 → 新增部署 → 網頁應用程式
//      執行身分: 我、存取權: 任何人
//   4. 複製網頁應用程式網址，貼到 index.html 的 API_URL
// ============================================================

// 不需要填 ID — 腳本綁定在試算表上會自動抓取

// ── 九大題庫（順序使用，用完自動從頭輪轉）────────────────────
const CATEGORIES = {
  '食物飲品': [
    ['蘋果', '梨子'],
    ['咖啡', '紅茶'],
    ['漢堡', '三明治'],
    ['壽司', '生魚片'],
    ['珍珠奶茶', '咖啡歐蕾'],
    ['冰淇淋', '雪糕'],
    ['水餃', '湯圓'],
    ['橘子', '柳丁'],
    ['巧克力', '糖果'],
    ['泡麵', '義大利麵'],
  ],
  '動物': [
    ['貓', '狗'],
    ['獅子', '老虎'],
    ['鸚鵡', '麻雀'],
    ['海豚', '鯨魚'],
    ['兔子', '松鼠'],
    ['青蛙', '蟾蜍'],
    ['鱷魚', '蜥蜴'],
    ['孔雀', '火烈鳥'],
    ['蝴蝶', '蜻蜓'],
    ['北極熊', '大熊貓'],
  ],
  '運動休閒': [
    ['足球', '籃球'],
    ['游泳', '潛水'],
    ['跑步', '散步'],
    ['下棋', '打牌'],
    ['爬山', '健行'],
    ['羽毛球', '桌球'],
    ['滑雪', '滑冰'],
    ['瑜珈', '皮拉提斯'],
    ['釣魚', '露營'],
    ['電玩', '桌遊'],
  ],
  '交通工具': [
    ['飛機', '直升機'],
    ['地鐵', '輕軌'],
    ['公車', '計程車'],
    ['摩托車', '腳踏車'],
    ['遊輪', '渡輪'],
    ['高鐵', '火車'],
    ['熱氣球', '滑翔翼'],
    ['轎車', '休旅車'],
    ['救護車', '消防車'],
    ['拖拉機', '挖土機'],
  ],
  '地點場所': [
    ['電影院', '劇場'],
    ['海灘', '游泳池'],
    ['城堡', '宮殿'],
    ['圖書館', '書店'],
    ['醫院', '診所'],
    ['超市', '便利商店'],
    ['公園', '植物園'],
    ['咖啡廳', '茶館'],
    ['健身房', '運動中心'],
    ['機場', '火車站'],
  ],
  '職業身份': [
    ['醫生', '護士'],
    ['老師', '教授'],
    ['警察', '保全'],
    ['廚師', '麵包師'],
    ['律師', '法官'],
    ['記者', '主播'],
    ['消防員', '救生員'],
    ['建築師', '室內設計師'],
    ['司機', '船長'],
    ['藝術家', '設計師'],
  ],
  '日常生活': [
    ['眼鏡', '太陽眼鏡'],
    ['手機', '平板'],
    ['書包', '公事包'],
    ['玫瑰', '鬱金香'],
    ['太陽', '月亮'],
    ['夏天', '冬天'],
    ['鋼琴', '吉他'],
    ['枕頭', '抱枕'],
    ['鬧鐘', '手錶'],
    ['雨傘', '雨衣'],
  ],

  // ── 新增：綠島主題 ────────────────────────────────────────
  '綠島時光': [
    ['潛水', '浮潛'],
    ['珊瑚', '海葵'],
    ['朝日溫泉', '海底溫泉'],
    ['機車', '腳踏車'],
    ['民宿', '帳篷'],
    ['海膽', '海參'],
    ['燈塔', '礁石'],
    ['夜浮潛', '夜釣'],
    ['貝殼', '海玻璃'],
    ['快艇', '玻璃船'],
  ],

  // ── 新增：趣味旅伴（用旅行夥伴的名字當詞對）─────────────
  '趣味旅伴': [
    ['黑輪', 'Sharon'],
    ['Xuan', '慶迪'],
    ['侑年', '孟辰'],
    ['黑輪', '慶迪'],
    ['Sharon', '孟辰'],
    ['Xuan', '侑年'],
    ['黑輪', 'Xuan'],
    ['慶迪', '侑年'],
    ['Sharon', 'Xuan'],
    ['孟辰', '黑輪'],
  ],
};

// ── 工具 ─────────────────────────────────────────────────────
function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const p = e.parameter;
    switch (p.action) {
      case 'createRoom':   return respond(createRoom(p.room, p.player, p.category));
      case 'joinRoom':     return respond(joinRoom(p.room, p.player));
      case 'getRoom':      return respond(getRoom(p.room));
      case 'startGame':    return respond(startGame(p.room, p.player));
      case 'getMyWord':    return respond(getMyWord(p.room, p.player));
      case 'submitDesc':   return respond(submitDesc(p.room, p.player, p.desc));
      case 'submitVote':   return respond(submitVote(p.room, p.player, p.target));
      case 'nextRound':    return respond(nextRound(p.room, p.player));
      default: return respond({success: false, error: '未知指令'});
    }
  } catch(err) {
    return respond({success: false, error: err.message});
  } finally {
    lock.releaseLock();
  }
}

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// ── 順序取得詞對（自動輪轉）─────────────────────────────────
function getNextPair(category) {
  const pairs = CATEGORIES[category];
  if (!pairs) throw new Error('無效題庫：' + category);

  const sheet = getSheet('CategoryIndex');
  const data = sheet.getDataRange().getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === category) {
      const idx = Number(data[i][1]) || 0;
      const pair = pairs[idx % pairs.length];
      // 更新索引
      sheet.getRange(i + 1, 2).setValue(idx + 1);
      return pair;
    }
  }

  // 首次使用此題庫
  sheet.appendRow([category, 1]);
  return pairs[0];
}

// ── 讀寫 GameState ────────────────────────────────────────────
function readState(roomName) {
  const sheet = getSheet('GameState');
  const data = sheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === roomName) {
      return { row: i + 1, state: JSON.parse(data[i][1]) };
    }
  }
  return null;
}

function writeState(row, state) {
  getSheet('GameState').getRange(row, 2).setValue(JSON.stringify(state));
}

// ── 創建房間 ──────────────────────────────────────────────────
function createRoom(roomName, playerName, category) {
  if (!roomName || !playerName) return {success: false, error: '參數缺失'};
  if (!CATEGORIES[category])   return {success: false, error: '無效題庫'};

  const rooms = getSheet('Rooms');
  const data  = rooms.getDataRange().getValues();
  for (const row of data) {
    if (row[0] === roomName) return {success: false, error: '房間名稱已存在'};
  }

  rooms.appendRow([roomName, playerName]);

  const initState = {
    host: playerName, category,
    players: [playerName],
    status: 'waiting',
    round: 0,
    words: {}, isUndercover: {}, alive: {},
    descriptions: [], votes: {},
    eliminated: [], lastResult: null,
    descOrder: [], currentDescIdx: 0, currentDescriber: null,
    wordPair: null, winner: null,
  };
  getSheet('GameState').appendRow([roomName, JSON.stringify(initState)]);

  return {success: true, players: [playerName], host: playerName, category};
}

// ── 加入房間 ──────────────────────────────────────────────────
function joinRoom(roomName, playerName) {
  if (!roomName || !playerName) return {success: false, error: '參數缺失'};

  const rooms = getSheet('Rooms');
  const data  = rooms.getDataRange().getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] !== roomName) continue;
    const row = data[i].filter(v => v !== '');

    if (row[row.length - 1] === 'PLAY') return {success: false, error: '遊戲已開始，無法加入'};
    for (let j = 1; j < row.length; j++) {
      if (row[j] === playerName) return {success: false, error: '名稱已被使用'};
    }

    rooms.getRange(i + 1, row.length + 1).setValue(playerName);

    const found = readState(roomName);
    if (!found) return {success: false, error: '找不到房間狀態'};
    found.state.players.push(playerName);
    writeState(found.row, found.state);

    return {success: true, players: found.state.players, host: found.state.host, category: found.state.category};
  }

  return {success: false, error: '房間不存在'};
}

// ── 取得房間狀態 ──────────────────────────────────────────────
function getRoom(roomName) {
  const found = readState(roomName);
  if (!found) return {success: false, error: '房間不存在'};
  const s = found.state;

  const safe = {
    success: true,
    host: s.host, players: s.players, status: s.status,
    category: s.category,
    round: s.round, alive: s.alive,
    descriptions: s.descriptions, votes: s.votes,
    eliminated: s.eliminated, lastResult: s.lastResult,
    currentDescriber: s.currentDescriber, winner: s.winner,
  };

  if (s.status === 'ended') {
    safe.revealWordPair    = s.wordPair;
    safe.revealIsUndercover = s.isUndercover;
  }

  return safe;
}

// ── 開始遊戲 ──────────────────────────────────────────────────
function startGame(roomName, playerName) {
  const found = readState(roomName);
  if (!found) return {success: false, error: '房間不存在'};
  const s = found.state;

  if (s.host !== playerName)  return {success: false, error: '只有房主可以開始'};
  if (s.players.length < 3)  return {success: false, error: '至少需要 3 名玩家'};
  if (s.status !== 'waiting') return {success: false, error: '遊戲已開始'};

  // 順序取得詞對
  const pair = getNextPair(s.category);
  const undercoverCount = s.players.length >= 7 ? 2 : 1;
  const shuffled = [...s.players].sort(() => Math.random() - 0.5);
  const undercoverSet = new Set(shuffled.slice(0, undercoverCount));

  s.wordPair = pair;
  s.words = {}; s.isUndercover = {}; s.alive = {};
  for (const p of s.players) {
    s.isUndercover[p] = undercoverSet.has(p);
    s.words[p] = undercoverSet.has(p) ? pair[1] : pair[0];
    s.alive[p] = true;
  }

  s.status = 'describing';
  s.round = 1;
  s.descriptions = []; s.votes = {}; s.eliminated = []; s.lastResult = null;
  s.descOrder = [...s.players].sort(() => Math.random() - 0.5);
  s.currentDescIdx = 0;
  s.currentDescriber = s.descOrder[0];

  writeState(found.row, s);

  // Rooms 表標記 PLAY
  const rooms = getSheet('Rooms');
  const data  = rooms.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === roomName) {
      const lastCol = data[i].filter(v => v !== '').length + 1;
      rooms.getRange(i + 1, lastCol).setValue('PLAY');
      break;
    }
  }

  return {success: true};
}

// ── 取得自己的詞語 ────────────────────────────────────────────
function getMyWord(roomName, playerName) {
  const found = readState(roomName);
  if (!found) return {success: false, error: '房間不存在'};
  const s = found.state;
  if (!s.words[playerName]) return {success: false, error: '找不到玩家詞語'};
  return {success: true, word: s.words[playerName], isUndercover: s.isUndercover[playerName]};
}

// ── 提交描述 ──────────────────────────────────────────────────
function submitDesc(roomName, playerName, desc) {
  const found = readState(roomName);
  if (!found) return {success: false, error: '房間不存在'};
  const s = found.state;

  if (s.status !== 'describing')    return {success: false, error: '現在不是描述階段'};
  if (s.currentDescriber !== playerName) return {success: false, error: '還沒輪到你'};
  if (!desc || !desc.trim())        return {success: false, error: '描述不能為空'};

  s.descriptions.push({player: playerName, desc: desc.trim(), round: s.round});

  let nextIdx = s.currentDescIdx + 1;
  while (nextIdx < s.descOrder.length && !s.alive[s.descOrder[nextIdx]]) nextIdx++;

  if (nextIdx >= s.descOrder.length) {
    s.status = 'voting';
    s.votes = {};
    s.currentDescriber = null;
  } else {
    s.currentDescIdx = nextIdx;
    s.currentDescriber = s.descOrder[nextIdx];
  }

  writeState(found.row, s);
  return {success: true, status: s.status, currentDescriber: s.currentDescriber};
}

// ── 提交投票 ──────────────────────────────────────────────────
function submitVote(roomName, playerName, target) {
  const found = readState(roomName);
  if (!found) return {success: false, error: '房間不存在'};
  const s = found.state;

  if (s.status !== 'voting')  return {success: false, error: '現在不是投票階段'};
  if (!s.alive[playerName])   return {success: false, error: '你已被淘汰'};
  if (!s.alive[target])       return {success: false, error: '目標已被淘汰'};
  if (playerName === target)  return {success: false, error: '不能投票給自己'};

  s.votes[playerName] = target;

  const alivePlayers = Object.keys(s.alive).filter(p => s.alive[p]);
  if (Object.keys(s.votes).length >= alivePlayers.length) {
    // 計票
    const counts = {};
    for (const t of Object.values(s.votes)) counts[t] = (counts[t] || 0) + 1;

    let maxVotes = 0, eliminated = null, tied = false;
    for (const [p, c] of Object.entries(counts)) {
      if (c > maxVotes)      { maxVotes = c; eliminated = p; tied = false; }
      else if (c === maxVotes) { tied = true; }
    }

    if (tied) {
      s.lastResult = {type: 'tie', round: s.round};
      s.status = 'result';
    } else {
      s.alive[eliminated] = false;
      const wasUndercover = s.isUndercover[eliminated];
      s.eliminated.push({player: eliminated, isUndercover: wasUndercover, round: s.round});
      s.lastResult = {type: 'eliminated', player: eliminated, isUndercover: wasUndercover, round: s.round};

      const aliveNow  = Object.keys(s.alive).filter(p => s.alive[p]);
      const aliveSpies = aliveNow.filter(p => s.isUndercover[p]);
      const aliveCivs  = aliveNow.filter(p => !s.isUndercover[p]);

      if (aliveSpies.length === 0)                     { s.status = 'ended'; s.winner = 'civilians'; }
      else if (aliveSpies.length >= aliveCivs.length)  { s.status = 'ended'; s.winner = 'undercover'; }
      else                                              { s.status = 'result'; }
    }
  }

  writeState(found.row, s);
  return {success: true, status: s.status, lastResult: s.lastResult, winner: s.winner};
}

// ── 下一輪（房主觸發）────────────────────────────────────────
function nextRound(roomName, playerName) {
  const found = readState(roomName);
  if (!found) return {success: false, error: '房間不存在'};
  const s = found.state;

  if (s.host !== playerName)  return {success: false, error: '只有房主可以推進'};
  if (s.status !== 'result')  return {success: false, error: '狀態錯誤'};

  const alivePlayers = Object.keys(s.alive).filter(p => s.alive[p]);
  s.round++;
  s.status = 'describing';
  s.votes = {};
  s.lastResult = null;
  s.descOrder = alivePlayers.sort(() => Math.random() - 0.5);
  s.currentDescIdx = 0;
  s.currentDescriber = s.descOrder[0];

  writeState(found.row, s);
  return {success: true, round: s.round};
}
