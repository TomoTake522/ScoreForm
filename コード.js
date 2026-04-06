function getActiveSheetName() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = ss.getSheetByName("index");
  return indexSheet.getRange("B1").getValue().toString().trim();
}

function doGet(e) {
  const page = e.parameter.page || 'index';
  const template = HtmlService.createTemplateFromFile('index');
  template.page = page; 
  return template.evaluate()
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setTitle('スコア入力システム');
}

function getQuizData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("問題");
  const data = sheet.getDataRange().getValues();
  
  const result = [];
  let currentSection = ""; 

  for (let i = 1; i < data.length; i++) {
    const mark = data[i][0];
    const content = data[i][1];
    const valC = data[i][2];
    const valD = data[i][3];

    if (mark === "*") {
      currentSection = content;
    } 
    else if (mark === "-") {
      result.push({
        type: "input",
        question: content,
        maxScore: valC !== "" ? valC : 0,
        minScore: valD !== "" ? valD : 0,
        section: currentSection
      });
    } 
    else if (content !== "") {
  result.push({
    type: "selection",
    question: content,
    score: Number(valC) || 0,
    penalty: (valD !== null && valD !== "") ? Number(valD) : 0, 
    section: currentSection
  });
}
  }

  console.log(result);
  return result;
}

function getInitialData(page) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const activeSheetName = getActiveSheetName();
    const indexSheet = ss.getSheetByName("index");
    
    const statusValues = indexSheet.getRange("A2:B20").getValues();
    const currentStatusRow = statusValues.find(row => row[0].toString().trim() === activeSheetName);
    const isLocked = !!(currentStatusRow && currentStatusRow[1].toString().includes("締切"));

    const sheet = ss.getSheetByName(activeSheetName);
    if (!sheet) throw new Error("シートが見つかりません: " + activeSheetName);
    
    const data = sheet.getDataRange().getValues();
    const rows = data.slice(1);
    const teams = rows.map(r => r[0]).filter(String);

    const allTeamData = {};
    rows.forEach(r => {
      if (r[0]) {
        allTeamData[r[0]] = {
          staffName: r[1],
          count: r[2],
          teamName: r[3],
          ruby: r[4],
          scores: r.slice(6)
        };
      }
    });

    const quizDefinitions = getQuizData();

    const rankingData = rows.filter(r => r[0] !== "" && r[5] !== "")
                            .map(r => ({num: r[0], name: r[3], score: Number(r[5])||0}))
                            .sort((a,b) => b.score - a.score);

    let adminData = null;
    if (page === 'admin') {
      const incompleteTeams = rows.filter(r => r[0] !== "" && (!r[5] || r[5] === 0))
                                  .map(r => ({
                                    num: r[0],
                                    staff: r[1] || "未設定",
                                    name: r[3] || "未設定"
                                  }));
      adminData = { incompleteTeams: incompleteTeams };
    }
    
    return { 
      currentSheet: activeSheetName, 
      isLocked: isLocked, 
      teams: teams, 
      allTeamData: allTeamData,
      rankingData: rankingData,
      adminData: adminData,
      quizDefinitions: quizDefinitions
    };
  } catch (e) { 
    return { error: e.message }; 
  }
}

function getRankings() {
  try {
    const activeSheetName = getActiveSheetName();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getActiveSheetName());
    const data = sheet.getDataRange().getValues().slice(1);
    const getSorted = (idx) => data.filter(r => r[0] !== "").map(r => ({num: r[0], name: r[3] || "未設定", score: Number(r[idx]) || 0})).sort((a,b) => b.score - a.score);
    return { overall: getSorted(5) };
  } catch (e) { return null; }
}

function getTeamData(teamNum) {
  const activeSheetName = getActiveSheetName();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(activeSheetName);
  const data = sheet.getDataRange().getValues();
  const r = data.find(row => row[0] == teamNum);
  if (!r) return null;
  
  return { 
    staffName: r[1],
    count: r[2], 
    teamName: r[3], 
    ruby: r[4], 
    scores: r.slice(6) 
  };
}

function updateData(teamNum, payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getActiveSheetName());
  const data = sheet.getRange("A:A").getValues();
  let idx = data.findIndex(r => r[0] == teamNum) + 1;
  if (idx > 0) {
    sheet.getRange(idx, 3, 1, 3).setValues([[payload.count, payload.teamName, payload.ruby]]);
    sheet.getRange(idx, 6).setValue(payload.totalScore);
    sheet.getRange(idx, 7, 1, payload.scores.length).setValues([payload.scores]);
    return "保存完了";
  }
}

function updateTeamCount(teamNum, count) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(getActiveSheetName());
  const data = sheet.getRange("A:A").getValues();
  let idx = data.findIndex(r => r[0] == teamNum) + 1;
  if (idx > 0) { sheet.getRange(idx, 3).setValue(count); return "保存完了"; }
}

function doPost(e) {
  const p = e.parameter;
  const command = p.command;
  let text = p.text ? p.text.trim() : "";
  const scriptUrl = ScriptApp.getService().getUrl();
  
  try {
    if (command === "/得点") return createSlackResponse(`🔗 *スコア入力フォームはこちら*\n${scriptUrl}`, "ephemeral");
    if (command === "/発表") {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const currentDefault = getActiveSheetName();
      let targetSheetName = text ? text : currentDefault;

      let sheet = ss.getSheetByName(targetSheetName);
      if (!sheet) return createSlackResponse(`エラー：シート「${targetSheetName}」が見つかりません。`, "ephemeral");

      const data = sheet.getDataRange().getValues().slice(1);
      const ranking = data.filter(r => r[0] !== "" && r[5] !== "").map(r => ({
        num: r[0], count: r[2] || 0,
        name: (r[3] || "未設定").toString().trim(),
        ruby: (r[4] || "").toString().trim(),
        score: Number(r[5]) || 0
      })).sort((a, b) => b.score - a.score);

      let msg = `【${targetSheetName}】の結果発表です！！！\n`;
      const icons = ["1位：", "2位：", "3位："];
      const emoji = ["🥇", "🥈", "🥉"];
      let found = false;

      for (let i = 0; i < 3; i++) {
        if (ranking[i]) {
          let displayName = ranking[i].name;
          if (ranking[i].ruby && ranking[i].name !== ranking[i].ruby) displayName += `（${ranking[i].ruby}）`;
          msg += `${emoji[i]} ${icons[i]}${ranking[i].num}班 ${displayName} ${ranking[i].count}名 ${ranking[i].score}点\n`;
          found = true;
        }
      }
      return createSlackResponse(found ? msg + "おめでとうございます🎊" : `【${targetSheetName}】まだスコアデータがありません。`, "in_channel");
    }
  } catch (err) { return createSlackResponse("エラー: " + err.message, "ephemeral"); }
}

function createSlackResponse(msg, type) {
  return ContentService.createTextOutput(JSON.stringify({"response_type": type, "text": msg})).setMimeType(ContentService.MimeType.JSON);
}