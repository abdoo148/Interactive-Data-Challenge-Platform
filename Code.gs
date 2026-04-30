// 1. Create Trainer Menu on Open
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🏆 Data Challenge')
      .addItem('Show Leaderboard (Trainer)', 'showLeaderboardSidebar')
      .addSeparator()
      .addItem('Setup Database & Settings', 'setupDatabase')
      .addToUi();
}

// 2. Show Leaderboard Sidebar
function showLeaderboardSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Live Trainer Dashboard')
      .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

// 3. Main Trainee Web App UI
function doGet(e) {
  try {
    var template = HtmlService.createTemplateFromFile('index');
    template.role = (e && e.parameter && e.parameter.role) ? e.parameter.role : '';

    var htmlOutput = template.evaluate();
    htmlOutput.setTitle('Interactive Data Challenge');
    htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
    return htmlOutput;
  } catch (err) {
    return HtmlService.createHtmlOutput('<div style="font-family:sans-serif; text-align:center; padding:50px;"><h1>⚠️ Error loading platform</h1><p>' + err.message + '</p></div>');
  }
}

// 4. Setup Database and Settings
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // A. Settings Sheet
  var settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Settings');
    var defaultSettings = [
      ['Setting Name', 'Value', 'Notes (Do not modify setting names)'],
      ['Theme Color', '#0ea5e9', 'Main color for buttons/glow (e.g., #0ea5e9 Blue)'],
      ['Background Color', '#0f172a', 'Main background color (e.g., #0f172a Dark Navy)'],
      ['Text Color', '#f8fafc', 'General text color (e.g., #f8fafc White)'],
      ['Option BG Color', '#1e293b', 'Background color for answers'],
      ['Option Hover Color', '#312e81', 'Color on mouse hover'],
      ['Pre-Quiz Questions (JSON)', '[{"id":"q1","label":"Job Title?","type":"text","required":true}]', 'Questions appearing before the challenge'],
      ['Rank 1 (Experts)', 'Data Maestro', 'Highest Rank Name'],
      ['Rank 1 %', '0.85', 'Requires 85% or more'],
      ['Rank 2 (Pros)', 'Table Ninja', 'Second Rank Name'],
      ['Rank 2 %', '0.65', 'Requires 65% or more'],
      ['Rank 3 (Intermediates)', 'Number Hunter', 'Third Rank Name'],
      ['Rank 3 %', '0.40', 'Requires 40% or more'],
      ['Rank 4 (Beginners)', 'Needs Coffee', 'Lowest Rank Name']
    ];
    settingsSheet.getRange(1, 1, 14, 3).setValues(defaultSettings);
    settingsSheet.getRange('A1:C1').setFontWeight('bold').setBackground('#1e293b').setFontColor('white');
    settingsSheet.setColumnWidth(1, 200); settingsSheet.setColumnWidth(2, 200); settingsSheet.setColumnWidth(3, 400);
  }

  // B. Questions Sheet
  var questionsSheet = ss.getSheetByName('Questions');
  if (!questionsSheet) {
    questionsSheet = ss.insertSheet('Questions');
    var qHeaders = ['Type', 'Question Text', 'Option 1', 'Option 2', 'Option 3', 'Option 4', 'Correct Answer/Order', 'Time (Seconds)', 'Points'];
    questionsSheet.appendRow(qHeaders);
    questionsSheet.getRange('A1:I1').setFontWeight('bold').setBackground('#1e293b').setFontColor('white');

    var sampleQuestions = [
      ["single", "In yesterday's session, 80% of enterprise data was classified as:", "Structured Data", "Unstructured Data", "Semi-structured Data", "Aggregated Data", "2", 15, 100],
      ["multiple", "Which of the following are dimensions of Data Quality? (Select all that apply)", "Completeness", "Speed", "Accuracy", "Price", "1,3", 20, 150],
      ["true_false", "Merging Cells is an excellent practice for organizing databases.", "True", "False", "", "", "2", 10, 50],
      ["order", "Order the following data lifecycle stages correctly:", "Analysis & Extraction", "Building Dashboards", "Raw Data Collection", "Data Cleaning", "3,4,1,2", 30, 250]
    ];
    questionsSheet.getRange(2, 1, 4, 9).setValues(sampleQuestions);
    questionsSheet.getRange("J1").setValue("Note: Types are (single, multiple, true_false, order). For multiple & order, use commas (e.g. 1,3,2).");
    questionsSheet.getRange("J1").setFontColor("red");
  }

  // C. Results Sheet
  var resultsSheet = ss.getSheetByName('Results');
  if (!resultsSheet) {
    resultsSheet = ss.insertSheet('Results');
    var rHeaders = ['Timestamp', 'Trainee Name', 'Pre-Quiz Answers', 'Points Earned', 'Max Score', 'Rank'];
    resultsSheet.appendRow(rHeaders);
    resultsSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#1e293b').setFontColor('white');
  }

  SpreadsheetApp.getUi().alert('Database and Settings successfully initialized!');
}

// 5. Fetch API Data
function getAppConfig() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("File is not properly connected to the script.");

    if (!ss.getSheetByName('Settings') || !ss.getSheetByName('Questions')) { 
      setupDatabase(); 
    }

    var settingsSheet = ss.getSheetByName('Settings');
    if (!settingsSheet) throw new Error("Settings sheet is missing or deleted.");
    
    var settingsData = settingsSheet.getRange(2, 1, 13, 2).getValues();

    var preQuizData = [];
    try { 
      if (settingsData[5][1]) {
        preQuizData = JSON.parse(settingsData[5][1]); 
      }
    } catch(e) { preQuizData = []; }

    var r1Min = parseFloat(settingsData[7][1]);
    var r2Min = parseFloat(settingsData[9][1]);
    var r3Min = parseFloat(settingsData[11][1]);

    var config = {
      themeColor: settingsData[0][1] || '#0ea5e9',
      bgMain: settingsData[1][1] || '#0f172a',
      textMain: settingsData[2][1] || '#f8fafc',
      optionBg: settingsData[3][1] || '#1e293b',
      optionHover: settingsData[4][1] || '#312e81',
      preQuizQuestions: preQuizData,
      ranks: [
        { title: settingsData[6][1] || 'Rank 1', min: isNaN(r1Min) ? 0.85 : r1Min, level: 1 },
        { title: settingsData[8][1] || 'Rank 2', min: isNaN(r2Min) ? 0.65 : r2Min, level: 2 },
        { title: settingsData[10][1] || 'Rank 3', min: isNaN(r3Min) ? 0.40 : r3Min, level: 3 },
        { title: settingsData[12][1] || 'Rank 4', min: 0, level: 4 }
      ]
    };

    var questionsSheet = ss.getSheetByName('Questions');
    if (!questionsSheet) throw new Error("Questions sheet is missing or deleted.");

    var qData = questionsSheet.getDataRange().getValues();
    var questions = [];
    var totalMaxScore = 0;

    var isValidOption = function(opt) {
      return opt !== undefined && opt !== null && String(opt).trim() !== "";
    };

    for (var i = 1; i < qData.length; i++) {
      if (qData[i][0] && String(qData[i][1]).trim() !== "") {
        var type = qData[i][0].toString().trim().toLowerCase();
        var pts = parseInt(qData[i][8]) || 100;
        totalMaxScore += pts;

        var correctStr = (qData[i][6] !== undefined && qData[i][6] !== null && qData[i][6] !== "") ? qData[i][6].toString() : "1";
        var correctArr = correctStr.split(',').map(function(s) { return parseInt(s.trim()); }).filter(function(n) { return !isNaN(n); });
        
        var qObj = {
          id: i, 
          type: type, 
          question: qData[i][1].toString(), 
          options: [],
          correctAnswers: type !== 'order' ? correctArr : [],
          correctOrder: type === 'order' ? correctArr : [],
          time: parseInt(qData[i][7]) || 20, 
          points: pts
        };

        if(type === 'true_false') {
          qObj.options.push({ text: isValidOption(qData[i][2]) ? String(qData[i][2]) : "True", id: 1 });
          qObj.options.push({ text: isValidOption(qData[i][3]) ? String(qData[i][3]) : "False", id: 2 });
        } else {
          if(isValidOption(qData[i][2])) qObj.options.push({ text: String(qData[i][2]), id: 1 });
          if(isValidOption(qData[i][3])) qObj.options.push({ text: String(qData[i][3]), id: 2 });
          if(isValidOption(qData[i][4])) qObj.options.push({ text: String(qData[i][4]), id: 3 });
          if(isValidOption(qData[i][5])) qObj.options.push({ text: String(qData[i][5]), id: 4 });
        }
        questions.push(qObj);
      }
    }

    config.totalMaxScore = totalMaxScore;
    return JSON.stringify({ config: config, questions: questions });
  } catch (error) {
    throw new Error("Failed to process app config: " + error.message);
  }
}

// 6. Submit Participant Result (Concurrency Safe)
function submitParticipantResult(dataStr) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 

    var data = JSON.parse(dataStr);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');
    
    if (!sheet) throw new Error("Results sheet is not available.");

    var preQuizAnswersStr = JSON.stringify(data.preQuiz || {});

    sheet.appendRow([
      new Date(), 
      data.name || '-', 
      preQuizAnswersStr, 
      parseInt(data.score) || 0, 
      parseInt(data.maxScore) || 0, 
      data.rank || '-'
    ]);

    return "success";
  } catch(e) { 
    throw new Error(e.message); 
  } finally {
    lock.releaseLock();
  }
}

// 7. Get Leaderboard Data for Sidebar
function getLeaderboard() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Results');
    if(!sheet) return JSON.stringify({ data: [] });
    
    var data = sheet.getDataRange().getValues();
    var leaderboard = [];

    for(var i = 1; i < data.length; i++) {
      if (data[i][1]) {
        leaderboard.push({
          name: String(data[i][1]), 
          score: parseInt(data[i][3]) || 0, 
          max: parseInt(data[i][4]) || 0, 
          rank: String(data[i][5]) || '-'
        });
      }
    }
    leaderboard.sort(function(a, b) { return b.score - a.score; });
    return JSON.stringify({ data: leaderboard }); 
  } catch (e) {
    return JSON.stringify({ data: [], error: e.message }); 
  }
}
