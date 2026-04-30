// 1. إنشاء قائمة المدرب داخل جوجل شيت عند فتح الملف
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🏆 منصة تحدي البيانات')
      .addItem('عرض لوحة الصدارة (المدرب)', 'showLeaderboardSidebar')
      .addSeparator()
      .addItem('تهيئة الجداول والإعدادات', 'setupDatabase')
      .addToUi();
}

// 2. دالة عرض لوحة الصدارة في قائمة جانبية داخل الشيت
function showLeaderboardSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('لوحة تحكم المدرب المباشرة')
      .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

// 3. دالة عرض الواجهة الأساسية للمتدربين (Web App) مع معالجة الأخطاء
function doGet(e) {
  try {
    var template = HtmlService.createTemplateFromFile('index');
    template.role = (e && e.parameter && e.parameter.role) ? e.parameter.role : '';

    var htmlOutput = template.evaluate();
    htmlOutput.setTitle('تحدي البيانات التفاعلي');
    htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no');
    return htmlOutput;
  } catch (err) {
    return HtmlService.createHtmlOutput('<div dir="rtl" style="font-family:sans-serif; text-align:center; padding:50px;"><h1>⚠️ حدث خطأ في تحميل المنصة</h1><p>' + err.message + '</p></div>');
  }
}

// 4. دالة تهيئة الجداول والإعدادات الشاملة (تُنشئ الشيتات إذا لم تكن موجودة)
function setupDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // أ. نافذة الإعدادات
  var settingsSheet = ss.getSheetByName('الإعدادات');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('الإعدادات');
    var defaultSettings = [
      ['اسم الإعداد', 'القيمة', 'ملاحظات (لا تقم بتعديل أسماء الإعدادات)'],
      ['اللون الرئيسي (Theme Color)', '#0ea5e9', 'لون الأزرار والتوهج (مثل: #0ea5e9 أزرق)'],
      ['لون الخلفية (Background)', '#0f172a', 'لون خلفية الشاشة (مثل: #0f172a كحلي غامق)'],
      ['لون النص (Text Color)', '#f8fafc', 'لون النصوص العامة (مثل: #f8fafc أبيض)'],
      ['خلفية أزرار الخيارات (Option BG)', '#1e293b', 'لون خلفية الإجابات'],
      ['تأثير التحويم للخيارات (Option Hover)', '#312e81', 'اللون عند مرور الماوس على الإجابة'],
      ['أسئلة التمهيد (JSON Format)', '[{"id":"q1","label":"المسمى الوظيفي؟","type":"text","required":true}]', 'مصفوفة أسئلة تظهر قبل التحدي'],
      ['رتبة 1 (الخبراء)', 'مايسترو البيانات', 'اسم أعلى رتبة'],
      ['نسبة رتبة 1', '0.85', 'يحصل عليها من يحقق 85% فأكثر'],
      ['رتبة 2 (المحترفون)', 'نينجا الجداول', 'اسم الرتبة الثانية'],
      ['نسبة رتبة 2', '0.65', 'يحصل عليها من يحقق 65% فأكثر'],
      ['رتبة 3 (المتوسطون)', 'صائد الأرقام', 'اسم الرتبة الثالثة'],
      ['نسبة رتبة 3', '0.40', 'يحصل عليها من يحقق 40% فأكثر'],
      ['رتبة 4 (المبتدئون)', 'محتاج قهوة ومراجعة', 'أقل رتبة لمن لم يحقق النسب السابقة']
    ];
    // ضبط النطاق ليطابق 14 صفاً بشكل دقيق
    settingsSheet.getRange(1, 1, 14, 3).setValues(defaultSettings);
    settingsSheet.getRange('A1:C1').setFontWeight('bold').setBackground('#1e293b').setFontColor('white');
    settingsSheet.setColumnWidth(1, 200); settingsSheet.setColumnWidth(2, 200); settingsSheet.setColumnWidth(3, 400);
  }

  // ب. نافذة الأسئلة
  var questionsSheet = ss.getSheetByName('الأسئلة');
  if (!questionsSheet) {
    questionsSheet = ss.insertSheet('الأسئلة');
    var qHeaders = ['نوع السؤال', 'نص السؤال', 'الخيار 1', 'الخيار 2', 'الخيار 3', 'الخيار 4', 'الإجابة/الترتيب الصحيح', 'الوقت (ثواني)', 'النقاط'];
    questionsSheet.appendRow(qHeaders);
    questionsSheet.getRange('A1:I1').setFontWeight('bold').setBackground('#1e293b').setFontColor('white');

    var sampleQuestions = [
      ["single", "في جلسة أمس، ذكرنا أن 80% من بيانات المؤسسات تُصنف كـ:", "بيانات منظمة", "بيانات غير مهيكلة", "بيانات شبه مهيكلة", "بيانات مجمعة", "2", 15, 100],
      ["multiple", "من أبعاد جودة البيانات (اختر كل ما يسبق):", "الاكتمال", "السرعة", "الدقة", "السعر", "1,3", 20, 150],
      ["true_false", "دمج الخلايا (Merge Cells) ممارسة ممتازة لتنظيم قواعد البيانات.", "صح", "خطأ", "", "", "2", 10, 50],
      ["order", "رتب المراحل التالية لدورة حياة البيانات بشكل صحيح:", "التحليل والاستخراج", "بناء لوحات القيادة", "جمع البيانات الخام", "تنظيف البيانات", "3,4,1,2", 30, 250]
    ];
    questionsSheet.getRange(2, 1, 4, 9).setValues(sampleQuestions);

    // إضافة ملاحظات للمدرب
    questionsSheet.getRange("J1").setValue("ملاحظات: الأنواع هي (single, multiple, true_false, order). المتعدد والترتيب يُكتب بفاصلة (1,3,2).");
    questionsSheet.getRange("J1").setFontColor("red");
  }

  // ج. نافذة النتائج
  var resultsSheet = ss.getSheetByName('النتائج');
  if (!resultsSheet) {
    resultsSheet = ss.insertSheet('النتائج');
    var rHeaders = ['التاريخ والوقت', 'اسم المتدرب', 'إجابات التمهيد', 'النقاط المكتسبة', 'أقصى درجة', 'الرتبة'];
    resultsSheet.appendRow(rHeaders);
    resultsSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#1e293b').setFontColor('white');
  }

  SpreadsheetApp.getUi().alert('تم تهيئة الجداول والإعدادات بنجاح!');
}

// 5. دالة جلب البيانات لتشغيل واجهة المتدرب (API)
function getAppConfig() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) throw new Error("الملف غير متصل بشكل صحيح بالسكريبت.");

    if (!ss.getSheetByName('الإعدادات') || !ss.getSheetByName('الأسئلة')) { 
      setupDatabase(); 
    }

    var settingsSheet = ss.getSheetByName('الإعدادات');
    if (!settingsSheet) throw new Error("شيت الإعدادات غير موجود أو محذوف.");
    
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
        { title: settingsData[6][1] || 'رتبة 1', min: isNaN(r1Min) ? 0.85 : r1Min, level: 1 },
        { title: settingsData[8][1] || 'رتبة 2', min: isNaN(r2Min) ? 0.65 : r2Min, level: 2 },
        { title: settingsData[10][1] || 'رتبة 3', min: isNaN(r3Min) ? 0.40 : r3Min, level: 3 },
        { title: settingsData[12][1] || 'رتبة 4', min: 0, level: 4 }
      ]
    };

    var questionsSheet = ss.getSheetByName('الأسئلة');
    if (!questionsSheet) throw new Error("شيت الأسئلة غير موجود أو محذوف.");

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
          qObj.options.push({ text: isValidOption(qData[i][2]) ? String(qData[i][2]) : "صح", id: 1 });
          qObj.options.push({ text: isValidOption(qData[i][3]) ? String(qData[i][3]) : "خطأ", id: 2 });
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
    throw new Error("فشل في معالجة إعدادات التطبيق: " + error.message);
  }
}

// 6. تسجيل نتيجة المتدرب بطريقة آمنة للتزامن (Concurrency Safe)
function submitParticipantResult(dataStr) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); 

    var data = JSON.parse(dataStr);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('النتائج');
    
    if (!sheet) throw new Error("شيت النتائج غير متوفر لتسجيل الإجابات.");

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

// 7. جلب لوحة الصدارة للقائمة الجانبية للمدرب بشكل آمن
function getLeaderboard() {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('النتائج');
    if(!sheet) return JSON.stringify({ data: [] });
    
    var data = sheet.getDataRange().getValues();
    var leaderboard = [];

    for(var i = 1; i < data.length; i++) {
      if (data[i][1]) { // حماية من الصفوف الفارغة بالكامل
        leaderboard.push({
          name: String(data[i][1]), 
          score: parseInt(data[i][3]) || 0, 
          max: parseInt(data[i][4]) || 0, 
          rank: String(data[i][5]) || '-'
        });
      }
    }
    leaderboard.sort(function(a, b) { return b.score - a.score; });
    
    // تم التعديل هنا لترجع البيانات داخل أوبجكت اسمه data لتتطابق مع كود القائمة الجانبية
    return JSON.stringify({ data: leaderboard }); 
  } catch (e) {
    return JSON.stringify({ data: [], error: e.message }); 
  }
}
