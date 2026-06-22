const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
        Header, Footer, PageNumber } = require('docx');
const fs = require('fs');

// ── קריאת נתוני המשחק מה-HTML (מקור האמת) ─────────────────
const htmlContent = fs.readFileSync('moadilim-poc.html', 'utf8');
const gcStart = htmlContent.indexOf('const GC = ') + 'const GC = '.length;
let depth = 0, inStr = false, escaped = false, idx = gcStart;
while (idx < htmlContent.length) {
  const c = htmlContent[idx];
  if (escaped) { escaped = false; }
  else if (c === '\\' && inStr) { escaped = true; }
  else if (c === '"' && !inStr) { inStr = true; }
  else if (c === '"' && inStr) { inStr = false; }
  else if (!inStr && c === '{') { depth++; }
  else if (!inStr && c === '}') { depth--; if (depth === 0) break; }
  idx++;
}
const GC = JSON.parse(htmlContent.slice(gcStart, idx + 1));

const NAVY  = '221f64';
const TEAL  = '1abe9b';
const LGRAY = 'f5f6fa';
const MGRAY = 'e3e6ef';
const WHITE = 'ffffff';

// A4: 11906 DXA wide, margins 1134 DXA (2 cm) each → content 9638
const TW = 9638;
const C1 = 1800; // מפתח בקוד
const C2 = 2600; // טקסט נוכחי
const C3 = 2400; // היכן מופיע
const C4 = 2838; // תיקון / הערה

const b = { style: BorderStyle.SINGLE, size: 1, color: MGRAY };
const borders = { top: b, bottom: b, left: b, right: b };
const cellMargins = { top: 90, bottom: 90, left: 160, right: 160 };

function hCell(text, width) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: { fill: NAVY, type: ShadingType.CLEAR },
    margins: cellMargins,
    children: [new Paragraph({
      alignment: AlignmentType.RIGHT, bidirectional: true,
      children: [new TextRun({ text, font: 'Arial', size: 21, bold: true, color: WHITE })]
    })]
  });
}

function dCell(text, width, shade, color) {
  return new TableCell({
    borders, width: { size: width, type: WidthType.DXA },
    shading: shade ? { fill: LGRAY, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    children: [new Paragraph({
      alignment: AlignmentType.RIGHT, bidirectional: true,
      children: [new TextRun({ text: text || '', font: 'Arial', size: 20, color: color || '2a2a2a' })]
    })]
  });
}

// rows: [key, currentText, whereItAppears]
function makeTable(rows) {
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: [C1, C2, C3, C4],
    rows: [
      new TableRow({
        tableHeader: true,
        children: [
          hCell('מפתח בקוד', C1),
          hCell('טקסט נוכחי', C2),
          hCell('היכן מופיע', C3),
          hCell('תיקון / הערה', C4),
        ]
      }),
      ...rows.map(([key, val, where], i) => new TableRow({
        children: [
          dCell(key,   C1, i%2),
          dCell(val,   C2, i%2),
          dCell(where, C3, i%2, '555577'),
          dCell('',    C4, i%2),
        ]
      }))
    ]
  });
}

const h1 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_1, bidirectional: true, alignment: AlignmentType.RIGHT,
  children: [new TextRun({ text, font: 'Arial', size: 34, bold: true, color: NAVY })]
});

const h2 = (text) => new Paragraph({
  heading: HeadingLevel.HEADING_2, bidirectional: true, alignment: AlignmentType.RIGHT,
  children: [new TextRun({ text, font: 'Arial', size: 24, bold: true, color: '555555' })]
});

const note = (text) => new Paragraph({
  bidirectional: true, alignment: AlignmentType.RIGHT,
  spacing: { after: 80 },
  children: [new TextRun({ text, font: 'Arial', size: 18, color: '888888', italics: true })]
});

const gap = () => new Paragraph({ spacing: { after: 180 }, children: [new TextRun('')] });

// ════════════════════════════════════════════════════════
const ch = [];

// ── כותרת ──────────────────────────────────────────────
ch.push(
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 480, after: 200 },
    children: [new TextRun({ text: 'מסמך קופי — מסילות ותכנונים', font: 'Arial', size: 52, bold: true, color: NAVY })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 100 },
    children: [new TextRun({ text: 'מיפוי כולל של כל הניסוחים בממשק המשחק', font: 'Arial', size: 26, color: '666666' })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 400 },
    children: [new TextRun({ text: 'יוני 2026', font: 'Arial', size: 22, color: '999999' })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 80 },
    children: [new TextRun({ text: 'איך משתמשים במסמך:', font: 'Arial', size: 22, bold: true, color: NAVY })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 60 },
    children: [new TextRun({ text: 'כל טבלה = קטגוריה אחת מהממשק. ארבע עמודות:', font: 'Arial', size: 20, color: '444444' })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 40 },
    children: [new TextRun({ text: 'מפתח בקוד — המיקום המדויק ב-gameContent.json או ב-HTML (לאיתור מהיר בעת עדכון)', font: 'Arial', size: 20, color: '444444' })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 40 },
    children: [new TextRun({ text: 'טקסט נוכחי — הניסוח הקיים במערכת', font: 'Arial', size: 20, color: '444444' })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 40 },
    children: [new TextRun({ text: 'היכן מופיע — איפה בממשק הטקסט נראה', font: 'Arial', size: 20, color: '444444' })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 400 },
    children: [new TextRun({ text: 'תיקון / הערה — מלאי כאן את השינויים הרצויים', font: 'Arial', size: 20, bold: true, color: TEAL })] }),
);

// ── 1. מטא ─────────────────────────────────────────────
ch.push(h1('1. כותרות ומטא'), note('מקור: gameContent.json → uiCopy'));
ch.push(makeTable([
  ['uiCopy.appTitle',     'מסילות ותכנונים',      'מסך פתיחה — כותרת ראשית גדולה'],
  ['uiCopy.pageTitle',    'מסילות ותכנונים - Proof of Concept', 'כותרת ה-tab בדפדפן'],
  ['uiCopy.gameSubtitle', 'התוכנית לא מתקדמת כי הזמן עובר - היא מתקדמת כי הנאמן מוביל אותה נכון.', 'מסך פתיחה — כיתוב אפור מתחת לכותרת'],
]));
ch.push(gap());

// ── 2. מסך פתיחה ────────────────────────────────────────
ch.push(h1('2. מסך פתיחה'), note('מקור: gameContent.json → uiCopy.welcome'));
ch.push(makeTable([
  ['uiCopy.welcome.title',                  'ברוכים הבאים למסילות ותכנונים!',   'מסך פתיחה — כותרת כרטיס מרכזי'],
  ['uiCopy.welcome.intro',                  'במשחק הזה תובילו תכנית עבודה חטיבתית לאורך השנה: מהגדרת המשימות, דרך דיוק המדדים ואבני הדרך, ועד לאישור תכנית העבודה ומודל התגמול.', 'מסך פתיחה — פסקת הסבר'],
  ['uiCopy.welcome.startButton',            '(הוסר — הוחלף בכפתור X)',             'מסך פתיחה — כפתור הכניסה ללוח הוחלף ב-X ללא טקסט'],
  ['uiCopy.welcome.openInstructionsButton', 'הנחיות המשחק',                       'מסך פתיחה — כפתור משני (קו תחתון)'],
  ['uiCopy.welcome.closeInstructionsButton','חזרה למפה',                           'חלונית הנחיות — כפתור סגירה'],
  ['hardcoded: station_start activatePhones', 'מתחילים',                           'תחנת כולם לעלות — כפתור התחלת המשחק (נועל טלפוני קבוצות לפי תור)'],
  ['hardcoded: station_start phonesActive',   '✓ המשחק החל — הקבוצות ננעלות לפי תור', 'תחנת כולם לעלות — הודעת אישור לאחר לחיצה על מתחילים'],
]));
ch.push(gap());

// ── 3. הוראות המשחק ─────────────────────────────────────
ch.push(h1('3. הוראות המשחק'), note('מקור: gameContent.json → uiCopy.welcome.sections[]'));
ch.push(makeTable([
  ['welcome.sections[0].title', 'איך משחקים?',        'חלונית הנחיות — כותרת פסקה 1'],
  ['welcome.sections[0].body',  'כל קבוצה מתקדמת על גבי מסלול המשחק. בתחנות המקצועיות הקבוצה בודקת את תיק תכנית העבודה שלה במובייל, מאתרת אי-דיוקים ומשיבה בעל פה למנחה.', 'חלונית הנחיות — גוף פסקה 1'],
  ['welcome.sections[1].title', 'תפקיד המנחה',        'חלונית הנחיות — כותרת פסקה 2'],
  ['welcome.sections[1].body',  'המנחה מפעילה את התחנות דרך המסך המוקרן, מציגה שאלות, מפעילה טיימר, חושפת תשובות ומעדכנת את מצב המשחק לפי איכות המענה.', 'חלונית הנחיות — גוף פסקה 2'],
  ['welcome.sections[2].title', 'מה המשתתפים עושים?', 'חלונית הנחיות — כותרת פסקה 3'],
  ['welcome.sections[2].body',  'כל קבוצה פותחת את תיק תכנית העבודה החטיבתי שלה, מתייעצת בזמן מוגבל ומנסה לזהות מה דורש תיקון. שימו לב — התיק נעול בזמן שקבוצה אחרת עונה.', 'חלונית הנחיות — גוף פסקה 3'],
  ['welcome.sections[3].title', 'תשובות והתקדמות',    'חלונית הנחיות — כותרת פסקה 4'],
  ['welcome.sections[3].body',  'תשובה נכונה — 10 נקודות, הקבוצה מתקדמת. תשובה חלקית — 5 נקודות, הקבוצה מתקדמת. תשובה שגויה — 0 נקודות, הקבוצה נשארת במקום ומדלגת על הסבב הבא.', 'חלונית הנחיות — גוף פסקה 4'],
  ['welcome.sections[4].title', 'בלת"מים',             'חלונית הנחיות — כותרת פסקה 5'],
  ['welcome.sections[4].body',  'בתחנות בלת"מ מסובבים גלגל ופועלים לפי ההוראה. בלת"מ יכול לגרום לעיכוב של סבב אחד או שניים — הקבוצה לא משחקת בסבבים אלה. קבוצה לא מסובבת שוב את אותו גלגל באותה תחנה.', 'חלונית הנחיות — גוף פסקה 5'],
  ['welcome.sections[5].title', 'איך מנצחים?',         'חלונית הנחיות — כותרת פסקה 6'],
  ['welcome.sections[5].body',  "המטרה היא להגיע הכי רחוק על המסלול. שוויון במיקום — נקודות שוברות שוויון. בסיום עוברים להתחייבות אישית: מה כל משתתף לוקח לעבודה שלו ב-72 השעות הקרובות.", 'חלונית הנחיות — גוף פסקה 6'],
  ['welcome.sections[6].title', 'סדר התורות',          'חלונית הנחיות — כותרת פסקה 7'],
  ['welcome.sections[6].body',  "המשחק מתנהל בסבבים — קבוצה 1, קבוצה 2, קבוצה 3, קבוצה 4, וחוזר חלילה. בראש המסך תמיד מוצג תור מי עכשיו ובאיזו תחנה. בסיום כל סבב מוצג חיווי 'סבב הושלם'.", 'חלונית הנחיות — גוף פסקה 7'],
]));
ch.push(gap());

// ── 4. לוח מנחה — כותרת ─────────────────────────────────
ch.push(h1('4. לוח המנחה — כותרת ומצב משחק'), note('מקור: gameContent.json → uiCopy'));
ch.push(makeTable([
  ['uiCopy.turnIndicator',           'תור {group} — {division} — תחנת {station}', 'לוח מנחה — שורת מצב עליונה (מעל המסלול)'],
  ['uiCopy.map.instruction',         'לחצו על תחנה להפעלתה',                       'לוח מנחה — כיתוב קטן על המסלול'],
  ['uiCopy.statusBar.activeTurn',    'תור פעיל',                                    'לוח מנחה — תג כחול בשורת הסטטוס'],
  ['uiCopy.statusBar.roundComplete', 'סבב הושלם — אפשר להתחיל סבב חדש',           'לוח מנחה — הודעה בשורת הסטטוס אחרי סיום סבב'],
  ['uiCopy.statusBar.noStation',     '-',                                            'לוח מנחה — ערך ריק כשאין תחנה פעילה'],
  ['uiCopy.statusBar.autoAdvance',   'יתקדם אוטומטית',                              'לוח מנחה — תג על קבוצה שממתינה להתקדמות אוטומטית'],
  ['uiCopy.statusBar.position',      'מיקום',                                       'לוח מנחה — תווית עמודת מיקום בשורת סטטוס'],
  ['uiCopy.statusBar.errors',        'טעויות',                                      'לוח מנחה — תווית עמודת טעויות בשורת סטטוס'],
  ['uiCopy.statusBar.nextTurn',      'תור הבא',                                     'לוח מנחה — תווית עמודת "תור הבא" בשורת סטטוס'],
]));
ch.push(gap());

// ── 5. לוח תוצאות — כרטיסי קבוצות ──────────────────────
ch.push(h1("5. לוח תוצאות — כרטיסי קבוצות (תחתית לוח)"), note('מקור: hardcoded HTML (renderFacilitator → CHIP_DEF, bScoreHtml) + uiCopy.statusBar'));
ch.push(makeTable([
  ['hardcoded: CHIP_DEF.active',   'תור פעיל',               'כרטיס קבוצה (תחתית לוח) — תג סטטוס: הקבוצה ששולטת עכשיו'],
  ['hardcoded: CHIP_DEF.played',   '✓ שיחקה',                'כרטיס קבוצה — תג: הקבוצה סיימה את תורה בסבב זה'],
  ['hardcoded: CHIP_DEF.waiting',  '⏳ ממתינה',               'כרטיס קבוצה — תג: הקבוצה מחכה (בעקבות בלת"מ)'],
  ['hardcoded: CHIP_DEF.notyet',   'מחכה לתורה',             'כרטיס קבוצה — תג: הקבוצה עוד לא הגיעה לתורה בסבב זה'],
  ['uiCopy.statusBar.played',      'שיחקה',                   'כרטיס קבוצה — שורת סטטוס מתחת לתג'],
  ['uiCopy.statusBar.waiting',     'ממתינה סבב',              'כרטיס קבוצה — שורת סטטוס: קבוצה שממתינה סבב'],
  ['uiCopy.statusBar.notYet',      'מחכה לתורה',             'כרטיס קבוצה — שורת סטטוס: עוד לא שיחקה'],
  ['uiCopy.points.label',          'נקודות',                  'כרטיס קבוצה — תווית מעל מספר הנקודות (מלא)'],
  ['hardcoded: bScoreHtml — נק׳',  'נק׳',                     'כרטיס קבוצה — קיצור "נקודות" בכרטיס הקומפקטי'],
  ['hardcoded: bScoreHtml — 📍',   '📍',                      'כרטיס קבוצה — emoji לפני שם התחנה הנוכחית'],
  ['hardcoded: nameByOrder — UE',  'גלגל הבלת"מ',             'כרטיס קבוצה — שם התחנה כשהקבוצה נמצאת בתחנת בלת"מ'],
]));
ch.push(gap());

// ── 6. פאנל קבוצה ───────────────────────────────────────
ch.push(h1('6. פאנל קבוצה — תחנה פעילה'), note('מקור: gameContent.json → uiCopy.groupCard'));
ch.push(makeTable([
  ['uiCopy.groupCard.waitingMessage',    'קבוצה זו ממתינה סבב',              'חלון תחנה — הודעה אפורה כשהקבוצה ממתינה'],
  ['uiCopy.groupCard.skipButton',        'דלגו כדי להמשיך',                  'חלון תחנה — כפתור לדילוג על קבוצה ממתינה'],
  ['uiCopy.groupCard.stationLocked',     'תחנה זו לא פעילה בתור הנוכחי',    'חלון תחנה — הודעה כשלוחצים על תחנה לא בתור'],
  ['uiCopy.groupCard.clickToOpen',       '▶ לחצו',                            'לוח מנחה — קישור מעל תחנה ניתנת ללחיצה'],
  ['uiCopy.groupCard.approveTransition', 'אישור מעבר',                        'חלון תחנת מעבר — כפתור אישור וסגירה'],
  ['uiCopy.groupCard.answerRecorded',    'התשובה נרשמה.',                     'חלון תחנה — הודעה אחרי בחירת נכון/חלקי/שגוי'],
  ['uiCopy.groupCard.mustAnswerFirst',   'יש לסמן תשובה לפני חזרה',          'חלון תחנה — הודעת שגיאה אם לוחצים "חזרה" לפני שעונים'],
]));
ch.push(gap());

// ── 7. כפתורים ──────────────────────────────────────────
ch.push(h1('7. כפתורים'), note('מקור: gameContent.json → uiCopy.buttons, uiCopy.timer'));
ch.push(makeTable([
  ['uiCopy.buttons.startGame',          'התחילו משחק',                  'מסך פתיחה — כפתור ראשי (כחול גדול)'],
  ['uiCopy.buttons.startTimer',         'הפעילו טיימר',                 'חלון תחנה מקצועית — כפתור הפעלת טיימר'],
  ['uiCopy.buttons.correct',            'נכון - התקדמו',                    'חלון תחנה — כפתור ירוק (לאחר חשיפת תשובה)'],
  ['uiCopy.buttons.partial',            'חלקי - טעות + המתנה',             'חלון תחנה — כפתור כתום'],
  ['uiCopy.buttons.wrong',              'שגוי - טעות + חזרה תחנה',         'חלון תחנה — כפתור אדום'],
  ['uiCopy.buttons.revealAnswer',       'חשפו תשובה',                   'חלון תחנה — כפתור שחושף את התשובה הנכונה'],
  ['uiCopy.buttons.openMobileFolder',   'פתח תיק קבוצה',               'חלון תחנה — כפתור שפותח תיק מובייל'],
  ['uiCopy.buttons.needReminder',       'צריכים תזכורת?',               'תיק מובייל — כפתור שפותח מגירת הגדרות/תזכורת'],
  ['uiCopy.buttons.openFolder',         'פתחו את תכנית העבודה השנתית', 'תיק מובייל — כפתור קישור לתכנית העבודה'],
  ['uiCopy.buttons.spinWheel',          'סובבו גלגל',                   'חלון בלת"מ — כפתור הפעלת הגלגל'],
  ['uiCopy.buttons.backToMap',          'חזרה למסלול',                  'חלון תחנה — כפתור חזרה ללוח (בפינה)'],
  ['uiCopy.buttons.confirmSoftWarning', 'המשך בכל זאת',                 'חלונית אזהרה — כפתור אישור המשך'],
  ['uiCopy.buttons.cancel',             'ביטול',                        'חלונית אזהרה — כפתור ביטול'],
  ['uiCopy.timer.stop',                 'עצרו',                         'טיימר פעיל — כפתור עצירה'],
  ['uiCopy.timer.resume',               'המשך',                         'טיימר מושהה — כפתור המשך'],
]));
ch.push(gap());

// ── 8. הודעות מערכת ─────────────────────────────────────
ch.push(h1('8. הודעות מערכת'), note('מקור: gameContent.json → uiCopy.messages'));
ch.push(makeTable([
  ['uiCopy.messages.mobileIntro',           'לפניכם הצצה למשימות נבחרות מתוך תכנית העבודה החטיבתית. במהלך המשחק תעבדו עם המשימות המוצגות כאן, אך זכרו: במציאות תכנית העבודה רחבה יותר, ובלת"מים יכולים להגיע גם ממשימות ומגורמים שאינם מוצגים בתיק.', 'תיק מובייל — פסקת הסבר בראש הדף'],
  ['uiCopy.messages.softWarning',           'שימו לב: לפי מצב המשחק, קבוצה זו עדיין לא הגיעה לתחנה הזו. להמשיך בכל זאת?', 'חלונית אזהרה — גוף הודעה (כשפותחים תחנה לפני הזמן)'],
  ['uiCopy.messages.stationCompleted',      'תחנה זו כבר הושלמה. בתור הבא הקבוצה מתקדמת חזרה קדימה.', 'חלון תחנה — הודעה כשהתחנה כבר טופלה'],
  ['uiCopy.messages.unexpectedAlreadyUsed', 'קבוצה זו כבר הפעילה את הבלת"ם הזה. ממשיכים הלאה.', 'חלון בלת"מ — הודעה כשניסו להפעיל בלת"מ שכבר השתמשו בו'],
  ['uiCopy.messages.correctOutcome',        'מצוין! הקבוצה מתקדמת לתחנה הבאה.',                            'לוח מנחה — באנר תוצאה (ירוק) אחרי לחיצת "נכון"'],
  ['uiCopy.messages.partialOutcome',        'נספרת טעות. המנחה מדייקת בקצרה. הקבוצה ממתינה לתור הבא.',   'לוח מנחה — באנר תוצאה (כתום) אחרי לחיצת "חלקי"'],
  ['uiCopy.messages.wrongOutcome',          'נספרת טעות. הקבוצה חוזרת תחנה אחת אחורה.',                  'לוח מנחה — באנר תוצאה (אדום) אחרי לחיצת "שגוי"'],
  ['uiCopy.messages.wrongOutcomeStayed',    'נספרת טעות. הקבוצה נשארת בתחנה הראשונה.', 'לוח מנחה — וריאנט של הודעת שגיאה (תחנה ראשונה)'],
  ['uiCopy.messages.autoAdvanceNotice',     'בתור הבא הקבוצה מתקדמת אוטומטית.',         'לוח מנחה — הודעת מידע על קבוצה שתתקדם אוטומטית'],
  ['uiCopy.messages.mobileTasksFallback',   'תיק המשימות לקבוצה זו יושלם בשלב הבא.',    'תיק מובייל — הודעה כשאין תוכן משימות עדיין'],
  ['uiCopy.messages.mobileNoCompensation',  'אין נתוני מודל תגמול.',                     'תיק מובייל — הודעה בלשונית "מודל תגמול" כשהיא ריקה'],
  ['uiCopy.messages.noGroupAnswer',         'אין תשובה ייחודית לקבוצה זו.',              'חלון תחנה — הודעה כשאין תשובה ספציפית לקבוצה זו'],
  ['uiCopy.messages.noStationQuestion',     'תחנת מעבר — אין שאלה.',                     'חלון תחנה — הודעה בתחנות מעבר שאין בהן שאלה'],
]));
ch.push(gap());

// ── 9. תצוגת מובייל ────────────────────────────────────
ch.push(h1('9. תצוגת מובייל — תיק הקבוצה'), note('מקור: gameContent.json → uiCopy.mobileFolder, uiCopy.tabs, uiCopy.reminderDrawer'));
ch.push(makeTable([
  ['uiCopy.mobileFolder.lockedMessage', 'תור {group} — המתינו לתורכם',      'תיק מובייל — כותרת מסך נעילה (גדולה, בעברית)'],
  ['uiCopy.mobileFolder.lockedWait',    'התיק שלכם ייפתח כשיגיע תורכם',     'תיק מובייל — כיתוב מתחת לכותרת הנעילה'],
  ['uiCopy.tabs.reviewTasks',           'משימות לבדיקה',                      'תיק מובייל — לשונית ראשונה (ברירת מחדל)'],
  ['uiCopy.tabs.compensationModel',     'מודל תגמול לבדיקה',                  'תיק מובייל — לשונית שנייה'],
  ['uiCopy.buttons.needReminder',       'צריכים תזכורת?',                     'תיק מובייל — כפתור בתחתית שפותח מגירת הגדרות'],
  ['uiCopy.buttons.openFolder',         'פתחו את תכנית העבודה השנתית',       'תיק מובייל — כפתור קישור לתכנית העבודה החטיבתית'],
  ['uiCopy.reminderDrawer.title',       'הגדרות בסיסיות',                     'מגירת תזכורת — כותרת החלק העליון'],
  ['uiCopy.completedBadge',             'הושלמה',                             'תיק מובייל — תג ירוק על משימה שסומנה כהושלמה'],
]));
ch.push(gap());

// ── 10. שדות טופס ───────────────────────────────────────
ch.push(h1('10. שדות טופס ותוויות'), note('מקור: gameContent.json → uiCopy.taskSections, taskFields, compensationFields, answerCard'));
ch.push(makeTable([
  ['uiCopy.taskSections.strategic',                  'שיוך אסטרטגי',          'תיק מובייל — כותרת קטע (סקשן) בכרטיס משימה'],
  ['uiCopy.taskSections.details',                    'פרטי המשימה',            'תיק מובייל — כותרת קטע בכרטיס משימה'],
  ['uiCopy.taskSections.planning',                   'תכנון לאורך השנה',       'תיק מובייל — כותרת קטע בכרטיס משימה'],
  ['uiCopy.taskFields.companyGoal',                  'מטרת חברה',              'תיק מובייל — תווית שדה בקטע "שיוך אסטרטגי"'],
  ['uiCopy.taskFields.goal',                         'יעד',                    'תיק מובייל — תווית שדה'],
  ['uiCopy.taskFields.outcomeMetric',                'מדד תוצאה',              'תיק מובייל — תווית שדה'],
  ['uiCopy.taskFields.classification',               'סיווג',                  'תיק מובייל — תווית שדה (Top 10 / שוטפת)'],
  ['uiCopy.taskFields.taskDescription',              'תיאור משימה',            'תיק מובייל — תווית שדה בקטע "פרטי המשימה"'],
  ['uiCopy.taskFields.annualAchievement',            'הישג שנתי נדרש',         'תיק מובייל — תווית שדה'],
  ['uiCopy.taskFields.outputMetricType',             'סוג מדד תפוקה',          'תיק מובייל — תווית שדה (בינארי / כמותי וכו׳)'],
  ['uiCopy.classificationLabels.top10',              'Top 10',                 'תיק מובייל — ערך בשדה "סיווג"'],
  ['uiCopy.classificationLabels.routine',            'שוטפת',                  'תיק מובייל — ערך בשדה "סיווג"'],
  ['uiCopy.compensationFields.rewardMetric',         'מדד תגמול',              'תיק מובייל — תווית שדה בלשונית "מודל תגמול"'],
  ['uiCopy.compensationFields.formula',              'נוסחה',                  'תיק מובייל — תווית שדה'],
  ['uiCopy.compensationFields.weight',               'משקל',                   'תיק מובייל — תווית שדה'],
  ['uiCopy.compensationFields.performanceThresholds','ספי ביצוע',              'תיק מובייל — תווית שדה'],
  ['uiCopy.compensationFields.dataSource',           'מקור נתונים',            'תיק מובייל — תווית שדה'],
  ['uiCopy.answerCard.taskLabel',                    'משימה:',                 'חלון תחנה — תווית בכרטיס התשובה (אחרי חשיפה)'],
  ['uiCopy.answerCard.whatIsWrongLabel',             'מה לא תקין:',            'חלון תחנה — תווית בכרטיס התשובה'],
  ['uiCopy.answerCard.correctFixLabel',              'תיקון נכון:',            'חלון תחנה — תווית בכרטיס התשובה'],
  ['uiCopy.answerCard.facilitatorNoteLabel',         'הערת מנחה:',             'חלון תחנה — תווית בכרטיס התשובה'],
]));
ch.push(gap());

// ── 11. תחנות ───────────────────────────────────────────
ch.push(h1('11. תחנות — שמות וחודשים'), note('מקור: gameContent.json → stations[].name / stations[].month'));
ch.push(makeTable([
  ['stations[station_start].name',           'כולם לעלות',              'לוח מנחה — שם תחנת הפתיחה על המסלול'],
  ['stations[station_start].month',          'יוני',                    'לוח מנחה — חודש על המסלול'],
  ['stations[station_company_goal].name',    'מטרת חברה ויעד',          'לוח מנחה — שם תחנה מקצועית על המסלול'],
  ['stations[station_company_goal].month',   'יוני',                    'לוח מנחה — חודש'],
  ['stations[station_unexpected_aug].name',  'בלת"מ',                   'לוח מנחה — שם תחנת בלת"מ (אוגוסט)'],
  ['stations[station_unexpected_aug].month', 'יוני',                    'לוח מנחה — חודש'],
  ['stations[station_outcome_metric].name',  'מדד תוצאה',               'לוח מנחה — שם תחנה מקצועית'],
  ['stations[station_outcome_metric].month', 'יולי',                    'לוח מנחה — חודש'],
  ['stations[station_top10].name',           'Top 10 / שוטפת',          'לוח מנחה — שם תחנה מקצועית'],
  ['stations[station_top10].month',          'אוגוסט',                  'לוח מנחה — חודש'],
  ['stations[station_measurable].name',      'ניסוח מדיד',              'לוח מנחה — שם תחנה מקצועית'],
  ['stations[station_measurable].month',     'ספטמבר',                  'לוח מנחה — חודש'],
  ['stations[station_unexpected_sep].name',  'בלת"מ',                   'לוח מנחה — שם תחנת בלת"מ (ספטמבר)'],
  ['stations[station_unexpected_sep].month', 'ספטמבר',                  'לוח מנחה — חודש'],
  ['stations[station_output_metric].name',   'מדד תפוקה',               'לוח מנחה — שם תחנה מקצועית'],
  ['stations[station_output_metric].month',  'אוקטובר',                 'לוח מנחה — חודש'],
  ['stations[station_unexpected_oct].name',  'בלת"מ',                   'לוח מנחה — שם תחנת בלת"מ (אוקטובר)'],
  ['stations[station_unexpected_oct].month', 'אוקטובר',                 'לוח מנחה — חודש'],
  ['stations[station_milestones].name',      'אבני דרך רבעוניות',       'לוח מנחה — שם תחנה מקצועית'],
  ['stations[station_milestones].month',     'נובמבר',                  'לוח מנחה — חודש'],
  ['stations[station_unexpected_nov].name',  'בלת"מ',                   'לוח מנחה — שם תחנת בלת"מ (נובמבר)'],
  ['stations[station_unexpected_nov].month', 'נובמבר',                  'לוח מנחה — חודש'],
  ['stations[station_reward_metric].name',   'מדד תגמול',               'לוח מנחה — שם תחנה מקצועית'],
  ['stations[station_reward_metric].month',  'דצמבר',                   'לוח מנחה — חודש'],
  ['stations[station_unexpected_comp].name', 'בלת"מ',                   'לוח מנחה — שם תחנת בלת"מ (דצמבר)'],
  ['stations[station_unexpected_comp].month','דצמבר',                   'לוח מנחה — חודש'],
  ['stations[station_weight].name',          'משקל וספי ביצוע',         'לוח מנחה — שם תחנה מקצועית'],
  ['stations[station_weight].month',         'ינואר',                   'לוח מנחה — חודש'],
  ['stations[station_end].name',             'מודל תגמול מאושר',        'לוח מנחה — שם תחנת הסיום'],
  ['stations[station_end].month',            'ינואר',                   'לוח מנחה — חודש'],
  ['stations[station_reflection].name',      'התחנה הבאה שלי',          'לוח מנחה — שם תחנת רפלקציה'],
  ['stations[station_reflection].month',     'ינואר',                   'לוח מנחה — חודש'],
  ['mapGates[gate_workplan_approved].name',  'אישור תכנית העבודה',      'לוח מנחה — שם שער (gate) על המסלול'],
]));
ch.push(gap());

// ── 12. שאלות המשחק ─────────────────────────────────────
ch.push(h1('12. שאלות המשחק'), note('מקור: gameContent.json → questions[].questionText / questions[].professionalPrinciple'));

const questions = [
  ['q_company_goal',  'מטרת חברה ויעד',
    'עיינו במשימות לבדיקה. אחת מהן אינה משויכת בצורה מדויקת למטרת החברה או ליעד. אתרו אותה, תקנו את השיוך ונמקו.',
    'תשובה נכונה מזהה את המשימה עם השיוך השגוי ומתקנת כך שהמשימה תיגזר ממטרת החברה ותשרת יעד מתאים.'],
  ['q_outcome_metric','מדד תוצאה',
    'במשימות לבדיקה מופיע בלבול: פעולה הוצגה כמדד תוצאה או מדד הוצג כמשימה. אתרו את הבלבול ותקנו.',
    'משימה היא פעולה מרכזית שמבצעים בפועל. מדד תוצאה בוחן אם היעד הושג. תשובה נכונה מסווגת מחדש את הרכיבים ומנמקת לפי ההבחנה הזו.'],
  ['q_top10',         'Top 10 / שוטפת',
    'עיינו במשימות לבדיקה. אחת מהן סווגה באופן לא מדויק כ-Top 10 או כשוטפת. אתרו אותה, תקנו את הסיווג ונמקו.',
    'משימת Top 10 צריכה להיות בנתיב הקריטי להשגת יעדי החברה, בעלת חשיבות גבוהה, המשכיות, התאמה למשאבים/תקציב והשפעה משמעותית. משימה שוטפת או אדמיניסטרטיבית לא תסווג כ-Top 10.'],
  ['q_measurable',    'ניסוח מדיד',
    'במשימות לבדיקה יש משימה שנוסחה באופן עמום מדי. אתרו אותה ונסחו אותה מחדש כך שיהיה ברור מה מבוצע, למי, באיזה היקף ועד מתי.',
    'ניסוח טוב כולל פעולה ברורה, תוצר מדיד, היקף או יעד, ומועד. ניסוחים כלליים כמו "הנגשת מידע", "שיפור ממשקים" או "קידום תהליך" אינם מספיקים ללא פירוט מדיד.'],
  ['q_output_metric', 'מדד תפוקה',
    'באחת מהמשימות לבדיקה נבחר סוג מדד תפוקה שאינו מתאים. אתרו את המשימה, בחרו סוג מדד מתאים ונמקו.',
    'בינארי מתאים לבוצע/לא בוצע; כמותי עולה לשאיפה למקסימום; כמותי יורד לשאיפה למינימום; אחוזים להתקדמות יחסית. התשובה צריכה לחבר בין אופי המשימה, ההישג השנתי וסוג המדד.'],
  ['q_milestones',    'אבני דרך רבעוניות',
    'עיינו במשימות לבדיקה. באחת מהן פריסת אבני הדרך אינה מאפשרת בקרה לאורך השנה. אתרו אותה והציעו פריסה מתוקנת.',
    'פריסה תקינה מאפשרת ניהול ובקרה לאורך השנה. משימה שכל אבני הדרך שלה מרוכזות ברבעון 4 נחשבת לתכנון לא מאתגר.'],
  ['q_reward_metric', 'מדד תגמול',
    'במודל התגמול לבדיקה יש מדד או נוסחה שאינם מספיק ברורים. אתרו את הבעיה ותקנו.',
    'מדד תגמול צריך להגדיר כיצד מודדים הצלחה של יחידה לצורך תגמול. נוסחה צריכה להבהיר איך מחשבים את הביצוע בפועל — מה נספר, איך מחושב ומה מקור הנתונים.'],
  ['q_weight',        'משקל וספי ביצוע',
    'בדקו את מודל התגמול לבדיקה: האם המשקולות מסתכמות ל-100? האם ספי הביצוע ריאליים ומאתגרים?',
    'כל המשקולות יחד צריכות להסתכם ל-100. ספי הביצוע צריכים להגדיר עמידה מלאה, חלקית או אי-עמידה, ולהיות ריאליים ומאתגרים.'],
];

questions.forEach(([id, station, q, principle]) => {
  ch.push(h2('תחנת ' + station));
  ch.push(makeTable([
    ['questions[' + id + '].questionText',          q,         'חלון תחנה — שאלה שמוצגת למנחה (מופיעה בראש החלון)'],
    ['questions[' + id + '].professionalPrinciple', principle, 'חלון תחנה — "תשובה נכונה" שמתגלה אחרי לחיצת "חשפו תשובה"'],
  ]));
  ch.push(gap());
});

// ── 13. מסכי מעבר ───────────────────────────────────────
ch.push(h1('13. מסכי מעבר (תחנות transition)'), note('מקור: gameContent.json → transitionMessages'));
ch.push(makeTable([
  ['transitionMessages.transition_start.title',             'כולם לעלות!',                    'חלון תחנת מעבר — כותרת (תחנת הפתיחה)'],
  ['transitionMessages.transition_start.body',              'סרקו את הברקוד של הקבוצה ופתחו את תיק תכנית העבודה החטיבתי שלכם. עברו על מבנה התיק: המשימות ומודל התגמול.', 'חלון תחנת מעבר — גוף טקסט'],
  ['transitionMessages.transition_gate_divisional.title',   'שער האישור החטיבתי',             'חלון תחנת מעבר — כותרת'],
  ['transitionMessages.transition_gate_divisional.body',    'הגעתם לשער האישור החטיבתי. אם השלמתם את תחנות הליבה הקודמות — ניתן להעביר את התוכנית למנהל אגף ולסמנכ"ל.', 'חלון תחנת מעבר — גוף טקסט'],
  ['transitionMessages.transition_deputy_ceo.title',        'משנה למנכ"ל — נעילה לעריכה',    'חלון תחנת מעבר — כותרת'],
  ['transitionMessages.transition_deputy_ceo.body',         'התוכנית הועברה למשנה למנכ"ל. מכאן ואילך המשימה נעולה לעריכה וניתנת לצפייה בלבד.', 'חלון תחנת מעבר — גוף טקסט'],
  ['transitionMessages.transition_workplan_approved.title', 'טו-טו! תכנית העבודה אושרה',     'חלון תחנת מעבר — כותרת חגיגית'],
  ['transitionMessages.transition_workplan_approved.body',  'תכנית העבודה אושרה על ידי מנכ"ל ודירקטוריון. פתחתם את מודל התגמול. מכאן ממשיכים: מדד תגמול, נוסחה, משקל וספי ביצוע.', 'חלון תחנת מעבר — גוף טקסט'],
  ['transitionMessages.transition_model_approved.title',    'טו-טו! מודל התגמול אושר',        'חלון תחנת מעבר — כותרת חגיגית (תחנת הסיום)'],
  ['transitionMessages.transition_model_approved.body',     'מודל התגמול אושר. הגעתם ליעד.', 'חלון תחנת מעבר — גוף טקסט'],
]));
ch.push(gap());

// ── 14. מסך סיום ────────────────────────────────────────
ch.push(h1('14. מסך סיום'), note('מקור: gameContent.json → uiCopy.endScreen'));
ch.push(makeTable([
  ['uiCopy.endScreen.title',              'המשחק הסתיים',            'מסך סיום — כותרת ראשית'],
  ['uiCopy.endScreen.winnerLabel',        'הקבוצה המנצחת',           'מסך סיום — תווית מעל שם הקבוצה הזוכה'],
  ['uiCopy.endScreen.commitmentQuestion', 'מה הפעולה הראשונה שתעשו ב-72 השעות הקרובות כדי לקדם את תכנית העבודה ואת מודל התגמול בחטיבה שלכם?', 'מסך סיום — שאלת ההתחייבות האישית (גדולה, מרכז המסך)'],
  ['uiCopy.endScreen.mentimeterUrl',      'יושלם בתיאום עם הלקוחה', 'מסך סיום — כתובת ה-QR/Mentimeter (טרם נקבעה)'],
  ['uiCopy.endScreen.mentimeterQrLabel',  'סרקו להתחייבות אישית',   'מסך סיום — כיתוב מתחת ל-QR Code'],
]));
ch.push(gap());

// ── 15. מילון מושגים ────────────────────────────────────
ch.push(h1('15. מילון מושגים — מגירת התזכורת'), note('מקור: gameContent.json → reminderDefinitions[]'));
ch.push(makeTable([
  ['reminderDefinitions[0].term',       'משימה',        'מגירת תזכורת — שם המושג (מודגש)'],
  ['reminderDefinitions[0].definition', 'פעולה מרכזית שמבצעים בפועל להשגת יעד.', 'מגירת תזכורת — הגדרה מתחת למושג'],
  ['reminderDefinitions[1].term',       'מדד תוצאה',   'מגירת תזכורת — שם המושג'],
  ['reminderDefinitions[1].definition', 'מדד הבוחן אם היעד הושג — נמדד מול היעד.', 'מגירת תזכורת — הגדרה'],
  ['reminderDefinitions[2].term',       'מדד תפוקה',   'מגירת תזכורת — שם המושג'],
  ['reminderDefinitions[2].definition', 'מדד הבוחן את ביצוע המשימה — נמדד מול המשימה. סוגים: בינארי, כמותי עולה, כמותי יורד, אחוזים.', 'מגירת תזכורת — הגדרה'],
  ['reminderDefinitions[3].term',       'מדד תגמול',   'מגירת תזכורת — שם המושג'],
  ['reminderDefinitions[3].definition', 'מדד המגדיר כיצד מודדים הצלחה של יחידה לצורך תגמול, כולל נוסחה, משקל וספי ביצוע.', 'מגירת תזכורת — הגדרה'],
]));
ch.push(gap());

// ── 16. קבוצות ──────────────────────────────────────────
ch.push(h1('16. קבוצות'), note('מקור: gameContent.json → groups[].label / groups[].divisionName'));
ch.push(makeTable([
  ['groups[group_1].label',        'קבוצה 1',                      'לוח מנחה + תיק מובייל — שם הקבוצה'],
  ['groups[group_1].divisionName', 'חטיבת נייד',                   'לוח מנחה + תיק מובייל — שם החטיבה'],
  ['groups[group_2].label',        'קבוצה 2',                      'לוח מנחה + תיק מובייל — שם הקבוצה'],
  ['groups[group_2].divisionName', 'חטיבת משאבי אנוש',             'לוח מנחה + תיק מובייל — שם החטיבה'],
  ['groups[group_3].label',        'קבוצה 3',                      'לוח מנחה + תיק מובייל — שם הקבוצה'],
  ["groups[group_3].divisionName", 'חטיבת בטיחות, ביטחון ואיכ"ס', 'לוח מנחה + תיק מובייל — שם החטיבה'],
  ['groups[group_4].label',        'קבוצה 4',                      'לוח מנחה + תיק מובייל — שם הקבוצה'],
  ['groups[group_4].divisionName', 'חטיבת פיתוח עסקי',             'לוח מנחה + תיק מובייל — שם החטיבה'],
]));
ch.push(gap());

// ── 17. חודשים ──────────────────────────────────────────
ch.push(h1('17. חודשים'), note('מקור: gameContent.json → uiCopy.months[]'));
ch.push(makeTable([
  ['uiCopy.months[0]', 'יוני',    'לוח מנחה — תווית חודש על ציר הזמן'],
  ['uiCopy.months[1]', 'יולי',    'לוח מנחה — תווית חודש על ציר הזמן'],
  ['uiCopy.months[2]', 'אוגוסט',  'לוח מנחה — תווית חודש על ציר הזמן'],
  ['uiCopy.months[3]', 'ספטמבר',  'לוח מנחה — תווית חודש על ציר הזמן'],
  ['uiCopy.months[4]', 'אוקטובר', 'לוח מנחה — תווית חודש על ציר הזמן'],
  ['uiCopy.months[5]', 'נובמבר',  'לוח מנחה — תווית חודש על ציר הזמן'],
  ['uiCopy.months[6]', 'דצמבר',   'לוח מנחה — תווית חודש על ציר הזמן'],
  ['uiCopy.months[7]', 'ינואר',   'לוח מנחה — תווית חודש על ציר הזמן'],
]));
ch.push(gap());

// ── 18. מיקרו-קופי שונות ────────────────────────────────
ch.push(h1('18. מיקרו-קופי שונות'), note('מקור: gameContent.json → uiCopy שונות'));
ch.push(makeTable([
  ['uiCopy.stationTypes.unexpectedEventLabel', 'בלת"מ',                        'תג על תחנות בלת"מ במסלול'],
  ['uiCopy.points.correct',                    '10 נקודות',                    'הודעת תוצאה — ניקוד תשובה נכונה'],
  ['uiCopy.points.partial',                    '5 נקודות',                     'הודעת תוצאה — ניקוד תשובה חלקית'],
  ['uiCopy.points.wrong',                      '0 נקודות',                     'הודעת תוצאה — ניקוד תשובה שגויה'],
  ['uiCopy.outcomeBanner.correct/partial/wrong','{group}: {msg}',              'לוח מנחה — תבנית פורמט באנר תוצאה'],
  ['uiCopy.boardSections.a',                   'מקטע א — הקמת תכנית עבודה',   'לוח מנחה — כותרת מקטע ראשון על ציר הזמן'],
  ['uiCopy.boardSections.aMonths',             'יוני — ספטמבר',               'לוח מנחה — תווית חודשים למקטע א'],
  ['uiCopy.boardSections.b',                   'המשך תכנית עבודה ואישורים',   'לוח מנחה — כותרת מקטע שני'],
  ['uiCopy.boardSections.bMonths',             'ספטמבר — נובמבר',             'לוח מנחה — תווית חודשים למקטע ב'],
  ['uiCopy.boardSections.c',                   'הקמת מודל תגמול',             'לוח מנחה — כותרת מקטע שלישי'],
  ['uiCopy.boardSections.cMonths',             'דצמבר — ינואר',               'לוח מנחה — תווית חודשים למקטע ג'],
]));
ch.push(gap());

// ── 19. חלון תחנה — בלת"מ ───────────────────────────────
ch.push(h1('19. חלון תחנה — בלת"מ'), note('מקור: hardcoded HTML (renderStationModal → unexpected_event branch)'));
ch.push(makeTable([
  ['hardcoded: UE title',                'גלגל הבלת"מ',                     'חלון בלת"מ — כותרת ראשית בראש החלון'],
  ['hardcoded: UE result header',        '⚡ תוצאת הגלגל',                  'חלון בלת"מ — כותרת קטע לאחר סיבוב הגלגל'],
  ['hardcoded: UE no-delay badge',       '✓ ממשיכים קדימה — אין עיכוב',    'חלון בלת"מ — תג ירוק כשתוצאת הגלגל: אין עיכוב'],
  ['hardcoded: UE delay badge',          '⏸ הקבוצה ממתינה סבב אחד',        'חלון בלת"מ — תג כתום כשתוצאת הגלגל: עיכוב'],
  ['hardcoded: UE confirm button',       'הבנו, ממשיכים ←',                 'חלון בלת"מ — כפתור אישור וסגירה (בלת"מ חדש)'],
  ['hardcoded: UE already-used button',  'ממשיכים קדימה ←',                 'חלון בלת"מ — כפתור כשהבלת"מ כבר שומש בעבר'],
  ['hardcoded: close tooltip',           'סגירה ללא מעבר תור',              'חלון בלת"מ — tooltip על כפתור ה-X'],
]));
ch.push(gap());

// ── 20. פופ-אפ תוצאה ─────────────────────────────────────
ch.push(h1('20. פופ-אפ תוצאה (אחרי נכון / חלקי / שגוי)'), note('מקור: hardcoded HTML (renderStationModal → outcomePopup)'));
ch.push(makeTable([
  ['hardcoded: outcome icon — correct',   '✅',                                       'פופ-אפ תוצאה — אייקון גדול בראש (תשובה נכונה)'],
  ['hardcoded: outcome icon — partial',   '⚠️',                                       'פופ-אפ תוצאה — אייקון (תשובה חלקית)'],
  ['hardcoded: outcome icon — wrong',     '❌',                                       'פופ-אפ תוצאה — אייקון (תשובה שגויה)'],
  ['hardcoded: outcome title — correct',  'תשובה נכונה!',                            'פופ-אפ תוצאה — כותרת ירוקה גדולה'],
  ['hardcoded: outcome title — partial',  'תשובה חלקית',                             'פופ-אפ תוצאה — כותרת כתומה'],
  ['hardcoded: outcome title — wrong',    'תשובה שגויה',                             'פופ-אפ תוצאה — כותרת אדומה'],
  ['hardcoded: outcome points — correct', '+10',                                      'פופ-אפ תוצאה — מספר הנקודות (גדול)'],
  ['hardcoded: outcome points — partial', '+5',                                       'פופ-אפ תוצאה — מספר הנקודות'],
  ['hardcoded: outcome points — wrong',   '+0',                                       'פופ-אפ תוצאה — מספר הנקודות'],
  ['hardcoded: outcome points suffix',    'נקודות',                                   'פופ-אפ תוצאה — טקסט ליד מספר הנקודות'],
  ['hardcoded: outcome next — wrong',     'מתקדמים לתחנה הבאה — ממתינים סבב אחד ⏳', 'פופ-אפ תוצאה — הודעת השלכה (תשובה שגויה)'],
  ['hardcoded: outcome next — correct',   'מתקדמים לתחנה הבאה 🚀',                  'פופ-אפ תוצאה — הודעת השלכה (נכון/חלקי)'],
  ['hardcoded: outcome dismiss button',   'הבנו, ממשיכים ←',                         'פופ-אפ תוצאה — כפתור סגירה'],
]));
ch.push(gap());

// ── 21. תצוגת מובייל — כפתור חזרה ───────────────────────
ch.push(h1('21. תצוגת מובייל — כפתור חזרה ללוח'), note('מקור: hardcoded HTML (renderMobile → backBar) — מוצג רק כשהמנחה פותחת תיק מתוך הלוח'));
ch.push(makeTable([
  ['hardcoded: mobile backBar button', '← חזרה ללוח', 'תיק מובייל (כשנפתח על-ידי המנחה) — כפתור בראש הדף לחזרה ללוח'],
]));
ch.push(gap());

// ── 22. תצוגת תוכן פיתוח (☰) ────────────────────────────
ch.push(h1('22. תצוגת תוכן פיתוח (כפתור ☰)'), note('מקור: hardcoded HTML (renderDevContentView) — נגישה דרך כפתור ☰ בפאנל הטלפונים'));
ch.push(makeTable([
  ['hardcoded: devContent header',               'סקירת תוכן — כל התחנות',  'תצוגת ☰ — כותרת ראשית של הדף'],
  ['hardcoded: devContent station count suffix', 'תחנות',                   'תצוגת ☰ — ״16 תחנות״ (כיתוב ספירה)'],
  ['hardcoded: devContent — no transition msg',  'תחנת מעבר',               'תצוגת ☰ — כשתחנת מעבר אין לה תוכן מיוחד'],
  ['hardcoded: devContent — no UE items',        'אין פריטים',              'תצוגת ☰ — כשתחנת בלת"מ ריקה מתוכן'],
  ['hardcoded: devContent — no question',        'אין שאלה',                'תצוגת ☰ — כשתחנה מקצועית ללא שאלה'],
  ['hardcoded: devContent type badge',           'מקצועית / בלת"מ / מעבר / רפלקציה', 'תצוגת ☰ — תג סוג תחנה על כל כרטיס'],
]));
ch.push(gap());

// ── 23. מיקרו-קופי DEV MODE בלבד ──────────────────────────
ch.push(h1('23. מיקרו-קופי DEV MODE בלבד'), note('מקור: hardcoded HTML — מופיע רק כשהאפליקציה רצה עם ?devmode=true. לא נראה למשתתפים.'));
ch.push(makeTable([
  ['hardcoded: nav-content button',               '📋 תוכן',                           'נאב-בר — כפתור ניווט לתצוגת תוכן פיתוח (נסתר; מופיע רק ב-DEV mode)'],
  ['hardcoded: DEV UE results header',           '⚙ כל התוצאות האפשריות (DEV MODE)', 'חלון בלת"מ — כותרת קטע בחירת תוצאה ידנית'],
  ['hardcoded: DEV UE select button',            '⚡ בחר',                            'חלון בלת"מ — כפתור בחירת תוצאה ספציפית'],
  ['hardcoded: DEV UE no-delay option',          '✓ אין עיכוב',                       'חלון בלת"מ — אפשרות: תוצאה ללא עיכוב'],
  ['hardcoded: DEV UE delay option',             '⏸ המתנה סבב',                       'חלון בלת"מ — אפשרות: תוצאה עם עיכוב'],
  ['hardcoded: DEV task error marker',           '⚠️ טעות מוטמנת',                    'תיק מובייל — תג על משימה שיש בה טעות מוטמנת'],
  ['hardcoded: DEV station hint prefix',         '📍 תחנה:',                          'תיק מובייל — תווית הינט לפיתוח'],
  ['hardcoded: DEV hint labels',                 'שאלה: / מה שגוי: / תיקון:',        'תיק מובייל — תוויות הינט לפיתוח'],
  ['hardcoded: renderDevView title',             '📋 סקירת תוכן — מצב פיתוח',        'תצוגת DEV — כותרת ראשית'],
  ['hardcoded: renderDevView incomplete badge',  '⚠ להשלמה',                          'תצוגת DEV — תג על תחנה עם שדות חסרים'],
  ['hardcoded: renderDevView complete badge',    '✓ הכל מלא',                         'תצוגת DEV — תג על תחנה שלמה'],
  ['hardcoded: renderDevView field suffix',      'שדות לא מלאים',                     'תצוגת DEV — כיתוב מספר שדות חסרים'],
  ['hardcoded: renderDevView note',              'שדות ... דורשים תיאום עם הלקוחה', 'תצוגת DEV — הערה על שדות הממתינים לאישור'],
  ['hardcoded: renderDevView section headers',   'תחנות מקצועיות / ⚡ תחנות בלת"מ / מודל תגמול / משימות גלויות / מסך סיום', 'תצוגת DEV — כותרות קטגוריות'],
  ['hardcoded: renderDevView waitRounds label',  'המתנה: {n} סבבים',                  'תצוגת DEV — כמה סבבי המתנה מוגדרים לתחנה'],
]));

// ── 24–27. תוכן תיק מובייל — לפי קבוצה ─────────────────

function taskRows(tid, t) {
  const cls = t.classification === 'top10' ? 'Top 10' : 'שוטפת';
  return [
    [tid + '.taskName',          t.taskName,          'תיק מובייל — כותרת כרטיס המשימה'],
    [tid + '.companyGoal',       t.companyGoal,       'תיק מובייל — שדה "מטרת חברה"'],
    [tid + '.goal',              t.goal,              'תיק מובייל — שדה "יעד"'],
    [tid + '.outcomeMetric',     t.outcomeMetric,     'תיק מובייל — שדה "מדד תוצאה"'],
    [tid + '.classification',    cls,                 'תיק מובייל — שדה "סיווג" (Top 10 / שוטפת)'],
    [tid + '.taskDescription',   t.taskDescription,   'תיק מובייל — שדה "תיאור משימה"'],
    [tid + '.annualAchievement', t.annualAchievement, 'תיק מובייל — שדה "הישג שנתי נדרש"'],
    [tid + '.outputMetricType',  t.outputMetricType,  'תיק מובייל — שדה "סוג מדד תפוקה"'],
    [tid + '.milestones.Q1',     t.milestones.Q1,     'תיק מובייל — אבן דרך Q1 (רבעון 1)'],
    [tid + '.milestones.Q2',     t.milestones.Q2,     'תיק מובייל — אבן דרך Q2 (רבעון 2)'],
    [tid + '.milestones.Q3',     t.milestones.Q3,     'תיק מובייל — אבן דרך Q3 (רבעון 3)'],
    [tid + '.milestones.Q4',     t.milestones.Q4,     'תיק מובייל — אבן דרך Q4 (רבעון 4)'],
    [tid + '.devHint',           t.devHint,           'DEV MODE — רמז "⚠️ טעות מוטמנת" (לא נראה למשתתפים)'],
  ];
}

function compRows(gid, m) {
  return [
    [gid + '.comp.modelName',            m.modelName,            'תיק מובייל — כותרת מודל התגמול'],
    [gid + '.comp.rewardMetric',         m.rewardMetric,         'תיק מובייל — שדה "מדד תגמול"'],
    [gid + '.comp.formula',              m.formula,              'תיק מובייל — שדה "נוסחה"'],
    [gid + '.comp.weight',               m.weight,               'תיק מובייל — שדה "משקל"'],
    [gid + '.comp.performanceThresholds',m.performanceThresholds,'תיק מובייל — שדה "ספי ביצוע"'],
    [gid + '.comp.dataSource',           m.dataSource,           'תיק מובייל — שדה "מקור נתונים"'],
    [gid + '.comp.notes',                m.notes,                'תיק מובייל — הערת DEV (לא נראית למשתתפים)'],
    [gid + '.comp.devHint',              m.devHint,              'DEV MODE — רמז "⚠️ טעות מוטמנת" במודל תגמול'],
  ];
}

const GROUP_LABELS = {
  group_1: 'קבוצה 1 — חטיבת נייד',
  group_2: 'קבוצה 2 — חטיבת משאבי אנוש',
  group_3: 'קבוצה 3 — חטיבת בטיחות, ביטחון ואיכ"ס',
  group_4: 'קבוצה 4 — חטיבת פיתוח עסקי',
};

const GROUPS_CONTENT = Object.keys(GC.visibleTasks).map(gid => ({
  gid,
  label: GROUP_LABELS[gid] || gid,
  tasks: GC.visibleTasks[gid],
  comp: GC.compensationModels[gid],
}));


GROUPS_CONTENT.forEach((g, gi) => {
  const sNum = 24 + gi;
  ch.push(h1(`${sNum}. תוכן תיק מובייל — ${g.label}`), note(`מקור: moadilim-poc.html → const GC → visibleTasks.${g.gid} + compensationModels.${g.gid}`));
  ch.push(h2('משימות לבדיקה'));
  g.tasks.forEach((t, ti) => {
    ch.push(h2(`משימה ${ti + 1}: ${t.taskName}`));
    ch.push(makeTable(taskRows(t.taskId, t)));
    ch.push(gap());
  });
  ch.push(h2('מודל תגמול לבדיקה'));
  ch.push(makeTable(compRows(g.gid, g.comp)));
  ch.push(gap());
});

// ── 28. תשובות לפי קבוצה ותחנה ─────────────────────────
ch.push(h1('28. תשובות לפי קבוצה ותחנה'), note('מקור: gameContent.json → stationQuestions[].answersByGroup — מוצגות בכרטיס התשובה אחרי לחיצת "חשפו תשובה"'));

const ANSWERS_BY_STATION = [
  { qid: 'q_company_goal', stationName: 'מטרת חברה ויעד', answers: [
    { g: 'קבוצה 1', taskName: 'שדרוג סלון נוסעים', whatIsWrong: 'המשימה שויכה בטעות למטרה "פיתוח מנופי צמיחה" וליעד עסקי.', correctFix: 'לשייך למטרה: השירות כערך עליון, וליעד: הקניית רמת שירות מיטבית ללקוחות הרכבת.' },
    { g: 'קבוצה 2', taskName: 'צמצום ימי הכשרות נהגי נוסעים בקורס ישיר', whatIsWrong: 'המשימה שויכה בטעות למטרה "שירות כערך עליון".', correctFix: 'לשייך למטרה "פיתוח ושימור ההון האנושי" וליעד של פיתוח מקצועיות, כשירות ותרבות למידה.' },
    { g: 'קבוצה 3', taskName: 'היערכות הרכבת בחירום', whatIsWrong: 'היעד שהוגדר הוא "עמידה בציות וברגולציה", אך זה לא היעד המדויק.', correctFix: 'לשייך ליעד "מימוש מיטבי של התוכנית להפחתת סיכונים".' },
    { g: 'קבוצה 4', taskName: 'הרחבת פעילות מסחרית בתחנות מרכזיות', whatIsWrong: 'המשימה שויכה בטעות למטרה הקשורה לשירות נוסעים בלבד.', correctFix: 'לשייך ל"פיתוח מנופי צמיחה" וליעד של פעילות כלכלית, רווחית ועסקית.' },
  ]},
  { qid: 'q_outcome_metric', stationName: 'מדד תוצאה', answers: [
    { g: 'קבוצה 1', taskName: 'קליטת קרונועים חשמליים מסוג DDEMU', whatIsWrong: '"קליטת 50 קרונועים" הוגדרה כמדד תוצאה.', correctFix: 'זו תפוקה/ביצוע משימה. מדד תוצאה מתאים יותר יהיה קשור לשיעור זמינות או כשירות הצי.' },
    { g: 'קבוצה 2', taskName: 'הקמת מתחם גיוס חדש', whatIsWrong: '"בניית מתחם גיוס" הוגדרה כמדד תוצאה.', correctFix: 'זו משימה/פעולה. מדד תוצאה צריך לבחון את השפעת המהלך, למשל שיפור חוויית מועמד או קיצור תהליך גיוס.' },
    { g: 'קבוצה 3', taskName: 'הטמעת בטיחות בכלל החטיבות', whatIsWrong: '"הטמעת בטיחות" הוגדרה כמדד תוצאה.', correctFix: 'זו פעולה רחבה. מדד תוצאה מתאים יכול להיות צמצום ימי היעדרות עקב תאונות עבודה או הפחתת אירועי בטיחות.' },
    { g: 'קבוצה 4', taskName: 'שיווק והשכרה של שטחי מסחר', whatIsWrong: 'המשימה עצמה הוגדרה כמדד תוצאה.', correctFix: '"שיווק והשכרה" היא פעולה; מדד תוצאה מתאים הוא סך הכנסות מסחר או שיעור תפוסת שטחי מסחר.' },
  ]},
  { qid: 'q_top10', stationName: 'Top 10 / שוטפת', answers: [
    { g: 'קבוצה 1', taskName: 'הגשת דוח רבעוני למנכ"ל', whatIsWrong: 'המשימה סווגה כ-Top 10.', correctFix: 'זו משימה שוטפת / אדמיניסטרטיבית ולא משימה בנתיב הקריטי.' },
    { g: 'קבוצה 2', taskName: 'הכנת חומרים לדיוני הנהלה בנושא הכשרות', whatIsWrong: 'המשימה סווגה כ-Top 10.', correctFix: 'זו משימה תומכת/שוטפת, אלא אם היא חלק ממהלך אסטרטגי מוגדר.' },
    { g: 'קבוצה 3', taskName: 'ניטור אוויר בתחנות הרכבת לפי רגולציה', whatIsWrong: 'המשימה סווגה כשוטפת בלבד.', correctFix: 'מאחר שמדובר במשימה רגולטורית ובנתיב קריטי לציות ולבטיחות, ניתן לסווגה כ-Top 10.' },
    { g: 'קבוצה 4', taskName: 'מעקב יומי אחרי מלאים/נכסים מסחריים', whatIsWrong: 'המשימה סווגה כ-Top 10.', correctFix: 'זו פעילות שוטפת, לא משימה אסטרטגית בנתיב הקריטי.' },
  ]},
  { qid: 'q_measurable', stationName: 'ניסוח מדיד', answers: [
    { g: 'קבוצה 1', taskName: 'שיפור זמינות הצי הנייד', whatIsWrong: 'הניסוח כללי מדי ואינו מגדיר פעולה, היקף או מועד.', correctFix: 'למשל: "קליטת 50 קרונועים חשמליים מסוג DDEMU עד סוף השנה, לפי אבני דרך רבעוניות".' },
    { g: 'קבוצה 2', taskName: 'הנגשת מידע למועמדים חדשים', whatIsWrong: '"הנגשת מידע" הוא ניסוח עמום ולא מדיד.', correctFix: 'למשל: "הקמת אזור מידע דיגיטלי למועמדים הכולל 5 תכנים מרכזיים עד סוף רבעון 2".' },
    { g: 'קבוצה 3', taskName: 'הטמעת תרבות בטיחות בכלל החטיבות', whatIsWrong: 'לא ברור מה ייחשב הטמעה, באילו חטיבות, באיזה היקף ומתי.', correctFix: 'למשל: "ביצוע 4 סדנאות בטיחות לחטיבות יעד והטמעת נוהל אחיד עד סוף רבעון 3".' },
    { g: 'קבוצה 4', taskName: 'קידום מסחר בתחנות', whatIsWrong: 'ניסוח כללי מדי.', correctFix: 'למשל: "השכרת 10 שטחי מסחר חדשים בתחנות מרכזיות עד סוף רבעון 3".' },
  ]},
  { qid: 'q_output_metric', stationName: 'מדד תפוקה', answers: [
    { g: 'קבוצה 1', taskName: 'קליטת קרונועים חשמליים מסוג DDEMU', whatIsWrong: 'סוג מדד התפוקה הוגדר כבינארי.', correctFix: 'כמותי עולה, כי מודדים כמות מצטברת של קרונועים שנקלטו.' },
    { g: 'קבוצה 2', taskName: 'צמצום ימי הכשרות נהגי נוסעים', whatIsWrong: 'סוג מדד התפוקה הוגדר ככמותי עולה.', correctFix: 'כמותי יורד או אחוז הפחתה, כי שואפים לצמצם ימי הכשרה.' },
    { g: 'קבוצה 3', taskName: 'צמצום ימי היעדרות עקב תאונת עבודה', whatIsWrong: 'סוג מדד התפוקה הוגדר ככמותי עולה.', correctFix: 'כמותי יורד, כי שואפים להפחית מספר ימים/מקרים.' },
    { g: 'קבוצה 4', taskName: 'הרחבת פעילות מסחרית בתחנות', whatIsWrong: 'סוג מדד התפוקה הוגדר כבינארי.', correctFix: 'כמותי עולה, למשל מספר שטחי מסחר שהושכרו או היקף שטח מסחרי פעיל.' },
  ]},
  { qid: 'q_milestones', stationName: 'אבני דרך רבעוניות', answers: [
    { g: 'קבוצה 1', taskName: 'קליטת קרונועים DDEMU', whatIsWrong: 'כל אבני הדרך מרוכזות ברבעון 4.', correctFix: 'לפרוס לאורך השנה: Q1: 10 קרונועים, Q2: 10, Q3: 10, Q4: 20.' },
    { g: 'קבוצה 2', taskName: 'הנגשת מידע למועמדים', whatIsWrong: 'כל העבודה הוגדרה לרבעון 4.', correctFix: 'Q1: אפיון תכנים, Q2: כתיבה והפקה, Q3: העלאה לאוויר, Q4: בדיקה ושיפור.' },
    { g: 'קבוצה 3', taskName: 'ניטור אוויר בתחנות', whatIsWrong: 'אבני הדרך מתחילות מאוחר מדי, רק Q3–Q4.', correctFix: 'Q1: אפיון צרכים ותקציב, Q2: רכש/התקשרות, Q3: התקנה, Q4: בדיקה והטמעה.' },
    { g: 'קבוצה 4', taskName: 'הרחבת פעילות מסחרית בתחנות', whatIsWrong: 'כל ההשכרות מתוכננות ל-Q4.', correctFix: 'Q1: איתור שטחים, Q2: מכרז/שיווק, Q3: חתימה, Q4: הפעלה.' },
  ]},
  { qid: 'q_reward_metric', stationName: 'מדד תגמול', answers: [
    { g: 'קבוצה 1', taskName: 'זמינות וכשירות צי נייד', whatIsWrong: 'המדד הוגדר "שיפור זמינות הצי" והנוסחה "לפי עמידה ביעד" – שניהם עמומים.', correctFix: 'להגדיר נוסחה ברורה: זמינות בפועל / יעד זמינות שנתי × 100, כולל מקור נתונים.' },
    { g: 'קבוצה 2', taskName: 'תהליכי גיוס והכשרות', whatIsWrong: 'המדד "שיפור תהליכי גיוס" עמום, והנוסחה "מספר מועמדים שטופלו" אינה מודדת שיפור.', correctFix: 'למשל: "קיצור משך תהליך הגיוס הממוצע", עם נוסחה: משך תהליך בפועל / יעד שנתי × 100.' },
    { g: 'קבוצה 3', taskName: 'בטיחות', whatIsWrong: 'המדד "שיפור בטיחות" כללי מדי, והנוסחה אינה מבהירה כיוון מדידה.', correctFix: 'למשל: "צמצום ימי היעדרות עקב תאונת עבודה", עם נוסחה ברורה ומקור נתונים.' },
    { g: 'קבוצה 4', taskName: 'הכנסות מסחר', whatIsWrong: 'הנוסחה "הכנסות בפועל" חלקית – לא ברור ביחס למה נמדדת הצלחה.', correctFix: 'הכנסות מסחר בפועל / יעד הכנסות שנתי × 100.' },
  ]},
  { qid: 'q_weight', stationName: 'משקל וספי ביצוע', answers: [
    { g: 'קבוצה 1', taskName: 'זמינות צי', whatIsWrong: 'המשקולות במדדים מסתכמות ל-105 במקום 100.', correctFix: 'לעדכן את המשקולות כך שסך כל המדדים יהיה 100.' },
    { g: 'קבוצה 2', taskName: 'גיוס והכשרות', whatIsWrong: 'המשקלים 30, 30, 25, 20 מסתכמים ל-105.', correctFix: 'לתקן כך שסך המשקולות יהיה 100.' },
    { g: 'קבוצה 3', taskName: 'בטיחות', whatIsWrong: 'ספי הביצוע לא מאתגרים: 100 = עד 200 אירועים, 70 = עד 250 אירועים, ללא הצדקה.', correctFix: 'להגדיר ספים ריאליים, מאתגרים וברורים, המבוססים על יעד ומקור נתונים.' },
    { g: 'קבוצה 4', taskName: 'הכנסות מסחר', whatIsWrong: 'הספים צפופים מדי: 100 = 5 מיליון, 70 = 4.9 מיליון, 0 = 4.8 מיליון.', correctFix: 'להגדיר מדרגות ביצוע מובחנות, ריאליות ומאתגרות.' },
  ]},
];

ANSWERS_BY_STATION.forEach(s => {
  ch.push(h2('תחנת ' + s.stationName));
  ch.push(makeTable(s.answers.flatMap(a => [
    [s.qid + '.answersByGroup.' + a.g + '.taskName',    a.taskName,    'כרטיס תשובה — שדה "משימה:"'],
    [s.qid + '.answersByGroup.' + a.g + '.whatIsWrong', a.whatIsWrong, 'כרטיס תשובה — שדה "מה לא תקין:"'],
    [s.qid + '.answersByGroup.' + a.g + '.correctFix',  a.correctFix,  'כרטיס תשובה — שדה "תיקון נכון:"'],
  ])));
  ch.push(gap());
});

// ── 29. אירועי בלת"מ ─────────────────────────────────────
ch.push(h1('29. אירועי גלגל הבלת"מ'), note('מקור: moadilim-poc.html → const GC → unexpectedEvents — מוצג בחלון הגלגל אחרי סיבוב'));

const UE_SET_NAMES = {
  ue_set_aug:   'סט אוגוסט',
  ue_set_sep:   'סט ספטמבר',
  ue_set_oct:   'סט אוקטובר',
  ue_set_nov:   'סט נובמבר',
  ue_set_comp:  'סט דצמבר (מודל תגמול)',
  ue_set_comp2: 'סט תגמול 2',
};

Object.entries(GC.unexpectedEvents).forEach(([setId, events]) => {
  const name = UE_SET_NAMES[setId] || setId;
  ch.push(h2(name + ' (' + setId + ')'));
  ch.push(makeTable(events.flatMap((e, i) => [
    [setId + '[' + i + '].label',       e.label,                  'חלון גלגל הבלת"מ — תווית תוצאה על הגלגל'],
    [setId + '[' + i + '].description', e.description,            'חלון גלגל הבלת"מ — תיאור מה קרה (מוצג אחרי הסיבוב)'],
    [setId + '[' + i + '].waitRounds',  String(e.waitRounds),     'חלון גלגל הבלת"מ — מספר סבבי המתנה (0 = ממשיכים, 1 = מחכים סבב)'],
  ])));
  ch.push(gap());
});

// ════════════════════════════════════════════════════════

const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Arial', size: 22, color: '2a2a2a' } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 34, bold: true, font: 'Arial', color: NAVY },
        paragraph: { spacing: { before: 440, after: 140 }, outlineLevel: 0,
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: TEAL, space: 4 } } } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 24, bold: true, font: 'Arial', color: '444444' },
        paragraph: { spacing: { before: 260, after: 100 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }
      }
    },
    headers: {
      default: new Header({ children: [
        new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true,
          border: { bottom: { style: BorderStyle.SINGLE, size: 2, color: MGRAY, space: 2 } },
          children: [new TextRun({ text: 'מסילות ותכנונים — מסמך קופי', font: 'Arial', size: 18, color: '999999' })]
        })
      ]})
    },
    footers: {
      default: new Footer({ children: [
        new Paragraph({ alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 2, color: MGRAY, space: 2 } },
          children: [
            new TextRun({ text: 'עמוד ', font: 'Arial', size: 18, color: '999999' }),
            new TextRun({ children: [PageNumber.CURRENT], font: 'Arial', size: 18, color: '999999' }),
          ]
        })
      ]})
    },
    children: ch
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('copy_doc.docx', buf);
  console.log('Done: copy_doc.docx');
}).catch(e => { console.error(e); process.exit(1); });
