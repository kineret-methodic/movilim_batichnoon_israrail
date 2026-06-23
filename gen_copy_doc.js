const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
        AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
        Header, Footer, PageNumber } = require('docx');
const fs = require('fs');

// ── Read GC from HTML (single source of truth) ───────────
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

// ── Helpers ───────────────────────────────────────────────
const NAVY  = '221f64';
const TEAL  = '1abe9b';
const LGRAY = 'f5f6fa';
const MGRAY = 'e3e6ef';
const WHITE = 'ffffff';

// A4: 11906 DXA wide, 2cm margins → content 9638
const TW = 9638;
const C1 = 1800; // key
const C2 = 2600; // current text
const C3 = 2400; // where it appears
const C4 = 2838; // correction / note

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
      children: [new TextRun({ text: String(text || '—'), font: 'Arial', size: 20, color: color || '2a2a2a' })]
    })]
  });
}

function makeTable(rows) {
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: [C1, C2, C3, C4],
    rows: [
      new TableRow({
        tableHeader: true,
        children: [hCell('מפתח בקוד', C1), hCell('טקסט נוכחי', C2), hCell('היכן מופיע', C3), hCell('תיקון / הערה', C4)]
      }),
      ...rows.map(([key, val, where], i) => new TableRow({
        children: [
          dCell(key,   C1, i % 2),
          dCell(val,   C2, i % 2),
          dCell(where, C3, i % 2, '555577'),
          dCell('',    C4, i % 2),
        ]
      }))
    ]
  });
}

const h1   = t => new Paragraph({ heading: HeadingLevel.HEADING_1, bidirectional: true, alignment: AlignmentType.RIGHT,
  children: [new TextRun({ text: t, font: 'Arial', size: 34, bold: true, color: NAVY })] });
const h2   = t => new Paragraph({ heading: HeadingLevel.HEADING_2, bidirectional: true, alignment: AlignmentType.RIGHT,
  children: [new TextRun({ text: t, font: 'Arial', size: 24, bold: true, color: '555555' })] });
const note = t => new Paragraph({ bidirectional: true, alignment: AlignmentType.RIGHT, spacing: { after: 80 },
  children: [new TextRun({ text: t, font: 'Arial', size: 18, color: '888888', italics: true })] });
const gap  = () => new Paragraph({ spacing: { after: 180 }, children: [new TextRun('')] });

// ── Shortcuts into GC ─────────────────────────────────────
const copy = GC.uiCopy;
const btn  = copy.buttons;
const msg  = copy.messages;
const sb   = copy.statusBar;

// ════════════════════════════════════════════════════════════
const ch = [];

// ── Cover ────────────────────────────────────────────────
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
    children: [new TextRun({ text: 'מפתח בקוד — המיקום המדויק ב-gameContent.json או ב-HTML', font: 'Arial', size: 20, color: '444444' })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 40 },
    children: [new TextRun({ text: 'טקסט נוכחי — הניסוח הקיים במערכת (נקרא ישירות מקובץ ה-HTML)', font: 'Arial', size: 20, color: '444444' })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 40 },
    children: [new TextRun({ text: 'היכן מופיע — איפה בממשק הטקסט נראה', font: 'Arial', size: 20, color: '444444' })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT, bidirectional: true, spacing: { after: 400 },
    children: [new TextRun({ text: 'תיקון / הערה — מלאי כאן את השינויים הרצויים', font: 'Arial', size: 20, bold: true, color: TEAL })] }),
);

// ── 1. Meta ──────────────────────────────────────────────
ch.push(h1('1. כותרות ומטא'), note('מקור: GC.uiCopy'));
ch.push(makeTable([
  ['uiCopy.appTitle',     copy.appTitle,     'מסך פתיחה — כותרת ראשית גדולה'],
  ['uiCopy.pageTitle',    copy.pageTitle,    'כותרת ה-tab בדפדפן'],
  ['uiCopy.gameSubtitle', copy.gameSubtitle, 'מסך פתיחה — כיתוב אפור מתחת לכותרת'],
]));
ch.push(gap());

// ── 2. Welcome screen ────────────────────────────────────
ch.push(h1('2. מסך פתיחה'), note('מקור: GC.uiCopy.welcome'));
ch.push(makeTable([
  ['uiCopy.welcome.title',                  copy.welcome.title,                  'מסך פתיחה — כותרת כרטיס מרכזי'],
  ['uiCopy.welcome.intro',                  copy.welcome.intro,                  'מסך פתיחה — פסקת הסבר'],
  ['uiCopy.welcome.startButton',            copy.welcome.startButton || '(הוסר — הוחלף בכפתור X)', 'מסך פתיחה — כפתור כניסה ללוח'],
  ['uiCopy.welcome.openInstructionsButton', copy.welcome.openInstructionsButton, 'מסך פתיחה — כפתור משני (קו תחתון)'],
  ['uiCopy.welcome.closeInstructionsButton',copy.welcome.closeInstructionsButton,'חלונית הנחיות — כפתור סגירה'],
  ['hardcoded: station_start activatePhones', 'מתחילים',                          'תחנת כולם לעלות — כפתור התחלת המשחק'],
  ['hardcoded: station_start phonesActive',   '✓ המשחק החל — הקבוצות ננעלות לפי תור', 'תחנת כולם לעלות — הודעת אישור לאחר לחיצה'],
]));
ch.push(gap());

// ── 3. Game instructions ─────────────────────────────────
ch.push(h1('3. הוראות המשחק'), note('מקור: GC.uiCopy.welcome.sections[]'));
const secRows = (copy.welcome.sections || []).flatMap((s, i) => [
  ['welcome.sections[' + i + '].title', s.title, 'חלונית הנחיות — כותרת פסקה ' + (i + 1)],
  ['welcome.sections[' + i + '].body',  s.body,  'חלונית הנחיות — גוף פסקה ' + (i + 1)],
]);
ch.push(makeTable(secRows));
ch.push(gap());

// ── 4. Facilitator board header ──────────────────────────
ch.push(h1('4. לוח המנחה — כותרת ומצב משחק'), note('מקור: GC.uiCopy'));
ch.push(makeTable([
  ['uiCopy.turnIndicator',           copy.turnIndicator,    'לוח מנחה — שורת מצב עליונה (מעל המסלול)'],
  ['uiCopy.map.instruction',         copy.map.instruction,  'לוח מנחה — כיתוב קטן על המסלול'],
  ['uiCopy.statusBar.activeTurn',    sb.activeTurn,         'לוח מנחה — תג כחול בשורת הסטטוס'],
  ['uiCopy.statusBar.roundComplete', sb.roundComplete,      'לוח מנחה — הודעה בשורת הסטטוס אחרי סיום סבב'],
  ['uiCopy.statusBar.noStation',     sb.noStation,          'לוח מנחה — ערך ריק כשאין תחנה פעילה'],
  ['uiCopy.statusBar.autoAdvance',   sb.autoAdvance,        'לוח מנחה — תג על קבוצה שממתינה להתקדמות אוטומטית'],
  ['uiCopy.statusBar.position',      sb.position,           'לוח מנחה — תווית עמודת מיקום'],
  ['uiCopy.statusBar.errors',        sb.errors,             'לוח מנחה — תווית עמודת טעויות'],
  ['uiCopy.statusBar.nextTurn',      sb.nextTurn,           'לוח מנחה — תווית עמודת "תור הבא"'],
]));
ch.push(gap());

// ── 5. Score chips (hardcoded in HTML) ───────────────────
ch.push(h1('5. לוח תוצאות — כרטיסי קבוצות (תחתית לוח)'), note('מקור: hardcoded HTML (renderFacilitator → CHIP_DEF, bScoreHtml) + GC.uiCopy.statusBar + GC.uiCopy.points'));
ch.push(makeTable([
  ['hardcoded: CHIP_DEF.active',   'תור פעיל',               'כרטיס קבוצה — תג סטטוס: הקבוצה ששולטת עכשיו'],
  ['hardcoded: CHIP_DEF.played',   '✓ שיחקה',                'כרטיס קבוצה — תג: הקבוצה סיימה תורה בסבב זה'],
  ['hardcoded: CHIP_DEF.waiting',  '⏳ ממתינה',               'כרטיס קבוצה — תג: הקבוצה מחכה (בעקבות בלת"מ)'],
  ['hardcoded: CHIP_DEF.notyet',   'מחכה לתורה',             'כרטיס קבוצה — תג: הקבוצה עוד לא הגיעה לתורה'],
  ['uiCopy.statusBar.played',      sb.played,                 'כרטיס קבוצה — שורת סטטוס מתחת לתג'],
  ['uiCopy.statusBar.waiting',     sb.waiting,                'כרטיס קבוצה — שורת סטטוס: קבוצה שממתינה סבב'],
  ['uiCopy.statusBar.notYet',      sb.notYet,                 'כרטיס קבוצה — שורת סטטוס: עוד לא שיחקה'],
  ['uiCopy.points.label',          copy.points.label,         'כרטיס קבוצה — תווית מעל מספר הנקודות'],
  ['hardcoded: bScoreHtml — נק׳',  'נק׳',                     'כרטיס קבוצה — קיצור "נקודות" בכרטיס הקומפקטי'],
  ['hardcoded: bScoreHtml — 📍',   '📍',                      'כרטיס קבוצה — emoji לפני שם התחנה הנוכחית'],
  ['hardcoded: nameByOrder — UE',  'גלגל הבלת"מ',             'כרטיס קבוצה — שם התחנה כשהקבוצה בתחנת בלת"מ'],
]));
ch.push(gap());

// ── 6. Group panel ───────────────────────────────────────
ch.push(h1('6. פאנל קבוצה — תחנה פעילה'), note('מקור: GC.uiCopy.groupCard'));
const gc = copy.groupCard;
ch.push(makeTable([
  ['uiCopy.groupCard.waitingMessage',    gc.waitingMessage,    'חלון תחנה — הודעה אפורה כשהקבוצה ממתינה'],
  ['uiCopy.groupCard.skipButton',        gc.skipButton,        'חלון תחנה — כפתור לדילוג על קבוצה ממתינה'],
  ['uiCopy.groupCard.stationLocked',     gc.stationLocked,     'חלון תחנה — הודעה כשלוחצים על תחנה לא בתור'],
  ['uiCopy.groupCard.clickToOpen',       gc.clickToOpen,       'לוח מנחה — קישור מעל תחנה ניתנת ללחיצה'],
  ['uiCopy.groupCard.approveTransition', gc.approveTransition, 'חלון תחנת מעבר — כפתור אישור וסגירה'],
  ['uiCopy.groupCard.answerRecorded',    gc.answerRecorded,    'חלון תחנה — הודעה אחרי בחירת נכון/חלקי/שגוי'],
  ['uiCopy.groupCard.mustAnswerFirst',   gc.mustAnswerFirst,   'חלון תחנה — הודעת שגיאה אם לוחצים "חזרה" לפני שעונים'],
]));
ch.push(gap());

// ── 7. Buttons ───────────────────────────────────────────
ch.push(h1('7. כפתורים'), note('מקור: GC.uiCopy.buttons, GC.uiCopy.timer'));
ch.push(makeTable([
  ['uiCopy.buttons.startGame',          btn.startGame,          'מסך פתיחה — כפתור ראשי (כחול גדול)'],
  ['uiCopy.buttons.startTimer',         btn.startTimer,         'חלון תחנה מקצועית — כפתור הפעלת טיימר'],
  ['uiCopy.buttons.correct',            btn.correct,            'חלון תחנה — כפתור ירוק (אחרי חשיפת תשובה)'],
  ['uiCopy.buttons.partial',            btn.partial,            'חלון תחנה — כפתור כתום'],
  ['uiCopy.buttons.wrong',              btn.wrong,              'חלון תחנה — כפתור אדום'],
  ['uiCopy.buttons.revealAnswer',       btn.revealAnswer,       'חלון תחנה — כפתור שחושף את התשובה הנכונה'],
  ['uiCopy.buttons.openMobileFolder',   btn.openMobileFolder,   'חלון תחנה — כפתור שפותח תיק מובייל'],
  ['uiCopy.buttons.spinWheel',          btn.spinWheel,          'חלון בלת"מ — כפתור הפעלת הגלגל'],
  ['uiCopy.buttons.backToMap',          btn.backToMap,          'חלון תחנה — כפתור חזרה ללוח'],
  ['uiCopy.buttons.confirmSoftWarning', btn.confirmSoftWarning, 'חלונית אזהרה — כפתור אישור המשך'],
  ['uiCopy.buttons.cancel',             btn.cancel,             'חלונית אזהרה — כפתור ביטול'],
  ['uiCopy.timer.stop',                 copy.timer.stop,        'טיימר פעיל — כפתור עצירה'],
  ['uiCopy.timer.resume',               copy.timer.resume,      'טיימר מושהה — כפתור המשך'],
]));
ch.push(gap());

// ── 8. System messages ───────────────────────────────────
ch.push(h1('8. הודעות מערכת'), note('מקור: GC.uiCopy.messages'));
ch.push(makeTable([
  ['uiCopy.messages.mobileIntro',           msg.mobileIntro,           'תיק מובייל — פסקת הסבר בראש הדף'],
  ['uiCopy.messages.softWarning',           msg.softWarning,           'חלונית אזהרה — גוף הודעה (כשפותחים תחנה לפני הזמן)'],
  ['uiCopy.messages.stationCompleted',      msg.stationCompleted,      'חלון תחנה — הודעה כשהתחנה כבר טופלה'],
  ['uiCopy.messages.unexpectedAlreadyUsed', msg.unexpectedAlreadyUsed, 'חלון בלת"מ — הודעה כשניסו להפעיל בלת"מ שכבר השתמשו בו'],
  ['uiCopy.messages.correctOutcome',        msg.correctOutcome,        'לוח מנחה — באנר תוצאה ירוק אחרי "נכון"'],
  ['uiCopy.messages.partialOutcome',        msg.partialOutcome,        'לוח מנחה — באנר תוצאה כתום אחרי "חלקי"'],
  ['uiCopy.messages.wrongOutcome',          msg.wrongOutcome,          'לוח מנחה — באנר תוצאה אדום אחרי "שגוי"'],
  ['uiCopy.messages.wrongOutcomeStayed',    msg.wrongOutcomeStayed,    'לוח מנחה — וריאנט הודעת שגיאה (תחנה ראשונה)'],
  ['uiCopy.messages.autoAdvanceNotice',     msg.autoAdvanceNotice,     'לוח מנחה — הודעת מידע על קבוצה שתתקדם אוטומטית'],
  ['uiCopy.messages.mobileTasksFallback',   msg.mobileTasksFallback,   'תיק מובייל — הודעה כשאין תוכן משימות עדיין'],
  ['uiCopy.messages.mobileNoCompensation',  msg.mobileNoCompensation,  'תיק מובייל — הודעה בלשונית "מודל תגמול" כשהיא ריקה'],
  ['uiCopy.messages.noGroupAnswer',         msg.noGroupAnswer,         'חלון תחנה — הודעה כשאין תשובה ספציפית לקבוצה'],
  ['uiCopy.messages.noStationQuestion',     msg.noStationQuestion,     'חלון תחנה — הודעה בתחנות מעבר שאין בהן שאלה'],
]));
ch.push(gap());

// ── 9. Mobile folder ─────────────────────────────────────
ch.push(h1('9. תצוגת מובייל — תיק הקבוצה'), note('מקור: GC.uiCopy.mobileFolder, GC.uiCopy.tabs, GC.uiCopy.reminderDrawer'));
ch.push(makeTable([
  ['uiCopy.mobileFolder.lockedMessage', copy.mobileFolder.lockedMessage, 'תיק מובייל — כותרת מסך נעילה (גדולה, בעברית)'],
  ['uiCopy.mobileFolder.lockedWait',    copy.mobileFolder.lockedWait,    'תיק מובייל — כיתוב מתחת לכותרת הנעילה'],
  ['uiCopy.tabs.reviewTasks',           copy.tabs.reviewTasks,           'תיק מובייל — לשונית ראשונה (ברירת מחדל)'],
  ['uiCopy.tabs.compensationModel',     copy.tabs.compensationModel,     'תיק מובייל — לשונית שנייה'],
  ['hardcoded: openRef button',         '📖 חומרי עזר · צריכים תזכורת?', 'תיק מובייל — כפתור בתחתית שפותח מסך חומרי עזר (hardcoded)'],
  ['uiCopy.buttons.needReminder',       btn.needReminder,                'ערך GC — כרגע לא בשימוש ישיר (הכפתור hardcoded למעלה)'],
  ['uiCopy.buttons.openFolder',         btn.openFolder,                  'תיק מובייל — כפתור קישור לתכנית העבודה החטיבתית'],
  ['uiCopy.reminderDrawer.title',       copy.reminderDrawer.title,       'מגירת תזכורת — כותרת החלק העליון'],
  ['uiCopy.completedBadge',             copy.completedBadge,             'תיק מובייל — תג ירוק על משימה שסומנה כהושלמה'],
]));
ch.push(gap());

// ── 9.5 Reference screen ─────────────────────────────────
ch.push(h1('9.5. מסך חומרי עזר — renderReferenceScreen'), note('מקור: hardcoded HTML (תוויות ממשק) + GC.companyGoals + GC.planningGlossary (תוכן)'));

ch.push(h2('תוויות ממשק (hardcoded ב-HTML)'));
ch.push(makeTable([
  ['hardcoded: ref screen title',      'חומרי עזר',                     'כותרת ראשית בפס העליון הצבעוני'],
  ['hardcoded: ref back button',       '→ חזרה לתיק העבודה',            'כפתור חזרה בפס העליון'],
  ['hardcoded: ref tab 1',             'מטרות ויעדי החברה',              'לשונית ראשונה בפס הטאבים'],
  ['hardcoded: ref tab 2',             'מילון מושגים',                   'לשונית שנייה בפס הטאבים'],
  ['hardcoded: ref goals intro',       'מטרות החברה ויעדיה לשנת 2026. הקישו על מטרה כדי לפתוח את היעדים שתחתיה.', 'לשונית מטרות — פסקת הסבר בראש'],
  ['hardcoded: ref glossary intro',    'מושגים בשפת התכנון — הגדרה, דוגמה וגורם אחראי.',                         'לשונית מילון — פסקת הסבר בראש'],
  ['hardcoded: ref example label',     'דוגמה',                          'מילון — תווית כרטיס דוגמה (💡)'],
  ['hardcoded: ref responsible label', 'גורם אחראי:',                    'מילון — תווית שם הגורם האחראי'],
]));
ch.push(gap());

ch.push(h2('מטרות החברה ויעדיה (GC.companyGoals)'));
const goalRows = (GC.companyGoals || []).flatMap(g => [
  ['companyGoals[' + g.n + '].title', g.title, 'מסך חומרי עזר — שם המטרה (כפתור מתקפל, order: ' + g.n + ')'],
  ...(g.objectives || []).map(o => [
    'companyGoals[' + g.n + '].objectives[' + o.code + ']',
    o.code + ' — ' + o.text,
    'מסך חומרי עזר — יעד תחת מטרה ' + g.n,
  ]),
]);
ch.push(makeTable(goalRows));
ch.push(gap());

ch.push(h2('מילון מושגי תכנון (GC.planningGlossary)'));
const glossRows = (GC.planningGlossary || []).flatMap((d, i) => [
  ['planningGlossary[' + i + '].term',        d.term,        'מסך חומרי עזר — שם המושג (בולט, בצבע)'],
  ['planningGlossary[' + i + '].definition',  d.definition,  'מסך חומרי עזר — הגדרה'],
  ...(d.example     ? [['planningGlossary[' + i + '].example',     d.example,     'מסך חומרי עזר — תוכן כרטיס דוגמה (💡)']] : []),
  ...(d.responsible ? [['planningGlossary[' + i + '].responsible', d.responsible, 'מסך חומרי עזר — שם הגורם האחראי']]        : []),
]);
ch.push(makeTable(glossRows));
ch.push(gap());

// ── 10. Form fields ──────────────────────────────────────
ch.push(h1('10. שדות טופס ותוויות'), note('מקור: GC.uiCopy.taskSections, taskFields, compensationFields, answerCard'));
const tf = copy.taskFields;
const cf = copy.compensationFields;
const ac = copy.answerCard;
ch.push(makeTable([
  ['uiCopy.taskSections.strategic',                  copy.taskSections.strategic,  'תיק מובייל — כותרת קטע "שיוך אסטרטגי"'],
  ['uiCopy.taskSections.details',                    copy.taskSections.details,    'תיק מובייל — כותרת קטע "פרטי המשימה"'],
  ['uiCopy.taskSections.planning',                   copy.taskSections.planning,   'תיק מובייל — כותרת קטע "תכנון לאורך השנה"'],
  ['uiCopy.taskFields.companyGoal',                  tf.companyGoal,               'תיק מובייל — תווית שדה "מטרת חברה"'],
  ['uiCopy.taskFields.goal',                         tf.goal,                      'תיק מובייל — תווית שדה "יעד"'],
  ['uiCopy.taskFields.outcomeMetric',                tf.outcomeMetric,             'תיק מובייל — תווית שדה "מדד תוצאה"'],
  ['uiCopy.taskFields.classification',               tf.classification,            'תיק מובייל — תווית שדה "סיווג"'],
  ['uiCopy.taskFields.taskDescription',              tf.taskDescription,           'תיק מובייל — תווית שדה "תיאור משימה"'],
  ['uiCopy.taskFields.annualAchievement',            tf.annualAchievement,         'תיק מובייל — תווית שדה "הישג שנתי נדרש"'],
  ['uiCopy.taskFields.outputMetricType',             tf.outputMetricType,          'תיק מובייל — תווית שדה "סוג מדד תפוקה"'],
  ['uiCopy.classificationLabels.top10',              copy.classificationLabels.top10,   'תיק מובייל — ערך בשדה "סיווג" (Top 10)'],
  ['uiCopy.classificationLabels.routine',            copy.classificationLabels.routine, 'תיק מובייל — ערך בשדה "סיווג" (שוטפת)'],
  ['uiCopy.compensationFields.rewardMetric',         cf.rewardMetric,              'תיק מובייל — תווית "מדד תגמול"'],
  ['uiCopy.compensationFields.formula',              cf.formula,                   'תיק מובייל — תווית "נוסחה"'],
  ['uiCopy.compensationFields.weight',               cf.weight,                    'תיק מובייל — תווית "משקל"'],
  ['uiCopy.compensationFields.performanceThresholds',cf.performanceThresholds,     'תיק מובייל — תווית "ספי ביצוע"'],
  ['uiCopy.compensationFields.dataSource',           cf.dataSource,                'תיק מובייל — תווית "מקור נתונים"'],
  ['uiCopy.answerCard.taskLabel',                    ac.taskLabel,                 'חלון תחנה — תווית "משימה:" בכרטיס התשובה'],
  ['uiCopy.answerCard.whatIsWrongLabel',             ac.whatIsWrongLabel,          'חלון תחנה — תווית "מה לא תקין:"'],
  ['uiCopy.answerCard.correctFixLabel',              ac.correctFixLabel,           'חלון תחנה — תווית "תיקון נכון:"'],
  ['uiCopy.answerCard.facilitatorNoteLabel',         ac.facilitatorNoteLabel,      'חלון תחנה — תווית "הערת מנחה:"'],
]));
ch.push(gap());

// ── 11. Stations ─────────────────────────────────────────
ch.push(h1('11. תחנות — שמות וחודשים'), note('מקור: GC.stations (ממוין לפי order) + GC.mapGates'));
const sortedSt = [...GC.stations].sort((a, b) => a.order - b.order);
const stRows = sortedSt.flatMap(s => [
  ['stations[' + s.id + '].name  (order:' + s.order + ')', s.name,        'לוח מנחה — שם תחנה על המסלול'],
  ['stations[' + s.id + '].month',                          s.month || '—', 'לוח מנחה — חודש על המסלול'],
]);
(GC.mapGates || []).forEach(g => {
  stRows.push(['mapGates[' + g.id + '].name  (afterOrder:' + g.afterOrder + ')', g.name, 'לוח מנחה — שם שער (gate) על המסלול']);
});
ch.push(makeTable(stRows));
ch.push(gap());

// ── 12. Station questions ────────────────────────────────
ch.push(h1('12. שאלות המשחק'), note('מקור: GC.stationQuestions'));
(GC.stationQuestions || []).forEach(q => {
  const s = GC.stations.find(x => x.id === q.stationId);
  ch.push(h2('תחנת ' + (s ? s.name : q.stationId)));
  ch.push(makeTable([
    ['stationQuestions[' + q.id + '].questionText',          q.questionText,          'חלון תחנה — שאלה שמוצגת למנחה'],
    ['stationQuestions[' + q.id + '].professionalPrinciple', q.professionalPrinciple, 'חלון תחנה — "תשובה נכונה" שמתגלה אחרי "חשפו תשובה"'],
  ]));
  ch.push(gap());
});

// ── 13. Transition messages ──────────────────────────────
ch.push(h1('13. מסכי מעבר (תחנות transition)'), note('מקור: GC.transitionMessages'));
const tmRows = Object.entries(GC.transitionMessages || {}).flatMap(([key, tm]) => [
  ['transitionMessages.' + key + '.title', tm.title, 'חלון תחנת מעבר — כותרת'],
  ['transitionMessages.' + key + '.body',  tm.body,  'חלון תחנת מעבר — גוף טקסט'],
]);
ch.push(makeTable(tmRows));
ch.push(gap());

// ── 14. End screen ───────────────────────────────────────
ch.push(h1('14. מסך סיום'), note('מקור: GC.uiCopy.endScreen'));
const es = copy.endScreen;
ch.push(makeTable([
  ['uiCopy.endScreen.title',              es.title,              'מסך סיום — כותרת ראשית'],
  ['uiCopy.endScreen.winnerLabel',        es.winnerLabel,        'מסך סיום — תווית מעל שם הקבוצה הזוכה'],
  ['uiCopy.endScreen.commitmentQuestion', es.commitmentQuestion, 'מסך סיום — שאלת ההתחייבות האישית (גדולה, מרכז המסך)'],
  ['uiCopy.endScreen.mentimeterUrl',      es.mentimeterUrl || '—', 'מסך סיום — כתובת ה-QR/Mentimeter'],
  ['uiCopy.endScreen.mentimeterQrLabel',  es.mentimeterQrLabel,  'מסך סיום — כיתוב מתחת ל-QR Code'],
]));
ch.push(gap());

// ── 15. Reminder definitions ─────────────────────────────
ch.push(h1('15. מילון מושגים — מגירת התזכורת'), note('מקור: GC.reminderDefinitions[]'));
const rdRows = (GC.reminderDefinitions || []).flatMap((d, i) => [
  ['reminderDefinitions[' + i + '].term',       d.term,       'מגירת תזכורת — שם המושג (מודגש)'],
  ['reminderDefinitions[' + i + '].definition', d.definition, 'מגירת תזכורת — הגדרה מתחת למושג'],
]);
ch.push(makeTable(rdRows));
ch.push(gap());

// ── 16. Groups ───────────────────────────────────────────
ch.push(h1('16. קבוצות'), note('מקור: GC.groups[]'));
const grpRows = (GC.groups || []).flatMap(g => [
  ['groups[' + g.id + '].label',        g.label,        'לוח מנחה + תיק מובייל — שם הקבוצה'],
  ['groups[' + g.id + '].divisionName', g.divisionName, 'לוח מנחה + תיק מובייל — שם החטיבה'],
]);
ch.push(makeTable(grpRows));
ch.push(gap());

// ── 17. Months ───────────────────────────────────────────
ch.push(h1('17. חודשים'), note('מקור: GC.uiCopy.months[]'));
const mthRows = (copy.months || []).map((m, i) => [
  'uiCopy.months[' + i + ']', m, 'לוח מנחה — תווית חודש על ציר הזמן',
]);
ch.push(makeTable(mthRows));
ch.push(gap());

// ── 18. Misc microcopy ───────────────────────────────────
ch.push(h1('18. מיקרו-קופי שונות'), note('מקור: GC.uiCopy'));
const ob = copy.outcomeBanner || {};
ch.push(makeTable([
  ['uiCopy.stationTypes.unexpectedEventLabel', copy.stationTypes.unexpectedEventLabel,     'תג על תחנות בלת"מ במסלול'],
  ['uiCopy.points.correct',                    copy.points.correct,                        'הודעת תוצאה — ניקוד תשובה נכונה'],
  ['uiCopy.points.partial',                    copy.points.partial,                        'הודעת תוצאה — ניקוד תשובה חלקית'],
  ['uiCopy.points.wrong',                      copy.points.wrong,                          'הודעת תוצאה — ניקוד תשובה שגויה'],
  ['uiCopy.outcomeBanner.correct',             ob.correct  || '{group}: {msg}',            'לוח מנחה — תבנית באנר תוצאה (נכון)'],
  ['uiCopy.outcomeBanner.partial',             ob.partial  || '{group}: {msg}',            'לוח מנחה — תבנית באנר תוצאה (חלקי)'],
  ['uiCopy.outcomeBanner.wrong',               ob.wrong    || '{group}: {msg}',            'לוח מנחה — תבנית באנר תוצאה (שגוי)'],
  ['uiCopy.boardSections.a',                   copy.boardSections.a,                       'לוח מנחה — כותרת מקטע ראשון על ציר הזמן'],
  ['uiCopy.boardSections.aMonths',             copy.boardSections.aMonths,                 'לוח מנחה — תווית חודשים למקטע א'],
  ['uiCopy.boardSections.b',                   copy.boardSections.b,                       'לוח מנחה — כותרת מקטע שני'],
  ['uiCopy.boardSections.bMonths',             copy.boardSections.bMonths,                 'לוח מנחה — תווית חודשים למקטע ב'],
  ['uiCopy.boardSections.c',                   copy.boardSections.c,                       'לוח מנחה — כותרת מקטע שלישי'],
  ['uiCopy.boardSections.cMonths',             copy.boardSections.cMonths,                 'לוח מנחה — תווית חודשים למקטע ג'],
]));
ch.push(gap());

// ── 19. UE modal (hardcoded) ─────────────────────────────
ch.push(h1('19. חלון תחנה — גלגל הבלת"מ'), note('מקור: hardcoded HTML (renderStationModal → unexpected_event branch)'));
ch.push(makeTable([
  ['hardcoded: UE title',               'גלגל הבלת"מ',                     'חלון בלת"מ — כותרת ראשית'],
  ['hardcoded: UE result header',       '⚡ תוצאת הגלגל',                  'חלון בלת"מ — כותרת קטע לאחר סיבוב'],
  ['hardcoded: UE no-delay badge',      '✓ ממשיכים קדימה — אין עיכוב',    'חלון בלת"מ — תג ירוק כשתוצאה: אין עיכוב'],
  ['hardcoded: UE delay badge',         '⏸ הקבוצה ממתינה סבב אחד',        'חלון בלת"מ — תג כתום כשתוצאה: עיכוב'],
  ['hardcoded: UE confirm button',      'הבנו, ממשיכים ←',                 'חלון בלת"מ — כפתור אישור וסגירה (בלת"מ חדש)'],
  ['hardcoded: UE already-used button', 'ממשיכים קדימה ←',                 'חלון בלת"מ — כפתור כשהבלת"מ כבר שומש'],
  ['hardcoded: close tooltip',          'סגירה ללא מעבר תור',              'חלון בלת"מ — tooltip על כפתור ה-X'],
]));
ch.push(gap());

// ── 20. Outcome popup (hardcoded) ────────────────────────
ch.push(h1('20. פופ-אפ תוצאה (אחרי נכון / חלקי / שגוי)'), note('מקור: hardcoded HTML (renderStationModal → outcomePopup)'));
ch.push(makeTable([
  ['hardcoded: outcome icon — correct',   '✅',                                        'פופ-אפ תוצאה — אייקון גדול (נכון)'],
  ['hardcoded: outcome icon — partial',   '⚠️',                                        'פופ-אפ תוצאה — אייקון (חלקי)'],
  ['hardcoded: outcome icon — wrong',     '❌',                                        'פופ-אפ תוצאה — אייקון (שגוי)'],
  ['hardcoded: outcome title — correct',  'תשובה נכונה!',                             'פופ-אפ תוצאה — כותרת ירוקה גדולה'],
  ['hardcoded: outcome title — partial',  'תשובה חלקית',                              'פופ-אפ תוצאה — כותרת כתומה'],
  ['hardcoded: outcome title — wrong',    'תשובה שגויה',                              'פופ-אפ תוצאה — כותרת אדומה'],
  ['hardcoded: outcome points — correct', '+10',                                       'פופ-אפ תוצאה — מספר הנקודות (גדול)'],
  ['hardcoded: outcome points — partial', '+5',                                        'פופ-אפ תוצאה — מספר הנקודות'],
  ['hardcoded: outcome points — wrong',   '+0',                                        'פופ-אפ תוצאה — מספר הנקודות'],
  ['hardcoded: outcome points suffix',    'נקודות',                                    'פופ-אפ תוצאה — טקסט ליד מספר הנקודות'],
  ['hardcoded: outcome next — wrong',     'מתקדמים לתחנה הבאה — ממתינים סבב אחד ⏳', 'פופ-אפ תוצאה — הודעת השלכה (שגוי)'],
  ['hardcoded: outcome next — correct',   'מתקדמים לתחנה הבאה 🚀',                   'פופ-אפ תוצאה — הודעת השלכה (נכון/חלקי)'],
  ['hardcoded: outcome dismiss button',   'הבנו, ממשיכים ←',                          'פופ-אפ תוצאה — כפתור סגירה'],
]));
ch.push(gap());

// ── 21. Mobile back button (hardcoded) ───────────────────
ch.push(h1('21. תצוגת מובייל — כפתור חזרה ללוח'), note('מקור: hardcoded HTML (renderMobile → floating backBtn) — מוצג רק כשהמנחה פותחת תיק מתוך הלוח'));
ch.push(makeTable([
  ['hardcoded: mobile backBtn emoji', '🏠', 'תיק מובייל (כשנפתח ע"י המנחה) — כפתור צף בצד ימין לחזרה ללוח'],
]));
ch.push(gap());

// ── 22. Dev content view (hardcoded) ─────────────────────
ch.push(h1('22. תצוגת תוכן פיתוח (כפתור ☰)'), note('מקור: hardcoded HTML (renderDevContentView) — נגישה דרך כפתור ☰'));
ch.push(makeTable([
  ['hardcoded: devContent header',               'סקירת תוכן — כל התחנות',            'תצוגת ☰ — כותרת ראשית'],
  ['hardcoded: devContent station count suffix', 'תחנות',                              'תצוגת ☰ — ספירת תחנות'],
  ['hardcoded: devContent — no transition msg',  'תחנת מעבר',                          'תצוגת ☰ — כשתחנת מעבר ללא תוכן מיוחד'],
  ['hardcoded: devContent — no UE items',        'אין פריטים',                         'תצוגת ☰ — כשתחנת בלת"מ ריקה'],
  ['hardcoded: devContent — no question',        'אין שאלה',                           'תצוגת ☰ — כשתחנה מקצועית ללא שאלה'],
  ['hardcoded: devContent type badge',           'מקצועית / בלת"מ / מעבר / רפלקציה', 'תצוגת ☰ — תג סוג תחנה'],
]));
ch.push(gap());

// ── 23. DEV MODE only (hardcoded) ────────────────────────
ch.push(h1('23. מיקרו-קופי DEV MODE בלבד'), note('מקור: hardcoded HTML — מופיע רק עם ?devmode=true. לא נראה למשתתפים.'));
ch.push(makeTable([
  ['hardcoded: nav-content button',               '📋 תוכן',                                         'נאב-בר — כפתור ניווט (מופיע רק ב-DEV mode)'],
  ['hardcoded: DEV UE results header',           '⚙ כל התוצאות האפשריות (DEV MODE)',               'חלון בלת"מ — כותרת קטע בחירת תוצאה ידנית'],
  ['hardcoded: DEV UE select button',            '⚡ בחר',                                          'חלון בלת"מ — כפתור בחירת תוצאה ספציפית'],
  ['hardcoded: DEV task error marker',           '⚠️ טעות מוטמנת',                                  'תיק מובייל — תג על משימה שיש בה טעות מוטמנת'],
  ['hardcoded: DEV station hint prefix',         '📍 תחנה:',                                        'תיק מובייל — תווית הינט לפיתוח'],
  ['hardcoded: renderDevView title',             '📋 סקירת תוכן — מצב פיתוח',                       'תצוגת DEV — כותרת ראשית'],
  ['hardcoded: renderDevView incomplete badge',  '⚠ להשלמה',                                        'תצוגת DEV — תג על תחנה עם שדות חסרים'],
  ['hardcoded: renderDevView complete badge',    '✓ הכל מלא',                                       'תצוגת DEV — תג על תחנה שלמה'],
  ['hardcoded: renderDevView section headers',   'תחנות מקצועיות / ⚡ תחנות בלת"מ / מודל תגמול', 'תצוגת DEV — כותרות קטגוריות'],
]));
ch.push(gap());

// ── 24–27. Mobile folder content per group (dynamic) ─────
function taskRows(tid, t) {
  const cls = t.classification === 'top10' ? 'Top 10' : 'שוטפת';
  return [
    [tid + '.taskName',          t.taskName,          'תיק מובייל — כותרת כרטיס המשימה'],
    [tid + '.companyGoal',       t.companyGoal,       'תיק מובייל — שדה "מטרת חברה"'],
    [tid + '.goal',              t.goal,              'תיק מובייל — שדה "יעד"'],
    [tid + '.outcomeMetric',     t.outcomeMetric,     'תיק מובייל — שדה "מדד תוצאה"'],
    [tid + '.classification',    cls,                 'תיק מובייל — שדה "סיווג"'],
    [tid + '.taskDescription',   t.taskDescription,   'תיק מובייל — שדה "תיאור משימה"'],
    [tid + '.annualAchievement', t.annualAchievement, 'תיק מובייל — שדה "הישג שנתי נדרש"'],
    [tid + '.outputMetricType',  t.outputMetricType,  'תיק מובייל — שדה "סוג מדד תפוקה"'],
    [tid + '.milestones.Q1',     t.milestones ? t.milestones.Q1 : '—', 'תיק מובייל — אבן דרך Q1'],
    [tid + '.milestones.Q2',     t.milestones ? t.milestones.Q2 : '—', 'תיק מובייל — אבן דרך Q2'],
    [tid + '.milestones.Q3',     t.milestones ? t.milestones.Q3 : '—', 'תיק מובייל — אבן דרך Q3'],
    [tid + '.milestones.Q4',     t.milestones ? t.milestones.Q4 : '—', 'תיק מובייל — אבן דרך Q4'],
    [tid + '.devHint',           t.devHint || '—',    'DEV MODE — רמז "⚠️ טעות מוטמנת"'],
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
    [gid + '.comp.notes',                m.notes || '—',         'תיק מובייל — הערת DEV'],
    [gid + '.comp.devHint',              m.devHint || '—',       'DEV MODE — רמז "⚠️ טעות מוטמנת" במודל תגמול'],
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
  comp:  GC.compensationModels[gid],
}));

GROUPS_CONTENT.forEach((g, gi) => {
  const sNum = 24 + gi;
  ch.push(h1(`${sNum}. תוכן תיק מובייל — ${g.label}`), note(`מקור: GC.visibleTasks.${g.gid} + GC.compensationModels.${g.gid}`));
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

// ── 28. Answers per group/station (dynamic) ───────────────
ch.push(h1('28. תשובות לפי קבוצה ותחנה'), note('מקור: GC.stationQuestions[].answersByGroup'));
(GC.stationQuestions || []).forEach(q => {
  const s = GC.stations.find(x => x.id === q.stationId);
  ch.push(h2('תחנת ' + (s ? s.name : q.stationId)));
  const rows = (GC.groups || []).flatMap(group => {
    const ans = q.answersByGroup[group.id];
    if (!ans) return [];
    const prefix = q.id + '.answersByGroup.' + group.id;
    return [
      [prefix + '.taskName',    ans.taskName,    'כרטיס תשובה — ' + group.label + ' — שדה "משימה:"'],
      [prefix + '.whatIsWrong', ans.whatIsWrong, 'כרטיס תשובה — ' + group.label + ' — שדה "מה לא תקין:"'],
      [prefix + '.correctFix',  ans.correctFix,  'כרטיס תשובה — ' + group.label + ' — שדה "תיקון נכון:"'],
    ];
  });
  if (rows.length) ch.push(makeTable(rows));
  ch.push(gap());
});

// ── 29. Unexpected events (dynamic) ──────────────────────
ch.push(h1('29. אירועי גלגל הבלת"מ'), note('מקור: GC.unexpectedEvents'));

const UE_SET_NAMES = {
  ue_set_aug:  'סט אוגוסט',
  ue_set_sep:  'סט ספטמבר',
  ue_set_oct:  'סט אוקטובר',
  ue_set_nov:  'סט נובמבר',
  ue_set_comp: 'סט נובמבר (מודל תגמול)',
};

Object.entries(GC.unexpectedEvents).forEach(([setId, events]) => {
  const name = UE_SET_NAMES[setId] || setId;
  ch.push(h2(name + ' (' + setId + ')'));
  ch.push(makeTable(events.flatMap((e, i) => [
    [setId + '[' + i + '].label',       e.label,              'חלון גלגל הבלת"מ — תווית תוצאה על הגלגל'],
    [setId + '[' + i + '].description', e.description,        'חלון גלגל הבלת"מ — תיאור מה קרה (אחרי הסיבוב)'],
    [setId + '[' + i + '].waitRounds',  String(e.waitRounds), 'חלון גלגל הבלת"מ — סבבי המתנה (0=ממשיכים, 1=מחכים)'],
  ])));
  ch.push(gap());
});

// ════════════════════════════════════════════════════════════
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
