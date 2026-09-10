// קוד זה רץ בתוך Google Sheets (Extensions > Apps Script)
// הוא הופך את הגיליון ל-API פשוט שהלוח הדיגיטלי מדבר איתו

const ADMIN_PASSWORD = 'כפרהנשיא2026'; // אותה סיסמה שכבר בלוח - אפשר לשנות, רק לוודא שזהה בשני המקומות

function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === 'Notes') {
      sheet.appendRow(['id', 'text', 'want', 'support', 'oppose', 'ts', 'authorId']);
    } else if (name === 'Comments') {
      sheet.appendRow(['id', 'noteId', 'text', 'ts', 'authorId']);
    } else if (name === 'Votes') {
      sheet.appendRow(['noteId', 'voterId', 'type']);
    }
  }
  return sheet;
}

function doGet(e) {
  const action = e.parameter.action;
  let result;
  try {
    if (action === 'getData') {
      result = getData();
    } else if (action === 'addNote') {
      result = addNote(e.parameter.id, e.parameter.text, e.parameter.want, e.parameter.authorId);
    } else if (action === 'editNote') {
      result = editNote(e.parameter.id, e.parameter.text, e.parameter.want, e.parameter.authorId);
    } else if (action === 'vote') {
      result = vote(e.parameter.noteId, e.parameter.type, e.parameter.voterId);
    } else if (action === 'addComment') {
      result = addComment(e.parameter.noteId, e.parameter.commentId, e.parameter.text, e.parameter.authorId);
    } else if (action === 'editComment') {
      result = editComment(e.parameter.noteId, e.parameter.commentId, e.parameter.text, e.parameter.authorId);
    } else if (action === 'deleteNote') {
      result = deleteNote(e.parameter.noteId, e.parameter.pw, e.parameter.authorId);
    } else if (action === 'deleteComment') {
      result = deleteComment(e.parameter.noteId, e.parameter.commentId, e.parameter.pw, e.parameter.authorId);
    } else {
      result = { error: 'unknown action' };
    }
  } catch (err) {
    result = { error: err.toString() };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getData() {
  const notesData = getSheet('Notes').getDataRange().getValues();
  const commentsData = getSheet('Comments').getDataRange().getValues();

  const notes = [];
  for (let i = 1; i < notesData.length; i++) {
    const r = notesData[i];
    if (!r[0]) continue;
    notes.push({ id: r[0], text: r[1], want: r[2], support: Number(r[3]) || 0, oppose: Number(r[4]) || 0, ts: Number(r[5]) || 0, authorId: r[6] || '' });
  }
  const comments = [];
  for (let i = 1; i < commentsData.length; i++) {
    const r = commentsData[i];
    if (!r[0]) continue;
    comments.push({ id: r[0], noteId: r[1], text: r[2], ts: Number(r[3]) || 0, authorId: r[4] || '' });
  }
  return { notes, comments };
}

function addNote(id, text, want, authorId) {
  if (!text) return { error: 'missing text' };
  getSheet('Notes').appendRow([id, text, want || '', 0, 0, Date.now(), authorId || '']);
  return { ok: true };
}

// הצבעה חכמה: לחיצה ראשונה = מוסיפה קול, לחיצה שוב על אותו כפתור = מבטלת,
// לחיצה על הכפתור השני = מחליפה צד. הכל נשמר לפי noteId+voterId בגיליון Votes.
function editNote(id, text, want, authorId) {
  if (!text) return { error: 'missing text' };
  const sheet = getSheet('Notes');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      const ownerId = data[i][6] || '';
      if (!authorId || ownerId !== authorId) return { error: 'unauthorized' };
      sheet.getRange(i + 1, 2).setValue(text);
      sheet.getRange(i + 1, 3).setValue(want || '');
      return { ok: true };
    }
  }
  return { error: 'note not found' };
}

function vote(noteId, type, voterId) {
  if (!voterId) return { error: 'missing voterId' };

  const notesSheet = getSheet('Notes');
  const notesData = notesSheet.getDataRange().getValues();
  let noteRow = -1;
  for (let i = 1; i < notesData.length; i++) {
    if (notesData[i][0] === noteId) { noteRow = i; break; }
  }
  if (noteRow === -1) return { error: 'note not found' };

  const votesSheet = getSheet('Votes');
  const votesData = votesSheet.getDataRange().getValues();
  let voteRow = -1, existingType = null;
  for (let i = 1; i < votesData.length; i++) {
    if (votesData[i][0] === noteId && votesData[i][1] === voterId) { voteRow = i; existingType = votesData[i][2]; break; }
  }

  const supportCol = 4, opposeCol = 5;
  function adjust(col, delta) {
    const current = Number(notesSheet.getRange(noteRow + 1, col).getValue()) || 0;
    notesSheet.getRange(noteRow + 1, col).setValue(Math.max(0, current + delta));
  }

  let action;
  if (voteRow === -1) {
    votesSheet.appendRow([noteId, voterId, type]);
    adjust(type === 'support' ? supportCol : opposeCol, 1);
    action = 'added';
  } else if (existingType === type) {
    votesSheet.deleteRow(voteRow + 1);
    adjust(type === 'support' ? supportCol : opposeCol, -1);
    action = 'removed';
  } else {
    votesSheet.getRange(voteRow + 1, 3).setValue(type);
    adjust(existingType === 'support' ? supportCol : opposeCol, -1);
    adjust(type === 'support' ? supportCol : opposeCol, 1);
    action = 'switched';
  }
  return { ok: true, action: action };
}

function addComment(noteId, commentId, text, authorId) {
  if (!text) return { error: 'missing text' };
  getSheet('Comments').appendRow([commentId, noteId, text, Date.now(), authorId || '']);
  return { ok: true };
}

function editComment(noteId, commentId, text, authorId) {
  if (!text) return { error: 'missing text' };
  const sheet = getSheet('Comments');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === commentId && data[i][1] === noteId) {
      const ownerId = data[i][4] || '';
      if (!authorId || ownerId !== authorId) return { error: 'unauthorized' };
      sheet.getRange(i + 1, 3).setValue(text);
      return { ok: true };
    }
  }
  return { error: 'comment not found' };
}

function deleteNote(noteId, pw, authorId) {
  const sheet = getSheet('Notes');
  const data = sheet.getDataRange().getValues();
  let rowIndex = -1;
  let ownerId = '';
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === noteId) { rowIndex = i; ownerId = data[i][6] || ''; break; }
  }
  if (rowIndex === -1) return { error: 'note not found' };

  const isAdmin = pw === ADMIN_PASSWORD;
  const isOwner = authorId && ownerId && authorId === ownerId;
  if (!isAdmin && !isOwner) return { error: 'unauthorized' };

  sheet.deleteRow(rowIndex + 1);
  const cSheet = getSheet('Comments');
  const cData = cSheet.getDataRange().getValues();
  for (let i = cData.length - 1; i >= 1; i--) {
    if (cData[i][1] === noteId) cSheet.deleteRow(i + 1);
  }
  return { ok: true };
}

function deleteComment(noteId, commentId, pw, authorId) {
  const sheet = getSheet('Comments');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === commentId && data[i][1] === noteId) {
      const ownerId = data[i][4] || '';
      const isAdmin = pw === ADMIN_PASSWORD;
      const isOwner = authorId && ownerId && authorId === ownerId;
      if (!isAdmin && !isOwner) return { error: 'unauthorized' };
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { error: 'comment not found' };
}
