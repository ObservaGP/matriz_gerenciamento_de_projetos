/**
 * * MATRIZ DE GERENCIAMENTO DE PROJETOS COM SINCRONIZAÇÃO BIDIRECIONAL COM AGENDA ONLINE - VERSÃO 8.9 - 14/12/2025
 * Versão 8.7 — 16/09/2025
 *
 * ------------------------------------------------------------------------
 * LICENÇA E REGISTRO
 * ------------------------------------------------------------------------
 * Este código está licenciado sob a Creative Commons Atribuição 4.0 Internacional (CC BY 4.0).
 *
 * Você pode copiar, redistribuir, remixar, transformar e criar a partir deste material,
 * para qualquer finalidade — inclusive comercial — desde que atribua o devido crédito
 * aos autores originais:
 *
 * AUTORES:
 * - Victor Gianordoli (Instituto Federal do Espírito Santo - IFES)
 *   ORCID: https://orcid.org/0000-0001-5905-0641
 *
 * - Taciana de Lemos Dias (Universidade Federal do Espírito Santo - UFES)
 *   ORCID: https://orcid.org/0000-0002-7172-1230
 *
 * É obrigatório citar os autores em qualquer reprodução, modificação ou derivação do código.
 *
 * Referência da licença completa:
 * https://creativecommons.org/licenses/by/4.0/
 *
 * Projeto registrado no INPI conforme Classificação:
 * - Campo de Aplicação: Administração — Código: AD04-Adm Publ
 * - Tipo de Programa: Planilhas Eletrônicas — Código: FA03-Planil Elet
 *
 * © 2025 - ObservaGP | Grupo de Pesquisa do Observatório de Gestão Pública
 * ------------------------------------------------------------------------
 *
 * Linguagem: JavaScript / Google Apps Script
 * Instituições: IFES (Instituto Federal do Espírito Santo) e UFES (Universidade Federal do Espírito Santo)
 *
 * Descrição:
 * Sistema para integração e sincronização bidirecional entre Google Planilhas e Google Agenda.
 * Permite controle de projetos, acompanhamento de prazos e visualização Gantt automatizada.
 */


// ======================
// CONSTANTES E CONFIGURAÇÕES
// ======================

const SHEET_EVENTS       = 'AGENDA';
const SHEET_ARCHIVE      = 'Arquivo';
const SHEET_ID           = 'Id';

const BEGIN_DATE         = new Date(2020, 0, 1);
const END_DATE           = new Date(2030, 0, 1);

const THROTTLE_SLEEP_MS  = 100;  // Ajustável conforme necessidade
const NO_SYNC_STRING     = "NOSYNC";
// Duração padrão (minutos) quando o início tem hora e o fim está vazio
const DEFAULT_TIMED_DURATION_MIN = 60;

// Limite de criação para evitar rate limit do CalendarApp
const CREATE_BATCH_SIZE = 125;
const CREATE_BATCH_SLEEP_MS = 15000;


// Coluna de sincronização (A=1, B=2, C=3 são ignoradas; usamos a partir de D=4)
const SYNC_START_COL     = 4;

const TITLE_ROW_MAP = {
  'title':       'Título',
  'description': 'Descrição',
  'location':    'Local',
  'starttime':   'Data e Hora de Início',
  'endtime':     'Data e Hora de Fim',
  'guests':      'Convidados (e-mail, e-mail, ...)',
  'color':       'Cor',
  'id':          'ID',
  'reg':         'Modificado na Agenda',
  'mod':         'Registro na Agenda'
};

// ======================
// FUNÇÕES AUXILIARES
// ======================

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet(name) {
  return getSpreadsheet().getSheetByName(name);
}

function getCalendarIdFromSheet() {
  const sheet = getSheet(SHEET_ID);
  if (!sheet) throw new Error('Aba "' + SHEET_ID + '" não encontrada.');
  return sheet.getRange('B17').getValue();
}

function getCalendar() {
  const calendarId = getCalendarIdFromSheet();
  const cal = CalendarApp.getCalendarById(calendarId);
  if (!cal) throw new Error('Calendário com ID ' + calendarId + ' não encontrado.');
  return cal;
}

// === Helpers para datas de dia-inteiro ===
function isDateOnly(d) {
  return d instanceof Date &&
         d.getHours() === 0 && d.getMinutes() === 0 &&
         d.getSeconds() === 0 && d.getMilliseconds() === 0;
}
function addDays(d, n) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate() + n);
}
function normalizeDateOnly(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}


function createIdxMap(headerRow) {
  const idxMap = {};
  headerRow.forEach((h, i) => {
    for (const key in TITLE_ROW_MAP) {
      if (TITLE_ROW_MAP[key] === h) {
        idxMap[key] = i;
        break;
      }
    }
  });
  return idxMap;
}

function convertCalEvent(calEvent) {
  const obj = {
    id:          calEvent.getId(),
    title:       calEvent.getTitle(),
    description: calEvent.getDescription(),
    location:    calEvent.getLocation(),
    guests:      calEvent.getGuestList().map(g => g.getEmail()).join(','),
    color:       calEvent.getColor() || '',
    reg:         calEvent.getLastUpdated().toISOString(),
    mod:         new Date().toISOString()
  };

  if (calEvent.isAllDayEvent()) {
    const start = normalizeDateOnly(calEvent.getAllDayStartDate());
    const endExcl = normalizeDateOnly(calEvent.getAllDayEndDate()); // exclusivo na Agenda
    obj.starttime = start;
    const days = Math.round((endExcl - start) / (24 * 3600 * 1000));
    // 1 dia → deixa fim vazio; >1 dia → grava último dia inclusivo
    obj.endtime = (days <= 1) ? '' : addDays(endExcl, -1);
  } else {
    obj.starttime = calEvent.getStartTime();
    obj.endtime   = calEvent.getEndTime();
  }
  return obj;
}

function datesDiffer(d1, d2) {
  if (!d1 && !d2) return false;
  if (!d1 || !d2) return true;
  return d1.getTime() !== d2.getTime();
}

function deleteRowsBatch(sheet, rowIndices) {
  if (!rowIndices || rowIndices.length === 0) return;
  rowIndices.sort((a, b) => b - a);
  let start = rowIndices[0], count = 1;
  for (let i = 1; i < rowIndices.length; i++) {
    if (rowIndices[i] === rowIndices[i - 1] - 1) {
      count++;
    } else {
      sheet.deleteRows(rowIndices[i - 1] - count + 1, count);
      count = 1;
    }
  }
  sheet.deleteRows(rowIndices[rowIndices.length - 1], count);
}

function calEventToSheet(calEvent, idxMap, dataRow) {
  const eventData = convertCalEvent(calEvent);
  for (const key in idxMap) {
    if (eventData[key] !== undefined) {
      dataRow[idxMap[key]] = eventData[key];
    }
  }
}

// ======================
// onEdit (dinâmico; não marca Alterado na Planilha ao clicar "Arquivar")
// ======================
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== SHEET_EVENTS) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row === 1) return; // ignora cabeçalho

  const lastCol = sheet.getLastColumn();
  const hdrs = sheet.getRange(1, SYNC_START_COL, 1, lastCol - SYNC_START_COL + 1).getValues()[0];
  const alterIdx = hdrs.indexOf('Alterado na Planilha');
  const archIdx  = hdrs.indexOf('Arquivar');

  // Colunas de interesse por nome (exclui Arquivar e Gantt)
  const iTitle  = hdrs.indexOf('Título');
  const iDesc   = hdrs.indexOf('Descrição');
  const iStart  = hdrs.indexOf('Data e Hora de Início');
  const iEnd    = hdrs.indexOf('Data e Hora de Fim');
  const iLocal  = hdrs.indexOf('Local');
  const iGuests = hdrs.indexOf('Convidados (e-mail, e-mail, ... )') >= 0
    ? hdrs.indexOf('Convidados (e-mail, e-mail, ... )') // variação rara
    : hdrs.indexOf('Convidados (e-mail, e-mail, ... )');
  const iGuests2 = hdrs.indexOf('Convidados (e-mail, e-mail, ...)');
  const guestsIdxFinal = iGuests2 >= 0 ? iGuests2 : hdrs.indexOf('Convidados (e-mail, e-mail, ...)');
  const iColor  = hdrs.indexOf('Cor');

  const interestAbsCols = [iTitle, iDesc, iStart, iEnd, iLocal, guestsIdxFinal, iColor]
    .filter(i => i >= 0)
    .map(i => SYNC_START_COL + i);

  if (alterIdx !== -1 && interestAbsCols.includes(col)) {
    sheet.getRange(row, SYNC_START_COL + alterIdx).setValue(new Date().toISOString());
  }

  // Dispara SOMENTE quando marcar (TRUE). Desmarcações não gastam recurso.
  if (archIdx !== -1 && col === SYNC_START_COL + archIdx && (e.value === true || e.value === 'TRUE')) {
    const lock = LockService.getDocumentLock(); // lock em nível de arquivo
    try {
      lock.waitLock(30000); // espera se houver outro clique em processamento
      processArchiveSelections(); // processa sempre
    } finally {
      lock.releaseLock();
    }
  }

  // (mantido) Lógica de prioridade
  const fullRow = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const prioTxt = (fullRow[6] || '').toString().trim().toUpperCase();
  const mapPrio  = { 'BAIXA':2, 'MÁXIMA':4, 'MÉDIA':5, 'ALTA':6, 'MÍNIMA':7 };
  let novoVal   = mapPrio[prioTxt];
  if (novoVal !== undefined) {
    const atual = fullRow[18];
    if (atual != novoVal) sheet.getRange(row, 19).setValue(novoVal);
  }
}

// ======================
// ARQUIVAR (zera a célula "Arquivar" antes de mandar)
// ======================
function processArchiveSelections() {
  const ss           = getSpreadsheet();
  const sheetEventos = ss.getSheetByName(SHEET_EVENTS);
  const sheetArquivo = ss.getSheetByName(SHEET_ARCHIVE);
  if (!sheetEventos || !sheetArquivo) return;

  const lastRow0 = sheetEventos.getLastRow();
  const lastCol0 = sheetEventos.getLastColumn();
  if (lastRow0 < 2) return;

  // Revarre até não existir mais TRUE (pega cliques feitos durante o processamento)
  while (true) {
    const lastRow = sheetEventos.getLastRow();
    const lastCol = sheetEventos.getLastColumn();
    if (lastRow < 2) break;

    const hdrs = sheetEventos.getRange(1, SYNC_START_COL, 1, lastCol - SYNC_START_COL + 1).getValues()[0];
    const archiveIdx = hdrs.indexOf('Arquivar');
    if (archiveIdx < 0) break;

    const colNumArch = SYNC_START_COL + archiveIdx;
    const marks = sheetEventos.getRange(2, colNumArch, lastRow - 1, 1).getValues();

    const rowsToDelete = [];
    const archiveBatch = [];

    // de baixo para cima (facilita deleção em blocos)
    for (let i = marks.length - 1; i >= 0; i--) {
      const v = marks[i][0];
      const isTrue = (v === true || v === 'TRUE'); // normaliza boolean
      if (isTrue) {
        const sheetRow = i + 2;
        const fullRow = sheetEventos
          .getRange(sheetRow, SYNC_START_COL, 1, lastCol - SYNC_START_COL + 1)
          .getValues()[0];

        // zera "Arquivar" no buffer (linha será removida da AGENDA)
        fullRow[archiveIdx] = '';
        archiveBatch.unshift(reorderToArchiveFormat(fullRow, hdrs));
        rowsToDelete.push(sheetRow);
      }
    }

    if (!rowsToDelete.length) break;

    // escreve no Arquivo (sempre desde A1)
    sheetArquivo
      .getRange(sheetArquivo.getLastRow() + 1, 1, archiveBatch.length, archiveBatch[0].length)
      .setValues(archiveBatch);

    // apaga linhas da AGENDA
    deleteRowsBatch(sheetEventos, rowsToDelete);
    // e volta ao while para capturar cliques que chegaram enquanto processava
  }
}


// Reordena linha da AGENDA para o layout da aba Arquivo
function reorderToArchiveFormat(row, headers) {
  const desiredOrder = [
    'Título',
    'Descrição',
    'Data e Hora de Início',
    'Data e Hora de Fim',
    'Local',
    'Convidados (e-mail, e-mail, ...)',
    'Cor',
    'PROJETO',
    'ETAPA',
    'STATUS',
    'DEMANDANTE',
    'CONTATOS',
    'AÇÕES RECOMENDADAS',
    'PRIORIDADE',
    'AÇÕES REALIZADAS',
    'OBSERVAÇÕES',
    'ID',
    'Modificado na Agenda',
    'Registro na Agenda',
    'Alterado na Planilha',
    'Gantt',
    'Arquivar'
  ];
  const idxMap = headers.reduce((m, h, i) => (m[h] = i, m), {});
  return desiredOrder.map(label => {
    const idx = idxMap[label];
    return (typeof idx === 'number' && row[idx] !== undefined) ? row[idx] : '';
  });
}

// ======================
// AGENDA → PLANILHA
// ======================
function syncFromCalendar() {
  const cal       = getCalendar();
  const calEvents = cal.getEvents(BEGIN_DATE, END_DATE);

  const ss           = getSpreadsheet();
  const sheetEventos = ss.getSheetByName(SHEET_EVENTS);
  if (!sheetEventos) return;

  const lastRow = sheetEventos.getLastRow();
  const lastCol = sheetEventos.getLastColumn();

  if (lastRow < 2) {
    const headerRow = sheetEventos.getRange(1, SYNC_START_COL, 1, lastCol - SYNC_START_COL + 1).getValues()[0];
    const idxMap = createIdxMap(headerRow);
    const rowsToAdd = calEvents.map(ce => {
      const newRow = new Array(headerRow.length).fill('');
      calEventToSheet(ce, idxMap, newRow);
      return newRow;
    });
    if (rowsToAdd.length) {
      sheetEventos
        .getRange(2, SYNC_START_COL, rowsToAdd.length, headerRow.length)
        .setValues(rowsToAdd);
    }
    return;
  }

  const dataRange = sheetEventos.getRange(1, SYNC_START_COL, lastRow, lastCol - SYNC_START_COL + 1);
  const data = dataRange.getValues();
  const headerRow = data[0];
  const idxMap = createIdxMap(headerRow);

  const sheetEventIds    = data.slice(1).map(r => r[idxMap['id']]);
  const calendarEventIds = calEvents.map(ev => ev.getId());

  const toRemoveRows = [];
  for (let i = sheetEventIds.length - 1; i >= 0; i--) {
    const id = sheetEventIds[i];
    if (id && calendarEventIds.indexOf(id) === -1) {
      toRemoveRows.push(i + 2);
      data[i + 1] = null;
    }
  }

  const rowsToAddBuffer = [];
  for (let j = 0; j < calEvents.length; j++) {
    const ce   = calEvents[j];
    const ceId = ce.getId();
    const idxPlan = sheetEventIds.indexOf(ceId);
    if (idxPlan < 0) {
      const newRow = new Array(headerRow.length).fill('');
      calEventToSheet(ce, idxMap, newRow);
      rowsToAddBuffer.push(newRow);
    } else {
      const rowIdx       = idxPlan + 1;
      const sheetModRaw  = data[rowIdx][idxMap['reg']];
      const sheetModDate = sheetModRaw ? new Date(sheetModRaw) : new Date(0);
      const calModDate   = ce.getLastUpdated();
      if (calModDate > sheetModDate) {
        const updatedRowArr = data[rowIdx];
        const newVals = new Array(headerRow.length).fill('');
        calEventToSheet(ce, idxMap, newVals);
        for (let c = 0; c < headerRow.length; c++) {
          updatedRowArr[c] = newVals[c];
        }
      }
    }
  }

  const cleanData = [headerRow];
  for (let k = 1; k < data.length; k++) {
    if (data[k] !== null) cleanData.push(data[k]);
  }

  sheetEventos
    .getRange(1, SYNC_START_COL, cleanData.length, headerRow.length)
    .setValues(cleanData);

  if (rowsToAddBuffer.length) {
    sheetEventos
      .getRange(cleanData.length + 1, SYNC_START_COL, rowsToAddBuffer.length, headerRow.length)
      .setValues(rowsToAddBuffer);
  }

  // Ordena apenas bloco com data
  if (toRemoveRows.length || rowsToAddBuffer.length) {
    customSortSheet();
  }
}

// ======================
// PLANILHA → AGENDA
// ======================
function syncToCalendar() {
  const cal       = getCalendar();
  const calEvents = cal.getEvents(BEGIN_DATE, END_DATE);
  const calEventIds = calEvents.map(ev => ev.getId());

  const ss           = getSpreadsheet();
  const sheetEventos = ss.getSheetByName(SHEET_EVENTS);
  const sheetArquivo = ss.getSheetByName(SHEET_ARCHIVE);
  if (!sheetEventos) return;

  const archivedIdsSet = new Set();
  if (sheetArquivo) {
    const lastRowArc = sheetArquivo.getLastRow();
    if (lastRowArc > 1) {
      const hdrsArc = sheetArquivo.getRange(1, 1, 1, sheetArquivo.getLastColumn()).getValues()[0]; // A1:...
      const idxIdArc = hdrsArc.indexOf('ID');
      if (idxIdArc !== -1) {
        const idsArcData = sheetArquivo.getRange(2, 1 + idxIdArc, lastRowArc - 1, 1).getValues();
        idsArcData.forEach(r => { if (r[0]) archivedIdsSet.add(r[0]); });
      }
    }
  }

  const lastRowEv   = sheetEventos.getLastRow();
  if (lastRowEv < 2) return;
  const lastColEv   = sheetEventos.getLastColumn();
  const allData     = sheetEventos.getRange(1, SYNC_START_COL, lastRowEv, lastColEv - SYNC_START_COL + 1).getValues();
  const headerRow   = allData[0];
  const idxMap      = createIdxMap(headerRow);

  const idIdx       = idxMap['id'];
  const regIdx      = idxMap['reg'];
  const modIdx      = idxMap['mod'];
  const altIdx      = headerRow.indexOf('Alterado na Planilha');
  const startIdx    = idxMap['starttime'];
  const endIdx      = idxMap['endtime'];
  const titleIdx    = idxMap['title'];
  const descIdx     = idxMap['description'];
  const locIdx      = idxMap['location'];
  const colorIdx    = idxMap['color'];
  const guestsIdx   = idxMap['guests'];

  let rowsChanged = false;
  const updatedData = [headerRow];

  // === Controle de lote para evitar "muitas entradas de uma vez" ===
  let createdInBatch = 0;

  function isRateLimitError(err) {
    const msg = (err && err.message) ? err.message.toString() : '';
    return /many|quota|rate|temporar|try again|invoked too many times/i.test(msg);
  }

  function sleepBetweenBatchesIfNeeded() {
    if (createdInBatch >= CREATE_BATCH_SIZE) {
      Utilities.sleep(CREATE_BATCH_SLEEP_MS);
      createdInBatch = 0;
    }
  }


  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const sheetId = row[idIdx];

    if (sheetId && archivedIdsSet.has(sheetId)) {
      updatedData.push(row);
      continue;
    }

    const rawStart = row[startIdx];
    const dtStart  = (rawStart instanceof Date) ? rawStart : null;
    if (!dtStart || isNaN(dtStart.getTime())) {
      updatedData.push(row);
      continue;
    }

    const sheetEvent = {
      id:       sheetId,
      title:    row[titleIdx],
      description: row[descIdx] || '',
      location: row[locIdx] || '',
      color:    row[colorIdx] || '',
      guests:   row[guestsIdx] || '',
      starttime: dtStart,
      endtime:   (row[endIdx] instanceof Date) ? row[endIdx] : null
    };

    if (sheetId) {
      const calIdx = calEventIds.indexOf(sheetId);
      if (calIdx >= 0) {
        let calEv       = calEvents[calIdx];
        const sheetRegRaw = row[regIdx];
        const sheetReg    = sheetRegRaw ? new Date(sheetRegRaw) : new Date(0);
        const calLastUpd  = calEv.getLastUpdated();
        const altPlanDate = new Date(row[altIdx] || 0);

        if (altPlanDate > calLastUpd && altPlanDate > sheetReg) {
          let numChanges = 0;
          const wantAllDaySingle = !sheetEvent.endtime && isDateOnly(sheetEvent.starttime);
          const wantAllDayMulti  = sheetEvent.endtime && isDateOnly(sheetEvent.starttime) && isDateOnly(sheetEvent.endtime);
          const wantTimed        = !wantAllDaySingle && !wantAllDayMulti; // há hora em início ou fim

          if (wantAllDaySingle) {
            const desiredStart = normalizeDateOnly(sheetEvent.starttime);
            if (calEv.isAllDayEvent()) {
              const curStart = normalizeDateOnly(calEv.getAllDayStartDate());
              if (datesDiffer(curStart, desiredStart)) {
                calEv.setAllDayDate(desiredStart);
                numChanges++;
              }
            } else {
              calEv.deleteEvent();
              calEv = cal.createAllDayEvent(sheetEvent.title, desiredStart);
              row[idIdx]  = calEv.getId();
              row[regIdx] = calEv.getLastUpdated().toISOString();
              numChanges++;
            }

          } else if (wantAllDayMulti) {
            const desiredStart   = normalizeDateOnly(sheetEvent.starttime);
            const desiredEndExcl = addDays(sheetEvent.endtime, 1);
            let needRecreate = true;
            if (calEv.isAllDayEvent()) {
              const curStart   = normalizeDateOnly(calEv.getAllDayStartDate());
              const curEndExcl = normalizeDateOnly(calEv.getAllDayEndDate());
              needRecreate = datesDiffer(curStart, desiredStart) || datesDiffer(curEndExcl, desiredEndExcl);
            }
            if (needRecreate || !calEv.isAllDayEvent()) {
              calEv.deleteEvent();
              calEv = cal.createAllDayEvent(sheetEvent.title, desiredStart, desiredEndExcl);
              row[idIdx]  = calEv.getId();
              row[regIdx] = calEv.getLastUpdated().toISOString();
              numChanges++;
            }

          } else { // evento com hora
            // preencher fim padrão se vier vazio
            let startEff = sheetEvent.starttime;
            let endEff   = sheetEvent.endtime;
            const hasStartTime = startEff && !isDateOnly(startEff);
            const hasEndTime   = endEff && !isDateOnly(endEff);
            if (!endEff && hasStartTime) {
              endEff = new Date(startEff.getTime() + DEFAULT_TIMED_DURATION_MIN * 60000);
            }
            if (calEv.isAllDayEvent()) {
              calEv.deleteEvent();
              calEv = cal.createEvent(sheetEvent.title, startEff, endEff);
              row[idIdx]  = calEv.getId();
              row[regIdx] = calEv.getLastUpdated().toISOString();
              numChanges++;
            } else if (datesDiffer(calEv.getStartTime(), startEff) ||
                       datesDiffer(calEv.getEndTime(),   endEff)) {
              calEv.setTime(startEff, endEff);
              numChanges++;
            }
          }
          if (calEv.getTitle() !== sheetEvent.title)       { calEv.setTitle(sheetEvent.title); numChanges++; }
          if (calEv.getDescription() !== sheetEvent.description) { calEv.setDescription(sheetEvent.description); numChanges++; }
          if (calEv.getLocation() !== sheetEvent.location) { calEv.setLocation(sheetEvent.location); numChanges++; }
          if (calEv.getColor() !== ('' + sheetEvent.color)) {
            if (sheetEvent.color && sheetEvent.color >= 1 && sheetEvent.color <= 11) {
              calEv.setColor('' + sheetEvent.color); numChanges++;
            }
          }
          if (numChanges > 0) { row[regIdx] = new Date().toISOString(); rowsChanged = true; }
        }
        updatedData.push(row);
      } else {
        row[titleIdx] = NO_SYNC_STRING + " " + row[titleIdx];
        row[idIdx]    = '';
        updatedData.push(row);
        rowsChanged = true;
      }
    } else {
      let newEv;
      const startOnly = isDateOnly(sheetEvent.starttime);
      const endOnly   = sheetEvent.endtime && isDateOnly(sheetEvent.endtime);
      const hasStartTime = sheetEvent.starttime && !isDateOnly(sheetEvent.starttime);
      const hasEndTime   = sheetEvent.endtime && !isDateOnly(sheetEvent.endtime);

      // Tenta criar; se bater rate limit, espera 15s e tenta mais 1x
      const tryCreate = () => {
        if (!sheetEvent.endtime && !hasStartTime) {
          return cal.createAllDayEvent(sheetEvent.title, normalizeDateOnly(sheetEvent.starttime));
        } else if (startOnly && endOnly) {
          return cal.createAllDayEvent(
            sheetEvent.title,
            normalizeDateOnly(sheetEvent.starttime),
            addDays(sheetEvent.endtime, 1)
          );
        } else {
          const startEff = hasStartTime
            ? sheetEvent.starttime
            : new Date(sheetEvent.starttime.getFullYear(), sheetEvent.starttime.getMonth(), sheetEvent.starttime.getDate(), 0, 0, 0);

          const endEff = sheetEvent.endtime
            ? sheetEvent.endtime
            : new Date(startEff.getTime() + DEFAULT_TIMED_DURATION_MIN * 60000);

          return cal.createEvent(sheetEvent.title, startEff, endEff);
        }
      };

      try {
        newEv = tryCreate();
      } catch (err) {
        if (isRateLimitError(err)) {
          Utilities.sleep(CREATE_BATCH_SLEEP_MS);
          newEv = tryCreate(); // retry 1x
        } else {
          throw err;
        }
      }

      if (sheetEvent.color && sheetEvent.color >= 1 && sheetEvent.color <= 11) {
        newEv.setColor('' + sheetEvent.color);
      }

      const newId  = newEv.getId();
      const nowTs  = new Date().toISOString();
      row[idIdx]   = newId;
      row[modIdx]  = nowTs;
      row[regIdx]  = newEv.getLastUpdated().toISOString();
      updatedData.push(row);
      rowsChanged = true;

      createdInBatch++;
      sleepBetweenBatchesIfNeeded();
    }
  }

  if (rowsChanged) {
    sheetEventos
      .getRange(2, SYNC_START_COL, updatedData.length - 1, headerRow.length)
      .setValues(updatedData.slice(1));
  }
}

// ======================
// MOVER SEM ID PARA ARQUIVO
// ======================
function moveEventsToArchive() {
  const ss           = getSpreadsheet();
  const sheetEventos = ss.getSheetByName(SHEET_EVENTS);
  const sheetArquivo = ss.getSheetByName(SHEET_ARCHIVE);
  if (!sheetEventos || !sheetArquivo) return;

  const lastRowEv  = sheetEventos.getLastRow();
  if (lastRowEv < 2) return;

  const lastColEv  = sheetEventos.getLastColumn();
  const allData    = sheetEventos.getRange(1, SYNC_START_COL, lastRowEv, lastColEv - SYNC_START_COL + 1).getValues();
  const headerRow  = allData[0];
  const idxTitle   = headerRow.indexOf('Título');
  const idxId      = headerRow.indexOf('ID');
  const idxReg     = headerRow.indexOf('Registro na Agenda');
  if (idxTitle < 0 || idxId < 0 || idxReg < 0) return;

  const iArch = headerRow.indexOf('Arquivar');
  const rowsToArchive = [];
  const rowsToDelete  = [];

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (row[idxTitle] && row[idxReg] && !row[idxId]) {
      const copy = row.slice();
      if (iArch >= 0) copy[iArch] = '';
      rowsToArchive.push(copy);
      rowsToDelete.push(i + 1);
    }
  }

  if (rowsToArchive.length) {
    const batchData = rowsToArchive.map(r => reorderToArchiveFormat(r, headerRow));
    sheetArquivo
      .getRange(sheetArquivo.getLastRow() + 1, 1, batchData.length, batchData[0].length)
      .setValues(batchData);

    deleteRowsBatch(sheetEventos, rowsToDelete);
  }
}

// ======================
// CHECKBOXES somente em linhas com data
// ======================
function applyCheckboxes() {
  const sheet = getSheet(SHEET_EVENTS);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sheet.getLastColumn();
  const hdrs = sheet.getRange(1, SYNC_START_COL, 1, lastCol - SYNC_START_COL + 1).getValues()[0];

  const idxTitle   = hdrs.indexOf('Título');
  const idxStart   = hdrs.indexOf('Data e Hora de Início');
  const idxArchive = hdrs.indexOf('Arquivar');
  const idxGantt   = hdrs.indexOf('Gantt');
  if ([idxTitle, idxStart, idxArchive, idxGantt].some(i => i < 0)) return;

  const colTitle   = SYNC_START_COL + idxTitle;
  const colStart   = SYNC_START_COL + idxStart;
  const colArchive = SYNC_START_COL + idxArchive;
  const colGantt   = SYNC_START_COL + idxGantt;

  const valuesTitle = sheet.getRange(2, colTitle, lastRow - 1, 1).getValues();
  const valuesStart = sheet.getRange(2, colStart, lastRow - 1, 1).getValues();

  const rowsValid = [];
  const rowsInvalid = [];
  for (let i = 0; i < valuesTitle.length; i++) {
    const title = valuesTitle[i][0];
    const start = valuesStart[i][0];
    if (title && (start instanceof Date) && !isNaN(start.getTime())) {
      rowsValid.push(i + 2);
    } else {
      rowsInvalid.push(i + 2);
    }
  }

  function insertForColumn(colNum) {
    if (!rowsValid.length) return;
    let start = rowsValid[0], count = 1;
    for (let j = 1; j < rowsValid.length; j++) {
      if (rowsValid[j] === rowsValid[j - 1] + 1) {
        count++;
      } else {
        sheet.getRange(start, colNum, count, 1).insertCheckboxes();
        start = rowsValid[j];
        count = 1;
      }
    }
    sheet.getRange(start, colNum, count, 1).insertCheckboxes();
  }

  insertForColumn(colArchive);
  insertForColumn(colGantt);

  if (rowsInvalid.length) {
    rowsInvalid.forEach(r => {
      sheet.getRange(r, colArchive).clearContent();
      sheet.getRange(r, colGantt).clearContent();
    });
  }
}

// ======================
// REMOVER EVENTOS DO CALENDÁRIO A PARTIR DO ARQUIVO (ID)
// ======================
function removeEventsFromFile() {
  const archiveSheet = getSheet(SHEET_ARCHIVE);
  if (!archiveSheet) return;

  const lastRowArc = archiveSheet.getLastRow();
  if (lastRowArc < 2) return;

  const lastColArc = archiveSheet.getLastColumn();
  const allArcData = archiveSheet.getRange(1, 1, lastRowArc, lastColArc).getValues(); // A1:...
  const hdrsArc    = allArcData[0];
  const idxIdArc   = hdrsArc.indexOf('ID');
  if (idxIdArc < 0) return;

  const toUpdate = [];
  for (let i = 1; i < allArcData.length; i++) {
    const evId = allArcData[i][idxIdArc];
    if (evId) {
      deleteEventById(evId);
      allArcData[i][idxIdArc] = '';
      toUpdate.push(i + 1);
    }
  }
  if (toUpdate.length) {
    archiveSheet.getRange(1, 1, allArcData.length, hdrsArc.length).setValues(allArcData);
  }
}

// ======================
//
// LIMPA LINHAS SEM TÍTULO
// ======================
function clearEmptyTitleRows() {
  const sheet = getSheet(SHEET_EVENTS);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sheet.getLastColumn();
  const range   = sheet.getRange(1, SYNC_START_COL, lastRow, lastCol - SYNC_START_COL + 1);
  const allData = range.getValues();
  const hdrs    = allData[0];

  const idxId        = hdrs.indexOf('ID');
  const idxModAgenda = hdrs.indexOf('Modificado na Agenda');
  const idxRegSheet  = hdrs.indexOf('Registro na Agenda');
  const idxAlt       = hdrs.indexOf('Alterado na Planilha');
  const idxTitle     = hdrs.indexOf('Título');
  if (idxId < 0 || idxModAgenda < 0 || idxRegSheet < 0 || idxAlt < 0 || idxTitle < 0) return;

  let changed = false;
  for (let i = 1; i < allData.length; i++) {
    if (!allData[i][idxTitle]) {
      allData[i][idxId]        = '';
      allData[i][idxModAgenda] = '';
      allData[i][idxRegSheet]  = '';
      allData[i][idxAlt]       = '';
      changed = true;
    }
  }
  if (changed) range.setValues(allData);
}

// ======================
// ORDENAR (Data → PROJETO → Título) só no bloco com data
// ======================
function customSortSheet() {
  const sheet = getSheet(SHEET_EVENTS);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sheet.getLastColumn();
  const header = sheet.getRange(1, SYNC_START_COL, 1, lastCol - SYNC_START_COL + 1).getValues()[0];

  const iProjeto = header.indexOf('PROJETO');
  const iStart   = header.indexOf('Data e Hora de Início');
  const iTitulo  = header.indexOf('Título');
  if (iProjeto < 0 || iStart < 0 || iTitulo < 0) return;

  const colStartAbs = SYNC_START_COL + iStart;
  const startVals = sheet.getRange(2, colStartAbs, lastRow - 1, 1).getValues();
  let lastDateRow = 1;
  for (let i = startVals.length - 1; i >= 0; i--) {
    const d = getDateValue(startVals[i][0]);
    if (d instanceof Date && !isNaN(d.getTime())) { lastDateRow = i + 2; break; }
  }
  if (lastDateRow <= 1) return;

  const sortRange = sheet.getRange(2, SYNC_START_COL, lastDateRow - 1, lastCol - SYNC_START_COL + 1);
  sortRange.sort([
    { column: SYNC_START_COL + iStart,   ascending: true },  // Data
    { column: SYNC_START_COL + iProjeto, ascending: true },  // PROJETO
    { column: SYNC_START_COL + iTitulo,  ascending: true },  // Título
  ]);
}

function getDateValue(cellValue) {
  if (cellValue instanceof Date) {
    return (cellValue.getHours() === 0 && cellValue.getMinutes() === 0 && cellValue.getSeconds() === 0)
      ? new Date(cellValue.getFullYear(), cellValue.getMonth(), cellValue.getDate())
      : cellValue;
  }
  if (typeof cellValue === 'string') {
    const [datePart, timePart = ""] = cellValue.split(" ");
    const [dd, mm, yy] = datePart.split('/');
    let year = parseInt(yy, 10);
    if (year < 100) year += 2000;
    const day   = parseInt(dd, 10);
    const month = parseInt(mm, 10) - 1;
    if (timePart) {
      const [hh, min, ss = "0"] = timePart.split(':');
      const hour   = parseInt(hh, 10);
      const minute = parseInt(min, 10);
      const second = parseInt(ss, 10);
      return (hour === 0 && minute === 0 && second === 0)
        ? new Date(year, month, day)
        : new Date(year, month, day, hour, minute, second);
    }
    return new Date(year, month, day);
  }
  return new Date(NaN);
}

// ======================
// FUNÇÕES AUXILIARES DE EVENTOS
// ======================
function updateEvent(calEvent, sheetEvent) {
  let numChanges = 0;
  if (sheetEvent.starttime && (datesDiffer(calEvent.getStartTime(), sheetEvent.starttime) ||
      datesDiffer(calEvent.getEndTime(), sheetEvent.endtime))) {
    if (!sheetEvent.endtime) calEvent.setAllDayDate(sheetEvent.starttime);
    else calEvent.setTime(sheetEvent.starttime, sheetEvent.endtime);
    numChanges++;
  }
  if (calEvent.getTitle() !== sheetEvent.title) { calEvent.setTitle(sheetEvent.title); numChanges++; }
  if (calEvent.getDescription() !== sheetEvent.description) { calEvent.setDescription(sheetEvent.description); numChanges++; }
  if (calEvent.getLocation() !== sheetEvent.location) { calEvent.setLocation(sheetEvent.location); numChanges++; }
  if (calEvent.getColor() !== ('' + sheetEvent.color)) {
    if (sheetEvent.color && sheetEvent.color >= 1 && sheetEvent.color <= 11) {
      calEvent.setColor('' + sheetEvent.color); numChanges++;
    }
  }
  return numChanges;
}

function deleteEventById(eventId) {
  const cal = getCalendar();
  try {
    const ev = cal.getEventById(eventId);
    if (ev) ev.deleteEvent();
  } catch (err) {
    // ignorar se já foi excluído
  }
}

// ======================
// MENU E TRIGGERS
// ======================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📅 SINCRONIZAR PLANILHA⇿AGENDA')
    .addItem('Sincronizar Agenda e Planilha', 'syncCalendarAndSheet')
    .addSubMenu(
      ui.createMenu('Outras opções')
        .addItem('Criar Sequência de Dias na Planilha Gantt', 'gerarCabecalhoGantt') // se não existir, remova essa linha do menu
        .addItem('Criar Sequência de Dias na Planilha AGENDA', 'createEventsWithPrompts')
        .addItem('Apagar Sequência de Dias Úteis e Finais de Semana na Planilha AGENDA', 'deleteEventsWithPrompts')
        .addItem('Apagar Todos os Eventos em um intervalo na Planilha AGENDA', 'deleteAllEventsWithPrompts')
        .addItem('Apagar Todos os Eventos da Agenda', 'deleteAllEventsFromCalendar')
        .addItem('Criar Nova Agenda', 'criarNovaAgenda')
    )
    .addToUi();
}

/**
 * Lista linhas onde "Fim" < "Início".
 */
function preSyncValidateDates() {
  const sheet = getSheet(SHEET_EVENTS);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const lastCol = sheet.getLastColumn();
  const range   = sheet.getRange(1, SYNC_START_COL, lastRow, lastCol - SYNC_START_COL + 1);
  const values  = range.getValues();
  const header  = values[0];

  const iTitle = header.indexOf('Título');
  const iStart = header.indexOf('Data e Hora de Início');
  const iEnd   = header.indexOf('Data e Hora de Fim');
  if (iTitle < 0 || iStart < 0 || iEnd < 0) return [];

  const issues = [];
  for (let r = 1; r < values.length; r++) {
    const row   = values[r];
    const start = getDateValue(row[iStart]);
    const end   = getDateValue(row[iEnd]);

    const startOk = (start instanceof Date) && !isNaN(start.getTime());
    const endOk   = (end   instanceof Date) && !isNaN(end.getTime());

    if (startOk && endOk && end.getTime() < start.getTime()) {
      issues.push({
        absoluteRow: r + 1,
        title: (row[iTitle] || '').toString(),
        start, end
      });
    }
  }
  return issues;
}

function syncCalendarAndSheet() {
  const issues = preSyncValidateDates();
  if (issues.length) {
    const ui = SpreadsheetApp.getUi();
    const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
    const fmt = d => Utilities.formatDate(d, tz, ((d.getHours() + d.getMinutes() + d.getSeconds()) ? 'dd/MM/yy HH:mm' : 'dd/MM/yy'));
    const maxShow = 50;
    const lines = issues.slice(0, maxShow).map(it => {
      const t = it.title && it.title.trim() ? it.title.trim() : '(sem título)';
      return `Linha ${it.absoluteRow}: ${t} — fim ${fmt(it.end)} < início ${fmt(it.start)}`;
    }).join('\n');
    const extra = issues.length > maxShow ? `\n… e mais ${issues.length - maxShow} item(s).` : '';
    ui.alert(
      'Datas inválidas encontradas',
      'A "Data e Hora de Fim" deve ser igual ou maior que a "Data e Hora de Início".\n\n' +
      lines + extra + '\n\nCorrija e execute novamente.',
      ui.ButtonSet.OK
    );
    return;
  }

  clearEmptyTitleRows();
  removeEventsFromFile();
  syncFromCalendar();
  syncToCalendar();
  moveEventsToArchive();
  applyCheckboxes();   // manter após sync
  customSortSheet();   // ordena Data → PROJETO → Título (apenas bloco com data)
  formatDateColumns();
  syncGanttFromAgenda();
}

// ======================
// Criar sequência de dias (sem índices fixos)
// ======================
function createEventsWithPrompts() {
  const ui = SpreadsheetApp.getUi();

  const resp1 = ui.prompt('Criar Sequência de Dias na Planilha', 'Informe a data inicial (dd/mm/aaaa):', ui.ButtonSet.OK_CANCEL);
  if (resp1.getSelectedButton() !== ui.Button.OK) { ui.alert('Operação cancelada.'); return; }
  const startDate = parseDate(resp1.getResponseText());
  if (!startDate) { ui.alert('Data inicial inválida!'); return; }

  const resp2 = ui.prompt('Criar Sequência de Dias na Planilha', 'Informe a data final (dd/mm/aaaa):', ui.ButtonSet.OK_CANCEL);
  if (resp2.getSelectedButton() !== ui.Button.OK) { ui.alert('Operação cancelada.'); return; }
  const endDate = parseDate(resp2.getResponseText());
  if (!endDate) { ui.alert('Data final inválida!'); return; }
  if (endDate < startDate) { ui.alert('A data final deve ser igual ou posterior à inicial!'); return; }

  const sheet = getSheet(SHEET_EVENTS);
  if (!sheet) throw new Error('Aba "' + SHEET_EVENTS + '" não encontrada.');

  const lastRow = sheet.getLastRow();
  const hdr = sheet.getRange(1, SYNC_START_COL, 1, sheet.getLastColumn() - SYNC_START_COL + 1).getValues()[0];

  const iTitle = hdr.indexOf('Título');
  const iStart = hdr.indexOf('Data e Hora de Início');
  const iColor = hdr.indexOf('Cor');
  if ([iTitle,iStart,iColor].some(i => i < 0)) throw new Error('Cabeçalhos esperados não encontrados.');

  const titles = lastRow >= 2 ? sheet.getRange(2, SYNC_START_COL + iTitle, lastRow - 1, 1).getValues().flat() : [];
  const dates  = lastRow >= 2 ? sheet.getRange(2, SYNC_START_COL + iStart, lastRow - 1, 1).getValues().flat() : [];
  const existingSet = new Set();
  if (lastRow >= 2) {
    for (let i = 0; i < titles.length; i++) {
      const t = titles[i];
      if (!t) continue;

      const tStr = t.toString().trim();
      if (!tStr.startsWith('-')) continue;

      const name = normDayName(tStr);
      if (!name) continue;

      const dk = toKeyDate(dates[i]);
      if (dk) existingSet.add(makeKey(dk, name));
    }
  }

  const tz    = sheet.getParent().getSpreadsheetTimeZone();
  const names = { 1:'Segunda', 2:'Terça', 3:'Quarta', 4:'Quinta', 5:'Sexta', 6:'Sábado', 0:'Domingo' };

  function normDayName(titleCell) {
    if (!titleCell) return '';
    return titleCell.toString().replace(/^\s*-\s*/,'').trim();
  }

  function toKeyDate(cell) {
    const d = (cell instanceof Date) ? cell : getDateValue((cell || '').toString());
    if (!(d instanceof Date) || isNaN(d.getTime())) return '';
    return Utilities.formatDate(d, tz, 'yyyy-MM-dd');
  }

  function makeKey(dateKey, name) {
    return dateKey + '|' + (name || '').trim();
  }


  const rowsToInsert = [];
  const templateRow = new Array(hdr.length).fill('');

  let cur = new Date(startDate);
  while (cur <= endDate) {
    const w = cur.getDay();
    if (names.hasOwnProperty(w)) {
      const name = names[w];
      const dk   = Utilities.formatDate(cur, tz, 'yyyy-MM-dd');

      if (!existingSet.has(makeKey(dk, name))) {
        const newRow = templateRow.slice();
        newRow[iTitle] = '- ' + name;

        // grava como Date (numérico) para manter padrão de coluna de data
        newRow[iStart] = new Date(cur.getFullYear(), cur.getMonth(), cur.getDate());

        newRow[iColor] = 8;
        rowsToInsert.push(newRow);

        // evita duplicar no mesmo batch
        existingSet.add(makeKey(dk, name));
      }
    }
    cur.setDate(cur.getDate() + 1);
  }

  if (rowsToInsert.length) {
    sheet.getRange(lastRow + 1, SYNC_START_COL, rowsToInsert.length, hdr.length).setValues(rowsToInsert);
    customSortSheet();
  }

  ui.alert(`Operação concluída: ${rowsToInsert.length} linha(s) adicionada(s).`);
}

// ======================
// APAGAR SEQUÊNCIA (zera "Arquivar" e usa reorderToArchiveFormat)
// ======================
function deleteEventsWithPrompts() {
  const ui = SpreadsheetApp.getUi();

  const resp1 = ui.prompt('Apagar Sequência de Dias na Planilha', 'Informe a data inicial (dd/mm/aaaa):', ui.ButtonSet.OK_CANCEL);
  if (resp1.getSelectedButton() !== ui.Button.OK) { ui.alert('Operação cancelada.'); return; }
  const startDate = parseDate(resp1.getResponseText());
  if (!startDate) { ui.alert('Data inicial inválida!'); return; }

  const resp2 = ui.prompt('Apagar Sequência de Dias na Planilha', 'Informe a data final (dd/mm/aaaa):', ui.ButtonSet.OK_CANCEL);
  if (resp2.getSelectedButton() !== ui.Button.OK) { ui.alert('Operação cancelada.'); return; }
  const endDate = parseDate(resp2.getResponseText());
  if (!endDate) { ui.alert('Data final inválida!'); return; }
  if (endDate < startDate) { ui.alert('A data final deve ser igual ou posterior à inicial!'); return; }

  const ss           = getSpreadsheet();
  const sheetEventos = ss.getSheetByName(SHEET_EVENTS);
  const sheetArquivo = ss.getSheetByName(SHEET_ARCHIVE);
  if (!sheetEventos || !sheetArquivo) throw new Error('Aba "' + SHEET_EVENTS + '" ou "' + SHEET_ARCHIVE + '" não encontrada.');

  const lastRow = sheetEventos.getLastRow();
  if (lastRow < 2) { ui.alert('Nenhuma linha para apagar.'); return; }
  const lastCol = sheetEventos.getLastColumn();

  const allData = sheetEventos.getRange(1, SYNC_START_COL, lastRow, lastCol - SYNC_START_COL + 1).getValues();
  const headerRow = allData[0];
  const idxTitle   = headerRow.indexOf('Título');
  const idxStart   = headerRow.indexOf('Data e Hora de Início');
  const iArch      = headerRow.indexOf('Arquivar');
  if (idxTitle < 0 || idxStart < 0) { ui.alert('Colunas "Título" ou "Data e Hora de Início" não encontradas.'); return; }

  const archiveBatch = [];
  const rowsToDelete = [];

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const titleCell = row[idxTitle];
    if (titleCell && titleCell.toString().startsWith('- ')) {
      const dateCell = row[idxStart];
      const dateObj = (dateCell instanceof Date) ? dateCell : getDateValue((dateCell || '').toString());
      if (dateObj instanceof Date && !isNaN(dateObj.getTime())) {
        if (dateObj >= startDate && dateObj <= endDate) {
          const rowCopy = row.slice();
          if (iArch >= 0) rowCopy[iArch] = '';
          archiveBatch.unshift(reorderToArchiveFormat(rowCopy, headerRow));
          rowsToDelete.push(i + 1);
        }
      }
    }
  }

  if (rowsToDelete.length) {
    sheetArquivo
      .getRange(sheetArquivo.getLastRow() + 1, 1, archiveBatch.length, archiveBatch[0].length)
      .setValues(archiveBatch);

    deleteRowsBatch(sheetEventos, rowsToDelete);
    ui.alert(`Foram movidas ${rowsToDelete.length} linha(s) para a aba Arquivo.`);
  } else {
    ui.alert('Nenhum evento encontrado para arquivar neste intervalo.');
  }
}

function deleteAllEventsWithPrompts() {
  const ui = SpreadsheetApp.getUi();

  const resp1 = ui.prompt('Apagar Todos os Eventos na Planilha', 'Informe a data inicial (dd/mm/aaaa):', ui.ButtonSet.OK_CANCEL);
  if (resp1.getSelectedButton() !== ui.Button.OK) { ui.alert('Operação cancelada.'); return; }
  const startDate = parseDate(resp1.getResponseText());
  if (!startDate) { ui.alert('Data inicial inválida!'); return; }

  const resp2 = ui.prompt('Apagar Todos os Eventos na Planilha', 'Informe a data final (dd/mm/aaaa):', ui.ButtonSet.OK_CANCEL);
  if (resp2.getSelectedButton() !== ui.Button.OK) { ui.alert('Operação cancelada.'); return; }
  const endDate = parseDate(resp2.getResponseText());
  if (!endDate) { ui.alert('Data final inválida!'); return; }
  if (endDate < startDate) { ui.alert('A data final deve ser igual ou posterior à inicial!'); return; }

  const ss           = getSpreadsheet();
  const sheetEventos = ss.getSheetByName(SHEET_EVENTS);
  const sheetArquivo = ss.getSheetByName(SHEET_ARCHIVE);
  if (!sheetEventos || !sheetArquivo) throw new Error('Aba "' + SHEET_EVENTS + '" ou "' + SHEET_ARCHIVE + '" não encontrada.');

  const lastRow = sheetEventos.getLastRow();
  if (lastRow < 2) { ui.alert('Nenhuma linha para apagar.'); return; }
  const lastCol = sheetEventos.getLastColumn();

  const allData = sheetEventos.getRange(1, SYNC_START_COL, lastRow, lastCol - SYNC_START_COL + 1).getValues();
  const headerRow = allData[0];
  const idxStart   = headerRow.indexOf('Data e Hora de Início');
  const iArch      = headerRow.indexOf('Arquivar');
  if (idxStart < 0) { ui.alert('Coluna "Data e Hora de Início" não encontrada.'); return; }

  const archiveBatch = [];
  const rowsToDelete = [];

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const dateCell = row[idxStart];
    const dateObj = (dateCell instanceof Date) ? dateCell : getDateValue((dateCell || '').toString());
    if (dateObj instanceof Date && !isNaN(dateObj.getTime())) {
      if (dateObj >= startDate && dateObj <= endDate) {
        const rowCopy = row.slice();
        if (iArch >= 0) rowCopy[iArch] = '';
        archiveBatch.unshift(reorderToArchiveFormat(rowCopy, headerRow));
        rowsToDelete.push(i + 1);
      }
    }
  }

  if (rowsToDelete.length) {
    sheetArquivo
      .getRange(sheetArquivo.getLastRow() + 1, 1, archiveBatch.length, archiveBatch[0].length)
      .setValues(archiveBatch);

    deleteRowsBatch(sheetEventos, rowsToDelete);
    ui.alert(`Foram movidas ${rowsToDelete.length} linha(s) para a aba Arquivo.`);
  } else {
    ui.alert('Nenhum evento encontrado para arquivar neste intervalo.');
  }
}

function parseDate(dateStr) {
  const parts = dateStr.split('/');
  if (parts.length !== 3) return null;
  const d = parseInt(parts[0], 10), m = parseInt(parts[1], 10) - 1, y = parseInt(parts[2], 10);
  const dt = new Date(y, m, d);
  return isNaN(dt.getTime()) ? null : dt;
}

// ======================
// EXCLUIR TODOS OS EVENTOS DO CALENDÁRIO
// ======================
function deleteAllEventsFromCalendar() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert('Confirmação', 'Apagar todos os eventos? Não tem volta.', ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    ui.alert('Cancelado'); return;
  }
  if (ui.alert('Confirmação', 'Confirma exclusão?', ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    ui.alert('Cancelado'); return;
  }
  if (ui.alert('Última chance', 'Tem certeza absoluta?', ui.ButtonSet.YES_NO) !== ui.Button.YES) {
    ui.alert('Cancelado'); return;
  }

  try {
    const cal       = getCalendar();
    const allEvents = cal.getEvents(new Date(2000, 0, 1), new Date(2100, 0, 1));
    let deletedCount = 0, failedCount = 0;
    const failedList = [];

    allEvents.forEach(ev => {
      try {
        ev.deleteEvent();
        deletedCount++;
      } catch (err) {
        const t = ev ? ev.getTitle() : 'Título indisponível';
        const i = ev ? ev.getId()    : 'ID indisponível';
        failedList.push({ title: t, id: i });
        failedCount++;
      }
    });

    let msg = `${deletedCount} evento(s) apagados.`;
    if (failedCount) {
      msg += ` (${failedCount} falharam)\n\n` +
             failedList.map(e => `Título: ${e.title} (ID: ${e.id})`).join('\n');
    }
    ui.alert(msg);
  } catch (err) {
    ui.alert('Erro: ' + err.message);
  }
}

function formatDateColumns() {
  const sheet = getSheet(SHEET_EVENTS);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const lastCol = sheet.getLastColumn();
  const hdrs = sheet.getRange(1, SYNC_START_COL, 1, lastCol - SYNC_START_COL + 1).getValues()[0];

  const idxStart = hdrs.indexOf(TITLE_ROW_MAP.starttime);
  const idxEnd   = hdrs.indexOf(TITLE_ROW_MAP.endtime);
  if (idxStart < 0 || idxEnd < 0) return;

  const colStart = SYNC_START_COL + idxStart;
  const colEnd   = SYNC_START_COL + idxEnd;

  const valuesStart = sheet.getRange(2, colStart, lastRow - 1, 1).getValues();
  const valuesEnd   = sheet.getRange(2, colEnd,   lastRow - 1, 1).getValues();

  const formatsStart = [];
  const formatsEnd   = [];
  for (let i = 0; i < valuesStart.length; i++) {
    const d1 = valuesStart[i][0], d2 = valuesEnd[i][0];
    const fmtFn = d => (d instanceof Date && !isNaN(d.getTime()))
      ? (d.getHours()||d.getMinutes()||d.getSeconds() ? "dd/mm/yy hh:mm" : "dd/mm/yy")
      : "";
    formatsStart.push([ fmtFn(d1) ]);
    formatsEnd.push([ fmtFn(d2) ]);
  }

  sheet.getRange(2, colStart, lastRow - 1, 1).setNumberFormats(formatsStart);
  sheet.getRange(2, colEnd,   lastRow - 1, 1).setNumberFormats(formatsEnd);
}

/**
 * Gantt dinâmico (coluna "Gantt" como checkbox indica quais linhas plotar)
 */
function syncGanttFromAgenda() {
  const ss     = SpreadsheetApp.getActiveSpreadsheet();
  const agenda = ss.getSheetByName(SHEET_EVENTS);
  const gantt  = ss.getSheetByName('Gantt');
  if (!agenda || !gantt) throw new Error('Aba AGENDA ou Gantt não encontrada.');

  const lastColA = agenda.getLastColumn();
  const hdrA     = agenda.getRange(1, SYNC_START_COL, 1, lastColA - SYNC_START_COL + 1).getValues()[0];

  const iTitle    = hdrA.indexOf(TITLE_ROW_MAP['title']);
  const iStart    = hdrA.indexOf(TITLE_ROW_MAP['starttime']);
  const iEnd      = hdrA.indexOf(TITLE_ROW_MAP['endtime']);
  const iGanttChk = hdrA.indexOf('Gantt');
  if ([iTitle,iStart,iEnd,iGanttChk].some(i => i < 0)) return;

  const rows = agenda.getRange(2, SYNC_START_COL, agenda.getLastRow()-1, hdrA.length).getValues();
  const eventos = rows.map(r => {
    const st = (r[iStart] instanceof Date) ? new Date(r[iStart].getFullYear(),r[iStart].getMonth(),r[iStart].getDate()) : null;
    const en = (r[iEnd]   instanceof Date) ? new Date(r[iEnd].getFullYear(),  r[iEnd].getMonth(),  r[iEnd].getDate()) : null;
    return { title: r[iTitle], start: st, end: en || st, mark: r[iGanttChk]===true || r[iGanttChk]==='TRUE' };
  }).filter(e => e.mark && e.start);

  if (eventos.length === 0) {
    const maxRowG = gantt.getLastRow(), maxColG = gantt.getLastColumn();
    if (maxRowG >= 4) gantt.getRange(4,1,maxRowG-3,maxColG).clearContent();
    return;
  }

  let minDate = eventos[0].start, maxDate = eventos[0].end;
  eventos.forEach(e => { if (e.start < minDate) minDate = e.start; if (e.end > maxDate) maxDate = e.end; });

  const oneDay = 24*60*60*1000;
  minDate = new Date(minDate.getTime() - 3*oneDay);
  maxDate = new Date(maxDate.getTime() + 3*oneDay);

  const totalDays = Math.round((maxDate - minDate)/oneDay) + 1;
  const row1 = [], row2 = [], row3 = [];
  for (let i=0; i<totalDays; i++) {
    const d = new Date(minDate.getTime() + i*oneDay);
    const dd = String(d.getDate()).padStart(2,'0');
    const mm = String(d.getMonth()+1).padStart(2,'0');
    const yyyy = d.getFullYear();
    const fmt = `${dd}/${mm}/${yyyy}`;
    row1.push(fmt);
    row2.push(fmt);
    const letter = ['D','S','T','Q','Q','S','S'][d.getDay()];
    row3.push(letter);
  }

  const maxColG = gantt.getLastColumn(), maxRowG = gantt.getLastRow();
  if (maxColG>=4) gantt.getRange(1,4,3,maxColG-3).clearContent();
  if (maxRowG>=4) gantt.getRange(4,1,maxRowG-3,maxColG).clearContent();

  gantt.getRange(1,4,1,totalDays).setValues([row1]);
  gantt.getRange(2,4,1,totalDays).setValues([row2]);
  gantt.getRange(3,4,1,totalDays).setValues([row3]);

  const outAC = eventos.map(e => [e.title, e.start, e.end]);
  gantt.getRange(4,1,outAC.length,3).setValues(outAC);

  const hdrDates = row1.map(s => {
    const [dd,mm,yyyy] = s.split('/');
    return new Date(+yyyy,+mm-1,+dd);
  });

  eventos.forEach((e, idx) => {
    const rowG = 4 + idx;
    const endDate = e.end || e.start;
    hdrDates.forEach((d,j) => {
      if (e.start <= d && d <= endDate) {
        gantt.getRange(rowG, 4 + j).setValue('█');
      }
    });
  });
}

// ======================
// CRIAR NOVA AGENDA (Arquivar em D, Título em E)
// ======================
function criarNovaAgenda() {
  const ui = SpreadsheetApp.getUi();
  const ss = getSpreadsheet();

  // Cabeçalho com "Arquivar" em D (coluna 4) e "Título" em E (coluna 5)
  const agendaHeader = [
    'MÊS', 'Nº / DIA SEMANA', 'PRAZO', 'Arquivar', // A..D
    'Título','Descrição',                          // E..F
    'Data e Hora de Início','Data e Hora de Fim',  // G..H
    'Local',                                       // I
    'Convidados (e-mail, e-mail, ...)',            // J
    'Cor','PROJETO','ETAPA','STATUS','DEMANDANTE', // K..O
    'CONTATOS','AÇÕES RECOMENDADAS','PRIORIDADE',  // P..R
    'AÇÕES REALIZADAS','OBSERVAÇÕES','ID',         // S..U
    'Modificado na Agenda','Registro na Agenda',   // V..W
    'Alterado na Planilha','Gantt'                 // X..Y
  ];

  const ganttHeader = ['Título','Data e Hora de Início','Data e Hora de Fim'];

  function ensureEmptySheet(name, header) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.getRange(1,1,1,header.length).setValues([header]);
    }
    return sh;
  }

  const shAgenda = ss.getSheetByName(SHEET_EVENTS);
  if (shAgenda) {
    const existing = shAgenda.getRange(1,1,1,agendaHeader.length).getValues()[0];
    if (JSON.stringify(existing) === JSON.stringify(agendaHeader)) {
      ui.alert('Nada a ser feito: a AGENDA já está criada com este cabeçalho.');
      return;
    }
  }

  ensureEmptySheet(SHEET_EVENTS, agendaHeader);
  ensureEmptySheet('Gantt',        ganttHeader);
  ensureEmptySheet(SHEET_ID,       ['']); // preencha B17 manualmente

  ui.alert('Nova agenda criada com sucesso (Arquivar em D)!');
}

