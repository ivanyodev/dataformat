/**
 * Excel & JSON Web Studio
 * Architecture: Store + IIFE Modules + App Orchestrator
 * Principles: SRP, DRY, encapsulation, low coupling
 */
(function (global) {
  'use strict';

  /* ========================================================================
   * CONSTANTS
   * ======================================================================== */
  const CONST = Object.freeze({
    ROW_HEIGHT: 28,
    VIRTUAL_BUFFER: 12,
    DEFAULT_COL_WIDTH: 120,
    MIN_COL_WIDTH: 50,
    MIN_ROW_HEADER_WIDTH: 40,
    BLANK_COLS: 26,
    BLANK_ROWS: 100,
    FILTER_CHUNK: 4000
  });

  /* ========================================================================
   * EVENT BUS (pub/sub for decoupled module communication)
   * ======================================================================== */
  const EventBus = (function () {
    const listeners = new Map();

    function on(event, handler) {
      if (!listeners.has(event)) listeners.set(event, new Set());
      listeners.get(event).add(handler);
      return function off() {
        listeners.get(event)?.delete(handler);
      };
    }

    function emit(event, payload) {
      const set = listeners.get(event);
      if (!set) return;
      set.forEach(function (fn) {
        try { fn(payload); } catch (err) { console.error('[EventBus]', event, err); }
      });
    }

    return Object.freeze({ on: on, emit: emit });
  })();

  /* ========================================================================
   * STORE — centralized application state
   * ======================================================================== */
  const Store = (function () {
    const state = {
      workbook: null,
      sheetName: '',
      activeFileId: null,
      fileIdCounter: 0,
      files: [],
      cellStyles: {},
      columnWidths: {},
      activeFilters: {},
      filteredRowIndices: [],
      selection: {
        cols: new Set(),
        rows: new Set(),
        cells: new Set(),
        range: null,
        anchorR: null,
        anchorC: null
      },
      editing: false,
      contextTarget: { type: null, index: null },
      currentFilterCol: null
    };

    function get(key) {
      if (key === undefined) return state;
      return state[key];
    }

    function set(key, value) {
      state[key] = value;
      EventBus.emit('store:' + key, value);
      return value;
    }

    function patch(partial) {
      Object.keys(partial).forEach(function (k) {
        state[k] = partial[k];
        EventBus.emit('store:' + k, partial[k]);
      });
    }

    function clearSelection() {
      state.selection.cols.clear();
      state.selection.rows.clear();
      state.selection.cells.clear();
      state.selection.range = null;
      state.selection.anchorR = null;
      state.selection.anchorC = null;
    }

    function hasSelection() {
      const s = state.selection;
      return s.cols.size > 0 || s.rows.size > 0 || s.cells.size > 0 || s.range !== null;
    }

    function cellKey(r, c) {
      return r + ',' + c;
    }

    function isCellSelected(R, C) {
      const s = state.selection;
      if (s.cols.has(C) || s.rows.has(R) || s.cells.has(cellKey(R, C))) return true;
      if (s.range) {
        const r = s.range;
        return R >= r.minR && R <= r.maxR && C >= r.minC && C <= r.maxC;
      }
      return false;
    }

    function getSheet() {
      if (!state.workbook || !state.sheetName) return null;
      return state.workbook.Sheets[state.sheetName] || null;
    }

    function getRange() {
      const sheet = getSheet();
      if (!sheet) return null;
      return XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
    }

    return Object.freeze({
      get: get,
      set: set,
      patch: patch,
      clearSelection: clearSelection,
      hasSelection: hasSelection,
      cellKey: cellKey,
      isCellSelected: isCellSelected,
      getSheet: getSheet,
      getRange: getRange
    });
  })();

  /* ========================================================================
   * UTILS
   * ======================================================================== */
  const Utils = (function () {
    const colLetterCache = [];

    function getColLetter(c) {
      if (colLetterCache[c] !== undefined) return colLetterCache[c];
      let n = c;
      let s = '';
      while (n >= 0) {
        s = String.fromCharCode((n % 26) + 65) + s;
        n = Math.floor(n / 26) - 1;
      }
      colLetterCache[c] = s;
      return s;
    }

    function cellAddr(R, C) {
      return getColLetter(C) + (R + 1);
    }

    function encodeCell(r, c) {
      return XLSX.utils.encode_cell({ r: r, c: c });
    }

    function decodeRange(ref) {
      return XLSX.utils.decode_range(ref || 'A1:A1');
    }

    function escapeHtml(str) {
      return String(str ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
    }

    function escapeXml(str) {
      return String(str ?? '')
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&apos;');
    }

    function escapeCsv(val) {
      const str = String(val ?? '');
      if (/[",\n\r]/.test(str)) return '"' + str.replace(/"/g, '""') + '"';
      return str;
    }

    function sanitizeXmlTag(name) {
      let tag = String(name ?? 'col').trim().replace(/[^a-zA-Z0-9_\-.]/g, '_');
      if (!/^[a-zA-Z_]/.test(tag)) tag = 'col_' + tag;
      return tag || 'col';
    }

    function getBaseFileName(filename) {
      let base = String(filename || 'Datos').replace(/\.[^/.]+$/, '');
      base = base.replace(/[\\\/\?\*\[\]:]/g, '_').trim() || 'Datos';
      if (base.length > 31) base = base.substring(0, 31);
      return base;
    }

    function makeUniqueSheetName(wb, desiredName) {
      let name = desiredName;
      let i = 1;
      while (wb.SheetNames.includes(name)) {
        const suffix = '_' + i;
        name = desiredName.substring(0, 31 - suffix.length) + suffix;
        i++;
      }
      return name;
    }

    function detectFileType(filename) {
      const n = String(filename).toLowerCase();
      if (n.endsWith('.json')) return 'json';
      if (n.endsWith('.csv')) return 'csv';
      if (n.endsWith('.txt')) return 'txt';
      if (n.endsWith('.sql')) return 'sql';
      if (n.endsWith('.xlsx')) return 'xlsx';
      if (n.endsWith('.xls')) return 'xls';
      if (n.endsWith('.xml')) return 'xml';
      if (n.endsWith('.html') || n.endsWith('.htm')) return 'html';
      return 'file';
    }

    function iconForType(type) {
      const map = {
        json: { icon: 'data_object', cls: 'tree__icon--json' },
        csv: { icon: 'table_view', cls: 'tree__icon--csv' },
        txt: { icon: 'description', cls: 'tree__icon--txt' },
        sql: { icon: 'storage', cls: 'tree__icon--sql' },
        xlsx: { icon: 'grid_on', cls: 'tree__icon--xlsx' },
        xls: { icon: 'grid_on', cls: 'tree__icon--xls' },
        xml: { icon: 'code', cls: 'tree__icon--xml' },
        html: { icon: 'language', cls: 'tree__icon--html' },
        file: { icon: 'insert_drive_file', cls: '' }
      };
      return map[type] || map.file;
    }

    function cloneWorkbook(wb) {
      try {
        return structuredClone(wb);
      } catch (e) {
        return JSON.parse(JSON.stringify(wb));
      }
    }

    function contrastingText(hexColor) {
      let hex = hexColor.replace('#', '');
      if (hex.length === 3) hex = hex.split('').map(function (c) { return c + c; }).join('');
      const r = parseInt(hex.substring(0, 2), 16);
      const g = parseInt(hex.substring(2, 4), 16);
      const b = parseInt(hex.substring(4, 6), 16);
      const yiq = (r * 299 + g * 587 + b * 114) / 1000;
      return yiq >= 128 ? '#000000' : '#ffffff';
    }

    function downloadBlob(content, filename, mimeType) {
      const blob = new Blob([content], { type: mimeType });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }

    function isPlainObject(val) {
      return val !== null && typeof val === 'object' && !Array.isArray(val);
    }

    return Object.freeze({
      getColLetter: getColLetter,
      cellAddr: cellAddr,
      encodeCell: encodeCell,
      decodeRange: decodeRange,
      escapeHtml: escapeHtml,
      escapeXml: escapeXml,
      escapeCsv: escapeCsv,
      sanitizeXmlTag: sanitizeXmlTag,
      getBaseFileName: getBaseFileName,
      makeUniqueSheetName: makeUniqueSheetName,
      detectFileType: detectFileType,
      iconForType: iconForType,
      cloneWorkbook: cloneWorkbook,
      contrastingText: contrastingText,
      downloadBlob: downloadBlob,
      isPlainObject: isPlainObject
    });
  })();

  /* ========================================================================
   * PARSERS — JSON / CSV / TXT / SQL
   * ======================================================================== */
  const Parsers = (function () {

    function flattenObject(obj, prefix, result) {
      prefix = prefix || '';
      result = result || {};
      Object.keys(obj).forEach(function (key) {
        const value = obj[key];
        const newKey = prefix ? prefix + '.' + key : key;
        if (value === null || value === undefined) {
          result[newKey] = value;
        } else if (Array.isArray(value)) {
          if (value.some(function (item) { return Utils.isPlainObject(item) || Array.isArray(item); })) {
            throw new Error('La propiedad "' + newKey + '" contiene un array de objetos o arrays anidados.');
          }
          result[newKey] = JSON.stringify(value);
        } else if (Utils.isPlainObject(value)) {
          flattenObject(value, newKey, result);
        } else if (['string', 'number', 'boolean'].indexOf(typeof value) !== -1) {
          result[newKey] = value;
        } else {
          throw new Error('La propiedad "' + newKey + '" tiene un tipo no soportado (' + typeof value + ').');
        }
      });
      return result;
    }

    function validateAndNormalizeJson(raw) {
      let data;
      try {
        data = JSON.parse(raw);
      } catch (e) {
        throw new Error('El archivo no es un JSON válido. Verifica la sintaxis.');
      }
      if (!Array.isArray(data)) {
        if (Utils.isPlainObject(data)) data = [data];
        else throw new Error('El JSON debe ser un array de objetos o un objeto.');
      }
      if (data.length === 0) throw new Error('El JSON está vacío.');
      const normalized = [];
      for (let i = 0; i < data.length; i++) {
        if (!Utils.isPlainObject(data[i])) {
          throw new Error('El elemento en la posición ' + i + ' no es un objeto.');
        }
        normalized.push(flattenObject(data[i]));
      }
      return normalized;
    }

    function detectDelimiter(text, candidates) {
      const lines = text.split(/\r\n|\n|\r/).filter(function (l) { return l.trim().length > 0; }).slice(0, 20);
      if (!lines.length) return null;
      let best = null;
      let bestScore = -1;
      candidates.forEach(function (d) {
        const counts = lines.map(function (line) {
          let count = 0;
          let inQuotes = false;
          for (let i = 0; i < line.length; i++) {
            const ch = line[i];
            if (ch === '"') {
              if (inQuotes && line[i + 1] === '"') { i++; continue; }
              inQuotes = !inQuotes;
            } else if (ch === d && !inQuotes) {
              count++;
            }
          }
          return count;
        });
        const first = counts[0];
        if (first === 0) return;
        if (counts.every(function (c) { return c === first; }) && first > bestScore) {
          bestScore = first;
          best = d;
        }
      });
      return best;
    }

    function parseDelimitedLine(line, delimiter) {
      const result = [];
      let current = '';
      let inQuotes = false;
      for (let i = 0; i < line.length; i++) {
        const ch = line[i];
        if (ch === '"') {
          if (inQuotes && line[i + 1] === '"') { current += '"'; i++; }
          else inQuotes = !inQuotes;
        } else if (ch === delimiter && !inQuotes) {
          result.push(current);
          current = '';
        } else {
          current += ch;
        }
      }
      result.push(current);
      if (inQuotes) throw new Error('Comillas desparejadas detectadas.');
      return result;
    }

    function validateAndParseCsv(raw) {
      const text = raw.replace(/^\uFEFF/, '');
      const lines = text.split(/\r\n|\n|\r/).filter(function (l) { return l.trim().length > 0; });
      if (!lines.length) throw new Error('El archivo CSV está vacío.');
      const delimiter = detectDelimiter(text, [',', ';', '\t', '|']);
      if (!delimiter) throw new Error('No se pudo detectar un delimitador válido en el CSV.');
      const rows = [];
      let expectedCols = null;
      for (let i = 0; i < lines.length; i++) {
        const fields = parseDelimitedLine(lines[i], delimiter);
        if (expectedCols === null) {
          expectedCols = fields.length;
          if (!expectedCols) throw new Error('La primera línea del CSV no contiene columnas.');
        } else if (fields.length !== expectedCols) {
          throw new Error('Línea ' + (i + 1) + ': tiene ' + fields.length + ' columnas, se esperaban ' + expectedCols + '.');
        }
        rows.push(fields);
      }
      return { rows: rows, delimiter: delimiter };
    }

    function validateAndParseTxt(raw) {
      const text = raw.replace(/^\uFEFF/, '');
      if (text.indexOf('\uFFFD') !== -1) {
        throw new Error('El archivo TXT no tiene codificación UTF-8 válida.');
      }
      const lines = text.split(/\r\n|\n|\r/).filter(function (l) { return l.trim().length > 0; });
      if (!lines.length) throw new Error('El archivo TXT está vacío.');
      const delimiter = detectDelimiter(text, ['\t', '|']);
      if (!delimiter) throw new Error('El TXT debe estar delimitado por tabuladores o pipes (|).');
      const rows = [];
      let expectedCols = null;
      for (let i = 0; i < lines.length; i++) {
        const fields = parseDelimitedLine(lines[i], delimiter);
        if (expectedCols === null) {
          expectedCols = fields.length;
          if (expectedCols < 2) throw new Error('La primera línea no tiene suficientes columnas.');
        } else if (fields.length !== expectedCols) {
          throw new Error('Línea ' + (i + 1) + ': columnas inconsistentes.');
        }
        rows.push(fields);
      }
      return { rows: rows, delimiter: delimiter };
    }

    /* --- SQL helpers --- */
    function extractBalancedParens(str, openIdx) {
      if (str[openIdx] !== '(') return null;
      let depth = 0;
      let inQuote = null;
      for (let i = openIdx; i < str.length; i++) {
        const ch = str[i];
        if (inQuote) {
          if (ch === inQuote) {
            if (inQuote === "'" && str[i + 1] === "'") { i++; continue; }
            if (str[i - 1] === '\\') continue;
            inQuote = null;
          }
          continue;
        }
        if (ch === "'" || ch === '"' || ch === '`') { inQuote = ch; continue; }
        if (ch === '(') depth++;
        else if (ch === ')') {
          depth--;
          if (depth === 0) return str.substring(openIdx + 1, i);
        }
      }
      return null;
    }

    function splitSqlList(str) {
      const parts = [];
      let current = '';
      let depth = 0;
      let inQuote = null;
      for (let i = 0; i < str.length; i++) {
        const ch = str[i];
        if (inQuote) {
          current += ch;
          if (ch === inQuote && str[i - 1] !== '\\') {
            if (inQuote === "'" && str[i + 1] === "'") { current += str[++i]; }
            else inQuote = null;
          }
          continue;
        }
        if (ch === "'" || ch === '"' || ch === '`') { inQuote = ch; current += ch; continue; }
        if (ch === '(') { depth++; current += ch; continue; }
        if (ch === ')') { depth--; current += ch; continue; }
        if (ch === ',' && depth === 0) { parts.push(current); current = ''; continue; }
        current += ch;
      }
      if (current.trim()) parts.push(current);
      return parts;
    }

    function parseCreateTableColumns(body) {
      const columns = [];
      splitSqlList(body).forEach(function (part) {
        const trimmed = part.trim();
        if (!trimmed) return;
        if (/^(PRIMARY|UNIQUE|KEY|INDEX|CONSTRAINT|FOREIGN|CHECK|FULLTEXT|SPATIAL)\b/i.test(trimmed)) return;
        const m = trimmed.match(/^[`"'\[\]]?([a-zA-Z0-9_]+)[`"'\[\]]?\s+([a-zA-Z0-9_]+(?:\s*\([^)]*\))?)/i);
        if (m) {
          columns.push({
            name: m[1],
            type: (m[2] || 'TEXT').toUpperCase().replace(/\s+/g, '')
          });
        }
      });
      return columns;
    }

    function parseSqlLiteral(token) {
      if (token === '' || /^null$/i.test(token)) return null;
      if (/^true$/i.test(token)) return true;
      if (/^false$/i.test(token)) return false;
      if (/^[+-]?\d+(\.\d+)?([eE][+-]?\d+)?$/.test(token)) {
        const n = Number(token);
        return Number.isFinite(n) ? n : token;
      }
      return token;
    }

    function parseInsertValues(sql, startIndex) {
      const rows = [];
      let i = startIndex;
      const len = sql.length;
      while (i < len && /\s/.test(sql[i])) i++;
      while (i < len) {
        while (i < len && /\s/.test(sql[i])) i++;
        if (sql[i] !== '(') break;
        i++;
        const values = [];
        let current = '';
        let inQuote = null;
        let depth = 0;
        while (i < len) {
          const ch = sql[i];
          if (inQuote) {
            if (ch === inQuote) {
              if (sql[i + 1] === inQuote) { current += inQuote; i += 2; continue; }
              inQuote = null; i++; continue;
            }
            if (ch === '\\' && i + 1 < len) { current += sql[i + 1]; i += 2; continue; }
            current += ch; i++; continue;
          }
          if (ch === "'" || ch === '"') { inQuote = ch; i++; continue; }
          if (ch === '(') { depth++; current += ch; i++; continue; }
          if (ch === ')') {
            if (depth === 0) {
              values.push(parseSqlLiteral(current.trim()));
              current = '';
              i++;
              rows.push(values);
              break;
            }
            depth--; current += ch; i++; continue;
          }
          if (ch === ',' && depth === 0) {
            values.push(parseSqlLiteral(current.trim()));
            current = '';
            i++;
            continue;
          }
          current += ch; i++;
        }
        while (i < len && /\s/.test(sql[i])) i++;
        if (sql[i] === ',') { i++; continue; }
        break;
      }
      return { rows: rows, endIndex: i };
    }

    function coerceBySqlType(val, sqlType) {
      if (val === null || val === undefined) return null;
      const t = String(sqlType).toUpperCase();
      if (/INT|SERIAL|BIGINT|SMALLINT|TINYINT|MEDIUMINT|YEAR/.test(t)) {
        if (typeof val === 'number') return Math.trunc(val);
        const n = Number(val);
        return Number.isFinite(n) ? Math.trunc(n) : val;
      }
      if (/DECIMAL|NUMERIC|FLOAT|DOUBLE|REAL|MONEY/.test(t)) {
        if (typeof val === 'number') return val;
        const n = Number(val);
        return Number.isFinite(n) ? n : val;
      }
      if (/BOOL|BIT/.test(t)) {
        if (typeof val === 'boolean') return val;
        if (val === 1 || val === '1' || /^true$/i.test(String(val))) return true;
        if (val === 0 || val === '0' || /^false$/i.test(String(val))) return false;
        return val;
      }
      return typeof val === 'string' ? val : String(val);
    }

    function coerceSqlRow(rawRow, tableColumns, insertCols) {
      let ordered = rawRow;
      if (insertCols && insertCols.length && tableColumns.length) {
        const byName = {};
        insertCols.forEach(function (colName, idx) {
          byName[colName.toLowerCase()] = rawRow[idx];
        });
        ordered = tableColumns.map(function (c) {
          const v = byName[c.name.toLowerCase()];
          return v === undefined ? null : v;
        });
      }
      return ordered.map(function (val, idx) {
        if (val === null || val === undefined) return null;
        const colType = (tableColumns[idx] && tableColumns[idx].type) || 'TEXT';
        return coerceBySqlType(val, colType);
      });
    }

    function parseSqlDump(raw) {
      if (!raw || typeof raw !== 'string') throw new Error('El archivo SQL está vacío.');
      let sql = raw.replace(/^\uFEFF/, '');
      sql = sql.replace(/\/\*[\s\S]*?\*\//g, ' ');
      sql = sql.replace(/^\s*--.*$/gm, ' ');
      sql = sql.replace(/^\s*#.*$/gm, ' ');

      const tables = new Map();

      const createStartRe = /CREATE\s+TABLE\s+(?:IF\s+NOT\s+EXISTS\s+)?[`"'\[\]]?([a-zA-Z0-9_]+)[`"'\[\]]?\s*\(/gi;
      let createStart;
      while ((createStart = createStartRe.exec(sql)) !== null) {
        const tableName = createStart[1];
        const openParenIdx = createStartRe.lastIndex - 1;
        const body = extractBalancedParens(sql, openParenIdx);
        if (body === null) continue;
        createStartRe.lastIndex = openParenIdx + body.length + 2;
        const columns = parseCreateTableColumns(body);
        if (!columns.length) continue;
        const key = tableName.toLowerCase();
        if (!tables.has(key)) {
          tables.set(key, { name: tableName, columns: columns, rows: [] });
        } else {
          const t = tables.get(key);
          if (!t.columns.length) t.columns = columns;
          t.name = tableName;
        }
      }

      const insertRe = /INSERT\s+INTO\s+[`"'\[\]]?([a-zA-Z0-9_]+)[`"'\[\]]?\s*(?:\(([^)]*)\))?\s*VALUES\s*/gi;
      let insertMatch;
      while ((insertMatch = insertRe.exec(sql)) !== null) {
        const tableName = insertMatch[1];
        const colListRaw = insertMatch[2];
        const key = tableName.toLowerCase();
        if (!tables.has(key)) tables.set(key, { name: tableName, columns: [], rows: [] });
        const table = tables.get(key);
        let insertCols = null;
        if (colListRaw) {
          insertCols = colListRaw.split(',').map(function (c) {
            return c.trim().replace(/^[`"'\[\]]|[`"'\[\]]$/g, '');
          }).filter(Boolean);
          if (!table.columns.length) {
            table.columns = insertCols.map(function (n) { return { name: n, type: 'TEXT' }; });
          }
        }
        const parsed = parseInsertValues(sql, insertRe.lastIndex);
        insertRe.lastIndex = parsed.endIndex;
        parsed.rows.forEach(function (rawRow) {
          table.rows.push(coerceSqlRow(rawRow, table.columns, insertCols));
        });
      }

      if (tables.size === 0) {
        throw new Error('No se encontraron tablas (CREATE TABLE / INSERT INTO) en el archivo SQL.');
      }

      const result = Array.from(tables.values());
      result.sort(function (a, b) {
        const ac = a.columns.length ? 0 : 1;
        const bc = b.columns.length ? 0 : 1;
        if (ac !== bc) return ac - bc;
        return a.name.localeCompare(b.name);
      });

      result.forEach(function (t) {
        if (!t.columns.length && t.rows.length) {
          const maxCols = Math.max.apply(null, t.rows.map(function (r) { return r.length; }));
          t.columns = Array.from({ length: maxCols }, function (_, i) {
            return { name: 'col' + (i + 1), type: 'TEXT' };
          });
        }
        const n = t.columns.length;
        t.rows = t.rows.map(function (r) {
          const out = r.slice(0, n);
          while (out.length < n) out.push(null);
          return out;
        });
      });

      return result;
    }

    function sqlTableToWorksheet(table) {
      const aoa = [table.columns.map(function (c) { return c.name; })];
      table.rows.forEach(function (row) {
        aoa.push(row.map(function (v) { return v === null || v === undefined ? '' : v; }));
      });
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      const range = Utils.decodeRange(ws['!ref']);
      for (let R = 1; R <= range.e.r; R++) {
        for (let C = 0; C <= range.e.c; C++) {
          const addr = Utils.encodeCell(R, C);
          const cell = ws[addr];
          if (!cell || cell.v === '' || cell.v === null || cell.v === undefined) continue;
          const colType = (table.columns[C] && table.columns[C].type) || 'TEXT';
          const t = String(colType).toUpperCase();
          if (/INT|SERIAL|BIGINT|SMALLINT|TINYINT|MEDIUMINT|YEAR|DECIMAL|NUMERIC|FLOAT|DOUBLE|REAL|MONEY/.test(t)) {
            const n = Number(cell.v);
            if (Number.isFinite(n)) { cell.v = n; cell.t = 'n'; }
          } else if (/BOOL|BIT/.test(t)) {
            if (typeof cell.v === 'boolean') cell.t = 'b';
            else if (cell.v === 1 || cell.v === '1' || cell.v === true) { cell.v = true; cell.t = 'b'; }
            else if (cell.v === 0 || cell.v === '0' || cell.v === false) { cell.v = false; cell.t = 'b'; }
          } else {
            cell.t = 's';
            cell.v = String(cell.v);
          }
        }
      }
      return ws;
    }

    return Object.freeze({
      validateAndNormalizeJson: validateAndNormalizeJson,
      validateAndParseCsv: validateAndParseCsv,
      validateAndParseTxt: validateAndParseTxt,
      parseSqlDump: parseSqlDump,
      sqlTableToWorksheet: sqlTableToWorksheet
    });
  })();

  /* ========================================================================
   * TOAST MODULE
   * ======================================================================== */
  const Toast = (function () {
    let hideTimer = null;

    function show(message, type) {
      type = type || 'success';
      const el = document.getElementById('toast');
      if (!el) return;
      el.textContent = message;
      el.classList.remove('toast--visible', 'toast--danger', 'toast--warning');
      if (type === 'danger') el.classList.add('toast--danger');
      else if (type === 'warning') el.classList.add('toast--warning');
      void el.offsetWidth;
      el.classList.add('toast--visible');
      clearTimeout(hideTimer);
      hideTimer = setTimeout(function () {
        el.classList.remove('toast--visible', 'toast--danger', 'toast--warning');
      }, type === 'danger' || type === 'warning' ? 4500 : 2200);
    }

    return Object.freeze({ show: show });
  })();

  /* ========================================================================
   * LOADER MODULE
   * ======================================================================== */
  const Loader = (function () {
    function show(msg) {
      const el = document.getElementById('globalLoader');
      const text = document.getElementById('loaderStatusText');
      if (text) text.textContent = msg || 'Cargando archivo...';
      updateProgress(0);
      if (el) el.classList.add('loader--visible');
    }

    function updateProgress(percent) {
      const p = Math.min(100, Math.max(0, Math.round(percent)));
      const bar = document.getElementById('progressBar');
      const pct = document.getElementById('loaderPercentage');
      if (bar) bar.style.width = p + '%';
      if (pct) pct.textContent = p + '%';
    }

    function hide() {
      const el = document.getElementById('globalLoader');
      if (el) el.classList.remove('loader--visible');
    }

    return Object.freeze({ show: show, updateProgress: updateProgress, hide: hide });
  })();


  /* ========================================================================
   * FILTER ENGINE
   * ======================================================================== */
  const FilterEngine = (function () {
    let token = 0;
    let debounceTimer = null;

    function rebuildSync() {
      const indices = [];
      const sheet = Store.getSheet();
      if (!sheet) { Store.set('filteredRowIndices', indices); return; }

      const range = Store.getRange();
      const headerRow = range.s.r;
      const startR = headerRow + 1;
      const endR = range.e.r;
      const startC = range.s.c;
      const endC = range.e.c;
      const globalQuery = (document.getElementById('globalFilterInput')?.value || '').trim().toLowerCase();
      const activeFilters = Store.get('activeFilters');
      const hasColFilters = Object.keys(activeFilters).length > 0;

      if (!globalQuery && !hasColFilters) {
        const total = endR - startR + 1;
        for (let i = 0; i < total; i++) indices.push(startR + i);
        Store.set('filteredRowIndices', indices);
        return;
      }

      const filterCols = hasColFilters
        ? Object.keys(activeFilters).map(function (k) {
            return { C: parseInt(k, 10), allowed: activeFilters[k], letter: Utils.getColLetter(parseInt(k, 10)) };
          })
        : [];

      const colLetters = [];
      if (globalQuery) {
        for (let C = startC; C <= endC; C++) colLetters[C] = Utils.getColLetter(C);
      }

      for (let R = startR; R <= endR; R++) {
        let show = true;
        if (hasColFilters) {
          for (let fi = 0; fi < filterCols.length; fi++) {
            const f = filterCols[fi];
            const cell = sheet[f.letter + (R + 1)];
            const val = cell && cell.v !== undefined && cell.v !== '' ? String(cell.v) : '(Vacío)';
            if (!f.allowed.has(val)) { show = false; break; }
          }
        }
        if (show && globalQuery) {
          let match = false;
          for (let C = startC; C <= endC; C++) {
            const cell = sheet[colLetters[C] + (R + 1)];
            if (!cell || cell.v === undefined || cell.v === '') continue;
            const val = typeof cell.v === 'string' ? cell.v.toLowerCase() : String(cell.v).toLowerCase();
            if (val.indexOf(globalQuery) !== -1) { match = true; break; }
          }
          if (!match) show = false;
        }
        if (show) indices.push(R);
      }
      Store.set('filteredRowIndices', indices);
    }

    function apply() {
      rebuildSync();
      EventBus.emit('filters:applied');
    }

    function scheduleApply() {
      if (debounceTimer) clearTimeout(debounceTimer);
      let rowCount = 0;
      const range = Store.getRange();
      if (range) rowCount = Math.max(0, range.e.r - range.s.r);
      const delay = rowCount > 15000 ? 280 : (rowCount > 5000 ? 150 : 60);
      debounceTimer = setTimeout(function () {
        debounceTimer = null;
        apply();
      }, delay);
    }

    return Object.freeze({ rebuildSync: rebuildSync, apply: apply, scheduleApply: scheduleApply });
  })();

  /* ========================================================================
   * SPREADSHEET MODULE
   * ======================================================================== */
  const Spreadsheet = (function () {
    let virtualScrollBound = false;
    let windowMouseUpBound = false;
    let isMouseDown = false;
    let anchorTd = null;
    let editingTd = null;

    function getRowHeaderWidth() {
      const range = Store.getRange();
      if (!range) return CONST.MIN_ROW_HEADER_WIDTH;
      const digits = String(range.e.r + 1).length;
      return Math.max(CONST.MIN_ROW_HEADER_WIDTH, digits * 9 + 20);
    }

    function ensureColumnWidths(sheetName) {
      const widths = Store.get('columnWidths');
      if (!widths[sheetName]) widths[sheetName] = {};
      const sheet = Store.get('workbook').Sheets[sheetName];
      const range = Utils.decodeRange(sheet['!ref']);
      const map = widths[sheetName];
      for (let C = range.s.c; C <= range.e.c; C++) {
        if (map[C] != null) continue;
        let maxLen = 8;
        const headerCell = sheet[Utils.encodeCell(range.s.r, C)];
        if (headerCell && headerCell.v !== undefined) maxLen = Math.max(maxLen, String(headerCell.v).length);
        const sampleEnd = Math.min(range.e.r, range.s.r + 80);
        for (let R = range.s.r + 1; R <= sampleEnd; R++) {
          const cell = sheet[Utils.encodeCell(R, C)];
          if (cell && cell.v !== undefined) maxLen = Math.max(maxLen, String(cell.v).length);
        }
        map[C] = Math.min(320, Math.max(CONST.MIN_COL_WIDTH, maxLen * 8 + 24));
      }
      return map;
    }

    function getColWidth(colIndex) {
      const widths = Store.get('columnWidths')[Store.get('sheetName')] || {};
      return widths[colIndex] != null ? widths[colIndex] : CONST.DEFAULT_COL_WIDTH;
    }

    function renderDropzone() {
      const container = document.getElementById('tableContainer');
      if (!container) return;
      container.innerHTML =
        '<div class="dropzone" id="dropzoneContainer">' +
          '<div class="dropzone__card" id="dropzone">' +
            '<div class="dropzone__icon"><span class="material-icons" style="font-size:3.2rem;">cloud_upload</span></div>' +
            '<div class="dropzone__title">Arrastra tu archivo aquí para comenzar</div>' +
            '<div class="dropzone__desc">o haz clic para explorar en tu computadora</div>' +
            '<div class="dropzone__formats">' +
              '<span class="dropzone__tag">.XLSX</span><span class="dropzone__tag">.XLS</span>' +
              '<span class="dropzone__tag">.JSON</span><span class="dropzone__tag">.CSV</span>' +
              '<span class="dropzone__tag">.TXT</span><span class="dropzone__tag">.SQL</span>' +
            '</div>' +
          '</div>' +
        '</div>';
      const dz = document.getElementById('dropzoneContainer');
      if (dz) dz.addEventListener('click', function () {
        document.getElementById('excelFile')?.click();
      });
    }

    function updateVirtualRows() {
      const tbody = document.getElementById('virtualTbody');
      const container = document.getElementById('tableContainer');
      if (!tbody || !container) return;

      const sheet = Store.getSheet();
      if (!sheet) return;
      const range = Store.getRange();
      const colCount = range.e.c - range.s.c + 1;
      const filtered = Store.get('filteredRowIndices');
      const total = filtered.length;

      const scrollTop = container.scrollTop;
      const viewHeight = container.clientHeight || 400;
      const visibleCount = Math.ceil(viewHeight / CONST.ROW_HEIGHT) + CONST.VIRTUAL_BUFFER * 2;
      let start = Math.floor(scrollTop / CONST.ROW_HEIGHT) - CONST.VIRTUAL_BUFFER;
      if (start < 0) start = 0;
      let end = start + visibleCount;
      if (end > total) end = total;

      const topPad = start * CONST.ROW_HEIGHT;
      const bottomPad = Math.max(0, (total - end) * CONST.ROW_HEIGHT);
      const styles = Store.get('cellStyles')[Store.get('sheetName')] || {};
      let html = '';

      if (topPad > 0) {
        html += '<tr aria-hidden="true"><td colspan="' + (colCount + 1) + '" style="height:' + topPad + 'px;padding:0;border:none;background:transparent;"></td></tr>';
      }

      for (let i = start; i < end; i++) {
        const R = filtered[i];
        const rowSel = Store.get('selection').rows.has(R) ? ' spreadsheet__row-header--selected' : '';
        html += '<tr data-r="' + R + '"><th class="spreadsheet__row-header' + rowSel + '" data-r="' + R + '">' + (R + 1) + '</th>';
        for (let C = range.s.c; C <= range.e.c; C++) {
          const addr = Utils.encodeCell(R, C);
          const cell = sheet[addr];
          const val = cell ? (cell.v !== undefined ? cell.v : '') : '';
          let styleAttr = '';
          if (styles[addr]) styleAttr = 'style="background-color:' + styles[addr].bg + ';color:' + styles[addr].font + ';"';
          const sel = Store.isCellSelected(R, C) ? ' spreadsheet__cell--selected' : '';
          const display = Utils.escapeHtml(String(val));
          html += '<td class="spreadsheet__cell' + sel + '" data-r="' + R + '" data-c="' + C +
            '" data-cell="' + addr + '" data-val="' + String(val).replace(/"/g, '&quot;') + '" ' + styleAttr + '>' + display + '</td>';
        }
        html += '</tr>';
      }

      if (bottomPad > 0) {
        html += '<tr aria-hidden="true"><td colspan="' + (colCount + 1) + '" style="height:' + bottomPad + 'px;padding:0;border:none;background:transparent;"></td></tr>';
      }
      if (total === 0) {
        html = '<tr><td class="spreadsheet__empty-msg" colspan="' + (colCount + 1) + '">Sin filas para mostrar</td></tr>';
      }
      tbody.innerHTML = html;
    }

    function bindVirtualScroll() {
      const container = document.getElementById('tableContainer');
      if (!container || virtualScrollBound) return;
      let ticking = false;
      container.addEventListener('scroll', function () {
        if (ticking) return;
        ticking = true;
        requestAnimationFrame(function () {
          if (Store.get('editing')) commitEdit(true);
          updateVirtualRows();
          ticking = false;
        });
      });
      virtualScrollBound = true;
    }

    function attachColumnResize() {
      const table = document.querySelector('.spreadsheet__table');
      if (!table) return;
      const range = Store.getRange();
      table.querySelectorAll('.spreadsheet__resize-handle').forEach(function (handle) {
        handle.addEventListener('mousedown', function (e) {
          e.preventDefault();
          e.stopPropagation();
          const colIndex = parseInt(handle.getAttribute('data-c'), 10);
          const startX = e.clientX;
          const startW = getColWidth(colIndex);
          handle.classList.add('spreadsheet__resize-handle--active');
          const onMove = function (ev) {
            const newW = Math.max(CONST.MIN_COL_WIDTH, startW + (ev.clientX - startX));
            const widths = Store.get('columnWidths');
            const sn = Store.get('sheetName');
            if (!widths[sn]) widths[sn] = {};
            widths[sn][colIndex] = newW;
            // re-apply col widths
            const cols = table.querySelectorAll('colgroup col[data-c]');
            cols.forEach(function (col) {
              const c = parseInt(col.getAttribute('data-c'), 10);
              col.style.width = getColWidth(c) + 'px';
            });
          };
          const onUp = function () {
            handle.classList.remove('spreadsheet__resize-handle--active');
            document.removeEventListener('mousemove', onMove);
            document.removeEventListener('mouseup', onUp);
          };
          document.addEventListener('mousemove', onMove);
          document.addEventListener('mouseup', onUp);
        });
      });
    }

    function renderSheet(sheetName) {
      Store.set('sheetName', sheetName);
      const styles = Store.get('cellStyles');
      if (!styles[sheetName]) styles[sheetName] = {};

      const sheet = Store.getSheet();
      const range = Store.getRange();
      ensureColumnWidths(sheetName);

      const rowHdrW = getRowHeaderWidth();
      let totalWidth = rowHdrW;
      for (let C = range.s.c; C <= range.e.c; C++) totalWidth += getColWidth(C);

      let html = '<table class="spreadsheet__table" style="width:' + totalWidth + 'px"><colgroup>';
      html += '<col style="width:' + rowHdrW + 'px;" />';
      for (let C = range.s.c; C <= range.e.c; C++) {
        html += '<col data-c="' + C + '" style="width:' + getColWidth(C) + 'px;" />';
      }
      html += '</colgroup><thead><tr><th class="spreadsheet__corner"></th>';

      const activeFilters = Store.get('activeFilters');
      for (let C = range.s.c; C <= range.e.c; C++) {
        const colLetter = Utils.getColLetter(C);
        const isFiltered = activeFilters[C] !== undefined;
        const headerCell = sheet[Utils.encodeCell(range.s.r, C)];
        const headerVal = headerCell && headerCell.v !== undefined && headerCell.v !== ''
          ? String(headerCell.v) : colLetter;
        const safeHeader = Utils.escapeHtml(headerVal);
        html += '<th class="spreadsheet__header" data-c="' + C + '" title="' + colLetter + ': ' + safeHeader + '">' +
          '<div class="spreadsheet__header-content">' +
            '<span class="spreadsheet__header-label">' + safeHeader + '</span>' +
            '<button type="button" class="spreadsheet__filter-btn' + (isFiltered ? ' spreadsheet__filter-btn--active' : '') +
            '" data-col="' + C + '">▼</button>' +
          '</div>' +
          '<div class="spreadsheet__resize-handle" data-c="' + C + '" title="Arrastra para ajustar ancho"></div>' +
          '</th>';
      }
      html += '</tr></thead><tbody id="virtualTbody"></tbody></table>';

      const container = document.getElementById('tableContainer');
      container.innerHTML = html;
      container.scrollTop = 0;

      FilterEngine.rebuildSync();
      updateVirtualRows();
      bindVirtualScroll();
      attachColumnResize();
      attachSelectionEvents();
      EventBus.emit('sheet:rendered', sheetName);
    }

    function clearSelectionUI() {
      Store.clearSelection();
      document.querySelectorAll('.spreadsheet__cell--selected').forEach(function (td) {
        td.classList.remove('spreadsheet__cell--selected');
      });
      document.querySelectorAll('.spreadsheet__header--selected, .spreadsheet__row-header--selected').forEach(function (th) {
        th.classList.remove('spreadsheet__header--selected', 'spreadsheet__row-header--selected');
      });
    }

    function selectCell(td, autoCalc) {
      if (autoCalc === undefined) autoCalc = true;
      const r = parseInt(td.getAttribute('data-r'), 10);
      const c = parseInt(td.getAttribute('data-c'), 10);
      const sel = Store.get('selection');
      if (sel.cols.size || sel.rows.size || sel.range) clearSelectionUI();
      sel.cells.add(Store.cellKey(r, c));
      td.classList.add('spreadsheet__cell--selected');
      sel.anchorR = r;
      sel.anchorC = c;
      anchorTd = td;
      updateFormulaBar(td);
      if (autoCalc) calculateStats();
    }

    function selectRow(rowIndex) {
      clearSelectionUI();
      const sel = Store.get('selection');
      sel.rows.add(rowIndex);
      sel.anchorR = rowIndex;
      updateVirtualRows();
      const firstTd = document.querySelector('tr[data-r="' + rowIndex + '"] td');
      if (firstTd) {
        anchorTd = firstTd;
        sel.anchorC = parseInt(firstTd.getAttribute('data-c'), 10);
        updateFormulaBar(firstTd);
      }
      calculateStats();
    }

    function selectColumn(colIndex) {
      clearSelectionUI();
      const sel = Store.get('selection');
      sel.cols.add(colIndex);
      sel.anchorC = colIndex;
      const filtered = Store.get('filteredRowIndices');
      sel.anchorR = filtered.length ? filtered[0] : null;
      updateVirtualRows();
      document.querySelector('.spreadsheet__header[data-c="' + colIndex + '"]')
        ?.classList.add('spreadsheet__header--selected');
      const firstTd = document.querySelector('td[data-c="' + colIndex + '"]');
      if (firstTd) { anchorTd = firstTd; updateFormulaBar(firstTd); }
      calculateStats();
    }

    function selectRange(startTd, endTd) {
      const startR = parseInt(startTd.getAttribute('data-r'), 10);
      const startC = parseInt(startTd.getAttribute('data-c'), 10);
      const endR = parseInt(endTd.getAttribute('data-r'), 10);
      const endC = parseInt(endTd.getAttribute('data-c'), 10);
      clearSelectionUI();
      const sel = Store.get('selection');
      sel.range = {
        minR: Math.min(startR, endR), maxR: Math.max(startR, endR),
        minC: Math.min(startC, endC), maxC: Math.max(startC, endC)
      };
      sel.anchorR = startR;
      sel.anchorC = startC;
      anchorTd = startTd;
      updateVirtualRows();
      calculateStats();
    }

    function updateFormulaBar(td) {
      if (!td) return;
      const addr = document.getElementById('activeCellAddress');
      const input = document.getElementById('formulaInput');
      if (addr) addr.textContent = td.getAttribute('data-cell') || 'A1';
      if (input) input.value = td.getAttribute('data-val') || '';
    }

    function calculateStats() {
      let sum = 0, numericCount = 0, count = 0;
      const sheet = Store.getSheet();
      const sel = Store.get('selection');
      const filtered = Store.get('filteredRowIndices');

      function addVal(raw) {
        count++;
        if (raw === null || raw === undefined || raw === '') return;
        const num = Number(raw);
        if (!isNaN(num)) { sum += num; numericCount++; }
      }

      if (sheet && sel.cols.size > 0) {
        sel.cols.forEach(function (C) {
          filtered.forEach(function (R) {
            const cell = sheet[Utils.encodeCell(R, C)];
            addVal(cell && cell.v !== undefined ? cell.v : '');
          });
        });
      } else if (sheet && sel.rows.size > 0) {
        const range = Store.getRange();
        sel.rows.forEach(function (R) {
          for (let C = range.s.c; C <= range.e.c; C++) {
            const cell = sheet[Utils.encodeCell(R, C)];
            addVal(cell && cell.v !== undefined ? cell.v : '');
          }
        });
      } else if (sheet && sel.range) {
        const r = sel.range;
        for (let R = r.minR; R <= r.maxR; R++) {
          if (filtered.length && filtered.indexOf(R) === -1) continue;
          for (let C = r.minC; C <= r.maxC; C++) {
            const cell = sheet[Utils.encodeCell(R, C)];
            addVal(cell && cell.v !== undefined ? cell.v : '');
          }
        }
      } else if (sheet && sel.cells.size > 0) {
        sel.cells.forEach(function (key) {
          const parts = key.split(',').map(Number);
          const cell = sheet[Utils.encodeCell(parts[0], parts[1])];
          addVal(cell && cell.v !== undefined ? cell.v : '');
        });
      }

      const avg = numericCount > 0 ? sum / numericCount : 0;
      const fmt = function (n) { return n.toLocaleString('es-ES', { maximumFractionDigits: 4 }); };
      const elSum = document.getElementById('sumValue');
      const elAvg = document.getElementById('avgValue');
      const elCount = document.getElementById('countValue');
      if (elSum) elSum.textContent = fmt(sum);
      if (elAvg) elAvg.textContent = fmt(avg);
      if (elCount) elCount.textContent = count;
    }

    function startEdit(td, replaceMode, initialChar) {
      if (Store.get('editing')) commitEdit(true);
      if (!td || !Store.get('workbook')) return;
      Store.set('editing', true);
      editingTd = td;
      const currentVal = replaceMode ? '' : (td.getAttribute('data-val') || '');
      const startValue = replaceMode ? (initialChar || '') : currentVal;
      td.classList.add('spreadsheet__cell--editing');
      td.classList.remove('spreadsheet__cell--selected');
      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'spreadsheet__cell-editor';
      input.value = startValue;
      td.innerHTML = '';
      td.appendChild(input);
      input.focus();
      input.setSelectionRange(input.value.length, input.value.length);

      input.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') { e.preventDefault(); commitEdit(true); navigate('ArrowDown'); }
        else if (e.key === 'Escape') { e.preventDefault(); commitEdit(false); }
        else if (e.key === 'Tab') { e.preventDefault(); commitEdit(true); navigate(e.shiftKey ? 'ArrowLeft' : 'ArrowRight'); }
        else if (e.key === 'ArrowUp' || e.key === 'ArrowDown') { e.preventDefault(); commitEdit(true); navigate(e.key); }
      });
      input.addEventListener('blur', function () {
        if (Store.get('editing') && editingTd === td) commitEdit(true);
      });
    }

    function commitEdit(save) {
      if (!Store.get('editing') || !editingTd) return;
      const td = editingTd;
      const input = td.querySelector('input.spreadsheet__cell-editor');
      const originalVal = td.getAttribute('data-val') || '';
      let newVal = originalVal;
      if (save && input) newVal = input.value;

      td.classList.remove('spreadsheet__cell--editing');
      td.textContent = newVal;
      td.setAttribute('data-val', newVal);

      if (save && newVal !== originalVal) {
        const cellAddress = td.getAttribute('data-cell');
        const sheet = Store.getSheet();
        if (newVal === '') {
          if (sheet[cellAddress]) delete sheet[cellAddress];
        } else {
          const num = Number(newVal);
          if (!isNaN(num) && newVal.trim() !== '') sheet[cellAddress] = { v: num, t: 'n' };
          else sheet[cellAddress] = { v: newVal, t: 's' };
        }
        const range = Store.getRange();
        const coords = XLSX.utils.decode_cell(cellAddress);
        if (coords.r > range.e.r || coords.c > range.e.c || coords.r < range.s.r || coords.c < range.s.c) {
          range.s.r = Math.min(range.s.r, coords.r);
          range.s.c = Math.min(range.s.c, coords.c);
          range.e.r = Math.max(range.e.r, coords.r);
          range.e.c = Math.max(range.e.c, coords.c);
          sheet['!ref'] = XLSX.utils.encode_range(range);
        }
      }
      Store.set('editing', false);
      editingTd = null;
      clearSelectionUI();
      selectCell(td, false);
      calculateStats();
    }

    function navigate(direction) {
      if (!anchorTd) return;
      const r = parseInt(anchorTd.getAttribute('data-r'), 10);
      const c = parseInt(anchorTd.getAttribute('data-c'), 10);
      if (direction === 'ArrowLeft' || direction === 'ArrowRight') {
        const nextC = c + (direction === 'ArrowRight' ? 1 : -1);
        const nextTd = document.querySelector('td[data-r="' + r + '"][data-c="' + nextC + '"]');
        if (nextTd) {
          clearSelectionUI();
          selectCell(nextTd, false);
          calculateStats();
          nextTd.scrollIntoView({ block: 'nearest', inline: 'nearest' });
        }
        return;
      }
      const filtered = Store.get('filteredRowIndices');
      const idx = filtered.indexOf(r);
      if (idx === -1) return;
      const nextIdx = direction === 'ArrowDown' ? idx + 1 : idx - 1;
      if (nextIdx < 0 || nextIdx >= filtered.length) return;
      const nextR = filtered[nextIdx];
      const container = document.getElementById('tableContainer');
      if (container) {
        const targetScroll = nextIdx * CONST.ROW_HEIGHT;
        const viewTop = container.scrollTop;
        const viewBottom = viewTop + container.clientHeight - CONST.ROW_HEIGHT * 2;
        if (targetScroll < viewTop || targetScroll > viewBottom) {
          container.scrollTop = Math.max(0, targetScroll - CONST.VIRTUAL_BUFFER * CONST.ROW_HEIGHT);
          updateVirtualRows();
        }
      }
      const nextTd = document.querySelector('td[data-r="' + nextR + '"][data-c="' + c + '"]');
      if (nextTd) {
        clearSelectionUI();
        selectCell(nextTd, false);
        calculateStats();
        nextTd.scrollIntoView({ block: 'nearest', inline: 'nearest' });
      }
    }

    function clearSelectedContent() {
      const sheet = Store.getSheet();
      if (!sheet) return;
      const sel = Store.get('selection');
      const filtered = Store.get('filteredRowIndices');
      let changed = false;

      function clearAddr(R, C) {
        const addr = Utils.encodeCell(R, C);
        if (sheet[addr]) { delete sheet[addr]; changed = true; }
      }

      if (sel.cols.size > 0) {
        sel.cols.forEach(function (C) { filtered.forEach(function (R) { clearAddr(R, C); }); });
      } else if (sel.rows.size > 0) {
        const range = Store.getRange();
        sel.rows.forEach(function (R) {
          for (let C = range.s.c; C <= range.e.c; C++) clearAddr(R, C);
        });
      } else if (sel.range) {
        const r = sel.range;
        for (let R = r.minR; R <= r.maxR; R++)
          for (let C = r.minC; C <= r.maxC; C++) clearAddr(R, C);
      } else if (sel.cells.size > 0) {
        sel.cells.forEach(function (key) {
          const p = key.split(',').map(Number);
          clearAddr(p[0], p[1]);
        });
      }
      if (!changed) return;
      updateVirtualRows();
      if (anchorTd) updateFormulaBar(anchorTd);
      calculateStats();
      Toast.show('Contenido eliminado');
    }

    function applyColor(bgColor) {
      if (!Store.hasSelection()) {
        Toast.show('Selecciona primero al menos una celda');
        return;
      }
      const styles = Store.get('cellStyles');
      const sn = Store.get('sheetName');
      if (!styles[sn]) styles[sn] = {};
      const fontColor = Utils.contrastingText(bgColor);
      const sel = Store.get('selection');
      const filtered = Store.get('filteredRowIndices');
      const range = Store.getRange();

      function applyTo(addr) {
        if (bgColor === '#1a1220' || bgColor === '#18181b') delete styles[sn][addr];
        else styles[sn][addr] = { bg: bgColor, font: fontColor };
      }

      if (sel.cols.size > 0) {
        sel.cols.forEach(function (C) { filtered.forEach(function (R) { applyTo(Utils.encodeCell(R, C)); }); });
      } else if (sel.rows.size > 0) {
        sel.rows.forEach(function (R) {
          for (let C = range.s.c; C <= range.e.c; C++) applyTo(Utils.encodeCell(R, C));
        });
      } else if (sel.range) {
        const r = sel.range;
        for (let R = r.minR; R <= r.maxR; R++)
          for (let C = r.minC; C <= r.maxC; C++) applyTo(Utils.encodeCell(R, C));
      } else if (sel.cells.size > 0) {
        sel.cells.forEach(function (key) {
          const p = key.split(',').map(Number);
          applyTo(Utils.encodeCell(p[0], p[1]));
        });
      }
      updateVirtualRows();
      Toast.show('Color de relleno aplicado');
    }

    function attachSelectionEvents() {
      const table = document.querySelector('.spreadsheet__table');
      if (!table) return;

      table.addEventListener('click', function (e) {
        const rowHeader = e.target.closest('.spreadsheet__row-header');
        if (rowHeader) {
          if (Store.get('editing')) commitEdit(true);
          const r = parseInt(rowHeader.getAttribute('data-r'), 10);
          if (!isNaN(r)) selectRow(r);
          return;
        }
        const colHeader = e.target.closest('.spreadsheet__header');
        if (colHeader) {
          if (e.target.closest('.spreadsheet__filter-btn')) return;
          if (Store.get('editing')) commitEdit(true);
          const c = parseInt(colHeader.getAttribute('data-c'), 10);
          if (!isNaN(c)) selectColumn(c);
        }
      });

      table.addEventListener('mousedown', function (e) {
        if (e.button !== 0) return;
        if (e.target.closest('.spreadsheet__row-header') || e.target.closest('.spreadsheet__header')) return;
        if (Store.get('editing')) {
          const clicked = e.target.closest('td');
          if (clicked !== editingTd) commitEdit(true);
          else return;
        }
        const td = e.target.closest('td.spreadsheet__cell');
        if (!td) return;
        isMouseDown = true;
        if (e.shiftKey && anchorTd) selectRange(anchorTd, td);
        else if (e.ctrlKey || e.metaKey) {
          const r = parseInt(td.getAttribute('data-r'), 10);
          const c = parseInt(td.getAttribute('data-c'), 10);
          const key = Store.cellKey(r, c);
          const sel = Store.get('selection');
          if (sel.cells.has(key)) {
            sel.cells.delete(key);
            td.classList.remove('spreadsheet__cell--selected');
          } else {
            sel.cells.add(key);
            td.classList.add('spreadsheet__cell--selected');
          }
          anchorTd = td;
          calculateStats();
        } else {
          clearSelectionUI();
          selectCell(td, false);
          calculateStats();
        }
        updateFormulaBar(td);
      });

      table.addEventListener('mouseover', function (e) {
        if (Store.get('editing') || !isMouseDown || !anchorTd) return;
        const td = e.target.closest('td.spreadsheet__cell');
        if (td && !e.ctrlKey && !e.metaKey) {
          selectRange(anchorTd, td);
          updateFormulaBar(td);
        }
      });

      table.addEventListener('dblclick', function (e) {
        const td = e.target.closest('td.spreadsheet__cell');
        if (!td) return;
        e.preventDefault();
        startEdit(td, false);
      });

      if (!windowMouseUpBound) {
        window.addEventListener('mouseup', function () { isMouseDown = false; });
        windowMouseUpBound = true;
      }
    }

    function copyToClipboard() {
      const sheet = Store.getSheet();
      if (!sheet) return;
      const range = Store.getRange();
      const sel = Store.get('selection');
      const filtered = Store.get('filteredRowIndices');

      function getVal(R, C) {
        const cell = sheet[Utils.encodeCell(R, C)];
        return cell && cell.v !== undefined ? String(cell.v) : '';
      }

      let copyText = '';
      if (sel.cols.size > 0) {
        const cols = Array.from(sel.cols).sort(function (a, b) { return a - b; });
        filtered.forEach(function (R) {
          copyText += cols.map(function (C) { return getVal(R, C); }).join('\t') + '\n';
        });
      } else if (sel.rows.size > 0) {
        Array.from(sel.rows).sort(function (a, b) { return a - b; }).forEach(function (R) {
          const vals = [];
          for (let C = range.s.c; C <= range.e.c; C++) vals.push(getVal(R, C));
          copyText += vals.join('\t') + '\n';
        });
      } else if (sel.range) {
        const r = sel.range;
        for (let R = r.minR; R <= r.maxR; R++) {
          const vals = [];
          for (let C = r.minC; C <= r.maxC; C++) vals.push(getVal(R, C));
          copyText += vals.join('\t') + '\n';
        }
      } else if (sel.cells.size > 0) {
        let minR = Infinity, maxR = -Infinity, minC = Infinity, maxC = -Infinity;
        const map = new Map();
        sel.cells.forEach(function (key) {
          const p = key.split(',').map(Number);
          if (!map.has(p[0])) map.set(p[0], new Map());
          map.get(p[0]).set(p[1], getVal(p[0], p[1]));
          if (p[0] < minR) minR = p[0]; if (p[0] > maxR) maxR = p[0];
          if (p[1] < minC) minC = p[1]; if (p[1] > maxC) maxC = p[1];
        });
        for (let r = minR; r <= maxR; r++) {
          if (!map.has(r)) continue;
          const rowMap = map.get(r);
          const vals = [];
          for (let c = minC; c <= maxC; c++) vals.push(rowMap.has(c) ? rowMap.get(c) : '');
          copyText += vals.join('\t') + '\n';
        }
      } else return;

      navigator.clipboard.writeText(copyText.trimEnd()).then(function () {
        Toast.show('Celdas copiadas al portapapeles');
      });
    }

    function resetView() {
      Store.patch({
        workbook: null,
        sheetName: '',
        cellStyles: {},
        columnWidths: {},
        activeFilters: {}
      });
      Store.clearSelection();
      renderDropzone();
      document.getElementById('tabsContainer').innerHTML = '';
      document.getElementById('activeCellAddress').textContent = 'A1';
      document.getElementById('formulaInput').value = '';
      document.getElementById('sumValue').textContent = '0';
      document.getElementById('avgValue').textContent = '0';
      document.getElementById('countValue').textContent = '0';
    }

    EventBus.on('filters:applied', function () {
      const container = document.getElementById('tableContainer');
      if (container) container.scrollTop = 0;
      clearSelectionUI();
      updateVirtualRows();
      document.getElementById('sumValue').textContent = '0';
      document.getElementById('avgValue').textContent = '0';
      document.getElementById('countValue').textContent = '0';
    });

    return Object.freeze({
      renderSheet: renderSheet,
      renderDropzone: renderDropzone,
      updateVirtualRows: updateVirtualRows,
      clearSelectionUI: clearSelectionUI,
      selectCell: selectCell,
      selectRow: selectRow,
      selectColumn: selectColumn,
      startEdit: startEdit,
      commitEdit: commitEdit,
      navigate: navigate,
      clearSelectedContent: clearSelectedContent,
      applyColor: applyColor,
      copyToClipboard: copyToClipboard,
      calculateStats: calculateStats,
      updateFormulaBar: updateFormulaBar,
      resetView: resetView,
      getAnchorTd: function () { return anchorTd; }
    });
  })();


  /* ========================================================================
   * EXPLORER MODULE
   * ======================================================================== */
  const Explorer = (function () {

    function updateActiveSnapshot() {
      const activeId = Store.get('activeFileId');
      if (!activeId || !Store.get('workbook')) return;
      const files = Store.get('files');
      const entry = files.find(function (f) { return f.id === activeId; });
      if (!entry) return;
      if (entry.isSql) {
        entry.cellStyles = JSON.parse(JSON.stringify(Store.get('cellStyles') || {}));
        entry.columnWidths = JSON.parse(JSON.stringify(Store.get('columnWidths') || {}));
        return;
      }
      entry.workbook = Utils.cloneWorkbook(Store.get('workbook'));
      entry.cellStyles = JSON.parse(JSON.stringify(Store.get('cellStyles') || {}));
      entry.columnWidths = JSON.parse(JSON.stringify(Store.get('columnWidths') || {}));
    }

    function registerFile(fileName, fileType) {
      const counter = Store.get('fileIdCounter') + 1;
      Store.set('fileIdCounter', counter);
      const realId = 'f_' + counter + '_' + Date.now();
      const entry = {
        id: realId,
        name: fileName,
        type: fileType || Utils.detectFileType(fileName),
        workbook: Utils.cloneWorkbook(Store.get('workbook')),
        cellStyles: JSON.parse(JSON.stringify(Store.get('cellStyles') || {})),
        columnWidths: JSON.parse(JSON.stringify(Store.get('columnWidths') || {}))
      };
      Store.get('files').push(entry);
      Store.set('activeFileId', realId);
      render();
      return realId;
    }

    function registerSqlFile(fileName, tables) {
      const counter = Store.get('fileIdCounter') + 1;
      Store.set('fileIdCounter', counter);
      const id = 'f_' + counter + '_' + Date.now();
      const entry = {
        id: id,
        name: fileName,
        type: 'sql',
        isSql: true,
        tables: tables,
        workbook: null,
        cellStyles: {},
        columnWidths: {},
        activeTableName: null,
        expanded: true
      };
      Store.get('files').push(entry);
      Store.set('activeFileId', id);
      render();
      return id;
    }

    function loadSqlTable(fileId, tableName) {
      const entry = Store.get('files').find(function (f) { return f.id === fileId; });
      if (!entry || !entry.isSql) return;
      if (Store.get('activeFileId') && Store.get('activeFileId') !== fileId) updateActiveSnapshot();

      const table = entry.tables.find(function (t) { return t.name === tableName; });
      if (!table) { Toast.show('Tabla no encontrada: ' + tableName, 'danger'); return; }

      Store.set('activeFileId', fileId);
      entry.activeTableName = tableName;

      const ws = Parsers.sqlTableToWorksheet(table);
      const sheetName = table.name.substring(0, 31);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, sheetName);

      Store.patch({
        workbook: wb,
        cellStyles: {},
        columnWidths: {},
        activeFilters: {}
      });
      Store.clearSelection();
      Tabs.render();
      Spreadsheet.renderSheet(sheetName);
      render();
      Toast.show('Tabla "' + table.name + '" cargada (' + table.rows.length + ' filas)');
    }

    function openFile(fileId) {
      const entry = Store.get('files').find(function (f) { return f.id === fileId; });
      if (!entry) return;
      updateActiveSnapshot();
      if (entry.isSql) {
        const tblName = entry.activeTableName || (entry.tables[0] && entry.tables[0].name);
        if (tblName) loadSqlTable(fileId, tblName);
        return;
      }
      Store.set('activeFileId', entry.id);
      Store.patch({
        workbook: Utils.cloneWorkbook(entry.workbook),
        cellStyles: JSON.parse(JSON.stringify(entry.cellStyles || {})),
        columnWidths: JSON.parse(JSON.stringify(entry.columnWidths || {})),
        activeFilters: {}
      });
      Store.clearSelection();
      const firstSheet = Store.get('workbook').SheetNames[0];
      Tabs.render();
      Spreadsheet.renderSheet(firstSheet);
      render();
      Toast.show('Abierto: ' + entry.name);
    }

    function removeFile(fileId) {
      const files = Store.get('files');
      const idx = files.findIndex(function (f) { return f.id === fileId; });
      if (idx === -1) return;
      const wasActive = Store.get('activeFileId') === fileId;
      files.splice(idx, 1);
      if (wasActive) {
        Store.set('activeFileId', null);
        Spreadsheet.resetView();
      }
      render();
      Toast.show('Archivo quitado del explorador');
    }

    function render() {
      const container = document.getElementById('treeChildren');
      if (!container) return;
      const files = Store.get('files');
      if (!files.length) {
        container.innerHTML = '<div class="tree__empty" id="treeEmpty">Sin archivos</div>';
        return;
      }

      let html = '';
      files.forEach(function (item) {
        const iconInfo = Utils.iconForType(item.type);
        const safeName = Utils.escapeHtml(item.name);

        if (item.isSql && item.tables && item.tables.length) {
          const parentActive = item.id === Store.get('activeFileId') ? ' tree__sql-parent--active' : '';
          const collapsed = item.expanded === false ? ' tree__sql-children--collapsed' : '';
          const chevron = item.expanded === false ? 'chevron_right' : 'expand_more';
          html += '<div class="tree__sql-group" data-id="' + item.id + '">';
          html += '<div class="tree__sql-parent' + parentActive + '" data-id="' + item.id + '">' +
            '<button type="button" class="tree__toggle tree__sql-toggle" data-id="' + item.id + '">' +
              '<span class="material-icons tree__chevron">' + chevron + '</span></button>' +
            '<span class="material-icons tree__icon ' + iconInfo.cls + '">' + iconInfo.icon + '</span>' +
            '<span class="tree__label">' + safeName + '</span>' +
            '<button type="button" class="tree__file-remove" data-id="' + item.id + '">' +
              '<span class="material-icons">close</span></button></div>';
          html += '<div class="tree__sql-children' + collapsed + '" data-parent="' + item.id + '">';
          item.tables.forEach(function (tbl, tIdx) {
            const tActive = (item.id === Store.get('activeFileId') && item.activeTableName === tbl.name)
              ? ' tree__table-item--active' : '';
            html += '<div class="tree__table-item' + tActive + '" data-file-id="' + item.id +
              '" data-table-idx="' + tIdx + '">' +
              '<span class="material-icons tree__icon tree__icon--table">grid_on</span>' +
              '<span class="tree__label">' + Utils.escapeHtml(tbl.name) + '</span>' +
              '<span class="tree__table-badge">' + (tbl.rows ? tbl.rows.length : 0) + '</span></div>';
          });
          html += '</div></div>';
        } else {
          const active = item.id === Store.get('activeFileId') ? ' tree__file--active' : '';
          html += '<div class="tree__file' + active + '" data-id="' + item.id + '" title="Doble clic para abrir">' +
            '<span class="material-icons tree__icon ' + iconInfo.cls + '">' + iconInfo.icon + '</span>' +
            '<span class="tree__label">' + safeName + '</span>' +
            '<button type="button" class="tree__file-remove" data-id="' + item.id + '">' +
              '<span class="material-icons">close</span></button></div>';
        }
      });
      container.innerHTML = html;
      bindEvents(container);
    }

    function bindEvents(container) {
      container.querySelectorAll('.tree__file').forEach(function (el) {
        el.addEventListener('dblclick', function (e) {
          if (e.target.closest('.tree__file-remove')) return;
          openFile(el.getAttribute('data-id'));
        });
        el.addEventListener('click', function (e) {
          if (e.target.closest('.tree__file-remove')) return;
          container.querySelectorAll('.tree__file, .tree__sql-parent, .tree__table-item').forEach(function (n) {
            n.classList.remove('tree__file--active', 'tree__sql-parent--active', 'tree__table-item--active');
          });
          el.classList.add('tree__file--active');
        });
      });
      container.querySelectorAll('.tree__sql-toggle').forEach(function (btn) {
        btn.addEventListener('click', function (e) {
          e.stopPropagation();
          const id = btn.getAttribute('data-id');
          const entry = Store.get('files').find(function (f) { return f.id === id; });
          if (!entry) return;
          entry.expanded = entry.expanded === false ? true : false;
          render();
        });
      });
      container.querySelectorAll('.tree__table-item').forEach(function (el) {
        el.addEventListener('click', function (e) {
          e.stopPropagation();
          const fileId = el.getAttribute('data-file-id');
          const tIdx = parseInt(el.getAttribute('data-table-idx'), 10);
          const entry = Store.get('files').find(function (f) { return f.id === fileId; });
          if (!entry || !entry.tables || !entry.tables[tIdx]) return;
          loadSqlTable(fileId, entry.tables[tIdx].name);
        });
      });
      container.querySelectorAll('.tree__file-remove').forEach(function (btn) {
        btn.addEventListener('click', function (e) {
          e.stopPropagation();
          removeFile(btn.getAttribute('data-id'));
        });
      });
    }

    function init() {
      const root = document.getElementById('treeRoot');
      const children = document.getElementById('treeChildren');
      const toggle = document.getElementById('treeRootToggle');
      if (toggle && root && children) {
        toggle.addEventListener('click', function () {
          root.classList.toggle('tree__node--collapsed');
          children.classList.toggle('tree__children--collapsed');
        });
      }
      render();
    }

    return Object.freeze({
      init: init,
      render: render,
      registerFile: registerFile,
      registerSqlFile: registerSqlFile,
      loadSqlTable: loadSqlTable,
      openFile: openFile,
      removeFile: removeFile,
      updateActiveSnapshot: updateActiveSnapshot
    });
  })();

  /* ========================================================================
   * TABS MODULE
   * ======================================================================== */
  const Tabs = (function () {

    function render() {
      const container = document.getElementById('tabsContainer');
      if (!container) return;
      container.innerHTML = '';
      const wb = Store.get('workbook');

      if (wb && wb.SheetNames.length) {
        const activeName = wb.SheetNames.indexOf(Store.get('sheetName')) !== -1
          ? Store.get('sheetName') : wb.SheetNames[0];

        wb.SheetNames.forEach(function (sheetName) {
          const btn = document.createElement('button');
          btn.className = 'statusbar__tab' + (sheetName === activeName ? ' statusbar__tab--active' : '');
          btn.dataset.sheet = sheetName;

          const nameSpan = document.createElement('span');
          nameSpan.className = 'statusbar__tab-name';
          nameSpan.textContent = sheetName;
          nameSpan.title = 'Doble clic para renombrar';

          const closeBtn = document.createElement('button');
          closeBtn.className = 'statusbar__tab-close';
          closeBtn.type = 'button';
          closeBtn.title = 'Cerrar';
          closeBtn.innerHTML = '<span class="material-icons" style="font-size:14px;">close</span>';

          btn.addEventListener('click', function (e) {
            if (e.target.closest('.statusbar__tab-close')) return;
            if (btn.querySelector('.statusbar__tab-rename')) return;
            document.querySelectorAll('.statusbar__tab').forEach(function (b) {
              b.classList.remove('statusbar__tab--active');
            });
            btn.classList.add('statusbar__tab--active');
            Store.set('activeFilters', {});
            Spreadsheet.renderSheet(sheetName);
          });

          nameSpan.addEventListener('dblclick', function (e) {
            e.stopPropagation();
            startRename(btn, nameSpan, sheetName);
          });

          closeBtn.addEventListener('click', function (e) {
            e.stopPropagation();
            closeSheet(sheetName);
          });

          btn.appendChild(nameSpan);
          btn.appendChild(closeBtn);
          container.appendChild(btn);
        });
      }

      const addBtn = document.createElement('button');
      addBtn.type = 'button';
      addBtn.className = 'statusbar__tab-add';
      addBtn.title = 'Nueva hoja en blanco';
      addBtn.innerHTML = '<span class="material-icons">add</span>';
      addBtn.addEventListener('click', function (e) {
        e.stopPropagation();
        createBlankSheet();
      });
      container.appendChild(addBtn);
    }

    function startRename(btn, nameSpan, oldName) {
      if (btn.querySelector('.statusbar__tab-rename')) return;
      const input = document.createElement('input');
      input.type = 'text';
      input.className = 'statusbar__tab-rename';
      input.value = oldName;
      nameSpan.style.display = 'none';
      btn.insertBefore(input, nameSpan);
      input.focus();
      input.select();

      function finish(save) {
        const newName = input.value.trim();
        input.remove();
        nameSpan.style.display = '';
        if (!save || !newName || newName === oldName) return;
        if (/[\\\/\?\*\[\]:]/.test(newName)) {
          Toast.show('Nombre inválido: no se permiten \\ / ? * [ ] :', 'danger');
          return;
        }
        if (newName.length > 31) { Toast.show('El nombre no puede superar 31 caracteres', 'danger'); return; }
        if (Store.get('workbook').SheetNames.indexOf(newName) !== -1) {
          Toast.show('Ya existe una hoja con ese nombre', 'danger');
          return;
        }
        renameSheet(oldName, newName);
      }

      input.addEventListener('keydown', function (e) {
        if (e.key === 'Enter') { e.preventDefault(); finish(true); }
        else if (e.key === 'Escape') { e.preventDefault(); finish(false); }
      });
      input.addEventListener('blur', function () { finish(true); });
    }

    function renameSheet(oldName, newName) {
      const wb = Store.get('workbook');
      if (!wb || !wb.Sheets[oldName]) return;
      wb.Sheets[newName] = wb.Sheets[oldName];
      delete wb.Sheets[oldName];
      const idx = wb.SheetNames.indexOf(oldName);
      if (idx !== -1) wb.SheetNames[idx] = newName;
      const styles = Store.get('cellStyles');
      if (styles[oldName]) { styles[newName] = styles[oldName]; delete styles[oldName]; }
      if (Store.get('sheetName') === oldName) Store.set('sheetName', newName);
      render();
      Toast.show('Hoja renombrada a "' + newName + '"');
    }

    function closeSheet(sheetName) {
      const wb = Store.get('workbook');
      if (!wb) return;
      if (wb.SheetNames.length <= 1) {
        Explorer.updateActiveSnapshot();
        Store.set('activeFileId', null);
        Spreadsheet.resetView();
        Explorer.render();
        Toast.show('Vista cerrada (el archivo sigue en el explorador)');
        return;
      }
      delete wb.Sheets[sheetName];
      wb.SheetNames = wb.SheetNames.filter(function (n) { return n !== sheetName; });
      const styles = Store.get('cellStyles');
      if (styles[sheetName]) delete styles[sheetName];
      if (Store.get('sheetName') === sheetName) {
        Store.set('activeFilters', {});
        render();
        Spreadsheet.renderSheet(wb.SheetNames[0]);
      } else {
        render();
      }
      Toast.show('Hoja "' + sheetName + '" cerrada');
    }

    function createBlankSheet() {
      Explorer.updateActiveSnapshot();
      const ws = {};
      ws['!ref'] = XLSX.utils.encode_range({
        s: { r: 0, c: 0 },
        e: { r: CONST.BLANK_ROWS - 1, c: CONST.BLANK_COLS - 1 }
      });

      if (!Store.get('workbook')) {
        const sheetName = 'Hoja1';
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
        Store.patch({ workbook: wb, cellStyles: {}, columnWidths: {}, activeFilters: {} });
        Store.clearSelection();
        Explorer.registerFile('Libro en blanco.xlsx', 'xlsx');
        render();
        Spreadsheet.renderSheet(sheetName);
        Toast.show('Hoja en blanco creada');
        return;
      }

      const wb = Store.get('workbook');
      let n = 1;
      let name = 'Hoja' + n;
      while (wb.SheetNames.indexOf(name) !== -1) { n++; name = 'Hoja' + n; }
      const uniqueName = Utils.makeUniqueSheetName(wb, name);
      XLSX.utils.book_append_sheet(wb, ws, uniqueName);
      const styles = Store.get('cellStyles');
      if (!styles[uniqueName]) styles[uniqueName] = {};
      Store.set('activeFilters', {});
      Store.clearSelection();
      const entry = Store.get('files').find(function (f) { return f.id === Store.get('activeFileId'); });
      if (entry && !entry.isSql) entry.workbook = Utils.cloneWorkbook(wb);
      render();
      Spreadsheet.renderSheet(uniqueName);
      Toast.show('Hoja "' + uniqueName + '" creada');
    }

    return Object.freeze({
      render: render,
      createBlankSheet: createBlankSheet,
      closeSheet: closeSheet
    });
  })();


  /* ========================================================================
   * FILE LOADER MODULE
   * ======================================================================== */
  const FileLoader = (function () {
    function processFile(file) {
      if (!file) return;
      Explorer.updateActiveSnapshot();
      Loader.show('Leyendo archivo...');
      const name = file.name.toLowerCase();
      const baseName = Utils.getBaseFileName(file.name);
      const isJson = name.endsWith('.json') || file.type === 'application/json';
      const isCsv = name.endsWith('.csv') || file.type === 'text/csv';
      const isTxt = name.endsWith('.txt') || file.type === 'text/plain';
      const isSql = name.endsWith('.sql') || file.type === 'application/sql' || file.type === 'text/x-sql';
      const fileType = Utils.detectFileType(file.name);
      const reader = new FileReader();
      reader.onprogress = function (e) {
        if (e.lengthComputable) Loader.updateProgress((e.loaded / file.size) * 50);
      };
      reader.onload = function (evt) {
        Loader.updateProgress(60);
        var statusEl = document.getElementById('loaderStatusText');
        if (statusEl) statusEl.textContent = 'Validando y procesando estructura...';
        setTimeout(function () {
          try {
            if (isSql) {
              var tables = Parsers.parseSqlDump(evt.target.result);
              Loader.updateProgress(80);
              var fileId = Explorer.registerSqlFile(file.name, tables);
              if (tables.length > 0) Explorer.loadSqlTable(fileId, tables[0].name);
              Loader.updateProgress(100);
              setTimeout(function () {
                Loader.hide();
                Toast.show('SQL cargado: ' + tables.length + ' tabla' + (tables.length !== 1 ? 's' : ''));
              }, 200);
              return;
            }
            var wb;
            if (isJson) {
              var normalized = Parsers.validateAndNormalizeJson(evt.target.result);
              var worksheet = XLSX.utils.json_to_sheet(normalized);
              wb = XLSX.utils.book_new();
              XLSX.utils.book_append_sheet(wb, worksheet, baseName);
            } else if (isCsv) {
              var parsed = Parsers.validateAndParseCsv(evt.target.result);
              var ws2 = XLSX.utils.aoa_to_sheet(parsed.rows);
              wb = XLSX.utils.book_new();
              XLSX.utils.book_append_sheet(wb, ws2, baseName);
            } else if (isTxt) {
              var parsed2 = Parsers.validateAndParseTxt(evt.target.result);
              var ws3 = XLSX.utils.aoa_to_sheet(parsed2.rows);
              wb = XLSX.utils.book_new();
              XLSX.utils.book_append_sheet(wb, ws3, baseName);
            } else {
              var data = new Uint8Array(evt.target.result);
              wb = XLSX.read(data, { type: 'array', cellStyles: true });
              if (wb.SheetNames.length === 1) {
                var oldName = wb.SheetNames[0];
                if (/^Sheet\d*$/i.test(oldName) || oldName === 'Hoja1') {
                  wb.SheetNames[0] = baseName;
                  wb.Sheets[baseName] = wb.Sheets[oldName];
                  if (baseName !== oldName) delete wb.Sheets[oldName];
                }
              }
            }
            Loader.updateProgress(80);
            if (statusEl) statusEl.textContent = 'Renderizando vista de celdas...';
            setTimeout(function () {
              var cellStyles = {};
              if (wb.Props && wb.Props.Comments) {
                try { cellStyles = JSON.parse(wb.Props.Comments); } catch (err) {}
              }
              Store.patch({ workbook: wb, cellStyles: cellStyles, columnWidths: {}, activeFilters: {} });
              Store.clearSelection();
              Tabs.render();
              Spreadsheet.renderSheet(wb.SheetNames[0]);
              Explorer.registerFile(file.name, fileType);
              Loader.updateProgress(100);
              setTimeout(function () {
                Loader.hide();
                var msg = isJson ? 'JSON cargado correctamente' : isCsv ? 'CSV cargado correctamente' : isTxt ? 'TXT cargado correctamente' : 'Excel cargado correctamente';
                Toast.show(msg);
              }, 200);
            }, 50);
          } catch (error) {
            Loader.hide();
            console.error(error);
            Toast.show('Error de validación:\n' + error.message, 'danger');
          }
        }, 50);
      };
      reader.onerror = function () {
        Loader.hide();
        Toast.show('Error al intentar leer el archivo desde el disco.', 'danger');
      };
      if (isJson || isCsv || isTxt || isSql) reader.readAsText(file, 'UTF-8');
      else reader.readAsArrayBuffer(file);
    }
    return Object.freeze({ processFile: processFile });
  })();

  /* ========================================================================
   * EXPORTER MODULE
   * ======================================================================== */
  const Exporter = (function () {
    function getVisibleData() {
      if (!Store.get('workbook') || !Store.get('sheetName')) {
        Toast.show('Carga un archivo primero');
        return null;
      }
      var sheet = Store.getSheet();
      var range = Store.getRange();
      var headers = [];
      for (var C = range.s.c; C <= range.e.c; C++) {
        var cell = sheet[Utils.encodeCell(range.s.r, C)];
        headers.push(cell && cell.v !== undefined ? String(cell.v) : Utils.getColLetter(C));
      }
      var filtered = Store.get('filteredRowIndices');
      if (!filtered.length) { FilterEngine.rebuildSync(); filtered = Store.get('filteredRowIndices'); }
      var dataRows = [];
      filtered.forEach(function (R) {
        var rowData = [];
        var hasData = false;
        for (var C = range.s.c; C <= range.e.c; C++) {
          var cell = sheet[Utils.encodeCell(R, C)];
          var val = cell && cell.v !== undefined ? String(cell.v) : '';
          rowData.push(val);
          if (val !== '') hasData = true;
        }
        if (hasData) dataRows.push(rowData);
      });
      return { headers: headers, dataRows: dataRows, sheetName: Store.get('sheetName') };
    }
    function toExcel() {
      var wbSrc = Store.get('workbook');
      if (!wbSrc) { Toast.show('No hay datos para exportar'); return; }
      var wb = XLSX.utils.book_new();
      wb.Props = { Title: 'Libro Exportado Studio', Comments: JSON.stringify(Store.get('cellStyles')) };
      var styles = Store.get('cellStyles');
      wbSrc.SheetNames.forEach(function (sheetName) {
        var original = wbSrc.Sheets[sheetName];
        var range = Utils.decodeRange(original['!ref']);
        var newSheet = {};
        for (var R = range.s.r; R <= range.e.r; R++) {
          for (var C = range.s.c; C <= range.e.c; C++) {
            var addr = Utils.encodeCell(R, C);
            var cellObj = original[addr];
            newSheet[addr] = cellObj ? Object.assign({}, cellObj) : { v: '', t: 's' };
            var custom = styles[sheetName] && styles[sheetName][addr];
            if (custom) {
              newSheet[addr].s = {
                fill: { patternType: 'solid', fgColor: { rgb: custom.bg.replace('#', '') } },
                font: { color: { rgb: custom.font.replace('#', '') } }
              };
            }
          }
        }
        newSheet['!ref'] = original['!ref'];
        XLSX.utils.book_append_sheet(wb, newSheet, sheetName);
      });
      XLSX.writeFile(wb, 'Datos_Exportados.xlsx');
      Toast.show('¡Archivo Excel generado exitosamente!');
    }
    function toCSV() {
      var data = getVisibleData();
      if (!data) return;
      var lines = [data.headers.map(Utils.escapeCsv).join(',')];
      data.dataRows.forEach(function (row) { lines.push(row.map(Utils.escapeCsv).join(',')); });
      Utils.downloadBlob(lines.join('\r\n'), data.sheetName + '_export.csv', 'text/csv;charset=utf-8');
      Toast.show('¡Archivo CSV generado exitosamente!');
    }
    function toTXT() {
      var data = getVisibleData();
      if (!data) return;
      var lines = [data.headers.join('\t')];
      data.dataRows.forEach(function (row) { lines.push(row.join('\t')); });
      Utils.downloadBlob(lines.join('\r\n'), data.sheetName + '_export.txt', 'text/plain;charset=utf-8');
      Toast.show('¡Archivo TXT generado exitosamente!');
    }
    function toHTML() {
      var data = getVisibleData();
      if (!data) return;
      var html = '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>' + Utils.escapeHtml(data.sheetName) +
        '</title><style>body{font-family:Segoe UI,sans-serif;margin:24px}table{border-collapse:collapse;width:100%}' +
        'th,td{border:1px solid #e2e8f0;padding:8px 12px}th{background:#a855f7;color:#fff}' +
        'tr:nth-child(even){background:#f1f5f9}</style></head><body><h1>' + Utils.escapeHtml(data.sheetName) +
        '</h1><table><thead><tr>';
      data.headers.forEach(function (h) { html += '<th>' + Utils.escapeHtml(h) + '</th>'; });
      html += '</tr></thead><tbody>';
      data.dataRows.forEach(function (row) {
        html += '<tr>';
        row.forEach(function (cell) { html += '<td>' + Utils.escapeHtml(cell) + '</td>'; });
        html += '</tr>';
      });
      html += '</tbody></table></body></html>';
      Utils.downloadBlob(html, data.sheetName + '_export.html', 'text/html;charset=utf-8');
      Toast.show('¡Archivo HTML generado exitosamente!');
    }
    function toXML() {
      var data = getVisibleData();
      if (!data) return;
      var safeHeaders = data.headers.map(Utils.sanitizeXmlTag);
      var root = Utils.sanitizeXmlTag(data.sheetName);
      var xml = '<?xml version="1.0" encoding="UTF-8"?>\n<' + root + '>\n';
      data.dataRows.forEach(function (row, idx) {
        xml += '  <row id="' + (idx + 1) + '">\n';
        row.forEach(function (cell, cIdx) {
          var tag = safeHeaders[cIdx] || ('col_' + (cIdx + 1));
          xml += '    <' + tag + '>' + Utils.escapeXml(cell) + '</' + tag + '>\n';
        });
        xml += '  </row>\n';
      });
      xml += '</' + root + '>';
      Utils.downloadBlob(xml, data.sheetName + '_export.xml', 'application/xml;charset=utf-8');
      Toast.show('¡Archivo XML generado exitosamente!');
    }
    function openJsonModal() {
      if (!Store.get('workbook') || !Store.get('sheetName')) { Toast.show('Carga un archivo primero'); return; }
      var sheet = Store.getSheet();
      var range = Store.getRange();
      var container = document.getElementById('jsonColumnList');
      container.innerHTML = '';
      for (var C = range.s.c; C <= range.e.c; C++) {
        var colLetter = Utils.getColLetter(C);
        var first = sheet[Utils.encodeCell(0, C)];
        var headerName = first && first.v ? first.v : 'Columna ' + colLetter;
        var label = document.createElement('label');
        label.className = 'filter-menu__item';
        label.innerHTML = '<input type="checkbox" class="json-col-cb" value="' + C + '" checked>' +
          '<span><b>' + colLetter + ':</b> ' + Utils.escapeHtml(String(headerName)) + '</span>';
        container.appendChild(label);
      }
      var selectAll = document.getElementById('jsonSelectAll');
      selectAll.checked = true;
      selectAll.onclick = function (e) {
        document.querySelectorAll('.json-col-cb').forEach(function (cb) { cb.checked = e.target.checked; });
      };
      document.getElementById('jsonModal').classList.add('modal--visible');
    }
    function closeJsonModal() {
      document.getElementById('jsonModal').classList.remove('modal--visible');
    }
    function processJsonExport() {
      var selectedCols = Array.from(document.querySelectorAll('.json-col-cb:checked')).map(function (cb) {
        return parseInt(cb.value, 10);
      });
      if (!selectedCols.length) { Toast.show('Selecciona al menos una columna'); return; }
      var sheet = Store.getSheet();
      var headers = {};
      selectedCols.forEach(function (c) {
        var cell = sheet[Utils.encodeCell(0, c)];
        headers[c] = cell && cell.v ? String(cell.v).trim() : Utils.getColLetter(c);
      });
      var filtered = Store.get('filteredRowIndices');
      if (!filtered.length) { FilterEngine.rebuildSync(); filtered = Store.get('filteredRowIndices'); }
      var jsonResult = [];
      filtered.forEach(function (R) {
        var rowObj = {};
        var hasData = false;
        selectedCols.forEach(function (c) {
          var cell = sheet[Utils.encodeCell(R, c)];
          var val = cell && cell.v !== undefined ? String(cell.v) : '';
          rowObj[headers[c]] = val;
          if (val !== '') hasData = true;
        });
        if (hasData) jsonResult.push(rowObj);
      });
      Utils.downloadBlob(JSON.stringify(jsonResult, null, 2), Store.get('sheetName') + '_export.json', 'application/json');
      closeJsonModal();
      Toast.show('¡Datos exportados a JSON!');
    }
    return Object.freeze({
      toExcel: toExcel, toCSV: toCSV, toTXT: toTXT, toHTML: toHTML, toXML: toXML,
      openJsonModal: openJsonModal, closeJsonModal: closeJsonModal, processJsonExport: processJsonExport
    });
  })();

  /* ========================================================================
   * CONTEXT MENU MODULE
   * ======================================================================== */
  const ContextMenu = (function () {
    function show(e, type, index) {
      Store.set('contextTarget', { type: type, index: index });
      var menu = document.getElementById('contextMenu');
      var ctxInsert = document.getElementById('ctxInsert');
      var ctxDelete = document.getElementById('ctxDelete');
      var ctxSetHeader = document.getElementById('ctxSetHeader');
      if (type === 'row') {
        ctxInsert.innerHTML = '<span class="material-icons">add</span> Insertar fila arriba';
        ctxDelete.innerHTML = '<span class="material-icons">delete</span> Eliminar fila';
        ctxInsert.classList.remove('context-menu__item--hidden');
        ctxDelete.classList.remove('context-menu__item--hidden');
        ctxSetHeader.classList.remove('context-menu__item--hidden');
      } else if (type === 'col') {
        ctxInsert.innerHTML = '<span class="material-icons">add</span> Insertar columna a la izquierda';
        ctxDelete.innerHTML = '<span class="material-icons">delete</span> Eliminar columna';
        ctxInsert.classList.remove('context-menu__item--hidden');
        ctxDelete.classList.remove('context-menu__item--hidden');
        ctxSetHeader.classList.add('context-menu__item--hidden');
      } else {
        ctxInsert.classList.add('context-menu__item--hidden');
        ctxDelete.classList.add('context-menu__item--hidden');
        ctxSetHeader.classList.add('context-menu__item--hidden');
      }
      menu.style.top = e.clientY + 'px';
      menu.style.left = e.clientX + 'px';
      menu.classList.add('context-menu--visible');
    }
    function hide() {
      var m = document.getElementById('contextMenu');
      if (m) m.classList.remove('context-menu--visible');
    }
    function insertRow(rowIndex) {
      var sheet = Store.getSheet();
      var range = Store.getRange();
      for (var R = range.e.r; R >= rowIndex; R--) {
        for (var C = range.s.c; C <= range.e.c; C++) {
          var oldCell = Utils.encodeCell(R, C);
          var newCell = Utils.encodeCell(R + 1, C);
          if (sheet[oldCell]) { sheet[newCell] = sheet[oldCell]; delete sheet[oldCell]; }
        }
      }
      range.e.r += 1;
      sheet['!ref'] = XLSX.utils.encode_range(range);
      Spreadsheet.renderSheet(Store.get('sheetName'));
      Toast.show('Fila insertada en posición ' + (rowIndex + 1));
    }
    function deleteRow(rowIndex) {
      var sheet = Store.getSheet();
      var range = Store.getRange();
      if (rowIndex === range.s.r) { Toast.show('No se puede eliminar la fila de encabezado'); return; }
      if (range.s.r === range.e.r) { Toast.show('No se puede eliminar la única fila existente'); return; }
      var headerCells = [];
      for (var C = range.s.c; C <= range.e.c; C++) {
        var cell = sheet[Utils.encodeCell(range.s.r, C)];
        headerCells.push(cell ? Object.assign({}, cell) : { v: Utils.getColLetter(C), t: 's' });
      }
      var dataRows = [];
      for (var R = range.s.r + 1; R <= range.e.r; R++) {
        if (R === rowIndex) continue;
        var rowCells = [];
        var hasAny = false;
        for (var C = range.s.c; C <= range.e.c; C++) {
          var cell = sheet[Utils.encodeCell(R, C)];
          if (cell && cell.v !== undefined && cell.v !== '') { rowCells.push(Object.assign({}, cell)); hasAny = true; }
          else rowCells.push(null);
        }
        if (hasAny) dataRows.push(rowCells);
      }
      Object.keys(sheet).forEach(function (key) { if (key[0] !== '!') delete sheet[key]; });
      for (var C = range.s.c; C <= range.e.c; C++) sheet[Utils.encodeCell(0, C)] = headerCells[C - range.s.c];
      dataRows.forEach(function (rowCells, idx) {
        rowCells.forEach(function (cell, cIdx) {
          if (cell) sheet[Utils.encodeCell(idx + 1, range.s.c + cIdx)] = cell;
        });
      });
      sheet['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: range.s.c }, e: { r: Math.max(0, dataRows.length), c: range.e.c } });
      Store.get('cellStyles')[Store.get('sheetName')] = {};
      Store.set('activeFilters', {});
      Spreadsheet.renderSheet(Store.get('sheetName'));
      Toast.show('Fila ' + (rowIndex + 1) + ' eliminada');
    }
    function insertColumn(colIndex) {
      var sheet = Store.getSheet();
      var range = Store.getRange();
      for (var C = range.e.c; C >= colIndex; C--) {
        for (var R = range.s.r; R <= range.e.r; R++) {
          var oldCell = Utils.encodeCell(R, C);
          var newCell = Utils.encodeCell(R, C + 1);
          if (sheet[oldCell]) { sheet[newCell] = sheet[oldCell]; delete sheet[oldCell]; }
        }
      }
      range.e.c += 1;
      sheet['!ref'] = XLSX.utils.encode_range(range);
      Spreadsheet.renderSheet(Store.get('sheetName'));
      Toast.show('Columna insertada en ' + Utils.getColLetter(colIndex));
    }
    function deleteColumn(colIndex) {
      var sheet = Store.getSheet();
      var range = Store.getRange();
      if (range.s.c === range.e.c) { Toast.show('No se puede eliminar la única columna existente'); return; }
      for (var R = range.s.r; R <= range.e.r; R++) delete sheet[Utils.encodeCell(R, colIndex)];
      for (var C = colIndex + 1; C <= range.e.c; C++) {
        for (var R = range.s.r; R <= range.e.r; R++) {
          var oldCell = Utils.encodeCell(R, C);
          var newCell = Utils.encodeCell(R, C - 1);
          if (sheet[oldCell]) { sheet[newCell] = sheet[oldCell]; delete sheet[oldCell]; }
        }
      }
      range.e.c -= 1;
      sheet['!ref'] = XLSX.utils.encode_range(range);
      Spreadsheet.renderSheet(Store.get('sheetName'));
      Toast.show('Columna ' + Utils.getColLetter(colIndex) + ' eliminada');
    }
    function setAsHeader(rowIndex) {
      var sheet = Store.getSheet();
      var range = Store.getRange();
      if (rowIndex === range.s.r) { Toast.show('Esta fila ya es el encabezado'); return; }
      var newHeaders = [];
      for (var C = range.s.c; C <= range.e.c; C++) {
        var cell = sheet[Utils.encodeCell(rowIndex, C)];
        if (cell && cell.v !== undefined && cell.v !== '') {
          newHeaders.push({ v: cell.v, t: cell.t || (typeof cell.v === 'number' ? 'n' : 's') });
        } else newHeaders.push({ v: Utils.getColLetter(C), t: 's' });
      }
      var dataRows = [];
      for (var R = range.s.r; R <= range.e.r; R++) {
        if (R === rowIndex || R === range.s.r) continue;
        var rowCells = [];
        var hasAny = false;
        for (var C = range.s.c; C <= range.e.c; C++) {
          var cell = sheet[Utils.encodeCell(R, C)];
          if (cell && cell.v !== undefined && cell.v !== '') { rowCells.push(Object.assign({}, cell)); hasAny = true; }
          else rowCells.push(null);
        }
        if (hasAny) dataRows.push(rowCells);
      }
      Object.keys(sheet).forEach(function (key) { if (key[0] !== '!') delete sheet[key]; });
      for (var C = range.s.c; C <= range.e.c; C++) sheet[Utils.encodeCell(0, C)] = newHeaders[C - range.s.c];
      dataRows.forEach(function (rowCells, idx) {
        rowCells.forEach(function (cell, cIdx) {
          if (cell) sheet[Utils.encodeCell(idx + 1, range.s.c + cIdx)] = cell;
        });
      });
      sheet['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: range.s.c }, e: { r: Math.max(0, dataRows.length), c: range.e.c } });
      Store.get('cellStyles')[Store.get('sheetName')] = {};
      Store.set('activeFilters', {});
      Spreadsheet.renderSheet(Store.get('sheetName'));
      Toast.show('Fila ' + (rowIndex + 1) + ' establecida como nuevo encabezado');
    }
    function handleAction(action) {
      hide();
      var target = Store.get('contextTarget');
      if (action === 'copy') Spreadsheet.copyToClipboard();
      else if (action === 'insert') {
        if (target.type === 'row') insertRow(target.index);
        if (target.type === 'col') insertColumn(target.index);
      } else if (action === 'delete') {
        if (target.type === 'row') deleteRow(target.index);
        if (target.type === 'col') deleteColumn(target.index);
      } else if (action === 'setHeader') {
        if (target.type === 'row') setAsHeader(target.index);
      }
    }
    function bind() {
      EventBus.on('sheet:rendered', function () {
        var table = document.querySelector('.spreadsheet__table');
        if (!table) return;
        table.addEventListener('contextmenu', function (e) {
          e.preventDefault();
          var rowHeader = e.target.closest('.spreadsheet__row-header');
          var colHeader = e.target.closest('.spreadsheet__header');
          var td = e.target.closest('td.spreadsheet__cell');
          if (rowHeader) {
            var r = parseInt(rowHeader.getAttribute('data-r'), 10);
            Spreadsheet.selectRow(r);
            show(e, 'row', r);
          } else if (colHeader) {
            var c = parseInt(colHeader.getAttribute('data-c'), 10);
            Spreadsheet.selectColumn(c);
            show(e, 'col', c);
          } else if (td) show(e, 'cell', null);
        });
      });
      var menu = document.getElementById('contextMenu');
      if (menu) menu.addEventListener('click', function (e) {
        var item = e.target.closest('.context-menu__item');
        if (item && item.dataset.action) handleAction(item.dataset.action);
      });
    }
    return Object.freeze({ bind: bind, hide: hide, handleAction: handleAction });
  })();

  /* ========================================================================
   * COLUMN FILTER MODULE
   * ======================================================================== */
  const ColumnFilter = (function () {
    function open(colIndex, buttonEl) {
      Store.set('currentFilterCol', colIndex);
      var filterMenu = document.getElementById('filterMenu');
      var filterList = document.getElementById('filterList');
      var sheet = Store.getSheet();
      var range = Store.getRange();
      var uniqueValues = new Set();
      for (var R = range.s.r + 1; R <= range.e.r; R++) {
        var cell = sheet[Utils.encodeCell(R, colIndex)];
        var val = cell && cell.v !== undefined && cell.v !== '' ? String(cell.v) : '(Vacío)';
        uniqueValues.add(val);
      }
      var sorted = Array.from(uniqueValues).sort();
      var currentAllowed = Store.get('activeFilters')[colIndex];
      var listHtml = '<label class="filter-menu__item"><input type="checkbox" id="selectAllFilters"' +
        (!currentAllowed ? ' checked' : '') + '><b>(Seleccionar todo)</b></label>';
      sorted.forEach(function (val) {
        var isChecked = !currentAllowed || currentAllowed.has(val);
        listHtml += '<label class="filter-menu__item"><input type="checkbox" class="val-checkbox" value="' +
          Utils.escapeHtml(val) + '"' + (isChecked ? ' checked' : '') + '><span>' + Utils.escapeHtml(val) + '</span></label>';
      });
      filterList.innerHTML = listHtml;
      var selectAllCb = document.getElementById('selectAllFilters');
      var valCbs = document.querySelectorAll('.val-checkbox');
      selectAllCb.addEventListener('change', function (e) {
        valCbs.forEach(function (cb) { cb.checked = e.target.checked; });
      });
      valCbs.forEach(function (cb) {
        cb.addEventListener('change', function () {
          selectAllCb.checked = Array.from(valCbs).every(function (c) { return c.checked; });
        });
      });
      var rect = buttonEl.getBoundingClientRect();
      filterMenu.style.top = (rect.bottom + window.scrollY + 2) + 'px';
      filterMenu.style.left = Math.min(rect.left + window.scrollX, window.innerWidth - 260) + 'px';
      filterMenu.classList.add('filter-menu--visible');
    }
    function apply() {
      var selectAllCb = document.getElementById('selectAllFilters');
      var valCbs = document.querySelectorAll('.val-checkbox');
      var col = Store.get('currentFilterCol');
      var filters = Store.get('activeFilters');
      if (selectAllCb.checked) delete filters[col];
      else {
        var allowed = new Set();
        valCbs.forEach(function (cb) { if (cb.checked) allowed.add(cb.value); });
        filters[col] = allowed;
      }
      var btn = document.querySelector('.spreadsheet__filter-btn[data-col="' + col + '"]');
      if (btn) {
        if (filters[col] !== undefined) btn.classList.add('spreadsheet__filter-btn--active');
        else btn.classList.remove('spreadsheet__filter-btn--active');
      }
      FilterEngine.apply();
    }
    function bind() {
      EventBus.on('sheet:rendered', function () {
        document.querySelectorAll('.spreadsheet__filter-btn').forEach(function (btn) {
          btn.addEventListener('click', function (e) {
            e.stopPropagation();
            open(parseInt(btn.getAttribute('data-col'), 10), btn);
          });
        });
      });
      var cancel = document.getElementById('filterCancel');
      if (cancel) cancel.addEventListener('click', function () {
        document.getElementById('filterMenu').classList.remove('filter-menu--visible');
      });
      var applyBtn = document.getElementById('filterApply');
      if (applyBtn) applyBtn.addEventListener('click', function () {
        apply();
        document.getElementById('filterMenu').classList.remove('filter-menu--visible');
      });
    }
    return Object.freeze({ bind: bind, open: open });
  })();

  /* ========================================================================
   * RIBBON MODULE
   * ======================================================================== */
  const Ribbon = (function () {
    function switchTab(tabName) {
      document.querySelectorAll('.ribbon__tab').forEach(function (btn) {
        btn.classList.toggle('ribbon__tab--active', btn.dataset.tab === tabName);
      });
      document.querySelectorAll('.ribbon__panel').forEach(function (panel) {
        panel.classList.toggle('ribbon__panel--active', panel.dataset.panel === tabName);
      });
    }
    function bind() {
      document.querySelectorAll('.ribbon__tab').forEach(function (btn) {
        btn.addEventListener('click', function () { switchTab(btn.dataset.tab); });
      });
      var fileInput = document.getElementById('excelFile');
      if (fileInput) fileInput.addEventListener('change', function (e) {
        if (e.target.files.length > 0) {
          FileLoader.processFile(e.target.files[0]);
          e.target.value = '';
        }
      });
      document.querySelectorAll('[data-action]').forEach(function (btn) {
        if (btn.closest('.context-menu')) return;
        btn.addEventListener('click', function () {
          switch (btn.dataset.action) {
            case 'new-sheet': Tabs.createBlankSheet(); break;
            case 'export-excel': Exporter.toExcel(); break;
            case 'export-json': Exporter.openJsonModal(); break;
            case 'export-csv': Exporter.toCSV(); break;
            case 'export-txt': Exporter.toTXT(); break;
            case 'export-html': Exporter.toHTML(); break;
            case 'export-xml': Exporter.toXML(); break;
          }
        });
      });
      document.querySelectorAll('.color-picker__swatch').forEach(function (swatch) {
        swatch.addEventListener('click', function () { Spreadsheet.applyColor(swatch.dataset.color); });
      });
      var closeBtn = document.getElementById('jsonModalClose');
      if (closeBtn) closeBtn.addEventListener('click', Exporter.closeJsonModal);
      var cancelBtn = document.getElementById('jsonModalCancel');
      if (cancelBtn) cancelBtn.addEventListener('click', Exporter.closeJsonModal);
      var dlBtn = document.getElementById('jsonModalDownload');
      if (dlBtn) dlBtn.addEventListener('click', Exporter.processJsonExport);
    }
    return Object.freeze({ bind: bind, switchTab: switchTab });
  })();

  /* ========================================================================
   * APP ORCHESTRATOR
   * ======================================================================== */
  class App {
    init() {
      Explorer.init();
      Ribbon.bind();
      ContextMenu.bind();
      ColumnFilter.bind();
      Tabs.render();
      this._bindGlobalEvents();
      this._bindDragDrop();
      this._bindGlobalFilter();
      Spreadsheet.renderDropzone();
    }
    _bindGlobalEvents() {
      document.addEventListener('keydown', function (e) {
        if (Store.get('editing')) return;
        var hasSel = Store.hasSelection();
        if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
          if (hasSel) { e.preventDefault(); Spreadsheet.copyToClipboard(); }
          return;
        }
        if (['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].indexOf(e.key) !== -1) {
          if (hasSel && Spreadsheet.getAnchorTd()) { e.preventDefault(); Spreadsheet.navigate(e.key); }
          return;
        }
        if (e.key === 'Enter' && Spreadsheet.getAnchorTd()) {
          e.preventDefault();
          Spreadsheet.startEdit(Spreadsheet.getAnchorTd(), false);
          return;
        }
        if ((e.key === 'Delete' || e.key === 'Backspace') && hasSel) {
          e.preventDefault();
          Spreadsheet.clearSelectedContent();
          return;
        }
        if (e.key === 'F2' && Spreadsheet.getAnchorTd()) {
          e.preventDefault();
          Spreadsheet.startEdit(Spreadsheet.getAnchorTd(), false);
          return;
        }
        if (Spreadsheet.getAnchorTd() && hasSel && !e.ctrlKey && !e.metaKey && !e.altKey && e.key.length === 1) {
          e.preventDefault();
          Spreadsheet.startEdit(Spreadsheet.getAnchorTd(), true, e.key);
        }
      });
      document.addEventListener('click', function (e) {
        var filterMenu = document.getElementById('filterMenu');
        if (filterMenu && !filterMenu.contains(e.target) && !e.target.classList.contains('spreadsheet__filter-btn')) {
          filterMenu.classList.remove('filter-menu--visible');
        }
        var contextMenu = document.getElementById('contextMenu');
        if (contextMenu && !contextMenu.contains(e.target)) {
          contextMenu.classList.remove('context-menu--visible');
        }
      });
    }
    _bindDragDrop() {
      var container = document.getElementById('tableContainer');
      if (!container) return;
      ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(function (ev) {
        container.addEventListener(ev, function (e) { e.preventDefault(); e.stopPropagation(); }, false);
        document.body.addEventListener(ev, function (e) { e.preventDefault(); e.stopPropagation(); }, false);
      });
      ['dragenter', 'dragover'].forEach(function (ev) {
        container.addEventListener(ev, function () { container.classList.add('spreadsheet--dragover'); }, false);
      });
      ['dragleave', 'drop'].forEach(function (ev) {
        container.addEventListener(ev, function () { container.classList.remove('spreadsheet--dragover'); }, false);
      });
      container.addEventListener('drop', function (e) {
        if (e.dataTransfer.files.length > 0) FileLoader.processFile(e.dataTransfer.files[0]);
      }, false);
      var dz = document.getElementById('dropzoneContainer');
      if (dz) dz.addEventListener('click', function () {
        var fi = document.getElementById('excelFile');
        if (fi) fi.click();
      });
    }
    _bindGlobalFilter() {
      var input = document.getElementById('globalFilterInput');
      var clearBtn = document.getElementById('globalFilterClear');
      if (!input || !clearBtn) return;
      function deselect() {
        if (Store.get('editing')) Spreadsheet.commitEdit(true);
        Spreadsheet.clearSelectionUI();
        document.getElementById('activeCellAddress').textContent = 'A1';
        document.getElementById('formulaInput').value = '';
        Spreadsheet.calculateStats();
      }
      input.addEventListener('focus', deselect);
      input.addEventListener('mousedown', deselect);
      input.addEventListener('input', function () {
        clearBtn.classList.toggle('filter-bar__clear--visible', input.value.trim().length > 0);
        FilterEngine.scheduleApply();
      });
      clearBtn.addEventListener('click', function () {
        input.value = '';
        clearBtn.classList.remove('filter-bar__clear--visible');
        FilterEngine.apply();
        input.focus();
      });
    }
  }

  document.addEventListener('DOMContentLoaded', function () {
    new App().init();
  });

})(typeof window !== 'undefined' ? window : this);
