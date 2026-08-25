// Turn "some CSV text" into a clean table.
//
// Port of budget_core/csvsniff.py. Handles the parts of real world exports
// that break naive CSV parsing: sep=; hint lines, semicolon/comma/tab/pipe
// separators, bank preambles above the real header, missing headers and
// ragged rows. Character-encoding detection (BOM/UTF-16/cp1252) happens in
// the browser before the text ever reaches here - see Encoding.gs - so
// this module only works with an already-decoded JS string, unlike the
// Python version which also reads raw bytes.

var CANDIDATE_DELIMITERS = [";", ",", "\t", "|"];

// Words that give away a header row. Compared token by token on folded text.
var HEADER_WORDS = {};
[
  "dato", "date", "datum", "bogfort", "bogforingsdato", "posteringsdato",
  "transaktionsdato", "valor", "valordato", "rentedato", "tid", "time",
  "tekst", "text", "beskrivelse", "description", "posteringstekst",
  "narrative", "details", "detaljer", "modtager", "afsender", "payee",
  "merchant", "reference", "memo", "note", "navn", "name", "titel",
  "belob", "amount", "sum", "value", "betrag", "montant", "beloeb",
  "saldo", "balance", "beholdning", "konto", "account", "kontonummer",
  "valuta", "currency", "type", "kategori", "category", "status",
  "debet", "kredit", "debit", "credit", "ind", "ud", "indsat", "haevet",
  "indbetaling", "udbetaling", "withdrawal", "deposit", "posting",
  "transaktion", "transaction", "art", "nummer", "number", "id",
  "afstemt", "gebyr", "fee",
].forEach(function (w) { HEADER_WORDS[w] = true; });

// Sniffing never looks at more than this many rows.
var SAMPLE_ROWS = 200;

/** A rectangular table of strings plus what we learned while reading it. */
class Table {
  constructor(rows, header, opts) {
    opts = opts || {};
    this.rows = rows;
    this.header = header || null;
    this.encoding = opts.encoding || "utf-8";
    this.delimiter = opts.delimiter || ";";
    this.preamble = opts.preamble || [];
    this.notes = opts.notes || [];
    this.source = opts.source || "";
    var widths = rows.map(function (r) { return r.length; });
    widths.push(header ? header.length : 0);
    this.n_columns = Math.max.apply(null, widths.concat([0]));
  }

  /** All values of one column (as strings). */
  column(index, limit) {
    var rows = limit === undefined || limit === null ? this.rows : this.rows.slice(0, limit);
    return rows.map(function (r) { return index < r.length ? r[index] : ""; });
  }

  header_name(index) {
    if (this.header && index < this.header.length) return this.header[index].trim();
    return "Kolonne " + (index + 1);
  }

  header_names() {
    var out = [];
    for (var i = 0; i < this.n_columns; i++) out.push(this.header_name(i));
    return out;
  }

  sample(count) {
    return this.rows.slice(0, count === undefined ? 5 : count);
  }

  get length() {
    return this.rows.length;
  }

  describe() {
    return this.rows.length + " rækker, " + this.n_columns + " kolonner, tegnsæt " +
      this.encoding + ", skilletegn " +
      (this.delimiter === "\t" ? "TAB" : "\"" + this.delimiter + "\"");
  }

  /** Build a table from already split rows (CSV rows or sheet cells). */
  static fromRows(rawRows, hasHeader, opts) {
    opts = opts || {};
    var rows = rawRows.map(function (row) {
      return row.map(function (c) { return c === null || c === undefined ? "" : String(c); });
    });
    var notes = (opts.notes || []).slice();
    var found = _findStart(rows);
    var start = found[0], nColumns = found[1];
    var preamble = rows.slice(0, start).filter(function (r) { return !isBlankRow(r); });
    if (preamble.length) notes.push("Sprang " + preamble.length + " indledende linje(r) over.");

    var body = rows.slice(start).filter(function (r) { return !isBlankRow(r); });
    if (!body.length) {
      return new Table([], null, { preamble: preamble, notes: notes, encoding: opts.encoding,
        delimiter: opts.delimiter, source: opts.source });
    }

    var header = null;
    var first = body[0];
    var useHeader = hasHeader === undefined || hasHeader === null ? looksLikeHeader(first, body.slice(1, 21)) : hasHeader;
    if (useHeader) {
      header = first.map(function (c) { return String(c).trim(); });
      body = body.slice(1);
      header = _uniquify(header);
    } else {
      notes.push("Ingen overskriftsrække fundet - kolonnerne navngives automatisk.");
    }

    var ragged = 0;
    var fixed = [];
    body.forEach(function (row) {
      if (row.length < nColumns) {
        if (row.length < 2) return;
        ragged += 1;
        row = row.concat(new Array(nColumns - row.length).fill(""));
      }
      fixed.push(row.slice());
    });
    if (ragged) notes.push(ragged + " række(r) havde færre kolonner end resten og blev fyldt ud.");

    return new Table(fixed, header, { preamble: preamble, notes: notes, encoding: opts.encoding,
      delimiter: opts.delimiter, source: opts.source });
  }
}

/** Make header names unique and non empty. */
function _uniquify(names) {
  var seen = {};
  var out = [];
  names.forEach(function (name, i) {
    name = (name || "").trim() || "Kolonne " + (i + 1);
    if (seen.hasOwnProperty(name)) {
      seen[name] += 1;
      name = name + " (" + seen[name] + ")";
    } else {
      seen[name] = 1;
    }
    out.push(name);
  });
  return out;
}

// --------------------------------------------------------------------------
// Delimiter
// --------------------------------------------------------------------------

/**
 * Split one CSV line respecting double-quoted fields (RFC 4180 style, "" is
 * an escaped quote). Mirrors Python's csv.reader for a single delimiter.
 */
function _splitCsvLine(line, delimiter) {
  var fields = [];
  var field = "";
  var inQuotes = false;
  var atFieldStart = true;
  for (var i = 0; i < line.length; i++) {
    var ch = line[i];
    if (inQuotes) {
      if (ch === '"') {
        if (line[i + 1] === '"') { field += '"'; i++; } else { inQuotes = false; }
      } else {
        field += ch;
      }
    } else if (atFieldStart && ch === " ") {
      continue; // skipinitialspace: spaces right after the delimiter (or line start)
    } else if (atFieldStart && ch === '"') {
      inQuotes = true;
      atFieldStart = false;
    } else if (ch === delimiter) {
      fields.push(field);
      field = "";
      atFieldStart = true;
    } else {
      field += ch;
      atFieldStart = false;
    }
  }
  fields.push(field);
  return fields;
}

function _parseCsv(text, delimiter) {
  var lines = text.split("\n");
  if (lines.length && lines[lines.length - 1] === "") lines.pop();
  return lines.map(function (line) { return _splitCsvLine(line.replace(/\r$/, ""), delimiter); });
}

/** Pick the separator that yields the most consistent, widest table. */
function detectDelimiter(text, candidates) {
  candidates = candidates || CANDIDATE_DELIMITERS;
  var lines = text.split(/\r\n|\r|\n/).filter(function (l) { return l.trim(); }).slice(0, 60);
  if (!lines.length) return ";";

  var best = null;
  var bestScore = 0.0;
  candidates.forEach(function (delimiter) {
    var rows = lines.map(function (line) { return _splitCsvLine(line, delimiter); });
    var widths = rows.filter(function (r) { return r.some(function (c) { return String(c).trim(); }); })
      .map(function (r) { return r.length; });
    if (!widths.length) return;
    var counts = {};
    widths.forEach(function (w) { counts[w] = (counts[w] || 0) + 1; });
    var width = 0, freq = -1;
    Object.keys(counts).forEach(function (w) {
      var wi = parseInt(w, 10);
      if (counts[w] > freq || (counts[w] === freq && wi > width)) { freq = counts[w]; width = wi; }
    });
    if (width < 2) return;
    var score = (freq / widths.length) * Math.min(width, 12);
    if (score > bestScore + 1e-9) { bestScore = score; best = delimiter; }
  });
  return best || ";";
}

var _SEP_HINT_PREFIXES = ["sep=", "SEP="];

/** Consume Excel's sep=; first line if present. */
function _stripSepHint(text) {
  var idx = text.indexOf("\n");
  if (idx === -1) return [text, null];
  var first = text.slice(0, idx);
  var rest = text.slice(idx + 1);
  var stripped = first.trim().replace(/^﻿/, "");
  for (var i = 0; i < _SEP_HINT_PREFIXES.length; i++) {
    var prefix = _SEP_HINT_PREFIXES[i];
    if (stripped.indexOf(prefix) === 0 && stripped.length === prefix.length + 1) {
      return [rest, stripped[prefix.length]];
    }
  }
  return [text, null];
}

// --------------------------------------------------------------------------
// Header detection
// --------------------------------------------------------------------------

/** Decide whether row is a header rather than the first data row. */
function looksLikeHeader(row, following) {
  var cells = row.map(function (c) { return String(c).trim(); });
  var nonEmpty = cells.filter(function (c) { return c; });
  if (!nonEmpty.length) return false;

  if (cells.some(function (c) { return looksLikeDate(c); })) return false;

  var hits = 0;
  nonEmpty.forEach(function (cell) {
    var folded = fold(cell);
    if (HEADER_WORDS.hasOwnProperty(folded)) { hits += 1; return; }
    var tokens = folded.split("/").join(" ").split("-").join(" ").split(" ").filter(function (t) { return t; });
    if (tokens.some(function (t) { return HEADER_WORDS.hasOwnProperty(t); })) hits += 1;
  });
  if (hits >= 2 || (hits === 1 && nonEmpty.length <= 3)) return true;

  following = (following || []).filter(function (r) { return !isBlankRow(r); });
  if (following.length) {
    var withDate = following.filter(function (r) { return r.some(function (c) { return looksLikeDate(c); }); }).length;
    if (withDate >= Math.max(1, Math.floor(following.length / 2))) return true; // rows below have dates, this one does not
    var numericBelow = following.filter(function (r) { return r.some(function (c) { return looksLikeAmount(c); }); }).length;
    if (numericBelow >= Math.max(1, Math.floor(following.length / 2)) &&
      !cells.some(function (c) { return looksLikeAmount(c); })) return true;
  }
  return hits > 0;
}

/** Find where the real table starts and how wide it is. */
function _findStart(rows) {
  var counts = {};
  rows.forEach(function (r) { if (!isBlankRow(r)) counts[r.length] = (counts[r.length] || 0) + 1; });
  var keys = Object.keys(counts);
  if (!keys.length) return [0, 0];
  var width = 0, freq = -1;
  keys.forEach(function (w) {
    var wi = parseInt(w, 10);
    if (counts[w] > freq || (counts[w] === freq && wi > width)) { freq = counts[w]; width = wi; }
  });

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (isBlankRow(row) || row.length !== width) continue;
    var following = rows.slice(i + 1, i + 5).filter(function (r) { return !isBlankRow(r); });
    if (!following.length || following.filter(function (r) { return r.length === width; }).length >=
      Math.max(1, following.length - 1)) {
      return [i, width];
    }
  }
  return [0, width];
}

// --------------------------------------------------------------------------
// Public entry point
// --------------------------------------------------------------------------

/**
 * Turn already-decoded CSV text into a Table. Encoding is decided by the
 * caller (the browser dialog, via TextDecoder - see Encoding.gs); pass the
 * resulting JS string plus the encoding name it used (for Table.describe()).
 */
function readTableFromText(text, opts) {
  opts = opts || {};
  text = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  var stripped = _stripSepHint(text);
  text = stripped[0];
  var hinted = stripped[1];
  var delimiter = opts.delimiter;
  if (hinted && !delimiter) delimiter = hinted;
  if (!delimiter) delimiter = detectDelimiter(text);

  var rawRows = _parseCsv(text, delimiter);
  return Table.fromRows(rawRows, opts.hasHeader, {
    encoding: opts.encoding || "utf-8", delimiter: delimiter, source: opts.source || "",
  });
}

if (typeof module !== "undefined") {
  module.exports = {
    Table: Table, readTableFromText: readTableFromText, detectDelimiter: detectDelimiter,
    looksLikeHeader: looksLikeHeader, CANDIDATE_DELIMITERS: CANDIDATE_DELIMITERS,
  };
  var fold = require("./TextUtils.gs").fold;
  var isBlankRow = require("./TextUtils.gs").isBlankRow;
  var looksLikeDate = require("./Parsing.gs").looksLikeDate;
  var looksLikeAmount = require("./Parsing.gs").looksLikeAmount;
}
