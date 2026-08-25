// Small text helpers shared by the parsing/detection modules.
//
// Port of budget_core/textutils.py - kept behaviourally identical.

var _LETTER_MAP = {
  "æ": "ae", "ø": "o", "ß": "ss", "đ": "d", "ł": "l", "þ": "th",
};

var _WS_RE = /\s+/g;

/**
 * Normalise text for keyword comparisons: case, accents and Nordic
 * letters are removed so "Café Øst", "CAFE OST" and "cafe øst" compare
 * equal.
 */
function fold(text) {
  if (text === null || text === undefined) return "";
  var s = String(text);
  s = s.normalize("NFKD").replace(/[\u0300-\u036f]/g, "");
  s = s.toLowerCase();
  for (var src in _LETTER_MAP) {
    if (s.indexOf(src) !== -1) {
      s = s.split(src).join(_LETTER_MAP[src]);
    }
  }
  return s.replace(_WS_RE, " ").trim();
}

/** Collapse whitespace but keep the original characters. */
function squeeze(text) {
  if (text === null || text === undefined) return "";
  return String(text).replace(_WS_RE, " ").trim();
}

/** True when every cell in row is empty or whitespace. */
function isBlankRow(row) {
  for (var i = 0; i < row.length; i++) {
    if (String(row[i]).trim()) return false;
  }
  return true;
}

if (typeof module !== "undefined") {
  module.exports = { fold: fold, squeeze: squeeze, isBlankRow: isBlankRow };
}
