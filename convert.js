/* Conversion Excel Twin Jet -> CSV FlightLog. Tout se fait localement dans le navigateur. */
(function (root) {
  'use strict';

  var COLS = ['Date', 'Registration', 'Departure', 'Arrival', 'SchedOut', 'SchedIn', 'Out', 'In', 'Off', 'On',
    'Takeoff', 'Day Landing', 'Night Landing', 'Approaches', 'Total Block', 'Total Flight', 'Night Time',
    'Distance', 'Pax', 'Payload', 'Fuel', 'Fuel Prev', 'Flight Number', 'PIC Time', 'SIC Time', 'DUAL Time',
    'INS Time', 'Sim Time', 'PIC Name', 'Notes', 'Function'];

  var LBS_TO_T = 0.45359237 / 1000;

  // ---------- utilitaires ----------
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function fmtHM(m) { return pad(Math.floor(m / 60)) + ':' + pad(m % 60); }
  function norm(s) { return String(s == null ? '' : s).trim(); }
  function key(s) { return norm(s).toLowerCase().replace(/[^a-z0-9]/g, ''); }

  // Heure/durée -> minutes. Accepte fraction de jour Excel, "HH:MM", "HH:MM:SS", "H.MM".
  function toMinutes(v) {
    if (v === '' || v == null) return null;
    if (typeof v === 'number') {
      var x = v >= 1 ? v - Math.floor(v) : v;   // date+heure Excel : on garde la partie heure
      return Math.round(x * 1440) % 1440;
    }
    var s = norm(v).replace('h', ':').replace('.', ':');
    var m = s.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
    if (!m) return null;
    return parseInt(m[1], 10) * 60 + parseInt(m[2], 10) + (m[3] && parseInt(m[3], 10) >= 30 ? 1 : 0);
  }

  // Durée (peut dépasser 24 h en théorie, mais ici < 24 h)
  function durMinutes(v) {
    if (typeof v === 'number') return Math.round(v * 1440);
    return toMinutes(v);
  }

  // Date -> "AAAA-MM-JJ". Accepte numéro de série Excel, "JJ/MM/AA", "JJ/MM/AAAA", "AAAA-MM-JJ".
  function toISODate(v) {
    if (v === '' || v == null) return null;
    if (typeof v === 'number') {
      var d = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000);
      return d.getUTCFullYear() + '-' + pad(d.getUTCMonth() + 1) + '-' + pad(d.getUTCDate());
    }
    var s = norm(v), m;
    if ((m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/))) return m[1] + '-' + pad(+m[2]) + '-' + pad(+m[3]);
    if ((m = s.match(/^(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})$/))) {
      var y = +m[3]; if (y < 100) y += 2000;
      return y + '-' + pad(+m[2]) + '-' + pad(+m[1]);
    }
    return null;
  }

  // ---------- astronomie (nuit EASA : soleil < -6°) ----------
  function sunElev(lat, lon, tMs) {
    var d = (tMs - Date.UTC(2000, 0, 1, 12)) / 86400000;
    var rad = Math.PI / 180;
    var g = ((357.529 + 0.98560028 * d) % 360) * rad;
    var q = (280.459 + 0.98564736 * d) % 360;
    var L = (q + 1.915 * Math.sin(g) + 0.020 * Math.sin(2 * g)) * rad;
    var e = (23.439 - 0.00000036 * d) * rad;
    var ra = Math.atan2(Math.cos(e) * Math.sin(L), Math.cos(L));
    var dec = Math.asin(Math.sin(e) * Math.sin(L));
    var gmst = ((18.697374558 + 24.06570982441908 * d) % 24 + 24) % 24;
    var ha = (gmst * 15 + lon) * rad - ra;
    var la = lat * rad;
    return Math.asin(Math.sin(la) * Math.sin(dec) + Math.cos(la) * Math.cos(dec) * Math.cos(ha)) / rad;
  }

  function gcNM(a, b) {
    var rad = Math.PI / 180;
    var la1 = a[0] * rad, la2 = b[0] * rad, dl = (b[1] - a[1]) * rad;
    var c = Math.pow(Math.sin((la2 - la1) / 2), 2) + Math.cos(la1) * Math.cos(la2) * Math.pow(Math.sin(dl / 2), 2);
    return 3440.065 * 2 * Math.asin(Math.sqrt(c));
  }

  function utcMs(isoDate, minutes) {
    var p = isoDate.split('-');
    return Date.UTC(+p[0], +p[1] - 1, +p[2]) + minutes * 60000;
  }

  // Minutes de nuit sur le temps bloc, position interpolée entre les deux terrains
  function nightMinutes(isoDate, outMin, blockMin, A, B) {
    var t0 = utcMs(isoDate, outMin), n = 0;
    for (var i = 0; i < blockMin; i++) {
      var f = (i + 0.5) / blockMin;
      var lat = A[0] + f * (B[0] - A[0]), lon = A[1] + f * (B[1] - A[1]);
      if (sunElev(lat, lon, t0 + (i + 0.5) * 60000) < -6) n++;
    }
    return n;
  }

  // ---------- lecture de l'export FlightLog existant ----------
  function parseCSV(text) {
    text = text.replace(/^﻿/, '');
    var rows = [], row = [], field = '', q = false;
    for (var i = 0; i < text.length; i++) {
      var c = text[i];
      if (q) {
        if (c === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else q = false; }
        else field += c;
      } else if (c === '"') q = true;
      else if (c === ',') { row.push(field); field = ''; }
      else if (c === '\n' || c === '\r') {
        if (c === '\r' && text[i + 1] === '\n') i++;
        row.push(field); rows.push(row); row = []; field = '';
      } else field += c;
    }
    if (field !== '' || row.length) { row.push(field); rows.push(row); }
    var head = rows.shift() || [];
    return rows.filter(function (r) { return r.length > 1; }).map(function (r) {
      var o = {}; head.forEach(function (h, j) { o[h] = r[j] || ''; }); return o;
    });
  }

  function indexExisting(records) {
    var idx = {};
    (records || []).forEach(function (r) {
      var k = r['Date'] + '|' + norm(r['Flight Number']).toUpperCase();
      // "prévu" = vol importé depuis le planning (pas d'immat, heures réelles = heures prévues)
      var planned = !norm(r['Registration']) && r['Out'] === r['SchedOut'] && r['In'] === r['SchedIn'];
      idx[k] = { schedOut: r['SchedOut'], schedIn: r['SchedIn'], planned: planned };
    });
    return idx;
  }

  // ---------- conversion ----------
  var ALIASES = {
    date: ['date'], imm: ['imm', 'immat', 'registration', 'reg'], vol: ['vol', 'flight', 'flightnumber', 'nvol'],
    dep: ['dep', 'depart', 'departure'], arr: ['arr', 'arrivee', 'arrival'],
    depblc: ['depblc', 'out', 'offblock'], arrblc: ['arrblc', 'in', 'onblock'], tpsblc: ['tpsblc', 'block', 'tempsbloc'],
    dec: ['dec', 'decollage', 'off', 'takeoff'], att: ['att', 'atterrissage', 'on', 'landing'],
    tpsvol: ['tpsvol', 'tempsvol', 'flighttime'], conso: ['conso', 'fuel'], pax: ['paxfret', 'pax'],
    cdb: ['cdb', 'pic'], opl: ['opl', 'sic', 'fo']
  };

  function mapHeaders(headers) {
    var map = {};
    headers.forEach(function (h) {
      var k = key(h);
      Object.keys(ALIASES).forEach(function (f) { if (!map[f] && ALIASES[f].indexOf(k) >= 0) map[f] = h; });
    });
    return map;
  }

  function flightNumber(v) {
    var s = norm(v).replace(/\.0+$/, '').toUpperCase();
    if (!s) return '';
    if (/^\d+$/.test(s)) return 'T7' + s;   // 5417 -> T75417
    return s;                                // ITI2911, T75417... laissés tels quels
  }

  function registration(v) {
    var s = norm(v).toUpperCase().replace(/\s/g, '');
    if (!s || s.indexOf('-') >= 0) return s;
    return s[0] === 'F' || s[0] === 'D' || s[0] === 'G' || s[0] === 'I' ? s[0] + '-' + s.slice(1) : s;
  }

  /**
   * rows : tableau d'objets {entête Excel: valeur}
   * opts : { trigram, picName, existing (index), airports }
   * Retourne [{ data: {col: val}, status, warnings: [] }]
   */
  function convert(rows, opts) {
    var airports = opts.airports || {};
    var trig = norm(opts.trigram).toUpperCase();
    var existing = opts.existing || {};
    if (!rows.length) return [];
    var map = mapHeaders(Object.keys(rows[0]));
    var missing = ['date', 'vol', 'dep', 'arr', 'depblc', 'arrblc'].filter(function (f) { return !map[f]; });
    if (missing.length) throw new Error('Colonnes introuvables dans l\'Excel : ' + missing.join(', '));
    function g(r, f) { return map[f] ? r[map[f]] : ''; }

    var out = [];
    rows.forEach(function (r, i) {
      var date = toISODate(g(r, 'date'));
      if (!date) return; // ligne vide / total
      var warn = [];
      var dep = norm(g(r, 'dep')).toUpperCase(), arr = norm(g(r, 'arr')).toUpperCase();
      var o = toMinutes(g(r, 'depblc')), inn = toMinutes(g(r, 'arrblc'));
      var off = toMinutes(g(r, 'dec')), on = toMinutes(g(r, 'att'));
      var blk = durMinutes(g(r, 'tpsblc'));
      if (blk == null && o != null && inn != null) blk = (inn - o + 1440) % 1440;
      var flt = durMinutes(g(r, 'tpsvol'));
      if (flt == null && off != null && on != null) flt = (on - off + 1440) % 1440;
      var vol = flightNumber(g(r, 'vol'));
      var cdb = norm(g(r, 'cdb')).toUpperCase(), opl = norm(g(r, 'opl')).toUpperCase();
      var isSIC = trig && opl === trig && cdb !== trig;

      var A = airports[dep], B = airports[arr];
      var night = null, dist = '', nightLanding = false;
      if (A && B && o != null && blk) {
        night = nightMinutes(date, o, blk, A, B);
        dist = String(Math.round(gcNM(A, B)));
        var landT = on != null ? on : inn;
        var landMs = utcMs(date, landT < o ? landT + 1440 : landT);
        nightLanding = sunElev(B[0], B[1], landMs) < -6;
      } else {
        warn.push('Terrain inconnu (' + (!A ? dep : arr) + ') : nuit et distance non calculées');
      }

      var conso = parseFloat(String(g(r, 'conso')).replace(',', '.'));
      var ex = existing[date + '|' + vol];
      var d = {};
      COLS.forEach(function (c) { d[c] = ''; });
      d['Date'] = date;
      d['Registration'] = registration(g(r, 'imm'));
      d['Departure'] = dep; d['Arrival'] = arr;
      d['Out'] = o != null ? fmtHM(o) : ''; d['In'] = inn != null ? fmtHM(inn) : '';
      d['Off'] = off != null ? fmtHM(off) : ''; d['On'] = on != null ? fmtHM(on) : '';
      d['SchedOut'] = ex && ex.schedOut ? ex.schedOut : d['Out'];
      d['SchedIn'] = ex && ex.schedIn ? ex.schedIn : d['In'];
      d['Total Block'] = blk != null ? fmtHM(blk) : '';
      d['Total Flight'] = flt != null ? fmtHM(flt) : '';
      d['Night Time'] = night ? fmtHM(night) : '';
      d['Distance'] = dist;
      d['Pax'] = norm(g(r, 'pax')).replace(/\.0+$/, '');
      d['Fuel'] = isNaN(conso) ? '' : (conso * LBS_TO_T).toFixed(2);
      d['Flight Number'] = vol;
      if (isSIC) {
        d['SIC Time'] = d['Total Block'];
        d['PIC Name'] = cdb;
        d['Function'] = 'SIC';
      } else {
        d['PIC Time'] = d['Total Block'];
        d['PIC Name'] = norm(opts.picName) || 'Self';
        d['Function'] = 'PIC';
      }
      if (trig && cdb !== trig && opl !== trig) warn.push('Trigramme ' + trig + ' absent (CDB ' + cdb + ', OPL ' + opl + ')');

      var status = !ex ? 'nouveau' : (ex.planned ? 'remplace prévu' : 'doublon');
      out.push({ data: d, nightLanding: nightLanding, pf: true, status: status, warnings: warn });
    });
    return out;
  }

  // Applique le choix PF/PM sur décollages/atterrissages
  function finalize(item) {
    var d = Object.assign({}, item.data);
    if (item.pf) {
      d['Takeoff'] = '1';
      d['Day Landing'] = item.nightLanding ? '' : '1';
      d['Night Landing'] = item.nightLanding ? '1' : '';
    } else {
      d['Takeoff'] = '0'; d['Day Landing'] = ''; d['Night Landing'] = '';
    }
    return d;
  }

  function csvField(v) {
    v = String(v == null ? '' : v);
    return /[",\n\r]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v;
  }

  function toCSV(items) {
    var lines = [COLS.join(',')];
    items.forEach(function (it) {
      var d = finalize(it);
      lines.push(COLS.map(function (c) { return csvField(d[c]); }).join(','));
    });
    return lines.join('\n') + '\n';
  }

  var api = { COLS: COLS, convert: convert, toCSV: toCSV, finalize: finalize, parseCSV: parseCSV,
    indexExisting: indexExisting, sunElev: sunElev, toMinutes: toMinutes, toISODate: toISODate };
  if (typeof module !== 'undefined' && module.exports) module.exports = api; else root.FLConvert = api;
})(this);
