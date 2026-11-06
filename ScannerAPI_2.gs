/**
 * Scanner JSONP API for GitHub Pages.
 * Requires these existing Admin.gs functions:
 * adminLogin, adminCheckSession,
 * adminPreviewScannedTicket, adminConfirmCheckIn.
 */
function doGet(e) {
  const p = e && e.parameter ? e.parameter : {};
  const api = String(p.api || '').trim().toLowerCase();
  const action = String(p.action || '').trim().toLowerCase();

  // `api=scanner` is the preferred route. Accepting a known scanner
  // action as a fallback keeps older deployed frontends compatible.
  if (api !== 'scanner' && !['login', 'ping', 'preview', 'confirm'].includes(action)) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: 'API tidak dikenal.'
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return handleScannerApi_(e);
}

function handleScannerApi_(e) {
  const p = e && e.parameter ? e.parameter : {};
  let callback;

  try {
    callback = scannerCallback_(p.callback);
  } catch (error) {
    return ContentService
      .createTextOutput(JSON.stringify({
        success: false,
        message: error && error.message
          ? error.message
          : String(error || 'Callback JSONP tidak valid.')
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const action = String(p.action || '').trim().toLowerCase();
  let result;

  try {
    if (action === 'login') {
      const login = adminLogin(
        String(p.username || ''),
        String(p.password || ''),
        String(p.operatorName || p.gate || p.username || 'Scanner').trim()
      );

      result = {
        success: true,
        token: login.token,
        username: login.username || 'AdminTicketing',
        operatorName: login.operatorName || String(p.operatorName || p.gate || '').trim()
      };
    } else if (action === 'ping') {
      const session = adminCheckSession(String(p.token || ''));

      result = {
        success: true,
        username: session.username || 'AdminTicketing'
      };
    } else if (action === 'preview') {
      result = scannerResult_(
        adminPreviewScannedTicket(
          String(p.token || ''),
          String(p.qr || '')
        )
      );
    } else if (action === 'confirm') {
      const seatIds = scannerSeatIds_(p.seatIds);

      if (seatIds.length) {
        result = scannerConfirmMany_(
          String(p.token || ''),
          String(p.qr || ''),
          seatIds,
          String(p.gate || 'Main Gate')
        );
      } else {
        result = scannerResult_(
          adminConfirmCheckIn(
            String(p.token || ''),
            String(p.qr || ''),
            String(p.seatId || ''),
            String(p.gate || 'Main Gate')
          ) || {}
        );
      }
    } else {
      throw new Error('Scanner API action tidak dikenal.');
    }
  } catch (error) {
    result = {
      success: false,
      message: error && error.message
        ? error.message
        : String(error || 'Scanner API error.')
    };
  }

  return ContentService
    .createTextOutput(
      callback + '(' + JSON.stringify(result) + ');'
    )
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function scannerCallback_(value) {
  const callback = String(value || '').trim();

  if (!callback) {
    throw new Error('Parameter callback JSONP wajib diisi.');
  }

  if (!/^[A-Za-z_$][0-9A-Za-z_$\.]*$/.test(callback)) {
    throw new Error('Callback JSONP tidak valid.');
  }

  return callback;
}

function scannerResult_(value) {
  const result = value && typeof value === 'object' ? value : {};

  return Object.assign({}, result, {
    success: result.success === true
  });
}


/**
 * Menerima seatIds dari JSONP sebagai JSON array (preferred)
 * atau comma-separated string untuk backward compatibility.
 */
function scannerSeatIds_(value) {
  const raw = String(value || '').trim();
  if (!raw) return [];

  let values = [];

  try {
    const parsed = JSON.parse(raw);
    if (Array.isArray(parsed)) {
      values = parsed;
    }
  } catch (_) {
    values = raw.split(',');
  }

  const unique = [];
  const seen = new Set();

  values.forEach(function(item) {
    const seatId = String(item || '').trim();
    if (!seatId || seen.has(seatId)) return;
    seen.add(seatId);
    unique.push(seatId);
  });

  if (unique.length > 50) {
    throw new Error('Maksimal 50 seat dalam satu kali check-in.');
  }

  return unique;
}

/**
 * Multi-seat check-in untuk scanner GitHub Pages.
 * Tetap memakai adminConfirmCheckIn() per seat sehingga proteksi
 * anti-double-scan dan ScriptLock yang sudah ada tetap digunakan.
 * Jika ada konflik dari device lain, seat lain tetap diproses dan
 * response akan menandai hasil sebagai partial.
 */
function scannerConfirmMany_(token, qr, seatIds, gate) {
  const confirmed = [];
  const failed = [];

  seatIds.forEach(function(seatId) {
    try {
      const item = adminConfirmCheckIn(
        token,
        qr,
        seatId,
        gate
      ) || {};

      if (item.success) {
        confirmed.push(seatId);
      } else {
        failed.push({
          seatId: seatId,
          message: item.message || 'Check-in gagal.'
        });
      }
    } catch (error) {
      failed.push({
        seatId: seatId,
        message: error && error.message
          ? error.message
          : String(error || 'Check-in gagal.')
      });
    }
  });

  if (!confirmed.length) {
    return {
      success: false,
      seatIds: [],
      failed: failed,
      message: failed.length
        ? failed.map(function(item) {
            return item.seatId + ': ' + item.message;
          }).join(' | ')
        : 'Tidak ada seat yang berhasil check-in.'
    };
  }

  const preview = adminPreviewScannedTicket(token, qr) || {};
  const message = failed.length
    ? confirmed.length + ' seat berhasil check-in; ' + failed.length + ' seat gagal.'
    : confirmed.length + ' seat berhasil check-in.';

  return Object.assign({}, preview, {
    success: true,
    partial: failed.length > 0,
    seatIds: confirmed,
    confirmedCount: confirmed.length,
    failed: failed,
    message: message
  });
}
