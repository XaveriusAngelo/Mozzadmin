const CONFIG = {
  DATA_SHEETS: { 1: 'Show_1', 2: 'Show_2' },
  PRICE_SHEET: 'Prices',
  PROMO_SHEET: 'Promos',
  BOOKING_SHEET: 'Bookings',
  EVENT_NAME: 'Mozzafiato 2026 — Resonant Voyage',
  EMAIL_SENDER_NAME: 'Mozzafiato 2026',
ETICKET: {
    enabled: true,

    // Template Google Slides berbeda untuk setiap show.
    templates: {
      1: {
        presentationId: '1OZdlXKfc_4V-IOu_GppUd2cKel0Z_O_Ye63-MJF3N4A',
        slideIndex: 0
      },

      2: {
        presentationId: '14reqppsNgh75_UuQYsU7pIbSDh51zavp-KHsFk2F6pc',
        slideIndex: 0
      }
    },

    tempFolderName: 'Mozzafiato E-Ticket Temp'
  },

  // Waktu isi data diri setelah kursi berhasil diamankan.
  HOLD_MINUTES: 5,

  PAYMENT: {
    // Waktu pembayaran setelah data customer berhasil disubmit.
      paymentMinutes: 10,
    // Total DI ATAS angka ini wajib transfer bank.
    qrisMaxTotal: 500000,

    qris: {
      enabled: true,
      imageUrl:
        'https://drive.google.com/file/d/1ntzHI7vWolPoKgaEZgoO__grgAkggjhH/view?usp=drivesdk'
    },

    banks: [
      {
        id: 'seabank',
        enabled: true,
        bankName: 'Seabank',
        accountNumber: '901822322868',
        accountHolder: 'Rapolo Joshua Napitupulu'
      }
    ]
  },

  // Narahubung customer per show.
  SUPPORT_CONTACTS: {
    1: {
      name: 'Betty',
      whatsapp: '6281806114730'
    },
    2: {
      name: 'Azka',
      whatsapp: '6285196118284'
    }
  },

  // Estimasi waktu verifikasi pembayaran oleh panitia.
  VERIFICATION_HOURS: 24,

  COLORS: {
    Oceanic: '#5D9FCC',
    Coraline: '#B483C8',
    Lagoon: '#FF8A73',
    Held: '#F4C542',
    Booked: '#0A53A8',
    Sold: '#B10202',
    AvailableText: '#FFFFFF',
    UnavailableText: '#000000'
  }
};

const REQUIRED_COLUMNS = [
  'Booking_ID',
  'Hold_Token',
  'Held_At',
  'Hold_Until',
  'Locked_Price',
  'Promo_ID',
  'Promo_Name',
  'Previous_Status',
  'Updated_At'
];

const FREEPASS_PASSWORD_PROPERTY = 'FREEPASS_PASSWORD';


const STUDENT_VOYAGE_CONFIG = Object.freeze({
  minCohort: 2021,
  maxCohort: 2026,
  maxPromoTicketsPerKtm: 3,
  ktmFolderName: 'Mozzafiato Student Voyage KTM',
  maxFileBytes: 5 * 1024 * 1024
});

function doGet(e) {
  const params = e && e.parameter
    ? e.parameter
    : {};

  if (
    String(params.api || '')
      .trim()
      .toLowerCase() === 'scanner'
  ) {
    return handleScannerApi_(e);
  }

  const page = String(
    params.page || ''
  ).trim().toLowerCase();

  if (page === 'admin') {
    return HtmlService
      .createHtmlOutputFromFile('AdminUI')
      .setTitle('Mozzafiato 2026 — Admin Portal')
      .addMetaTag(
        'viewport',
        'width=device-width, initial-scale=1, minimum-scale=1, viewport-fit=cover'
      )
      .setXFrameOptionsMode(
        HtmlService.XFrameOptionsMode.ALLOWALL
      );
  }

  const show = normalizeShow_(params.show);
  const template =
    HtmlService.createTemplateFromFile('Index');

  template.initialShow = show;
  template.initialBookingId = String(
    params.booking ||
    params.bookingId ||
    ''
  ).trim();
return template
    .evaluate()
    .setTitle(
      'Mozzafiato 2026 — Seat Selection'
    )
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1, minimum-scale=1, viewport-fit=cover'
    )
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );
}

function getSeatMap(showValue, freepassPassword) {
  const show = normalizeShow_(showValue);
  const freepassUnlocked =
    isValidFreepassPassword_(freepassPassword);

  SpreadsheetApp.flush();
  releaseExpiredHolds_(show);

  const sheet = getDataSheet_(show);
  validateLayoutColumns_(sheet);

  const table = readTable_(sheet);
  const prices = getPrices_();
  const promos = getPromos_();
  const promoUsage = getPromoUsageMap_();
  const now = new Date();

  const publicPromos = buildPublicPromoInfo_(
    show,
    prices,
    promos,
    promoUsage,
    now
  );

  const promoByCategory = new Map(
    publicPromos.map(item => [item.category, item])
  );

  const seats = table.rows
    .map(row => {
      const category = String(row.Category || '').trim();
      const status = String(row.Status || 'Available').trim();
      const seatNumberRaw = row.Seat_Number;
      const seatNumber = String(
        seatNumberRaw === null || seatNumberRaw === undefined
          ? ''
          : seatNumberRaw
      ).trim();

      const regularPrice =
        getCategoryPrice_(prices, category);

      const activePromo =
        promoByCategory.get(category);

      const lockedPrice =
        toNumber_(row.Locked_Price);

      /*
       * Promo bundle tidak ditampilkan sebagai harga satu seat,
       * karena diskon baru sah setelah jumlah bundle terpenuhi.
       * Total pilihan tetap dihitung secara live di frontend.
       */
      const canShowPromoAsSeatPrice =
        activePromo &&
        activePromo.remaining > 0 &&
        String(
          activePromo.promoType || 'SINGLE'
        ).toUpperCase() !== 'BUNDLE';

      const displayPrice =
        status.toLowerCase() === 'available' ||
        (
          status.toLowerCase() === 'reserved' &&
          freepassUnlocked
        )
          ? (
              canShowPromoAsSeatPrice
                ? activePromo.price
                : regularPrice
            )
          : (
              lockedPrice !== null
                ? lockedPrice
                : regularPrice
            );

      return {
        seatId: String(row.Seat_ID || '').trim(),
        category,
        status,
        rowLabel: String(row.Row_Label || '').trim(),
        seatNumber: seatNumber || String(row.Seat_ID || '').trim(),

        x: toNumber_(row.X),
        y: toNumber_(row.Y),
        width: toNumber_(row.Width, 24),
        height: toNumber_(row.Height, 24),

        price: displayPrice,
        regularPrice,
        promoId:
          activePromo && activePromo.remaining > 0
            ? activePromo.promoId
            : '',
        promoName:
          activePromo && activePromo.remaining > 0
            ? activePromo.promoName
            : '',
        promoRemaining:
          activePromo
            ? activePromo.remaining
            : 0,

        holdUntil: toIsoDateString_(row.Hold_Until),

        isSelectable:
          status.toLowerCase() === 'available' ||
          (
            status.toLowerCase() === 'reserved' &&
            freepassUnlocked
          ),
        isFreepass:
          status.toLowerCase() === 'reserved'
      };
    })
    .filter(seat => seat.seatId);

  return {
    success: true,
    show,
    freepassUnlocked,
    serverTime: now.toISOString(),
    holdMinutes: CONFIG.HOLD_MINUTES,
    paymentMinutes: getPaymentMinutes_(),
    contact: getShowContact_(show),
    verificationNotice: buildVerificationNotice_(show),
    prices,
    promos: publicPromos,
    nextPromoChangeAt:
      getNextPromoChangeAt_(
        promos,
        show,
        now
      ),
    seats
  };
}

function holdSeats(request) {
  if (
    !request ||
    !Array.isArray(request.seatIds) ||
    !request.seatIds.length
  ) {
    throw new Error('Pilih minimal satu kursi.');
  }

  const show = normalizeShow_(request.show);

  const requestedIds = [
    ...new Set(request.seatIds.map(id => String(id).trim()))
  ].filter(Boolean);

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    releaseExpiredHolds_(show);

    const sheet = getDataSheet_(show);
    ensureOperationalColumns_(sheet);

    const table = readTable_(sheet);
    const prices = getPrices_();
    const promos = getPromos_();
    const promoUsage = getPromoUsageMap_();
    const pricingNow = new Date();

    const bySeat = new Map(
      table.rows.map((row, index) => [
        String(row.Seat_ID || '').trim(),
        {
          row,
          sheetRow: index + 2
        }
      ])
    );

    const unavailable = [];

    requestedIds.forEach(id => {
      const item = bySeat.get(id);

      if (!item) {
        unavailable.push(id);
        return;
      }

      const itemStatus =
        String(item.row.Status || '').trim().toLowerCase();

      if (
        itemStatus === 'reserved' &&
        !isValidFreepassPassword_(request.freepassPassword)
      ) {
        unavailable.push(id);
        return;
      }

      if (!isSelectableSeatStatus_(item.row.Status)) {
        unavailable.push(id);
      }
    });

    if (unavailable.length) {
      throw new Error(
        `Kursi ini baru saja tidak tersedia: ${unavailable.join(', ')}`
      );
    }

    /*
     * Harga dikunci saat hold dibuat.
     * Promo memakai waktu server + kuota real-time.
     *
     * Bundle dihitung dari seluruh pilihan customer,
     * bukan dari satu seat secara terpisah.
     */
    const lockedPricing =
      buildLockedPricingForSelection_(
        show,
        requestedIds,
        bySeat,
        prices,
        promos,
        promoUsage,
        pricingNow
      );

    const bookingId = createBookingId_(show);
    const holdToken = Utilities.getUuid();
    const heldAt = new Date();

    const holdUntil = new Date(
      heldAt.getTime() +
      CONFIG.HOLD_MINUTES * 60 * 1000
    );

    const headers = getHeaders_(sheet);

    requestedIds.forEach(id => {
      const item = bySeat.get(id);

      setCellByHeader_(
        sheet,
        item.sheetRow,
        headers,
        'Previous_Status',
        String(item.row.Status || '').trim()
      );

      setCellByHeader_(
        sheet,
        item.sheetRow,
        headers,
        'Status',
        'Held'
      );

      setCellByHeader_(
        sheet,
        item.sheetRow,
        headers,
        'Booking_ID',
        bookingId
      );

      setCellByHeader_(
        sheet,
        item.sheetRow,
        headers,
        'Hold_Token',
        holdToken
      );

      setCellByHeader_(
        sheet,
        item.sheetRow,
        headers,
        'Held_At',
        heldAt.toISOString()
      );

      setCellByHeader_(
        sheet,
        item.sheetRow,
        headers,
        'Hold_Until',
        holdUntil.toISOString()
      );

      const pricing = lockedPricing.get(id);

      setCellByHeader_(
        sheet,
        item.sheetRow,
        headers,
        'Locked_Price',
        pricing.price
      );

      setCellByHeader_(
        sheet,
        item.sheetRow,
        headers,
        'Promo_ID',
        pricing.promoId || ''
      );

      setCellByHeader_(
        sheet,
        item.sheetRow,
        headers,
        'Promo_Name',
        pricing.promoName || ''
      );

      setCellByHeader_(
        sheet,
        item.sheetRow,
        headers,
        'Updated_At',
        heldAt
      );
    });

    syncVisualSheet_(show);
    SpreadsheetApp.flush();

    const selected = requestedIds.map(id => {
      const row = bySeat.get(id).row;
      const pricing = lockedPricing.get(id);

      return {
        ...row,
        Locked_Price: pricing.price,
        Promo_ID: pricing.promoId || '',
        Promo_Name: pricing.promoName || ''
      };
    });

    const total = requestedIds.reduce(
      (sum, id) =>
        sum + Number(lockedPricing.get(id).price || 0),
      0
    );

    const categorySummary = summarizeCategories_(selected);
    const promoSummary = summarizePromos_(selected);
    const regularTotal = selected.reduce(
      (sum, row) =>
        sum + getCategoryPrice_(prices, String(row.Category || '').trim()),
      0
    );
    const studentVoyageActive = selected.some(isStudentVoyagePromoRow_);

    return {
      bookingId,
      holdToken,
      holdUntil: holdUntil.toISOString(),
      show,
      seatIds: requestedIds,
      total,
      regularTotal,
      categorySummary,
      promoSummary,
      studentVoyageActive
    };
  } finally {
    lock.releaseLock();
  }
}

function cancelSeatHold(request) {
  if (!request) {
    throw new Error('Data hold tidak ditemukan.');
  }

  const bookingId = String(request.bookingId || '').trim();
  const holdToken = String(request.holdToken || '').trim();
  const show = normalizeShow_(request.show);

  if (!bookingId || !holdToken) {
    return {
      success: true,
      released: 0
    };
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = getDataSheet_(show);
    ensureOperationalColumns_(sheet);

    const table = readTable_(sheet);
    const headers = getHeaders_(sheet);
    const now = new Date();
    let released = 0;

    table.rows.forEach((row, index) => {
      const status = String(row.Status || '').trim().toLowerCase();
      const rowBookingId = String(row.Booking_ID || '').trim();
      const rowHoldToken = String(row.Hold_Token || '').trim();

      // Hanya kursi yang masih berstatus Held yang boleh dilepas.
      // Kursi Booked/Sold tidak akan tersentuh.
      if (
        status === 'held' &&
        rowBookingId === bookingId &&
        rowHoldToken === holdToken
      ) {
        const sheetRow = index + 2;

        setCellByHeader_(
          sheet,
          sheetRow,
          headers,
          'Status',
          restoreSeatStatus_(row)
        );

        [
          'Booking_ID',
          'Hold_Token',
          'Held_At',
          'Hold_Until',
          'Locked_Price',
          'Promo_ID',
          'Promo_Name',
          'Previous_Status'
        ].forEach(header => {
          setCellByHeader_(sheet, sheetRow, headers, header, '');
        });

        setCellByHeader_(
          sheet,
          sheetRow,
          headers,
          'Updated_At',
          now
        );

        released += 1;
      }
    });

    SpreadsheetApp.flush();

    return {
      success: true,
      released
    };
  } finally {
    lock.releaseLock();
  }
}


function submitBooking(request) {
  if (!request) {
    throw new Error('Data booking tidak ditemukan.');
  }

  const bookingId =
    String(request.bookingId || '').trim();

  const holdToken =
    String(request.holdToken || '').trim();

  const name =
    String(request.name || '').trim();

  const email =
    String(request.email || '').trim();

  const whatsapp =
    String(request.whatsapp || '').trim();

  const referral =
    String(request.referral || '').trim();

  const requestedCustomerType =
    String(request.customerType || '').trim().toLowerCase();

  const studentUniversity =
    String(request.studentUniversity || '').trim();

  const studentCohortRaw =
    String(request.studentCohort || '').trim();

  if (!bookingId || !holdToken) {
    throw new Error(
      'Booking ID atau hold token tidak valid.'
    );
  }

  if (!name) {
    throw new Error(
      'Nama lengkap wajib diisi.'
    );
  }

  if (!email) {
    throw new Error(
      'Email wajib diisi.'
    );
  }

  if (!isValidEmail_(email)) {
    throw new Error(
      'Format email tidak valid.'
    );
  }

  if (!whatsapp) {
    throw new Error(
      'Nomor WhatsApp wajib diisi.'
    );
  }

  if (!referral) {
    throw new Error(
      'Referral Penyanyi wajib diisi. Jika tidak ada, isi dengan tanda (-).'
    );
  }

  const lock =
    LockService.getScriptLock();

  lock.waitLock(15000);

  try {
    let holdData =
      findHeldBooking_(
        bookingId,
        holdToken
      );

    if (!holdData.rows.length) {
      throw new Error(
        'Data hold tidak ditemukan atau ' +
        'masa hold sudah berakhir.'
      );
    }

    const now = new Date();

    if (
      !holdData.holdUntil ||
      holdData.holdUntil.getTime() <=
        now.getTime()
    ) {
      throw new Error(
        'Masa hold telah berakhir. ' +
        'Silakan pilih kursi kembali.'
      );
    }

    const bookingSheet =
      getBookingSheet_();

    ensureBookingHeaders_(
      bookingSheet
    );

    if (
      bookingExists_(
        bookingSheet,
        bookingId
      )
    ) {
      throw new Error(
        'Booking ini sudah pernah disubmit.'
      );
    }

    let studentPromoRows =
      holdData.rows.filter(isStudentVoyagePromoRow_);

    const studentVoyageActive =
      studentPromoRows.length > 0;

    const freepassBooking =
      holdData.rows.length > 0 &&
      holdData.rows.every(row =>
        String(row.Previous_Status || '').trim().toLowerCase() ===
        'reserved'
      );

    let customerType = 'General';
    let resolvedStudentUniversity = '';
    let resolvedStudentCohort = '';
    let studentKtmUrl = '';
    let studentKtmHash = '';
    let studentPromoTickets = 0;
    let preparedStudentKtm = null;

    if (studentVoyageActive) {
      if (!['general', 'student'].includes(requestedCustomerType)) {
        throw new Error('Pilih status penonton terlebih dahulu.');
      }

      if (requestedCustomerType === 'general') {
        repriceHeldStudentVoyageToRegular_(
          holdData.show,
          bookingId,
          holdToken,
          holdData.prices,
          now
        );

        holdData = findHeldBooking_(bookingId, holdToken);

        if (!holdData.rows.length) {
          throw new Error(
            'Data hold tidak ditemukan setelah penyesuaian harga. Silakan pilih kursi kembali.'
          );
        }
      } else {
        customerType = 'Student';

        if (!studentUniversity) {
          throw new Error('Asal universitas wajib diisi.');
        }

        if (!/^\d{4}$/.test(studentCohortRaw)) {
          throw new Error('Angkatan harus berupa tahun 4 digit.');
        }

        const studentCohort = Number(studentCohortRaw);

        if (
          studentCohort < STUDENT_VOYAGE_CONFIG.minCohort ||
          studentCohort > STUDENT_VOYAGE_CONFIG.maxCohort
        ) {
          throw new Error(
            'Mohon maaf, kamu belum bisa mendapatkan promo Student Voyage 🙏'
          );
        }

        studentPromoTickets = studentPromoRows.length;

        if (
          studentPromoTickets >
          STUDENT_VOYAGE_CONFIG.maxPromoTicketsPerKtm
        ) {
          throw new Error(
            'Satu KTM hanya dapat digunakan untuk maksimal 3 tiket promo Student Voyage 🙏'
          );
        }

        preparedStudentKtm =
          prepareStudentKtmUpload_(request.studentKtm);

        const existingStudentPromoTickets =
          countStudentVoyageTicketsByKtmHash_(
            bookingSheet,
            preparedStudentKtm.hash
          );

        if (
          existingStudentPromoTickets + studentPromoTickets >
          STUDENT_VOYAGE_CONFIG.maxPromoTicketsPerKtm
        ) {
          throw new Error(
            'KTM ini sudah mencapai batas maksimal 3 tiket promo Student Voyage 🙏'
          );
        }

        resolvedStudentUniversity = studentUniversity;
        resolvedStudentCohort = studentCohort;
        studentKtmHash = preparedStudentKtm.hash;
      }
    }

    /*
     * Batas pembayaran dihitung setelah
     * data customer berhasil disubmit.
     */
    const paymentMinutes =
      getPaymentMinutes_();

    const paymentDeadline = new Date(
      now.getTime() +
      paymentMinutes * 60 * 1000
    );

    const seatIds =
      holdData.rows.map(row =>
        String(
          row.Seat_ID || ''
        ).trim()
      );

    const total =
      holdData.rows.reduce(
        (sum, row) =>
          sum +
          getLockedPriceForRow_(
            row,
            holdData.prices
          ),
        0
      );

    const categorySummary =
      summarizeCategories_(
        holdData.rows
      );

    const promoSummary =
      summarizePromos_(
        holdData.rows
      );

    if (preparedStudentKtm) {
      studentKtmUrl = saveStudentKtmFile_(
        bookingId,
        name,
        preparedStudentKtm
      );
    }

    const bookingHeaders =
      getHeaders_(bookingSheet);

    const bookingRow =
      bookingSheet.getLastRow() + 1;

    const bookingData = {
      Booking_ID: bookingId,
      Hold_Token: holdToken,
      Show: holdData.show,
      Seats: seatIds.join(', '),
      Category_Summary: categorySummary,
      Promo_Summary: promoSummary,
      Total: total,
      Name: name,
      Email: email,
      WhatsApp: whatsapp,
      Referral_Penyanyi: referral,

      Customer_Type: customerType,
      Student_University: resolvedStudentUniversity,
      Student_Cohort: resolvedStudentCohort,
      Student_KTM_URL: studentKtmUrl,
      Student_KTM_Hash: studentKtmHash,
      Student_Promo_Tickets: studentPromoTickets,

      Status: freepassBooking ? 'Verified' : 'Waiting for Payment',
      Payment_Status: freepassBooking ? 'Verified' : 'Unpaid',
      Payment_Method: freepassBooking ? 'Freepass' : '',
      Payment_Deadline: freepassBooking
        ? ''
        : paymentDeadline.toISOString(),
      Payment_Proof_URL: '',

      Created_At: now,
      Updated_At: now
    };

    /*
     * Menulis data berdasarkan nama header,
     * sehingga urutan kolom tetap aman.
     */
    bookingHeaders.forEach(
      (header, index) => {
        const value =
          Object.prototype
            .hasOwnProperty.call(
              bookingData,
              header
            )
            ? bookingData[header]
            : '';

        bookingSheet
          .getRange(
            bookingRow,
            index + 1
          )
          .setValue(value);
      }
    );

    /*
     * Kursi diubah menjadi Booked supaya
     * tidak bisa dipilih pembeli lain.
     *
     * Status transaksi di sheet Bookings
     * tetap Waiting for Payment.
     */
    updateHeldSeatsToBooked_(
      holdData.show,
      bookingId,
      holdToken,
      now
    );

    if (freepassBooking) {
      finalizeVerifiedBooking_(
        bookingId,
        bookingData,
        bookingRow,
        bookingHeaders,
        now
      );
    }

    syncVisualSheet_(
      holdData.show
    );
    SpreadsheetApp.flush();

    return {
      success: true,
      bookingId,
      show: holdData.show,
      seatIds,
      total,
      categorySummary,
      promoSummary,
      customerType,
      studentPromoTickets,

      status: freepassBooking ? 'Verified' : 'Waiting for Payment',
      paymentStatus: freepassBooking ? 'Verified' : 'Unpaid',
      freepass: freepassBooking,

      paymentDeadline:
        freepassBooking
          ? ''
          : paymentDeadline.toISOString(),

      contact: getShowContact_(holdData.show),
      verificationNotice:
        buildVerificationNotice_(holdData.show)
    };
  } finally {
    lock.releaseLock();
  }
}

function getPublicAppHealth() {
  const result = {
    ok: true,
    spreadsheetId: SpreadsheetApp.getActive().getId(),
    sheets: {}
  };

  [1, 2].forEach(show => {
    const sheet = getDataSheet_(show);
    const headers = getHeaders_(sheet);

    result.sheets[`show${show}`] = {
      name: sheet.getName(),
      rows: Math.max(0, sheet.getLastRow() - 1),
      hasX: headers.includes('X'),
      hasY: headers.includes('Y'),
      hasSeatId: headers.includes('Seat_ID'),
      hasStatus: headers.includes('Status')
    };
  });

  result.prices = getPrices_();
  result.promos = getPromos_().map(promo => ({
    promoId: promo.promoId,
    promoName: promo.promoName,
    category: promo.category,
    quota: promo.quota,
    active: promo.active,
    show: promo.show,
    startAt: promo.startAt.toISOString(),
    endAt: promo.endAt.toISOString()
  }));
  return result;
}


function releaseExpiredHolds() {
  Object.keys(CONFIG.DATA_SHEETS).forEach(show => releaseExpiredHolds_(Number(show)));
}

function releaseExpiredHolds_(show) {
  const sheet = getDataSheet_(show);
  ensureOperationalColumns_(sheet);
  const table = readTable_(sheet);
  const headers = getHeaders_(sheet);
  const now = new Date();
  let changed = false;

  table.rows.forEach((row, i) => {
    const status = String(row.Status || '').toLowerCase();
    const until = parseStoredDate_(row.Hold_Until);
    if (status === 'held' && until && until.getTime() <= now.getTime()) {
      const sheetRow = i + 2;
      setCellByHeader_(
        sheet,
        sheetRow,
        headers,
        'Status',
        restoreSeatStatus_(row)
      );
      [
        'Booking_ID',
        'Hold_Token',
        'Held_At',
        'Hold_Until',
        'Locked_Price',
        'Promo_ID',
        'Promo_Name',
        'Previous_Status'
      ].forEach(h =>
        setCellByHeader_(sheet, sheetRow, headers, h, '')
      );
      setCellByHeader_(sheet, sheetRow, headers, 'Updated_At', now);
      changed = true;
    }
  });

  if (changed) syncVisualSheet_(show);
}

function syncVisualSheet_(show) {
  // Tidak digunakan.
  // Seat plan pada website dirender langsung dari kolom X dan Y
  // di sheet Show_1 / Show_2.
  return;
}

function installTriggers() {
  const handlerNames = [
    'releaseExpiredHolds',
    'releaseExpiredPayments',
    'runBookingMaintenance'
  ];

  ScriptApp.getProjectTriggers()
    .filter(trigger =>
      handlerNames.includes(
        trigger.getHandlerFunction()
      )
    )
    .forEach(trigger =>
      ScriptApp.deleteTrigger(trigger)
    );

  ScriptApp.newTrigger('runBookingMaintenance')
    .timeBased()
    .everyMinutes(1)
    .create();
}


/**
 * Menjalankan pembersihan hold dan pembayaran kedaluwarsa.
 * Fungsi ini dipanggil otomatis oleh trigger setiap 1 menit.
 */
function runBookingMaintenance() {
  releaseExpiredHolds();
  releaseExpiredPayments();
}


/**
 * Mengubah booking yang belum dibayar setelah deadline menjadi Expired,
 * lalu mengembalikan kursinya menjadi Available.
 */
function releaseExpiredPayments() {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const bookingSheet = getBookingSheet_();
    ensureBookingHeaders_(bookingSheet);

    const table = readBookingTable_(bookingSheet);
    const headers = getHeaders_(bookingSheet);
    const now = new Date();

    let bookingChanged = false;
    const affectedShows = new Set();

    table.rows.forEach((booking, index) => {
      const status = String(
        booking.Status || ''
      ).trim().toLowerCase();

      const paymentStatus = String(
        booking.Payment_Status || ''
      ).trim().toLowerCase();

      const deadline =
        parseStoredDate_(booking.Payment_Deadline);

      const isWaitingForPayment =
        status === 'waiting for payment' ||
        paymentStatus === 'unpaid';

      const alreadyProcessed =
        [
          'waiting for verification',
          'verified',
          'paid',
          'expired',
          'cancelled',
          'canceled'
        ].includes(paymentStatus) ||
        [
          'waiting for verification',
          'verified',
          'expired',
          'cancelled',
          'canceled'
        ].includes(status);

      if (
        !isWaitingForPayment ||
        alreadyProcessed ||
        !deadline ||
        deadline.getTime() > now.getTime()
      ) {
        return;
      }

      const bookingId = String(
        booking.Booking_ID || ''
      ).trim();

      const show = Number(
        booking.Show || 0
      );

      if (!bookingId || !CONFIG.DATA_SHEETS[show]) {
        return;
      }

      releaseBookedSeatsForExpiredPayment_(
        show,
        bookingId,
        now
      );

      const sheetRow = index + 2;

      setCellByHeader_(
        bookingSheet,
        sheetRow,
        headers,
        'Status',
        'Expired'
      );

      setCellByHeader_(
        bookingSheet,
        sheetRow,
        headers,
        'Payment_Status',
        'Expired'
      );

      setCellByHeader_(
        bookingSheet,
        sheetRow,
        headers,
        'Updated_At',
        now
      );

      bookingChanged = true;
      affectedShows.add(show);
    });

    if (bookingChanged) {
      SpreadsheetApp.flush();

      affectedShows.forEach(show =>
        syncVisualSheet_(show)
      );
    }
  } finally {
    lock.releaseLock();
  }
}


/**
 * Mengembalikan kursi milik booking kedaluwarsa menjadi Available.
 */
function releaseBookedSeatsForExpiredPayment_(
  show,
  bookingId,
  updatedAt
) {
  const sheet = getDataSheet_(show);
  ensureOperationalColumns_(sheet);

  const table = readTable_(sheet);
  const headers = getHeaders_(sheet);

  table.rows.forEach((row, index) => {
    const matchesBooking =
      String(row.Booking_ID || '').trim() ===
      bookingId;

    const status =
      String(row.Status || '')
        .trim()
        .toLowerCase();

    if (
      !matchesBooking ||
      status !== 'booked'
    ) {
      return;
    }

    const sheetRow = index + 2;

    setCellByHeader_(
      sheet,
      sheetRow,
      headers,
      'Status',
      restoreSeatStatus_(row)
    );

    [
      'Booking_ID',
      'Hold_Token',
      'Held_At',
      'Hold_Until',
      'Locked_Price',
      'Promo_ID',
      'Promo_Name',
      'Previous_Status'
    ].forEach(header => {
      setCellByHeader_(
        sheet,
        sheetRow,
        headers,
        header,
        ''
      );
    });

    setCellByHeader_(
      sheet,
      sheetRow,
      headers,
      'Updated_At',
      updatedAt
    );
  });
}


/**
 * Memastikan dataset memiliki kolom koordinat yang dibutuhkan Index.html.
 */
function validateLayoutColumns_(sheet) {
  const headers = getHeaders_(sheet);
  const required = ['X', 'Y'];
  const missing = required.filter(header => !headers.includes(header));

  if (missing.length) {
    throw new Error(
      `Kolom koordinat ${missing.join(', ')} belum ada di sheet ${sheet.getName()}.`
    );
  }
}

function ensureOperationalColumns_(sheet) {
  const headers = getHeaders_(sheet);
  const missing = REQUIRED_COLUMNS.filter(h => !headers.includes(h));
  if (!missing.length) return;
  const start = headers.length + 1;
  sheet.getRange(1, start, 1, missing.length).setValues([missing]);
}

function readTable_(sheet) {
  ensureOperationalColumns_(sheet);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const rows = values.slice(1).map(row => Object.fromEntries(headers.map((h, i) => [h, row[i]])));
  return { headers, rows };
}

function getHeaders_(sheet) {
  return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
}

function setCellByHeader_(sheet, row, headers, header, value) {
  const col = headers.indexOf(header) + 1;
  if (!col) throw new Error(`Kolom ${header} tidak ditemukan.`);
  sheet.getRange(row, col).setValue(value);
}

function getDataSheet_(show) {
  const name = CONFIG.DATA_SHEETS[show];
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) throw new Error(`Sheet ${name} tidak ditemukan.`);
  return sheet;
}

function normalizeShow_(value) {
  const show = Number(value || 1);
  if (![1, 2].includes(show)) throw new Error('Show tidak valid.');
  return show;
}

function isSelectableSeatStatus_(status) {
  return ['available', 'reserved'].includes(
    String(status || '').trim().toLowerCase()
  );
}

function isValidFreepassPassword_(password) {
  const configuredPassword =
    PropertiesService
      .getScriptProperties()
      .getProperty(FREEPASS_PASSWORD_PROPERTY);

  return Boolean(
    configuredPassword &&
    String(password || '') === configuredPassword
  );
}

function restoreSeatStatus_(row) {
  const previousStatus = String(row.Previous_Status || '').trim();
  return previousStatus.toLowerCase() === 'reserved'
    ? 'Reserved'
    : 'Available';
}

function createBookingId_(show) {
  const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMddHHmmss');
  const random = Math.random().toString(36).slice(2, 6).toUpperCase();
  return `MZF-S${show}-${stamp}-${random}`;
}

function summarizeCategories_(rows) {
  const counts = {};
  rows.forEach(row => {
    const category = String(row.Category).trim();
    counts[category] = (counts[category] || 0) + 1;
  });
  return Object.entries(counts).map(([name, count]) => `${count} ${name}`).join(', ');
}

function buildPrefilledFormUrl_(data) {
  const cfg = CONFIG.FORM;

  if (
    !cfg ||
    !cfg.url ||
    String(cfg.url).includes('PASTE_') ||
    !cfg.entries
  ) {
    return '';
  }

  const params = {};
  params[cfg.entries.bookingId] = data.bookingId;
  params[cfg.entries.show] = `Show ${data.show}`;
  params[cfg.entries.seats] = data.seatIds.join(', ');
  params[cfg.entries.categorySummary] = data.categorySummary;
  params[cfg.entries.total] = String(data.total);

  const query = Object.entries(params)
    .filter(([key]) => key && !String(key).includes('XXXX'))
    .map(([key, value]) =>
      `${encodeURIComponent(key)}=${encodeURIComponent(value)}`
    )
    .join('&');

  return query
    ? `${cfg.url}${cfg.url.includes('?') ? '&' : '?'}${query}`
    : cfg.url;
}

/**
 * Mengubah nilai spreadsheet menjadi angka yang aman.
 * Nilai kosong menghasilkan fallback.
 */
function toNumber_(value, fallback) {
  const defaultValue = fallback === undefined ? null : fallback;

  if (value === '' || value === null || value === undefined) {
    return defaultValue;
  }

  const number = Number(value);
  return Number.isFinite(number) ? number : defaultValue;
}

/**
 * Membaca deadline/timestamp dari Sheets secara tahan format.
 * Mendukung Date object, ISO string, epoch milliseconds, dan serial date Sheets.
 * Timestamp baru disimpan sebagai ISO text agar timezone/format cell tidak menggeser waktu.
 */
function parseStoredDate_(value) {
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    // Epoch milliseconds.
    if (value > 100000000000) {
      const epochDate = new Date(value);
      return Number.isNaN(epochDate.getTime()) ? null : epochDate;
    }

    // Google Sheets serial date (days since 1899-12-30).
    if (value > 10000 && value < 100000) {
      const serialDate = new Date((value - 25569) * 86400000);
      return Number.isNaN(serialDate.getTime()) ? null : serialDate;
    }
  }

  const text = String(value || '').trim();
  if (!text) return null;

  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function toIsoDateString_(value) {
  const date = parseStoredDate_(value);
  return date ? date.toISOString() : '';
}

function titleCase_(text) {
  const s = String(text || '').toLowerCase();
  return s ? s.charAt(0).toUpperCase() + s.slice(1) : '';
}

/**
 * Membaca harga berdasarkan kategori dari sheet Prices.
 *
 * Struktur wajib:
 * Category | Price
 */
function getPrices_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(CONFIG.PRICE_SHEET);

  if (!sheet) {
    throw new Error(
      `Sheet harga "${CONFIG.PRICE_SHEET}" tidak ditemukan.`
    );
  }

  const values = sheet.getDataRange().getValues();

  if (values.length < 2) {
    throw new Error(
      `Sheet "${CONFIG.PRICE_SHEET}" belum memiliki data harga.`
    );
  }

  const headers = values[0].map(header =>
    String(header || '').trim()
  );

  const categoryIndex = headers.indexOf('Category');
  const priceIndex = headers.indexOf('Price');

  if (categoryIndex === -1 || priceIndex === -1) {
    throw new Error(
      `Sheet "${CONFIG.PRICE_SHEET}" harus memiliki kolom Category dan Price.`
    );
  }

  const prices = {};

  values.slice(1).forEach((row, index) => {
    const category = String(row[categoryIndex] || '').trim();
    const price = toNumber_(row[priceIndex]);

    if (!category) {
      return;
    }

    if (price === null || price < 0) {
      throw new Error(
        `Harga untuk kategori "${category}" tidak valid pada baris ${index + 2}.`
      );
    }

    if (Object.prototype.hasOwnProperty.call(prices, category)) {
      throw new Error(
        `Kategori "${category}" muncul lebih dari satu kali di sheet Prices.`
      );
    }

    prices[category] = price;
  });

  if (!Object.keys(prices).length) {
    throw new Error('Tidak ada harga yang terbaca dari sheet Prices.');
  }

  return prices;
}


/**
 * Mengambil harga kategori dan memberikan error jika belum diatur.
 */
function getCategoryPrice_(prices, category) {
  if (!Object.prototype.hasOwnProperty.call(prices, category)) {
    throw new Error(
      `Harga kategori "${category}" belum diatur di sheet "${CONFIG.PRICE_SHEET}".`
    );
  }

  return prices[category];
}


function getPaymentMinutes_() {
  return Math.max(
    1,
    Number(
      CONFIG.PAYMENT &&
      CONFIG.PAYMENT.paymentMinutes
        ? CONFIG.PAYMENT.paymentMinutes
        : 10
    )
  );
}


/**
 * Promo sheet:
 * Promo_ID | Promo_Name | Category | Promo_Price |
 * Start_At | End_At | Quota | Active | Show
 *
 * Show boleh kosong / ALL / 1 / 2.
 */
function getPromos_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(CONFIG.PROMO_SHEET);

  // Promo bersifat opsional. Jika sheet belum dibuat,
  // harga reguler tetap berjalan.
  if (!sheet || sheet.getLastRow() < 2) {
    return [];
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(value =>
    String(value || '').trim()
  );

  const required = [
    'Promo_ID',
    'Promo_Name',
    'Category',
    'Promo_Price',
    'Start_At',
    'End_At',
    'Quota',
    'Active'
  ];

  const missing = required.filter(
    header => !headers.includes(header)
  );

  if (missing.length) {
    throw new Error(
      `Sheet "${CONFIG.PROMO_SHEET}" belum lengkap. ` +
      `Kolom yang kurang: ${missing.join(', ')}.`
    );
  }

  const index = Object.fromEntries(
    headers.map((header, i) => [header, i])
  );

  const seenIds = new Set();
  const promos = [];

  values.slice(1).forEach((row, offset) => {
    const sheetRow = offset + 2;
    const promoId =
      String(row[index.Promo_ID] || '').trim();

    if (!promoId) return;

    if (seenIds.has(promoId)) {
      throw new Error(
        `Promo_ID "${promoId}" duplikat di sheet Promos.`
      );
    }
    seenIds.add(promoId);

    const category =
      String(row[index.Category] || '').trim();

    const promoNameRaw =
      String(row[index.Promo_Name] || '').trim() ||
      promoId;

    const promoNameKey =
      promoNameRaw.toLowerCase();

    const promoName =
      (
        promoNameKey.includes('student voyage') ||
        promoNameKey.includes('student crew') ||
        promoNameKey.includes('mahasiswa')
      )
        ? 'Student Voyage'
        : promoNameRaw;

    const promoPrice =
      toNumber_(row[index.Promo_Price]);

    const quota =
      Math.floor(Number(row[index.Quota] || 0));

    const startAt = row[index.Start_At];
    const endAt = row[index.End_At];

    const activeValue = row[index.Active];
    const active =
      activeValue === true ||
      ['true', 'yes', 'y', '1', 'aktif', 'active']
        .includes(
          String(activeValue || '')
            .trim()
            .toLowerCase()
        );

    const showValue =
      index.Show === undefined
        ? ''
        : String(row[index.Show] || '').trim();

    const promoTypeRaw =
      index.Promo_Type === undefined
        ? 'SINGLE'
        : String(row[index.Promo_Type] || 'SINGLE')
            .trim()
            .toUpperCase();

    const promoType =
      promoTypeRaw || 'SINGLE';

    const bundleQtyRaw =
      index.Bundle_Qty === undefined
        ? 1
        : Math.floor(Number(row[index.Bundle_Qty] || 1));

    const bundleQty =
      promoType === 'BUNDLE'
        ? bundleQtyRaw
        : 1;

    const bundleTotalRaw =
      index.Bundle_Total === undefined
        ? null
        : toNumber_(row[index.Bundle_Total]);

    if (!category) {
      throw new Error(
        `Category promo kosong pada baris ${sheetRow}.`
      );
    }

    if (
      promoPrice === null ||
      promoPrice < 0
    ) {
      throw new Error(
        `Promo_Price tidak valid pada baris ${sheetRow}.`
      );
    }

    if (
      !Number.isFinite(quota) ||
      quota < 0
    ) {
      throw new Error(
        `Quota tidak valid pada baris ${sheetRow}.`
      );
    }

    if (
      !(startAt instanceof Date) ||
      isNaN(startAt.getTime()) ||
      !(endAt instanceof Date) ||
      isNaN(endAt.getTime())
    ) {
      throw new Error(
        `Start_At / End_At promo pada baris ${sheetRow} ` +
        'harus berupa tanggal dan waktu Google Sheets.'
      );
    }

    if (endAt.getTime() <= startAt.getTime()) {
      throw new Error(
        `End_At harus setelah Start_At pada baris ${sheetRow}.`
      );
    }

    if (!['SINGLE', 'BUNDLE'].includes(promoType)) {
      throw new Error(
        `Promo_Type pada baris ${sheetRow} harus SINGLE atau BUNDLE.`
      );
    }

    if (
      promoType === 'BUNDLE' &&
      (
        !Number.isFinite(bundleQty) ||
        bundleQty < 2
      )
    ) {
      throw new Error(
        `Bundle_Qty promo bundle pada baris ${sheetRow} minimal 2.`
      );
    }

    const bundleTotal =
      promoType === 'BUNDLE'
        ? (
            bundleTotalRaw !== null &&
            bundleTotalRaw >= 0
              ? bundleTotalRaw
              : promoPrice * bundleQty
          )
        : null;

    promos.push({
      promoId,
      promoName,
      category,
      promoPrice,
      startAt,
      endAt,
      quota,
      active,
      show: showValue,
      promoType,
      bundleQty,
      bundleTotal
    });
  });

  return promos;
}


function promoAppliesToShow_(promo, show) {
  const value =
    String(promo.show || '')
      .trim()
      .toLowerCase();

  return (
    !value ||
    value === 'all' ||
    value === '*' ||
    Number(value) === Number(show)
  );
}


function getActivePromoForCategory_(
  promos,
  category,
  show,
  now
) {
  const matches = promos.filter(promo =>
    promo.active &&
    promo.category === category &&
    promoAppliesToShow_(promo, show) &&
    promo.startAt.getTime() <= now.getTime() &&
    promo.endAt.getTime() >= now.getTime() &&
    promo.quota > 0
  );

  if (matches.length > 1) {
    throw new Error(
      `Ada lebih dari satu promo aktif untuk kategori "${category}" ` +
      `pada Show ${show}. Periksa jadwal di sheet Promos.`
    );
  }

  return matches.length
    ? matches[0]
    : null;
}


/**
 * Kuota terpakai dihitung dari kursi Held, Booked, dan Sold
 * yang sudah mempunyai Promo_ID.
 */
function getPromoUsageMap_() {
  const usage = {};

  Object.keys(CONFIG.DATA_SHEETS).forEach(showKey => {
    const show = Number(showKey);
    const sheet = getDataSheet_(show);
    ensureOperationalColumns_(sheet);
    const table = readTable_(sheet);

    table.rows.forEach(row => {
      const status =
        String(row.Status || '')
          .trim()
          .toLowerCase();

      const promoId =
        String(row.Promo_ID || '').trim();

      if (
        !promoId ||
        !['held', 'booked', 'sold'].includes(status)
      ) {
        return;
      }

      usage[promoId] =
        Number(usage[promoId] || 0) + 1;
    });
  });

  return usage;
}


/**
 * Menghitung harga untuk SATU PILIHAN kursi sekaligus.
 *
 * SINGLE:
 *   promo diterapkan per tiket sampai kuota habis.
 *
 * BUNDLE:
 *   promo hanya diterapkan pada kelompok lengkap dalam kategori
 *   yang sama. Contoh Bundle_Qty 3:
 *   pilih 1-2 tiket = reguler,
 *   pilih 3-5 tiket = 3 promo + sisanya reguler,
 *   pilih 6 tiket = 6 promo (selama kuota cukup).
 *
 * Kuota Promos dihitung dalam jumlah tiket/seat, bukan jumlah bundle.
 */
function buildLockedPricingForSelection_(
  show,
  requestedIds,
  bySeat,
  regularPrices,
  promos,
  promoUsage,
  now
) {
  const lockedPricing = new Map();
  const idsByCategory = {};

  requestedIds.forEach(id => {
    const item = bySeat.get(id);
    const category =
      String(item.row.Category || '').trim();

    if (!idsByCategory[category]) {
      idsByCategory[category] = [];
    }

    idsByCategory[category].push(id);
  });

  Object.entries(idsByCategory)
    .forEach(([category, seatIds]) => {
      const regularPrice =
        getCategoryPrice_(
          regularPrices,
          category
        );

      const promo =
        getActivePromoForCategory_(
          promos,
          category,
          show,
          now
        );

      let promoCount = 0;
      let remaining = 0;

      if (promo) {
        const used =
          Number(
            promoUsage[promo.promoId] || 0
          );

        remaining =
          Math.max(
            0,
            promo.quota - used
          );

        if (promo.promoType === 'BUNDLE') {
          const bundleQty =
            Math.max(
              2,
              Number(promo.bundleQty || 2)
            );

          const eligibleFromSelection =
            Math.floor(
              seatIds.length / bundleQty
            ) * bundleQty;

          const eligibleFromQuota =
            Math.floor(
              remaining / bundleQty
            ) * bundleQty;

          promoCount =
            Math.min(
              eligibleFromSelection,
              eligibleFromQuota
            );
        } else {
          promoCount =
            Math.min(
              seatIds.length,
              remaining
            );
        }
      }

      seatIds.forEach((id, index) => {
        const usePromo =
          Boolean(
            promo &&
            index < promoCount
          );

        const bundleQty =
          promo &&
          promo.promoType === 'BUNDLE'
            ? Math.max(
                2,
                Number(promo.bundleQty || 2)
              )
            : 1;

        lockedPricing.set(id, {
          price:
            usePromo
              ? promo.promoPrice
              : regularPrice,
          regularPrice,
          promoId:
            usePromo
              ? promo.promoId
              : '',
          promoName:
            usePromo
              ? promo.promoName
              : '',
          promoType:
            usePromo
              ? promo.promoType || 'SINGLE'
              : '',
          bundleQty:
            usePromo
              ? bundleQty
              : 1,
          bundleTotal:
            usePromo &&
            promo.promoType === 'BUNDLE'
              ? Number(
                  promo.bundleTotal ||
                  promo.promoPrice * bundleQty
                )
              : null,
          remaining
        });
      });

      if (promo && promoCount > 0) {
        promoUsage[promo.promoId] =
          Number(
            promoUsage[promo.promoId] || 0
          ) + promoCount;
      }
    });

  return lockedPricing;
}


function getEffectivePricing_(
  show,
  category,
  regularPrices,
  promos,
  promoUsage,
  now
) {
  const regularPrice =
    getCategoryPrice_(
      regularPrices,
      category
    );

  const promo =
    getActivePromoForCategory_(
      promos,
      category,
      show,
      now
    );

  if (!promo) {
    return {
      price: regularPrice,
      regularPrice,
      promoId: '',
      promoName: '',
      promoType: '',
      bundleQty: 1,
      bundleTotal: null,
      remaining: 0
    };
  }

  const used =
    Number(
      promoUsage[promo.promoId] || 0
    );

  const remaining =
    Math.max(
      0,
      promo.quota - used
    );

  /*
   * Helper ini adalah pricing satu seat.
   * Promo BUNDLE sengaja tidak diterapkan di sini.
   * Bundle harus dihitung lewat buildLockedPricingForSelection_().
   */
  if (
    remaining <= 0 ||
    promo.promoType === 'BUNDLE'
  ) {
    return {
      price: regularPrice,
      regularPrice,
      promoId: '',
      promoName: '',
      promoType: '',
      bundleQty: 1,
      bundleTotal: null,
      remaining
    };
  }

  return {
    price: promo.promoPrice,
    regularPrice,
    promoId: promo.promoId,
    promoName: promo.promoName,
    promoType: promo.promoType || 'SINGLE',
    bundleQty: 1,
    bundleTotal: null,
    remaining
  };
}


function buildPublicPromoInfo_(
  show,
  regularPrices,
  promos,
  promoUsage,
  now
) {
  return Object.keys(regularPrices)
    .map(category => {
      const promo =
        getActivePromoForCategory_(
          promos,
          category,
          show,
          now
        );

      if (!promo) return null;

      const used =
        Number(
          promoUsage[promo.promoId] || 0
        );

      const remaining =
        Math.max(
          0,
          promo.quota - used
        );

      const bundleQty =
        promo.promoType === 'BUNDLE'
          ? Math.max(2, Number(promo.bundleQty || 2))
          : 1;

      return {
        promoId: promo.promoId,
        promoName: promo.promoName,
        category: promo.category,
        price: promo.promoPrice,
        regularPrice:
          getCategoryPrice_(
            regularPrices,
            category
          ),
        quota: promo.quota,
        remaining,
        promoType: promo.promoType || 'SINGLE',
        bundleQty,
        bundleTotal:
          promo.promoType === 'BUNDLE'
            ? Number(
                promo.bundleTotal ||
                promo.promoPrice * bundleQty
              )
            : null,
        bundlesRemaining:
          promo.promoType === 'BUNDLE'
            ? Math.floor(remaining / bundleQty)
            : null,
        startAt:
          promo.startAt.toISOString(),
        endAt:
          promo.endAt.toISOString()
      };
    })
    .filter(Boolean);
}


function getNextPromoChangeAt_(
  promos,
  show,
  now
) {
  const nowMs = now.getTime();
  const candidates = [];

  promos.forEach(promo => {
    if (
      !promo.active ||
      !promoAppliesToShow_(promo, show)
    ) {
      return;
    }

    const startMs =
      promo.startAt.getTime();

    const endMs =
      promo.endAt.getTime();

    if (startMs > nowMs) {
      candidates.push(startMs);
    }

    if (endMs > nowMs) {
      // Tambah sedikit waktu agar request berikutnya
      // benar-benar melewati batas End_At.
      candidates.push(endMs + 1000);
    }
  });

  if (!candidates.length) {
    return '';
  }

  return new Date(
    Math.min(...candidates)
  ).toISOString();
}


function getLockedPriceForRow_(
  row,
  regularPrices
) {
  const locked =
    toNumber_(row.Locked_Price);

  if (
    locked !== null &&
    locked >= 0
  ) {
    return locked;
  }

  const category =
    String(row.Category || '').trim();

  return getCategoryPrice_(
    regularPrices,
    category
  );
}


function summarizePromos_(rows) {
  const counts = {};

  rows.forEach(row => {
    const promoId =
      String(row.Promo_ID || '').trim();

    if (!promoId) return;

    const promoName =
      String(row.Promo_Name || '').trim() ||
      promoId;

    const key =
      `${promoName} [${promoId}]`;

    counts[key] =
      Number(counts[key] || 0) + 1;
  });

  return Object.entries(counts)
    .map(([name, count]) =>
      `${count} ${name}`
    )
    .join(', ');
}


/**
 * Jalankan sekali jika ingin membuat sheet Promos otomatis.
 */
function setupPromoSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(CONFIG.PROMO_SHEET);

  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.PROMO_SHEET);
  }

  const headers = [
    'Promo_ID',
    'Promo_Name',
    'Category',
    'Promo_Price',
    'Start_At',
    'End_At',
    'Quota',
    'Active',
    'Show',
    'Promo_Type',
    'Bundle_Qty',
    'Bundle_Total'
  ];

  if (sheet.getLastRow() === 0) {
    sheet
      .getRange(1, 1, 1, headers.length)
      .setValues([headers]);

    sheet.setFrozenRows(1);
  } else {
    const currentHeaders =
      getHeaders_(sheet);

    const missing =
      headers.filter(
        header =>
          !currentHeaders.includes(header)
      );

    if (missing.length) {
      sheet
        .getRange(
          1,
          currentHeaders.length + 1,
          1,
          missing.length
        )
        .setValues([missing]);
    }

    sheet.setFrozenRows(1);
  }

  return {
    success: true,
    sheet: sheet.getName(),
    headers
  };
}


/**
 * URL gambar untuk header website.
 * Mendukung URL image biasa dan link Google Drive /file/d/... .
 */
function getBookingSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(CONFIG.BOOKING_SHEET);

  if (!sheet) {
    throw new Error(
      `Sheet "${CONFIG.BOOKING_SHEET}" tidak ditemukan.`
    );
  }

  return sheet;
}

function ensureBookingHeaders_(sheet) {
  const requiredHeaders = [
    'Booking_ID',
    'Hold_Token',
    'Show',
    'Seats',
    'Category_Summary',
    'Promo_Summary',
    'Total',
    'Name',
    'Email',
    'WhatsApp',
    'Referral_Penyanyi',
    'Customer_Type',
    'Student_University',
    'Student_Cohort',
    'Student_KTM_URL',
    'Student_KTM_Hash',
    'Student_Promo_Tickets',
    'Status',
    'Payment_Status',
    'Payment_Method',
    'Payment_Deadline',
    'Payment_Proof_URL',
    'Ticket_Code',
    'Ticket_QR_URL',
    'Verified_At',
    'Email_Sent_At',
    'WhatsApp_Template',
    'WhatsApp_Link',
    'Created_At',
    'Updated_At'
  ];

  if (sheet.getLastRow() === 0) {
    sheet
      .getRange(
        1,
        1,
        1,
        requiredHeaders.length
      )
      .setValues([
        requiredHeaders
      ]);

    return;
  }

  const headers =
    getHeaders_(sheet);

  const missing =
    requiredHeaders.filter(
      header =>
        !headers.includes(header)
    );

  if (!missing.length) {
    return;
  }

  const startColumn =
    headers.length + 1;

  sheet
    .getRange(
      1,
      startColumn,
      1,
      missing.length
    )
    .setValues([
      missing
    ]);
}


function bookingExists_(sheet, bookingId) {
  const table = readBookingTable_(sheet);

  return table.rows.some(row =>
    String(row.Booking_ID || '').trim() === bookingId
  );
}


function readBookingTable_(sheet) {
  ensureBookingHeaders_(sheet);

  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(value =>
    String(value || '').trim()
  );

  const rows = values.slice(1).map(row =>
    Object.fromEntries(
      headers.map((header, index) => [
        header,
        row[index]
      ])
    )
  );

  return {
    headers,
    rows
  };
}


function isStudentVoyagePromoRow_(row) {
  const promoText = [
    row && row.Promo_ID,
    row && row.Promo_Name
  ]
    .map(value => String(value || '').trim().toLowerCase())
    .join(' ');

  return (
    promoText.includes('student voyage') ||
    promoText.includes('student crew') ||
    promoText.includes('mahasiswa')
  );
}


function repriceHeldStudentVoyageToRegular_(
  show,
  bookingId,
  holdToken,
  regularPrices,
  updatedAt
) {
  const sheet = getDataSheet_(show);
  ensureOperationalColumns_(sheet);

  const table = readTable_(sheet);
  const headers = getHeaders_(sheet);
  let changed = false;

  table.rows.forEach((row, index) => {
    const matches =
      String(row.Booking_ID || '').trim() === bookingId &&
      String(row.Hold_Token || '').trim() === holdToken &&
      String(row.Status || '').trim().toLowerCase() === 'held';

    if (!matches || !isStudentVoyagePromoRow_(row)) {
      return;
    }

    const sheetRow = index + 2;
    const category = String(row.Category || '').trim();
    const regularPrice = getCategoryPrice_(regularPrices, category);

    setCellByHeader_(
      sheet,
      sheetRow,
      headers,
      'Locked_Price',
      regularPrice
    );

    setCellByHeader_(
      sheet,
      sheetRow,
      headers,
      'Promo_ID',
      ''
    );

    setCellByHeader_(
      sheet,
      sheetRow,
      headers,
      'Promo_Name',
      ''
    );

    setCellByHeader_(
      sheet,
      sheetRow,
      headers,
      'Updated_At',
      updatedAt
    );

    changed = true;
  });

  if (changed) {
    SpreadsheetApp.flush();
  }
}


function prepareStudentKtmUpload_(studentKtm) {
  if (!studentKtm) {
    throw new Error('Upload KTM aktif untuk menggunakan promo Student Voyage.');
  }

  const fileName =
    String(studentKtm.fileName || 'ktm').trim();
  const mimeType =
    String(studentKtm.mimeType || '').trim();
  const base64 =
    String(studentKtm.base64 || '').trim();

  if (!['image/jpeg', 'image/png'].includes(mimeType)) {
    throw new Error('KTM harus berupa JPG atau PNG.');
  }

  if (!base64) {
    throw new Error('File KTM kosong.');
  }

  let decoded;

  try {
    decoded = Utilities.base64Decode(base64);
  } catch (error) {
    throw new Error('File KTM tidak valid. Silakan upload ulang.');
  }

  if (decoded.length > STUDENT_VOYAGE_CONFIG.maxFileBytes) {
    throw new Error('Ukuran KTM maksimal 5 MB.');
  }

  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    decoded
  );

  const hash = digest
    .map(byte => (byte < 0 ? byte + 256 : byte).toString(16).padStart(2, '0'))
    .join('');

  return {
    fileName,
    mimeType,
    bytes: decoded,
    hash
  };
}


function saveStudentKtmFile_(bookingId, customerName, prepared) {
  const extension =
    prepared.mimeType === 'image/png'
      ? 'png'
      : 'jpg';

  const safeCustomerName =
    String(customerName || 'student')
      .replace(/[^a-zA-Z0-9_-]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .slice(0, 50) || 'student';

  const safeOriginalName =
    String(prepared.fileName || 'ktm')
      .replace(/\.[^.]+$/, '')
      .replace(/[^a-zA-Z0-9_-]+/g, '-')
      .replace(/^-+|-+$/g, '')
      .slice(0, 50) || 'ktm';

  const finalName =
    `${bookingId}_${safeCustomerName}_${safeOriginalName}_${Date.now()}.${extension}`;

  const blob = Utilities.newBlob(
    prepared.bytes,
    prepared.mimeType,
    finalName
  );

  const folder = getStudentKtmFolder_();
  const file = folder.createFile(blob);

  return file.getUrl();
}


function getStudentKtmFolder_() {
  const folders =
    DriveApp.getFoldersByName(
      STUDENT_VOYAGE_CONFIG.ktmFolderName
    );

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(
    STUDENT_VOYAGE_CONFIG.ktmFolderName
  );
}


function countStudentVoyageTicketsByKtmHash_(bookingSheet, ktmHash) {
  const normalizedHash =
    String(ktmHash || '').trim().toLowerCase();

  if (!normalizedHash) {
    return 0;
  }

  const inactiveStatuses = new Set([
    'expired',
    'declined',
    'cancelled',
    'canceled'
  ]);

  const table = readBookingTable_(bookingSheet);

  return table.rows.reduce((total, row) => {
    const rowHash =
      String(row.Student_KTM_Hash || '').trim().toLowerCase();

    if (rowHash !== normalizedHash) {
      return total;
    }

    const status =
      String(row.Status || '').trim().toLowerCase();

    if (inactiveStatuses.has(status)) {
      return total;
    }

    const count = Math.max(
      0,
      Math.floor(Number(row.Student_Promo_Tickets || 0))
    );

    return total + count;
  }, 0);
}


function findHeldBooking_(bookingId, holdToken) {
  const prices = getPrices_();

  for (const showKey of Object.keys(CONFIG.DATA_SHEETS)) {
    const show = Number(showKey);

    releaseExpiredHolds_(show);

    const sheet = getDataSheet_(show);
    const table = readTable_(sheet);

    const rows = table.rows.filter(row => {
      return (
        String(row.Booking_ID || '').trim() === bookingId &&
        String(row.Hold_Token || '').trim() === holdToken &&
        String(row.Status || '').trim().toLowerCase() === 'held'
      );
    });

    if (rows.length) {
      const holdDates = rows
        .map(row => parseStoredDate_(row.Hold_Until))
        .filter(Boolean);

      return {
        show,
        rows,
        prices,
        holdUntil: holdDates.length ? holdDates[0] : null
      };
    }
  }

  return {
    show: null,
    rows: [],
    prices,
    holdUntil: null
  };
}


function updateHeldSeatsToBooked_(
  show,
  bookingId,
  holdToken,
  updatedAt
) {
  const sheet = getDataSheet_(show);
  const table = readTable_(sheet);
  const headers = getHeaders_(sheet);

  table.rows.forEach((row, index) => {
    const matches =
      String(row.Booking_ID || '').trim() === bookingId &&
      String(row.Hold_Token || '').trim() === holdToken &&
      String(row.Status || '').trim().toLowerCase() === 'held';

    if (!matches) return;

    const sheetRow = index + 2;

    setCellByHeader_(
      sheet,
      sheetRow,
      headers,
      'Status',
      'Booked'
    );

    setCellByHeader_(
      sheet,
      sheetRow,
      headers,
      'Updated_At',
      updatedAt
    );
  });
}



function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Mozzafiato Booking')
    .addItem(
      'Verify & Send Ticket',
      'verifySelectedBooking'
    )
    .addToUi();
}


/**
 * Jalankan dari menu Mozzafiato Booking ketika satu sel
 * pada baris booking sedang dipilih.
 */
function verifySelectedBooking() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSheet();

  if (sheet.getName() !== CONFIG.BOOKING_SHEET) {
    ui.alert(
      `Buka sheet "${CONFIG.BOOKING_SHEET}" terlebih dahulu.`
    );
    return;
  }

  ensureBookingHeaders_(sheet);

  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    ui.alert('Pilih salah satu baris booking.');
    return;
  }

  const headers = getHeaders_(sheet);
  const bookingIdColumn =
    headers.indexOf('Booking_ID') + 1;

  if (!bookingIdColumn) {
    ui.alert('Kolom Booking_ID tidak ditemukan.');
    return;
  }

  const bookingId = String(
    sheet.getRange(row, bookingIdColumn).getValue() || ''
  ).trim();

  if (!bookingId) {
    ui.alert('Booking ID pada baris ini kosong.');
    return;
  }

  try {
    const result =
      verifyBookingAndSendTicket(bookingId);

    ui.alert(
      'Berhasil',
      `Tiket ${result.bookingId} sudah diverifikasi dan dikirim ke ${result.email}.`,
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert(
      'Gagal',
      error && error.message
        ? error.message
        : String(error),
      ui.ButtonSet.OK
    );
  }
}

function finalizeVerifiedBooking_(
  bookingId,
  booking,
  bookingRowNumber,
  bookingHeaders,
  now
) {
  const email = String(booking.Email || '').trim();
  const whatsapp = String(booking.WhatsApp || '').trim();

  if (!email || !isValidEmail_(email)) {
    throw new Error('Email booking kosong atau tidak valid.');
  }

  const ticketCode =
    String(booking.Ticket_Code || '').trim() ||
    createTicketCode_(bookingId);
  const qrUrl = buildQrImageUrl_(
    buildTicketQrPayload_(bookingId, ticketCode)
  );
  const statusUrl = buildBookingStatusUrl_(bookingId);

  sendTicketEmail_({
    ...booking,
    Booking_ID: bookingId,
    Ticket_Code: ticketCode,
    Ticket_QR_URL: qrUrl,
    Status_URL: statusUrl
  });

  const whatsappTemplate = buildWhatsappTemplate_({
    ...booking,
    Booking_ID: bookingId,
    Ticket_Code: ticketCode,
    Status_URL: statusUrl
  });
  const whatsappLink = buildWhatsappLink_(
    whatsapp,
    whatsappTemplate
  );

  bookingHeaders.forEach((header, index) => {
    const updates = {
      Status: 'Verified',
      Payment_Status: 'Verified',
      Ticket_Code: ticketCode,
      Ticket_QR_URL: qrUrl,
      Verified_At: now,
      Email_Sent_At: now,
      WhatsApp_Template: whatsappTemplate,
      WhatsApp_Link: whatsappLink,
      Updated_At: now
    };

    if (Object.prototype.hasOwnProperty.call(updates, header)) {
      SpreadsheetApp
        .getActive()
        .getSheetByName(CONFIG.BOOKING_SHEET)
        .getRange(bookingRowNumber, index + 1)
        .setValue(updates[header]);
    }
  });

  updateBookedSeatsToSold_(
    Number(booking.Show),
    bookingId,
    now
  );

  return {
    ticketCode,
    qrUrl,
    statusUrl,
    whatsappTemplate,
    whatsappLink
  };
}


/**
 * Verifikasi pembayaran, ubah seat menjadi Sold,
 * kirim email tiket, dan tulis template WhatsApp.
 */
function verifyBookingAndSendTicket(bookingIdValue) {
  const bookingId =
    String(bookingIdValue || '').trim();

  if (!bookingId) {
    throw new Error('Booking ID wajib diisi.');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const sheet = getBookingSheet_();
    ensureBookingHeaders_(sheet);

    const table = readBookingTable_(sheet);
    const index = table.rows.findIndex(row =>
      String(row.Booking_ID || '').trim() === bookingId
    );

    if (index === -1) {
      throw new Error('Booking tidak ditemukan.');
    }

    const booking = table.rows[index];
    const email = String(booking.Email || '').trim();
    const whatsapp = String(
      booking.WhatsApp || ''
    ).trim();

    if (!email || !isValidEmail_(email)) {
      throw new Error(
        'Email booking kosong atau tidak valid.'
      );
    }

    const currentStatus = String(
      booking.Status || ''
    ).trim().toLowerCase();

    if (currentStatus === 'expired') {
      throw new Error(
        'Booking sudah Expired dan tidak dapat diverifikasi.'
      );
    }

    const now = new Date();
    const ticketCode =
      String(booking.Ticket_Code || '').trim() ||
      createTicketCode_(bookingId);

    const qrPayload =
      buildTicketQrPayload_(
        bookingId,
        ticketCode
      );

    const qrUrl =
      buildQrImageUrl_(qrPayload);

    const statusUrl =
      buildBookingStatusUrl_(bookingId);

    sendTicketEmail_({
      ...booking,
      Booking_ID: bookingId,
      Ticket_Code: ticketCode,
      Ticket_QR_URL: qrUrl,
      Status_URL: statusUrl
    });

    const whatsappTemplate =
      buildWhatsappTemplate_({
        ...booking,
        Booking_ID: bookingId,
        Ticket_Code: ticketCode,
        Status_URL: statusUrl
      });

    const whatsappLink =
      buildWhatsappLink_(
        whatsapp,
        whatsappTemplate
      );

    const row = index + 2;
    const headers = getHeaders_(sheet);

    const updates = {
      Status: 'Verified',
      Payment_Status: 'Verified',
      Ticket_Code: ticketCode,
      Ticket_QR_URL: qrUrl,
      Verified_At: now,
      Email_Sent_At: now,
      WhatsApp_Template: whatsappTemplate,
      WhatsApp_Link: whatsappLink,
      Updated_At: now
    };

    Object.entries(updates).forEach(
      ([header, value]) =>
        setCellByHeader_(
          sheet,
          row,
          headers,
          header,
          value
        )
    );

    const show = Number(booking.Show);

    if (CONFIG.DATA_SHEETS[show]) {
      updateBookedSeatsToSold_(
        show,
        bookingId,
        now
      );

      syncVisualSheet_(show);
    }

    SpreadsheetApp.flush();

    return {
      success: true,
      bookingId,
      email,
      ticketCode,
      qrUrl,
      statusUrl,
      whatsappTemplate,
      whatsappLink
    };
  } finally {
    lock.releaseLock();
  }
}


function getBookingStatus(bookingIdValue) {
  const bookingId =
    String(bookingIdValue || '').trim();

  if (!bookingId) {
    throw new Error('Masukkan Booking ID.');
  }

  const sheet = getBookingSheet_();
  ensureBookingHeaders_(sheet);

  const table = readBookingTable_(sheet);
  const booking = table.rows.find(row =>
    String(row.Booking_ID || '')
      .trim()
      .toLowerCase() ===
    bookingId.toLowerCase()
  );

  if (!booking) {
    throw new Error(
      'Booking ID tidak ditemukan. Periksa kembali kode booking-mu.'
    );
  }

  return {
    bookingId:
      String(booking.Booking_ID || '').trim(),

    show:
      Number(booking.Show || 0),

    seats:
      String(booking.Seats || '').trim(),

    categorySummary:
      String(
        booking.Category_Summary || ''
      ).trim(),

    total:
      Number(booking.Total || 0),

    status:
      String(booking.Status || '').trim(),

    paymentStatus:
      String(
        booking.Payment_Status || ''
      ).trim(),

    paymentDeadline:
      toIsoDateString_(booking.Payment_Deadline),

    ticketCode:
      String(
        booking.Ticket_Code || ''
      ).trim(),

    ticketQrUrl:
      String(
        booking.Ticket_QR_URL || ''
      ).trim(),

    emailSentAt:
      booking.Email_Sent_At instanceof Date
        ? booking.Email_Sent_At.toISOString()
        : '',

    customerName:
      String(booking.Name || '').trim()
  };
}


function updateBookedSeatsToSold_(
  show,
  bookingId,
  updatedAt
) {
  const sheet = getDataSheet_(show);
  const table = readTable_(sheet);
  const headers = getHeaders_(sheet);

  table.rows.forEach((row, index) => {
    const matches =
      String(row.Booking_ID || '').trim() ===
      bookingId;

    const status =
      String(row.Status || '')
        .trim()
        .toLowerCase();

    if (!matches || status !== 'booked') {
      return;
    }

    const sheetRow = index + 2;

    setCellByHeader_(
      sheet,
      sheetRow,
      headers,
      'Status',
      'Sold'
    );

    setCellByHeader_(
      sheet,
      sheetRow,
      headers,
      'Updated_At',
      updatedAt
    );
  });
}


function createTicketCode_(bookingId) {
  const hash = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    `${bookingId}|${Utilities.getUuid()}`
  );

  const shortCode = hash
    .slice(0, 5)
    .map(byte =>
      (byte + 256)
        .toString(16)
        .padStart(2, '0')
    )
    .join('')
    .toUpperCase();

  return `MZF-${shortCode}`;
}


function buildTicketQrPayload_(
  bookingId,
  ticketCode
) {
  return [
    'MOZZAFIATO-2026',
    `BOOKING:${bookingId}`,
    `TICKET:${ticketCode}`
  ].join('|');
}


function buildQrImageUrl_(payload) {
  return (
    'https://quickchart.io/qr' +
    '?size=500' +
    '&margin=2' +
    '&text=' +
    encodeURIComponent(payload)
  );
}


function buildBookingStatusUrl_(bookingId) {
  const webAppUrl =
    ScriptApp.getService().getUrl();

  if (!webAppUrl) {
    return '';
  }

  return (
    webAppUrl +
    '?booking=' +
    encodeURIComponent(bookingId)
  );
}


function sendTicketEmail_(booking) {
  const bookingId =
    String(booking.Booking_ID || '').trim();

  const ticketCode =
    String(booking.Ticket_Code || '').trim();

  const email =
    String(booking.Email || '').trim();

  const name =
    String(booking.Name || 'Guest').trim();

  const show =
    String(booking.Show || '').trim();

  const seats =
    String(booking.Seats || '').trim();

  const qrUrl =
    String(booking.Ticket_QR_URL || '').trim();

  const statusUrl =
    String(booking.Status_URL || '').trim();

  const subject =
    `🚢 MOZZAFIATO CONCERT 2026: Resonant Voyage 🌊 — ${bookingId}`;

  let eticketAssets = null;

  try {
    eticketAssets = buildEticketAssets_({
      ...booking,
      Booking_ID: bookingId,
      Ticket_Code: ticketCode,
      Ticket_QR_URL: qrUrl
    });
  } catch (error) {
    console.error(
      'E-ticket assets gagal dibuat:',
      error
    );
  }

  const eticketInlineImage =
    eticketAssets &&
    eticketAssets.pngBlob
      ? eticketAssets.pngBlob
      : null;

  const eticketAttachment =
    eticketAssets &&
    eticketAssets.pdfBlob
      ? eticketAssets.pdfBlob
      : null;

  const calendarUrl =
    buildGoogleCalendarUrl_(booking);

  const statusButton =
    statusUrl
      ? `
        <a
          href="${escapeHtmlServer_(statusUrl)}"
          style="
            display:inline-block;
            width:100%;
            max-width:270px;
            padding:14px 18px;
            border-radius:999px;
            background:#14324a;
            color:#ffffff;
            text-decoration:none;
            font-family:Arial,sans-serif;
            font-size:14px;
            font-weight:700;
            text-align:center;
            margin:0 6px 10px;
          "
        >
          Periksa Status Ticket
        </a>
      `
      : '';

  const calendarButton =
    calendarUrl
      ? `
        <a
          href="${escapeHtmlServer_(calendarUrl)}"
          style="
            display:inline-block;
            width:100%;
            max-width:300px;
            padding:14px 18px;
            border-radius:999px;
            background:#e9f2f8;
            color:#14324a;
            text-decoration:none;
            font-family:Arial,sans-serif;
            font-size:14px;
            font-weight:700;
            text-align:center;
            margin:0 6px 10px;
            border:1px solid #d3e1ea;
          "
        >
          Tambahkan Acara Ke Google Calendar
        </a>
      `
      : '';

  const heroBlock =
    eticketInlineImage
      ? `
        <tr>
          <td
            align="center"
            style="padding:0 16px 0;"
          >
            <img
              src="cid:eticket_visual"
              width="620"
              alt="Mozzafiato 2026 E-Ticket"
              style="
                display:block;
                width:100%;
                max-width:620px;
                height:auto;
                margin:0 auto;
                border:0;
                border-radius:18px;
              "
            >
          </td>
        </tr>
      `
      : `
        <tr>
          <td
            align="center"
            style="padding:24px 24px 8px;"
          >
            <img
              src="${escapeHtmlServer_(qrUrl)}"
              width="300"
              alt="Mozzafiato Ticket QR"
              style="
                display:block;
                width:100%;
                max-width:300px;
                height:auto;
                border:0;
              "
            >
          </td>
        </tr>
      `;

  const htmlBody = `
    <!doctype html>
    <html>
      <body
        style="
          margin:0;
          padding:0;
          background:#eef4f8;
        "
      >
        <table
          role="presentation"
          width="100%"
          cellspacing="0"
          cellpadding="0"
          border="0"
          style="
            width:100%;
            border-collapse:collapse;
            background:#eef4f8;
          "
        >
          <tr>
            <td
              align="center"
              style="padding:18px 8px;"
            >
              <table
                role="presentation"
                width="620"
                cellspacing="0"
                cellpadding="0"
                border="0"
                style="
                  width:100%;
                  max-width:620px;
                  border-collapse:collapse;
                  overflow:hidden;
                  border-radius:22px;
                  background:#ffffff;
                  box-shadow:0 12px 34px rgba(20,50,74,.12);
                "
              >
                <tr>
                  <td
                    align="center"
                    style="
                      padding:26px 24px 14px;
                      font-family:Arial,sans-serif;
                      color:#14324a;
                      background:#ffffff;
                    "
                  >
                    <div
                      style="
                        margin:0 0 8px;
                        color:#14324a;
                        font-size:18px;
                        line-height:1.5;
                        font-weight:800;
                      "
                    >
                      🚢 MOZZAFIATO CONCERT 2026: Resonant Voyage 🌊
                    </div>

                    <div
                      style="
                        margin:0;
                        font-size:28px;
                        line-height:1.2;
                        font-weight:800;
                        color:#14324a;
                      "
                    >
                      Tiket kamu sudah terkonfirmasi!
                    </div>

                    <div
                      style="
                        margin:12px auto 0;
                        max-width:500px;
                        color:#627487;
                        font-size:14px;
                        line-height:1.6;
                      "
                    >
                      Hai ${escapeHtmlServer_(name)}, simpan e-ticket ini
                      dan tunjukkan QR saat check-in.
                    </div>
                  </td>
                </tr>

                ${heroBlock}

                <tr>
                  <td
                    align="center"
                    style="
                      padding:18px 24px 6px;
                      font-family:Arial,sans-serif;
                      color:#627487;
                      font-size:13px;
                      line-height:1.7;
                    "
                  >
                    <strong style="color:#14324a;">Booking ID:</strong>
                    ${escapeHtmlServer_(bookingId)}
                    <br>
                    <strong style="color:#14324a;">Show:</strong>
                    Show ${escapeHtmlServer_(show)}
                    &nbsp;•&nbsp;
                    <strong style="color:#14324a;">Seat:</strong>
                    ${escapeHtmlServer_(seats)}
                  </td>
                </tr>

                <tr>
                  <td
                    align="center"
                    style="padding:16px 24px 4px;"
                  >
                    ${statusButton}
                    ${calendarButton}
                  </td>
                </tr>

                <tr>
                  <td
                    align="center"
                    style="
                      padding:4px 24px 18px;
                      font-family:Arial,sans-serif;
                      color:#627487;
                      font-size:13px;
                      line-height:1.7;
                    "
                  >
                    <strong style="color:#14324a;">
                      Sabtu, 7 November 2026
                    </strong>
                    <br>
                    Reminder:
                    6 Jam Sebelum, H-1, H-7
                  </td>
                </tr>

                <tr>
                  <td
                    align="center"
                    style="
                      padding:18px 24px 24px;
                      border-top:1px solid #e5edf2;
                      font-family:Arial,sans-serif;
                      color:#7a8997;
                      font-size:12px;
                      line-height:1.6;
                    "
                  >
                    Jangan bagikan QR kepada orang lain.
                    Setiap tiket hanya dapat digunakan sesuai status check-in.
                  </td>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </body>
    </html>
  `;

  const plainBody = [
    '🚢 MOZZAFIATO CONCERT 2026: Resonant Voyage 🌊',
    '',
    'Tiket kamu sudah terkonfirmasi!',
    `Booking ID: ${bookingId}`,
    `Ticket Code: ${ticketCode}`,
    `Show: ${show}`,
    `Seats: ${seats}`,
    'Sabtu, 7 November 2026',
    'Reminder: 6 Jam Sebelum, H-1, H-7',
    statusUrl
      ? `Periksa Status Ticket: ${statusUrl}`
      : '',
    calendarUrl
      ? `Tambahkan Acara Ke Google Calendar: ${calendarUrl}`
      : '',
    '',
    'Simpan e-ticket ini dan tunjukkan QR saat check-in.'
  ].filter(Boolean).join('\n');

  const emailOptions = {
    to: email,
    subject,
    body: plainBody,
    htmlBody,
    name:
      CONFIG.EMAIL_SENDER_NAME ||
      'Mozzafiato 2026'
  };

  if (eticketInlineImage) {
    emailOptions.inlineImages = {
      eticket_visual:
        eticketInlineImage
    };
  }

  if (eticketAttachment) {
    emailOptions.attachments = [
      eticketAttachment
    ];
  }

  MailApp.sendEmail(emailOptions);
}

function buildGoogleCalendarUrl_(booking) {
  const title =
    '🚢 MOZZAFIATO CONCERT 2026: Resonant Voyage 🌊';

  const show =
    String(booking.Show || '').trim();

  const seats =
    String(booking.Seats || '').trim();

  const bookingId =
    String(booking.Booking_ID || '').trim();

  const details = [
    'Mozzafiato 2026 — Resonant Voyage',
    show ? `Show ${show}` : '',
    seats ? `Seats: ${seats}` : '',
    bookingId ? `Booking ID: ${bookingId}` : '',
    '',
    'Reminder yang disarankan:',
    '- 6 Jam Sebelum',
    '- H-1',
    '- H-7'
  ].filter(Boolean).join('\n');

  // All-day event for Saturday, 7 November 2026.
  const startDate = '20261107';
  const endDate = '20261108';

  const params = [
    'action=TEMPLATE',
    `text=${encodeURIComponent(title)}`,
    `dates=${encodeURIComponent(startDate + '/' + endDate)}`,
    `details=${encodeURIComponent(details)}`
  ];

  return 'https://calendar.google.com/calendar/render?' +
    params.join('&');
}



function getEticketTemplateConfig_(showValue) {
  const cfg = CONFIG.ETICKET || {};
  const templates = cfg.templates || {};

  const rawShow =
    String(showValue || '').trim();

  const match =
    rawShow.match(/\d+/);

  const show =
    match
      ? Number(match[0])
      : Number(rawShow);

  if (![1, 2].includes(show)) {
    throw new Error(
      `Show e-ticket tidak valid: ${rawShow || '(kosong)'}`
    );
  }

  const template =
    templates[show];

  if (
    !template ||
    !String(
      template.presentationId || ''
    ).trim()
  ) {
    throw new Error(
      `Template e-ticket untuk Show ${show} belum dikonfigurasi.`
    );
  }

  return {
    show,

    presentationId:
      String(
        template.presentationId
      ).trim(),

    slideIndex:
      Math.max(
        0,
        Number(
          template.slideIndex || 0
        )
      )
  };
}

function buildEticketAssets_(booking) {
  const cfg = CONFIG.ETICKET || {};

  if (cfg.enabled !== true) {
    return null;
  }

  const templateConfig =
    getEticketTemplateConfig_(
      booking.Show ||
      booking.show
    );

  const templateId =
    templateConfig.presentationId;

  const slideIndex =
    templateConfig.slideIndex;

  const templateFile =
    DriveApp.getFileById(
      templateId
    );

  const tempFolder = getOrCreateFolder_(
    cfg.tempFolderName || 'Mozzafiato E-Ticket Temp'
  );

  const bookingId =
    String(booking.Booking_ID || '').trim();

  const copy = templateFile.makeCopy(
    `TEMP_${bookingId}_${Date.now()}`,
    tempFolder
  );
  const copyId = copy.getId();

  try {
    const presentation =
      SlidesApp.openById(copyId);

    const slide =
      presentation.getSlides()[slideIndex];

    if (!slide) {
      throw new Error(
        `Slide index ${slideIndex} tidak ditemukan pada template.`
      );
    }

    const replacements = {
      '{{BOOKING_ID}}': String(booking.Booking_ID || ''),
      '{{TICKET_CODE}}': String(booking.Ticket_Code || ''),
      '{{SHOW}}': booking.Show ? `Show ${booking.Show}` : '',
      '{{SEATS}}': String(booking.Seats || ''),
      '{{NAME}}': String(booking.Name || ''),
      '{{TOTAL}}': `Rp${Number(
        booking.Total || 0
      ).toLocaleString('id-ID')}`
    };

    Object.entries(replacements).forEach(([placeholder, value]) => {
      presentation.replaceAllText(placeholder, value);
    });

    replaceQrPlaceholder_(
      slide,
      String(booking.Ticket_QR_URL || '').trim()
    );

    presentation.saveAndClose();
    Utilities.sleep(900);

    const slideObjectId = slide.getObjectId();

    const pdfBlob = exportGoogleFileAsPdf_(copyId);
    pdfBlob.setName(`E-Ticket_${bookingId}.pdf`);

    const pngBlob = exportSlideAsPng_(copyId, slideObjectId);
    pngBlob.setName(`E-Ticket_${bookingId}.png`);

    return {
      pdfBlob,
      pngBlob
    };
  } finally {
    try {
      DriveApp.getFileById(copyId).setTrashed(true);
    } catch (cleanupError) {
      console.error(cleanupError);
    }
  }
}


function buildEticketPdf_(booking) {
  const assets = buildEticketAssets_(booking);

  return assets && assets.pdfBlob
    ? assets.pdfBlob
    : null;
}


function replaceQrPlaceholder_(slide, qrUrl) {
  if (!qrUrl) {
    throw new Error('URL QR ticket kosong.');
  }

  const placeholder = slide.getPageElements().find(element => {
    const title = String(element.getTitle() || '').trim();
    const description = String(element.getDescription() || '').trim();

    return (
      title === 'QR_PLACEHOLDER' ||
      description === 'QR_PLACEHOLDER'
    );
  });

  if (!placeholder) {
    throw new Error(
      'Shape dengan Alt text QR_PLACEHOLDER tidak ditemukan pada template.'
    );
  }

  const left = placeholder.getLeft();
  const top = placeholder.getTop();
  const width = placeholder.getWidth();
  const height = placeholder.getHeight();

  placeholder.remove();

  const response = UrlFetchApp.fetch(qrUrl, {
    muteHttpExceptions: true
  });

  if (
    response.getResponseCode() < 200 ||
    response.getResponseCode() >= 300
  ) {
    throw new Error(
      `QR tidak dapat diunduh. HTTP ${response.getResponseCode()}.`
    );
  }

  slide.insertImage(
    response.getBlob().setName('ticket-qr.png'),
    left,
    top,
    width,
    height
  );
}


function exportGoogleFileAsPdf_(fileId) {
  const url =
    'https://www.googleapis.com/drive/v3/files/' +
    encodeURIComponent(fileId) +
    '/export?mimeType=' +
    encodeURIComponent('application/pdf');

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });

  if (
    response.getResponseCode() < 200 ||
    response.getResponseCode() >= 300
  ) {
    throw new Error(
      'Gagal mengekspor e-ticket PDF. HTTP ' +
      response.getResponseCode()
    );
  }

  return response.getBlob();
}


function exportSlideAsPng_(presentationId, slideObjectId) {
  const url =
    'https://docs.google.com/presentation/d/' +
    encodeURIComponent(presentationId) +
    '/export/png?pageid=' +
    encodeURIComponent(slideObjectId);

  const response = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
    },
    muteHttpExceptions: true
  });

  if (
    response.getResponseCode() < 200 ||
    response.getResponseCode() >= 300
  ) {
    throw new Error(
      'Gagal mengekspor e-ticket PNG. HTTP ' +
      response.getResponseCode()
    );
  }

  return response.getBlob();
}


function getOrCreateFolder_(folderName) {
  const name =
    String(folderName || '').trim() ||
    'Mozzafiato E-Ticket Temp';

  const folders = DriveApp.getFoldersByName(name);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(name);
}



function testEticketSlidesShow1() {
  return testEticketSlidesForShow_(1);
}


function testEticketSlidesShow2() {
  return testEticketSlidesForShow_(2);
}


function testEticketSlidesForShow_(show) {
  const bookingId =
    `TEST-SHOW${show}-${Date.now()}`;

  const ticketCode =
    `MZF-TEST-S${show}`;

  const qrUrl =
    buildQrImageUrl_(
      buildTicketQrPayload_(
        bookingId,
        ticketCode
      )
    );

  const assets =
    buildEticketAssets_({
      Booking_ID: bookingId,
      Ticket_Code: ticketCode,
      Show: show,
      Seats: show === 1
        ? 'A1, A2'
        : 'B1, B2',
      Name: `Test Show ${show}`,
      Total: 190000,
      Ticket_QR_URL: qrUrl
    });

  if (!assets || !assets.pdfBlob || !assets.pngBlob) {
    throw new Error(
      `E-ticket Show ${show} gagal dibuat.`
    );
  }

  const folder =
    getOrCreateFolder_(
      'Mozzafiato E-Ticket Test'
    );

  const pdfFile =
    folder.createFile(
      assets.pdfBlob.copyBlob().setName(
        `TEST_E-Ticket_Show_${show}.pdf`
      )
    );

  const pngFile =
    folder.createFile(
      assets.pngBlob.copyBlob().setName(
        `TEST_E-Ticket_Show_${show}.png`
      )
    );

  Logger.log(
    `SHOW ${show} PDF: ${pdfFile.getUrl()}`
  );

  Logger.log(
    `SHOW ${show} PNG: ${pngFile.getUrl()}`
  );

  return {
    show,
    pdfUrl: pdfFile.getUrl(),
    pngUrl: pngFile.getUrl()
  };
}

function testEmbeddedEticketEmail() {
  const recipient = Session.getActiveUser().getEmail();

  if (!recipient) {
    throw new Error(
      'Email akun aktif tidak terbaca. Isi manual pada fungsi ini bila perlu.'
    );
  }

  const booking = {
    Booking_ID: 'TEST-EMBED-001',
    Ticket_Code: 'MZF-EMBED001',
    Show: 1,
    Seats: 'B3, B4, B5, B6',
    Name: 'Test Mozzafiato',
    Email: recipient,
    Total: 380000
  };

  const qrUrl = buildQrImageUrl_(
    buildTicketQrPayload_(
      booking.Booking_ID,
      booking.Ticket_Code
    )
  );

  // Tes aset terlebih dahulu supaya error PNG tidak tertutup fallback email.
  const assets = buildEticketAssets_({
    ...booking,
    Ticket_QR_URL: qrUrl
  });

  if (!assets || !assets.pngBlob) {
    throw new Error(
      'PNG embedded tidak berhasil dibuat.'
    );
  }

  sendTicketEmail_({
    ...booking,
    Ticket_QR_URL: qrUrl
  });

  return {
    success: true,
    recipient,
    pngBytes: assets.pngBlob.getBytes().length,
    pdfBytes: assets.pdfBlob
      ? assets.pdfBlob.getBytes().length
      : 0
  };
}


function buildWhatsappTemplate_(booking) {
  const name =
    String(booking.Name || 'Kak').trim();

  const bookingId =
    String(booking.Booking_ID || '').trim();

  const showValue =
    String(booking.Show || '').trim();

  const showLabel = showValue
    ? `Show ${showValue}`
    : '-';

  const seats =
    String(booking.Seats || '').trim();

  const statusUrl =
    String(booking.Status_URL || '').trim();

  const lines = [
    `Halo ${name}! 👋`,
    '',
    'Pembayaran tiket Mozzafiato 2026 kamu sudah kami konfirmasi 🥳',
    '',
    `🎫 Booking ID: ${bookingId}`,
    `🎙️ Show: ${showLabel}`,
    `💺 Seat: ${seats}`,
    '',
    'E-ticket berisi QR booking sudah kami kirimkan melalui email yang digunakan saat pemesanan. Silakan cek folder Inbox, Promotions, atau Spam.',
    '',
    'Cek status booking:',
    statusUrl ? `🔗 ${statusUrl}` : '',
    '',
    'Terima kasih dan sampai jumpa di Mozzafiato Concert 2026! ⛵'
  ];

  return lines.join('\n');
}


function buildWhatsappLink_(
  whatsapp,
  message
) {
  const number = normalizeWhatsappNumber_(
    whatsapp
  );

  if (!number) {
    return '';
  }

  return (
    'https://wa.me/' +
    number +
    '?text=' +
    encodeURIComponent(message)
  );
}


function getShowContact_(showValue) {
  const show = normalizeShow_(showValue);
  const contacts = CONFIG.SUPPORT_CONTACTS || {};
  const contact = contacts[show] || {};

  const name = String(contact.name || '').trim();
  const whatsapp = normalizeWhatsappNumber_(
    contact.whatsapp || ''
  );

  return {
    show,
    name,
    whatsapp,
    waUrl: whatsapp
      ? 'https://wa.me/' + whatsapp
      : ''
  };
}


function buildVerificationNotice_(showValue) {
  const contact = getShowContact_(showValue);
  const hours = Math.max(
    1,
    Number(CONFIG.VERIFICATION_HOURS || 24)
  );

  const contactText = contact.name
    ? ` Jika setelah ${hours} jam belum terverifikasi, silakan hubungi ${contact.name} melalui WhatsApp.`
    : '';

  return (
    `Tiket akan diverifikasi maksimal 1x24 jam setelah bukti pembayaran dikirim.${contactText}`
  );
}


function normalizeWhatsappNumber_(value) {
  let number =
    String(value || '')
      .replace(/\D/g, '');

  if (!number) {
    return '';
  }

  if (number.startsWith('0')) {
    number = '62' + number.slice(1);
  } else if (number.startsWith('8')) {
    number = '62' + number;
  }

  return number;
}


function escapeHtmlServer_(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}


function isValidEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(
    String(email || '').trim()
  );
}

function getPaymentInfo() {
  const payment =
    CONFIG.PAYMENT || {};

  const qris =
    payment.qris || {};

  const banks =
    Array.isArray(payment.banks)
      ? payment.banks
          .filter(bank =>
            bank &&
            bank.enabled !== false
          )
          .map((bank, index) => {
            const rawId =
              bank.id ||
              bank.bankName ||
              `bank-${index + 1}`;

            return {
              id: String(rawId)
                .trim()
                .toLowerCase()
                .replace(
                  /[^a-z0-9_-]/g,
                  '-'
                ),

              bankName: String(
                bank.bankName || ''
              ).trim(),

              accountNumber: String(
                bank.accountNumber || ''
              ).trim(),

              accountHolder: String(
                bank.accountHolder || ''
              ).trim()
            };
          })
          .filter(bank =>
            bank.bankName &&
            bank.accountNumber &&
            bank.accountHolder
          )
      : [];

  const originalQrisUrl =
    String(
      qris.imageUrl || ''
    ).trim();

  return {
    paymentMinutes: Math.max(
      1,
      Number(
        payment.paymentMinutes || 15
      )
    ),

  qrisMaxTotal: Math.max(
    0,
    Number(
      payment.qrisMaxTotal || 500000
    )
  ),

    qris: {
      enabled:
        qris.enabled !== false &&
        Boolean(originalQrisUrl),

      imageUrl:
        convertDriveImageUrl_(
          originalQrisUrl
        )
    },

    banks
  };
}


function convertDriveImageUrl_(url) {
  const value =
    String(url || '').trim();

  if (!value) {
    return '';
  }

  /*
   * Format Google Drive:
   * drive.google.com/file/d/FILE_ID/view
   */
  const filePathMatch =
    value.match(
      /drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/
    );

  if (
    filePathMatch &&
    filePathMatch[1]
  ) {
    return (
      'https://drive.google.com/thumbnail' +
      '?id=' +
      encodeURIComponent(
        filePathMatch[1]
      ) +
      '&sz=w1200'
    );
  }

  /*
   * Format:
   * drive.google.com/open?id=FILE_ID
   * drive.google.com/uc?id=FILE_ID
   */
  const queryIdMatch =
    value.match(
      /[?&]id=([a-zA-Z0-9_-]+)/
    );

  if (
    queryIdMatch &&
    queryIdMatch[1]
  ) {
    return (
      'https://drive.google.com/thumbnail' +
      '?id=' +
      encodeURIComponent(
        queryIdMatch[1]
      ) +
      '&sz=w1200'
    );
  }

  /*
   * Jika bukan link Drive,
   * URL dikembalikan apa adanya.
   */
  return value;
}

function uploadPaymentProof(request) {
  if (!request) {
    throw new Error('Data bukti pembayaran tidak ditemukan.');
  }

  const bookingId = String(request.bookingId || '').trim();
  const paymentMethod = String(request.paymentMethod || '').trim();
  const fileName = String(request.fileName || 'payment-proof').trim();
  const mimeType = String(request.mimeType || '').trim();
  const base64 = String(request.base64 || '').trim();

  if (!bookingId) {
    throw new Error('Booking ID tidak valid.');
  }

  if (!paymentMethod) {
    throw new Error('Pilih metode pembayaran terlebih dahulu.');
  }

  if (!base64) {
    throw new Error('File bukti pembayaran kosong.');
  }

  if (!['image/jpeg', 'image/png'].includes(mimeType)) {
    throw new Error('Bukti pembayaran harus berupa JPG atau PNG.');
  }

  const bookingSheet = getBookingSheet_();
  ensureBookingHeaders_(bookingSheet);

  const table = readBookingTable_(bookingSheet);
  const bookingIndex = table.rows.findIndex(row =>
    String(row.Booking_ID || '').trim() === bookingId
  );

  if (bookingIndex === -1) {
    throw new Error('Booking tidak ditemukan.');
  }

  const booking = table.rows[bookingIndex];

  const total = Number(
    booking.Total || 0
  );

  const qrisMaxTotal = Math.max(
    0,
    Number(
      (CONFIG.PAYMENT && CONFIG.PAYMENT.qrisMaxTotal) ||
      500000
    )
  );

  const normalizedPaymentMethod =
    paymentMethod.toLowerCase();

  if (
    normalizedPaymentMethod === 'qris' &&
    total > qrisMaxTotal
  ) {
    throw new Error(
      'Total pembayaran di atas Rp500.000 wajib menggunakan transfer bank.'
    );
  }

  const deadline = parseStoredDate_(booking.Payment_Deadline);
  const now = new Date();

  if (
    deadline &&
    deadline.getTime() < now.getTime()
  ) {
    throw new Error('Batas waktu pembayaran sudah berakhir.');
  }

  const decoded = Utilities.base64Decode(base64);

  if (decoded.length > 5 * 1024 * 1024) {
    throw new Error('Ukuran file maksimal 5 MB.');
  }

  const extension = mimeType === 'image/png' ? 'png' : 'jpg';
  const safeOriginalName = fileName
    .replace(/\.[^.]+$/, '')
    .replace(/[^a-zA-Z0-9_-]+/g, '-')
    .replace(/^-+|-+$/g, '')
    .slice(0, 60) || 'payment-proof';

  const finalName =
    `${bookingId}_${safeOriginalName}_${Date.now()}.${extension}`;

  const blob = Utilities.newBlob(
    decoded,
    mimeType,
    finalName
  );

  const folder = getPaymentProofFolder_();
  const file = folder.createFile(blob);

  const sheetRow = bookingIndex + 2;
  const headers = getHeaders_(bookingSheet);

  setCellByHeader_(
    bookingSheet,
    sheetRow,
    headers,
    'Payment_Method',
    paymentMethod
  );

  setCellByHeader_(
    bookingSheet,
    sheetRow,
    headers,
    'Payment_Status',
    'Waiting for Verification'
  );

  setCellByHeader_(
    bookingSheet,
    sheetRow,
    headers,
    'Status',
    'Waiting for Verification'
  );

  setCellByHeader_(
    bookingSheet,
    sheetRow,
    headers,
    'Payment_Proof_URL',
    file.getUrl()
  );

  setCellByHeader_(
    bookingSheet,
    sheetRow,
    headers,
    'Updated_At',
    now
  );

  return {
    success: true,
    bookingId,
    paymentMethod,
    status: 'Waiting for Verification',
    paymentProofUrl: file.getUrl(),
    contact: getShowContact_(booking.Show),
    verificationNotice:
      buildVerificationNotice_(booking.Show)
  };
}


function getPaymentProofFolder_() {
  const folderName = 'Mozzafiato Payment Proofs';
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(folderName);
}
