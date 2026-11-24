// Public-safe starter defaults (no personal records, no personal signature, no private URLs).
// Your real monthly records & signature sync privately from your own PIN-protected Google Sheet.
(function () {
  const now = new Date();
  const y = now.getFullYear();
  const m = now.getMonth() + 1;
  const lastDay = new Date(Date.UTC(y, m, 0)).getUTCDate();
  const pad = (n) => String(n).padStart(2, '0');
  const startIso = `${y}-${pad(m)}-01`;
  const endIso = `${y}-${pad(m)}-${pad(lastDay)}`;
  const monthNames = [
    'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'
  ];

  window.TOUR_DEFAULTS = {
    profile: {
      companyName: 'Arkan Group',
      reportTitle: 'Tour Expenses (Out of Khulna)',
      destination: 'Working',
      purpose: 'Delar Visit & Surcing',
      designation: 'ASM',
      lockDatesOnMove: true,
      signatureLabels: [
        'Prepared by',
        'Dept. Head',
        'Audit Dept.',
        'Accounts',
        'GM (Operation)',
        'Director'
      ]
    },
    places: [
      'Jhenaidah',
      'Jhenaidah Sadar',
      'charatola',
      'Shohagpur',
      'Chulkani Bazar',
      'KoatChadpur',
      'Jalalpur',
      'Hat Gopalpur',
      'Narkelbaria',
      'Sadhuhati',
      'Bongkira',
      'Bazar Gopalpur',
      'Dakbangla',
      'Kushtia',
      'Mirpur',
      'Moheshpur',
      'Kaliganj',
      'Khalishpur',
      'Houseing',
      'Garagonj',
      'Soilkupa',
      'Chuadanga Sadar',
      'Darshona'
    ],
    defaultSignature: '',
    sampleMonth: {
      id: `${y}-${pad(m)}`,
      label: `${monthNames[m - 1]} ${y}`,
      startDate: startIso,
      endDate: endIso,
      destination: 'Working',
      purpose: 'Delar Visit & Surcing',
      designation: 'ASM',
      showSigDate: true,
      sigDate: `${pad(lastDay)}-${pad(m)}-${y}`,
      days: [
        {
          date: startIso,
          routes: [
            { from: '', to: '', bus: '', metro: '', auto: '', bike: '', rickshaw: '' },
            { from: '', to: '', bus: '', metro: '', auto: '', bike: '', rickshaw: '' }
          ],
          breakfast: '',
          lunch: '',
          dinner: '',
          teaBreak: '',
          hotel: ''
        }
      ]
    }
  };
})();
