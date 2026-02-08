// === HELPER: Parse progres data ===
function parseProgres(text, userRow, username) {
  let data = {
    channel: '',
    scOrderNo: '',
    serviceNo: '',
    customerName: '',
    workzone: '',
    contactPhone: '',
    odp: '',
    memo: '',
    symptom: '',
    ao: '',
    workorder: '',
    tikor: '',
    snOnt: '',
    nikOnt: '',
    stbId: '',
    nikStb: '',
    dateCreated: new Date().toLocaleDateString('id-ID', {
      weekday: 'long',
      day: 'numeric',
      month: 'long',
      year: 'numeric',
      timeZone: 'Asia/Jakarta',
    }),
    teknisi: (username || '').replace('@', ''),
  };

  const patterns = {
    channel: /CHANNEL\s*:\s*([A-Za-z0-9]+)/i,
    scOrderNo: /SC\s*ORDER\s*NO\s*:\s*(.+?)(?=\n|$)/i,
    serviceNo: /SERVICE\s*NO\s*:\s*([0-9]+)/i,
    customerName: /CUSTOMER\s*NAME\s*:\s*(.+?)(?=\n|$)/i,
    workzone: /WORKZONE\s*:\s*([A-Za-z0-9]+)/i,
    contactPhone: /CONTACT\s*PHONE\s*:\s*([0-9\+\-\s]+)/i,
    odp: /ODP\s*:\s*(.+?)(?=\n|$)/i,
    memo: /MEMO\s*:\s*(.+?)(?=\n|$)/i,
    symptom: /SYMPTOM\s*:\s*(.+?)(?=\n|$)/i,
    ao: /AO\s*:\s*(.+?)(?=\n|$)/i,
    workorder: /WORKORDER\s*:\s*([A-Za-z0-9]+)/i,
    tikor: /TIKOR\s*:\s*(.+?)(?=\n|$)/i,
    snOnt: /SN\s*ONT\s*:\s*(.+?)(?=\n|$)/i,
    nikOnt: /NIK\s*ONT\s*:\s*([0-9]+)/i,
    stbId: /STB\s*ID\s*:\s*(.+?)(?=\n|$)/i,
    nikStb: /NIK\s*STB\s*:\s*([0-9]+)/i,
  };

  for (const [key, pattern] of Object.entries(patterns)) {
    const match = text.match(pattern);
    if (match && match[1]) {
      data[key] = match[1].trim();
    }
  }

  return data;
}

// === HELPER: Parse aktivasi data ===
function parseAktivasi(text, username) {
  let data = {
    channel: '',
    dateCreated: '',
    scOrderNo: '',
    workorder: '',
    ao: '',
    ncli: '',
    serviceNo: '',
    address: '',
    customerName: '',
    workzone: '',
    contactPhone: '',
    bookingDate: '',
    paket: '',
    package: '',
    odp: '',
    mitra: '',
    symptom: '',
    memo: '',
    tikor: '',
    snOnt: '',
    nikOnt: '',
    stbId: '',
    nikStb: '',
    teknisi: (username || '').replace('@', ''),
  };

  const patterns = {
    channel: /CHANNEL\s*:\s*(.+?)(?=\n|$)/i,
    dateCreated: /DATE\s*CREATED\s*:\s*(.+?)(?=\n|$)/i,
    scOrderNo: /SC\s*ORDER\s*NO\s*:\s*(.+?)(?=\n|$)/i,
    workorder: /WORKORDER\s*:\s*(.+?)(?=\n|$)/i,
    ao: /AO\s*:\s*(.+?)(?=\n|$)/i,
    ncli: /NCLI\s*:\s*(.+?)(?=\n|$)/i,
    serviceNo: /SERVICE\s*NO\s*:\s*(.+?)(?=\n|$)/i,
    address: /ADDRESS\s*:\s*(.+?)(?=\n|$)/i,
    customerName: /CUSTOMER\s*NAME\s*:\s*(.+?)(?=\n|$)/i,
    workzone: /WORKZONE\s*:\s*(.+?)(?=\n|$)/i,
    contactPhone: /CONTACT\s*PHONE\s*:\s*(.+?)(?=\n|$)/i,
    bookingDate: /BOOKING\s*DATE\s*:\s*(.+?)(?=\n|$)/i,
    paket: /PAKET\s*:\s*(.+?)(?=\n|$)/i,
    package: /PACKAGE\s*:\s*(.+?)(?=\n|$)/i,
    odp: /ODP\s*:\s*(.+?)(?=\n|$)/i,
    mitra: /MITRA\s*:\s*(.+?)(?=\n|$)/i,
    symptom: /SYMPTOM\s*:\s*(.+?)(?=\n|$)/i,
    memo: /MEMO\s*:\s*(.+?)(?=\n|$)/i,
    tikor: /TIKOR\s*:\s*(.+?)(?=\n|$)/i,
    snOnt: /SN\s*ONT\s*:\s*(.+?)(?=\n|$)/i,
    nikOnt: /NIK\s*ONT\s*:\s*(.+?)(?=\n|$)/i,
    stbId: /STB\s*ID\s*:\s*(.+?)(?=\n|$)/i,
    nikStb: /NIK\s*STB\s*:\s*(.+?)(?=\n|$)/i,
  };


  const row = [
        parsed.dateCreated,    // A: DATE CREATED
        parsed.channel,        // B: CHANNEL
        parsed.workorder,      // C: WORKORDER
        parsed.ao,             // D: AO
        parsed.scOrderNo,      // E: SC ORDER NO
        parsed.serviceNo,      // F: SERVICE NO
        parsed.customerName,   // G: CUSTOMER NAME
        parsed.workzone,       // H: WORKZONE
        parsed.contactPhone,   // I: CONTACT PHONE
        parsed.odp,            // J: ODP
        parsed.symptom,        // K: SYMPTOM
        parsed.memo,           // L: MEMO
        parsed.tikor,          // M: TIKOR
        parsed.snOnt,          // N: SN ONT
        parsed.nikOnt,         // O: NIK ONT
        parsed.stbId,          // P: STB ID
        parsed.nikStb,         // Q: NIK STB
        parsed.teknisi,        // R: NAMA TELEGRAM TEKNISI
      ];