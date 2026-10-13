   const CONFIG = {
  notificationEmail: 'elghifarin@gmail.com',
  folderName: 'IndoProperti Data',
  spreadsheetName: 'Data Pengajuan IndoProperti',
  folderProperty: 'INDOPROPERTI_DRIVE_FOLDER_ID',
  spreadsheetProperty: 'INDOPROPERTI_SPREADSHEET_ID'
};

function setupIndoProperti() {
  MailApp.getRemainingDailyQuota();
  const properties = PropertiesService.getScriptProperties();
  let folderId = properties.getProperty(CONFIG.folderProperty);
  let folder = folderId ? DriveApp.getFolderById(folderId) : null;

  if (!folder) {
    folder = DriveApp.createFolder(CONFIG.folderName);
    properties.setProperty(CONFIG.folderProperty, folder.getId());
  }

  let spreadsheetId = properties.getProperty(CONFIG.spreadsheetProperty);
  let spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : null;

  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create(CONFIG.spreadsheetName);
    DriveApp.getFileById(spreadsheet.getId()).moveTo(folder);
    properties.setProperty(CONFIG.spreadsheetProperty, spreadsheet.getId());
  }

  let sheet = spreadsheet.getSheetByName('Data');
  if (!sheet) {
    sheet = spreadsheet.getSheets()[0];
    sheet.setName('Data');
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Waktu', 'Jenis Data', 'ID', 'Nama / Judul', 'Jenis Pengajuan',
      'Kategori', 'Nomor WhatsApp', 'Lokasi', 'Harga / Budget',
      'Luas', 'Catatan / Deskripsi', 'Status', 'File Drive', 'Data JSON'
    ]);
    sheet.setFrozenRows(1);
  } else if (sheet.getRange(1, 14).getValue() !== 'Data JSON') {
    sheet.getRange(1, 14).setValue('Data JSON');
  }

  const result = {
    folderUrl: folder.getUrl(),
    spreadsheetUrl: spreadsheet.getUrl()
  };
  Logger.log(JSON.stringify(result));
  return result;
}

function doGet(e) {
  try {
    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty(CONFIG.spreadsheetProperty);
    if (!spreadsheetId) throw new Error('Jalankan setupIndoProperti terlebih dahulu.');

    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('Data');
    if (!sheet) throw new Error('Sheet Data tidak ditemukan.');

    const values = sheet.getDataRange().getValues();
    const ackToken = e && e.parameter && e.parameter.ack;
    if (ackToken) {
      const ackRecordType = e.parameter.recordType === 'listing' ? 'listing' : 'submission';
      const ackFound = values.slice(1).some(function(row) {
        const rowType = row[1] === 'Listing Diposting' ? 'listing' : 'submission';
        if (rowType !== ackRecordType || !row[13]) return false;
        try {
          const savedRecord = JSON.parse(row[13]);
          return savedRecord.clientRequestId === ackToken;
        } catch (error) {
          return false;
        }
      });
      return jsonpResponse(e.parameter.callback, {
        ok: true,
        acknowledged: ackFound
      });
    }

    const records = { listings: [] };
    values.slice(1).forEach(function(row) {
      const recordType = row[1] === 'Listing Diposting' ? 'listing' : 'submission';
      const record = recordFromRow(row, recordType);
      if (!record || record.deleted) return;
      if (recordType === 'listing') {
        records.listings.push(publicListing(record));
      } else if (!record.isPosted) {
        records.listings.push(publicSubmissionListing(record));
      }
    });

    records.listings.reverse();
    const callback = e && e.parameter && e.parameter.callback;
    if (callback) {
      if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) throw new Error('Callback tidak valid.');
      return ContentService
        .createTextOutput(callback + '(' + JSON.stringify({ ok: true, ...records }) + ');')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return jsonResponse({ ok: true, ...records });
  } catch (error) {
    console.error(error);
    const callback = e && e.parameter && e.parameter.callback;
    if (callback && /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
      return jsonpResponse(callback, { ok: false, error: error.message || String(error) });
    }
    return jsonResponse({ ok: false, error: error.message || String(error) });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Request kosong atau tidak valid.');
    }
    if (e.postData.contents.length > 12 * 1024 * 1024) {
      throw new Error('Ukuran data melebihi batas 12 MB.');
    }

    const data = JSON.parse(e.postData.contents);
    if (!data || data.source !== 'indoproperti') {
      throw new Error('Sumber data tidak valid.');
    }
    if (data.action !== 'upsert' || data.recordType !== 'submission') {
      throw new Error('Endpoint publik hanya menerima pengajuan baru.');
    }

    const allowedTypes = ['Titip Jual Properti', 'Cari / Request Properti'];
    const allowedCategories = [
      'Lahan Industri', 'Bangunan Pabrik', 'Gudang Logistik',
      'Ruko / Komersial', 'Lahan Lainnya'
    ];
    if (allowedTypes.indexOf(data.type) === -1 || allowedCategories.indexOf(data.category) === -1) {
      throw new Error('Jenis atau kategori pengajuan tidak valid.');
    }
    ['name', 'phone', 'location'].forEach(function(field) {
      if (!String(data[field] || '').trim()) throw new Error('Kolom wajib belum diisi: ' + field);
    });
    if (String(data.name).length > 120 || String(data.phone).length > 32 ||
        String(data.location).length > 200 || String(data.price || '').length > 120 ||
        String(data.area || '').length > 120 || String(data.notes || '').length > 4000) {
      throw new Error('Panjang salah satu kolom melebihi batas.');
    }
    const attachments = Array.isArray(data.attachments) ? data.attachments : [];
    if (attachments.length > 5 || attachments.some(function(file) {
      return !file || typeof file.dataUrl !== 'string' || file.dataUrl.length > 2 * 1024 * 1024;
    })) {
      throw new Error('Lampiran tidak valid atau melebihi batas ukuran.');
    }

    const properties = PropertiesService.getScriptProperties();
    const folderId = properties.getProperty(CONFIG.folderProperty);
    const spreadsheetId = properties.getProperty(CONFIG.spreadsheetProperty);
    if (!folderId || !spreadsheetId) {
      throw new Error('Jalankan setupIndoProperti terlebih dahulu.');
    }

    const folder = DriveApp.getFolderById(folderId);
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName('Data');
    if (!sheet) {
      throw new Error('Sheet Data tidak ditemukan. Jalankan setupIndoProperti kembali.');
    }

    const recordType = 'submission';
    const id = String(data.id || '');
    if (!/^IP-[A-Za-z0-9-]{1,80}$/.test(id)) throw new Error('ID pengajuan tidak valid.');
    const rows = sheet.getDataRange().getValues();
    const rowNumber = findRecordRow(rows, recordType, id);
    if (rowNumber) throw new Error('Pengajuan dengan ID tersebut sudah pernah diterima.');

    const folderLinks = saveAttachments(folder, attachments, false);
    const record = {
      recordType,
      id,
      timestamp: String(data.timestamp || ''),
      type: data.type,
      category: data.category,
      name: String(data.name).trim(),
      phone: String(data.phone).trim(),
      location: String(data.location).trim(),
      price: String(data.price || '-').trim(),
      area: String(data.area || '-').trim(),
      notes: String(data.notes || '-').trim(),
      status: 'Baru',
      isPosted: false,
      attachments: folderLinks,
      clientRequestId: String(data.clientRequestId || '')
    };
    const nameOrTitle = record.name;
    const type = record.type;
    const row = [
      new Date(),
      'Pengajuan',
      sheetSafeCell(id),
      sheetSafeCell(nameOrTitle),
      sheetSafeCell(type),
      sheetSafeCell(record.category),
      sheetSafeCell(record.phone),
      sheetSafeCell(record.location),
      sheetSafeCell(record.price),
      sheetSafeCell(record.area),
      sheetSafeCell(record.notes),
      'Baru',
      folderLinks.map(function(file) { return file.url; }).join('\n'),
      JSON.stringify(record)
    ];

    sheet.appendRow(row);

    if (!data.skipNotification) {
      const subject = 'Pengajuan baru IndoProperti: ' + (nameOrTitle || id);
      const emailBody = [
        'Data baru telah disimpan ke Google Drive.',
        '',
        'Jenis data: Pengajuan',
        'ID: ' + id,
        'Nama / Judul: ' + (nameOrTitle || '-'),
        'Lokasi: ' + (record.location || '-'),
        'Spreadsheet: ' + spreadsheet.getUrl(),
        'Folder Drive: ' + folder.getUrl(),
        folderLinks.length ? 'File:\n' + folderLinks.map(function(file) { return file.url; }).join('\n') : 'File: tidak ada'
      ].join('\n');
      MailApp.sendEmail(CONFIG.notificationEmail, subject, emailBody);
    }

    return jsonResponse({ ok: true, spreadsheetUrl: spreadsheet.getUrl(), folderUrl: folder.getUrl() });
  } catch (error) {
    console.error(error);
    return jsonResponse({ ok: false, error: error.message || String(error) });
  }
}

function publicListing(record) {
  return {
    id: record.id,
    title: record.title || '',
    category: record.category || '',
    statusTag: record.statusTag || '',
    price: record.price || '',
    location: record.location || '',
    lt: record.lt || '',
    lb: record.lb || '',
    sertifikat: record.sertifikat || '',
    zonasi: record.zonasi || '',
    image: record.image || '',
    desc: record.desc || ''
  };
}

function publicSubmissionListing(record) {
  const isRequest = /cari|request/i.test(String(record.type || ''));
  const areas = String(record.area || '').split('/');
  return publicListing({
    id: record.id,
    title: `${isRequest ? 'Mencari ' : ''}${record.category || 'Properti'} - ${record.location || 'Lokasi belum diisi'}`,
    category: record.category || '',
    statusTag: isRequest ? 'DICARI' : 'DIJUAL',
    price: record.price || '',
    location: record.location || '',
    lt: areas[0] ? areas[0].trim() : '',
    lb: areas[1] ? areas[1].trim() : '',
    image: 'https://placehold.co/800x600/0f172a/f59e0b?text=Properti+Industri',
    desc: record.notes || ''
  });
}

function jsonpResponse(callback, value) {
  if (!callback || !/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback)) {
    return jsonResponse(value);
  }
  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(value) + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function saveAttachments(folder, attachments, publishImages) {
  if (!Array.isArray(attachments)) return [];

  return attachments.map(function(attachment) {
    if (attachment && attachment.url && !attachment.dataUrl) {
      const metadata = {
        name: attachment.name || '',
        type: attachment.type || '',
        size: attachment.size || 0,
        url: attachment.url,
        viewUrl: attachment.viewUrl || ''
      };
      if (publishImages && metadata.type.indexOf('image/') === 0) {
        const fileIdMatch = metadata.url.match(/\/d\/([A-Za-z0-9_-]+)/) ||
          metadata.url.match(/[?&]id=([A-Za-z0-9_-]+)/);
        if (fileIdMatch) {
          const file = DriveApp.getFileById(fileIdMatch[1]);
          file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          metadata.viewUrl = 'https://drive.google.com/uc?export=view&id=' + file.getId();
        }
      }
      return {
        ...metadata
      };
    }
    if (!attachment || typeof attachment.dataUrl !== 'string') {
      throw new Error('Format lampiran tidak valid.');
    }

    const match = attachment.dataUrl.match(/^data:([^;,]+);base64,([A-Za-z0-9+/=]+)$/);
    if (!match) throw new Error('Data file tidak valid: ' + (attachment.name || 'lampiran'));

    const bytes = Utilities.base64Decode(match[2]);
    const blob = Utilities.newBlob(
      bytes,
      match[1],
      String(attachment.name || 'lampiran').replace(/[\\/:*?"<>|]/g, '_')
    );
    const file = folder.createFile(blob);
    if (publishImages && match[1].indexOf('image/') === 0) {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }
    return {
      name: blob.getName(),
      type: match[1],
      size: bytes.length,
      url: file.getUrl(),
      viewUrl: match[1].indexOf('image/') === 0
        ? 'https://drive.google.com/uc?export=view&id=' + file.getId()
        : ''
    };
  });
}

function findRecordRow(rows, recordType, id) {
  const expectedType = recordType === 'listing' ? 'Listing Diposting' : 'Pengajuan';
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][1] === expectedType && String(rows[i][2]) === id) return i + 1;
  }
  return 0;
}

function recordFromRow(row, recordType) {
  if (row[13]) {
    try {
      return JSON.parse(row[13]);
    } catch (error) {
      console.error('JSON record tidak valid pada baris ' + (row[0] || 'unknown'), error);
    }
  }

  if (!row[2]) return null;
  if (recordType === 'listing') {
    const areas = String(row[9] || '').split('/');
    const fileUrl = String(row[12] || '').split('\n')[0];
    return {
      recordType: 'listing',
      id: String(row[2]),
      title: String(row[3] || ''),
      statusTag: String(row[4] || 'DIJUAL'),
      category: String(row[5] || ''),
      phone: String(row[6] || ''),
      location: String(row[7] || ''),
      price: String(row[8] || ''),
      lt: areas[0] || '',
      lb: areas[1] || '',
      desc: String(row[10] || ''),
      status: String(row[11] || 'Diposting'),
      image: fileUrl || '',
      attachments: []
    };
  }
  return {
    recordType: 'submission',
    id: String(row[2]),
    name: String(row[3] || ''),
    type: String(row[4] || ''),
    category: String(row[5] || ''),
    phone: String(row[6] || ''),
    location: String(row[7] || ''),
    price: String(row[8] || ''),
    area: String(row[9] || ''),
    notes: String(row[10] || ''),
    status: String(row[11] || 'Baru'),
    isPosted: false,
    attachments: []
  };
}

function sheetSafeCell(value) {
  const text = String(value == null ? '' : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function jsonResponse(value) {
  return ContentService
    .createTextOutput(JSON.stringify(value))
    .setMimeType(ContentService.MimeType.JSON);
}
