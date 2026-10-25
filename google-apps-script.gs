   const CONFIG = {
  notificationEmail: 'elghifarin@gmail.com',
  folderName: 'IndoProperti Data',
  spreadsheetName: 'Data Pengajuan IndoProperti',
  spreadsheetId: '1ONdzYzsd477Am8R8CWI3xB6tP6ohOyK4fALDbHcL1U0',
  folderId: '1oy2W6st6ryxtRt5IVDvvImAwYHBKdaZ1',
  folderProperty: 'INDOPROPERTI_DRIVE_FOLDER_ID',
  spreadsheetProperty: 'INDOPROPERTI_SPREADSHEET_ID',
  adminTokenProperty: 'INDOPROPERTI_ADMIN_TOKEN'
};

function setupIndoProperti() {
  MailApp.getRemainingDailyQuota();
  const properties = PropertiesService.getScriptProperties();
  let folderId = CONFIG.folderId || properties.getProperty(CONFIG.folderProperty);
  let folder = folderId ? DriveApp.getFolderById(folderId) : null;

  if (!folder) {
    folder = DriveApp.createFolder(CONFIG.folderName);
  }
  properties.setProperty(CONFIG.folderProperty, folder.getId());

  let spreadsheetId = CONFIG.spreadsheetId || properties.getProperty(CONFIG.spreadsheetProperty);
  let spreadsheet = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : null;

  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.create(CONFIG.spreadsheetName);
    DriveApp.getFileById(spreadsheet.getId()).moveTo(folder);
  }
  properties.setProperty(CONFIG.spreadsheetProperty, spreadsheet.getId());

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
    const ackToken = e && e.parameter && e.parameter.ack;
    if (ackToken) {
      const cachedAck = CacheService.getScriptCache().get('ack:' + ackToken);
      if (cachedAck) return jsonpResponse(e.parameter.callback, JSON.parse(cachedAck));
    }

    const properties = PropertiesService.getScriptProperties();
    const spreadsheetId = properties.getProperty(CONFIG.spreadsheetProperty);
    if (!spreadsheetId) throw new Error('Jalankan setupIndoProperti terlebih dahulu.');

    const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('Data');
    if (!sheet) throw new Error('Sheet Data tidak ditemukan.');

    const values = sheet.getDataRange().getValues();
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
  let data = null;
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('Request kosong atau tidak valid.');
    }
    if (e.postData.contents.length > 12 * 1024 * 1024) {
      throw new Error('Ukuran data melebihi batas 12 MB.');
    }

    data = JSON.parse(e.postData.contents);
    if (!data || data.source !== 'indoproperti') {
      throw new Error('Sumber data tidak valid.');
    }
    const clientRequestId = String(data.clientRequestId || '');
    if (!/^REQ-\d+-[a-f0-9]{32}$/.test(clientRequestId)) {
      throw new Error('ID permintaan tidak valid.');
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

    const rows = sheet.getDataRange().getValues();
    if (data.action === 'adminListSubmissions') {
      requireAdminToken(data.adminToken);
      const submissions = rows.slice(1).map(function(row) {
        if (row[1] !== 'Pengajuan') return null;
        return recordFromRow(row, 'submission');
      }).filter(function(record) { return record && !record.deleted; }).reverse();
      const offset = Math.max(0, Math.floor(Number(data.offset) || 0));
      const limit = Math.min(10, Math.max(1, Math.floor(Number(data.limit) || 10)));
      const result = {
        ok: true,
        submissions: submissions.slice(offset, offset + limit),
        hasMore: offset + limit < submissions.length
      };
      cacheAcknowledgement(clientRequestId, result);
      return jsonResponse({ ok: true, acknowledged: true });
    }

    const adminOperation = Boolean(data.adminToken);
    if (adminOperation) requireAdminToken(data.adminToken);
    if (!adminOperation && (data.action !== 'upsert' || data.recordType !== 'submission')) {
      throw new Error('Endpoint publik hanya menerima pengajuan baru.');
    }

    if (data.action === 'delete') {
      if (!adminOperation) throw new Error('Otorisasi admin diperlukan.');
      const recordType = validateRecordType(data.recordType);
      const id = validateRecordId(data.id, recordType);
      const rowNumber = findRecordRow(rows, recordType, id);
      const previous = rowNumber ? recordFromRow(rows[rowNumber - 1], recordType) : {};
      const deletedRecord = Object.assign({}, previous, {
        recordType: recordType,
        id: id,
        deleted: true,
        status: previous.status || 'Dihapus',
        attachments: previous.attachments || [],
        clientRequestId: clientRequestId
      });
      writeRecord(sheet, rowNumber, deletedRecord, recordType);
    } else if (data.action === 'upsert') {
      const recordType = validateRecordType(data.recordType);
      if (recordType === 'listing' && !adminOperation) throw new Error('Otorisasi admin diperlukan.');
      const requestedId = validateRecordId(data.id, recordType);
      const existingRow = findRecordRow(rows, recordType, requestedId);
      if (!adminOperation && existingRow) {
        cacheAcknowledgement(clientRequestId, {
          ok: true,
          spreadsheetUrl: spreadsheet.getUrl(),
          duplicate: true
        });
        return jsonResponse({ ok: true, acknowledged: true });
      }
      const record = buildRecord(data, recordType, adminOperation, folder, clientRequestId);
      writeRecord(sheet, existingRow, record, recordType);

      if (!adminOperation && !data.skipNotification) {
        const subject = 'Pengajuan baru IndoProperti: ' + (record.name || record.id);
        const emailBody = [
          'Data baru telah disimpan ke Google Drive.',
          '',
          'Jenis data: Pengajuan',
          'ID: ' + record.id,
          'Nama / Judul: ' + (record.name || '-'),
          'Lokasi: ' + (record.location || '-'),
          'Spreadsheet: ' + spreadsheet.getUrl(),
          'Folder Drive: ' + folder.getUrl(),
          record.attachments.length ? 'File:\n' + record.attachments.map(function(file) { return file.url; }).join('\n') : 'File: tidak ada'
        ].join('\n');
        try {
          MailApp.sendEmail(CONFIG.notificationEmail, subject, emailBody);
        } catch (notificationError) {
          console.error('Pengajuan tersimpan, tetapi notifikasi email gagal dikirim:', notificationError);
        }
      }
    } else {
      throw new Error('Operasi tidak didukung.');
    }

    cacheAcknowledgement(clientRequestId, {
      ok: true,
      spreadsheetUrl: spreadsheet.getUrl(),
      folderUrl: folder.getUrl()
    });
    return jsonResponse({ ok: true, acknowledged: true });
  } catch (error) {
    console.error(error);
    if (data && /^REQ-\d+-[a-f0-9]{32}$/.test(String(data.clientRequestId || ''))) {
      try {
        cacheAcknowledgement(String(data.clientRequestId), {
          ok: false,
          error: error.message || String(error)
        });
      } catch (cacheError) {
        console.error('Gagal menyimpan status error permintaan:', cacheError);
      }
    }
    return jsonResponse({ ok: false, error: error.message || String(error) });
  }
}

function requireAdminToken(token) {
  const expected = PropertiesService.getScriptProperties().getProperty(CONFIG.adminTokenProperty);
  if (!expected || expected.length < 32) {
    throw new Error('Kunci admin belum dikonfigurasi pada Script Properties.');
  }
  if (typeof token !== 'string' || token !== expected) {
    throw new Error('Kunci admin tidak valid.');
  }
}

function cacheAcknowledgement(requestId, result) {
  const succeeded = !result || result.ok !== false;
  CacheService.getScriptCache().put('ack:' + requestId, JSON.stringify({
    ok: succeeded,
    acknowledged: succeeded,
    result: result
  }), 300);
}

function validateRecordType(recordType) {
  if (recordType !== 'submission' && recordType !== 'listing') {
    throw new Error('Jenis record tidak valid.');
  }
  return recordType;
}

function validateRecordId(value, recordType) {
  const id = String(value || '');
  const pattern = recordType === 'listing' ? /^LIST-[A-Za-z0-9-]{1,80}$/ : /^IP-[A-Za-z0-9-]{1,80}$/;
  if (!pattern.test(id)) throw new Error('ID data tidak valid.');
  return id;
}

function buildRecord(data, recordType, adminOperation, folder, clientRequestId) {
  const id = validateRecordId(data.id, recordType);
  const attachments = Array.isArray(data.attachments) ? data.attachments : [];
  if (attachments.length > 5 || attachments.some(function(file) {
    if (!file || String(file.name || '').length > 255 || String(file.type || '').length > 100) return true;
    if (!adminOperation && typeof file.dataUrl !== 'string') return true;
    return file.dataUrl && (typeof file.dataUrl !== 'string' || file.dataUrl.length > 12 * 1024 * 1024);
  })) {
    throw new Error('Lampiran tidak valid atau melebihi batas ukuran.');
  }
  let savedAttachments = [];
  const record = {
    recordType: recordType,
    id: id,
    timestamp: String(data.timestamp || ''),
    clientRequestId: clientRequestId,
    attachments: savedAttachments
  };

  if (recordType === 'submission') {
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
    record.type = data.type;
    record.category = data.category;
    record.name = String(data.name).trim();
    record.phone = String(data.phone).trim();
    record.location = String(data.location).trim();
    record.price = String(data.price || '-').trim();
    record.area = String(data.area || '-').trim();
    record.notes = String(data.notes || '-').trim();
    record.status = adminOperation ? String(data.status || 'Baru') : 'Baru';
    if (['Baru', 'Di-Follow Up', 'Selesai', 'Ditolak'].indexOf(record.status) === -1) {
      throw new Error('Status pengajuan tidak valid.');
    }
    record.isPosted = adminOperation && data.isPosted === true;
  } else {
    if (!String(data.title || '').trim() || String(data.title).length > 180 ||
        String(data.category || '').length > 100 || String(data.location || '').length > 200 ||
        String(data.desc || '').length > 4000 || String(data.price || '').length > 120) {
      throw new Error('Data listing tidak valid atau melebihi batas.');
    }
    record.title = String(data.title).trim();
    record.category = String(data.category || '').trim();
    record.statusTag = String(data.statusTag || 'DIJUAL').slice(0, 40);
    record.price = String(data.price || '-').trim();
    record.location = String(data.location || '').trim();
    record.lt = String(data.lt || '').slice(0, 80);
    record.lb = String(data.lb || '').slice(0, 80);
    record.sertifikat = String(data.sertifikat || '').slice(0, 100);
    record.zonasi = String(data.zonasi || '').slice(0, 100);
    record.desc = String(data.desc || '').trim();
    record.status = String(data.status || 'Diposting').slice(0, 40);
  }
  savedAttachments = saveAttachments(folder, attachments, recordType === 'listing');
  record.attachments = savedAttachments;
  if (recordType === 'listing') {
    const image = String(data.image || '');
    record.image = /^https:\/\/(?:drive\.google\.com|lh3\.googleusercontent\.com|placehold\.co)\//.test(image)
      ? image
      : (savedAttachments.filter(function(file) { return file.type.indexOf('image/') === 0; })[0] || {}).viewUrl || '';
  }
  return record;
}

function writeRecord(sheet, rowNumber, record, recordType) {
  const isListing = recordType === 'listing';
  const storedType = isListing ? 'Listing Diposting' : 'Pengajuan';
  const area = isListing ? [record.lt || '', record.lb || ''].join('/') : record.area;
  const row = [
    new Date(),
    storedType,
    sheetSafeCell(record.id),
    sheetSafeCell(isListing ? record.title : record.name),
    sheetSafeCell(isListing ? record.statusTag : record.type),
    sheetSafeCell(record.category),
    sheetSafeCell(record.phone || ''),
    sheetSafeCell(record.location),
    sheetSafeCell(record.price),
    sheetSafeCell(area || ''),
    sheetSafeCell(isListing ? record.desc : record.notes),
    sheetSafeCell(record.deleted ? 'Dihapus' : record.status),
    record.attachments.map(function(file) { return file.url; }).join('\n'),
    JSON.stringify(record)
  ];
  if (rowNumber) sheet.getRange(rowNumber, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
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
    desc: isRequest ? 'Informasi pencarian properti.' : 'Informasi properti yang ditawarkan.'
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
