/*************************************************
 * INVENTARIS LAB TKJ
 * SMK MADYATAMA PALEMBANG
 *************************************************/


const SHEET_BARANG = 'Barang';
const SHEET_KATEGORI = 'Kategori';
const SHEET_LOKASI = 'Lokasi';
const SHEET_PENGGUNA = 'Pengguna';
const SHEET_PENGATURAN = 'Pengaturan';


/*************************************************
 * WEB APP
 *************************************************/

function doGet() {

  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Inventaris Lab TKJ')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/*************************************************
 * SETUP DATABASE
 *************************************************/

function setupDatabase() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();


  // ================================
  // BARANG
  // ================================

  const barangSheet =
    getOrCreateSheet(
      ss,
      SHEET_BARANG,
      [
        'ID',
        'Kode Inventaris',
        'Nama Barang',
        'Kategori',
        'Merk',
        'Model/Tipe',
        'No. Seri',
        'Tahun',
        'Sumber Dana',
        'Harga',
        'Lokasi',
        'Kondisi',
        'Jumlah',
        'Keterangan',
        'Foto URL',
        'Dibuat',
        'Diubah'
      ]
    );


  // ================================
  // KATEGORI
  // ================================

  const kategoriSheet =
    getOrCreateSheet(
      ss,
      SHEET_KATEGORI,
      [
        'ID',
        'Nama Kategori'
      ]
    );


  if (
    kategoriSheet.getLastRow() < 2
  ) {

    const kategoriDefault = [
      'Komputer',
      'Monitor',
      'Laptop',
      'Printer',
      'Jaringan',
      'Server',
      'UPS',
      'Keyboard',
      'Mouse',
      'Kabel',
      'Peralatan Jaringan',
      'Peralatan Lab',
      'Lainnya'
    ];

    kategoriDefault.forEach(
      function(nama) {

        kategoriSheet.appendRow([
          generateId('KAT'),
          nama
        ]);

      }
    );

  }


  // ================================
  // LOKASI
  // ================================

  const lokasiSheet =
    getOrCreateSheet(
      ss,
      SHEET_LOKASI,
      [
        'ID',
        'Nama Lokasi'
      ]
    );


  if (
    lokasiSheet.getLastRow() < 2
  ) {

    const lokasiDefault = [
      'Lab TKJ Atas',
      'Lab TKJ Bawah',
      'Ruang Guru',
      'Ruang Kaprodi TKJ',
      'Gudang',
      'Lainnya'
    ];

    lokasiDefault.forEach(
      function(nama) {

        lokasiSheet.appendRow([
          generateId('LOK'),
          nama
        ]);

      }
    );

  }


  // ================================
  // PENGGUNA
  // ================================

  const penggunaSheet =
    getOrCreateSheet(
      ss,
      SHEET_PENGGUNA,
      [
        'ID',
        'Nama',
        'Username',
        'Password',
        'Role',
        'Status'
      ]
    );


  if (
    penggunaSheet.getLastRow() < 2
  ) {

    penggunaSheet.appendRow([
      generateId('USR'),
      'Administrator',
      'admin',
      'admin123',
      'Admin',
      'Aktif'
    ]);

  }


  // ================================
  // PENGATURAN
  // ================================

  getOrCreateSheet(
    ss,
    SHEET_PENGATURAN,
    [
      'Nama Pengaturan',
      'Nilai'
    ]
  );


  return 'Database berhasil disiapkan.';

}


/*************************************************
 * CREATE / GET SHEET
 *************************************************/

function getOrCreateSheet(
  ss,
  sheetName,
  headers
) {

  let sheet =
    ss.getSheetByName(sheetName);


  if (!sheet) {

    sheet =
      ss.insertSheet(sheetName);

  }


  if (
    sheet.getLastRow() === 0
  ) {

    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .setValues([headers]);

    sheet
      .getRange(
        1,
        1,
        1,
        headers.length
      )
      .setFontWeight('bold');

    sheet.setFrozenRows(1);

  }


  return sheet;

}


/*************************************************
 * GENERATE ID
 *************************************************/

function generateId(
  prefix
) {

  return (
    prefix +
    '-' +
    new Date()
      .getTime()
      .toString(36)
      .toUpperCase() +
    '-' +
    Math.random()
      .toString(36)
      .substring(2, 7)
      .toUpperCase()
  );

}


/*************************************************
 * LOGIN ADMIN
 *************************************************/

function loginAdmin(
  username,
  password
) {

  setupDatabase();


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        SHEET_PENGGUNA
      );


  const data =
    sheet.getDataRange()
      .getValues();


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    const row =
      data[i];


    if (
      String(row[2]) ===
        String(username) &&
      String(row[3]) ===
        String(password) &&
      String(row[5]).toLowerCase() ===
        'aktif'
    ) {

      return {
        success: true,
        user: {
          id: row[0],
          nama: row[1],
          username: row[2],
          role: row[4]
        }
      };

    }

  }


  return {
    success: false,
    message:
      'Username atau password salah.'
  };

}


/*************************************************
 * GET PUBLIC BARANG
 *************************************************/

function getPublicBarang() {

  setupDatabase();

  return ambilDataBarang();

}


/*************************************************
 * GET ALL BARANG
 *************************************************/

function getAllBarang() {

  setupDatabase();

  return ambilDataBarang();

}


/*************************************************
 * AMBIL DATA BARANG
 *************************************************/

function ambilDataBarang() {

  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        SHEET_BARANG
      );


  if (
    sheet.getLastRow() < 2
  ) {

    return [];

  }


  const data =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        17
      )
      .getValues();


  return data.map(
    function(row) {

      return {

        id: row[0],

        kode: row[1],

        nama: row[2],

        kategori: row[3],

        merk: row[4],

        model: row[5],

        seri: row[6],

        tahun: row[7],

        sumberDana: row[8],

        harga: row[9],

        lokasi: row[10],

        kondisi: row[11],

        jumlah:
          Number(row[12]) || 0,

        keterangan: row[13],

        foto: row[14],

        dibuat:
          formatTanggal(row[15]),

        diubah:
          formatTanggal(row[16])

      };

    }
  );

}


/*************************************************
 * FORMAT TANGGAL
 *************************************************/

function formatTanggal(
  value
) {

  if (!value) {
    return '';
  }


  if (
    Object.prototype
      .toString
      .call(value) ===
    '[object Date]'
  ) {

    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'dd/MM/yyyy HH:mm'
    );

  }


  return String(value);

}


/*************************************************
 * MASTER DATA
 *************************************************/

function getMasterData() {

  setupDatabase();


  return {

    kategori:
      getKategori(),

    lokasi:
      getLokasi()

  };

}


/*************************************************
 * KATEGORI UNTUK DROPDOWN
 * HASIL = ARRAY STRING
 *************************************************/

function getKategori() {

  setupDatabase();


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        SHEET_KATEGORI
      );


  if (
    sheet.getLastRow() < 2
  ) {

    return [];

  }


  const data =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        2
      )
      .getValues();


  return data
    .map(
      function(row) {
        return String(row[1] || '').trim();
      }
    )
    .filter(
      function(item) {
        return item !== '';
      }
    );

}


/*************************************************
 * KATEGORI UNTUK ADMIN
 * HASIL = OBJECT
 *************************************************/

function getKategoriAdmin() {

  setupDatabase();


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        SHEET_KATEGORI
      );


  if (
    sheet.getLastRow() < 2
  ) {

    return [];

  }


  const data =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        2
      )
      .getValues();


  return data
    .map(
      function(row) {

        return {

          id: row[0],

          nama: row[1]

        };

      }
    );

}


/*************************************************
 * TAMBAH KATEGORI
 *************************************************/

function tambahKategori(
  nama
) {

  setupDatabase();


  nama =
    String(nama || '').trim();


  if (!nama) {

    return {
      success: false,
      message:
        'Nama kategori tidak boleh kosong.'
    };

  }


  const kategori =
    getKategori();


  const sudahAda =
    kategori.some(
      function(item) {

        return item.toLowerCase() ===
          nama.toLowerCase();

      }
    );


  if (sudahAda) {

    return {
      success: false,
      message:
        'Kategori tersebut sudah ada.'
    };

  }


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        SHEET_KATEGORI
      );


  sheet.appendRow([
    generateId('KAT'),
    nama
  ]);


  return {
    success: true,
    message:
      'Kategori berhasil ditambahkan.'
  };

}


/*************************************************
 * UPDATE KATEGORI
 *************************************************/

function updateKategori(
  id,
  namaBaru
) {

  setupDatabase();


  namaBaru =
    String(namaBaru || '').trim();


  if (!namaBaru) {

    return {
      success: false,
      message:
        'Nama kategori tidak boleh kosong.'
    };

  }


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      SHEET_KATEGORI
    );


  const data =
    sheet.getDataRange()
      .getValues();


  let namaLama = '';


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      String(data[i][0]) ===
      String(id)
    ) {

      namaLama =
        String(data[i][1]);

      sheet
        .getRange(i + 1, 2)
        .setValue(namaBaru);

      break;

    }

  }


  if (!namaLama) {

    return {
      success: false,
      message:
        'Kategori tidak ditemukan.'
    };

  }


  // Update nama kategori
  // pada data barang

  const barangSheet =
    ss.getSheetByName(
      SHEET_BARANG
    );


  if (
    barangSheet &&
    barangSheet.getLastRow() >= 2
  ) {

    const lastRow =
      barangSheet.getLastRow();


    const kategoriRange =
      barangSheet.getRange(
        2,
        4,
        lastRow - 1,
        1
      );


    const kategoriData =
      kategoriRange.getValues();


    kategoriData.forEach(
      function(row) {

        if (
          String(row[0]) ===
          namaLama
        ) {

          row[0] =
            namaBaru;

        }

      }
    );


    kategoriRange.setValues(
      kategoriData
    );

  }


  return {
    success: true,
    message:
      'Kategori berhasil diperbarui.'
  };

}


/*************************************************
 * DELETE KATEGORI
 *************************************************/

function deleteKategori(
  id
) {

  setupDatabase();


  const ss =
    SpreadsheetApp
      .getActiveSpreadsheet();


  const sheet =
    ss.getSheetByName(
      SHEET_KATEGORI
    );


  const data =
    sheet.getDataRange()
      .getValues();


  let namaKategori = '';
  let rowNumber = 0;


  for (
    let i = 1;
    i < data.length;
    i++
  ) {

    if (
      String(data[i][0]) ===
      String(id)
    ) {

      namaKategori =
        String(data[i][1]);

      rowNumber =
        i + 1;

      break;

    }

  }


  if (!rowNumber) {

    return {
      success: false,
      message:
        'Kategori tidak ditemukan.'
    };

  }


  // Cek apakah sedang dipakai

  const barangSheet =
    ss.getSheetByName(
      SHEET_BARANG
    );


  if (
    barangSheet &&
    barangSheet.getLastRow() >= 2
  ) {

    const kategoriData =
      barangSheet
        .getRange(
          2,
          4,
          barangSheet.getLastRow() - 1,
          1
        )
        .getValues();


    const digunakan =
      kategoriData.some(
        function(row) {

          return String(row[0]) ===
            namaKategori;

        }
      );


    if (digunakan) {

      return {
        success: false,
        message:
          'Kategori tidak dapat dihapus karena masih digunakan oleh data barang.'
      };

    }

  }


  sheet.deleteRow(
    rowNumber
  );


  return {
    success: true,
    message:
      'Kategori berhasil dihapus.'
  };

}


/*************************************************
 * LOKASI
 *************************************************/

function getLokasi() {

  setupDatabase();


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        SHEET_LOKASI
      );


  if (
    sheet.getLastRow() < 2
  ) {

    return [];

  }


  const data =
    sheet
      .getRange(
        2,
        1,
        sheet.getLastRow() - 1,
        2
      )
      .getValues();


  return data
    .map(
      function(row) {

        return String(
          row[1] || ''
        ).trim();

      }
    )
    .filter(
      function(item) {

        return item !== '';

      }
    );

}


/*************************************************
 * TAMBAH BARANG
 *************************************************/

function tambahBarang(
  data
) {

  setupDatabase();


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        SHEET_BARANG
      );


  const now =
    new Date();


  const id =
    generateId('BRG');


  sheet.appendRow([

    id,

    data.kode || '',

    data.nama || '',

    data.kategori || '',

    data.merk || '',

    data.model || '',

    data.seri || '',

    data.tahun || '',

    data.sumberDana || '',

    Number(data.harga) || 0,

    data.lokasi || '',

    data.kondisi || '',

    Number(data.jumlah) || 0,

    data.keterangan || '',

    data.foto || '',

    now,

    now

  ]);


  return {
    success: true,
    message:
      'Data barang berhasil ditambahkan.',
    id: id
  };

}


/*************************************************
 * UPDATE BARANG
 *************************************************/

function updateBarang(
  data
) {

  setupDatabase();


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        SHEET_BARANG
      );


  const values =
    sheet.getDataRange()
      .getValues();


  let rowNumber = 0;


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (
      String(values[i][0]) ===
      String(data.id)
    ) {

      rowNumber =
        i + 1;

      break;

    }

  }


  if (!rowNumber) {

    return {
      success: false,
      message:
        'Data barang tidak ditemukan.'
    };

  }


  const dibuat =
    values[rowNumber - 1][15] ||
    new Date();


  sheet
    .getRange(
      rowNumber,
      1,
      1,
      17
    )
    .setValues([[
      
      data.id,

      data.kode || '',

      data.nama || '',

      data.kategori || '',

      data.merk || '',

      data.model || '',

      data.seri || '',

      data.tahun || '',

      data.sumberDana || '',

      Number(data.harga) || 0,

      data.lokasi || '',

      data.kondisi || '',

      Number(data.jumlah) || 0,

      data.keterangan || '',

      data.foto || '',

      dibuat,

      new Date()

    ]]);


  return {
    success: true,
    message:
      'Data barang berhasil diperbarui.'
  };

}


/*************************************************
 * DELETE BARANG
 *************************************************/

function deleteBarang(
  id
) {

  setupDatabase();


  const sheet =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getSheetByName(
        SHEET_BARANG
      );


  const values =
    sheet.getDataRange()
      .getValues();


  for (
    let i = 1;
    i < values.length;
    i++
  ) {

    if (
      String(values[i][0]) ===
      String(id)
    ) {

      sheet.deleteRow(
        i + 1
      );


      return {
        success: true,
        message:
          'Data barang berhasil dihapus.'
      };

    }

  }


  return {
    success: false,
    message:
      'Data barang tidak ditemukan.'
  };

}


/*************************************************
 * DASHBOARD
 *************************************************/

function getDashboard() {

  setupDatabase();

  return hitungStatistik();

}


/*************************************************
 * HITUNG STATISTIK
 *************************************************/

function hitungStatistik() {

  const data =
    ambilDataBarang();


  let totalBarang = 0;

  let jenisBarang =
    data.length;

  let barangBaik = 0;

  let barangRusakRingan = 0;

  let barangRusakBerat = 0;


  data.forEach(
    function(item) {

      const jumlah =
        Number(item.jumlah) || 0;


      totalBarang +=
        jumlah;


      if (
        item.kondisi === 'Baik'
      ) {

        barangBaik +=
          jumlah;

      }


      if (
        item.kondisi ===
        'Rusak Ringan'
      ) {

        barangRusakRingan +=
          jumlah;

      }


      if (
        item.kondisi ===
        'Rusak Berat'
      ) {

        barangRusakBerat +=
          jumlah;

      }

    }
  );


  return {

    totalBarang:
      totalBarang,

    jenisBarang:
      jenisBarang,

    barangBaik:
      barangBaik,

    barangRusak:
      barangRusakRingan +
      barangRusakBerat,

    barangRusakRingan:
      barangRusakRingan,

    barangRusakBerat:
      barangRusakBerat

  };

}