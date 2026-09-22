const SPREADSHEET_ID = "1Nv49FeNTl-fCjkeLjDeIRVA2cQJomCk0oN-6eAGV6tM";
const SHEET_SLT = "SLT";
const SHEET_SPDT = "SPDT";
const SHEET_MASTER = "MASTER";
const ADMIN_PASSWORD = "0800";

const MASTER_CACHE_KEY = "MASTER_BARCODE_CACHE";
const MASTER_CACHE_TIME = 21600;

// Kolom data:
// A = Timestamp
// B = Barcode
// C = Qty
// D = Transaction ID
const DATA_COLUMNS = 4;


/* =========================================================
   RESPONSE
========================================================= */

function jsonResponse(data){
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function textResponse(text){
  return ContentService.createTextOutput(String(text));
}


/* =========================================================
   MASTER CACHE
========================================================= */

function getMasterData(){
  const cache = CacheService.getScriptCache();
  const cached = cache.get(MASTER_CACHE_KEY);

  if(cached){
    try{
      return JSON.parse(cached);
    }catch(err){
      console.log("Cache MASTER rusak: " + err);
    }
  }

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SHEET_MASTER);

  if(!sheet) return [];

  const lastRow = sheet.getLastRow();

  if(lastRow < 2) return [];

  const values = sheet.getRange(2,2,lastRow-1,2).getValues();
  const result = [];

  for(let i=0;i<values.length;i++){
    const barcode = values[i][0];
    const nama = values[i][1];

    if(barcode==="" || barcode==null) continue;

    result.push({
      barcode:String(barcode).trim(),
      nama:String(nama||"").trim()
    });
  }

  try{
    cache.put(
      MASTER_CACHE_KEY,
      JSON.stringify(result),
      MASTER_CACHE_TIME
    );
  }catch(err){
    console.log("Cache gagal: " + err);
  }

  return result;
}


/* =========================================================
   SHEET
========================================================= */

function getTargetSheet(ss,mode){
  return String(mode||"").toLowerCase()==="opname"
    ? ss.getSheetByName(SHEET_SPDT)
    : ss.getSheetByName(SHEET_SLT);
}


/* =========================================================
   NORMALIZER
========================================================= */

function normalizeBarcode(value){
  return String(value==null?"":value)
    .trim()
    .replace(/\s+/g,"")
    .toLowerCase();
}

function normalizeTransactionId(value){
  return String(value==null?"":value)
    .trim();
}


/* =========================================================
   FIND TRANSACTION
   Mencari Transaction ID di kolom D.
========================================================= */

function findTransactionRow(sheet,transactionId,lastRow){
  const tx=normalizeTransactionId(transactionId);

  if(!tx || lastRow<2) return 0;

  /*
    TextFinder menghindari pembacaan seluruh kolom D
    ke memory JavaScript.
  */
  const range=sheet.getRange(2,4,lastRow-1,1);

  const found=range
    .createTextFinder(tx)
    .matchEntireCell(true)
    .matchCase(true)
    .findNext();

  return found ? found.getRow() : 0;
}


/* =========================================================
   DO POST
========================================================= */

function doPost(e){
  try{
    if(!e || !e.postData || !e.postData.contents){
      return textResponse("Error: Data POST kosong");
    }

    const data = JSON.parse(e.postData.contents);

    /*
      getAllBarcode bisa langsung memakai cache MASTER.
      Tidak perlu membuka Spreadsheet lebih dulu.
    */
    if(data.action==="getAllBarcode"){
      return jsonResponse(getMasterData());
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);


    /* =====================================================
       GET HISTORY
    ===================================================== */

    if(data.action==="getHistory"){
      const sheetHistory = getTargetSheet(ss,data.mode);

      if(!sheetHistory) return jsonResponse([]);

      const lastRow = sheetHistory.getLastRow();

      if(lastRow<=1) return jsonResponse([]);

      const startRow = Math.max(2,lastRow-49);
      const numRows = lastRow-startRow+1;

      /*
        Sekarang membaca A:D.
        D berisi Transaction ID.
      */
      const historyData =
        sheetHistory
          .getRange(startRow,1,numRows,DATA_COLUMNS)
          .getValues();

      const masterMap = {};

      getMasterData().forEach(item=>{
        masterMap[normalizeBarcode(item.barcode)] = item.nama;
      });

      const result = [];

      for(let i=historyData.length-1;i>=0;i--){
        const timestamp = historyData[i][0];
        const originalBarcode = String(historyData[i][1]||"").trim();
        const qty = historyData[i][2];
        const transactionId =
          String(historyData[i][3]||"").trim();

        if(
          !timestamp &&
          originalBarcode==="" &&
          (qty==="" || qty==null) &&
          transactionId===""
        ) continue;

        const actualRow = startRow+i;
        let searchBarcode = originalBarcode;

        if(
          originalBarcode.startsWith("20") &&
          originalBarcode.length===13 &&
          /^\d+$/.test(originalBarcode)
        ){
          searchBarcode = originalBarcode.substring(2,7);
        }

        result.push({
          row:actualRow,

          timestamp:
            timestamp instanceof Date
              ? timestamp.toISOString()
              : String(timestamp||""),

          barcode:originalBarcode,
          searchBarcode:searchBarcode,

          nama:
            masterMap[normalizeBarcode(searchBarcode)]||"",

          qty:qty,

          transactionId:transactionId
        });
      }

      return jsonResponse(result);
    }


    /* =====================================================
       DELETE DATA
    ===================================================== */

    if(data.action==="deleteData"){
      if(data.password!==ADMIN_PASSWORD){
        return textResponse("Password salah");
      }

      const sheetDelete = getTargetSheet(ss,data.mode);

      if(!sheetDelete){
        return textResponse("Error: Sheet delete tidak ditemukan");
      }

      const lastRow = sheetDelete.getLastRow();

      if(lastRow<=1){
        return textResponse("Error: Tidak ada data");
      }

      const row = Number(data.row);

      if(
        !Number.isInteger(row) ||
        row<2 ||
        row>lastRow
      ){
        return textResponse("Error: Baris data tidak valid");
      }

      /*
        Baca A:D.
      */
      const rowValues =
        sheetDelete.getRange(row,1,1,DATA_COLUMNS).getValues()[0];

      const actualBarcode =
        String(rowValues[1]||"").trim();

      const actualQty =
        Number(rowValues[2]);

      const actualTransactionId =
        String(rowValues[3]||"").trim();

      const requestedBarcode =
        String(data.barcode||"").trim();

      const requestedQty =
        Number(data.qty);

      const requestedTransactionId =
        normalizeTransactionId(data.transactionId);


      /*
        Kalau HTML mengirim Transaction ID,
        pastikan row yang dihapus benar-benar transaksi yang sama.
      */
      if(
        requestedTransactionId &&
        actualTransactionId!==requestedTransactionId
      ){
        return textResponse(
          "Error: Transaction ID data sudah berubah. Silakan refresh riwayat."
        );
      }


      if(
        requestedBarcode &&
        normalizeBarcode(actualBarcode)!==
        normalizeBarcode(requestedBarcode)
      ){
        return textResponse(
          "Error: Data sudah berubah. Silakan refresh riwayat."
        );
      }


      if(
        Number.isFinite(requestedQty) &&
        Number.isFinite(actualQty) &&
        Math.abs(actualQty-requestedQty)>0.0000001
      ){
        return textResponse(
          "Error: Quantity data sudah berubah. Silakan refresh riwayat."
        );
      }


      /*
        Hapus A:D agar Transaction ID ikut bersih.
      */
      sheetDelete
        .getRange(row,1,1,DATA_COLUMNS)
        .clearContent();

      return textResponse("Delete berhasil");
    }


    /* =====================================================
       RESET
    ===================================================== */

    if(data.action==="reset"){
      if(data.password!==ADMIN_PASSWORD){
        return textResponse("Password salah");
      }

      const sheetReset = getTargetSheet(ss,data.mode);

      if(!sheetReset){
        return textResponse(
          "Error: Sheet reset tidak ditemukan"
        );
      }

      const lock = LockService.getScriptLock();

      try{
        lock.waitLock(10000);

        const lastRow = sheetReset.getLastRow();

        if(lastRow>1){
          /*
            A:D dibersihkan.
            Transaction ID ikut di-reset.
          */
          sheetReset
            .getRange(
              2,
              1,
              lastRow-1,
              DATA_COLUMNS
            )
            .clearContent();
        }

        return textResponse("Reset berhasil");

      }finally{
        try{
          lock.releaseLock();
        }catch(err){}
      }
    }


    /* =====================================================
       SAVE DATA
       1 PRODUK = 1 REQUEST = LANGSUNG SIMPAN
    ===================================================== */

    if(data.action==="saveData"){

      const sheet = getTargetSheet(ss,data.mode);

      if(!sheet){
        return textResponse(
          "Error: Sheet tujuan tidak ditemukan"
        );
      }


      /*
        Transaction ID wajib untuk sistem anti-duplikat baru.
      */
      const transactionId =
        normalizeTransactionId(data.transactionId);

      if(!transactionId){
        return jsonResponse({
          status:"error",
          success:false,
          error:"transaction_id_required",
          message:"Transaction ID wajib dikirim."
        });
      }


      let qtyText =
        String(data.qty==null?"":data.qty)
        .trim()
        .replace(/\s/g,"");


      if(
        qtyText.includes(",") &&
        qtyText.includes(".")
      ){
        if(
          qtyText.lastIndexOf(",")>
          qtyText.lastIndexOf(".")
        ){
          qtyText =
            qtyText
              .replace(/\./g,"")
              .replace(",",".");
        }else{
          qtyText =
            qtyText.replace(/,/g,"");
        }

      }else if(qtyText.includes(",")){
        qtyText =
          qtyText.replace(",",".");
      }


      const qty = Number(qtyText);

      if(!Number.isFinite(qty)){
        return textResponse(
          "Error: Qty tidak valid"
        );
      }

      if(qty<=0){
        return textResponse(
          "Error: Qty harus lebih dari 0"
        );
      }


      const barcode =
        String(data.barcode||"").trim();

      if(barcode===""){
        return textResponse(
          "Error: Barcode kosong"
        );
      }


      /*
        Lock tetap digunakan supaya dua HP yang mengirim
        bersamaan tidak mendapatkan row yang sama.
      */
      const lock = LockService.getScriptLock();

      try{
        lock.waitLock(10000);


        /*
          CEK ANTI DUPLIKAT
          Kalau Transaction ID sudah ada, jangan tulis
          baris baru.
        */
        const lastRowBefore =
          sheet.getLastRow();

        const existingRow =
          findTransactionRow(
            sheet,
            transactionId,
            lastRowBefore
          );

        if(existingRow){

          const existingValues =
            sheet
              .getRange(
                existingRow,
                1,
                1,
                DATA_COLUMNS
              )
              .getValues()[0];

          const existingTimestamp =
            existingValues[0];

          const existingBarcode =
            String(existingValues[1]||"").trim();

          const existingQty =
            Number(existingValues[2]);

          return jsonResponse({
            status:"already_saved",
            success:true,
            alreadySaved:true,
            duplicate:false,
            row:existingRow,

            timestamp:
              existingTimestamp instanceof Date
                ? existingTimestamp.toISOString()
                : String(existingTimestamp||""),

            barcode:existingBarcode,
            qty:existingQty,

            transactionId:transactionId
          });
        }


        /*
          TRANSAKSI BARU
        */
        const nextRow =
          Math.max(
            2,
            sheet.getLastRow()+1
          );

        const savedAt =
          new Date();


        /*
          Tulis A:D sekaligus.
        */
        sheet
          .getRange(
            nextRow,
            1,
            1,
            DATA_COLUMNS
          )
          .setValues([
            [
              savedAt,
              barcode,
              qty,
              transactionId
            ]
          ]);


        return jsonResponse({
          status:"success",
          success:true,
          alreadySaved:false,
          row:nextRow,
          timestamp:savedAt.toISOString(),
          barcode:barcode,
          qty:qty,
          transactionId:transactionId
        });


      }finally{
        try{
          lock.releaseLock();
        }catch(err){}
      }
    }


    return textResponse(
      "Error: Action tidak dikenal"
    );

  }catch(err){

    console.error(err);

    return textResponse(
      "Error: " + err.message
    );
  }
}


/* =========================================================
   DO GET
========================================================= */

function doGet(e){
  return textResponse(
    "THE BUTCHER API ONLINE"
  );
}


/* =========================================================
   CLEAR MASTER CACHE
========================================================= */

function clearMasterCache(){
  CacheService
    .getScriptCache()
    .remove(MASTER_CACHE_KEY);

  return "MASTER CACHE CLEARED";
}
