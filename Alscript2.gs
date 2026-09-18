const SPREADSHEET_ID = "1Nv49FeNTl-fCjkeLjDeIRVA2cQJomCk0oN-6eAGV6tM";
const SHEET_SLT = "SLT";
const SHEET_SPDT = "SPDT";
const SHEET_MASTER = "MASTER";
const ADMIN_PASSWORD = "0800";
const MASTER_CACHE_KEY = "MASTER_BARCODE_CACHE";
const MASTER_CACHE_TIME = 21600;

function jsonResponse(data){
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function textResponse(text){
  return ContentService.createTextOutput(String(text));
}

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

function getTargetSheet(ss,mode){
  return String(mode||"").toLowerCase()==="opname"
    ? ss.getSheetByName(SHEET_SPDT)
    : ss.getSheetByName(SHEET_SLT);
}

function normalizeBarcode(value){
  return String(value==null?"":"")
    .trim()
    .replace(/\s+/g,"")
    .toLowerCase();
}

function doPost(e){
  try{
    if(!e || !e.postData || !e.postData.contents){
      return textResponse("Error: Data POST kosong");
    }

    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

    if(data.action==="getAllBarcode"){
      return jsonResponse(getMasterData());
    }

    if(data.action==="getHistory"){
      const sheetHistory = getTargetSheet(ss,data.mode);
      if(!sheetHistory) return jsonResponse([]);

      const lastRow = sheetHistory.getLastRow();
      if(lastRow<=1) return jsonResponse([]);

      const startRow = Math.max(2,lastRow-49);
      const numRows = lastRow-startRow+1;
      const historyData = sheetHistory
        .getRange(startRow,1,numRows,3)
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

        if(
          !timestamp &&
          originalBarcode==="" &&
          (qty==="" || qty==null)
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
          qty:qty
        });
      }

      return jsonResponse(result);
    }

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

      const rowValues =
        sheetDelete.getRange(row,1,1,3).getValues()[0];

      const actualBarcode = String(rowValues[1]||"").trim();
      const actualQty = Number(rowValues[2]);
      const requestedBarcode =
        String(data.barcode||"").trim();
      const requestedQty = Number(data.qty);

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

      sheetDelete
        .getRange(row,1,1,3)
        .clearContent();

      return textResponse("Delete berhasil");
    }

    if(data.action==="reset"){
      if(data.password!==ADMIN_PASSWORD){
        return textResponse("Password salah");
      }

      const sheetReset = getTargetSheet(ss,data.mode);
      if(!sheetReset){
        return textResponse("Error: Sheet reset tidak ditemukan");
      }

      const lock = LockService.getScriptLock();

      try{
        lock.waitLock(10000);

        const lastRow = sheetReset.getLastRow();

        if(lastRow>1){
          sheetReset
            .getRange(2,1,lastRow-1,3)
            .clearContent();
        }

        return textResponse("Reset berhasil");

      }finally{
        try{
          lock.releaseLock();
        }catch(err){}
      }
    }

    if(data.action==="saveData"){
      const sheet = getTargetSheet(ss,data.mode);
      if(!sheet){
        return textResponse("Error: Sheet tujuan tidak ditemukan");
      }

      let qtyText =
        String(data.qty==null?"":data.qty)
        .trim()
        .replace(/\s/g,"");

      if(qtyText.includes(",") && qtyText.includes(".")){
        if(qtyText.lastIndexOf(",")>qtyText.lastIndexOf(".")){
          qtyText = qtyText
            .replace(/\./g,"")
            .replace(",",".");
        }else{
          qtyText = qtyText.replace(/,/g,"");
        }
      }else if(qtyText.includes(",")){
        qtyText = qtyText.replace(",",".");
      }

      const qty = Number(qtyText);

      if(!Number.isFinite(qty)){
        return textResponse("Error: Qty tidak valid");
      }

      if(qty<=0){
        return textResponse("Error: Qty harus lebih dari 0");
      }

      const barcode = String(data.barcode||"").trim();

      if(barcode===""){
        return textResponse("Error: Barcode kosong");
      }

      const lock = LockService.getScriptLock();

      try{
        lock.waitLock(10000);

        const nextRow = Math.max(2,sheet.getLastRow()+1);
        const savedAt = new Date();

        sheet
          .getRange(nextRow,1,1,3)
          .setValues([
            [savedAt,barcode,qty]
          ]);

        return jsonResponse({
          status:"success",
          success:true,
          row:nextRow,
          timestamp:savedAt.toISOString()
        });

      }finally{
        try{
          lock.releaseLock();
        }catch(err){}
      }
    }

    return textResponse("Error: Action tidak dikenal");

  }catch(err){
    console.error(err);
    return textResponse("Error: " + err.message);
  }
}

function doGet(e){
  return textResponse("THE BUTCHER API ONLINE");
}

function clearMasterCache(){
  CacheService
    .getScriptCache()
    .remove(MASTER_CACHE_KEY);

  return "MASTER CACHE CLEARED";
}
