// CONFIGURATION
const LEADS_SHEET = "Leads";
const AIRTABLE_BASE_ID = "appl6e4pwMTM3zmm2";
const AIRTABLE_TABLE_ID = "tbl1hIorYHhwyyBsr"; 
const AIRTABLE_API_KEY = "your_api_key";

function syncLeadsByLastUpdated() {
  const ss = SpreadsheetApp.getActive();
  const leadsSheet = ss.getSheetByName(LEADS_SHEET);
  const data = leadsSheet.getDataRange().getValues();
  const headers = data.shift();
  
  const CREATED_COL = headers.indexOf("Created Date");
  const UPDATED_COL = headers.indexOf("Last Updated");
  const STAGE_COL = headers.indexOf("Pipeline Stage");

  let buckets = {};

  data.forEach((row) => {
    const createdDate = new Date(row[CREATED_COL]);
    const updatedDate = new Date(row[UPDATED_COL]);
    const stage = (row[STAGE_COL] || "").toLowerCase();

    // 1. COUNT THE LEAD (Based on Created Date)
    if (!isNaN(updatedDate.getTime())) {
      const bId = getBucketId(updatedDate);
      initBucket(buckets, bId);
      buckets[bId].leads++;
    }

    // 2. COUNT THE ACTION (Based on Last Updated Date)
    // Only count actions if the stage is not "New Lead"
    if (!isNaN(updatedDate.getTime()) && !stage.includes("new lead")) {
      const bId = getBucketId(updatedDate);
      initBucket(buckets, bId);
      const b = buckets[bId];

      if (stage.includes("hot")) b.res++;
      if (stage.includes("booked") || stage.includes("buddy")) b.book++;
      if (stage.includes("showed") || stage.includes("didn't purchase")) b.show++;
      if (stage.includes("purchased") || stage.includes("won") || stage.includes("sale")) b.sale++;
    }
  });

  // 3. SYNC TO AIRTABLE
  for (let id in buckets) {
    const b = buckets[id];
    const fields = {
      "Month": b.month,
      "From": b.from,
      "To": b.to,
      "Leads": b.leads,
      "Responses": b.res,
      "Bookings": b.book,
      "Shows": b.show,
      "Sales": b.sale,
      "Response Rate": calcRate(b.res, b.leads),
      "Booking Rate": calcRate(b.book, b.leads),
      "Show Up Rate": calcRate(b.show, b.book),
      "Closing Rate": calcRate(b.sale, b.show)
    };
    upsertToAirtable(fields);
  }
}

// Helper to initialize a bucket object
function initBucket(buckets, id) {
  if (!buckets[id]) {
    const parts = id.split("|");
    buckets[id] = { month: parts[0], from: parts[1], to: parts[2], leads: 0, res: 0, book: 0, show: 0, sale: 0 };
  }
}

// Helper to get Bucket ID from a Date
function getBucketId(date) {
  const year = date.getFullYear();
  const monthIdx = date.getMonth();
  const monthName = Utilities.formatDate(date, Session.getScriptTimeZone(), "MMMM");
  const day = date.getDate();
  
  let f, t;
  if (day <= 7) { f = fDate(year, monthIdx, 1); t = fDate(year, monthIdx, 7); }
  else if (day <= 15) { f = fDate(year, monthIdx, 8); t = fDate(year, monthIdx, 15); }
  else if (day <= 23) { f = fDate(year, monthIdx, 16); t = fDate(year, monthIdx, 23); }
  else { f = fDate(year, monthIdx, 24); t = fDate(year, monthIdx, new Date(year, monthIdx + 1, 0).getDate()); }
  
  return `${monthName}|${f}|${t}`;
}

function upsertToAirtable(fields) {
  const url = `https://api.airtable.com/v0/${AIRTABLE_BASE_ID}/${AIRTABLE_TABLE_ID}`;
  // Precise search using DATETIME_FORMAT to ensure Airtable finds the record
  const formula = `AND({Month}='${fields.Month}', DATETIME_FORMAT({From},'YYYY-MM-DD')='${fields.From}', DATETIME_FORMAT({To},'YYYY-MM-DD')='${fields.To}')`;
  const searchUrl = `${url}?filterByFormula=${encodeURIComponent(formula)}`;
  
  const response = UrlFetchApp.fetch(searchUrl, {
    headers: { "Authorization": "Bearer " + AIRTABLE_API_KEY },
    muteHttpExceptions: true
  });
  
  const records = JSON.parse(response.getContentText()).records || [];
  let method = "POST", finalUrl = url, payload = { records: [{ fields: fields }] };

  if (records.length > 0) {
    method = "PATCH";
    finalUrl = `${url}/${records[0].id}`;
    payload = { fields: fields };
  }

  UrlFetchApp.fetch(finalUrl, {
    method: method,
    headers: { "Authorization": "Bearer " + AIRTABLE_API_KEY, "Content-Type": "application/json" },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function fDate(y, m, d) { return y + "-" + ("0" + (m + 1)).slice(-2) + "-" + ("0" + d).slice(-2); }
function calcRate(num, den) { return den === 0 ? 0 : Number((num / den).toFixed(4)); }