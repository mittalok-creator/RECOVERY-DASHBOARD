function initApp(DATA){
/* ---------- Recovery Dashboard (branch portal): who's logged in ----------
   js/login.js writes this on a successful Sol ID + PIN login (the login
   PIN itself IS the Sol ID -- Alok's own explicit choice, "simple login,"
   not real access control). Read as a plain string; every consumer below
   resolves it to that dataset's own actual branch-name spelling rather
   than assuming a canonical one, since C.SOL_DESC/KC.BRANCH can differ
   slightly from BRANCH_LIST's spelling. */
function loggedInSolId(){ try{ return sessionStorage.getItem('upgb-sol-id'); }catch(e){ return null; } }
/* ---------- NPA column map ---------- */
const C = {
  HELPER:0, PROVISION:1, MULTI:2, SOL_ID:3, SOL_DESC:4, CUST_ID:5, ACCT_NO:6,
  NAME:7, ADDR:8, PHONE:9, AADHAR:10, PAN:11, OPN_DT:12, SCHEME:13, SANCT_DT:14,
  SANCT_LIM:15, OUTBAL:16, UNCHG:17, URI:18, ASSET:19, USER_CLASS_DT:20,
  SYS_SUBCLASS:21, SYS_CLASS_DT:22, NPA_DT:23, SB_ACCT:24, SB_BAL:25, REGION:26
};
const NPA_COLUMN_COUNT = 27;
const PROV_RATES = {SUB_STD:.10, DA1:.20, DA2:.30, DA3:1, LOSS:1};
/* Minimum settlement % of O/S Balance under Lok Adalat, by Asset Code, as
   on 30-06-2026 (Alok's own circular). Substandard accounts aren't in the
   circular at all -- Lok Adalat OTS doesn't apply to them -- so they're
   deliberately absent from this table rather than defaulting to some
   guessed rate; lokAdalatMin() below returns an explicit "not eligible"
   result for SUB_STD instead of falling through to undefined. This app
   has no separate "PWO" Asset Code -- every loss-category account is
   just LOSS -- so a PWO account gets the same 40% as any other LOSS
   account here. */
const LOK_ADALAT_RATES = {DA1:.80, DA2:.70, DA3:.50, LOSS:.40};
function lokAdalatMin(s){
  if(s.assetCode==='SUB_STD') return {eligible:false};
  const rate = LOK_ADALAT_RATES[s.assetCode];
  if(rate===undefined || s.os==='') return null;
  return {eligible:true, amount:s.os*rate, pct:rate};
}

/* ---------- Lazy-loaded vendor libraries ----------
   msal-browser, xlsx, exceljs, html2canvas and jsPDF used to be plain
   blocking <script src> tags in index.html -- ~2.76MB combined, parsed
   and executed on every single page load before app.js even started,
   even though the huge majority of visits are just viewing NPA data and
   never touch OneDrive login, Excel import/export, or the WhatsApp PDF
   share button. That was the single biggest fixable cause of "the app
   is slow/blank to open" reported across several branch computers on
   different networks (2026-09-17) -- fixed by injecting each library's
   <script> tag on first actual use instead, cached so a second use on
   the same page never re-fetches it. Version query strings here must
   stay in sync with the versions these libraries were last bumped to. */
const __vendorScriptPromises = {};
function loadVendorScript(src){
  if(!__vendorScriptPromises[src]){
    __vendorScriptPromises[src] = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = src;
      s.onload = () => resolve();
      s.onerror = () => { delete __vendorScriptPromises[src]; reject(new Error('Failed to load '+src)); };
      document.head.appendChild(s);
    });
  }
  return __vendorScriptPromises[src];
}
function ensureXLSX(){ return typeof XLSX!=='undefined' ? Promise.resolve() : loadVendorScript('js/vendor/xlsx.full.min.js?v=20260915h'); }
function ensureExcelJS(){ return typeof ExcelJS!=='undefined' ? Promise.resolve() : loadVendorScript('js/vendor/exceljs.min.js?v=20260912a'); }
function ensureHtml2Canvas(){ return typeof html2canvas!=='undefined' ? Promise.resolve() : loadVendorScript('js/vendor/html2canvas.min.js?v=20260912a'); }
function ensureJsPDF(){ return window.jspdf ? Promise.resolve() : loadVendorScript('js/vendor/jspdf.umd.min.js?v=20260915h'); }
function ensureMsal(){ return typeof msal!=='undefined' ? Promise.resolve() : loadVendorScript('js/vendor/msal-browser.min.js?v=20260912a'); }

/* ---------- Build indexes once ---------- */
const npaByAcct = new Map();
const npaByHelper = new Map();
const byCustId = new Map();
DATA.npa.rows.forEach(r=>{
  if(r[C.ACCT_NO]!=='') npaByAcct.set(String(r[C.ACCT_NO]), r);
  if(r[C.HELPER]!=='') npaByHelper.set(String(r[C.HELPER]), r);
  const cid = String(r[C.CUST_ID]);
  if(cid && !byCustId.has(cid)) byCustId.set(cid, r);
});
const oldOtsByAcct = new Map();
DATA.oldots.rows.forEach(r=>{
  if(r[0]!=='' && !oldOtsByAcct.has(String(r[0]))) oldOtsByAcct.set(String(r[0]), {date:r[1], amount:r[2]});
});
/* Branch-wise total advance, uploaded separately from the daily NPA file
   (see handleBranchAdvUpload) -- lets the Dashboard show NPA % (NPA
   outstanding / total advance) per branch. Persisted through Publish, but
   not reset/carried-forward on a daily NPA update since it changes on its
   own, much slower schedule. */
DATA.branchAdvances = DATA.branchAdvances || {};
/* Branch Manager / Recovery Officer contacts -- keyed by Sol ID (string),
   uploaded via Update Data -> Branch Contacts. Same "own slow-moving
   schedule, not reset on a daily NPA update" treatment as branchAdvances
   above. */
DATA.branchContacts = DATA.branchContacts || {};
/* Special Note -- keyed by Account No. (string), added one at a time by
   the Admin via Update Data -> Special Note (see wireChrome's
   specialNoteSaveBtn wiring). {note, updatedAt, updatedBy} per account.
   Same "own slow-moving schedule, not reset on a daily NPA update"
   treatment as branchAdvances/branchContacts above -- a note stays on an
   account across daily uploads until the Admin removes it. Shown to
   every viewer (not just Admin) as a banner at the top of that account's
   Loan Detail screen -- see drawDetailBody. */
DATA.specialNotes = DATA.specialNotes || {};

/* Lok Adalat -- proposed-OTS accounts where token money has already been
   collected ahead of a Lok Adalat (Alok's request, 2026-09-11, following
   up on an earlier throwaway version of this same idea). Keyed by Account
   No. (string): { date, ots, token, remark }. Uploaded via Settings ->
   Update Data -> "Lok Adalat" (see handleLokAdalatUpload), one full
   replace per upload -- "jo last database upload hoga wahi accounts show
   hon", so this is never merged with a previous upload, only the latest
   one's accounts are ever shown. Same "own slow-moving schedule, not
   reset on a daily NPA update" treatment as branchAdvances/branchContacts/
   specialNotes above, and part of the publish payload like they are. */
DATA.lokAdalat = DATA.lokAdalat || {};

/* Interest Reversal master list -- keyed by Account No. (string) -> amount
   (₹). Uploaded via Settings -> Update Data -> "Interest Reversal (master
   list)" (see handleInterestReversalMasterUpload). Head Office's own daily
   NPA export carries an "Intt Rev" column, but it's proven too unreliable
   to trust alone -- it silently went blank for every account the very
   upload after Alok's file was first merged in on 2026-09-08, since there
   was nowhere for that figure to persist. This master list is now the
   authoritative source instead: applyInterestReversalMaster() overrides
   Interest Reversal with this list's figure, by Account No., on every
   future daily NPA upload (see processDailyParsed()). Same "own
   slow-moving schedule, not reset on a daily NPA update" treatment as
   branchAdvances/branchContacts/specialNotes/lokAdalat above, and part of
   the publish payload like they are -- so it only ever needs re-uploading
   when Alok has an updated/corrected list, never alongside a routine
   daily NPA file. */
DATA.interestReversalMaster = DATA.interestReversalMaster || {};

/* Address list -- two maps built from the SAME upload (see
   handleAddressListUpload/buildAddressListMap): DATA.addressList keyed by
   Account No. (string) -> address, and DATA.addressListByCustomer keyed by
   Customer ID (string) -> address. Uploaded via Settings -> Update Data ->
   "Address List". Both are app-wide rather than scoped to one tool: (1) a
   final fallback inside mergeCustomerDetails() -- if an account's address
   is still blank after Customer Master + carry-forward, Account No. is
   checked first, then Customer ID; (2) exposed read-only to same-origin
   Utility Hub iframe tools via window.UPGB_getAddressList()/
   window.UPGB_getAddressListByCustomer() (see below) so tools/branch-
   split.html can enrich its own output without re-implementing PIN
   decryption. Alok's own request (2026-09-22): "1 list main upload karun
   and entire app main wo use ho jahan bhi address available nahi hai" --
   one list, used everywhere, not just one utility.
   Customer ID support added 2026-09-23: Alok's real source list is
   naturally keyed by Customer ID, not Account No. -- one address per
   customer, not re-typed per account -- and buildAddressListMap() used to
   flatly REJECT a file with no Account Number column at all ("Could not
   find an Account Number column"), so his upload never succeeded, which
   is why Branch Split had nothing to fetch AND Publish stayed disabled
   (its own enabling only happens inside a successful upload handler).
   Both key types are now accepted, together or separately, with Account
   No. taking priority wherever a row supplies both (more precise: a
   customer's registered address can differ from where a specific account
   was opened). Same "own slow-moving schedule, not reset on a daily NPA
   update" treatment as branchAdvances/branchContacts/
   interestReversalMaster above, and part of the publish payload like they
   are. */
DATA.addressList = DATA.addressList || {};
DATA.addressListByCustomer = DATA.addressListByCustomer || {};
// Re-key both maps through normId() on every load, not only at upload
// time -- self-heals a list already published under the old code (before
// normId() existed, 2026-09-23) whose keys carry a leading zero or a
// stray ".0" that would otherwise keep failing to match even after this
// fix ships, until Alok happens to re-upload and re-publish the exact
// same file again. Cheap (runs once per page load, not per lookup).
DATA.addressList = renormalizeIdMap(DATA.addressList);
DATA.addressListByCustomer = renormalizeIdMap(DATA.addressListByCustomer);
/* Read-only accessors for same-origin Utility Hub iframe tools (currently
   just tools/branch-split.html) that need the live, already-decrypted
   address maps without re-implementing PIN decryption themselves. Each
   returns a fresh shallow copy so a tool can't mutate the parent's live
   DATA by reference. Deliberately NOT the general pattern for tool <-> app
   data sharing -- every other Utility Hub tool stays fully self-contained/
   independent; these exist only because address data must be genuinely
   single-sourced per Alok's own "1 list ... entire app main" requirement
   (2026-09-22). */
window.UPGB_getAddressList = function(){ return Object.assign({}, DATA.addressList||{}); };
window.UPGB_getAddressListByCustomer = function(){ return Object.assign({}, DATA.addressListByCustomer||{}); };

/* ---------- Date helpers (NPA dates are raw Excel serials) ---------- */
const XL_EPOCH = new Date(1899,11,30);
function excelSerialToDate(n){ return new Date(XL_EPOCH.getTime() + n*86400000); }
function dateToExcelSerial(d){ return Math.round((d.getTime()-XL_EPOCH.getTime())/86400000); }
// Strict on purpose: this app only ever produces/consumes date strings as
// DD-MM-YYYY (see CLAUDE.md) or raw Excel serial numbers. The previous
// version split on '-' and fell back to parseFloat() for anything that
// didn't match -- parseFloat() parses a *leading* numeric prefix, not the
// whole string, so a stray-whitespace date ("  22-07-2022  "), an ISO
// date ("2022-07-22"), or a US-format date ("07-22-2022") all silently
// parsed as a plausible-looking but WRONG date (an Excel serial number
// near 1900, or a rolled-over invalid month) instead of failing loudly.
// Every date is now validated end-to-end -- including a round-trip check
// that rejects invalid calendar dates a naive constructor would otherwise
// silently roll over (e.g. day 31 in a 30-day month) -- and returns null
// rather than guessing when the input doesn't match exactly.
function toDate(v){
  if(v===''||v===null||v===undefined) return null;
  if(v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if(typeof v==='number') return isFinite(v) ? excelSerialToDate(v) : null;
  if(typeof v==='string'){
    const s = v.trim();
    const m = /^(\d{1,2})-(\d{1,2})-(\d{4})$/.exec(s);
    if(m){
      const day = +m[1], month = +m[2], year = +m[3];
      if(month<1 || month>12 || day<1 || day>31) return null;
      const d = new Date(year, month-1, day);
      return (d.getFullYear()===year && d.getMonth()===month-1 && d.getDate()===day) ? d : null;
    }
    if(/^-?\d+(\.\d+)?$/.test(s)){
      const n = parseFloat(s);
      return isFinite(n) ? excelSerialToDate(n) : null;
    }
  }
  return null;
}
function endOfMonth(d){ return new Date(d.getFullYear(), d.getMonth()+1, 0); }
function sameDate(a,b){ return a.getFullYear()===b.getFullYear() && a.getMonth()===b.getMonth() && a.getDate()===b.getDate(); }
function daysBetween(a,b){ return Math.round((a-b)/86400000); }
function fmtDate(d){ if(!d) return '—'; return String(d.getDate()).padStart(2,'0')+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+d.getFullYear(); }
/* Date part always goes through fmtDate() (DD-MM-YYYY, never locale-
   dependent) -- only the time-of-day portion uses toLocaleTimeString,
   since that carries no date-format ambiguity. */
function fmtDateTime(d){ if(!d) return ''; return fmtDate(d)+', '+d.toLocaleTimeString('en-IN',{hour:'2-digit',minute:'2-digit'}); }
function fmtINR(n){ if(n===''||n===null||n===undefined||isNaN(n)) return '—'; return '₹'+Number(n).toLocaleString('en-IN',{maximumFractionDigits:2}); }
function fmtCr(n){
  if(n===''||n===null||n===undefined||isNaN(n)) return '—';
  const abs = Math.abs(n);
  if(abs>=1e7) return '₹'+(n/1e7).toFixed(2)+' Cr';
  if(abs>=1e5) return '₹'+(n/1e5).toFixed(2)+' L';
  return '₹'+Number(n).toLocaleString('en-IN',{minimumFractionDigits:2,maximumFractionDigits:2});
}
function esc(s){ return (s===null||s===undefined)?'':String(s).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c])); }
/* Shared "no rows to show" placeholder for every filterable table body in
   the app (KCC Overdue/PNPA branch summaries, account-list modals, etc.) --
   one consistent icon + message instead of each call site hand-rolling its
   own bare, colorless text. */
function emptyStateRowHtml(colspan, msg){
  return `<tr><td colspan="${colspan}" class="table-empty-row"><span class="table-empty-ic" aria-hidden="true"><svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg></span>${esc(msg)}</td></tr>`;
}

/* Two different print jobs on this one page want two different @page
   sizes (OTS Calculator: portrait A4; Daily NPA Projection grid: landscape
   A4) -- but @page has no selector to scope it by view, so two static
   @page rules in the stylesheet just fight over the "size" property and
   whichever is later in source order silently wins for BOTH print jobs.
   Swapping a single <style> tag's @page rule right before each print call
   keeps only one @page declaration in the document at any moment, so each
   button reliably gets its own layout regardless of the other's CSS. */
function printWithPageSize(pageCss){
  let el = document.getElementById('dynamicPrintPage');
  if(!el){ el = document.createElement('style'); el.id = 'dynamicPrintPage'; document.head.appendChild(el); }
  el.textContent = `@page{${pageCss}}`;
  window.print();
}
function printOtsSheet(){ printWithPageSize('size:A4;margin:12mm'); }
window.printOtsSheet = printOtsSheet;

/* Slices a tall rendered canvas into as many A4 pages as it needs and
   returns the finished PDF as a Blob. html2canvas gives back one single
   image regardless of how many accounts are linked (the print sheet's
   table just grows taller with more accounts), so a settlement with 3-4
   linked accounts can easily run past one page -- this crops the source
   canvas into page-height strips rather than squashing everything onto
   one page and making it unreadable. */
function canvasToPdfBlob(canvas){
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit:'pt', format:'a4', orientation:'portrait', compress:true });
  const margin = 24;
  const usableW = doc.internal.pageSize.getWidth() - margin*2;
  const usableH = doc.internal.pageSize.getHeight() - margin*2;
  const scale = usableW / canvas.width; // canvas px -> PDF pt
  const pxPerPage = Math.floor(usableH / scale);
  // addImage(canvas) with the live HTMLCanvasElement embeds it essentially
  // uncompressed (width*height*4 bytes, ~9.6MB for one plain settlement
  // page) -- jsPDF only actually PNG-compresses when it's handed a real
  // encoded PNG to parse, so every canvas goes through toDataURL() first.
  if(canvas.height * scale <= usableH){
    doc.addImage(canvas.toDataURL('image/png'), 'PNG', margin, margin, usableW, canvas.height * scale);
  } else {
    let renderedPx = 0, first = true;
    while(renderedPx < canvas.height){
      const sliceH = Math.min(pxPerPage, canvas.height - renderedPx);
      const slice = document.createElement('canvas');
      slice.width = canvas.width; slice.height = sliceH;
      slice.getContext('2d').drawImage(canvas, 0, renderedPx, canvas.width, sliceH, 0, 0, canvas.width, sliceH);
      if(!first) doc.addPage();
      doc.addImage(slice.toDataURL('image/png'), 'PNG', margin, margin, usableW, sliceH * scale);
      renderedPx += sliceH;
      first = false;
    }
  }
  return doc.output('blob');
}
/* Shared by both share paths below: hands a finished file to the Web
   Share API when the browser actually supports sharing files (Android
   Chrome, iOS Safari 15+) so the OS's own share sheet opens with
   WhatsApp sitting right there -- one tap closer than "download, open
   WhatsApp, attach it" was before. navigator.canShare({files}) is the
   real capability check; most desktop browsers return false here even
   though navigator.share itself exists, which is exactly the case the
   download+wa.me fallback covers -- wa.me can't attach a file itself,
   so that one manual tap to attach what just downloaded is unavoidable,
   and the message says so rather than silently doing less than promised.

   Desktop-specific wrinkle Alok hit: WhatsApp Web only allows ONE active
   browser tab per login -- opening wa.me in a fresh tab while another
   WhatsApp Web tab is already logged in doesn't reach the compose screen
   at all, it shows WhatsApp's own "used in another tab" interstitial
   with a "Use Here" button first. No website can skip that screen (it's
   WhatsApp's own session lock, not something this app controls), so
   rather than silently failing to look like it "shared," the pre-filled
   text now names the interstitial directly. Opening the SAME named
   window on every share (instead of a fresh "_blank" each time) at least
   stops repeated shares from this app piling up their own duplicate
   WhatsApp Web tabs against each other. */
async function shareFileOrFallback(blob, fileName, mimeType, shareText){
  const file = new File([blob], fileName, {type:mimeType});
  if(navigator.canShare && navigator.canShare({files:[file]})){
    try{
      await navigator.share({ files:[file], title:'UPGB OTS Calculator', text:shareText });
    }catch(err){
      if(err && err.name==='AbortError') return; // user closed the share sheet -- not a failure
      throw err;
    }
    return;
  }
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = fileName;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url), 30000);
  const kind = mimeType==='application/pdf'?'PDF':'image';
  const msg = `${shareText} — ${kind} downloaded to your computer. If WhatsApp shows "used in another tab", tap Use Here, then attach the ${kind} from your Downloads.`;
  window.open(`https://wa.me/?text=${encodeURIComponent(msg)}`, 'npaOtsWhatsAppShare');
}
/* Renders the exact same print sheet as printOtsSheet() to a real PDF
   file client-side (html2canvas -> jsPDF, no server involved). #printArea
   is display:none outside of @media print, so html2canvas (which renders
   the live DOM, not a print simulation) would otherwise capture nothing;
   it's floated on-screen at -9999px just long enough to rasterize, then
   restored exactly as it was. */
async function shareOtsPdf(){
  const slots = window.__slots; const custRow = window.__custRow;
  if(!slots || !custRow) return;
  await Promise.all([ensureHtml2Canvas(), ensureJsPDF()]);
  renderPrintView();
  const printEl = document.getElementById('printArea');
  const prevCss = printEl.style.cssText;
  printEl.style.cssText = 'display:block;position:fixed;left:-9999px;top:0;width:760px;background:#fff;padding:24px;z-index:-1';
  if(document.fonts && document.fonts.ready){ try{ await document.fonts.ready; }catch(e){} }
  await new Promise(r=>setTimeout(r, 60));
  let blob;
  try{
    const canvas = await html2canvas(printEl, { scale:2, backgroundColor:'#ffffff' });
    blob = canvasToPdfBlob(canvas);
  } finally {
    printEl.style.cssText = prevCss;
  }
  const safeName = String(custRow[C.NAME]||'borrower').replace(/[\\/:*?"<>|]/g,' ').replace(/\s+/g,' ').trim().slice(0,40);
  const fileName = `OTS_${safeName}_${dateToInputValue(new Date())}.pdf`;
  const shareText = `${custRow[C.NAME]||'—'} — ${custRow[C.SOL_DESC]||'—'}`;
  await shareFileOrFallback(blob, fileName, 'application/pdf', shareText);
}
window.shareOtsPdf = () => shareOtsPdf().catch(err=>{
  console.error(err);
  alert('Could not prepare the PDF to share. Please try again.');
});

/* The compact, mobile-shaped summary card -- Total O/S as the headline,
   each linked account's own Dues/P&L/Asset Code, Total Dues closing it
   out -- built to be shared as an image instead of the full A4 Settlement Statement,
   which reads as a wall of small print in a WhatsApp thumbnail. Approved
   mockup direction: navy/gold "Hero Stat" card, one card = one image.
   NPA date in the header uses the earliest across linked accounts when
   they differ (rare, but two accounts under one customer aren't
   guaranteed to have slipped into NPA on the same date). */
const WA_ICON_COIN = '<svg class="wa-icon" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><circle cx="12" cy="12" r="9"/><path d="M12 7v10M9 9.5c0-1.1 1.3-2 3-2s3 .9 3 2-1.3 1.5-3 2-3 .9-3 2 1.3 2 3 2 3-.9 3-2"/></svg>';
const WA_ICON_TREND = '<svg class="wa-icon" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><polyline points="3 17 9 11 13 15 21 6"/><polyline points="14 6 21 6 21 13"/></svg>';
const WA_ICON_DOC = '<svg class="wa-icon" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>';
const WA_ICON_SEAL = '<svg class="wa-icon" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><circle cx="12" cy="8" r="6"/><path d="M9 13.5L7 22l5-3 5 3-2-8.5"/></svg>';
const WA_ICON_CLOCK = '<svg width="9" height="9" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>';
/* Calendar-accurate years+months since the NPA date, not a /365.25
   decimal approximation (Alok: "npa ki age hai use year and month main
   kar k likh do" -- a decimal year reads as a made-up stat next to a
   real rupee figure, whole years+months reads like something a banker
   would actually say). Borrows a month from years when the day-of-month
   hasn't been reached yet this month, same logic as any calendar-age calc. */
function npaAgeStr(from){
  const now = new Date();
  let years = now.getFullYear() - from.getFullYear();
  let months = now.getMonth() - from.getMonth();
  if(now.getDate() < from.getDate()) months--;
  if(months < 0){ years--; months += 12; }
  if(years <= 0 && months <= 0) return '<1 mo in NPA';
  const yPart = years>0 ? `${years} yr${years===1?'':'s'}` : '';
  const mPart = months>0 ? `${months} mo${months===1?'':'s'}` : '';
  return `${[yPart,mPart].filter(Boolean).join(' ')} in NPA`;
}
function renderShareCard(){
  const slots = window.__slots; const custRow = window.__custRow;
  if(!slots || !custRow) return;
  const totalOS = slots.reduce((a,s)=>a+((s.os!=='')?s.os:0),0);
  let totalPL = 0;
  slots.forEach(s=>{ totalPL += (s.totalPL!==''?s.totalPL:0); });
  const npaDates = slots.map(s=>toDate(s.npaDate)).filter(Boolean);
  const earliestNpa = npaDates.length ? new Date(Math.min(...npaDates.map(d=>d.getTime()))) : null;
  const npaAge = earliestNpa ? npaAgeStr(earliestNpa) : null;
  const solId = esc(custRow[C.SOL_ID])||'';
  const logoSrc = document.querySelector('.nav-logo')?.src || '';
  // Same partial-settlement convention as renderPrintView()'s aggregate
  // totals: only accounts that actually have an OTS Amount typed in
  // count toward the OTS/impact/gauge sums -- an account with no figure
  // yet isn't silently treated as zero (and doesn't count toward the
  // Recovery % denominator either).
  let totalOtsSum = 0, totalImpact = 0, totalSacrifice = 0, osOfOtsAccts = 0, anyOts = false;
  // Lok Adalat minimum is only meaningful next to an account that's both
  // settled AND eligible (Substandard accounts have no Lok Adalat floor
  // at all) -- an ineligible account with an OTS Amount typed in just
  // sits outside this comparison entirely, same as an unsettled account
  // sits outside the OTS/impact sums above.
  let totalLokAdalatMin = 0, anyLokAdalatEligible = false;
  const acctCards = slots.map(s=>{
    const pl = s.totalPL; const pct = s.ratio!==''?` (${(s.ratio*100).toFixed(1)}%)`:'';
    const ots = parseOtsAmount(otsAmounts[s.acctNo]);
    if(ots!==null){
      totalOtsSum += ots; totalImpact += (ots-pl); totalSacrifice += (s.os-ots);
      osOfOtsAccts += s.os; anyOts = true;
      const la = lokAdalatMin(s);
      if(la && la.eligible){ totalLokAdalatMin += la.amount; anyLokAdalatEligible = true; }
    }
    return `<div class="wa-acct" data-asset="${esc(s.assetCode)}">
      <div class="info">
        <div class="no">A/c ${esc(s.acctNo)}</div>
        <div class="scheme">${esc(s.scheme)||'—'} &middot; ${esc(s.assetCode)||'—'}</div>
      </div>
      <div class="amts">
        <div class="os wa-num">${fmtINR2(s.os)}</div>
        <div class="pl">P&amp;L ${fmtINR2(pl)}${pct}</div>
      </div>
    </div>`;
  }).join('');
  // Once at least one account has a settlement figure typed in, the
  // Settlement Offer becomes the headline outcome worth sharing -- with
  // a Recovery vs. Sacrifice gauge underneath, the same visual language
  // as the live Loan Detail screen's own Recovery Scale.
  let offerHtml = '';
  if(anyOts){
    const arrow = totalImpact>0?'▲':(totalImpact<0?'▼':'');
    const cls = totalImpact>0?'pos':(totalImpact<0?'neg':'');
    const recoveryPct = osOfOtsAccts>0 ? (totalOtsSum/osOfOtsAccts*100) : 0;
    const sacrificePct = 100-recoveryPct;
    offerHtml = `<div class="wa-offer">
      <div class="title">${WA_ICON_SEAL}Settlement Offer</div>
      <div class="amt wa-num">${fmtINR2(totalOtsSum)}</div>
      <div class="row2">
        <div class="grp"><div class="k">Total Sacrifice</div><div class="v wa-num">${fmtINR2(totalSacrifice)}</div></div>
        <div class="grp"><div class="k">P&amp;L Impact</div><div class="v ${cls} wa-num">${arrow} ${fmtINR2(Math.abs(totalImpact))}</div></div>
      </div>
      <div class="wa-gauge"><div class="fill" style="width:${recoveryPct.toFixed(1)}%"></div><div class="dot" style="left:${recoveryPct.toFixed(1)}%"></div></div>
      <div class="wa-gauge-labels">
        <span>Recovered <b>${recoveryPct.toFixed(1)}%</b></span>
        <span class="sac">Sacrificed <b>${sacrificePct.toFixed(1)}%</b></span>
      </div>
      ${anyLokAdalatEligible ? `<div class="wa-cmp">
        <div><div class="k">Lok Adalat Minimum</div><div class="v wa-num">${fmtINR2(totalLokAdalatMin)}</div></div>
        <span class="wa-met ${totalOtsSum>=totalLokAdalatMin?'yes':'no'}">${totalOtsSum>=totalLokAdalatMin?'✓ Meets minimum':'⚠ Below minimum'}</span>
      </div>` : ''}
    </div>`;
  }
  document.getElementById('shareCardArea').innerHTML = `
    <div class="wa-card">
      <div class="wa-band">
        ${logoSrc?`<img class="wa-logo-img" src="${logoSrc}" alt="">`:''}
        <div class="wa-bank-group">
          <div class="wa-bank">Uttar Pradesh Gramin Bank</div>
          <div class="wa-bank-sub">Regional Office Hathras</div>
        </div>
      </div>
      <div class="wa-content">
        <div class="wa-profile">
          <div>
            <div class="wa-name">${esc(custRow[C.NAME])||'—'}</div>
            <div class="wa-sub">Cust ID ${esc(custRow[C.CUST_ID])||'—'} &middot; ${esc(custRow[C.SOL_DESC])||''}${solId?` (${solId})`:''}</div>
          </div>
          ${npaAge!==null?`<div class="wa-npa-pill">${WA_ICON_CLOCK} ${npaAge}</div>`:''}
        </div>
        <div class="wa-stats">
          <div class="wa-stat"><div class="k wa-lbl">${WA_ICON_COIN}Total O/S Balance</div><div class="v wa-num">${fmtINR2(totalOS)}</div></div>
          <div class="wa-stat warn"><div class="k wa-lbl">${WA_ICON_TREND}Total P&amp;L</div><div class="v wa-num">${fmtINR2(totalPL)}</div></div>
        </div>
        <div class="wa-section-lbl">${WA_ICON_DOC}Account Breakdown</div>
        ${acctCards}
        ${offerHtml}
        <div class="wa-foot-line">as on ${fmtDate(new Date())} &middot; UPGB OTS Calculator</div>
      </div>
    </div>
  `;
}
async function shareOtsImage(){
  const slots = window.__slots; const custRow = window.__custRow;
  if(!slots || !custRow) return;
  await ensureHtml2Canvas();
  renderShareCard();
  const cardEl = document.getElementById('shareCardArea');
  const prevCss = cardEl.style.cssText;
  cardEl.style.cssText = 'display:block;position:fixed;left:-9999px;top:0;z-index:-1';
  if(document.fonts && document.fonts.ready){ try{ await document.fonts.ready; }catch(e){} }
  await new Promise(r=>setTimeout(r, 60));
  let blob;
  try{
    const canvas = await html2canvas(cardEl.firstElementChild, { scale:2, backgroundColor:null });
    blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/png'));
  } finally {
    cardEl.style.cssText = prevCss;
  }
  const safeName = String(custRow[C.NAME]||'borrower').replace(/[\\/:*?"<>|]/g,' ').replace(/\s+/g,' ').trim().slice(0,40);
  const fileName = `OTS_${safeName}_${dateToInputValue(new Date())}.png`;
  const shareText = `${custRow[C.NAME]||'—'} — ${custRow[C.SOL_DESC]||'—'}`;
  await shareFileOrFallback(blob, fileName, 'image/png', shareText);
}
window.shareOtsImage = () => shareOtsImage().catch(err=>{
  console.error(err);
  alert('Could not prepare the image to share. Please try again.');
});

/* Small anchored menu on the WhatsApp share button -- Alok wanted both
   the full compliance-record PDF AND the quick-glance image available,
   not one replacing the other, so the button asks which one instead of
   picking for him. Built and torn down on demand rather than living in
   the static markup, same pattern as other one-off popovers in this
   app; a single document-level click listener (registered with `once`)
   closes it on any outside click without needing to track focus state. */
function toggleShareOtsMenu(evt){
  evt.stopPropagation();
  const existing = document.getElementById('shareOtsMenu');
  if(existing){ existing.remove(); return; }
  const btn = evt.currentTarget;
  const menu = document.createElement('div');
  menu.id = 'shareOtsMenu';
  menu.className = 'share-ots-menu';
  menu.innerHTML = `
    <button type="button" onclick="event.stopPropagation();document.getElementById('shareOtsMenu').remove();shareOtsPdf()">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M14 2H6a2 2 0 00-2 2v16a2 2 0 002 2h12a2 2 0 002-2V8z"/><polyline points="14 2 14 8 20 8"/></svg>
      <span>Full PDF</span>
    </button>
    <button type="button" onclick="event.stopPropagation();document.getElementById('shareOtsMenu').remove();shareOtsImage()">
      <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><path d="M21 15l-5-5L5 21"/></svg>
      <span>Summary Image</span>
    </button>
  `;
  document.body.appendChild(menu);
  const r = btn.getBoundingClientRect();
  menu.style.top = (r.bottom + window.scrollY + 6) + 'px';
  menu.style.right = (window.innerWidth - r.right) + 'px';
  setTimeout(()=>document.addEventListener('click', function closeMenu(){
    document.getElementById('shareOtsMenu')?.remove();
    document.removeEventListener('click', closeMenu);
  }), 0);
}
window.toggleShareOtsMenu = toggleShareOtsMenu;
/* Illustrative severity bands for NPA % (NPA outstanding / total advance),
   not a claim of official RBI benchmark thresholds -- just enough to spot
   a high-NPA branch/region at a glance. */
function npaPctSeverity(pct){
  if(pct>=10) return {color:'var(--red)', soft:'var(--red-soft)'};
  if(pct>=5) return {color:'var(--amber)', soft:'var(--amber-soft)'};
  return {color:'var(--green)', soft:'var(--green-soft)'};
}
const ASSET_LABELS = {SUB_STD:'Substandard asset', DA1:'Doubtful — up to 1 year', DA2:'Doubtful — 1 to 3 years', DA3:'Doubtful — more than 3 years', LOSS:'Loss asset'};
function assetLabel(code){ return ASSET_LABELS[code] || code; }
function titleCase(s){ return String(s||'').toLowerCase().replace(/\b\w/g,c=>c.toUpperCase()); }

/* Sol ID / Branch reference list -- [oldSolId, newSolId, branchName], from
   Alok's SOL_ID.xlsx. Static reference data (doesn't change with daily NPA
   updates), so it's embedded directly rather than published/uploaded like
   the NPA dataset -- one row per branch, "R O Hathras" (the Regional
   Office) listed first same as in the source file, rest in Sol ID order. */
/* Frozen source of truth: UPGB_NEW_SOL_ID.xlsx (the official Old/New Sol ID
   + branch name master Alok supplied 2026-08-15). Verified against this
   sheet's own "District Name" column that the Hathras/Mathura split used by
   branchGroups() below (9270-9309 vs 9310-9325) is exactly right, and
   corrected 4 branch names to the sheet's official long form (Agra Road ->
   Hathras Agra Road, Aligarh Road -> Hathras Aligarh Road, Service Branch ->
   Hathras Service Branch, Hatisa -> Hatisa Bhagwantpur). All 57 Old/New Sol
   ID pairs matched the sheet exactly, no other differences. */
const BRANCH_LIST = [[15990,9269,"R O Hathras"],[15010,9270,"Agsauli"],[15020,9271,"Bamnai"],[15030,9272,"Bandhnoo"],[15040,9273,"Baraus"],[15050,9274,"Bastoi"],[15060,9275,"Bisawar"],[15070,9276,"Chandpa"],[15080,9277,"Chhonda Gadua"],[15090,9278,"Devinagar"],[15100,9279,"Eihan"],[15110,9280,"Hathras Agra Road"],[15120,9281,"Hathras Aligarh Road"],[15130,9282,"Mursan Gate"],[15140,9283,"Hathras Service Branch"],[15150,9284,"Hatisa Bhagwantpur"],[15160,9285,"Jarera"],[15170,9286,"Komari"],[15180,9287,"Kota"],[15190,9288,"Ladpur"],[15200,9289,"Mahow"],[15210,9290,"Meetai"],[15220,9291,"Mendu"],[15230,9292,"Mughal Garhi"],[15240,9293,"Mursan"],[15250,9294,"Parsara"],[15260,9295,"Pora"],[15270,9296,"Purdil Nagar"],[15280,9297,"Ratibhanpur"],[15290,9298,"Ruheri"],[15300,9299,"Sadabad"],[15310,9300,"Sahpau"],[15320,9301,"Salempur"],[15330,9302,"Sasni"],[15340,9303,"Sikandra Rao"],[15350,9304,"Tuksan"],[15360,9305,"Wazidpur"],[15370,9306,"Adarshnagar"],[15380,9307,"Hasayan"],[15390,9308,"Jaleser Road"],[15400,9309,"Naugaon"],[16010,9310,"Bajna"],[16020,9311,"Baldev"],[16030,9312,"Bati"],[16040,9313,"Damodarpura"],[16050,9314,"Farah"],[16060,9315,"Goverdhan"],[16070,9316,"Maant"],[16080,9317,"Mathura City"],[16090,9318,"Laxmi Nagar"],[16100,9319,"Pali Kheda"],[16110,9320,"Raya"],[16120,9321,"Ronchi Bangar"],[16130,9322,"Sonai"],[16140,9323,"Tarsi"],[16150,9324,"Vrindavan"],[16160,9325,"Jajan Patti"]];
/* Branch master data from the same frozen UPGB_NEW_SOL_ID.xlsx source as
   BRANCH_LIST above -- branch code, official branch email, RO/Branch type,
   Urban/Rural/Semi Urban area, district, registered address, PIN, and date
   opened. Keyed by new Sol ID (number). This is the authoritative source
   for branch district (used by branchGroups() below) rather than a
   hardcoded Sol ID range, since the sheet's own District Name column is
   ground truth, not an inference from the numbering. */
const BRANCH_META = {9269:{code:"ROHATH",email:"recovery.rohath@upgb.bank.in",type:"Regional Office",area:"Urban",district:"Hathras",address:"MUNSHI GAJADHAR MARG ALIGARH ROAD",pin:"204101",dateOpen:"01-04-2013"},9270:{code:"AGSAUA",email:"AGSAUA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.AGSAULI HATHRAS",pin:"204210",dateOpen:"08-08-1983"},9271:{code:"BAMNHA",email:"BAMNHA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.LUHETA HATHRAS",pin:"204101",dateOpen:"11-08-1982"},9272:{code:"BANDHA",email:"BANDHA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.KGW SASNI HATHRAS",pin:"202139",dateOpen:"10-12-1983"},9273:{code:"BARAHA",email:"BARAHA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"VILL. BARAUS PO- BANS AMRU",pin:"281306",dateOpen:"31-12-2012"},9274:{code:"BASTOA",email:"BASTOA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O. BASTOI HATHRAS",pin:"204215",dateOpen:"20-12-1983"},9275:{code:"BISAWA",email:"BISAWA@upgb.bank.in",type:"Branch",area:"Semi Urban",district:"Hathras",address:"MOHALLA PENTH BAZAR, BISAWAR BLOCK- SADABAD",pin:"281302",dateOpen:"14-03-2012"},9276:{code:"CHANHA",email:"CHANHA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.CHANDPA HATHRAS",pin:"204101",dateOpen:"17-09-1981"},9277:{code:"CHHONA",email:"CHHONA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"VILLAGE- CHHONDA GADUA P.O. GADUA",pin:"204216",dateOpen:"14-03-2012"},9278:{code:"DEVINA",email:"DEVINA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.HATHRAS JUNCTION HATHRAS",pin:"204102",dateOpen:"09-08-1994"},9279:{code:"EIHANA",email:"EIHANA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.EIHAN HATHRAS",pin:"204102",dateOpen:"19-08-1982"},9280:{code:"HATRDA",email:"HATRDA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"AGRA ROAD HATHRAS",pin:"204101",dateOpen:"08-11-1994"},9281:{code:"HATHHA",email:"HATHHA@upgb.bank.in",type:"Branch",area:"Urban",district:"Hathras",address:"MUNSHI GAJADHAR MARG ALIGARH ROAD",pin:"204101",dateOpen:"22-03-2012"},9282:{code:"MURSNA",email:"MURSNA@upgb.bank.in",type:"Branch",area:"Urban",district:"Hathras",address:"MURSAN GATE HATHRAS HATHRAS",pin:"204101",dateOpen:"08-02-1994"},9283:{code:"HATHRA",email:null,type:"Service Branch",area:"Urban",district:"Hathras",address:"MUNSHI GAJADHAR MARG ALIGARH ROAD",pin:"204101",dateOpen:"16-07-2012"},9284:{code:"HATISA",email:"HATISA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.HATISA BHAGWANTPUR HATHRAS",pin:"204101",dateOpen:"28-09-1984"},9285:{code:"JARERA",email:"JARERA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O. NAGLA VEER SAHAI HATHRAS",pin:"204214",dateOpen:"10-08-1983"},9286:{code:"KOMARA",email:"KOMARA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.KOMRI HATHRAS",pin:"202139",dateOpen:"11-12-1981"},9287:{code:"KOTAHA",email:"KOTAHA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"VILLAGE and P.O. KOTA BLOCK - MURSAN",pin:"204213",dateOpen:"14-03-2012"},9288:{code:"LADPUA",email:"LADPUA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.LADPUR HATHRAS",pin:"204101",dateOpen:"27-05-1981"},9289:{code:"MAHOWA",email:"MAHOWA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.,MAHOW HATHRAS",pin:"204121",dateOpen:"02-09-1981"},9290:{code:"MEETAA",email:"MEETAA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.MEETAI HATHRAS",pin:"204101",dateOpen:"26-04-1982"},9291:{code:"MENDUA",email:"MENDUA@upgb.bank.in",type:"Branch",area:"Semi Urban",district:"Hathras",address:"P.O.MENDU HATHRAS",pin:"204105",dateOpen:"18-09-1985"},9292:{code:"MUGHAA",email:"MUGHAA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"VILL. and PO- MUGHALGARHI TEHSIL- S. RAO",pin:"204215",dateOpen:"29-03-2013"},9293:{code:"MURSAA",email:"MURSAA@upgb.bank.in",type:"Branch",area:"Semi Urban",district:"Hathras",address:"P.O. MURSAN HATHRAS",pin:"204213",dateOpen:"31-10-1984"},9294:{code:"PARSRA",email:"PARSRA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"VILLAGE and P.O. PARSARA BLOCK- HATHRAS",pin:"204101",dateOpen:"14-03-2012"},9295:{code:"PORAHA",email:"PORAHA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.PORA HATHRAS",pin:"204215",dateOpen:"14-10-1982"},9296:{code:"PURDIA",email:"PURDIA@upgb.bank.in",type:"Branch",area:"Semi Urban",district:"Hathras",address:"P.O.PURDILNAGAR HATHRAS",pin:"204214",dateOpen:"10-01-1995"},9297:{code:"RATIHA",email:"RATIHA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O. PIPAL GAVAN HATHRAS",pin:"204215",dateOpen:"09-08-1983"},9298:{code:"RUHERA",email:"RUHERA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.RUHERI HATHRAS",pin:"204101",dateOpen:"19-10-1982"},9299:{code:"SADABA",email:"SADABA@upgb.bank.in",type:"Branch",area:"Semi Urban",district:"Hathras",address:"HIGHWAY PLAZA, AGRA- ALIGARH ROAD SADABAD",pin:"281306",dateOpen:"29-02-2008"},9300:{code:"SAHPAA",email:"SAHPAA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"MAIN ROAD, MOHALLA- BAJARIA VILLAGE and P.O. SAHPAU",pin:"281307",dateOpen:"14-03-2012"},9301:{code:"SALEMA",email:"SALEMA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.SALEMPUR HATHRAS",pin:"202124",dateOpen:"15-10-1981"},9302:{code:"SASNIA",email:"SASNIA@upgb.bank.in",type:"Branch",area:"Semi Urban",district:"Hathras",address:"P.O.SASNI HATHRAS",pin:"204216",dateOpen:"05-02-1994"},9303:{code:"SIKADA",email:"SIKADA@upgb.bank.in",type:"Branch",area:"Semi Urban",district:"Hathras",address:"P.O. S.RAO HATHRAS",pin:"204215",dateOpen:"09-02-1994"},9304:{code:"TUKSAA",email:"TUKSAA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.TUKSAN HATHRAS",pin:"204101",dateOpen:"16-09-1983"},9305:{code:"WAZIDA",email:"WAZIDA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"P.O.WAZIDPUR HATHRAS",pin:"204215",dateOpen:"25-01-1982"},9306:{code:"ADARSA",email:"ADARSA@upgb.bank.in",type:"Branch",area:"Urban",district:"Hathras",address:"-Adarshnagar Maindu Road -Hathras-204101 std-05722",pin:"204101",dateOpen:"03-05-2016"},9307:{code:"HASAYA",email:"HASAYA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"Hasayan -Sikandra rao-Hathras- pin-204212 Std code-05721",pin:"204212",dateOpen:"03-05-2016"},9308:{code:"JALESA",email:"JALESA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"Jaleser Road  P- JaleserRS Tahsil -Sadabad -pin 281104 STD-05745",pin:"281104",dateOpen:"15-05-2016"},9309:{code:"NAUGGA",email:"NAUGGA@upgb.bank.in",type:"Branch",area:"Rural",district:"Hathras",address:"V+P Nagaonn-tahsil -Sahabad-Hathras.  pin-281502 std 0565",pin:"281502",dateOpen:"18-05-2016"},9310:{code:"BAJNAA",email:"BAJNAA@upgb.bank.in",type:"Branch",area:"Rural",district:"Mathura",address:"MOHALLA SHIVAJI NAGAR BAJNA",pin:"281201",dateOpen:"20-01-2010"},9311:{code:"BALDEA",email:"BALDEA@upgb.bank.in",type:"Branch",area:"Rural",district:"Mathura",address:"JAWAHAR ROAD, NEW POST OFFICE BUILDING BALDEV",pin:"281301",dateOpen:"19-03-2012"},9312:{code:"BATIHA",email:"BATIHA@upgb.bank.in",type:"Branch",area:"Rural",district:"Mathura",address:"VILL. and PO- BATI, MAIN ROAD",pin:"281004",dateOpen:"29-03-2013"},9313:{code:"DAMODA",email:"DAMODA@upgb.bank.in",type:"Branch",area:"Rural",district:"Mathura",address:"VILL. DAMODARPURA PO- AURANGABAD",pin:"281006",dateOpen:"31-12-2012"},9314:{code:"FARAHA",email:"FARAHA@upgb.bank.in",type:"Branch",area:"Rural",district:"Mathura",address:"NEAR BUS STAND POST FARAH",pin:"281122",dateOpen:"30-06-2008"},9315:{code:"GOVERA",email:"GOVERA@upgb.bank.in",type:"Branch",area:"Semi Urban",district:"Mathura",address:"SARAI BARA BAZAR GOVERDHAN",pin:"281502",dateOpen:"10-03-2008"},9316:{code:"MAANTA",email:"MAANTA@upgb.bank.in",type:"Branch",area:"Rural",district:"Mathura",address:"RAYA NAUJHIL ROAD MAANT",pin:"281202",dateOpen:"26-03-2010"},9317:{code:"MATHUA",email:"MATHUA@upgb.bank.in",type:"Branch",area:"Urban",district:"Mathura",address:"17-A RADHA NAGAR, OPPOSITE MADHUVAN HOTEL KRISHNA NAGAR",pin:"281004",dateOpen:"05-03-2008"},9318:{code:"LAXMIA",email:"LAXMIA@upgb.bank.in",type:"Branch",area:"Urban",district:"Mathura",address:"BEHIND - MAA CHANDRAWALI PETROL PUMP LAXMI NAGAR",pin:"281001",dateOpen:"14-03-2012"},9319:{code:"PALIKA",email:"PALIKA@upgb.bank.in",type:"Branch",area:"Rural",district:"Mathura",address:"OM NAGAR COLONY, PALIKHERA SONKH ROAD",pin:"281004",dateOpen:"24-03-2012"},9320:{code:"RAYAHA",email:"RAYAHA@upgb.bank.in",type:"Branch",area:"Semi Urban",district:"Mathura",address:"SUBEDAR ATAR SINGH MARKET, HATHRAS- MATHURA ROAD RAYA",pin:"281204",dateOpen:"11-03-2008"},9321:{code:"RONCHA",email:"RONCHA@upgb.bank.in",type:"Branch",area:"Rural",district:"Mathura",address:"KADAMB VIHAR ROAD RONCHI BANGAR",pin:"281006",dateOpen:"31-12-2012"},9322:{code:"SONAIA",email:"SONAIA@upgb.bank.in",type:"Branch",area:"Rural",district:"Mathura",address:"VILL. and PO- SONAI MATHURA",pin:"281206",dateOpen:"31-12-2012"},9323:{code:"TARSIA",email:"TARSIA@upgb.bank.in",type:"Branch",area:"Rural",district:"Mathura",address:"VILL. TARSI, PO- DHANGAO MATHURA",pin:"281005",dateOpen:"29-03-2013"},9324:{code:"VRINDA",email:"VRINDA@upgb.bank.in",type:"Branch",area:"Semi Urban",district:"Mathura",address:"MUDGAL RISHI BHAWAN, MOTI JHEEL MARG VRINDAVAN",pin:"281121",dateOpen:"17-03-2009"},9325:{code:"JAJANA",email:"JAJANA@upgb.bank.in",type:"Branch",area:"Rural",district:"Mathura",address:"v -Jajanpatti Block Goverdhan- Mathura.pin code-281123",pin:"281123",dateOpen:"17-08-2016"}};
/* Branch contacts (Manager + Recovery Officer) now live on DATA.branchContacts,
   uploaded via Update Data -> Branch Contacts (see buildBranchContactsMap/
   handleBranchContactsUpload below) and published alongside the rest of the
   data, same pattern as DATA.branchAdvances -- replaces the earlier
   hardcoded MANAGER_CONTACTS constant now that collection is an ongoing,
   self-serve process rather than a one-off code ship. */

function branchMatchesQuery(name, oldId, newId, q){
  const bc = DATA.branchContacts[String(newId)];
  return String(newId).includes(q) || String(oldId).includes(q) || name.toLowerCase().includes(q) ||
    (bc && ((bc.mgr||'').toLowerCase().includes(q) || (bc.roName||'').toLowerCase().includes(q)));
}
function branchRowHtml([oldId,newId,name]){
  return `
    <div class="edge-row" onclick="showBranchCard(${newId})" role="button" tabindex="0" aria-label="View full contact card for ${esc(name)}">
      <div class="edge-row-top">
        <div class="edge-row-branch">${esc(name)}</div>
        <div class="edge-row-ids"><span class="edge-solid">Sol ID ${esc(newId)}</span><span class="edge-oldid">Old ${esc(oldId)}</span></div>
      </div>
    </div>`;
}
/* Grouped view of BRANCH_LIST, sorted by new Sol ID ascending (low to
   high) -- branch staff look branches up by Sol ID, not alphabetically, so
   that's the order that's actually useful here. BRANCH_LIST's own declared
   order already happens to be ascending by Sol ID, but this sorts
   defensively rather than relying on that, since other code (Excel
   template, showBranchCard lookup) depends on the source array's own
   order and shouldn't be reordered.
   Groups follow the real administrative split within UPGB Hathras Regional
   Office -- each branch's actual district, straight from BRANCH_META
   (the frozen official sheet), not an inferred Sol ID range. */
const BRANCH_DISTRICT_LETTER = {Hathras:'HTH', Mathura:'MTH'};
function branchGroups(){
  const ro = BRANCH_LIST.find(([,,name])=>name==='R O Hathras');
  const rest = BRANCH_LIST.filter(([,,name])=>name!=='R O Hathras')
    .slice().sort((a,b)=>a[1]-b[1]);
  const groups = [];
  if(ro) groups.push({id:'ro', letter:'★', label:'Regional Office', rows:[ro]});
  rest.forEach(entry=>{
    const dist = (BRANCH_META[entry[1]]||{}).district || 'Other';
    const last = groups[groups.length-1];
    if(!last || last.id!==dist){
      groups.push({id:dist, letter:BRANCH_DISTRICT_LETTER[dist]||dist.slice(0,3).toUpperCase(), label:`${dist} District`, rows:[entry]});
    } else last.rows.push(entry);
  });
  return groups;
}
function renderBranchList(filter){
  const q = (filter||'').trim().toLowerCase();
  const body = document.getElementById('branchListBody');
  const rail = document.getElementById('branchEdgeRail');
  const countEl = document.getElementById('branchListCount');
  if(!body) return;

  if(q){
    const rows = BRANCH_LIST.filter(([oldId,newId,name])=>branchMatchesQuery(name,oldId,newId,q));
    if(countEl) countEl.textContent = `${rows.length} match${rows.length===1?'':'es'}`;
    if(rail) rail.innerHTML = '';
    body.innerHTML = rows.length ? rows.map(branchRowHtml).join('') : `<div class="edge-empty">No branch matches "${esc(filter)}"</div>`;
    return;
  }

  if(countEl) countEl.textContent = `${BRANCH_LIST.length} branches`;
  const groups = branchGroups();
  body.innerHTML = groups.map(g=>`
    <div class="edge-grp" id="edgeGrp-${esc(g.id)}"><b>${esc(g.label)}</b><i></i><em>${g.rows.length}</em></div>
    ${g.rows.map(branchRowHtml).join('')}
  `).join('');
  if(rail) rail.innerHTML = groups.map(g=>`<button type="button" onclick="jumpBranchGroup('${esc(g.id)}')" aria-label="Jump to ${esc(g.label)}">${esc(g.letter)}</button>`).join('');
}
function jumpBranchGroup(id){
  document.getElementById('edgeGrp-'+id)?.scrollIntoView({block:'start', behavior:'smooth'});
}
window.jumpBranchGroup = jumpBranchGroup;
/* WhatsApp deep link for a mobile number -- wa.me needs the full
   international number with no "+"/spaces, so a bare 10-digit Indian
   mobile gets "91" prefixed; a number that already carries a country
   code is left as-is. Android/iOS don't offer WhatsApp in the tel: "Open
   with" chooser (it isn't registered as a tel: handler), so this is a
   separate icon/link next to the phone number rather than relying on
   that chooser to surface it. */
function toWaNumber(raw){
  const digits = String(raw||'').replace(/\D/g,'');
  if(digits.length===10) return '91'+digits;
  return digits;
}
function waIconLink(num){
  const wa = toWaNumber(num);
  if(!wa) return '';
  return `<a class="wa-link" href="https://wa.me/${wa}" target="_blank" rel="noopener" title="Chat on WhatsApp" aria-label="Chat on WhatsApp" onclick="event.stopPropagation()"><svg width="14" height="14" viewBox="0 0 448 512" fill="currentColor" aria-hidden="true"><path d="M380.9 97.1C339 55.1 283.2 32 223.9 32c-122.4 0-222 99.6-222 222 0 39.1 10.2 77.3 29.6 111L0 480l117.7-30.9c32.4 17.7 68.9 27 106.1 27h.1c122.3 0 224.1-99.6 224.1-222 0-59.3-25.2-115-67.1-157zm-157 341.6c-33.2 0-65.7-8.9-94-25.7l-6.7-4-69.8 18.3L72 359.2l-4.4-7c-18.5-29.4-28.2-63.3-28.2-98.2 0-101.7 82.8-184.5 184.6-184.5 49.3 0 95.6 19.2 130.4 54.1 34.8 34.9 56.2 81.2 56.1 130.5 0 101.8-84.9 184.6-186.6 184.6zm101.2-138.2c-5.5-2.8-32.8-16.2-37.9-18-5.1-1.9-8.8-2.8-12.5 2.8-3.7 5.6-14.3 18-17.6 21.8-3.2 3.7-6.5 4.2-12 1.4-32.6-16.3-54-29.1-75.5-66-5.7-9.8 5.7-9.1 16.3-30.3 1.8-3.7.9-6.9-.5-9.7-1.4-2.8-12.5-30.1-17.1-41.2-4.5-10.8-9.1-9.3-12.5-9.5-3.2-.2-6.9-.2-10.6-.2-3.7 0-9.7 1.4-14.8 6.9-5.1 5.6-19.4 19-19.4 46.3 0 27.3 19.9 53.7 22.6 57.4 2.8 3.7 39.1 59.7 94.8 83.8 35.2 15.2 49 16.5 66.6 13.9 10.7-1.6 32.8-13.4 37.4-26.4 4.6-13 4.6-24.1 3.2-26.4-1.3-2.5-5-3.9-10.5-6.6z"/></svg></a>`;
}
/* A handful of BRANCH_META addresses (the branches opened 2016 -- Adarshnagar,
   Hasayan, Jaleser Road, Naugaon) already have the PIN typed into the
   address text itself ("...std-05722"), unlike the rest which don't --
   append meta.pin only when it isn't already there, so those four don't
   show the same PIN twice. */
function masterAddressOf(meta){
  if(!meta.address) return null;
  if(meta.pin && !meta.address.includes(meta.pin)) return meta.address + ' - ' + meta.pin;
  return meta.address;
}
/* R O Hathras (the Regional Office, Sol ID 9269 -- BRANCH_META type
   "Regional Office") carries different job titles than every branch: its
   "Manager" contact is the Region Head, and its "Recovery Officer" is a
   Senior Manager Recovery, not a branch-level Recovery Officer. Every other
   branch keeps the plain Branch Manager / Recovery Officer labels. */
function branchRoleLabels(newId){
  const isRO = (BRANCH_META[newId]||{}).type === 'Regional Office';
  return isRO
    ? {mgrLabel:'Region Head', mgrShort:'RH', roLabel:'Senior Manager Recovery', roShort:'SMR'}
    : {mgrLabel:'Branch Manager', mgrShort:'MGR', roLabel:'Recovery Officer', roShort:'RO'};
}
/* Full branch detail card, reusing the same generic title/sub/info-grid
   modal already built for Quick Account Detail (quickAcctModalOverlay) --
   shows everything collected for that branch (Old + New Sol ID, Manager,
   Recovery Officer, and whatever else has been uploaded so far). */
function showBranchCard(newId){
  const entry = BRANCH_LIST.find(([,nid])=>nid===newId);
  if(!entry) return;
  const [oldId,,name] = entry;
  const meta = BRANCH_META[newId] || {};
  const bc = DATA.branchContacts[String(newId)] || {};
  const plain = (label,val) => [label, val ? esc(val) : null];
  const tel = (label,num) => [label, num ? `<span class="v-with-wa"><a href="tel:${esc(num)}">${esc(num)}</a>${waIconLink(num)}</span>` : null];
  const mail = (label,addr) => [label, addr ? `<a href="mailto:${esc(addr)}">${esc(addr)}</a>` : null];
  /* Manually-collected fields (bc.*, via the Branch Contacts upload) win
     over the frozen master sheet's own address when both are present --
     bc.address is a human keeping it current, meta.address is a one-time
     snapshot. */
  const masterAddress = masterAddressOf(meta);
  const roleLabels = branchRoleLabels(newId);
  const fields = [
    plain('Branch Type', meta.type),
    plain('District', meta.district),
    plain('Area', meta.area),
    plain('Branch Code', meta.code),
    mail('Branch Email', meta.email),
    plain('Date Opened', meta.dateOpen),
    plain('Address', bc.address || masterAddress),
    mail(roleLabels.mgrLabel+' Email', bc.mgrEmail),
    plain('Branch Landline', bc.landline),
    plain('Category', bc.category),
    plain('IFSC Code', bc.ifsc),
    plain('Remarks', bc.remarks),
  ];
  document.getElementById('quickAcctTitle').textContent = name;
  document.getElementById('quickAcctSub').innerHTML = `Sol ID ${esc(newId)} &middot; Old Sol ID ${esc(oldId)}`;
  document.getElementById('quickAcctGrid').innerHTML = fields.map(([k,v])=>`<div><div class="k">${esc(k)}</div><div class="v">${v!==null?v:'—'}</div></div>`).join('');
  document.getElementById('quickAcctModalOverlay').classList.add('show');
}
window.showBranchCard = showBranchCard;
function filterBranchList(){ renderBranchList(document.getElementById('branchListSearch').value); }
window.filterBranchList = filterBranchList;
// All edge panels are position:fixed at the same top-right corner and
// their pull-tabs stack down the same right edge -- with 4 of them now
// (Branch/Sol ID, Lok Adalat, Sacrifice Delegation, Account Details),
// letting more than one open at once stacked their headers directly on
// top of each other (a real report from Alok, back when there were only
// two). Every panel now goes through this one generic toggle: opening
// one force-closes every other panel, and hides every OTHER handle
// while any panel is open -- a handle's own glass/blur styling was
// designed to sit over page content, not over another panel's opaque
// background, so leaving a sibling handle visible while a panel covers
// part of it just looked broken (text half-swallowed) rather than
// actually being unclickable; hiding it entirely is the honest fix.
const EDGE_PANEL_KEYS = ['branch','lokAdalat','sacrifice','ledgerAccounts'];
function edgePanelEls(key){
  return {
    panel: document.getElementById(key+'EdgePanel'),
    backdrop: document.getElementById(key+'EdgeBackdrop'),
    handle: document.getElementById(key+'EdgeHandle'),
  };
}
function toggleEdgePanel(key, force){
  const {panel, backdrop, handle} = edgePanelEls(key);
  if(!panel) return;
  const open = force===undefined ? !panel.classList.contains('open') : force;
  EDGE_PANEL_KEYS.forEach(k=>{
    if(k===key) return;
    const other = edgePanelEls(k);
    if(other.panel){
      other.panel.classList.remove('open');
      other.backdrop.classList.remove('open');
      other.handle.classList.remove('active');
      other.handle.setAttribute('aria-expanded','false');
    }
  });
  panel.classList.toggle('open', open);
  backdrop.classList.toggle('open', open);
  handle.classList.toggle('active', open);
  handle.setAttribute('aria-expanded', open ? 'true' : 'false');
  EDGE_PANEL_KEYS.forEach(k=>{
    if(k===key) return;
    edgePanelEls(k).handle?.classList.toggle('edge-handle-hidden', open);
  });
  if(key==='branch' && open){
    renderBranchList('');
    const search = document.getElementById('branchListSearch');
    search.value = '';
    setTimeout(()=>search.focus(), 260);
  }
}
function toggleBranchPanel(force){ toggleEdgePanel('branch', force); }
window.toggleBranchPanel = toggleBranchPanel;
function toggleLokAdalatPanel(force){ toggleEdgePanel('lokAdalat', force); }
window.toggleLokAdalatPanel = toggleLokAdalatPanel;
function toggleSacrificePanel(force){ toggleEdgePanel('sacrifice', force); }
window.toggleSacrificePanel = toggleSacrificePanel;
function toggleLedgerAccountsPanel(force){ toggleEdgePanel('ledgerAccounts', force); }
window.toggleLedgerAccountsPanel = toggleLedgerAccountsPanel;
async function onedriveTryResume(){
  try{
    const app = await onedriveMsal();
    if(app.getAllAccounts().length){
      onedriveFolderStack = [{id:null, name:ONEDRIVE_ROOT_LABEL}];
      await onedriveLoadCurrentFolder();
    }
  }catch(e){ /* not signed in yet -- the Connect screen already showing is correct */ }
}

/* ---------- OneDrive browser (Alok's request, 2026-09-07) ----------
   "Mujhe is dashboard se kuch bhi download nahi karna bas ek one way
   ftp chahiye... jo bhi data one drive k folders main save hai wo seen
   ho sake and download kar saken" -- read-only, browse + download only,
   nothing from this app is ever written back to OneDrive. Whoever opens
   this screen signs into THEIR OWN personal Microsoft account via a
   popup -- it's a general "browse your own OneDrive from here" tab, not
   tied to any one person's account. Originally a slide-out edge panel;
   moved to a full page/tab (2026-09-07, same day) per Alok's follow-up
   ("tool tabs main set karo... pure page par data show ho, back k
   options hon") -- same auth/Graph logic below, now with the room for a
   proper sortable table, a filter box, and an explicit Back button
   alongside the breadcrumb, on top of what he originally asked for.

   Auth: MSAL.js (js/vendor/msal-browser.min.js, self-hosted like every
   other vendor lib in this app) against an app registration Alok created
   himself in Microsoft Entra ID ("NPA DASHBOARD", Personal Microsoft
   account users only) -- an unavoidable one-time step only he could do,
   since it requires signing into a Microsoft/Azure account. The Client ID
   and redirect URI below are that registration's own public identifiers,
   not secrets (an SPA app registration has no client secret at all --
   that's the whole point of the "Single-page application" platform type).
   Files.Read (not Files.ReadWrite) is the only Graph permission requested,
   matching the one-way, read-only intent at the API level too, not just
   in this UI.

   Scoped to one folder tree, not Alok's whole OneDrive: he asked for
   "D:\OneDrive\UPGB\Recovery\ALOK_MITTAL\HATHRAS" (his own OneDrive
   sync client's local mount point) and everything under it, specifically
   -- "D:\OneDrive\" is just where OneDrive happens to be mounted on his
   own PC, not part of the actual cloud path, so the path Graph itself
   needs is everything after that: UPGB/Recovery/ALOK_MITTAL/HATHRAS.
   Opening the tab loads straight into that folder's own contents
   (resolved by path, once); every subfolder from there on is navigated
   by its own Graph item ID, same as any folder. There's no way back out
   above HATHRAS from inside this screen -- the breadcrumb's root entry
   *is* HATHRAS, not OneDrive's real root, and the Back button disables
   itself there rather than exiting the scoped tree. */
const ONEDRIVE_CLIENT_ID = 'ad7b5590-643c-4b07-9814-8fd890e1568d';
const ONEDRIVE_REDIRECT_URI = 'https://npadashboard.alokmittal.net';
const ONEDRIVE_ROOT_PATH = 'UPGB/Recovery/ALOK_MITTAL/HATHRAS';
const ONEDRIVE_ROOT_LABEL = 'HATHRAS';
let onedriveMsalApp = null;
let onedriveMsalReady = null; // the in-flight/completed initialize() promise, shared across concurrent callers
let onedriveFolderStack = []; // [{id,name}, ...] -- stack[0].id is always null (root path lookup)
// Current folder's raw Graph items plus the live filter text and sort
// column/direction -- kept separate from the fetch itself so typing in
// the filter box or clicking a column header only ever re-renders the
// list/grid area (onedriveRenderListArea), never re-hits the network, and never
// touches the filter <input> itself so it never loses focus/cursor
// position mid-keystroke.
// view: 'list' or 'grid' (Google Drive-style toggle, 2026-09-07 UI
// refresh); remembered across visits the same way the theme toggle is,
// since it's a pure display preference with no bearing on the data itself.
let onedriveListState = {rawItems:[], filterText:'', sort:{key:'name', dir:'asc'}, view: onedriveSavedView()};
function onedriveSavedView(){
  try{ return localStorage.getItem('upgb-onedrive-view')==='grid' ? 'grid' : 'list'; }catch(e){ return 'list'; }
}
// msal-browser v3 requires `await instance.initialize()` before calling
// any other MSAL API (loginPopup, getAllAccounts, acquireTokenSilent) --
// the actual real-world failure hit here ("uninitialized_public_client_
// application") -- unlike v2, where the constructor alone was usable
// immediately. onedriveMsalReady caches that one initialize() call so
// concurrent callers (e.g. onedriveTryResume firing right as the tab
// opens) all await the same promise instead of racing separate ones.
async function onedriveMsal(){
  if(!onedriveMsalApp){
    await ensureMsal();
    onedriveMsalApp = new msal.PublicClientApplication({
      auth: { clientId: ONEDRIVE_CLIENT_ID, authority: 'https://login.microsoftonline.com/consumers', redirectUri: ONEDRIVE_REDIRECT_URI },
      cache: { cacheLocation: 'localStorage' },
    });
    onedriveMsalReady = onedriveMsalApp.initialize();
  }
  await onedriveMsalReady;
  return onedriveMsalApp;
}
async function onedriveGetToken(interactive){
  const app = await onedriveMsal();
  const accounts = app.getAllAccounts();
  if(accounts.length){
    try{
      const r = await app.acquireTokenSilent({ scopes:['Files.Read'], account: accounts[0] });
      return r.accessToken;
    }catch(e){ /* falls through to interactive below */ }
  }
  if(interactive===false) throw new Error('sign-in required');
  // prompt:'select_account' forces Microsoft's own login page to show its
  // account-chooser screen every time, even when the browser already has
  // an active Microsoft SSO session for some account -- without it,
  // login.live.com silently continues straight into that cached account's
  // own sign-in flow (e.g. an Authenticator "Get a sign-in request" push
  // screen) with no visible way to pick a different account at all. Real
  // production report: Alok couldn't switch to a different Microsoft
  // account on another device because this screen never appeared.
  const r = await app.loginPopup({ scopes:['Files.Read'], prompt:'select_account' });
  return r.accessToken;
}
function onedriveConnectScreenHtml(statusMsg){
  return `<div class="onedrive-connect">
      <svg width="34" height="34" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true"><path d="M7.5 17a4 4 0 0 1-.5-7.97A5.5 5.5 0 0 1 17.5 8a4 4 0 0 1 .5 7.97"/><path d="M9 17h9"/></svg>
      <p>Sign in with your personal Microsoft account to browse and download files from your own OneDrive. Nothing from this app is ever uploaded there.</p>
      <button type="button" class="onedrive-connect-btn" onclick="onedriveConnect()">Connect OneDrive</button>
      <div class="onedrive-status">${esc(statusMsg||'')}</div>
    </div>`;
}
async function onedriveConnect(){
  document.getElementById('onedrivePageBody').innerHTML = onedriveConnectScreenHtml('Signing in…');
  try{
    await onedriveGetToken(true);
    onedriveFolderStack = [{id:null, name:ONEDRIVE_ROOT_LABEL}];
    await onedriveLoadCurrentFolder();
  }catch(err){
    console.error(err);
    // Surfaces MSAL's own error code/message (e.g. "popup_window_error",
    // an AADSTS#### redirect-URI mismatch, "user_cancelled") right in the
    // screen instead of one generic string for every failure -- otherwise
    // diagnosing a real sign-in problem needs someone to open DevTools
    // and read the console, which isn't realistic for most users of this
    // app to be asked to do.
    const detail = (err && (err.errorCode || err.name)) ? `${err.errorCode||err.name}${err.errorMessage?': '+err.errorMessage:(err.message?': '+err.message:'')}` : (err && err.message) || 'Unknown error';
    document.getElementById('onedrivePageBody').innerHTML = onedriveConnectScreenHtml(`Could not sign in — ${detail}`);
  }
}
window.onedriveConnect = onedriveConnect;
// Explicit sign-out (Alok's request) returns to the same Connect screen
// used for a fresh sign-in -- both states are always one click away from
// each other, never just one or the other.
async function onedriveSignOut(){
  const app = await onedriveMsal();
  const accounts = app.getAllAccounts();
  if(accounts.length) app.logoutPopup({ account: accounts[0] }).catch(()=>{});
  onedriveFolderStack = [];
  onedriveListState = {rawItems:[], filterText:'', sort:{key:'name', dir:'asc'}};
  document.getElementById('onedrivePageBody').innerHTML = onedriveConnectScreenHtml();
}
window.onedriveSignOut = onedriveSignOut;
function onedriveFmtSize(bytes){
  if(typeof bytes!=='number') return '';
  if(bytes<1024) return bytes+' B';
  if(bytes<1024*1024) return (bytes/1024).toFixed(1)+' KB';
  if(bytes<1024*1024*1024) return (bytes/1024/1024).toFixed(1)+' MB';
  return (bytes/1024/1024/1024).toFixed(2)+' GB';
}
async function onedriveLoadCurrentFolder(){
  const cur = onedriveFolderStack[onedriveFolderStack.length-1];
  const body = document.getElementById('onedrivePageBody');
  body.innerHTML = `<div class="onedrive-loading">Loading…</div>`;
  try{
    const token = await onedriveGetToken(true);
    // lastModifiedDateTime/webUrl added for the Modified column and the
    // "Open in OneDrive" action (both new in the full-page redesign) --
    // both are ordinary DriveItem properties Graph already returns for
    // folders as well as files, no extra request needed.
    const select = '$select=id,name,folder,file,size,lastModifiedDateTime,webUrl,%40microsoft.graph.downloadUrl';
    // Graph's $orderby only accepts primitive fields -- "folder" is a
    // complex facet object ({childCount:N}), not a primitive, so sorting
    // by it (to put folders first) throws a 400 BadRequest, confirmed by
    // a real request: "The $orderby expression must evaluate to a single
    // value of primitive type." Sort by name here (Graph's own job); the
    // column-sortable table below does its own client-side ordering on
    // top of whatever Graph returns, including folders-first regardless
    // of the chosen column.
    const order = '$orderby=name';
    const url = cur.id===null
      ? `https://graph.microsoft.com/v1.0/me/drive/root:/${ONEDRIVE_ROOT_PATH.split('/').map(encodeURIComponent).join('/')}:/children?${select}&${order}&$top=200`
      : `https://graph.microsoft.com/v1.0/me/drive/items/${encodeURIComponent(cur.id)}/children?${select}&${order}&$top=200`;
    const res = await fetch(url, { headers:{ Authorization:'Bearer '+token } });
    if(!res.ok){
      // Same reasoning as the sign-in error fix above -- Graph's own
      // error body (e.g. "itemNotFound: The resource could not be
      // found" for a wrong/renamed folder path, or "accessDenied") is
      // far more useful than a bare HTTP status for actually diagnosing
      // a real failure, and there's no DevTools access to fall back on.
      let detail = `HTTP ${res.status}`;
      try{ const errBody = await res.json(); if(errBody && errBody.error) detail += ` — ${errBody.error.code}: ${errBody.error.message}`; }catch(e){}
      throw new Error(detail);
    }
    const data = await res.json();
    onedriveListState.rawItems = data.value||[];
    onedriveListState.filterText = ''; // a filter is contextual to the folder it was typed in
    onedriveRenderFolderView();
  }catch(err){
    console.error(err);
    // Retry alone strands anyone signed in with the wrong Microsoft
    // account (e.g. a different personal account on another device) --
    // retrying just re-fails with the same account every time, and this
    // was the only OneDrive error screen with no way to sign out and
    // switch accounts. Sign out is offered alongside Retry here too.
    body.innerHTML = `<div class="onedrive-error">Could not load this folder — ${esc(err.message||String(err))}<div class="onedrive-error-actions"><button type="button" class="onedrive-retry-btn" onclick="onedriveLoadCurrentFolder()">Retry</button><button type="button" class="onedrive-signout-btn" onclick="onedriveSignOut()">Sign out</button></div></div>`;
  }
}
window.onedriveLoadCurrentFolder = onedriveLoadCurrentFolder;
function onedriveOpenFolderFromEl(el){
  onedriveFolderStack.push({ id: el.dataset.id, name: el.dataset.name });
  onedriveLoadCurrentFolder();
}
window.onedriveOpenFolderFromEl = onedriveOpenFolderFromEl;
function onedriveGoTo(index){
  onedriveFolderStack = onedriveFolderStack.slice(0, index+1);
  onedriveLoadCurrentFolder();
}
window.onedriveGoTo = onedriveGoTo;
// Explicit Back button (Alok's request), alongside the breadcrumb rather
// than instead of it -- pops exactly one level, same as clicking the
// breadcrumb's second-to-last entry, disables itself at HATHRAS (the
// scoped root) since there's nowhere to go back to from there.
function onedriveGoBack(){
  if(onedriveFolderStack.length<=1) return;
  onedriveFolderStack.pop();
  onedriveLoadCurrentFolder();
}
window.onedriveGoBack = onedriveGoBack;
// Applies the live filter text and current sort to the folder's raw
// items, shaping each into the plain-object form applySort()/the row
// template expect (same {key: value} convention as every other sortable
// list in this app -- see resultListState/acctListState elsewhere).
function onedriveVisibleItems(){
  const q = onedriveListState.filterText.trim().toLowerCase();
  const raw = q ? onedriveListState.rawItems.filter(it=>String(it.name||'').toLowerCase().includes(q)) : onedriveListState.rawItems;
  const shaped = raw.map(it=>({
    id: it.id, name: it.name, isFolder: !!it.folder,
    size: typeof it.size==='number' ? it.size : null,
    modified: it.lastModifiedDateTime || null,
    childCount: it.folder && it.folder.childCount!==undefined ? it.folder.childCount : null,
    downloadUrl: it['@microsoft.graph.downloadUrl'] || null,
    webUrl: it.webUrl || null,
  }));
  const sorted = applySort(shaped, onedriveListState.sort);
  // Folders lead regardless of the chosen sort column -- the same
  // convention Explorer and OneDrive's own web app use -- via a second
  // stable pass on top of applySort()'s own stable sort, so each group
  // keeps whatever order the column just gave it.
  return sorted.slice().sort((a,b)=> a.isFolder===b.isFolder ? 0 : (a.isFolder?-1:1));
}
// Google Drive-style file icons: one flat folder glyph, plus a page
// glyph colour-coded by extension (red=PDF, green=spreadsheet, blue=doc,
// orange=slides, purple=image, grey=archive/other) so files are
// scannable by icon colour/shape alone, the same convention Drive and
// OneDrive's own web apps both use -- not a copy of either's actual
// trademarked logo art, just the same colour-by-type idea.
const ONEDRIVE_EXT_CLASS = {
  pdf:'od-ic-pdf',
  doc:'od-ic-doc', docx:'od-ic-doc',
  xls:'od-ic-sheet', xlsx:'od-ic-sheet', csv:'od-ic-sheet',
  ppt:'od-ic-slide', pptx:'od-ic-slide',
  jpg:'od-ic-img', jpeg:'od-ic-img', png:'od-ic-img', gif:'od-ic-img', webp:'od-ic-img', svg:'od-ic-img', bmp:'od-ic-img',
  zip:'od-ic-zip', rar:'od-ic-zip', '7z':'od-ic-zip',
  txt:'od-ic-text',
};
function onedriveIconSvg(name, isFolder){
  if(isFolder) return `<svg class="od-ic od-ic-folder" width="20" height="20" viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M3 6.5A1.5 1.5 0 0 1 4.5 5h4.44a1.5 1.5 0 0 1 1.06.44l1.06 1.06a1.5 1.5 0 0 0 1.06.44H19.5A1.5 1.5 0 0 1 21 8.44v8.56a1.5 1.5 0 0 1-1.5 1.5h-15A1.5 1.5 0 0 1 3 17V6.5Z"/></svg>`;
  const ext = String(name||'').split('.').pop().toLowerCase();
  const cls = ONEDRIVE_EXT_CLASS[ext] || 'od-ic-generic';
  return `<svg class="od-ic ${cls}" width="20" height="20" viewBox="0 0 24 24" aria-hidden="true">
    <path fill="currentColor" fill-opacity=".16" d="M6 2h7l5 5v13a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2Z"/>
    <path fill="currentColor" d="M13 2v5h5z"/>
    <path fill="none" stroke="currentColor" stroke-width="1.6" d="M6 2h7l5 5v13a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2Z"/>
  </svg>`;
}
function onedriveRowsHtml(items){
  if(!items.length){
    const msg = onedriveListState.filterText.trim() ? 'No items match your filter.' : 'This folder is empty.';
    return `<tr><td colspan="4" class="onedrive-empty-cell">${esc(msg)}</td></tr>`;
  }
  return items.map(it=>{
    const modifiedTxt = it.modified ? fmtDate(new Date(it.modified)) : '—';
    if(it.isFolder){
      return `<tr class="onedrive-row onedrive-folder" data-id="${esc(it.id)}" data-name="${esc(it.name)}" onclick="onedriveOpenFolderFromEl(this)">
        <td class="tal onedrive-name-cell">
          ${onedriveIconSvg(it.name, true)}
          <span class="onedrive-name">${esc(it.name)}</span>
        </td>
        <td class="tal">${modifiedTxt}</td>
        <td>${it.childCount!==null?it.childCount+' item'+(it.childCount===1?'':'s'):'—'}</td>
        <td class="onedrive-actions-cell"></td>
      </tr>`;
    }
    return `<tr class="onedrive-row onedrive-file">
      <td class="tal onedrive-name-cell">
        ${onedriveIconSvg(it.name, false)}
        <span class="onedrive-name">${esc(it.name)}</span>
      </td>
      <td class="tal">${modifiedTxt}</td>
      <td>${onedriveFmtSize(it.size)}</td>
      <td class="onedrive-actions-cell">
        ${it.webUrl ? `<a class="onedrive-action-btn" href="${esc(it.webUrl)}" target="_blank" rel="noopener" title="Open in OneDrive" aria-label="Open ${esc(it.name)} in OneDrive"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" aria-hidden="true"><path d="M18 13v6a2 2 0 01-2 2H5a2 2 0 01-2-2V8a2 2 0 012-2h6"/><path d="M15 3h6v6"/><path d="M10 14L21 3"/></svg></a>` : ''}
        ${it.downloadUrl ? `<a class="onedrive-action-btn onedrive-dl-btn" href="${esc(it.downloadUrl)}" target="_blank" rel="noopener" title="Download" aria-label="Download ${esc(it.name)}"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" aria-hidden="true"><path d="M12 3v12m0 0l-4-4m4 4l4-4"/><path d="M4 19h16"/></svg></a>` : ''}
      </td>
    </tr>`;
  }).join('');
}
// Grid view (Drive's card-tile layout) -- same shaped items as the list,
// just rendered as tiles with a bigger icon and the name/meta below it,
// instead of table rows. Folders and files share one grid; folders still
// sort ahead per onedriveVisibleItems().
function onedriveGridHtml(items){
  if(!items.length){
    const msg = onedriveListState.filterText.trim() ? 'No items match your filter.' : 'This folder is empty.';
    return `<div class="onedrive-empty-cell">${esc(msg)}</div>`;
  }
  return `<div class="onedrive-grid">${items.map(it=>{
    const modifiedTxt = it.modified ? fmtDate(new Date(it.modified)) : '';
    const meta = it.isFolder
      ? (it.childCount!==null ? it.childCount+' item'+(it.childCount===1?'':'s') : '')
      : [onedriveFmtSize(it.size), modifiedTxt].filter(Boolean).join(' · ');
    const openAttrs = it.isFolder ? ` data-id="${esc(it.id)}" data-name="${esc(it.name)}" onclick="onedriveOpenFolderFromEl(this)"` : '';
    return `<div class="onedrive-card ${it.isFolder?'onedrive-folder':'onedrive-file'}"${openAttrs} tabindex="0" role="${it.isFolder?'button':'group'}" aria-label="${esc(it.name)}">
      <div class="onedrive-card-icon">${onedriveIconSvg(it.name, it.isFolder)}</div>
      <div class="onedrive-card-name" title="${esc(it.name)}">${esc(it.name)}</div>
      <div class="onedrive-card-meta">${esc(meta)}</div>
      ${!it.isFolder ? `<div class="onedrive-card-actions">
        ${it.webUrl ? `<a class="onedrive-action-btn" href="${esc(it.webUrl)}" target="_blank" rel="noopener" title="Open in OneDrive" aria-label="Open ${esc(it.name)} in OneDrive" onclick="event.stopPropagation()"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" aria-hidden="true"><path d="M18 13v6a2 2 0 01-2 2H5a2 2 0 01-2-2V8a2 2 0 012-2h6"/><path d="M15 3h6v6"/><path d="M10 14L21 3"/></svg></a>` : ''}
        ${it.downloadUrl ? `<a class="onedrive-action-btn onedrive-dl-btn" href="${esc(it.downloadUrl)}" target="_blank" rel="noopener" title="Download" aria-label="Download ${esc(it.name)}" onclick="event.stopPropagation()"><svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" aria-hidden="true"><path d="M12 3v12m0 0l-4-4m4 4l4-4"/><path d="M4 19h16"/></svg></a>` : ''}
      </div>` : ''}
    </div>`;
  }).join('')}</div>`;
}
// Renders just the list/grid content area (table or card grid, per
// onedriveListState.view) -- called on every filter keystroke, sort
// click, and view-toggle click, never re-fetching from Graph and never
// touching the toolbar/filter <input> itself (so typing never loses
// focus/cursor position, and the view toggle's own active state survives
// a content-only re-render).
function onedriveRenderListArea(){
  const area = document.getElementById('onedriveListArea');
  if(!area) return;
  const items = onedriveVisibleItems();
  if(onedriveListState.view==='grid'){
    area.innerHTML = onedriveGridHtml(items);
  }else{
    area.innerHTML = `<div class="onedrive-table-wrap">
      <table class="onedrive-list">
        <thead id="onedrivePageHead"><tr>
          <th class="tal sortable" data-key="name" tabindex="0" role="button" aria-sort="none" onclick="onedriveSortBy('name')">Name<span class="sort-ic">▾</span></th>
          <th class="tal sortable" data-key="modified" tabindex="0" role="button" aria-sort="none" onclick="onedriveSortBy('modified')">Last modified<span class="sort-ic">▾</span></th>
          <th class="sortable" data-key="size" tabindex="0" role="button" aria-sort="none" onclick="onedriveSortBy('size')">File size<span class="sort-ic">▾</span></th>
          <th class="onedrive-col-actions" aria-hidden="true"></th>
        </tr></thead>
        <tbody id="onedrivePageRows">${onedriveRowsHtml(items)}</tbody>
      </table>
    </div>`;
    updateSortIcons('onedrivePageHead', onedriveListState.sort);
  }
}
function onedriveSortBy(key){
  onedriveListState.sort = nextSort(onedriveListState.sort, key);
  onedriveRenderListArea();
}
window.onedriveSortBy = onedriveSortBy;
function onedriveFilterInput(value){
  onedriveListState.filterText = value;
  onedriveRenderListArea();
}
window.onedriveFilterInput = onedriveFilterInput;
// List/Grid toggle (Google Drive's signature control) -- swaps only the
// content area, leaves the toolbar (breadcrumb/search/Back/sign-out)
// untouched, and remembers the choice in localStorage the same way the
// theme toggle does.
function onedriveSetView(view){
  if(onedriveListState.view===view) return;
  onedriveListState.view = view;
  try{ localStorage.setItem('upgb-onedrive-view', view); }catch(e){}
  document.querySelectorAll('.onedrive-view-btn').forEach(b=>b.classList.toggle('active', b.dataset.view===view));
  onedriveRenderListArea();
}
window.onedriveSetView = onedriveSetView;
// Full render (toolbar + list/grid area) -- called only after a genuine
// folder fetch, since the breadcrumb/Back-button state actually changes
// then; filtering, sorting, and switching view within an already-loaded
// folder go through onedriveRenderListArea() above instead, which never
// rebuilds this toolbar shell.
function onedriveRenderFolderView(){
  const canGoBack = onedriveFolderStack.length>1;
  const crumbSep = '<svg class="onedrive-crumb-sep" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" aria-hidden="true"><path d="M9 6l6 6-6 6"/></svg>';
  const breadcrumb = onedriveFolderStack.map((f,i)=>{
    const last = i===onedriveFolderStack.length-1;
    return `<span class="onedrive-crumb${last?' current':''}"${last?'':` onclick="onedriveGoTo(${i})"`}>${esc(f.name)}</span>`;
  }).join(crumbSep);
  const view = onedriveListState.view;
  document.getElementById('onedrivePageBody').innerHTML = `
    <div class="onedrive-toolbar">
      <button type="button" class="onedrive-back-btn" onclick="onedriveGoBack()" ${canGoBack?'':'disabled'} aria-label="Back one folder">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.3" aria-hidden="true"><path d="M15 18l-6-6 6-6"/></svg>
      </button>
      <div class="onedrive-breadcrumb">${breadcrumb}</div>
      <div class="onedrive-toolbar-spacer"></div>
      <div class="onedrive-search-pill">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
        <input type="text" placeholder="Search in this folder" value="${esc(onedriveListState.filterText)}" oninput="onedriveFilterInput(this.value)" aria-label="Filter files in this folder">
      </div>
      <div class="onedrive-view-toggle" role="group" aria-label="View">
        <button type="button" class="onedrive-view-btn${view==='list'?' active':''}" data-view="list" onclick="onedriveSetView('list')" title="List view" aria-label="List view">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" aria-hidden="true"><line x1="4" y1="6" x2="20" y2="6"/><line x1="4" y1="12" x2="20" y2="12"/><line x1="4" y1="18" x2="20" y2="18"/></svg>
        </button>
        <button type="button" class="onedrive-view-btn${view==='grid'?' active':''}" data-view="grid" onclick="onedriveSetView('grid')" title="Grid view" aria-label="Grid view">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" aria-hidden="true"><rect x="4" y="4" width="7" height="7" rx="1.5"/><rect x="13" y="4" width="7" height="7" rx="1.5"/><rect x="4" y="13" width="7" height="7" rx="1.5"/><rect x="13" y="13" width="7" height="7" rx="1.5"/></svg>
        </button>
      </div>
      <button type="button" class="onedrive-icon-btn" onclick="onedriveLoadCurrentFolder()" title="Refresh" aria-label="Refresh this folder">
        <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" aria-hidden="true"><path d="M21 12a9 9 0 11-2.64-6.36"/><polyline points="21 3 21 9 15 9"/></svg>
      </button>
      <button type="button" class="onedrive-signout-btn" onclick="onedriveSignOut()">Sign out</button>
    </div>
    <div id="onedriveListArea"></div>`;
  onedriveRenderListArea();
}
/* Sol ID -> ledger account number prefix, live as Alok types. Each of the
   6 accounts in the Account Numbers panel is XXXX + a fixed 10-digit
   suffix in the source ledger -- the XXXX stands in for whichever
   branch's own Sol ID it actually is. Typing a Sol ID into the small
   input next to the heading fills those 4 placeholder characters in
   live, one digit at a time (a colour pulse on each account marks the
   change, so it reads like the digits are typing themselves in);
   clearing the input reverts every account back to the XXXX placeholder. */
function initLedgerAcctNumbers(){
  document.querySelectorAll('#ledgerAccountsEdgePanel .ledger-acct-no').forEach(el=>{
    if(el.dataset.acctSuffix===undefined) el.dataset.acctSuffix = el.textContent.slice(4);
  });
}
function onLedgerSolIdInput(value){
  const digits = String(value||'').replace(/\D/g,'').slice(0,4);
  const input = document.getElementById('ledgerSolIdInput');
  if(input && input.value!==digits) input.value = digits;
  const prefix = (digits + 'XXXX').slice(0,4);
  document.querySelectorAll('#ledgerAccountsEdgePanel .ledger-acct-no').forEach(el=>{
    el.textContent = prefix + (el.dataset.acctSuffix||'');
    el.classList.remove('pulse'); void el.offsetWidth; el.classList.add('pulse');
  });
}
window.onLedgerSolIdInput = onLedgerSolIdInput;
// Recovery Dashboard (branch portal): the Sol ID used to always start
// blank (XXXX placeholders) and had to be typed in by hand every visit --
// pointless here since the whole portal is already locked to one Sol ID
// at login. Alok, 2026-09-25: "accounts slidebar main xxxx bydefault jo
// sol id open hai wahi set kar do" -- prefill it with the logged-in Sol ID
// by default; still a plain editable input, so it can be overridden for a
// one-off lookup on a different branch's ledger accounts.
(function prefillLedgerSolId(){
  const solId = loggedInSolId();
  if(solId) onLedgerSolIdInput(solId);
})();
initLedgerAcctNumbers();
document.addEventListener('keydown', (e)=>{
  if(e.key==='Escape'){ EDGE_PANEL_KEYS.forEach(k=>toggleEdgePanel(k, false)); }
});

updateReportDateDisplay();

/* ---------- Core formula engine (1:1 with the OTS sheet) ---------- */
// Extracted out of computeUCI so the UCI tenure shown next to "UCI @ 8.5%"
// on screen (uciTenureTag(), below) can compute the exact same anchor date
// without duplicating this scheme-dependent rule a second time.
function uciAnchorDate(npaDateRaw, scheme){
  const npaDate = toDate(npaDateRaw);
  if(!npaDate) return null;
  if(scheme==='CC004'){
    const y = npaDate.getFullYear();
    const sep24=new Date(y,8,24), mar24=new Date(y,2,24);
    return npaDate>sep24?sep24:(npaDate>mar24?mar24:new Date(y-1,8,24));
  }
  const eom = endOfMonth(npaDate);
  return sameDate(npaDate,eom) ? new Date(npaDate.getFullYear(),npaDate.getMonth(),29) : endOfMonth(new Date(npaDate.getFullYear(),npaDate.getMonth()-1,1));
}
function computeUCI(os, npaDateRaw, scheme, rate){
  rate = rate===undefined ? 8.5 : rate;
  // Was `if(!os || !npaDateRaw)` -- os===0 is a real, valid O/S Balance
  // (a fully-recovered-but-not-yet-declassified NPA, say), not "blank", but
  // `!os` treats 0 the same as '' (the actual blank sentinel used
  // elsewhere), so it silently returned '' for a zero O/S account. That
  // then propagated into computeSlot(): Total P&L came out ₹0 (correct,
  // since it doesn't gate on uci) while Total Dues came out "—" instead of
  // ₹0 (wrong), and the account was then silently dropped from every
  // Total-Dues-based aggregate even though it's a perfectly real account.
  if(os==null || os==='' || !npaDateRaw) return '';
  const anchor = uciAnchorDate(npaDateRaw, scheme);
  if(!anchor) return '';
  return os*rate/100*(daysBetween(new Date(),anchor)/365);
}
/* "UCI @ 8.5% (anchor date to today)" -- Alok's request, in the row's own
   HEADING, not appended to each account's figure. The row label is one
   shared cell across every account/column though (th.lt-label, not a
   per-column th), so a genuinely per-account tenure isn't representable
   there when a borrower's linked accounts have different NPA dates (and
   therefore different anchors) -- picks the first account that actually
   has a computable UCI/anchor as the one the heading shows, which is
   exact for the overwhelmingly common single-account case and a
   reasonable, clearly-labelled-as-one-figure approximation otherwise. */
function uciLabelWithTenure(slots){
  const s = (slots||[]).find(x=>x.uci!=='');
  const anchor = s ? uciAnchorDate(s.npaDate, s.scheme) : null;
  return anchor ? `UCI @ 8.5% (${fmtDate(anchor)} to ${fmtDate(new Date())})` : 'UCI @ 8.5%';
}
/* Row -> loan-slot shape. Split out of lookupLoanSlot so the OTS
   Worksheet can build the same slot straight from an account number,
   without needing the borrower's cust ID and slot position first. */
function slotFromRow(row){
  if(!row) return null;
  return {
    acctNo: row[C.ACCT_NO], scheme: row[C.SCHEME]||'', sanctionDate: row[C.SANCT_DT]||'',
    sanctionLimit: row[C.SANCT_LIM]===''?'':row[C.SANCT_LIM], assetCode: row[C.ASSET]||'',
    npaDate: row[C.NPA_DT]||'', osBalance: row[C.OUTBAL]===''?'':row[C.OUTBAL], uri: row[C.URI]===''?0:row[C.URI],
  };
}
function lookupLoanSlot(custId, slotNo){
  return slotFromRow(npaByHelper.get(custId+':'+slotNo));
}
function computeSlot(slot){
  if(!slot) return null;
  const today = new Date();
  const npaDate = toDate(slot.npaDate);
  const daysNpa = npaDate ? daysBetween(today, npaDate) : '';
  const os = typeof slot.osBalance==='number' ? slot.osBalance : '';
  const uri = typeof slot.uri==='number' ? slot.uri : 0;
  const uci = os!=='' ? computeUCI(os, slot.npaDate, slot.scheme, 8.5) : '';
  const uci125 = os!=='' ? computeUCI(os, slot.npaDate, slot.scheme, 12.5) : '';
  // Total Dues = O/S + UCI@8.5% + Interest Reversal.
  const totalDues = (os!=='' && uci!=='') ? os+uci+uri : '';
  // Total Contractual Dues = O/S + UCI@12.5% + Interest Reversal -- same
  // Interest Reversal that folds into Total Dues above, previously left
  // out here entirely (Alok: typing Interest Reversal updated Total Dues
  // but not Total Contractual Dues).
  const totalContractualDues = (os!=='' && uci125!=='') ? os+uci125+uri : '';
  // Net O/S is always identical to O/S Balance -- not a separate figure --
  // so Provision is calculated directly on O/S Balance (by asset code).
  const netOutstanding = os;
  let provision = '';
  if(os!=='' && PROV_RATES[slot.assetCode]!==undefined) provision = os*PROV_RATES[slot.assetCode];
  // Total P&L = O/S - Provision (Interest Reversal already flows into
  // Total Dues/Total Sacrifice above, not into Total P&L).
  const totalPL = (os!==''&&provision!=='') ? os-provision : '';
  const eligibleCompromise = totalPL!=='' ? Math.max(0,totalPL) : '';
  const ratio = (eligibleCompromise!=='' && os) ? eligibleCompromise/os : '';
  const notEligible = (daysNpa!=='' && daysNpa<=180);
  return {...slot, daysNpa, os, uri, uci, uci125, totalDues, totalContractualDues, netOutstanding, provision, totalPL, eligibleCompromise, ratio, notEligible};
}

/* ---------- Search ---------- */
const SEARCH_MODES = [
  {id:'acct', label:'Account No.', col:C.ACCT_NO, ph:'e.g. 160835110000679'},
  {id:'cust', label:'Cust ID', col:C.CUST_ID, ph:'e.g. 700962400'},
  {id:'mobile', label:'Mobile No.', col:C.PHONE, ph:'e.g. 9876543210'},
  {id:'aadhar', label:'Aadhar No.', col:C.AADHAR, ph:'e.g. 913206620914'},
  {id:'pan', label:'PAN', col:C.PAN, ph:'e.g. BJAPV4204K'},
  {id:'sb', label:'SB No.', col:C.SB_ACCT, ph:'e.g. 152910100005105'},
];
let searchMode = 'acct';
let __lastSearchMatches = null, __lastSearchMode = null;
const pillsEl = document.getElementById('modePills');
const searchInputEl = document.getElementById('searchInput');
SEARCH_MODES.forEach(m=>{
  const b = document.createElement('button');
  b.textContent = m.label; b.dataset.mode = m.id;
  if(m.id===searchMode) b.classList.add('active');
  b.onclick = ()=>{
    searchMode=m.id;
    pillsEl.querySelectorAll('button').forEach(x=>x.classList.toggle('active',x===b));
    searchInputEl.placeholder = m.ph;
    updateBranchPeek(searchInputEl.value.trim()); // hides if leaving Account No., shows if arriving with digits already typed
    if(searchInputEl.value.trim()) runSearch(); else renderEmpty();
  };
  pillsEl.appendChild(b);
});

const searchInput = document.getElementById('searchInput');
const clearBtn = document.getElementById('clearBtn');
// Live search: results appear as the account no./name/etc is typed, no
// need to press Enter or tap Search first -- but only once 6+ characters
// are in, so it doesn't try to match against a handful of stray digits.
// No cap on match count either -- the full match set renders every time.
let __liveSearchTimer = null;
searchInput.addEventListener('input', ()=>{
  clearBtn.style.display = searchInput.value ? 'flex' : 'none';
  clearTimeout(__liveSearchTimer);
  const q = searchInput.value.trim();
  updateBranchPeek(q);
  if(!q){ renderEmpty(); return; }
  if(q.length<6) return; // wait for at least 6 characters before suggesting
  __liveSearchTimer = setTimeout(()=>runSearch(), 160);
});
searchInput.addEventListener('keydown', e=>{ if(e.key==='Enter'){ clearTimeout(__liveSearchTimer); runSearch(); } });
function clearSearch(){ searchInput.value=''; clearBtn.style.display='none'; clearTimeout(__liveSearchTimer); hideBranchPeek(); renderEmpty(); }

/* ---------- Branch-prefix peek (Alok's request, 2026-09-08) ----------
   "jab main pahle 2 digit type karun to... branches k name right corner
   main show hon... jab 4 digit se jyada type kar dun to wo na dikhe" --
   purely informational, nothing to click: while searching by Account No.,
   typing 2-4 digits shows which branches' old Sol ID (BRANCH_LIST's first
   element, e.g. 16010) starts with those digits, live-narrowing on every
   keystroke, gone the moment a 5th digit goes in or the field empties.
   Deliberately scoped to searchMode==='acct' only -- a Sol ID prefix has
   no meaning for Cust ID/Mobile/Aadhar/PAN/SB No. searches. */
const BRANCH_PEEK_MIN = 2, BRANCH_PEEK_MAX = 4, BRANCH_PEEK_LIMIT = 10;
function updateBranchPeek(raw){
  const digits = raw.replace(/\D/g,'');
  if(searchMode!=='acct' || digits.length<BRANCH_PEEK_MIN || digits.length>BRANCH_PEEK_MAX){
    hideBranchPeek();
    return;
  }
  const matches = BRANCH_LIST.filter(([oldId])=>String(oldId).startsWith(digits)).map(([oldId,,name])=>({oldId, name}));
  if(!matches.length){ hideBranchPeek(); return; }
  renderBranchPeek(digits, matches);
}
function hideBranchPeek(){
  const panel = document.getElementById('branchPeekPanel');
  if(!panel) return;
  panel.classList.remove('show');
  panel.setAttribute('aria-hidden','true');
}
function renderBranchPeek(digits, matches){
  const panel = document.getElementById('branchPeekPanel');
  if(!panel) return;
  const shown = matches.slice(0, BRANCH_PEEK_LIMIT);
  const extra = matches.length - shown.length;
  // Sol ID sits in its own static span, never scrambled -- Alok asked for
  // it to be "clearly visible", so it stays crisp throughout the decode
  // animation instead of dissolving into glyphs along with the name.
  panel.innerHTML = `<div class="bp-head">SOL ${esc(digits)}<span class="bp-cursor">▌</span></div>`
    + shown.map((m,i)=>`<div class="bp-row" data-i="${i}"><span class="bp-sol">${esc(m.oldId)}</span><span class="bp-sep">–</span><span class="bp-name"></span></div>`).join('')
    + (extra>0 ? `<div class="bp-more">+${extra} more</div>` : '');
  panel.classList.add('show');
  panel.setAttribute('aria-hidden','false');
  repositionBranchPeekForKeyboard(); // the panel is only ever shown while actively typing, so the on-screen keyboard is almost always up on mobile right now
  panel.querySelectorAll('.bp-row').forEach((row,i)=>{
    scrambleInto(row.querySelector('.bp-name'), shown[i].name, {duration:280, delay:i*40});
  });
}
// Generic "decode" text reveal -- characters resolve left-to-right out of
// a scrambled glyph set, classic terminal/hacker-console effect. Used only
// for the branch peek above for now, kept generic in case something else
// wants the same treatment later. Respects prefers-reduced-motion like
// every other animation in this app (see animateNumber()).
const BRANCH_PEEK_GLYPHS = '!<>-_\\/[]{}=+*^?#$%01';
function scrambleInto(el, text, opts){
  opts = opts || {};
  if(__reduceMotion){ el.textContent = text; return; }
  const duration = opts.duration || 280, delay = opts.delay || 0;
  const startAt = performance.now() + delay;
  function frame(now){
    if(now<startAt){ el.__peekRaf = requestAnimationFrame(frame); return; }
    const t = Math.min(1, (now-startAt)/duration);
    const lockCount = Math.ceil(t*text.length);
    let out = '';
    for(let i=0;i<text.length;i++){
      out += (i<lockCount || text[i]===' ') ? text[i] : BRANCH_PEEK_GLYPHS[(Math.random()*BRANCH_PEEK_GLYPHS.length)|0];
    }
    el.textContent = out;
    if(t<1) el.__peekRaf = requestAnimationFrame(frame); else el.textContent = text;
  }
  if(el.__peekRaf) cancelAnimationFrame(el.__peekRaf);
  el.__peekRaf = requestAnimationFrame(frame);
}
// Mobile fix (Alok's report, screenshot showed the panel gone once the
// on-screen keyboard opened): the CSS below anchors the mobile panel to a
// fixed distance from the bottom of the viewport, which is exactly where
// the keyboard sits -- and the panel is only ever visible while actively
// typing digits, i.e. almost always with the keyboard already up. A fixed
// pixel guess can't account for how tall any given device's keyboard
// actually is, so this reads the real gap from visualViewport (the area
// the keyboard covers) and pushes the panel up above it live, resizing as
// the keyboard opens/closes/changes (e.g. switching to a suggestions bar).
const BRANCH_PEEK_MOBILE_MQ = window.matchMedia('(max-width:640px)');
function repositionBranchPeekForKeyboard(){
  const panel = document.getElementById('branchPeekPanel');
  if(!panel) return;
  if(!BRANCH_PEEK_MOBILE_MQ.matches || !window.visualViewport){ panel.style.bottom = ''; return; }
  const covered = Math.max(0, window.innerHeight - window.visualViewport.height - window.visualViewport.offsetTop);
  panel.style.bottom = (88 + covered) + 'px'; // 88px clears the bottom tab bar when no keyboard is up
}
if(window.visualViewport){
  window.visualViewport.addEventListener('resize', repositionBranchPeekForKeyboard);
  window.visualViewport.addEventListener('scroll', repositionBranchPeekForKeyboard);
}

// Result-list sort state, same shape as acctListState below (list of
// plain {acctNo,name,...} objects rather than raw NPA rows, so the
// existing generic applySort()/nextSort()/updateSortIcons() machinery --
// already used by the Dashboard's All Accounts table and the Branch/Sol
// ID list modals -- can drive this table too instead of a bespoke sort).
let resultListState = {list:[], sort:{key:'name',dir:'asc'}};
function renderResultsTable(){
  const sorted = applySort(resultListState.list, resultListState.sort);
  updateSortIcons('resultsHead', resultListState.sort);
  const tbody = document.getElementById('resultsBody');
  if(tbody) tbody.innerHTML = resultRowsHtml(sorted);
}
// Alok's request, 2026-09-05: "yahan bhi sort ka filter de do account and
// customer par... z to a like" -- clickable Account/Customer headers,
// toggling ascending/descending on repeat clicks the same way every other
// sortable table in this app already works.
function sortResultsBy(key){
  resultListState.sort = nextSort(resultListState.sort, key);
  renderResultsTable();
}
window.sortResultsBy = sortResultsBy;
function runSearch(){
  const q = searchInput.value.trim().toLowerCase();
  if(!q){ renderEmpty(); return; }
  const mode = SEARCH_MODES.find(m=>m.id===searchMode);
  const seen = new Set();
  const matches = [];
  for(const r of DATA.npa.rows){
    const val = r[mode.col];
    if(val==='' || val===null) continue;
    if(String(val).toLowerCase().includes(q)){
      const cid = String(r[C.CUST_ID]);
      const key = mode.id==='acct' ? String(r[C.ACCT_NO]) : cid;
      if(seen.has(key)) continue;
      seen.add(key);
      matches.push(r);
    }
  }
  // Alok's original request -- results list reads more naturally sorted
  // A-Z by borrower name than in raw data order. Refined for the Account
  // No. search specifically (2026-09-05): typing digits of an account
  // number is scanning for a near-match among the results, so those default
  // to sorting by account number itself (numeric-aware -- see applySort()'s
  // acctNo handling below) rather than name, which is unrelated to what was
  // actually typed. Every other mode (Cust ID/Mobile/Aadhar/PAN/SB No.)
  // keeps the by-name default. A fresh search (new query or mode change)
  // always resets to this default, even if a header click had switched the
  // sort during the previous result set.
  resultListState.sort = {key: mode.id==='acct' ? 'acctNo' : 'name', dir:'asc'};
  renderResults(matches, mode);
}

/* ---------- OTS start screen (the Search tab before anything is searched)
   Replaces what used to be a bare icon + one line of text on an otherwise
   empty screen. Approved "Action Hub" layout: recently-opened borrowers
   first (the overwhelmingly common next action -- back to yesterday's
   account), then top branches by O/S, then a small portfolio line for
   context. Deliberately kept lighter than the Dashboard tab so it informs
   without duplicating it. ---------- */

/* Recently-opened borrowers, newest first, capped at RECENT_MAX. Stored
   only in this browser's localStorage -- never published, never sent
   anywhere -- so it stays per-person even though the app itself needs no
   login. Keyed by custId so re-opening the same borrower moves it back to
   the top instead of adding a duplicate row. */
const RECENT_KEY = 'upgb-recent-borrowers';
const RECENT_MAX = 50;
function getRecentBorrowers(){
  try{
    const raw = JSON.parse(localStorage.getItem(RECENT_KEY) || '[]');
    return Array.isArray(raw) ? raw.filter(r=>r && r.custId).slice(0,RECENT_MAX) : [];
  }catch(e){ return []; }
}
function rememberBorrower(custRow){
  if(!custRow) return;
  const custId = String(custRow[C.CUST_ID]||'');
  if(!custId) return;
  /* O/S is summed across the borrower's linked loan accounts, the same way
     openDetail() builds its slots -- storing only custRow's own balance
     would under-report a multi-account household (e.g. showing one loan's
     38k for a borrower whose two loans total 1.19 L). */
  const slots = [1,2,3,4].map(n=>lookupLoanSlot(custId,n)).filter(Boolean);
  const totalOs = slots.reduce((a,s)=>a+(typeof s.osBalance==='number'?s.osBalance:0),0);
  const entry = {
    custId,
    acctNo: String(custRow[C.ACCT_NO]||''),
    name: custRow[C.NAME]||'',
    branch: custRow[C.SOL_DESC]||'',
    asset: custRow[C.ASSET]||'',
    os: slots.length ? totalOs : (typeof custRow[C.OUTBAL]==='number' ? custRow[C.OUTBAL] : ''),
    n: slots.length,
  };
  try{
    const list = getRecentBorrowers().filter(r=>String(r.custId)!==custId);
    list.unshift(entry);
    localStorage.setItem(RECENT_KEY, JSON.stringify(list.slice(0,RECENT_MAX)));
  }catch(e){ /* private mode / quota -- recents are a convenience, not critical */ }
}
function initialsOf(name){
  const parts = String(name||'').trim().split(/\s+/).filter(Boolean);
  if(!parts.length) return '—';
  return ((parts[0][0]||'') + (parts.length>1 ? (parts[1][0]||'') : '')).toUpperCase();
}

function openRecentBorrower(custId, acctNo){
  openDetail(String(custId), acctNo ? String(acctNo) : undefined);
}
window.openRecentBorrower = openRecentBorrower;

/* Total of the OTS Amounts saved on this device for one borrower, summed
   across their linked loan accounts -- the same slots openDetail() builds,
   so this matches the "Total OTS Amount" the detail screen shows. Returns
   null when nothing has been entered for any of them yet. */
function savedOtsFor(custId){
  const slots = [1,2,3,4].map(n=>lookupLoanSlot(String(custId),n)).filter(Boolean);
  let total = 0, any = false;
  slots.forEach(s=>{
    const v = parseOtsAmount(otsAmounts[s.acctNo]);
    if(v!==null){ total += v; any = true; }
  });
  return any ? total : null;
}
/* ---------- OTS Worksheet ----------
   Every account this device has an OTS Amount saved for, in one table with
   running totals. The amounts were already being kept (and restored) per
   account, but they were only visible one borrower at a time -- this is
   the view that makes a whole settlement batch reviewable and exportable.
   Figures are recomputed from live data through the same computeSlot /
   totalDuesFor path the detail screen uses, so nothing here can drift from
   what that screen shows. */
function otsWorksheetRows(){
  const rows = [];
  Object.keys(otsAmounts).forEach(acctNo => {
    const ots = parseOtsAmount(otsAmounts[acctNo]);
    if(ots===null) return;
    const raw = npaByAcct.get(String(acctNo));
    if(!raw) return; // account no longer in the book (regularized/closed)
    const s = computeSlot(slotFromRow(raw));
    const totalDues = totalDuesFor(s);
    rows.push({
      acctNo: String(acctNo),
      custId: String(raw[C.CUST_ID]||''),
      name: raw[C.NAME]||'',
      branch: raw[C.SOL_DESC]||'',
      asset: s.assetCode||'',
      os: s.os===''?0:s.os,
      ots,
      sacrifice: totalDues==='' ? '' : totalDues-ots,
      impact: s.totalPL==='' ? '' : ots-s.totalPL,
    });
  });
  return rows.sort((a,b)=>b.os-a.os);
}
function otsWorksheetTotals(rows){
  const sum = k => rows.reduce((a,r)=>a+(typeof r[k]==='number'?r[k]:0),0);
  return { os:sum('os'), ots:sum('ots'), sacrifice:sum('sacrifice'), impact:sum('impact') };
}
function renderOtsWorksheet(){
  const rows = otsWorksheetRows();
  const t = otsWorksheetTotals(rows);
  document.getElementById('wsSub').textContent = rows.length
    ? `${rows.length} account(s) with an OTS Amount saved on this device`
    : 'No OTS Amount has been entered yet';
  const body = document.getElementById('wsBody');
  const foot = document.getElementById('wsFoot');
  const sum = document.getElementById('wsSum');
  if(!rows.length){
    body.innerHTML = `<tr><td colspan="9" style="text-align:center;color:var(--sub);padding:26px 10px">
      Open a borrower and type an OTS Amount — every account you work out will be listed here.</td></tr>`;
    foot.innerHTML = '';
    sum.innerHTML = '';
    return;
  }
  sum.innerHTML = [
    ['O/S Balance', fmtINR2(t.os), ''],
    ['OTS Amount', fmtINR2(t.ots), 'ws-ots'],
    ['Total Sacrifice', fmtINR2(t.sacrifice), ''],
    ['Impact on P&amp;L', (t.impact>0?'+':(t.impact<0?'−':''))+fmtINR2(Math.abs(t.impact)),
      t.impact>0?'ws-pos':(t.impact<0?'ws-neg':'')],
  ].map(([lbl,val,cls])=>`<div class="ws-sum-tile">
      <span class="ws-sum-lbl">${lbl}</span>
      <span class="ws-sum-val ${cls}">${val}</span>
    </div>`).join('');
  body.innerHTML = rows.map(r=>`<tr class="clickable" onclick="closeOtsWorksheet();openDetail('${esc(r.custId)}','${esc(r.acctNo)}')">
    <td>${esc(r.acctNo)}</td>
    <td class="tal">${esc(r.name)||'—'}</td>
    <td class="tal">${esc(r.branch)||'—'}</td>
    <td>${r.asset?`<span class="badge-pill ${esc(r.asset)}">${esc(r.asset)}</span>`:'—'}</td>
    <td>${fmtINR2(r.os)}</td>
    <td class="ws-ots">${fmtINR2(r.ots)}</td>
    <td>${r.sacrifice===''?'—':fmtINR2(r.sacrifice)}</td>
    <td class="${r.impact===''?'':(r.impact>0?'ws-pos':(r.impact<0?'ws-neg':''))}">${r.impact===''?'—':(r.impact>0?'+':(r.impact<0?'−':''))+fmtINR2(Math.abs(r.impact))}</td>
    <td><button type="button" class="ws-del" title="Remove this account's saved OTS Amount"
      aria-label="Remove saved OTS Amount for account ${esc(r.acctNo)}"
      onclick="event.stopPropagation();removeSavedOts('${esc(r.acctNo)}')">✕</button></td>
  </tr>`).join('');
  foot.innerHTML = `<tr class="ws-total">
    <td colspan="4" class="tal">Total — ${rows.length} account(s)</td>
    <td>${fmtINR2(t.os)}</td>
    <td class="ws-ots">${fmtINR2(t.ots)}</td>
    <td>${fmtINR2(t.sacrifice)}</td>
    <td class="${t.impact>0?'ws-pos':(t.impact<0?'ws-neg':'')}">${(t.impact>0?'+':(t.impact<0?'−':''))+fmtINR2(Math.abs(t.impact))}</td>
    <td></td>
  </tr>`;
}
function openOtsWorksheet(){
  renderOtsWorksheet();
  document.getElementById('wsModalOverlay').classList.add('show');
}
function closeOtsWorksheet(){ document.getElementById('wsModalOverlay').classList.remove('show'); }
window.openOtsWorksheet = openOtsWorksheet;
window.closeOtsWorksheet = closeOtsWorksheet;
function removeSavedOts(acctNo){
  delete otsAmounts[acctNo];
  delete interestReversalOverrides[acctNo];
  saveOtsAmounts(); saveUriOverrides();
  renderOtsWorksheet();
  renderEmpty();
}
window.removeSavedOts = removeSavedOts;
/* Built with ExcelJS, not SheetJS, for the same reason the single-borrower
   export is (see the note above exportOtsExcel): the free SheetJS build
   writes number formats but silently drops fonts and borders, and this
   sheet is meant to be handed to a branch or filed, not just read on
   screen. Same plain treatment as that export -- bold, real borders, no
   fill colour -- and the same XL_* constants, so the two sheets look like
   they came from one system. */
async function exportOtsWorksheet(){
  const rows = otsWorksheetRows();
  if(!rows.length){ alert('There is no saved OTS Amount to export yet.'); return; }
  await Promise.all([ensureXLSX(), ensureExcelJS()]);
  const t = otsWorksheetTotals(rows);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('OTS Worksheet', { views: [{showGridLines:false, state:'frozen', ySplit:5}] });
  const set = (addr, value, opts={}) => {
    const cell = ws.getCell(addr);
    cell.value = value;
    if(opts.numFmt) cell.numFmt = opts.numFmt;
    if(opts.font) cell.font = opts.font;
    if(opts.align) cell.alignment = opts.align;
    if(opts.border!==false) cell.border = opts.border || XL_BORDER_ALL;
    return cell;
  };

  ws.columns = [
    {width:20},{width:34},{width:18},{width:10},{width:17},{width:17},{width:17},{width:17},
  ];
  ws.mergeCells('A1:H1');
  set('A1', 'UPGB HATHRAS — OTS WORKSHEET',
    {font:{bold:true, size:16}, align:{horizontal:'center'}, border:false});
  ws.mergeCells('A2:H2');
  set('A2', `Uttar Pradesh Gramin Bank (Regional Office Hathras) · Data as on ${fmtAsOnDisplay()} · Prepared ${fmtDateTime(new Date())}`,
    {font:{size:10, color:{argb:'FF333333'}}, align:{horizontal:'center'},
     border:{bottom:{style:'medium', color:{argb:'FF555555'}}}});
  ws.getRow(1).height = 26;

  const headerRow = 5;
  ['Account No.','Customer','Branch','Asset','O/S Balance','OTS Amount','Total Sacrifice','Impact on P&L']
    .forEach((h,i)=>{
      set(XLSX.utils.encode_col(i)+headerRow, h,
        {font:{bold:true}, align:{horizontal:i<4?'left':'right', vertical:'middle', wrapText:true}});
    });
  ws.getRow(headerRow).height = 26;

  rows.forEach((r,i)=>{
    const n = headerRow + 1 + i;
    set(`A${n}`, String(r.acctNo), {align:{horizontal:'left'}});
    set(`B${n}`, r.name, {align:{horizontal:'left'}});
    set(`C${n}`, r.branch, {align:{horizontal:'left'}});
    set(`D${n}`, r.asset, {align:{horizontal:'left'}});
    set(`E${n}`, r.os===''?null:r.os, {numFmt:XL_INR_FMT});
    set(`F${n}`, r.ots, {numFmt:XL_INR_FMT});
    /* Sacrifice and Impact are written as live formulas off the same row's
       O/S and OTS cells, so a settlement amount edited in Excel updates the
       two derived columns and the totals -- the way the single-borrower
       export already behaves. The constants they need (Total Dues and Total
       P&L, which no column on this sheet carries) are folded into the
       formula as the row's own difference, so nothing silently goes stale.
       Deliberately NOT rounded to the paisa here (that used to happen, "to
       match fmtINR2 on screen") -- the on-screen Worksheet total instead
       sums every row's raw, full-precision Sacrifice/Impact and rounds only
       that one final sum for display. Pre-rounding each row's constant
       before Excel's own SUM() ran on them made the exported grand total a
       sum-of-rounded-values rather than a round-of-the-sum, drifting a few
       paisa-to-rupees from the on-screen total on a large book. Passing the
       raw values through (numFmt still rounds how they're DISPLAYED, just
       not what's stored) makes both totals agree exactly. */
    const dues = r.sacrifice==='' ? null : r.ots + r.sacrifice;
    const pl   = r.impact===''    ? null : r.ots - r.impact;
    set(`G${n}`, dues===null ? null : {formula:`(${dues})-F${n}`}, {numFmt:XL_INR_FMT});
    set(`H${n}`, pl===null   ? null : {formula:`F${n}-(${pl})`},   {numFmt:XL_INR_FMT_PL});
  });

  const totalRow = headerRow + rows.length + 1;
  const first = headerRow + 1, last = headerRow + rows.length;
  ws.mergeCells(`A${totalRow}:D${totalRow}`);
  set(`A${totalRow}`, `TOTAL — ${rows.length} ACCOUNT(S)`,
    {font:{bold:true}, align:{horizontal:'left'}});
  ['E','F','G'].forEach(col=>{
    set(`${col}${totalRow}`, {formula:`SUM(${col}${first}:${col}${last})`},
      {numFmt:XL_INR_FMT, font:{bold:true}});
  });
  set(`H${totalRow}`, {formula:`SUM(H${first}:H${last})`},
    {numFmt:XL_INR_FMT_PL, font:{bold:true}});

  const noteRow = totalRow + 2;
  ws.mergeCells(`A${noteRow}:H${noteRow}`);
  set(`A${noteRow}`, 'Total Sacrifice and Impact on P&L recalculate from the OTS Amount in column F. Figures are as on the data date above.',
    {font:{italic:true, size:9, color:{argb:'FF666666'}}, border:false});

  ws.pageSetup = {
    paperSize: 9, orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0,
    horizontalCentered: true, printTitlesRow: `${headerRow}:${headerRow}`,
    margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
    printArea: `A1:H${noteRow}`,
  };

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = `UPGB_OTS_Worksheet_${dateToInputValue(new Date())}.xlsx`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url), 30000);
}
/* The button's onclick can't await, so a rejection here would surface only
   as an unhandled promise in the console -- the user would just see nothing
   download. This turns that into a message they can act on. */
window.exportOtsWorksheet = () => exportOtsWorksheet().catch(err=>{
  console.error(err);
  alert('The worksheet could not be exported. Please try again.');
});

/* ---------- Backup / restore of this device's own OTS work ----------
   Everything the app saves per-person lives in this browser's
   localStorage, so clearing browser data or moving to another phone loses
   it. These two put that work in a file the user holds. Nothing is sent
   anywhere -- the file is written and read locally. */
const OTS_BACKUP_VERSION = 1;
function backupOtsWork(){
  const payload = {
    app: 'upgb-ots', version: OTS_BACKUP_VERSION,
    exportedAt: new Date().toISOString(),
    otsAmounts, uriOverrides: interestReversalOverrides,
    recents: getRecentBorrowers(),
  };
  const blob = new Blob([JSON.stringify(payload, null, 2)], {type:'application/json'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = `UPGB_OTS_Backup_${dateToInputValue(new Date())}.json`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url), 30000);
}
window.backupOtsWork = backupOtsWork;
function restoreOtsWork(evt){
  const file = evt.target.files[0];
  if(!file) return;
  evt.target.value = ''; // let the same file be picked again after a failed try
  const reader = new FileReader();
  reader.onerror = () => alert('Could not read that file.');
  reader.onload = e => {
    let data;
    try{ data = JSON.parse(String(e.target.result)); }
    catch(err){ alert('That file is not a valid backup — it could not be read as JSON.'); return; }
    const isMap = v => v && typeof v==='object' && !Array.isArray(v);
    if(!isMap(data) || data.app!=='upgb-ots' || !isMap(data.otsAmounts)){
      alert('That file is not a UPGB OTS backup.');
      return;
    }
    const n = Object.keys(data.otsAmounts).length;
    if(!confirm(`Restore ${n} saved OTS Amount(s) from this backup?\n\nThis replaces what is currently saved on this device.`)) return;
    otsAmounts = data.otsAmounts;
    interestReversalOverrides = isMap(data.uriOverrides) ? data.uriOverrides : {};
    saveOtsAmounts(); saveUriOverrides();
    if(Array.isArray(data.recents)){
      try{ localStorage.setItem(RECENT_KEY, JSON.stringify(data.recents.slice(0,RECENT_MAX))); }catch(err){}
    }
    renderOtsWorksheet();
    renderEmpty();
    alert(`Restored ${n} saved OTS Amount(s).`);
  };
  reader.readAsText(file);
}
window.restoreOtsWork = restoreOtsWork;

function clearRecentBorrowers(){
  // Only the visited-list is dropped. Saved OTS Amounts are keyed by
  // account, not by this list, and are real work -- they stay.
  try{ localStorage.removeItem(RECENT_KEY); }catch(e){}
  renderEmpty();
}
window.clearRecentBorrowers = clearRecentBorrowers;

/* Drops a single borrower from the visited-list. Same rule as the Clear
   button above: the list is a convenience, so removing a row never touches
   that borrower's saved OTS Amount -- that stays in the worksheet. */
function removeRecentBorrower(custId){
  try{
    const list = getRecentBorrowers().filter(r=>String(r.custId)!==String(custId));
    localStorage.setItem(RECENT_KEY, JSON.stringify(list));
  }catch(e){}
  renderEmpty();
}
window.removeRecentBorrower = removeRecentBorrower;

/* The way into the worksheet, and -- when nothing is saved yet -- the way
   to a backup file, which is exactly the state a fresh device is in. That
   is why the bar renders in both cases instead of only when work exists. */
function otsWorksheetBarHtml(){
  const rows = otsWorksheetRows();
  const t = otsWorksheetTotals(rows);
  const sub = rows.length
    ? `${rows.length} account(s) · O/S ${fmtCr(t.os)} · OTS ${fmtCr(t.ots)}`
    : 'Nothing saved on this device yet — open to restore a backup';
  return `
      <button type="button" class="start-ws${rows.length?'':' is-empty'}" onclick="openOtsWorksheet()">
        <span class="start-ws-ic" aria-hidden="true">
          <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.9"><rect x="3.5" y="3" width="17" height="18" rx="2.5"/><line x1="3.5" y1="9" x2="20.5" y2="9"/><line x1="9.5" y1="9" x2="9.5" y2="21"/></svg>
        </span>
        <span class="start-ws-txt">
          <span class="start-ws-nm">OTS Worksheet</span>
          <span class="start-ws-sub">${esc(sub)}</span>
        </span>
        <span class="start-ws-go">Open</span>
      </button>`;
}

function renderEmpty(){
  const mode = SEARCH_MODES.find(m=>m.id===searchMode);
  const recents = getRecentBorrowers();

  // Before anything has been opened there is no list to show, so the screen
  // carries the search hint instead of an empty heading.
  if(!recents.length){
    document.getElementById('mainArea').innerHTML = `
      <div class="ots-start">
        <div class="start-hint">
          <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
          <span>Search by <b>${esc(mode.label)}</b> above. Borrowers you open will be listed here for quick access.</span>
        </div>
        ${otsWorksheetBarHtml()}
      </div>`;
    return;
  }

  document.getElementById('mainArea').innerHTML = `
    <div class="ots-start">
      ${otsWorksheetBarHtml()}
      <div class="start-block">
        <div class="start-head">
          <span class="start-lbl">Recently Opened</span>
          <button type="button" class="start-clear" onclick="clearRecentBorrowers()">Clear</button>
        </div>
        ${recents.map(r=>{
          const ots = savedOtsFor(r.custId);
          const nm = esc(r.name)||'—';
          return `
        <div class="start-rec-row">
          <button type="button" class="start-rec" onclick="openRecentBorrower('${esc(r.custId)}','${esc(r.acctNo||'')}')">
            <span class="start-rec-av">${esc(initialsOf(r.name))}</span>
            <span class="start-rec-txt">
              <span class="start-rec-nm">${nm}</span>
              <span class="start-rec-sub">${esc(r.branch)||'—'}${r.asset?' · '+esc(r.asset):''}${r.n>1?' · '+r.n+' accounts':''}</span>
            </span>
            <span class="start-rec-figs">
              <span class="start-rec-amt">${r.os===''?'—':fmtCr(r.os)}</span>
              ${ots!==null ? `<span class="start-rec-ots">OTS ${fmtCr(ots)}</span>` : ''}
            </span>
          </button>
          <button type="button" class="start-rec-del" title="Remove from Recently Opened"
            aria-label="Remove ${nm} from Recently Opened"
            onclick="removeRecentBorrower('${esc(r.custId)}')">✕</button>
        </div>`;
        }).join('')}
      </div>
    </div>`;
}

// Compact, table-style result list -- matches the "All Accounts" list
// already used on the KCC Overdue / PNPA tabs (Account,
// Customer, Branch, Asset, O/S Balance), so the search behaves the same
// way as every other account list in the app: type -> a plain scrollable
// list of matches -> tap a row -> the full OTS Calculator detail opens.
function resultRowsHtml(list){
  return list.map(c=>`<tr class="clickable" onclick="openDetail('${esc(c.custId)}','${esc(c.acctNo)}')">
      <td>${esc(c.acctNo)}</td>
      <td class="tal">${esc(c.name)||'—'}</td>
      <td class="tal">${esc(c.branch)||'—'}</td>
      <td>${c.asset?`<span class="badge-pill ${esc(c.asset)}" title="${esc(assetLabel(c.asset))}">${esc(c.asset)}</span>`:'—'}</td>
      <td>${fmtINR2(c.os)}</td>
    </tr>`).join('');
}
function renderResults(matches, mode){
  __lastSearchMatches = matches; __lastSearchMode = mode;
  const el = document.getElementById('mainArea');
  if(!matches.length){
    el.innerHTML = `<div class="ots-results"><div class="results-hint">0 matches found</div>` +
      `<div class="no-results">` +
      `<svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" aria-hidden="true"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>` +
      `<div>No borrower matches that ${esc(mode.label)}.<br>Try a different value or search mode.</div></div></div>`;
    return;
  }
  resultListState.list = matches.map(r=>({
    acctNo: String(r[C.ACCT_NO]),
    name: r[C.NAME]||'',
    branch: r[C.SOL_DESC]||'',
    asset: r[C.ASSET]||'',
    os: typeof r[C.OUTBAL]==='number' ? r[C.OUTBAL] : 0,
    custId: String(r[C.CUST_ID]),
  }));
  const sorted = applySort(resultListState.list, resultListState.sort);
  /* .ots-results carries the brass tokens, the same way .ots-start and the
     hero card above it do -- without it the tab reads brass at the top,
     sapphire through the result list, then brass again on the detail
     screen the list leads into. Account/Customer headers are sortable
     (Alok's request) via the same generic applySort()/nextSort() engine
     already driving the Dashboard's All Accounts table and the Branch/Sol
     ID list modals -- clicking either toggles ascending/descending the
     same way those do, keyboard-activatable too via the shared th.sortable
     Enter/Space handler registered once, near applySort() itself. */
  el.innerHTML = `<div class="ots-results">` +
    `<div class="results-hint-row">` +
      `<div class="results-hint">${matches.length} match${matches.length>1?'es':''} found</div>` +
      `<button type="button" class="export-xl-btn" onclick="exportSearchResults()">${EXPORT_XL_ICON} Export to Excel</button>` +
    `</div>` +
    `<div class="dash-table-wrap acct-list-scroll">
      <table class="dash-table">
        <thead id="resultsHead"><tr>
          <th class="sortable" data-key="acctNo" tabindex="0" role="button" aria-sort="none" onclick="sortResultsBy('acctNo')">Account<span class="sort-ic">▾</span></th>
          <th class="tal sortable" data-key="name" tabindex="0" role="button" aria-sort="none" onclick="sortResultsBy('name')">Customer<span class="sort-ic">▾</span></th>
          <th class="tal">Branch</th><th>Asset</th><th>O/S Balance</th>
        </tr></thead>
        <tbody id="resultsBody">${resultRowsHtml(sorted)}</tbody>
      </table>
    </div></div>`;
  updateSortIcons('resultsHead', resultListState.sort);
}
/* "What's on screen right now" export -- resultListState.list + its
   current sort, the exact same array resultRowsHtml() just rendered, so
   the file always matches what the banker is looking at. */
function exportSearchResults(){
  const sorted = applySort(resultListState.list, resultListState.sort);
  if(!sorted.length) return;
  const rows = sorted.map(c=>[c.acctNo, c.name, c.branch, c.asset, c.os]);
  exportRowsToExcel(
    `Search_Results_${dateToInputValue(new Date())}.xlsx`, 'Search Results',
    ['Account No','Customer','Branch','Asset','O/S Balance'], rows,
    [null,null,null,null,XL_INR_FMT]
  );
  showToast(`✓ ${sorted.length} row${sorted.length>1?'s':''} exported`);
}
window.exportSearchResults = exportSearchResults;

/* ---------- Detail view ----------
   Typed OTS Amounts and Interest Reversal overrides are kept in this
   device's own localStorage, so a settlement being worked out survives a
   reload, a phone restart, or coming back the next day. Nothing here is
   ever published or sent anywhere -- it stays on the one device it was
   typed on, and each person's working figures stay their own.

   Interest Reversal is persisted alongside the OTS Amount deliberately:
   it feeds Total Dues, which feeds Total Sacrifice, so restoring one
   without the other would show a different sacrifice figure than the one
   on screen when the account was last left. */
const OTS_AMOUNTS_KEY = 'upgb-ots-amounts';
const URI_OVERRIDES_KEY = 'upgb-uri-overrides';
function loadStoredMap(key){
  try{
    const raw = JSON.parse(localStorage.getItem(key) || '{}');
    return (raw && typeof raw === 'object' && !Array.isArray(raw)) ? raw : {};
  }catch(e){ return {}; }
}
function persistStoredMap(key, map){
  try{ localStorage.setItem(key, JSON.stringify(map)); }
  catch(e){ /* private mode / quota -- on-screen values still work */ }
}
let otsAmounts = loadStoredMap(OTS_AMOUNTS_KEY);          // key: acctNo -> typed OTS Amount
let interestReversalOverrides = loadStoredMap(URI_OVERRIDES_KEY); // key: acctNo -> typed Interest Reversal
function saveOtsAmounts(){ persistStoredMap(OTS_AMOUNTS_KEY, otsAmounts); }
function saveUriOverrides(){ persistStoredMap(URI_OVERRIDES_KEY, interestReversalOverrides); }
// Resolves the live Interest Reversal for a slot: the user's typed override
// if present, else the value loaded from the daily NPA data.
function uriFor(s){
  const raw = interestReversalOverrides[s.acctNo];
  if(raw===undefined) return s.uri;
  const v = parseFloat(raw);
  return (raw===''||isNaN(v)) ? 0 : v;
}
// Total Dues = O/S + UCI@8.5% + Interest Reversal -- computed live off
// uriFor() so it reacts to the editable field. Total P&L (O/S - Provision)
// does NOT depend on Interest Reversal, so it stays a static computeSlot()
// value and needs no live helper.
function totalDuesFor(s){
  return (s.os!=='' && s.uci!=='') ? s.os + s.uci + uriFor(s) : '';
}
// Total Contractual Dues = O/S + UCI@12.5% + Interest Reversal -- same
// live uriFor() dependency as totalDuesFor() above, so typing an Interest
// Reversal override updates this figure too, not just Total Dues.
function totalContractualDuesFor(s){
  return (s.os!=='' && s.uci125!=='') ? s.os + s.uci125 + uriFor(s) : '';
}

function openDetail(custId, jumpAcct){
  const custRow = byCustId.get(custId);
  if(!custRow) return;
  rememberBorrower(custRow);
  switchView('search');
  const slots = [1,2,3,4].map(n=>{
    const s = lookupLoanSlot(custId, n);
    return s ? computeSlot(s) : null;
  }).filter(Boolean);
  const prevOts = oldOtsByAcct.get(String(custRow[C.ACCT_NO]));

  const pane = document.getElementById('detailPane');
  document.getElementById('shell').classList.add('detail-active');
  pane.classList.add('open');
  pane.innerHTML = `
    <div class="detail-head">
      <div class="detail-headrow">
        <button class="back-btn" onclick="closeDetail()" aria-label="Back to search results">
          <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><polyline points="15 18 9 12 15 6"/></svg>
        </button>
        <div class="detail-headtext">
          <h2>${esc(custRow[C.NAME])||'—'}</h2>
          <p>${esc(custRow[C.SOL_DESC])||''} · Cust ID ${esc(custRow[C.CUST_ID])}</p>
        </div>
        <div class="rm-scale-badge">
          <span class="rm-scale-title">RM SCALE - V</span>
          <span class="rm-scale-tenure-group">
            <span class="rm-scale-tenure">29-03-2021 to 31-03-2022</span>
            <span class="rm-scale-tenure">23-09-2024 to 11-09-2026</span>
          </span>
        </div>
        <button class="share-btn" onclick="exportOtsExcel()" title="Export to Excel (live formulas — edit OTS Amount and everything else recalculates)" aria-label="Export to Excel with formulas">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><rect x="3" y="3" width="18" height="18" rx="2"/><path d="M8 3v18M16 3v18M3 9h18M3 15h18"/></svg>
        </button>
        <button class="share-btn" onclick="printOtsSheet()" title="Print / Share" aria-label="Print or share this report">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M4 12v8a2 2 0 002 2h12a2 2 0 002-2v-8"/><polyline points="16 6 12 2 8 6"/><line x1="12" y1="2" x2="12" y2="15"/></svg>
        </button>
        <button class="share-btn" onclick="toggleShareOtsMenu(event)" title="Share on WhatsApp" aria-label="Share this report on WhatsApp">
          <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 11.5a8.38 8.38 0 01-.9 3.8 8.5 8.5 0 01-7.6 4.7 8.38 8.38 0 01-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 01-.9-3.8 8.5 8.5 0 014.7-7.6 8.38 8.38 0 013.8-.9h.5a8.48 8.48 0 018 8v.5z"/></svg>
        </button>
      </div>
    </div>
    <div class="detail-inner${slots.length>=1?' has-agg':''}">
      ${slots.length>=1?`<aside id="aggBar" aria-label="Account totals">
        <div class="agg-title">${slots.length>1?`All ${slots.length} Accounts`:'This Account'}</div>
        <div class="agg-hero">
          <div class="agg-hero-top">
            <div>
              <div class="agg-hero-label">Net Settlement Impact</div>
              <div class="agg-hero-value" id="aggTotImpact">—</div>
              <div class="agg-hero-sub" id="aggHeroSub">—</div>
            </div>
            <div class="agg-hero-ring" id="aggHeroRing" style="--pct:0"><span id="aggHeroRingPct">—</span></div>
          </div>
          <div class="agg-hero-pcts">
            <div class="agg-hero-pct-chip"><span class="k">of Total Dues</span><span class="v" id="aggPctDues">—</span></div>
            <div class="agg-hero-pct-chip"><span class="k">of O/S Balance</span><span class="v" id="aggPctOs">—</span></div>
          </div>
        </div>
        <div class="agg-mini-grid">
          <div class="agg-mini"><div class="k">Total OTS Amount</div><div class="v" id="aggTotOts">—</div></div>
          <div class="agg-mini"><div class="k">Total O/S Balance</div><div class="v" id="aggTotNetOs">—</div></div>
          <div class="agg-mini"><div class="k">Total P&amp;L</div><div class="v" id="aggTotPL">—</div></div>
          <div class="agg-mini"><div class="k">Total Sacrifice</div><div class="v" id="aggTotSac">—</div></div>
        </div>
        <div class="agg-scale">
          <div class="agg-block-head">${ltIcon('gauge')}Recovery Scale</div>
          <div class="agg-scale-track">
            <div class="agg-band loss"></div>
            <div class="agg-band safe" id="aggBandSafe"></div>
            <div class="agg-needle" id="aggNeedle"><span class="agg-needle-val" id="aggNeedleVal">—</span></div>
          </div>
          <div class="agg-scale-labels">
            <div class="agg-slab be" id="aggLabBE">Break-even<b id="aggBEVal">—</b></div>
            <div class="agg-slab" id="aggLabOS">O/S<b id="aggOSVal">—</b></div>
            <div class="agg-slab" id="aggLabDues">Dues<b id="aggDuesVal">—</b></div>
          </div>
        </div>
        <div class="agg-wf collapsed" id="aggWfBlock">
          <div class="agg-block-head agg-wf-toggle" onclick="toggleAggWf()" onkeydown="if(event.key==='Enter'||event.key===' '){event.preventDefault();toggleAggWf();}" role="button" tabindex="0" aria-expanded="false" aria-controls="aggWfBody">
            ${ltIcon('list')}Where The Dues Go
            <svg class="agg-wf-chev" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" aria-hidden="true"><polyline points="6 9 12 15 18 9"/></svg>
          </div>
          <div class="agg-wf-body" id="aggWfBody">
            <div class="agg-wf-bar">
              <span id="aggWf1"></span><span id="aggWf2"></span><span id="aggWf3"></span>
            </div>
            <div class="agg-wf-key">
              <div><span class="agg-wf-dot" style="background:#1B2A44"></span>Recovered in cash (OTS)<b id="aggWfCash">—</b></div>
              <div><span class="agg-wf-dot" style="background:#D4A544"></span>Ledger sacrifice (BDWO)<b id="aggWfLedger">—</b></div>
              <div><span class="agg-wf-dot" style="background:#7A8798"></span>Unrealised interest (UCI)<b id="aggWfUci">—</b></div>
            </div>
          </div>
        </div>
      </aside>`:''}
      <div id="detailBody" style="padding-top:14px"></div>
    </div>
  `;
  drawDetailBody(custRow, slots, prevOts);
  pane.scrollTop = 0;
}

// "Where The Dues Go" starts collapsed on mobile (Alok's request -- it was
// eating too much of the fixed bottom dock) but stays permanently expanded
// on desktop, where the sidebar has room -- the .collapsed class only has
// any visual effect inside the mobile media query in styles.css.
function toggleAggWf(){
  const block = document.getElementById('aggWfBlock');
  if(!block) return;
  const collapsed = block.classList.toggle('collapsed');
  const head = block.querySelector('.agg-wf-toggle');
  if(head) head.setAttribute('aria-expanded', String(!collapsed));
}
window.toggleAggWf = toggleAggWf;

function closeDetail(){
  const pane = document.getElementById('detailPane');
  pane.classList.remove('open');
  pane.innerHTML = '';
  document.getElementById('shell').classList.remove('detail-active');
  document.getElementById('railLeft').classList.remove('show');
  document.getElementById('railRight').classList.remove('show');
  document.getElementById('eligibleBanner').classList.remove('show');
  document.getElementById('specialNoteBanner')?.classList.remove('show');
  /* Coming back from a borrower, the start screen behind it is stale -- the
     visit just entered Recently Opened, and any OTS Amount typed changes
     both that row and the worksheet bar's totals. Only redrawn when the
     start screen is what's showing; a result list is left as it was. */
  if(document.querySelector('#mainArea .ots-start')) renderEmpty();
}

function drawDetailBody(custRow, slots, prevOts){
  const body = document.getElementById('detailBody');
  const totalOS = slots.reduce((a,s)=>a+((s.os!=='')?s.os:0),0);
  const totalDues = slots.reduce((a,s)=>a+((s.totalDues!=='')?s.totalDues:0),0);
  const totalNetOS = slots.reduce((a,s)=>a+((s.netOutstanding!=='')?s.netOutstanding:0),0);
  const totalContractualDues = slots.reduce((a,s)=>a+((s.totalContractualDues!=='')?s.totalContractualDues:0),0);
  // Total P&L (O/S - Provision) no longer depends on Interest Reversal, so
  // it's a stable per-render snapshot again -- only Total Dues needs live
  // recomputation (recalcAggregate), since Interest Reversal folds into it.
  const totalPL = slots.reduce((a,s)=>a+((s.totalPL!=='')?s.totalPL:0),0);

  body.innerHTML = `
    <div class="card borrower-card">
      <div class="bcard-top">
        <div class="bavatar" aria-hidden="true"><svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="8" r="3.6"/><path d="M4.5 20c1.6-3.6 4.8-5.5 7.5-5.5s5.9 1.9 7.5 5.5"/></svg></div>
        <div class="bcard-title-col">
          <div class="bname">${esc(custRow[C.NAME])||'—'}</div>
          <div class="baddr">${esc(custRow[C.ADDR])||'—'}</div>
        </div>
        <div class="bcard-branch">${esc(custRow[C.SOL_DESC])||'—'}</div>
      </div>
      <div class="info-grid">
        <div><div class="k">Cust ID</div><div class="v">${esc(custRow[C.CUST_ID])||'—'}</div></div>
        <div><div class="k">Sol ID</div><div class="v">${esc(custRow[C.SOL_ID])||'—'}</div></div>
        <div><div class="k">Mobile</div><div class="v">${esc(custRow[C.PHONE])||'—'}</div></div>
        <div><div class="k">Aadhar</div><div class="v">${esc(custRow[C.AADHAR])||'—'}</div></div>
        <div><div class="k">PAN</div><div class="v">${esc(custRow[C.PAN])||'—'}</div></div>
        <div><div class="k">Branch</div><div class="v">${esc(custRow[C.SOL_DESC])||'—'}</div></div>
        <div><div class="k">SB A/C</div><div class="v">${esc(custRow[C.SB_ACCT])||'—'}</div></div>
        <div><div class="k">SB Balance</div><div class="v">${fmtINR2(custRow[C.SB_BAL]===''?0:custRow[C.SB_BAL])}</div></div>
      </div>
      ${prevOts?`<div class="linked-note">⏱ Previous OTS on record: ${esc(prevOts.date)} — ${esc(prevOts.amount)}</div>`:''}
      <div class="linked-note">🔗 ${slots.length} loan account${slots.length>1?'s':''} linked</div>
    </div>

    <div class="loans-col">
    <div class="section-label">Loan Accounts</div>
    <div class="section-sub">All accounts side-by-side · Enter OTS amount below</div>

    ${loanTableHTML(slots)}
  </div>
  `;

  const tableWrap = body.querySelector('.loan-table-wrap');
  if(tableWrap){
    const updateFade = () => tableWrap.classList.toggle('at-end', tableWrap.scrollLeft + tableWrap.clientWidth >= tableWrap.scrollWidth - 4);
    tableWrap.addEventListener('scroll', updateFade);
    updateFade();
  }

  window.__slots = slots;
  window.__totalDues = totalDues;
  window.__totalPL = totalPL;
  window.__totalNetOS = totalNetOS;
  window.__totalContractualDues = totalContractualDues;
  window.__totalOS = totalOS;
  window.__custRow = custRow;
  window.__prevOts = prevOts;

  slots.forEach((s,i)=>recalcLoan(i));
  recalcAggregate();

  const notEligibleAccts = slots.filter(s=>s.notEligible).map(s=>s.acctNo);
  const banner = document.getElementById('eligibleBanner');
  if(notEligibleAccts.length){
    document.getElementById('eligibleBannerText').textContent =
      `Not eligible — A/c ${notEligibleAccts.map(a=>esc(String(a))).join(', ')} NPA not aged 6 months`;
    banner.classList.add('show');
  } else {
    banner.classList.remove('show');
  }

  // Special Note -- Admin-authored, per Account No., set via Update Data ->
  // Special Note. Shown to every viewer, not just Admin, the same way the
  // "Not eligible" banner above is -- but it's informational (a hold, an
  // instruction, a reminder), not a warning, so it reads in brass rather
  // than red. Collected across every linked account on this borrower, not
  // just the one initially searched, since Alok tags a note by Account No.
  // and a borrower can have several.
  const notedSlots = slots
    .map(s => ({acctNo: s.acctNo, note: (DATA.specialNotes||{})[String(s.acctNo)]}))
    .filter(x => x.note && x.note.note);
  const noteBanner = document.getElementById('specialNoteBanner');
  if(noteBanner){
    if(notedSlots.length){
      const text = notedSlots.length===1
        ? notedSlots[0].note.note
        : notedSlots.map(x => `A/c ${x.acctNo}: ${x.note.note}`).join('  ·  ');
      document.getElementById('specialNoteBannerText').textContent = text;
      noteBanner.classList.add('show');
    } else {
      noteBanner.classList.remove('show');
    }
  }
  // Lok Adalat (see DATA.lokAdalat above). Same "collect across every
  // linked account" treatment as the Special Note banner just above --
  // but the Account No. itself is left out of the message (Alok's
  // request, 2026-09-11: "account no hata dena... wahi ac to open hai"),
  // since whichever account this is is already the one on screen. In the
  // rare case two DIFFERENT linked accounts on the same borrower both
  // match, their lines still show side by side without a label telling
  // them apart -- an accepted tradeoff for the common one-account case.
  const lokAdalatSlots = slots
    .map(s => ({acctNo: s.acctNo, hit: (DATA.lokAdalat||{})[String(s.acctNo)]}))
    .filter(x => x.hit);
  const lokAdalatBanner = document.getElementById('lokAdalatBanner');
  if(lokAdalatBanner){
    if(lokAdalatSlots.length){
      const oneLine = x => {
        const base = x.hit.ots
          ? `OTS ${fmtINR(x.hit.ots)} already received (Token ${fmtINR(x.hit.token)})`
          : `Token ${fmtINR(x.hit.token)} already received`;
        const dateBit = x.hit.date ? ` on ${x.hit.date}` : '';
        const remarkBit = x.hit.remark ? ` — ${x.hit.remark}` : '';
        return base + dateBit + remarkBit;
      };
      document.getElementById('lokAdalatBannerText').textContent = lokAdalatSlots.map(oneLine).join('  ·  ');
      lokAdalatBanner.classList.add('show');
    } else {
      lokAdalatBanner.classList.remove('show');
    }
  }
  positionBanners();
}
/* The "Not eligible", "Special Note", and "Lok Adalat" banners share the
   same fixed top-center spot (see .eligible-banner/.special-note-banner/
   .lok-adalat-banner in styles.css) so a single banner always lands dead
   center -- but each is an independent condition and more than one can be
   true for the same borrower at once. Rather than hard-coding fixed
   offsets (leaving an odd gap whenever an earlier one is hidden), measure
   each shown banner's actual rendered bottom edge, since its text (and
   therefore height) varies -- then stack the next one directly below it. */
function positionBanners(){
  // offsetHeight, not a sibling's getBoundingClientRect() after moving it --
  // .special-note-banner/.lok-adalat-banner animate `top` (transition:
  // top .2s), so reading a just-repositioned banner's rect immediately
  // after setting its top returns its stale pre-transition position, which
  // would throw off whichever banner stacks below it. Height alone isn't
  // animated, so accumulating a running offset from each banner's own
  // offsetHeight sidesteps the timing issue entirely.
  let top = 18; // matches the banners' own base CSS top:18px
  ['eligibleBanner','specialNoteBanner','lokAdalatBanner'].forEach(id => {
    const el = document.getElementById(id);
    if(!el) return;
    if(el.classList.contains('show')){
      el.style.top = top + 'px';
      top += el.offsetHeight + 10;
    } else {
      el.style.top = '';
    }
  });
}

// Small stroke-icon library for the loan table's row/section labels --
// purely visual (aria-hidden), makes a long particulars list scannable
// instead of reading like a plain spreadsheet.
const LT_ICONS = {
  loanTerms: '<rect x="3" y="4" width="18" height="17" rx="2"/><path d="M3 9h18M8 2v4M16 2v4"/>',
  dues: '<path d="M12 3 3 8l9 5 9-5-9-5Z"/><path d="M3 13l9 5 9-5" opacity=".55"/>',
  settlement: '<path d="M9 11 12 14l7-7"/><circle cx="12" cy="12" r="9"/>',
  calendar: '<rect x="3" y="4" width="18" height="17" rx="2"/><path d="M3 9h18M8 2v4M16 2v4"/>',
  doc: '<path d="M7 3h7l4 4v14H7z"/><path d="M14 3v4h4M9 12h6M9 16h6"/>',
  warn: '<path d="M12 3 2 20h20L12 3Z"/><path d="M12 10v4M12 17h.01"/>',
  coin: '<circle cx="12" cy="12" r="9"/><path d="M12 7v10M9 9.5h4.5a1.8 1.8 0 0 1 0 3.6H10a1.8 1.8 0 0 0 0 3.6H15"/>',
  rotate: '<path d="M4 12a8 8 0 0 1 14.9-4M20 12a8 8 0 0 1-14.9 4"/><path d="M18 4v4h-4M6 20v-4h4"/>',
  percent: '<circle cx="12" cy="12" r="9"/><path d="M9 15l6-6M9.5 9h.01M14.5 15h.01"/>',
  layers: '<path d="M12 3 3 8l9 5 9-5-9-5Z"/><path d="M3 13l9 5 9-5" opacity=".55"/>',
  shield: '<path d="M12 3l7 3v6c0 5-3.2 7.6-7 9-3.8-1.4-7-4-7-9V6l7-3Z"/>',
  trend: '<path d="M3 17l6-6 4 4 8-8M15 7h6v6"/>',
  badge: '<circle cx="12" cy="9" r="5"/><path d="M8.5 13.5 7 21l5-2.5L17 21l-1.5-7.5"/>',
  bars: '<path d="M4 20V10m6 10V4m6 16v-7"/>',
  list: '<path d="M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01"/>',
  tag: '<path d="M20.6 12.6 12.6 20.6a2 2 0 0 1-2.8 0l-7.4-7.4a2 2 0 0 1 0-2.8L10.4 2.4A2 2 0 0 1 11.8 2H18a2 2 0 0 1 2 2v6.2a2 2 0 0 1-.6 1.4Z"/><path d="M14 8h.01"/>',
  clock: '<circle cx="12" cy="12" r="9"/><path d="M12 7v5l3.5 2"/>',
  avatar: '<circle cx="12" cy="8" r="3.6"/><path d="M4.5 20c1.6-3.6 4.8-5.5 7.5-5.5s5.9 1.9 7.5 5.5"/>',
  gauge: '<path d="M12 3a9 9 0 0 0-7.6 13.9M12 3a9 9 0 0 1 7.6 13.9"/><path d="M12 12 16 8"/>',
  scale: '<path d="M12 3v18M5 7h14"/><path d="M5 7l-3 6a4 4 0 0 0 8 0z"/><path d="M19 7l-3 6a4 4 0 0 0 8 0z"/>',
};
function ltIcon(name, size){
  const s = size||13;
  return `<svg class="lt-row-icon" width="${s}" height="${s}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">${LT_ICONS[name]}</svg>`;
}
// Icon in a colored circular badge, matching the aggregate sidebar's
// per-stat treatment -- one consistent tone for every row (see
// .lt-icon-badge in styles.css) so the whole table reads as one coherent
// system instead of the sidebar's badges vs. the table's plain gray
// outline icons looking like two different designs.
function ltIconBadge(name, extraId){
  return `<span class="lt-icon-badge"${extraId?` id="${extraId}"`:''}>${ltIcon(name,10)}</span>`;
}

function loanTableHTML(slots){
  const cols = slots.map(s=>`
    <th scope="col">
      <div class="lt-acc">A/c · ${esc(s.acctNo)}</div>
      <div class="lt-scheme">${esc(s.scheme)||''}</div>
      ${s.assetCode?`<span class="badge-pill ${esc(s.assetCode)}" title="${esc(assetLabel(s.assetCode))}">${esc(s.assetCode)}</span>`:''}
    </th>`).join('');
  // Sticky, like every other row's label cell (th.lt-label) -- previously
  // a single colspan'd td, which isn't sticky, so the group heading text
  // (LOAN TERMS / DUES & PROVISIONING / SETTLEMENT & IMPACT) scrolled off
  // to the left as soon as the table was scrolled horizontally, while
  // every other row's label correctly stayed pinned in view.
  const group = (label, icon) => `<tr class="lt-group"><th scope="row" class="lt-label">${ltIcon(icon)}<span class="lt-label-text"><span class="lt-label-inner">${label}</span></span></th>${slots.map(()=>'<td></td>').join('')}</tr>`;
  const row = (label, icon, fn, cls='') => `<tr class="${cls}"><th scope="row" class="lt-label">${ltIconBadge(icon)}<span class="lt-label-text"><span class="lt-label-inner">${label}</span></span></th>${slots.map(s=>`<td>${fn(s)}</td>`).join('')}</tr>`;
  const statRow = (label, icon, idPrefix, iconId) => `<tr><th scope="row" class="lt-label">${ltIconBadge(icon,iconId)}<span class="lt-label-text"><span class="lt-label-inner">${label}</span></span></th>${slots.map((s,i)=>`<td id="${idPrefix}-${i}">—</td>`).join('')}</tr>`;
  const otsRow = () => `<tr class="lt-ots-row"><th scope="row" class="lt-label">${ltIconBadge('coin')}<span class="lt-label-text"><span class="lt-label-inner">Settlement (OTS) Amount</span></span></th>${slots.map((s,i)=>`
      <td><div class="lt-ots-cell">
        <span class="lt-cur">₹</span>
        <input type="number" class="lt-ots-input" id="otsInput-${i}" placeholder="0" value="${otsAmounts[s.acctNo]||''}"
          aria-label="OTS amount for account ${esc(String(s.acctNo))}"
          oninput="onOtsInput(${i},'${esc(String(s.acctNo))}')">
        <span class="pct-tag" id="pctNetOs-${i}"></span>
      </div></td>`).join('')}</tr>`;
  const uriRow = () => `<tr><th scope="row" class="lt-label">${ltIconBadge('rotate')}<span class="lt-label-text"><span class="lt-label-inner">Interest Reversal</span></span></th>${slots.map((s,i)=>`
      <td><div class="lt-ots-cell">
        <span class="lt-cur">₹</span>
        <input type="number" class="lt-ots-input" id="uriInput-${i}" placeholder="0" value="${uriFor(s)||''}"
          aria-label="Interest reversal for account ${esc(String(s.acctNo))}"
          oninput="onUriInput(${i},'${esc(String(s.acctNo))}')">
      </div></td>`).join('')}</tr>`;
  const totalDuesRow = () => `<tr class="lt-strong"><th scope="row" class="lt-label">${ltIconBadge('layers')}<span class="lt-label-text"><span class="lt-label-inner">Total Dues</span></span></th>${slots.map((s,i)=>`<td id="totalDues-${i}">—</td>`).join('')}</tr>`;
  const totalContractualDuesRow = () => `<tr class="lt-strong lt-divider"><th scope="row" class="lt-label">${ltIconBadge('layers')}<span class="lt-label-text"><span class="lt-label-inner">Total Contractual Dues</span></span></th>${slots.map((s,i)=>`<td id="totalContractualDues-${i}">—</td>`).join('')}</tr>`;
  // Settlement Progress: OTS Amount as a share of Total Dues, drawn as a
  // thin fill bar plus a printed percentage -- lets four accounts be
  // compared by eye instead of reading six-figure numbers column by column.
  const settleRow = () => `<tr><th scope="row" class="lt-label">${ltIconBadge('gauge')}<span class="lt-label-text"><span class="lt-label-inner">Settlement Progress</span></span></th>${slots.map((s,i)=>`<td id="settleCell-${i}"><span class="dash">—</span></td>`).join('')}</tr>`;
  const eligRow = slots.some(s=>s.notEligible) ? `<tr><th scope="row" class="lt-label"></th>${slots.map(s=>`<td>${s.notEligible?'<span class="eligibility-warn">⚠ Not aged 6mo</span>':''}</td>`).join('')}</tr>` : '';
  // Shown right above the actual OTS input -- the mandated floor, seen
  // before typing a proposed figure, not after. Substandard accounts
  // read "Not Eligible" rather than a rate, since Lok Adalat OTS simply
  // doesn't apply to them (Alok's own instruction) -- distinct from an
  // account whose Asset Code isn't in the circular at all, which
  // shouldn't come up in practice but falls back to "—" defensively.
  const lokAdalatRow = () => `<tr class="lt-ots-row lt-lokadalat-row"><th scope="row" class="lt-label">${ltIconBadge('scale')}<span class="lt-label-text"><span class="lt-label-inner">OTS Amt as per Lok Adalat</span></span></th>${slots.map(s=>{
    const la = lokAdalatMin(s);
    if(!la) return '<td>—</td>';
    if(!la.eligible) return '<td><span class="lt-lokadalat-na">Not Eligible</span></td>';
    return `<td>${fmtINR2(la.amount)} <span class="pct-tag">(${(la.pct*100).toFixed(0)}%)</span></td>`;
  }).join('')}</tr>`;

  return `
  <div class="loan-table-wrap">
  <table class="loan-table">
    <thead><tr><th scope="col" class="lt-label">${ltIcon('list')}<span class="lt-label-text"><span class="lt-label-inner">Particulars</span></span></th>${cols}</tr></thead>
    <tbody>
      ${eligRow}
      ${group('Loan Terms', 'loanTerms')}
      ${row('Sanction Date', 'calendar', s=>fmtDate(toDate(s.sanctionDate)))}
      ${row('Sanction Limit', 'doc', s=>fmtINR2(s.sanctionLimit))}
      ${row('NPA Date', 'warn', s=>fmtDate(toDate(s.npaDate)))}
      ${row('O/S Balance', 'coin', s=>fmtINR2(s.os), 'lt-strong')}
      ${group('Dues &amp; Provisioning', 'dues')}
      ${uriRow()}
      ${row(uciLabelWithTenure(slots), 'percent', s=>fmtINR2(s.uci))}
      ${totalDuesRow()}
      ${totalContractualDuesRow()}
      ${row('Provision', 'shield', s=>fmtINR2(s.provision))}
      ${row('Total P&amp;L', 'trend', s=>fmtINR2(s.totalPL) + (s.ratio!==''?` <span class="pct-tag">(${(s.ratio*100).toFixed(1)}%)</span>`:''), 'lt-strong lt-divider')}
      ${group('Settlement &amp; Impact', 'settlement')}
      ${lokAdalatRow()}
      ${otsRow()}
      ${settleRow()}
      ${statRow('Total Sacrifice', 'percent', 'totalSac')}
      ${statRow('Ledger Sacrifice (BDWO Amount)', 'badge', 'ledgerSac')}
      ${statRow('P&amp;L Impact', 'bars', 'impact')}
    </tbody>
  </table>
  </div>
  <div class="lt-hint">
    <svg width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 11v5M12 8h.01"/></svg>
    <span>Enter the OTS Amount and Interest Reversal for each account to calculate Total Sacrifice, Ledger Sacrifice, and P&amp;L Impact automatically.</span>
  </div>`;
}

// Parses a raw typed OTS Amount into a valid, non-negative number, or null
// if blank/invalid/negative. Centralizes what "a usable OTS Amount" means
// so every consumer (screen calc, aggregate, print, Excel) agrees, instead
// of each re-parsing the raw value slightly differently. A negative OTS
// Amount has no real-world meaning (a bank can't receive negative money in
// a settlement) -- previously nothing rejected it, so typing e.g. "-5000"
// flowed straight through into Total Sacrifice/Impact on P&L with no
// validation anywhere catching it.
function parseOtsAmount(raw){
  if(raw===undefined || raw==='') return null;
  const v = parseFloat(raw);
  return (isNaN(v) || v<0) ? null : v;
}

function onOtsInput(i, acctNo){
  const v = document.getElementById('otsInput-'+i).value;
  if(v==='') delete otsAmounts[acctNo]; else otsAmounts[acctNo] = v;
  saveOtsAmounts();
  recalcLoan(i);
  recalcAggregate();
}

function onUriInput(i, acctNo){
  const v = document.getElementById('uriInput-'+i).value;
  // A blank field means literal 0, same as typing "0" -- NOT "clear the
  // override and fall back to the master-data default". That fallback
  // used to be harmless because every account's master default was 0,
  // but broke visibly once real per-account Interest Reversal defaults
  // were seeded (2026-09-08): clearing the field silently brought back a
  // non-zero master value instead of 0, so Total Dues/Total Contractual
  // Dues looked like they weren't reacting to "set it to 0" at all.
  interestReversalOverrides[acctNo] = v==='' ? '0' : v;
  saveUriOverrides();
  recalcLoan(i);
  recalcAggregate();
}
window.onUriInput = onUriInput;

const __reduceMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
function animateNumber(el, from, to, render, dur){
  if(el.__raf) cancelAnimationFrame(el.__raf);
  if(__reduceMotion || from===to){ el.textContent = render(to); return; }
  dur = dur || 440;
  const start = performance.now();
  const step = (now)=>{
    const t = Math.min(1,(now-start)/dur);
    const e = 1-Math.pow(1-t,3);
    el.textContent = render(from + (to-from)*e);
    if(t<1){ el.__raf = requestAnimationFrame(step); } else { el.__raf = 0; }
  };
  el.__raf = requestAnimationFrame(step);
}

function recalcLoan(i){
  const s = window.__slots[i];
  const raw = otsAmounts[s.acctNo];
  const otsParsed = parseOtsAmount(raw);
  const ots = otsParsed===null ? '' : otsParsed;
  const totalSacEl = document.getElementById('totalSac-'+i);
  const ledgerEl = document.getElementById('ledgerSac-'+i);
  const impactEl = document.getElementById('impact-'+i);
  const pctEl = document.getElementById('pctNetOs-'+i);
  const totalDuesEl = document.getElementById('totalDues-'+i);
  const totalContractualDuesEl = document.getElementById('totalContractualDues-'+i);
  const settleCellEl = document.getElementById('settleCell-'+i);

  // Total Dues (O/S + UCI + Interest Reversal) and Total Contractual Dues
  // (O/S + UCI@12.5% + Interest Reversal) both depend on Interest Reversal,
  // not on OTS Amount, so both must update unconditionally -- even while
  // OTS Amount is still blank.
  const totalDues = totalDuesFor(s);
  if(totalDuesEl) totalDuesEl.textContent = fmtINR2(totalDues);
  if(totalContractualDuesEl) totalContractualDuesEl.textContent = fmtINR2(totalContractualDuesFor(s));

  if(ots===''||isNaN(ots)){
    [totalSacEl,ledgerEl,impactEl].forEach(e=>e.textContent='—');
    impactEl.classList.remove('pos','neg');
    impactEl.__val = 0;
    if(pctEl) pctEl.textContent='';
    if(settleCellEl) settleCellEl.innerHTML = '<span class="dash">—</span>';
    return;
  }
  if(settleCellEl){
    const settlePct = (totalDues && totalDues>0) ? Math.max(0,Math.min(100,(ots/totalDues)*100)) : 0;
    // Alok's request -- shows the O/S-based share alongside the Dues-based
    // bar/percentage that was already here, not just one or the other.
    const settlePctOs = (s.os && s.os>0) ? Math.max(0,(ots/s.os)*100) : null;
    const osLine = settlePctOs!==null ? ` · ${settlePctOs.toFixed(1)}% of O/S` : '';
    settleCellEl.innerHTML = `<div class="settle-cell"><div class="settle-bar"><span style="width:${settlePct.toFixed(1)}%"></span></div><span class="settle-pct">${settlePct.toFixed(1)}% of dues${osLine}</span></div>`;
  }
  // Total Sacrifice = Total Dues - OTS Amount (Interest Reversal is already
  // folded into Total Dues above); Ledger Sacrifice (BDWO Amount) = O/S -
  // OTS Amount; Impact on P&L = OTS Amount - Total P&L.
  const totalSac = (totalDues!=='') ? totalDues-ots : '';
  const ledgerSac = s.os!=='' ? s.os-ots : ''; // also shown as "(BDWO Amount)" -- same figure
  const impact = s.totalPL!=='' ? ots - s.totalPL : '';
  totalSacEl.textContent = fmtINR2(totalSac);
  ledgerEl.textContent = fmtINR2(ledgerSac);
  if(pctEl) pctEl.textContent = (s.os) ? (ots/s.os*100).toFixed(1)+'%' : '—';
  impactEl.classList.remove('pos','neg');
  if(impact!=='' && !isNaN(impact)){
    const prev = (typeof impactEl.__val==='number') ? impactEl.__val : 0;
    impactEl.__val = impact;
    animateNumber(impactEl, prev, impact, (v)=>{
      const sign = v>0.5?'+':(v<-0.5?'−':'');
      return sign + fmtINR2(Math.abs(v)).replace('₹','₹ ');
    });
    impactEl.classList.add(impact>0?'pos':(impact<0?'neg':''));
  } else {
    impactEl.textContent = fmtINR2(impact);
    impactEl.__val = 0;
  }
  impactEl.classList.remove('flash');
  void impactEl.offsetWidth;
  impactEl.classList.add('flash');
}

function recalcAggregate(){
  const slots = window.__slots;
  let totalOts=0;
  slots.forEach(s=>{
    const v = parseOtsAmount(otsAmounts[s.acctNo]);
    if(v!==null) totalOts+=v;
  });
  // Alok's explicit call: for a multi-account borrower, this whole panel
  // only ever shows real numbers once EVERY linked account has an OTS
  // Amount typed in -- not as soon as any one of them does. It used to sum
  // Total Dues over all the accounts but OTS Amount over only the filled
  // ones, so Total Sacrifice/Impact/Recovery Scale could look meaningful
  // while actually mixing two different sets of accounts. Now the whole
  // panel stays "—" (identical to the "nothing typed yet" state) until
  // every account is filled, at which point totalOts already correctly
  // covers the same full set liveTotalDues does.
  const allFilled = slots.length>0 && slots.every(s=>parseOtsAmount(otsAmounts[s.acctNo])!==null);
  // Total Dues is live (Interest Reversal is editable and folds into it),
  // so the aggregate sum must be recomputed fresh, not read from a stale
  // snapshot. Total P&L no longer depends on Interest Reversal, so
  // window.__totalPL (the static per-render snapshot) stays valid as-is.
  let liveTotalDues=0;
  slots.forEach(s=>{ const td = totalDuesFor(s); liveTotalDues += (td!==''?td:0); });
  // Total Sacrifice = Total Dues - OTS Amount (matches the per-account
  // formula in recalcLoan()); Interest Reversal is already folded into
  // Total Dues, so it isn't added again here.
  const aggTotalSac = liveTotalDues - totalOts;
  document.getElementById('aggOts') && (document.getElementById('aggOts').textContent = allFilled?fmtINR2(totalOts):'—');
  document.getElementById('aggSac') && (document.getElementById('aggSac').textContent = allFilled?fmtINR2(aggTotalSac):'—');
  const otsTxt = allFilled?fmtINR2(totalOts):'—';
  const railOts = document.getElementById('railOts'); if(railOts) railOts.textContent = otsTxt;
  const railOts2 = document.getElementById('railOts2'); if(railOts2) railOts2.textContent = otsTxt;
  const railDues = document.getElementById('railDues'); if(railDues) railDues.textContent = fmtINR2(liveTotalDues);
  const railPLLeft = document.getElementById('railPLLeft');
  if(railPLLeft){
    const impact = allFilled ? (totalOts - window.__totalPL) : '';
    railPLLeft.textContent = impact===''?'—':(impact>0?'+':(impact<0?'−':'')) + fmtINR2(Math.abs(impact));
    railPLLeft.classList.remove('pos','neg');
    if(impact!==''){ if(impact>0) railPLLeft.classList.add('pos'); else if(impact<0) railPLLeft.classList.add('neg'); }
  }
  const railSac = document.getElementById('railSac'); if(railSac) railSac.textContent = allFilled?fmtINR2(aggTotalSac):'—';
  // Live aggregate summary panel (shown for multi-account borrowers) --
  // the hero ring shows settlement progress (OTS as a share of Total
  // Dues), the same figure the old unlabeled #aggBar::after ring drove,
  // now with a real percentage printed inside it.
  const pct = (allFilled && liveTotalDues>0) ? Math.max(0,Math.min(100,(totalOts/liveTotalDues)*100)) : 0;
  const heroRingEl = document.getElementById('aggHeroRing');
  if(heroRingEl) heroRingEl.style.setProperty('--pct', pct.toFixed(1));
  const heroRingPctEl = document.getElementById('aggHeroRingPct');
  if(heroRingPctEl) heroRingPctEl.textContent = (allFilled && liveTotalDues>0) ? pct.toFixed(0)+'%' : '—';
  const heroSubEl = document.getElementById('aggHeroSub');
  if(heroSubEl) heroSubEl.textContent = `across ${slots.length} account${slots.length>1?'s':''} · O/S ${fmtCr(window.__totalOS)}`;
  // Alok's request -- the ring only ever showed OTS as a share of Total
  // Dues; this pair makes both readings visible together (Dues share can
  // run well above/below the O/S share depending on how much UCI/Interest
  // Reversal is in play), not just whichever one the ring happens to draw.
  const pctDuesEl = document.getElementById('aggPctDues');
  if(pctDuesEl) pctDuesEl.textContent = (allFilled && liveTotalDues>0) ? pct.toFixed(1)+'%' : '—';
  const pctOsEl = document.getElementById('aggPctOs');
  if(pctOsEl){
    const totalOsForPct = window.__totalOS;
    const pctOs = (allFilled && totalOsForPct>0) ? Math.max(0,(totalOts/totalOsForPct)*100) : null;
    pctOsEl.textContent = pctOs!==null ? pctOs.toFixed(1)+'%' : '—';
  }

  const aggOtsEl = document.getElementById('aggTotOts');
  if(aggOtsEl){
    // .innerHTML + fmtINR2Wrap (not the plain otsTxt used by the rail
    // above) -- the tight aggBar sidebar column needs the single <wbr>
    // after ₹ so a figure that doesn't fit wraps cleanly onto its own
    // line instead of splitting mid-digit.
    aggOtsEl.innerHTML = allFilled ? fmtINR2Wrap(totalOts) : '—';
    const aggNetOsEl = document.getElementById('aggTotNetOs');
    if(aggNetOsEl) aggNetOsEl.innerHTML = fmtINR2Wrap(window.__totalNetOS);
    const aggPLEl = document.getElementById('aggTotPL');
    if(aggPLEl) aggPLEl.innerHTML = fmtINR2Wrap(window.__totalPL);
    const aggSacEl = document.getElementById('aggTotSac');
    if(aggSacEl) aggSacEl.innerHTML = allFilled?fmtINR2Wrap(aggTotalSac):'—';
    const aggImpEl = document.getElementById('aggTotImpact');
    if(aggImpEl){
      const impact = allFilled ? (totalOts - window.__totalPL) : '';
      aggImpEl.classList.remove('pos','neg');
      if(impact===''){ aggImpEl.innerHTML='—'; }
      else {
        aggImpEl.innerHTML = (impact>0?'+':(impact<0?'−':'')) + fmtINR2Wrap(Math.abs(impact));
        const sign = impact>0?'pos':(impact<0?'neg':'');
        if(sign) aggImpEl.classList.add(sign);
      }
    }
  }

  // Bonus: Recovery scale (aggregate-level needle gauge) + Where The Dues
  // Go (waterfall). Break-even = O/S - Provision, i.e. the same figure
  // already computed per-render as window.__totalPL -- OTS above this
  // point means the settlement is P&L-positive, below it means an
  // additional charge. Scale ceiling is live Total Dues.
  const totalOs = window.__totalOS;
  const BE = window.__totalPL;
  const scaleMax = liveTotalDues;
  const bePct = scaleMax>0 ? Math.max(0,Math.min(100,(BE/scaleMax)*100)) : 0;
  const needlePct = scaleMax>0 ? Math.max(0,Math.min(100,(totalOts/scaleMax)*100)) : 0;
  const bandSafeEl = document.getElementById('aggBandSafe');
  if(bandSafeEl){ bandSafeEl.style.left = bePct+'%'; bandSafeEl.style.width = (100-bePct)+'%'; }
  const needleEl = document.getElementById('aggNeedle');
  if(needleEl) needleEl.style.left = 'calc('+needlePct+'% - 1px)';
  const needleValEl = document.getElementById('aggNeedleVal');
  if(needleValEl){
    // Bubble is centered on the needle by default (left:50%,
    // translateX(-50%)) -- fine everywhere except right at the two ends
    // of the track, where a centered bubble would hang half off the
    // sidebar's edge. Re-anchor it inward there instead of clipping.
    needleValEl.style.transform = needlePct<10 ? 'translateX(-6px)' : (needlePct>90 ? 'translateX(calc(-100% + 6px))' : 'translateX(-50%)');
    const prev = (typeof needleValEl.__val==='number') ? needleValEl.__val : 0;
    if(allFilled){
      needleValEl.__val = totalOts;
      animateNumber(needleValEl, prev, totalOts, v=>fmtCr(v), 450);
    } else {
      needleValEl.__val = 0;
      needleValEl.textContent = '—';
    }
  }
  const beValEl = document.getElementById('aggBEVal'); if(beValEl) beValEl.textContent = fmtCr(BE);
  const osValEl = document.getElementById('aggOSVal'); if(osValEl) osValEl.textContent = fmtCr(totalOs);
  const duesValEl = document.getElementById('aggDuesVal'); if(duesValEl) duesValEl.textContent = fmtCr(liveTotalDues);

  const ledgerSac = Math.max(totalOs-totalOts,0);
  const uci = slots.reduce((a,s)=>a+((s.uci!=='')?s.uci:0),0);
  const wA = liveTotalDues>0 ? Math.max(0,Math.min(100,totalOts/liveTotalDues*100)) : 0;
  const wB = liveTotalDues>0 ? Math.max(0,Math.min(100-wA,ledgerSac/liveTotalDues*100)) : 0;
  const wC = Math.max(100-wA-wB,0);
  // Fixed (non-theme-flipping) chart swatches -- a UI token like --ink
  // inverts between dark/light themes, which would make one segment's
  // text unreadable in one of the two themes (caught in the feasibility
  // mockup review). Chart categories get literal colors instead.
  const WF1='#1B2A44', WF2='#D4A544', WF3='#7A8798';
  // Update the 3 segment spans in place (style.width + textContent) rather
  // than rebuilding via innerHTML -- innerHTML replacement destroys and
  // recreates the elements every render, which gives the CSS width
  // transition nothing to animate from (a freshly-created element just
  // appears at its final width, no motion). Same reasoning applies to the
  // key row's <b> values below, using animateNumber() like the P&L Impact
  // figure already does, instead of an innerHTML dump.
  const wf1 = document.getElementById('aggWf1');
  if(wf1){ wf1.style.width = wA.toFixed(1)+'%'; wf1.style.background = WF1; wf1.style.color = '#fff'; wf1.textContent = wA>7?wA.toFixed(0)+'%':''; }
  const wf2 = document.getElementById('aggWf2');
  if(wf2){ wf2.style.width = wB.toFixed(1)+'%'; wf2.style.background = WF2; wf2.style.color = '#241d08'; wf2.textContent = wB>7?wB.toFixed(0)+'%':''; }
  const wf3 = document.getElementById('aggWf3');
  if(wf3){ wf3.style.width = wC.toFixed(1)+'%'; wf3.style.background = WF3; wf3.style.color = '#fff'; wf3.textContent = wC>7?wC.toFixed(0)+'%':''; }

  const animateWfVal = (el, newVal, active) => {
    if(!el) return;
    const prev = (typeof el.__val==='number') ? el.__val : 0;
    if(active===false){ el.__val = 0; el.textContent = '—'; return; }
    el.__val = newVal;
    animateNumber(el, prev, newVal, v=>fmtINR2(v), 450);
  };
  animateWfVal(document.getElementById('aggWfCash'), totalOts, allFilled);
  animateWfVal(document.getElementById('aggWfLedger'), ledgerSac, allFilled);
  animateWfVal(document.getElementById('aggWfUci'), uci, true);

  renderPrintView();
}

function renderPrintView(){
  const slots = window.__slots; const custRow = window.__custRow;
  if(!slots || !custRow) return;
  const totalOS = slots.reduce((a,s)=>a+((s.os!=='')?s.os:0),0);
  // Total Dues is live (Interest Reversal folds into it), so it's summed
  // fresh here rather than read from the static window.__totalDues snapshot.
  let totalDues = 0;
  slots.forEach(s=>{ const td = totalDuesFor(s); totalDues += (td!==''?td:0); });

  function otsFor(s){
    return parseOtsAmount(otsAmounts[s.acctNo]);
  }
  let totalOtsSum = 0, totalLedgerSac = 0;
  slots.forEach(s=>{ const v = otsFor(s); if(v!==null){ totalOtsSum+=v; totalLedgerSac+=(s.os-v); } });
  // Same "all or nothing" rule as the live aggregate sidebar
  // (recalcAggregate()): Total Dues above sums every linked account, so
  // Total OTS Amount/Ledger Sacrifice/Sacrifice below only show real
  // figures once every account has one typed -- otherwise this printed/
  // shared sheet would mix "all accounts' dues" against "only some
  // accounts' OTS Amount" the same way the sidebar used to.
  const allOtsFilled = slots.length>0 && slots.every(s=>otsFor(s)!==null);
  // Same aggregate Settlement % the on-screen sidebar chips already show
  // (recalcAggregate()) -- see the "OTS Amount" row below for the per-
  // account version of the same gap.
  const aggPct = totalDues>0 ? (totalOtsSum/totalDues)*100 : 0;
  const aggPctOs = totalOS>0 ? (totalOtsSum/totalOS)*100 : null;

  // Rows the sheet is actually read for -- bolded/enlarged in print (see
  // .pv-table tr.pv-strong in styles.css) so they stand out from the
  // supporting particulars around them, same "which numbers matter"
  // convention as the on-screen loan table's own lt-strong rows.
  // Total Contractual Dues is deliberately NOT in this print/PDF table --
  // it stays on-screen only (loanTableHTML) per Alok's review; Total
  // Sacrifice below reads off Total Dues (+ Interest Reversal), not it.
  const STRONG_ROWS = new Set(['O/S Balance','Total Dues','Total P&L','OTS Amt as per Lok Adalat','OTS Amount','Total Sacrifice','Impact on P&L']);
  // Scheme moved here from the page footer (was repeating the branch name a
  // third time alongside the header and the borrower info grid) -- one row
  // per account, right above O/S Balance where the settlement figures start.
  // Net O/S is deliberately NOT a separate row -- it's always identical to
  // O/S Balance (Net O/S = O/S Balance, no exceptions), so showing both was
  // just the same number twice; O/S Balance is the one that stays.
  const rows = [
    ['Sanction Date', 'calendar', s=>fmtDate(toDate(s.sanctionDate))],
    ['Sanction Limit', 'doc', s=>fmtINR2(s.sanctionLimit)],
    ['Asset Code', 'tag', s=>esc(s.assetCode)||'—'],
    ['NPA Date', 'warn', s=>fmtDate(toDate(s.npaDate))],
    ['Days in NPA', 'clock', s=>s.daysNpa!==''?s.daysNpa.toLocaleString('en-IN')+' days':'—'],
    ['Scheme', 'tag', s=>esc(s.scheme)||'—'],
    ['O/S Balance', 'coin', s=>fmtINR2(s.os)],
    [uciLabelWithTenure(slots), 'percent', s=>fmtINR2(s.uci)],
    ['Total Dues', 'layers', s=>fmtINR2(totalDuesFor(s))],
    ['Interest Reversal', 'rotate', s=>fmtINR2(uriFor(s))],
    ['Provision', 'shield', s=>fmtINR2(s.provision)],
    ['Total P&L', 'trend', s=>fmtINR2(s.totalPL) + (s.ratio!==''?` (${(s.ratio*100).toFixed(1)}%)`:'')],
    ['OTS Amt as per Lok Adalat', 'scale', s=>{const la=lokAdalatMin(s); if(!la) return '—'; if(!la.eligible) return 'Not Eligible'; return fmtINR2(la.amount)+` (${(la.pct*100).toFixed(0)}%)`;}],
    // Same Settlement Progress % the on-screen loan table already shows per
    // account (recalcLoan()'s settlePct/settlePctOs) -- this print/PDF sheet
    // never had it, so it read as "missing" even though the underlying
    // figures were always live (Alok: "ots one pager print karte main ye %
    // nahi aata"). Computed identically here rather than reading a DOM value,
    // since this sheet is rebuilt straight from data, not a screen capture.
    ['OTS Amount', 'coin', s=>{
      const v=otsFor(s); if(v===null) return '—';
      const td=totalDuesFor(s);
      const pct = td>0 ? Math.max(0,Math.min(100,(v/td)*100)) : 0;
      const pctOs = s.os>0 ? Math.max(0,(v/s.os)*100) : null;
      const osPart = pctOs!==null ? ` · ${pctOs.toFixed(1)}% of O/S` : '';
      return `${fmtINR2(v)} (${pct.toFixed(1)}% of dues${osPart})`;
    }],
    ['Total Sacrifice', 'percent', s=>{const v=otsFor(s); return v===null?'—':fmtINR2(totalDuesFor(s)-v);}],
    ['Ledger Sacrifice (BDWO Amount)', 'badge', s=>{const v=otsFor(s); return v===null?'—':fmtINR2(s.os-v);}],
    // Arrow mirrors the up/down icon-set convention from Excel's conditional
    // formatting -- up for a positive (better-than-booked) P&L impact, down
    // for a negative one -- so the sign reads at a glance, not just from the
    // minus sign buried in the number.
    ['Impact on P&L', 'bars', s=>{const v=otsFor(s); if(v===null) return '—'; const impact=v-s.totalPL; const arrow=impact>0?'▲ ':(impact<0?'▼ ':''); return arrow+fmtINR2(impact);}],
  ];
  const tableRows = rows.map(([label,icon,fn])=>`<tr${STRONG_ROWS.has(label)?' class="pv-strong"':''}><td class="pv-label">${label}</td>${slots.map(s=>`<td>${fn(s)}</td>`).join('')}</tr>`).join('');

  // Sol ID now rides along with the branch name in the header ("Branch:
  // MENDU (9291)") instead of repeating as its own row in the info grid
  // below -- same one-mention-only convention as the earlier branch-name
  // dedup fix.
  const solId = esc(custRow[C.SOL_ID])||'';
  const logoSrc = document.querySelector('.nav-logo')?.src || '';
  document.getElementById('printArea').innerHTML = `
    <div class="pv-topbar"></div>
    <div class="pv-brandrow">
      ${logoSrc?`<img class="pv-logo" src="${logoSrc}" alt="">`:''}
      <div><div class="pv-bank">Uttar Pradesh Gramin Bank</div><div class="pv-bank-sub">Regional Office Hathras</div></div>
    </div>
    <div class="pv-doctitle">OTS Settlement Statement<div class="pv-doctitle-rule"></div></div>
    <div class="pv-meta"><span>Report Date <b>${fmtDate(new Date())}</b></span><span>Branch <b>${esc(custRow[C.SOL_DESC])||''}${solId?` (${solId})`:''}</b></span></div>
    <div class="pv-borrower">
      <div class="pv-name">${esc(custRow[C.NAME])||'—'}</div>
      <div class="pv-addr">${esc(custRow[C.ADDR])||'—'}</div>
      <div class="pv-grid">
        <div><div class="k">Cust ID</div><div class="v">${esc(custRow[C.CUST_ID])||'—'}</div></div>
        <div><div class="k">Mobile</div><div class="v">${esc(custRow[C.PHONE])||'—'}</div></div>
        <div><div class="k">PAN</div><div class="v">${esc(custRow[C.PAN])||'—'}</div></div>
        <div><div class="k">Aadhar</div><div class="v">${esc(custRow[C.AADHAR])||'—'}</div></div>
        <div><div class="k">SB A/c</div><div class="v">${esc(custRow[C.SB_ACCT])||'—'}</div></div>
        <div><div class="k">SB Balance</div><div class="v">${fmtINR2(custRow[C.SB_BAL]===''?0:custRow[C.SB_BAL])}</div></div>
      </div>
    </div>
    <div class="pv-sec-lbl">Particulars</div>
    <table class="pv-table">
      <thead><tr><th>Particulars</th>${slots.map(s=>`<th>${esc(s.acctNo)}</th>`).join('')}</tr></thead>
      <tbody>${tableRows}</tbody>
    </table>
    <div class="pv-agg">
      <div class="pv-agg-title">Aggregate Totals</div>
      <div class="pv-agg-row"><span>Total O/S Balance</span><span>${fmtINR2(totalOS)}</span></div>
      <div class="pv-agg-row"><span>Total Dues</span><span>${fmtINR2(totalDues)}</span></div>
      <div class="pv-agg-row pv-agg-hero"><span>Total OTS Amount</span><span>${allOtsFilled?`${fmtINR2(totalOtsSum)} (${aggPct.toFixed(1)}% of dues${aggPctOs!==null?` · ${aggPctOs.toFixed(1)}% of O/S`:''})`:'—'}</span></div>
      <div class="pv-agg-row"><span>Total Ledger Sacrifice</span><span>${allOtsFilled?fmtINR2(totalLedgerSac):'—'}</span></div>
      <div class="pv-agg-row pv-agg-hero"><span>Total Sacrifice</span><span>${allOtsFilled?fmtINR2(totalDues-totalOtsSum):'—'}</span></div>
    </div>
    <div class="pv-foot">UPGB OTS Calculator &middot; Designed &amp; Developed by Alok Mittal</div>
  `;
}

/* Excel export with LIVE formulas, not just the computed snapshot printed
   above -- every figure that depends on another cell (UCI, Total Dues,
   Provision, Total P&L, and everything downstream of the OTS Amount you
   type in) is a real =formula, so editing OTS Amount (or O/S Balance, if
   a payment changes it) recalculates every dependent cell in Excel itself,
   exactly like the on-screen calculator does. Only the true source-data
   fields (Sanction Date/Limit, Asset Code, NPA Date, O/S Balance,
   Interest Reversal) are plain values -- everything else is derived.
   The visible rows here match the print sheet exactly (same set Alok
   reviewed on paper) -- Scheme, UCI Anchor Date, and the Provision Rate
   lookup table are helper/intermediate values the PDF never showed, so
   they live on a second "Calculation Details" sheet instead of cluttering
   this one; formulas here just reference across to that sheet. */
/* Exactly the print sheet's 16 rows, in the print sheet's order -- Scheme
   included (it used to sit on the helper sheet), Net O/S excluded (it is
   always identical to O/S Balance, which is why the print sheet dropped it),
   and "OTS Amount" plainly named. The two sheets are now row-for-row the
   same document. */
const OTS_XL_ROW_LABELS = [
  'Sanction Date','Sanction Limit','Asset Code','NPA Date','Days in NPA','Scheme',
  'O/S Balance','UCI @ 8.5%','Total Dues','Interest Reversal','Provision','Total P&L',
  'OTS Amount','Total Sacrifice','Ledger Sacrifice (BDWO Amount)','Impact on P&L',
];
const OTS_XL_CALC_ROW_LABELS = ['UCI Anchor Date'];
/* SheetJS (the "xlsx" global used elsewhere in this file, e.g. Daily NPA
   Projection's export) is the free Community Edition, which can only
   READ cell styles, not write them -- .z (number format) writes fine, but
   fonts/fills/borders are silently dropped, so a SheetJS-built workbook
   always comes out plain black-on-white regardless of what's set on the
   cell object. ExcelJS (window.ExcelJS, js/vendor/exceljs.min.js) writes
   real styling, so this export uses it instead. Deliberately kept plain/
   functional (bold text, real borders, live formulas, no fill colors) --
   Excel is for editing and calculation, not decoration; the brass/color
   treatment stays on the app screen and the print/PDF sheet. */
const XL_BORDER_THIN = {style:'thin', color:{argb:'FF555555'}};
const XL_BORDER_ALL = {top:XL_BORDER_THIN, bottom:XL_BORDER_THIN, left:XL_BORDER_THIN, right:XL_BORDER_THIN};
const XL_INR_FMT = '"₹"#,##,##0.00;[Red]-"₹"#,##,##0.00';
// Impact on P&L only: same currency format, plus a profit/loss arrow driven
// by the cell's live sign (green ▲ for profit, red ▼ for loss).
const XL_INR_FMT_PL = '[Green]"▲ ₹"#,##,##0.00;[Red]"▼ -₹"#,##,##0.00';
const XL_DATE_FMT = 'dd-mm-yyyy';
const XL_CALC_SHEET = 'Calculation Details';

async function exportOtsExcel(){
  const slots = window.__slots; const custRow = window.__custRow;
  if(!slots || !custRow) return;
  await Promise.all([ensureXLSX(), ensureExcelJS()]);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('OTS Calculator', { views: [{showGridLines:false}] });
  /* Hidden: it only carries the UCI anchor dates and the provision-rate
     lookup the formulas point at. The printed sheet never showed either, so
     the workbook now opens on one sheet that matches the PDF. Right-click
     the tab strip and Unhide to inspect it. */
  const wsCalc = wb.addWorksheet(XL_CALC_SHEET, { views: [{showGridLines:false}], state:'hidden' });
  const colLetter = i => XLSX.utils.encode_col(i+1); // account 0 -> B, 1 -> C, ...
  const cols = slots.map((s,i)=>colLetter(i));
  const lastCol = cols[cols.length-1];
  const lastColIdx = cols.length + 1; // 1-indexed: A=1, B=2...
  /* One span for every merged row -- title, subtitle, meta, name, address,
     aggregates and footer all end on the same column as the table below
     them. Previously the header block merged out to column E while the table
     stopped at D, leaving a permanently empty column hanging off the right
     of every export. The floor of 4 keeps a single-account sheet wide enough
     for the footer line without stretching a 3-account one. */
  const SPAN = Math.max(lastColIdx, 4);
  const SPAN_COL = XLSX.utils.encode_col(SPAN - 1);

  const setOn = (sheet, addr, value, opts={}) => {
    const cell = sheet.getCell(addr);
    cell.value = value;
    if(opts.numFmt) cell.numFmt = opts.numFmt;
    if(opts.font) cell.font = opts.font;
    if(opts.fill) cell.fill = {type:'pattern', pattern:'solid', fgColor:{argb:opts.fill}};
    if(opts.align) cell.alignment = opts.align;
    if(opts.border!==false) cell.border = opts.border || XL_BORDER_ALL;
    return cell;
  };
  const set = (addr, value, opts) => setOn(ws, addr, value, opts);
  const setCalc = (addr, value, opts) => setOn(wsCalc, addr, value, opts);
  const dateVal = jsDate => jsDate || null;
  const formula = f => ({formula: f});

  // ---- Header, matching the print sheet exactly: no logo (dropped per
  // Alok's request -- it never fit cleanly at export width either), a
  // plain title/subtitle/meta block, then the borrower's name and address,
  // then the same two-column info grid the PDF uses (Cust ID/Mobile/PAN/
  // Aadhar beside SB A/c/SB Balance; Sol ID rides in the Branch line, not
  // its own field). Built with a running row counter, not fixed row
  // numbers, so the header can grow or shrink without hand-recalculating
  // every row below it. ----
  const solId = String(custRow[C.SOL_ID]||'');
  /* Guidance the paper sheet has no need for (paper cannot be edited) is
     attached as a cell note rather than its own row, so the sheet keeps the
     PDF's exact shape. Wrapped because note support varies by ExcelJS build
     and a missing note must never cost the whole export. */
  const addNote = (addr, text) => { try{ ws.getCell(addr).note = text; }catch(e){} };
  let r = 1;
  ws.mergeCells(r,1,r,SPAN);
  set(`A${r}`, 'UPGB OTS CALCULATOR', {font:{bold:true, size:16, color:{argb:'FF000000'}}, align:{horizontal:'center'}, border:false});
  ws.getRow(r).height = 26;
  r++;
  ws.mergeCells(r,1,r,SPAN);
  set(`A${r}`, 'Uttar Pradesh Gramin Bank (Regional Office Hathras)', {font:{size:11, color:{argb:'FF333333'}}, align:{horizontal:'center'}, border:{bottom:{style:'medium', color:{argb:'FF555555'}}}});
  r += 2;

  const reportDateRow = r;
  set(`A${reportDateRow}`, 'Report Date', {font:{bold:true, color:{argb:'FF333333'}}, border:false});
  /* On screen, "days since X" (computeUCI, Days in NPA, etc.) is
     daysBetween(today, anchor) = Math.round((today-anchor)/86400000),
     with `today` carrying the current time-of-day -- so it rounds UP to
     the next day once past noon, not just at midnight. Excel's live
     formulas below do plain subtraction between two date serials with no
     rounding, so writing the raw `new Date()` (full timestamp) here made
     every day-count -- and everything downstream of it (UCI, Total Dues,
     Total Sacrifice, Impact on P&L) -- silently drift from the on-screen
     figure by up to a day's interest on every export, worse the later in
     the day it ran. Snapping to the SAME nearest-midnight the on-screen
     Math.round would resolve to (before noon -> today 00:00, at/after
     noon -> tomorrow 00:00) makes Excel's plain subtraction land on the
     identical whole-day count instead. */
  const reportDateSnapshot = new Date();
  reportDateSnapshot.setHours(reportDateSnapshot.getHours()>=12 ? 24 : 0, 0, 0, 0);
  set(`B${reportDateRow}`, dateVal(reportDateSnapshot), {numFmt:XL_DATE_FMT, font:{bold:true, color:{argb:'FF000000'}}, border:XL_BORDER_ALL});
  addNote(`B${reportDateRow}`, 'Editable. Every UCI, Days in NPA and dues figure below recalculates off this date.');
  set(`C${reportDateRow}`, 'Branch', {font:{bold:true, color:{argb:'FF333333'}}, border:false});
  ws.mergeCells(reportDateRow,4,reportDateRow,SPAN);
  set(`D${reportDateRow}`, `${custRow[C.SOL_DESC]||''}${solId?` (${solId})`:''}`, {font:{color:{argb:'FF000000'}}, border:false});
  const reportDateRef = `$B$${reportDateRow}`;
  r += 2;

  ws.mergeCells(r,1,r,SPAN);
  set(`A${r}`, custRow[C.NAME]||'', {font:{bold:true, size:12, color:{argb:'FF000000'}}, border:false});
  r++;
  ws.mergeCells(r,1,r,SPAN);
  set(`A${r}`, custRow[C.ADDR]||'', {font:{size:10, color:{argb:'FF333333'}}, border:false});
  r += 2;

  const infoPairs = [
    ['Cust ID', String(custRow[C.CUST_ID]||''), 'SB A/c', String(custRow[C.SB_ACCT]||'')],
    ['Mobile', String(custRow[C.PHONE]||''), 'SB Balance', custRow[C.SB_BAL]===''?0:custRow[C.SB_BAL]],
    ['PAN', String(custRow[C.PAN]||''), '', ''],
    ['Aadhar', String(custRow[C.AADHAR]||''), '', ''],
  ];
  infoPairs.forEach((row,i)=>{
    const rr = r+i;
    set(`A${rr}`, row[0], {font:{bold:true, color:{argb:'FF333333'}}, border:false});
    set(`B${rr}`, row[1], {font:{color:{argb:'FF000000'}}, border:false});
    if(row[2]){
      set(`C${rr}`, row[2], {font:{bold:true, color:{argb:'FF333333'}}, border:false});
      ws.mergeCells(rr,4,rr,SPAN);
      set(`D${rr}`, row[3], {font:{color:{argb:'FF000000'}}, numFmt: row[2]==='SB Balance'?XL_INR_FMT:undefined, border:false});
    }
  });
  r += infoPairs.length + 1;

  // ---- Calculation Details sheet (hidden): the UCI anchor date per account
  // plus the Provision Rate lookup table. Scheme no longer lives here -- it
  // is a proper row on the main sheet now, and the anchor formula reads it
  // from there, so the same value is not stored in two places. ----
  wsCalc.mergeCells(1,1,1,Math.max(lastColIdx,4));
  setCalc('A1', 'Calculation Details', {font:{bold:true, size:14, color:{argb:'FF000000'}}, border:false});
  wsCalc.mergeCells(2,1,2,Math.max(lastColIdx,4));
  setCalc('A2', 'Helper values feeding the "OTS Calculator" sheet\'s formulas -- not shown on the printed sheet.', {font:{italic:true, size:10, color:{argb:'FF666666'}}, border:false});
  const calcHeaderRow = 4;
  setCalc(`A${calcHeaderRow}`, 'Particulars', {font:{bold:true, color:{argb:'FF000000'}}});
  slots.forEach((s,i)=>setCalc(`${cols[i]}${calcHeaderRow}`, s.acctNo, {font:{bold:true, color:{argb:'FF000000'}}, align:{horizontal:'center'}}));
  const RC = { anchor: calcHeaderRow+1 };
  OTS_XL_CALC_ROW_LABELS.forEach((label,i)=>setCalc(`A${calcHeaderRow+1+i}`, label, {font:{bold:true, color:{argb:'FF000000'}}}));

  const RATE_ROWS = [['SUB_STD',0.10],['DA1',0.20],['DA2',0.30],['DA3',1],['LOSS',1]];
  const rateHeadRow = calcHeaderRow + OTS_XL_CALC_ROW_LABELS.length + 2;
  const rateTable = `'${XL_CALC_SHEET}'!$A$${rateHeadRow+1}:$B$${rateHeadRow+RATE_ROWS.length}`;
  setCalc(`A${rateHeadRow}`, 'Provision Rate reference (by Asset Code)', {font:{bold:true, size:10.5}, border:false});
  RATE_ROWS.forEach(([code,rate],i)=>{
    setCalc(`A${rateHeadRow+1+i}`, code, {font:{color:{argb:'FF000000'}}});
    setCalc(`B${rateHeadRow+1+i}`, rate, {numFmt:'0%', font:{color:{argb:'FF000000'}}});
  });

  // ---- Particulars table (main sheet) ----
  const headerRow = r;
  set(`A${headerRow}`, 'Particulars', {font:{bold:true, color:{argb:'FF000000'}}});
  slots.forEach((s,i)=>set(`${cols[i]}${headerRow}`, s.acctNo, {font:{bold:true, color:{argb:'FF000000'}}, align:{horizontal:'center'}}));

  // Same rows the print sheet bolds, so the two read identically on paper.
  const STRONG_ROWS = new Set(['O/S Balance','Total Dues','Total P&L','OTS Amount','Total Sacrifice','Impact on P&L']);
  const rowOf = label => headerRow + 1 + OTS_XL_ROW_LABELS.indexOf(label);
  const R = {
    sanctionDate: rowOf('Sanction Date'), sanctionLimit: rowOf('Sanction Limit'), assetCode: rowOf('Asset Code'),
    npaDate: rowOf('NPA Date'), daysNpa: rowOf('Days in NPA'), scheme: rowOf('Scheme'),
    os: rowOf('O/S Balance'), uci85: rowOf('UCI @ 8.5%'), totalDues: rowOf('Total Dues'),
    uri: rowOf('Interest Reversal'), provision: rowOf('Provision'), totalPL: rowOf('Total P&L'),
    ots: rowOf('OTS Amount'), totalSac: rowOf('Total Sacrifice'),
    ledgerSac: rowOf('Ledger Sacrifice (BDWO Amount)'), impact: rowOf('Impact on P&L'),
  };
  addNote(`A${R.ots}`, 'Type a settlement amount here. Total Sacrifice, Ledger Sacrifice, Impact on P&L and the aggregate totals all recalculate from it.');

  OTS_XL_ROW_LABELS.forEach((label,i)=>{
    const r = headerRow + 1 + i;
    set(`A${r}`, label, {font:{bold:true, color:{argb:'FF000000'}}});
  });

  slots.forEach((s,i)=>{
    const c = cols[i];
    const rowStyle = r => ({border:XL_BORDER_ALL, align:{horizontal:'right'}, font:{color:{argb:'FF000000'}, bold:STRONG_ROWS.has(OTS_XL_ROW_LABELS[r-headerRow-1])}});

    // UCI Anchor Date (formula) for this account on the hidden sheet, read by
    // this sheet's UCI @ 8.5% below. Both its inputs -- NPA Date and Scheme --
    // are read back off the main sheet, so editing either there flows through.
    const npaRefMain = `'OTS Calculator'!${c}${R.npaDate}`;
    const schemeRefCalc = `'OTS Calculator'!${c}${R.scheme}`;
    // Anchor date replicates computeUCI()'s scheme-dependent rule exactly:
    // CC004 (KCC) uses fixed 24-Mar/24-Sep half-year edges; every other
    // scheme anchors to end of NPA month (or the previous month's end, if
    // the NPA date itself isn't a month-end).
    const anchorF = `IF(${schemeRefCalc}="CC004",`+
      `IF(${npaRefMain}>DATE(YEAR(${npaRefMain}),9,24),DATE(YEAR(${npaRefMain}),9,24),`+
        `IF(${npaRefMain}>DATE(YEAR(${npaRefMain}),3,24),DATE(YEAR(${npaRefMain}),3,24),DATE(YEAR(${npaRefMain})-1,9,24))),`+
      `IF(${npaRefMain}=EOMONTH(${npaRefMain},0),DATE(YEAR(${npaRefMain}),MONTH(${npaRefMain}),29),EOMONTH(${npaRefMain},-1)))`;
    setCalc(`${c}${RC.anchor}`, formula(anchorF), {border:XL_BORDER_ALL, align:{horizontal:'right'}, font:{color:{argb:'FF000000'}}, numFmt:XL_DATE_FMT});
    const anchorRefCalc = `'${XL_CALC_SHEET}'!${c}${RC.anchor}`;

    set(`${c}${R.sanctionDate}`, dateVal(toDate(s.sanctionDate)), {...rowStyle(R.sanctionDate), numFmt:XL_DATE_FMT});
    set(`${c}${R.sanctionLimit}`, s.sanctionLimit===''?0:s.sanctionLimit, {...rowStyle(R.sanctionLimit), numFmt:XL_INR_FMT});
    set(`${c}${R.assetCode}`, s.assetCode, rowStyle(R.assetCode));
    set(`${c}${R.npaDate}`, dateVal(toDate(s.npaDate)), {...rowStyle(R.npaDate), numFmt:XL_DATE_FMT});
    set(`${c}${R.daysNpa}`, formula(`${reportDateRef}-${c}${R.npaDate}`), {...rowStyle(R.daysNpa), numFmt:'0'});
    set(`${c}${R.scheme}`, s.scheme||'', rowStyle(R.scheme));
    set(`${c}${R.os}`, s.os===''?0:s.os, {...rowStyle(R.os), numFmt:XL_INR_FMT});
    set(`${c}${R.uci85}`, formula(`${c}${R.os}*8.5/100*((${reportDateRef}-${anchorRefCalc})/365)`), {...rowStyle(R.uci85), numFmt:XL_INR_FMT});
    set(`${c}${R.uri}`, uriFor(s), {...rowStyle(R.uri), numFmt:XL_INR_FMT});
    // Total Dues = O/S + UCI@8.5% + Interest Reversal.
    set(`${c}${R.totalDues}`, formula(`${c}${R.os}+${c}${R.uci85}+${c}${R.uri}`), {...rowStyle(R.totalDues), numFmt:XL_INR_FMT});
    // Provision reads O/S Balance directly. It used to go through a Net O/S
    // row, but that row only ever mirrored O/S Balance -- which is exactly
    // why the print sheet dropped it -- so the indirection is gone with it.
    set(`${c}${R.provision}`, formula(`${c}${R.os}*VLOOKUP(${c}${R.assetCode},${rateTable},2,FALSE)`), {...rowStyle(R.provision), numFmt:XL_INR_FMT});
    // Total P&L = O/S - Provision (Interest Reversal already flows into
    // Total Dues above, not into Total P&L).
    set(`${c}${R.totalPL}`, formula(`${c}${R.os}-${c}${R.provision}`), {...rowStyle(R.totalPL), numFmt:XL_INR_FMT});
    const otsNum = parseOtsAmount(otsAmounts[s.acctNo]);
    set(`${c}${R.ots}`, otsNum===null ? 0 : otsNum, {border:XL_BORDER_ALL, align:{horizontal:'right'}, font:{bold:true, color:{argb:'FF000000'}}, numFmt:XL_INR_FMT});
    // Total Sacrifice = Total Dues - OTS Amount (Interest Reversal is
    // already folded into Total Dues above, not added a second time).
    set(`${c}${R.totalSac}`, formula(`${c}${R.totalDues}-${c}${R.ots}`), {...rowStyle(R.totalSac), numFmt:XL_INR_FMT});
    set(`${c}${R.ledgerSac}`, formula(`${c}${R.os}-${c}${R.ots}`), {...rowStyle(R.ledgerSac), numFmt:XL_INR_FMT});
    // Impact on P&L gets its own number format with a profit/loss arrow
    // baked into the format string (not XL_INR_FMT, which every other
    // currency cell also uses) -- Excel/Sheets pick the arrow from the
    // formula's live sign, so it stays correct as OTS Amount is edited.
    set(`${c}${R.impact}`, formula(`${c}${R.ots}-${c}${R.totalPL}`), {...rowStyle(R.impact), numFmt:XL_INR_FMT_PL});
  });

  // ---- Aggregate totals: the print sheet's five, in its order (O/S, Dues,
  // OTS, Ledger Sacrifice, Sacrifice). Ledger Sacrifice was missing here
  // entirely, and the order did not match the paper. ----
  const aggTitleRow = headerRow + OTS_XL_ROW_LABELS.length + 2;
  ws.mergeCells(aggTitleRow,1,aggTitleRow,SPAN);
  set(`A${aggTitleRow}`, 'A G G R E G A T E   T O T A L S', {font:{bold:true, size:12, color:{argb:'FF000000'}}, align:{horizontal:'center'}, border:false});
  const sumRange = row => `SUM(B${row}:${lastCol}${row})`;
  const AGG = [
    ['Total O/S Balance', R.os],
    ['Total Dues', R.totalDues],
    ['Total OTS Amount', R.ots],
    ['Total Ledger Sacrifice', R.ledgerSac],
    ['Total Sacrifice', R.totalSac],
  ];
  AGG.forEach(([label,srcRow],i)=>{
    const rr = aggTitleRow+1+i;
    set(`A${rr}`, label, {font:{bold:true, color:{argb:'FF000000'}}, border:false});
    ws.mergeCells(rr,2,rr,SPAN);
    set(`B${rr}`, formula(sumRange(srcRow)), {numFmt:XL_INR_FMT, font:{bold:true, size:12, color:{argb:'FF000000'}}, align:{horizontal:'right'}, border:false});
  });

  // The per-account "scheme · branch" strip that used to sit here is gone:
  // Scheme is a table row now, and the branch already prints in the header,
  // so it was the same two facts repeated once per account.
  const footerRow = aggTitleRow + AGG.length + 2;
  ws.mergeCells(footerRow,1,footerRow,SPAN);
  set(`A${footerRow}`, 'Designed & Developed by ALOK MITTAL · Uttar Pradesh Gramin Bank', {font:{italic:true, size:9.5, color:{argb:'FF666666'}}, align:{horizontal:'center'}, border:false});

  // ---- Column widths + freeze header row/label column ----
  // Every column out to SPAN gets a width, including any beyond the last
  // account -- an unsized column at the right edge reads as a stray blank.
  ws.getColumn(1).width = 30;
  for(let ci = 2; ci <= SPAN; ci++) ws.getColumn(ci).width = 17;
  ws.views = [{state:'frozen', xSplit:1, ySplit:headerRow, topLeftCell:`B${headerRow+1}`, showGridLines:false}];
  wsCalc.getColumn(1).width = 30;
  cols.forEach((c,i)=>{ wsCalc.getColumn(2+i).width = 17; });

  // ---- A4 print setup, so this sheet prints exactly like the PDF -- one
  // page, portrait, scaled to fit regardless of how many accounts (2-4)
  // are linked. Only the "OTS Calculator" sheet is set up this way; the
  // "Calculation Details" sheet is helper data, not meant to be printed.
  ws.pageSetup = {
    paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 1,
    horizontalCentered: true, printTitlesRow: `${headerRow}:${headerRow}`,
    margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
  };
  ws.pageSetup.printArea = `A1:${SPAN_COL}${footerRow}`;

  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const safeName = String(custRow[C.NAME]||'borrower').replace(/[\\/:*?"<>|]/g,' ').replace(/\s+/g,' ').trim().slice(0,40);
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = `OTS_${safeName}_${dateToInputValue(new Date())}.xlsx`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url), 30000);
}
window.exportOtsExcel = exportOtsExcel;

function toggleUpdateModal(show){
  document.getElementById('updateModalOverlay').classList.toggle('show', show);
  closePublishReview();
  if(show){ loadVersionHistory(); renderSpecialNoteList(); updateLokAdalatClearBtn(); }
  if(!show){
    document.getElementById('uploadStatus').innerHTML='';
    document.getElementById('uploadSummary').innerHTML='';
    document.getElementById('applyDataBtn').disabled = true;
    document.getElementById('fileInput').value = '';
    document.getElementById('uploadDropLabel').textContent = 'Tap to choose the daily NPA file';
    renderValidationReport(null);
    const asOnRow = document.getElementById('asOnDateRow');
    if(asOnRow) asOnRow.style.display = 'none';
    __pendingData = null;
    __pendingAsOnDate = null;
    __lastValidation = null;
    const acctInput = document.getElementById('specialNoteAcctInput');
    if(acctInput) acctInput.value = '';
    document.getElementById('specialNoteStatus').innerHTML = '';
    onSpecialNoteAcctInput();
  }
}
function openUpdateModal(){ toggleUpdateModal(true); }

let __pendingData = null;
let __pendingMaster = null;
let __masterFileName = null;
let __pendingAsOnDate = null;
let __lastValidation = null;
// Address List, Customer Master (uploaded on its own), Branch Advance,
// Interest Reversal, and Lok Adalat all apply straight into DATA in memory
// with no separate "Apply" step -- but that's still only this browser
// tab's memory until Publish actually commits it. Alok reported Branch
// Split's Address column coming up completely blank even after uploading
// a fresh Address List; the likely cause, found by re-reading every one of
// these upload handlers, is that none of their success messages ever said
// Publish was still needed -- "applied immediately" reads as "done" to a
// non-technical user, so it's easy to close the tab (or just come back
// another day) never having clicked Publish, silently losing the upload
// exactly like the KCC Overdue bug this same beforeunload guard already
// protects against below. This flag extends that same protection to these
// five uploads too.
let __hasUnpublishedRefData = false;

function xlsxDateToDMY(d){
  return String(d.getUTCDate()).padStart(2,'0')+'-'+String(d.getUTCMonth()+1).padStart(2,'0')+'-'+d.getUTCFullYear();
}
function normalizeCell(v){
  if(v instanceof Date) return xlsxDateToDMY(v);
  if(v===undefined || v===null) return '';
  return v;
}
function findSheet(wb, candidates){
  const names = wb.SheetNames;
  for(const cand of candidates){
    const hit = names.find(n=>n.toLowerCase().replace(/[\s_]/g,'')===cand);
    if(hit) return hit;
  }
  return null;
}

function parseCSV(text){
  const rows = []; let row = [], field = '', inQuotes = false;
  for(let i=0;i<text.length;i++){
    const c = text[i];
    if(inQuotes){
      if(c === '"'){ if(text[i+1] === '"'){ field+='"'; i++; } else inQuotes = false; }
      else field += c;
    } else {
      if(c === '"') inQuotes = true;
      else if(c === ','){ row.push(field); field=''; }
      else if(c === '\n'){ row.push(field); rows.push(row); row=[]; field=''; }
      else if(c === '\r'){ /* skip */ }
      else field += c;
    }
  }
  if(field!=='' || row.length){ row.push(field); rows.push(row); }
  return rows;
}
function normHeader(h){ return String(h||'').toLowerCase().replace(/[^a-z0-9]/g,''); }
function looksScientific(s){ return /^[0-9]+(\.[0-9]+)?e\+?\d+$/i.test(String(s).trim()); }
function expandSci(s){ const n = Number(s); if(!isFinite(n)) return String(s).trim(); return BigInt(Math.round(n)).toString(); }
// Canonical normalization for an ID used purely as a cross-file matching
// key (Address List <-> daily NPA upload, by Account No. or Customer ID)
// -- never for what's actually stored/displayed elsewhere in the app, so
// a real leading zero someone deliberately typed still shows correctly
// on screen. Absorbs the three real-world mismatches that otherwise
// silently break an address match even though both sides plainly mean
// the same account/customer: scientific notation (a large numeric cell
// Excel sometimes displays as e.g. 1.51E+14), a leading zero present
// when a cell was kept as text but dropped when the same value was
// parsed as a number elsewhere, and a trailing ".0"/".00" a spreadsheet
// sometimes adds to a whole number. Added 2026-09-23 after Alok reported
// addresses still not matching even with the Customer ID fallback
// already in place -- mirrors tools/branch-split.html's own normId()
// (this app has no cross-file shared-module setup, so it's a second
// copy, same as every other small helper this tool duplicates rather
// than imports).
function normId(v){
  if(v==null || v==='') return '';
  const s = String(v).trim();
  if(!s) return '';
  if(looksScientific(s)) return expandSci(s);
  const n = Number(s);
  if(isFinite(n) && /^[0-9.]+$/.test(s)) return String(Math.round(n));
  return s;
}
// Re-keys an Account No./Customer ID -> Address map through normId() --
// used right after DATA.addressList/addressListByCustomer are loaded
// (see their own init comment) so a map published under an older version
// of this code, before normId() existed, self-heals on the very next
// load instead of silently staying broken until someone re-uploads.
function renormalizeIdMap(map){
  const out = {};
  Object.keys(map||{}).forEach(k=>{ const nk = normId(k); if(nk) out[nk] = map[k]; });
  return out;
}
// Recovery Dashboard (branch portal): shared Account No. -> Address lookup
// used by KCC Overdue and PNPA Slippage (Alok, 2026-09-25: "har jagah
// table main branch name ki jagah address column add kar do") -- their own
// rows carry no address column of their own. The REAL, populated address
// data on this app lives on DATA.npa.rows' own C.ADDR field (merged in at
// upload time from Customer Master, or from DATA.addressList as a
// fallback there already) -- DATA.addressList itself is frequently empty
// (it's only Alok's separate, optional legacy "Address List" upload), so
// looking it up directly here would silently return "—" for almost every
// account even when the address is known. Built once per DATA.npa.rows
// reference (cheap: a single pass over data already resident in memory)
// and rebuilt automatically the moment that reference changes (a fresh
// upload/publish swaps in a new DATA.npa object rather than mutating the
// old one), so this can never serve a stale map after a refresh.
let __addrByAcctCache = { forRows: null, map: null };
function addressForAcctNo(acctNo){
  const rows = DATA.npa && DATA.npa.rows;
  if(__addrByAcctCache.forRows !== rows){
    const m = new Map();
    (rows||[]).forEach(r=>{ if(r[C.ADDR]) m.set(normId(r[C.ACCT_NO]), r[C.ADDR]); });
    __addrByAcctCache = { forRows: rows, map: m };
  }
  const key = normId(acctNo);
  return __addrByAcctCache.map.get(key) || DATA.addressList[key] || DATA.addressListByCustomer[key] || '';
}

/* ---------- Cleaning rules for mobile / PAN / Aadhar (confirmed against real HO data) ---------- */
function cleanMobile(raw){
  const digits = String(raw==null?'':raw).replace(/\D/g,'');
  let ten = null;
  if(digits.length===10) ten = digits;
  else if(digits.length===12 && digits.slice(0,2)==='91') ten = digits.slice(-10);
  if(ten && /^[6-9]/.test(ten)) return ten;
  return 'N/A';
}
function cleanPan(raw){
  const s = String(raw==null?'':raw).trim().toUpperCase();
  return /^[A-Z]{5}[0-9]{4}[A-Z]$/.test(s) ? s : 'N/A';
}
function cleanAadhar(raw){
  const digits = String(raw==null?'':raw).replace(/\D/g,'');
  return /^\d{12}$/.test(digits) ? digits : 'N/A';
}

/* ---------- As-on date, parsed from the uploaded filename, Admin confirms/edits it ---------- */
function parseAsOnDateFromFilename(filename){
  const name = String(filename||'');
  let m = name.match(/as[_\s]?on[_\s]?(\d{2})(\d{2})(\d{4})/i);
  if(m) return new Date(+m[3], +m[2]-1, +m[1]);
  m = name.match(/(\d{2})[.\-](\d{2})[.\-](\d{4})/);
  if(m) return new Date(+m[3], +m[2]-1, +m[1]);
  return null;
}
function dateToInputValue(d){
  return d.getFullYear()+'-'+String(d.getMonth()+1).padStart(2,'0')+'-'+String(d.getDate()).padStart(2,'0');
}

/* ---------- HO daily file mapping (works for both .csv and .xlsx, multi-region aware) ----------
   Maps the daily "e-AB NPA AC WISE" CBS export (one row per loan account) into the
   internal NPA_COLUMN_COUNT-wide layout, grouping each customer's accounts into slots 1..N. */
function detectHoHeader(headerCells){
  const header = headerCells.map(normHeader);
  return header.indexOf('accountno')>=0 && header.indexOf('customerid')>=0 && header.indexOf('category')>=0;
}
function parseHoDate(v){
  // A stray Date instance here would only ever come from SheetJS's own
  // cellDates:true conversion, which we deliberately no longer request
  // (see the XLSX.read() call sites -- 2026-09-08 fix) precisely because
  // it silently returns a date shifted almost a full day off from the
  // one it was given whenever the browser's local timezone is ahead of
  // UTC (as IST always is): reading a genuine Excel date cell for
  // 15-Mar-2026 through cellDates:true came back reading as 14-Mar in
  // BOTH local and UTC getters, since the underlying instant itself is
  // shifted, not just its interpretation. Excel date-typed cells now
  // arrive here as their raw serial number instead, handled correctly by
  // excelSerialToDate() (a fixed local-midnight epoch + whole-day
  // increments, immune to this since India has no DST). This branch
  // stays only as a defensive fallback for a genuinely-clean Date object
  // constructed elsewhere (e.g. new Date(y,m-1,d) local time).
  if(v instanceof Date) return v;
  if(typeof v==='number') return excelSerialToDate(v);
  if(typeof v==='string' && v.trim()){
    let p = v.trim().split('-');
    if(p.length===3 && p[2].length===4) return new Date(+p[2], +p[1]-1, +p[0]);
    p = v.trim().split('.');
    if(p.length===3 && p[2].length===4) return new Date(+p[2], +p[1]-1, +p[0]);
  }
  return null;
}
function earlierRaw(rawA, rawB){
  const dA = parseHoDate(rawA), dB = parseHoDate(rawB);
  // Format off the already-correct parsed Date rather than re-deriving
  // from the raw cell via normalizeCell() -- normalizeCell() only knows
  // how to turn a Date INSTANCE into text; a raw Excel serial number
  // (which is what genuine date-typed cells are now, per the comment in
  // parseHoDate() above) would otherwise pass straight through
  // unconverted and land in NPA Date as a bare number like "46096"
  // instead of "15-03-2026" -- exactly what CLAUDE.md's date-format rule
  // warns against.
  if(dA && dB) return fmtDate(dA<=dB ? dA : dB);
  if(dA) return fmtDate(dA);
  if(dB) return fmtDate(dB);
  return '';
}
function cellStr(row, i){ return i>=0 ? String(row[i]==null?'':row[i]).trim() : ''; }

function mapHoRowsToNpa(headerCells, dataRows){
  const header = headerCells.map(normHeader);
  const idx = (name) => header.indexOf(normHeader(name));
  const iSol=idx('sol'), iRegion=idx('region'), iBranch=idx('branch'), iAcct=idx('accountno'),
    iCust=idx('customerid'), iScheme=idx('schemecode'), iName=idx('accountname'), iBal=idx('balanceamount'),
    iNpaDate=idx('accountnpadate'), iCustNpaDate=idx('custnpadate'), iSba=idx('sbaaccbalance'),
    iCategory=idx('category'), iSanctDt=idx('sanctiondate'), iLimit=idx('limit'),
    iMobile=idx('mobileno'), iInttRev=idx('inttrev');

  const missing = [];
  if(iAcct<0) missing.push('Account No');
  if(iCust<0) missing.push('Customer ID');
  if(iCategory<0) missing.push('Category');
  if(iBal<0) missing.push('Balance Amount');
  if(iBranch<0) missing.push('Branch');
  if(missing.length){
    throw new Error('Missing required column(s): '+missing.join(', ')+'. Check this file matches the HO "e-AB NPA AC WISE" export layout.');
  }

  let sciCount = 0;
  let badBalCount = 0;
  let blankCustCount = 0;
  const slotCounter = new Map();
  const outRows = [];
  for(const row of dataRows){
    if(!row || row.length<3) continue;
    const acctRaw = cellStr(row, iAcct);
    if(!acctRaw) continue;
    let acctNo = acctRaw;
    if(looksScientific(acctRaw)){ acctNo = expandSci(acctRaw); sciCount++; }
    const custId = cellStr(row, iCust);
    if(!custId){ blankCustCount++; continue; }
    const slot = (slotCounter.get(custId)||0) + 1;
    slotCounter.set(custId, slot);

    const balRaw = row[iBal];
    if(balRaw==null || balRaw==='' || isNaN(parseFloat(balRaw))) badBalCount++;
    const branchRaw = cellStr(row, iBranch);

    let sbAcct='', sbBal='';
    const sbaRaw = cellStr(row, iSba);
    if(sbaRaw.includes('->')){
      const parts = sbaRaw.split('->');
      sbAcct = parts[0].trim();
      sbBal = parseFloat(parts[1]) || 0;
    }
    const cat = cellStr(row, iCategory);
    const region = cellStr(row, iRegion);
    const npaDate = earlierRaw(iNpaDate>=0?row[iNpaDate]:'', iCustNpaDate>=0?row[iCustNpaDate]:'');

    const out = new Array(NPA_COLUMN_COUNT).fill('');
    out[0] = custId+':'+slot; out[2] = slot; out[3] = cellStr(row,iSol); out[4] = branchRaw;
    out[5] = custId; out[6] = acctNo; out[7] = cellStr(row,iName);
    out[9] = cellStr(row,iMobile);
    // fmtDate(parseHoDate(...)) here, not normalizeCell() -- same reason
    // as earlierRaw() above: a genuine date-typed Sanction Date cell now
    // arrives as a raw Excel serial number, which normalizeCell() would
    // otherwise pass straight through unconverted.
    const sanctDt = iSanctDt>=0 ? parseHoDate(row[iSanctDt]) : null;
    out[13] = cellStr(row,iScheme); out[14] = sanctDt ? fmtDate(sanctDt) : normalizeCell(iSanctDt>=0?row[iSanctDt]:'');
    out[15] = parseFloat(row[iLimit])||0; out[16] = parseFloat(row[iBal])||0;
    out[18] = (iInttRev>=0 && row[iInttRev]!=='' && row[iInttRev]!=null) ? (parseFloat(row[iInttRev])||0) : '';
    out[19] = cat; out[20] = npaDate; out[21] = cat; out[22] = npaDate; out[23] = npaDate;
    out[24] = sbAcct; out[25] = sbBal; out[26] = region;
    outRows.push(out);
  }
  return { rows: outRows, sciCount, badBalCount, blankCustCount };
}

/* ---------- Customer Master parsing + merge (Address / Aadhar / PAN, ~80k rows, refreshed rarely) ---------- */
/* Scans the first few rows for the real header (skips title/instruction rows some
   templates — including ours — put above the actual column headers). */
function findHeaderRowIndex(allRows, mustContainAnyNormalized){
  for(let i=0;i<Math.min(10, allRows.length);i++){
    const normed = (allRows[i]||[]).map(normHeader);
    if(mustContainAnyNormalized.some(w=>normed.includes(w))) return i;
  }
  return 0;
}
function buildCustomerMasterMap(headerCells, dataRows){
  const header = headerCells.map(normHeader);
  const idx = (...names) => { for(const n of names){ const i = header.indexOf(normHeader(n)); if(i>=0) return i; } return -1; };
  const iCust = idx('customeridcif','customerid','cif');
  const iAddr = idx('address');
  const iMobile = idx('mobileno','mobile');
  const iAadhar = idx('aadharno','aadhar');
  const iPan = idx('pan');
  if(iCust<0) throw new Error('Customer Master file needs a "Customer ID" column.');
  const map = new Map();
  for(const row of dataRows){
    const cid = cellStr(row, iCust);
    if(!cid || map.has(cid)) continue;
    map.set(cid, {
      address: cellStr(row, iAddr),
      mobile: cleanMobile(row[iMobile]),
      aadhar: cleanAadhar(row[iAadhar]),
      pan: cleanPan(row[iPan]),
    });
  }
  return map;
}
/* Matches the real HO "Daily Follow-up Sheet" layout: a header row with
   plain "Sol ID"/"Branch Name" columns, but the Advance column's own header
   cell just says generic "AMT" -- its real label ("Advances <as-on-date>")
   lives in a merged cell 1-3 rows above, since the date changes every time
   this file is refreshed. Falls back to a plain "Total Advance" column
   directly in the header row for a manually-filled template. Matches
   branches by Sol ID (a stable numeric code), not branch name, since the
   same branch can appear under different name spellings/abbreviations
   across different HO reports (e.g. "MURSAN GATE" vs "M.G.Hathras") --
   Sol ID is the one thing guaranteed to match the NPA data's own Sol ID
   column. Figures are entered in the same unit UPGB already reports them
   in, Lakhs, and converted to plain rupees here to match the NPA data's
   units. NPA March/June are optional (older, simpler advance-only files
   still work) -- matched by prefix ("npamarch"/"npajune") since the
   header's own year suffix moves forward every year (MARCH 26 -> 27 -> ...). */
function buildBranchAdvanceMap(allRows, hIdx){
  const header = (allRows[hIdx]||[]).map(normHeader);
  const idx = (...names) => { for(const n of names){ const i = header.indexOf(normHeader(n)); if(i>=0) return i; } return -1; };
  const idxPrefix = (name) => header.findIndex(h=>h.startsWith(normHeader(name)));
  const iSol = idx('solid','sol');
  if(iSol<0) throw new Error('Could not find a "Sol ID" column -- branches are matched by Sol ID, not name, since branch names vary between reports.');
  const iBranchName = idx('branchname','branch');
  let iAdv = idx('totaladvance','advance','advancelakhs','totaladvancelakhs');
  if(iAdv<0){
    for(let r=Math.max(0,hIdx-3); r<hIdx && iAdv<0; r++){
      const row = allRows[r]||[];
      for(let c=0;c<row.length;c++){
        if(/^advances?\b/i.test(String(row[c]||'').trim())){ iAdv = c; break; }
      }
    }
  }
  if(iAdv<0) throw new Error('Could not find an "Advances" column (checked the header row and the few rows above it).');
  const iNpaMar = idxPrefix('npamarch');
  const iNpaJun = idxPrefix('npajune');
  const toRupees = (v) => {
    const lakhs = parseFloat(String(v==null?'':v).replace(/[^0-9.\-]/g,''));
    return isNaN(lakhs) ? null : lakhs*100000;
  };
  const map = {};
  for(const row of allRows.slice(hIdx+1)){
    const sol = cellStr(row, iSol);
    if(!sol) continue;
    const adv = toRupees(row[iAdv]);
    if(adv===null || adv<=0) continue;
    map[sol] = {
      adv,
      branchName: iBranchName>=0 ? cellStr(row, iBranchName) : '',
      npaMar26: iNpaMar>=0 ? toRupees(row[iNpaMar]) : null,
      npaJun26: iNpaJun>=0 ? toRupees(row[iNpaJun]) : null,
    };
  }
  return map;
}
/* Interest Reversal master list -- Account No. + amount, matched by exact
   Account No. (not Customer ID/Sol ID, since Interest Reversal is an
   account-level figure). A row with an unparseable amount is skipped
   rather than defaulting to 0, so a genuinely blank/malformed cell in
   Alok's file doesn't silently zero out an account that already had a
   real figure from an earlier upload of this same list. */
function buildInterestReversalMasterMap(allRows, hIdx){
  const header = (allRows[hIdx]||[]).map(normHeader);
  const idx = (...names) => { for(const n of names){ const i = header.indexOf(normHeader(n)); if(i>=0) return i; } return -1; };
  const iAcct = idx('accountnumber','accountno','acctno');
  if(iAcct<0) throw new Error('Could not find an "Account Number" column.');
  const iAmt = idx('interestreversal','intreversal','amount');
  if(iAmt<0) throw new Error('Could not find an "Interest Reversal" amount column.');
  const map = {};
  for(const row of allRows.slice(hIdx+1)){
    const acct = cellStr(row, iAcct);
    if(!acct) continue;
    const amt = parseFloat(row[iAmt]);
    if(isNaN(amt)) continue;
    map[acct] = amt;
  }
  return map;
}
/* Address list -- Account No. and/or Customer ID, plus Address. A row
   needs at least one of the two identifier columns; it's fine for the
   whole file to have only one or the other (Alok's real source list is
   Customer ID only -- one address per customer, not re-typed per
   account), or a mix of both column types row to row. Both key types go
   through normId() (scientific notation, leading zeros, a trailing
   ".0"/".00" -- see its own comment), since address lists this size (tens
   of thousands of rows) are exactly the kind of file where those show up,
   and a mismatched key format is enough to silently break an otherwise-
   correct match. A row with a blank address is skipped rather than
   storing an empty string, so it can never overwrite a real address a
   future upload of this same list omits by mistake. Returns both maps
   separately (never merged into one) so a caller can prefer the more
   precise Account No. match and only fall back to Customer ID when that
   misses. */
function buildAddressListMap(allRows, hIdx){
  const header = (allRows[hIdx]||[]).map(normHeader);
  const idx = (...names) => { for(const n of names){ const i = header.indexOf(normHeader(n)); if(i>=0) return i; } return -1; };
  const iAcct = idx('accountnumber','accountno','acctno','account');
  const iCust = idx('customerid','custid','customerno','custno');
  if(iAcct<0 && iCust<0) throw new Error('Could not find an "Account Number" or "Customer ID" column.');
  const iAddr = idx('address');
  if(iAddr<0) throw new Error('Could not find an "Address" column.');
  const byAccount = {}, byCustomer = {};
  for(const row of allRows.slice(hIdx+1)){
    const addr = cellStr(row, iAddr);
    if(!addr) continue;
    if(iAcct>=0){
      const acct = normId(cellStr(row, iAcct));
      if(acct) byAccount[acct] = addr;
    }
    if(iCust>=0){
      const cust = normId(cellStr(row, iCust));
      if(cust) byCustomer[cust] = addr;
    }
  }
  return { byAccount, byCustomer };
}
/* Branch Contacts (Manager + Recovery Officer) template/upload -- matches
   branches by Sol ID same as buildBranchAdvanceMap above, not by name.
   Every contact field is optional (collection is ongoing); a row with a
   Sol ID but nothing else is simply skipped rather than stored empty. */
function buildBranchContactsMap(allRows, hIdx){
  const header = (allRows[hIdx]||[]).map(normHeader);
  const idx = (...names) => { for(const n of names){ const i = header.indexOf(normHeader(n)); if(i>=0) return i; } return -1; };
  const iSol = idx('solid','sol');
  if(iSol<0) throw new Error('Could not find a "Sol ID" column -- branches are matched by Sol ID, not name.');
  const iMgr = idx('branchmanagername','managername','branchmanager');
  const iMgrMobile = idx('managermobileno','managermobile');
  const iMgrEmail = idx('manageremailid','manageremail');
  const iRoName = idx('recoveryofficername','recoveryofficer');
  const iRoMobile = idx('recoveryofficermobileno','recoveryofficermobile');
  const iLandline = idx('branchlandlineno','landlineno','landline');
  const iCategory = idx('branchcategory','category');
  const iAddress = idx('branchaddress','address');
  const iIfsc = idx('ifsccode','ifsc');
  const iRemarks = idx('remarks');
  const map = {};
  let count = 0;
  for(const row of allRows.slice(hIdx+1)){
    const sol = cellStr(row, iSol);
    if(!sol) continue;
    const rec = {
      mgr: iMgr>=0 ? cellStr(row, iMgr) : '',
      mgrMobile: iMgrMobile>=0 ? cellStr(row, iMgrMobile) : '',
      mgrEmail: iMgrEmail>=0 ? cellStr(row, iMgrEmail) : '',
      roName: iRoName>=0 ? cellStr(row, iRoName) : '',
      roMobile: iRoMobile>=0 ? cellStr(row, iRoMobile) : '',
      landline: iLandline>=0 ? cellStr(row, iLandline) : '',
      category: iCategory>=0 ? cellStr(row, iCategory) : '',
      address: iAddress>=0 ? cellStr(row, iAddress) : '',
      ifsc: iIfsc>=0 ? cellStr(row, iIfsc) : '',
      remarks: iRemarks>=0 ? cellStr(row, iRemarks) : '',
    };
    Object.keys(rec).forEach(k=>{ if(!rec[k]) delete rec[k]; });
    if(!Object.keys(rec).length) continue;
    map[sol] = rec;
    count++;
  }
  if(!count) throw new Error('No rows with a Sol ID and at least one contact field found.');
  return map;
}
/* Lok Adalat proposed-OTS list -- matched by Account No. (not Sol ID),
   since this tracks specific borrower accounts already in the NPA book,
   not branch-level figures. OTS Amount is optional (blank until fixed);
   a row with neither OTS Amount nor Token Amount is skipped, since
   there's nothing to show. Multiple rows for the same account (part-
   payments of the token on different dates) are combined: Token Amount
   sums across all of them, while OTS Amount/Remark/Date each take
   whichever row's value comes last in the file (OTS only gets fixed
   once; a later remark or date supersedes an earlier one). Date is
   stored already formatted DD-MM-YYYY (CLAUDE.md's date-format rule),
   not a raw serial/Date object, since this map round-trips through the
   publish JSON. */
function buildLokAdalatMap(allRows, hIdx){
  const header = (allRows[hIdx]||[]).map(normHeader);
  const idx = (...names) => { for(const n of names){ const i = header.indexOf(normHeader(n)); if(i>=0) return i; } return -1; };
  const iAcct = idx('accountno','acno','acctno','account');
  if(iAcct<0) throw new Error('Could not find an "Account No." column.');
  const iDate = idx('date');
  const iOts = idx('otsamount','ots');
  const iToken = idx('tokenamount','token');
  const iRemark = idx('remark','remarks');
  const toRupees = (v) => {
    if(v===''||v==null) return null;
    const n = parseFloat(String(v).replace(/[^0-9.\-]/g,''));
    return isNaN(n) ? null : n;
  };
  const map = {};
  for(const row of allRows.slice(hIdx+1)){
    let acctRaw = cellStr(row, iAcct);
    if(!acctRaw) continue;
    if(looksScientific(acctRaw)) acctRaw = expandSci(acctRaw);
    const ots = iOts>=0 ? toRupees(row[iOts]) : null;
    const token = iToken>=0 ? toRupees(row[iToken]) : null;
    if(ots===null && token===null) continue;
    const remark = iRemark>=0 ? cellStr(row, iRemark) : '';
    const dateObj = iDate>=0 ? toDate(row[iDate]) : null;
    if(!map[acctRaw]) map[acctRaw] = { date: '', ots: 0, token: 0, remark: '' };
    const rec = map[acctRaw];
    if(token!==null) rec.token += token;
    if(ots!==null) rec.ots = ots;
    if(remark) rec.remark = remark;
    if(dateObj) rec.date = fmtDate(dateObj);
  }
  return map;
}
async function handleBranchContactsUpload(evt){
  const file = evt.target.files[0];
  if(!file) return;
  await ensureXLSX();
  const labelEl = document.getElementById('branchContactsUploadDropLabel');
  if(labelEl) labelEl.textContent = file.name;
  const statusEl = document.getElementById('branchContactsUploadStatus');
  statusEl.innerHTML = `<div class="upload-status info">Reading Branch Contacts file…</div>`;
  const isCsv = /\.csv$/i.test(file.name);
  const reader = new FileReader();
  reader.onerror = function(){ statusEl.innerHTML = `<div class="upload-status err">⚠ Failed to read the file from disk.</div>`; };
  reader.onload = function(e){
    try{
      const headerHints = ['solid','sol'];
      let allRows, hIdx;
      if(isCsv){
        allRows = parseCSV(String(e.target.result));
        hIdx = findHeaderRowIndex(allRows, headerHints);
      } else {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        allRows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1, raw:true, defval:''});
        hIdx = findHeaderRowIndex(allRows, headerHints);
      }
      const map = buildBranchContactsMap(allRows, hIdx);
      const count = Object.keys(map).length;
      DATA.branchContacts = map;
      const label = document.getElementById('branchContactsStatusLabel');
      if(label) label.textContent = `${count.toLocaleString('en-IN')} branch(es) loaded (${file.name})`;
      statusEl.innerHTML = `<div class="upload-status ok">✔ ${count.toLocaleString('en-IN')} branch contact record(s) parsed. Tap a branch in the Branch/Sol ID panel to see the full card. Not live for anyone else until you hit Publish below.</div>`;
      __hasUnpublishedRefData = true;
      clearStalePublishStatus();
      const publishBtn = document.getElementById('publishBtn');
      if(publishBtn) publishBtn.disabled = false;
      if(document.getElementById('branchEdgePanel')?.classList.contains('open')) filterBranchList();
    } catch(err){
      statusEl.innerHTML = `<div class="upload-status err">⚠ Could not read this file: ${esc(err.message||err)}</div>`;
    }
  };
  if(isCsv) reader.readAsText(file); else reader.readAsArrayBuffer(file);
}
async function handleLokAdalatUpload(evt){
  const file = evt.target.files[0];
  if(!file) return;
  await ensureXLSX();
  const labelEl = document.getElementById('lokAdalatUploadDropLabel');
  if(labelEl) labelEl.textContent = file.name;
  const statusEl = document.getElementById('lokAdalatUploadStatus');
  statusEl.innerHTML = `<div class="upload-status info">Reading Lok Adalat file…</div>`;
  const isCsv = /\.csv$/i.test(file.name);
  const reader = new FileReader();
  reader.onerror = function(){ statusEl.innerHTML = `<div class="upload-status err">⚠ Failed to read the file from disk.</div>`; };
  reader.onload = function(e){
    try{
      const headerHints = ['accountno','acno','acctno','account'];
      let allRows, hIdx;
      if(isCsv){
        allRows = parseCSV(String(e.target.result));
        hIdx = findHeaderRowIndex(allRows, headerHints);
      } else {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        allRows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1, raw:true, defval:''});
        hIdx = findHeaderRowIndex(allRows, headerHints);
      }
      const map = buildLokAdalatMap(allRows, hIdx);
      const count = Object.keys(map).length;
      if(!count) throw new Error('No rows with an Account No. and at least one of OTS Amount/Token Amount found.');
      DATA.lokAdalat = map;
      const label = document.getElementById('lokAdalatStatusLabel');
      if(label) label.textContent = `${count.toLocaleString('en-IN')} account(s) loaded (${file.name})`;
      statusEl.innerHTML = `<div class="upload-status ok">✔ ${count.toLocaleString('en-IN')} Lok Adalat account(s) parsed. Opening one of these accounts now shows the "already received" banner. Not live for anyone else until you hit Publish below.</div>`;
      __hasUnpublishedRefData = true;
      clearStalePublishStatus();
      const publishBtn = document.getElementById('publishBtn');
      if(publishBtn) publishBtn.disabled = false;
      updateLokAdalatClearBtn();
    } catch(err){
      statusEl.innerHTML = `<div class="upload-status err">⚠ Could not read this file: ${esc(err.message||err)}</div>`;
    }
  };
  if(isCsv) reader.readAsText(file); else reader.readAsArrayBuffer(file);
}
function updateLokAdalatClearBtn(){
  const btn = document.getElementById('lokAdalatClearBtn');
  if(btn) btn.disabled = !Object.keys(DATA.lokAdalat||{}).length;
}
/* Bulk "Clear List" -- the only way to empty this list used to be uploading
   a fresh file with zero matching rows, which was clumsy for the common
   case of "this Lok Adalat is over, wipe it." Mirrors the tail end of a
   successful handleLokAdalatUpload() above, just with an empty map instead
   of a parsed one. */
function clearLokAdalat(){
  if(!Object.keys(DATA.lokAdalat||{}).length) return;
  DATA.lokAdalat = {};
  const label = document.getElementById('lokAdalatStatusLabel');
  if(label) label.textContent = 'not loaded yet';
  const dropLabel = document.getElementById('lokAdalatUploadDropLabel');
  if(dropLabel) dropLabel.textContent = 'Tap to choose the Lok Adalat file';
  const fileInput = document.getElementById('lokAdalatFileInput');
  if(fileInput) fileInput.value = '';
  const statusEl = document.getElementById('lokAdalatUploadStatus');
  if(statusEl) statusEl.innerHTML = `<div class="upload-status ok">✔ Lok Adalat list cleared. Not live for anyone else until you hit Publish below.</div>`;
  __hasUnpublishedRefData = true;
  clearStalePublishStatus();
  const publishBtn = document.getElementById('publishBtn');
  if(publishBtn) publishBtn.disabled = false;
  updateLokAdalatClearBtn();
}

/* Special Note -- Admin types an Account No. one at a time (no file
   upload), sees the matching name/branch immediately via the same
   npaByAcct index Search uses, types a note and saves it. Writes straight
   into DATA.specialNotes and unlocks Publish, the same "no separate Apply
   step" pattern handleBranchContactsUpload above uses for a small
   side-dataset -- goes live for every viewer, not just Admin, the next
   time Publish to Live Site is pressed. Displayed on that account's Loan
   Detail screen by drawDetailBody's specialNoteBanner block. */
function onSpecialNoteAcctInput(){
  const input = document.getElementById('specialNoteAcctInput');
  const lookupEl = document.getElementById('specialNoteLookup');
  const textEl = document.getElementById('specialNoteText');
  const removeBtn = document.getElementById('specialNoteRemoveBtn');
  const saveBtn = document.getElementById('specialNoteSaveBtn');
  if(!input || !lookupEl || !textEl || !removeBtn || !saveBtn) return;
  const acctNo = input.value.trim();
  if(!acctNo){
    lookupEl.innerHTML = '';
    textEl.value = '';
    textEl.disabled = true;
    removeBtn.disabled = true;
    saveBtn.disabled = true;
    saveBtn.textContent = 'Save Note';
    return;
  }
  const row = npaByAcct.get(acctNo);
  if(!row){
    lookupEl.innerHTML = `<div class="special-note-lookup-row notfound">⚠ Account not found in current NPA data</div>`;
    textEl.value = '';
    textEl.disabled = true;
    removeBtn.disabled = true;
    saveBtn.disabled = true;
    saveBtn.textContent = 'Save Note';
    return;
  }
  const existing = (DATA.specialNotes||{})[acctNo];
  lookupEl.innerHTML = `<div class="special-note-lookup-row found">✔ ${esc(row[C.NAME])||'—'} · ${esc(row[C.SOL_DESC])||'—'}</div>`;
  textEl.disabled = false;
  textEl.value = existing ? existing.note : '';
  removeBtn.disabled = !existing;
  saveBtn.textContent = existing ? 'Update Note' : 'Save Note';
  onSpecialNoteTextInput();
}
function onSpecialNoteTextInput(){
  const textEl = document.getElementById('specialNoteText');
  const saveBtn = document.getElementById('specialNoteSaveBtn');
  if(!textEl || !saveBtn) return;
  saveBtn.disabled = textEl.disabled || !textEl.value.trim();
}
function saveSpecialNote(){
  const input = document.getElementById('specialNoteAcctInput');
  const textEl = document.getElementById('specialNoteText');
  const statusEl = document.getElementById('specialNoteStatus');
  if(!input || !textEl) return;
  const acctNo = input.value.trim();
  const note = textEl.value.trim();
  const row = npaByAcct.get(acctNo);
  if(!acctNo || !row || !note) return;
  const user = (window.UPGBAuth && window.UPGBAuth.getCurrentUser()) || {};
  DATA.specialNotes = DATA.specialNotes || {};
  DATA.specialNotes[acctNo] = { note, updatedAt: new Date().toISOString(), updatedBy: user.login || null };
  if(statusEl) statusEl.innerHTML = `<div class="upload-status ok">✔ Note saved for A/c ${esc(acctNo)} — ${esc(row[C.NAME])||'—'}. Goes live for everyone on Publish.</div>`;
  __hasUnpublishedRefData = true;
  clearStalePublishStatus();
  const publishBtn = document.getElementById('publishBtn');
  if(publishBtn) publishBtn.disabled = false;
  onSpecialNoteAcctInput();
  renderSpecialNoteList();
}
function removeSpecialNote(){
  const input = document.getElementById('specialNoteAcctInput');
  if(input) removeSpecialNoteByAcct(input.value.trim());
}
function removeSpecialNoteByAcct(acctNo){
  if(!acctNo || !DATA.specialNotes || !DATA.specialNotes[acctNo]) return;
  const row = npaByAcct.get(acctNo);
  delete DATA.specialNotes[acctNo];
  const statusEl = document.getElementById('specialNoteStatus');
  if(statusEl) statusEl.innerHTML = `<div class="upload-status ok">✔ Note removed for A/c ${esc(acctNo)}${row?(' — '+esc(row[C.NAME])):''}. Goes live for everyone on Publish.</div>`;
  __hasUnpublishedRefData = true;
  clearStalePublishStatus();
  const publishBtn = document.getElementById('publishBtn');
  if(publishBtn) publishBtn.disabled = false;
  const input = document.getElementById('specialNoteAcctInput');
  if(input && input.value.trim()===acctNo) onSpecialNoteAcctInput();
  renderSpecialNoteList();
}
function loadSpecialNoteIntoEditor(acctNo){
  const input = document.getElementById('specialNoteAcctInput');
  if(!input) return;
  input.value = acctNo;
  onSpecialNoteAcctInput();
  input.focus();
}
function renderSpecialNoteList(){
  const listEl = document.getElementById('specialNoteList');
  const countEl = document.getElementById('specialNoteCountLabel');
  const notes = DATA.specialNotes || {};
  const acctNos = Object.keys(notes);
  if(countEl) countEl.textContent = acctNos.length ? `${acctNos.length.toLocaleString('en-IN')} saved` : 'none saved';
  if(!listEl) return;
  if(!acctNos.length){ listEl.innerHTML = ''; return; }
  listEl.innerHTML = acctNos.map(acctNo => {
    const row = npaByAcct.get(acctNo);
    const name = row ? (row[C.NAME]||'—') : 'not in current NPA data';
    return `<div class="special-note-list-item">
      <div class="special-note-list-body">
        <div class="special-note-list-acct">A/c ${esc(acctNo)} <span>${esc(name)}</span></div>
        <div class="special-note-list-text">${esc(notes[acctNo].note)}</div>
      </div>
      <button type="button" class="special-note-list-edit" onclick="loadSpecialNoteIntoEditor('${esc(acctNo)}')" aria-label="Edit note for account ${esc(acctNo)}">Edit</button>
      <button type="button" class="special-note-list-remove" onclick="removeSpecialNoteByAcct('${esc(acctNo)}')" aria-label="Remove note for account ${esc(acctNo)}">✕</button>
    </div>`;
  }).join('');
}
window.loadSpecialNoteIntoEditor = loadSpecialNoteIntoEditor;
window.removeSpecialNoteByAcct = removeSpecialNoteByAcct;

function carryForwardMapFromCurrentData(){
  const map = new Map();
  DATA.npa.rows.forEach(r=>{
    const cid = String(r[C.CUST_ID]||'');
    if(!cid || map.has(cid)) return;
    map.set(cid, { address:r[C.ADDR]||'', mobile:r[C.PHONE]||'', aadhar:r[C.AADHAR]||'', pan:r[C.PAN]||'' });
  });
  return map;
}
function mergeCustomerDetails(npaRows, masterMap, carryForwardMap){
  npaRows.forEach(r=>{
    const cid = String(r[C.CUST_ID]||'');
    const fresh = masterMap ? masterMap.get(cid) : null;
    const prior = carryForwardMap ? carryForwardMap.get(cid) : null;
    const src = fresh || prior;
    r[C.ADDR] = src ? src.address : '';
    r[C.AADHAR] = src ? src.aadhar : 'N/A';
    r[C.PAN] = src ? src.pan : 'N/A';
    const dailyMobileClean = cleanMobile(r[C.PHONE]);
    r[C.PHONE] = dailyMobileClean!=='N/A' ? dailyMobileClean : ((src && src.mobile && src.mobile!=='N/A') ? src.mobile : 'N/A');
    // Final fallback: Customer Master and carry-forward both come up empty
    // for a genuinely new/never-seen account -- the app-wide, admin-
    // uploaded Address List (see its own comment near DATA.addressList) is
    // checked last: by Account No. first (more precise), then by Customer
    // ID if that misses (Alok's real source list is Customer-ID-only), and
    // only ever fills a blank, never overwrites a real address the steps
    // above already supplied.
    if(!r[C.ADDR]){
      const la = DATA.addressList[normId(r[C.ACCT_NO])] || DATA.addressListByCustomer[normId(r[C.CUST_ID])];
      if(la) r[C.ADDR] = la;
    }
  });
}
/* Re-applies the same DATA.addressList fallback mergeCustomerDetails()
   already does going forward, but as its own pass over whatever's already
   loaded -- called right after a fresh Address List upload so accounts
   already on screen get fixed immediately instead of waiting for the next
   daily NPA upload to pick it up. */
function applyAddressListFallback(npaRows){
  const byAcct = DATA.addressList, byCust = DATA.addressListByCustomer;
  if(!byAcct && !byCust) return;
  npaRows.forEach(r=>{
    if(r[C.ADDR]) return;
    const addr = (byAcct && byAcct[normId(r[C.ACCT_NO])]) || (byCust && byCust[normId(r[C.CUST_ID])]);
    if(addr) r[C.ADDR] = addr;
  });
}

/* ---------- Validation engine: run before "Apply Update" is enabled ---------- */
function validateNpaRows(rows){
  const errors = [], warnings = [];
  const acctSeen = new Set();
  let dupCount=0, blankBranch=0, blankCust=0, badBal=0, badNpaDate=0, badSanctDate=0;
  rows.forEach(r=>{
    const acct = String(r[C.ACCT_NO]||'');
    if(acct){ if(acctSeen.has(acct)) dupCount++; else acctSeen.add(acct); }
    if(!r[C.SOL_DESC]) blankBranch++;
    if(!r[C.CUST_ID]) blankCust++;
    if(r[C.OUTBAL]===''||r[C.OUTBAL]==null||isNaN(r[C.OUTBAL])) badBal++;
    if(r[C.NPA_DT] && !toDate(r[C.NPA_DT])) badNpaDate++;
    if(r[C.SANCT_DT] && !toDate(r[C.SANCT_DT])) badSanctDate++;
  });
  if(dupCount>0) errors.push(`${dupCount.toLocaleString('en-IN')} duplicate Account No. found.`);
  if(blankBranch>0) errors.push(`${blankBranch.toLocaleString('en-IN')} row(s) have a blank Branch.`);
  if(blankCust>0) errors.push(`${blankCust.toLocaleString('en-IN')} row(s) have a blank Customer ID.`);
  if(badBal>0) errors.push(`${badBal.toLocaleString('en-IN')} row(s) have a missing/non-numeric Balance Amount.`);
  if(badNpaDate>0) warnings.push(`${badNpaDate.toLocaleString('en-IN')} row(s) have an NPA date that couldn't be read.`);
  if(badSanctDate>0) warnings.push(`${badSanctDate.toLocaleString('en-IN')} row(s) have a Sanction date that couldn't be read.`);
  return { ok: errors.length===0, errors, warnings, totalRows: rows.length };
}
function renderValidationReport(result){
  const el = document.getElementById('validationReport');
  if(!el) return;
  if(!result){ el.innerHTML=''; return; }
  const cls = result.ok ? 'ok' : 'err';
  const title = result.ok ? '✔ Validation passed' : '⚠ Validation failed — fix the file before applying';
  let html = `<div class="validation-report ${cls}"><h4>${title}</h4>`;
  if(result.errors.length){
    html += `<ul>${result.errors.map(e=>`<li>${esc(e)}</li>`).join('')}</ul>`;
  } else {
    html += `<div style="font-size:12px">${result.totalRows.toLocaleString('en-IN')} rows checked — no duplicate accounts, blank branch/customer/amount, or unreadable dates.</div>`;
  }
  if(result.warnings.length){
    html += `<div style="margin-top:8px;font-size:11.5px;color:var(--sub)">Warnings (won't block Apply):<ul>${result.warnings.map(w=>`<li>${esc(w)}</li>`).join('')}</ul></div>`;
  }
  html += `</div>`;
  el.innerHTML = html;
}

/* Alok's own Interest Reversal list (DATA.interestReversalMaster) is the
   authoritative source, overriding whatever HO's own file carried that day
   for any account it covers -- HO's "Intt Rev" column has proven too
   inconsistent to rely on alone (see DATA.interestReversalMaster's own
   comment above). Accounts the master doesn't cover keep HO's own file's
   value (or blank/0), unchanged. Called from both parsing branches below. */
function applyInterestReversalMaster(rows){
  const master = DATA.interestReversalMaster;
  if(!master) return;
  rows.forEach(r=>{
    const v = master[String(r[C.ACCT_NO])];
    if(v!=null) r[C.URI] = v;
  });
}
function processDailyParsed(parsed, filename, statusEl, summaryEl){
  if(parsed.isHoFormat){
    const {rows, sciCount, badBalCount, blankCustCount} = mapHoRowsToNpa(parsed.header, parsed.rows);
    if(!rows.length) throw new Error('No account rows found in this file.');
    const carryForward = carryForwardMapFromCurrentData();
    mergeCustomerDetails(rows, __pendingMaster, carryForward);
    applyInterestReversalMaster(rows);
    const validation = validateNpaRows(rows);
    if(blankCustCount>0) validation.errors.unshift(`${blankCustCount.toLocaleString('en-IN')} row(s) had a blank Customer ID and were excluded from the upload entirely.`);
    if(badBalCount>0) validation.errors.unshift(`${badBalCount.toLocaleString('en-IN')} row(s) have a missing/non-numeric Balance Amount.`);
    validation.ok = validation.errors.length===0;
    __lastValidation = validation;
    __pendingData = { npa: {headers: DATA.npa.headers, rows}, oldots: DATA.oldots };
    renderValidationReport(validation);

    const sciPct = rows.length ? sciCount/rows.length : 0;
    if(sciPct > 0.3){
      statusEl.innerHTML = `<div class="upload-status err">⚠ ${sciCount.toLocaleString('en-IN')} of ${rows.length.toLocaleString('en-IN')} account numbers in this file are stored in scientific notation (e.g. 1.51E+14) — the CBS export truncates them, so Account No. search/display will be unreliable after applying. Customer ID and Mobile No. search still work fine. Ask for the "Account No" column to be exported as plain text/number to fix this at the source.</div>`;
    } else {
      statusEl.innerHTML = `<div class="upload-status ok">✔ Parsed successfully. Review below, then Apply.</div>` +
        (sciCount ? `<div class="upload-status err" style="margin-top:8px">⚠ ${sciCount.toLocaleString('en-IN')} account number(s) were stored in scientific notation and may be missing trailing digits.</div>` : '');
    }
    summaryEl.innerHTML = `
      <div class="upload-summary">
        <div class="box"><div class="k">Loan accounts found</div><div class="v">${rows.length.toLocaleString('en-IN')}</div></div>
      </div>`;
    document.getElementById('applyDataBtn').disabled = !validation.ok;

    const guessed = parseAsOnDateFromFilename(filename);
    const row = document.getElementById('asOnDateRow');
    const input = document.getElementById('asOnDateInput');
    const hint = document.getElementById('asOnDateHint');
    if(row && input){
      row.style.display = 'flex';
      if(guessed){
        input.value = dateToInputValue(guessed);
        hint.textContent = '(read from the filename — adjust if this looks wrong)';
      } else if(!input.value){
        input.value = dateToInputValue(new Date());
        hint.textContent = "(couldn't read a date from the filename — please set it)";
      }
      __pendingAsOnDate = input.value;
    }
  } else {
    const wb = parsed.wb;
    const npaSheetName = findSheet(wb, ['npa']);
    if(!npaSheetName){
      throw new Error('This doesn\'t match the daily HO export layout, and no sheet named "NPA" was found for the legacy format either.');
    }
    const npaWs = wb.Sheets[npaSheetName];
    const npaRaw = XLSX.utils.sheet_to_json(npaWs, {header:1, raw:true, defval:''});
    const npaHeaders = (npaRaw[0]||[]).slice(0,NPA_COLUMN_COUNT).map(h=>String(h||''));
    const npaRows = npaRaw.slice(1)
      .filter(r=>r[6]!=='' && r[6]!==undefined && r[6]!==null)
      .map(r=>{ const row=[]; for(let i=0;i<NPA_COLUMN_COUNT;i++) row.push(normalizeCell(r[i])); return row; });
    applyInterestReversalMaster(npaRows);

    let oldOtsRows = [];
    const oldOtsSheetName = findSheet(wb, ['oldots']);
    if(oldOtsSheetName){
      const oldWs = wb.Sheets[oldOtsSheetName];
      const oldRaw = XLSX.utils.sheet_to_json(oldWs, {header:1, raw:true, defval:''});
      oldOtsRows = oldRaw.slice(1)
        .filter(r=>r[0]!=='' && r[0]!==undefined && r[0]!==null)
        .map(r=>[normalizeCell(r[0]), normalizeCell(r[1]), normalizeCell(r[2])]);
    }
    const validation = validateNpaRows(npaRows);
    __lastValidation = validation;
    __pendingData = { npa: {headers: npaHeaders, rows: npaRows}, oldots: {headers:['Account Number','Date','Amount'], rows: oldOtsRows} };
    renderValidationReport(validation);
    statusEl.innerHTML = `<div class="upload-status ok">✔ Parsed successfully (legacy workbook format). Review below, then Apply.</div>`;
    summaryEl.innerHTML = `
      <div class="upload-summary">
        <div class="box"><div class="k">NPA rows found</div><div class="v">${npaRows.length.toLocaleString('en-IN')}</div></div>
        <div class="box"><div class="k">OLD OTS rows found</div><div class="v">${oldOtsRows.length.toLocaleString('en-IN')}</div></div>
      </div>`;
    document.getElementById('applyDataBtn').disabled = !validation.ok;
    const row = document.getElementById('asOnDateRow');
    if(row) row.style.display = 'none';
    __pendingAsOnDate = null;
  }
}

async function handleFileUpload(evt){
  const file = evt.target.files[0];
  if(!file) return;
  await ensureXLSX();
  document.getElementById('uploadDropLabel').textContent = file.name;
  const statusEl = document.getElementById('uploadStatus');
  const summaryEl = document.getElementById('uploadSummary');
  summaryEl.innerHTML = '';
  renderValidationReport(null);
  document.getElementById('applyDataBtn').disabled = true;
  const isCsv = /\.csv$/i.test(file.name);
  statusEl.innerHTML = `<div class="upload-status info">Reading file…</div>`;
  const reader = new FileReader();
  reader.onerror = function(){ statusEl.innerHTML = `<div class="upload-status err">⚠ Failed to read the file from disk.</div>`; };
  reader.onload = function(e){
    try{
      let parsed;
      if(isCsv){
        const csvRows = parseCSV(String(e.target.result));
        parsed = { header: csvRows[0]||[], rows: csvRows.slice(1), isHoFormat: true };
      } else {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        const firstRaw = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1, raw:true, defval:''});
        const header = firstRaw[0]||[];
        parsed = { header, rows: firstRaw.slice(1), isHoFormat: detectHoHeader(header), wb };
      }
      processDailyParsed(parsed, file.name, statusEl, summaryEl);
    } catch(err){
      statusEl.innerHTML = `<div class="upload-status err">⚠ Could not read this file: ${esc(err.message||err)}</div>`;
    }
  };
  if(isCsv) reader.readAsText(file); else reader.readAsArrayBuffer(file);
}

async function handleMasterFileUpload(evt){
  const file = evt.target.files[0];
  if(!file) return;
  await ensureXLSX();
  const labelEl = document.getElementById('masterUploadDropLabel');
  if(labelEl) labelEl.textContent = file.name;
  const statusEl = document.getElementById('masterUploadStatus');
  statusEl.innerHTML = `<div class="upload-status info">Reading Customer Master…</div>`;
  const isCsv = /\.csv$/i.test(file.name);
  const reader = new FileReader();
  reader.onerror = function(){ statusEl.innerHTML = `<div class="upload-status err">⚠ Failed to read the file from disk.</div>`; };
  reader.onload = function(e){
    try{
      const headerHints = ['customeridcif','customerid','cif'];
      let header, rows;
      if(isCsv){
        const csvRows = parseCSV(String(e.target.result));
        const hIdx = findHeaderRowIndex(csvRows, headerHints);
        header = csvRows[hIdx]||[]; rows = csvRows.slice(hIdx+1);
      } else {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        const sheetName = wb.SheetNames.find(n=>!/field\s*reference/i.test(n)) || wb.SheetNames[0];
        const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {header:1, raw:true, defval:''});
        const hIdx = findHeaderRowIndex(raw, headerHints);
        header = raw[hIdx]||[]; rows = raw.slice(hIdx+1);
      }
      __pendingMaster = buildCustomerMasterMap(header, rows);
      __masterFileName = file.name;
      const label = document.getElementById('masterStatusLabel');
      if(label) label.textContent = `${__pendingMaster.size.toLocaleString('en-IN')} customers loaded (${file.name})`;
      if(__pendingData){
        // A fresh daily NPA file is already staged this session -- merge into
        // that pending row set, same as before, so Apply Update ships both
        // together.
        const carryForward = carryForwardMapFromCurrentData();
        mergeCustomerDetails(__pendingData.npa.rows, __pendingMaster, carryForward);
        const validation = validateNpaRows(__pendingData.npa.rows);
        __lastValidation = validation;
        renderValidationReport(validation);
        document.getElementById('applyDataBtn').disabled = !validation.ok;
        statusEl.innerHTML = `<div class="upload-status ok">✔ ${__pendingMaster.size.toLocaleString('en-IN')} customer record(s) parsed and merged into the pending daily upload.</div>`;
      } else {
        // No daily file pending -- a Customer Master refresh is its own
        // periodic thing (Alok re-uploads it every 6-8 months, independent
        // of any daily file), so it applies immediately against whatever
        // NPA data is currently live, same as Branch Advance/Interest
        // Reversal/Address List already do. Without this, uploading a
        // Customer Master on its own did nothing -- Apply Update only ever
        // enables from a fresh daily-file upload, so a Customer Master
        // refresh here had no way to ever take effect.
        const carryForward = carryForwardMapFromCurrentData();
        mergeCustomerDetails(DATA.npa.rows, __pendingMaster, carryForward);
        __hasUnpublishedRefData = true;
        clearStalePublishStatus();
        const publishBtn = document.getElementById('publishBtn');
        if(publishBtn) publishBtn.disabled = false;
        statusEl.innerHTML = `<div class="upload-status ok">✔ ${__pendingMaster.size.toLocaleString('en-IN')} customer record(s) parsed and applied to this session's data. Not live for anyone else until you hit Publish below.</div>`;
      }
    } catch(err){
      statusEl.innerHTML = `<div class="upload-status err">⚠ Could not read this file: ${esc(err.message||err)}</div>`;
    }
  };
  if(isCsv) reader.readAsText(file); else reader.readAsArrayBuffer(file);
}

/* Total advance is a much slower-moving figure than daily NPA data, so this
   upload applies immediately (no separate Apply step) rather than staging
   alongside the NPA file -- there's no risk of it corrupting account data,
   only of a bad NPA% showing until the next Publish. Uploading always fully
   replaces the previous figures (a stale branch just silently loses its %
   until re-uploaded, rather than guessing which branches carry forward). */
async function handleBranchAdvUpload(evt){
  const file = evt.target.files[0];
  if(!file) return;
  await ensureXLSX();
  const labelEl = document.getElementById('branchAdvUploadDropLabel');
  if(labelEl) labelEl.textContent = file.name;
  const statusEl = document.getElementById('branchAdvUploadStatus');
  statusEl.innerHTML = `<div class="upload-status info">Reading Branch Advance file…</div>`;
  const isCsv = /\.csv$/i.test(file.name);
  const reader = new FileReader();
  reader.onerror = function(){ statusEl.innerHTML = `<div class="upload-status err">⚠ Failed to read the file from disk.</div>`; };
  reader.onload = function(e){
    try{
      const headerHints = ['solid','sol'];
      let allRows, hIdx;
      if(isCsv){
        allRows = parseCSV(String(e.target.result));
        hIdx = findHeaderRowIndex(allRows, headerHints);
      } else {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        const sheetName = wb.SheetNames.find(n=>/daily\s*follow[\s-]*up/i.test(n))
          || wb.SheetNames.find(n=>!/field\s*reference|npa\s*list|holiday|gap/i.test(n))
          || wb.SheetNames[0];
        allRows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {header:1, raw:true, defval:''});
        hIdx = findHeaderRowIndex(allRows, headerHints);
      }
      const map = buildBranchAdvanceMap(allRows, hIdx);
      const count = Object.keys(map).length;
      if(!count) throw new Error('No valid Sol ID/Advance rows found.');
      DATA.branchAdvances = map;
      const label = document.getElementById('branchAdvStatusLabel');
      if(label) label.textContent = `${count.toLocaleString('en-IN')} branch(es) loaded (${file.name})`;
      statusEl.innerHTML = `<div class="upload-status ok">✔ ${count.toLocaleString('en-IN')} branch advance figure(s) parsed. NPA % is now shown on the Dashboard. Not live for anyone else until you hit Publish below.</div>`;
      __hasUnpublishedRefData = true;
      clearStalePublishStatus();
      const publishBtn = document.getElementById('publishBtn');
      if(publishBtn) publishBtn.disabled = false;
      if(document.querySelector('.view.active')?.dataset.view==='dashboard') renderDashboard();
    } catch(err){
      statusEl.innerHTML = `<div class="upload-status err">⚠ Could not read this file: ${esc(err.message||err)}</div>`;
    }
  };
  if(isCsv) reader.readAsText(file); else reader.readAsArrayBuffer(file);
}

/* Interest Reversal is Alok's own list, not HO's daily file -- so, same as
   Branch Advance above, this applies immediately (no separate Apply step)
   and always fully replaces the previous list. Once loaded, it's applied
   automatically to every future daily NPA upload by applyInterestReversalMaster()
   (see processDailyParsed()), so it never needs re-uploading alongside a
   routine daily file -- only when Alok has an updated/corrected list. */
async function handleInterestReversalMasterUpload(evt){
  const file = evt.target.files[0];
  if(!file) return;
  await ensureXLSX();
  const labelEl = document.getElementById('intReversalMasterUploadDropLabel');
  if(labelEl) labelEl.textContent = file.name;
  const statusEl = document.getElementById('intReversalMasterUploadStatus');
  statusEl.innerHTML = `<div class="upload-status info">Reading Interest Reversal file…</div>`;
  const isCsv = /\.csv$/i.test(file.name);
  const reader = new FileReader();
  reader.onerror = function(){ statusEl.innerHTML = `<div class="upload-status err">⚠ Failed to read the file from disk.</div>`; };
  reader.onload = function(e){
    try{
      const headerHints = ['accountnumber','accountno'];
      let allRows, hIdx;
      if(isCsv){
        allRows = parseCSV(String(e.target.result));
        hIdx = findHeaderRowIndex(allRows, headerHints);
      } else {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        allRows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1, raw:true, defval:''});
        hIdx = findHeaderRowIndex(allRows, headerHints);
      }
      const map = buildInterestReversalMasterMap(allRows, hIdx);
      const count = Object.keys(map).length;
      if(!count) throw new Error('No valid Account Number/Interest Reversal rows found.');
      DATA.interestReversalMaster = map;
      const label = document.getElementById('intReversalMasterStatusLabel');
      if(label) label.textContent = `${count.toLocaleString('en-IN')} account(s) loaded (${file.name})`;
      statusEl.innerHTML = `<div class="upload-status ok">✔ ${count.toLocaleString('en-IN')} Interest Reversal figure(s) parsed. Will apply to every future daily NPA upload automatically. Not live for anyone else until you hit Publish below.</div>`;
      __hasUnpublishedRefData = true;
      clearStalePublishStatus();
      const publishBtn = document.getElementById('publishBtn');
      if(publishBtn) publishBtn.disabled = false;
    } catch(err){
      statusEl.innerHTML = `<div class="upload-status err">⚠ Could not read this file: ${esc(err.message||err)}</div>`;
    }
  };
  if(isCsv) reader.readAsText(file); else reader.readAsArrayBuffer(file);
}

/* Address List is app-wide reference data (see DATA.addressList's own
   comment), not scoped to one upload flow -- so, same as Branch Advance
   and Interest Reversal above, this applies immediately and always fully
   replaces the previous list (both the Account No. and Customer ID maps
   together, even if this particular file only has one of the two column
   types -- re-uploading is a full replace, not a merge, same as every
   other reference-data upload in this app). Also retroactively re-applies
   the fallback to whatever's already loaded in DATA.npa.rows right now,
   so uploading fixes already-visible blank addresses immediately rather
   than only affecting the next daily NPA upload. */
async function handleAddressListUpload(evt){
  const file = evt.target.files[0];
  if(!file) return;
  await ensureXLSX();
  const labelEl = document.getElementById('addressListUploadDropLabel');
  if(labelEl) labelEl.textContent = file.name;
  const statusEl = document.getElementById('addressListUploadStatus');
  statusEl.innerHTML = `<div class="upload-status info">Reading Address List…</div>`;
  const isCsv = /\.csv$/i.test(file.name);
  const reader = new FileReader();
  reader.onerror = function(){ statusEl.innerHTML = `<div class="upload-status err">⚠ Failed to read the file from disk.</div>`; };
  reader.onload = function(e){
    try{
      // Both identifier types are valid header rows -- a Customer-ID-only
      // list (Alok's real source file, 2026-09-23) has no
      // Account Number-like column at all, so the header-row scan has to
      // recognize "Customer ID" too or it never finds row 1 in the first
      // place and buildAddressListMap() below fails for the wrong reason.
      const headerHints = ['accountnumber','accountno','customerid','custid'];
      let allRows, hIdx;
      if(isCsv){
        allRows = parseCSV(String(e.target.result));
        hIdx = findHeaderRowIndex(allRows, headerHints);
      } else {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        allRows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1, raw:true, defval:''});
        hIdx = findHeaderRowIndex(allRows, headerHints);
      }
      const { byAccount, byCustomer } = buildAddressListMap(allRows, hIdx);
      const acctCount = Object.keys(byAccount).length, custCount = Object.keys(byCustomer).length;
      if(!acctCount && !custCount) throw new Error('No valid Account Number/Customer ID + Address rows found.');
      DATA.addressList = byAccount;
      DATA.addressListByCustomer = byCustomer;
      const label = document.getElementById('addressListStatusLabel');
      const countParts = [];
      if(acctCount) countParts.push(`${acctCount.toLocaleString('en-IN')} by account`);
      if(custCount) countParts.push(`${custCount.toLocaleString('en-IN')} by customer`);
      if(label) label.textContent = `${countParts.join(', ')} loaded (${file.name})`;
      statusEl.innerHTML = `<div class="upload-status ok">✔ ${countParts.join(', ')} address(es) parsed and applied to this session (including Branch Split, right now). ⚠ Not saved yet — anyone opening the app fresh, on any device, still sees the OLD address list until you hit Publish below.</div>`;
      applyAddressListFallback(DATA.npa.rows);
      __hasUnpublishedRefData = true;
      clearStalePublishStatus();
      const publishBtn = document.getElementById('publishBtn');
      if(publishBtn) publishBtn.disabled = false;
    } catch(err){
      statusEl.innerHTML = `<div class="upload-status err">⚠ Could not read this file: ${esc(err.message||err)}</div>`;
    }
  };
  if(isCsv) reader.readAsText(file); else reader.readAsArrayBuffer(file);
}

function applyNewData(){
  if(!__pendingData || (__lastValidation && !__lastValidation.ok)) return;
  const applyBtn = document.getElementById('applyDataBtn');
  applyBtn.classList.add('is-loading');
  applyBtn.disabled = true;
  setTimeout(()=>{ applyNewDataNow(); applyBtn.classList.remove('is-loading'); }, 10);
}
function applyNewDataNow(){
  const newRows = __pendingData.npa.rows;
  /* A daily upload is the full current state of the book -- any account no
     longer present has regularized/closed and should disappear, so the
     new file always fully replaces the old data rather than merging. */
  const newAcctSet = new Set(newRows.map(r=>String(r[C.ACCT_NO]||'')));
  const oldAcctSet = new Set((DATA.npa.rows||[]).map(r=>String(r[C.ACCT_NO]||'')));
  const staleRemovedCount = (DATA.npa.rows||[]).filter(r=>!newAcctSet.has(String(r[C.ACCT_NO]||''))).length;
  const newAddedCount = newRows.filter(r=>!oldAcctSet.has(String(r[C.ACCT_NO]||''))).length;

  DATA.npa = { headers: __pendingData.npa.headers, rows: newRows };
  if(__pendingData.oldots) DATA.oldots = __pendingData.oldots;
  if(__pendingAsOnDate) DATA.asOnDate = __pendingAsOnDate;

  npaByAcct.clear(); npaByHelper.clear(); byCustId.clear(); oldOtsByAcct.clear();
  DATA.npa.rows.forEach(r=>{
    if(r[C.ACCT_NO]!=='') npaByAcct.set(String(r[C.ACCT_NO]), r);
    if(r[C.HELPER]!=='') npaByHelper.set(String(r[C.HELPER]), r);
    const cid = String(r[C.CUST_ID]);
    if(cid && !byCustId.has(cid)) byCustId.set(cid, r);
  });
  DATA.oldots.rows.forEach(r=>{
    if(r[0]!=='' && !oldOtsByAcct.has(String(r[0]))) oldOtsByAcct.set(String(r[0]), {date:r[1], amount:r[2]});
  });

  /* A daily upload used to wipe every typed OTS Amount. Now that these are
     saved on the device, wiping would throw away real work each morning --
     so entries are pruned to accounts still present in the new file
     (regularized/closed ones go) and everything else carries forward. */
  [[otsAmounts, saveOtsAmounts], [interestReversalOverrides, saveUriOverrides]].forEach(([map, save])=>{
    Object.keys(map).forEach(acct=>{ if(!newAcctSet.has(String(acct))) delete map[acct]; });
    save();
  });
  updateReportDateDisplay();
  const staleMsg = staleRemovedCount>0 ? ` (${staleRemovedCount.toLocaleString('en-IN')} account(s) from the previous data no longer appear — regularized/closed accounts removed.)` : '';
  const addedMsg = newAddedCount>0 ? ` (${newAddedCount.toLocaleString('en-IN')} new account(s) added.)` : '';
  document.getElementById('uploadStatus').innerHTML = `<div class="upload-status ok">✔ Data updated — ${DATA.npa.rows.length.toLocaleString('en-IN')} NPA rows now active.${staleMsg}${addedMsg}</div>`;
  document.getElementById('downloadAppBtn').disabled = false;
  const publishBtn = document.getElementById('publishBtn');
  if(publishBtn) publishBtn.disabled = false;
  __lastApplyMeta = {
    staleRemovedCount,
    newAddedCount,
    newRowCount: newRows.length,
  };
  document.getElementById('searchHeader').style.display='';
  renderEmpty();
  renderDashboard();
  __pendingData = null;
  __pendingAsOnDate = null;
  // Runs after __pendingData is cleared above -- pendingUnpublishedLabel()
  // reads that flag to mean "uploaded but not yet applied," which is no
  // longer true the instant Apply Update finishes, so recomputing the
  // banner here (rather than before the clear, like every other upload
  // handler does) avoids it showing that stale, now-wrong wording.
  clearStalePublishStatus();
}
let __lastApplyMeta = null;

function fmtAsOnDisplay(){
  if(DATA.asOnDate){
    const parts = DATA.asOnDate.split('-');
    if(parts.length===3) return `${parts[2]}-${parts[1]}-${parts[0]}`;
  }
  return fmtDate(new Date());
}
function updateReportDateDisplay(){
  document.querySelectorAll('.report-date-val').forEach(e=>e.textContent = fmtAsOnDisplay());
}

function csvField(v){
  const s = String(v==null?'':v);
  return /[",\n]/.test(s) ? '"'+s.replace(/"/g,'""')+'"' : s;
}
function downloadCsvRows(filename, headers, dataRows){
  const csv = [headers, ...dataRows].map(r=>r.map(csvField).join(',')).join('\r\n');
  const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url), 30000);
}
function downloadCsvTemplate(filename, headers, exampleRow){
  downloadCsvRows(filename, headers, [exampleRow]);
}
const EXPORT_XL_ICON = '<svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" aria-hidden="true"><path d="M12 3v12m0 0l-4-4m4 4l4-4"/><path d="M4 19h16"/></svg>';
/* Lightweight "what you see on screen" export -- unlike exportOtsWorksheet/
   exportOtsExcel above (formal, formula-driven, printable sheets), this is
   for a quick working copy of whatever list/table is currently on screen,
   so it stays plain: bold header row, real numbers (not pre-formatted
   strings) via optional per-column numFmt, no borders/merges/print setup. */
async function exportRowsToExcel(filename, sheetName, headers, rows, numFmts){
  await ensureExcelJS();
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet(sheetName, { views: [{showGridLines:false, state:'frozen', ySplit:1}] });
  ws.columns = headers.map(h => ({ header: h, width: Math.max(14, h.length+6) }));
  ws.getRow(1).font = { bold: true };
  rows.forEach(r => {
    const row = ws.addRow(r);
    if(numFmts) numFmts.forEach((fmt, i) => { if(fmt) row.getCell(i+1).numFmt = fmt; });
  });
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url), 30000);
}
/* Shared bottom-center confirmation toast -- first use anywhere in the app
   (nothing to reuse), so it's kept generic on purpose for whatever the next
   fire-and-forget action needs (export, copy, etc.), not export-specific. */
let __toastTimer = null;
function showToast(msg){
  let el = document.getElementById('appToast');
  if(!el){
    el = document.createElement('div');
    el.id = 'appToast';
    el.className = 'app-toast';
    document.body.appendChild(el);
  }
  el.textContent = msg;
  el.classList.add('show');
  clearTimeout(__toastTimer);
  __toastTimer = setTimeout(()=>el.classList.remove('show'), 2800);
}
function downloadDailyTemplate(){
  const headers = ['Sol','Region','Branch','Account No','Customer ID','Intt Rev','Scheme Code','Account Name','Balance Amount','Turnover','Interest Charge Amount','Continuous Excess Date','Review Date','KCC Disbursement Date/Stock Date','Due date','Demand Amount','Adjustment Amount','Reasons','Exempted','Account NPA Date','Cust NPA Date','SBA Acc/Balance','Remarks','Category','Prov Amt','CADU','Sanction Date','Limit','Disb Date','ROI','Mobile No','SMA Status','Sec Val','Sec OS','Unsec OS'];
  const example = ['9316','HATHRAS','MAANT','160720303013711','705760143','','AG203','EXAMPLE BORROWER NAME','38155.85','','','','','','','38155.85','','CBS NPA','','30-11-2012','30-11-2012','124610100004372 -> 0','Marked in CBS','DA3','38155.85','1009','23-11-2010','40000','23-11-2011','9','9999999999','SMA0','80000','38155.85','0'];
  downloadCsvTemplate('UPGB_Daily_NPA_Template.csv', headers, example);
}
function downloadMasterTemplate(){
  const headers = ['Customer ID (CIF)','Customer Name','Address','Aadhar No','PAN'];
  const example = ['705760143','EXAMPLE BORROWER NAME','VILL EXAMPLE, POST EXAMPLE, DISTRICT, UP - 000000','123456789012','ABCDE1234F'];
  downloadCsvTemplate('UPGB_Customer_Master_Template.csv', headers, example);
}
function downloadBranchAdvTemplate(){
  const headers = ['Sol ID','Branch Name','Advance (₹ Lakhs)','NPA MARCH 26 (₹ Lakhs)','NPA JUNE 26(₹ Lakhs)'];
  const example = ['9282','M.G.Hathras','1877.53','71.53','75.45'];
  downloadCsvTemplate('UPGB_Branch_Advance_Template.csv', headers, example);
}
function downloadInterestReversalMasterTemplate(){
  const headers = ['Account_Number','Interest Reversal'];
  const example = ['151635110000123','5525'];
  downloadCsvTemplate('UPGB_Interest_Reversal_Template.csv', headers, example);
}
// Two example rows on purpose -- one Account Number-keyed, one Customer
// ID-keyed -- so the template itself shows that either column works
// (together or separately) rather than only documenting it in prose Alok
// might not read. A row needs whichever one it has, plus Address.
function downloadAddressListTemplate(){
  const headers = ['Account Number','Customer ID','Address'];
  const rows = [
    ['151635110000123','','P.O. AGSAULI HATHRAS'],
    ['','4521178','MAIN ROAD, HATHRAS'],
  ];
  downloadCsvRows('UPGB_Address_List_Template.csv', headers, rows);
}
/* Branch Contacts template -- unlike the other "blank + one example row"
   templates above, this one pre-fills Sol ID/Old Sol ID/Branch Name for
   every branch from BRANCH_LIST (the app's own reference list), and
   carries forward whatever's already in DATA.branchContacts, so
   re-downloading after a partial upload doesn't lose what's already
   collected -- only the still-blank fields need filling in. */
function downloadBranchContactsTemplate(){
  const headers = ['Sol ID','Old Sol ID','Branch Name','Branch Manager Name','Manager Mobile No.','Manager Email ID','Recovery Officer Name','Recovery Officer Mobile No.','Branch Landline No.','Branch Category','Branch Address','IFSC Code','Remarks'];
  const rows = BRANCH_LIST.map(([oldId,newId,name])=>{
    const bc = DATA.branchContacts[String(newId)] || {};
    return [newId, oldId, name, bc.mgr||'', bc.mgrMobile||'', bc.mgrEmail||'', bc.roName||'', bc.roMobile||'', bc.landline||'', bc.category||'', bc.address||'', bc.ifsc||'', bc.remarks||''];
  });
  downloadCsvRows('UPGB_Branch_Contacts_Template.csv', headers, rows);
}
function downloadLokAdalatTemplate(){
  const headers = ['Date','Account No.','OTS Amount','Token Amount','Remark'];
  const example = ['01-09-2026','150281010000028','','20000','Token received in cash'];
  downloadCsvTemplate('UPGB_Lok_Adalat_Template.csv', headers, example);
}

function downloadUpdatedApp(){
  const json = JSON.stringify({ npa: DATA.npa, oldots: DATA.oldots, asOnDate: DATA.asOnDate||null });
  const blob = new Blob([json], {type:'application/json'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  const dateTag = (DATA.asOnDate||'').replace(/[^0-9]/g,'') || 'backup';
  a.href = url;
  a.download = `UPGB_NPA_data_backup_${dateTag}.json`;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url), 60000);
}

/* ---------- Publish to live site (commits data/latest.json straight to
   this repo via GitHub's Git Data API, using the Admin's own already-repo-
   scoped OAuth token -- see js/publish.js). Only the final ref-update step
   changes what's live; anything that fails before that leaves production
   untouched. ---------- */
let __pendingPublish = null; // { type: 'publish'|'rollback', dataObj, meta, versionId } staged for confirmPublish()
let __lastHistoryList = []; // last-loaded version history, so rollback review can show metadata without a separate fetch

/* #publishStatus (the "Published — live at ..." banner) sits as a sibling
   AFTER #publishReviewPanel in the DOM, not inside it -- closePublishReview()
   only hides the review panel, so a successful publish's confirmation text
   stays on screen indefinitely afterward. If new data (KCC Overdue, Daily
   PNPA, Branch Advance/Contacts, or a fresh daily NPA Apply
   Update) gets staged after that, the old "Published" banner is still
   sitting right there looking current -- a real report from Alok: he
   uploaded a KCC Overdue rollover file, saw the still-visible banner from
   an earlier publish, and reasonably read that as confirmation his new
   upload had gone live, when the actual commit never touched
   data/kcc-overdue.json at all. Every place that stages new pending data
   for publish calls this first, so a stale success (or failure) message
   can never be mistaken for feedback on what's about to be published. */
function clearStalePublishStatus(){
  const el = document.getElementById('publishStatus');
  if(el) el.innerHTML = '';
  updateUnpublishedBanner();
}
// A persistent, hard-to-miss reminder that something is staged only in
// this browser tab's memory and hasn't actually gone live yet -- the inline
// "✔ ... parsed" message each upload shows scrolls out of view the moment
// another panel is touched, and the beforeunload confirm() dialog below is
// easy to dismiss without reading. This banner stays visible right above
// the Publish button itself for as long as anything is unpublished, reusing
// the exact same pendingUnpublishedLabel() list the beforeunload guard
// already computes, so the two can never disagree about what's pending.
function updateUnpublishedBanner(){
  const el = document.getElementById('unpublishedRefDataBanner');
  if(!el) return;
  const parts = pendingUnpublishedLabel();
  if(parts.length){
    el.style.display = 'block';
    el.textContent = '⚠ Not live yet — ' + parts.join('; ') + '. Click "Publish to Live Site" below to make it available everywhere.';
  } else {
    el.style.display = 'none';
  }
}

/* Real bug this guards against: __pendingPnpaData/
   __pendingKccOverdueData/__pendingData are plain JS variables -- a file
   upload stages data ONLY in browser memory until Publish actually sends
   it. A full page reload (the Refresh button, browser F5, closing the
   tab) wipes all of it silently, with no error, no warning. Alok hit this
   exactly: uploaded a KCC Overdue file, then (most likely) hit Refresh
   before Publishing -- the reload wiped the staged upload, and the
   Publish that followed went through "successfully" but with nothing
   KCC-related to include, since Refresh giving no indication that the
   upload it just discarded even existed. Confirmed via git history: that
   publish's commit only touched data/history/*, never data/kcc-overdue.json,
   matching this exact failure mode.
   window.confirm() lets the Refresh button show wording that actually
   names what's about to be lost; beforeunload is the safety net for every
   OTHER way the page can go away (F5, closing the tab, navigating off)
   where the browser only allows its own generic "leave site?" prompt, not
   custom text -- both check the same one source of truth below. */
function pendingUnpublishedLabel(){
  const parts = [];
  if(__pendingData) parts.push('the uploaded daily NPA file (not yet applied)');
  if(typeof __pendingPnpaData!=='undefined' && __pendingPnpaData) parts.push('the Daily PNPA upload');
  if(typeof __pendingKccOverdueData!=='undefined' && __pendingKccOverdueData) parts.push('the KCC Overdue upload');
  if(__hasUnpublishedRefData) parts.push('a reference-data update (Address List, Customer Master, Branch Advance/Contacts, Interest Reversal, Lok Adalat, or a Special Note)');
  return parts;
}
window.addEventListener('beforeunload', (e) => {
  if(pendingUnpublishedLabel().length){ e.preventDefault(); e.returnValue = ''; }
});

function computeCurrentDataSummary(){
  return { rowCount: DATA.npa.rows.length, asOnDate: DATA.asOnDate||null };
}
/* One row per dataset that could go into this publish's commit -- each
   dataset gets its own icon/name/detail instead of every publish being
   labelled generically "Publish NPA data" regardless of what actually
   changed (Alok's report: the log said "NPA data" even for a KCC-only
   upload). The NPA Book row is marked "checked automatically" rather than
   a plain "included" badge, since whether it truly changed since the last
   publish can only be known server-side (publishData() compares blob
   shas) -- it always ships as part of data/latest.json, but only earns a
   new version-history entry and a mention in the commit message when its
   content actually differs from what's already live. */
function publishReviewItemRow({icon, title, sub, maybe}){
  return `<div class="publish-item${maybe?' is-maybe':''}">
    <div class="publish-item-icon">${svgIcon(icon)}</div>
    <div class="publish-item-body">
      <div class="publish-item-title">${esc(title)}</div>
      <div class="publish-item-sub">${esc(sub)}</div>
    </div>
    <span class="publish-item-badge${maybe?'':' on'}">${maybe?'Checked automatically':'Included'}</span>
  </div>`;
}
function openPublishReview(){
  const summary = computeCurrentDataSummary();
  const meta = __lastApplyMeta || {};
  const user = (window.UPGBAuth && window.UPGBAuth.getCurrentUser()) || {};
  const staleLine = meta.staleRemovedCount>0
    ? `<div class="pr-warn">${meta.staleRemovedCount.toLocaleString('en-IN')} account(s) removed as regularized/closed.</div>`
    : '';
  const addedLine = meta.newAddedCount>0
    ? `<div class="pr-good">${meta.newAddedCount.toLocaleString('en-IN')} new account(s) added.</div>`
    : '';

  const npaLabel = `NPA data (${summary.rowCount.toLocaleString('en-IN')} accounts, as on ${fmtAsOnDisplay()})`;
  const items = [publishReviewItemRow({
    icon: ICON_BANKNOTE, title: 'NPA Book', maybe: true,
    sub: `${summary.rowCount.toLocaleString('en-IN')} accounts · as on ${fmtAsOnDisplay()}`,
  })];
  let pnpaLabel = null, kccovLabel = null;
  if(__pendingPnpaData){
    pnpaLabel = `Daily PNPA (${__pendingPnpaData.rows.length.toLocaleString('en-IN')} accounts, as on ${__pendingPnpaData.asOnDate||''})`;
    items.push(publishReviewItemRow({ icon: ICON_ALERT_CIRCLE, title: 'Daily PNPA', sub: `${__pendingPnpaData.rows.length.toLocaleString('en-IN')} accounts · as on ${esc(__pendingPnpaData.asOnDate||'')}` }));
  }
  if(__pendingKccOverdueData){
    kccovLabel = `KCC Overdue (${__pendingKccOverdueData.rows.length.toLocaleString('en-IN')} accounts, as on ${__pendingKccOverdueData.asOnDate||''})`;
    items.push(publishReviewItemRow({ icon: ICON_TARGET, title: 'KCC Overdue', sub: `${__pendingKccOverdueData.rows.length.toLocaleString('en-IN')} accounts · as on ${esc(__pendingKccOverdueData.asOnDate||'')}` }));
  }
  const specialNoteCount = Object.keys(DATA.specialNotes||{}).length;
  if(specialNoteCount){
    items.push(publishReviewItemRow({ icon: ICON_NOTE, title: 'Special Notes', maybe: true, sub: `${specialNoteCount.toLocaleString('en-IN')} account(s) with a note` }));
  }
  const lokAdalatCount = Object.keys(DATA.lokAdalat||{}).length;
  if(lokAdalatCount){
    items.push(publishReviewItemRow({ icon: ICON_NOTE, title: 'Lok Adalat (Token Received)', maybe: true, sub: `${lokAdalatCount.toLocaleString('en-IN')} account(s)` }));
  }
  document.getElementById('publishReviewSummary').innerHTML = `
    ${addedLine}
    ${staleLine}
    <div class="publish-item-list">${items.join('')}</div>
    <div style="color:var(--sub)">Publishing as <b>${esc(user.login||'unknown')}</b>. Goes live on npadashboard.alokmittal.net within about a minute.</div>
  `;
  __pendingPublish = {
    type: 'publish',
    dataObj: { npa: DATA.npa, oldots: DATA.oldots, asOnDate: DATA.asOnDate||null, branchAdvances: DATA.branchAdvances||{}, branchContacts: DATA.branchContacts||{}, specialNotes: DATA.specialNotes||{}, lokAdalat: DATA.lokAdalat||{}, interestReversalMaster: DATA.interestReversalMaster||{}, addressList: DATA.addressList||{}, addressListByCustomer: DATA.addressListByCustomer||{} },
    meta: {
      asOnDate: summary.asOnDate,
      rowCount: summary.rowCount,
      npaLabel,
      publishedBy: user.login || null,
      isRollback: false,
    },
    labels: { pnpaLabel, kccovLabel },
  };
  document.getElementById('publishConfirmBtn').textContent = 'Confirm & Publish';
  document.getElementById('publishReviewPanel').style.display = 'block';
  document.getElementById('publishStatus').innerHTML = '';
}
function closePublishReview(){
  const panel = document.getElementById('publishReviewPanel');
  if(panel) panel.style.display = 'none';
  __pendingPublish = null;
}
async function confirmPublish(){
  if(!__pendingPublish || !window.UPGBPublish) return;
  const confirmBtn = document.getElementById('publishConfirmBtn');
  const cancelBtn = document.getElementById('publishCancelBtn');
  const statusEl = document.getElementById('publishStatus');
  confirmBtn.disabled = true; cancelBtn.disabled = true;
  confirmBtn.classList.add('is-loading');
  const onProgress = (msg) => { statusEl.innerHTML = `<div class="upload-status ok">⏳ ${esc(msg)}</div>`; };
  try{
    const labels = __pendingPublish.labels || {};
    let extraFiles;
    if(__pendingPublish.type!=='rollback' && __pendingPnpaData){
      extraFiles = (extraFiles||[]).concat([{ path:'data/pnpa.json', content: __pendingPnpaData, label: labels.pnpaLabel }]);
    }
    if(__pendingPublish.type!=='rollback' && __pendingKccOverdueData){
      extraFiles = (extraFiles||[]).concat([{ path:'data/kcc-overdue.json', content: __pendingKccOverdueData, label: labels.kccovLabel }]);
    }
    const result = __pendingPublish.type === 'rollback'
      ? await window.UPGBPublish.rollbackToVersion(__pendingPublish.versionId, onProgress)
      : await window.UPGBPublish.publishData(__pendingPublish.dataObj, __pendingPublish.meta, onProgress, extraFiles);
    statusEl.innerHTML = `<div class="upload-status ok">✔ ${esc(result.commitMessage||'Published')} — live at npadashboard.alokmittal.net within ~30-60s (commit ${esc(result.commitSha.slice(0,7))}).</div>`;
    document.getElementById('publishBtn').disabled = true;
    __pendingPnpaData = null;
    __pendingKccOverdueData = null;
    __hasUnpublishedRefData = false;
    updateUnpublishedBanner();
    closePublishReview();
    loadVersionHistory();
  } catch(err){
    statusEl.innerHTML = `<div class="upload-status err">⚠ Publish failed: ${esc(err.message||err)}. Nothing changed on the live site — safe to retry.</div>`;
  } finally {
    confirmBtn.disabled = false; cancelBtn.disabled = false;
    confirmBtn.classList.remove('is-loading');
  }
}
async function loadVersionHistory(){
  const listEl = document.getElementById('versionHistoryList');
  const countEl = document.getElementById('versionHistoryCount');
  if(!listEl || !window.UPGBPublish) return;
  listEl.innerHTML = '<div style="padding:8px 0;color:var(--sub);font-size:11.5px">Loading…</div>';
  try{
    const history = await window.UPGBPublish.getHistoryIndex();
    __lastHistoryList = history;
    if(countEl) countEl.textContent = history.length ? `(${history.length})` : '';
    if(!history.length){ listEl.innerHTML = '<div style="padding:8px 0;color:var(--sub);font-size:11.5px">No published versions yet.</div>'; return; }
    listEl.innerHTML = history.map((v,i)=>`
      <div class="version-row${i===0?' current':''}">
        <div>
          <span class="vr-meta">${esc(v.date||'Unknown date')} — ${(v.rowCount||0).toLocaleString('en-IN')} accounts</span>
          <span class="vr-sub">${v.isRollback?'rollback · ':''}published ${v.publishedAt?fmtDateTime(new Date(v.publishedAt)):''}${v.publishedBy?' by '+esc(v.publishedBy):''}</span>
        </div>
        ${i===0?'':`<button type="button" onclick="openRollbackReview('${esc(v.file)}')">Rollback to this</button>`}
      </div>
    `).join('');
  } catch(err){
    listEl.innerHTML = `<div style="padding:8px 0;color:var(--red);font-size:11.5px">Could not load version history: ${esc(err.message||err)}</div>`;
  }
}
function openRollbackReview(fileName){
  const version = __lastHistoryList.find(v=>v.file===fileName);
  if(!version) return;
  document.getElementById('publishReviewSummary').innerHTML = `
    <div class="pr-warn">You are about to roll back the LIVE site to an older version.</div>
    <div>Version date: <b>${esc(version.date||'unknown')}</b></div>
    <div>Accounts in this version: <b>${(version.rowCount||0).toLocaleString('en-IN')}</b></div>
    <div style="margin-top:8px;color:var(--sub)">This publishes the old version again as the new current version — nothing in your current session's applied data is used.</div>
  `;
  __pendingPublish = { type: 'rollback', versionId: fileName };
  document.getElementById('publishConfirmBtn').textContent = 'Confirm Rollback';
  document.getElementById('publishReviewPanel').style.display = 'block';
  document.getElementById('publishStatus').innerHTML = '';
}

/* ---------- Cmd+K quick search palette ---------- */
const cmdkOverlay=document.getElementById('cmdkOverlay'), cmdkInput=document.getElementById('cmdkInput'), cmdkResults=document.getElementById('cmdkResults'), cmdkClose=document.getElementById('cmdkClose');
let cmdkMatches=[], cmdkActive=0;
/* Quick Search only ever looked at DATA.npa.rows -- confirmed by direct
   data check that KCC Overdue's ~9.7k accounts have zero overlap with that
   dataset's account numbers (completely separate report universes, not
   just a filtered subset), so a borrower who only appears in KCC Overdue
   could never be found here no matter what was typed. Prefetch it in the
   background the first time the palette opens (if not already loaded from
   visiting that tab) so search works regardless of which tabs have been
   visited this session. Recovery Dashboard (branch portal) note: the PNPA
   prefetch that used to sit alongside this was removed -- PNPA isn't part
   of this portal's surface at all (no nav item, no view section), so
   fetching it here would just be a wasted network call on every Cmd+K
   open for data nothing ever reads. */
function openCmdk(){
  if(!cmdkOverlay) return;
  cmdkOverlay.classList.add('show'); cmdkInput.value=''; renderCmdk(''); setTimeout(()=>cmdkInput.focus(),30);
  if(!KCC_OVERDUE_DATA){
    fetchJson(DATA_ORIGIN + 'data/kcc-overdue.json?t=' + Date.now())
      .then(d=>{ KCC_OVERDUE_DATA=d; if(cmdkOverlay.classList.contains('show')) renderCmdk(cmdkInput.value); })
      .catch(()=>{});
  }
}
window.openCmdk = openCmdk;
function closeCmdk(){ if(cmdkOverlay) cmdkOverlay.classList.remove('show'); }
function cmdkItemHtml(m, idx){
  const r = m.row;
  if(m.source==='npa'){
    const asset=r[C.ASSET]||''; const initials=(String(r[C.NAME]||'?').trim().charAt(0)||'?').toUpperCase();
    return `<div class="cmdk-item${idx===0?' active':''}" data-idx="${idx}">
      <div class="ci-ic">${esc(initials)}</div>
      <div class="ci-main"><div class="ci-name">${esc(r[C.NAME])||'—'}</div>
        <div class="ci-sub">A/c ${esc(r[C.ACCT_NO])} · ${esc(r[C.SOL_DESC])||''} · Cust ${esc(r[C.CUST_ID])}</div></div>
      ${asset?`<span class="badge-pill ci-badge ${esc(asset)}">${esc(asset)}</span>`:''}
    </div>`;
  }
  const isKcc = m.source==='kccov';
  const name = isKcc ? r[KC.NAME] : r[PC.NAME];
  const acct = isKcc ? r[KC.ACCT] : r[PC.ACCT];
  const branch = isKcc ? r[KC.BRANCH] : r[PC.BRANCH];
  const initials=(String(name||'?').trim().charAt(0)||'?').toUpperCase();
  return `<div class="cmdk-item${idx===0?' active':''}" data-idx="${idx}">
    <div class="ci-ic">${esc(initials)}</div>
    <div class="ci-main"><div class="ci-name">${esc(name)||'—'}</div>
      <div class="ci-sub">A/c ${esc(acct)} · ${esc(branch)||''}</div></div>
    <span class="badge-pill ci-badge" style="background:${isKcc?'var(--accent-soft);color:var(--accent)':'var(--amber-soft);color:var(--amber)'}">${isKcc?'KCC Overdue':'PNPA'}</span>
  </div>`;
}
function renderCmdk(q){
  q=String(q||'').trim().toLowerCase();
  const out=[]; const seen=new Set();
  // Recovery Dashboard (branch portal): scoped to the logged-in Sol ID --
  // same "simple login + auto-filter" model as Dashboard/KCC Overdue
  // above, not real access control. No PNPA loop here (unlike the
  // production site's own renderCmdk) -- PNPA isn't part of this portal's
  // surface at all (no nav item, no view section), so searching it would
  // just be dead code reachable from nowhere else.
  const mySolId = loggedInSolId();
  if(q){
    for(const r of DATA.npa.rows){
      if(mySolId && String(r[C.SOL_ID])!==String(mySolId)) continue;
      const name=String(r[C.NAME]||'').toLowerCase(), acct=String(r[C.ACCT_NO]||'').toLowerCase(),
        cust=String(r[C.CUST_ID]||'').toLowerCase();
      if(name.includes(q)||acct.includes(q)||cust.includes(q)){
        const cid=String(r[C.CUST_ID]); if(seen.has(cid)) continue; seen.add(cid);
        out.push({source:'npa', row:r});
        if(out.length>=12) break;
      }
    }
    if(out.length<15 && KCC_OVERDUE_DATA && KCC_OVERDUE_DATA.rows){
      for(const r of KCC_OVERDUE_DATA.rows){
        if(mySolId && String(KCCOV_BRANCH_SOL[String(r[KC.BRANCH]).toUpperCase()])!==String(mySolId)) continue;
        const name=String(r[KC.NAME]||'').toLowerCase(), acct=String(r[KC.ACCT]||'').toLowerCase();
        if(name.includes(q)||acct.includes(q)){
          out.push({source:'kccov', row:r});
          if(out.length>=15) break;
        }
      }
    }
  }
  cmdkMatches=out; cmdkActive=0;
  if(!q){ cmdkResults.innerHTML='<div class="cmdk-empty">Type a name, account no. or customer ID…</div>'; return; }
  if(!out.length){ cmdkResults.innerHTML='<div class="cmdk-empty">No borrower found for that.</div>'; return; }
  /* out is already contiguous by source (the three loops above push npa,
     then kccov, then pnpa matches in that order) -- grouping here only
     inserts a label whenever the source changes, it never reorders
     anything, so cmdkMatches/idx stays exactly aligned with what's on
     screen and arrow-key navigation keeps working across group labels. */
  let lastSource=null;
  const parts=[];
  out.forEach((m,idx)=>{
    if(m.source!==lastSource){ parts.push(cmdkGroupLabelHtml(m.source)); lastSource=m.source; }
    parts.push(cmdkItemHtml(m,idx));
  });
  cmdkResults.innerHTML=parts.join('');
  cmdkResults.querySelectorAll('.cmdk-item').forEach(it=>{
    it.addEventListener('click',()=>pickCmdk(+it.dataset.idx));
    it.addEventListener('mousemove',()=>setCmdkActive(+it.dataset.idx));
  });
}
function cmdkGroupLabelHtml(source){
  const map = {
    npa: ['NPA Accounts', 'var(--accent)'],
    kccov: ['KCC Overdue', 'var(--accent)'],
    pnpa: ['Daily PNPA', 'var(--amber)'],
  };
  const [label, dotColor] = map[source] || [source, 'var(--accent)'];
  return `<div class="cmdk-group-label"><span class="cmdk-group-dot" style="background:${dotColor}"></span>${esc(label)}</div>`;
}
function setCmdkActive(idx){ cmdkActive=idx; cmdkResults.querySelectorAll('.cmdk-item').forEach(it=>it.classList.toggle('active',+it.dataset.idx===idx)); }
/* NPA results link to the real OTS settlement detail (openDetail); KCC
   Overdue/PNPA rows have no customer ID or the fuller record that detail
   view needs (confirmed separate datasets, not just a filtered view of
   the same one), so they open a small read-only info card instead. */
function pickCmdk(idx){
  const m=cmdkMatches[idx]; if(!m) return;
  closeCmdk();
  showQuickAcctDetail(m.source, m.row);
}
/* Recovery Dashboard (branch portal) has no OTS Calculator at all -- an
   'npa' source used to route to openDetail() (the settlement screen) on
   the production site; here it renders the same read-only field-grid card
   as KCC Overdue/PNPA instead, via a new isNpa branch below. openDetail()
   itself (and the whole settlement-calculation code block it depends on)
   is left in this file, dormant/unreachable, rather than deleted -- safer
   than a large deletion pass on a copy of a production-adjacent codebase;
   nothing in this portal's UI calls it. */
function showQuickAcctDetail(source, row){
  const isKcc = source==='kccov';
  const isNpa = source==='npa';
  const title = isNpa ? row[C.NAME] : (isKcc ? row[KC.NAME] : row[PC.NAME]);
  const branch = isNpa ? row[C.SOL_DESC] : (isKcc ? row[KC.BRANCH] : row[PC.BRANCH]);
  const sourceLabel = isNpa ? 'NPA' : (isKcc ? 'KCC Overdue' : 'Daily PNPA');
  const sub = `${esc(branch)||'—'} · ${sourceLabel}`;
  const fields = isNpa ? [
    ['Account No', row[C.ACCT_NO]], ['Customer ID', row[C.CUST_ID]], ['Scheme', row[C.SCHEME]],
    ['Outstanding', fmtINR2(row[C.OUTBAL])], ['Asset Class', row[C.SYS_SUBCLASS]||row[C.ASSET]],
    ['Sanction Date', fmtDate(toDate(row[C.SANCT_DT]))], ['Sanction Limit', fmtINR2(row[C.SANCT_LIM])],
    ['NPA Date', fmtDate(toDate(row[C.NPA_DT]))],
  ] : isKcc ? [
    ['Account No', row[KC.ACCT]], ['Scheme', row[KC.SCHEME]], ['Outstanding', fmtINR2(row[KC.OS])],
    ['CADU', fmtINR2(row[KC.CADU])], ['Limit', fmtINR2(row[KC.LIMIT])], ['Cust NPA Date', row[KC.CUSTNPADATE]],
    ['F.Y.', row[KC.FY]], ['Category', row[KC.CATEGORY]], ['SMA', row[KC.SMA]], ['Reason', row[KC.REASON]],
  ] : [
    ['Account No', row[PC.ACCT]], ['Scheme', row[PC.SCHEME]], ['Outstanding', fmtINR2(row[PC.OS])],
    ['CADU', fmtINR2(row[PC.CADU])], ['Limit', fmtINR2(row[PC.LIMIT])], ['Review Date', row[PC.REVIEW]],
    ['Reason', row[PC.REASON]],
  ];
  document.getElementById('quickAcctTitle').textContent = title || '—';
  document.getElementById('quickAcctSub').innerHTML = sub;
  // Account No (15 digits) and Reason (can be several comma-joined codes)
  // are the two fields most likely to be longer than a narrow grid column
  // on a phone -- give them their own full-width row instead of letting
  // them wrap mid-digit/mid-word inside a half-width cell (Alok's own
  // screenshot, 2026-09-25, showed exactly this: "1501351100 02325" split
  // across two lines).
  const FULL_WIDTH_KEYS = ['Account No', 'Reason'];
  document.getElementById('quickAcctGrid').innerHTML = fields.map(([k,v])=>
    `<div${FULL_WIDTH_KEYS.includes(k)?' class="full"':''}><div class="k">${esc(k)}</div><div class="v">${esc(v!==null&&v!==undefined&&v!==''?v:'—')}</div></div>`
  ).join('');
  document.getElementById('quickAcctModalOverlay').classList.add('show');
}
window.showQuickAcctDetail = showQuickAcctDetail;
/* Tapping a row inside the NPA/PNPA/KCC Overdue account-list modal opens
   the same Quick Account Detail card as a search result -- looked up by
   account no. against the raw dataset rather than threading the raw row
   through the list-modal's already-transformed {acctNo,name,os,...}
   display objects, since account numbers are unique within each report. */
function showQuickAcctDetailByAcct(source, acctNo){
  if(source==='npa'){
    const row = npaByAcct.get(String(acctNo)); // reuses the existing "Build indexes once" Map
    if(row) showQuickAcctDetail('npa', row);
    return;
  }
  const data = source==='kccov' ? KCC_OVERDUE_DATA : PNPA_DATA;
  if(!data || !data.rows) return;
  const col = source==='kccov' ? KC.ACCT : PC.ACCT;
  const row = data.rows.find(r=>String(r[col])===String(acctNo));
  if(row) showQuickAcctDetail(source, row);
}
window.showQuickAcctDetailByAcct = showQuickAcctDetailByAcct;
/* Repoints every former openDetail(custId) call site (Dashboard's
   account/customer list-modal rows, Cmd+K's npa branch) at the read-only
   card instead. Reuses the same byCustId Map (js/app.js, "Build indexes
   once") openDetail() itself already looks up by -- no new indexing. */
function showNpaQuickDetail(custId){
  const row = byCustId.get(String(custId));
  if(row) showQuickAcctDetail('npa', row);
}
window.showNpaQuickDetail = showNpaQuickDetail;
function cmdkEnsureVisible(){ const el=cmdkResults.querySelector('.cmdk-item.active'); if(el) el.scrollIntoView({block:'nearest'}); }
if(cmdkOverlay){
  cmdkInput.addEventListener('input',()=>renderCmdk(cmdkInput.value));
  cmdkClose.addEventListener('click',closeCmdk);
  cmdkOverlay.addEventListener('click',(e)=>{ if(e.target===cmdkOverlay) closeCmdk(); });
  cmdkInput.addEventListener('keydown',(e)=>{
    if(e.key==='ArrowDown'){ e.preventDefault(); setCmdkActive(Math.min(cmdkActive+1,cmdkMatches.length-1)); cmdkEnsureVisible(); }
    else if(e.key==='ArrowUp'){ e.preventDefault(); setCmdkActive(Math.max(cmdkActive-1,0)); cmdkEnsureVisible(); }
    else if(e.key==='Enter'){ e.preventDefault(); pickCmdk(cmdkActive); }
    else if(e.key==='Escape'){ closeCmdk(); }
  });
}
document.addEventListener('keydown',(e)=>{
  if((e.metaKey||e.ctrlKey) && (e.key==='k'||e.key==='K')){ e.preventDefault(); (cmdkOverlay&&cmdkOverlay.classList.contains('show'))?closeCmdk():openCmdk(); }
  else if(e.key==='Escape'){
    if(cmdkOverlay && cmdkOverlay.classList.contains('show')) closeCmdk();
    else if(document.getElementById('wsModalOverlay')?.classList.contains('show')) closeOtsWorksheet();
    else if(document.getElementById('detailPane').classList.contains('open')) closeDetail();
  }
});

/* ==================================================================
   Dashboard — Portfolio Intelligence (additive; reads DATA, never
   mutates it; every figure below is derived with the exact same
   PROV_RATES / netOutstanding / totalPL formulas used in the
   per-borrower settlement engine above — nothing recomputed differently) */
const ASSET_ORDER = ['SUB_STD','DA1','DA2','DA3','LOSS'];
const ASSET_SEV_COLOR = { SUB_STD:'var(--sev-1)', DA1:'var(--sev-2)', DA2:'var(--sev-3)', DA3:'var(--sev-4)', LOSS:'var(--sev-5)' };
const SLAB_DEFS = [
  {id:'s1', label:'Upto ₹2 Lakh', max:200000},
  {id:'s2', label:'₹2 Lakh – ₹5 Lakh', max:500000},
  {id:'s3', label:'₹5 Lakh – ₹10 Lakh', max:1000000},
  {id:'s4', label:'₹10 Lakh & above', max:Infinity},
];
const HIGH_VALUE_CUST_THRESHOLD = 1000000; // ₹10 Lakh

function computeDashboardStats(branchFilter){
  const rows = DATA.npa.rows;
  const today = new Date();
  const assetMix = {};
  const branchMap = new Map();
  const allBranches = new Set();
  const buckets = [
    {id:'ne', label:'Not yet eligible (≤ 6 months)', count:0, os:0},
    {id:'y1', label:'6 months – 1 year', count:0, os:0},
    {id:'y13', label:'1 – 3 years', count:0, os:0},
    {id:'y3p', label:'3+ years', count:0, os:0},
  ];
  const slabs = SLAB_DEFS.map(sl=>({...sl, count:0, os:0}));
  const schemeMix = { KCC:{count:0,os:0}, NONKCC:{count:0,os:0} };
  const custMap = new Map();
  const acctList = [];
  let totalOS=0, totalNetOS=0, totalProvision=0, totalBookValue=0;
  let eligibleCount=0, notEligibleCount=0, matchedAccounts=0;
  const seen = new Set();
  for(const r of rows){
    const acct = String(r[C.ACCT_NO]);
    const branch = r[C.SOL_DESC] || 'Unassigned';
    if(branch) allBranches.add(branch);
    if(acct==='' || seen.has(acct)) continue;
    seen.add(acct);
    if(branchFilter && branch!==branchFilter) continue;
    matchedAccounts++;
    const asset = r[C.ASSET]||'(unclassified)';
    const os = typeof r[C.OUTBAL]==='number' ? r[C.OUTBAL] : 0;
    const uri = typeof r[C.URI]==='number' ? r[C.URI] : 0;
    const netOs = os-uri;
    const rate = PROV_RATES[asset];
    const provision = rate!==undefined ? netOs*rate : 0;
    const bookValue = Math.max(0, os-uri-provision);

    totalOS+=os; totalNetOS+=netOs; totalProvision+=provision; totalBookValue+=bookValue;

    if(!assetMix[asset]) assetMix[asset]={count:0,os:0};
    assetMix[asset].count++; assetMix[asset].os+=os;

    if(!branchMap.has(branch)) branchMap.set(branch,{count:0,os:0,solId:String(r[C.SOL_ID]||'')});
    const b=branchMap.get(branch); b.count++; b.os+=os;

    const scheme = r[C.SCHEME]||'';
    const schemeKey = scheme==='CC004' ? 'KCC' : 'NONKCC';
    schemeMix[schemeKey].count++; schemeMix[schemeKey].os+=os;

    const slab = slabs.find(sl=>os<=sl.max);
    if(slab){ slab.count++; slab.os+=os; }

    const custId = String(r[C.CUST_ID]||'');
    if(custId){
      if(!custMap.has(custId)) custMap.set(custId, {custId, name:r[C.NAME]||'', branch, address:r[C.ADDR]||'', os:0, count:0});
      const cu = custMap.get(custId); cu.os+=os; cu.count++;
    }

    const npaDate = toDate(r[C.NPA_DT]);
    let bucketId = null;
    if(npaDate){
      const days = daysBetween(today, npaDate);
      if(days<=180){ buckets[0].count++; buckets[0].os+=os; notEligibleCount++; bucketId='ne'; }
      else {
        eligibleCount++;
        if(days<=365){ buckets[1].count++; buckets[1].os+=os; bucketId='y1'; }
        else if(days<=1095){ buckets[2].count++; buckets[2].os+=os; bucketId='y13'; }
        else { buckets[3].count++; buckets[3].os+=os; bucketId='y3p'; }
      }
    }

    acctList.push({ acctNo:acct, custId, name:r[C.NAME]||'', branch, address:r[C.ADDR]||'', os, asset, scheme:schemeKey, slabId: slab?slab.id:null, bucketId, npaDate: npaDate ? fmtDate(npaDate) : '' });
  }
  let oldOtsSum=0, oldOtsCount=0;
  DATA.oldots.rows.forEach(r=>{
    if(r[0]==='') return;
    oldOtsCount++;
    const n = parseFloat(String(r[2]||'').replace(/[^0-9.\-]/g,''));
    if(!isNaN(n)) oldOtsSum+=n;
  });

  const custList = [...custMap.values()];
  const highValueCust = custList.filter(c=>c.os>=HIGH_VALUE_CUST_THRESHOLD);
  const highValueOS = highValueCust.reduce((a,c)=>a+c.os,0);
  const highValueCustList = [...highValueCust].sort((a,b)=>b.os-a.os);
  const allAcctSorted = [...acctList].sort((a,b)=>b.os-a.os);

  return {
    totalAccounts:matchedAccounts, totalOS, totalNetOS, totalProvision, totalBookValue,
    eligibleCount, notEligibleCount, assetMix, branchMap, buckets, oldOtsCount, oldOtsSum,
    branchCount: branchMap.size, allBranches: [...allBranches].sort((a,b)=>a.localeCompare(b)),
    schemeMix, slabs, custCount: custList.length,
    highValueCustCount: highValueCust.length, highValueOS, highValueCustList,
    acctList, allAcctSorted,
  };
}

function fmtINR2(n){ if(n===''||n===null||n===undefined||isNaN(n)) return '—'; return '₹'+Number(n).toLocaleString('en-IN',{minimumFractionDigits:2,maximumFractionDigits:2}); }
/* Same figure, but as HTML with a <wbr> after the ₹ symbol and after every
   comma -- used only in the tight aggregate sidebar (#aggBar), where a
   large multi-account total ("₹1,04,50,000.00") genuinely doesn't fit on
   one line at that column width. A single <wbr> after ₹ alone wasn't
   enough -- the remaining "1,04,50,000.00" chunk could still be too wide
   by itself, and with no further break point the browser fell back to
   `word-break:break-word`, splitting mid-digit or mid-decimal
   ("₹1,81,205" / ".58") -- unreadable at a glance, the whole point of a
   summary figure. Indian-format grouping commas are natural digit-group
   boundaries, so a <wbr> after each one means any forced wrap lands
   between whole groups ("1,04," / "50,000.00") rather than through one --
   nothing after the last comma is ever a candidate, so the decimal pair
   always stays attached to its own group. On any line that fits, none of
   these are used. Caller must assign via .innerHTML, not .textContent. */
function fmtINR2Wrap(n){ const s = fmtINR2(n); return s==='—' ? s : s.replace(/([₹,])/g, '$1<wbr>'); }

/* Recovery Dashboard (branch portal) -- resolves the Sol ID that logged
   in (js/login.js, sessionStorage['upgb-sol-id']) to the exact branch-name
   spelling used in the real NPA rows (C.SOL_DESC can differ slightly from
   BRANCH_LIST/BRANCH_META's canonical spelling, per the existing comment
   on dashboardBranchInfoCard below -- so this is resolved from DATA.npa
   itself, not from BRANCH_LIST). Resolved once and cached; '' means no Sol
   is logged in (shouldn't happen -- login.js gates the whole app -- but
   defensive rather than assuming). "Simple login + auto-filter" model
   (Alok's explicit choice): this is a display filter, not access control. */
let __recoveryBranchName = null;
function loggedInBranchName(){
  if(__recoveryBranchName !== null) return __recoveryBranchName;
  const solId = loggedInSolId();
  if(!solId){ __recoveryBranchName = ''; return __recoveryBranchName; }
  const row = DATA.npa.rows.find(r=>String(r[C.SOL_ID])===String(solId));
  __recoveryBranchName = row ? (row[C.SOL_DESC]||'') : '';
  return __recoveryBranchName;
}
function populateBranchFilter(branches){
  const sel = document.getElementById('dashBranchFilter');
  if(!sel) return;
  const current = sel.value;
  sel.innerHTML = `<option value="">Regional Office</option>` + branches.map(b=>`<option value="${esc(b)}">${esc(b)}</option>`).join('');
  sel.value = branches.includes(current) ? current : '';
}
function updateDashTitle(){
  const el = document.getElementById('dashTitle');
  if(!el) return;
  const first = DATA.npa.rows.find(r=>r[C.REGION]);
  el.textContent = first ? `UPGB ${titleCase(String(first[C.REGION]))} region NPA Portfolio` : 'UPGB NPA Portfolio';
}

function svgDonut(segments, size){
  size = size || 130;
  const strokeW = 18;
  const r = size/2 - strokeW/2 - 2;
  const c = 2*Math.PI*r, cx=size/2, cy=size/2;
  const total = segments.reduce((a,s)=>a+s.value,0) || 1;
  let acc = 0;
  const circles = segments.map(s=>{
    const frac = s.value/total;
    const len = Math.max(0, frac*c - (segments.length>1?1.5:0));
    const dash = `${len.toFixed(2)} ${(c-len).toFixed(2)}`;
    const rotate = (acc/total)*360 - 90;
    acc += s.value;
    const pct = Math.round(frac*100);
    return `<circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="${s.color}" stroke-width="${strokeW}" stroke-dasharray="${dash}" stroke-linecap="round" transform="rotate(${rotate} ${cx} ${cy})" data-label="${esc(s.label)}" data-value-label="${esc(s.valueLabel||'')}" data-pct="${pct}%"></circle>`;
  }).join('');
  return `<svg width="${size}" height="${size}" viewBox="0 0 ${size} ${size}" class="donut-svg">
    <circle cx="${cx}" cy="${cy}" r="${r}" fill="none" stroke="var(--track-bg)" stroke-width="${strokeW}"></circle>
    ${circles}
  </svg>`;
}

function donutCard(segments, size, centerValue, centerLabel){
  return `<div class="donut-wrap">
    ${svgDonut(segments, size)}
    <div class="donut-center"><div class="donut-center-value">${esc(centerValue)}</div><div class="donut-center-label">${esc(centerLabel)}</div></div>
  </div>`;
}

function donutLegend(segments){
  return segments.map(s=>`<div class="legend-row${s.onclick?' clickable':''}"${s.onclick?` onclick="${s.onclick}"`:''}><span class="legend-dot" style="background:${s.color}"></span>${esc(s.label)}<span class="legend-val">${s.valueLabel}</span></div>`).join('');
}
/* Donut hover tooltip -- one delegated listener for every donut on the
   page rather than per-chart wiring, since svgDonut()'s output is inserted
   via innerHTML (no live DOM reference to attach to at construction time).
   The tooltip element itself is created lazily on first hover, same
   pattern as showToast()'s #appToast. */
document.addEventListener('mousemove', (e)=>{
  const seg = e.target && e.target.closest && e.target.closest('.donut-svg circle[data-label]');
  let tip = document.getElementById('chartTooltip');
  if(!seg){ if(tip) tip.classList.remove('show'); return; }
  if(!tip){ tip = document.createElement('div'); tip.id='chartTooltip'; tip.className='chart-tooltip'; document.body.appendChild(tip); }
  tip.innerHTML = `<b>${esc(seg.dataset.label)}</b> · ${esc(seg.dataset.pct)}${seg.dataset.valueLabel?`<br>${esc(seg.dataset.valueLabel)}`:''}`;
  tip.style.left = e.clientX + 'px';
  tip.style.top = e.clientY + 'px';
  tip.classList.add('show');
});

function acctRows(list, opts){
  opts = opts || {};
  const offset = opts.offset||0;
  if(!list.length) return emptyStateRowHtml(6, 'No accounts');
  // Recovery Dashboard (branch portal): no OTS Calculator, so this opens
  // the read-only Quick Account Detail card instead of openDetail().
  return list.map((a,i)=>`<tr class="clickable" onclick="showNpaQuickDetail('${esc(a.custId)}')">
    <td>${opts.rank?`<span class="dash-rank">${i+1+offset}</span>`:''}${esc(a.acctNo)}</td>
    <td class="tal">${esc(a.name)||'—'}</td>
    <td class="tal">${esc(a.address)||'—'}</td>
    <td>${a.asset?`<span class="badge-pill ${esc(a.asset)}" title="${esc(assetLabel(a.asset))}">${esc(a.asset)}</span>`:'—'}</td>
    <td>${fmtINR2(a.os)}</td>
    <td class="tal">${esc(a.npaDate)||'—'}</td>
  </tr>`).join('');
}

const ACCT_LIST_BATCH = 300;
function renderAcctListBatch(list, tbody, shownRef){
  if(shownRef.n>=list.length) return;
  const next = list.slice(shownRef.n, shownRef.n+ACCT_LIST_BATCH);
  tbody.insertAdjacentHTML('beforeend', acctRows(next, {rank:true, offset:shownRef.n}));
  shownRef.n += next.length;
}

/* ---------- Shared column-sort helper (dashboard table + list modal) ---------- */
function applySort(list, sort){
  if(!sort || !sort.key) return list;
  const key = sort.key, dir = sort.dir;
  return [...list].sort((a,b)=>{
    let av=a[key], bv=b[key];
    if(typeof av==='string') av=av.toLowerCase();
    if(typeof bv==='string') bv=bv.toLowerCase();
    if(av<bv) return dir==='asc'?-1:1;
    if(av>bv) return dir==='asc'?1:-1;
    return 0;
  });
}
function nextSort(current, key){
  if(current && current.key===key) return {key, dir: current.dir==='asc'?'desc':'asc'};
  return {key, dir:(key==='name'||key==='branch'||key==='acctNo')?'asc':'desc'};
}
function updateSortIcons(theadId, sort){
  const thead = document.getElementById(theadId);
  if(!thead) return;
  thead.querySelectorAll('th[data-key]').forEach(th=>{
    th.classList.remove('sort-asc','sort-desc');
    const active = sort && th.dataset.key===sort.key;
    if(active) th.classList.add(sort.dir==='asc'?'sort-asc':'sort-desc');
    th.setAttribute('aria-sort', active ? (sort.dir==='asc'?'ascending':'descending') : 'none');
  });
}
/* Keyboard support for sortable column headers (Enter/Space triggers the same click handler) */
document.addEventListener('keydown', (e)=>{
  if(e.key!=='Enter' && e.key!==' ') return;
  const th = e.target.closest && e.target.closest('th.sortable');
  if(!th) return;
  e.preventDefault();
  th.click();
});
/* Same treatment for the Branch/Sol ID panel's clickable rows (opens the
   branch contact card). */
document.addEventListener('keydown', (e)=>{
  if(e.key!=='Enter' && e.key!==' ') return;
  const row = e.target.closest && e.target.closest('.edge-row[role="button"]');
  if(!row) return;
  e.preventDefault();
  row.click();
});

/* On mobile, #aggBar is a fixed-position dock pinned to the bottom of the
   screen (see styles.css) sitting on top of whatever loan-table content
   happens to be scrolled underneath it -- so tapping into OTS Amount (or
   any lt-ots-input) right after Interest Reversal, when the row had only
   been scrolled minimally into view, could land the field's screen
   position right behind the dock, silently swallowing every further tap
   there. Re-center any of these inputs on focus so they're never left
   sitting in that dead zone, regardless of what scroll position got them
   into view in the first place. */
document.addEventListener('focusin', (e)=>{
  if(e.target && e.target.classList && e.target.classList.contains('lt-ots-input')){
    e.target.scrollIntoView({block:'center', behavior:'smooth'});
  }
});

/* ---------- Dashboard: "All Accounts" table (sortable, lazy-scrolled) ---------- */
// Recovery Dashboard (branch portal): Alok, 2026-09-25: "har table main
// filter bhi add karo jisse npa date se filter kar saken ya address se
// filter kar saken" -- Address text filter + NPA Date range, matching the
// same filter shape already added to KCC Overdue/PNPA Slippage.
let dashAddressFilter = '';
let dashNpaDateFrom = '';
let dashNpaDateTo = '';
function dashFilterAcctList(list){
  let out = list;
  if(dashAddressFilter){
    const q = dashAddressFilter.trim().toLowerCase();
    out = out.filter(a=>String(a.address||'').toLowerCase().includes(q));
  }
  if(dashNpaDateFrom || dashNpaDateTo){
    const from = dashNpaDateFrom ? new Date(dashNpaDateFrom+'T00:00:00') : null;
    const to = dashNpaDateTo ? new Date(dashNpaDateTo+'T23:59:59') : null;
    out = out.filter(a=>{
      const dt = toDate(a.npaDate);
      if(!dt) return false;
      if(from && dt < from) return false;
      if(to && dt > to) return false;
      return true;
    });
  }
  return out;
}
let acctListState = {list:[], sort:{key:'os',dir:'desc'}};
function renderAcctListTable(resetScroll){
  const tbody = document.getElementById('acctListBody');
  if(!tbody) return;
  const sorted = applySort(acctListState.list, acctListState.sort);
  acctListState.sortedList = sorted;
  updateSortIcons('acctListHead', acctListState.sort);
  tbody.innerHTML = '';
  const shownRef = {n:0};
  acctListState.shownRef = shownRef;
  renderAcctListBatch(sorted, tbody, shownRef);
  if(resetScroll){ const wrap = document.getElementById('acctListWrap'); if(wrap) wrap.scrollTop = 0; }
}
function sortAcctListBy(key){
  acctListState.sort = nextSort(acctListState.sort, key);
  renderAcctListTable(true);
}
window.sortAcctListBy = sortAcctListBy;
function initAcctListScroll(list){
  const wrap = document.getElementById('acctListWrap');
  if(!wrap) return;
  acctListState.list = list;
  acctListState.sort = {key:'os',dir:'desc'};
  renderAcctListTable();
  wrap.onscroll = ()=>{
    if(wrap.scrollTop + wrap.clientHeight > wrap.scrollHeight - 400) renderAcctListBatch(acctListState.sortedList, document.getElementById('acctListBody'), acctListState.shownRef);
  };
}

function custRows(list){
  if(!list.length) return emptyStateRowHtml(4, 'No customers');
  // Recovery Dashboard (branch portal): no OTS Calculator, so this opens
  // the read-only Quick Account Detail card instead of openDetail().
  return list.map(c=>`<tr class="clickable" onclick="showNpaQuickDetail('${esc(c.custId)}')">
    <td class="tal">${esc(c.name)||'—'}<br><span style="color:var(--ink-mute);font-weight:600;font-size:11px">Cust ID ${esc(c.custId)}</span></td>
    <td class="tal">${esc(c.address)||'—'}</td>
    <td>${c.count} A/C</td>
    <td>${fmtINR2(c.os)}</td>
  </tr>`).join('');
}

const ACCT_LIST_HEAD = '<tr>'
  +'<th class="sortable" data-key="acctNo" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'acctNo\')">Account<span class="sort-ic">▾</span></th>'
  +'<th class="tal sortable" data-key="name" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'name\')">Customer<span class="sort-ic">▾</span></th>'
  +'<th class="tal sortable" data-key="address" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'address\')">Address<span class="sort-ic">▾</span></th>'
  +'<th class="sortable" data-key="asset" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'asset\')">Asset<span class="sort-ic">▾</span></th>'
  +'<th class="sortable" data-key="os" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'os\')">Amount<span class="sort-ic">▾</span></th>'
  +'<th class="tal sortable" data-key="npaDate" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'npaDate\')">NPA Date<span class="sort-ic">▾</span></th>'
  +'</tr>';
const CUST_LIST_HEAD = '<tr>'
  +'<th class="tal sortable" data-key="name" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'name\')">Customer<span class="sort-ic">▾</span></th>'
  +'<th class="tal sortable" data-key="address" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'address\')">Address<span class="sort-ic">▾</span></th>'
  +'<th class="sortable" data-key="count" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'count\')">Accounts<span class="sort-ic">▾</span></th>'
  +'<th class="sortable" data-key="os" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'os\')">Amount<span class="sort-ic">▾</span></th>'
  +'</tr>';

/* ---------- Generic list modal (sortable, lazy-scrolled for account lists) ---------- */
let __listModalScrollHandler = null;
let listModalState = {list:[], type:'acct', sort:{key:'os',dir:'desc'}, addressFilter:''};
function renderListModalBody(resetScroll){
  const body = document.getElementById('listModalBody');
  let filtered = listModalState.list;
  if(listModalState.addressFilter){
    const q = listModalState.addressFilter.trim().toLowerCase();
    filtered = filtered.filter(r=>String(r.address||'').toLowerCase().includes(q));
  }
  const sorted = applySort(filtered, listModalState.sort);
  listModalState.sortedList = sorted;
  updateSortIcons('listModalHead', listModalState.sort);
  if(listModalState.type==='cust'){ body.innerHTML = custRows(sorted); }
  else if(listModalState.type==='pnpa'){ body.innerHTML = pnpaAcctRows(sorted); }
  else if(listModalState.type==='kccov'){ body.innerHTML = kccovAcctRows(sorted); }
  else{
    body.innerHTML = '';
    const shownRef = {n:0};
    listModalState.shownRef = shownRef;
    renderAcctListBatch(sorted, body, shownRef);
  }
  if(resetScroll){ const wrap = body.closest('.list-modal-scroll'); if(wrap) wrap.scrollTop = 0; }
}
function sortListModalBy(key){
  listModalState.sort = nextSort(listModalState.sort, key);
  renderListModalBody(true);
}
window.sortListModalBy = sortListModalBy;
function showListModal(title, sub, headHTML, type, list, defaultSort){
  document.getElementById('listModalTitle').textContent = title;
  document.getElementById('listModalSub').textContent = sub || '';
  document.getElementById('listModalHead').innerHTML = headHTML;
  listModalState = {list, type, sort: defaultSort || {key:'os',dir:'desc'}, addressFilter:''};
  const addrInput = document.getElementById('listModalAddressFilterInput');
  if(addrInput){
    addrInput.value = '';
    addrInput.onchange = () => { listModalState.addressFilter = addrInput.value; renderListModalBody(true); };
  }
  renderListModalBody();
  document.getElementById('listModalOverlay').classList.add('show');
  const wrap = document.getElementById('listModalBody').closest('.list-modal-scroll');
  if(__listModalScrollHandler) wrap.removeEventListener('scroll', __listModalScrollHandler);
  __listModalScrollHandler = ()=>{
    if(listModalState.type!=='acct') return;
    if(wrap.scrollTop+wrap.clientHeight>wrap.scrollHeight-400) renderAcctListBatch(listModalState.sortedList, document.getElementById('listModalBody'), listModalState.shownRef);
  };
  wrap.addEventListener('scroll', __listModalScrollHandler);
}
function closeListModal(){ document.getElementById('listModalOverlay').classList.remove('show'); }
function showAcctListModal(title, sub, list){ showListModal(title, sub, ACCT_LIST_HEAD, 'acct', list, {key:'os',dir:'desc'}); }
function showCustListModal(title, sub, list){ showListModal(title, sub, CUST_LIST_HEAD, 'cust', list, {key:'os',dir:'desc'}); }
window.showAcctListModal = showAcctListModal;
window.showCustListModal = showCustListModal;

function jsq(s){ return String(s).replace(/\\/g,'\\\\').replace(/'/g,"\\'").replace(/"/g,'&quot;'); }

function barRows(items){
  const max = Math.max(1, ...items.map(i=>i.value));
  const anyBadge = items.some(i=>i.badge);
  return items.map(it=>{
    const pct = Math.max(2, (it.value/max*100));
    return `<div class="bar-row${anyBadge?' has-npa':''}${it.onclick?' clickable':''}"${it.onclick?` onclick="${it.onclick}"`:''}>
      <div class="bar-label" title="${esc(it.label)}">${esc(it.label)}</div>
      <div class="bar-track"><div class="bar-fill" style="width:${pct.toFixed(1)}%;background:${it.color||'var(--accent)'};color:${it.color||'var(--accent)'}"></div></div>
      <div class="bar-value">${it.valueLabel}</div>
      ${anyBadge?`<div class="bar-npa-badge" style="color:${it.badge?(it.badgeColor||'var(--ink)'):'var(--ink-mute)'}">${it.badge?esc(it.badge)+'<span class=\"bar-npa-tag\">NPA</span>':'—'}</div>`:''}
    </div>`;
  }).join('');
}

function kpiTile(label, value, sub, onclick){
  return `<div class="kpi-tile${onclick?' clickable':''}"${onclick?` onclick="${onclick}"`:''}>
    <div class="kpi-label">${esc(label)}</div>
    <div class="kpi-value">${value}</div>
    ${sub?`<div class="kpi-sub">${sub}</div>`:''}
  </div>`;
}

/* Lucide-style icons (rounded, 2px stroke, 24x24 viewBox) for the hero KPI
   row and insight strip -- hand-drawn to match the icon convention already
   used throughout the app (stroke="currentColor" so each card tints its own
   icon via CSS). */
const ICON_BANKNOTE = '<rect x="2" y="6" width="20" height="12" rx="3"/><circle cx="12" cy="12" r="2.5"/><path d="M6 12h.01M18 12h.01"/>';
const ICON_USERS = '<path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/>';
const ICON_ALERT_TRIANGLE = '<path d="m21.7 18-8-14a2 2 0 0 0-3.5 0l-8 14A2 2 0 0 0 4 21h16a2 2 0 0 0 1.7-3Z"/><path d="M12 9v4"/><path d="M12 17h.01"/>';
const ICON_TICKET = '<path d="M2 9a3 3 0 0 1 0 6v2a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-2a3 3 0 0 1 0-6V7a2 2 0 0 0-2-2H4a2 2 0 0 0-2 2Z"/><path d="M13 5v2M13 11v2M13 17v2"/>';
const ICON_ALERT_CIRCLE = '<circle cx="12" cy="12" r="10"/><path d="M12 8v4"/><path d="M12 16h.01"/>';
const ICON_LANDMARK = '<path d="M3 21h18"/><path d="M3 10h18"/><path d="M5 6l7-3 7 3"/><path d="M4 10v11"/><path d="M20 10v11"/><path d="M8 14v3"/><path d="M12 14v3"/><path d="M16 14v3"/>';
const ICON_MAP = '<path d="M14.106 5.553a2 2 0 0 0 1.788 0l3.659-1.83A1 1 0 0 1 21 4.619v12.764a1 1 0 0 1-.553.894l-4.553 2.277a2 2 0 0 1-1.788 0l-4.212-2.106a2 2 0 0 0-1.788 0l-3.659 1.83A1 1 0 0 1 3 19.381V6.618a1 1 0 0 1 .553-.894l4.553-2.277a2 2 0 0 1 1.788 0z"/><path d="M15 5.764v15"/><path d="M9 3.236v15"/>';
const ICON_STAR = '<path d="M11.525 2.295a.53.53 0 0 1 .95 0l2.31 4.679a2.123 2.123 0 0 0 1.595 1.16l5.166.756a.53.53 0 0 1 .294.904l-3.736 3.638a2.123 2.123 0 0 0-.611 1.878l.882 5.14a.53.53 0 0 1-.771.56l-4.618-2.428a2.122 2.122 0 0 0-1.973 0L6.396 21.01a.53.53 0 0 1-.77-.56l.881-5.139a2.122 2.122 0 0 0-.611-1.879L2.16 9.795a.53.53 0 0 1 .294-.906l5.165-.755a2.122 2.122 0 0 0 1.597-1.16z"/>';
const ICON_TARGET = '<circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="6"/><circle cx="12" cy="12" r="2"/>';
const ICON_NOTE = '<path d="M12 17v5"/><path d="M9 10.76a2 2 0 0 1 1.11-1.79l1-.5a2 2 0 0 1 1.78 0l1 .5A2 2 0 0 1 15 10.76V15H9Z"/><path d="M8 15h8l1 2H7l1-2Z"/>';
function svgIcon(pathData){ return `<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${pathData}</svg>`; }

function heroKpiCard(opts){
  const side = (opts.badge||opts.corner) ? `<div class="hero-kpi-side">${opts.badge||''}${opts.corner||''}</div>` : '';
  return `<div class="hero-kpi-card${opts.onclick?' clickable':''}"${opts.onclick?` onclick="${opts.onclick}"`:''} style="--hero-tint:${opts.tint};--hero-color:${opts.color}">
    <div class="hero-kpi-main">
      <div class="hero-kpi-icon">${svgIcon(opts.icon)}</div>
      <div class="hero-kpi-label">${esc(opts.label)}</div>
      <div class="hero-kpi-value" id="${opts.id}">${opts.fallback||'—'}</div>
      <div class="hero-kpi-sub">${opts.sub}</div>
    </div>
    ${side}
  </div>`;
}

/* Small search-icon button dropped into a section heading wherever a raw
   account list is shown (directly, or via the shared list-drill-down
   modal) -- reuses the existing Quick Search (Cmd+K) palette rather than
   building a second search UI, since that already looks up a borrower by
   name/account no./customer ID/mobile and opens their settlement detail. */
function sectionSearchBtn(){
  return `<button type="button" class="section-search-btn" onclick="openCmdk()" title="Search a borrower by name or account no." aria-label="Search a borrower">
    <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4"><circle cx="11" cy="11" r="7"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>
  </button>`;
}

let currentDashStats = null;
const BUCKET_LABELS = {ne:'Not yet eligible (≤ 6 months)', y1:'6 months – 1 year', y13:'1 – 3 years', y3p:'3+ years'};

/* Filter changes (region/branch) call this instead of renderDashboard()
   directly, so the swap reads as a soft cross-fade — dim briefly, replace
   the numbers/charts while still dimmed, then ease back in — instead of
   the whole panel abruptly flashing blank. */
function renderDashboardSmooth(){
  const el = document.getElementById('dashboardArea');
  if(!el){ renderDashboard(); return; }
  el.classList.add('dash-updating');
  el.classList.add('no-card-anim');
  setTimeout(()=>{
    renderDashboard();
    requestAnimationFrame(()=>{ el.classList.remove('dash-updating'); });
  }, 90);
}

function drillBranch(branch){
  const sel = document.getElementById('dashBranchFilter');
  if(sel){ sel.value = branch; renderDashboardSmooth(); }
}
function showAssetList(code){
  if(!currentDashStats) return;
  const list = currentDashStats.acctList.filter(a=>a.asset===code).sort((a,b)=>b.os-a.os);
  showAcctListModal(assetLabel(code)+' — Accounts', list.length.toLocaleString('en-IN')+' account(s)', list);
}
function showBucketList(bucketId){
  if(!currentDashStats) return;
  const list = currentDashStats.acctList.filter(a=>a.bucketId===bucketId).sort((a,b)=>b.os-a.os);
  showAcctListModal('NPA Ageing — '+(BUCKET_LABELS[bucketId]||bucketId), list.length.toLocaleString('en-IN')+' account(s)', list);
}
function showSchemeList(schemeKey){
  if(!currentDashStats) return;
  const list = currentDashStats.acctList.filter(a=>a.scheme===schemeKey).sort((a,b)=>b.os-a.os);
  showAcctListModal((schemeKey==='KCC'?'KCC (CC004)':'Non-KCC')+' — Accounts', list.length.toLocaleString('en-IN')+' account(s)', list);
}
function showSlabList(slabId){
  if(!currentDashStats) return;
  const def = SLAB_DEFS.find(sl=>sl.id===slabId);
  const list = currentDashStats.acctList.filter(a=>a.slabId===slabId).sort((a,b)=>b.os-a.os);
  showAcctListModal('Outstanding Slab — '+(def?def.label:slabId), list.length.toLocaleString('en-IN')+' account(s)', list);
}
function showHighValueCustList(){
  if(!currentDashStats) return;
  showCustListModal('Customers ≥ ₹10 Lakh O/S', currentDashStats.highValueCustList.length.toLocaleString('en-IN')+' customer(s), high → low', currentDashStats.highValueCustList);
}
window.drillBranch = drillBranch;
window.openRollbackReview = openRollbackReview;
window.showAssetList = showAssetList;
window.showBucketList = showBucketList;
window.showSchemeList = showSchemeList;
window.showSlabList = showSlabList;
window.showHighValueCustList = showHighValueCustList;

/* Shown in the top-right corner of the "Total Outstanding" hero card,
   below the NPA% badge -- a March/June + gap treatment using the
   per-branch NPA March/June figures from the Branch Advance upload.
   Only aggregates
   over branches that actually have a Mar/Jun figure (and only compares
   against THOSE branches' current O/S), same safeguard as the advance
   aggregation just above -- so a partial upload never produces a
   misleading gap by comparing against branches with no baseline. */
function dashboardCornerStats(s){
  let marOS=0, marBase=0, marN=0, junOS=0, junBase=0, junN=0;
  s.branchMap.forEach((v)=>{
    const rec = DATA.branchAdvances[v.solId];
    if(rec && rec.npaMar26!=null){ marOS+=v.os; marBase+=rec.npaMar26; marN++; }
    if(rec && rec.npaJun26!=null){ junOS+=v.os; junBase+=rec.npaJun26; junN++; }
  });
  if(!marN && !junN) return '';
  const gapLine = (v) => { const improved = v<=0; return `<span style="color:${improved?'var(--green)':'var(--red)'}">${improved?'▼':'▲'} ${fmtCr(Math.abs(v))}</span>`; };
  let html = '<div class="hero-kpi-corner-stats">';
  if(marN) html += `<div class="hero-kpi-corner-group"><div class="hero-kpi-corner-row"><span>Mar</span><b>${fmtCr(marBase)}</b></div><div class="hero-kpi-corner-gap">${gapLine(marOS-marBase)}</div></div>`;
  if(junN) html += `<div class="hero-kpi-corner-group"><div class="hero-kpi-corner-row"><span>Jun</span><b>${fmtCr(junBase)}</b></div><div class="hero-kpi-corner-gap">${gapLine(junOS-junBase)}</div></div>`;
  html += '</div>';
  return html;
}
/* Recovery Dashboard (branch portal) only -- Today's NPA total, the gap
   against the existing March 2026 baseline (DATA.branchAdvances/npaMar26,
   the same figure dashboardCornerStats() above already shows in
   miniature), and a placeholder for the branch's target ahead of the
   NEXT fiscal year-end (March 2027) -- not yet available anywhere in this
   app's data, shown honestly as "Not set" rather than guessed. These are
   all NPA-position figures, not slippage figures, so they belong here on
   the Dashboard rather than on the PNPA Slippage page (Alok's own
   correction, 2026-09-25 -- they had briefly been placed there instead). */
function dashboardNpaTargetStrip(branchName){
  const solId = loggedInSolId();
  const todayNpa = computeDashboardStats(branchName || null).totalOS;
  const adv = solId ? DATA.branchAdvances[String(solId)] : null;
  const marGapHtml = (adv && adv.npaMar26!=null)
    ? (()=>{ const gap = todayNpa - adv.npaMar26; const improved = gap<=0;
        return `<div class="npa-target-tile"><span class="lbl">Gap from March 2026</span><span class="val" style="color:${improved?'var(--green)':'var(--red)'}">${improved?'▼':'▲'} ${fmtCr(Math.abs(gap))}</span></div>`; })()
    : `<div class="npa-target-tile"><span class="lbl">Gap from March 2026</span><span class="val muted">Baseline not uploaded</span></div>`;
  return `<div class="npa-target-strip">
    <div class="npa-target-tile"><span class="lbl">Today's NPA</span><span class="val">${fmtCr(todayNpa)}</span></div>
    ${marGapHtml}
    <div class="npa-target-tile"><span class="lbl">Target for March 2027</span><span class="val muted">Not set yet</span></div>
  </div>`;
}

/* Branch profile card shown at the top of the Dashboard. A single branch
   picked from #dashBranchFilter reads its Sol ID off s.branchMap (captured
   straight from the real NPA rows during computeDashboardStats, so it
   matches even though the raw branch-name spelling differs slightly from
   BRANCH_LIST/BRANCH_META's canonical form). "Regional Office" (blank
   filter, the whole book) has no branchMap entry of its own -- R O Hathras
   (Sol ID 9269) never carries NPA accounts, it's the administrative office,
   not a lending branch -- so that case is hardcoded to 9269 instead, since
   "Regional Office" on this dropdown always means that one office. */
function dashboardBranchInfoCard(branchFilter, s){
  const solId = branchFilter ? (s.branchMap.get(branchFilter)||{}).solId||'' : '9269';
  const branchName = branchFilter || 'R O Hathras';
  const meta = BRANCH_META[Number(solId)] || {};
  const bc = DATA.branchContacts[solId] || {};
  const listEntry = BRANCH_LIST.find(([,nid])=>String(nid)===String(solId));
  const oldId = listEntry ? listEntry[0] : '';
  const item = (label,val) => val ? `<div><div class="k">${esc(label)}</div><div class="v">${val}</div></div>` : '';
  const masterAddress = masterAddressOf(meta);
  const address = esc(bc.address) || (masterAddress ? esc(masterAddress) : '');
  const items = [
    item('District', meta.district ? esc(meta.district) : ''),
    item('Branch Email', meta.email ? `<a href="mailto:${esc(meta.email)}" onclick="event.stopPropagation()">${esc(meta.email)}</a>` : ''),
    item('Address', address),
  ].filter(Boolean).join('');
  if(!items) return '';
  return `<div class="card branch-info-card"${solId?` onclick="showBranchCard(${solId})" role="button" tabindex="0"`:''}>
    <div class="branch-info-head">
      <div><div class="bname">${esc(branchName)}</div><div class="baddr">${solId?`Sol ID ${esc(solId)} &middot; Old ${esc(oldId)}`:''}</div></div>
      ${solId?'<div class="branch-info-cta">Full details →</div>':''}
    </div>
    <div class="info-grid">${items}</div>
  </div>`;
}
function renderDashboard(){
  const el = document.getElementById('dashboardArea');
  if(!el) return;
  const filterSel = document.getElementById('dashBranchFilter');
  const lockedBranch = loggedInBranchName();
  const branchFilter = lockedBranch || (filterSel ? filterSel.value : '');
  const s = computeDashboardStats(branchFilter || null);
  currentDashStats = s;
  populateBranchFilter(s.allBranches);
  if(lockedBranch && filterSel){
    filterSel.value = lockedBranch;
    // Swap the (still-functional, just now redundant) dropdown for a
    // plain "Your Branch" label the first time this renders -- the whole
    // promise of Sol-login is "this shows your branch," and a still-
    // visible picker (even disabled) undermines that. Idempotent: only
    // acts once, since a repeat renderDashboard() call would otherwise
    // find the select already hidden.
    if(filterSel.style.display !== 'none'){
      filterSel.style.display = 'none';
      const label = document.createElement('span');
      label.className = 'recovery-branch-lock';
      label.textContent = lockedBranch;
      filterSel.insertAdjacentElement('afterend', label);
    }
  }
  updateDashTitle();

  const assetItems = ASSET_ORDER.filter(k=>s.assetMix[k]).map(k=>({
    label: assetLabel(k)+' ('+k+')', value:s.assetMix[k].os, color:ASSET_SEV_COLOR[k],
    valueLabel:`${s.assetMix[k].count.toLocaleString('en-IN')} · ${fmtCr(s.assetMix[k].os)}`,
    onclick:`showAssetList('${k}')`
  }));
  Object.keys(s.assetMix).filter(k=>!ASSET_ORDER.includes(k)).forEach(k=>assetItems.push({
    label:k, value:s.assetMix[k].os, color:'var(--ink-mute)',
    valueLabel:`${s.assetMix[k].count.toLocaleString('en-IN')} · ${fmtCr(s.assetMix[k].os)}`,
    onclick:`showAssetList('${jsq(k)}')`
  }));

  const branchTop = [...s.branchMap.entries()].sort((a,b)=>b[1].os-a[1].os).slice(0,10)
    .map(([branch,v])=>{
      const rec = DATA.branchAdvances[v.solId];
      const npaPct = rec && rec.adv>0 ? (v.os/rec.adv*100) : null;
      return {label:branch, value:v.os, color:'var(--accent)',
        valueLabel:`${v.count.toLocaleString('en-IN')} · ${fmtCr(v.os)} · ${(s.totalOS?(v.os/s.totalOS*100):0).toFixed(2)}%`,
        badge: npaPct!==null ? npaPct.toFixed(1)+'%' : null,
        badgeColor: npaPct!==null ? npaPctSeverity(npaPct).color : null,
        onclick:`drillBranch('${jsq(branch)}')`};
    });

  const agingItems = s.buckets.map(b=>({label:b.label, value:b.os, color:'var(--accent-2)',
    valueLabel:`${b.count.toLocaleString('en-IN')} · ${fmtCr(b.os)}`,
    onclick:`showBucketList('${b.id}')`}));

  const kccPct = s.totalOS ? (s.schemeMix.KCC.os/s.totalOS*100) : 0;
  const nonKccPct = s.totalOS ? (s.schemeMix.NONKCC.os/s.totalOS*100) : 0;
  const kccSeg = [
    {label:'KCC (CC004)', value:s.schemeMix.KCC.os, color:'var(--green)',
      valueLabel:`${s.schemeMix.KCC.count.toLocaleString('en-IN')} A/C · ${fmtCr(s.schemeMix.KCC.os)} · ${kccPct.toFixed(1)}%`,
      onclick:`showSchemeList('KCC')`},
    {label:'Non-KCC', value:s.schemeMix.NONKCC.os, color:'var(--accent-2)',
      valueLabel:`${s.schemeMix.NONKCC.count.toLocaleString('en-IN')} A/C · ${fmtCr(s.schemeMix.NONKCC.os)} · ${nonKccPct.toFixed(1)}%`,
      onclick:`showSchemeList('NONKCC')`},
  ];
  const slabColors = ['var(--sev-1)','var(--sev-2)','var(--sev-3)','var(--sev-4)'];
  const slabSeg = s.slabs.map((sl,i)=>({label:sl.label, value:sl.os, color:slabColors[i],
    valueLabel:`${sl.count.toLocaleString('en-IN')} A/C · ${fmtCr(sl.os)}`,
    onclick:`showSlabList('${sl.id}')`}));

  const highRiskOS = (s.assetMix.DA3?s.assetMix.DA3.os:0) + (s.assetMix.LOSS?s.assetMix.LOSS.os:0);
  const highRiskPct = s.totalOS ? (highRiskOS/s.totalOS*100) : 0;
  const avgTicket = s.totalAccounts ? s.totalOS/s.totalAccounts : 0;

  /* NPA % (NPA outstanding ÷ total advance) for whatever's currently in
     view -- the whole book when "Regional Office" is selected, or just that
     branch when one is picked from the filter, since s.branchMap already
     reflects that filter. Only aggregates over branches with an uploaded
     advance figure, so a partially-uploaded advance file never silently
     understates the ratio by dividing by a smaller, incomplete total. */
  let advOsSum=0, advSum=0, advBranchCount=0;
  s.branchMap.forEach((v)=>{
    const rec = DATA.branchAdvances[v.solId];
    if(rec && rec.adv>0){ advOsSum+=v.os; advSum+=rec.adv; advBranchCount++; }
  });
  const aggNpaPct = advSum>0 ? (advOsSum/advSum*100) : null;
  const heroCorner = dashboardCornerStats(s);
  let heroNpaBadge = '';
  if(aggNpaPct!==null){
    const sev = npaPctSeverity(aggNpaPct);
    heroNpaBadge = `<div class="hero-kpi-badge" style="background:${sev.soft};color:${sev.color}">${aggNpaPct.toFixed(1)}% NPA</div>`;
  }

  /* "What should happen next" -- the single largest concentration of aged,
     actionable exposure (excludes the "not yet eligible" bucket, since that
     one isn't actionable yet), computed fresh from real data every render
     rather than a fixed/fabricated callout. */
  const actionableBuckets = s.buckets.filter(b=>b.id!=='ne' && b.os>0);
  const topBucket = actionableBuckets.length ? actionableBuckets.reduce((max,b)=>b.os>max.os?b:max) : null;

  el.innerHTML = `
    ${dashboardBranchInfoCard(branchFilter, s)}
    ${lockedBranch ? dashboardNpaTargetStrip(lockedBranch) : ''}
    <div class="hero-kpi-row">
      ${heroKpiCard({id:'heroTotalOs', label:'Total Outstanding', fallback:fmtCr(s.totalOS), sub:s.totalAccounts.toLocaleString('en-IN')+' accounts', icon:ICON_BANKNOTE, tint:'var(--accent-soft)', color:'var(--accent)', badge:heroNpaBadge, corner:heroCorner})}
      ${heroKpiCard({id:'heroTotalAccts', label:'Total Accounts', fallback:s.totalAccounts.toLocaleString('en-IN'), sub:s.custCount.toLocaleString('en-IN')+' unique customers', icon:ICON_USERS, tint:'var(--gauge-track)', color:'var(--accent-2)'})}
      ${heroKpiCard({id:'heroHighRisk', label:'High-Risk Exposure', fallback:fmtCr(highRiskOS), sub:'DA3 + Loss · '+highRiskPct.toFixed(1)+'% of book', icon:ICON_ALERT_TRIANGLE, tint:'var(--red-soft)', color:'var(--red)', onclick:(s.assetMix.LOSS||s.assetMix.DA3)?`showAssetList('${s.assetMix.LOSS?'LOSS':'DA3'}')`:''})}
      ${heroKpiCard({id:'heroAvgTicket', label:'Average Ticket Size', fallback:fmtINR2(avgTicket), sub:'per account, this book', icon:ICON_TICKET, tint:'var(--amber-soft)', color:'var(--amber)'})}
    </div>

    ${topBucket ? `
    <div class="insight-strip clickable" onclick="showBucketList('${topBucket.id}')">
      <div class="insight-icon">${svgIcon(ICON_ALERT_CIRCLE)}</div>
      <div class="insight-body">
        <div class="insight-title">Recovery focus: ${esc(BUCKET_LABELS[topBucket.id]||topBucket.label)}</div>
        <div class="insight-text">${fmtCr(topBucket.os)} across ${topBucket.count.toLocaleString('en-IN')} account(s) — the largest concentration of aged exposure in this book.</div>
      </div>
      <div class="insight-cta">View list →</div>
    </div>` : ''}

    <div class="chart-grid">
      <div class="chart-card">
        <div class="chart-title">Total Outstanding — KCC vs Non-KCC<span class="chart-sub">scheme CC004 = KCC · every other scheme = Non-KCC</span></div>
        <div class="kcc-total-strip">
          <div><div class="lbl">Total A/C Amount</div><div class="val">${fmtCr(s.totalOS)}</div></div>
          <div><div class="lbl">Total Accounts</div><div class="val">${s.totalAccounts.toLocaleString('en-IN')}</div></div>
        </div>
        <div class="donut-flex">
          ${donutCard(kccSeg, undefined, fmtCr(s.totalOS), 'Total O/S')}
          <div class="donut-legend">${donutLegend(kccSeg)}</div>
        </div>
        <div class="split-stat-grid">
          <div class="split-stat kcc clickable" onclick="showSchemeList('KCC')">
            <div class="split-stat-label">KCC (CC004)</div>
            <div class="split-stat-amt">${fmtCr(s.schemeMix.KCC.os)}</div>
            <div class="split-stat-count">${s.schemeMix.KCC.count.toLocaleString('en-IN')} A/C · ${fmtINR2(s.schemeMix.KCC.os)} · ${kccPct.toFixed(1)}% share</div>
          </div>
          <div class="split-stat nonkcc clickable" onclick="showSchemeList('NONKCC')">
            <div class="split-stat-label">Non-KCC</div>
            <div class="split-stat-amt">${fmtCr(s.schemeMix.NONKCC.os)}</div>
            <div class="split-stat-count">${s.schemeMix.NONKCC.count.toLocaleString('en-IN')} A/C · ${fmtINR2(s.schemeMix.NONKCC.os)} · ${nonKccPct.toFixed(1)}% share</div>
          </div>
        </div>
      </div>

      <div class="chart-card">
        <div class="chart-title">Outstanding by Amount Slab<span class="chart-sub">account-wise O/S buckets</span></div>
        <div class="donut-flex">
          ${donutCard(slabSeg, undefined, fmtCr(s.totalOS), 'Total O/S')}
          <div class="donut-legend">${donutLegend(slabSeg)}</div>
        </div>
      </div>
    </div>

    <div class="chart-grid">
      <div class="chart-card">
        <div class="chart-title">Asset Classification Mix<span class="chart-sub">by outstanding balance · RBI IRAC norms · tap a row for the list</span></div>
        <div class="bar-list">${barRows(assetItems)}</div>
      </div>
      <div class="chart-card">
        <div class="chart-title">NPA Ageing<span class="chart-sub">days since NPA date · tap a row for the list</span></div>
        <div class="bar-list">${barRows(agingItems)}</div>
      </div>
      ${branchFilter ? '' : `
      <div class="chart-card chart-card-wide">
        <div class="chart-title">Top Branches by Exposure<span class="chart-sub">top 10 of ${s.branchCount.toLocaleString('en-IN')} branch(es) · tap to drill into a branch</span></div>
        <div class="bar-list">${barRows(branchTop)}</div>
      </div>`}
    </div>

    <div class="section-label">Customer-Wise Outstanding</div>
    <div class="kpi-grid">
      ${kpiTile('Total Unique Customers', s.custCount.toLocaleString('en-IN'), fmtCr(s.totalOS)+' combined outstanding')}
      ${kpiTile('Customers ≥ ₹10 Lakh O/S', s.highValueCustCount.toLocaleString('en-IN'), fmtCr(s.highValueOS)+(s.custCount?' · '+((s.highValueCustCount/s.custCount)*100).toFixed(1)+'% of customers':'')+' · tap to view list', 'showHighValueCustList()')}
    </div>

    <div class="section-label">All Accounts by Outstanding<span class="chart-sub">${s.totalAccounts.toLocaleString('en-IN')} account(s) · tap a column to sort · scroll for more</span>${sectionSearchBtn()}</div>
    <div class="bank-filter-row" style="margin-bottom:10px">
      <input type="text" id="dashAddressFilterInput" class="dash-select" placeholder="Filter by Address…" value="${esc(dashAddressFilter)}" style="max-width:220px">
      <input type="date" id="dashNpaDateFromInput" class="dash-select" value="${esc(dashNpaDateFrom)}" style="max-width:170px" title="NPA Date from">
      <span style="color:var(--ink-mute);font-size:12px;align-self:center">to</span>
      <input type="date" id="dashNpaDateToInput" class="dash-select" value="${esc(dashNpaDateTo)}" style="max-width:170px" title="NPA Date to">
    </div>
    <div class="dash-table-wrap acct-list-scroll" id="acctListWrap">
      <table class="dash-table">
        <thead id="acctListHead"><tr>
          <th class="sortable" data-key="acctNo" tabindex="0" role="button" aria-sort="none" onclick="sortAcctListBy('acctNo')">Account<span class="sort-ic">▾</span></th>
          <th class="tal sortable" data-key="name" tabindex="0" role="button" aria-sort="none" onclick="sortAcctListBy('name')">Customer<span class="sort-ic">▾</span></th>
          <th class="tal sortable" data-key="address" tabindex="0" role="button" aria-sort="none" onclick="sortAcctListBy('address')">Address<span class="sort-ic">▾</span></th>
          <th class="sortable" data-key="asset" tabindex="0" role="button" aria-sort="none" onclick="sortAcctListBy('asset')">Asset<span class="sort-ic">▾</span></th>
          <th class="sortable" data-key="os" tabindex="0" role="button" aria-sort="none" onclick="sortAcctListBy('os')">Amount<span class="sort-ic">▾</span></th>
          <th class="tal sortable" data-key="npaDate" tabindex="0" role="button" aria-sort="none" onclick="sortAcctListBy('npaDate')">NPA Date<span class="sort-ic">▾</span></th>
        </tr></thead>
        <tbody id="acctListBody"></tbody>
      </table>
    </div>
  `;
  initAcctListScroll(dashFilterAcctList(s.allAcctSorted));
  const dashAddrInput = document.getElementById('dashAddressFilterInput');
  if(dashAddrInput) dashAddrInput.onchange = () => { dashAddressFilter = dashAddrInput.value; initAcctListScroll(dashFilterAcctList(s.allAcctSorted)); };
  const dashNpaFromInput = document.getElementById('dashNpaDateFromInput');
  if(dashNpaFromInput) dashNpaFromInput.onchange = () => { dashNpaDateFrom = dashNpaFromInput.value; initAcctListScroll(dashFilterAcctList(s.allAcctSorted)); };
  const dashNpaToInput = document.getElementById('dashNpaDateToInput');
  if(dashNpaToInput) dashNpaToInput.onchange = () => { dashNpaDateTo = dashNpaToInput.value; initAcctListScroll(dashFilterAcctList(s.allAcctSorted)); };

  const heroOs = document.getElementById('heroTotalOs');
  if(heroOs) animateNumber(heroOs, 0, s.totalOS, fmtCr, 900);
  const heroAccts = document.getElementById('heroTotalAccts');
  if(heroAccts) animateNumber(heroAccts, 0, s.totalAccounts, n=>Math.round(n).toLocaleString('en-IN'), 900);
  const heroRisk = document.getElementById('heroHighRisk');
  if(heroRisk) animateNumber(heroRisk, 0, highRiskOS, fmtCr, 900);
  const heroTicket = document.getElementById('heroAvgTicket');
  if(heroTicket) animateNumber(heroTicket, 0, avgTicket, fmtINR2, 900);
}

/* ---------- Daily PNPA (Potential NPA) -- whole-bank, branch-wise, bucketed by scheme ----------
   A separate dataset from DATA.npa: the source file is the whole-bank HO
   "Daily PNPA" export (all 65 regions), but this tab only ever keeps
   Hathras's own rows (Alok's ask -- this is a Hathras-scoped app, the
   other 64 regions' potential-NPA accounts aren't his to work), and drops
   zero-balance accounts (an SMA flag with a ₹0 outstanding isn't
   actionable). Rows are stored as compact arrays (see PC below) instead
   of the full 35-column HO layout -- only the fields this tab actually
   uses are kept. */
const PC = {REGION:0, BRANCH:1, SCHEME:2, ACCT:3, NAME:4, OS:5, CADU:6, LIMIT:7, REVIEW:8, REASON:9, CUSTNPADATE:10};
/* "Limit Review" is its own bucket, pulled out ahead of the scheme-based
   split -- an account flagged Limit Review is routed there regardless of
   scheme code, so KCC/KCC-AH/Other only ever show accounts NOT already
   called out for a limit review (no double-counting across buckets). */
const PNPA_BUCKETS = [
  {key:'kcc', label:'KCC', sub:'Scheme code CC004 · reason "KCC-Disbrsmnt-36" only'},
  {key:'kccah', label:'KCC — Animal Husbandry', sub:'Scheme code CC043, excluding Limit Review'},
  {key:'limitreview', label:'Limit Review', sub:'Flagged "Limit Review", any scheme'},
  {key:'other', label:'Other Schemes', sub:'All remaining scheme codes, excluding Limit Review'},
];
function pnpaBucketOfRow(row){
  if(String(row[PC.REASON]||'').includes('Limit Review')) return 'limitreview';
  const scheme = row[PC.SCHEME], reason = String(row[PC.REASON]||'');
  if(scheme==='CC004') return reason.includes('KCC-Disbrsmnt-36') ? 'kcc' : 'other';
  return scheme==='CC043' ? 'kccah' : 'other';
}
/* The source file's own "Remarks" column is almost always just "-" (no real
   content) -- the actual why-is-this-flagged info lives in "Reasons"
   instead (e.g. "LAANPA,LimReview"), so that's what gets shown and searched
   as this tab's reason/remark field. "LimReview" is spelled out as "Limit
   Review" since Alok specifically calls that one out; the other codes are
   shown as-is rather than guessed-translated. */
function formatPnpaReasons(raw){
  return String(raw||'').split(',').map(s=>s.trim()).filter(Boolean)
    .map(s=>s==='LimReview'?'Limit Review':s).join(', ');
}
function parsePnpaRows(headerCells, dataRows){
  const header = headerCells.map(normHeader);
  const idx = (name) => header.indexOf(normHeader(name));
  const iRegion=idx('region'), iBranch=idx('branch'), iAcct=idx('accountno'), iScheme=idx('schemecode'),
    iName=idx('accountname'), iBal=idx('balanceamount'), iCadu=idx('cadu'), iLimit=idx('limit'),
    iReview=idx('reviewdate'), iReasons=idx('reasons'), iCustNpa=idx('custnpadate');
  const missing = [];
  if(iAcct<0) missing.push('Account No');
  if(iBranch<0) missing.push('Branch');
  if(iScheme<0) missing.push('Scheme Code');
  if(iBal<0) missing.push('Balance Amount');
  if(iCadu<0) missing.push('CADU');
  if(iRegion<0) missing.push('Region');
  if(missing.length) throw new Error('Missing required column(s): '+missing.join(', ')+'. Check this file matches the "Daily PNPA" export layout.');
  const rows = [];
  for(const row of dataRows){
    if(!row || row.length<3) continue;
    const region = cellStr(row, iRegion);
    if(region.toUpperCase()!=='HATHRAS') continue;
    const acctRaw = cellStr(row, iAcct);
    if(!acctRaw) continue;
    const bal = parseFloat(row[iBal])||0;
    if(bal===0) continue;
    let acctNo = acctRaw;
    if(looksScientific(acctRaw)) acctNo = expandSci(acctRaw);
    const reviewDt = toDate(iReview>=0?row[iReview]:'');
    const custNpaDt = toDate(iCustNpa>=0?row[iCustNpa]:'');
    rows.push([
      region, cellStr(row, iBranch), cellStr(row, iScheme), acctNo, cellStr(row, iName),
      bal, parseFloat(row[iCadu])||0,
      iLimit>=0 ? (parseFloat(row[iLimit])||0) : 0,
      reviewDt ? fmtDate(reviewDt) : '',
      iReasons>=0 ? formatPnpaReasons(cellStr(row, iReasons)) : '',
      custNpaDt ? fmtDate(custNpaDt) : '',
    ]);
  }
  return rows;
}
let PNPA_DATA = null;
let __pendingPnpaData = null;
let pnpaBucketTab = 'kcc';
let pnpaBranchFilter = '';
function setPnpaBucketTab(tab){ pnpaBucketTab = tab; renderPnpaDashboardBody(); }
window.setPnpaBucketTab = setPnpaBucketTab;

async function handlePnpaUpload(evt){
  const file = evt.target.files[0];
  if(!file) return;
  await ensureXLSX();
  const labelEl = document.getElementById('pnpaUploadDropLabel');
  if(labelEl) labelEl.textContent = file.name;
  const statusEl = document.getElementById('pnpaUploadStatus');
  statusEl.innerHTML = `<div class="upload-status info">Reading Daily PNPA file…</div>`;
  const isCsv = /\.csv$/i.test(file.name);
  const reader = new FileReader();
  reader.onerror = function(){ statusEl.innerHTML = `<div class="upload-status err">⚠ Failed to read the file from disk.</div>`; };
  reader.onload = function(e){
    try{
      let header, dataRows;
      if(isCsv){
        const allRows = parseCSV(String(e.target.result));
        header = allRows[0]||[]; dataRows = allRows.slice(1);
      } else {
        const data = new Uint8Array(e.target.result);
        const wb = XLSX.read(data, {type:'array'});
        const sheetName = wb.SheetNames.find(n=>/pnpa/i.test(n)) || wb.SheetNames[0];
        const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {header:1, raw:true, defval:''});
        header = raw[0]||[]; dataRows = raw.slice(1);
      }
      const rows = parsePnpaRows(header, dataRows);
      if(!rows.length) throw new Error('No account rows found in this file.');
      const guessed = parseAsOnDateFromFilename(file.name);
      const asOnDate = guessed ? dateToInputValue(guessed) : dateToInputValue(new Date());
      __pendingPnpaData = { asOnDate, rows };
      PNPA_DATA = __pendingPnpaData;
      const label = document.getElementById('pnpaStatusLabel');
      if(label) label.textContent = `${rows.length.toLocaleString('en-IN')} accounts loaded (${file.name})`;
      statusEl.innerHTML = `<div class="upload-status ok">✔ Parsed ${rows.length.toLocaleString('en-IN')} accounts, as on ${esc(asOnDate)}. Goes live the next time you hit Publish.</div>`;
      clearStalePublishStatus();
      const publishBtn = document.getElementById('publishBtn');
      if(publishBtn) publishBtn.disabled = false;
      if(document.querySelector('.view.active')?.dataset.view==='pnpa') renderPnpaDashboardBody();
    } catch(err){
      statusEl.innerHTML = `<div class="upload-status err">⚠ Could not read this file: ${esc(err.message||err)}</div>`;
    }
  };
  if(isCsv) reader.readAsText(file); else reader.readAsArrayBuffer(file);
}

function renderPnpaDashboard(){
  const el = document.getElementById('pnpaDashboardArea');
  if(!el) return;
  if(PNPA_DATA){ renderPnpaDashboardBody(); return; }
  el.innerHTML = `<div class="empty-state"><div class="data-loading-spinner" aria-hidden="true" style="position:static;border-color:rgba(58,123,255,.25);border-top-color:var(--accent)"></div><p style="margin-top:14px">Loading Daily PNPA data…</p></div>`;
  fetchJson(DATA_ORIGIN + 'data/pnpa.json?t=' + Date.now())
    .then(d => { PNPA_DATA = d; renderPnpaDashboardBody(); })
    .catch(() => {
      el.innerHTML = `<div class="empty-state"><h2>Could not load Daily PNPA data</h2><p>Check your internet connection, then tap Refresh.</p></div>`;
    });
}
// Recovery Dashboard (branch portal): loader for the NEW "PNPA Slippage"
// view (below) -- shares PNPA_DATA with the old hidden renderPnpaDashboard()
// above (never wired into this portal's nav) so a fetch is never duplicated
// if both happen to run in the same session.
function ensurePnpaDataLoaded(onReady, onError){
  if(PNPA_DATA){ onReady(); return; }
  fetchJson(DATA_ORIGIN + 'data/pnpa.json?t=' + Date.now())
    .then(d => { PNPA_DATA = d; onReady(); })
    .catch(onError);
}
// Weekly/Monthly PNPA (Alok, 2026-09-25): their own separate published
// files, same fetch pattern as Daily PNPA above. Deliberately NOT treated
// as fatal if missing/unreachable -- Alok may not have uploaded either one
// yet (a brand-new upload type), and the "Today" tab (Daily PNPA) should
// keep working regardless -- __pnpaWeeklyFailed/__pnpaMonthlyFailed record
// a load attempt's own failure separately from "never tried yet", so
// renderPnpaSlipView() (below) knows to stop retrying and just render an
// empty Week/Month tab rather than looping forever or blocking the page.
let PNPA_WEEKLY_DATA = null, __pnpaWeeklyFailed = false, __pnpaWeeklyLoading = false;
function ensurePnpaWeeklyDataLoaded(onReady){
  if(PNPA_WEEKLY_DATA || __pnpaWeeklyFailed){ onReady(); return; }
  if(__pnpaWeeklyLoading) return;
  __pnpaWeeklyLoading = true;
  fetchJson(DATA_ORIGIN + 'data/pnpa-weekly.json?t=' + Date.now())
    .then(d => { PNPA_WEEKLY_DATA = d; __pnpaWeeklyLoading = false; onReady(); })
    .catch(() => { __pnpaWeeklyFailed = true; __pnpaWeeklyLoading = false; onReady(); });
}
let PNPA_MONTHLY_DATA = null, __pnpaMonthlyFailed = false, __pnpaMonthlyLoading = false;
function ensurePnpaMonthlyDataLoaded(onReady){
  if(PNPA_MONTHLY_DATA || __pnpaMonthlyFailed){ onReady(); return; }
  if(__pnpaMonthlyLoading) return;
  __pnpaMonthlyLoading = true;
  fetchJson(DATA_ORIGIN + 'data/pnpa-monthly.json?t=' + Date.now())
    .then(d => { PNPA_MONTHLY_DATA = d; __pnpaMonthlyLoading = false; onReady(); })
    .catch(() => { __pnpaMonthlyFailed = true; __pnpaMonthlyLoading = false; onReady(); });
}

function pnpaBranchAgg(rows, bucket){
  const map = new Map();
  for(const r of rows){
    if(pnpaBucketOfRow(r)!==bucket) continue;
    const key = r[PC.BRANCH];
    let e = map.get(key);
    if(!e){ e = {branch:r[PC.BRANCH], count:0, os:0}; map.set(key,e); }
    e.count++; e.os += r[PC.OS];
  }
  return [...map.values()].sort((a,b)=>b.os-a.os);
}

function renderPnpaDashboardBody(){
  const el = document.getElementById('pnpaDashboardArea');
  const d = PNPA_DATA;
  if(!el) return;
  if(!d || !d.rows){ el.innerHTML = `<div class="empty-state"><h2>No Daily PNPA data yet</h2><p>Upload the Daily PNPA file from Update Data to populate this tab.</p></div>`; return; }

  document.querySelectorAll('.pnpa-report-date-val').forEach(e=>{
    const parts = (d.asOnDate||'').split('-');
    e.textContent = parts.length===3 ? `${parts[2]}-${parts[1]}-${parts[0]}` : (d.asOnDate||'—');
  });

  const allBranches = [...new Set(d.rows.map(r=>r[PC.BRANCH]))].sort((a,b)=>a.localeCompare(b));
  const branchFilterOptions = `<option value="">Regional Office</option>` +
    allBranches.map(b=>`<option value="${esc(b)}"${pnpaBranchFilter===b?' selected':''}>${esc(b)}</option>`).join('');
  const toolbar = `<div class="dash-toolbar">
    <span class="dash-toolbar-label">Branch</span>
    <select id="pnpaBranchFilterSelect" class="dash-select">${branchFilterOptions}</select>
  </div>`;

  // Hero blocks total whichever rows are currently in scope -- Regional
  // Office (all Hathras) by default, or just the selected branch's own
  // rows once one is picked, so the KCC/KCC-AH/Limit Review/Other numbers
  // always match what the Branch filter above them is set to.
  const scopedRows = pnpaBranchFilter ? d.rows.filter(r=>r[PC.BRANCH]===pnpaBranchFilter) : d.rows;
  const bucketTotals = {};
  PNPA_BUCKETS.forEach(b=>{ bucketTotals[b.key]={count:0,os:0,branches:new Set()}; });
  for(const r of scopedRows){
    const bk = pnpaBucketOfRow(r);
    bucketTotals[bk].count++; bucketTotals[bk].os += r[PC.OS]; bucketTotals[bk].branches.add(r[PC.BRANCH]);
  }
  const bucketIcon = {kcc:ICON_TARGET, kccah:ICON_STAR, limitreview:ICON_ALERT_CIRCLE, other:ICON_LANDMARK};

  const heroRow = `<div class="hero-kpi-row bank-hero-row">${PNPA_BUCKETS.map(b=>{
    const t = bucketTotals[b.key], isActive = pnpaBucketTab===b.key;
    return heroKpiCard({
      id:'pnpaHero_'+b.key, icon: bucketIcon[b.key],
      tint: isActive?'var(--accent-soft)':'rgba(120,120,140,.12)', color: isActive?'var(--accent)':'var(--ink-mute)',
      onclick:`setPnpaBucketTab('${b.key}')`,
      label: b.label,
      fallback: fmtCr(t.os),
      sub: pnpaBranchFilter
        ? `${t.count.toLocaleString('en-IN')} accounts in ${esc(pnpaBranchFilter)}`
        : `${t.count.toLocaleString('en-IN')} accounts · ${t.branches.size.toLocaleString('en-IN')} branches`,
      badge: isActive ? `<div class="hero-kpi-badge" style="background:var(--accent-soft);color:var(--accent)">Viewing</div>` : '',
    });
  }).join('')}</div>`;

  el.innerHTML = toolbar + heroRow +
    `<div class="chart-card" style="margin-top:20px">
      <div class="section-label" id="pnpaTableLabel"></div>
      <div id="pnpaBranchTableCard"></div>
    </div>`;

  const filterSel = document.getElementById('pnpaBranchFilterSelect');
  if(filterSel) filterSel.onchange = () => { pnpaBranchFilter = filterSel.value; renderPnpaDashboardBody(); };
  renderPnpaBranchTable();
}

function renderPnpaBranchTable(){
  const d = PNPA_DATA;
  const wrap = document.getElementById('pnpaBranchTableCard');
  const labelEl = document.getElementById('pnpaTableLabel');
  if(!wrap || !d) return;
  const activeBucket = PNPA_BUCKETS.find(b=>b.key===pnpaBucketTab);
  let branchAgg = pnpaBranchAgg(d.rows, pnpaBucketTab);
  if(pnpaBranchFilter) branchAgg = branchAgg.filter(r=>r.branch===pnpaBranchFilter);
  const scopeLabel = pnpaBranchFilter ? esc(pnpaBranchFilter) : 'Regional Office (all branches)';
  if(labelEl) labelEl.innerHTML = `${esc(activeBucket.label)} — Branch-wise Summary, highest O/S first<span class="chart-sub">${esc(activeBucket.sub)} · ${scopeLabel} · ${branchAgg.length.toLocaleString('en-IN')} branch(es) shown · tap a branch to see the account list</span>`;
  const rowsHtml = branchAgg.map((r,i)=>{
    return `<tr class="clickable" onclick="pnpaShowBranchAccounts('${pnpaBucketTab}','${esc(r.branch)}')">
      <td><span class="dash-rank">${i+1}</span></td>
      <td class="tal">${esc(r.branch)}</td>
      <td>${r.count.toLocaleString('en-IN')}</td>
      <td>${fmtCr(r.os)}</td>
    </tr>`;
  }).join('');
  wrap.innerHTML = `<div class="dash-table-wrap acct-list-scroll">
    <table class="dash-table">
      <thead><tr><th class="tal">Rank</th><th class="tal">Branch</th><th>Accounts</th><th>Total O/S</th></tr></thead>
      <tbody>${rowsHtml || emptyStateRowHtml(4, 'No branches match this filter')}</tbody>
    </table>
  </div>`;
}

const PNPA_ACCT_LIST_HEAD = '<tr>'
  +'<th class="sortable" data-key="acctNo" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'acctNo\')">Account<span class="sort-ic">▾</span></th>'
  +'<th class="tal sortable" data-key="name" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'name\')">Customer<span class="sort-ic">▾</span></th>'
  +'<th class="sortable" data-key="os" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'os\')">O/S<span class="sort-ic">▾</span></th>'
  +'<th class="sortable" data-key="cadu" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'cadu\')">CADU<span class="sort-ic">▾</span></th>'
  +'<th class="sortable" data-key="limit" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'limit\')">Limit<span class="sort-ic">▾</span></th>'
  +'<th class="tal sortable" data-key="reviewDate" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'reviewDate\')">Review Date<span class="sort-ic">▾</span></th>'
  +'<th class="tal sortable" data-key="reason" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'reason\')">Reason<span class="sort-ic">▾</span></th>'
  +'</tr>';
function pnpaAcctRows(list){
  if(!list.length) return emptyStateRowHtml(7, 'No accounts');
  return list.map(a=>`<tr class="clickable" onclick="showQuickAcctDetailByAcct('pnpa','${esc(a.acctNo)}')">
    <td>${esc(a.acctNo)}</td>
    <td class="tal">${esc(a.name)||'—'}</td>
    <td>${fmtINR2(a.os)}</td>
    <td>${fmtINR2(a.cadu)}</td>
    <td>${fmtINR2(a.limit)}</td>
    <td class="tal">${esc(a.reviewDate)||'—'}</td>
    <td class="tal">${esc(a.reason)||'—'}</td>
  </tr>`).join('');
}
function showPnpaListModal(title, sub, list){ showListModal(title, sub, PNPA_ACCT_LIST_HEAD, 'pnpa', list, {key:'os',dir:'desc'}); }
window.showPnpaListModal = showPnpaListModal;
function pnpaShowBranchAccounts(bucket, branch){
  const rows = PNPA_DATA.rows.filter(r=>pnpaBucketOfRow(r)===bucket && r[PC.BRANCH]===branch);
  const list = rows.map(r=>({ acctNo:r[PC.ACCT], name:r[PC.NAME], os:r[PC.OS], cadu:r[PC.CADU], limit:r[PC.LIMIT], reviewDate:r[PC.REVIEW], reason:r[PC.REASON] }));
  const bLabel = (PNPA_BUCKETS.find(b=>b.key===bucket)||{}).label || bucket;
  showPnpaListModal(`${branch} — ${bLabel}`, `Hathras · ${list.length.toLocaleString('en-IN')} account(s)`, list);
}
window.pnpaShowBranchAccounts = pnpaShowBranchAccounts;

/* ==================================================================
   Recovery Dashboard (branch portal) only -- "PNPA Slippage" tracker.
   Alok's own confirmed definition (asked directly, since the real "Daily
   PNPA" export's Cust NPA Date turned out to be a RECORD of when an
   account already became NPA, not a future prediction -- confirmed by
   direct inspection of his real file, 2026-09-25: every row with a
   non-zero Balance Amount already has this date set, and none of the
   Balance-Amount-zero rows do): "Today/This Week/This Month" means
   accounts that slipped INTO NPA that recently, grouped by Cust NPA Date,
   not a forward projection. Reuses PC/parsePnpaRows/PNPA_DATA (the
   existing "Daily PNPA" pipeline) unchanged -- this is a new view on the
   same data, not a new upload. */
const PNPA_SLIP_TABS = [
  {key:'today', label:'Today'},
  {key:'week', label:'This Week'},
  {key:'month', label:'This Month'},
];
// "Today" is still bucketed off Daily PNPA's own Cust NPA Date, since that
// file is a whole-book snapshot (every account currently on Head Office's
// PNPA watch, whenever it slipped) rather than something already scoped to
// just today -- a Cust NPA Date of today, relative to the viewer's own
// clock, is what actually means "slipped today" here.
function pnpaSlipBucketOf(custNpaDateStr, today){
  const d = toDate(custNpaDateStr);
  if(!d) return null;
  const startOfToday = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const startOfRow = new Date(d.getFullYear(), d.getMonth(), d.getDate());
  const daysAgo = Math.round((startOfToday - startOfRow) / 86400000);
  if(daysAgo !== 0) return null; // in the future (shouldn't occur) or any day but today
  return 'today';
}
function pnpaSlipSchemeKey(row){ return row[PC.SCHEME]==='CC004' ? 'kcc' : 'nonkcc'; }
function pnpaEmptySlipTotals(){ return { kcc:{cnt:0,amt:0}, nonkcc:{cnt:0,amt:0}, all:{cnt:0,amt:0} }; }
function pnpaAddRowToSlipTotals(t, row){
  const k = pnpaSlipSchemeKey(row);
  t[k].cnt++; t[k].amt += row[PC.OS];
  t.all.cnt++; t.all.amt += row[PC.OS];
}
// Today: filtered out of the Daily PNPA whole-book snapshot by Cust NPA
// Date, scoped to one branch (the locked branch -- PC.BRANCH's exact
// spelling can differ from C.SOL_DESC's, same caveat as KCC Overdue's own
// branch lock, resolved via pnpaLoggedInBranchName() below).
function pnpaAggregateToday(rows, branch, today){
  const totals = pnpaEmptySlipTotals(), list = [];
  for(const r of rows){
    if(branch && r[PC.BRANCH]!==branch) continue;
    if(pnpaSlipBucketOf(r[PC.CUSTNPADATE], today)!=='today') continue;
    pnpaAddRowToSlipTotals(totals, r);
    list.push(r);
  }
  list.sort((a,b)=>b[PC.OS]-a[PC.OS]);
  return { totals, list };
}
// This Week/This Month: Alok confirmed directly (2026-09-25) that these
// are their OWN separate files from Head Office, already scoped to that
// exact period ("this week's slippage" / "this month's slippage" as a
// genuine, authoritative report -- not a whole-book snapshot needing a
// date-window filter applied on this end) -- so every row is shown as-is,
// branch-filtered only, same shape as Daily PNPA (parsePnpaRows/PC), just
// published as its own file (data/pnpa-weekly.json/data/pnpa-monthly.json).
function pnpaAggregatePeriod(rows, branch){
  const totals = pnpaEmptySlipTotals(), list = [];
  for(const r of (rows||[])){
    if(branch && r[PC.BRANCH]!==branch) continue;
    pnpaAddRowToSlipTotals(totals, r);
    list.push(r);
  }
  list.sort((a,b)=>b[PC.OS]-a[PC.OS]);
  return { totals, list };
}
// Resolves the logged-in Sol ID to THIS dataset's own branch-name spelling,
// same KCCOV_BRANCH_SOL-based technique KCC Overdue's branch lock already
// uses (PC.BRANCH strings come from the same kind of HO export as
// KC.BRANCH, not guaranteed to match C.SOL_DESC's spelling exactly).
let __pnpaLockedBranch = null;
function pnpaLoggedInBranchName(rows){
  if(__pnpaLockedBranch !== null) return __pnpaLockedBranch;
  const solId = loggedInSolId();
  if(!solId){ __pnpaLockedBranch = ''; return __pnpaLockedBranch; }
  const allBranches = [...new Set(rows.map(r=>r[PC.BRANCH]))];
  __pnpaLockedBranch = allBranches.find(b=>String(KCCOV_BRANCH_SOL[String(b).toUpperCase()])===String(solId)) || '';
  return __pnpaLockedBranch;
}

/* ---------- Remark: device-local only for now (Alok's own explicit
   instruction, 2026-09-25 -- "abhi provision karo, mere paas Synology ka
   home server hai, future mein use configure karenge") -- a real synced
   backend is planned but not built yet. Stored in this browser's own
   localStorage, keyed by Account No., same non-published pattern as this
   app's existing OTS Amount/Interest Reversal overrides (js/app.js,
   'upgb-ots-amounts'/'upgb-uri-overrides') -- explicitly NOT the Special
   Note pattern (which publishes and is shared with every viewer), since
   there is no publish pipeline on this read-only portal at all. Framed
   plainly in the UI as device-only so a branch manager doesn't assume a
   colleague on another device/phone will see the same remark. ---------- */
const PNPA_REMARK_KEY = 'upgb-pnpa-remarks';
function getPnpaRemarks(){
  try{ return JSON.parse(localStorage.getItem(PNPA_REMARK_KEY) || '{}'); }catch(e){ return {}; }
}
function savePnpaRemark(acctNo, text){
  try{
    const map = getPnpaRemarks();
    if(text && text.trim()) map[acctNo] = text.trim(); else delete map[acctNo];
    localStorage.setItem(PNPA_REMARK_KEY, JSON.stringify(map));
  }catch(e){ /* private mode / quota -- remark is a convenience, not critical */ }
}
window.savePnpaRemark = function(acctNo, inputEl){
  savePnpaRemark(acctNo, inputEl.value);
  const status = inputEl.closest('tr')?.querySelector('.pnpa-remark-status');
  if(status){ status.textContent = 'Saved on this device'; status.classList.add('show'); setTimeout(()=>status.classList.remove('show'), 1600); }
};

let pnpaSlipTab = 'today';
let pnpaAddressFilter = '';
function setPnpaSlipTab(tab){ pnpaSlipTab = tab; renderPnpaSlipView(); }
window.setPnpaSlipTab = setPnpaSlipTab;

// PNPA rows carry no per-row address of their own -- resolved via the
// shared addressForAcctNo() (see its own comment near normId()), which
// looks at the main NPA book's own already-merged addresses first.
function pnpaAddressFor(acctNo){
  return addressForAcctNo(acctNo);
}
// Sortable via the same generic applySort()/nextSort()/updateSortIcons()
// engine Dashboard's All Accounts table and every other sortable table in
// this app already share (Alok, 2026-09-25: "table main sort filters lagao
// har heading par" -- every heading, including Remark). Raw PC-indexed
// rows are mapped into plain {acctNo,name,...} objects only for sorting/
// rendering purposes -- applySort() does a[key] property lookups, which a
// raw array index can't do directly.
let pnpaSlipSort = {key:'os', dir:'desc'};
function sortPnpaSlipBy(key){ pnpaSlipSort = nextSort(pnpaSlipSort, key); renderPnpaSlipView(); }
window.sortPnpaSlipBy = sortPnpaSlipBy;
function renderPnpaSlipTable(list, emptyMessage){
  if(pnpaAddressFilter){
    const q = pnpaAddressFilter.trim().toLowerCase();
    list = list.filter(r=>pnpaAddressFor(r[PC.ACCT]).toLowerCase().includes(q));
  }
  if(!list.length) return `<div class="empty-state"><p>${esc(emptyMessage || 'No accounts slipped in this period.')}</p></div>`;
  const remarks = getPnpaRemarks();
  const objs = list.map(r=>({
    acctNo:r[PC.ACCT], name:r[PC.NAME], address:pnpaAddressFor(r[PC.ACCT]), os:r[PC.OS], cadu:r[PC.CADU],
    custNpaDate:r[PC.CUSTNPADATE], remark: remarks[r[PC.ACCT]] || '',
  }));
  const sorted = applySort(objs, pnpaSlipSort);
  const rowsHtml = sorted.map(o=>{
    const acct = esc(o.acctNo);
    const remark = esc(o.remark);
    return `<tr>
      <td class="clickable" onclick="showQuickAcctDetailByAcct('pnpa','${acct}')">${acct}</td>
      <td class="tal clickable" onclick="showQuickAcctDetailByAcct('pnpa','${acct}')">${esc(o.name)||'—'}</td>
      <td class="tal">${esc(o.address)||'—'}</td>
      <td>${fmtINR2(o.os)}</td>
      <td>${fmtINR2(o.cadu)}</td>
      <td class="tal">${esc(o.custNpaDate)||'—'}</td>
      <td class="tal">
        <input type="text" class="pnpa-remark-input" value="${remark}" placeholder="Remark…" onchange="savePnpaRemark('${acct}', this)">
        <span class="pnpa-remark-status">Saved on this device</span>
      </td>
    </tr>`;
  }).join('');
  return `<div class="dash-table-wrap acct-list-scroll"><table class="dash-table pnpa-slip-table">
    <thead id="pnpaSlipTableHead"><tr>
      <th class="sortable" data-key="acctNo" tabindex="0" role="button" aria-sort="none" onclick="sortPnpaSlipBy('acctNo')">Account<span class="sort-ic">▾</span></th>
      <th class="tal sortable" data-key="name" tabindex="0" role="button" aria-sort="none" onclick="sortPnpaSlipBy('name')">Name<span class="sort-ic">▾</span></th>
      <th class="tal sortable" data-key="address" tabindex="0" role="button" aria-sort="none" onclick="sortPnpaSlipBy('address')">Address<span class="sort-ic">▾</span></th>
      <th class="sortable" data-key="os" tabindex="0" role="button" aria-sort="none" onclick="sortPnpaSlipBy('os')">Balance<span class="sort-ic">▾</span></th>
      <th class="sortable" data-key="cadu" tabindex="0" role="button" aria-sort="none" onclick="sortPnpaSlipBy('cadu')">CADU<span class="sort-ic">▾</span></th>
      <th class="tal sortable" data-key="custNpaDate" tabindex="0" role="button" aria-sort="none" onclick="sortPnpaSlipBy('custNpaDate')">Cust NPA Date<span class="sort-ic">▾</span></th>
      <th class="tal sortable" data-key="remark" tabindex="0" role="button" aria-sort="none" onclick="sortPnpaSlipBy('remark')">Remark <span class="pnpa-remark-note">(saved on this device only)</span><span class="sort-ic">▾</span></th>
    </tr></thead>
    <tbody>${rowsHtml}</tbody>
  </table></div>`;
}
function pnpaSlipSummaryChips(totals){
  return `<div class="pnpa-slip-summary">
    <div class="pnpa-slip-chip"><span class="lbl">KCC</span><span class="cnt">${totals.kcc.cnt.toLocaleString('en-IN')} A/C</span><span class="amt">${fmtCr(totals.kcc.amt)}</span></div>
    <div class="pnpa-slip-chip"><span class="lbl">Non-KCC</span><span class="cnt">${totals.nonkcc.cnt.toLocaleString('en-IN')} A/C</span><span class="amt">${fmtCr(totals.nonkcc.amt)}</span></div>
    <div class="pnpa-slip-chip total"><span class="lbl">Total</span><span class="cnt">${totals.all.cnt.toLocaleString('en-IN')} A/C</span><span class="amt">${fmtCr(totals.all.amt)}</span></div>
  </div>`;
}
/* Alok's own explicit instruction (2026-09-25): lead the PNPA Slippage
   page with 3 hero blocks for TODAY specifically -- Total Slippage, KCC
   Slippage, and "Non-KCC & Technical" -- ahead of the Today/This Week/
   This Month tabs below, since today's fresh slippage is the most
   actionable figure for a branch manager. "Technical" is a category he
   has not defined yet ("use baad mein bataunga") -- this block currently
   shows the same figure as the existing Non-KCC total until he defines
   what should be split out of it; labelled honestly below rather than
   inventing a filter for a category with no definition yet. Clicking any
   card jumps to the Today tab, which already lists these exact accounts. */
function pnpaTodayHeroBlocks(todayTotals){
  return `<div class="pnpa-today-hero">
    <div class="pnpa-today-hero-head">Today's Slippage</div>
    <div class="pnpa-today-hero-row">
      <div class="pnpa-today-card total clickable" onclick="setPnpaSlipTab('today')">
        <span class="lbl">Total Slippage</span>
        <span class="cnt">${todayTotals.all.cnt.toLocaleString('en-IN')} A/C</span>
        <span class="amt">${fmtCr(todayTotals.all.amt)}</span>
      </div>
      <div class="pnpa-today-card kcc clickable" onclick="setPnpaSlipTab('today')">
        <span class="lbl">KCC Slippage</span>
        <span class="cnt">${todayTotals.kcc.cnt.toLocaleString('en-IN')} A/C</span>
        <span class="amt">${fmtCr(todayTotals.kcc.amt)}</span>
      </div>
      <div class="pnpa-today-card nonkcc clickable" onclick="setPnpaSlipTab('today')">
        <span class="lbl">Non-KCC &amp; Technical</span>
        <span class="cnt">${todayTotals.nonkcc.cnt.toLocaleString('en-IN')} A/C</span>
        <span class="amt">${fmtCr(todayTotals.nonkcc.amt)}</span>
        <span class="note">Technical breakdown to be added</span>
      </div>
    </div>
  </div>`;
}
function renderPnpaSlipView(){
  const el = document.getElementById('pnpaSlipArea');
  if(!el) return;
  if(!PNPA_DATA){
    el.innerHTML = `<div class="empty-state"><div class="data-loading-spinner" aria-hidden="true" style="position:static;border-color:rgba(58,123,255,.25);border-top-color:var(--accent)"></div><p style="margin-top:14px">Loading PNPA data…</p></div>`;
    ensurePnpaDataLoaded(renderPnpaSlipView, () => {
      el.innerHTML = `<div class="empty-state"><h2>Could not load PNPA data</h2><p>Check your internet connection, then tap Refresh.</p></div>`;
    });
    return;
  }
  // Weekly/Monthly are optional -- a load attempt (success or failure) is
  // required before rendering so the tab counts/table don't flash "no
  // data" and then correct themselves a moment later, but a failure (most
  // likely: the file simply hasn't been uploaded yet, a brand-new upload
  // type) never blocks the whole page -- only that one tab shows its own
  // clear empty state below, "Today" (Daily PNPA) keeps working regardless.
  if(!PNPA_WEEKLY_DATA && !__pnpaWeeklyFailed){ ensurePnpaWeeklyDataLoaded(renderPnpaSlipView); return; }
  if(!PNPA_MONTHLY_DATA && !__pnpaMonthlyFailed){ ensurePnpaMonthlyDataLoaded(renderPnpaSlipView); return; }
  const branchName = pnpaLoggedInBranchName(PNPA_DATA.rows);
  const today = new Date();
  const agg = {
    today: pnpaAggregateToday(PNPA_DATA.rows, branchName, today),
    week: pnpaAggregatePeriod(PNPA_WEEKLY_DATA ? PNPA_WEEKLY_DATA.rows : [], branchName),
    month: pnpaAggregatePeriod(PNPA_MONTHLY_DATA ? PNPA_MONTHLY_DATA.rows : [], branchName),
  };
  const tabsHtml = `<div class="bank-tab-row">${PNPA_SLIP_TABS.map(t=>
    `<button type="button" class="bank-tab-btn${pnpaSlipTab===t.key?' active':''}" onclick="setPnpaSlipTab('${t.key}')">${t.label} <span class="pnpa-tab-count">${agg[t.key].totals.all.cnt}</span></button>`
  ).join('')}</div>`;
  const active = agg[pnpaSlipTab];
  const emptyMessages = {
    today: 'No accounts slipped today.',
    week: __pnpaWeeklyFailed ? 'No Weekly PNPA file uploaded yet.' : 'No accounts in this week\'s slippage file.',
    month: __pnpaMonthlyFailed ? 'No Monthly PNPA file uploaded yet.' : 'No accounts in this month\'s slippage file.',
  };
  el.innerHTML = `
    ${pnpaTodayHeroBlocks(agg.today.totals)}
    ${tabsHtml}
    <div class="bank-filter-row">
      <input type="text" id="pnpaAddressFilterInput" class="dash-select" placeholder="Filter by Address…" value="${esc(pnpaAddressFilter)}" style="max-width:220px">
    </div>
    ${pnpaSlipSummaryChips(active.totals)}
    ${renderPnpaSlipTable(active.list, emptyMessages[pnpaSlipTab])}
  `;
  const addrInput = document.getElementById('pnpaAddressFilterInput');
  if(addrInput) addrInput.onchange = () => { pnpaAddressFilter = addrInput.value; renderPnpaSlipView(); };
  updateSortIcons('pnpaSlipTableHead', pnpaSlipSort);
}
window.renderPnpaSlipView = renderPnpaSlipView;

/* ---------- KCC Overdue -- Hathras-only, restricted to 3 schemes, rich filters ----------
   Unlike PNPA, the source "KCC Overdue" file is already Hathras-scoped (confirmed
   against a real file: all rows were Region=HATHRAS), so no whole-bank filtering is
   needed -- but the parser still defensively drops any stray non-Hathras row in
   case a future export widens scope. Only rows matching one of the 3 known scheme
   codes are kept; there is no "Other" catch-all bucket here (unlike PNPA). */
const KC = {BRANCH:0, SCHEME:1, ACCT:2, NAME:3, OS:4, CADU:5, LIMIT:6, REVIEW:7, CUSTNPADATE:8, FY:9, CATEGORY:10, SMA:11, REASON:12};
/* KCC Overdue rows carry only the branch name string (uppercase, e.g.
   "HATHRAS AGRA ROAD") -- match it back to the frozen BRANCH_LIST to show
   Sol ID alongside it in the Datewise Calendar view. */
const KCCOV_BRANCH_SOL = Object.fromEntries(BRANCH_LIST.map(([,newId,name])=>[name.toUpperCase(), newId]));
const KCC_OVERDUE_SCHEMES = [
  {key:'kcc', code:'CC004', label:'KCC'},
  {key:'kccah', code:'CC043', label:'KCC — Animal Husbandry'},
  {key:'od023', code:'OD023', label:'OD-023 (Tatkal)'},
];
function kccOverdueBucketOf(scheme){
  const m = KCC_OVERDUE_SCHEMES.find(s=>s.code===scheme);
  return m ? m.key : null;
}
/* The source file's F.Y. column stores its value with literal double-quote
   characters around it (e.g. the cell's actual text is ["MAR-27"], not just
   MAR-27) -- almost certainly the HO export's own guard against Excel trying
   to auto-parse "MAR-27" as a date. Stripped for display/filtering. */
function stripQuoteChars(s){ return String(s||'').replace(/^"+|"+$/g,'').trim(); }
function parseKccOverdueRows(headerCells, dataRows){
  const header = headerCells.map(normHeader);
  const idx = (name) => header.indexOf(normHeader(name));
  const idxPrefix = (name) => header.findIndex(h=>h.startsWith(normHeader(name)));
  const iRegion=idx('region'), iBranch=idx('branch'), iAcct=idx('accountno'), iScheme=idx('schemecode'),
    iName=idx('accountname'), iBal=idxPrefix('balanceamount'), iCadu=idx('cadu'), iLimit=idx('limit'),
    iReview=idx('reviewdate'), iCustNpa=idx('custnpadate'), iFy=idx('fy'), iCategory=idx('category'),
    iSma=idx('smastatus'), iReason=idx('reasons');
  const missing = [];
  if(iAcct<0) missing.push('Account No');
  if(iBranch<0) missing.push('Branch');
  if(iScheme<0) missing.push('Scheme Code');
  if(iBal<0) missing.push('Balance Amount');
  if(iCustNpa<0) missing.push('Cust NPA Date');
  if(missing.length) throw new Error('Missing required column(s): '+missing.join(', ')+'. Check this file matches the daily KCC "Data" export layout.');
  const rows = [];
  for(const row of dataRows){
    if(!row || row.length<3) continue;
    if(iRegion>=0){ const region = cellStr(row, iRegion); if(region && region.toUpperCase()!=='HATHRAS') continue; }
    const scheme = cellStr(row, iScheme);
    if(!kccOverdueBucketOf(scheme)) continue;
    const acctRaw = cellStr(row, iAcct);
    if(!acctRaw) continue;
    let acctNo = acctRaw;
    if(looksScientific(acctRaw)) acctNo = expandSci(acctRaw);
    const reviewDt = toDate(iReview>=0?row[iReview]:'');
    const custNpaDt = toDate(row[iCustNpa]);
    rows.push([
      cellStr(row, iBranch), scheme, acctNo, cellStr(row, iName),
      parseFloat(row[iBal])||0, iCadu>=0?(parseFloat(row[iCadu])||0):0,
      iLimit>=0?(parseFloat(row[iLimit])||0):0,
      reviewDt ? fmtDate(reviewDt) : '',
      custNpaDt ? fmtDate(custNpaDt) : '',
      iFy>=0 ? stripQuoteChars(cellStr(row, iFy)) : '',
      iCategory>=0 ? cellStr(row, iCategory) : '',
      iSma>=0 ? cellStr(row, iSma) : '',
      iReason>=0 ? cellStr(row, iReason) : '',
    ]);
  }
  return rows;
}
let KCC_OVERDUE_DATA = null;
let __pendingKccOverdueData = null;
let kccovSchemeTab = 'kcc';
let kccovBranchFilter = '';
let kccovFyFilter = '';
let kccovDateMode = 'month';
let kccovMonthFilter = '';
let kccovDateFrom = '';
let kccovDateTo = '';
let kccovAddressFilter = '';
let kccovView = 'summary'; // 'summary' | 'calendar' | 'fymonth'
let kccovBranchSort = {key:'os', dir:'desc'};
function sortKccovBranchBy(key){ kccovBranchSort = nextSort(kccovBranchSort, key); renderKccOverdueBody(); }
window.sortKccovBranchBy = sortKccovBranchBy;
// Datewise Calendar's own column sort -- {key:'sol'|'branch'|'total'|<date
// string>, dir:'asc'|'desc'}. Starts null so renderKccOverdueCalendar()
// can tell "never sorted yet" apart from "user explicitly re-sorted" and
// pick today's date column as the one-time default only in the former case.
let kccovCalSort = null;
function setKccovSchemeTab(tab){ kccovSchemeTab = tab; renderKccOverdueBody(); }
window.setKccovSchemeTab = setKccovSchemeTab;
function setKccovDateMode(mode){ kccovDateMode = mode; renderKccOverdueBody(); }
window.setKccovDateMode = setKccovDateMode;
function setKccovView(v){
  kccovView = v;
  /* Datewise Calendar renders one column per distinct Cust NPA Date --
     with no month/range filter picked yet, kccovFilteredRows() lets every
     date in the whole upload through (which spans years, since Cust NPA
     Date is a forward-looking projected-classification date, not just
     "this month"), producing an unusably wide table. Default the filter
     to the upload's own as-on month the first time Calendar is opened, so
     it always starts scoped to something sane; leave it alone once the
     user has picked their own month/range. */
  if(v==='calendar' && KCC_OVERDUE_DATA && KCC_OVERDUE_DATA.asOnDate){
    const ym = KCC_OVERDUE_DATA.asOnDate.slice(0,7);
    if(kccovDateMode==='month' && !kccovMonthFilter) kccovMonthFilter = ym;
    if(kccovDateMode==='range' && !kccovDateFrom && !kccovDateTo){
      const [y,m] = ym.split('-').map(Number);
      kccovDateFrom = dateToInputValue(new Date(y, m-1, 1));
      kccovDateTo = dateToInputValue(new Date(y, m, 0));
    }
  }
  renderKccOverdueBody();
}
window.setKccovView = setKccovView;

async function handleKccOverdueUpload(evt){
  const file = evt.target.files[0];
  if(!file) return;
  await ensureXLSX();
  const labelEl = document.getElementById('kccOverdueUploadDropLabel');
  if(labelEl) labelEl.textContent = file.name;
  const statusEl = document.getElementById('kccOverdueUploadStatus');
  statusEl.innerHTML = `<div class="upload-status info">Reading KCC Overdue file…</div>`;
  const isCsv = /\.csv$/i.test(file.name);
  const reader = new FileReader();
  reader.onerror = function(){ statusEl.innerHTML = `<div class="upload-status err">⚠ Failed to read the file from disk.</div>`; };
  reader.onload = function(e){
    try{
      let header, dataRows;
      if(isCsv){
        const allRows = parseCSV(String(e.target.result));
        header = allRows[0]||[]; dataRows = allRows.slice(1);
      } else {
        const data = new Uint8Array(e.target.result);
        // No cellDates:true (Alok's audit, 2026-09-08): it silently
        // returned Cust NPA Date one full day early, every time, in IST
        // -- see the comment on parseHoDate() for the exact mechanism.
        // Every date column below goes through toDate(), which already
        // converts the raw Excel serial number correctly on its own.
        const wb = XLSX.read(data, {type:'array'});
        const sheetName = wb.SheetNames.find(n=>/kcc|overdue|^data$/i.test(n)) || wb.SheetNames[0];
        const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {header:1, raw:true, defval:''});
        header = raw[0]||[]; dataRows = raw.slice(1);
      }
      const rows = parseKccOverdueRows(header, dataRows);
      if(!rows.length) throw new Error('No account rows found in this file.');
      const guessed = parseAsOnDateFromFilename(file.name);
      const asOnDate = guessed ? dateToInputValue(guessed) : dateToInputValue(new Date());
      __pendingKccOverdueData = { asOnDate, rows };
      KCC_OVERDUE_DATA = __pendingKccOverdueData;
      const label = document.getElementById('kccovStatusLabel');
      if(label) label.textContent = `${rows.length.toLocaleString('en-IN')} accounts loaded (${file.name})`;
      statusEl.innerHTML = `<div class="upload-status ok">✔ Parsed ${rows.length.toLocaleString('en-IN')} accounts, as on ${esc(asOnDate)}. Goes live the next time you hit Publish.</div>`;
      clearStalePublishStatus();
      const publishBtn = document.getElementById('publishBtn');
      if(publishBtn) publishBtn.disabled = false;
      if(document.querySelector('.view.active')?.dataset.view==='kccov') renderKccOverdueBody();
    } catch(err){
      statusEl.innerHTML = `<div class="upload-status err">⚠ Could not read this file: ${esc(err.message||err)}</div>`;
    }
  };
  if(isCsv) reader.readAsText(file); else reader.readAsArrayBuffer(file);
}

function renderKccOverdue(){
  const el = document.getElementById('kccOverdueArea');
  if(!el) return;
  if(KCC_OVERDUE_DATA){ renderKccOverdueBody(); return; }
  el.innerHTML = `<div class="empty-state"><div class="data-loading-spinner" aria-hidden="true" style="position:static;border-color:rgba(58,123,255,.25);border-top-color:var(--accent)"></div><p style="margin-top:14px">Loading KCC Overdue data…</p></div>`;
  fetchJson(DATA_ORIGIN + 'data/kcc-overdue.json?t=' + Date.now())
    .then(d => { KCC_OVERDUE_DATA = d; renderKccOverdueBody(); })
    .catch(() => {
      el.innerHTML = `<div class="empty-state"><h2>Could not load KCC Overdue data</h2><p>Check your internet connection, then tap Refresh.</p></div>`;
    });
}

// KCC Overdue's own rows carry no per-row address of their own -- resolved
// via the shared addressForAcctNo() (see its own comment near normId()),
// which looks at the main NPA book's own already-merged addresses first.
function kccovAddressFor(acctNo){
  return addressForAcctNo(acctNo);
}
function kccovFilteredRows(d){
  let rows = d.rows;
  if(kccovBranchFilter) rows = rows.filter(r=>r[KC.BRANCH]===kccovBranchFilter);
  if(kccovFyFilter) rows = rows.filter(r=>r[KC.FY]===kccovFyFilter);
  if(kccovAddressFilter){
    const q = kccovAddressFilter.trim().toLowerCase();
    rows = rows.filter(r=>kccovAddressFor(r[KC.ACCT]).toLowerCase().includes(q));
  }
  if(kccovDateMode==='month' && kccovMonthFilter){
    const [y,m] = kccovMonthFilter.split('-').map(Number);
    rows = rows.filter(r=>{ const dt = toDate(r[KC.CUSTNPADATE]); return dt && dt.getFullYear()===y && (dt.getMonth()+1)===m; });
  } else if(kccovDateMode==='range' && (kccovDateFrom || kccovDateTo)){
    const from = kccovDateFrom ? new Date(kccovDateFrom+'T00:00:00') : null;
    const to = kccovDateTo ? new Date(kccovDateTo+'T23:59:59') : null;
    rows = rows.filter(r=>{
      const dt = toDate(r[KC.CUSTNPADATE]);
      if(!dt) return false;
      if(from && dt < from) return false;
      if(to && dt > to) return false;
      return true;
    });
  }
  return rows;
}
function renderKccOverdueBody(){
  const el = document.getElementById('kccOverdueArea');
  const d = KCC_OVERDUE_DATA;
  if(!el) return;
  if(!d || !d.rows){ el.innerHTML = `<div class="empty-state"><h2>No KCC Overdue data yet</h2><p>Upload the KCC Overdue file from Update Data to populate this tab.</p></div>`; return; }

  document.querySelectorAll('.kccov-report-date-val').forEach(e=>{
    const parts = (d.asOnDate||'').split('-');
    e.textContent = parts.length===3 ? `${parts[2]}-${parts[1]}-${parts[0]}` : (d.asOnDate||'—');
  });

  const allBranches = [...new Set(d.rows.map(r=>r[KC.BRANCH]))].sort((a,b)=>a.localeCompare(b));
  const allFy = [...new Set(d.rows.map(r=>r[KC.FY]).filter(Boolean))].sort();
  // Recovery Dashboard (branch portal): resolve the logged-in Sol ID to
  // THIS dataset's own branch-name spelling via the existing
  // KCCOV_BRANCH_SOL map (built off BRANCH_LIST) -- not assumed to match
  // C.SOL_DESC's spelling on the NPA side, matching how the Datewise
  // Calendar view already resolves each KCC row's Sol ID the same way.
  // Locked (not just defaulted): force kccovBranchFilter to it every
  // render, ignoring any stray manual selection.
  const lockedSolId = loggedInSolId();
  const lockedKccBranch = lockedSolId ? allBranches.find(b=>String(KCCOV_BRANCH_SOL[String(b).toUpperCase()])===String(lockedSolId)) : null;
  if(lockedKccBranch) kccovBranchFilter = lockedKccBranch;
  const branchFilterOptions = `<option value="">Regional Office</option>` +
    allBranches.map(b=>`<option value="${esc(b)}"${kccovBranchFilter===b?' selected':''}>${esc(b)}</option>`).join('');
  const fyFilterOptions = `<option value="">All F.Y.</option>` +
    allFy.map(f=>`<option value="${esc(f)}"${kccovFyFilter===f?' selected':''}>${esc(f)}</option>`).join('');

  const dateModeRow = `<div class="bank-tab-row" style="margin-bottom:10px">
    <button type="button" class="bank-tab-btn${kccovDateMode==='month'?' active':''}" onclick="setKccovDateMode('month')">Cust NPA Date — By Month</button>
    <button type="button" class="bank-tab-btn${kccovDateMode==='range'?' active':''}" onclick="setKccovDateMode('range')">By Date Range</button>
  </div>`;
  const dateInputsRow = kccovDateMode==='month'
    ? `<input type="month" id="kccovMonthInput" class="dash-select" value="${esc(kccovMonthFilter)}" style="max-width:200px">`
    : `<input type="date" id="kccovDateFromInput" class="dash-select" value="${esc(kccovDateFrom)}" style="max-width:170px">
       <span style="color:var(--ink-mute);font-size:12px;align-self:center">to</span>
       <input type="date" id="kccovDateToInput" class="dash-select" value="${esc(kccovDateTo)}" style="max-width:170px">`;

  const toolbar = `<div class="dash-toolbar">
      <span class="dash-toolbar-label">Branch</span>
      ${lockedKccBranch
        ? `<span class="recovery-branch-lock">${esc(lockedKccBranch)}</span>`
        : `<select id="kccovBranchFilterSelect" class="dash-select">${branchFilterOptions}</select>`}
    </div>
    <div class="bank-filter-row">
      <select id="kccovFyFilterSelect" class="dash-select">${fyFilterOptions}</select>
      <input type="text" id="kccovAddressFilterInput" class="dash-select" placeholder="Filter by Address…" value="${esc(kccovAddressFilter)}" style="max-width:220px">
    </div>
    ${dateModeRow}
    <div class="bank-filter-row">${dateInputsRow}</div>`;

  const filteredRows = kccovFilteredRows(d);
  // The hero scheme-tab row and its bucketTotals are only meaningful for
  // Branch Summary/Calendar -- the 3 bifurcation views (F.Y./Month
  // Summary, Branch Report, All Branches Overview) exist specifically to
  // show KCC/KCC-AH/OD-023 side by side, so a still-highlighted "active
  // scheme" card above a table that visibly spans all 3 would read as
  // contradictory. Skipped entirely (not just hidden) when one of those
  // views is active, rather than computed and thrown away.
  const showHero = kccovView==='summary' || kccovView==='calendar';
  let heroRow = '';
  if(showHero){
    const bucketTotals = {};
    KCC_OVERDUE_SCHEMES.forEach(s=>{ bucketTotals[s.key]={count:0,os:0,branches:new Set()}; });
    for(const r of filteredRows){
      const bk = kccOverdueBucketOf(r[KC.SCHEME]);
      bucketTotals[bk].count++; bucketTotals[bk].os += r[KC.OS]; bucketTotals[bk].branches.add(r[KC.BRANCH]);
    }
    const bucketIcon = {kcc:ICON_TARGET, kccah:ICON_STAR, od023:ICON_ALERT_TRIANGLE};
    heroRow = `<div class="hero-kpi-row bank-hero-row">${KCC_OVERDUE_SCHEMES.map(s=>{
      const t = bucketTotals[s.key], isActive = kccovSchemeTab===s.key;
      return heroKpiCard({
        id:'kccovHero_'+s.key, icon: bucketIcon[s.key],
        tint: isActive?'var(--accent-soft)':'rgba(120,120,140,.12)', color: isActive?'var(--accent)':'var(--ink-mute)',
        onclick:`setKccovSchemeTab('${s.key}')`,
        label: s.label,
        fallback: fmtCr(t.os),
        sub: `${t.count.toLocaleString('en-IN')} accounts · ${t.branches.size.toLocaleString('en-IN')} branches`,
        badge: isActive ? `<div class="hero-kpi-badge" style="background:var(--accent-soft);color:var(--accent)">Viewing</div>` : '',
      });
    }).join('')}</div>`;
  }

  // 3 view modes in one row: Branch Summary/Calendar (scoped to one scheme
  // via the hero row above) plus F.Y./Month Summary (always showing KCC/
  // KCC-AH/OD-023 side by side) -- a thin .bank-tab-sep divider marks the
  // boundary between the two groups without a second tab row, same
  // flex-wrap this row already relies on for narrow screens. Recovery
  // Dashboard (branch portal): "Branch Report" and "All Branches Overview"
  // were dropped entirely (not just hidden) -- both exist on the main
  // npadashboard site to let an HO user pick/compare branches, which is
  // meaningless here since this portal is permanently locked to one
  // branch already (Alok, 2026-09-25: "all branch overview ki need hai
  // kya... branch report dono ko hata do").
  const viewToggleRow = `<div class="bank-tab-row" style="margin-top:18px">
    <button type="button" class="bank-tab-btn${kccovView==='summary'?' active':''}" onclick="setKccovView('summary')">Branch Summary</button>
    <button type="button" class="bank-tab-btn${kccovView==='calendar'?' active':''}" onclick="setKccovView('calendar')">Datewise Calendar</button>
    <span class="bank-tab-sep" aria-hidden="true"></span>
    <button type="button" class="bank-tab-btn${kccovView==='fymonth'?' active':''}" onclick="setKccovView('fymonth')">F.Y./Month Summary</button>
  </div>`;

  const isBifurcationView = kccovView==='fymonth';
  const showPrintPdf = kccovView==='fymonth';
  const actionButtons = kccovView==='summary'
    ? `<button type="button" class="export-xl-btn" onclick="exportKccOverdueSummary()">${EXPORT_XL_ICON} Export to Excel</button>`
    : isBifurcationView
      ? `<button type="button" class="export-xl-btn" onclick="exportKccOverdueBifurcation()">${EXPORT_XL_ICON} Export to Excel</button>`
        + (showPrintPdf ? `<button type="button" class="export-xl-btn" onclick="printKccOverdueBifurcation()">Print</button>
           <button type="button" class="export-xl-btn" onclick="exportKccOverdueBifurcationPdf('${kccovView}')">Save as PDF</button>` : '')
      : '';

  el.innerHTML = toolbar + heroRow + viewToggleRow +
    (kccovView==='calendar' ? `<div id="kccovInsightWrap"></div>` : '') +
    `<div class="chart-card" style="margin-top:16px">
      <div class="chart-card-head-row">
        <div class="section-label" id="kccovTableLabel"></div>
        ${actionButtons}
      </div>
      ${kccovView==='calendar' ? `<div class="kccov-cal-legend" id="kccovCalLegend"></div>` : ''}
      <div id="kccovBranchTableCard"></div>
    </div>`;

  const branchSel = document.getElementById('kccovBranchFilterSelect');
  if(branchSel) branchSel.onchange = () => { kccovBranchFilter = branchSel.value; renderKccOverdueBody(); };
  const fySel = document.getElementById('kccovFyFilterSelect');
  if(fySel) fySel.onchange = () => { kccovFyFilter = fySel.value; renderKccOverdueBody(); };
  const addrInput = document.getElementById('kccovAddressFilterInput');
  if(addrInput) addrInput.onchange = () => { kccovAddressFilter = addrInput.value; renderKccOverdueBody(); };
  const monthInput = document.getElementById('kccovMonthInput');
  if(monthInput) monthInput.onchange = () => { kccovMonthFilter = monthInput.value; renderKccOverdueBody(); };
  const fromInput = document.getElementById('kccovDateFromInput');
  if(fromInput) fromInput.onchange = () => { kccovDateFrom = fromInput.value; renderKccOverdueBody(); };
  const toInput = document.getElementById('kccovDateToInput');
  if(toInput) toInput.onchange = () => { kccovDateTo = toInput.value; renderKccOverdueBody(); };

  if(kccovView==='calendar') renderKccOverdueCalendar(filteredRows);
  else if(kccovView==='fymonth') renderKccOverdueFyMonth(filteredRows);
  else renderKccOverdueBranchTable(filteredRows);
}

let kccovLastExport = null;
// Recovery Dashboard (branch portal): this used to be a branch-wise
// ranking table (Rank/Branch/Accounts/Total O/S) -- but with the whole
// tab already locked to one branch, that table always had exactly one
// row, which just restated the hero cards above it. Alok, 2026-09-25:
// "branch summary k blocks hain to niche summary table ki need nahi hai
// yahan account wise list de do with filtered" -- the hero blocks already
// are the summary, so this now lists the underlying accounts directly
// (scoped to whichever scheme tab is active), with Address in place of
// the no-longer-meaningful Branch column and its own Address filter.
function renderKccOverdueBranchTable(filteredRows){
  const wrap = document.getElementById('kccovBranchTableCard');
  const labelEl = document.getElementById('kccovTableLabel');
  if(!wrap) return;
  const activeScheme = KCC_OVERDUE_SCHEMES.find(s=>s.key===kccovSchemeTab);
  const unsorted = filteredRows.filter(r=>kccOverdueBucketOf(r[KC.SCHEME])===kccovSchemeTab).map(r=>({
    acctNo:r[KC.ACCT], name:r[KC.NAME], address:kccovAddressFor(r[KC.ACCT]), os:r[KC.OS], cadu:r[KC.CADU], limit:r[KC.LIMIT],
    custNpaDate:r[KC.CUSTNPADATE], fy:r[KC.FY], category:r[KC.CATEGORY], sma:r[KC.SMA],
  }));
  const list = applySort(unsorted, kccovBranchSort);
  const scopeLabel = kccovBranchFilter ? esc(kccovBranchFilter) : 'Regional Office (all branches)';
  kccovLastExport = { list, schemeLabel: activeScheme.label, schemeCode: activeScheme.code, scopeLabel };
  if(labelEl) labelEl.innerHTML = `${esc(activeScheme.label)} — Account-wise list, highest O/S first<span class="chart-sub">Scheme ${esc(activeScheme.code)} · ${scopeLabel} · ${list.length.toLocaleString('en-IN')} account(s) shown · tap an account for details</span>`;
  const rowsHtml = list.map((r,i)=>`<tr class="clickable" onclick="showQuickAcctDetailByAcct('kccov','${esc(r.acctNo)}')">
    <td><span class="dash-rank">${i+1}</span>${esc(r.acctNo)}</td>
    <td class="tal">${esc(r.name)||'—'}</td>
    <td class="tal">${esc(r.address)||'—'}</td>
    <td>${fmtINR2(r.os)}</td>
    <td>${fmtINR2(r.cadu)}</td>
    <td class="tal">${esc(r.custNpaDate)||'—'}</td>
  </tr>`).join('');
  wrap.innerHTML = `<div class="dash-table-wrap acct-list-scroll">
    <table class="dash-table">
      <thead id="kccovBranchTableHead"><tr>
        <th class="tal sortable" data-key="acctNo" tabindex="0" role="button" aria-sort="none" onclick="sortKccovBranchBy('acctNo')">Account<span class="sort-ic">▾</span></th>
        <th class="tal sortable" data-key="name" tabindex="0" role="button" aria-sort="none" onclick="sortKccovBranchBy('name')">Customer<span class="sort-ic">▾</span></th>
        <th class="tal sortable" data-key="address" tabindex="0" role="button" aria-sort="none" onclick="sortKccovBranchBy('address')">Address<span class="sort-ic">▾</span></th>
        <th class="sortable" data-key="os" tabindex="0" role="button" aria-sort="none" onclick="sortKccovBranchBy('os')">O/S<span class="sort-ic">▾</span></th>
        <th class="sortable" data-key="cadu" tabindex="0" role="button" aria-sort="none" onclick="sortKccovBranchBy('cadu')">CADU<span class="sort-ic">▾</span></th>
        <th class="tal sortable" data-key="custNpaDate" tabindex="0" role="button" aria-sort="none" onclick="sortKccovBranchBy('custNpaDate')">Cust NPA Date<span class="sort-ic">▾</span></th>
      </tr></thead>
      <tbody>${rowsHtml || emptyStateRowHtml(6, 'No accounts match this filter')}</tbody>
    </table>
  </div>`;
  updateSortIcons('kccovBranchTableHead', kccovBranchSort);
}
/* Exports exactly the account list currently on screen (same scheme tab /
   branch / F.Y. / address / date filters) -- kept WYSIWYG so the file
   never surprises whoever downloads it with a different set of rows than
   what they were looking at. */
function exportKccOverdueSummary(){
  const x = kccovLastExport;
  if(!x || !x.list.length) return;
  const rows = x.list.map(r=>[r.acctNo, r.name, r.address, r.os, r.cadu, r.custNpaDate]);
  exportRowsToExcel(
    `KCC_Overdue_${x.schemeCode}_${dateToInputValue(new Date())}.xlsx`, 'KCC Overdue Summary',
    ['Account No','Customer','Address','O/S','CADU','Cust NPA Date'], rows,
    [null,null,null,XL_INR_FMT,XL_INR_FMT,null]
  );
  showToast(`✓ ${x.branchAgg.length} branch row${x.branchAgg.length>1?'s':''} exported`);
}
window.exportKccOverdueSummary = exportKccOverdueSummary;

/* ---------- KCC Overdue Bifurcation: Excel export (Report/Data/Branch Report-with-
   dropdown/All Branches, 4 sheets) -- ported from tools/kcc-overdue-summary.html,
   using this app's own lazy ensureExcelJS() instead of that standalone tool's plain
   <script> tag. Separate from exportKccOverdueSummary() above, which stays scoped
   to the current on-screen Branch Summary table exactly as its own comment already
   documents -- this is the F.Y./Month Summary/Branch Report/All Branches Overview
   views' own export, always covering every branch (the Branch Report sheet's own
   dropdown lets a viewer pick any branch inside Excel itself, no re-export needed). */
function kccovRound2(n){ return Math.round(n*100)/100; }
function kccovThinBorder(){ return { top:{style:'thin',color:{argb:'FFD7DED9'}}, left:{style:'thin',color:{argb:'FFD7DED9'}}, bottom:{style:'thin',color:{argb:'FFD7DED9'}}, right:{style:'thin',color:{argb:'FFD7DED9'}} }; }
const KCCOV_FY_PALETTE_ARGB = ['FF2F527D','FF1F7A5C','FF8B5E1F','FF6F3D8E','FFA5432B','FF2F7D7D'];
const KCCOV_FY_PALETTE_SOFT_ARGB = ['FFDCE6F1','FFDCEEE6','FFF3E7CE','FFEBDFF3','FFF6E2DC','FFDCF0F0'];

function kccovWriteDataRow(ws, r, label, g, bold){
  ws.getCell(r,1).value = label;
  if(bold) ws.getCell(r,1).font = {bold:true};
  KCCOV_BIFURCATION_GROUPS.forEach((gd,i)=>{
    const col = 2+i*2, b = g[gd.key];
    const cCnt = ws.getCell(r,col); cCnt.value = b.cnt; cCnt.numFmt = '0';
    const cAmt = ws.getCell(r,col+1); cAmt.value = kccovRound2(b.amt/100000); cAmt.numFmt = '0.00';
  });
  if(bold){
    for(let c=1;c<=13;c++){ const cell = ws.getCell(r,c); cell.font = {bold:true}; cell.border = {top:{style:'thin'}}; cell.fill = {type:'pattern',pattern:'solid',fgColor:{argb:'FFF3E7CE'}}; }
  }
  return r+1;
}
function kccovWriteFyBand(ws, r, fy, argb){
  ws.mergeCells(r,1,r,13);
  const c = ws.getCell(r,1);
  c.value = 'F.Y. '+fy; c.font = {bold:true, color:{argb:'FFFFFFFF'}, size:12}; c.alignment = {horizontal:'center', vertical:'middle'};
  for(let col=1; col<=13; col++) ws.getCell(r,col).fill = {type:'pattern',pattern:'solid',fgColor:{argb:argb}};
  ws.getRow(r).height = 20;
  return r+1;
}
function kccovWriteTwoRowHeader(ws, r, softArgb){
  ws.mergeCells(r,1,r+1,1);
  const mc = ws.getCell(r,1); mc.value = 'Month'; mc.font = {bold:true}; mc.alignment = {vertical:'middle',horizontal:'center'};
  KCCOV_BIFURCATION_GROUPS.forEach((gd,i)=>{
    const col = 2+i*2;
    ws.mergeCells(r,col,r,col+1);
    const gc = ws.getCell(r,col); gc.value = gd.label; gc.font = {bold:true}; gc.alignment = {horizontal:'center'};
    ws.getCell(r,col).fill = {type:'pattern',pattern:'solid',fgColor:{argb:softArgb}};
    ws.getCell(r,col+1).fill = {type:'pattern',pattern:'solid',fgColor:{argb:softArgb}};
    ws.getCell(r+1,col).value = 'A/C Count'; ws.getCell(r+1,col).font = {bold:true};
    ws.getCell(r+1,col+1).value = 'Amount (₹ Lakh)'; ws.getCell(r+1,col+1).font = {bold:true};
  });
  for(let col2=1; col2<=13; col2++){ ws.getCell(r,col2).border = kccovThinBorder(); ws.getCell(r+1,col2).border = kccovThinBorder(); }
  return r+2;
}
function kccovWriteBifurcationSheet(ws, title, rows){
  ws.getColumn(1).width = 16;
  for(let c=2;c<=13;c++) ws.getColumn(c).width = 15;
  let r = 1;
  ws.mergeCells(r,1,r,13);
  ws.getCell(r,1).value = title; ws.getCell(r,1).font = {bold:true, size:14};
  r += 2;
  const byFy = kccovAggregateFyMonth(rows);
  const fyKeys = kccovSortFyKeys(Array.from(byFy.keys()));
  fyKeys.forEach((fy, fi)=>{
    const argb = KCCOV_FY_PALETTE_ARGB[fi % KCCOV_FY_PALETTE_ARGB.length], soft = KCCOV_FY_PALETTE_SOFT_ARGB[fi % KCCOV_FY_PALETTE_SOFT_ARGB.length];
    r = kccovWriteFyBand(ws, r, fy, argb);
    r = kccovWriteTwoRowHeader(ws, r, soft);
    const monthMap = byFy.get(fy), monthKeys = kccovSortMonthKeys(Array.from(monthMap.keys()));
    const fyTotal = kccovEmptyGroupTotals();
    monthKeys.forEach(mk=>{ const m = monthMap.get(mk); kccovAddGroupInto(fyTotal, m.g); r = kccovWriteDataRow(ws, r, m.label, m.g, false); });
    r = kccovWriteDataRow(ws, r, 'F.Y. '+fy+' TOTAL', fyTotal, true);
    r += 1;
  });
  if(!fyKeys.length){ ws.getCell(r,1).value = 'No qualifying rows.'; }
}
// Recovery Dashboard (branch portal): the "Branch Report" (per-branch
// dropdown drill-down) and "All Branches" (one row per branch) sheets
// were dropped along with their on-screen views -- both exist on the
// main npadashboard site to compare across branches, meaningless here
// since every row already belongs to the one locked branch (same reason
// as the on-screen tab removal above). This raw Data sheet is kept as a
// plain reference dump.
function kccovWriteDataSheet(ws, rows){
  ws.addRow(['Sol','Branch','Account No','Balance Amount','Scheme','Cust NPA Date','F.Y.','Month Label']);
  ws.getRow(1).font = {bold:true};
  ws.getRow(1).fill = {type:'pattern',pattern:'solid',fgColor:{argb:'FFDCEEE8'}};
  rows.forEach(r=>{
    const groupLabel = (KCCOV_BIFURCATION_GROUPS.find(g=>g.key===r.bucket)||{}).label || r.bucket;
    ws.addRow([r.sol, r.branch, r.acctNo, r.bal, groupLabel, r.npaDate, r.fy, r.monthLabel]);
  });
  ws.getColumn(6).numFmt = 'dd-mm-yyyy';
  for(let c=1;c<=8;c++) ws.getColumn(c).width = (c===2?22:(c===3?16:14));
  ws.views = [{state:'frozen', ySplit:1}];
}
async function exportKccOverdueBifurcation(){
  await ensureExcelJS();
  const {rows: mapped} = kccovMapBifurcationRows(kccovFilteredRows(KCC_OVERDUE_DATA));
  if(!mapped.length) return;
  const wb = new ExcelJS.Workbook();
  wb.calcProperties = { fullCalcOnLoad: true };
  kccovWriteBifurcationSheet(wb.addWorksheet('Report'), 'KCC Overdue — F.Y./Month Bifurcation', mapped);
  kccovWriteDataSheet(wb.addWorksheet('Data'), mapped);
  const buf = await wb.xlsx.writeBuffer();
  const blob = new Blob([buf], {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = `KCC_Overdue_Bifurcation_${dateToInputValue(new Date())}.xlsx`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url), 30000);
  showToast('✓ Workbook exported');
}
window.exportKccOverdueBifurcation = exportKccOverdueBifurcation;

/* ---------- KCC Overdue Bifurcation: Print + Save as PDF (F.Y./Month Summary and
   All Branches Overview only, not Branch Report, matching the standalone tool's own
   scope decision) ---------- */
// Print reuses the app's own single existing print mechanism (#printArea +
// the @media print rule in css/styles.css, already serving the OTS one-
// pager) instead of porting the standalone tool's separate [data-printing]-
// attribute mechanism -- introducing a second, parallel print-visibility
// system in this same stylesheet for only this feature would have no
// functional benefit over reusing the one that's already here. The
// already-rendered table HTML is copied straight into #printArea (no
// re-render needed -- it's already exactly what should print).
function printKccOverdueBifurcation(){
  const source = document.getElementById('kccovBranchTableCard');
  const printArea = document.getElementById('printArea');
  if(!source || !printArea) return;
  printArea.innerHTML = `<div class="kccov-print-wrap">${source.innerHTML}</div>`;
  printWithPageSize('size:A4 landscape;margin:10mm');
}
window.printKccOverdueBifurcation = printKccOverdueBifurcation;

// Save as PDF needs the standalone tool's own row-aware, header-repeating,
// per-page-capture technique ported in as new code -- the app's existing
// OTS PDF pipeline (canvasToPdfBlob()) rasterizes one single flowing
// element and blind-slices it by raw pixel height, which is the wrong
// shape for a multi-table, multi-F.Y.-section landscape document (it
// would never repeat a section's header past its first page, and could
// cut a data row in half exactly where a slice boundary lands -- both
// real bugs the standalone tool's own captureTablePages() was built to
// fix earlier the same day). This measures each table's own real thead/
// row heights, hides out-of-range tbody rows, captures with html2canvas,
// un-hides, and repeats -- so a page break can only ever fall between two
// rows, never through one.
function kccovAddCanvasAsPdfPage(doc, canvas, margin, usableW){
  const scale = usableW / canvas.width;
  doc.addImage(canvas.toDataURL('image/png'), 'PNG', margin, margin, usableW, canvas.height * scale);
}
async function kccovCaptureTablePages(table, doc, margin, usableW, usableH, isFirstPageOfDoc){
  const thead = table.querySelector('thead');
  const tbody = table.querySelector('tbody');
  const rows = tbody ? Array.prototype.slice.call(tbody.rows) : [];
  if(!rows.length){
    if(!isFirstPageOfDoc) doc.addPage();
    kccovAddCanvasAsPdfPage(doc, await html2canvas(table, {scale:2, backgroundColor:'#ffffff'}), margin, usableW);
    return;
  }
  const tableWidthPx = table.getBoundingClientRect().width;
  const theadHeightPx = thead ? thead.getBoundingClientRect().height : 0;
  const rowHeightPx = rows[0].getBoundingClientRect().height || 24;
  const scalePt = usableW / tableWidthPx;
  const pageHeightPx = usableH / scalePt;
  // -1 row of headroom: a measured row height is an average, not an exact
  // bound (the F.Y. TOTAL row, bolder text, can render a hair taller) --
  // safer to leave one row's worth of slack than let a page's image
  // overflow past the printable margin.
  const rowsPerPage = Math.max(1, Math.floor((pageHeightPx - theadHeightPx) / rowHeightPx) - 1);
  let pageStarted = isFirstPageOfDoc;
  for(let start=0; start<rows.length; start+=rowsPerPage){
    const end = Math.min(start+rowsPerPage, rows.length);
    rows.forEach((r, i)=>{ r.style.display = (i>=start && i<end) ? '' : 'none'; });
    if(!pageStarted) doc.addPage();
    pageStarted = false;
    kccovAddCanvasAsPdfPage(doc, await html2canvas(table, {scale:2, backgroundColor:'#ffffff'}), margin, usableW);
  }
  rows.forEach(r=>{ r.style.display = ''; });
}
async function exportKccOverdueBifurcationPdf(which){
  const container = document.getElementById('kccovBranchTableCard');
  if(!container) return;
  const tables = container.querySelectorAll('table.bifurcation-table');
  if(!tables.length) return;
  await Promise.all([ensureHtml2Canvas(), ensureJsPDF()]);
  if(document.fonts && document.fonts.ready){ try{ await document.fonts.ready; }catch(e){} }
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF({ unit:'pt', format:'a4', orientation:'landscape', compress:true });
  const margin = 24;
  const usableW = doc.internal.pageSize.getWidth() - margin*2;
  const usableH = doc.internal.pageSize.getHeight() - margin*2;
  for(let t=0; t<tables.length; t++){
    await kccovCaptureTablePages(tables[t], doc, margin, usableW, usableH, t===0);
  }
  const blob = doc.output('blob');
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = `KCC_Overdue_${which}_${dateToInputValue(new Date())}.pdf`;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  setTimeout(()=>URL.revokeObjectURL(url), 30000);
}
window.exportKccOverdueBifurcationPdf = exportKccOverdueBifurcationPdf;

/* Datewise NPA Slippage Calendar: Branch x Cust-NPA-Date heatmap, inspired
   by Head Office's own "Datewise Calendar of KCC PNPA" MIS sheet. Built
   entirely from fields already collected by the KCC Overdue upload (no new
   file) -- respects whatever Month/Date-Range window is already selected
   via kccovFilteredRows, so "Grand Total" always matches that window. */
function kccovCellSeverity(v, maxCell){
  const p = v / maxCell;
  if(p < 0.18) return 'kccov-cal-low';
  if(p < 0.40) return 'kccov-cal-mod';
  if(p < 0.68) return 'kccov-cal-high';
  return 'kccov-cal-severe';
}
function renderKccOverdueCalendar(filteredRows){
  const wrap = document.getElementById('kccovBranchTableCard');
  const labelEl = document.getElementById('kccovTableLabel');
  const legendEl = document.getElementById('kccovCalLegend');
  const insightWrap = document.getElementById('kccovInsightWrap');
  if(!wrap) return;
  const activeScheme = KCC_OVERDUE_SCHEMES.find(s=>s.key===kccovSchemeTab);
  const scopeLabel = kccovBranchFilter ? esc(kccovBranchFilter) : 'Regional Office (all branches)';

  if(legendEl) legendEl.innerHTML = `
    <span><span class="sw" style="background:transparent;border:1px dashed var(--line)"></span>No slippage</span>
    <span><span class="sw" style="background:var(--green-soft)"></span>Low</span>
    <span><span class="sw" style="background:var(--amber-soft)"></span>Moderate</span>
    <span><span class="sw" style="background:var(--red-soft)"></span>High</span>
    <span><span class="sw" style="background:var(--red)"></span>Severe</span>`;

  const bucketRows = filteredRows.filter(r=>kccOverdueBucketOf(r[KC.SCHEME])===kccovSchemeTab && r[KC.CUSTNPADATE]);
  const dateSet = new Set(bucketRows.map(r=>r[KC.CUSTNPADATE]));
  const dates = [...dateSet].sort((a,b)=>toDate(a)-toDate(b));

  const byBranch = new Map();
  bucketRows.forEach(r=>{
    const b = r[KC.BRANCH];
    let e = byBranch.get(b);
    if(!e){ e = { name:b, sol: KCCOV_BRANCH_SOL[b.toUpperCase()], cells:{} }; byBranch.set(b, e); }
    e.cells[r[KC.CUSTNPADATE]] = (e.cells[r[KC.CUSTNPADATE]] || 0) + r[KC.OS];
  });
  const matrix = [...byBranch.values()].map(e=>{
    const row = dates.map(d=>e.cells[d]||0);
    return { name:e.name, sol:e.sol, row, total: row.reduce((a,c)=>a+c,0) };
  }).filter(m=>m.total>0);

  /* Default sort, first time this view is opened this session: today's own
     date column, highest slippage first -- that's the one column a banker
     opening this screen almost always wants to see first. Falls back to the
     previous default (worst branch overall) if today has no slippage
     recorded at all, rather than sorting by a column that's all zeros. */
  if(!kccovCalSort){
    const todayStr = fmtDate(new Date());
    kccovCalSort = dates.includes(todayStr) ? {key:todayStr, dir:'desc'} : {key:'total', dir:'desc'};
  }
  const kccovCalSortValue = (m, key) => {
    if(key==='sol') return m.sol==null ? -Infinity : m.sol;
    if(key==='branch') return m.name.toLowerCase();
    if(key==='total') return m.total;
    const ci = dates.indexOf(key);
    return ci>=0 ? m.row[ci] : 0;
  };
  matrix.sort((a,b)=>{
    const av = kccovCalSortValue(a, kccovCalSort.key), bv = kccovCalSortValue(b, kccovCalSort.key);
    if(av<bv) return kccovCalSort.dir==='asc'?-1:1;
    if(av>bv) return kccovCalSort.dir==='asc'?1:-1;
    return 0;
  });

  if(labelEl) labelEl.innerHTML = `${esc(activeScheme.label)} — Datewise Slippage, worst branch first<span class="chart-sub">Scheme ${esc(activeScheme.code)} · ${scopeLabel} · ${matrix.length.toLocaleString('en-IN')} branch(es) shown · amounts in ₹ Lakh · tap a cell to see the account list</span>`;

  if(!dates.length || !matrix.length){
    if(insightWrap) insightWrap.innerHTML = '';
    wrap.innerHTML = `<div class="dash-table-wrap"><div style="padding:30px;text-align:center;color:var(--ink-mute)">No Cust NPA Date data in this window</div></div>`;
    return;
  }

  const colTotals = dates.map((_,ci)=>matrix.reduce((s,r)=>s+r.row[ci],0));
  const grandTotal = colTotals.reduce((a,c)=>a+c,0);
  const maxCell = Math.max(1, ...matrix.flatMap(m=>m.row));

  if(insightWrap){
    const worstIdx = colTotals.indexOf(Math.max(...colTotals));
    const branchesThatDay = matrix.filter(m=>m.row[worstIdx]>0).length;
    insightWrap.innerHTML = `<div class="insight-strip">
      <div class="insight-icon">${svgIcon(ICON_ALERT_TRIANGLE)}</div>
      <div class="insight-body">
        <div class="insight-title">Worst single day in this window: ${esc(dates[worstIdx])}</div>
        <div class="insight-text">${fmtCr(colTotals[worstIdx])} slipped to NPA across ${branchesThatDay.toLocaleString('en-IN')} branch(es) on this one date — ${grandTotal ? ((colTotals[worstIdx]/grandTotal)*100).toFixed(0) : 0}% of this window's ${esc(activeScheme.label)} slippage.</div>
      </div>
    </div>`;
  }

  const fmtLakh = v => (v/1e5).toLocaleString('en-IN',{minimumFractionDigits:2,maximumFractionDigits:2});
  const calSortTh = (key, label, cls) => {
    const active = kccovCalSort.key===key;
    const classes = [cls, 'sortable', active ? (kccovCalSort.dir==='asc'?'sort-asc':'sort-desc') : ''].filter(Boolean).join(' ');
    return `<th class="${classes}" data-key="${esc(key)}" tabindex="0" role="button" aria-sort="${active?(kccovCalSort.dir==='asc'?'ascending':'descending'):'none'}" onclick="sortKccovCal('${esc(key)}')">${label}<span class="sort-ic">▾</span></th>`;
  };
  const thead = `<tr>${calSortTh('sol','Sol ID','kccov-cal-sol')}${calSortTh('branch','Branch','kccov-cal-branch')}${dates.map(d=>calSortTh(d, esc(d.slice(0,5)))).join('')}${calSortTh('total','Total','kccov-cal-total')}</tr>`;
  const rowsHtml = matrix.map((m,i)=>{
    const cells = m.row.map((v,ci)=>{
      if(v<=0) return `<td><span class="kccov-cal-cell kccov-cal-blank">–</span></td>`;
      return `<td><span class="kccov-cal-cell ${kccovCellSeverity(v,maxCell)}" onclick="kccovShowBranchAccounts('${kccovSchemeTab}','${esc(m.name)}','${esc(dates[ci])}')">${fmtLakh(v)}</span></td>`;
    }).join('');
    return `<tr><td class="kccov-cal-sol">${m.sol ? esc(String(m.sol)) : '—'}</td><td class="kccov-cal-branch"><span class="dash-rank">${i+1}</span>${esc(m.name)}</td>${cells}<td class="kccov-cal-total">${fmtLakh(m.total)}</td></tr>`;
  }).join('');
  const footRow = `<tr><td class="kccov-cal-sol"></td><td class="kccov-cal-branch">Grand Total</td>${colTotals.map(v=>`<td>${fmtLakh(v)}</td>`).join('')}<td class="kccov-cal-total">${fmtLakh(grandTotal)}</td></tr>`;

  wrap.innerHTML = `<div class="kccov-cal-wrap">
    <table class="kccov-cal-table">
      <thead>${thead}</thead>
      <tbody>${rowsHtml}</tbody>
      <tfoot>${footRow}</tfoot>
    </table>
  </div>`;
}
/* Sol ID/Branch default to A-Z on first click (nextSort()'s own
   convention for name-like columns elsewhere in the app); every date
   column and Total default to highest-first, since that's almost always
   what you're scanning for in a slippage table. Re-renders the whole KCC
   Overdue body, same as every other filter/tab change on this view. */
function sortKccovCal(key){
  if(kccovCalSort && kccovCalSort.key===key){
    kccovCalSort = {key, dir: kccovCalSort.dir==='asc'?'desc':'asc'};
  } else {
    kccovCalSort = {key, dir: (key==='branch'||key==='sol') ? 'asc' : 'desc'};
  }
  renderKccOverdueBody();
}
window.sortKccovCal = sortKccovCal;

const KCCOV_ACCT_LIST_HEAD = '<tr>'
  +'<th class="sortable" data-key="acctNo" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'acctNo\')">Account<span class="sort-ic">▾</span></th>'
  +'<th class="tal sortable" data-key="name" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'name\')">Customer<span class="sort-ic">▾</span></th>'
  +'<th class="tal sortable" data-key="address" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'address\')">Address<span class="sort-ic">▾</span></th>'
  +'<th class="sortable" data-key="os" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'os\')">O/S<span class="sort-ic">▾</span></th>'
  +'<th class="sortable" data-key="cadu" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'cadu\')">CADU<span class="sort-ic">▾</span></th>'
  +'<th class="sortable" data-key="limit" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'limit\')">Limit<span class="sort-ic">▾</span></th>'
  +'<th class="tal sortable" data-key="custNpaDate" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'custNpaDate\')">Cust NPA Date<span class="sort-ic">▾</span></th>'
  +'<th class="tal sortable" data-key="fy" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'fy\')">F.Y.<span class="sort-ic">▾</span></th>'
  +'<th class="tal sortable" data-key="category" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'category\')">Category<span class="sort-ic">▾</span></th>'
  +'<th class="tal sortable" data-key="sma" tabindex="0" role="button" aria-sort="none" onclick="sortListModalBy(\'sma\')">SMA<span class="sort-ic">▾</span></th>'
  +'</tr>';
function kccovAcctRows(list){
  if(!list.length) return emptyStateRowHtml(10, 'No accounts');
  return list.map(a=>`<tr class="clickable" onclick="showQuickAcctDetailByAcct('kccov','${esc(a.acctNo)}')">
    <td>${esc(a.acctNo)}</td>
    <td class="tal">${esc(a.name)||'—'}</td>
    <td class="tal">${esc(a.address)||'—'}</td>
    <td>${fmtINR2(a.os)}</td>
    <td>${fmtINR2(a.cadu)}</td>
    <td>${fmtINR2(a.limit)}</td>
    <td class="tal">${esc(a.custNpaDate)||'—'}</td>
    <td class="tal">${esc(a.fy)||'—'}</td>
    <td class="tal">${esc(a.category)||'—'}</td>
    <td class="tal">${esc(a.sma)||'—'}</td>
  </tr>`).join('');
}
function showKccovListModal(title, sub, list){ showListModal(title, sub, KCCOV_ACCT_LIST_HEAD, 'kccov', list, {key:'os',dir:'desc'}); }
window.showKccovListModal = showKccovListModal;
function kccovShowBranchAccounts(bucket, branch, custNpaDate){
  const filteredRows = kccovFilteredRows(KCC_OVERDUE_DATA);
  let rows = filteredRows.filter(r=>kccOverdueBucketOf(r[KC.SCHEME])===bucket && r[KC.BRANCH]===branch);
  if(custNpaDate) rows = rows.filter(r=>r[KC.CUSTNPADATE]===custNpaDate);
  const list = rows.map(r=>({ acctNo:r[KC.ACCT], name:r[KC.NAME], address:kccovAddressFor(r[KC.ACCT]), os:r[KC.OS], cadu:r[KC.CADU], limit:r[KC.LIMIT], custNpaDate:r[KC.CUSTNPADATE], fy:r[KC.FY], category:r[KC.CATEGORY], sma:r[KC.SMA] }));
  const sLabel = (KCC_OVERDUE_SCHEMES.find(s=>s.key===bucket)||{}).label || bucket;
  const subLabel = custNpaDate ? `Hathras · Cust NPA Date ${custNpaDate} · ${list.length.toLocaleString('en-IN')} account(s)` : `Hathras · ${list.length.toLocaleString('en-IN')} account(s)`;
  showKccovListModal(`${branch} — ${sLabel}`, subLabel, list);
}
window.kccovShowBranchAccounts = kccovShowBranchAccounts;

/* ---------- KCC Overdue: F.Y./Month Bifurcation, Branch Report, All Branches Overview ----------
   Ported from tools/kcc-overdue-summary.html (a standalone, no-login Utility
   Hub tool that stays live as a fallback per Alok's own choice) now that he
   wants these 3 report views blended into this tab. Confirmed via direct
   data analysis that the 3 scheme-code buckets here (KCC/KCC-AH/OD-023) and
   that tool's 3 reason-code buckets are the exact same 3-way split of the
   exact same rows (CC004<->KCC-Disbrsmnt-36, CC043<->KCC-Disbrsmnt-15,
   OD023<->KCC-Disbrsmnt-24, zero exceptions across Alok's real reference
   file) -- so this reuses the already-proven kccOverdueBucketOf() for
   grouping and never looks at the Reasons column at all, which is what
   makes a 15/24/36-backwards mixup structurally impossible here, not just
   unlikely (2026-09-25). */
const KCCOV_BIFURCATION_GROUPS = [
  ...KCC_OVERDUE_SCHEMES.map(s=>({key:s.key, label:s.label})),
  {key:'all', label:'All Data'},
  {key:'l5', label:'5 Lakh+ A/C'},
  {key:'l10', label:'10 Lakh+ A/C'},
];
const KCCOV_MONTH_ABBR = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
function kccovMonthLabelFor(monthKey){ const y = Math.floor(monthKey/100), m = monthKey%100; return KCCOV_MONTH_ABBR[m]+'-'+String(y).slice(-2); }
function kccovFmtLakh(v){ return (v/1e5).toLocaleString('en-IN', {minimumFractionDigits:2, maximumFractionDigits:2}); }
function kccovFmtCnt(n){ return (n||0).toLocaleString('en-IN'); }

// Maps kccov's KC-indexed rows (already Branch/F.Y./date-filtered via
// kccovFilteredRows) into the lightweight shape the ported aggregation
// functions below consume. A row with an unparseable/missing Cust NPA Date
// can't be placed into any month bucket -- skipped, with a count returned
// so callers can surface it rather than silently under-reporting on a
// banking dashboard. A blank F.Y. is NOT skipped (unlike the standalone
// tool) -- bucketed under 'Unspecified' (sorted last) instead, so no row
// a user can already see in Branch Summary silently vanishes from these
// views just because F.Y. happens to be blank.
function kccovMapBifurcationRows(rows){
  const mapped = []; let skippedNoDate = 0;
  for(const r of rows){
    const bucket = kccOverdueBucketOf(r[KC.SCHEME]);
    if(!bucket) continue; // defensive; parseKccOverdueRows already restricts to the 3 schemes
    const npaDate = toDate(r[KC.CUSTNPADATE]);
    if(!npaDate){ skippedNoDate++; continue; }
    const fy = r[KC.FY] || 'Unspecified';
    const monthKey = npaDate.getFullYear()*100 + npaDate.getMonth();
    mapped.push({
      sol: KCCOV_BRANCH_SOL[String(r[KC.BRANCH]).toUpperCase()] || '',
      branch: r[KC.BRANCH], acctNo: r[KC.ACCT], bal: r[KC.OS], bucket, fy,
      npaDate, monthKey, monthLabel: kccovMonthLabelFor(monthKey),
    });
  }
  return { rows: mapped, skippedNoDate };
}

/* ---------- aggregation (ported from tools/kcc-overdue-summary.html) ---------- */
function kccovEmptyGroupTotals(){
  const g = {}; KCC_OVERDUE_SCHEMES.forEach(s=>{ g[s.key] = {cnt:0, amt:0}; });
  g.all = {cnt:0, amt:0}; g.l5 = {cnt:0, amt:0}; g.l10 = {cnt:0, amt:0};
  return g;
}
function kccovAddRowToGroup(g, row){
  const b = g[row.bucket]; b.cnt++; b.amt += row.bal;
  g.all.cnt++; g.all.amt += row.bal;
  if(row.bal>=500000){ g.l5.cnt++; g.l5.amt += row.bal; }
  if(row.bal>=1000000){ g.l10.cnt++; g.l10.amt += row.bal; }
}
function kccovAddGroupInto(target, src){ Object.keys(target).forEach(k=>{ target[k].cnt += src[k].cnt; target[k].amt += src[k].amt; }); }
function kccovAggregateFyMonth(rows){
  const byFy = new Map();
  rows.forEach(r=>{
    if(!byFy.has(r.fy)) byFy.set(r.fy, new Map());
    const byMonth = byFy.get(r.fy);
    if(!byMonth.has(r.monthKey)) byMonth.set(r.monthKey, { monthKey:r.monthKey, label:r.monthLabel, g:kccovEmptyGroupTotals() });
    kccovAddRowToGroup(byMonth.get(r.monthKey).g, r);
  });
  return byFy;
}
// F.Y. labels are "MAR-YY" (fiscal year ending March 20YY) -- sorted
// chronologically off the trailing 2-digit year; 'Unspecified' (blank F.Y.
// rows) always sorts last rather than first.
function kccovFyYear(fy){ if(fy==='Unspecified') return Infinity; const m = /-(\d{2})$/.exec(fy); return m ? 2000 + +m[1] : 0; }
function kccovSortFyKeys(keys){ return keys.slice().sort((a,b)=>{ const ya=kccovFyYear(a), yb=kccovFyYear(b); return ya!==yb ? ya-yb : String(a).localeCompare(String(b)); }); }
function kccovSortMonthKeys(keys){ return keys.slice().sort((a,b)=>a-b); }

/* ---------- drill-down: reuses the existing showKccovListModal/kccovAcctRows infra
   (js/app.js above), filtering the RAW KC-indexed toolbar-filtered rows -- not the
   mapped shape above, which drops name/CADU/limit/category/SMA the modal needs. ---------- */
function kccovFilterByBucketKey(scoped, bucketKey){
  if(bucketKey==='all') return scoped;
  if(bucketKey==='l5') return scoped.filter(r=>r[KC.OS]>=500000);
  if(bucketKey==='l10') return scoped.filter(r=>r[KC.OS]>=1000000);
  return scoped.filter(r=>kccOverdueBucketOf(r[KC.SCHEME])===bucketKey); // kcc/kccah/od023
}
function kccovRowsToModalList(rows){
  return rows.map(r=>({ acctNo:r[KC.ACCT], name:r[KC.NAME], address:kccovAddressFor(r[KC.ACCT]), os:r[KC.OS], cadu:r[KC.CADU], limit:r[KC.LIMIT], custNpaDate:r[KC.CUSTNPADATE], fy:r[KC.FY], category:r[KC.CATEGORY], sma:r[KC.SMA] }));
}
function kccovShowBifurcationAccounts(fy, monthKey, monthLabel, bucketKey){
  const base = kccovFilteredRows(KCC_OVERDUE_DATA).filter(r=>{
    const npaDate = toDate(r[KC.CUSTNPADATE]); if(!npaDate) return false;
    if((r[KC.FY]||'Unspecified')!==fy) return false;
    return (npaDate.getFullYear()*100+npaDate.getMonth())===monthKey;
  });
  const list = kccovRowsToModalList(kccovFilterByBucketKey(base, bucketKey));
  const groupLabel = (KCCOV_BIFURCATION_GROUPS.find(g=>g.key===bucketKey)||{}).label || bucketKey;
  showKccovListModal(`F.Y. ${fy} · ${monthLabel} · ${groupLabel}`, `${list.length.toLocaleString('en-IN')} account(s)`, list);
}
window.kccovShowBifurcationAccounts = kccovShowBifurcationAccounts;

/* ---------- rendering: F.Y./Month bifurcation table ----------
   One real <table> PER F.Y. section (each with its own <thead>/<tbody>), not one
   giant flat table -- lets both native print pagination and the PDF export
   (below) repeat each section's own band+header wherever it breaks across a
   page, and keeps a page break from ever landing mid-row. All tables share the
   same .bifurcation-scroll scrolling ancestor, so the on-screen sticky-header
   behavior (position:sticky computes off the nearest scrolling ancestor, not
   the <table>) works the same as if it were one continuous table. */
const KCCOV_FY_PALETTE = ['#2F527D','#1F7A5C','#8B5E1F','#6F3D8E','#A5432B','#2F7D7D'];
function kccovBuildFySectionTable(fy, monthMap, color){
  const monthKeys = kccovSortMonthKeys(Array.from(monthMap.keys()));
  let html = '<table class="bifurcation-table fy-section-table"><thead>';
  html += `<tr class="fy-band" style="background:${color}"><td colspan="13">F.Y. ${esc(fy)}</td></tr>`;
  html += `<tr><th rowspan="2" class="month-head hdr-row1">Month</th>${KCCOV_BIFURCATION_GROUPS.map(gd=>`<th class="hdr-row1" colspan="2">${esc(gd.label)}</th>`).join('')}</tr>`;
  html += `<tr>${KCCOV_BIFURCATION_GROUPS.map(()=>`<th class="hdr-row2">A/C Count</th><th class="hdr-row2">Amount (₹ Lakh)</th>`).join('')}</tr>`;
  html += '</thead><tbody>';
  const fyTotal = kccovEmptyGroupTotals();
  monthKeys.forEach(mk=>{
    const m = monthMap.get(mk);
    kccovAddGroupInto(fyTotal, m.g);
    html += `<tr class="data-row"><td class="month-cell">${esc(m.label)}</td>`;
    KCCOV_BIFURCATION_GROUPS.forEach(gd=>{
      const b = m.g[gd.key];
      html += `<td class="num clickable" onclick="kccovShowBifurcationAccounts('${esc(fy)}',${mk},'${esc(m.label)}','${gd.key}')">${kccovFmtCnt(b.cnt)}</td>`;
      html += `<td class="num clickable" onclick="kccovShowBifurcationAccounts('${esc(fy)}',${mk},'${esc(m.label)}','${gd.key}')">${kccovFmtLakh(b.amt)}</td>`;
    });
    html += '</tr>';
  });
  // Recovery Dashboard (branch portal): a F.Y. section with only ONE month
  // makes this TOTAL row byte-for-byte identical to that single month's
  // own row -- on the main multi-branch npadashboard site that's rare (an
  // F.Y. usually spans several months across many branches' combined
  // rows), but here, already scoped to one branch, a single-month F.Y. is
  // the common case and the repeated-looking row reads as a bug ("ye
  // galat hai na" -- Alok, 2026-09-25, screenshot). Skip the row entirely
  // when it would just restate the one row already shown; still shown
  // normally once a second month appears under the same F.Y.
  if(monthKeys.length > 1){
    html += `<tr class="fy-total-row"><td>F.Y. ${esc(fy)} TOTAL</td>`;
    KCCOV_BIFURCATION_GROUPS.forEach(gd=>{ const b = fyTotal[gd.key]; html += `<td class="num">${kccovFmtCnt(b.cnt)}</td><td class="num">${kccovFmtLakh(b.amt)}</td>`; });
    html += '</tr>';
  }
  html += '</tbody></table>';
  return html;
}
function kccovRenderBifurcationTable(rows){
  const byFy = kccovAggregateFyMonth(rows);
  const fyKeys = kccovSortFyKeys(Array.from(byFy.keys()));
  if(!fyKeys.length) return '<div class="empty-state"><p>No qualifying rows.</p></div>';
  let html = '<div class="bifurcation-scroll">';
  fyKeys.forEach((fy, fi)=>{ html += kccovBuildFySectionTable(fy, byFy.get(fy), KCCOV_FY_PALETTE[fi % KCCOV_FY_PALETTE.length]); });
  html += '</div>';
  return html;
}

/* ---------- dispatch target, called from renderKccOverdueBody() ---------- */
function renderKccOverdueFyMonth(filteredRows){
  const wrap = document.getElementById('kccovBranchTableCard');
  const labelEl = document.getElementById('kccovTableLabel');
  if(!wrap) return;
  const { rows: mapped, skippedNoDate } = kccovMapBifurcationRows(filteredRows);
  const scopeLabel = kccovBranchFilter ? esc(kccovBranchFilter) : 'Regional Office (all branches)';
  if(labelEl) labelEl.innerHTML = `F.Y./Month Bifurcation — KCC / KCC-AH / OD-023 side by side<span class="chart-sub">${scopeLabel} · tap any figure to see the account list${skippedNoDate?` · ${skippedNoDate} account(s) excluded (no Cust NPA Date)`:''}</span>`;
  wrap.innerHTML = kccovRenderBifurcationTable(mapped);
}

/* ---------- Nav / view switching ---------- */
// OneDrive/PassSheet are reached only via the Utility hub now (2026-09-08),
// not their own nav-rail items -- the "Utility" nav-item stays highlighted
// as their parent while viewing either, so the rail never shows nothing
// active at all.
// Recovery Dashboard (branch portal) trims this to just the one tool
// Alok asked to keep for now ("Abhi k liye keval pass sheet") -- more can
// be added the same way later, per his own "baaki add-ons baad mein
// karenge." The other 9 tools' files simply aren't copied into this repo.
const UTILITY_CHILD_VIEWS = ['passsheet'];
/* Screen switches used to be an instant cut -- .view{display:none} has no
   transition of its own, so the outgoing screen just vanished the moment a
   nav item was clicked, then the incoming one popped in a beat later (its
   own viewIn entrance animation was always there, but with nothing before
   it the whole thing read as "blank, then a new screen" rather than one
   continuous motion). Now the outgoing view gets a brief (120ms) fade+
   settle first via the .view-leave class/keyframe in styles.css, and only
   once that's done does the actual class swap + re-render + the existing
   (now slightly richer, fade+scale) viewIn entrance happen -- one
   unbroken transition instead of two disconnected snaps. Skipped entirely
   under prefers-reduced-motion, and on the very first call (no current
   view yet, e.g. app startup), so neither adds any actual delay there. */
/* The 7 tool iframes (see index.html for why) carry data-src, not src, so
   they don't load until a viewer actually opens that specific tab -- each
   one loaded here exactly once (a second visit to the same tab is a no-op,
   the iframe just stays as it already was). */
const TOOL_IFRAME_BY_VIEW = {
  passsheet: 'passSheetFrame',
};
function loadToolIframeIfNeeded(view){
  const frameId = TOOL_IFRAME_BY_VIEW[view];
  if(!frameId) return;
  const frame = document.getElementById(frameId);
  if(frame && !frame.getAttribute('src') && frame.dataset.src) frame.src = frame.dataset.src;
}
function switchView(view){
  loadToolIframeIfNeeded(view);
  const current = document.querySelector('.view.active');
  const target = document.querySelector(`.view[data-view="${view}"]`);
  const doSwitch = () => {
    document.querySelectorAll('.view').forEach(v=>{ v.classList.toggle('active', v.dataset.view===view); v.classList.remove('view-leave'); });
    document.querySelectorAll('.nav-item[data-view]').forEach(b=>b.classList.toggle('active',
      b.dataset.view===view || (UTILITY_CHILD_VIEWS.includes(view) && b.dataset.view==='utility')));
    if(view==='dashboard') renderDashboard();
    if(view==='pnpa') renderPnpaDashboard();
    if(view==='pnpaslip') renderPnpaSlipView();
    if(view==='kccov') renderKccOverdue();
    // Resume a still-valid OneDrive sign-in silently (no popup) whenever
    // this tab is opened while it's still showing the Connect screen --
    // once signed in, coming back to the tab should go straight into the
    // last folder, not ask again every time. Skipped once a folder is
    // already showing, so switching away and back doesn't reset browsing
    // state (filter text, current folder, scroll position).
    if(view==='onedrive' && document.querySelector('#onedrivePageBody .onedrive-connect')) onedriveTryResume();
    const mainCol = document.getElementById('mainCol');
    if(mainCol) mainCol.scrollTop = 0;
  };
  const reduceMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  if(current && current!==target && !reduceMotion){
    current.classList.add('view-leave');
    setTimeout(doSwitch, 120);
  } else {
    doSwitch();
  }
}
window.switchView = switchView;

/* ---------- Utility Hub search (2026-09-22, part of the "elegant/
   refreshing" redesign) ---------- */
// Cards are static hand-authored HTML (unique icon SVGs, existing
// switchView() onclick handlers) -- there's no data array behind them the
// way OneDrive/Telephone Directory have, so this never rebuilds innerHTML;
// it only toggles .is-hidden on the .utility-card/.utility-category
// elements already in the DOM, keyed off each card's own title+desc text.
function utilityFilterInput(value){
  const q = value.trim().toLowerCase();
  let anyVisible = false;
  document.querySelectorAll('#viewUtility .utility-category').forEach(cat => {
    let catHasVisible = false;
    cat.querySelectorAll('.utility-card').forEach(card => {
      const title = card.querySelector('.utility-card-title')?.textContent || '';
      const desc = card.querySelector('.utility-card-desc')?.textContent || '';
      const match = !q || (title + ' ' + desc).toLowerCase().includes(q);
      card.classList.toggle('is-hidden', !match);
      if (match) catHasVisible = true;
    });
    cat.classList.toggle('is-hidden', !catHasVisible);
    if (catHasVisible) anyVisible = true;
  });
  const emptyState = document.getElementById('utilityEmptyState');
  if (emptyState) {
    emptyState.hidden = anyVisible;
    const qSpan = document.getElementById('utilityEmptyQuery');
    if (qSpan) qSpan.textContent = value.trim();
  }
}
window.utilityFilterInput = utilityFilterInput;

/* ---------- Light / dark theme toggle ---------- */
/* `persist` defaults to true (an actual user click via toggleTheme should
   always be remembered). wireChrome's own startup call passes false --
   it only exists to sync the label/attribute to whatever the <head>
   script already resolved (stored choice, else the OS's own light/dark
   setting), and used to unconditionally re-save that resolution to
   localStorage on every single load. That silently turned "no explicit
   preference yet, currently showing dark" into "explicitly chose dark"
   after a person's very first visit -- before they'd even looked at the
   screen -- which permanently defeated the OS-preference check above
   from then on. Only a real toggle-button click should ever write to
   localStorage. */
function applyTheme(theme, persist){
  document.documentElement.setAttribute('data-theme', theme==='light'?'light':'dark');
  if(persist!==false){ try{ localStorage.setItem('upgb-theme', theme); }catch(e){} }
  const label = document.getElementById('themeToggleLabel');
  if(label) label.textContent = theme==='light' ? 'Dark Mode' : 'Light Mode';
}
function toggleTheme(){
  const current = document.documentElement.getAttribute('data-theme')==='light' ? 'light' : 'dark';
  applyTheme(current==='light' ? 'dark' : 'light');
}

/* ---------- Settings flyout (Alok's request, 2026-09-08) ----------
   Refresh/Theme/Sign-in used to be three separate rail buttons; now
   #settingsBtnNav opens this one flyout instead of the Update Data modal
   directly. The menu is a persistent element (see .settings-menu in
   styles.css) -- toggled via a class, never recreated -- so auth.js's own
   references to #githubSignInBtn/#authUserInfo etc. stay valid no matter
   how many times this opens and closes. */
function toggleSettingsMenu(){
  const menu = document.getElementById('settingsMenu');
  const trigger = document.getElementById('settingsBtnNav');
  if(!menu || !trigger) return;
  if(menu.classList.contains('show')){ closeSettingsMenu(); return; }
  const r = trigger.getBoundingClientRect();
  menu.style.left = (r.right + 10) + 'px';
  // Settings sits near the bottom of the rail, so opening straight down
  // from its top edge can push the menu (and the sign-in widget at its
  // very end) below the viewport on shorter screens -- measure the menu's
  // real height once it's actually laid out (.show first, invisible via
  // opacity, then position, then reveal) and anchor it to the trigger's
  // BOTTOM edge growing upward whenever growing downward wouldn't fit.
  menu.style.visibility = 'hidden';
  menu.classList.add('show');
  const menuH = menu.getBoundingClientRect().height;
  const fitsBelow = r.top + menuH <= window.innerHeight - 8;
  menu.style.top = (fitsBelow ? Math.max(8, r.top) : Math.max(8, r.bottom - menuH)) + 'px';
  menu.style.visibility = '';
  trigger.setAttribute('aria-expanded','true');
}
function closeSettingsMenu(){
  document.getElementById('settingsMenu')?.classList.remove('show');
  document.getElementById('settingsBtnNav')?.setAttribute('aria-expanded','false');
}
window.toggleSettingsMenu = toggleSettingsMenu;
// Close on outside click, Escape, or after any action taken inside the
// menu itself (Refresh/Theme/Upload Data/Sign in/Sign out all trigger
// their own effect first -- this listener is on the menu itself so it
// only ever fires after that effect's own click handler already ran,
// same-element listeners fire in attach order, then this fires again on
// the bubble to the ancestor -- so nothing here is skipped or reordered).
document.addEventListener('click', (e)=>{
  const menu = document.getElementById('settingsMenu');
  const trigger = document.getElementById('settingsBtnNav');
  if(!menu || !menu.classList.contains('show')) return;
  if(menu.contains(e.target) || trigger?.contains(e.target)) return;
  closeSettingsMenu();
});
document.getElementById('settingsMenu')?.addEventListener('click', ()=>closeSettingsMenu());
document.addEventListener('keydown', (e)=>{ if(e.key==='Escape') closeSettingsMenu(); });

// Loan table's label column is a fixed width on mobile (sized to fit
// "Interest Reversal" -- see .lt-label in styles.css), so the handful of
// labels that genuinely run longer (Total Contractual Dues, Settlement
// (OTS) Amount, OTS Amt as per Lok Adalat, the UCI row's date-embedding
// one) need a way to be read in full, while every ordinary label that
// already fits must stay completely put -- Alok caught a real bug in the
// first version of this (every row shared one offset, so a short label
// like "O/S Balance" got dragged clean off its own cell once the drag
// went far enough to reveal the much-longer UCI label). Each row is now
// clamped to its OWN overflow (0 for a label that already fits, meaning
// it simply never moves) while still following one shared drag amount --
// so it's one pane to the finger, but a label that has nothing to reveal
// truly has nothing to reveal.
// Deliberately snaps back to closed on release rather than staying
// wherever it was dragged to -- press-and-drag to peek at the rest of a
// label, let go and it returns -- which sidesteps a real gotcha with
// persisting position instead: the table is rebuilt via innerHTML on
// every drawDetailBody() render, so a persisted "current offset" read
// back from one drag would no longer match the fresh DOM's actual
// (always-zero) position on the next one, and every cell's own clamp can
// legitimately differ row to row, so there's no single value that
// correctly re-describes every row's position anyway.
// A native overflow-x:auto per cell was tried first and never actually
// moved -- it's nested inside .loan-table-wrap, which already owns
// native horizontal touch-scroll for the whole table, and a drag
// starting on the tiny label strip was simply captured by that outer
// scroller instead of the inner one. Driving it manually with pointer
// events (and touch-action:pan-y on .lt-label-text, so the browser
// leaves horizontal drags alone for this to handle while still
// scrolling the page normally on a vertical one) sidesteps that
// nested-same-axis-scroll ambiguity entirely -- delegated on document,
// like the settings-menu listener above, since the table is rebuilt via
// innerHTML on every render.
(function wireLabelDrag(){
  let dragging=false, startX=0, maxOffset=0;
  const cells = () => Array.from(document.querySelectorAll('.loan-table .lt-label-inner'));
  const ownOverflow = (el) => el.scrollWidth - el.closest('.lt-label-text').clientWidth;
  const apply = (offset) => {
    cells().forEach(el => {
      const ownMax = ownOverflow(el);
      el.style.transform = ownMax > 0 ? `translateX(${-Math.min(offset, ownMax)}px)` : '';
    });
  };
  document.addEventListener('pointerdown', (e)=>{
    if(!e.target.closest('.loan-table .lt-label-text')) return;
    const all = cells();
    maxOffset = all.reduce((max, el) => Math.max(max, ownOverflow(el)), 0);
    if(maxOffset <= 0) return;
    dragging = true; startX = e.clientX;
  });
  document.addEventListener('pointermove', (e)=>{
    if(!dragging) return;
    apply(Math.max(0, Math.min(maxOffset, startX - e.clientX)));
  });
  const endDrag = () => { if(dragging) apply(0); dragging=false; };
  document.addEventListener('pointerup', endDrag);
  document.addEventListener('pointercancel', endDrag);
})();

/* ---------- Wire static chrome (nav, header icons, modals) ---------- */
(function wireChrome(){
  const on = (id, evt, fn) => { const e=document.getElementById(id); if(e) e.addEventListener(evt, fn); };
  applyTheme(document.documentElement.getAttribute('data-theme')==='light' ? 'light' : 'dark', false);
  on('themeToggleBtn','click',()=>toggleTheme());
  on('themeToggleBtnMobile','click',()=>toggleTheme());
  const openUpdateModalAsAdmin = () => {
    if(window.UPGBAuth) UPGBAuth.requireAdmin(openUpdateModal); else openUpdateModal();
  };
  on('updateDataBtn','click',openUpdateModalAsAdmin);
  on('settingsBtn','click',openUpdateModalAsAdmin);
  // settingsBtnNav (sidebar) no longer opens Update Data directly (Alok's
  // request, 2026-09-08) -- it now toggles the settings flyout, which
  // holds Refresh/Theme/Sign-in (all still completely ungated -- see the
  // long comment on .settings-menu-wrap in index.html) plus an "Upload
  // Data" item inside the menu that still runs through
  // openUpdateModalAsAdmin exactly as this button itself used to, via the
  // generic [data-open-data] wiring a few lines down.
  on('settingsBtnNav','click',(e)=>toggleSettingsMenu(e));
  on('cmdkBtnNav','click',()=>openCmdk());
  on('cmdkBtnNavMobile','click',()=>openCmdk());
  on('listModalCloseX','click',()=>closeListModal());
  document.getElementById('listModalOverlay')?.addEventListener('click',(e)=>{ if(e.target.id==='listModalOverlay') closeListModal(); });
  const closeQuickAcctModal = () => document.getElementById('quickAcctModalOverlay')?.classList.remove('show');
  on('quickAcctCloseX','click',closeQuickAcctModal);
  document.getElementById('quickAcctModalOverlay')?.addEventListener('click',(e)=>{ if(e.target.id==='quickAcctModalOverlay') closeQuickAcctModal(); });
  on('wsCloseX','click',closeOtsWorksheet);
  document.getElementById('wsModalOverlay')?.addEventListener('click',(e)=>{ if(e.target.id==='wsModalOverlay') closeOtsWorksheet(); });
  on('clearBtn','click',()=>clearSearch());
  on('searchGoBtn','click',()=>runSearch());
  on('uploadDrop','click',()=>document.getElementById('fileInput').click());
  on('fileInput','change',(e)=>handleFileUpload(e));
  on('masterUploadDrop','click',()=>document.getElementById('masterFileInput').click());
  on('masterFileInput','change',(e)=>handleMasterFileUpload(e));
  on('branchAdvUploadDrop','click',()=>document.getElementById('branchAdvFileInput').click());
  on('branchAdvFileInput','change',(e)=>handleBranchAdvUpload(e));
  on('intReversalMasterUploadDrop','click',()=>document.getElementById('intReversalMasterFileInput').click());
  on('intReversalMasterFileInput','change',(e)=>handleInterestReversalMasterUpload(e));
  on('addressListUploadDrop','click',()=>document.getElementById('addressListFileInput').click());
  on('addressListFileInput','change',(e)=>handleAddressListUpload(e));
  on('branchContactsUploadDrop','click',()=>document.getElementById('branchContactsFileInput').click());
  on('branchContactsFileInput','change',(e)=>handleBranchContactsUpload(e));
  on('downloadBranchContactsTemplateBtn','click',()=>downloadBranchContactsTemplate());
  on('lokAdalatUploadDrop','click',()=>document.getElementById('lokAdalatFileInput').click());
  on('lokAdalatFileInput','change',(e)=>handleLokAdalatUpload(e));
  on('lokAdalatClearBtn','click',()=>clearLokAdalat());
  on('downloadLokAdalatTemplateBtn','click',()=>downloadLokAdalatTemplate());
  on('pnpaUploadDrop','click',()=>document.getElementById('pnpaFileInput').click());
  on('pnpaFileInput','change',(e)=>handlePnpaUpload(e));
  on('kccOverdueUploadDrop','click',()=>document.getElementById('kccOverdueFileInput').click());
  on('kccOverdueFileInput','change',(e)=>handleKccOverdueUpload(e));
  on('downloadDailyTemplateBtn','click',()=>downloadDailyTemplate());
  on('downloadMasterTemplateBtn','click',()=>downloadMasterTemplate());
  on('downloadBranchAdvTemplateBtn','click',()=>downloadBranchAdvTemplate());
  on('downloadIntReversalMasterTemplateBtn','click',()=>downloadInterestReversalMasterTemplate());
  on('downloadAddressListTemplateBtn','click',()=>downloadAddressListTemplate());
  on('asOnDateInput','change',(e)=>{ __pendingAsOnDate = e.target.value; });
  on('updateCancelBtn','click',()=>toggleUpdateModal(false));
  on('applyDataBtn','click',()=>applyNewData());
  on('downloadAppBtn','click',()=>downloadUpdatedApp());
  on('publishBtn','click',()=>openPublishReview());
  on('publishCancelBtn','click',()=>closePublishReview());
  on('publishConfirmBtn','click',()=>confirmPublish());
  on('eligibleBanner','click',()=>{ document.getElementById('eligibleBanner').classList.remove('show'); positionBanners(); });
  // Unlike #eligibleBanner (click anywhere to dismiss), only the explicit
  // close button dismisses #specialNoteBanner -- a note is important
  // enough that a stray tap on the banner itself shouldn't hide it.
  on('specialNoteBannerCloseBtn','click',()=>{ document.getElementById('specialNoteBanner').classList.remove('show'); positionBanners(); });
  on('lokAdalatBannerCloseBtn','click',()=>{ document.getElementById('lokAdalatBanner').classList.remove('show'); positionBanners(); });
  on('specialNoteAcctInput','input',()=>onSpecialNoteAcctInput());
  on('specialNoteText','input',()=>onSpecialNoteTextInput());
  on('specialNoteSaveBtn','click',()=>saveSpecialNote());
  on('specialNoteRemoveBtn','click',()=>removeSpecialNote());
  on('dashBranchFilter','change',()=>renderDashboardSmooth());
  // One consolidated Refresh button (top header/sidebar) always does a full
  // page reload, for every view. It used to branch per-view -- Bank
  // Dashboard/Daily PNPA/KCC Overdue only re-fetched that tab's own data
  // JSON and re-rendered with whatever app.js was already loaded in memory,
  // while only Dashboard/Search fell back to location.reload(). That meant
  // Refresh on those three tabs could never pick up newly shipped app code
  // (a bug fix, a new feature) -- a real case of this: the Datewise
  // Calendar view shipped to KCC Overdue, and a user sitting on that exact
  // tab hit Refresh repeatedly and never saw it, because their browser's
  // service worker was still serving the old app.js it had already loaded
  // and Refresh never asked the browser to re-evaluate the page at all.
  // A full reload's request for index.html/app.js still goes through the
  // service worker's network-first fetch handler (sw.js), which always
  // gets whatever is actually live rather than a stale cached copy, as
  // long as there's a connection -- so this is not slower in any way that
  // matters, just reliably correct for every view instead of only two.
  const refreshCurrentView = (e) => {
    const pending = pendingUnpublishedLabel();
    if(pending.length && !confirm(`You have unpublished data staged: ${pending.join(', ')}. Refreshing will discard it -- Publish first if you want to keep it.\n\nRefresh anyway?`)) return;
    e.currentTarget.classList.add('is-spinning');
    location.reload();
  };
  on('refreshCurrentBtnMobile','click',refreshCurrentView);
  on('refreshCurrentBtnNav','click',refreshCurrentView);
  // "Download for Offline" (Alok's request, 2026-09-12): re-fetches
  // data/latest.json right now, same URL Refresh/the initial load already
  // use -- the service worker's fetch handler (sw.js) is the one actually
  // doing the saving, into its own DATA_CACHE_NAME, keyed on the path
  // alone so this always overwrites the same single entry rather than
  // piling up a new one. This button doesn't reload the page; it's purely
  // "make sure the copy I'll fall back to offline is fresh right now",
  // e.g. right before heading out somewhere with no signal. When the
  // network is genuinely unreachable, this fetch fails with nothing
  // (yet) to overwrite -- the previous offline copy, if any, is untouched.
  const downloadForOfflineNow = () => {
    // Checked up front, not left to the fetch to fail: sw.js's own fetch
    // handler falls back to the cached copy when the network request
    // fails, so a fetch made while offline still resolves successfully
    // (just served from the existing cache) -- fetchJson() alone can't
    // tell the difference, and reporting that as a fresh save would be
    // dishonest about what actually just happened (nothing did).
    if(!navigator.onLine){
      showToast('Could not save for offline use — check your internet connection and try again.');
      return;
    }
    showToast('Saving data for offline use…');
    fetchJson(DATA_ORIGIN + 'data/latest.json?t=' + Date.now())
      .then(data => {
        const rowCount = (data.npa && data.npa.rows) ? data.npa.rows.length : 0;
        // DATA.asOnDate is stored as a plain YYYY-MM-DD string (not a
        // format toDate() parses -- that helper is for NPA-row dates,
        // which come in as DD-MM-YYYY or a raw Excel serial), same
        // split-and-reverse fmtAsOnDisplay() already uses for this field.
        const dateParts = data.asOnDate ? String(data.asOnDate).split('-') : [];
        const asOn = dateParts.length===3 ? `${dateParts[2]}-${dateParts[1]}-${dateParts[0]}` : '';
        const bits = [];
        if(asOn) bits.push('data as on ' + asOn);
        if(rowCount) bits.push(rowCount.toLocaleString('en-IN') + ' accounts');
        showToast('✔ Saved for offline use' + (bits.length ? ' — ' + bits.join(', ') : '') + '.');
      })
      .catch(() => {
        showToast('Could not save for offline use — check your internet connection and try again.');
      });
  };
  on('downloadOfflineBtnNav','click',downloadForOfflineNow);
  on('downloadOfflineBtnMobile','click',downloadForOfflineNow);
  // Lets a viewer see at a glance whether they're looking at live data or
  // their last saved offline copy -- important since the service worker's
  // offline fallback (see sw.js) is otherwise silent to this page: a fetch
  // served from cache resolves exactly like a normal successful one.
  window.addEventListener('offline', () => showToast('You are offline — showing your last saved data.'));
  window.addEventListener('online', () => showToast('Back online.'));
  if(!navigator.onLine) showToast('You are offline — showing your last saved data.');
  document.querySelectorAll('.nav-item[data-view]').forEach(b=>{
    b.addEventListener('click',()=>switchView(b.dataset.view));
  });
  /* #sideNav's rail is only 76px in normal document flow -- .nav-shell
     (the full-width, 252px, position:absolute panel with the logo/labels)
     is an overlay that expands over it on #sideNav:hover/:focus-within.
     A mouse click on any button inside it leaves that button focused
     (standard browser behaviour), which keeps :focus-within matching --
     and therefore the rail stuck expanded, covering ~130px of whatever
     content sits behind it -- until focus happens to land on something
     else entirely unrelated. Confirmed via Playwright: clicking a nav
     item and moving the mouse away still left .nav-shell at 252px and the
     Dashboard's leftmost KPI card genuinely hidden underneath it. Blurring
     on mouseleave releases any focus still trapped inside the rail the
     moment the cursor actually leaves it, so it collapses back to 76px
     exactly when a real user's eyes would expect it to. */
  const sideNavEl = document.getElementById('sideNav');
  if(sideNavEl){
    sideNavEl.addEventListener('mouseleave', ()=>{
      if(sideNavEl.contains(document.activeElement)) document.activeElement.blur();
    });
  }
  document.querySelectorAll('[data-open-data]').forEach(b=>{
    b.addEventListener('click',openUpdateModalAsAdmin);
  });
})();

renderEmpty();
switchView('dashboard');

window.openDetail = openDetail;
window.closeDetail = closeDetail;
window.onOtsInput = onOtsInput;
}

/* ---------- Encrypted-data decrypt helpers (2026-09-17) ----------
   data/latest.json, data/pnpa.json and data/kcc-overdue.json used to be
   plain, fully-readable JSON on a public GitHub Pages URL -- real customer
   names, addresses, Aadhaar-shaped numbers, loan amounts, staff phone
   numbers, no access control at all. The PIN screen (js/splash.js) looked
   like it gated access but never actually did -- it's a client-side-only
   overlay, and this file's own loadNpaData() fetched the data regardless
   of PIN state. Fixed by encrypting these 3 files (AES-256-GCM, key
   derived via PBKDF2 from the same PIN the splash screen already asks
   for) so the raw files are useless without it, while every viewer's
   experience stays identical -- same PIN screen, same 4 digits. A wrong
   or missing PIN can no longer even see the JSON shape, only ciphertext.

   Mirror of the equivalent block in js/publish.js (which also needs
   encrypt, for Publish/rollback) -- keep both in sync. js/splash.js and
   this file are separate script-tag IIFEs with no shared JS scope, so the
   PIN is handed off via sessionStorage (see js/splash.js's unlock()). */
const PIN_STORAGE_KEY = 'upgb-splash-pin';
const DEFAULT_PBKDF2_ITER = 200000;
function getStoredPin(){
  try { return sessionStorage.getItem(PIN_STORAGE_KEY) || null; } catch(e){ return null; }
}
function base64ToBytes(b64){
  const binary = atob(b64);
  const bytes = new Uint8Array(binary.length);
  for(let i=0;i<binary.length;i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}
async function deriveAesKey(pin, saltBytes, iterations){
  const pinBytes = new TextEncoder().encode(pin);
  const baseKey = await crypto.subtle.importKey('raw', pinBytes, {name:'PBKDF2'}, false, ['deriveKey']);
  return crypto.subtle.deriveKey(
    { name:'PBKDF2', salt: saltBytes, iterations, hash:'SHA-256' },
    baseKey,
    { name:'AES-GCM', length:256 },
    false,
    ['decrypt']
  );
}
function isEncryptedEnvelope(obj){
  return !!(obj && typeof obj==='object' && obj.enc===1
    && typeof obj.data==='string' && typeof obj.iv==='string' && typeof obj.salt==='string');
}
/* Encrypted bytes are high-entropy and don't gzip -- the first shipped
   version of this (2026-09-17) skipped compression and it cost dearly:
   data/latest.json's transfer size over the wire went from ~1.0MB (the
   plain JSON gzips to about a quarter of its size, being mostly repeated
   keys/structure) to ~4.15MB (the encrypted+base64 blob barely compresses
   at all) -- a 4x regression that undid the same day's separate "slow/
   blank to open" fix and very plausibly caused new load failures on the
   exact weak branch-office connections that fix targeted. Fixed by
   compressing the plaintext BEFORE encrypting (deflate-raw via
   CompressionStream, ~75-80% smaller for this data), so the ciphertext
   itself is small and gzip-over-HTTP on top no longer matters much either
   way. Falls back to uncompressed if CompressionStream isn't available
   (old/locked-down browsers) -- the envelope's own `comp` field records
   which happened, so decrypt always knows whether to decompress. */
async function decompressBytes(bytes){
  const ds = new DecompressionStream('deflate-raw');
  const writer = ds.writable.getWriter();
  writer.write(bytes); writer.close();
  const chunks = [];
  const reader = ds.readable.getReader();
  while(true){ const {done, value} = await reader.read(); if(done) break; chunks.push(value); }
  const total = chunks.reduce((n,c)=>n+c.length,0);
  const out = new Uint8Array(total);
  let off = 0; for(const c of chunks){ out.set(c, off); off += c.length; }
  return out;
}
/* Rejects with .isDecryptError=true on a missing PIN or a failed decrypt
   (wrong PIN or corrupted ciphertext -- AES-GCM's auth tag can't tell
   those apart, so one message covers both correctly) -- loadNpaData()
   below checks that flag to show "reload and re-enter PIN" instead of
   its normal network-error retry. */
async function decryptEnvelope(envelope){
  const pin = getStoredPin();
  if(!pin){
    const e = new Error('Could not unlock data -- PIN session missing.');
    e.isDecryptError = true;
    throw e;
  }
  try{
    const key = await deriveAesKey(pin, base64ToBytes(envelope.salt), envelope.iter || DEFAULT_PBKDF2_ITER);
    const plainBuf = await crypto.subtle.decrypt(
      { name:'AES-GCM', iv: base64ToBytes(envelope.iv) }, key, base64ToBytes(envelope.data));
    let bytes = new Uint8Array(plainBuf);
    if(envelope.comp === 'deflate-raw') bytes = await decompressBytes(bytes);
    return JSON.parse(new TextDecoder('utf-8').decode(bytes));
  } catch(err){
    const e = new Error('Could not unlock data -- it may be corrupted or the PIN session is invalid.');
    e.isDecryptError = true;
    throw e;
  }
}

/* Data lives in data/latest.json, committed straight to this repo by
   js/publish.js -- no separate backend/database. The timestamp query param
   bypasses HTTP/CDN caching -- this is live banking data and must never be
   served stale while a real connection is available (same reasoning as the
   service worker's network-first fetch).

   30s timeout (added 2026-09-17): plain fetch() never rejects on a merely
   stalled connection (as opposed to an outright failure), only on one that
   actually errors out -- so loadNpaData()'s existing retry-then-error-with-
   Retry-button handling below never used to fire for that case, and the
   page could sit on its loading spinner forever with nothing wrong showing
   at all. This was a real contributor to "sometimes the app just goes
   blank," reported across several branch computers on different networks.
   30s is generous enough not to abort a merely slow (not stalled) transfer
   of this file's several-MB size on a weak connection. */
// Recovery Dashboard (branch portal) reads the SAME already-published data
// the production site's own users see -- no separate upload/publish flow
// exists here (Alok keeps publishing only from npadashboard.alokmittal.net).
// Both data/latest.json and data/kcc-overdue.json are served by GitHub
// Pages with `access-control-allow-origin: *` (confirmed live), so this is
// a plain public cross-origin fetch, no auth/relay/CORS workaround needed.
const DATA_ORIGIN = 'https://npadashboard.alokmittal.net/';
function fetchJson(url, timeoutMs){
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs || 30000);
  return fetch(url, { signal: controller.signal })
    .then(r => { if(!r.ok) throw new Error('HTTP '+r.status); return r.json(); })
    .then(parsed => isEncryptedEnvelope(parsed) ? decryptEnvelope(parsed) : parsed)
    .finally(() => clearTimeout(timer));
}
function loadNpaData(isRetry){
  fetchJson(DATA_ORIGIN + 'data/latest.json?t=' + Date.now())
    .then(data => {
      const overlay = document.getElementById('dataLoadingOverlay');
      if(overlay) overlay.classList.add('hidden');
      initApp(data);
    })
    .catch(err => {
      const overlay = document.getElementById('dataLoadingOverlay');
      // A wrong/missing PIN or corrupted ciphertext can't be fixed by
      // retrying the same fetch -- skip the auto-retry-once below (that's
      // for network blips only) and point at the one thing that actually
      // helps: going back through the splash screen to re-enter the PIN.
      if(err && err.isDecryptError){
        if(overlay){
          overlay.classList.remove('hidden');
          overlay.innerHTML = '<div class="data-loading-text err">Could not unlock NPA data. Please reload this page and re-enter the 4-digit PIN.</div>'
            + '<button type="button" class="data-loading-retry-btn" id="dataLoadingRetryBtn">Reload Page</button>';
          const btn = document.getElementById('dataLoadingRetryBtn');
          if(btn) btn.onclick = () => {
            // Reload alone isn't enough -- index.html skips the splash
            // screen whenever 'upgb-splash-unlocked' is still set, which
            // would just loop back into this same error with no way to
            // re-enter the PIN. Clearing both keys forces splash to show.
            try{ sessionStorage.removeItem('upgb-splash-unlocked'); sessionStorage.removeItem(PIN_STORAGE_KEY); }catch(e){}
            location.reload();
          };
        }
        console.error('Failed to decrypt NPA data', err);
        return;
      }
      // A single blip (phone switching towers/wifi) shouldn't scare a non-technical
      // user with an error screen -- retry once automatically before giving up.
      if(!isRetry){ setTimeout(() => loadNpaData(true), 2000); return; }
      if(overlay){
        overlay.classList.remove('hidden');
        overlay.innerHTML = '<div class="data-loading-text err">Could not load NPA data. Check your internet connection.</div>'
          + '<button type="button" class="data-loading-retry-btn" id="dataLoadingRetryBtn">Retry</button>';
        const btn = document.getElementById('dataLoadingRetryBtn');
        if(btn) btn.onclick = () => {
          overlay.innerHTML = '<div class="data-loading-spinner" aria-hidden="true"></div><div class="data-loading-text">Loading NPA data…</div>';
          loadNpaData(false);
        };
      }
      console.error('Failed to load NPA data', err);
    });
}
// data/latest.json is now encrypted against the splash screen's PIN, but
// this script finishes executing well before a human finishes typing 4
// digits into that screen -- calling loadNpaData() unconditionally here
// would race straight past an unlock that hasn't happened yet and always
// lose, hitting the decrypt-failure branch above on every fresh session.
// If a PIN is already in sessionStorage (returning visit within the same
// tab session -- index.html's own skip-check hides the splash screen
// entirely in that case, so no unlock event is coming), load right away;
// otherwise wait for splash.js's unlock() to say the PIN is ready.
if(getStoredPin()){
  loadNpaData(false);
} else {
  window.addEventListener('upgb-pin-unlocked', () => loadNpaData(false), { once: true });
}

/* Alok, 2026-09-25: "dono app har 1 ghante par auto refresh ho jayen jisse
   jo bhi main kuch naya push karun wo apne aap live ho jaye" -- a branch
   PC's tab left open all day should never drift more than an hour behind
   a fresh Publish on the main site, without anyone remembering to hard-
   refresh it. Also a real mitigation for the stale-app-shell-cache class
   of bug seen today (sw.js's stale-while-revalidate for index.html/the JS
   bundle can otherwise leave a long-lived tab running yesterday's code
   until it's manually reloaded). This portal is read-only (no
   Upload/Publish of its own -- js/publish.js doesn't exist here), and any
   per-device state a viewer has typed (PNPA remarks, the address/branch
   filters) is either already saved to localStorage or trivial UI state,
   so an unconditional reload is safe here unlike the main app's own
   version of this same timer. */
setInterval(() => location.reload(), 60 * 60 * 1000);
