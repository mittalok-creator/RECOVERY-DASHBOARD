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
/* Branch/Region NPA Target for the fiscal year, uploaded on NPA-DASHBOARD's
   own Settings -> Update Data -> "Branch/Region NPA Target" panel -- this
   portal has no upload of its own, it only ever reads whatever the main
   app has Published. Powers dashboardNpaTargetStrip()'s "Target for <final
   month>" tile below. */
DATA.branchTargets = DATA.branchTargets || {};
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
/* Full Customer Master address book (Customer ID -> Address), published
   by the main app independent of the NPA book -- see its own init comment
   in NPA-DASHBOARD's js/app.js. This is what lets addressForAcctNo()
   resolve an address for a KCC Overdue/PNPA account, neither of which
   exists in DATA.npa.rows at all (confirmed: disjoint account universes,
   zero overlap) -- Alok, 2026-09-25: "fir kcc overdue main kyun nahi aa
   raha" (why isn't it showing in KCC Overdue). */
DATA.customerAddressMap = renormalizeIdMap(DATA.customerAddressMap || {});
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
      await navigator.share({ files:[file], title:'UPGB Recovery Dashboard', text:shareText });
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

/* ==================================================================
   Application Form -- Hindi "OTS Compromise Settlement" borrower
   application letter (new tab, Alok's request 2026-09-27). Ported from
   his own ALOK_OTS_UTILITY.html reference tool's buildApplicationFormHTML/
   numberToHindiWords/formatDMonY -- same wording he already uses today,
   auto-filled from this app's own NPA data instead of typed by hand, with
   only Outstanding/Purpose/OTS Amount/Token Amount/Token Date left as
   manual input (no data source exists for a negotiated settlement).
   Print/Save-as-PDF/WhatsApp Share reuse this file's own existing
   printWithPageSize()/canvasToPdfBlob()/shareFileOrFallback() -- the exact
   same machinery the (dormant) OTS Calculator above already proved out. */
function otsAppFormatDMonY(date){
  if(!date) return '';
  const months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
  const dd=String(date.getDate()).padStart(2,'0');
  return dd+'-'+months[date.getMonth()]+'-'+date.getFullYear();
}
/* Plain flex rows instead of a real <ol>/<li> -- html2canvas (used by both
   Save-as-PDF and WhatsApp Share) doesn't reserve any marker-box width for
   native list numbering, so a real <ol> renders as "1यह कि..." with the
   number jammed straight against the text, no gap at all (confirmed
   against a real generated PDF, 2026-09-27). Hand-numbering with a
   fixed-width column renders identically in native print, on-screen and
   in the html2canvas raster -- one shared implementation for all three
   output paths instead of a print-only fix that would leave the PDF/
   WhatsApp paths still broken. */
function otsAppPoint(num, text, bold){
  return '<div style="display:flex;gap:6px;margin-bottom:12px;' + (bold?'font-weight:bold;':'') + '">'
    + '<div style="flex:none;min-width:20px;">' + num + '.</div>'
    + '<div style="flex:1;">' + text + '</div>'
    + '</div>';
}
const OTS_APP_HINDI_0_99=['शून्य','एक','दो','तीन','चार','पांच','छह','सात','आठ','नौ','दस',
'ग्यारह','बारह','तेरह','चौदह','पंद्रह','सोलह','सत्रह','अठारह','उन्नीस','बीस',
'इक्कीस','बाईस','तेईस','चौबीस','पच्चीस','छब्बीस','सत्ताईस','अट्ठाईस','उनतीस','तीस',
'इकतीस','बत्तीस','तैंतीस','चौंतीस','पैंतीस','छत्तीस','सैंतीस','अड़तीस','उनतालीस','चालीस',
'इकतालीस','बयालीस','तैंतालीस','चौंतालीस','पैंतालीस','छियालीस','सैंतालीस','अड़तालीस','उनचास','पचास',
'इक्यावन','बावन','तिरपन','चौवन','पचपन','छप्पन','सत्तावन','अट्ठावन','उनसठ','साठ',
'इकसठ','बासठ','तिरेसठ','चौंसठ','पैंसठ','छियासठ','सड़सठ','अड़सठ','उनहत्तर','सत्तर',
'इकहत्तर','बहत्तर','तिहत्तर','चौहत्तर','पचहत्तर','छिहत्तर','सतहत्तर','अठहत्तर','उनासी','अस्सी',
'इक्यासी','बयासी','तिरासी','चौरासी','पचासी','छियासी','सत्तासी','अट्ठासी','नवासी','नब्बे',
'इक्यानवे','बानवे','तिरानवे','चौरानवे','पचानवे','छियानवे','सत्तानवे','अट्ठानवे','निन्यानवे'];
function otsAppHindiUpTo999(n){
  if (n < 100) return OTS_APP_HINDI_0_99[n];
  const h = Math.floor(n/100), r = n%100;
  return OTS_APP_HINDI_0_99[h] + ' सौ' + (r ? ' ' + OTS_APP_HINDI_0_99[r] : '');
}
function otsAppNumberToHindiWords(num){
  num = Math.round(Math.abs(Number(num)||0));
  if (num === 0) return 'शून्य';
  let crore = Math.floor(num / 10000000); num %= 10000000;
  let lakh = Math.floor(num / 100000); num %= 100000;
  let thousand = Math.floor(num / 1000); num %= 1000;
  const rest = num;
  const parts = [];
  if (crore) parts.push(otsAppHindiUpTo999(crore) + ' करोड़');
  if (lakh) parts.push(otsAppHindiUpTo999(lakh) + ' लाख');
  if (thousand) parts.push(otsAppHindiUpTo999(thousand) + ' हज़ार');
  if (rest) parts.push(otsAppHindiUpTo999(rest));
  return parts.join(' ');
}
function otsAppHindiRupeesOnly(num){ return 'रुपये ' + otsAppNumberToHindiWords(num) + ' मात्र'; }
/* Real Uttar Pradesh Gramin Bank logo (supplied by Alok, 2026-09-27),
   inlined as a data URI -- same convention this file's own nav-logo/OTS
   hero head-logo already use, rather than a separate asset file. */
const OTS_APP_LOGO_DATA_URI = 'data:image/webp;base64,UklGRgptAwBXRUJQVlA4WAoAAAAwAAAA/wMA/wMASUNDUMgBAAAAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADZBTFBIdhMCAAEhh23bSBKs/Vc3/Rd8nrkrIaL/E5C+7dN/AWy1ABu97bCtPbAv7Ttsz8rPdp6VxxEnTjMdBR1JVFFDEg32qAWp8h2IiZqPSUJS0LaNtOTCn/R9FCJiAuaM57l2OVc5r1u+5teq9UTlIyqfQsCcufqaglxF7lXc0jAoqE2hsugKIFcHAdxISpocYAD/BvC41TB8iZiACfCH7f9620rFPa/35zvGnGutvdfuzaYbAUUEFcXuxBYV49gd2F1HVESP3YGt2IGB2ElYhCDdsIHdtdacY3w/n9frjzFmrLXxHJf8IiImwBttbcsmSY50zvW8n7l7QFJhlljVomqVqHGYmZmZmZmZmZmnmRkFJdb0qNUk5mLMyswId7Pvfe7rx2fuUYoZ+Vpp9qMiYgIsN5IkSJIk6nX8U3wV9qiepSEiJsAb///rZ9tp9np9fv+ZtfY+lpNz4gkSkuDa0OBQnJQ2uLu7F4eU4FrcKlhxLVSoCw4FWiQlQNAYsZMc2Xuvmf/v87oxs9deZ3Ndvc51LyImQKX+r//r//q//q//6//6v/6v/+v/Fnj3hJ7xFfz1WX9jj+d6LgJP8O/nfK7Hv+H4qft381z/j0lc1R//mWd6L99ztXg619/3PO/Newtwc5v2v/ccb28Vs7AKrn76B5/f5YwUqjY6/rbR53Qir6cJglpLjheXv46D+Fwu8KeMREpAtQ1/k9dOn8k1nmY2Q0KwGf/kTzc8kxevvLE5IyKKcGaAv1J647O3ztkbCFGoANqabPh6ebTP3eBFo9RGCSS7sRAIvxXLc/hfCE56h1AABYrX+3+U+gzupd2h1GZEtgy3vffPfh3z7O18drnMHN4IskCJgPPy8tfieu72ue+T1y0mEkqBzvBXsPZ+3vbZXqEWkhU2oEinx/rGH4E+a/trERpIbuojadnJ5k9c7058vva/0v0ZB6fIiChAqnPQg//DfoA+V/vbVw/OhQMSFkAUoEPp8AMsfab22hfMsndBCExQlEcl2q5P/ijTPkvbfz4dczkfgRAEeevQjcPvek15hn7s7+XScbmwWshDieIjjR5w4f/AYdZ6bsaXLwkOIA8XEMtQALntyW9l6+xnZSd/79lQQdAWTQWBMHLbQscf+zGkfU4m/+Y9SAGoiaAkWEgGShWm6wdwZp6PGfuVXHNCyuHhggit0GawFnYXfwXyfDzgD7KAUEKQw8MQBZVQmM7qXH8beT4G/DG214rHLeJRIVpoKVS7TMlP/2TOZ2Pyee8/9qYAoTDpERBRUKQCzIz+jcxzseC1W2Yk3QSG5dEtgHaA8pZGw/+CFxB8/sX193OwZQlYDUCMAgwSoOVmam3On3IA+vzrj9KZ550Chi5nPAwIBDtAgSEMsP+RH+1elmff//oVD5nLhps8lIe5IdACpBiKkNiVv5F00+deX/IqK1oxZIAkxFg4YkpBBEZo24f+v1jqc6//Ezl1GEDcUS2GUBgCtGihYHh096d/iKnzvOsXzl4nFjYBo1CWx4OBQmhtLXQYel4vv/n1ap93/YO9h10mBVLk4QAUEpWKRVEEg0DyV3LgMy7v/v7Lh+MBtQGFRQQpIKBIgQIUpgA5Lvwhpj7bij/A/S7nOk5QkECBtAQgDIXS0rZUKczOT/+xtQvkWdb84T1kZ2p3UMmCpNaJcG1bZxBJBy12S1v/5oc9aZxnWH7x57keZgyxUoCyOGSVBErNggIGqoMVD/4PJJP2+dXJb5wLe5pD5jIwhQNkMyq0GAWola2Eapi0D78XxzbPrvhD13NgA6QGMHWxBwsBjgK1wM0soEVG2/zmM5Bhnlu95294TnntGmNgESMhTQAxAKUibQMWUGiKi/8G4ObZ9U/g3GYvjqQEk0AUIpA8TAulVKU82m7A4frp76GDz60+7z2XuTBiELpggKNCGI8WUBERWlFAD6R5uPv7z8tx2OdVX/mHhwoUK5AkHm4J2gNpgfLWWspQoKVy8D9g7c7zqs+8u3NsoWZZd2MBlinaxY1a0ILQFrYUcKAFuptv/J7LXuezqq/4E7csbCljOcNQOzQDTo0rQ2grhaLDggyooJCE/+79eH1GJb+Q+9aL9lAtEgYIDIRZgAIZKJjiAALcgLRn/w/kwedSxr/++5eTClTHJFECoYmokbKC2KbcGqCh3A4bevzBn3i1o8+iAr7fDCdAK5sYAW6AEMiGIkJp9LFHK/YGh3M2/0oeKM+l/9EXum3pEFaDBTMAicBAEQUQKQqUMsg0QNG4/eN/UGL7HOrghx3eHFKtsyglURvJAK4kYoWWFmlpUwklvF3X5V/F3ZzT9fzpPP6nf3G5v1okjMtCCAiyFE3xqEBHkTI3AAU3lEILs37s968zLPbzJz6NowYHAhgUVjAAH+iwQEulKbcTR0trwSJgy2T/Gxia8vz5i/+3ObsKsduQu+0EVUNjACutYimpj9AG28htLYJp80N/34qd5tnTf9jTPdwzHMp1mHCdIkOJAacQ0IEpdSnCgm4hKVOwrPNfQKB0njm97+fAedLBICZNuhMzhQGWsM4uRSE11imEYcQFdETbGa+/8x+50B195vSbvviyc9qVzURiQdmclscDl5mV2wHbSpAShLYdg0Xl6/tfiVDa501f7/64P2ZdDSycBhJSty0lPUMqUjKtCDhIYQ/VwCBMP/3gv/sTd2d7DM+av+T3z+14HAMKSg8RATrAKcxDQBUKKAUKMhUTUhhSqXx0+e96ZPV83vTjvvSswt1lQ7DKEkADQx4uD0sFKEJtIbALVtrh1vovvX95nweeN3+jQ09qlCRxDIgkBHcjIKMFRApiqbS2sDatxViZ7vWP/U0fXp87fcGvp6MA94gHMIVIQjRHuTAKgzAAloICBBSC0DIAdc1veDEH61nTz/3i83TMnVNhgUaAhjCXESGixdpaqIIwQloLRGxkCde1+edeD64zfbZ08MlzQFKz6ISwWoJAUzx0bYYVURzk0QaYWoGWsdChl9fHH/x7Xs419bnSyXs/gyAOmArCtQhgVZRkGSAmmBbAoSAINIGWURjAyHk8fN+Hnj07z5Tgv/4rnlQIApaVAzEssQFDkCBBEByDvGMpSihRYHo+XH7l4TD02dK3fu24RyUpQ8eBMhhkJEAECpEW6KS8YyuPF0HoDv3gf/ZnvnjO+Fxp+CiOYVlgCNqMhwIsIEQlSEBBBdP2sba+DVpuXVwvP/If/ejTpvM8aX3fNziuDMNGGvJ4FNkjAhYgIVIoUCIbQJkbZRBbSjkv/vKPNj15ntz8/V9yuWcDGHZYgFUKHixEbArSLlAQHKgDC1qq0BbkEYSJX21hzzxHOtZf8U8udxMasBwChgoGEyCHPFyOiUAojDQNU6o4QpTSYhzYk3/dj13YV4rPjk7yU+4v9yCSAphWBgcSEkAtA4hIW4SIFSIWJpSWUZlypvT4PX/lxykD9rnRw3fPUAFKtbZAAwHCAhGPBhu3MsAAAgNIU9zalKqknvnaP+5jriPD86PmfV/9uE6RoCCOKCEgjSYyD4yRgIoB5HGBlhYWj4qU83R/x6tjM5P6TOfrfosXx/ATuRDZQipUgUAL8qg8LkCB8Wj0oAeiQywiFc4BN+++uz+v1+vSi1rPXH7Y//GXfsML8+RrsOdMxAJtq+MCgVSxFBEtDxWigCaINQhokyFCpGbwHXocl2F4geuyeO769X7vO/envrDk4zdhViREaHOJTgMKBFccWB4uagBSDAQrWIRZLkTwzvNIkBfoP/7jDftZy+VvP8nb7/yilLcp0iyuGysNuOCB5niAJBYrGEOsQYmcrDlgmsGJO8nO8tEzwjjzxr7v//4H/nngs5a/fXPF1377C4MPI7jWAYIODUsDuKwsLKCQjBSgDYCie3BALA0JeNAgjvDvnHQINm/o//VLv+dv/8/T5yz/0TfY45pf7SNe1PISnRwjEwwEMQ64yAghUIs8VIMlzAiEAAbBTUqDKvkqeTijkz7fV79vzvmv/RN5xvoJv3bn3K79whcUxxO4rESrLAkQtSLABJsg9iAeDgKIIZEAASNhRCLxsSxcNCyf61/O9Nr/+3OW/4nTLhx8/xt8ETrsOJt4QKMoD2dsgRgQ0JIFQUBCCGNVJORRRQYZHJ5eZs/mRhbQ18t/6DJJf9l//fnK9/wULt0cx8WP+eFcDl8A9/dCIw8jBAIMFDB6BASmeBiL8bpTBAJWvG5E8Em3r8h5vd+OocIA8k/4VV2XteY//xsf85nJ8GF/tNXVi/wUrmdvbPIyJEEhUAGUENECrGtAhFDZkIAPEiiKUHpMaOlmvIh6AAgdsPz7gWrzt8abPjNZfs+Tpuss6/WT3srwxvewZNdd7NSIwG1pddPNiSANCY3ahQrILW0FbLGoE1Q5gcuMrjtHPFr453XsSL/tv9mbZ6cf8/0ALwDOW75lxwsozmJYAo9yVdvRGCKaAQGMCgsbOTaSh0KMGDRg6QGs7e3N4c6h46VTEVl89xcIpAn/yV/Ieoby1+7Qk2pn9xfh8caGTc+FOTgpOYCYUMQ55FExaLTgRBjUCQRHeahMgAaBdOxJCmaOUQv8a1qDlJ2/iv3cRL7fVxU7QI9rftrtXW8skQgDciiJYew0Hg1KE4XUDimAMBKoEKA1cAt5uPf340XggFkw6ubfqoNk9+76j/8nkmcm+fueGZJTR3e335J5AcO1nUZQt1XQapmEJQMhHtWFoQRMIAFMOBccJGak0JlnK140opus7eILv2pbbA+uc/1fMc9M+CMvd7AhWO7+Z5wvIL4SdZOGcVigfHRjkgEMiYBhF0xAQheIoZ0JKOTRCPl8ihygTqVl8497RVdRXZfrr/tXkOckBx/zA3oKQxDjzRM+5iNvnDeyV/7uqLgYnODyqIGAVAACCZSDCSwIuEDgAIHGw5zYe+65felQJ2fauBx3i/8pChQ4Li/mr/3CXJ6PGH+5PWMxwHZ66WfdiT7fIX+dltkB4QgGnA0JoFOpFdLYFBYDaoGcMwgiHkYEBix8PrfOHDrMOLl3z6/+quuiDkDH4a9kAz4PYb/TJ54DDlQn4578CJZ6Pg7+F3aEtrY8gQXZcOV0CpEgOCFgKHZHHi5HZqsoJQ8DZOCf7sFMErsuIfxXFWxKsXH/875ngD4TmT91dZU4zQG5uX7Ut2EG8HmYL2CuLRojHQuSBB3LsZDpSQiDILU2Q7PhwJwDE56h6PogEv7102Na8vBYkHDlX0dmoIHWCn89GJ6H9kPe7iGhh+TCCH+EFuh5dr8yA0lg7WBpY8yWlKmaknCwVg8DRIPqWAg45NEBETjbfyyrxLlLxl78kz+eIQhTDHr+Zb+QzjOR43cuhFSMyJb7cZ98DvKcs13ur8dGjNbkYigWHCiuo85UBOUEeCYIATIgUFAABe5w3H/urTMAHjMA2fwPqEOx0nRI+jehz0T+k5c9rs1uSSarxi+j51vW6/u9IaHQzonRZSMkQIhKB0CJEzykeHSXEOOhbuQQtb7nvTcOg7WbwjG/7C/tmjUFTDbLDb/pX1yehz75CdcTZxsHhLiU8e8/icXXo5Ev45yRlHYOhA10AHxwIuguChAeAFuGhgMQgAAjQiBzfM5rLSsTuyxS/uuzGxMqTGCSh/736HMPD7jhv7i5OThmBnlUQi/79h/PrZe5vB6X+Od5Qi4qsTAOhEYERyDM8KjIBgwDgQ4IIhAB7JZ0nvtveEKdMLcvPZnLcXPMx/+GDqsFC2hTXvCr/wkQyPON4YT7t/wggmBYClgUgs/iHNzBR27u5A+7ULhgNZDlEGJABkVs8bAhwEWhhxuwhUgx44CH8+fnpTlkGupmODv/bXIBncFiK+k4/2sOZJ5vLCC/EqTYGgwaesjdO77zOVfIQDg9+Z9Z0HUKARNc1tjFCEsB41FDgImHaiMxDBEWEMz9hf/tyU4jJt0b++K/0dCBpBWFjs756379WeQZZ9DtT4pIgG0HBkTU80/SHCsD4RSf94pnS60W7QIIroyLuRngRoBgAOEWFkaWQZJLuSHv+zeXC+6uOOMu8+/9wmAEWgGkgXj+vznaPON4+ENurou6yAhFIIuXffv3ufpsYhHYSzf80+NsHAIcxgcNweIQBlYwSrAxD4BRQDTJBEiPgHGZ/zYWhLPzTC/ny/8tKC2DYaAAu7Tf99VT5nnHzW+93giiawgiwATe/d7LOjIGXnbkvzt0wVkCsUXC2SCEQ5ANEDZ5GMXDeCiIEVQohBd+7dwePIzDaPiZH46tNgUyhQnLRv63Xfi849u8pS0Ip3KBgHXoPN/5U7rQLCOds+N/pnhaMq3AVFBzsBLUCR5OgKMGwTwopQAsUYaHRievfvaHzRbEXs/y7h0/76SLXQGK2tDR7Pt/4YfbPu/4L6/MDCkxMIDAAN149+vo8LR1aLu7/afvf4IHjsUAQlMT24hmhwBJPOpiMYQC+iAFAwWqzT95vu16XvdsDyYO9g/cxuBKEcQiGMqd+fd18Yxz+FoffdwAGAiWBLDA4e3NzX9zdSctd9l7flcuW5hsuA2CO1IBsg9Eg4KhESAeD5CH8ugyzE2//S289P5nz+6fnfd7f8w++0WfcVEpCKUgb43Of/Fy+owjfgdXqJBgETBAAqc+/Ru+Njs6MDjw22ObAheDcckUUpEaohZgk0AoBCgQoHhY0cAsn/2P374Moh3T5dm7f8VqeXxEYIBS2cd88S8lPuO4fNvzAhohThtlBQW4/BnW3TbWNv/Vl0gnszkFpOdgS1QsQ4haNsPjKY9q8dAgUmC5XuWX75y7blNde8ZfJoACbVpoQ9tSlov/PemzDfn+BwJNEqCA6LojsNx/4nfRoYbpBuR3IcNqi8MCs+WRgszCQqyGPIwyWIOQHmCZ4Ml4Yd77F18+5RLnRrP33/drXo8qCIo3go0I7Pm+r1wlzzTi1y9dwwQJMFg4IE5S9y98/ILHgifBb3jfEWqN7o7rwyXQWPEAgYYIECQYIECBUJHgYOPk515fuvFyt2uxc37YH29Wbnd5tFIIIww9+c9S5pkGL30Eh2MIuwgKDMDEgTOX681vkTHLxHnlh0Jx6sJQA7LDQJny+gIJgssAlAYEEAUCSDfv+wPvSFYAT55d/5YNDyU8LgggnOmx+PfcEXym8V0vteICDhFl+6CBAG72M77L7j6IOYD/+X6SgwLQAogAUQIWpkp52A4UDGA8NFGgMpyferxl7rueFXk9f/nXWDPeOkN4tAANxfnCX8KAzzN+U8tI8wBAHDRClkfzlT/8TgCnOeGJr/5nhltHMJDAxBucSBV64AAKCBhCEo8K4773T3/4s+PC6e65sd/2F5/KwxbFAFgQxNTt8L8mFH12IR/z1boEUQBJLKUAibHo0w//m8yTkYMj7sZfA+f9OMjDgbZ50BkhgAgtCAQBFM+ZIAuZe73OL7x9+e15zPXu7LrPPvzPM64mOgWBggBTqKwe/f5vnwD63CI+g1VQKbC1SQN3FViRuf+kX7s7N4Mzclz+1f92dAGDAFqGcHFE2ALYGIACpNIHJUBFRlDn5V//vg+/ubu4c93w2Vs++60nDTkdEJgBOgUCMEiH/wDLdvrcgvkpTQuQ0cKcrCGbswvpbjPnz/6UO8+TTWRuf9D1ADCAxVgNUVoQ91wMoNSNlIITLVAfsCxnt3z7l5+ex54LM57+wY9ylvsoBmiLgBGoxUCFfz9XFvLs8p2fZBwBporNioDYTEw4nof7V196dhkvBZ1Pv+Tnzp5sLjTlzi5B0YQxx6xuhNCA4AoHgMQiQzqqv/NfvttZmPMy969eP/N7YXMeOTDNoGGgMB2xxR3sd34bayA+t/jORwMuIATp4dYDgmiAxN71v3u3LLHs9fgt/1yuKxOCc+AEaQNBgDgiBAixB0ABSUjn2LnNF/2Cd992cnK/1zq/y2ee7KBzNs3UtqhIg4BQLeHfQl2lzy1+7CyAbEklgArJhAgtnd2cn/q7Nj1ALteb49vOycAO4UbuuPKoCMuji4DQilAJJhwA0zJM3/X+LSbb3h1z/43+O2SWg9XQRSMgFFEYwCB7/5tBnmH+O6GbKeDYQgYsJLWgHkN79+M/g7PEvXb/9PN/8m2zDStBJQwQQIIEIhAQEqBiPL6IM5zzK//Fx9yfXLc975994Bv9HZb7lBIoCJaCwNAGC615870f7pQi+Izi3W/XdmRdQ2oADCZkYUqrmLn//V9nd+Gwo336O//+NMDw6GgrMGAmIoDUAErI4yLEziLndvM3f/U7Wk6vc3J396l/yR1uZJnDLalCwULZQWmhavvRXwjHUQp9FuEjn86MExwrRg60PJQlxwJ0jvHm5sP+4Sdxq82RXOeb/GuvdC4wQjDQAsjDBYgUCOS5C3Zyt7me7/vhwuW9r7z27NXD4xP/1yfEtA6RWAujPL64FXFojv2/4byO8pxSfgi7QCQgspvHI0iyMQKdAMXfv7w6krA+uX67V7hnJKsVIObBCbuMldijsBAEFbiYjObNp//rD+e83u/99foVX/EN/95lOXYdApBaaChIyxRKbQmM/QtXtcRnEg3ENwBaXHA3lhlbepDC6AZ4bGXXp3/veHazLQs8/Zc/6OLpLpjCAkZwLA6LPKoooIakFYa7kZcf+TffWdndHdfXPvGv3nKJmYWowloEpwMhzDgWKGJffBdrtPV5hIvcfgQCMoKQC9v0EAh53RgOm0tf88/xmtXNyvHkv/uThx3s5rodZSYwCbOwkWDGEBHWNDCLY+uv+8PvTNqDru/98L8FQ7QTpgNj02nrpFJsGqQtyJV/HruurPosojE+amKQM+iRqXF1ACwWxo0AFrLv/udOjuUq7Jw//P/yeh/IkZQIsTWwkIcJJEYghjKdZNIef+Qz3/bSULyvZx94+fPeegIS8SBssCiyOgUIwlZSVP7tlE5Ln0VMXfhWAMZEMropTCGBmxDD+mBKrudn/Jb7+7MdiKd9p381h0eUwUSwiDzUgAgJUASI6EBdr5fP/6lP3nEeuVz2vPmKy327PNQQKhRtwZqbAsNC67Te/1KEpcmziI34nmwxORCFLmtKIIyA4FjCSPnqT/uP4+Ratjcf+PRnnjuriJRwEECDpyAJGBCAAM2K6/Xm1U/zw/eY2eN6/763ft7c3agrCA9UBkCZWoLchrZAXPbLaCv6LAJYvh4ECKugQgMJGQpRwKkBcR7n7d1nfk92Q/fgpX/879zfEIK0GsJYSB0QBoSQLI8awzJ74evtS2s7h8+eveUfvMNbDjJGGAK2WHCMW5miwIAIzNEX3z13pJ3q84j4aPawgiHkoZHkxAIaAgILdpmbueVPfb/jpdubw9s6n3zOd7zb5f56njEi8VBAAWoAQ6QBkvVc2GuvfP33vv2dL18u57WbZ+/4px/FEo5EFcZOU9oKpDQAQ6OlhbUu+Sf1cjk1U54/CvDkbR2tClH1YGGAXYxlVzYwhCCBnT/+TT9wf1Zznr70l38ox/U4mmGXVtjcJfZkQeJ146FdU5frsd/u8999cz579myfnV9685c/bJmgEgShZi1W6CCAQ1PLWIUZ+Tcy65JI6rOHQD7yQg0twZC760iU0jQ7A1hSwOqyeN7/tW94f73bluV425/+0R3393bukAPhOnXOILZCtKFsnV4s5rr+B3/vnXP26nne71fO3/qUZFsUIIyA0TaoDsWSYSikttN09zfZ5YLUPnvgwdfgOgaDmjSjtBRk0rA0INqgMA0el/2/vsa957m69dY//v13JvWEJTTLmeAUBlAUYmRgODxvn/7Uv/rxx11cLu2rN//zJ+8hMwdAoAgjldAWtAHbNAK1GrO8/4411+RQVPwM9oGwfFrnmcRGBieWxMgwSM0KEkixgFud52d/3N1xudI08+TP/8CFZmcUiBRy5VIuAUoIa2fQ9dj/8D/7iL3ewt1r969c/5tvstKy22JGFZAUpB2UAawMclvY1BcfbErRofQzfCZ+S2PAJkg9NoVxDaCAkcoHpCQsXLj8o3e/du7SwT1v/e++35XreVDCiiDLEWgTjwvkqon8yv/xo2e9u7+6d/d/61tch5TJAXF9CCwU1AxgESXTilhs7C8619KAIO/ue9/sixk+qRshaYiEkZoaXIhBALUCB5aYw9m9f8s/ettr59m5O15f/l9+7DHniUIDnrIHtIBIEI+KDtDxO3/jx9748rH3r/Hq9a9+w/MCYh20SMMG4VhsoUKRcrsot1KKvzwuslZEwHfpu54QL3z5uD0sUxIBNkQhQ3nOECBFgA4v9+/4x+/gvOf0QuftH//3L7cVMEBHDDTDwyQEYtE9d5n//Ve+7dY5Ad+3f/lTrwPQJjiFNPKwKmiBKmNvwEcoOvzFx2TNDKsL6Dt87//k5en9TuDnfNmX/xpu5sUccj1SEVArQukERnfBByegPQBheXg5bi4f/T++09uWuz2v9zf//df68ploC6QAgQAUCBnORvc1Xru/3Fxubi+3F55e/zyXWNJZAiUIElQpCjMYhMiAQJlbfvjd3F1H2bnGo3Mcl3nLP/1p//If/stZy/WULuA/+rCXf+aH3e+8EORloCBqU2Aoh41thDY6HmSALUm5W+d+0394e890np4eb/snn/oV7FXWOhOCyIKNJdg9Yrxc2meyBx3H5aWnbw0bBWYlwgyKoIeAInFG0JqBwJws757zcszMAcPrcO7+LuDyN7PLfkoH/MkPv3r7gzjGeQHncjBoGym7wYYsQ0MLjVkAEeVQGFMc7rs/52a7aziCt/+Lr/vK0RXaOSo8sZKVBM6Z3No73nVznr3KcT30rr3KQrQDBLXCUJthikHBCC0DupwRjvuWp8PRcAqCAB771u+zZ/od/ylInta9/TM4jvtfzS77hhxgOA0ZIHRSaY/S3FkgFrYwURBBGCe8fNzfu1nm5Eq7b/+yb/KBGY6dXRCOzTENgYNU86mHN3Px3L1rb/+P43IamRYmDAYyiNgEDdUg6AwJNkJD8fT+5riJw0OGAB2+90tdYP7LXxh5Wv8XecZ5vO17RPOGWjmOa7sQKwmxrgKyojXpuKrJQiSbIbuwfNKfu+vcK0vDS//iW9xdzx2bGR66IQIULjIM753789nNKMz1H8GRuRO6sMnsAxoQ1oYgR4ogigeUtsiQaIAI2Bm/8Fyu7g//J+x5UvdJ/+7cXOfcX8WTi/oGZHw7UwPD0KBaEwM1IWM8PIoUizVdIGaWcb/Tr7lD2Obcfds//YavutvU2jYrFKh1GjFXvspZ82yvRzNPX3kviZmWDMGSrPJwZzAoIGwBEYhBKIUnc3jcOKDOCN6cX/WTPfbSzr/pu3rnU7o/w70u1098+e62Ap8n42WQdiUliMMIVIyADdgxAtVJOGAzD/bc82d/32d02jpd3/KF3/nZsMKUMkgEDUflAu+62cOnO3tzdnnvYrPpRi4Ojw9UAEHISDswAwsNFSGwR7yMcppBGVk/zrXpZV/+u3jjkzj5ml9zhznc+Z6nh0DPQ7bIwbjZJgmlQBAKO0RTCkGkyKIIMLZ/4ps9q4XmuHD7D77Ts8PFq0NhDCA005UOXrr9QE/QLserXN/PygAjIsDOg0jEwUgTaACCIUFloDxOsDnPnEOLPQ6Y44fATmjn3/or7w30Sdzh376Zy8W9XC6/ifvbyxzy3MLgNDihAwQOEBCyNioiIA8TFkZlQedyc/CXPuF8aWDOuTs7/s53+opz3GE5BVsFqTl0edW3feDy9MnlMm9/9/p7APbBKiA7KJgQFYAB5RCgvK4ROCG8q7uZQT1ubmeUdt+Zw5AXc/1bsgMSn75dv8VbGneOveEjPzY9HMVHhCRWWBCsR6rF3KjCWAqJAIsGYInK3fV6+Rsf8QEexgkv/a0ffs+Zp1xYaoDk4VVub5+95ZU7uzmvT955/bw464QGsGXACEkDs8hKKCOoJQIpKozDccajEY+Byy3veBs2YUh/5T8HEpg+fePPclonbnz7ex2a5NGAOCCYKsIFCClzlEPyNCTMijGjmFMhh/G4/+j/yjtYdlnmrf/rr8LWNebMWMBTvTY3e86z8zavd5fbv38/dBxlFEy1mEWB1j4gFLIVFxlFKF09SJIjZ41ZBlj+XU6soYb/5pq2PIX/dh9xf1mOI1l+Hh8YDwMfAR0uCEUj2qwLJKBBbHH0KLmjAGtIkyK5cenT/ovrfdd7JmaOJ7/jdwwyyzbKTjIwNz396O0uz4Gz49+8r3ETDDF5WApmIA0BWcgIGRUBbgjsxAfOJx444zFHMJx3PxdnZZzJ/o3fP0clPn37QxxpzOL1a71tGuY46hGZYthYnU5IJhtQi32As8nEZHIGKMTaEq4F8eyH/IB7jmmVa8Nn/qHjfgHMkkUe/WqX6/V4xqtwO2/5Mk5ciRQCdBOWySQngtWpWHbZlMlkYaYHswfn7KTmwjSnO9+4vc5OJNT/OzuuPnWTr/fxu8aSjJdvtiMHzgCieeEpc+0APEpyNASEqYNZyGFIlKOgUGAOYoc5JtbrH/rk17Zzu47n/XH9j/7aDYmwTm1By2dc9/Ls8qrLs317f/ZCcy6nAJY1yApi58AwMhOoDnM0AzCwTIkVFVcvHKN6jODAOz5cLswqAeb7vzpL6nraFn8ytGSFKz9srxNKAG7b8gpMC4SG0QJRBcF55CyQsbBjNSUC6OxGznC3f/X2Kud1r8L1yft/3HvdDUZSyOQbXc95ch67913vn/znnFcZjvNKISiCVOFBEQQZIFTUA3AgYaWcuB+HMepsLnmeX5/ziIYCi/kvM5SZp218668Vxx7jehZ9R9ZJ5kE5DHwgRmElLFBCUGDiQJOYJWNABmUtEhjT6ub6rt99NwadHsfd03/27c4LLJidyMjTt796nChc6fKP3nc5dhjmuAoKsKjuCpsLECBucq6pbg9gGWBg5BzuO7YThwN35cIPo8lteyo9/80UL64nbv9VDWzqHHNz/ejvcE1OPBgh2HO+ZKtgAlDDJgJDCEIDBjUK4QRQggZcmCe3L/vDvu77b44h7p7d3e/Tv//9r+ezA5Ax6nzXO895+60jvP813/tXmANALwiE50AnxwGMHAkDBhgNOwT42AAswF5nee3Ja6t4mS4zHsfp98NTZiWrGj74T81F3E/WfPDV3rlDjkibw4+8nMflmHVQZGzbUkCEFmAoASwlWB4PQNg4QIABFxKoTv6P27vd5XS5OX3rX/j1Bz2LgGWO4yN4ZS/HMeNLd09ufzPsVkQUwFFxtG0LZRRAlYLysCEgoMXw5EK8pnPMXHbUg+rtT8/jOIFOLF6u/xnnINEnaj34b64rBq2y2Pe9pdNV2CBa2IQqCqcgpoUAt5QpAeUsYCge7RoNIKTO+eF/8LUrbLsLd8fTX/439Oismj35mvfnl52A01ufvfUvf0Ucq7mmEIiAyhBh0p6KkpCSVNIaKNjOsHbX7cFMGNtk3+iGvR6AAZmub/sND5XQJ2oPP+7rCGQbKNLtJz8bmMtc5jhURnltzFrCwng4FVLIRkIPmTFIpUcuYkEowHF+v+9Nc3YO49w/6ft+ETLWMfKNkfeMk0dvOc8/LnuAiWxSgAkYIoA6YGVgCctoNBiIMADH++48nGC4OW5Y8sfXXliB0knL/x5I1vGE7c/udUqDYZmIH3A6mleClorzvLYzA7LQhMEx4Y6SrpgPF9IsVpeoUDIIJP7QS9twNGfN9emXf9cvOE5lio+es5v3n+uh18uT330dOVmMRjQEEE5Ig1iNVFfYgZhgRRaBSoD41x3BMHk9y9m3fFe9tNAbhTW/8cv32cN+qiZv/1RvVmMYEik+HWewiyKoF953ORtAYAApIIQJGRBwA0YMU1wGRx4diHbhuL70Q87jcNtNl8sXftqX33DPvXzVt90dl7sP0I4e+y8+N7EBkQUGysDDBLEMk2UZQUgxhmBOAJWkg/9rL8fMSBwyefe13nqaB0BDmcV1/RfocZg8UYtf55kRCGogn3xzHowlTu4C/5hjlhWj1lTKgASQNhnpPCHJyAEipAWUSbbZ3/j02bnjarA3fuG3+Rft3c3dy5PPbuaBcvP+n+WeYxEkUGASAbQ0akHDuEBivL5wEBB4wvLnnr6kOZrSBb7T4QiDSCGQ/e8gi8jxNI23/egaYUgq9LCXvvorxTRLO8MEn6OnTMCJFA8F3YAgBCLmIKkCKESJIR6OOHLzx++Pzam57tTTf/rdf8/90y/4WXnFa+8/98TLk7/x/pF1IBRSRCAREnlUBHACKQss5NFIQIR/+PKzgZNmWvC4fDstXlelm/NLv/KaoH2SdvDfzt2FgTYwYVE+da3LS0tgcAy/j+MGMJxRhRCimccYBQEhcBwkHi4guA8E0Zt//xufN3Px3JXrCfO5v+zbfdfv+CeePL319uYSN1zOl97/O3eYBQVCoCKtTsaIVh5u8nADtQUEChQWur8/4T0d94vX2pu3Pnly+9rX+3bI46WlZR3r/G3nuW/ME7Tz9lt0Cwi6YMsk15/IfXO9Hh4aY1c+j7AALHBD9kEbUBFbGGBEESIhQCFL1BLn/tnX7r2nNtidm6fv/Zy/+cWeIx3nuZd96bj86tkalgABQpOEY4kHQNQAbQALICDFo4sGx+W8vOcrLidzqBy7cN59/ZsWHwFFbY9+1/c/GJNjfHrGD57UgsVWoQW+2UsntttZyqavQgstIO1OhDtM45kIBtQAGFSAtPFQEZAUYb/qZzwL2zqCvV6f3B4Hr55Xtn326j67P77yd4wuZLQsYBkiK5o2ZJCBUjrAiRVSBhMIR/LZ9xyXWV32CdvVH4kDUbXDBqTw3+5InPTp2fymTmgAkWFnnbq/fI9ncz/HXG4UPRze94F26jy2cBgGMkE4QgZIlVhQZpINDmQpFgaYFTzq93J1j7ET9+bmArrneLYxetz8/AVDSB6KEYDDwoS7QyghokDQLCrI8NAk4+APHZfzcBI82+6ffFPANaG2i0I89j/jw/M4aoxPzr72h3HTCWdMW0yiwC9mZM5dKHIPvuhyrmPiQm0qU1kJtDAsVLpRKzuuyCoiJ0QCnHP3kT/uNbruuXTO3bnLeZ1nr0V7XnfPl97/e1gdyIapQDwhlolTmqlkGSgLPOWEMR5PAFnolP/25Zdur207sfd6fu23njyalQAOM8yrf+4+N6J9cvbruNLRXrQZJaR19lPefndcr4ejwgzyV3agkRo4HAHESrEQZcJRpkwYkmBIgCNMYTsO+U3ct87Q7hxy7nk9z2vnPjtvj3n6CxgeGoDYmgd4DtixDMjRBjGuosz6EHogRhDW2P6bt9ztqFDe7tz/eDaSBafaEoyZ/+byxaHEp2Z+Zxx2gQRBATnujm9+dYgGrfXgTx6V5lhQQSAjICkBSUSMKIEUpJUgETEuN3dv+Yw7p8089/6a2P15j/QB759+5X/BFRYFShzcogNLJAicA4RGgmqYAgUiRBDHmL+/N3uzgS1db5/d/sB7RTnEAIUiyK/4/tfLmHliJt/zciVnBgiIBfSkZ//xOjdHbUJ7neP/4gIklaG2wUIADVWLnADCiSgUQ5IIcAoIee7e/T6vu7UzuV53uef+/kq83/vbX3DOImAIEUwgkMAKxsMoQRUYNt0oEAEigD35rU8+zOWs8WjP85Vv8PZTCGBBiFBKhv/e3A8u+qRs4jfuHIBsCNgAMXNz+80+8tXj6mX0fs/hOM8v+sr7YzBksMAZgJgHrDNOcYAsHBCAIwhgDwblYcc88Z3f6j5378/7Xc/z7v61657X+2s397z8xb97loB46IN4GAZMwEAhykMFA6UB5NFAgLbpf7u89yXOQ8/Vmyfn8XO4CD4ASqGK08N/HFfm3MGnYgrLu78agshgAUKZTvyU6zlyGW84uOH+9vrPLnsWZSFAGyAEMBDAwrICBAGcPHSDymoXnAJ+BzEjzW7sFquNdfuzXjmy0kUCUiGzIBdioahipXNBkBWyx4QCcC9f8EVvPcNzl2nY8+2fcZ0pgKYVQAYifPz9wz5o+lSsAH4WChG4GMujbnP04+eE2XXtjDn4A8tEEi4RjEDhNkvYpjgJgSS7TEAMlqAwEY7y9T9yO9doO6PlvN533fPVJ9efbRo0kFoJGIxpgowkohnMDAFh4PogCnVp8Y9we3hBKPb+ePZNn1449dzFdG8BUjbM+R86iaPrqdjjP4GiBWgWd6KAMc6P/Gp32Qh6s078t4cMg9g0hQU4NNsg4WABAWTgoBsPSwQMoLO9nPzu2aPZYZhz1wVO8u7J7/43c8+ACxROQEiboBXQtEEYoluGKwsDkCiQeNzwB17y0g5ztgtXfynNAYeylkXLbReT81+Xuevy4vFkTPjq77xvnWNp7SgJYEE2fuN5LqdM7kFX/+V7CQgSQRKmOEU2TgVWs1nIKFiYklWNh4PoERvf451XbPDcVU+u93vSee+nMYs1rJgIK6u6sSiAqOGDNhCINbQeWLBtnjBf9rkvnzXn9Wq1vPrx3/S8LmdQUBoKAzLiq994neW6dp6MBf/ZMkOAgyVNqoAcfcbNzKiC1Dy5+wuwYTxqgSAcoAMTkGCDJQyFwAATQgHCGejMvT8GpEJdtqQV/uWfPACFySDAEmRsImkpm5bQIcJpGgMECF3UOSf4zb7Mng5L1NNnP+zgpg5wlICmBGnK6X91U3MkeSoGT791ByAjbE6CCdFV7v0RHzh3ODh21135rYCFbLQjGyFnuWuYIGXrwIoOEggFCCzUQSnO/lx2lr22Z9Oc1z2pgx/xDJZOEkFJlYcrQjhogqObgQgEqSAgMUiMwu9629EeXOM07vMn084ViDUqII9ue+x/mnvROH069oNv1pUz2sYwQVDW68Jv4r49cZGU4x++NoCsIw0xrsAoM0CwAiFABlJEQEOgDSgoIec7v/m9HMcBu+cqe+7uzSs/e07xIAgoCVhgeChlEK8/PihSgoJdSGh3EM6//t6nezM869xd6D3f6sOfsbhCA4UCFOqya17+s8/Lpd4+GfuppzvtIQ4kyMnjMzfHzfnWH3K9nF69vV2wK6/+o/Pk3AACECZCYEkFKoQBUQJEogZYQBZIWOG42d/MPVeVIbbb4wJefs+/OCISIBRcGB4uGbqCyMYEBCobMGbgkIEOnGfH7z2e4i6v3b+29yPP/vP0iAEEK7QgFZCs/+m5c5mup2M3X3tFpgWKrZglIgnjN3l3M8NVUZDffISMhEMSkCcFy7aiCLAVC4a7YlBYEVMASdT1Uz885UEP6OD+9vot70GEdaWIIaJlEAl5VKlpgYDNKUgGBBEQOGbe+8ff9XS3pb3sXDv/3U+43ghHJIEtWmiBYXn+yi9vurJonoh919vDwRpqXJHVNVfY3eN89zd4dm2YojX8s+dxEkKRJGlGG6ibkW4jObtwMrMnhi4jSQILFrR+xnnouV6dGs6dV2++4OcNZyvT0SJRLDEbwMK52S6JMRBACBJCPAyiLU75Gb311dg2N3x2/t45AhMIEJhhkGkzav6p3wxrbTNPxH4di6Cl2xGrOYTkKPqfdelEpgOOOd73/wzMBgZkBQ7KoeRUWlOLMcJAgwi5oMSKIgzC9Wdxnl0ODuDoZpzT/+RfIg8FA4MSGVnQjll0GPZBCiDjSkLrI4Y5Wr36Fz7s4i5wo1337lt9ytqAxJlQbUJolFmk/U8wrtyh4pMvuf2aTA/RQNTcajJSkm/0sa+6E2tYw49vOek0C9FhYVGCQA1AWatAsmE103NZkbKAbdhPevm8bLXgjtte+Lp3o0DkFIpHFjULWEcUywFgeLIIOyTFBCwZshH+6lffslude3+eDfy+I9ozNg+BDlIYEcI25286oHYZ0idf8G0811OkRBJiEExZmM79z3eOc1EPgstf/TzX0yPUhWgNCVwFkng4Jg4bASuJNCSAoIEBT7/ZeQTOEA3H9fKBb0xQ7QC6COJmDcYIOzURsxFHtpkkoGtgIiHb5QO/6603yW4LzLz2iZ/UcuHImbBIYCgZKGUxl1e/9tNjhSDlyXf8CprTRAgTkAVJcjDnte/y8a9chnVPcLjhx1ziCAQHCJVJa2LzRDeq0aoDTMfcq5KaLhKacIGfyx5GoRXdfvavgEUGIoYkOGAoLLZB14ARWx9CauIyGAohOvAzn711ryG5617vf8d5DCiFgTAQAim2DOuB/+q1gawWn3zx9Ju8Cje1oAKE7gDYSqC1v55n3k7TqbiXv/IVlyoeT3m4gANBF+gQlWU0DCFcDnJsg4FtQVzm7lvNcTlggSn3evySf3JGAmoh0CKIpMjDIaEEAeXcYVfaIZAWBJbzeP+fecu12at6bvXs4779/R4LjEBIBBGKNMHN7j9pda3UHM3Tr294czlkGFggMDAjtJSYt/sDvvF5A8fx5LJZT175ObOFECQYMEIQo4C4AgLxusEonKADC+NAOJebecs3+wA7x8VojpvLDV9TEBZIW9ABFkMWDDghiJDHD4dGcsAgB2pXlp//gSdPxznOxZu3Pn3r/u+3T4cj1oRZoAVKBUqZxQs+/uL19aS7sJ9+/UJAxkoiWlhmeagQQtsf3Ssz48ycHvvS7/vi4xwXINkTWaAsjUcXeVAwCy7IAjgExUBGUGf3/Cx2p3GMk73lI44lBjAcil0aYlkTsYOp0LbYACMplqAQQnQ73vsHPtz7o8DmfPX+y7/+1zgpZIJoQFsebQmwGHv5Z/RIktUEn3j57XMrqcWGBgYGrHAF1mO/5r//2h51nixsznenaWph8wBmSZC0YpngBCSbbDbgBAIDNgIS1I7rd7w8M5tJz+3ynu9x3NmyUFSiUTqND4pqcgDHhgrioYIx8mhI9/N95one03Z29Tjv/ptCQQwxWgKFokwZWJb/+L46K6XaJ15f9aUF1hgsY5JYbE2ahUzu/9BbbLyIo/bk//ybs4UHjJXtRBbDCg1hBxEGFQPMASCoKYIk0Fz2re9ub+ZYZhyOy3/1X9/eO8AqNIEgnBAEaIAbuSViINhODBgPczW9/dt/9e13nZflftny/d/x4+4HaNumWHBad5lAR8luR379R6d34ErNE6/vdzAjBDNJQkuxuFMiYDu98w994Hq2UqnHfN+DMJJGY0BXYCclElnRFBSkgIUAZHndhaDjW3Sc26HRObf/wRfdniAKCLSKzWLNA5BEYHIBJgNQhADioS1w/sCbmye3x9nV3d3zlds/d1XOHIUVCZAVQms8RdB5+d1nrgm1M0+7/Omcs+xKwTaE4TDiIsjCTJf97l/3ehqQF3efftFvH8tcXJIqBkkEN2NhaYEVAjEZ4SR0WEthmWX4WTAj2F6G45PvzkWkU5CZFnBaRWQnYQochoVMHrXONcRaF6c9fsYXftTOa2dFe90+8HPf4kUOgeAIMWgZihZWwaHh3zMeErPi0663f8R5Wc8LAjiUMkihSoACM+3fuF6XxtmrR/rzv8wihkjAiFZ3AQ6FQRFICVgwCA+MalDAIU/Pr0nhHDjN8f6vd3tugEduAMPDGaBAXEOFLUZM2AchMwggojHzj37vR8xxd3Bd1+Xu2Uf9IjQKmMDFFbDh1hYKy2z+DXPFI1SeeP07szkHpiRim24JWBLQ7nm6b/khnLBC57nn7f13eeWYHR5PV4ThnIEAZNFlBCkCqUWiTKR4QGa+83I3VB1779N/8iNvtgjEACEWghUQRAh2sgCDAcLQjcfNgH/16e++vbs8Na7ndet89pfmOtsmGxPN7hACyq1qS0k//MrVIZdon3INP2iOEQQQKBxgZjZAekTnOOaG/+yjjmO83Dw5z9jr/O3fyhlXUBNmXIKB14GBBijw0ARnYEUxAKmA5nLT8Q1vlvM8O+fZ9eb3/zeXPRGC4VGdEIYHiwuxTIMGCSwguDQ+Uold+YVf+vJx6zncH3ddXro9fu4n7oWZGRjJU2ZAKLcdpO0SC3e/7FP2diR5qiWwfFeyBYslFFaAHVsSRJCHPfk/7u+POc5mLnFcjl/45ds5PDQ4SYY1aGBbKGFRIpLaKKEWIEghBlg+7V4u48S1+e7vufDQBTZYoGABYmg6c8pi3SSYEIIBKMBx43r5Y3/mnU8PuKznvZ7XVz7iV6wFtEvJxAI4AK1S1LYj8B84DaKDT7QCeOnjPILdxp3cDFtiWCZwOdklCq+f9Jn3r3E9hGtTN8c3uIzDJg8HK6YpThim1mKAzLVVRWXB3YIKICD5bu6sZufw4c+O8xEIYxZGmLOQCMddABle3yCERVBOaPHu9l/9tHc/7WSvu1vdv+L/6CCwDkSLuLSkDm6sZUZxW/5Z3HO4ktX6NOvRr0WVKhCLWblGicLsrAOtcjz7hZ92ntf7Dg7X8ekX/EyvLS7Lo7N0hnTEkhw8XCSxDpAIRhjHEk5hbIGvv5c5ljn0Ce+I+5MAiRCCMgKB1KFcYHqdeBANBduAI7fnd+rJU90zdtr785Vf9CkwK80259gQD4c6lgMYQCoBv/pxLxLDYQGfXPnIjydGgW0cJRyTcmApxiH0IHD/p3dcj9mry8Cct7/3b16uZ4lutnvITCAhEmlMm6zKCZkE0BTDMCzkwH7s05PzGNbz5Wdf/eBcJFSwgpwxKHMxRnpgATWQZJCCGVyBH/N577ptr7jXute7r/8L7uLRRWZ3SJxkKtpujI0D2sCveZOEDrsAfXIVCJ8OJMUIEFC6DFCGPLpErcf15q8/uXdlBsPb89///LmGlYETLiLMIogL5Mw1S9xEYFFRIEGgmL35cI5jGtm3/qvvfDNnD+R84MAATaFCQDCeZRAK2CKICzC1HfAn/vTHXM9D292ux/W1t/z1a5asIzguEBShAU2BltjChn/rw0Fcibl5ov3ko64mKFuATEQDpKBk6WCO7Vw/8Q+d1zk31J2efuW3fu0lOxkVxDgIDGyxxDW6WNIoJ8BAFFMgiCnw8fdtR+Pd0//x51wOzpSOIICiUAHGlLJDGZWAxAABYcEZ1n/8o99yc8v9htv9q/fXu7/5pCEMkIBhAARsKSglhVFI7//ZnhamuJ5qxafcLhQaCFoJ44aggSEUjx5zefY9fvJ5b7vLZej+9os+9b2uFEAkIAXDDog0ieSAuByRu4nC8lBKGPkmKsywL/3G/3zSBRF8oCEsjwsELj0kCcMhiDAQODv+8bd82zsOh+uey16Pu1f/5ieeNzfnmAMFCCDp1oAAAwihKP1Fd/fmxVqlfaolP5RtlIcGRMNsTLIAAo/oI6x1vOrvW/lw5e5IuwjzSz4suAsoUgARDgBDBBIoGrDWgVgYhQU1Fr7zMrvndY/hmzHDAMEiBCTQYyc5JoAKASgkYohEnnzYK97UPrtc5upx++TJ2/5PPJmwACQoHh2H2oESwSkV4Lj7vlnndc+SAXya9RncWARObiGBwsBAOC23pbUUgfNnHi4PxvZso9fv7tnMytgBkAFK2QghOmENKQgcMqBkgZYq4GPpuMwwxrsBCqThYWiBUHAQnWAYkEE8HvFoALdvu2Mb73ZvcW6v7/9sDhMS0JIUesCxRgy1DCWOZZD+2k+5cx01CH16pfHVznQX2nVdadyoE5YSEepAqwxh03V+wpuRZqcD2/ukU8f2piUUyxgPXUWR1EUCyRas2ZRyZPg4OzE9l6dBysOgkJSEtAVjSEViUKSFlQmWAnjpHffXOuW1cy+7J/svgQqhCKmTTSta0mlnKiKbDM2ek3/+3q0sepSn2NFbb72EY47MDkoWHDBRWRCRLimLfcaQnzwWZ/fWBdFVka02UKGILEjQ7IAQpVUG5IzBDqBD2b7tZo7w6II3LEIbASZIagABgQQVBgUkThAHDN4wT2+bmdXl7nK9f7ItM05kQCAHRCq4xWKWBaarBDHuvwxsw3J4mn3wdbkWE8XCYXCCkwGiA+DAiJ26Ya1t+7WcbfbYbWzX4RwaaKZQxym5gqzMtoRIKNKKLGc25EoIy+3QODszx63RiWMjIAsIgcsUCg6nWnJihMEqMOV2zOX04JLu7E3LP92CXHcM4RwYMMBGIGxABClAU/muDx/WccHydPt7YwNEDUB0qLhCK0GQCimgXRQzV87XpJhRSu7W1/dlsGJhkIAqtBXIIIC0MQCiHDIgnkg451yaUee4PB1mPQIISIg2JIRxwahjQWNYQAjYOBnkWXOcXBKvZ+PluPsA7bnrToF1LDQFkawy4G4HoAQsFV59157EJJU8zfp2WJsxysIgVbjEILQDoxaK4q7A3uePrGPvYYhA7eujtEITFJi0gsMYBhQtY9IMBSWVMZCCx8vnLs2edzezsiygkDSCgkLJUKYNkOkYIqUMsHtejrl1PQ4PrteLcmcdyoSKNBUPRRmAqBELCLQia/tPndlt8QLzFOvKV78q2gwtw8NQZVDAUEraEW1ndoTZyqe6DnVlSVrPnzWNYFspxKFAEgbSGNghYQmMElQSmOm2cC4CtcAQUC4QGVAiICJJEfLQQB5KHK843N3vzcJcLteut//m6cmRgJzEIjiIVAABFK0KAapI4/6Xny5WMmPI0yvh9oKCQRJtYYABRFiqgJkCJoFhXfLq8n2v5jxwJRfWZObhyulmtxqGAVBAYSFgwiZcsMmdHCxAto3LV72/7LPrs5PlzplBXMBERGRxkZYgWFKGIMQFQe6v8AWfc/PWt778sl68dvtkLpe3nMeBKOuBaVngBlLrIwIFCgW5dce/dM3Ry4u7F5e10qdX8eFvvYcFhKFisI0qqlArN5VGKAwSGDY/ee47luturh0n55/eBU6lA5qCYCkg2ZBgwQWXJlmw2FLCd+3hjQxdbpFYGHCFAjeGDsIBMIwgJNsYdqPLAf/suI/k9q7L7Hheb264gsQsJcDCxiCEw4OBUhmgllKAcH77S+NBZKFPr4aPH1E7gVMKaIYEkaYWKO1GKjNkiC1MAhnOXrIfNlq4vPnBOImUDFZDEGjWxXjoEQw2EbtCY6MGuy9xlXPnygUWmAB2iLUQyjB6gKkFkDiy4FX34Et3r1fO5Ql1nmd3H/YWbFkJesRH2XKRPR8RKJAWEcfS6Vp+dd+FtbOU+uQqPo2OIFmOtILNAgsWJ6Wtrj5q2gJKZcIlL+R+d6vxSL/xewiFBgLkQo9ko7MGJBLiokqA8VB49xzL3Dje+si2wYANCC5SuRBsCCOGEKwQXOAfXJ5eDha8veHYOXrXFYshB5VsAxICdo5iwRa0QOlA2hoof/nrrnC0VZ9ewTeGpBmsbQ5QdJTSGWKwCidIlMoqdJBNu15e3R4TqwmXN//wvYoOU1YgQUcjgDShAJTVMBYDJuYl9sKc4cGIzYwLC3lCMGyDjI2NsMsCy6NDzBn8Ky6vzUW3cxyu127nbCyssFyPAhsdmFCltlBoicUB06LlnzmvGGoszpMr+XQ8gwgaYIqBCAdqqVIqHAhlEAZGoOZ+H9c2PrhDenRdjn/0gV2nWgYGgzZWoMlNSsIQTwRx4+FuT0GFy/UlcJNHc5UhKGUlYJNkRGoI2FU2es/Fazu5czdze3Nz+1Ec7a6cKmDDjrYTUAHhmRiwRAEMLQEov9HrEo6C5MnV8hF5LIowsiUPDVjQA63cCBREKXGYzKSfnDsfVzYeWXg997F+5mdzwQiLKh4VjAAGEhwg7IAQ1LDjZl8m5EjeSjNbwSAmamDrBMRIAGYOiTHk8N4vuT0ud4tejkt3bLx7mJlZLAKIIRRpBwSXcbHVAru0FBBsuMz3KEYOOvTJFe+4PcUF0mJAKVJ3AjAIGChUHCzSkTTOPS/mcnc324gHtht//CfO2sEuBgIDAxQKmcGGMOwGaEG63Ow2jux0whESLLJAKh4gonQij0dCWnr9Z91f7w/uZjtH7uf+8rW5h5YdREQBS8DBIAYctgixLFEFAdru/fKLb1zDjIs8ubrhwy/YHryuSIVATLKQQgGkRUAQMKZiPvjw7tW3/bJZL4/kzn1OLrnvT/3QMefQGRkBhEBniAQJRjEYBhZU3DzeuU/2HPfysgIhDycUQCDidT2EQBRbXPXkfz7xpftn++rNenWeHLz0L8+b4dCBxQWCkocFgoHhULl1KBWgN6LH5fLJuQt7TxufWMHHUx5VFrahSuZizALtYxQpU6Q3Q6HA7A7fc57TfbfPU7rX4eW3ABSIhSWAggXCM4agR09phVge7s0yw+lxfcoJTj3A1g3YAKFHgAKIQNuB++N3f/+4b/MVr+d9XF/zGx8t4MLQEIIiLdIJIUTgpFDAKXagpVShfNeZDnoQeWp98mlJKCag0gYtsLgNgAxQmCITW6DQgpRh7eQvo9t1nU0HOC7+/k9WvGliMMAIYOqgggB1BAQZ0EziynE6zMmjPpBVaqIIK6A0WJGF6MrNz/qZH37tkTuOc+89P/4bA1FDICu7wQJDcojVqkHhFBgTaooCFrL5C2dlDYsl9IlVfDdWrDLd2hrYCZyTiYQCiCUVA0MLoPJoCPxlL1jDblTieeFP/gyFJgQyYMiNJEIBkoCkeCh0mWnnmIO34jwSEDvFOTIJAgkJOzDQaep5ufsP/tOPecmjvV4797K81sd8IgmHrLREI0XRGhXgBJ4gRWhToK0gDILwT3qoubQKOZ5YyVdjoBoWkFFhh6OVqQmaikKxQoeyMFMIZaBQ7vuVb5tzX2e3UHqQr/1QexUBAtxIHSTVDdHAGHoQcWXo2RzjrmCxEmHDMR0sKFBJuCBBeNB6vXzhN/urH/OyinTudl79VEKjJKcDw0dhAC1tZRkMWxG5iVJAhinlL58pE9naeWIVb6NqBoZAEKAaQUWACNgKYAm20ya0sKOCp/OVL80k0VpzaV68+YedTTrJwJCgRRB7qLGCDiRyXNfLceiKB1qTAyELlQAuqriMCItseO3m//6WX/ARx7Oz0by/n+N8+buhYbICUgpZSrpMpbsmwI5UgEBhsDCk2mO+xHWRCBJ8WjXztlAoEQQBhyRgeagWCnmk6lBcCigs2A2kffldHlcrkdWzD+f5B1+dpQ4mxSEApxZSCcqBgEGCLvO0TY9xb1iHhyHbgPhgGxaBgRYaIOPGP/Rd+sizp1PowWv4jo/jXsE4EMIgxABhYARnFJJJbW9opeE2VpF85aVGD1IG86Sq28uVMgUkWiSaVh63AMF2bgQQLeV219JFd8n+4NevNy3qHGTVzNf5xxNSFIAggHKK7CEOLVZALjM7otH1jIgNFEZI2UBLHg0GBkJGf+gveuvLr13gHFjn8pV93Q/vhoeeEEkDAg0Qj8ujBSCUiC23QgWUMicvL+sIM2YOmXlS5Xg4B0DLo1O5g0oIgbQF0ACtMBREpnSJCLPI8YK7F5dXd8fS2Xu/Pvf1i7/Jf3O9v7teER8gQAoHwAABA2IgszRcb/dqzpMPh4ihgKCAkUcfBCC7LDLAT/xveclb59pl9HzydL/my7ZAXEgIWSgoEiJgoVChQGGoVKDcDoDkMvzir+91WZVMlCfV5zvU0CUjInk0WSOAulWAYSjQjtrSQWinwODIHv4DX5zuwIN2R186v8ev4HLFjY0IQgKCgK1tITIIM5pjjDwJZ0oJSoNiReh1diWD4NP/2DveSmVeD242Pr1tIAkBWQaUNCHwBDclKgWWKQMwgNSNFKDCdyVjogdPrz9qlmjnJEER2KFtBy0AXQMtUBCdwoxWqWgYClrd8z3fcc3d9EGGZt09ffIrv/erN9eF1cAWWmMlIFSG0BUITq9es5mZQEJaWKkgbQKRh6Iobr72rf7WOy+2y3XvT+nu5Y913JYSSho2qiEiZY+NkUVAdoEIYGs6LbJq7YBz8tWwinRs69Oqj2NxlYmCqJrC3GgW0DLcDpCBQirA1IbSGqQAntm/WN+c91lAuNmX3/rffasvmU6RKBFceSgL4hTEBCCsc7vNONcbWiAbnBwGDFfi4QBhJvn+r/dP3nn75NKc1/O6r17Pu5uP49R42EoI42qErbSTGAznBoznTpQScABpO1QMQPgFZ2QAXBcK+HTqazZKi8ekC8KwKo2njcEASxiDkKldEwgsalGddqJSDl78St94pRvr9bjylr/3Tb/wmNgz06kHK4hqCwiVILszdn8743C54uApQzwuD4cANoCVqe14z9d6zzuO7a778zyvZ9fz3TfX89pMkqQUJ0Mk5wzoqhRxzAA5NkgjFYhFCbYwEPkVPdvoysxeAH069bEhKpSBLNYETAyCQBBKFyAg2EoKxbFAheCMRPie4eGyYkbb6R2f/43+4mEO51BKBwpn0TIBMkSgnjuXg8jp5DjvzwngfGCUAD5QoIOw9Z9+3T7qnGFWr3d3rz3zk57OuO0gjYB2FI4xsMQAjEmsIemStdYhQ+nGihYZ4NtGEh0NypPqj5elZlPEGFGgCQji4UAbLB1sK7ja2iqO9MbONghXvvfFvaDo9VzOt732vf60w84RWcoGHJKHiWWIBMxcOLWjKc451CUHCpAlQCAIS65z+V3fYt7pzalx7sAr5ycixzSBxqOBBEqIARWGycOJCSYIjYTgAgsVCeWrOSBJZq6Qp1UfRaIpCCFIyKAoiRQEpQUDGIAtjNAqyI1Z0H2m/PJiW/bqqcfTvs/PqaOtQQEMiGGBRVHZNcezaceZ835PjRhEVAREHh8SPLfDH/uznr772drec57r+cp+S8+DmZMI6YEIECjSAPkANB+87gJVbqsAtkJnBj5kpDUcB2d9UvVJZoB1RhgEAgvEoyEDUFQebWlZQKByW6BCi2vlMr/so5cfXtYxOm67l7f89v/o/rX7u3E7YR0FDCYxIGAGoePlvem6e94+beQYgXg8QmCBhWr3vLD7m//8O47X5vrs2bP3v3r/2quvPbv85etFikDIBYwIiQAChhEWsH1MYygI0EILpZW2KvDm4XIxGkyKT6nedh4EFrLMSQs9tMVtlpQhtAMtDG0RbEuLjEORoeW2sM2rF/fXw03oNHP4W37YuXTm7AnntoBRAEtALOzWvXF4cnJxAQKoAgICqykB9jyPX/mZb33b7Xmecfras3M/4N/8xi16DEQyEALTZj6QpYh4dACWxSVt50YKUFCmhjJz5JjG9KApVp9O3e61WDQ0jpBWK1iGIGC4NXYsAjiV0lpuQkGVQYbCnPu73Ztlc2Gvl6d/5rtfz4ZlDjkUziWAyBBgwpFrp550KgRJArqkIqBQiZ17/eG/4V23R3vC2dXhy46//4kNJxsLFK1UpK7rrq1YxgQrG84pO9UiDjO1AkiZBAxFtIizci6H9snUcXOYTQgxMQtcYhAMm5WhkwK0LKAdrBRIdIhuAyNoO1ZWjuO7F3s7QpnLwUv/x3eVe2MzaJiRmnASCiKpm+PYdY5LlygMQoiGFlgBZkLnst/1L37U5XTw/v56v/d3r3zEP/m42gNE0V0lCRIpyXETlEAaF0YGZCpijTiyKTaUto0vc5BhtOsooE+n5JhhNxwkOAACfEiiTKZHZUQpEFZ1AxY2wXFtqIMbZwEDZ37jvm7XTvLkPHjpr379124XqiiCc5BqJ1FRirnzOk5yyVHWdHEdiAlkEdBd3/fd/+93HMdx3HdWu9f3ffw/f9c6u2K0HBaeHquwHIksw8MaAmPAWogAqpVBMRZSpBL7yulYsDMK9OnUuVUqZMhDm5CAJHcbkhECFJG601kKlghNCVjaJUvCAvl+7sFqga48+bxveXeB0WkGYFCUyWgB9KS7Y9Q8rrfY2SDYSgRB4m4E+cXf/e993M2d58ky1/P62ge+2j94CWZxcZNsaGKgDhqsCTUUMNwgnBxo5VYI2EJoBcKjFWLaoEABn0bdnwzqcOopFLLCAgIoM9KWQClVKI6ApWiFEacwiRSKODPw0b6yp8uy481cLn//k9937EbxeFthJAosewzT1TiiCzsjC+mIkppsM4DX49Vv80/ePT254rYn3b//6/2Dcs45yRFEQAOULCihDKKEEXSTcI6k8minFkApj5p1l+oq7cAShD6NOu+Phdk44gAQtByQLQgCVQARp9RiaioySAkQIe0WoEzV8j8dd+2FWtgZnnzRJ33x2JmABqNjJWDsNLp3x+GBwgUDBGEBHAgYy63Ls2/4RR92IJd5tg9fu/sxf+c4h7iOpG2C0AMwVCEywUYej5GSEJDWzhCJhRYpbWXuT+ggUQYJT6V9CVwM3CghRDCYBuSh3BZKQSlWpEIREAq1IKtyG9DjO3zO0+vcHx43MjDHzVd+8t/Zu3vbXdqMwhkeH1zw9jWvr93dPTtv6tBHgIF4qNBuZ9v9t3nPx12vz57dvfraa3PMCd/jd+bT8rgsD4eAZXgdI0jENlxeVwgFAxUoomXAogWnMvPROl/eLXMc25kB9pOpKTCEJN221iqgWWqj0Gl5XGobKgO9qdBHKJRSKFNOLecnfvaTjg6BtqEn7/8W/+iyLU44a0on0IOlNuRwpq7cG0DAEkG1nZszrly/yz96993NCSdXe7bn+7/Jn7gxBFyI1WJarQgIKwgNBiKogJZFaCkYCGPp4BQBRZrLQcGaCVCeTl+5nkvSJqbMMKvBtrioKVVh7NYphcKQTijYoYNs0soMsFnMAXj/Cf/49jXY81yo61zffvmmf/cSnrHhCnQQgDsAlesR4NlJ4NYAAqrHmHS/+23+7kc1dS4n3dt7Pvl/BTxjW1Mo2QxwW8DXgaYKo0DMAQiGBFpqQDNECtsy+DBLbss6FciTKbwgU8Ka0G47MSVHzC5pbQeom0yFmxoqA4jWtM0eaQlFqrMZ/ej/+92vXMKW6bLH3eX6bf/p5EADnMSACQx1jOrsznjcvwUxmASBFQQSfDLf+p9+1JngSUt95Tf4G57AseoQ0gBeBRnDZTkSwkVdcM2AZQVjm5gZRWCYETp0Vq12XS44qPawCD6ZunIM5eogi46yoNLSMeooMW4wNYOAzjCQzA1aNCaNLMlplXi9cP0q/+fLX37dXTabnetLfYv/7eZek7MZqAACcFq87v3EHue8xoTBEBSycIIU893/6TufXfA8l3PPc9/zVf/PlzwoDNhgeHgT1HLE0ZwNKDtCQxyZ5DDG6rTWJIW2XSaC2jAoOt2jlLZonT6ZugdKRmgBeShYcRhFDUBhKYUuKbQrCG2ETQl0xhFonRXZ00vcnh/zf798J8e2V47qvLz67T/7ZsM9BhLcehAk154NKtx2Z54YwALInBwRCz/sb7xLz7K8nt1/ySd89uVq5UoLRMuAACMQe0GkPYqMAbMMd8lgDCi3qrDLgFBEKHuRCGxFqDyZfjazM7oIEyGIu4iSCda2kGlpCbfWdHNrC7PaTsHstFShM0ROmPuv+jfnlZbarjEwN//OFx7nybRo2jgrGC2HQnrMeX+yHUSEGMVB5jo/8r/+sFX33O1q1/d++N/f8wKYqCC6CSlLLIoFDvrgBBoekCMIqYlUWqZMQBgituh9lQTIVAt9OvUaa4CEqITkTOqZhoQFoRE0lAFDIwgECKhiZ1UJEAzukLf3X+cve4cq6+4xzvkpX3a4CwsWxMhCMu5Vd5g9z5caT0TUTRTqbJmf/Kc+7IbzwKWzzle/5j96ygGCggQ4UoCDjAjyaItyQFIjCGEhMCAIWCRamUAxDHm9x05ZRbaRJ9Sv3W9AQguJAAGRwFCBjYCMhUIpe3i0QuVWChhsoUCgMwiX/Wb/+rhcbm6P0+psfHL3tb/svF6dCRaQYFAnj/fx9IIXb1+6VEcIS6NQuEc7f+2Pf8STJy/dUPfX89zz2Uf8tXcxEAmUBmwiVBkUGIQSBJJSAAoC7CKFgiXQYgkggDM/Op9OO7Mnkb37lOoqSYQIU8UCQQ5AaiiBsksqt9EmiODAY21TxhaFtoNABdPH/pn3bhwuwjKnX/Lp97uU20RRQHFuXNPDc7veTSKBKwRsO+zli77T299yU8DpXvN86R++LQIREqBlLFjkYbglSSIJIaEkG0HlAAi2YHm8wECnNj/DUeliUULIE6rX2kCltXApkFBWNiqpwEgoTGnFURgsWGCgCJBCZ6htWhBwOz/9s95/nWpj49Snf+cHDVuAgqduaE77vjbO5n5uCMBoCGkPYC9f8U3e9tJu5+617bh/5f5/fFsCxMOSHbYddtjCFJA1WFt5dI1FGM5HdDd2YYqAnUJlIDPG6/CP0jVYR+saxqdT/0bnpKADhMmkxpYBgSatbUmL1FigpAwplgqkbYa2lJIgsUBJztEv+R535yUM6/Y4vP0Lv+52mka0HRuLQMdjdw5uTmmhZEJwB8zzO+y7d6eu1zb3+uzPfsNT4iwEV7BBGhghdhEYBEOASsCTwmPKIaYOgYzUYqUU2yZtl/xILrNmyoHp9TA8nf7XFyCLKSlBDEpotYZdVSwCOhSpIBaLzW53Ja0NKA6UYjgE4f7Zr/++67WCup46Tz/rP4cTqmZMitXiFU7zKpcLIw+D2MDV8Tt9zlvPuT3vr1vX63n/yi/7DueyOgFnHLScCME6qYZtsCJQjYqrKGzjkjUuUVilUJZo1QqI4QeJu3FmOmS3fTr1DwH2HCkIeVRpKgWNxhaoAKVpAYFCuZXgAnERxEpAVCJSkr/xEa8l05Je3Us/+y8fgEEkMWpTV4o2GfBcXB4gA2f+1L/9cYf3Z7tcr+f99QP/4c99zZtsVRhJUFOY5XFzhEE0BAKGiWKIgQ5jcgUqQSktgaYDhfKDd0uPFilB8OnU53BwIOSELPpgVwwoNRAFaQHLmN4MQFqYoh1podPOKEAFWBXSbv7+sbsrBIecT179/l9xXHcny1Aimn0fS83uOdShDDlMmeNf+kMf1XnXXnfX9v4DH/+Hz5u5XmUwMIgBFtiJVh8EBRuSBRGYkhYEhrg4lRY6YhnaIhLCz8qUZVposH069T7anQgYaBDddcAxBVAs0CKCI4IAVkQhtCiFopr0BhsIgifuR/wPXDc8gvJyXL78W91dZmI1eSjjzGsDlxk5L9FJCwbVop/9/d7d3tx4znYSz17+ywfHzg2yDCQqOeCgTjyAFDSQEhGQxYAYSMSRatFiwJAalWG2/Ky7TGWOoWNYT6bmnzHHcKwWBBWIEQRYBUVAy6NhpApggcECDmKVWlphoKARMDR8iz8+T2eObs9tuz/n8o9/v/ecjBgb4u7Jh/Pk2Ghvvqpw6CCCM93t3fd49enb3nI0PKP12Jd/70ddhpkCOCGCEKoNEKAHgAk1AsTDAVEFBAJQkAKFQiu3hTh88Xi4memDMzOb7VMpX2ECZlOQ0MBWQYNEqQU6UMEBW9tORWYCAs20DbcFobQSj2a08/1+1AfgKHdq8Lj8lM+ec6YFGAnUvtzDi4snNEAQwJ5c5rv+q3ffXHdZzrvX7u5ffd+P/F7nGmDMDrlsQsgYbYGgRFIIEQZVAAEEIAELUGtLBYQC2oH58cuqKRgPItKnUvvqtRAmokbCQKsWY3ChYDcKFmgJRYUWpO0gNAPMDFDbmgIrhgV297u/5jPupubK5HRz8x2+zOsyDwiWqvd7XyesrxGBIG0ee3zmX/q4y915nudez/T+vd/kN6EtiWdCmsUKraEtEMTK6bDpKbmURBWs0AnuIjBjHTNMAeamJs39N3F1mNBe7VCeTHv3yi7VzrLQ2m6SqbN4FmOpYErbChWqUIJyq+yOdg0QQJk6hBCyKkzX/+0trwyMF8OFmy/9rqdL5SLYICdTl8tckssDCBjl5v/6lR9/3uDZtu7evvrh/xW4HVkc4Q7qxJAEMnPyMIoDnGJA2AFkHR62DCUBqGMzEwJCLNhSvvHpWrVN0WNqXE+m4ktlaw3VTJBgNooOhQoZGCGNGwwwUlvaQWwTUtzLpjrSZS0kaJCAPuIvsmy7WMfKy//gV9xcEwlCq32l61Ac85TZeQAru9z9B29/cpwtdT1fu98P3P9XH32/TLlsYMMmCOwoUMOjHgLUgBAzgXGAIY5BjAYgQk0pMICFscuvXde5t9Oh57YZ95Mp+bwj8rChZVQjsQkOjHgohaZA2gUdMEBFTWstInSVznYCUkfExdhmO27uv8VPeu1sDud0zz3vZ37N37kxYTZZiLv3HiOdHXvS9EBmm+G7vO8jP3ChOPfcLp0/7Zvf37jOnqgNxhBZA7CoRAiwj0hFwBABsQYYY0DRIohVgYB1I938g7wIaltT3QxPqP4uJmQ1QTkoCwYB9qDQIDqgLepQUCxQFQuISLImnSmJqAEZdxx77Td9g7tx68DmMueTm+/0HovAEhzef1RKBVjUNtzz2//aO+5vz7OW7f7isw//jWtOqpgQILga6+xJIMtDwREEfJBIoAmlJWmpU1qxw203dhGo//BFLCtt64Sn1MPfhpFFPXCRtHVIHg2jCAhtSXeL0AagHZgSBgQo5dZgIp7iUAuMS5N/9uXzrHLW63Xvb9/3re+1HMAU7nR19nr/pNg2HenmX/3Cdw4t5y675949+9udF4GdcXmogTHE0ByIOBIIFKDysCQQZWNCQHakEtWCIoOrCO2UP/0SptO6u8k8qVo+F1FcthIUCBQNFCwoKa54XAat8+empjL/cR0U2MCS45OXj0/8Lef1Zia4vbty7mv+4z/vPbucDm6+97r3c382r91+jl6Hw61dz+97866nN08v9/vs7ux44rNf9e69gdRApDZ5GA5m85+fMvaTw4uNK9FDj5YCAhYE5NYs/tG74RJnEi7UJ1Un/woCGAythZylCosgy99RxYnp9zZZEep5YmZTFQ9na8IFvf7ob7v3ncdxvQo4l+MnfoWnnRNg3F3HYOF6R0eEOfnb/+HHtFzZM7bdD3ydn4uAJA+ChkcTdqwp+9ecPT3BtvjMZPPP0ZFxmKHwyNBbK/3hy3TsrEaBhU+n5CvuXHYoxhDJNMCTYWErtvdqfivqVePtmkWZrPLmZc+zYSuWWZHGDfnzt/fbec7kMuvNl3/bdhgEkPfcTztIx6tnmjBcj8/+rA9jWa+617tr7+e/4ToVsNDiLhm00qnt2/DdbH+0idmzXnm2HpM3Q7Enjk2xY9lkFIvj/lpeDEN2Tzt126dT8dqrIxNEK7C5Aiscm9jMqu6ihz1JN6HmIBX1tJvT56qWQ23oWFKY9eAtv/N0PN2DCwxz+//82tvrkUTK+54dwpjnGe2A24XvffO2c5md1p293v3aj+yItQAcGFtcFJwzfdZtdavtPWJUzeOVOJl39Gz2jFpZgyoLnEpo18/sC2lmcJUl6tMpuH5513NLaWg52Ck9KAIL0sKzts6koj+nq7ziqxF5vUxjsrAyUdC4V37oN3mNDt3WNufpr/0Cl3PBjX8T96OMvnLxykLKb/3Xb7kI2+6z+5Pjlf/gp6I6pNMZwRwHQBY+KwsdTT7l9WrrW92th9Fqn7EUVygdaG1LaxzA8Md613M2zmbPzDB9QjV8kcaIG6E5K9FKIyIP450To7aNMZJ5a+o+XV1nd8UASRASZPX06PwvWc9rHUjb3Jzftz0PhZF/fT0u7p4+udzDLGB80We966atzmW47Cv8Rk7YhQkYGChAFGhXezAby2xvt1s4Q11zXwsTQQFFgoVWoQEof5g1QjZO6ipPqoc/zUwswSrJDhqiBrZo81s1NWrVstX629NWo2Jz+yIYSZKHwlTnp3yfO45Zzxa4ntfjb/7FC1eJjfecnufEdXlPzgANP5jD83TPzrvr7PXZT/z4Pc51DK4QwGJhixvsGb0XTZN1b1CwYdhHz/4YG6koMEKAwWlp+c13OcrMYma3XX1SJf+dIyvMDChMBSkhMaJL1aZrFWWlmrxlularKerq0yJAS9ACmKH7/+Ltd8cxG7scssflB37lDRVz8CV5kbG9f+9VzmWX//r/fseSq13j2fUDH//Lrq4jmBdmBAYUnDjEunPn3SkIlythKYM0Oy40XSsCGBSoEAX4/V+4bDZTO6l7m6dUV//Zq/epqFRAEEoUCDYKZU+R7RVtoQujaDWs9vQnEcrDkCRfvrztb3Hj5fbm5mbYYOfZd35lyd07PjuPmePm5gMvH5zn5QYPfsP95SWOmfNsfPnlo//uLTfMABVbEFDxeFj+eX4XM2Nmw1j5O5HfMilAKRREoHBu+KnuTbBjcueRkidU+dqrlx3ZAJVK7PQKAxDR5sH8+eDatFIPKo9htYnjQRBStIsSC8h+8n/w5e4wKoP4lr/1OxTnevuBL71hLrJeXvuym+OA4n/7qx92Xc7O3WVfuX/vj/mU62EPxYXiofngFNnZ2c+wyYQRYvL12M+wYZtfodyWAiO3q/7MT9y92anUhjLYPqECPm8qYCHYmkXlkJVgrWDfzpmVppeNVjYme/VqU6voDcolXZxkV9b2MvyhD3v2ZD2RA0/G2//4n7Oct9y/Z+Z6NU6ePnvlRLZXv9eHXdxzz7PrXu2Vt/82BCnDQSFYBpYkIi9rPG9feea7NO2Z1dvHb89s/sAeQgWsWghFOq4/xUW3p26n3aZPq+JPm+UOSSJAc7JZLMO8TB1btLCuO0uTaxmn7GprM5tFnAQzyGNHBs5nb/01dzPeEFtUL11/0Fdc2CuvvvIkdhm5+YpXz+s2x8++f+s2w17v99w97n7jnBAw4DILDAy755zOinNBzNo+2fDxLM+H2Y5YTlsO0yrdgoUWCpSWht9698FeYV0bxdVBfUIFf5Y5gYMTwLFp6QBLh0hX4qqnl1qbLrVXMneJt4LXdXenK9mJAw+YLJnL/viPeu9yMrPEtbtu/v53ud7D3fW+zbry2vX9/+DJnPJP//jbODzvrlt7PXnvJ/+wq8PmAAEj7JUVj4ShMt2bb9z79PDRnrrtRr8Lj8c6740scJW2VYsDYbT0b32xkn1Ohs2eAUqfUMW/OJcyDGyhSWHCiEVkXxt3XfKsDSG/tTV1U1wyCwbICQI8yVXxqD9w14XZYpHZfelv/ZFhb/+b8wLC7tOzX/iBubz2Rd9pX7rvbtuz8zyvr776Z5tZDgHUKPFgGNJgQ4td+/B25Nmi3iRm1ugs26aY4raA0EIKXYPc/9GPUqe7M4y7DpSn1PP+993ulDtCTMhBmSPVROrO3Xme6t3hrNY186y3Lm6UoiCAIKHgREOd4P13/mr31xrn5Npy7dnLP/ln/5X3/a3ffLqLeb52vf2cr/NrftnP/JZf8vL9XGPPc9nrs1e+wyfvwMbDxdBWoRX0ZBTNHtrqZ84oKmsqXlt4uEwEWgVpUWBz2/z41z5kQdPVaZGDJ9b295kOEzk1CDCEcEbacMwpSFc18lV11Wdb8SRvWI2C3BQgUEGPbYY/fD1nVZZzzy5zPvuD3+O7/6gPPL2/7nlezx1f+8ov/axf9wfe82E3Uc/2el7P5bz9M8+WFSmYbQgHQiI6wGLpGiWImrE2LUa3euqje3aSVsGKIAyppPC7dh46dPbAtHs2Pq06+eM8NEBSaUkgjCVr/aqG8mAldxGsk+xq3fUrKcgAhSKPbmPHftrXul7H4/Lkcrm5HJe5vPz2ufncL3n2Nm4nb25v5G1vu7195zvf+vT2yVwu3Iu+9OTyA99+ezAa0jKyFmBMimwoavON+Tt0mpR//H3z5/FCym0ApCjQAr/NfeZ8aHtuVgnQp1XyP6YFLEOBDRFoZcLbyF4MtrfGZuwZZw1Ftj0ZAVRIVFAb4C7o/OFnzdop1HHGveu5595tu3cbe3Zzf73hep7Xu+auu/P+lXf/jnb2VMMdAEEWbAcoCWbeytnMsN837/BsqL3JljdmHQiUll1aa7Gg+2//+DiatXd3W2grT6vjn33lNa+JxMMFcGuRDXNbrNu0R3Wpp7wqWc3LsolYTLaFCKQgjTwElq/3Te46OIbO03u7ej2vdZ577l7b9irneb27u7vufdf789q++tovfxmZA/BsomoDa3EJU1st1xcvTXtS2O2LVzC1p02eNvYtU4ssKi0KMHz9pz8wnd1NuBnSJ1bD9fMP9yYBg8XZZkcmZoZcH09x2aV4et0oL1zrqmu1pcesahQtAT2ZEIlC/GPn9Xof48UF3bq/ntfiet0i7s/zXM693nfdtbj/mB+ALAtxhCuMU8pkGRM093Y2n9ltGmfR43VTYyPrxrZur5vAQQFFU7pb9B+dlztdMZsKXYw+rdqDPzw7DMHJYDA1Jm0gD89Y21mPdbUQk1qzbLw+UXUFxghMkuhOwaqi95/4ja/HNFeu6O5199y9cnd/d57X63k9TzBeu7+/P++v5+55v9df9tLiMASMSQTKgow82vSU6eGee5LXd3xyT6q+w+or1U1nKa4CA9JbVqLrr1jZZzlnurtnGGmfVrH8RWe3WEYUYFgqhwgu8VLkrjvcudSdmlVp5WBRI3MjDJJIoAZgbc8/eD1rx2PO63ml6P56bnWe9y1xvV5rt3Ove797+9rH/ChMthWgYNCKCSMCF+y0ldNWu2a7fdozu+zNPvXKPryxtplHWwKFqrQd9g+8XDh7Os5UKE+uPficOxthEIJgm0RhUFCqVNj0yBrZVzf2FCWX0gEZHgpDjoEKoXmzX/fj725AtpHdvV+Y657X0+W61+t1dl/bPZft7NqV33hcAUaTYAQ3gQyBaRkXm8n2UNDEV6nBxTeNDZdWrQAiUFCYgvbND300bsrQA7elT66CV750D1aFB1qKwq5FVK+MUaPJitDnmtZlQiisiAADQYlBoADZ2evvvj9vL+eV3fB67f6szut2apzneV7bve+6170/9/rqh//ALAAtE6IBRQEFppg2W1n1BHvK+kztHeJ1U8/FbcRqKVLQwaIV+//5+roKm2F2tOUJ9hX+O+5PewgThizJDII7RXXZ63fUq7bMIohWhTeoACRj2YggCG1xjovf7eu8enDcXi5ml5uZ81yaGXCO29vDyxxn1/M+9ebSz6uZeQAIgWsLBAaUYVipTW87Rh2DWPfm95TtGK/mPwslQDVzbv/GDz5+uUh6sCKOPsFi+K2gCWxBI5hLG8TkdzMyng3blq4hNvtO3vs5q0cmi9g0okKhwATcxq/p/uYyo8c5xGbslb1ybsDZyVnm6cX72x9zIq8fCoFTSEuAaZK/M+fQzF/zJYuZPcO8WU9mAgMwUiDTivPXfSm7TB1nQO1TrOXvrec5lFOt0RKeSkkYVKnVGV2sqWmUz9Xrmrbb1vawBsCAgnYMiAjOy92nv/uZMzVM25XZ3fb+zrs49+703LPr9cp953lef8LbE9gt0qIywAWQh7ZQmdlX+44xL4Te8cjS60bm7tWNhkXYBGBAmPr1b36B25at07PDE+33OBLABBAWDtqw0qjZS+VCQ35vTmOpbbWoLx5ZBLghECeyIPLo8bPu98YZPZDZpaw9KXPP2GX2bpfh8gvzgWiCJaMlU8zashgx5z253lxjVs/Yy+VZY3Zeve022evNre2i1CqWhX9zXnZKweyNSXyi9ZuP3XVYUFcdh4ydTbbKUue3TrUOdVWQUFfK645uYUSCIzKGoxLXLICffLx2heJ0W3bq2kls55XY5dzrDpuvfMuP3M4WBCiGpJyIRBiEQ+Xr3We8rrP71tK9KEaXltd2T516rXNrFaRWlBIL/5s7HpDpntGezvSJ1m+/n2HPGXaEiqBkGKvS37aFGkpt036St9mqOXG3RmAXgYkFBCSsZetyvvxj7o5jLkCEnG0gUS7bnhPT1rPL73VvVIJiIQkEQsFKIs7Nx222w6t9LOPMmvx9auXk9/CeWrCABWgLx9f/yIeaInSYMsNT7fd+zrEoNdBojCmPjogxSm2La1FtFW0rChrVwAYMQORAPCrKOBft+kuvd4fcHODiBC0NRCCwhdZ+yid1s2Jj0AEmGAjY5gRkhwapH9Fbk5Fj/3jIvM3Gm74uIyAWKcSh/uY3H+6SmT17SBmfbPE7gNgEpm0Crd0CVltRqfU3Fd62J4pqs3WXsVojDFAyRIpCxFDXezw/4ptem8EL0knycHk8hrZzsfN3cUWuwKogCCGg4QgDmtiY3xHDJVvDViO04e1Km8MnpEBLi5TNDP+bvIQWN3v2iTzh+gvMosKyGo/qiECq+R3BGw1z97k7VYq6LrPqKIlHJUAMAsXKLjfMwS96LU84nlxeurkwMx4K8wgsjG3Tx3ybJjwWBCUiHrbxMApAgtpCGhzDUwtFmCn6BO+aGEKhqAVmpX7tt3/n/cG+3l97NZNu+2TLz//i6ZFkYCAeJo9+R3gNeV1kk7KFwVNmlNGGoIBVApaCOEnlSB2+8zve79EBhx4O2ID7OmjrsPx4dhSHiAVyVUB8sCxiPNqG+Y8jFGNmHjx/PjO7txiKQhywmLHrbzo9rsx2mm6GyJPt+FUoRC5AUCxuEXeGd2W8bI+Vvm+rzSaGb6Gt/D4kJKaw1ULiKB4Ka8dPeHbX0RbFtnJGvH4LLb38H3F4AlgwYQtcoYYCmDpxW9uojA0P+9Xe8jLcF2qP21hjeyHdFFLGDin8zz/8SDaZ7mHjtH2yJX90uw+g6RRFmpokVnR5RmkzXnwaqk5hF814WROPW0YOjsUURMG5015+wgTdzM0coBPy/DIg3+1p7BHRYBkwy4UMnJiFPUAEjz1b/DhW9XDlvHiOtXV8u1c6dZQuyuAEBfzGH365i+cu2KN1yJOt5ov/1bi7jXQURJQd7TzIq3ddWJ+Le10pTfRSW69Vbp3YpkbI4BLh2a5qeu0QDz723XfOgbubNG1vgBbh14hirk1IOAhLQwUceLAgq5o7Lc68H7wPjc21bs1VDzf5822cQBDCYLmuv/L+23vubbtnGCvMky0u/K5j6GBZV3aHEXGNUmd3ZUsmu9BJhXS0s1YrFY1QK7AlIeoMSHVIdHL84NfExuGITsTnC5z9mE+KJBCzMMENBDCWCBPFbmxzG8vNbDs8Jn82zAtt2d5zENpCp7jHkf/Dx2dt2TOzN7IJT7l/93nWNo2DSiExe6yzked1lX5igyz9rFJXTf48FY9KCAzZLEWMyg64r/ys4xqcLc0E0PMJyS/CjeIUWMRYGtwNxUwBAZZNVr3xdjKV3ihsxmjLz3aZOrZVC5Euputn/sTHWbuMs3e7t/Z8wnVe3vs5x07E66aA7TBdFNWGsWpplYoiyitksyoKzkQOWoEOGA+XqZE++uu+RhcndhuINxiwlx90B4IOwAwxiMhMUKpgEDg8bWUuFbTI+bMgIb+T58+HOkBtCUD+ijffFtz73HWnMA1PuI2ffVxyMii1gBBwhTJ/Tk2pKEiFyStEjQpqMSEAbYEBgwRJn/iHcbrp5jhmKd74cbP/3jvxMMXC5WEE5A7oCQGIBLlWtpA/aIOIbS0jvwlHSRiBoAPZzn/3g17SQ9ocO87oEy5P+B9eCRcKpQAkCrINvVgd29TeG/XKsGldb28/smGoZFFAKxGAoFXgcf26b7u7PTs89JA3rufyq5AwSECwtGCQDQ7wDIgM2zwWnn8nZubPm63BTKPBKKHIQKbEf+DHv7Lu4UqHs5HKfsJVp/vnvTY1dvJoiyQbN9u7SEvYq9tsd99X+39Re9VZL1YeEALSwCo7skCwJ5DudvDdXz2f5MTUCyj26Tc6yZLIjTbddfesdmxhGzc8yT0tPr7xdp4Nxoa1N3qV1/taz9uzNwMLdgiQnuN/+cWL8+j4UGfOmY3bJ1yP/gJkp1ByCWHIsTqd6gW1o1yUeusye8ql171udodtRBh2FFTTxjLHAjLMyc+9LjMeNjNvbG5u/LTLDgmpGOg2OB5iuwxZMctUHh++Q/Yubhlslv7YxMqd6PqsArViSgse6wd/4Muyh8l5TqVhhzzpkn/xhWbTRnA0EMnSlGmvW7X1+6h6dldFd5eeurRVY/EQhQTKERB0IGCW9XS/9jtfu4io2xvb6+5vXF0RYnFSiEFbEWF1poFU5NHn5qU2Wyu+qWO6T8XW8hq2adIIFERbmv9xPt5783CeTHf3nhH6pIsLv+SGtbERmoJVyVGtu+J1SlGk0iYN7kz+XL0AahEfIEHhSIGK4TF3fKtncDCXceaNHcd5+7VrJwIpmMIBWidHYACqEQk0e3JxrPTYB43Ik2J2eqXQCEKhlNZP/povvugF2Lt7OtDM9CmX1OUPvn+MVhJhxGLispU9ZfT0HPLnasrfRc9VqCEMAgoBUR7dlSUYOIefce7TFQb3je0cX+9lLyRCHQAJEUCwDyAaPQuMrQhmFHMWfbGZnQ3v9PasNrMKRYEW5O/6sS9de7/Ps93O3t00POmO4vzPr5AHS4RsUsxuFa6tq1y/atXe0bXpd0+7QooQHGBTWVw0II8JN6VT7r/N02BL1DfUnr+0xFjUDIYUcUJEW0JhFDXlH5KgmT0+iHhCt6m7Jiopt2MxPvwXv/LivJjSmRkZZgafcgF75RfMbgsoQDM4whRk1WZFg6Ark2ubugqlNgRWkjHooAkI5KEzAcPh8bXf+xpD13PzEPA5ns58WnEoEwS2QVShLJQwbCSPGmYv9vw5k+qwDDoaCiZ/irYQYs/+yT/zHbwcF+c5ncIg0CdewKv/8NKwWmsYFcWe/RjBe1Z6RdnmZRm1NzFbvwXD45VEA7qxpBRBsXz/Oy4dzoURYXoO79/9ERUEhiwMGBgEmmI0Yu0WGGsFt3lvLRjmN7y9J3vboseCUimljv47X35B7LWdvdtGytPw7wlnrMpKYCTnxdYr5HVsnt6TVR2rGrsbNe09I7s8rIBwXVqmQB6VZqwftF4ue9g4LS7POfPpdmy4JTixtlNpwCaEAa0mlPUWetOW2hrvsWVevuTaSDzuyRDAEWbyg7/9i7PbmXP2HvY++2Tsn/1rH9A5QNYkHLSqabqTNZcLrZZrmtXci1s3FgjgKojLcVLDTkBAbrPpfsS7Xx0vpKbE8+75I0ZmYBAgNYZmW92mWMJzY+IAtdc35Pua7pklcxa74YxVmpdjfstUMrTx3/7xl2ede9qSnqSWW59+8d0vA557YKcFhrwKvSqzrsqrrCaIq9ZyvJ5+pVBITCDMNDXMQOEgcHBw8y1e2ZrjiEd9nm4/dR1bHi0BJTmUh8ooHINAWjy3T/ny+Z+Z2/XxKLVeOpLyULc3qxlYMoTAT/7OD4+9geneJ9Od4dE+/Tr+3j85kplQYR5Ux2o3GlnZ7rbaoPbtV6Q5qVjEAIFSGAJSREMQMWD9CLmZ4wBnGHqe88NfdkFIFgHcfYAsiuwiCBsI1Y7Gzn/sjfPncjZjfg/GxV7SUinj+g/57U2daaedDgPePAGXHzvHurI0CjWoByto8bqj0qmhvH1tqbQ7WwkEoW2NwLaEgokSDezS9dsdi84xsSTPe/+d7dKuAIJtjAwbDGzksOe2jhbNWGPQbG7YxRAeQ3v9bDzZjI9ia8vOT/71335lqD2vezrS3fZJmDj/+1+fxWWQh4Kb4k41/TkaVfau7tRdV34LpYLCSCq4qDNGgjyUqcVT3v5x9x14XOYgxNcRfjYnOmRZOIaBIjmDhscgZ2lCaEToTZH9YU/B8Fe4OKpBDLrlP/TiCw/s2bM7TB1QnobH/cEP5JxkgVoxxj8XdqF/DqsLqzDqh/xO8ZwmEI+KED1IdGZuDv1hX3px77fpcjNL4GNf/Rt0AUEECQIEKgwoE7QjAUK/+Xt8/L1W6IgcRVrzZzyBAuiP/T1fWN92uvfDppVmXTJPxGD2+Bd//emdMIQOZGBI7M3vhnxXZIK0mSaUaRTbgyoYsq1aVoSFACrkh9563Nw4c4gOEDDxDToDEBIkNABDIFICyoYQQW9jjTnWY8P8PWP2Nt6wzJ9zqANl/bv3t/fMZtot084+eUK+8iNXcIOWloJ60jriacbubE9lFjR9u3k8jTQ20mpjpAuC0xDLrNQSMnzck1dP87DOdUFA5Hs7RkRuLLIZkCyL2y7CSRmee8Ly0pZt7UvMy8TYfJtWmUa+GjMy1u7pH/7NX97xZOPMbk+cmadjHd3+o79yVGZODkzo3ki3R2r2Mjs80Vg/t1qrlyFyiQl3XRd30GhhZ5kkoeh88g1fPbgdBw9WHm2b78isVxHTHSJ56DpA0wzAcdIahwNzvgUz6WPPtcPLYvUZpm+xWYeBpshi/UdefDzT1Zy7DKzZgk/HNvn+C+gagAbc58x3XeV1JburYdVKXYQu5UpWJqWcOQfoKJAZHABh3Blh+HZ3x4XLHKEDBjC9612dw0FFNiwcIEsHMchABpdxChSYfSyLw5cOi8O2HZZ9vM/oaFHClqH+7t/5bW9irvt0z55rt9D0ydjQvvSFfwI5pRACmHFVz6YU+TM0Ttiklj9LawCKETpqmKAiIiROlwH0yr+/18vNTVzGIwgYDr7F7amYiBmHwRmDZRBUQMFKwNLSsHtZ85kZZ3NcZdzYTt5EDwOI4/wHPvw4Xs929t7ndMq0myfju9N5/Og7O7KhMCKVnyns61Ztleyi4m61/vneEjYECEBifUhgIHBgEHD0CU+6zMwCXORhs3wfGEAgxEg6gFYDQR1chGEqJj0odt7prVk2hfeeJx6VPfczmRgV+/v+kS/vNec5e0/mvDKd1fCU/Oy83P+Kc8LhVFC3vY3NweuuMkVIQin/sc8lGEk2QRAEWhUUMMAQxJtPfO20yxy1MY+wfOp6brCiPJ4ZDWhJECkYjIKtzyMUh0tFwdydgxpP7UfSKhQ//Tf9orvrWke79/Va23Z2fVIGY6zGJ9ZQsAsIaBgFVIiEgpKHOpAQFI8HIA8DMEce1oNSgOKhYQI1fNjHvzZebm+9uRyteFw4jidf1eMYYIACMQEZIJS1TKOVR0NIoRQolNvyGYdbEcJNoUDRAXi4+/Gf/ejjDxaZdbkc8bKStRZ9YrZ5j7dXJagwiHACCBU+WJSUCgggEJJQEOiBtJUsLKUQSG3oIxkGyNK3uqanF2c4cuCm/bCX3CoggAVcjIcBMStIwFDECrRUoPWRtjd9xLakwAwUoCgICCwGe/wFv/NyqdidurBZdCb69IwvYc2Qhcpg17ATNggoYCXUhVhYoKaktiKL0pTkoZw+IJ0CgsBWFmT5Nm0HYMwWXK/HfkMQxoLFnTMiDDYoySAEAWmJNQN063SwIhahdKiMFAO2II8W2gIG/fIPXI4O1gdyGrqh1vLkXNPyi7pcVSEUmRXAA0YzmCOQxwdk8IiNBEIhrpo4BcIEHYtnIQ3tWkVNzZZzfsJtN+fl4nGBceiG5XuCtmkIcBgCxoBIoit4raRxaBoY0gm1YG1bxlYDtFRuHYCWRy0YjB4/ZHRaTtcxmpDQ+vQM9Q84NSQCKBCDQoQhGsEJTHBGsCw24roJJB7hCYqcIAk4JonOACMOiAB++FvvjyeHEE4IF/iOnOAIKBiryqOCnGYNCxeVRJRto2ay0FIMkhIKeAu0kCmjIJTAYBl/0W/58Fy6dK4zpY/yNL2miSe4TZFQCAxoDyQhc6VDHnaYMEwZIqPZINJBPDx4XEACCAKRQECgJ598dyweehBw4bLzCSvZhiDCALQIsDCILAOcLAJt1DpMuG0DtUVQqJQCKkQsUBqANY3xg98xL3aKbU9Q2oX4JA1qffnDyWhsrSJYQUMhXFKglcCEUISkZXjYGg8DFh+UDNjSisXjiXT6NdeOyyoXZc89+Kgn5wSOQLRCBAOVA7CgUMwOtDACTVOYEktFqbQgWNvyeBWwAC0R9ub6Z16GirHXgsuWp+paJDy6UYrMpq5gCBIwoRAgMcQDl8fLgQxEDKSGpCQeOjiVgkECMnzLZ3s73QZ54Qa2T2TSABZzABEDJEAGkDAGcAjQWkEYrNwKyqOipTcV2l0EhIFqfvnfnztOpxs517TG5VM1qmquX1wPylQwwABIAEKImAQjIYA8KiSACQvIQ0NADAwQRB7aKQTBN/M9X/Hq/V5Gin1yaT+JJXwgGI8Wghm0vK4oFDBArQA2INMbaHnHVAoCugRo23QwvPorHl5cXhy5u1ySy5wI6tO1hVxcfeVyiRuII4BgK6hAaouANlpil4aHMVU8aKkH5layEAhsYLAbHkEx9FEv7bHqDAfHXIsfz6FQJPH68VAEBoowCqIAFqFIuVWAohTmZgMIMEBpqdqZaPGnfuhju4ehtQRpK0/Wa5jGe2OTjBggHg4OabuEI7hgNCYjEG0QSlGDQIuhwuwIAVJJCmfU2G5veem8vcx44sQxLF+dE7ZMKiNYHo/AJcSFhBKqgFIYOq0zUApIs0sXUKaQFollsK6Wtt/793+5V1Oolr040erTNRfufIG71BADIQhBGM1EsbgDqOhyGlmSSmgSBEMGBCwhpeJi2ghSc3B8rWdwHMcN7MGu13krA5MIToQxOw9yZTFoGXDFCU9sqMIqEcKgQKlAAUsjgFAGhUVlL/pdf4tfTveuuGvGLFTikzXK7GT5kBKEWmB4fMANjqCRg8QY4kyOdEmTLUkjDYIJjBmqNE4RJhDdkFm+Dvdj4LAg82FvWRY0CIap4CBYUIYQhcAC5OJKAYqDlDBYxLJUKMqjBQkIhTFd5+Wf+MmHV2cxMzDd50w7DOXJeii0HpfdZaiwOAg2IhAGQBygLKTMYOm0cbFRDJBHTQODRiU5IiCEQGGXb9jTu73E4SjM9a2eCMYuUeEASWOwNUJAFmA87JG0oRCa3oBQ2ioDBRwAp4CTIum/zBcHoWzAzmypRZ6yB7VWPARaqhEZY2bAogELwTDwAFGU0XLgSjEUKGw8aliwQjA8XLABmOnr9drTIy9z0b3O8HVhJMDRVKGFErRRQoSpAZAAoUMsDm2R9gZE5NEWCVAClYyW8Bd98PrSSJlpCbMLBnzKBqgE74mSbCYSIhIWtgxJsYLCAwhddfNYheHRnBAWQUpAC3aYWE8wu77zJV66HENwmbm/8h2pEhWQghxAggQQ5OGEAMaYAaQq0hrbGxBKaAQoI4iFTKX9rl97PS9tjbszM6ZOaZ+8eayH+JjMKkaEECAStRAQtAESEkphihEHxEfGtWCAEB88Lg+HA9Dj5sN7/2sztze3DR4vX/haTMijAVoACoI8WjwUrQQBoYCAVgClFApCBdpKAIo4Kec03/bJNHcHM7gux3HkorY8fa/Kclq3B6hlYQGMgICWRDeAsIClaBUi3SCKh0GgwO7CCj1QoNYIOJe9e+ldch30OGTu7/l6zCGPu1AoLA+LHiT7IECFoKBAhbY3LYjUDWB3UUoLyAwZuuz8M877h4dMKbvOdPcE1Dx5I2npuceiM7RTgAu7U4uEGxZAcKI529AWW1QmRIYQnYxjswBRtMgkBIccXt5xBWe1HTh5JyzEWTG7wEIGYRThLnXiFiy06aYFGC2FkaEga6DtsiCjPGobMFd/g2+W36yD5YqD2HFqn76hBo0H7UQcQ5bGRkRllnEyFtRjWWsMRNRR6YERE3AEmw1ZIE6I+2CTnf2kbhgvOzNyeDwtwz2EGMspsSTPEUOVg501BtRtlLaAlFkldAoNLTK2NQUogsj4It+xHq4P3Ku75SRpRQ1P4pUK8EgdLojLZJjtAMtA2cAUixxhmCUDxENBJQnCEYmYlYWBAKeRPc75mNechdGm9i03EmBo4QwloiQH1MKY7XAkLMpIiXKTBpgmDhbVBqXIgKC7YqeXF/NwsO7PXco+r3uGnrb0aVybHImHbMoGE0KIEzA8FHcHwQQBsZkAdvLB7oYI0VILSgNCbEgrZ4N8zGU99GDsuD0+kt0VdGMkFgeqMGNxkkoBt5UqHk0JCI4CmREEWxAg0ALpqJ2PLl1y9sSbjkBFik/iqBGKVX20YiqPKtUibGwCumTSgkENUCxAMDNCoDrFYOgCJQzm0NjG1zydnLlwoL4LRoiQitdVBdYGKFQM0KEQYBiBthS5jRSG2w6PVmGaUM3Lr/hiX9I1HbBzzoBQAJ/IkSrV296upC3JFqDDwxiCRRHRYREUEpkBi9cVJJUQGAJjAFk0ZviYl9tL7cyw9/tJcMziyYAK8whBMSviCBJSgo4IYgER5NEBQUBQ+gjdZqYzMR+d54eZJePuxmw2m0h4Mm/AZV7+Os+up65CqEAkMAMKIwYFMREPDSDQxwwIEQQIRNwomAByv8bevyU6X33ten3y5Px6nQQcAAtuECSIDMBGhYAWQQFIW1qgj0TKCLa23ArRYmHv8zfw6b778O6yLiUAR2Iys+jTOZBa0y33tzosgGY8ai0EWxvABuWu5xQFEkBQBbFEAQQsSOdYFkACb7mduznUS7P3fYTDDCQwCwkuERGPDphBkUAYLFYqoNwWJDAIlj5SLDGwdr7yJg8zmipQGJwpIE/qmpzu953HuEhSEkUBtoEyLCCaBMMplazsCaS7itqDxXAIThcLEsDO48kzMWaxgw9nAQJYpqAT10DdhRZ3gZN2wVODgAzYkaEUKUDDbk3ToTDQYRYz/bRXv4FTsXRwzzlMOyVP6YCqPacDdkBBitSRxDHItFoWC7ELMASN0tRAuTw6IEuy09QBVAy5PGmaS8yh8K482ceG2MEDWFxgbJkQ2YM5sA7iqohMlYZCgW4QloIggxFtxlxyJRfeZB0bUpdnXRymmvZpnSyGeOtBBMUpDcsCkywCuMqMs8phLK3giiTpkAEzSwGu7HiAoCOBDO/Y6+Qx4sq7jKsoC1dVAOewAUJF1KYThig9LMcotWQEcEEL8qiEwg1p40+F18tPdpe087A7s9tdoeWJvVLX6eZmeHYMhjI4NDtBwkgAw9puA5B74BhDBg2x5gSQQcxmQvGcKV/9ejlqDvGCbz0PLiZAB8jCQhQnAkQYehCoggSUyq1ALX2kAC0VuZki5WvHkX2hTZShQVqw8iTfo8v13L1XAsjYJrEHEGGVeUwEwVBguZKLDA+LgSFpRGDARwSVd3DLgWJ7nRk5I5qUBdkpMgbETJJgBYENBjI3pWBFAK1QTKVA2zAS9gGXdhfkCN0Dy5aCT/AqF46YG77sA6BSIFMWukEyoqK0k4gFgmLKRGg5gYAoEc8rQCcf0asuyYU4nrCOiBj4YBB1MAB5mNAEUBpIBKy23IoiRVopULV063xzpQ/3F7MdzsZh9jmk5Yl+OPckX7v5is/f8w7X2CVcg/QBLLRCTmyIYMTDyCyAJihoC63NgAiKg68xr1yOy80e3hxvubtRSwoEiIdSgBQgCSJGoBAwN4D6yNDuCtiixc7Y5Hp9/dp59ZWPjjVHV6LeebbdY8SneQromOPlmw98rsfBes5qDuOwbsjSRANIKmxEsgjiuolmElaCrSnQiqHGxzID6TYn7SghwaZEC4vBghArQEss0aYAU4Q+CoGuAm1khtZ0wvXF8dM/tT9405gUTBJ6AW0wPNUPmsnX7uizzzo9OadiCRDYTaym4Dwt3DmzFIhtlqGFWAqaNAdkA1oKhJfubbZRr16wUEp2DAtZwTCgxTKYbUdxAjgjpVSwTLWWYqeEOhWaqz/+Zp3HN4chzKlFmuke90yf7MGAct5evHzOpB1z5MyCUY0zGglCHSvgeQjJiY07LE664qAJLFunkzIoD4+54rQJnk84Ncia8Gxw13URDTzSIJmZZdmdMzoGt00rIKhFUR0dsOiFP3uutN/synXjJJzBS7JAyRM+ltm8iWP/9as30rAT4RmoFAUDrOvRsMZhBkPS4BTB4UKlIcw4EEoRydu5snhMkAfDuTEowXAC4yAEU8E2FCdIA3JQwCItA9XSBgbEDEld0xn+v9fjHC8P516r7Os00M05lNLe+EQPZz3m/ma9/4rddKFZp8EQAxW0QWAblBDUECKDWkVmK4ZIGaANabnhcnscB3BcLnPUeghuDChwtFYQCDBGw1QjykPpMBJB2qGUDJTGqZQk5x/vi7tj9nm2tey9wiIcgDz5P6LLcrsHxxe9OgYMDasExNij4EIo7rLKCUk1kFNMSzKOtAZBrTPADCOMc+HgencwzVagFDCAhkIBYrEHYBJIiKQCYmskCCkWakU23P/E5cMLLw+9trrsPWxD02Ke/HEevdyQR7fv/2wcJAQpggofLQZxEYepA6Q0xnBkJqIlGEbackKA5enLgDKDV+TShVwkhNUWVWIsIVTKZhHQkASBQkHKgCDVQEn4kT/6svuYXNrhOid6X7oBQ9unfxjPrp0dd1fvf+fdeV8KsTpYo0CAEthIhBLQgJAQEQiDQdTIQBCcFz5w7dXr9fra/fXuyTEnpzBNgNCAw6MKQSgLigwRrIBii2plIAClQnfNeeVv/NnLq1dn4LjL3tfz+jDZ41BCeS6YsEdntxf+7/fdOC1Snq1hBLgUuQWkshAMC4tFkVltQYJEhLjM8vLTGz0cT0+52wMIBFkqoragZVoJBGohgaaCFihTixZaSjrUMpfjp/+K9VF6EW1zodrBsdKG5wPAQdyd59x88T9nrsWCo8OEBewAQUoQTAWL2wTCIiCDgZ1xqstQA/L03Lhet8JOT9Ba2CaTYM1wYgYsXGJIAENZpWBxi4xuW1ulkt/3p15+eL3HmbMtq1Zndt1D6/MBGWjocnD46l+8dzwASwiDASZQmYiAZYRJW5NGasI1IZGRHQPBYDgAPNwRgTBhRIadPJIAlwKEaTwrEGRpBguBRYuMlkk1A9f//d3Ll11yZ9zTvGjY5SIejcRnA+TJ1LWr00vXP/V5jpuplCkCJyIRNmgjBIszp2UjEg1JcNiyQwIL7sy5CQfDcX8wtDkuLAjTwrgCDEryMMaxWnCijgxQtCBYbUvkH/4H8/E1c7rXeS27+5iR7ilThinPCdWTppuh4Z/9KVpEAlgeBlLhKrYI2XYADC4SMQjGCCCCIQN4ukOXgbn1mMzDaEIe5gAIEQ8FAREQBVZEVouATq1YYeTT/9HrL37UtbifOZDszKXRvbompDw33NbaY2l7y7M//OocAZk6BliDQBJwADJIADGFAIGLBDg8ujxcnlxmlIGzZ+dFhERA4qFLJSYEtLAEVcTCYBSltJAglFDL3/17++K47PudOU9nF9jZxXE6M/jsQJxDpst0evQP/mSCiRBJbASYrIvxcAhIT5WUJA7jdYUconBHHjTec0ixPJTHhSGHkpBiYAR8lOH1i7WlyBSGOD/1//jGl1988vrNpdfd6eTKRsq1m1wilOeHjbtZxeW8ur/38znv2So4d7GZCBGnWR8hBMOFgAADExYCrMRmX3r51evS/XX3Mk8XZYyAHsTDIRQoSmIh6IwEWASoMMBMxc4ns/6a3/HqSx988OGX1hyOO+f14eH1w7WN7vM65XmitFXKdfHykn/jD+zt5Xrg7uJhGcIWBMs+orTJTClgFWXkApEKjM15UHEkrQuQEZAsD1sINoyZhRUtFCMa6A3JTB3b4gf54f/L/Rc/WD33vBgYhjsaYKDY8Gwxk1PohGafvI3f/feRvTZuspKLO0KyzFYsJAtIJ4BogMUgwLIBV7o/MwyaawPhiZJsAwuRwrDNwmzaIg+TAE6hYDIN9krd/8UfevVqzt2HacrMdLC7lJGxPGcMmUQU5Obm3/ysr4DrxM7QgWAc8bqDNjA6EXKEZIpk7gmwooHmMcwuoEOLcGxkTOtODPKoNtQADJggTdAB4XYhyXLtv/Lv8yv7zTZz7mHsHhQ5U1hmAJ8vACu0W44Hx9P967/l/nIE47pEp4aJELCyC20FEZsJkAwOyQECjd5pzbjLHCotHCLJzE6QFOVgYAOwMAYriLSFsWNmJvzA//LNl77t7vzmeaVnT69TxnHaKbPZIOVZo7RE05Ren7x9f88fPecCwXkkSC6yLZiLRDNKkzsi7AaZAUQRuOvqmh7USSgPY6KdE9IkIgISECkMN+FcBOnqbsLv+x//2e/4wnFPrjNwck2dXplV6mYCyHNIK3LPo/Po6cvv+eX/UjppJkAxqRGQSWQCgRwC4hBcGSwRQGiN8aCIpxCwkCQMh1KUCikSFqYgGOwhHbGlrh//H/6efuHl9dWC+9nD7nWfM2lruoNbGHyLzxcSKaDb9Qk5/rOf8SUcl1iUIBXNwEhAWGTYeDhhDWxo4lI7yjDpxWfXm2VEJgxBgPBhmUIhGEhAMTCFxbiO82v/nb+v3/XmfHMdzf1cr3t3NufeLePsQgDK4+VZozOHhNflbp7c8NLNX/xxX8xZ4wkk1a4YCULQADE+EFlhmyBYZnR68hHHcazXPXyy772eJw8FAwhWlxKTQCWaAotGWgSh9/C1/9L/my+vL8xlHSmXuV6v+2H3+vDwcD2vbG27eUZZW0aLHeMr3t70F3/plxw3nCusC6SLiLIIEsvDCmCWNSRleHw5aDwG5RYOYoOQQBiaFCCKChE422RoHZjS8yVf+y/81Q/f/R0fdOimM73Ycdo17UCU8uxy2bi2O2zvrXnnO/aPftYXc7DndadKsTYCCE5chVXZikGnBdoKyHmGntD2jHtPioGVxSDKJQA3YXUr8nAGQq6099e58IP/5r9uvvqVta/t3rNn73Nmzu6aduzMLs8vhWYVTV678cLxYe++/r4f8n91jK2NAqIZE3HEUDsQqmyFAzEisdxfiaNFhotztxEMGJ6IMVrCrHacO8kCBLnM0c7N6h/4N/yAX3rFtTN73O11D1Y79Ujacig+uyDoZGmsV9Tj4ty8/fK/f78/+do4XAVWkEc1CqQBlABnAANBkOF6nud1c7apc26PYyUWiIMgAGXxAJhhVikssysefP7f/G/7Bz/47q/49dnsPVfOsz33OXvD7rB32ymU55k6up37muCx+6pv+6gP//wf/x2//DgOznNdXECAwSFwCRAIEEiBDZY9788pKi7Hsncns8JsGAhHtZVQCYE1Cey5efHZ3/gPPvXPfPV7vnTtaqe9zn2n49nNuYeze1kdyvPNZcM973Tjei5z7e3vuvmr3+G/u/dygDULPTgBQoUegvaQCdgx5DifaXCpfTY3x8yxDS4zIA9DJ7UkqBbNrWWO4XN/0tf6rn/7bd/+cV/edZ8zI3N9mH125uzsqRQK4Xnn7i67rMccXer+ej3ecvNh/9cP+H7/aA1Mk+KgbXhcHChBgmKo9uw83S0Yuosd2moAWiASmkhhYNSIBo/pC37R1/7Gf+b2Uz72ycO+5mCuPa8PM33Tc59TdnevIA4DPudIo6mLlXXYdXvHg/v+G/vb/+lPvBr3anIBrrkVQWBBKqSiCCI2XVh63aNO72o973f2dnXmxdd44uCICJwGo0EESfUmIuIGnPmDH3jVzqUX70xu86wxsMYa2/HY0/BIerPNqIS39rSBzFFWsliulRxeri7X9Zvf3LnPF3zE1cRCbkQQARYpK0UiAfRGiuBK/fUnXXFivRrrPvrJ/b2X3n00WYIgAggSClyWAIwSV4M7nvl1r/TVV51MetTxyCCD2fFI9zpjhKE4Rm7jrT5uM2pYunB5iYf76LG85MY79nL/x37KAxqxJdpIE4wSEFJYKSIiMJL78qe+9S7TxjzGju94+d2H0gBLLhIJKwglQn0iQ2lifsVv//2/+cord6g+RkX2yEhGOmMN9pg9TEhGFN4a1AKmsBqXZnWtlcmlK73sabm6Y79f8Y7v+ml3L+KMpRtWRFAUKYIopAiI6Ft/9pqs5tVs94n18+6VpsgUQDRK4CKKEms0FXD2GT/wvJuWl528Zt8hCzOjZIzMMcM9HTNma12RHUL5rUAetRkla1ZXzPI4L2YdZK+pMab1TdElj378w0+VYLjkpiQlVwQQBUW4TH3i314x7c1j7nPadOtzHto1jRYRYYkgXAmDShWw9+9P+/OXr05ddOqkpq55Gi1OrFmzbadnxIMMVokrwQpvVWrFOVh7uVgczVoeZL3wYohp0c74jr06/ZhPfdjFjZlKMcpqpIJVAImQa3zQiy5dnJmzmkemxc3//tChiU1XUtEgZVEpoL/h3/7g2TfOuxefWEw1Jk+zorkSZ9DxMLN7ZGfGWoGGrIDy1iSPZoKaEPFYy6xhTaVQrZaZs3c709X3/KgPuuZEA4bSYhEacQWkjMbD9nba3ip9Pae12//ifYbSLCWSlRo0Nn3bq//o99+8KnYu7ZcOzTRqKDUYvTnOGKN67Hmkeg163JEc3mp1tTVJKjmSoym1TDUWqVSqq1bzatV08UMf/iH3u+REm5AgSQsCyLy46ZEnzlx8dmXP7tN086f/rIoKCBKheH3zi1/wT6943X7zxZdo2k1fNRVFuWNInD6wvZ56Bk5vA8fqDm/lKkoOelQuMSVRLVmIkppaGfnMbWN0TzuXvsvdHvjAa644uWRRRMUY7frP+KeTy9Tsbo2+3Gt/9Y4LAK/6uOHGW17+78+/fn3Gy52TJ2oXGtjTgODgYA/UZzxije50kxE5FkF5KxcQ4qrHuDjCJC9SUatMaVFDTdOIurxe7w9nuayTy4suvvjSk2dX+7fd+vL1Rcua6R4MZnnUpW9z9SU3vO7s3l53t3aXO9PO6cYqi14DyXLKxIy0ORlyVolHzbUyHhqVQVB461gbCS5iWloDWIw2IWg0IaHRpKl6G2LVvZ6HG6OGTrSpRu1pJfogeNcai1aaFm1aZIpSZNLcFMVyybgywE48lEH3sNUZg7k0sDBvXbu0IVItosSEp1JR5Wo0q/WqsVgtKlPFHZSIPYtS1ypj1tBQIXazONmWPYu0IeYpLFJyJHU0YCM9JIk7cdIzI7qTAYTw1rbaZZcLNXnqU8tUaaOVWtHmktRGQ2lKMleUZoaTmtZ1ZmQEQ7TwbqvWWlFpLqEKaQS57SuILssmHRhxnJFgh1mxCG+FK04aqZRypdHUolQrSpVCtJRFQ64VKSWoXL3NYzUSM5TGoi+rTU0tTYtZUqBFkbG6IscKoWcImxHHmJHUMGVQ3voGSGkhVUCSioKqmgJSo0WEUpl0pEK1X6rgeT2sDEdoYtLUqlEsNFSSrbLEMIYkdsVkDE9x1GPLUk8Ib+WrQEnEjSoBpVaioioEVhMEp1QQewrzGMNDScqZWDZNpCaVEiSU1iPDYCjEbk5ie+okJDZpxryVb4ZNbbqi5lZANVMoamlNRAUlEsUkHY2zDqaN1igoTQ1ES1MUKkMJbg5JIJEHwbGVECeA8tY+xFuChCiEWlJpo9pQUZOAymRImtucchz2B0m5Kmo019IVWjRFWG4OcZsbSVyxklgOwSE1trxVsSLRkhaaiygSNDUXUlIgIxNlJGsGFTJpVAUtZVpaqoJ6c2pujCJJaijB1IgpB0h462FlQ1GgokymVexmUSWVJQS4ihj1iDEYgUqLhoo2QagWJEtg1FNKwgAhq2s0Ew3qrYsOrZQLQOoSUpobRRs1WqJyzUpEb+uyUhYKBaWUJNo8EeGaK0KjkqToZUUh1XmrZyFLGmVFSqtU5DaKzZorgVFDYVOkLFRKC5RbEKNcwQSBITVAVhQov3UTSAaQNEAl5AqyKi6TIkNRUNgUpECUCxTKAdLiRIjqBKASAJG3cgIUAAVQVEBMSWSAFJkjlhVRyApSlFgiIBdh21tRKxxVhM0yx68gChMpUCTQRngr7/KWAxUFBWWbcpCCQnmy2RRRUMpv7XU+Kxys8JazjbcCUs6TQypsVe4c5S2SFRTVeOuf81rhzr8zJxy9/NZSbxEVYHHqmvc5fWnzG19y/a033hreqn730vf/1IdcsWwVN0H3vHr1f/7Cv936VnLTw7/r3U42GM7UFVCxnjR5XPcNv/LWbUoZHvU7d5MEQURREJsuYP2Pj72Jwm+lhsKD/ulKj0lRQBBBFJkKqc7EH3zCzFuvT3/8gThtLLAQloxkCUMB6QvOfvA/vtXaw/7srqbNKimRIiIXQVEgZSmj6ekfgvJWaMovfWJTaVhNQMRopoiIoqAI15zlz30ub53+Fd8jSklRQURRFFIQkCWILPQzn//WZ5K56nff3WpJWgQuIAIIYjOREmmv/dhXvdXZ5uKzf6SNVkMSWyM2LbFp3MZkknnxmb9K+a3OmG456Rb5oIOzjYggkoIP/vPdebzVGV/6I1EUcXAQB0ccGBCn9qu/9Rl7uxFAFBERRbgIgMASQZh67b156/Opf+zvuiDCRURkKSKKiNhqgRjtJ774rc/aWN500qXR2Bqx1QVERBERo8FQ3et1b3VG+Vu+w5QriiI2k4oAIhKJyErRef47orx1mWpcfMOSoYIgghIJoqAoAjai8lD7tF97azMUnvoRsZDYdLEZRWxGbCZ46qnr78pbmQty9eum9ZICXINKaKZI0iKO6Hrs7zTGW5NVavAJv9mnPjEaQBBE69bYGgVBQJBce0+qevmtxgBFf/CRmILrX7h+1Q03X3e/d3r4PelMELFpis0wprf/jzJvRV4GRd/7uZfsPednnn0HonloffIxX3OvOS0TLjYjXMBof/ahU0d5q7Gtiz697/LaN3Nyp7WduWZn7/pTX/Fli7nVaACRi62rRT+9Ut6arKyU2s5iOnGqNKE+4Sk3rc+852+eGC0IIixFEZ3p/f5mEfpbjwGidNVoy0VVa2LqQdfPdfZjf0GWODAC8Nz+7COWvd66DKqd2Mm0qF21MdWo+PreuPnJn+4C1xYiiNz7RQtp/61yFOUtFSwmFsuUd9oUWX0/Zu7/eV8YktsGERHr5fv+7WIuEuUg5WgKUH4LpPwPvjJvyVtNU5bWIstk6uwxWp8f9ee7uCAIVwRkvfMzn7/jLszBZc6llKC8xfgf/8pbLqlaSQuKkFrh0tBfvXNSAUVbIvp07d3bcsaWXEYBlONSOM7y/0yynxf7mP282P//UAYUjlM5X8rblG3lbQoomtooqWiMMixG7S3H3gf/0YRbIoGLzaG9S727XxmgsH3qR2pjS11x2f12dHPP6+9YndmmcJ4q50C5sFW1MzwTJJINQRBFcBkpVLF2LOIABrtTkElL0o7UDMUm28GmhWQyazagA5aEmZuU3ggFtDWdw82sjr0BrAW0vdEBiwVSw7TpOwi9CRUpNcBlUM6DTS2H5iMoKChlSiZHknyYIytQBhAEtV6gUSIKXmrsTv1s+88H4IriSmHKgvnuNyzJCJrGACrhWC95xwe8x7vdY1FEjLbOmX97+h+/NBy5jKZZIBKhglhWOLJSVrZc0L7y1x+aBStfHIzSelCZaSZICCVp0ZnW1pV9Hpy6hOgDzFqnn9xhitDRcXJcYz2xeuWS7Kzm9coZux7WkIxzTOzOQxJHPz1Wg73GLmd9Evfq5m45dmSdxFl7rZl136SG7SR71X2VOaQc1TZ4SnaV5mJMZyaNpt6u2nvJR9xYPhfv96Qr7DNtt8pN0VIgOCBlJYB7aVTssSyIaxS90SdNAVJzcxqI1plCrSkIfbEYq7A37XQPL6rlovVtyhsfcPdKE64ICIqAm193WfNY72redTKdnp726bcd6dL3/4qHn4oSp41yRVXQb/+NH3g5KIfh7n9615w5c1fsSXhIvcWjkkgsxjRXy7wYNfel6Ld+4JsucP3ypxE5BSBTgdRwUyoMVIAGpKkDEJgagCIVOmu0DGYHp2ZsrotAHUuopxAnU1PrSAnsyaow2K7JVBSmEeruEupoBQawZNdAs8tCmIlYpyqwBau7TEG4rAh1VZyf+1LnHFz6yssAV4wXw6VUyqEE4HKUGuWSYKT1EnJUiBgaRE4lSrOZLGShKMqo4vARWKJPBDEaSRFBRITLCnLKFix4v785wr2feo/LXesa08gCnEgFrvK48ct+1+VDXPyvD3IyJXINCQoM5RplVABDlAtc8JNfhFq/gHXdVevWG2MaFUkjblbkNorNcrmi8Mg602Ddqc04dtXJMAEcZtU6NlLaHVJnNROc9hhxowB7zaqTMqvgdlbjjE1TBip1llQynYVjaWozVJg1Nlh3WbXMmlUnmx4jjFIKGjVaKjKjrU6nkmN76L8tFK1VRhWNFqgkAtpQpIiQoWYFSHO5TKYYeeqT5QopFEIqRXAqlXhKoeA2miugLk2zREsKiNiepCVhlBXoE0rqG77/UPWrjyNDxq0wDWMxjUbUrbrtA/6dg6f7/WdHbmWlolGpCIZcVHBZSGG7Mv3SZ1Beri9cveFuoxQRgStiM3IREbnY6oIgdtiMFgK8CMjTWGAqcnmBIIhNC8RWB5gwCrUXgUnvpFJHYfQIVA4qVHiEIjaoYmuEhQgigqSiiCXbIzxlwULRKDEaUcSSCCBZtKapxji2z9wdlHdSgCUSyNXohhpI9ikQJmQmkMkTMhPYwwREyUJgKQKIyCJTGmQJWeBCbAqIsKRCBUs2fRKI26Mb1WWlDd7xLy7DNWVCRhPQMglqlEYt4bJ/fcW70jZEv09rO4vd1hbTok1tmhZtaq0tFjvTolpbtMWiTW1Sm6ZqrU28LYtWs7Cabj41WoSCiCKipIKAIAJyEYwEEIFQQIVIKkBhWUSEKESjErkIGgVUZALFWiaIAMWGc4ENRajWQqUAYkkuimgDUgaQBSAlIWLTBaEAgihIA7E12aBCm9JLOa6rcVNJuCyBJDLsZOpOwJBgkGAmlARgsIoYlCDhogAiECgCIiKIikgbREAKRMTWiMKF4GSUSpnBH36kh8pUEomgSJCoQYElP+BVdx0UgK6hkAswhBhIFDtCUhBEissY79SkUsOqbjoNoSJcJBKJRISFgoAobhUXkRRALgIQxbQoVG8gkMsCmlWRS3YBuENToXtRMiltqOLkKEKbtnG0QEEK7pYgAIErQlSMKICkEBFAAgFFcgVwgYtIUSQXCiGlUYjj3icEp5UViLQKuGg61MGO7WDQeeGcgcUaMo3ZBrawIKkaLQIEkSlCkCULmVKIgpQIIgwkRQJEBRFNbQwFeMmDTDEaVrEZICklUVIEaz750ssIBO0yTAFRIqmwhVXNhlAhSEgpKfOiDZWCla6/wkhsWgIIwopAbAQBuEoRgItUpIiAaoALTy4KwEVkyaUUlFPUsEYmVBYNzQ4EAiBYCsFJScusscjtAnkLWADRqApERBSAFBGlRREMyY1IgQJFQqiNRoqRgVqmjo8HpoJKVCkicGUHJ5CMidnByORgD0mhCc8Jae5vaKJKRRW1yFLQaCmSGi1UrAgqKQuUDVegXLhFSosATGhTNQLccsloCY0giKKyoKJmQkXN1JSTb74qbAYqkCgCEVRBBCHjKAKBUJBPLTKhigHV9Lq7uEAEUEQEQbggIBfQJ0aLomg0NYBEJIWlhorRKIBRFRASImCpiAIIjFYqO11gkVrcEYFCsCDWCkjBlYLgBhsuGq5EoFG4RCJcaYBGk4oQIeEGKKAIkaI3qzWlWud4TzAWpmQQIeYBww64kzlAxRARMhAIhhwAO9xQhIAIkSIRUYAlyhIpEJUQKUqiQhBUpJJQgFwUM82jM3V+8rQbShEVRBYUhBABuFFJatQbGk3UWUAxogCiCFBElaUShcGYraO1qsoSqNubrgDkYjNyASQSW4OIghpCES2KgAqiRpMil6uR1CiVAEuQCqSIAiJIwyDQWUCdlFpsqEUE2lQIYIFaC42GaFBEFJsuIGqgCAgVEUUFRQqBksamAmIzamC1TOVyjmUJLqJiU2BNOGkmAukRDwUIEQgpIRsSp0JsClAiKYpSARURoQAVELmKCIISRZEAFBVAXJmqivf6gmhUCAKiFEQBKgKnVRBKtNLLGYiZFhQ2w/Y4BBCEzSAKMFp4oomHGFC98cq+iBsuwEVAEEUREZtBELEZEJuKoKEQUQBFM8JCguCKiIQC1FKQUl0UbeoIVG7Lo4pQ2UFuBWi42JqIREiQFBAFFIRABgVtRGyP2AxUBAxXuUlRxcdxAlUkbxGAhiCAPCrggkCChGBpPBQghAiSYlNBKAAyEFGBIHBZKSIIYmvEls2IqC4q13L/5J/OlKKGJVAECFEQIjUEMEraWd37B0g8UQlNGwIiqIZCkFASEAIZRVlGk6a1BKpX3gOKo0YQsTUogiBARBBtoAiCACILoIIKRFIUAkUgIpUMQQIgYCVUYEJNGwAZIl3UOlLZLDajAiRBhAQkECFcRBRBhCKKoooQIBAgQK65arle7inmOFcIRRVxsAkEkgmQNQBpCpAgISSJCLgAFUGAIhAoAZcgCAGIAhAKIggQWyOBQI3dGtPgD5ZqyLIQlhQURQAKCAJSI3L5K+7BlIswYvsWNkUhACOlIBBJiEmNys2A6qIX3c1FbYmijQiioKQiIAggIAKCgNgUYEkUIYoAXKgAV4QgQrXQcFsLFXFSoCK3Sq2QQq0iATm0ggBZCNhQITYrEhGbCiRCQRGbEQRFZMl08vZp2ifhOJc3RBEgIIIEIQUyeSgOJIIk8nBHAHl81QIXRABBABFIVFAEri0BRAQCAQiCQEAEDtlp8+0P+gDNKSpFRUUkRzBKIFdQRIhCLXbFN6NcAcSYOICKBAhhswTCCBTQtLzd3QwLqE6/6OosRoswEhBEBInkgiC2BgEKsaQIBJGlShSEhSKiAohSCAgiYmSkgEVQgDCmAtQKUGuxzaQyOimKNqIIAURFgGxsRkQIRgMByFVAALFVgEC4RpvmFk+d4yzXCbZHRICQIA9FnlsSdgRIkJ1MkBBJUC5ESInNCIQLhIByAUFsRiIbLkBEHDJ0av/kt7PaISBAEURFJCkKICKSEADhz3/IzCNBLioqNoMUUAgQYkyRECE4OU1SppsYTO2+6urehBCmwEVIERFFlkgUCZkIhJQoREGkiERSRBCBC+QSuCK2x8zWzmKgVsems9ZMKkKlReqkKIG9Qi0gIgIiKYhCFFAiEZEIU1FFbEaY2hCbLiCyhIuGeto0StN8NDNKxAWCuPCQSbiISYAPFk2mFORxd5IdN0kgCkIRERBBoALgAiogJQiEBbiQqSBZgggScZv35wc/ri9HkyvIJUhhyZChEoFEUtBGFJ5Mm8hcMQqQgZIoRsIkIAsiBQhOuaoiYUzxmnsNWlIQREBYchFSbB0lF9RoQACRitiUpdGIXEqxWZEoSIoEyaC0GZlFnVAhE1iQQR6tBsamVsa6ig2PJkIAFQWUAogQCJeRhEAQBTeJoMo2QQRyQwBroYl5ueI4PTGqAUFppMAEYa6XMGkAkgEThkx2ZHJ2aGB2kIhIADKCoAgBiogASxBARJHGBAUgBEgQuUgzTRO/AGmgUaKAFKlQo1KVKIWIlEYURMujHzLdMTzJZRShRoRQmbhCISQgJgp4pk2pdoPVHz1mPRHJpeBGkIvIZSmCCIiCGi7GRBQhQFEiCgpBETEaEDRKRYqhKiJN0C5GYaRVKjuECdXKZEeLUNKAbaC6KFxAhJAjGC0I4SJFBQUhLLloJHKBXFgSEYpaIIJaLUsBysdxB9WiCKggdpAMuSCQlmRTAyBINJB7MDWcB8MOgiIiQrhACMAFICAFwS0i4KaJyClJrg2IgIhQ0+qBb880cFGYciECsSsJLsBSoCWVAcJ8dp6xqnIBCIKFCAg1UIwgKGUloGrFpAobUH3sb7sZVRRBFEgFRVgKgiArBSFqAawiAlADBW0QhUI1WmhsbwQFBGvVMpYGakVogGprF2BRgNLYCi2A4pBSlBaIRBERiYgIChSRCiJCoQACEYgIyLJlSlNbK0fTq004UEQDAS4k0CSTCpgAmQMZHDzMAXYIhAUCWZEAgsAKSkoRqMLWxtaykCsgQFGBDEn7ckYrVxQVCAJQTGf39i4+sXApBKKo0JQ45c+7/K9vL8sUEKASUFBwhAQgRzSpQJl3FpHsEoOpq15wNRghS0EGEQkgCBgNUxVSEnKKBihCbFXYqiCImguMNqxURKxYsQRCGKmEakWqAQZBqNCkCjQcUhChCCFIRJALkiJERBEuXBQGAREpwAUIwqhp0c6G2llxdBdKRSgIVyQBhoc5PO4qgqsgJMjDnCAckLiKKAoFRBEIkAgUERGKtoALoICKiIukgCC4eM57A6OhsClAADd+8X/esbri3b/j8lFVRBwYpOSiT3/92naPDQqAAIFSipGjUJIJgclTpe1MIhKGFB/+G7tVnRalIJFMRUQBgohEEIpIaIYgBBAFpA25RJCiAqfcsMpERJU2U2XQgQwZi3vJQIbMPgK1KJMwqRNSRRsuiDRKmCIaDRCWK1RARMLNFVVklbUhIkDFgepuhWBY0MYRSBlEUkBFJI8m4A6GsEOyqkk70usIgiQBFSFrVCAhigRCIGKJEEjcFEiBICISoYiKgELxprb70HmaW19EgItIEeNFH+/lxe2Wp33Rdz5gNLkg2lIRlfH5Z/6jtaaeIkggEggIVKEUEVNRpHALywU4TI7FFA/703uMFg0pBR5TRpVGKqMVYywUQUaTQZLDaABBCEFEBEUkz6VO4YysVGK0ITRKhm1OUlh1djlXe81cDWXoTriudslkAgQJt2KzYmE1nIpFWDdAGmmJWkYKGo0KoBQUQIQAbHArK+21Z2/fZXd3vdcF5qh/uFqmM6nmhQuRPEyTHJa5XsRTY0AED8Ay5FFBEEBBNZLMVdBLvVBWpYq0aqoaJaMGAkWQFAhLuDRqVMY0vMOrVpermJgQgCIRols+554n5r5/l9ue/qSfWjgNEEBChUzTvZZPur5nObVwyBAOLaNiKKOj/etvua03mkWYYP0u/+suV7x9jfVo1JJuT9KJYXfBxVGi0U9mnSycBQUCIhGBKUViNLjx1Imu9DQPqAF24aTixe7UXjqex3VsINe5CHNsGFLPObsPL4uZ1glUAIk2UFcyDYWUC4p4NBAjqe5Sr2bJSgHCFUUoMsV8dr0eO8vMp8b1HzudgsWZkEI50qs++5vf3OrEzu58TbuI+MgESRCY+1efvnZ/OcQdunI5ZjsCBSGTR197+tplWsM82ScUJ91xhtvMlGm/T43lHq6cXVxydncpgBBQERQEN56c3cv7O3X2dF7ypLoPhGK7SKSE79Slfb3u+4uLfuPhL2ocUgpEpD3i37/mib/+tD957WvfcNuNLwfHqkgkJAh40fNe/8abb7/hja99zXVvfM1zP+7URQ1kMoRpldJ20xtjklgsYqpR05ilosmdWqh2Rk9Y1tve974ffwWAiASUcQUVnHnHxTi1diGqjeFawi59aCrdu1kK7M1K1tx3udiZnQVndhO8Wo7jYf83/+UzhmrBtsRoqw945aXzQFMbPaOoTHsNSU1WT6o0dgbN471+TuWCqOIGoJRZ3b/Vuo00nTp7+7h41x4aYSgcee8pDy9NUxa54/RLT6throIrsvzuX/hRr93klatc9HrbXR/4Wv/9W2U1wR3IfO2bfMXbZozeRnmxr5kWLdfuSLQiU5TpkumMamHprhfn1xsusWkpSeGbPui20/OyFmK9v+KOaboYLwIQuSCKEc9ve3ur1ViN9cWv+tCnRUAUBRSRwOPGP/7cNz/+C9/7Qz7i0Q9+VygRiyAsJcP90z7gvd/j0Q99+MPe6dFv/zbv8KGr06EyIqiqLLJsmhqlak1zajSq1hXTVEmNqdzOetHJHv+xn3f4+ctGA8bEZrlAyJWd5elbW1sxQK45RUvKqkz3SMO61lOMsBvzoGdWA24wOdee9P53YgSrUAIRvnZeXOSE3lp2eyEW1kCiTKSM5brmrjt+f70kBEUSW2W49cSJVvNKrhUX7Xh0ORVUzpEmWg2yZOdEm0ZVusPDARq+4nj60s31yavny1cmnlzncnd57SkEBBDBDse+7SmJZlntpDM3lXYHnkiV6Qvg5lqU5nm65ZW7J159/z4BLpCQhWt1yz0uuiXz2tgTF/X9M2BtaLSISoXk9Bv3xrwefWZc/LdvvksUCQFBQ4p4h4tOn7rIV5yY77hD++tApRKBEiRRO/e810UX7Z664vTJiy+/7PTFUytqU0QGUC3FtJO2LE9qk0YJijZHpKiaRiJPTMWQa1rM//QbX5ikmIhwUUQBXDt3PfXmmNlRlWUXnmKo2lhsTgc0l14BjvuM2GN7rpaunMB87bg/1ggCCJBRu/faP5FZw4whE5JSsBFFLNIqCbcCLiIChTao4Usvasu+c8LqJ+YxM2eu0VMkHNmT3Sp9h2l34YyDhqxJspOXbt513/lWa4fzinv50k8kGHZIDqQ97SM+8DH3d2C52zttGDBKImeRuakNaVGws15dMq/OvPDDZYpiNAHFKGbt9hMjvc1ZjXVf+0TMaCAawGhE0rv82pms1mPf8zLjt77Uje0CWky4f1tN3p1OnD55aufEUgnFZmhyMHTl1KmLTi1P7C5bTtZkp3sdAvVi0FRNu2JKU5NcyIthimShsoyrzW42Wt6qe/zpurm5IKSSQi4Yi7qoz3hdZFRnTD2u2k8pY5oq7HVCyqzVM3WnmdXutWV1dpJTOQ0zVqn0EU+Kb7tqbZE1MTKjZsZkR1JkuUZc8xjLBKggLOECEtrQsJrXNWStpp6K8eAYzWhjnuZR664Kj4qDUHK33l3zyvXYswbOZ1ed2JlVgLVL58zl2X3GqB6TMcoZFUbZ1lnNbdRo3hfjxLpu3vXpX/ltCFjlIm6U29n5zJRaGfe+msfQmVFMowSgpGE02iPPmHWf59Hi9ns0gAg3ARTiymnyNK9X69WYTZcCEmA2w7w/r/f2xxj2WM1DjNEzLAIqhdYyaVFVlZYqiQhAocKUEFgODcG8gttOrk42ylIhVKYKqOxdlHlO4dTAzDarVnRZ+zYy3legXoWz4yx7ZXsfvHLmuoJT9o8PNACVWjEl7CZIs9WT4GS0aGQkDSmz5kXC5GVSA0KUIioA1VicnEnGYIxEi9msZhyRo9VawUO17iWpPfIwYWJevTx93/VynlhxHveXc47j/iWbUAhl2OP+WW4g1emWhpN1S4atNWENY4+00eY2rrz14jefxiVcIFAjKZaunf39sj36SHPmsSBFBKACivD+V197yXrMw14xvXBeREGhgY0AdndvYz3N6/RV6xwyKihRKmo0GQUvGGMGJ2sJ2JOKXaiJ5oUpISuRgJZQ4C4VAXua9ytAsTUqiKB5XuzvybbovTmpeGSENkymrjIFwzQzkTKMsmcVSPdcj+k+AlO5FcEo0f6kMj0MO45wbCjjEWiM2Jrj1J4rKAIQgCXVGM2KZPe16GFgx+FY7SGNKk5ktIgkJAnNpffPcRIntOfB/XFyewvuACkI5+z1pWdPriQdD6cHMdJJFCfu6T3tnldffa/7nVpc+pCd5Uuu398FKoIIUIGmeb9njKHeVfs91ZKGwtaIClPa9z5mcXqve8Q7uePaewsBIpQAAVd4IYl1YzQmUDbEdo3KtFy4PHeaNOYxnG5DC6rUKlILUqQalEimPiFHyVI4NC8y57qqeRK4AEWggHLp+kY1jbl394uqnMyrE6q57NWrJccD184xqdcVrtA3dxzn6vnwYTxfPOz7S/DIef+FV33MWqkQXkx7K3uenT7oDvSpMiSiRq/gZUmjvF4XCQJBEBRENd1+ei5ovZZzX3v0PkM4zgEd7XFz9TOTC0gBQbDz6bO3Pzv3etm9X+ru5npeXvsq909wYEcy8NjDZ+95z7pN69lZT2IQ+n7Um8N+9Zl7fNr7veeirCAlj64xSiBQ2Iy4NGd3xtpj9ljNo1Z6yTRKhO2CiKi/14983ZsXq3Jrmvtr6ROgSLAlcJWpzHP2Mq/YVNiaAEzj7N4Ygybvr/rQpPnseqwPQVtSWoXFaK6IUixBWW00hpy1IlhNvU/zTeqFpQKwgAjQ8kzrw/ZMT1frMytW3vF+mTfe8bNrz9kXPGRO+mqG8Z47rzn70Bf9NOfeuXBd8/C13TgBZgkCpC2WdPfhmXlgJ/JaI2VVBHENJLehRShhlAJBUFKcFDDcx8iYh3sQ5rgVAvYauSgSCEk5Nnu2e+6ebt21V+4vQDKQgblb3V3beoyuufogZqQPsZb7irt/8RefBttWgVOiiQOFCwSjgtXt7tFrpX7DLZcOEBFEbBo3Pv2uLx0tbcx7zDfTcEXikMbWlAaYzcQKEEuEYh6aRjyP0ceqjz5DRgnekqsiUZHSElwVpayyK6uiysxL9/leux0p3lJDklzMzYyyxtrprJmH+siS20Wl3+1PCfQhn6zGmX5SqZ1PcfX6xk9Sd/fI3bnXjw3d60wIBUSMVpOx5z5GZ051ZcooWZEGNYrF5Oa4KVEkVERRRESwMvrcNWteGzOswbGHGOO1KNCFGktjr91d789r591Z9RJnN3ejnEOIhcW0dinufbbmMFfslEzGavnOv/hQ0mEqSgQqIkmBi6CCyGisu4dnz2tnPdzOvObhFYEwRYRSEPQtn7CreVFDxRnmpRLAkBAndNNJjY4DKCCMSBAG9XnuTu+Z3Ufvw3NmvqgEWgqlLJJpICRXL1CtF2shrSpnPgeEylFEmiEVLcZakRyPOblx0YPmxZmlZzNYrq3Nbq6uujMDEI6r6b7Y2ZbVnLlYCOEotU4AkXFWCus+eh+dEciEVqoASjQy9ttZ1Tg5VBHCZVAUKWWa0HDm3qw+EifKsW11KUMRqTjgRjO49/fX86y7YM9re+luakDYsYPMYb139OwzxtBgT5ChEa3e5TfvDaPUQCIRCmOSG5tBgAvRVKMP6IN5BBrf96sWW4uIRJEVPu7R/7psDhXC5PJYBAEjqgpdY3iUprERVyAFZWGLqdxtmd7HGBmsh2iAHVKpIKtGlWtEkai51pT6NKY7rnrMoExJAZQCAUNnV7t9zOticXv/nkedmirTeu1SFuZ88cDFdV2nLMft4cPdPkDSXF7vF+ehre3i+NovvLQOGXHCrUTNvc9x714nXVjriiXXKCKilUBlDSJcQIVAQAGv+zpjzJ0x4gxCOLcZvbXAeSBJMjXxGud5Xq97Pa/FPtv2yd0Zeh6BeYBJx03ZvcoYcwY2sXq58/fvCaM1KnELEFxqoYykpIgoXJABjPS17VijftfN4EIgoopSJPzEu1Dt7EJwDZ5SE26AJgVgX6MPMbqzhlSUiJDCVYYpUc3useexHiOkh7RDKiiAAlSi0dDa/Q92ShYhBJJAJLxY71+2N894ePWYL1wCDNUoYYrKoxNoCVC32QbKRB7dC6gwVpoWh4BrNI0i2H30hA40jzKmrHRFo0RfhFhQowkCUGz6jMawx7Dcx2CQKOfGEomnqYlVQE5nnHN3zq6d17h6eu3aEQMYSUAc0D1OZvcY4wFFP3XtadIbkVuVACSCIlfkSrlEhPqURYXeR8o1IF7s/+ujPKoiIEokFInx9g95lXu1/Z3pHi4rUC4SIIGhMPf9eVUuqKiUREjIAY+RYeKMPhw7wxi5IzpxjZyZdn/9XUMkIigQgoJhnne7TeYx9ucn4DSbzGlKgl3S0TlZG/aEDLMn25aUyey019XhUsLAwoo8ngA4Iz0mChhMwCgK6SVqZH2iotCCJQRgOTVqTo89MhI5hHCOpoyosMEdHjbQdp7Xe7s/N9pzjz3XhUTAQIDzulccTjSGsbpk17jkuqVLVESjN4JEJICKKESBFWkKaHFWSga9Ekp81gsmJ4qCkACEB/0HHtM8ehs7F6tPoKiIRCSFW8Uox+k1ACURYlMlm5ahiuPYY8yOgvhjQKdOf+7X3GU1RRWJrSnARSv+cCnSPVRX3H0ZUlRhUbgBNQcsygHUMdRWoZaQC1SA1ja4F9THTOZZYzi2Y7aGg0MAjPrEMESJIhkgILV+y0Ijc48zPBxS5hyXBSEIDJAJO9OV5nper1c29jreyE3yePJoweVUvc89GW0khHDJaxd96lNbL8GaEhmFRBJJEEBUAK5MfRDSe2zA63rhvz8yjYAAWQIozYv7rE56OS/6RXcdJYwgIhGkcaua+xjuMiQhEskGoNSa2fPk3h3s0Uc35m9n9BNXPeAr3uci2AEiNiMEoSIW4/Ver9vuvC9fkqg5E1BR1KgVKJKpAFLdKxRFmFBkUgnGskAmNzWqrllxYo+0SYEiMtPIFgUFSg4z4xJJKAgUEVRE1+5dtpzX09QzLfvsgDnXXnT3BQg1pEh6feUpOD19dnfXdWri7Hz2hIUBhDC7uXb35P1tMdVYMyysJl/+uh2mTGEZpRzhIiIFoIogQhAR+db1ReWBLp73V91KKnzsv57e35VttVBYAmtXtz/ipbt3nOy++7wwFIBAwqFWp1vf2Zk82O8TEqogsZmgtnOiXbwwyxOLaVqUmu0Y8uvUpXXpvdviAz/oqhOAK4pEQAgSiQi49nmXxxngeMkocAVkuRCByjYBRsAaQKDuWKRaEZyRR1MhAuZlDRfKFNDAbXTax+n31wABMoA2ThUEAZEMAgLXjx2psKUMS1bOWVnrUkQpSIbr6R5wtudyby7sAsEAuJMhppfr2YjV5nIG0Xzpa3YigoSQK6ZAILYrSqIISFQaNaqivkpCQC2v+ZLfWDEabURAQVSGS0rrxombn7hwEW1sdQVW1y0VJzRiSCBshgga7mpSxiAZjCEAcXefb/3QtrhEAO5mLBRLuBSASEWwKL5zuvjMtO8+youJZqGkIAigUusqI4RbkS1asQr1PHi0YJGCQymBuYUhhkfrDGDwoM///FN56ce9MGwNwODt2SqXEnHwjdQYQ6QTD0w454qrjQJFQjOEV/Z+O8/78349Aa4DJdWwMwuE7lYzYg97tEEb885LT1ikLOKioiKCpAgiSFAgECpHfZ4Uksxmc8DiN//5HcF9MQERUaDsySxWnHno+1OgCDAQReGNYzGnI9ZDA1SgOEIBYUaXUTSMHTMbCHan3nwqGpWKNTfSkkqRAplIRC4kcdvTrx4dnN6yM4EAiSgCqBOQ2gDUjgjrPECaykSOAhZrEdw2bJp14amPkZgy8Glf+Gh619s84z2fv01h8/0VXKEARASJeH1vaXGs4Qoi5XM1irigZkHI3A5fOSyqy2snVpzL3cQwOTRnDSz3LIcs28KqoTz9KsRoFKFGA4hcqIIMKODGVgsQOyMiwNgotxHe47XXjIWUFAFSMNR/7nUn96qvf7oNGtFoUIRIiN/dWQKBhQLEGyJJlKRQSyLR7Ahl4oL+2z2XckVKyw64IhoJQQIYjSIIvqR2zp7Y8yB97IxFNpDFBDAhCBOkWAiVykFlggUK8vgqMGEBuKAQY6xlD1xR3v037gEDaT79Fw+9cUuA8uWfN1pKUdxUAC7KvDARFezhLgfMuRZIEESScxAP2PN0O2evQZylHJNIsMPRwM5eEvvokJ7RhvJh722kZiGXJiIiCiCVIsKlgpCisGzmUsHKg83gyu7+g1949zGhIgYZ+sTP/mSbyNlPercR4aJBFFEJ8IfLoeAWtkZACUsVpE6mOBPuhM4Y48Leg17sEgzhwgIRoiA2rRYEhFf/1T2ymDNqrmRh3FxAGQhAkSptpA211ArWACJCHbXUCdY6UkSiKVIkI3Pv333HvlqEKVYufs69KIMCbX7S6QpCo1oUESokvNqMEcUhnKcxNdoo50GOO3DOJpXQAJQXuV8RWK2xZq7TcZwnkhJLo3zpH0YRFEkpCpEsBaQIRAAFJACNotqAASMHxDVz60Oed18MlRaXhiYe/1ROnOHMNT+jIRAEKUAk6+xrdge253JtFIFEBYgRcKDbFASbXNh7+L80CehqEDeIXGkQKRSRGYr4hEtO3eZ1MnfsnQVKEUUVgRCsIISJYYKIPF7AAgV53GoRESQ3WoFxRk782Kcs1lqgRG103+Uf3tNAoOYHfFqCCA2IogKC+ksXYaszRs4PwFOKCVHMcLxZgvY42wcQyz0GDLEHKMzZdqxsy4owvz31SbhAUUAAAgTCUlINV9JcRKGUuBSNgzaT4vaHPOVjChKlksb+B7549+Kz40z93YmxcEUREMmIKP955kRoqNsGwqaCYlCTSIicOInN4ML+u/3tJFw4o2iEECBIkYBQaWk3fO8dl9zRVogwl041ChcICAKxIkUIFKvQtwmQSS3SekORIhVC0OOCKn/sky9dpzltSAQt+zv90JeLAD7xHCEUIEJRFKRRb7rhlIdHhjx3K+dHMVyqECQaAniy24n35/lIsMMrxw4hOdmDneN6bN4sENztA/pEUqZIEFsFAosKKiJRpiwhyM2NTFINDu2Fp9XH3uVDn3RFq4jipU/9zlOX6Oz61oue8oB5xyYIhEIFJH7u5O7crEHv2iIBRSiAyNMkk8qIo5gL/Pd/idwQaWwqBQUaJbYWzEvx3U87HdbDc5cZ9rwIAohSABUrldJQgQpWqLzVWosgVuQznpp18la7X/ptX4gnWJgGqCXLevyj/vXE/lJnT7/qIjZFAAFCfRJ18/6Ck2dKrS2YRjg/k3UkBBIZAOnp3Vfe3y8eGg+PHb7KwQ7QQLDKcPPslXe13NlftppGlL8OdhMiKoggAuhTKpa6FJATUwZS8R2VVYhVyrYF1aWd+Zbfftv7fPr7rM686cxfP+/GE741tT//4AfUDrWgAAVIZKw3/O7l7WSroqbmqVCAABIEFE2LtZpnVZXUkC7wvXRUSyQsFAkIpBERAc5yvu0n/ngxrTsZNVy41utTpCIQigAmI0glVSpSESowekOtQIrcFqlCgL7IemI5f8wnBeRyEREi6It/uKRrPR72l5cPVBCJTYFrMo0nFztrMpQag3G+EBJyh4fLAMycXac9QgA5j3wZ3AkDUYFTjy0hqjDNV9wvVGM0sWkRAZbaaKjBxFb3VkaGUL8x7dwxMoRH2Ko4kubm3P78rzm1v8OeFst5qt3rFn/wmAoigiBFkMjil3WphSwFMJJjAoQIqPQ5i3hCREqSC3v3epWbJaJyBVwE4Uhha5k7Pui1D1h57ZHMnXksvB4ho8EoYQlIhYoTLNpZpTK2VaVW3KkFd6SNU5mALDqj5unszvsviVVUhBEVmr18ziN2rnziY6dRgeDaFkFUTn5nWjhJzWMWnfN2mkYhJqQhcX7s3nu9G87tQQOXXrkeitF0TsOOecxckMmI1L6seXJaGQpQgIDoU7vl5he97tX98svvdfGt97v/cqf3visqDf7q5Noe8xgOmwq9qslCqpxkWo6LprQ9jWvf5u+u6CgBS0pSISKo//QlshkYQ8CUQAQlYKxUTKUMI1zwv/fLaiNRCkQRGYqgiAhqvvTP3+2ldzk9PNLHNJZ7nXkyLYASqRiNOhQgtQgGENI0PC5l7dDUFAiTvQhAIG13v82an/CvoSJcUEERI8qD/+he91niChVFMpUUUcFoeeX19xhE7pGoqZ8ninsFcDlHmPNILl939jxvruucD+xgvdPYocHjemkPmL2sljQYcsaXuqBF5SJCFW1E7T++9YU15p0wr0/f2nff4V0+5OqL/zO6+FR780fcdpH76DEoGwjMWnKFEl7iqY+Tq91LfvujyURAFEBEqiTBD956qSuOQ2FDmSCssBnIRDKUUT0lcmFPufxVp0oAEUAUURFBEAlTk6543rvdePsYnYwKFnc0ajS7yqHiRiGMFC0yKsiEQq1gZYJUxBS6D1YLkAZ4HrWzv/uCV7xNhBEuBCBV4Q8Al6iRAFRCISIG8+Jzdi9Zr7tlup3BeRoWpUGNjCBHQl/P+/EkTx6NCY9wcFi9lFDHxNrGGpmvvqiKIlAgXICLUW/82n9YXiadqD63ee9K9p7zI4273eFptch17K7P9nnMY5itgZ1LLrtj7O7te1+LO3ZW85Q2Vp/5Ho8tcEVKIoEUBBBWP3rxydHbhl1VAFHhAlK4MBZRJZmUkQt8cPHrTlVhSREuNqOgRMJSWteYdp/xbq+/2GMMhzW93XzHZU6XY2uWsJwJt4JT5R2DQ3esIIRCK0xctcgURCgelpb+qH+6fG4FVEQIUSKSFAoSqIJbgsBV4sbnXLqK0oftEJ0v0D1B7sSjO4vnx739S286LXugi8dX29llzlHIIWmgKQUTeNfFsFLi4Igaxb/9r7rnggz2YD3v7Hvnktb3rg29hnwS71WfBwaFzZP/cveLbd1x62te9pzX1ofN69f4Xvd45AMLXA0lYjNQBnDV1569shfd8TCdviGCglJBoZKUxggOIC70140XqbmI5GqAXAIQQAWYWITdP3j3W3eH19nZU6b52qd9KpaqIVJKpUIVi6Tc1JLiKKkV0EkFU1QEhIAE02SKjJ2XfsFvTcISRKiClMFoBSjlchEVCJAQfPF0UZt6TxSGDMp5MoKsgyU5THTzHf7Qy7HtnA+WI97y7QeBEUFCYO9PskvTLPO/iChXFARlBUtP/5IrrtjvqPZTfTCP9cQqOrG6Y6I4uW77km0LAqD62ocOvOwX3ydk/0Tm7DKCLFU0CpcqIIIEdF76y/ckjkkGiQuQIyQQICKF9LIiBL7gN73mGpcIUGy6iBC4XAiiUWPBvf7qvc62KD2jT233i9/rK952UdOSeTnO1LIVgKUiVCqIjFapIo9mAlRLmBR5NIBGy+xplP7wTz6iT4yGQKDRKNxwWoigiAREgEzd9heXtd5xn8cIIoTzde90YYIQD0V++Ve/u7m5r+WhnRyf+XUTQUigyYKLKxEspncfRSIEAlJEjX//+ru1PU9aDSfpWs/an4Zmn3WHHcvGcRDbW33dukbZrElXx5oZo05IILeoBUSEEOBp/ZFXTPseYx7dGR4MMVQi4AQBSiwJ0Mhw48L/Lz1uQRgVmZQoAmMKRQEppbIg40HPe/R1p/tczdPa8Mcf0BatLbhY0+7O/ne93wdSbCdUqTAEoExga2nciZvV6yGjIsUKirjs7GK5s7c/LZbju585jUyEzVACCmsKKoXNIIFBNVbveeryvuNMU4taDc7bxrjNCCHJR+CGr/JPvun7L3Vzf9ju5fYDfM+fPiD0AHIQb65P7jdDuoqrL2uRhAsBERFj/8tz8bL39WA/7mNo3pdHjWnv7II23UaIKWwf0H97t084CwFRcoKtowGRJeQCgRUyVB9//cI7qnh/zliPSisaEEEhgCLa2XEffYyMEaMLfh/+M3cbg7FrUgkBhRaEJRCjVKBp3OVpb3d2oQlHQie0SFuqTvWJ6/XoU0wEdS8otykgEkHONQekp1qdJFT2QgSisFcFqLHQP3/RT5RAlkC0IaJUAjJABOAKKGlf8uqHDtOn/bmHrtbPm1GsMpaQgAkJ7Hziz/5FN+N5aHL9wFt/3Y8eqAFyQhK27bI7RZKi+00WggqKEAlpf//itzmx6iv1uc/pg5GZSX2cTUutpmG3VsOTxYHTew6ta7IiIqQoCBoulCKpAogqlSZ91jPvckrK3IezTjnGQCICSthsivtuSwVFJhf8+KrvmRxXRVZmChQRkNhsJkqqjbf5xN+6ek4hRRAv1KLbdx2/zamxUkBIwdpZnQWfDh6+OZL7S66enx77aD+9+5Ffu4dYF7cVymBGSmLsxc+98G1Hi8oBRc0JWLJcTkNRoorA8DO//WD1mj0zjwx6ct5gXtMaKbAjIeTkL/ytP+++XYeTj/rJP+WdhA5Z4AbsONzjRENKuKr1KSgSgSBVwo9ftLva72Osxz5rZzhz7XdluHPJO979yuvO/tFa0XyY3UHUKIggKSyIG5VIkQoIEhaj6rueepdLBhmjj9HHsCxCEEQEbSTJcEuvQVIWuuA3+WM/6d5Xv7k/qKZ+arGADBcIZAmgItxA40fu50XUKAnQXHulE3s7ddvVJzogIFIAF4Rv/k9/z0990vPuVVcb+uK4v57nZH/6z/zf0lWsjwikADHFSBPrj/yPk1FULhCWG5RdvYUGuJAQNVJ/8j13u/gO5jEvOn12KuP8aeO3f3wAAQKeY5L5U//V7/qNd3cnx8f90h/6EisRkhKzkEg5NbWeiCuRK6NFpEBAan7pRZzt6z5672tWJniPkHGf6Rset2yw/x8/9Tu3JU39gFO3KQUgXKECblEjijRakCUBCDf94M/c45KV1vOw1x5hohFQoiGjpDqbETWQ1Iz4H4BT1cmdE3tmsTx58d3u907vfdkSnGZJELlQ1IAal3/YH1xG0eY2WlpvaNTe5L3xkWi0ToBKpfa8/Og//9OPPv5oHrjvvDi791oPXR32/OltcKAKBaWC6H2USHrTKx//M26JygVGDQRFKaJPlIZwWal/+aKLr9rvLnvVR5zYnL/hzzOxh5TsdAAEcv/OX/Iv/58v+uKv+e985C3noVCDBImuQdzCMTyUMqfYbASKCBfKuq/7qnvl/b4/mJPBXHOZD/mY/7WbMeTFOz85T/u0m2gcuPefH10Nl4tCIVVBEKGkSCGyhTGNT/6R0xePCWWeZ9xHBkhAREMBMREBtRCqwShkcsFvSdvZXS7knXl1/XPO1IkHfMAXXFy2FBRSRAoCiu96qBettUJjdLxuYrG63afeo1d2QCi3xYd1+dF/2etf8cndvXs/9NXatNf9MnXv04+hKHLbgBBlCRDlUcsn/9YnrJaEiqCCQiAVJRMozZRd05981RWXZC6PedY8gp3y+SJcbzh7EQMIDQKE58HBfsJ3BTrvjoEcXAWTzAQvDkOr1rqzMygEAhApAmN1Se99Xo39PrrdCZ7LfP/jG32aXNPo5Q+74ct/xAeN3//pXSiIoogCRSBQxVVE2AXxNH/Wv566tGdkzMOj94xh0AAqCRFhMwQ3lBCKiAv/k+vEianacnFqatKZ9qofueIH3q+iIBSxKTbje9332ouCUGDQmaGxu77r1eWuSsFHcF/6yb//a7/sfp3dc2b2m7KvUkqv9PV5EeRxgVSgGShpNFv12c+7F5ECCIGoBIQSgVIM8K9/60XLjAx1jxp9xMScrwGv/uYj1zergDwuuNrcd8z9eCQRqsAOCQaUTR6iJE4DViQiiECwPOGxdu/rVR+2cQjm/b9YDIGgNLS//OGf/CJABObXPu9dLVFBEVvlAiJUIiISWO36z3rJpVcWrYbHvO4GItsFBAmFckAS4LiBGVL+B0CDZZ3cse9orRXpyyvWn/rdn3giAhCbERGB9o5/kBKL9WJkaK5ES++deeC0asUaRivS7uN3/WO/8pM3lzP3Ha9iuzt2tXTfwwQm74Bh0SseCXSknfkDnnO6iEgAIpJiU2yKiP6lT7+m1l5l5TFmeh+D8/5Tb1s4ZO5AAqEEDeel9YhhJUmHBBGQbo5zN04iTipGBQgggqjd5brV3Ps8Ru+mI7kMT2mkNJpAKbesv/DzfgYhyeg3v2c3EIiIIlwFIECRIkEo/vGL1ldaZK+7z+u5z90kY+ACBEhQAkJiJCPbSfsfADCplrs1T85gykWLRd30fT/9bsQloggSAirzNY9WGiNl0BgQC2YvRkAqQi1trv+Ar1lwbpnes89dzsu6dpJPf3wDlLdWSIPKmEI8xhR7Nb36YS/a0RTJE1tVEBEFmVGq+Rt+7+r51M7uqb22Go6KeJx3t7/58umyIwMIIpByGV4/FMFVBHbIjuMt770/U0p3QCAOFptZfPY3v/rEalZGWUxDaNq5/X6XU73RCEg0vMiPvst/ZibVePYPLruXvQqBECkgKYjYntToP/Y9J06tPa/t/f3Vaq/Pfd0D0yQmCIcXqsVqrGt4WkwS6EKdcoiSFtPOXNOc5ZDndV3yps/4zUcYQMhCoth6r922nJvamIxJGXp4cHmaABIKWO359zO750Nn5zrXmX1mONltmFpBmdxICKyVXi2lMgS313/IP9rlEhEiwhIIddWYxBse86b7nhhhXve5ew7zGENRzivlU/6Sc0cgeU4R87GUgHAgswHhXI5N1VQKL1eUaCMIiOR8QvV0esXpSky58wNt7iVLighCqeWf3luLPlUWe094bRtMSYoIUAQqF4pEYql4yee97MqLlcq8Xs/zKl57JJYwMSknBwSgPDrNmQBEXbA7vCamadlqmvYXWu32qsuv/bK/3ulxERUIgoJy0eVDqK00pjGg083gvRkVHq3WCfS4/3H3/fGwz9nT6957fOisXKHrfsXWFiswJYBe1jBOiMOg/dNPfkGfQkMWyFUkhavJE/3J39Pedpxtfaznde+rsWpJFMJ5Hf7q1V+NlQMkhB1WsiFDcFXOgzLDndPJ2T2do2oQN/4VLIJwOQhZ4aL3+62ekZkMGcCLtHcdTaIA4cKiwtWP/z9tYbO/fNMn/DbDCyrIEigiKkzhktY7rB/3b3XfWpfHat5frVfzau5zd8hAVgcq8hYzGAkpSFlBwdb/AEBTTeu2WLVpXWMlql/2pn999ypcIrjhIgpul9+SBVVRKAMMxMWhrBOwQtrQ+zc+XO/m3PfdPefc22G/euhV1o7FGmTEAEOUajAEaV2jKLXHv/weQ5ZSQSlcqVCAbnvmV910l4tWo83Q12O97h05g/Nf+aB/cEwuAy4wOTkMeB4BzNYcacjAdJQ11+GcUEbhNbNSRKRcbNao8BMPunlnbkRlBcTIAy/RqBYlRQrKAsaXXLpXiZKdP/75TyeAoAhEREEIijD7L77w5nueYlQfzPO8v5qHSQKwSMOFcI1KANJSitokJ1C2o5b/EVCZNK01rZhW1ALm1V9+U0VEEQ0KFyK6ZCxZzjXLmG3F5Xi4CEVwQpiaPvT+fh72Obvn7HEPb1rOXPtgijAhFCAVEmQlVjdFRlV/r1eUE4QIpCgILm5925y6+oTppON5f557dx9DOf/Qy37Ob9pxDZTHJVk98AEqEEOYCOJ0WcU9lZgbb7lqJkSiiBjNDcZFP/9x+zuzBiZAaLmySY2gAhFUoPQrP+rXTs1pdOlLf/1jFBRFiEiBFEVSiBvfY69derLDSr2v9vZ7n9d93QebxirDEA1EqCIkgWql2BJhcMFdOQagqe8v94u2txjLsRh7a1AqKKmkIraeTmuTShWLAGkwgRa1FtrQgZXXL3jjPq+95+rM7LJ7OvS6ztUKECreUIMGoaGkSHURnXjNNz9RlSBAKIoEBbunVjt7Y1qn9zGmPjyP0RdGnP+B3/bffAbXAwFZFSAYgGUSAQQQ5GHSztAlaoXC/LwPUIAIUNRSWMqHPP6JezsZk0gZYMxTJigiCAoQVDyoVNEcpv2P/4v3M6OiIEvIVBCEKC9p06kez/I85tV67XXvI8pBkwXFVoeEioQZGQKJChfmr/LqjiOhmf2COjtpwWp/v5QIEQkVAoi5e69p1ZpMIjbHxG2p0CJgrYEOrx6uZj+cu9ezj246bJ0hydyAvDWwkEJwGdIFjL78njfepdMUhGuUCwtGdr/zM69W9oPnsV7O84yTYe6sP/Tvf1WJR5XHk5ABwwBWIalBoDqZyKAh+P6nVYoQQUQSkRrjCX/9CXe4JYiyMi8f2xuWEAFQEBBVa4uZqWe09mF//AEMAoqIKEiFNITe+zOfdDnrucZ6jHl/PdYe8ejbgkYB0pYSUqiQUfTJwyMZFV2AW/z0J558zfu+6kgxc1Fa1+Isy1VvcgASAUQRatxelYhi6mQL5vWSSioCAqMcX/mhh/MyD+c5u6fTAt07Y6e7FKgweURA4WSkkGa2jrnxoFsrIFAhKapAmj/hf/2JvNboffRVn+fRPWJ0pxDv+wb/7MOdfAQSGFjlcRNySAQmIWHvL1x3EihR/4IMIkQCpBTRcrzvf3znb59dl2MIl3/ip5+5JJUgBEQiArUHtBP75ZhU/HHf+9nTqCiAAKuQC1hP07c94RfO3E55jHmMee4mIWyf7RQQjAgBgaRFUtBEqJALbzs/8IWC/MSXoFDeUIRCLDur2pmL3UfXhAsKCAhBYGenTp7dX+zut/XUNyba6qf/ZrKCsiGIKlNOXXbbFfNtTZIBKgMiT64gpyFAQLE1KaZkmiqmMg2CvMC64we+em8npthUIkFT8aPv8qrrLlr1mjVXa6CKijtn4Lb7/fPbdC+JiCIgCsXWCEIkVzASYlN4Mn3SWWX0QMYDrgqNCAFBgAStrvicjzvxmutY71x2xaMe8Pn3p0AgDhmU1B98Rlvi/YwMt7548id3ZcJFEEQEZAngq39ucVEf4PUws3c6aR+5dICJpIj0pizXHnans6fYb7VN3/9lSUW86KEcPgRErSurLHzi/awUQITYjAQ3qrJojWlgto5wEylHhSKCRFS97az7PJbz6EOd2BjQUKcNsZ42ICgIAmF/qqBq1EgUQhK1b3jpfdOnECRcbAoy/cU73nKFMtMZDpvJnYQm3fbwP/xgMzIRwmbEgYqrIGgUFFFQEJRVU0e2K67ae+pXKYUiNhUJRRqgS9/uoz/8vd//Xd/xEXe5D44iEBBtRBC5nvjEy9q+ztqzxbRuP/FZlpsrhAICCiIK+fTfvWyxGuqZzp5sbbm12GARRIqMBmBtEMjmW+9f+OOjlHLxkkeePQJgCIbO/RcJFoiAkiIY7phqDAkn2TJYc8ajFU5IpMQNV+kR6wz1jJE1fbLZDIahVBUuIiUIEBFrW7Y7XQmbTljM7/3SHQypCCUVt4Ty5d/xxTcvRsNrsAQp7qQZaHf/Iz7/5JMaVm1xgSsiimi25LSKyoASMBU5UmnMceTmZ37SqQbIgrgUIlBTFNMia7gvRYoIBSAFloynT33qlbXqqzGPIC9mfvlxgxZFrgisVCIF3Pi037jy7DJ9zJzj0A6Pt3J2FgUY7QTYgSI65Mq33r/8BxMpILjp3V/M1FE2Di6kxZlTfwrlIopwhYrkCtcvaaGqptTYkKf+xtsvwqooBSgtFNKH5g7NS7snkThqk1KWQC42I6DhAqEqZUOp9Ol1H/EnrUK5iKzQkAz5tK/6jstMT0nVUyR3FmDar+WPXfWxT9kZC1fDFUglQEAIhWYEFUUKkkBuYYozbDSm1QsfOyTTLFSgSIgo5aVAyJpMIoVCLgmRVIDx/i89Nbr7PM9DUtO6/cInjcbWDUYpimIVjHf9jyWxqAMgj2cscVVsSYEyy31cDsZytNLsb6194J82QCQU8DufdTsKgHIA653cceJll1iUJQIVRFBZ8+0Xh1aMDA02g9r+S995LkCB0SIURhsP2t3f9WoMy2Q+kgsj4aKIiBRRjIySkmGxOYrQF3/12b8wIIUrksUoibh920f/4WV7LXKUGHEnnjWdued7fEGlKJxG4YICAZGAIi0hQgAuUBblGmHEMVR+/ctOllwgwAi5EWgIpyhLBWWlhNyIcIOA8tpHXdz6PK/H3EmyhL329PdPJIgYDZIKMkq5zj7qxUtcnZFC7SPThhOotVYMLgh3QxadWBi+pX76umUUCRe4ojd81j/sQZlD9sWKK15wTVCoRBW5iASgO25+4DpSlHQUQDUv5j/91lSARA0FoDFOvddfuC8cDQLKoUL1hYgpICEQgocihCkONMDQrz73HeMKFUGFUsrl+Fce/KaLazZQKu7cWb/9TzziZKTRKEDIBUFBCjBefr9FooDLJeQiqLUhdaUzpCzO/sY3zlNjwxIIkShAqoRLopBSKaUhuVBScelLnnLf2+e19tfDs1M7nhf7J//qXU0lBS4ggCCpJLe83ZtOzbZ1gJFHBU/BVsQKdSg9mXPGMpZ+a+1FD6KAuET6gnnBm77y9/c5UIHMJ774f59wiRQKIQWgoLF4yapF2aAIgCH85mwKUQGiIEa1fNLnrluGkwDhyCpLYbMQkYQQiYJJtm2VF+/zxlMZzSREgCs1lPWJf35wN0ZA7mR8+1fsOKQSsRkVAoEi3PY/iV88DcgIARTI6YMa6WldKVx//aiHracIISIRISSQBFQkJEgJFJSGnDYy+caHXnb69t77vJqdbu068+7eiWc+LAVukFRSRIGUa1Uvf/Ttu3OHWqCPTdtMG8utBRoGyqRsmvIt9qd+RGdiaxSZGrVqt/79T//Hjbu3AamaTr/Pk+4DskYDFASgSHjiN0jkEIi2KAy9vO/KBQgQggD94y4943Iijl5ex1VsdUHE5hSSMQAdTtXq7Hv+8wLjggKkiGad2L/LL3/cxdU8RriTn3zWw6yZFCikiNiUS6hP/MPXX//YS11AHSAg8pims8LuqiSq6Su+/1KQAAQRRICAWGlCRiqSSEgogohbPul5d1mPvpdx1t3RKJWlvvjtx1RExIGJIkVJzz9/YOm+5d1th2j1EWTCo20FpfZbaw959skWFLnYDAKX5dtf+4pnv+IOLj/5Xo94OyAIFVsFRAgo87eXrBdM+yzaftgagPmpH70jNoMIQo1a8F6/d2IvCeUjmR3klhRQESCUsYxbSVosuzhsZq13nvcDXzcvCeCKUIgEu3zM4379VK/FolpJd6KveaIoJhUIUSAigAKY9r/xD0/fPvZ3oyCAiAgEY577ImQEIuz5Lz9PFBFBEZEAEaFKAS5BJCRARGBp/TPff3mb1/N6te7uPTLNU5fmxT+8S0wL2kiKzSiKeOaHiMuuwH7LuNzQABSBQB09L8e6taxtv4X26nvNCwVFQQRE5JBSCly4ICmUYruIUMRmvenaixerOKxdqMymAj/7NIYKEIAAotF++P4zlgYoh1NAKVQQkUhsyiVHlCkph4BMs77hpQ9wYSkhckGApJ784DcuoJQC3RmKhv74wwhQECUFERFbI3jG17z+0hPLswlRCIoIQJJO92JgJwog/SGfbKEIRUFRQFEECIgEiiRLCIUxsfetT7no4lrt932vVyOzRlkSlTLPfaiVIgoILBSEU6k//+hNukN5+7iAFKxAhUo6uATQjHzrvL7zaypkwhXhwlAEZUh4ipIMLyQCkaIA0mgofQHRr42Txknors72AH9z0zUtkRIBllxKy11+68Mbwwrh8IFquIRLkdia1MqWheMQDq/R/Kg3LntVLAgVRSHU2H3aO8+VVLhzJsnOVz4yqUQCoWI0IgERTuMPv2lcs7+uvRJDbIqgRBBp2fCcGokqUPnZ1Rd10iIUKbgJIkUIMIWREimIoJC66cGLS072vprnufe1HSNoBKh6xiONlFABkRSWCCm+5794Yqd5W20tlWmhSrHSDq06tNRvnS1eeD8zAbIUEY0GkUmLG6lRLQssEwFKCqAiWACZf/R0Y6iP4Wrr8jYF8iPfmTYCwqJIgVz+4Pd+lkYPFD4UEIGIQJa2UImkNoKA8uEic8sH/oPkUkS5FKmIqP6Qr/nO0ZqqiXb+qeDir7urU12NKApQiGwBNb7yt66aRu+9TS6hICBAAYrUgkdXJQGi+o2rPmpIiIiIShoFwjKiSANJRAg0WobqM//k0hNjrMbcvR77RBGto0axvuwf3rZPQaPkIrIUgT0F/af/O5dzNev6FrugWJ1FrVRsiedKZTXDt9Db+NyfSfoCEQRJQcCNQKiiBRhCUWRJirCUIMD1gzc8qDpRPNwTtgeUH5kXkSxGg0i4KKLfebsbmlDMkXsBWygiIoSbR1zgmHD4Tol//I1PckAuwmYUUPnb3kNGkpLzL+Fh33pJ0jzFpdEQRC4EINCtH/+qa9qqkiwKJMBIBSJlJkg6DrNAUIHvP/l5EXIjQqRAAU8QFDdQACEgaWRMdzzUJzS698cYcw8BBEqbqNUVz7gfLo2GywUERZBI/5V/2u/KtuWthUyhdTEBCFWcF5muh27T2m+ZlXn6B3W1IBIqCoAgILOpgAoqUAQRFVEUgW/7+WsoZcQOHAIlnHnG+4ymMs0VuSjAbVz5Wx84qsxxjsmAIhcCgUBViQgoKIeSguszX38pyGU1QAQRavhvnuRqEnfK9o1ftOglNvvUOPxoRP/82eu77twxzeyt91IhCmWIEskp0YFeg7A1qPGUF3zdYiBQFJqlSCIIoSIiQhAUuVlj4he/qS3X7uvhMcZgaxmqakzzRS+9y1AlFUAwJEEidPxd3/eDR/d+B/eLQkOdVIG2WEdbNSPlW+aSedV9GC2CCJLKNoUIAiJFBBAEEVFCgDF91W13b6N3D3AXhwzAdz2tBReUKxWRSJYf/ZOf26dy+UjyKCiUCogomBZLaSYQDh/jTPvv89yaFwyJpKwiFoTFU999oKo7w10e9E6jakwIMRERuVJbKqlfeMIld1Pf90BaF3JZonChgsAYC2OGTLQBaBp/++ffcLHLSosECIQiiQhEGiKSkEiI1D/geafi3kfPHLIFBaqoxXzpC66MiEAWSgoXLni4+wPffV3Hto9RNtBCMwEqAdt0Z53Dtn4LrQy7b7g8iO0RQAhUXDgNQiQQoCiKRAFE9a9/eM3Ft7nsPpw4hwDK//jM9+oUo1EIQcpVQ/Mn/a9fTC2GcpQVcoNIEYCQhVIwRiPNHD2tL1/wRT89pYCgFCDHrfXpk599OVqodN7dkQ/sQAiuwhVEIdeGnMc/5X47IyOzpXEaRksBSQG4rGQaYWQklWwrW/rX//1h79cmIhFJSBESIG8IEUEEpqKR6V/effdEeT13e+iATTXV2Fnd73k7AgIpNiMskbuHD/+x77t+cL3yeIZRACFUROi0q702HRXst8q2XvmiK+dqiScXh3ZBkZQstgslELFVnPn0qy/zqfUKaaQIhzerb/9rBs0AUVC5osZJfuH+r1yUUI4wLZQIgYgiQG6ts7O7XFWband2jsJgzZNf9oAhGmKrq9GiBZd86wuXOyd2d08spPPrVV/XJ7kkKKAQLqCAuOWHfucBl3g1vNdn6+ztroYiXMVmoWK1mvtUY/QkbA+KdNuPvvLrT8CQxGbEgXIBoyEAQcRm9J7/eNG6tTbcx+DQSiXl+/zHyepNbAZxYNarfu8f+yXX3O22QI9pQ7it3Fb0wof7xatvsBYo5VvrF//9OwxUDNUoRRuRC5elgiLagotCASJc88fe+uC155E49qijgP7mX965N7cghCBCQNDf3otqGRwxQyqACEXgQrgqFhANwrGOd7jhxHpyEYhUgCTIJd/zlftjYtEa5/fZb/wHuSEOWxFb3a7/qNfe7/R67msPj0wZlQhIAUFEU0olpjESuQ4QSCx55jec/LS7TazVIqLRDqCt1VIQhAhie85e1HaaqdlzlMNsltSvfMHVYwICiO0RTOYX/33/rCyzpQSPqxNGARliMXCuTpZa+db74nu+ks7kwlBYIooSykkRKciKKgpblRSf8bKHdJJ5jNXwsI9WeewrmyUUEKAQWVHu+a3f6dF0lMXUxaYgAJWEWaSD8cAc7+Lsh/yVPI0mqwgKEgHe7+l3n1sWxXlev/y2t1eXi2hbpERiqP3tF+jBO3vrzHNPn5mH50WNwhXhAsJgUFVtdKVjbdtafVrW3333fz784XddooHcevOQnPQdl8fkBgKlAoikfv6zT7i8M89hcERFE/Ndnn+1KYgiRiMITJBf+2v/4bv9kA4wc2SkaZFJmAi1VNeDD+QB7LfW2uBt/+GyXlSUMQFJjRYE5QKBgihQtM3UeO8b73nitszzGJ39YYmj+zW/8wkgIgsXKIpU4E+/z2t23aZ+uEqxPSkikIq1ytUNqqpxPL39/U98mWmhHCkIF8hqX/lVTezsLHVeJXznw6OhSK4tcqlgtL1v/J1Lr2Bfs+exdhiduZIWKopwEaVUSpRoRSTloEoV1Re71/7G7z7rig+9elkUZSbTJM7uvemKk/MyIiINcBEN8f5/v0TLecwJKIcBUdPq6pdfvF4CLiBiNKINm+/8D/wHj/3qNYAcU4o4QmqhjfVIh/Tu9XYo32JvjN2f+Aw5JSAoCnIRUUFECIUI4doobvzo6+/W7kjP8Lxai/TkSKONz3jTyT5JSClciAign/6EH1isZeVQE24bLooIFwqTuz0VzpzB8SZ85b+8Q5cKlJQLBbGC93h8b1NNNel8wr75y7+mSYCCiCg2242f95J7nawx1p7n1Vj39MVY40K4EAVRxGjpibpxhsLBQQBNy91x+9/98N9d1+59r+VuJtKvf+ULfu/Xn/aca764QBKbAogMb3ybvcUiHiPmGKupX/nSiyi2yhJEZC+Hv5i415xlnayeoRIqrYG0cteJu5tO+db71EYe/E+nQ4pAQQQRgACxmYArFRHxT1+6f9flGB5jXvVhzwnHmGnvi351BUEwqgAEltryIVfcsZhb+qF2VIKoIBAVmymRhEidY64wHvPi3VGhUCIicJV9r+fuLpaLxbTgfP+Xj/ojjQYoHNibXvHJ/ZpTt5XneT3Pcx8ezEmRuNVokSu4sFotRbPpHFGpZKFhdtoJxsv/7al/8k///C/Pes6/POt5L3vzybtezjM/8OqACBJEBMG8+NrvWyyxM+c4WpsXdfWLdkHIlAui0Vzk8vrfuS4PWQ/AuC+THW6tSukAvDLnLkax33JrXtR6+sHPXSRSxJYIRRABriCBUES0/ranXnwN6gzvz330PlfIMdj82t+8Z18ApEREEK5Zt/1WO7WiehuHWWBAgIg4cBKUBKM47oR27Yf8bTVXUCMBV6H2xp9anTyxU63adJ5l/tUHv67hUoiEi7Q8/ZsXdx17i9VYz/M8r9ZtXaPoKWihIRdQEY67ghnyOBwawmkTaZ02nTxR0940zO7i5GJn3q2bn/Z5nkQkog0Ig9Hf6cUXZzgeOYZCXuo+L8LlckHE5vXQ3/xv/clv3wmL20xrnUBFEGk6Dz1XZ2aE8i13ldWG7/LVn3tJIhHhAhCgKKoIoggU6S++5baLdnrmeazH3FeM4YE4TsNjn3dXASkiQECj1p/63B1K1FAOUBpEMFoUEQSGlRN6kDKOCxiLZ3z5j1JsjaAgt//9r962M7zTpNJ5BvqhB994hUBAVKAz3/SXl1xia6av1mO1Gn05FxUpUgRRAUKk7bcllGeclsNBSxEq5eaaVVmgVKpptTibnX//4NMFCEBsjdt68fxHlzpjRD6aJGVn/U5/dUIEgSuC2Df/t//i629f94Qhs6lgDRMpUlDwzIFO3fgtOBTRNC+u+JrPuIz0aikiNiOx6UIIU/r7n/3Hq65p1Lzy3Fcz2p/HGOa4b/y0v7YYagixNeM/vvxlV+3k+sW8n3BgeNNtpx2nEIoEpMw0u9VytZwWs3JMAWZ+9o13x3JFBXSm1S1/duqu97zrFZecPn2ihM4zznzb+z/1KluqCIg+518uuWzZ1+vetVysp8W6zY2Mdul6QehqbAqgccLpbb/wsDmiUi5FRCxYOC3TTpMV15Qpb/yXTweIQBAEYCmPeN5yGmOE4yyR5vHIfyNBYjOif/Jf/ce/69XLb37jcm8YWsqHP/3xHWSCgIDbzP0nZ3rE4yrfildlgsXcT73v1z3ihENKcpkikUAwaxqt8jff/crp3if3h/bt9dzH6CMesXJc+pvf+MRRk4nYdHHdV/9tu2axGnIqygGwd4JKs4gQgEVxu6Y5giic02l++DMehGI2PfEfn/LruUhjqgVeTBTn/Rse/8H/VpVCAG/8mFuuuazmPnsVr0fvwzGJuI7CNeHaOPDmMw0N28g6ApFTcnM0MDVET4LDunrxd596clQFCSAbQa5/fefFpD4YxwFFw+O9/xYXGGnQ+M3/hvklL8/77D1ttmlX58UCCO/YNefPrFff3JQZx34LDoqWRXbX9hXv9e33W4KHkvQayqIPGgvIS37+z245ce/SGeQ++rw/+jzmkRHCcUeP+7MPXpUsgD4VP/uDe9ecmAcZHoNwyGtvvGa1zKDRWxSThT3d6F4aGYxxTtTbje/3yuVqoYyG2/ip7z55+tSIR9IqU+n8403PeZ+/qxgo/uCrF/dcjVrN7mOfvj/6XsZIvOMX7V8SjWb1Bi5kiqXXzKY3zRw9YFqHjiXLozrqyqrszs2770R5EqNBQK4I13Xv/MYdj25yLIhSW33eTyc1L0a85Plf8o99/AvenO3unMPQnc66f72hsxhIK+318vobrS1DpnwrXqFU1apNi3X6RY/6oHd/gE72JrIAJkh/8Z/9x3/e5NP3mMait1Uyu6/7fuYx7Jo5h+EjX3W3tQGsNv/HV7z2ipNtJY8IiUOffcpXLEnVcM0tqbFIpsy0rmGgOecgst74Sb8nZ9Hcxm//0LWnT188LVGpqBZyJ9D099/2bS5P7H3B31x5zS3Vmdfp64yRfZwkyprrbrqvcReYGq2GarTF2ivh2ncLOgqg4QaEOZQyqo+WZFb31H/5q3dQkQaMRmSJ9Omxfzq1PXuIHAdVqtGf8kkxUcu1X/VXJ3/xizfNXDPDrrjOY1/5vf+Cs6tjhnEbV2btNbvd7J39LTmQqGpT07Rjsl4tTlxz2QPuetcrrjihvXu/6tX//pLX7+lEuxvV57ETz85++jy7r+2RwTnV+u2ed/c+qvXS677gPy+924LZfZAej8PpW//Xt56ctJggI6VFOjf/yxsWjC4nGeJcBvz7v/ipHvOi/vH7Xry8bKfNjsdAlsSdccx8x+vu3ZM3f/Kr73vRHTVWY8xjrMbsdYZ75MQ1f9333RscF8WSUaLxmjtOlNceLeGYa5CZFPRyTIYz4qTqzbvvkiLRKEURo4HH4ilfMqt7jjleZfLuvn/3Y0eDvW/+vdVdFnN1HmYeHvZml5bNcf7Xf+eSoyWVCyPLv+On/KTdM7L51ny50qpYTk07ariNWaPiVHmRarsnrtlZL90ZMLr7nPXwenYf3Ro5N1nc/PYvvFqi3fHEp053pWmevfY8d1k5lO/4ifudvrcu907CCS/3PEZuSZvdR8Yc5RwowNQ/8w8fo7P/+q1vXFwu90XG8PA6rmql80wACo/4w0eOGz5U99HNU2fMfaz77DGvs2YQM43id9/5gz/g7ieubJw9u/beTr/j4fmrHxw7ezXNybEpiQBCLJI4jOoaQr/ydYWEgCBLANHfPtYja6859vLEyI998sV7v/hjt190Uj66z73Pk4ECtWuO3/sr/7Lv/LiX73n5wXH5xjEf++mf/r2/c17xMOeU8q16RWrFsjQ1Jmlip6bmVm7riaUr6mPlll77u91jrOeeecROhjnH0S2P+u0H+faXfcX6buqVlXtnPa9HwuFrkZx5eRsLioma9sK06Mt5bwzb6eEcBkDwuE/Ts1552WXT3Pq0P6/X3cMjlrmz3vz+l9V197vqzMTac1u7e+6zRwYztnAobvr1RaZ2inWYpjH1ZFov7oicDPmYNgNJjONEzgxRG/gN93ubUWkRES4IiHzf92Yk69nHtrlQu/rUmdSlw4uc1+ncz8N57magZTzXn/qbk+W+6P0lsPb0K696ct3T4Vv5WliTFpNQ04KlxpKmis5MzVMWYq29GtZ6PfqYPdfsYcfi3EpKveGDLlveNF1+xbrU3fvc3ec1wzmc1Xqi3RNpoU0sdpwFq+pr9945D2fV3s/rHlfteqRrVX09r0fH6smdowbkzM6J3ZuqIQ2tPGb3WbMNiDDamGpRLJnEbqNVqnUv7pi8LneHc55gehKAQKjwXf/7ZCZLiiKIiG758See6Epfux+XotBq0sW+FLWFeh969tzn3ucAFArruOzlepX6QS8WfPXpHLtr7z2D/VZdc5omtWpjR61paoUQxTRV4jkzGoni9B6kUU0C+RzFc9VidcNyunR3XsyG7mR0kyrncBVlHtmfhhjrtGhaixpkDCHl3IiqZe1WFm2Z3uM+8OjOSJXuFBlA1j5x86JC4WLMTrV9x44CMEpen6yqaWqekKjUlJ2RxbwW+NxIraXVPNEiWZKgWL7xWV/Y0+QGcSHX2Rf84j/eeslY7636ep6PabvaVFcuqxnGfc/Omeu6m5npUBiufPjSIx47R0Ol59Fz9sOe1pZv2QuqNFFMUptqcrWWUsuqiFLdYmjIPT2a8RhjyOJcSyRt0cS67UMv9THm3sdIwuHda71cTbu9QSGvm61B94jj6pzbFGM9rXdWe6dWGsaOnZGoCHfesbPaGespaUNxxmDIgxkNbZB5wWqxblObrFZuUmc4HswDcm5AbqNVl9OsANGoUX/+EXeLRBQ32HvB9/3b/vI0fe6rPa+8Pi5lY8ri0uVE2uCc2e199+Ychket+81KDkII6Zqh55xtO3yrX1NKmmhTWkNVDUlKKcocU3NqeJgxQubQOfcxalkvpr1FqQ3QWGee3QfmqEnW02p0lZB62xcZ6XRvcq5NeaxXy9W6LA33uAtnHtyZa9Rqb6GqfQJz6Ix1Qoy3aT1lPbV5mt3WTeuiq8szszSkcxWCATQ4UInO/OT3xOulJM6+9Bf/4bqd0z455xbP87qv+ziurUJ1elfV4ryh093zmLMOj5eel/PY9zGZKCln9nU2s7Hf4gNVqZpbU3MtolZSaqjEOupKyEhPPJo1xPmolB17PS2GhKd5rD3Sa3B0jzZabzVTzcwagcHIypRyrlDUp3T1MEjGSFYjUkN3IjujrydLIiOzjBXCISNGG2P0lnXbbyUxa217aEyDO6WYXvTUr5xOlN/0F7/18r3FxTsnZ7dbdceY97Lu3edAQdIlU7lJvadXx8y2w7t2r5OjO3g12uENZ/culCeALSpaqVkTamlWi+TakywHDwaxZgLKeRAGjHmx36oglbjbPZ3jzGA5N0lKzYWVgWWgi3Me7MxMZ1GUztwFchzuzB7TXKXJYoaBOoMjZjQzMk+jiq5MsxN7uMvcafNPv/i66z/ul/9pni6/+NJbHcv7vmP0ufc+zsXWdlLVkLkO7Cl0ePc9W6+GTMisDWd7Xac8FVQ1VxptLgoJKkAnilMjCYkdCOdrl+ZFp7kZ0TVnhOMN3Z0KEh2cuGZsk3MHrjFXXxXEYXbcM7izu2satdaoLqljAcphiN36nMmqNpOpM5JuxJ22BuN1r3/tqenUqeVibxGS2Suf7cOrMYdzrRNFucxZOh3LZ233OkNyMmcaGHbnylNCtShqqRaBlGaY2xAxweUknNcK7ggVo3nIg+MCz6IUQhpxRhuKOE+Nx1gXaFhjHgnJnQ3XXDSH4FYDA+HwHnPNaZWikiixRsKdN8DONN9tUW47aGTUWHu9N+Z4dJ+7VoXijEOG8tnPXNeuao+9Y9vslj4hQMhaSaWCQIvsFdJQ1DVXapxfAXc3kNwGic5BxlAJJRFl2UPh/HX1clHRrHgk4s5vBnOKmA7FccYDDUhBRV1WuJMry34JC+0uEmF7pJ9dj16dnDPaaCmXWzj9OWE48VzFBzzd1srTwzQRnCoyUJCFwNMgMee7EXHNLWsZDc7hKFtpDCVW0gnncU2mILGMBPUWAFCaoTDycSgQUqmg0FvEnT9j2mlLKq0PNCdZa9jhfFQkuXXAlp9LA11txCENGfsEQYZUGj2kUrHmNFc6RDnvIAwGgzKEc2ri6kXMAIXzey4LjQoWkLxFiBkpI445ELmNckRwL6Pc2Qht2qmUycxAnWHMeatQAwUyn61T1tipZJxSeIoQkKzeHGpE0YDBYDPcic15acz2cH6bgYEOMW9JgyGQY9kMgwEZgQwU7vyp1lqNMhliZijcOYefw8IG7LBrAYaniYEoAUsoZRQU/m9X4b9mBQIoEN4SVkkpFDICLt9J/lwWwPLUMQYIgCEQ/q83/Bcdtoe3nEKiOsi8xS7Pm0VoDv+DV9mm/Dfd/wAO28P/h1KRt6ZTDhLZFv5fWvkff+FAhc0i/w+kHOKtAIW8JUAZAzS3kf+X2T51ycr/+AsG3fte9377e9ynbr309bf8+yte/eo9Ov8vqjZddM0bL73sxHr4zOvY+lYAsHzkDz1swQKIulrU5+t/58nX3fD/Io/6g/e526VCu6z2z5x97d8+89/P8j/+T37qN9+96HObRytGjTVU3e0r++qZX/piUP6f4mG/su6XYlCyANT3X/iDT7sNlP+xd/cf/+BdDzwtMg2NqahILq+m9p4vevknPCf8P+UTVzNRhEbD6k2p6eZf+dZb+Z/6l/3TA5uoSbikiSUTRETAUnnAv/3JZ96gbJT/H6DyiTBJLgiyBFGU/NUHgfI/8P7PV+EAggIiNgUQpBg+/I2P+x0kY/4fMHwISEGWXCQohUers+/2PP7n/WUvuWqm1WhEESAL5CKMklGmPqbf/qf3X5n/NxTvtVABJJIL1yhSDBpf/f3nQPmfaXd7bfVFlAa4gEggFygtFKkxpU/v/oaHXlf+fwHBe4siYjQiwFIEsdKe/+izyjH9D/V7vcR9QigpFaaCgAoI5HIVyZT9K1531W38P6F41EkXLiJIBCAiQsbitfcb/A/6K148qVyityJCILbKFYyoCBosh9986ar8/wDwvlEUuaJEgpCCyNWL1973WOrie914/fw/z9prdrHbaCmigMJmBGVREdZo5cINXnU38/+C9W47yEUERAARSUWQtV74yCO1T/+C+5/uw7f90+Pf9D/L2stPDSkoqkC5UIAIkgIUREWKq9q469+87/8T7L7taSGsLSm2u0aJeL34rm8H5RDv87uXM6JZWvLUT1j9zzH7uV9tj/OACQQGEDDRTCQBOeCGvt13/t+f/b++aeqmT6MVRESAi6hobJ4Y3/a+f7OzYrOqn3jFXYEWdiH6yBve/59BUZ5+zb7rl1/Zg50EEhLYQcBiFZNVdsL9s1/1Sx/Y/4tjHmcwDVyKCkIFsRnUEL90v6FsuC+uvdSFJTDlnHr2VTcQwhMwftexszC5E5LlzmSYtkekOSwDtS//j9/4Qfy/+C6WaNCoSAAiwgW4iKjc88OeurMOm0+6tDen2ZWUy0P/8BD+Z/j1bd/nnB3AHQS8XmrCnUVSTHAHrYHZb/j9/sRxGva9bTH4iy9sVCCRiyiQAlxELqf98hWETX25FcpU1EahNj/obV5S/p9g/AlHd3YYoIajhiaBGiTIJmSPYPV3fMRJEO/tc/Gtf/iuZRciIiBcRBQIqPiST//5ciU8RmqxBIGGpZEv+VIp/wPsyXfaHdlJEiV3UmI5hxDPMbE8Uo6Td33zv3ac/L/5xeC2x77vR33kAhjlElEgiiAgUPFjJ1Ga/ThCalQoSEIyPlhu/X+A/eTjOkYIhgaTtMNyMDtwHkCSsOPucP97vsEp/b+4FOP6p9z3Kx93+k1nbr/fpcKKsFAQKRMxTnzQX2RKeFeKokQEKaMpVymt0idfP+c850Y3QwNMyei1y43nkAcgMKQ2h+3X/yr/ysW+j001QdXlF9dq8Tvf9O4fdkWduPaqmggFkUBFXKNVfvl+Z7t08ZVDiqASt6Ls8u4110m1n3odH8MtyACCILQj1z/zO/71Wz72m/+kj+687PCcgxhP+V6/+bJ2vnct+pMfV1ONIkkNN2qMvUs9BblAQAQNN9BV7/4XlbTJC4Qi1CASE7k4bZH7p14fscjD2CFZhgP+0Xf78g9/+fo5v/vrfOrv+7id62VFQGrIrpef/Qm58t7d26cse9FIWdUKPHGpzMhCLjYF4CJQT3jXac1eG1hCEYBACWfdBtenXn8ZCCQyBCsh/+Y73Xwz3u9r+96//k3/909erhdYDZ0w5OM+8ssu8v6dk43qjdRctvZ3IG4jTB4ltkZEFQnr4RetpqzOnowIIsJCeCx8407xP7/+U8geADvnIQys9Z1f/iheeebd/eVjvuw7/y9fvcuqgmGUx3l81z946tqZ9ys8L5hcxZRWWSpy0RhVFrhgtCBCnwK77/enTKt/+UBDhCBlhSI3nD3Vmx+eev2GZQZkhyHZWZs/+fkfu++/nt1dXzs+/PN/6P/wdBUJGszhhB9wGSjD+3Ut9y8uKholBUAloigCV4RQGI0pQfMXfLSLb30mlQIXoiIY+h7U23jz1OuDKWiPSRMG5Pq7X3rLq+3d3Xm9v/nKt/6d//NbnxeAaMjhnOn8+sdxd5T3be/g0YgbwRSKIJmyEQQFpEGiRO9x8tY9nvOy+w8VERBIarrhhxetze31E6+7kNUAOwIkftG/7Lxyf72/v79X+FGfo7nDlBIH2Ec+ub9spu9ZLFRBfRExGhABFagwKgqKBERlKSc/+DcXWn3Cs5c92rDkiOnj61So8fDE68srHQZ2BiC9jq/N2657vb/eec/Z/eVf/cNvcB7IYAQCuzdf9Z88fTW8d7fdSF4Q1EjkQqAAbhDhYlNRkbH8gvf12H3+F/zU0pFCEZXm737mKYF46vXB+UTIGBIwL/c3lw8cbz/P+9eu1/vztvb+z/+8C0UCrjw8zk/9/Sf1fauumZopKCJULrEpCFEEFW1EALt5rwe/9KKx/JXP+tnduU9U+jSy/M7v2ynv090nW6mTnhd3AGElQRi+9HKzu9dieu2Y8S/FziRIA9BxHh/jPPmkmfesWQAuUy4oDrQKEESQFBAh3N7ll4R2fu3eP/GYk87QjnL7x/35crGOkObJ1jDQ1+eTWYGdKSDhxO7au+v5zPM8nt1c5+998cdPSg9ITo74xI5nZPN+7VvZVMRmUq6N1GgWbgg3CUABwofRR9ru9Y990Kc89uorWF33Cz9866Vzt2XSJ1u3A8OCMZzHDgnFq6++PPfX++W6Bzd1fOW//GbDDgKzrscyfNzlnmz7ftXedMdpIuGiLBUFRKJRUQPSiCxFCpp4W82LHp1evfTbOHHPO24/My2v9GgacTRPumCmIR5OE4YK1/vrmVt0PDuO653WkNvRoNT6Vbam5T17dKQBBY5ikTbESEspciFGi4SI6Mrb3u0NrdaLNTvTyBumOjX1O4o5CtF+upVq39xTg1GD5HZ36TLbzGVunl3W26ez5z8FgaaVnZ1c3308e9rynu35575WNEik6q0BaUE6e9KhyWU1RJSUK23FT37kfi0xWYml52lO7EiqyhOuwc18Ldk5BwmLWe5v7dmXcX/dOmE+4Nz9+r/+VNqJgVmZZt/+9q+4G963S0/85k/TXFfNteyaoGdv0dbTQovVboRAYjMqiilafdDV151cpaUNvKrqNWqshTL4H4Tr5QrDpLsTPdWMZ1aR4hD2Otf1sHVfCrZ0Yaa1ELrjLPRU19SeSFnjkLRDuhjPS5XJ7GrKTsw222qg7JDtPnBNWxrCDMHput7Niwv8sj+TsxzEEA7wpD7qJ37xF7563bsuXo+P7/3n7TfaqTlWcmcA5Mm3+GvHQ1jT676bmKtjKV0PWYV239F0HNNQpyXda6+bWfPqahquzT6mtoVR0kDpthk7F8Y17pZjai2WnEtVsfb+OEflE81UPOT9H7m6LrlIq2n3hHfaad2yPlFMV+z0a+5zn8tPTJWwKSKA0fYWf/qZeNb+2qMvmrw6uWp3CTV6LF1Dex6EbZmaDlRoZjFrYFJZsBkFt8rgLktEhmZZ3k0il3dnxfS1c0FE/sK/4pfCItkWFh1tzV5TlNmhRK9pKF1DQ9fDGoWFzRRPHcyYDspe1co3XzSN7tVx1bFM3CxOpQF0pIK1UsyuMJNF11i6zaxy0OZkzZECBoQ7g/Nxv2I2xs6jzd1bDgRmhwwZ5tKf27tVCLipDblmxGZqEc/lkBbNjIQdenQshdQd61ZgJ3vVWW0pExaeuo+p3E5zLmovWGZtlBHx/F//6VkROZ5LvvgbvdBwmX4CFKX2lymqI9wXJdGVqTcXKiAIEZGmU7zjc+jyZJWHUIBUagidRzOw3hhkC01ryEhZtCOy05RUOkndZTXUIhXsRYgqvRRFXs6D6rzwy599QYT+lb+mTkBKGzJ0WH3B4KRhYBZ+0ApIC8FXtaiMHKA7W5HaZlZrm+ZL0K6mpYKMJcBkSNORWrPRaieDvfNcM8eslsRmAMTuY6/VO/YwoVVgIuZmUdk5hjhvQHYgDUDYAc55kuE2bIvyckSoxSptyxqlNJVaBcZhQVuQMFiRjtiwgTTMCywZAGcx3GaAyh2WVXrUD9zzzTXCsS7mU994ohJZEeWkJe0ip5mlsEIhFkUW4kARuQLlQIqdUW5iwZBVKUh2ZDFWCLsCi1OHHq00s9ggTbGzqruQLrZOGqeC3MiGIKBQGQJazly5WswXPvrqF2YmE6qjdmcbasmQdq/spIxMKnsBluuSDE4cgYkjdMJI2XZkOHVs2TJih7SMTbbqaMUO2VZHppbJ2i7YVOZgZ0ZPF9m4Mx4JNESuAsF6LMB16miN2Y4CawIXMOrUtFDLEFtBGLtOw6y2pQ1lHwwdGnY6JTbYvWhbxnbS2Sub4DBk7A5bds4wa/BcG6nT2miDpN0F6ViUzl13+6TqzZCuyGQgBRERcFlKVJgQCUiUIhJEUWJqyMJK0pujmMh20jgIThjQ0gkZdGNomQCkrWtjm5FbuSWp0YIQo7lMOZWy+n3ejAW5sMHlGMIqyppVdDnapqsW63HGopc5mGXY0qO6V6vJHIym2GDYC8pqV0GOfaGrHGdglURGPCrV0NUIq6FhsQPIsUGxe5VFLpVBF7YpFgK8vzAss+LBzh54uMd5OZCQWc2oo2lnQlYtEncIprOmmjVhSY4uVjWzykGYdFY9uijOUXQR9poucuywimtLdblDj3pU5oDjXMxiea6ylOYmlx8KVHIMpfHQVIVGY7SkubBERUKUrEQqhuRGShCBXLhAloBGmFyUkvJEC8JVLIzKhAkwAdhrr86iVmqdphWsDlpkRxFggaJECAMpOS1JHv5SSzYXNC2QWsKOQtiLUJ0gwiyWkynQamW10LiqA0mbVgGcdDG6FxYZXUwKPZgUCQ1QqdQKWgnnYsJidMKiupfLcQiwdUJbS6M7nEc3PJzgPNqDYfE8ukShBiYHHISQW6GECYtJqwvSCQtgIsXsxSoUCQasxRawUgwplUV14tGmSOwBIE0bljsgB9ooJRdZ0Mrh6G7SiakDCLe4VIpIIRFCJCWjUgq4ECEURRQpkoVbJMRoUKm4CawgUJkAoxNgtFILArQodUSoo6O1NVJkyRVE5IqIcBGC+iNfsJwhFzbKfugEO1kgrQFoDcCERcEqBEWqYsTMCIQgt5YgKMFCCbcFuZUKeFMrIEx49BgeLwFokyqgAAtCQdk5D4Y9MFxMOlJYDXEl2MFAHp6HJHoAQg1AJtyqIhShoRDAIm8VESQAoyCA3FapMuF2IkCpAVm8XRAItcaU1stHItEpu08klNXsIhQgkqAIRCkhRUVsj6IoKSOlQEZpAE4rAAmoUHnXOoIAe0FltCDlBoZUoAgQRWxGwAaKIpyWR//zsuMLHPZ8sJ21ClQU26AwgVRaoWFugKHWCqCFClAaEKAKVqFSayptCGBRoEIAKhGZQEgFhTYYuc0EKpUiBdwDYmqEARoSa0CggQNIHIh0EJTnrzRQQawVSxCsAliogLxjGwyPC0UsVgi37QIBASqPFrGSSDLapVdkjrHmlQoooAhEAgKShAvkgoKKEBEiEYqQKIQCKVyAaFG2cMDbK28vAnUEK0Voc/No5dGII0ZsjUhqvNezTu3VTC5oFBchICAImCoQAKkBUxWoBAjlEaBya1N5VG4FsFpBanhcHrXyqCBAeFRAgIBMbgggCMiEtxoN50GugmCuPnAHYNjhoQOQPLeAILejIAJSK9hHEOSzmsqjFUCogrxd+bksVgQKqPhGGm2tHKlsLh2eMEiAZAQRAYGIKEAGAQgBIlJQIAJFgXJFgFDEMdfHHq+8tVKpUAQq7x5EBEQcgtA/4O9PzaMN50IGvHpRaYUKBRosVqBSixNaqUgFqhNuRagUsXhTmgkUBARKqZMbKE6QygQmsGN9hApQKZbRIkyAWiAVGEgZOkCUzMCJh80OJO5wHmSeI7b6OlC5LTatlUcLUtt6QwWoFSi2wUplMqFSASpUqAi0BCZFmBTEYgVBSVstT6zHzuhHchOLqZuCRAgqAoGiKKAoCMoFEWQDFCkULgChiIogIgiiQ0zqY0CtVCa0gdFipVJhdMJt3YjY6iIoItoSQT7+9y5eWTMXOO8vjk2FWiwpCrYCiIwQWieUSbVmMgGoWGUyCpQmhGqt1Fqh7DWEgUDalBpGgTipNxOoWKWGQShOqgi1gIkAOyEYJBJJDjSkoXBkuQOgJlAQAaqtSKECxRFQoUKBKjKpQCqCEFItImBFqKMVMpQgtaGhKIiSgvKyds4s9ptyFMZE4ckVKUAUccgoIBAHKkgRRBApqSICFyISkYKI2FqByo4wYcJta5VKsbaEG4rc1lGmIAKI2DpKQORyAf1Tf/viPTxygeNyZldow45llhSpYZxUmlJEaDVjdZJai+5V21ih4gCj7ghVim1omj3RUvdKwdFMTUsKFFP3amMtjBnRmlKpSKgIsMyUkGawY6RJA6VLsEM2KwQ7G7BSmUyYMChgpbU2pQBOGthrblIrYK0TGHWUClTrBEjLbSaFiQCZUIDdUEWuXsrU2xyOHnZJGkQKUhThwqUoKKE0GqHAJYOIABGpACVVRGyGKKBImECxSN2x0olUSm0Xk4Iw2WsbRm+oMApTEGEhXLggbgBB8yf/3ul11owLHIc7FkhdAyxuq4CmAjttsLMGqLYGwFG6JtXiSKEZSGtahFragEMJdBYscIcAmWDA8yhOOBezKiN7DU4KlZEKnQWEqwOoKXQeMeAOCXMeDDs2NDE7TUkDU6nUMQQyWovAkDCm0FRGhK69KtW92oCVADkXoVqr1L2YVCdgyYRmwgRTsuMswIXQ2Z2pqqZKjgRCBQpJESKoqEaTBVCkBRkoqNEgQaPFLchUKsINYcoKuDRaJCY7CJPzKAIT6kgzYp1URhpqZS9GGVIBXEQRoxQBLiJIClc+/o8vWnt0nAsaBZCmE8VSwakTZGTigrAXITAEB6odpRAKDQHZiwC2hkJFDHsRCJXVwuiich4TKhMOEPCAxSzDXgCxspdrRqoBadAddng0hxqAIdnhoBRhhwYBN2USkBGsVCq0YdJwa7FSCBOQBRRYmDbIhJFQaBVgLxZNoVJsKpTCAtqwILgoUvHIpFIIRxYzFUBQRAQXggZFQAEQFRREI1REoQZQpABaIiiKCKABjEhlwqMTaKUKlb2gyIRJpVYqxQp1Y+toQNBogAsXEDE/7qmXrL1258Lm+YBtKHECpSAVKeIEalGQkZRJQGp2kNuaCjCLCaNCZQIT9kKKzLIoZcJIWaQCTCZWqUAXQCYBsK5CGKBym2LCeRAJMawSD5PQwiZNi3NEVh7NkMptaLBKoCINQICmUppWRilKMUXRFAUcA9BxlkCbWgZWBUZhMigghJYJlFbKkcJtjIURIgKUgghChCLXqJAxERGBQoJbABERAUhBAHIJQOa2NDxaa62FpjKWtFZrnRShIDAEggiIzVFKkJKCKBotn/DUS/Yzjwsd18OZVCmWvSqTCViLhIkAGRmbUimCVaRCoNyaYlEoQBDpKhmFgC0pxzTFMlIgnRSxVhAmGECoAJYgQEUeyiCIhIGeRwNMPMDI8wABOqIckEqlTlqwTBArCDYVoUgFxwYQECxSobQKTGi4nUAm1EAnDVCg2hIIREAUtamtp044cnEVE5UIQAEiBJCCgIwqTGwqKFC4gYCA2OpCckEKWSIKNK1SmlEqVOqEyUi1ArUIVKhQYBYBRUBQBLgSCUwF8Yl/dMlqHvPIhY2jZLVhUncOaEOgAhUIIJCiAQSpVQLIo0IFrJVKA9aJBZlAlYmArSJAuC2k8qjcNlQoVgGcAMVRbs5DskkBQRYOHk0aqCEOVsAhEQEqSGkmoUoRaoUKWBlShEoo0BuoUqWIAoznIR1TAaQADQglgGCdpAARxEt6qao4eliCIgUBqSCIiIpEilQBiESkoAAVCYgCROCijCIEKSJBEaAIxTbcVCjCiEBBbisjIBQrriCICAri8C5/3NMum1esuy9s3H8STudop9LJ2FoRoVChjyCzAGGClTb0bYBUEM9DCIBAHUCotUIFBJhYYFRAbp08Io82SAUJgJBOAA5gx0hI2FFgHgACDjvJACQPh/bA2kZuq5VVRCwCgjUAQ+S2MusRAQEBuU2PNiwQQKCAUKQBKlWyl7gQFOs3n90TGVGO0gZ/mtYLyVRUKC4EAlQgxFYBigQCIoAIRQCjzVMKsVUgIn5uJ9xOKlAro0Ct0IlALXJwxHG/6wtPjTHm/VzQ8OE/+ldxocIM2cd2N+MBVAAB5LZmbKtDy+u7lSr1La0AE2Sztsyy3l/24lEBBCpSOdfEE1it95fsozKBCVSopILn0ZGurVvwzKxmB+EC7DDIYnuJ8zzWg+EUoCYDE/CAYXd17Mp0AUWo1CKPCt0Na+xwax8O5gBqBdoUK7X7OI8pIl7nbhZAhRoAQQB3EFEvan7NTvUaFSmHK4e/fsX95ZG49TFZSRvIFcrUQJ4s9SLlxmqhAnBFgFyIKGoRQC9QhAWy2ECKpT0mM9LlZPQ8ZEgFGERgL+o2J0BtvZlyyYYpEKCgtyjldva+e7uw8p5zIaP81f+zf9UXuKzhuFxXXT3qN45XM4dFqLQVaIvMARxAXkGhCnsVaSaVwBrKqgeTvpjh0TrZQhsAAa1HrYWXcCBNMVTbruKszuJ6WQCrHlCU4ZxMQGom3GHIC3JhLwNw7Bo7yFADJwcQbiPn4YLCLB6dMKnAxhy1pDlqHY5hX6BaoJoinXQ8OK4XHr+MAzAZsVgL1pu0jaQKXnbbxamBQzi8oa0f97enXIylcGEKXA0VzUwALVoQMcGiAFyFLLnKShHTKCqaIqiouSwhb2/DO1be3pIKFJxMzmNAx5KKCIIYhYiIwBTb//ORF6uyXu/jCxnAf+yDbzu2r17+k9ddXnzw4gs/ml/343/Zr8xRBGoVOQ+8/3TtHpx3sMZPHta3H8HirFqaBuFnvnB2vV5vHi7roP7Yx1/65AtHLVQWFGrlm596PXr/JXy48/rya3x07P1FKKNYcVVYfOOnvnJ2n6u926sna+Kld6JlspMDLnC9njdXr+srzvX2ePXpV34VAw7geolJlK88fvjhgzcvjiUP686P98dAdnBSqggTrsdPnuys9fJre+0VXszxhROQOhOhDgTyyd7l8sm1uXs5fHL92K8ANi2INK380LfbN8ytnfW0uPlvf/zEwquUCUdUGPzrXT/rne5ay6XWi3ulmcbCotc0t8xTpkYVcY1q+4ul5l3K1JCooBSocUu1rKoyrVFbzrEGmSb0uOqabDzWNx96rB597eUu+4vs1QZgMlqqf+yLX3jgSk2Py0rxaCz3kKVoPwleLG/o85nlov/aX7/+5KD3vhryhY315kcuPQ7/r0kCudif/Wf+VfsyC0AEYbH/4L/mZ7+yuftGH9RL7376S3/XL94LRKgVAX7sn5JeXl2OH5qex453X3v4+L/xb5hLsRbBUOF/9z9+cbmG2b0e4XJ3z/Xh237HlxpWEZARhvy2f/9XrnXPw3H1cq2TeVgf97u+yXmDEQMk69Xrt//i+302HXOH95ee9fF/8mvuha3pAgHI+/+9fXjgumbxkJ7K3/brkUhLoAI4/5I/+NHP4l6u45szQ1e++Mn1+r/9Fy+6CNAGqewf+Sd944s5Ll/bnC9Sj/PT7/1bPj5UdGRCAK7/qj/4wd1xHDe2kTHSc9FwashHCkDd/sMXOZXQpOEKqi6XW83EaJGduQZ98t3jm3dfepGoNFxJpdh8yfudvaxpcZvtxbqmMaZeTs3p9roOz12HddlTunhxv5j80W8fqxP2okJh/4af+PD+ytntDjud1Nxam4mIFhluvWrtKDWnXTwlY933ZC5s+mq/mAM/PrrSyfbFd3zHWnI7yqOd9YWf/coX15UvHzP15PJLvvJt52Ui2KZSYa/LN76adbx58cHlOu3DpV95M9+3HipYSqVQiV8d7pNca881H3D3U1+kaRUKkE5m7n76xYu+uq6Hoefl+vIN4/311XdxZEIknIcqX/i2l64rd8d53NPcnb0LFhFAgNUnvO3Duc7sPXdc7/lkXkG1uibVgrPNl7/wcLmfNQ/9juMs+8wnr/bry+V6BKAaEGG9mI++kLW+e23q1f3hT2icVZG32l6+/pWvnHw6VzVrVavFLdOqMo+OcgRAMW1qYaqGWyTSRJ8kT6lYslaLUQltf3niYiHLRYQqCQFO3n7PE6vRrtIIdqOll/a17PVuT2Zt9jolI1zXq6P5af7BLznK6F61YP3GXzTf8c1Vr6H3F9eIRl+mDRXzzrpcstUFGfNy1UzPel5HubDRB18fD3B+moBexvPLTNcQArRW3S+//IW7+1mHd3uukOu+y0gRQRBYPPjyo589c76cF53zVR/mZe9m9gHFlDZU9nf46pu7Dy8G92bfufPy26dFqVLEAHzUcLG5nF0e16z7a9bxEjsJkelBGO84bl87uuPs3k5sbleUIBNhj+Mtb50s7rv2N0NeP+TjtguBIsjO9uvXVy9/+oFcS69r7xl6yRd/7IvnODdWilD2cXlxXDvHiLP2ujcv6AEVJTfIGy8ffPO86DlknVVzVUWEY4yCPRfrWrhmKqi3oayosw0NCogcNO8m8z1OJhQIQFIk1iXmVEzzoLxIFqNP40yuGVqa4exLju7p4WG/efc7vu2ocjuB4nl88o//5sdyX08yZ0Fkq1ogKlkYZ51iSMOIeXje54Ln6LB4qEbdw6dv2BIeN4DMSe9W7jdu0/m0mGmFGipMgPO8/5S6Xxtnt6/z6fEzEyfMAmxAWNfzYeZ4eI3d4D2X6+vx4SVUqXLbsR/My/OTB3ba2SlecV69kWEHAQEG5t65v1t2z1P3PHp2tzw8ZycEVrhez557zofuwcGXLxsqTBZAiR1f8c37bSbnnrG4eX15/fpVqBVwIhWY65mDuZ+02eUhD+eiBUEmUKH54GfuV2drGgPL09rrYYdjDVJHWTCYKTfVWhWrRXtppqFeksq1WuzP7gsCEURRBFS1mtbL3iEaZmVENK5rOmE/hMePVksvF8+Xf89XFaBIBbq5/pPvP5j7r3donBhNrsFmay4ROV4HsxzuieLZsy90nJA369iXOKpf5jqbl0U6i3c8Ll98881zN1znQfaHn14YA2BFIFDn4nXGc+bs2Z5fzH7wmACrCHJbPz1mHTt3XGfKOTnPT998nDYC1rpXzfinvv7pp8dleDjbuWT2hOt5csDQAyABn33ph73l5np/bof3557rVz4V5WBwB5iK+ZdfnevDee4Hc+W8v54AFWutUo5++PWvv3h13XMy53LmZM5PXr0+73I0PBqYVLk7P1xd7A+796bzIv3eXKBCJTwuH68Xry9sO6M7ntC8zujHA/Ge1PZbmRooDIhkVZvGDK3v1oxV80LRdauLxwIECKEoTBmtTe7LDA2r15TMuH8q591cH4au4zz2rLbY9eI6d/wDv/hmQpuCAG/+ojdHvs5x7ut1qqkYs/spswjhIhnCemniWIr7SPrgAuikV3uG0n3l5fwUE1oXQK0VfLDS7n3uXMN5Pe8Eilaggp192fAw59m9c+UeXj8wAMXKKMUZbPf22mt7uqfdmRSoIBJoSnrJOefscc85HTpcn/Co4QPDOGH3PDvhPE+up7vs8DCERDr2wp65ztmzszuvG4CJ1ArI5Lrn4cWeczjn3J2Zjffdw75jryKUIHW/mHO7+zDs3c3u+UkAJzujN3XS3Sln5zFmJ87ozrCOC2L1NqihUSIaXkiZ02p58SW7e2eu2vNYr86sr7gj67WEXBG4IuQCjXTGNMbIYNQexIx83TWTa+/X2ownedFX+zw+eH355lz+4V+8JkWgQq3z8CvfvOj1CgP3V2W0KsZSZCGzrOZgGCRyFsPDmatfCKFncuZqpJvr9YSJtDJxAmUeHtjDnvO+m6s7D4sC7jRFWhzPechc53yYh3a6J9friUDb1ZoOKfNwXne7u++Z01NPWyYO8ugANrT3eu6Hc+/uC3sY9yHsAAbUWHFde7bXc+8v+yxPut5A+aBjBSpfu+x5uJ7nfkjv15yz4VSLjpXKwHV7zrn3OefqQznbce/XtS6aitTaTc61u8+9Hzx39yseXs8QMqKVAs70enb2g9w9x5nHesTO8UFwVyOMuFXLCrjXIz7g/e+yM7U2Jjzi269//d7Nf/93YkmZIKHRUkHrjJHskzVOxsQgg7ymOw+svAFe/Pq/9Nf/hi+84PLp5u5rf/R/8M2/56uLVpmwg1A/+aX99b/0V736nu/+0Mv1p17DC//21TfCjlbLVsmuglwzGVLihXvSlVwYwel2k7B7FuYo2dKMTqw1e/PQ8+w9u/u8fzM8akUggNI+DOd1X/e1uO9Xr68ZJrImVOcAvHp2rnPueRi6j86cl1iypQKkAmzamYfzfrf7YbiS9gJ7JCnBgNnNBOeV6/3N9bzu2DWYMiYmwbXr7Id52Nc958zVOS3HVoCMVcice+ze93t6XScPuHte3AGYBCbUsemmu9c91/O8ticPOc+xFKEVrJiS6z2dGmM9iEen2zXOBYaaW2+OBnDqgz//UZcCjOpMrmiRqx7KaJ0JorJApEYhhqETzcl+xJAyB/l1vF7YJ+vX/zv/ed/+8gBqhcp+8+IoTirAXlDh7/7V3/mSd864/jd+6LVMFUmTpg4yZhRJ09yVhAuk25M1dvom3WWYxWqDEoCZc7zfc/+wbzlpsO4VKFYede3uc1/f9DzrtNdOKQWQ24MJzdmr3ed5v2emR+lgu1hMmJBt3UL31Yfr/Tl7Nx1qeemVBAhASKQbdq+79/fXV7cr902bAA2ch+xwSDNzzvU8p7vtyTGLuBeVUG6zr33Iw7nPfQ7saYeZetHuRVUqoYXCnvbh3NeZm71XSzMiVKTIPqfzsDU6Hsk8HDM412YWIdS9v+cxJ0lvbJZLRPEsJk9EFilcBDVw1Y66QtaZu2NkBriulm4++uf913817B3lXDQTwoQdqbx1AtRKOzqzSqMEvPBzn0mjpAKEo6HgUuiEC6fFEzLOeXIOzaQ6gQp2hevZntPZ19lczxMKq4Ig1JZy7d67092TYe+uAyc1FKSS1hfK9WEe9p59tnvmskNTIZBKLGugnmvO2fucpkyp15cUckh5mNC6nef9uWfL6ToRE4AHNHQ417Nz9px9TsqGxYQCUuR2mL1neu7rPmt3d4E92ZChgjzqpJ4X5+w5D3Oe0ytnJrsQCpXH54A593Uxj9E9iEcL52U67H7md19Keto0L6ClrEBLczEayAlARQKIXBU39QzPjBFgaBDIWb763/7X33ElpNAuxlDrKJRZ1EqbytDFo90LUg6FXvnp/wiLTC5HRsSyGna4oFpg6AMP14vLWiHYVmpzce8953nuvbler1uwuhePi0Lnck/nOied9hxPh9QKlNs2c5bzZM/euyezzzW+qVgpIlhbujNxz8x5DjJ15HpdtiHloexBdN3avZ7bwlmXk2GnWSEFIpnd6e6Za8w16ihABSrN9P6c68Pevc65OrXdx8AaJghUmYDQNef0Og9zdmYPNlQBK48XH2b1OnvuY8QZM5lrnBdQ3/+5p7A1McaEC7lZmZDLkguKRCRiM6JoqjEqnrszHCUkYW2+9L/7Vxi2uiE0ZVAr1CKP1sqjtT4CGASuOfIL3v82pTUSpxkUUwFyQeWtzsMZz4sCUg0gM2vWyfR67s1096wAXRWo3Hazr73uk3OmpSfrYYdSxyAVtK0nsDnPc+ie7m5ABKEIIts1l9mcD9dpJ5vaFMCjAhLoiGPvbc+N6y47p5wHNYDCABgz990dZu/FyQFHSFlUAKHuRXJmX8/rnHvvKRSciU2lIjSA9XKc65zOPs92z875CV3lVmqthjP7oXRneDaJO0Ocj8oj/vrSMaqq+tSSFIQICFK5NEogQBUsISBTWKzSR2dkYAygHOe/+v/2cncxNjSFMWPDbatQaQNQaZVKfaRw4eoJU//f32MqilxGBLRxQfZ+EbkgjwoU4fjo2//QHE3vHq4hF/liZ9kJVBAoeLwJI3j3kA48HPeUZSUAtYJFXr6+z+pKwNzfnffcHSePV6BSD75+l9PXJJ2W1WmGJ9e5v8GEBBDt5eP+uFHnsOueM72UK4FAArnzji/k5f1al/3iTHvtB4UKiBREVu9W4VhZZxxwtDPtGzsBgSpAXfuDb+ay3/SIHTR8EPAREIHKi3NNs9YsV4w6g5yTMrSh5Ce/wLQoxeSiwKVKCgSoQQMioiCkIBJpp3YWXVVzehK2tvD7vx/oLG6LPFqpFWqtFLmd8K6VIg7/jK9+llrShoUFuXDD3n1NWpoijBUYrtO19doZh557WrTK7QQr0wydstOdrTtzpTCBWicVpDBspC0NDwbJW6QIdV0ZZz109xybaZR07hhAdgoh2UvOWXG21z1lvQdF2GEHQKTb1aEMTnbXa4tQYbQC1M0w1z17dtkOlO7lSWm9kUqVdsMqXPe0ow+9joVJmwq0THYx2eDMabMw59iAyUXPfuiQURRRCChAhSsiIopAhAhQBEix6QwzMkzYbPT7/KEXVCZMeHxyg3zmUaBgrZViJQoiGi1EOvuu/zF5wiOKCOULNbYt1yxEKqktHfZmz95c29m0WECKN6m1rafS3d7rprANYOqEWRlq23MYep29dycnFKb7wqRY7Qg18NB2+lDPnjp0Z86DdoDJJhSGc2au07m7XeG08RTYmYLOCxT3wEz3nLubnCJjBdICFcvIpnu3ZzpQkGmHmlKZSIVdOszs6bUz3acdCqS1VEbM0D0O2+lKN+JcKyhc/uqLLRQZxRICSxGUK1KkCIhkRbgIKaRgXB5xZFBg8IDnfLTjXp3FKFRaQmUHaKgVqNM4odBUJqNTWJEiTGG3fu83Fqouo3ABt4ELTWUUqHuBU4VzHK9MOedaHnV41AJZzT5n6uQEqBRK7V6wgErjuTYlrWcZQPhAafaigraCHJ7X9nRqYt1AtB0gaRCyaRq4v9Z53LknuDMnO0PshS480jju6R5nndO1WidWgFIhm3NIZ8+U8ngDb0iKEwiVgt1tZ1LWw3Qo2TelCtBMIJxrRlXaPpTL5yoAV7/8YhcBiJAEUKMJcAVcRKOEQo0GhXARgTI0RDIA0tR5wHNP1Y5Ka7VSGmopgYI71sqOE6AWqRWIGKWkoiHKc1tfubccwhqIXLhhisTBUJkQijJM4dzDpoynad2L1BukIufqtGd5tLgGmYRKIUCgV6adc886KUA5WBPChIoUHkLOh9nT5spkZ2P2SocaYEgQ2eM4N9h8lsB64oG50wED4NDu1gFqa3upLiZMMBUoXPd2Zvd8F60vgyOhUpEqM7Zz7g0FZO6WpWpngYRa5sx0HBlKrMF5mCteemKMpZFGgaSgUSoAF4giohlIUUQEFVHEcAY1jAIwxInnL51KmyIVmACVCRQpTkZHCzKhgJXiI7ii0djaM669P0tDYi7krsJrmJSGapEyiN1779m04ADpLMChCyg7xDK73T4ieGGqVerIpIDW7s7ucGsPTvZqLSPQ7HqByfXqVrIZaxk8Z0mywISMri7E7rIGsIrbJOwkCec1U5jdnYlOhUJNBaGStnsN2ZZ3HNYwzUB5e5uDMYWeZWCc7MuA1NQJ0HFi6dShegjnXKGee5FrgjISoAhKiCgKBETkSqqAAGLTil2QwKix0QavPzEvpgtG21BwRKgVoMKEHUBGrBV6QwWIiCACS7aXT/6cojE0Luhswk8htUFwpJk0QNsZENhDUnlUoUKDzkUz42mB7bEDWAAJk1o9jdNW+kg52WBDCVNlQeGTfV/3SI7uAWx6zSmgAeRhzFSMYgREKwx1vWDyqMe4od3ugmscbAAqFYGhmemsK+84rKSAPOooTjYs9+oUClzDVdKACCkiZBPrxLE4HwNPusc89QlRFlHYFBBFEBFQKFQkpSgKAkSqRVYSDQUY/J/L5zbihArFCbcVoD5WpAK1ILcVJowSubZsj4nbI5/fFs4guYADJ99LU+TRUElXJpcIdA8wc04XpVYBqaHaih5vLIDSQYujQIPQvuh16nZP9g2c/Kc5kEkgQGUyfNvD/fVhb6a0oGiPE0CShwkMyzFFJo/v8qheQACBWWpBEJwcQLmVKkBT13l0s3lnOy+RCgIEIHVdJzVp5yb0QVOgAgg3dW9LVXZyPizm9/mq1UKLjDZaEaEAQZEQIBBBbC8QigBENPUsmlM2ARbz3b/CFYhQJzw+IsCk8tYKlQpUKo9WICICcEVAGH35xnuwq+5YuWAjyPTjFOKQsiC69tWtqy/iAXBcXryYiYgVKgLuXGai+7hyu939JjKLRyuAeFTUe3pXFWJixX3fmc1yBQQKRfHADz1xH/51fzCa+s7eHZmd/tXvNpViBQQZ92W4O5Ksa1qwx13Oo9JAm0opv+gPeNFwgJDOQhTaQkSUFDkZ1Asc5RDxuIGKABcHBk5c8xqrmHZXczeMlsV6tzQaspQtwOJUY0S1qGlZ0nkw6k9TYqjREiBGASEAF4yyOFiAKbE9tbjIc/roUuHyrOeXG1sjFAERBQRRkesAMAIBLlCUFMgFCEQiITZVE1f6+fNEFRG6ACOA6cqv/t7uFhJFQFAUxc+7Y6etcY8BrV43Qp0IOOG2O8pgedS2BVZRwCJUJmkIBhnWYL7oHh/1/QtXUBSxVUD0gU/ZHZn6LK3HuHQNbW5nXwKtMmEURrEQp9ZMgdrdRcBig6BDj2kPqUWFkOayRJDLFQkVMEyCAuHwBdJoFEAQpJwRt3XrIZRRiCoqSwXCFUVYRZUj063o3OUjTrhMC1sVlRAijmpenp1UlGkj0wauSIClQMVoxBVIKm184OWpiKAQiU1FACLCMkBcKBVZG0UEUAAFWCmBFEUQRViP/1p21VyKKhdYJgNXf/L3vfv7nu2OxiSNQpYiCL294PZm+lCvABGhigUklcHTjXUy8xh00Yl1tDqjFI9uJyImi7f9zkedPAUVESJGwxJJiVqGuO3IF5nphPApLpfuSM2I7BUozMoZSlsezbHTOtIKe0Hhk9IMziZJo+Jyi6goIsLVHCVzCIfPAqC5IkABZLNWoZEMaQDGcqEKgVRFkYl6FJJ0ZHLu9Ou0EhYOioWTkPXSXmrJyTFWiyxEc7RRrkAkSCUZHhCqHFyjfgurkiJyAYQIF0TJaBnNRdSCC1F2ilRciilGbTSkIKEoAgngite/atElozIXWAd8wO+8L+ud3kYYC0CSW6RIQF5022JY1dUBOpsC1UoBJ5lmsgt2ZW4qDKR0pAUUYNWgIVPwvj/1QHXWzVNALmihHCgiatQCqreBRstAYdVlBVLLYkTIqtPScGuv5LyjgezA4pFamSJUigIOaDQXICupqFIjVGqEoyoIyoULFEHVWoJU0fYt2QuCAtJoUIALqHliSFAMcz6+x8luFEFFUcUaUrTDgv66H3zG+77b3jNe9oabLvuwD3nU7ivfZqjQaCAQCnIHI9GnNVgfdclIcEVYRAQ3UkSxJEujgXC52KzRcEFFpuTCFUVBikARCAKuT//u5TS3KGVdYOE9/+y+9HkxSlIqCOSGiIg0+I87pkD3YHMvTQNVkIpUyxnAmfLI0gCTDBgqICxmUiUIO38l5lbRULmogFKJokiBEmNyqyocqUDNpOCkAjTAHHN9MUOH6U3p8qi0uuisijA1gwCjTI2AKIogyhKgJLEcjt6URCpcuBCIMMqRugdlCzOVGwZZKFIRgQY4Zad0XvwgDSONpAhxaQLI/q3f+9w35KLffcOa5SXc9mt15XTjy0suNSJcbFAVoNI1UGjfQSZSCkIRrhSjgSigRgPmBREVRFJFCorQ1n3sLNmMiCKMZBQkMA/5g9VOBSuVCyw//LnDo+RWkaiSIlxYCFzpL9g/MdYapgLMvLBMsI5VbhijtbWPBY472ky0pal0llsWYmSM0jLzYoEZFTciBQpViKKQCo0CkQaNhNbaiWUCUHT2UJYWsdzuC1LTUQUGUEWlGEVYwy0RlhKlAiRlWy1RjiCFUIEIis2gblUNSlU9BGjzqdEUFU6xNVgSOAJyPuy8bTeRIwkrFCYZfuO3/8OD3vF9Hn35n/3qH82XL8ai7d20d7aNVi6CSwRpNJJJspy0Ub7mwRaj9SYQiQRKjRL7L/qTF71p9aC3f497n96JCGioogiiiBuue8Y/3frCla7+xM+8zEpViCRLpCkoU1/e8qKTI4LkAstlT5iL1otUJkuuRAiUiFSZ586TGYOYjdpJSlsNULEZONh2hkfHwxekEREyUkKNam20BMtzdSmSXGFTECkCnJYoAZCLIFLIrTiGCpZmlhJLa7ndUzFAAJEwBBK1KAGRjFCAEAiEkkpRPY5yOKIQDhkFUQiNljGcGkAYzQARakCEKA2ppTFGRefDQ2tKqxRKKCEFoL3iRR/xko+9v4v3/+mnfNWZEwuvxqRLeo1WkSLkgohqUKkiRPBpElGaIAIQGCR/3y/f5mXzn6x7u+qLv3BRilKigABu80e9cTpZJ+bx5u+pT37iDrilIgQURFLEB/zkTpPjdqHlrg9porXWK21MIDdFAiTSROPfVjsewSMg9r7fBiSVW8FQhxWwj4WxhCoVqoCMlq7ShGqcmU+kKERZCghQEIDKFCpwqSUkKuKQOjWAVUSPB5LdpApgmZF3FGCkDYrVQkhzlSukSCQQW5fNvUzC4ZWmEgIBQUEJ3Wq4Kk4Ai7XVUAQiQkDZU7dEqJDz4Suql0tIUkAxFel3T7387T5vZ9ZqJ4vP/JGvXvTJ8XrCfTcCERWRKpoGi1YpZIVPlCsSLgQFJBK85Bueu7xkVx7XXjrqxu99KQgQuIjYzE3Lu2jVd5rnMz/3fk/clRDKhiBNlqj7/zbLISNzYfUN93tgqpdUckMSAjYgEcmCfzkRu6kGSPvy4DwcVqRyW+Gu+eTcwNq9YbWbWVChSKt0sajQoJXbqVUbAqIIRDYQSQVRAVypCEUFIEIhUwHk1h4vXn8cqGtjobk09cZaK5R9HG8yMq/7CEq6at5RBAhBBK5qU5untQlHTF4hE8FogFAQnD79mn0mJcRAhRNLNxBJsdWFdrUzzdOqqkl17kofqIkyWwUBQu545Yu+dAFkJ0v4pNOvPtnbso8TWSxGAwiACMBuRh9jTKOsk29TuAwVEUSksvnhH9PdT6372eLiyfPyxL5coCCBImi0S8ZiXqTWo528y4u+44enxqaIiBCKlIv+dn92A0rKhRR++n+bklpSioDoAJeQhl7x5h0G+4yFi+z7K0EpFqCNwHS5HEO5zQ5QqIBQhYkaE4jWI+MKihDhQBdSRAQR5WJTQQBDpIbJBKBN6bxsg+BwW+XxCswCCTtNVCtUA6h5XRGJqCgABWEUEkcWAyGIIhSkqGKLGqRUAILhRQQRICAiDFEikRXXOUu7hAhBTAioZPklL3iWCAqy2PmNd0/G/qK7UYowBURREqksaF3UOLUbRCUFICJI8c/fe/q+Y312PXfPbdB3miqAcEWIKKGPVFZjzHOfT/7nV/9yc1NEBEIWAull1y9wUZa5oPrK97pmbkISCCK5ia1CCYtXvW5peuIQWvON62VNtCDVMBm7dyhMeWsnMJEWqrWkGekRgxruMwgrCGEqcpUlBYQLXLKACEXg0EYmGWsbQ5tVt05n7CMUKmCFNYGGB1WOnTEFFPfIKoIUha1pO5lWHjqKOQEBl2RJDmBnHpX0ZNZGGWk0CBZBkEKDGhBBGflcVS7dKUESFUCUIG7/4mfe3EiNRiriEdfcfGqVsZAQiiRckADE8ylHpZniwQgHgQJBQfL4qsuvWI29sefB3EKXXCgioAACZR13VlnPc3cufc3ffTFCAIoiRRAv7njhjpTRFC6w/vgTTBURLoLcgrakkRp6xXUnWNecGhS+PPzjLjkz6qQINKhmMEZ7MxppRitI2YudWktSSq7M+1a5ueLCFXABAREQpCK54QISVK1CAzYQoKZNMwVSbh1AoHVbAlgGR6HUSDlCRC65cAXAAu2KiZqP0sYJEJQLBOUCal/LYJWm9UZvGvNuIIVAgMCNuTmgRgvn3FxdQamIREMgUeN7X93e1CdFoIhM3/7Zl5k4GY0IkUqKgrITI9GmEe4HqSQFCgIF6jmvvMt6Ws3ruZPsT2OMphCIChRBLJfHqkbv83qQvRO/89mXRICQIlAUivk/y6koXGh90UecdmWjgpSSBEQSiFHPOYuTtRxlOebf+7/KHAWxPCoVOiK05VEttzYFahezSFlJhWlu3dMcIVSkAEQosykAAapIQRSAGSlMDJUJMsk4LZvN4zWttcGU28pqxQhimQohakSkABeSzLJ7jXOUwd2RXBRUBAUoe7PsyAmb1bIzhVAU20UADRKYidC5Kq4iAEUiGgIY+884ffYWAooYU0b7pJN3KLtzaCBIBBUFIWYGELvgLhgoXCA2BfyRa2+5v+5jP5kzu6+t0SJEUISFynKPe1/37k4t9/7p07cAEhFCIrvPLleqt1xoGU//wvRCIBRJEG0EEZHnnV0w41Cj1vp//GvPM4AFHysUakqhmUfGddA6QhtrZUuJpqFEcDtUoihSECgCARERQYoUEFtdECqEijVAme1OUHm8loYCiHMjHMQSiqhoNM5GYBGBUi5BcfXeZTLKEdi5hkpFJGK7lXb7TjWTUAbSWPQlSkAhAlASWkNOWZWcK5ioERFLQIYBreezZWQUV0OTT37pkxblqXlesCkARY6SVGtRqtrgFJUoiO0ugNcs50af+3oexKs22xECSyCQZRjQK72P0bta6t8/zRKRZISAuHTJ3+9mjKpB5YKK+OPP21WLhEgDC7EpRcjTv+611ExPY5x6xq8/XbtWahsYBVA8zJThdnQACpVAKxocVVpal3gpIUIRIABBNgSETREhIgiCCVSogDBqlzA6tI8xndRglRGokCNKNSsKQzBLiSIRoRQWY/HYz18PU+NwyqWXhyCiAghClL0cqRGz1WtGqaII5IoiEC0s2ignI41zLHErSEAhAlU46qt5PoFoqAhgvvrkOvTe7CISoKBAgIJpxiTsQgoERAERxDzw/tzXc+82NdqaMY0SUkSQiqJ39xrzPPfYFmP13k2RAAkRoUpO/H0WWQxbKBdS0PjR74ZIQgLEZiQiEPWv+7uz3RPlqv/06YxYsAYIj24HMUseFWYI2gI1FDPZTSwVYcE/MCEFQUQkiMSmSwCuSBEuiICGCrRUJoSSacMmi+zHCnsBDTUAFhJJHlUoQPEGkKvYLkAx781eOYOjvtOpoChiM8Kl6OwqC1GKvYG0RmyNCkVAMWp0laqCx7lKOAOYCECRQmlccqmnsTohF5EYiCs/88krZ0cpEFuFVVBtZCrQwBMLAhFEQgBiXpRGd+99Hl7tuE+ZPVmAkEuAopBSxphH9/Boo/rDLgNxsIAofnaTI2OHC6qVf32ve0SSXSKCCAQoEpy5/fY77Gk1Lef7/C1coEgAqTxe7o4Prveia+2bFD+5wVSoIBiX05OZzPwrjkARAUEEEEVFlDGVXeCqLjEajKSCQQgVK4GkVOwj2QJU6l43vbkfi6kJlIh0XjBQRQRBFBVth6su2VuZwyviW1S4EBGRIkBMa8+0mqXyhjMN3CIgigRETBrrOSpXNXSOCNd1owAEFEQ86QvXN77wZBnMaGlh+Mvf7fZWu2OyFBBEFASykCHT7vDMm1GhIEUQAUx81Cdesn+JvWbWYu2pa7noTSRFKhswpp30ue+PxYrecZ+0fOSOXBBtIYrSbnvOxW0pqoqBLqREPPGHJKwiEkQACVJQu+4ZC2VtpvXD/yJdAVuBWoEJVQaNVXm8dJhjIIWmABWoaLCccT2fhgsEZSkSIhIiKJnmRTVkuSasBNYZmEDlVipMRGO0PNZaEWi41U1dRKhSIdD6bxPJkYgiBKPB6Xf/S8mHy2J+xLtYErgCikQxSi/ui32ZpmJToUMhCwURUgm9VNXWkhJyjqqm/QJQQAYFUllcuTpx22vuSyGNaLRetXvVa3dufdtmBAogAkpl7kUSoky8FgIIIIKIhEfQYewHZ4k8rde33CuWBApb1bLqQ13a72MwWAwv3/60p0SKgEjIjVufe2quKJ1EuZACvP6mx8yLKGlGiCiSIrdEL792Xaud/eXZ9/xt5kWgYgttoQZk2p6IDfPI0C5o2CMF95pQDcYIDI2JM2+4e8qKCAViU2HTbaxa9neD010EidFmdQxVmPDozmmdtjPlcQ+YxaS1MKETToZMRbEIg+edPZUBSRFFuJqV+u37vaEph6q5/iFhFAhBRORq5o92WqEkHQUiWjEmJEazJJQKaASwZYpzHPq8KiIgVEBJgIdodfc//lI5SguVwhd/yhPW7/QUEARhCQSQoUhJNIDnB9yCKYSQS+L+d79uMVxr4ZVqdEZzlYnLBSFKGg6Vtbs1YLTkIygkCxEpiEQvfO1JaaxVXTIXWPPDn/kgN6WQiFyKAFpU+c/rqk978pO+2JEEVAzNpA20gUXzMBQ4rjeODLWkTEIzAbOvqlLtrKf1xFO/EMoFYmsEpLbkVZ+1u7hpnFn0dmJH6/1FXKN9+Fd/55ZaGQFqk2MOWyNvb9pFMzprIihopRJgWgPh1ue950ChQKORIuXK8uXX3LpcH0Jx3XwqogFygQAKy39+aRxSaj1ACSfCkEYEuCVVoyVBLS05J4KE+TYkQ0RkkApx9SNvf/UvGrcIl6Wobrrf13/SYpQb5aIgGQ2QMKGmMa0WXAtqEZIlS9RoGu2HPunsJdXbKm1oAKdOV6RQKUJU1AhpNoRulBadfC8ikBsRSKSKV97QnJBBjy60kC/9xnsPCTlpoDQsiUg3PfVet9LX/vPTaaI0ESZFxAAG6DhKOx1uiyxCg1KQULFHJcRQ8G3XhVRALsCFi2K0DOn2V991ys26Y46r5AyHHL/1XyF1AqGCyMhMpZv9lp1GAElTKxaIrZJlBWB8+18BiIhGhCvFyO5LH/kGQCgy4cH/eBpJkYsiCgKc/MdN911JVsdsLQoJCoUiohHFHlJId4pzGsDk5YNAFIuKQRD+6WNzwx99JDKRmpW2/j99wqWyVES4KqSIrBG5cGW65faLEoEiISAC8fHf8U237ZBlr6lXxPddDAgFgkSEVMk8bAeVayzHpyghQrhFQIL4Oy88XA6EC65affUXve0QSRMIRRBFY3rZS+ZLb7vlXk/HUi+UCmKLUHm0zKq0FiyAVYZMRByoHWvZ1BSQE13/unsSF0EQFRTgZrU1Zy+5bCyvmkZYjzKme99vvkpxAlSgiNt9qYGaeWwVahsqgLRwRxKFyGHrs1fLUcNNEBEJSJmrXvkL3/U6gjGc+NWPUSAmCiIRQGpuvz9VMpw2U2NjsGPcKkKErTUaraaSESLnBCiK7zaIshRAIUDnd5/17V/9tZ/pKdJQ3HI5U+xFryICV4GQpVCLZo2MGsq/vC8icsGGK6Zc33jd+74QKb0wd/2+TwKIFCwQGS2IjN57H3JwzfrxoUqBhdiQjP7w4sWsocRJLrQo3PGVX/Kw4GkUAjeEBFP+8cbcvPN5PzHmVmMBUB6VUuTxBbpsppbbYg+EAEiKEGqPppSoNo3Wv/5XSkRCRERsFhKMk2fmtp6hetcYI81zcfkgZ00FEQSQO1qsm7fuUQhgwVosiQVFmwMKcObVb6NIEIFgQ5oyFp/9/Df/3i9d+4Dpjk967L1O2RWkpAIIQMit/vSSlFKMmE2F9QQgACEiUkhSS4sJOleWebaEoEBBYouG3+ZXb79p1eZTpCIchb4Ik0FEFbFZRFLS1NSGiyf/2gYFRESUnAxd9Zyf+KnXaDW4+iM/5REXEwmIkCBSpXD6jDtxDUh9wtWWUIRAbG2M376HDWYYXXABTeNLLvvCqaUsgbAQVF7/k5fnqr+5e/dCmYggVNkLkdsizLhTJmvmkWBKeFQAayBtjAmp9oeUxe/tnQIREYGACBBpVScbq3LWqa79amN0a2g6UnlH03Ole0AfsV1MYBQBoZUJIQMCBBrj8/5OKYjYHoGsNuqSB37tmHfSKvQskMtVaIsLQHrVm67MuhjINbY0OsWhBSJMw4WqSHKuQLrtejJIEAJEQUqhXXrfrC5eLVAlSTEvoq4WgUBsD5MySZrXKTE9fUyjKSJSJIBKs8b0+NWr3nj7Q07enYwALiEQoAglSJlHhhMULvr1XqTSQBJABHnDv91lJcgYSbgA45bf//HvuFfm5lEOFPujXffL33bJ6bv/SzJJCCSsBVblrULNet37a9dcN4+Ou19jKlApwjzS9ucxvBpOgA/6B4nIJSACAaMRdOsbHro3TyOrMehNcx/IU+fg8SIVdq69Xp2zYgYo++GjCoEigJly3I/YvSflAIP29zefZjQECgIRqFFTpDbtliA1BRBFxNbCQuud71wva2ePaW8eVoAhljYF0TYgwJllX2XYSJzzJPMvkQZKIGyNpAahTrADkSqKFogJBEQQQRTU5b25z0m34abnvmPImCwQUUQU0WDnwQAREyhFFBTEZuTJ2fNqWo6x2Bv8nycRShQIQoQySv/yqlO7Uq0zxYkuuIBau/7xlzzubijdC09k/e9/8Fvr0/m0n86Alo1IhQC1jk6gYmXuPqW7E3tDYUBGBCmkmLisJoUAPOMZ7xElBRECXKRBhVvkaRoz7ukZndGnrr2R2wkCVnDm0Mpkyq02JxWwUHn0TBnTUkJh62P/sk9sRgEiRTQQQgVEWBGKEAcLZee237/65BmKoRIB0GAtCVzRNpdIRJsqJAp1rja/u5cRIhwoIiJXAlIgioiiCCACFIUoFLhoUhmFx/9TRUURBUVBUUSQpQgrpBRIKiIIkEvrRtdqWo8s7/lzFw2kILZKCFT2N1y+0+O5WFdG5QIMjB3/6U/dcP93O71oyfySp/7BS3z3vXf+6QfGktxkEMhYKU5KKgKdk2uiQHnUegHITmXiCE6GQWMkZnP6iJs0GluDXEoEAU6OQJiH504bncxOLlCsqUCxysF2pNk83h4VmTDu5YQWTl1WGSsc+Fc3XwJEVgSRFZICiAgWEYorLnDhIhIJ36RTw47H2gmbo41SRqPCgYqGrJFUSxjB58Md310FIEXJhlFI2pCQUQi4FLmICJCUkAsxFIUocyXw79dftV4WkSKj0QKBiAiXS6GUFEHgUhCgMZ09zd6JnfXuNe/6DW/TqMhFRIIAJUG3/uG9M0Y8SBfhAqw0rRcnbn/+H//8M172shf801++6sTpq/r+Q34HNwoBRILixKY6oVIrKTlet2TJ43WglQANINbKMlAltbFRt3zmL0IEuKAQCBfCiUeGVyP2NNIjvBYtCBRr2wEOTCv6iBUKpLNmQUDAxSgHJGVLyp/9e6Oh0SIiUamkCKBIJkWEmiuASIGA1G2/felJysOglDckCokIIAJSSKqGKjVVC+elvndFgESWCKlErrgQqYjRkIhSICIHFZELC1kKFaYZlPUn/nUbDRGloCFACASFohpRKrBRBBeKWnTxvVff+cCLq5ZKb5ILKYAUIimZvvmKy71KhlEG0gUYgsyp5jZff8Ot48RitDt23v93+noZESmSRNSgI0XBKtLGuQZo+xYl6MRSqYFKqZeEHLM5lr/0748EXICLCBGgOKHK6H328JxVbDLa+QBpnQgVzHDGXUq3+5FSXVS6GoCKkO0MKcFsd+n3r71b5IZAWECK0SIQLiDCRbEpDhafvbyMPkrpY07YjNhZEEDBhYuCKOliYA9L5wVnvgbbIApQhBSx1SIgIkZDSSVFEXClsBTLARyA4L955X0d5EK4cKFEAEHI0EABgYsIsADFv/yCdzoRwaAsIAJJAglSWn3/vXf2KmOQMcpcmFUjo5ZVJyfGtFzdcc33fwQpkOVCEShCCgSEChW6TdWK9hF6eUGbiTxaYSK1UgFJeWPonW+52A1UKAgUgUNqZSV2d5e2JL02tbUWQZodJ/SIKpbbE7AGgSqPtpBcklA2CLzTG0xhKaIA5GoRJCU2IwFRFKFow/X6Z1yyaCvbHkHlDWDZdwKRUiC2hmqgQSqV88L60VeURYDEAYSCKnIFhULQIlSoIgJUFJWTjIUaSLQBKO/xegUUESIiREQEWAUBBJEVKaooNQr29N4ZNGhJgZCIiIgAjek3uaKH2QNrcKE2bmCkYeTOR/4yUSMCahSCSBSqwITKRNpQWRKY8mhTCVjrI9yGUJZcMkBzGw98xcnIElYRgYuEVavS8BjqWU9xFNmhFgSpgNjZR6fSDeW2wARqpQJFjhCnxUNhq1LTGz/nycYFgiBCgdgqXAgUAUIRYrP4yBOXr7uDR0b1sCnTGyCIIhREikpa9UocztPGezLZFdAmBBeEKIASNBoaLQIEAlAEhVRpvYmEzZY3ftsT+jRaiBBFEJsRFBARQUJFgaAUDRa31aI3NiUCEigSKWS57X3dfRf7Q/R0jCsXZgChBrhX2q+95/56CQE3UUSRoNgANCATwLQPMAuT8KhUoRWE0oDZ2vFoStqK7Up7872ec7cmRQ1CUFnNdDKj7jmJmCtG7JQRiiBQMRytBUl5fE0XYq0CCCRBhIbYHtynn/+nd08BRAIXEBGpCBUUJBBBYjMa9QVvus9qdVH3qpskbA3shRaBEJGIhC13JJD6edKnN74zTUBAEEQZSxGOIqDALWJrNiIRETRoLU2RN1ztu/79kRYCik2BIrGZSEG4VCAEBBFxx+nr//7jxHaXhABEJAkN1XdeeVWb3SsaGSFcsHXG5HUWizP3+iUoQapUoKRBgppJZRuopEIlH8Jc7ai9Kd0PlNtRLMVZ2UWmcmK96gYGwE0P/dj/9V5LRl/2cgU1j75zRb9+1+ysl73anKUDxasj1wtUGIUKzZ5Tz4GcmUe+IAWqWKnU8uH0RSNSRdmy2fmA/3jAeqpeFZGKAkQuIlAQm0EcqPXyj3/vitXutBr7+6v1PFC2UJyypSgKEkBStXMxHmMgV50ndP7tE9CAuGdG8UAU4CoiAEEhAgiCIIoi0pZ95bFqO3M54DHlnV91r9UOEOFCuBDbVSCgXBARhABGOzEu+tHfvtJCERQIiEiIgFbXPelBU6a4dzwgXLgtamhirvf/JYYEgiBABFEiV7LqohZAwNr2xcMRWMhtBof9gtFQoQjENOEqKRyss7/6gZe+15e84+n1opJkSXlxyzfUzmqpPjzUS4O4hesVAYsp4CjsJeNthkfP+wJYRyYymXJl2pdSlMNha/9BL33AsOw0gkiBKCKiQEQQIgJcsPinT7nb1X099fWwM8SBBY6EIiDCEsDAhSiJnC+g33703IZLlZZIhcAlEYmtUUCgAAIUgpAyV8pqCVE0ucsPec2VrJcyFFHE9ghwsSkXwpUoimg0li/4zc9fDGjQIm8oEhQR8NF3v+xszXFXL3Uu4CqiUuvld30Bab0iUgKQCyFiL3/4u74oYIU2QJ0+sNtp6TzSlIOMmSLnIVMd1rKRjaUckKrF2T9+zNX3ept3+uh+cvLe665/9nPeeO3FWg6NeWRmKFG5WBcGQAowybTVuLW0PtZVqHshtbPSwhtEp5zB4U0e+Dfv7TSLxIVcOJL6pFEgl2yRRK5itOd8zN3utk+tZ6/7mD2SbTK0RCggIgFlHEfC1JDOn3rW/V8y9ZWVSBbEpZDIFUAkgggXIcIlF1ixBiYDp0x60GL/4f9xeetqBAlEUFKQVFlEbuWyykUEZjSyn2978VVRASgtEW4hUnopX3TVQ87GzPZgpvkCDlAa4+5Py1ySPMVTJEuAJJLip/41f5MVsEImTHAdOI3q2jeyricNuA84KApxL0uqSdMeWxVEBhePm5437e4XoY2dheqSocFquA+LGjJEXgiVap2kpsGE7CrrfIsM7ao0dYEIK5pMURE5BEq973d+E0hRBVSuAjxBc0VFKkoqrqTar3/L6XuerXU3I+uBymwfSwARqU+MBoqA9ClG1ajkvEna6x7yvcvFoBWUgwqUggoKIEIqKMhFCiqiICiiQbmAQHpd+6AnffbIaIIIEFAuVElZlRbJlRQoqCKyV6eu/6lvUdyIEBIoEoqE/veND5/XVs+aISVc4B289x+kL+VCoLhogFyAq77+r/2TX6SMTgACAShnKEzLIyUl7MViwgSK1OjuIXZXtgTiRuta7kqLJUzLpVyrPtpgnucRumYFcF1PQCiIPFp6Xjt2DyeP7+YgQ9GhVpSZxEUSh0OH6Js/8tcuSoRICkFIAciKLGIqKqj5S/7o4itm1k5fzWuGu3MAYSlFgYk0oigUylwJDM7nkEW+7iFvWNoKpgKuBEtKtDFahAKKggCUUMSxq0aM2ayEdsvnfMKvLkafCCFuRIjRoKIExiQllRQokpyd1Tj1q6+759wIjAYWRAhR9U2vePjcY/eu4cGFn8W3fGkfDRejCVdBZIEQqtv+7X/s++92hiIgtZb2Ek2nI48OWzuzRtoi1Y5Um1rvgqqxZWv6oi+Yq+0sUlX7O82lzMx9mAFhe3YRaJhUCsg4lzHaOo8ZoSFVW6TA0YYEqXBERcun3vXv38EZVVEECEGEJBOlIIrQyx57/T0XbV3yPLvb7lHa2DaYnDRBJJcQIkQllRPczh+IF4sX3+/zXXEiCEJFBZSg0DLUMqoiCJCUFGv2Qj0AtSFZ4Fa/dZe/fgeGqwHNRVIpSFJSmIjAEAShN1qr6abPfHqNJtKIELgsbH7qhY9Y7cuOBmsSc4H35HO8WsopMeECcElREbk9+3Ff+s5PXtCugQpYcLTp5Y1p20dK9hKwGbEFRhlSRlsPEg7r5sYYbV6uJntMWmdWujx3uzqQbVeQKohFqNpUlKQ8vq/nGoodDUApEpIyg6MG1ouzj/yQX7iGWC7E9hAJaEAAcd3nPueiuxdDffSx7r3PhMHg4FERCoIiiiAlUeWKTHT+JO6j6cmXfsUdTS2xQhABI9mUme543dtUkCuCIIJEFLkNHAHEhNFG7d7+ju/6w+9EEkRFEQoQgQWjQAgCIMW7N07oudddEyIBkkUjsvo3vuid05mHxhhDXeEC7wOeMdaLoSxGiTTJ0MTWxvjmP3zbX/EnPt6XrgbLrdB0ru1AcXhcKksICmBdndodqRDNHErQ21yTM5hr4TH3UfHA8whW2K6wGGREEEBYc73AEOtbOFRoRR7NgEiJkZQjoNDbyT+760f9wL1LLa4oiAjhqoAQXl37jX9x6tJdzSfXrMe6r73qHhxxFRSJrUIRSspukjAj5w+tJSO1/qErH/2He9CCAAQCiwRe8g0vfQFARRBFEURIy2nAsDfKgJqcqT/7ne//LZ+4I0UQFIkoinorlTawBAKiU9MKbv7Kp6RPLkuhJaSkvY+9+O12zlipdWYPS75w09QG3/3toaFMtAhAksCuVDR/Uh688yPHA3shILUCYu6uD5y7xyf6GPWegpTSgCNmak1SMq05VLRbKtFS9H1EWh/VmavZCgfxwZeAGt6x9vTlNz9in2PnLS/IeQEBijz+4pyj9BBx1CCNeTH/8UM+6IkP2FU8VKMVAgo0BOz96Q+88qIHL4B5ZLVed/dVz5jJYdrcLZRtESjAMifW87BYtOI8Hp1gZ5l//ri7f8Zz1xJAUjCmgnrlN7/kpuuuhKAIEELR0Dih0Z11ptYDGCBzUCtPr/7Mq771I6/aBdwLUlWFzuxfwWYUCgJC5dtOXLQ/td/9zU/IvECAkDqlV33I/R5yao+MMbySMcmFm8j1Y2/XJycSlgISYKRy9elNH3u3d3Cf/fp8JVBEoELtOccpjNTedJwbChiAVNKQy8QB5QDjeSLLoaKZSmUeJiGJo4OsWinQBio71CkDdeStHVKhCHLrsA45lJCPBEHolPOnj7j6Q77ivrsLKEZzsbX1N//jk/95deoel47Z09qs1n095q6BCYc1ATeCiIQLkEdzkwiW0Pmz1ZaWi8XZX33Xu7zdF97t3R9wcVu0MY/bXvT3t/zLyy7amRfGlRAERK60ELchK044ZMpkNFW77fE79370xz/kqhM7AB669aaX/vob//hkAKEQgaJGOdVq547Hv/RiraYSLoiX/dt//AFXrscq62SQYUfWhRvgX/CEhFsKSYGQEpCxfMZn3ufBe/vt9X7YBbBMKMJkp2xWmbHcrnEOFAZDhULLHIMdWYRD2vutVo1Wq0xrRepghj0gBxFDJ1QzKl2tnZKrnp13uDDUymBBKDwANnY4zpSn4cUltf/b73jFXR73wXc7WdN6CfiGF//9i/79zT55+uTJnb1OVvNI3+99HrMzzKEbVEMRiYSrAkS0uEoxPcX5Lo951ImrLsqrfrXXvDuNnHDWY6HpsmJaLKmIsFVUROQ2FsLJ4NAmo5mW6eK69res01ecXiyZb/HtnvesW66RArIQCoqY++Qxj+WbP/CZam4gZ0z1wk+85G0u3u8wM2b3UY4GF3A1/S0jzUJECEgphbCs9uSfffDVN7u653nxppYaoCEeudwjkceHdQ0U1mQCFEGtC1orNK04WAG3oRp40FspIQORpAzKATRWEBoAKzk9Fpkt2se2ClgDQ3j0IY7sKDoGZbRRpGvnlPKy/x3vLE9edH1f9jHGdOLUfXrI2VPS2hpj9LGaGVHI4TzhCEW4kQJFEM2tUJWq2TnfDCje8+LUKQqk5XTrcpnmTUYxL0DWBowilVFIUpjmQ0FlaPKofmKxa9pr3bqmtuhe7q7mi5XIRUHkAlxS4UXViX/+/q9ZlxFNde03/sH97157PR5aGXVZPdIFHBafSVpvCClitDQQrjR05ste83aX3qYanvOhABMMFJFK502GsvdjdLsoE0JwIlRIiIN7B2VbAAbQtVKlQ/U2RMJgMxwcupxQHJEJHbLP2ei0PF5GqVSqQCsLewgsH0Mgo2FXn7I4sdN6GGeWZKftTGoLObjNjGHmzPO8po90zOEldgQkEhEuIBL0Es7o5k6YEY/uMYZ7a5KYl9Vd6U1SLVMEFRGuNAK00EfkDI5owHKf9hb7LcsdLcpTq1uXpu2MCIqIgIhE+pJRdhuL737aI8uN2n/O9/3G3d5+WWeUHs9t7S7bUriQ25/wPSfdkKwECURQw8WrPushDztx06KPYR5IgQBM5NFtl2t3tjzezKLTxYTpstUWCiEDKocjZlBiTL1leHjqhKOP0owWcQK4HtjSOC1vda0NLVKZAIFczVAyOGYTpe+M0ZfpNVXQ7uRWmisrysoYnpnTh9019yGO7LAbXAWFIhGhUVRJlZjozrA5Y/cur1rT3jRQ6yMeg0zC4IKoIhchUE0moBwOCI4y5lqyHtpvVO1Moy9Xyx1XBeQSmyYaeIqXo023fOSXf9CD18/5xV9+5T3ud/VqT9HMGIN5JC4SLuz+5z8/nuoTioiEIhCj8Qdf96grx9mFWU373r9cTHhUoAJhUJD2MSa+VKnAgipW3Ko80TPCcVrRYFZw5hocx546KVYaUhncTE5H3zbnCAGQBgqSVsooxwUZXas2q8ueJ+F+djEDNReddWPlkTH6jIeBQY6CySKBiCCEgFILUw1ZBp1/0D2P1Zh7T8npkzutixqzGROiXCBAKVSUWuQiwxzjoGdRyXqajXoz817FWsgVRCoCQYUILFHr9e7t37rzJ8+67up7XHrZtL9fNebMGnSsPhl0gYc/frtHzw0hJAAhRdW/5l/ec2+kqw/mOvcSKT5WAeEKpbS8ffYFGYQAWs0u2RWjEY43pEshBlM+mnfjWKoglNF1nthi5207FagAApa6DRCmfnyYvojTe2pB02hrNWlu5UwMPJw+RtdsKxCOLM4SKQCCiM24hpGKJLkzCGwnozJk3Ic6QvPu0NlBCnGgAEbLQCIxx6pUxpjobSRTayOtT713TxEgNEogIperZYxQZHd/nO6LRe2f3UX7rdbJGFbicCH4CV98j74IFrhBIsTqy8Y7rC/yqL5eZfTTr97IW8sEyovj7rwn3l2vj4W5Y80CRhGcsPK2dUdU1MJ9HMtmrCAI4uhf+ujYIkAnYsNOdnn5xrvr9eEtr958tA8p0AZAph+d+6pq1aZBjgvU91s1w1CvScIoK8AmxehJ4lINSiJHCdfevjNFgEsISIq8adob6zlq4k4ZqtoY6W1SjQgaONLNO25vXC/7pMhFUGSppS/Pntr3TBvyMQT31tc1zTSaZzmt9/U8IjBpNCCiYHf3lnXnhHt3vLg4O8Nzac3Kfb/GVFgZJeUC0OoJ/+e0U01Kw6KIXv9R93g3sqfeBx15Hn5iFkDxJoRpmul1zcOezeMTTvaFigLFQIVrJA3HKMcEgQA1lCO9vj9tFzBZQK04kw5tw+O+/pFfM52UYritPd8c0xxQInP8gzZqDHUmWqeUinrKrgEeGMeOwlA48qAtKiKoFAhSgcEoiXPnEHZXuryWqlKGQq7qGmftiVEIUKASKXuXjmkuEnOcCtao3rpKktuYB8wuMDSASERhfTa1mBV5VG/V0FBXBhltRsMZCHNB+Lpf+IaTIZJCo6vpqV/1gLe9bR7M8TzaPA+7zzugdptRBlbzejPZdfctDD/M7GUrdeyQkT2jxFEP4VybcLTFXJiUjlanwpt9SevMfkvJ7GUXToWK5Hh90uJ0HI5dSXVlp9NqLU2VaivFCViDmqOugboSjnc+c6kJVoVUhcgamhkwIt85AmTMWavcWlofSjVH1d0LXMICgWmEMIUxRjjmQNzsdRQ1I9QbHgBuEEREcBsBp2emBoOot87AszLXTEYxxIXif3vWl5wgKSRg4uv++mFX39adIc+1muXQn7i0zrA6tducjZ2stR2jfcz+sYbTTMqeOA1l0WQnGcWdcx8ysJ21laGr8joHTivv+IOdblJzvbRM2D3OHVciUI4t4DKDNhelQa3VsKErCp1BRsrloByHzizboKv1qnjQhkzGahkH4Xbn2BqCeuulWUmbJUWjWHdrLlkFpM1DuHZcQ0rRxrFsHZqnURrrgpoXzloaErF6c/UG0E/ecWpY1aezoVK9BpIdrUYN4VpXuHD8dy//QqUFRbTXf8R930Vn5+DheczWSD34jR/6xZwrjozanQz+8fVJenZfh7eW322YY7ebVeoxkeuUIM3pd5bf9c8+c+3dNZuOOReGfbp3y+47/Kd+98oDsad7ewgLRo+O43BuDUOdqvVijVTzUEdRkmgA1sxmOM68+ZZ7EmpAl9BApq3uGCRKPO5E4AwGQypnVUqaNKvfcg+PxUiGFHXSq3h1upxaE85hcAbQ0UId1v/xXmKAa9EZ8pho/IdrjfvKgS4zylnTGQKwI3MBefzp6Y+qEUnwd1/+sLuN/RGPMUbSNRiZS/+l/8/vuBzr0mrdx8jwo/+D7A6nY9/C+l3/nv/MRx++Cl37qpvd1w83/eyIvZ5t7qz/yh/43kWucah4ofDm//6a84Qp7/j7/7X/tV++rscrgLsKMH/iHzs1gohyboAQzY2OChcyqZ5MGUkUg8Lxqvq7/vg7n961FyysyIXm/e+59US31+nhLWHcEQzFgurv8qtvewlpqVSnys2+7gvmxexhzvXgwIEr9Y2/fLdlxBQmgMUYWn/ZcvI+cyzH1SDDRIlsc+G5/ciJzyQN+hN/5+EXr/vcLY/RnbhmwtU/8Jvo5eNPaI0fvTkueJ2f/TCeKOXt+/5//72+XOu6yXFs1sNha+/0HvPwnegbv8G7y+XavY7tynk3a72+rtl7t/Qd+Bt+07HQuWPIC9jzcH952cd64IRzb43m0FDRSaBkZgNlgHDMGbr+Y0ChRJviJlUYy2QMkrcMm8EKGcqYrn2/Vllgpka1Njd21reeKkaSc3VwBrjGPz1kbotFFm2qvhtPo+XGnd01rY9gGVN2c0jUmwnJBSZNv/ZnTzw5/IyvqLfTPjOZM4ZmZuyGAbqO+fToMVnztePT6GXu+lDObt792Oy+uXsR53pML8tZX6qhJNxpc7k/zod5sT7ceUU60/P44n59PeecMm/L5br3dWVyfzng/vRqvuy1d7tTnI9hoAyAwmU6KADmHIdpuVo2S1L1RU+1qauNfWIP3pIGQqBPC8/TejFFjDqrqjrDZd4DD85bhYRT82IRmFW9SmFitVuMrDUcFwpD6sJOCBAuOHdf+4Z363/3Q/e+7+h9MCx3wdzCUFUrs5tjrchhhTHcc73uvc/PcOKeuwvHOu4P6MpZGn0EfKeZe/by44kvatIvHCMcL994nb3l7XPPhbs7e6eauS4rXG6tMQ9z3gZAMZjNcH7OWrd52WrSQkYlxaU5sxXeIstzLXbRxFS9mlBLl7tjnz8BzIrFotU0TbpjqajhatjjxLyb4aR6RECSuHB99pd/5l/Wd79IHhXNw4PMStcQFp3gecyezbGztXHoSUnzGYCe55nOzMNlMoPBZqDcWW5de7WnaTNj7dk5O2VnvwU4j309el3x2Jmd1m26PcrnzdYAypbzVi4Nx9OY5qqRRoYZrhVvMZVtCkSFam7qi9FsRhYDh4yEc65s0wb0ycPD7lIveVCelZEe41hdIwKFC9qaRNtdj4DnpJOZTuaqQYDJrL2P6/KMXZtdvM61e5/zmQp755p1roc1DY6TOT3cidvrJXPXtTWsIdOec5Zd3rlzXpdZWQ85N5JpxzA1OP/D+W2t5THGaN3u0kyIB2v3txhhewDS2yjNra2r2avSLObEHmOcs7A9bNUYfT3NU7W50dRVow0ye3Q6zNW5IG41N5hl5MFsOeVFD5A6nXGu65osvWZr2cN183M6lzOXnnNdD0oSegYodx57HqfHzkOWi9Nr7HS7gb4T047TcxHH2cig6vwXmCCN0ZvoZVQMZa3ReQutQDRPNfUwratUXcwxOWdHjjDz1KvsWrdC6podJ9GIrAtjtIhYjlklYGutweYwtHtds8x6CIBt7rvLzM/B+JDLfsjVFUQzIJ1w5y105yGra7N2MqfjQ0vLZ7TnsY+HnCwfFChrkvwXAGOesNdt3RSJNpIM5y1Vkq6BM7des9zT9is4Pu+g16Qxq/VCXbRejj0yh81cIBsaCyUj7lLHBMLBdjgzZ9YDgbTUK+2Un8uOp+Y8zjSVhC6HO3V3j7nsnOnqcTJOh5bPPKTMPs51v+i2TEP4LzHEmhtTMjRAjun1FosYu0an9WkulLJNRrgTyjU31WhKL1dChpMoEC6Yt2iIdFsOMuGwBWbOrD05wbqdmcrP9bmuxwNXYqsY0sidjDJzNV69rit4ZTY/l7ZzJqcPXLeW7Zj/OkdpbmtmVAYGqcFb7kBv5XTPNYiUHg3dCcKQ2twYJSooJDgmXFAP8siAka4k8qEebc+YXo9NtZ3Sn7OebAaInG1EvCUcBonNoOXk57RwMmckkxMK5b9SoxG5RKAGGryFH1CzACFjBXOnjDsjzKgKY5QoXGDfeU2A4TonLzYQz+v0JGNkhHNrA/IY6ZC3BDZIY4szys91cQ0Ur6Tlz6+CIqkEEcJbfpFKRVjmzq6wVZESLsQvC3gOy4tPsD2GE4tzXUB4C9rUkU4YSnbm5wh6yngCA2D+CwHKpaRGWeG/yMIoUL6THaiIcCE/Tj7oAUhyzsxb2mEYdzMAm+HPYbHlreWtLe2D9l93BorDn3ML5c/nyv/wiv/XOtyWP/fl5/3/8/7/ef//vP9//tGzL+py9sZmPxj2hqQXNjvrZc8XZG8GewN6y6S8pTlO+/lQzrdK/i9L5PiUDWVb+S2EooByuHKRHKD822aP2PcgO9y92U2aBefEeLG2czKzIWuPXa4YYA9s1mmuDE3Tzurqg8lzL5VuR9FxwnEay9L95eaeYwlwXIIGcGK2yZOhOZmOzmnaSXIvV7LnCAhYwDSpMYlWVTVRpfNFzVagJdpQSMOICPyWockJ1qbaYk4OyjaoOZCNoJwDEWjOFECDhis4UyLLkg+lnG8FbVKpQdVUarAU7fyQ0kYNSoBGWlkhKYXyOIwI535nLohpTiNxMy1R5EyRU64ypM1qrow2mnMODmyXX3ZPljsnFjde/+ZbbwGUqW85TtM1ZqeO83JlFjjOFxSw2nUem9DMulrIvIfACdwDwQKdEC84OmGvwEI8egUCi4exdHKFhZMrnCwsz3kFFjgBTuAkuJwBV+Dk0U4eX4ATrnACLJywXGE5YVlY7oF4zhIBZug9gx6GbXfsnC/pEGCANzYHYDCHLd+p5vIaShkGYE44AQrpHKgQzmEABnS2d+gcY3nLeaxthtHjDLC7M2AdBnU+nJ9CCW/RlWO73wd98MOvOiUXuMjq7Bv/7C//5aZw/NESLAsnV1iAkw/ihpMTNpxcKTC8h/rkm3zyu9/RS7d3nLdXnl6+5CW+7HP/9/fsC4H5Ht9quj577ebp9Tp3e395y8e89x/8pwDN8k2++7z6zmeXZ8e9vP16Prnc7XzEV7z87HKV2eZ4NnN97by4F28+cHjofPgHjpt5yxf/vr8dMG/95O92vd1nd+7Hruec9/thntebe/rw187mOnN5DbybuyevzjlzecZXcLM3e/2oT/i8n89zOwB3+bAnfut3PuE7/vd3f8f3PeEJ3/u93/nE7/rR7/rhH3yMzofFx3/gm/cWi7XEqLkWjXVadrzqcytPvvFNN77m9rP7wUi+s/yyf/Wv/ujPXGHmxas3D32R+8uLrz14FnvxbXdc9+Y71vuvMKjEOLbLP/Ij3nT7/kVt3/JiNRWD9ZSWi/a6NMv/eNP+9ft7gwOV8ybbHvVTP/fTP/oj/+cHvv87v+Pbv/3bvu27n/ST//tHfuhzT+Nzd+mnvt+Nt46TO3vVs7Ov1XLtvhy1e9Pe/txWPvsfN67PrtksK+fu5Ie/O/1Mb2tdMeNkUaf2FXUWZ9cZtOyePDv35dlaLG+L14sT7d7X/+cvcOzTIz//Yy6DKJY1UsqEknrjfzz5qftA+Rhm+Rbf8m378e87s2dPLu98L77j1b/1337BC/tX/6Xf8Y175ujXZLH4vp+5/MH/0f37yH7jP/tRPJub3V2cIw7gs7/x3Qv6Wb/uBNiOhrXDVy83n/WZAJe7b/Q3OGcwwVN3Gk9MhozoOGcbc4+VhtMY+fx/758w+33+wNzIKbSHp0hTrLMSHJQ7Cw3BnE7MebAcP+M3PxcgHvzyP/mab/7Gb/zGb/m6b/mKb/zGr/yqr/26b/iSb/yyL/2j3zsffuFTYC40igjVUFAaxpnE9r3Vq3/3V14/zJ3RdX7hd38nbEtGSEE404KUiKIxzr7oKX/5Is7hz34sMFSDYJVcRNJIKYrE5urGP//x580b53v7s2d/3md/3hd+0eO//Cu/4Zu/5Zu/7du+/mu/4Ou++PE/8zrOwx///AZ2qhK5ICWneVTAqgbpt1z/60992T7nvo2XPMhK4hYJK5U0AsLIhSsVV2ooRGlypg95+jHd5Ws+57TogqhcVtCY3Loqxd61P/nzN4N6Q8sP+sPc7wVo9tL1wnmB3/FTXtRP++98ABT2aiN9uJx3f+2/6n2Eb/MRnRfuwUPonkPn63zaX3oxfi+8jh3FdWTOw9dufwoP7/lOl/vDzWHd9pg6hT1qNztwzrahlvs58VxPZ/fmY37oLz1OfufTZ+cqsQOHi3uqDsRQnrOT2Tq10/0Mxt31pR8D2PMw/e5FibsZHtiJRp/HMB9x9Xnw4cHNpkYJpQcqDGI1j0WGSzVd9A5PyJmnfv3rKJ93Lb/iO2ZoV4e6NWyZnTjSU0Kj2qWP/hH2f/rHXn5cy3eNPFpmKE0G9waGpGSRpJolsf7PJ/3J7eeZ9KgPZqyHoz56t2LbC7uf/pZzN73pdAFW0pREWIsRgEwDATQ29//t2//t5nOkME8jzZIZqckKlqSYUALSW5WJDBVRZ73zygcfy31+8gMbiZEKuhQBVKwaNWhWnv/xr2T2jehP5rzsLAxW92Pn5Yfwon/Si2scKtc41XkovxEy7x8fO+fl5FJgdROX6+mH8WJ7iY6hzTk4ajgue39wwvTR7HHsHjhwxFDZqAx7kZga2YajS3Q5laNz384Jb7u72WPdlTnHWA/QsWNHrnTZcU66qamcY3SWm3vfAcTz5uQDEy2nSS1I01SladFobp9yHgxXs6WaAKOiUgPVVLAItVCq3Efb+eTX3vDxf8f5X777jFotJYCUVa2TFoow0qW++3hu+Kw/zrEsdoDFzEQkKiwihCFSEpGSkoyaGH//yW86r8jH0KvV1CimaZJKyVBVPv/c6Q2XSxlSs9AgRUEvjBCBIFmoSH/pF/+Dz03tNYFwJIQcCZEoUYVEFURvyBJGY0zX3jtHu88zr5wymkfR3GAINixEKkYjNXHtJ/5l3ujsWxggp9kd3Yvezwua3t6lq4WlbHq5djj28P750XQenZedhFya6+EL4mwSL2HkxOixHuecN+/kOHbGGJYRVIccYARhTpsjhGS4LOIcF27uL9cne+yhAIcwl8wYhoM8Go6aw9WJqZkJD67tCbd3z8WnLXoKBlpIlahaGJH71527OinTlklFYcIVQUtEkBSNRdIiO/s69bev/JCXnXfDwypz4ITMqmUxmTQchEpVIGhp69Kn3vyhzz6OdhKKBS4pKg6sKIJAUm6RM2hR+7d37edR+HRUTai1VkFSbQ6Pa+57znaumxoVN7a7IAKSigCnuSKT0dxPrH7zC88eX6EztbAEQWyNIra6iDg4YqtrPb35Aau2Plx903fQ1aIhKZEAXEm5XKYASyNR+yc/4q+9AfswxL2sMHCzN+Ec9iJcDmASzmNSj3JdZ36Q4/oe8k5wHQZ2YgfmMi9o+JLzMgSCsSPOpQ//soN98uyjURvIFJIQSHIIzCN3LDRrJnfihtiOk4MSMDMRgQRzEk0GAbTJBG8vJ3PH6wv4ZrcWoYKSRKBok5qufhA6R54rAhWKgEIUqCIAVRa0IGqKgPu96NOecr5lfsn9CyEIiFAJwQq0VAQUpKbAZc/88w+ZMo4yzXJFyBVFQQVEAoiIIAJcUQT52v9z/vAlV3WRCCBFEqKWVmpP0Lla375kawQutkYQEUVBQQBBrgh+87POIvk4gPQJgrZFSQFRQGyJIOKQI9Mz3n3KOEz5Zz8HIrZGbHVFQETEZiJQPul//89+1mvPdV7ecjoDw/BwEHwaLzLe0kFTObCCXLzHc/EeesLABA4yYXG+mJ37o9wjCXJ1LjQuc+UVTFhBlhEGMsAMxG0cciVxyqS5sB6LBQJRIA8TyQQJzFwxZUXyOBmy1wvc9apeCBUEiNiMo+jLK+3cTJNTKYgEURSxKYJwKUIAo6IIfv0vP/A8G45lQ7XoEJlQGStCCBGXC5QRffAt97qNo7ZmNlUIIQgCCGJ7RAQRBEWvuh/KefI1qUIKCQIJiiglHnsVOjft1lJtbI+IgAgXREBAUQREjmr1ab99fF1IgAsLQDBKBBEEEEUREEWQMT39Q9vIYXje23e1KNog2hIQRLgIIgKy+RN/wWc+l3AskJkA69Lygp9NnQUVJjIBWez3ka6cQ0M1PGxhsRdA93vMHpmyBOdxzr6WMLzGeQEwYNjZ2UHKSRBqaLUdVjJnh+Qpc53zPG9Zc0fMFQOpsdhJlwymkJxIivOeWXsd4GNaxcIOMYBIEB7Nn3k549zYNUJAJCKKInBBFBUoiIiWWFH6B1x7t/Nr7e0AThDCNgGpwBSRK6gsIGpD45Kb7vnGI7lGxKYripAjERG0xQVEhEimuOOud3A+iivujgBjKRAQAcl4+gJybnLzEqWICIIIIoKwEJtDRUSIQiVpr3j4GY55bmJUiogQyQUERSRUEC4OGeL29++nwWGf9/Z2UQQUgSU2R4uIiAAXriiegx/N8+8zHoorJCLNi+pwRKxKawWuhHn/2MvdRVw9D3Yw5vqAF4LgDrIzO3Gs3t7QCK8RgOdBiOcgqx7tpBlDIgy4hwkkdM/BdXCxQUKcRM4DB4RBmKABBxDcqbHLCXPyvI+3XEEIVdAACUTFu1+Gzk2GJQVcEjKE0SigoghC2IxSChPru15313E+lb3XorP2YgKrUMFJJ0SpiFAmEilan15/tzcfAVNhM9UngFRCisi1hSiIEMkF9LN3PXs+hI9sJolSoKABKMmQ0Jdxrm9ZKJVUILXNBRCNFgQEuSJGwzVKI1N/xAuOqbeIIBcQEQFBAW1ARBBEW0SfnvGe5rD//oiuSkVEQJQUYMmFKyIiinDF7YQPAPY6000PoAzAnexFXdyrRazZS6jOA5T3zwvHtIMMDCDE0Iuw264Xhx1IV88ZWGY9uOOoxoMd0oMVRBogQNZchh2PshxYHaCCCZJkSICDZKdhNbCGDBPYgWHhAOw57vuAYSWoQAQahqCKlC8i54a1IhAE5ALRcOECBC4KFy5KiLBcX/XGq8+ncLkIXV00TJDKBEuIEAI3yhW5cNpoL7n0CCHUhospCBeCYmtEn9g6GhGRi3lxyxU+D+D/BAkkwmYjYUgVpEuW5+rGnQYRriiI0YgiRglcRIQUowVGU5C1anrwK45n1GiEFBYiSiQ2g8C1AS4gIiJuz3o05YO+5+tcjDZUbI0IAhcBNBoQMdqGOOVPAfG6670mOx47kEhzvqirFJkUWCB7oQzvoU+XZXbY4XUTmfMFhB41TMgydBBe2IMZYY+BHVMSCRKjEEqB2RlAcEjm9I6OJ6zmUA6sJA9LeSgCloKUMCt5nJ148rzfLMkAwaYAyyhDFXHF1ZxbTUmESwCKRIRwBSAUm65iM5FYctVTHncelVeTOcBaCLcVKh1DFJAwFShSiXLRG+96OEMgSC5kS0QkBRFOAxe4XJELoE/rxSseeO7E3S8fU4JQZCAgMQkk0j74XN1wIiUXREAQG5GSiqIoAitFZMk1WmLx4U8/ll4iKRdRRFAUWURiM1YRl7aBXM969NSVbe/yrJFpNFcERBFAEFtdrg03o5Spjv3zHCevL/c0OeRkkHTyos+rtKTy6E5or+i8f9xDwxCsPoJEvNBCAAHl4em53G7PurAHkQgkNEwSMsQyEzAN0aREocfLbEcuJkIyIZAMPAigJgQiJTAlejZOz/NxqlYWUSQ2KwWqAOGx58gtkBRiNA4OhXBSbIaCoKCC0bL6pEf++/l0psswQR4tIlRCuQBRCBBYIDFf/aufeqgmsVUVUREBKAgiIopwEUVJEcToO1/2o+csfBckCEIKbWy14mj1f87VTQuVIsBSBLiCgIgDg6IAkRK52W3dP+BZxzE3BRHhchGxNQJcEUDEwRF0PeN9Uj7g9Xf1ZAmXi63RlkSKAEu4iAjIl331szl5zg4BeVgIErMv6jJSnAATiICQ8j4STe4wPK68/0XN2iQQJiZcX7td9nCRpglyJQeCBEgmMFOQQYgmsRtYL2SCRDSEmECaAzuTAUzyqAMYN8NcX6/e75IeVYQiKYRQYAhV5NvOkdYVUUCaSwFEigiKAiJUEIEgiD7t5E/u2c8buI5OSH2kCKWCNQIBASHLUoRV/XH3eP1hFj0CQUSEAEWKODhyERFFwBBU+rjbzeeqFh8/NzBGEiFFIFCuyqgHXn5u6rqdooCITQtFACG1EUVAUkQkIAWtet7uNcdRFQVxYBQRbYkgitgeBJCRZ71v+YDP+2kXMBoRm0FEABFbIzZHIbD25m+/b+o5DALSHCBRwV7IcZm1aIAdKmOZlXV9D7mrmwB5dAeCD3tBL73rkjwqmMAxh3eX63H2CcQgggwIIAiEkJDJwkCIiDTFlfdeZ5CHAgMCsQOCAEL5AEPk8Z15chz3vp5/FGRPgEJEogQEFVK66oGazkUtSQQgCoQlNLdyRQARkYvDloju8iF/fN6EVQkgUMSiUEEQRCCCq9JAUCxdf/uAFh+gHZRIgLAQCultlCAiApfLxfYoIuIFD9tQjs93O0VQi9hU5JitFVqTvvDcrM4ugzjmgLA2EokIIiCCUcnE1I8SFxFbAyJiNKcRccgggCBg78RzH6U1W69542hARAREEZAhlcsVgalRrhTGE7nwawl7va4Hrx+SGS/+fGVRhLrAmkY+zvY9hJkEyAcTNjxj9kVwLykJO7jjZqfrceWL2YUeJI9HQ2DJw3LHtZiIBpDioOMqr18DEDA8r6C8rgkJO3PlA8+e8pw79x4NFSaRlRAEyFgonc9KdA7UjNgMAlIGFh1MUhxSQYAImNTP3vW8GapC5VZuKyN01lxNEIFSbFoCUO5/n1cXB5GkAIIKIGIILyxF0WgREdFGALH1rm/iHD8WRULKBorUjAgIUv3Loc7B8qZFgQBLAWEKgtjuirBSLg4dQr3uXi0+SkixmRQQjYpQb+LoEUBWu3/xwTv7277/K7dsRmxPiiQW0BJQMWjBLSl45S8SzysRIJhgBS29oFoywdJgsV3X3feRg/ZAjBUwlgvIizyvaSg5gaFFcIV347l4gg0sU4csDAjFOnvOruV5nd2jkho44d4TtiGp3WEkCWQ1hGInuRdX5/4CXYMLz57nPU9FwhTBRVSAHCoVjPhiRs6BqUREaBSbRRhelyKZNCctRAoQhBuiX3W/V54v0lXAghQYoGb22unCRgEIkERSRCTzkx+agxhysakEEWVexLMmQkVRn8RWUxARQM5TPu1cfYNVYJKEBAWXjUww0ZV3w+dAN5dEEUFAoxE6odWoaAuNpIaoKCKIpMD12T/fxlFckFQQWMLSbBUqcGFFQyISljYy/dN7cOB1V/UJIogAS4lwqCCl6IocLbrKLWLk8keZteeAAJIQdtTg5EWfWEIRRy3Ctel7iGaYribQxEkvZNZEUnCHnVMWsIPPZeA6noeKbiMNKxkgAx7r5d6NndFdrh70eVyufPa3ursf9nS8P+acAfCcphwQSHHnOnOdBjrO0dntCTd383pfwGQCRCkoUgYEFZBqvvge6BwwRIiANhpRxFBNgRqtAIXRhkCjIaLmclV+5DHn0aRSJzRVLLB1z5uc3T0ZJS5EpDIgXFB+v6kfIi5tKBoFQgtorWs0LwBCoAAiIiIXls+ePkd3v7QnjEJliEy5qDI1BBJ8Huf0pqVGIwKIImDUaAH1RmpWixSKKIqAoGg0HC6/jaMAjEJA3AKKS1i9FJUVNErKKPWqyF489xGVbFx+XSa24SIiGkqaYOzfeMM9p+liXHEFhCsp/mN2luecCcnwHBtwx5MX7e0kQ2TSFpikB+U9dFEI0AhFWV7oNO0hCMnQEI7AwH/6lW89mPYo9mAu7g4wrDXADjvnQRcBajgG8rj/c+Nx/vj/5cNuePwGjk6vMxyIC+cgWOzBpeOGJoVVuPHPcsdzvm8sqRxRGAzCRQeCGMVnk3OQDArARYMA/O4X3DqNBFhe/QH3/OrTjAEuGi4CZVHrd+b8SQp7ESr7QCfz5W/w1sXdv+gzrqh5kssSiCAivLzvyw5qg7A9LRYpbv/AZ9MGmzuXPfBb32uZURVtwCgBrv3de7zh3HxWtUkpTJoiAGGAFBRU/7Jz4ht3oVxEkBSblrw/uZUKHFMlgKQYDZKiNzTqpQ8+2iiQCxckFa33dzLv0iDlxroAipEKmthr+dC/YftH/OGYIiAiAiJ6Jt3wfb964zps7iwf/dEfd0ULo5BMffHnwfK82xWThoOdkGlH7IW07RBuw4DQlvoeIgZIRAJ0vX1By1VDMiGFy9nM4Cmf/TW++vXtH/b0uq8+PXl63/H+V972556ww8QAWDJcX/uzfO6X2vs6Om9vb24/9vqX/irnzmd/8rd8y91xefVJlz29f+148p6/+PKZrMJRmU3d/9TPfcu9r3TeOEf3PLn865N/COzrvPvyERELBAxwcEEzCJD64zmnAUGkICBunL1jZ0x9Z2SsX/+L6J3+7PJ0NwBFFaUScsmpM+fJ5uVkH6ugdVGbzqcv7DVn3Pzqr+Vhf3OpS1EQkWA0aDz+Sw7qIS5cQahGE335JhQtRq0JwEc/ZXcUirBEBHHb3/3Obzk3XzzsKAgchWLIMhWUJILL7gI6Nm7aSSNuBBGIp+vf4bZpVXtVWZy899e/95Vi9JoE4IJtUUSfHvjyI4wKAogiQoqXvqMqZmdRo0++6KyRlcW8kKvatBpnbmOr8scfPlRAOACX64YP/VcASZ56AK74mG+5u9QXXdFP4Dh5/obrbQ3ZACS1xIspNGMjNaVAhL6HvHbtZhWQfEQ56EXQBRoQAQGi+xgI3vN3dooL2Mwy90+f1IQSiEw7p//mF3zg7sq0usd1ijbced9fHNDi6Qrcn9dFEGscENPjb/1Dz/Hc21ihdo8FjvOx/zjVEkoCqIK1JAMVATR82dVq50J2CZQAEvPi9mqLYKfRMvIvVzz/YZNFhAABKaiH/Ot5Amf2Ahkq1qHrPlPwaFGt/B9X7y9GA0EE0IJc78vBnlSAIiJXg7mvFjuwlqUK0zyt7/b8K1xERBBQ0P70soeek7tcXlJRgEoR2wskJKHE3wI5Nt+8rCKIiM252p9+9E7WsWrIgRPv/DNvw2guiCAIIKngetX9j9AbQRxsyc9/hE6sBipFwi0pRkUDSLkNCG0AvOGubhwyIijPer/9VumAgCyYy3DVP95nQUj+AvH8x8iAJJI8HOUFT1rSAAIG2gfCvH88vZyrmWAJMVxf1JOdksdDYI9jDk8Anzybm52cbZ5c7rp7AtIAZgJ1dnnX2159dc7qLFfYXoVYeLrmQTdL1+NmbhQknB0oyPHrvu14FtNphauvZXby+A8gSIUhAVEtTgERApLB96TqHFBsRhUhBG+zPlXDXRnTUICHv+reWGIzSBQj7b1+4DyRNxMmNQgTSLlO3dnBKLQTe4/5k+ZCiAMVcq8TewfUmk2ZiAJUOnl3UWcBWnArj/u9cJcDXWxdt9Ul/Vz8xjRwzFAACc4GCRK2Bq8+g3Oom04gNg9gtNe8rWqeh0ipuoErf+KjJ7GZFBEwKmmsFh/z1MPFihAQxOb8ondSUx8VRAOMqwJWKGS6KbN5ZodCB20mtz74up2RwaZCGVXLDA/64ffdgedd2TewK2BJQmZqL+rpTScgCFVKO8fKyfvnW9pLiACSGLyV43wR834QkkzZwZG763FKyyuX89lxh2Dv5XrxPQ+Gh2LS0In/5vqu3WezN3c3xz2dl+tlYz3OZybO3g4dXF555V0BCAzpACyf84G33V5Bq12We4CAWeCrvEsNAhWkJPSsCCBBAGmZD6fn2BS0RREWQjyIPUwMHdRG8iV/1HeAKAjAqlbvwnkz+wgVKkDDrDf74hpmIDD2+POnfSi4iCACocpFCw7cX1YQFGIzyrS6db9cLrNZBn7wy11RBBFBZGhcunc8ClA3aJoaRSGy9bbrxpxEIgmqmpan7kEdW+6YIlJBQIRHu+kqyRz15M3LjSDAxfaI199zQzkgYntQBK4/+/h1XDYKgAKggMJRFzefGi2HSqRnPVoKrW8c+T7/ePf5nrzhuF/AQJAd7JxjdvZFvHYn9Ty4FbCR49jZ7x/3tAdAApgs3LO80LspDajEjD3jZBaYUzb22E6yI6whw0UyaOe84Xpu3Mmy2z0swTmdLrNHu3Tcc1YIJAk1LAynUe5WnMPrL9jPwAEBAoSYfmqPxtYQQJWr75k6tkXFFQFBYqtpMwdnpPzM9YIIQEBUpHL/86UcKxMrVFIowLTNcMiPu+WECwgiiiBw7/84QObgKEgorAhlDizDu/+jpSC2RtDRXW48nq33eRU2AsWRkPQ131E1IEECQjOfSTu2uq0oAURs7e2mK6fexuEUeMPdkoIoIgIignjgyxUONoq2RQKSP/kEJ2SwVeEcnr6pAVEEBLlMPenrS+ocVQHJ00e899e8MdgJkIc52wity3PbA3c8XQdApQKt19ny/nlp3QGoIY3gwos92xkiZFgtQZLmhLU25orZcCWdEEH2ANnM2HtOlq7b5uYCnLC2ngfndM+qniO4Q4AUXOm60JXOgvM5gPjhlZLAALAcvuoNypYExUjms9GxGQUgqkTKaKIxOLy58fX3QlGQiwjCmK48X8IlE3DQoltk04HhHdv8lx9pghiSgKB5eRUHO6O2YRFUlggHiwD3eMWSuFCEJWBdevDLj0359celKoAlbDl5yi9bJUuJhQWNryDHltun3ii2ugjzdNNVqI4AKDddZsqFC3BhpXB+9xM5rBG42BpE+NOP7VScbedSufI6sdWFCwuF57yjwnkdizw8DxAG9uCKPH887HjWtDLBiTABWt9HbmYk050U7AhfEDunw/CwAEMEZg0s4gTFdbBGyLCDgIG7U3buF7jfucdNjMCI9dzmdLb7ZQ8gJ0lA4HroyW7HehrPacC7PhwQShrEAhi/JG8RoALkL6IfX9WIiIgIaYDIoQLwvI/BSkERRUGwQDkvhvu9oGQAIc3M8JkHT3lKIUiziKKkeMghZkMQIAQw5MbRL3n1pUFsD6Dhxbs989jQHSdFRkhBUhG3jnm/dbeEFBSAL7tnPzZum5DYaikwpuuvRuHIbVzy5h0gwgUQuSBj/+JDjQrqTRGRC5f/6BPmlMvnDi67vhCRiwjoE9Yt97+Zc1npDYk82rFDQs2Jcpz2HK97XnDtVSlMaGrK5n305GgiGMpEll6M3TYSwDkeu7J6IIbAwvCwBQQyQRrYGRJr59rZ7nLTfcvOlcfjYbsre4XX5jokgCAJntxc967O7X4t19cL4Nd1OihAopLRM/hd5BAoAQT5qntz/B4UERYVFyMROhQK7IciQBBCxF4Szstwd7RgQ6XK7OPAzwL/NgSMliJCqODqg0KLIogoXFDBKEfZffWVDSCyxObau8enPPSFDpqEMqiWhKfRfxnFQQUQUPhOjv/mRRVARORioFuuUJQj0cbH//aQUhBFAEasdh/2gsMYBBGbUWD87qc4nKcX3+bCFbE9svCz3502ju84d5YDSIasQZrFnXjjc1MrFSdUKbKb9xFJ2gGYeLgx+CJiCSbhCISOMw93gFmAfeTRBgZ2QpMBJLWNtk7uNmtP3vB1dk6hApNHE9JYNs7z9CBaer2Hxw9cSwkIBKSPL6zX31rISgghQuLLzkFZRCBwyZKSowRYukWCKAABab98fkDO4zxEEEpDzqGfRdzqIioEshAUy4NQ3BSkCCpQo5lw1NOvvBwRBQEEOvXofzkWBXjaB/c4CVHJEfaXiScypQSExAzhxx5f3bhUYYnNIHrddFUb5WPAL31An9i6LXL19qv/6zCjckAQm/NTHxvO12lvGoWIkoqAhPXur3zBWY65fCyzVyJNEAdJ/+VcrgAmLroP3CeMQSBUSmi2+F4iIWSAJAK9CM9xAHd4KFl7wlFrvFG5J2B2wJIEkqG9ntvuXup06Y2xnLNyMDuIJIlAAGeYnfECP+bpKUgAIljL+VXMv0EQEkKEyOOzz0FKKAKlRiuUFONwwOIRRkAAooCgm/OznJByWwnAXMJnbv2yxQBERAoIgkq2qFEgIEQID88Ljqo89AWA2HQRxJo85FXHsqmbT1KSQAjJtJtuaHrTakzeUIQoGCfuf2y5ZUFxcERGbryryjmayKOeTZ8itkdJsa7brjhIsdgMgggX/U8+1ueLcsMVfQpicyOh5lrf8F0/fUzHrJfCJAlZhOsXn3c9iFhogdv75gNWQB6t0Nba95CTcBZESEh5xoudIwSGDECImuPEeMOhssMACMjDdM627ZRdWF70ItfkcUES4YbX2DiTF/vL9/oUQEoMlPiH/R1+BoaAAMgyXHL3YwtFBMhKAyuISo5w97sJiESEIkjpes7T4dO2UEEmtV4QP0vn/giNEpFcURRutdhaRmwtBFBmGiiHgx/8crYHsRn3utebysf0rs/sDYHYGqHf1GD+O3UiQEhDQ8WTj63duJC2RYA1t5uuLvDRWPT8x9uNxoERQIZ19+uVg1AUgWsDwt99wDhfyj//mb2JQ0YExmjTHb//5L87b+qJEkgKTDiXb/2bb99yy+XDXro7L8TbL9xx+9bz9uN6dMKjRSj07nC/hzzdhFklYUdaYrY31vHyDsnDgbDlpq7HyRs28Liex/BoQhIOiHe7ZzHLB3XkLb7eoyZ4Xq5euneyNzbL9/JpG45UAGt+e17rubf2MoQAKarMN6Bjk4WAwrgKt1BWtijbdgMoQoAiqKrzZV2/tLqoUAnCdD3QzwI/gV0IEVVEdXgJ2TYaQRtABGhRuyzWR7n+yo0IshEEZy82x/0HH+VAAgkB1nxrmHgSJSCB0NwQ79XqmOYbLooUEUEEzHXHpUyd4333f+wTEAQQbWTxvn9LeQNiCQji4Oc//CSex6TRzDQK2hg1WHTUy4f7oD8bLQVBG1sjiMKNf/nE50MZBVC2KW/MjjYD2AF2QG5+GmbI6ye5ieIjASrJ5bz4HnI3hsypghNwIMULnPtxJZNVIz1ZTuyNBAQlZEqaAnVwmZYzPujLZTkeCYQMkyMikje+fNzb724hQQCBeeq/w5j8z8tRgJSgQUp8XksdD9MQEOEqBEws2TnD9rDZ/pgxWQKXJQDBXyFyPpQHbFosQJFyhcxnuP+9HUYDUARhMi/TATWRYmskYCg25oiP/6E+EREBUQSj/fs7KseicGYpBQQSsvp09mbazLPSN0SILFTRO7t0LHXjTiaARAqCoTsuUThWRXsLIYAoIgh7+vb/zcGBiICwtJEf/prFpP1FpaJpZqB5qXWdONvmTKtDwaXXNyUVAURRBFikgKxe+AO/t8+BysZxytUVWRHIAHGXIMNHCJgoctuGKtSTmfcQaA9y2sGUrgeDvMjrdY8pcGeWNZw1jBcqjCGwAwFJIdYmH/SF63GyA7gDO1gQMxuwL+A4fxM3gQBhELSbb0adb5MMARE1oxr17viYXCTSaIpGI2pDuI1tW6eXXjqKikEUWw2/QDgfhWSolNRJQYFsPuOvkUUUESSCsHlF2G7IUG3JRrNYaZQ3lC3v/beDtgEugCipT/tVjld5n7/phSAGh2amZ0Cnzt6w6CKgiCKQ/lM4x5JbdudFlBQR4OrttkuPC/j3dxARJAVGcjI9/UMOUHqJKCIioZLyat139tsu6d4lbbBGzUqtF8s3v9erD8WL3waMBBEQERGiRIxW5DU/8OSzoLBVOYaTZwgwQSCsrs7ODiE1NS6SOEX2Mm0qwOl6eA+5pQPMOZuEupy88PNIVxqYBqa1g3jBV0jWmDQLEO43sP3gAWc2JJgIiJlOLS/Q+I4IoCFSFhE/AYF/mV2kEhCUgfbjHPdMoqQhXKC4lTStDvOwv7liSEQSgsgilfllnJ8VT5VxtCEd28Hsd1D4pHexUEEIUEBqtceBFVMQRWUigOVFaYNDXvXUd/VYAC5cQMTmVTcA5aPB3753UiCKlHDKX4aQ+UUSNEQECKi3X6Jj0Q0nI0CAC4vRbrmc4/9fv2RFKAJGAwbtVfc7AEYRjUYEuAiCUbgBFgKwQgF9ir78hw8oA2/7gtHYjAAXW0dLymUHFvDGb//tmwEUjnW4xMNMgDVCmCZMGpqdOXk4oXVVGiZid/fa7x+vcexQeYDsOMHii7DbbZwQQDLsskAvJCiYDEJ8wOBZuydvymckSEpYAwTnbvQi2o995+lgaEQUBH4GoPWXLRyQ2KqIPOjYetRQ4UIC5mWcs+3kWUA5cc+v+uirGS0glA1RCeifx3kCNBU0rUwU3LyjQnjoL1CeFyRF4SKCegnKAauKQAgUIKIHgQLUzgf/77cHsARBuLbl5Q/iuNuZJSljgUgq/eyLCBFPkpy0UBiFCL03OZ7rTnoCiIiIRurM6WNT7vnaUQIIAhIluv3SA8oJgqQAokQbcpWViBQiiZQo0D7vZw/Y/qa7WAggAoKCQgoXSZTRGtzwEz96A8cd4wHsAEl6rMMOIgSzDtGIzQLkNu2ixlB5/3xCrgoRI0TEvog8FRDYgdhBHcBeBLN7gUxACCCzTpx9k5ghgKVAy+ky8WJ/orgoEAVI2tlrgZhvphOCSYIkvHwfdCwhiYIqgkRLdPUbcNXcdk6c6AvSmTANErmIIkbji1DOlzWrrRQnFmnh2Gs/EuArv59eagwJoKJA40mI7aMEJBVFxCXY/fMbFqtLdKqmphSWXC6IcEUQEJ/1C1M/pge/KEGKTAwa2nkJ22+5WRIRiUDB1b+f480Np/qEJaLIVGatLj42uPi6HbcokYgAV/f6Em+DUSFF3FwQUARGYjMoAogAnPYRf3yo8iP+3SlQog2sFIDLEulNoTDFS57wa8d0AMvOgAiUUw4kMdAAs0KaCS0CNRX2xfeTBSYMGR5NTl6o+wRLkgFQuJpclhfqkhmAwHkQlc6yYW+GKRHcIcSEYabdeUE/lRykBEQg+hsaSDzdgSEkCWRK/UfJsQgQCBSBFEV3sxK5aSgjC6wCUFkgRBr7zydQ/hxs9jElRRoQqlceNrdVddVXfv4puiL65C2IAvhD6oA+jQIKEKiA5NL7I0ciaeuaYG5pEQeGjWmUj+l3Pm40I1FQaKjz/TQg8JRWFo4UA7T0By6P6aYdGhEREMTMHZcfXxs3XZZUUAQWiO46vbdNCQQFgQtGYwOCNg7rIvJYfPQfHqSA8k1PkCtiewSuQCoRMihlRgu55bt+6FgWkQkhREt5aAoggD5IM6ZilQm2lfAeeiIIkoCAcs++CLzgABKmCHY2Mi8km0Tk0QOEamZnduVNOWcCDBjgjgfPMvaF+CnvuLshgiQgWHxtFS1uZ1+8o2pBEAUhePAJdCwuRQIQCGJGZSaUU0ILq2A0gCKISPjb2TSfQ3kYU26lUIBv+1W/5OPz8vF3fPAx071OtB5VmZooDnSN9oozy/mAyRWXIkFSAK6RNjymkluG0jTKjc3Ixeb47cdx3CdvVpMbKAiiXvwBAii+DxcUUAUlpN3HHE+7/gRiewSM3m6+uuQcj/K6u4sIiICIjUv2yAaEKAKCiAAitka4gCAIwsSL9/ubg7Yvfv5TcQUUsTUIsFKjFLE9rBesH/cHxwC3AALIwynZySR5/YAJCBWwGIDOOtjvIbXEniAUIBGVHC0lqwgCARGERhz6sZB9FKwVKEI1NvvI5U0/Fx8BJOwAAkzFW+5u9toL6b+AHdkaUdjTizUwmC/7Swu5MNur6qufkOPwaSK2BgEF0JYuthdFlAIikBNlXtz8pC3npdyCJCSIcOkj/u7lAJGHlxiGHV4/kv+Y7XVekmnloUJywIUDCCgCoggggsjFaCznY3u3f5rVQEBEUtQte61vhFefOTVKKIogiuQnX3YsHkpIMGlueCt7iRcbf+kHkAkCO8AcfPjnDa+byE4NmOwICTs0ALaD5XScfP0fOQqLa69wIQMpEW3DFUFAWMJpuPiXD/qK2Td0RJIQghgMYsrr5gDJo5UqFYq5VN5DX1NpZotkW5IrCsdYq0YiEblAEJCrRpRjkVIAS0MB60wmdmc+F7NaMiEkKztkvODLNz0vBVEEEbR6sctlJTxztexVFUkJAuIv5VgbTlMEiAiSwkURbVFIgXCJrTXS9JGcv8tiw06BkGlPF8BMVglXIIFkmvf8l3PyqDkFBjvxMJkEIgiCIA4f+uLnPpdj/8d3N0LBJVANtacRtihP+/i4EUUQIZbzpSfPKkdrgYlk1dCIV174FxSRJAxA47CvRzQDswOkJCazIw/DQMiF608ciUvevLNuUm8CgVMIRgsigogoCrL6Yv8dd7zBAVdgRYCEhEyABIIEgVJSnSBg9c3Q95DpXMhJjFE2DsRxjtGbFJBFgISi5qg43gGnawAtwkTUdHeGz6VzBQFLPI9ZkMV9Qf/uTTFsEYIihycgTICzb75XMarihlxA46pTZ44jKltFBIGgIopcAqIgy0VU0Wi4CLQfe8Z5NNwc1yMnzGV2yIbU9nAZQBogyVyJn8rOY3GNc4Qd41EBeZ2tEUQwmivIyXTbFf3Ydm5f9IlIKUVYtHwN3kJ4/GuLIRKCCBqN7/kyjjFFQroDEzXEi/+iKSUkV1ftGa+/CgkNIYAQ4bQOaWaJJO95ypHaWL78nkO4XBQRUUgapja2u7BQ5wvf3RuIex46OzyUsD3CJIUkzGRjEat1UhRoeA+9RciaMxDsgAHlGFo1oiIqUhEquxFhKcchjmmoYAUrw65RPp/bJA+NPHI41hmP+xfza2AvMFCjBQV4GqCw+b0/2jBIIEFk6uu+tXw0YaGgCDdEEIICXAJFyFIUV1QgTX//pZzXuzaRxsDswU5HuB0wJ2SaILgyNP/4D0OPwYVCHCRJMh4fbUtSbI3Y6lrf8zqO/R3/xamCIBeUG/0NqMYGvOmWyxM5FTEKIvHJ6BiKMBFmh8BZPpjvCtvBYKGJDeyxTPdAyM0EdoB0AAEEifW84SltHIGp8+uPI+uF1kuBCxdBEBQBEVHkNJO3f97HvSEToMmacI8cJEAIcZIVZqKCQNMGaNS+f6yDoBxigS6QcHTRrGJrhMBlhYIyx5mhgNRO2glCY8pZPqeBhDvI49XsKS/4m9FQghqAadfdzCAAyq/2BgIhFCDwVZijp4vCRUjDBWJrRAEaDSgSKkpSuJ773pzPwzELoe2YHCkBTZkDUkp4Hmly9d/juc7lIo/LDgkCLsCFi62jHTTa/MDXcPx//gGI0UAUkJB/gQw2FX/rjwMtCBoR5XHFNW86hnEaJskQ7G5gL+yrI4oh1mDjlcfNwoNkhyFrmppMTHbIJDkWLr+iHGXzwf988eilKtcWSy6iCBcE4cJIefaxf+r7vJH1SHKHDJjC80AgFKA9HIBSASZQQtV0y/vnDRHRVgkRSDjGpEoukCUIgmYlrijHMQEboK5J2gYqtB0zn4szUzABMhEOYM4X8UmXq7OEIJKrzF9ycLj1NfdLhQiCkjTmk/d83TGgUaCIKBAJBERsTQOSkkDR0DzVb34S5/XJ7f3lelHwWICdZAC3AzFzApADqetx8/2+gOc+M9ghJF5/B4giiIgIKApKX+zf+zqOf9qHGs0qSKCMvp2DA7/SJwIQARE183VfgXIUuUeWJrgOC8QL/0RCIoNhhz2NHhjnhSR3hDQGcnYyoUQoIZWbn+cYRfSV33GSnhJpRElZAiKACCwUq45n3/sb/93nOljOy3nYgBiJO5OZlhAeGWBGKAg4o8AQ9/vHrsCmoPhA5nigjYkokiCCIkmkyTPHfK5UBAQrQGU77fD5nB1CAQLMQK9X7AX8OhuMAgqF5G89QAG+6TdQITaDBORHPgblKEmFzQIhNuOSSIogiKRIuLC04JN+E1DOH7ge2bADIg4BIUOyI4KAmZn6h/4kb+DGEwYEGJAEARcbREQRES5cL3/4WdU4tnd9htPcAJIiNPJPhwBuv/kqKYAiZGCM/wXhqNYmHmYMuPFBvHySBc0AxQBeX2UWCNhjEwZJdqI9CMi1GXYKkdh4+w8dB0wdvvjbr8R4iktBEWCJpNjqSsTMsz/7ifs8YnCk5wHiEg6eznaAsEpke2BKrWxThe4Wyvvn3QiOKBAYwx06ljrZKFAEYauSNZNSPiaDFSiAhVbW5UFrPwdGgTzMQWDn4FkYL/K7NMgUIgJGO/uKAwLKM8RobFUUp9E/iuM0cgkg2yI1ABURgIggqqiYX//oNyuc18MH9hAYwLWJARBAEGCHBIPS47/+YTy01znYC49m8lDALba62EwKIhj1859TRjm2v34fKzSCUCTG4tYzh+Prfp4UQYpSQBuXX3YzR9ZaDmA7QMxIL+zrvJ2ddUgwgeXZa8Pr3tSAPBQGwR0HCKYdYwqRA/7V7x0PLYZ7/MCHnQwCAVu2RlsigJzwVb75X5l9vRBBOlKI4dEjBhIaIM0DaidYY4WRbpJ5D3lpVzRAYJHglhzLPAIRApQt0dR6KeHolqGAlQJFGNycZdh8HiOFfPAwdNrjyfVm34BhfOOXGTIHEJDmt7/s5PL2S27c2Z8Wvee2f35UCAJFqKBFb/cCHSm7xfYoikC4HEpBURCIzaHIP/KNMwHlPDrZTRMymzwPQEjosUmTdphTvvt/w6PxunUEAiKQsAMiCCICF0QAefN3/tw6SiXH1G7fFUAQIpDFuP9VF+O6cXcU1dOe+cZrYqUgEcAUfvbjlaPMN5M8qiEwhr2ob31wMLyucHZ4bV/HNR7d4XX1tdsDjCkPcq+HgMD7/udbj+ngh37Sl10CGdVbAxdYIIjYftv6Jz4uXn8dghB7MFCTbB64sxs3u54741ECYMVKJse6m/Aeem8I7gboLobHhAuL7YogEmtaV42jFdYGiiCPSobZa0wznwe4rfYAyESSlfNozud7aD8bhCRZCIt60QJMsdW9WZaiQIQrteYHPqiOhEzbhiKAUfbUiUAobFeQ+MYfXi+7KXMeL0jCDtgkQ0jKw0zweljgeRz79f4hb9jORh7uAAKToNEiiILYjCJ4779n0YPCkcsbj/j3iAggsgDa65aTMyarRJ/2tEsZCCICZH8I4qiN5fFzTACBeLH202GnBmEnGa/z5by+e3HnwbBD5v17ftRnv3zvS8txz+V88grX9ebi7Ut1+bD/+ys59/d9jy95xAKXKYjiRhRhCYjax37ElzyHrYGFALNTcvepn3fbdn+8dJ7XublcXvM1L+eH/71XWwVsA5Bu75l5D7lIcWUgwoGGF+1ESkARVkQgIcIca5EWaMNYGK2udb8ZPp85yEMJdqTaYy8nzx/Q5Tuv7ghQEciNRTY6SuS2jAZKQAEK0fye9KM5RRCURWQJuWZnoJbqE0FsNo/6+rd7Qzcy57M9OSIxXj+DApogoFmJvR78q2/wlbzxbgEhmzUbaNfKBY0IIIoCcr3x7W7imJWNP3oMBNBoSpm45QSuUaneosw7FSKDXC5Gi8TJh/3HkcbAMuQB5Dp8ELv9hPOYEpIICj7vOVac80h2DHC5/It/DnCcvMFZkMw5U2D3fl/9QXcHK8lEBEEEkaTn5Uf8+udYxNzhPASYNfbV/6fl+Y8N+Nwv/UU6i1EDFA2FzPvHHSGXmEBgVng35eOoJlSWokSuUI4KEx0PBCsYJhNGsd37SNb+fNxtM4+wQwN07DBznM9nyL/7EcSsJshSGqNG08Qo0+QCKADhIsKV7L7zvxwtJVIAQrg8mBdZ4CpFTKBRAuTQL/2FDzKE84uNkdxJsYak2YEJm5o9iJ2L/Izfwos8TiUTk0FojwSiiENGuMB/8SHHtXW5V0RRUMCFGmgoqsrkllKhQCaQKylERV/1GUeazhSQ2IMdYl+Yfc2nkADCsbDT8j8/x8M8DxjkOjDHkw+73JzXuVwDd26uBbM313b25vS5U3WgLnnkj9/jFCMtULgiLAFIP5HndG5AZuc4x+Q8OOKlY29PT1DOWcfbc4cn76bDJEABgXZ4L/0yhqDQeHiwfDWOd3lZgEpEpEYgMpTj42qJbZiQklYiOdvy+dwUYAcBktm4P/fk+cP4eZzSLGKlSOFmYYGaFQVShAIaARfN/t73PQZCRRuGovB0Zqef2V0G+qQ+0XDRJ5qrf+DlN3G+y0w7DRNwHg07XC/JDrSyVlcvN+yf/MlfwgudrQnZkUByuOd9L7zLNN0bGG3baEkl2rv3DefiHZ7DNgVRbMqFKFe1pDEaoMZoqiCAMB479aPkPBbDnQ6Yhfgg/uFoGHYANDkv/MXnux4DUNNld9hnnNdZ7qDg+qouwGvK6T3no0JhoN7xVx+opFxEriiC5NiPfZ4FjgQ6SBgoj90t10WWznl22Sv37SVUJraVicjMe8iVsRoxIsLlUxDK0XZPKwIRFUBEG9qD1DGVSpFJqDAKksKw9ueCCdgZMEtAGDm254K4+W7XQ0RSESEq1IJbA6WSXigAEQYKl/Xu5SO1CHBBiiCG7rjrGRaVetp7VzIBKVezUXV+5/3Ou5jTTUDiYGjOg3C5HsF2cqM38C8/64+8dtzcvxBskASQHVx/5m+tlrb4qu/oUxQREUWmwi9/xvEpz3j0kBThClFAoyqyKVwVS6BYaRFKBK6ek/d65VFoMUOEdbg2L6yP+/rXC6uYOxkx7Ge/jgGXhTwHaDrF++GeUxbAYQ3w2Ig38TJaw+Xf+78WFqQAIleJTz76i15vznv22BkoYCdt5Tg7By0eytW8nlLHvQDTHakLMu8fX9GCAGmAIN+NQTj65TujRTAaRCDBWN6MFI5rM0J4q8KQpqSbz+UBrEAmA2QLsPGGP/Xm7nqTuIYEAlCkInIBmlyWMAJBoso0L9/h348isz0ICLFu2Tu5XM/Dj3vtwkCkFKFCyu/7wJedb7A6GZYCyMFh6YXs4OF7P/uz/sZ74jJ3vNBnh5AyEMiyx5PLS6+tAF74kPWiIiAiwtKeHvoq5Zhgf5FG5IoUENCMQhEJkCAqMRqACiA18w2fexQfgAkJAiwv/i+dNMujE8Fwzpe8MvtIPBywcRAYmO1cWR6Nk8evMvumKFMhbbD17i/fdUEUBDDE8m5eX0UGEoUJ2Mv1uJeB5XXXou5o0EUADNJuKO+f7/MIWKKGRy+8i+N9u6UQ0CAiUkC8gmkexwQHaQWKCBQH23Sdnwc7qGFAXn+74TrGG/+917kkOQCmsGRAAcsFECRSZBtkTOGXHnaURLhUQYAK5N2Tmeks3vQFTw5EIDbmpeLxJw8635a72CPBIWEHeu/NDRN3N929+r7P/vP/4B+8D5iZq4C9sQuIQIjEdH36rsuVFZrmt32BkaItBA319tx3Qjmmx/36vEgqgghlI5WoMCESkRVAgdEISKlp/QkctVZAkESqFmZfyC/+6vc3MICQPOrxJ1le3xpCHqatex47x24PZk8el1jetJWgQEV+57+4xAhBUBDO4ZOfo40AJCEsL2FLmT02J4C3TbECFIVOm+N8D/lyQhxUQVrh6cP+43im9ZKAiEAgeuPPWSkix2FbEapUoNCUD39mUj6P8RZ5/WrYcfb+9hqXs+eS+CiPIPNBESqUCxQVAuGGInBFRBRuLHnozmpDOYAqQqQALiwtTt6xHEmZn7/hyj4JgmiwSCbGA9/r79s4r47zVm5YSBCWmlc+uhvq8uwZz71cCYg3fjhQA0JC3fDas8sdcuF/e7iQOPxob/8fHPe/v31UitgUAiJUgYrYmmKjIgoESGg6/UF/foRMPEwEUi6vwfLQAAPs5v6rfBY3AJkgwOrwi3h9Zx0QIDE8dj7q8O7eeLi8/vJva01/9j5jIiICsuNl+YznwJeXJDEUDPBZBhGPnzycVywWJoDA5ODbZngP/Qc4Ai1AIdD+5kNRjvalrkhREC4iaOJfaDPhWLN1gFoRwMLMejhs7ecAtn0kGQYG6Ogw4rkbPuUjrkkEJCIBGFzIUBHIlBAFWC6CIPrUJ5chHDjFFAIZFVREIC5c/t7vJSAZl6UUhKdeOlrN59EirpOGEFNdX5tnZy6vO/vgg+lyHk4CLsKcuHt0Nu/++owmIiII2rjxSlCOYzpbzQiBqYBAmFKQog1VhCwRKYIgpeDHHnSEafKB9ECcTo7tQQBhs8f9zee4HTxniDX/z/ueIxrWRywxzsvd3e2M5yP/n7CNJ3zzXA0iNoc8by7PMd3PvWkISdrOSrzIq86qAEWAzPiG3feQz8kTQFgIdpbmba68IRz5k69xidGiuIoEMiZeT3Rcm2K3ARgbRkS9lpbPpdID2QGSAIIrb3D5bY1pBkKykUqC03AAVMMSEYAqUkzZT7ibOfQo7AJXGYuQDNyGSOr7bz+lWFJE2QWeuOSxv+2Z81g6TsVEzsNAjjkczPvHlg/29VgDkvYgDG5prQD8w3sMChQRREi53ufvpnEsn/KrUYShcBQioJxILkYJiKzR4iJAgIgo9zl92+GydI6AgaTAyfPHcn7EF7NzrlJFg4DxH/K8JyJhiQ3J3tK14U0tcg52Vp/6CyoEURSE4ulzwBmYFLLjqqfxgpdsC2KHDIkA9v3jn72f3QVmpYFZUT3nniiHUi7/KVJQo4UCXLic9X6NcMxrrwJUISM7EoZJK5/P8xwzk5WHZjRers+n+y3ZSFd2oEbJRZKiQioE3JqDC0gRBUVJu+upM4cjKSIKahQpOiFdDJI86TuhJRVF5cjFuv3qTto4j4bLdaDmPOioOQ/2SVaGPfLBP55JhImJyG5jsuXSG6nR2OpyQdTPXiL5SNXml93PVcAouYIsNpViFGqOIpAVN5ArArA0qr3fHxxOm8hDzwGwK4/aIx5X4Lf+VBhEmNWGWMxX/ulzCeAO2vDQjifLce6b5oM/80f/ZQ2IHFcb3/ataVHEKFngDv/sOXYOJszZoeF1lV7ISCq33RekWHB4D/1bXmZYFolgwl3f7ae/IOXD5O7/eXI9xVID4aJBBf0zCVM/nqHoKjAh7WKvJucWPy9XkSQmoUTP2/O8nLzB/apvQSB1p0iNFK6yAIMscWAiQcSYKrLC5/3g4TIqUtAoSpGtgCsDCN+7t5RTJAQJ0du8+8OPN+f1zZE1HAQMR4xdldZ4k7owCOQA7ABcMdv/7r16I8KNIFxkXn7193OcV9xgl1JERImKiCASEVEkJFKlIBIQN5SUxje/0+FqVbI9OAKv2GPxaFf4an/zXQksIQokuufxk7DnyGUQkJSV6164N3uTXPnmIv/wjf+64hxe9aLL3IQliNhZ4AufwwNH2APTmoFsjRd7XYvbkkuViZXy/jn8CiCGGTKKKRbj8379C27lsB/360sXRQARiKIQvh6gc7zr3AuBSah2DEDaSuZzMXOFUjJ3gIqjtOc6rr+YBgQhEYQIUliSosIlKAIIolAElIyvvvvhSFAEpYigabhRg61t/ob/E0RAQFAmpvFlu/vn17Fcj4GwlbSFYCHerN2AtIdBwrA3B3bQldcnlRQR4AJGrS4aRyt//1f2abQIAlIgGpWKABmIILiMQChyNVDhrB+2XB9qXAjMSawZvACXK7OPcPsd/8t3ntcLiTXsUEi0874/TLyusUPKMkMi6LPjOBGwF6Uc4st+uJMFq1/54RfmuO76V28TKSIFgQHl/2L2sTmXPc45QMABYjVmX8RpGtogWJC5CGu/dyz/2wc8XdoSQRrHPvPYl3ztU2/Zpgf+8X1apGQLqCKImMa/IdU4pg1OkFAKBFrkArvD5/JYCDGyARJB4vmv/CBweXQWWUTAUIiiJEEUEAoFCKpEhjZfddEdh8NCgCxQRBSGCCiDn3zxQ8ekURBQECnz+x+qnD9LYIDgEGzHUsub+SyBAyQRigXHAfzbI/vEZkQEcc2Lb/muo8Eb7jZabyCiiIyWSB4FqTDaoES5nEzgQhYhBdKiffrPHkpAAAY1zP288u4v36DLZd7ykd/957ybvd8biwZgQAGJ46ddec5gsoSBnWywLzyvywLxJvyubzRRr4nX/MovvdLHce9/uTwl5EoqIJg9+ByWx1PiKHd4dEePM4gXWRowCEyA0kHePz36PQQ5iAgIpe32K5987Ru+/UVul37gZzxI1WtCqcIFhEhA+/vVpM40cixlEYBWU6U8utHy+TxnckgJIY0DyZ7v+JSXr5fzuELIoCISRuWqBAhFhsqNRCKCgFRJvPOdX36oKapsUIAUyTOTAQLT+kNfh0u4BIhYyfwh930V56+8usdEAoQ04525Zm+aNkAeykPlbm7v+0EPfOkoRZYiIElqXpxcHakuuXEwmYoiQKUU5QIokCiFIZHCqIgKREXI0jfe91BTGpgmk1z86L//xe99+hG+fHvT5ZZ91u3cIIO5OwAPTjv+zR/kDQ4NMMAgQ/St/8cP/wpYL8u14y3PPM5r59vPGzyP8wHXX/MrOcQff3ifElye4JbX/97vvWwvh9DbPe4zr2GkCiACiAXv+bJ5jl0ExCRkcspeEJXJ+Eh4dJrjtO8bnfx8NmJdCImqwU64F08nKYhYQgsRFYHASkU/w4gGxx6AakqlWoGTXuvnhQAEhAQk2JbnP/8MJRewBHCFMSv0kryMibxUpdjfJUREAiIqjM8B2jigt0IAERHCKRPM1s7rr70mkhSIQA1I/vWK8yhO8EBIDAxEwrA3yzH2Oq97wT2vHPLv3qO3hovDjvb0D5n6EfixL15P0mgQEaG5StR+o0tM1OiDqgXgVOFKRBACqnKfS28pH1SXWRNBShA+8hsQQpBzK1DuIAOQCQ58U97gAULCDiTwlF/jNO3lGpeT83LdUV0RWvTxBb9UBhG44fSkiCgpCFrftv7Pf7yWy2+8y9u83YNPQEQARWzPubv92894/bw5AVPABSiPy/UFXY4Swm1BkOOgLe+h9//T7R1ks2iATqIgg0YZBUtEQCTANRpK8nvUEMe9dq1QcQKCpc2O08/JpSShRDFpOlje4NNPvF6UhxKGcj3kxRzcBptVdi33/+gxMSAIKIKdfuo+r2ZwYCkcLABRC2rmYH3YcxMoLLGZFLn8Y3/vvDF0FlgFSKwrBEG8SW+mVR4mkJyXg+Xg8kW3uyA6KErQqb02jvCmK2lRREBAXzAuOdMYbEpmU8UA+M+HOAIQ22vm03/YHDzt7oWHqwNwMoTsPBiWh9Lw+I7IeQz8uH89+zwWEgJTYwLnhcbD60gX3ZuBvew2Lje4/eqnceDOmUbEgUZyccgYFaMiBK4t0sGfY/Z1PI9ZHiYwBALFi7TKBKATRx69dsp76fd+NmEygBCkApBGRCqigovNKBqNshj1B3sa+Ng2jCMW0pu9FF0tn5ecGsBJSAm88oa/4+2yw3kASMjllzPZKOUyW20p+3zCzTuFpUgEiCi+77HLtbINaowWRRCUROmYQ+Z511+OIBWiEFWc/Mop5TxJ9ghAyGRn6WyWN/XKuCSwk8IQmANQ/va9XQDRlj6NNuqvPpAj3v31ToMICxSa601nGGyP2ZpRbH7OPxAUBBGQLMc33qWNg4K0BzsDO+SImFPazpCQPJpMTR3A//Ofsjxv7PUWCcHhoc2pEkdwOVvwdM6ZnZ0W2p885qCr3uyShXDhipwWdZS0IKsEmGIzCsqF38fyutfioVBDuu7sZi8iDssouCCMKEX7XvLeP3Y5pykezTAKFSg3i025IEij5BYp6qlPBZRjy5g6Qah7wQJlprr25+Msh0fLHQijozfyn8eaB7BDArypL2dQcBnKUIprwN4Nd48kRKTREBQfxVocKIaCUGQKVXpr5DDwqU/vAREXoiBtjBNP+ro2zg9krwfsWO7ANOtBx/mm4goNAgo7QEMzDoLL37SAoD5FBCUVsbs6wo9/0cgUAbiIcPgxppFtCAKIKMBzRwUBkSJcoq48fZtyQLF2sAM05A7XCwnu5ABBIjsgmTDJe78OLzJpkmSHhGPHHXSPnDw63GPnaBhx/eYnKdse9WwLRRFELra6iPqEpYjIxfYoxPAXXpl6nWPZAcgdEuY8NsheAFxoRoBai2BDs99L+FHXI0C2BRCChFvSgFRcCCJDySJQ6bu/cbZMlGOycC4We7FX4TyoQEez+XzGkI8MDWktnL6B2w8L53phB0lc5tfohE0DAQwMgB9/QrmCQlQkEn3xDs/NQelF4UIUmynqCPrzmy4bmKqAazSUNviKNpTz5TJGgEkCdYMnb+5EdkjEAfI1Wg6pPP2D2B4UYcrFn374EV53dzdGic2IruLJDHFgEEBAAGdf/cDeQipsH1P4um8KB+Z6nMMOO0hEDkhwYO4kwA4DsMOwcr184J28wANBgmjIMF1zZ3IAcWVw2GFm8fsfy1blDz4qKSBiNLa6TI1GhEnDxVYXEDTgD4B43fM4j5LzyBp29tjzMOKFnns1VCZYQn1keC+d/U5Hiy1aw8NIyFUuiUggohTREBC0Wu59CpIclGNBRqxOhMWxlxC0na79uRjDVR4NwInljXy326tzzjogZlH8SAWIwubGIX8yCogUCCVQfOvHcliNRgFEJCqncfjwcX89DSkSyC2I4OlXPjmcrze1CAJIUhFvcgVMw9gDTmDNoa56s1sQELHVGtoZhzp9QyEqKSISqbS6nvggyAaQDb7294oIBApiYrTP5ZC57IJguOjQkCEQTPJweDSB3ONmv+raGzAWHrgjQqshNAzIDjQBCbjjvnz6h8ob5Re87RAFAUUEFAHEKiBWEW1ERESc8ydeYXl9r/edlxRhUoZxy+xF5AQ8l0UpMsc9Y99H3ON//8NzbDigEEhRFFLIBbJAiEADiniqd3UNDITjLXN00oCFGqCVu29ayufS7QDzEYHc4eSN/mISKR4mDN32+kJAOGq5jVvecG8HqxgtogFDHwbKtmIzggRUMEqk8EHib2+6DJSIqEBRxPoTL7n1PFmYJmuWx3OGN/u2AwKSHYC8xLGHAf7+PecFERBFAbFqv/YZyiF+5ItHKajYsERf/AGVHOaQlcA/rRYUhYIQm/3Ku7y5jQPkgD2QECFDhAQBEnk0jKTh/Jgvn0V6nsDcAQR2UCAZHgoDiAPyYLD52w9R34B/e+RogAsgIiJysTWIzY0oCIjub97/g4DjfJ3xXJEJQSRzTiBeqAPMARQQaFV5H42TH/KXGBAEkAUiCkUpZBVAhCACsNpnPH9CAVCOpWFhEIqDCKWNJFc/H2kOEiIJKes+39u+cc5OEwkCXvwFkzvHaYbyBU+TGwotQBRVdt7huWG7BESAKmKzhxZzcMTH/h1GiQTEomL7dz/gPJH73AtM+kAcI3tTHas8TESAOLl5doRr3ghRBLiCIHa/dK0cdOtFo6WZIooAwtez7P1oCgS4/tp7jYmkBGKr+dqvGmzXlXYOWEXYEQHCHUmQR3OHihy+4Gu8NovE8zswgCCTkg6P9gCEhB6hFP/lh05xgPu/WEICcEURJBUEBAFEHNGX/XoncPK6dXucQC5AKA57tC/mIMVwW6HSHsvTvoc8+m2+uPs924A2hYAgRUTikJFc0dzqa35pkSE2w/EOFiwiKTfWHWYP/Xw8MROQhxIRN9lzfduDo9lhkARY8NN0cdzhOdAoUgiiiNZ54oe2sS2NCi4OtMZiRcIhFf7+thMFShSlIFKdyPvf/xXnx3F/2cMkAwjiYbyp73InQR6P4X3cxxFf+NCkiIgbW+3pNx4HUEa56jo2IyBic3nzq9nnGANm87t/roMSsZnUwnwWB+9dZgVSSIbXFYYEIRBk0r3secxv/6mwEG9UBXZ4VEAgIYQdHsoOD1Pj6a8/4iwo8Khn1hBFkCUO6SLikKaCy6C0Pv1f8vxys/FwSAR28OYqL3TeilCZILXW2Gta3ls/6ktur5O0NYgpSwkBIVyAIAKNWi/0v/8PjrzluO2EqcVSGwDlehfMfD5OBMIQYGVWDuL1Z/nJIAxQyuN/g0b5mODNt1zaJygsIRTR8h7LNdtlU1QUSEQx1cAcNsDH/JUxBUJGgOiLv7nX+XEPrhKUYFjxph9KqAESF06eWwHe/nkMVSIUQBCXpgEowOf9NC4iNiMi8Ysox3GwnsIENCJcqKwalz70Pw+oNgBZJcnH2ImQHeR5ux437/8Wf58XOlvUDEBC7hAPBVYyE5NEkHnxRx83jJA/8+dMiwQR4IrAVATg2pJQwBBiTPy8/56H9jpXXp19kAFsM7H35IuYPe0EkCEidmdE3mM/6q8+Oe8WI4hKIAiCcBW4AAVhT/rin0AB5Vw0MrUWecfJ+YYyfD6nTg4BcnUy4I7XNRb/PShpcEpiXr60Y8zx/9g3O6SouEaLXHDqXf4x2zwTS4qgYhEIR//rmy9FDpAqpxiqhe959ze08SY4ziMkO4caMvd8051HPHTICEYu9DzbX33vIbfILSLCshff9h0c+LIHQFIkFQ0VDj/AOT57+8WjBSkickkDPu/LD+iuD3YGkMhSGNghjSZkpTuOo9/wC67MvoilcNY0SJkdc3UHYCehdEfCYNSffLQcgLu+atEXihtEHHKUIgjIaU4hhsTIgnf+P3k0njMGkGCnyVjWehHrUpkgoU3Rg4c1fZ/hW/6oyw17HjsoIVKfQG5QgEU0mpLe6o73eg6KKXFOR1x70ZRMHuN8cUDX/nxcmz0AJAYQWHy9AL7GE3YmJ8HzQIonUqZ8bPp+miiC5OaiYtX3P4oDJQEICOUSiNI4Av/rj1xFKkUqUIDHs+45+OArMmTqOmmzc4u9udD2SIBkJbvF51C23PflKEVAuIh6ES2HsnHZTQARUcRoQdl7EwHl2Gp82w+MIFwuXBGCz+NAzRgwQBKCysMOEmmSnTnZ4+i//Flfgi0vdg5odkAQYMjzqIFyktyBAQzB9TufaLZ//G+0WSVtc0W4iIiCIiAiAlwZ0/WPeAMYb/C1U2CHZnaCZs6Z8zhfADdPBrsYKs2kbThH3m9/30f+g8M9DXBR0IjUCAiKAAk0Pe3jz0qGMudSymZJw0gmY1OP+zeE8vm8nz0AdoCdGlTe6O+AIpAdjhrwb2Ewx55b7rjIBWOCoghqI++0s9o2IooIV5AhzRMWR/3jF70NRAQXYAE17vboZ6oPWnHcrO40CKHtDPHm3pMZhFVAJJ+wz3Hwq+4zKmmm2BoyFt/zDW1A+UOeZqgIiAgC/VNXIBz74KeyiEwhLNwo+8Tdrt3WcodkB0ES3GGHKVlhAXaZm/vf8ts+l5v77AUtBMN58GiSHDSkQ4QMOyGZDNXvPJaD7/WCi3FSVMQRXWwGRCBDE3/00Qbi+fXZAAwO50GT7XET8catsBRLCWLdvBjed7/kG32H990c7fVMQYkgAkUcKKt4xQe9kkUPYCQfX0PTAcZVsAquvTLVz8kFgR1IWAcoer1ZeOt3gCMnQBLXuvmsAuVjU37984wzRQqkgKHFR/3WNiYSCYQbSrAnpn44pY2PeFFLixMiWaO51OrX7xs++PJS2wWEZEBmxd5cHrJDJsQesCy9gAe8LEqkyEVSENyuuNUBnv2oXgVGAohG6QlMM+eytPeKB4xSiCgjIPAdn6NscDLu7JjQHuzA7EySpFV3l44LX/BL/vyXMl4hXvAAlgdls3MeAcy6oxgQeB6lmJi/+qAyKLCYT3/hV18Ja4omrNRokBQREFCUMFrxose+UOGNuy95DsIOR0oTmzT7huIKWIsAZZQhkHm/gf/lw/7DrzguuFKCAAUEEHCLGq//tL/VjjsCkZjjj1RkqgNQm9Lo2sznJE5lc4eH1RR3r7fAqydgkFnFXur3EOdUfPu1U8polEvFaEyDH7nmADKmQYkkGk01EKN8qDB4xVP+l12pSK6MKUU87n3fV70Jlq/ovJxuE7HorhBv7mfXzks7O611KeNg9ggK3HzpkJSEFhC45sW3PoHN664icYtLkKDUeAHmnFp8y1OiNBehAn3CfPLU2dy7bQ82KOpAsz3AjfPCCXoB3vPf/4p/+oyDkw/m7lSsjLies+60xzCck8gme4F2yoif9Z4boFCm7v6Fn301kYu4AZaCiIiikHJx5u8//3VAG28s3ta6UQUZbc4Kx/lGrOtLSkuxMKtMecXh9T1H+Avv/Jp/7mRQWUqCIghS1OL+D1/1fC6eQ+sGwjmd8ObFYlidlEKtUx9odz8n//RTHbnQkBNMzu2zfb2HH33sUDJ7HmN6XKcfoM2Ujw/e+Or7yDEaTRAlteDqNjbKL36nooWgUiacvryD4sjK577+itBQIA0YbUh69K+/CY77ey9ODUOLMpfuebPf/6tvOHYAHgDChfdx3B9h83G/7oIhsVWB5OarylB37JQJJQsFWXn1TYxzA/zJmVO9KFMgRkstxol3+/st45W3zw7ipMcq7hxhDrBHcX3vH/8//soXxOXl+04+uMuQAwl2BEMHIeMAOc2xMORs1RdP/xC2l6kEuOYTPuf+O6VEUSqKACKIgPHqr/qHmzjuOb9MmKVJO53ZA54dW2tvIOaCQqZM2jXC4utk854bwD/+nm//UZ9/8nKkdTN9kVQzUJz9x697voHblYv7GHE4x8P/4D/20TmBak8Zr3cr3/xPf5OTz+tP/SMft8zBs5vj3AucBtef9iXP53ve/d6X5qTRK3GdvuAZL2YGc/yGD/qza8YkFFeqjTCSF0+rDfPk765IkzJCGmNXN30642ihv93vPYhdjTSDWHhJ8oqn8Sa853f/qf/g6V45l8N1r3F7/t433Ws//488OfOmc+u86O15/57/kpOjl//gPdvU6G5jTEyWSP/BbwfQ8x6WM7bcoiUe0e3PfTybyjkQt//zo/enmg3VvIyVWtz8Qrb/he97f2XwA3PksXsD5Agn5/n+z/8r/+pv/tP3njx6fYUP+p/9rsczd+boTIJj267j7ZVZZICzGwTPlOX+qx5z7QGHny56h296u8uXENkVMRYAXt/69O99xT7nML7jX3vX072/zZqmRY+7v3vlxe4f+Nd+dIJrT6d76cvwE/9lrrwn3/2BT2hXfMBnPPT0lMUO6t6/PW/4p595+T5QhcutD8ech/9VPmuGz/vf+BoX78ebuzUMA4zn/7yPOOZK4hrYLOfly+4PLBMXmfoAyRz801eo9WYlHFjmWH3deyCZA8tAmfNzvx//n/AvfvjU7BGuc53C2BewqWl40YlanzqHV+scbziHhvdbrlE4rAgH/hCOkxdq/Fv5/SEeziww6yztLLPonHMCTGOUy7neveqd3/bel9/tit3bl+0NZ17zome98obBOV7+zSdiRxBz8lDihf9YjtNaPmvflzgu47pf/yBozTjTDAqbC9OlxCOd89Ri32KHW/s5sissr0EQxAPe4OwpEZERy5t4DXRYCYg5dDoDc0gj+RgogytkiwHM+XqcYG9g9k0HC8sJsFcWiBdbpgwKSJE3lC0KbRzTudYasalsC4c+eWjPZ9i/FfFQ6gpwBQkWFtor5yywnDxcUI5LCtRQ2N4Gh1XOARJx8vDEAOSDeVLFviWDOJn3pPNajRj3pDTT5BJKeqkSyR7KedIMmI7NaJuhrP25yXg4CyYWxOxzLa8bzDkLHOPcKVvKAEIigHIIBZTyhhBJzHEagikUBdUoA7UfPOM8ltc349/SWQADwwbPF6MAiiAqZBKOV9konxeBaCOUN0A5wHg0HpeApvi3OOhYV5A1gInAaRZkNgPh2KAMNPUyokw5MPWNcxsPJTAwaJl9YVIYmgGUkmF4f9aUmQxTBK5yQhQjJVbn/B1KN5ShAwUYPr/ZgwUiFh4sb7RHgJMFWM592Go2g9kaDhkgmM1gzrURAQYGMB/8gJPnjEf334blYUAQywdX4WBtKEc777MBZns4MN5oPLr8W3+yBAvxcHnYycM4QSDOeVlKyqLYOH8DCOLx5YUXKDAAHWCQzvuSIENXVcbURiBII0IR5k79SD9HxP87VgwQzl/jjfbg/xOafTAOX+b/CY03urxZFUCBMijny7/FFsv7cxGoxJF6KxwqgWBcJncOe4Plz4P91xPO+/j/1PH/A9tzxb/NCii0ofAWv1Der5WqTsVFFLVhzHZzZy2Plj8P+qZR3lIo/N9k/zYo/3cV/59QAVDY+l/B4/Y9KrijoRYlSali1dj4v+E4X8NbyvA/Jcv7doghRGEIwlvhBzABovit8Q4d/v/+//9saX8O7P//Z/9/tF5m1s5+p/j/suU7T/m/reL/JdqbZ/ZDLmv6GFc42fz53Nx5zVte+xDBm196kyjnWbx5lw+5boAPvvAdf8EXv8In3/ipr/3QT/ThsdkXNPtBsn+Ljrl8rhQ2y29xeLPPvrma8xZHeUtjbxqbfZ543fE52lS2KYdSDpIeGG9i40Ovv+hn//J/2z/xOy+zeNf5oR/7X//AP9vlBa8Bl/OQafacBaRpWYh/6xeFcCxZ1BzOddhqaJUyvaKNvbDsB80Aj6VAQ/axfQG3GV6XMfaNLbczDbx6mY4tTD8uBWpaCwTJnWBqosaqRWVDqaJ43PnmBsNnM1PdvMYc5/lmiIXb6+XiHCd7vga6LMd+sZszoGeSy6w5qk6kZV7klWFu7HrHC7/lYO7j9pzq7Mn9fXyI1X/if/fXvVrAOT0zoUy2+ov6f/Lf/O0/8j9/0Yt592/86s/m+uyjvc5BvnrD9ehkhmdP59x3vfb+L/xX//f//JVj787wgO+7+2pZWU8zUoUTQ4uX6tQzv3Oci6369Ld7v3iuMXHbojItOq+9fH/5mFf//k/60g9OP+AHfNzZcdxz+YpbjuO86Z97PT/m/b/hr7yRl//I1+rufhy+8jifvfS2vugP/vHz+T7t98/dM82vvDlWT05M0/J1759jQfn4L7vLDbO9w7zzWu+euOYlX3zT+fPu33HV3q2rymWMtLm1TG2uJPMuLJZizhtf9tx/fen6zfdhv/Wrn+99Pzdf5bXpNnw2577r6f/1K//eB423/9Afu6++xkefR259YPf4sJt/9dv/4Lje7q/O3tHmW3cWMMbJ1tuNl33iiw71nr++vnWM2dN9y9Mn95f98f/Xi3n5z33kucecT3au3IyX63H9wA/6xx9S+cf/Z/65dx0f9mWcHpppYpy5Dse7vssfvP/CX/knv/KNffz/89a8NrM0Ycmcs0foCcGB+wov/JpnjPNNz7pilAK2GhEM1RzvfOOTztmTviZYRKmMilwcm+73/e++6Qfne/whh3KdvMqOE+7xP37X/+MN/IVvT+bZZVdmD/7Tb/83nuvpn/gqVOd4csAy5RrTCx52PNzltYverBBNvS9rXrz8gefPH11GZJdGiUSyDClSSlKyir66+S/+jn/8ypvq8vc+YYe9zhGwTHsMfd//9eNe+6D9yN98Tuux4olzNnMef+rDn35Mv/ApCydWVVKYNvPG++QQi9fcFeJ4EuHk2D/xKS/knf/wowvPDmmK2Tj+xsd/4EMn3/+3/aJe9+ghs85VLAdMepzHjnvl+ITfzV/7QZ/XG/iNbz3nXO4vO7VzssdOu9edPbLymcxTvePf7//7p71mnD+F3+mKdaUPoXjWmKwwiP0dO8o5UODjPU+uKBrBQvGz9LLn1/qWf+VFmfF79Rpa95zHOXBex2Of+d1/7hv4Bm1A3Nl0du3lT+a5P/qrXJlTOtl2cBUXi3d/xvE8+h/s5nJB+lhA7V3m82bHQ5guF4qGRkXRkCWRkdFS7Nzjx/yd82//7L/yJvphn7D3l+U4z529xLLnuH7YH/2eH7TvcV6JbrOBM/eS12P/stXx/NEHTgmBlqRZsed73F4+4JI3V0BRjlPGdl+efQHH3/iYq5vMeZxznbhy7ek7XuZDoQuPk0/48gosBJHSMgVVykWl0Why7ey0au3dXnv2Q47wSCu1WO5Oy2nRFm2x2JmWbTEtFzuLnWlq02JnubPcadNcePotPvcrv9GbZ5df0XHhcnt7c3Nze3N5etxebm+ePn3p5b39U/lBCPgoj7k5Lsc4Nzc3l+NyuX3y9MmNl3nLLwR8IQEf+67Zm9uby+Vy3Ny+dHu53F5uXn76ZG5vn343PuNHex1H1uVyXO7WsY7j2C/f7dc/rJVj3V3uLi+Pu+NyHJebJzeX83++fTFfhblcbi+3l8vl8uStt09ub146eLNeLnPYLa+fhEASssOjG03v8o8vfXuo8+SrUzttsWiLtpyWtZimy5Obw2P69Ld80D51jidPnzw9Lscxc1xubp4cevDSF/Biv+TW4+bm5ub2uNzcHjfHDS+97X0sr/u1j8vt5eZye9zO7c3NzMG88ykv8O9/8nk5bm5vbi43c3u5mctxuXnppbm/+xCIMvN133JqzIshQYQrinCJAnBVQikGo+z8wZdCHWJZKZFAkhAMECBBkhIphVXXl/7uZ7/lzQJ8W2EgEMsMbOS7Oh+Eh53j1oJsQHCuId8W6YXA8B3m/kICynYAtMjy7s+yj5lKSae2Zc27rWUHC7YAQnQ+/bkv5n1HQwApsHD7ppE1SXZCBCRzx2A4R9Ik3Q943ssWPmfamN5GgySIAIiK5Lj5xA/ajXIii5hCJeeHf/oD3xBUyKMBT1ie97ywIw2e0TLd7RsQ+KNf++6obUsXCNvl137l8aGP8JA/v4fNRAsWkSJECldEJAthhCJ7ZPkdl+BDnCW2EKAoUIHEEBIMYBEeHufd137PT3jTfMST+zlDCs4xVmM4zrd8yvlBmkJSYrYQDpHt6ddLXqzLf8xlMXJjCEhz6f6zuAmM1HQMbd+tJYgdvCls2/SZH/VCgLCIWLL1TRMmuTOLmUASE0VzALYlT0m/701ve86y8TFFGxAsJQOiBO7+7A/asTQuE2cuZWb9kQe9odeQTgIi7MobvDqVuxyA28YbDL7FD+iS+BAmFq50/L1fyvmhD77jm9IbuIUUUVCUFNZoRFAuXCGWoho+cTmHHRHF1lEWQUYqKCQXyEqEnnTc7M3v+mtvDvlJO4wnwjbkDkxU/HQ+WA6IJE0U5GozP51ezNCTj12mWFEQOgdZxM+y4jgptAHIfPpuLzrNtAiZFnCSjv/uBcgTGjJMh/bc65uGXVNmZ8IEkwZs1oFSDmguSaf+7ePP1dYnQAopBYpskU3r9TM+aLGxU3bAIOcYc77jW/MiXz6FAytXgnquJwI7MMCped4vgq93w1v+F66CrgG40sHxpf8Ow4c89UufZjWFcqEE3FBUuGgRiSgo4ipSMtEbDxUQgaCGUKBkxSAiRwjJC5fSuV6/2b94UzQ/Az1r0dHkIAgHvu8HRyxEdlgfJRgYzu+OL2bh2z4NkgMoUA5OnFmPz9JpWBYQGcB3kRCKlYGIrBwz7jf+GHwjwwDJwBDO4eXN05ImgGWYK0Am1BgQUBBj+dsfdx7UA4clAiNQpcqKMPTWT/hgMTMOikmLxzJ19udfyLgImowJez6H3LopTQtDE0HQ693zeU+uB7swAAvOAq99yhXwQxw7L7/H2ou4oCK5khYhUwUIJCKIJKGQSXXt/qEmBBHbnQoBSwkQCaQgOpmBvXh+1b/+ZuCj32G1lwkWCgqGa929+8M/KHMMVDI5JG1SsZzvfEsvRviPoatIKxLYTg+3zzJSigVEmc2775kiGQIFGDZo7v5XeiMBnQWcIbvniW+aEYdWdxAEG0gehkMTQCSlGKtfe/i5anxokRCUEgmKpqD1ur/0g3UXdQabBZAET+4+7Ju/iKsiLdJGLfscYZWwTonZIW/wz7/tejlzogdDdoTf8z0QfWjjildcglpcKUZBIRAuCXCRoAAWwhBqSL/CoZcEBNpQKYqoQgCSkhCicQLpZJ992g97M/woOLyJrNnVYqCOucC/90E5z50V2iQERlpt2MsPxBdSL3/tkxkBBgmI0WD0s1AZq0oFsh7eRV6KMKOlYmE1bT/l6/LG7xs0OuTE2fbpm2YlGEAeD1clMYAdAAG4ZRLTPyypczL4jo5ICYlIWCCie+n7fbBWwwASJ5TF0/PPv4iTBRJTSTt9Pdk9BnFAEch5A9/uu51wBGYAQau/+S8i8aHN+7ysDalIJZ5CBEQIEVSRKhJQJJTTZDSegg4SqLZkAwxiqxSQFYGUVhGO2dkP/IFPeBP81Fw4kgEHRIg8O/ZXfFCgJnQkAQQ8OJXxF9ALgU99B64Qj+qiwaj2WVoIwhQH7eRdhktQjIiARdfy2P/mDS0IrgAdyeU47t40Fxz2sXyADaQAhqYsUVTgYnXqp/E5oR4mgwIKghSEoMrbvxF+UJ6xOIKKCxBHOvfv/vfwDV1iagwRBz14/bhA8rAEZOz5PvZ/wrl2VKYF5szn/JwhYD500ca9XiG1TGzWhFCE05ClFCigpHARAQJJZ15EDgq7mADaVoKEDAAnGsVIiMPjXnzyUv/bxQ/Wp3wkO5CAmCBRw4x+yts+CA4npwIVEBDCAR1+gsyLGP6ozggiBA0gLMTls0SFQoRAu453kQ9IUd5RxGY8/Krf+Y3Ip3AMw0MBvFze+qY5D9eBBHkokgIkhawUIUiKxcn6jPvXuWh8xkJVQoAgiEQeP6Y/SB+UgQHClkQSxrk84Y/xRuUtc0F5zjqOXk/eduA+SIPs9ub53vKPjsuNl4NDESye7f35/q91Lg/3QxfjxItrtBBSCCxIpDYvQK6ABEihoKxUMNL3cfg5REBC2IwlDCOSgsZQaYjYR1ZHPvE/7IP14xYXhiVqWRZxWpr8yfjCWMZjiVOkDASCvOzNJ7EvYm8+pkF2CZDkYTsB18+CDNjSgY7mAvgY4EB7U8rtksCc/Yk3AidrkUAL0L5pLsd1SCRJIMGaUHCuR7qjAKhIMfg1n4vwXU4LYTMxCSqCnOnrDx/UY23RcFDi8cX9yG/dG4rdFoQe6A74eivrA8tFPI8TfOzgr72NxQAKwOWG4/gUPvQ5veJEL4KkiEQKUmDhggC4GA2iKCrIAIffOkKEApHYDBJESg0ngAAauFPEQHL9nXywf/qo7DYnVx0GNhZ14fzZ9MIayIEODCkKouQ8+HX4IvjmLtAMBotV4CyxfZZdMtNaQ5vp3kDfdq1OubVAN8MUcdM7/r03ECekGeBRJ9ub5tqRQQpmsu3iMfejO3OsrKIQMAXiXS47F776GhoCAkbVEVEImiM+/YMz0ABGBUa0TMR/wxtcVgtYcIs9Z6HH4C4mArSB7TwGesTzJ3/9KxPQiQSFcvzkL/wQyA/edTQiBY2GKoSKMEql9USe8EiRIjiioaq9lx3BbgEEEURAGE4fzCOOI4ihQTGKmI/6pA/Sxx9rMMLB0VI5cIB7wPGud/PiI20tdsIVBVA64DvTC/kdnLFSV2nCcIkJ8bNMBjAIQwB511IlSgEtLCBdwf2v3gDIWK4EiMz1TdMes4LkjolyWZgu16PpOrPNQmSRAg91HncOGl9GAsgiFSRw6BEBftIH59mew8NWDUgcsuP+bZ/+BmBDBoZS9jiuPPeTWTwlwhPyuPL6fY3f3iVMOoJIqeMv/U4+9PnIzx9FpAANCEgmSEM9KikRYrRQUQxRAr+LDpcapYQgQATENLVpmnYWNdVQ5EgrDwNRPH/bB+k/QqwIJEx5uDA9/HkfhMHdHUYG4ZAKhTWuL380L/QTOYCCC1gwNgXj8oalTdoyRBrGdwkHAxYBKhaQBuD+bT/2uTw4CESQQFffNJwwrK4IkuLR9dnTbgGue1CshUBJqhQ++RwMvtCIQBECKKgSB1m4+7YfnGMO2IUZisUBWIFj/8Lbnk/YZohOFabT8zkGECcCGBqWHpPLP+yKIbihgNt86bfBD3380zRXIVcAh2YEioYaE/O1i9vP9J2Ll7uXAW5rKpZl9a/miCEgEIAd0XP7z3zjF/zmn/zR7//FXz/jhaMaSSwPUxo6d7/1B+kHU3jaUMLaRu3AwvjsR3wQ4FgOiNdfkIcmy897IR//0mqpUqcT4Do8rM9CIxVIASaBvK2UUJy5EcrDCPD2/D3P1bKcAVIBFW+iKRCUeHR/9J/ZZ87VY24/+vv+4icZWSBjSWgo9zg+cfmlNA8RkCGuMkMexuDx5FM+KKsACwROQCcD5M5vfL64Hm5ABwE55OstBOQpglBy9ZGDf3ScN2UACkLM8KlM86GO79ztNeFRLqJyUWxKE/O//vxzX3FmXSFV02X3++APe9hyh4RoKLe+mehQlQILDKhEJv3qFz7xyY97zEd+7Ae9/7u/3amLHndDK1kCpAAu3L/0jT4I8tXfxeOzqzsz1ITTNoJ++Nd+cbJWaw8KRglIl64/kHkB38eGpgAdMliBrZk3lAJOoUArPtC3yatSIbkBdAEVwvypzwOcjBKokE3HmybwMZjME/7HV44rwGX7nF/x0vf4wy/tHpBQFBCadNnxhS8qoECQpiKEqggSGI7f8UHxwWAAgrAetCjTj33rc8FJWsjDwq3XkzNjOORRNYbHf/7X8HJ/CtGDqIjv9i9Z9kMc0zfCNERJQa5ia8G4+St+FZRpwZgCqzf+07fCFQ/7yFsGklj/ApBDoZACaoOgwTWXX3lSVW1aLlvb+42rn4AFAqGCzDE//IMw/ETE5QIjDcAcOeAIGvx68AXlMdEkhC4PZRGmSx/1tn1j8wvOc1EeN1eYB4P6hh5PLIKxXfi2UitvL4A8ntfT+//kueICgTwUGn32phF5NETovLl+9D84nuy5nly87p/9qP/tm7QDAkgAKtNFx1b15S4kNgVRQwUiSHCe12//QYGAUVACGkAIR/6bN3ADxhAIMzny+nEYzy9zmXzw9X/13ZGCoi3izHn5pX9xkA9hKuysvlvricamoIAoils+7M8BQgfWHHzj337lXU6/2/Ns7XwTRz2BENsNSpL5hltuS5y+nocjvvVXW4+HxuMefL0Pwvn2n3wyjNBV3GvY2QYBhfJdD/TFbGs7GBAMUDAF6fA9ecO+7SOO28GQs9ptoACy5fwsl4EyTAEKzQf0bZAR2kfqOwXHMXN8d/R14EuZeHwDa/fNM+2RhJB2dNPdcQLhPXB530/9K3qN2CoE4cSVx3e/K2wrBgIW8/DACxQtx4X5jvNBeFrE4wEksQwG9W0+6floUJgFIq7z0nPAEyWCfay9+cBbLzcjX/3vdAsXAgKjOs/jj/xylvwQRmDFV/RFwBUpICJGJe11D9pXFLYqhwDG6tmPuM8/9h9fH6kTsi0CJGQGBENQ0vic25cQhNQDLudXe3HylpclHnoE4mv/0ObkYWA92eNbQy8GYg4KyFoKAgVX+GnMG7DvV0toHZIjWxCR0meZgKRCW0GOdxFCwZtCoQVLGuH+udt4zisMFTFWlG+aARIkMdVueN0Azvmrf/cb3T0RECKAKD59bP0roIAiiFT64sZ1SxMQxhrPX7MfhHg0lswAxIAiz/8B8Dk44+EEYkfnc93zaPiY3j25v+6T3vV/c4LXg0ro4IQ9/sUP5mEfwgDlK6dZwoVAmAjkqlc+IIKwPRz9de/7wC/lyOt4IxBCDEoQBxtGrb9NsiQhBHudd724+Jks7p67p55Xlj/yH3HuAQm70Gm/GnmxogcssCAm6LnApuw3pDcQv3EsKu08pxPMB1LsfpYT2R3xtrTzgG+DDUwLhY4gLoNhNvUn6HXkBJCSJcHrvGmqOA92XJHqvGMfe2jLr//Dsl0cmMWxnfosM4qAiUly9htuAx7kprF+w5c+CFMLyVQAtYQUD4+v/imYr3MfFZC1MNhzbbiLDxYoFq5X/tu3nVMdpy4ntMje3H9tPgQanugJIhchVa4A0pvuTyWcQ4nrOPpQi4hAiJRRVrgO2hS/Sxwra7hyDvuWFwc/uUtNxzghDP/JX+PWMTcYmmP8pjenL2YJTo7ccXEB2YNoELh88hvhI9/WGuC5ecQRjw6ZcYfvltVGUWAEOenbhlqoIAQoMVCC1+k4PuOtPBQgIkCaBtmLbxpVJgYGatju6Hnicv07DhADBKyg9bHda5e0UNCSinoWT/17VhQrGnM99vItPwgMmhuMASZJ7th07X8jep1bdAGEhnCP58KakaBZcC4cXO5/9b/blHCwRwfMdHTwaa99KISPWkb0BQIQKClGre4J5lwq0f4xmKBAIEohkCGHCzetimGHYGj2hJsXNXyVl0+Sg+hkLuv5L77iPZwnpGi2XI/vSi/mgEUY2oLhlB0Eafeg38kb/XHszAjl6LILSmQHjfRdpLtadymEqlfW2wIUA6USsC6RlBdkz7/4SDCcgGgYyNibh/NIgFVA1uWNxhMWRxJbK67+qmN7HFlAGEECWO7d8kMsWBmi4EJ85geDAVKEaXEMkina4/ox34rnND1cWJSg3eeQSABZGXa5r7l+/M+/0jEtnYMUp8hP+7vHh0R+1OApEqQCiFGtPrJzjgMhR+tI2QhFYjHCkcxON6gYgPIye3lRy2cxUQAznHT8s/fzZ9w59xTkgct/wguXBuDQA2x2LYWS3fNbHW/kp14niPUgTpYJMEDaePeiUlmt3E45ecehiGWoBTosiSALrXff8m0PgIhlQUuJzuN885AEDgkxnftGTr8VsxIiCQFq3LB3bF/sxChNQYnFd+XlvXEEhVM057eYF1dICcQ6FImkwOFx/tHnuc4JUZDkMvh6MRgLGxAT3N/tk799R1amPYCD/L9+m+GHPk7f0wlKhAsRRaj+9c+5cy4IQZQgqGyYOWqjmpAKLMy4e717UfB9rqUCMg3wH+Fv5cIcl80SL1z2qx8vzoUgbbdJBAJHhn36ic81fOS7JZADhIMrmruDLIzHuwF1CkSmpTJl3hYOtCHIjVHjYQ6Fnn/ldSSGYUF59NLxpvEwHk1wYK/3b8S+Nxo2JSlJ2r+N47ryciQlDiKqqvyCr/tnVggKZIN+E+ZFnbKMQDKJgiyCGtPHf2N8nVeY8CHAMGzj6/hAJsYB4eC4Hvz5d6YkwAgo5Ps/9biJPtShfDRpLhdREVIRzfVhtHEeqI4iGoUKCFGFijRTRzCLKYMhmgD0A72wT3oJRwBZ8Hrwv8Pfu5+rrSACdDc/+AUtoyFIaBApAsQ5w69+ruXHTknInhuUyzADNDHcfpZVDLeNElkH+BZYMA5Ywu3oLJKEA0/363/tx5aTqAkIaOz65sEDIeThOpfzjXR8GxZtCEAq8Ssc95OJrEIChNDLb0h+jMVCKNB67P5eFvBFsAmxCMkuTE0tgO31/O+m13kChgQCpG4+BkdlQrxuTz/wY77zSszOuFAYyTc6b+7iQ541PpeiUkQRSsFA9XfXM9hUDqUcJT5KEAqbgiDQ0F2YDlENwkdMqoGICzVfsS/s941ee0Bi+C+/guG/3jAEQfDgp3N5Eca28qiLw9wPCwU6B5e++3MNP50UF+aYOZ/9D7cMFIAIzGcxCL0BCs4ntG8ZXg8hBbkd8lrNa/cIsgx78L8+FgdHioSL0HHzpkFNMHIHvXnLE+d8zLh4/1lHNxyYxBk6+zvHtfhgWlVQBTBQz8yCf94bggEMB1yOb/xVedHNachQyPLKV1xPF52idZ76Ud+WRy98mJOAAgt2XKLHltdSBHxQcfPsU38PT26Pkdk08lyvV37C5/KafOhz8I5DkSKhCCzaUP94tiscOhyobcc5xkAbQQJSdjEfwgPEt0Igjw948NnAvICD45tCFzdAwYM/xSy/0mJ4/XPm/AZvub6IYDkKdumozo51SIOC9fYdz7PvfvfeQAzg9TyeHNfhYUAQXD/LlgKFDLVFeec5CuXRMsL9eZ63N0Aw4HB+1Dd4ZJDkdYe4urxpLxQ1iAxoXuKhbnDlx/7CzgvaJqki/TjH/ZE7OIhEQBH8g/T2ppVZMCCL4DsA9iKGSyKEwPTS555DQJJTsX/isbi0LA+DWThgeM6mjUcXUPiCP7gsEIxo5DQ3f/73gH0IhGtOdBREkEgqGdEbbzggnHsdoRfKhiCE0LgIHUJokQfczRjAQgV78hcZ1jcGn3GbA1PRAsuvRvnbr15crKDlyL18O18EB3BqOcm4wAcigkWZ4fwZzyHf00UMYC9d73/bHYTI4ob0blJwRoHUkT2XdzvHtoUCKP19v1LOALcgjD//yLK8rkBWvXnCnT1CIFndV9jjhBbgk/7mOzsvOxBDgD4qX39sXwwFYdPYyXgu0Zn/aA4hIYDtfwLEi2zbADRiO37Lz7WBBMLKd39/BE4+oGsPLGa7Gq8vrFNLy5QAX/NrG4Q8DOeU63zJZwDxodCPxiIylQgppAZfxvGKGpIR2iYgRygw2y0Elk/ggyRpvuj5BgwY0GjkvwXoRfyUGoCJHFiffRkn8P/4DMGCCYXfxgtd6EJCRAz+se9kAQwQnJef9hzxmcXjMSfzL//qsrLWgE35bmVZagHqyKzVd9vWGpBaqf/iN7567AMGSDiefdWv+2BYerCwrGDsm4bd4QCBHZqab/bH3/Lah30FPnnHV/u0H/P1wSOWiO3TxGf7uKZ3J04EgooqvJZi8MQpApBYWM63f1VecK5AdAqsfeQfAkhSabL+wBEwnHAkAS7ogfuYwAGgMCAPewmdDE7AajwvfG0+ZPoukZDSIgA32lD+6ngUBjEEb1FAHNFQGAJFgKKtRTsgTt7u+hOjFoQIGIiv+DyWF3k+/eYzwLkocK7z2che+FVzsCuKXMGbvuo7XshQLaC6Inydz7t3HnC21sHbnoN3fMx5cAVK7o/L8Zve9xRgUhaHlXcXullaGIksxncJFyACjEjl09/3o/ecJVAEhT/7mMDCwCAI9aaJXJBEdhB+zzx91lc+vb3cDlxnMLSAiGA9/xc47ocsjUogDNKQfoVU8TejFOCEQ8bll7woPAYKD5JD5z3/vWeGwMox9eT38ch7zhqSpQNkAx8J0W0yOUGIYDglOAhSy2/zZR86eefIFbdAQIWVdvbMQQrc5xk7ZzwGuOJpX8szbq9+7ate84ZXvfI1L3nFtTf+5aEmhDDbQwHXhnHAiXt+yRv+YzE3pgATYvP4Dbzor3l7snGAEM7JZxGc/Pfn5epUEBeWPef7vBCYGBBiYOVjvvgrzlXocJCzm2+JjwzfdyaOQJdL8X9eBhcshsCOdyuwABykLbuKbxuAUqYEoCdfnz/xntt7EKAAnuwnfy0g7jgRaPdB1L5pRhlWTRlI3tZ5ux+zs93LDLEoABlbr38Hjrn4UUQgMcJB3T9CPNh/JS5QhwQ1+x8+8AVcEU4FsM6zly8/gQVjQSDh+7wFWL7sgGgRabcRegTCxoAYgi2u1HEyAUQkl9/9lzg+ZHIJFASEIpSi9DIfFJSfvGbv7hJbLQHRPe56r3ve9Z73utcD7nnVqfc+dRiUUAUkSEDx2Of95zNf+cZrb7jljr3bX/Wj16CmoBBSovjbH/EFfNaIhgDibv8jA/mBf96N0IDANPHTXxRnAggw595dPpellHbh8PpL6JH44TQJgsD5vi+4gADKw3DfTagUEKsS+0DfdvAKrYTHszgu/NwDl0cl4Hr+D8Bwx+wDZgBmnONNc6VkIJPExLicNLbNdjCtkgJR9aoH5bhc79YtQAhAyeKOmwGKnx0GUAjQefauT+CgF5PyyCruxzz9V3//xkUGbFtmX/o9j5yAOBg4CHthHhtuirXEgNGZlY54oBx479/4iRf2QyYXBSwQICJC+AMOG+6bqZcQJgoqJx5jeB7zel6TSw+zgwQgJLbrxNs9+F3vc9nVl15yaqdpPVOKgBBA0P1fvxKQXsC3ux/ER4zO419+gAUOfpe7ODz0BI6+5s2LWPQg4uG514tv74+PKOQM0n57Hp156ZueVjyaXN73JZc7GF5XHG7fDSz4SAUQF0AeGVIQ8DGclz6cP/z53AuBIFduzk/4unIyMADyMBzPN82BAiSC5A7e7BxuOHsyBSAB8vjFh+xxrKJ414UbSpAiKpSe3CYo85vPKIB4/Dj6LSwv8thgeHSAY75Ox4+ahte1Y3C/zwXkCoY8DIZZIx+JY+TAARQwmDSRR+O8ee+3paMPmUytR2pAEIqi4u8OI/IyLSBSpiAgE1nsLCbtLBe7J07U4uGHAWVjazYmLJmdakbUtAQVgEQZXHc+AyBe4Dd++4WABUq8HJ8FAie/95yCFurA8PIzX0RcB8SA85iLvH1+z7IBQgseNz+IAfb8mk+aGYTQXX703Vc8uQIEBBVvsEiB0gqU4guEecub4V3L1Ns7nn3rG1wsHl7ywm/oJv6f1477BwSLteftm+Z0IDFh1T1gGWAgRmAyAO+P173N566BOoZgfjJtICkECJl51uhgeO1Xnrl0irLE+B1pXsTLcyUeDfB6fvUP3P+ff9P7vX+QCuiT3/HgK3YwIDHgXk7qEch4tAhYG1zkwcLK8b5PesZ58iHThSWAoAAyg3btIRQWJzWkgECEVBBABQigw8gmBGLQBpaAQQpAgQAbiGzB/K/ve/BCfx0nDyfQRfmTTEC8/0tXSUFZnPiP0DcEorECB3RyXF/9IpeFFgW2/4gdgD9YPAyhOa//E3MdogIDln03mFgQpAxYFuUd16JvKTbaDv/icy9QPrIQfdd33N9i95cHCbJu8aY93OGhmdKcRwyxkyAJ4fCA/Y97+Cs7gI8BaA8ZTQEUiVhMeToH/s/HAh4sZHLpbe9iX0RNPgiB4xAOvuMrddnAAsp+1EdycHWvCEhYHJ487z0CgWIwSDweAwKf+yV8KLUtbQskJ+AqHM4eIogTRCoXiktRlGyEDJQw0w5qShSCBGBIGEEtMjEaiQgDURLnwffkRT/5Dg0tYGzuHe95P49f+I3HFaBacAPe9VH1hg5yKYlgjfvhj86Z4LSwO3592AeflIDI0u7Nv7he9riHbQxAZPHdWmyLRQK7DO/8ZuPsgSKFzunKd577UXwwa8P5h3nGcfGETmNlyHrTnLgC6c5CHEvQTCY7AsO2qsHJX3rRB7Dk2N95ASVwFCJ51KtvPei34Wkts6jLyo/khbptLAlte3qwvO9H3u4iixiWx+/jnq/MiUct3cXniLutMyCAKGiHAIuAPvltzIdQTgiUIiqLFGwccb4tyGpRClBEKsSCBjZiHJQUSCJKoIJKjSLlQgVTBApXEI7GX/mVL0i+o7TOAjYy8LtxYeDKHzghVh0eTX86b/ykSEoA11P53Rw8zMDhPL4hD7/eS9cDYNvB41V/DtlNKAm2WryzEHAEaDcYru/mUSQUCoLzPuKf/d83bT3CNDj//teBDwwXQMTYrH3TCIMguAOT05EmCMPawBBgccld/+KNH9SO7adJCEFsivAjHPy375d0ltlIZvh5LwZRJglH4QbgT3ylSQMmKHyXl+Q2fB2QdJ9HGJwkH8C6DSFAosBb/tllP4RyFk0UBKSIiAp1KNVCEKGIuCFSBAkgTMWaOmitCMIQUghBRAQKICASAkEYxj/xi31B8UvXwxoCCeA3I7AgX/xvuuq4EQ7I9Se/AHFZHUFwnMvyz5+RccIh4vIrHvklsCc0CDGv/fnbs85UhIWhJO/WGQjg6BLBdxE7uFodkeo5HPLvkwpwgtDyX8LNK52sAwiDHeebpvsbebjjkARGnodAnuMkWLRA1lf/0bMvP6YTD7NMUIUgpDZ+6RB3f/Eml4QBBe7f/WEvxJkpAAtqLwzwEy8uEFgIO/9dvoJKsCCnVM8BB+EQXUE7UEYgGKDo+u6fy4dQq7dUUEWCCKjO8lCZdrErIpEaQFwiAiECIQc1ISwa26UEGQkTIAiEjyyw/M0fxIvi8k1d0E0CFt7/xcSjyZ+6nQKByKK3fpUXAR08WkC7M/uPLrDnWHDSzf23Q/DTmTl4vDj+bidwLw9jINZhv8tQgdISKB0G8jaoFnQIgJvzwtl8/t8RiGWC2uP6jb86LrsTxLqo9qZhWN3BTEFSyiT2GB7uhDjAQvM7vvYex1E8RjUVBSFgh/ammw/B74FcDCSACz/ihZwEuwK6jbucwB/9Uk0eWsD6rZ70eUdrhECeQPN6ywKVeqGoIigKaPGcg1/5bT+EEkugCCCFEomrD1X9VqqsqAwQgVAABTxgiQ4iAUK8RQChIFQEEqAgQbLMn/tme+lFfQshXYyCg8ufwnAA4rcyAIElyfR73xg4wxLkA0T5aRzrCMLBsde3fnTwiW/Zi7EiYLc/g4CLFSuV0ck7i013QKBCmivztsVBGDrh0Ya9u4zw7XORQUMO2v9hvvycQwQYJjDfPAswPJREoBAgTMBEholYaO/Uf1x8DOZbgAxARFRhfgYd4q+dAwZFPjj5uS8ExkYeroQgwI+dYAVQ4nJ3/Bk+gKIC4RzOcx3AgPJojYoFjIAjc9LdX5wPnYwQBAgoQBjueihnAZEbqaRQVAAiEKo6wtuEaFBCIoiAxNawGQgoJOhO//H37Gkv7DecIcysjEjzOxDaB/C5/wCMsRUa2fNbg28gpjMB3KmdmV3+xjLDc87woxh+LfcgU4ly/QeMl0ptAAZn+KwpriKAReUVedtw0AYDFNB521xj+cp/ACdAGcJxfo2v/dpXHmOEAIaXN8+gPBRSEldhooHzSJA6MItC0C995THo6od5IAVQiZDiO8lBx5f/61lxNYQF+phPehF1NgLEaLA0wJ97/9EShCDCp/NKSJDIw2329ZYnkMujjhCRChBAEN780w+dSLbTFEEEUSV80Qf6EOHlg1SFPqkSBQgiigAWLGjbwkklBlAwiNiZA6L3hAO1PDpf+vV/Gbx68oLf8k11ApiFAP/Glz9598d9ja/9tT/xU77GV/n4T3r3f/sBYzdHAZrj5U8F38heRjGY1Wmxm2f/PXcEPiiGn/eO5Wt7u0Ep7fJPXpmzeXIrIKSAXN6tLAB5XGAo8y77NDwuDGQPCvhupwCLCDSH/8dL/yEgEATucvumOW+lhofWcI7313OuT6/1tCdHW7M3LKtqERGmEy1XfiV1lHwCHRRBwBHSm+vEvS678rKrLr/8rvf6mHd9Oy7owAAEl5NvdbyAcQgBCeTmgIULX5tzZkFggdv4ejyj2Hjd7fI88NoiPgYs+v4vn+ueDyzIYziOr/rzP2QSxqIXKEhEIoJ3CIe+tJ09maBGlEQuRKIIBQmRbTBABQRRwJCjKXbmaRFtCGBa4HM+7iP+IS9++OaMy6OCG3zaP3rvv/5n/8/f+b//0d/77H/6Lz/nX/2UtyzOCJHAceXHMvlcMNkCwQA7J9zzi2GHZQFtzrd/hDdfu0bAQPS/oJClAiQC5t1gpwVKKbTIxrcBoVDKrYAnj37B/8rDMYDEPvwj37MAIkgw7JtmKgRIHGzn5/z+d7xy3twzb3n60kf/wk99G4lAEFIA870X+Sh8U1oZEQQFCVc/9TUvePGLX/HiF730eX//1//y/3myLK97sLC/8nwBywpQreSlA2GufN4/vKkhWIZY+O/fd7Y1So/MFK8/7MA+0oOBzp/1rw58ANKc5TS/+i0fKmF1kasAJSlcicfi4stuOpRo6xYDiqZ1NSKkBCIAo4NsO2BhDKlISrdG6zZDEMBQy1//AH4Qlt8LDLSstRncHHPc3M5xexzx0lvQXQgjoNv9QcdJzyXL6dTa8jBW+DuvjTG4FBzw0/vqwrnsWp3s/hekd4EWhBXcf5ZWoVIstbM3faeNHcaBIjO7j/EfMhC4C4Yn//37vgSIIGBdetNcEQEyk3Vf/bvv/3wwHv0OTz/1l31bT8EJMQJq+E856n2uRqMCClsl8qFXnLr68ksvv/KyK+7xcV/z3Z6uwAaRMR/70S+gJQnRYiuCZfhBKEjTFXfk3d/mK67JLsgCxfMsewK7cMoG5NM//SNX72AR4FjC8/qFHzJ5wbv2ZQTRaKgsUWPxjn9xqJ/9y0WNKaMqjCwtVxxSRECYqIOiAigQRWSqSHltqRqNbbW42w/6VPog8NavDtDp7IAcBDvNuRNQ5wk1xUNL9S0fhTx3yFF44LDbcBwh//C4J2BS3Pje/GRAmQNzdt775VBHR6UlOLK+myRlpm2ACS437yhrlcYQoJAcPO6rf8wrAiMgHX3Vb/0S4EqCMPmmsesgYO7kPJvj3XNcb9o47g949pc+/Vv/6acXKAkKYkmL9zx9lE8ppxElDEBGI1E5DCBrMwyAhmHe8yNehAiSqQ4MgODf+6ckD73AbF3+9EfdXNYZAgFn5zngmA2HjvCB5/i//NPOCxggHUBzvPUPfqjkH79mYcFoVAQRqcEXfuyh/uGLvmDiDMu5zWff9Y5xlQBEgcjGjA/qQQKCAIEIsKiFWjwYpNh0Ds+bP3rw4uV7clJzsLIyu8RwpB7OXMbj1DmhTUCW5TPp+Q5YUAjWYbcrDL8Bd0kiGPiY+UGno5xsk/nzEaeuqVALsIjvUjqbBKQY2rLIu5xtA4WiMF1fhx/FRMEDUC5/+q0AA3LCzHZ985y3S0I0PFxe2e0O4J4Tjpu7//G/+x52gIiBKinzuUf5LGgkRLQEKqghJLVCDQSXQIxkjvi5LwAP4kyk2O3yABa+My4hsDAyHDCdJCvBcj6HkA4kiwMxxxP+3S+6cKIFK5Du9Yd+8w+RPA03QoGSlEAmj+GwWv9kLZomr+jL/brvs65gVIWEBFS0w6QQCRBi3HrP3t/8+e+/7sYbbhNlIbbtjnzEZ3wQ4FdygBSvK7Y4SEDUERwxihADfi+eW6PFWCwNUzj5b3QHaXLP8wI/8y2c2MoMnu6fAmZ1WEAV4gXKgtEit0OLb4OFFtw82tLrLK/9Je93l4fLgpcn0/L4QHu2bxqmYFUgwnhKPKesfNbfoIEgFJvGH3yE3XuYJJKiKAFZIYEkGkxkEW2BLXu5vvNjXgBgI7DGKA8D/ef/kIGgpogl0GONKU+S54yZJdZFooAPfID3/gecArpMgSfNs//p+NDIP1NBJEKjoUTCiwe8/AAl4NmSSK0WY7kUZYQUSbExdVALEAFIcQX44w/9uE97u7tcfZ83Ns8VgMCOePL7PwjdfAJQKDIFTQHndAqQ8lAzHg2v7/4E5jkKYGGHELAawPd9/nEauIB2zy+7cQAFgssXfxmx2p0DrFDASN9FqEWQPiIZ+ja5w0IbpNxOj3HD9/JYj0eGCViYHnu46ptmc1CBNJj7eY19Htbls7/sEzDBoC1o/fAJHeZTpxBVRYrY7IhNITJloU0EVWTijF/+xhZCAaaBdh4AK98FUFA0oAlcVtAdbJ9juGdWbRhA5d7b4e/81psBaECVgwv75K9+SKSdef09RgMENICKmvmZ9zsgAAplF6PS7jeFCAjCEaGRbWKooEJQaBWH0mJHe+Hmh8AiIiA0ld72Q17YwQ9kQYFh0WYw7cADzJmCkJJQjJPfTM8BwZgDCnlyuA/4nV4QGEkc3nJJUANO+VUAdnO5EDCsKnLwGbVWAW8U7hjwkQEqVQuCi/F1rnzp3zmCCKgHIhKQtDNc3zRjhiSCtDy557ndhf7xd1QBje2qtMsvJYf5WtJQECiUkhJIAqGiltUxEWAB5+boe72xIR5WScgyj7B84V9fiHhURwhGi2VsGnwdeIpDCiCP9pUrP/eLeSgPl4fH8dq/+1PwQx+D3/uyQBKZAkFIy/s88GUKClsDAwaMcpa9WhIwFARWYntYMgiiiAJqi5C9/f2m4tYXlkM0AJU87D97YSe/8ozHQ8CCkYEhICAZYAI2yLk5v8MbyjMgIjuOVnT5E3u/RMDhcADyME/oPP/LB8urtQBhRHF9t8oUKVCgFQ+APlLuCzpAAYrX1wM+AyeR0KHkoQThody8eTB3MFkGuz17vgXmyV9fBznahtG+OOzJ+2lKhe0JbAgMAcjhCBoIzAog7/wab+SQfCAC4iXWR2C/45ZIgAEoiyITtOP0OstTFigKCKiXiWefznlyPTmjCYibl/c3f40+9AH/57UBSakAgaIifvVdgsKRa1ovEoKwCAiBcgAwQGy12DRmGAwfrWgkBURJ8ORXvKDh+CgBKcA6QeIEFmRxcc4lakAkY3jHN+a5BY5hERmE8YkF//LVw+XRbSMYaJEjxi96zwOYHSFEM2neDSYZHnVaATYgb504JUVpocy8TvB5/+BY6kweFaCH6BJz9iY6BYhzjLg9Hew5gOH68pRNEcAhVWl1iOKDm6xIBAIgjQQoEAFZCUIETEDaH3gjEEQAAkxBr8MHfoFnC0YJFAi4BcwxR7yuhOepotCDyxHA3/7dR+fl4DAQIGX+jwv4IQ7ldfvLCIVNF7FgjEd9+J8Qjm5xh9UACRRIYB8dVEWyrQwOk9eIYeDlr7RaJIxFiPz8eTHyPWYPikROnbbNgxoWoGE52IWFVlaE+MnPB3kCLNXGumcA9/9Uhh6OcEKw2S52Xn4/PrI24EIY6bPPUCC9KbG17Qn0be7qHhio2Ha+zsPvJsIBBNUDg2GZOON402yzPfAIstOrxPMvvgYiRgaE4rF74hDhe6AMJGyKqIVAxEgCWLEJEIu1sv67voGoU6CIouS5f+vdZTcQDFZ3gRhZOrerr7fcLQfXhWCT2isI/MQv81w2hB6wdf3Y3w99iCPw9I8ajUAARQIo+alX3sxxhitauXDkIkFs5qCoIBsaIiWoFSMTQrz/TlyAwElOuD/+2IuJnwuLwAkdFcPMuSysNLHMRq4HKA3EDt//jbCzzTmhVE08+itvWHY0mD0MGLCBgd9Gj8QR50kJAef1MzBh7IClArp598CsSCAtII/74PP+xrAFGeIDzHag5uRNPTs54MoO3fIC+w5AWaKIECHj5GEuub+DAAoUiKAsIGmRQlgNQcAsQjv3N98Wn4tjOkCcUxI557nufhk7QrjCwCQQdcRxdPD6ch3IKWOU7Hrw6Nc6LqwuIIu5czl/8NfHD3FQ/sJri6QIEYAI1cRL7332OAbLqrgiCCUjWB1GQSKElBCQBKsjwqtfbuIMWBg62Jvr933yQvb2Gx8XA+YAW5Sg6SjcUjecQysQ4ZSZ863f5LkGNmh46GgtPvivryvHuc0VhnRBV6DjA/8GQNDdnBGW0eyzZHZwFUBtad5tQAuwQVF8LEA+nXUENoiFxQFZZpy5vmmu0ZGQTE559kaMj/o0BigOQZCqSAeJRzWxYTApAgZlDKGUkoQzZqlLACJXfjk914KwLBwFxA3P/6vuj3NZYFggjkJwQpbtec6W4wCThIlDQPjSnzw7BkLDOcDZ/pVLH+oIb3ztvVyKRFwb0UDiyhsf8upjmFZvrKGikVJskWLikDMxNoogBjolBAb4hFIqxRgAF4k//EL49Et4BkHrELEdyACiABaQ8jAOloXf/VwwjeiDDBfowav/3NoDGJIWS5TOu/mfedhAoyMsD6Mn+G6DlHpz2yq+S3SUTgVKud7gg4f55X8SNnBgBxeDjIG2etM8ngFYM3vyRrP/jKvKlAwEZFjdfFD4AVsYCioCFIhoIkBABrEpcKSk1IVvyhs8mQdDqGD4fPyUmTkxYHi4skDpSTCvFyiFWAbZk4wYfufnHNdlgjwbOPTwpb/MhzzaeL+XFzLgBiTKhDC7L3vM049mHrK3awjCSCIB6qAiRTUQIgLKFQwgeM5tAkQLosG8fu+3vZCf67U9IuLxFWSppmbbNGATApJ4uN/o9nkWTB6KCDHDo79SmR0ZFhQxgBh+CQePjtch4EC2UvpuVsQWp2X4jJJioZqCLRO9Hgc/fCcF5KFowAAIc/PmmT2wfABD9IE3Qp/xXe4vlSrTkIRdaWcPQbuPAYkIxKbKCQZFEQWsdCR1LhiAzvnka72BhTh4uAln9gZ+72veHyIQYAPDow1w+nrDnLNKgUnAkEgHn3q2iwspUSPnp31//NAG8IqbLgsFVASSBFCk/uyvHv/Co8z8L9FGASkpscIOOshSENuLeLilFZuh+DDksuwCwv165b96EfNNrrM6iTjBxLDWKoE7Y1jHAJJCme7xPZ6nB8ViwJ5dOgOEP72yE8H0IBHN9dnfYx8sFzkxIFBn+KyRW6UR1qIvb3ykxEGIUymmXnje8/LqT4DlkQEFRB52elzv3yy2F0B2MgU9Duy5XvpNP+GVWyMiOQlRkfaqvUN83EVrjBIRAxJYmxVhRxCekGTHSAracX/9z9/ApUBAxoLwMs/FT5hJYJGAEEDQnTzxMRgZQCNByUFG9nj/t705CFGpUPa4/4PvDj+kMcof9qwoqYitFkSWxPu94Lon/t7rDZRpg0u++itOARVAgVJV43JyCBIOdETV5LsNDjTPuEkxmMGoy3E84d/7qMsb+wFP5jgkkYW0I2Y4PJLxhEDu5lwEMEilufiTngccHSZsUYoHwfv/+V43eWiAQICXy//J8PiVGSFAHr3/LGsD0wKUQgyFPgIWKRChAMfdc3XlP/8SOKFggXgYnHiwXt8s3SAhk5h0dP28C7u6szzx43/hD3zKSyYQUSKESftFDvkZmVKAiAR2BiVhoQySAYS7CgEERMjtfNrB8TyvGlEEyZ5ez5PnPeYPfMVZCwMYGo8azGtcLvTYycuCACpFTBIFJ3/1c8+q5TRnDMbLzT/mQxtl8+yXPojeIiJwSUSpRGhc/YPzG1/1rH+89vYTvufHvf3DduMKqiQkgRg646AeZALBEkLUILUN8SlCMBGAurj7t69v7DdW8WARhnz/z/35v+pn/tLP+szP+pW/8lf8is/85b/0F//8z/qsX/kjfvj3/g9/J+eyIeiird/6yXOtsLsGbiLyuv/rdAk8A6HiocCvYHvssrpoyELu+Vm2DEGwKpTyztcwSIGSUnaf6+F3OBhQGhYLVpi26PJmYWgHQFIQb37yH32rN9ebD/ukT/moT3r3Ow8oAiKpFJBG6a8P8z6ZWhKC0DBY/MNXfd03fvXXf9NXf8XXfvM3f9O3fctX/Lxf+vM+85f/ys/6VbQtaYiwncencT6PoGg8lMsxtM/huXz6bQotRCTghrS3nievG9sCSxCAWOgCwr97XNjlPDBKQK/v/KMwH8IwwMP3nGhIEeUCN40WnDZc93pvCAJiwGmRSJAGkpiZDooiUJAFcSiWxAek/uwMKYIFLLMxH/dt39DbPgKQBgw4r8v3/Y2//hf95l/+mZ/1mb/kF/+SX/LLf9mv+DW/7jM/8xf/l3/sz/zFn/d3j5Oh4mTOtPXbPZcsg0gOrvH6n4UnwCEF6Ebt1et/z+tf2WbBYlqWVz7LYKZCcTpM53w3djJTKc5Mmc439Pe/cIyAkrSmUlnqTXN/FDsA7RDCz0uuFx7NBQwIAQvHYtzy4kM8/EQKKQJCCYrx0T/wvd/zg0/67u//4R984pO+57u+84d+w6/89b/il33mZ/6ifz3rABjFcsTv4LmP042QEqDluQP+2udzyiqoFcTQpjk3ryf3O1xPMVxiQyio9Mt/EQoD5wBLO92cP+BbrB/C2NT+x/0OcksqIpBKmgs1t4FDA3qlWiJXFKlcUCahOKRCCguEoISy4rDmM9swAhNmcz35U2/oh8tEAgIMw/9yubm5dQ7nOAYuqDcdF175VQxRdsARDNdf8JgP4gCXGqJpjtf7/GezBSCSNETWP+k52qMmUDBB8m5kamEsqYgP79a0Kgi6xMu8Ib49p9gKIhki2bHqm4XzcuwANCcgnOPJscz9xNG0B4xRIAWi8PSH6KAnhgiEQANFo/7iJgQgKKAmRh7+BI7dHrqCw+Xu6x7PZ4hhAHPOG3j02x+zTAIEkuTqKnO+XpzCgSW5gO3wuHX86n/SDtIBZh7bSf/L0/NDHPB7z3mkZWkUNAuNRlmBFAJcTAApKkSBRgJUWOGDZimxEkQChgR0EPz2qpkEQ6COs+/+lm/kPzYlOEPBe//Pu/Ous609z4Ur1T02/Nf3e+q4dpLAXL/5zSPB4HbFAy2Vk/P1+l8usbUhwYo2nv58nqfN5azCg4x5NweUhkKccXwXgSK0FAU9eeP/6AuPkxpGIoEEG4jeLO4ek7AyNeSxTTXnHJkoRUKB6BHOon8OOeDiD0bKiJwQKcJ8F5WNgAF3tkf+59cGVoOBPE+7+TbPFcdAgEQ7V17gP/7cCAvQIUnGZNt9vaE5ccpWDxPqdYKTb/LegxY4YRTyOPbJ/8iHOAPvtLeTJI0IKVEEgQIymkoQCAThlhAiihQmByUpFQKk4MK1hBxihydIAQLJkuJPvoEPe1cmc21s2Tj4LGauwvIGrye89sVjG66CxXr7bR95tLxUtEi4PO8vY6C0IAT0xPO/fj3JndXDEJKGdxUqcjvQInplva0Ui1CgMAODb2D4VnBCRRRMLQssxpvmMiE5YXMe0hRzcrBeL8AAUwqKaWVp5Gc7B79nWSAiiMDl2v9nwjE/+5zYESUQDo5+/fPZbkrIiXG8AL7TQGjREls8PGHOy+uFCsIphCyaPvLo+7+r6dYERLF4/VY/Ez+kAeQuN5VcIFAEclVEiChGA+RKS4o0ChMhiNlBBzWJRIgQVED2OfSKJ63wEGhx3SvRx3yL5/sZGEtIDMrt/g8kV17oDb/7IIBhAsXuf8fzHMcCAp5ZzjzH33mGo5Ixj9TBfN6+Xjio0FgRPn2ngqS1bWxFG9lvW3x40xosVFnpDez8s388pgU4FDOYCfOmOa+Xg5ACOwBzSFeZ64UdnBRCBChjOvP1HPKJiE4JSaJoyfTMNcf/y90LIYlwpPt1b57nNIZYpBFh8I34z//yIJukyShZA/fb6wnXIfTAjJ3oiNcV+Rt/7sgg4tFYPX/dR/ehjunW+71sITcgEogCgVBUCIgqAiIBUiQCCZ0cNBASDooqm9qnHYZF/4FGkTkgx8HhXv87wNf7MaXlJXUT+FfPDnfmhZzxBxpHhTSwC5/4BB9ZytWQDsiD6+sd9//X5QrCrAaQw978TMRHIExAAnE8yDsAtlojNUDKHb5tgAlAQUcpF96gxHdRZAYFVEjLg33TXFAEHHaERR6KMXPUrEIkKBHkiXe77SBd8dB9pRkUtjt8P2RDx/Hf4mmFAa3n3j/91s8zDCgjDJCr84b49DVGFEMQcYYZE3wAL19IVlDkEEdkHkHke37Oca8oIcjg0fHX5kMd2Nv/5Tt29uDRhIQdQh5mAiEQqSLEqEV/ThsHzQgXFVAkokbIoWa+eTAWQEBAQBHv+Y8cfMUNQRAURRaIL/0xzumbr4pEFLE9bp/7c2wqGFJsHeUaZPegwWNfixtBFUSEcNXq9wk54HQ1RbgAgYsd5t3utiCAADX7oG8rIRSwQmpa/SghvPo5D/WJpCKAoCijkXhanzfthkujiK1BEdFGBBCxNSLCkCmufByHzBcjF5vacFnT3l+zPcex95xHDApAgEi1ie97x4OWDSoiikDBC5JxFODbvn20IBdHHeWJQ37kH7rAlUhRBPPZ+99Y5tC7151Yn1AUxNYogr9+///uAC5f8BFth0DsgMQ504NldkJ2SBIkgIZDONhOE0CQAGGWWDkE5E8e06dEAohwRV033uUQj/+hURGpiCCIctWN50L5tm/vtEhgyUU0L17+QIWtvYEFAlxZT4eB1912KqCMhgi4klHP5ZBKR8RSUExVb+azZlIasLgX5cDeWGDSIiCjAjnS1o983YyDAJIiVJI214rzdnnLrosDtkZJ4QqKoigCkkZQOWOqj+DQX4mwBEFAjQn9Lef0659uF5iKABn8DhysiaDRCEEoLS3KMbC/SIuioC1JEWS8OMxH/GEkF5sREf2O+9/EUR/wspGmuEUQNkX0AX/13x720t/5Wnu9aALRbAfFhOQOSa5ZcUuEoFr5EKNGgNEgiNGAFYTDf8ItreSNiIiEIo9+9kFvukufcLkYJVeUXHdXzu20TpRSRBSBa15ceivbhwQuCMLqU794dQj+5qMGIA1FgEjQ5x8mdIKIFFclbll/lhImwESQ9ryn3BasE1thElppciyvf9k900yESyGFy4z0fv7ozC6kcAXcTOFyERFQUkCEG8lkh93Vu3LohzOaGxECcDUX33Bu/nZuEAoroQClPfLfD3Dr5WqIyK6AXIPj/JIfNUiQDRdWKLuNQygf8kcTrtG2BCVj/163Honf+Zh5KQzCjVEp3Jc33+Psf3dg/IT/5GnnhcwVGjJJqAHYCTOUpQjw/lkOKZSKGgRBQTHTxkFKG9r7nm8luBERRUTAjVdb2bj6jUolIBERWfzuY88R68WgBRGNBkGu7/26qW/plWLQQipy6Jdw2K97EQQJCQVnUu0/7zAwRZZQIpIU7HwWHNPUAFuayz3vWJKRmwCESTse3vNNtkZLKqpAisKqxfnD7aeSxqYpNoNwEVKAK4iI0SpmqWvficN/NTUaKUioFKA891yI9RvuMyKIRMGoofCT73KAxnIsgCAXQeBwvDefTiOI0YCI0aJEng4qf+QfRkQQhAvXut/rxqO0wY2npyhuAAGB8fQXH/TfHTb17j/ynTg7zFJIEtiBJGEnSQEUg1bPvuUwu7gCWAKQyywZHBwG4Um9oaANxBbX5/8MgPK9X50aNCLoTQmV+7z2XP37I3pjQ0SW7DbvX1IGpJXKjSggYCSHe8k8uUYhAkklfedlHH5fDgEqLsljsfwsgqTW2tDMXti32KFmqFQEW4v1sbz52e8MFWGhDVwpkfPo1osjBBYgcEEiuQBLLiwIrStt6o99Fodv758sDIxJwiVX6tWcyxSf/Ve2PCGA3mj2eDgHrhe1AIIQCIo6HuVBLxktIICQskC45uVhPv63R4MIXIkg48y9bj8KcL9X2BMBSAER68nTJ/z21P+bA4yv/r9/AneO2OyEELgae2RCYjQ0mmr804ejHLRGCFcRJakyCOWg7Wd/4ktCRSk2gwK64ZrRBnDdFWNypTCKUvTqJ3yuHvZ8k0LgstAQnh7wCgUm3boUFYVUFHW4+FA8873XC5QiLsnl4vsO1Yao4SpwKZFq8FkrtYwULHs5lHd8OUyACdS2QRxdgQ9+0V1cGLlSJgZZTOdPvenKFEZyMZorCMhGAI0WrMhiQZ76pRz1HXoqRI0IlDRcTzoniH9YT31RERZqimrqO/d4/TYGRkkBCVLcdCzAa+451KIoshCBVG/7Jw+CT/nVIaQAimI05rus5qPx7d82hKIoFY3mGpp3brjf7fz3psGl7Vv9qY9ihWx4uAPsLJrhTgqEhfZ/++veCOLgW0gZHBEFUaYR5Qh87c0nhkgLoKQAFx/wVxsnbx+TgZakIJUx/cljyueo3bE7BAjhlEYbqV/6LIA29pejEBbl0ByfOIzyFf8+pezJJSBJ8ceHGoQBjBpKcAPtfJaJ3aluaKtNecfsyXRnCyNs6zyOIcCtH/1suyQrWFYFg/r5w/WXuQXJ5ea0QBSlYppJQcpRFdV/5TtAOcKnZG7C8sIIIBS/ReHjS5t//xM1WlyKyNaWX/jAA2aIBLiM3NDSx7OYr3ljZEVEKKSwNNi7+CDlQ/+UIUgLSkSYz973Vo6qwEsf6FB9AlCCehviG7/vvz2OVoJv8N9/hEqukhCSlIA8FIu86We+9ywl9fIBb46iSJachovGa1gMjqjVu/2jSzBaUkAEo/3YlwJ84m+MNqRIOFKE+MTf4twq/PKnWbJpKMLCo91yTWdzNVVERADkuu2avUOE5545EWgucKHQnn3HoWA3bVRKU4QmVNRnOV+MQVm7GaX75bsMwUCkqXMwmzoGQPnnf3ukROQWR4plqV57Hr3kgRECQ1SRU5HiAkUZU0BkvfeiH/8TaunOEd+nT1ZDI2WkBO+cWTUPzgE84QUyUlQmRGrD769skyogLAsJ2n4F+WjA373XoFJJkShIltttl6CU2fzgPyMIiEIonJvuewbJhwKmnec+IIgEAYQxDeU3P+/sf288/0vf9D//uJfg3nOmleM8TvX0XC7OoPnWX/uRa1eAAXNgk/ZEzRfZcakc3/FsZo4aePz3rZfeIUOxQIH9T/rTjYtfeLd5fSo41QiO8muf43MEtBfdBxYw5MFC62ZPf/LRVAJ/9177LSNoUlcj+ukv4UAF8soHn51ac4SY7Lr+w+cjnH39PbtkSWjUUOY/s+/2v/8P3R8d1zFzzbC8/vDf9C7l//GnfzFzbUzpcJfrn62PJ/C+/3aPJhbqGVCR6M2/cR59xk+1uZXHQvQwASNkytTpUY0Tq77/+r/+gxesAdw58q2nlmMN2rFrlqYC/xyDc5nBC//+vfYzhFuDde1L03ztcrXtaY/prVVfZgSVKV5QHXOsV7/0EveUpZakEorGr38Kh7zrc69oY42qEXVpkv/kIznee/3LFfNyLJgNNOnsbnLmUS8v/7cJ4Kd9z0/9+u8aWBeyUR69//Iv/I0/ec3tII7+2Efe/fX7TDvrRUZfpt22X2+YOVZdUdpZTBc5e3hM+5yd+xm21iXsXlFZa9lqNfaWJ3jJPufhdNElJ1XVx0yr6eaozzezvU6duKR6r2XtnrHtvLEfFIBHXHz57mT36lKfzi6vv52jvs3Jq9PFjhfWavbODTfyGf/Dr766et877l6/4oFj5k/Nu+Af/1Xfd1wfznV99VPHNZdXl4eXcOy3v+3Ji6Z2EY0eMs8NuPHm84jTV+9lmcWCEXbdWs1drZcyVqsa/exZc04/4hUsZ5bZWWvWvEB3jA6gHJsCfOAlF6c0puKEs1ozLebbVmz/pLuon2i1sIYW0m0n+7WdYz95enElXXvLLPaW3SjJrTcN2jiAE/c52efR1juL6mdp+yduu5bjPnHP9dK1u58x10T2+s7t6z3++1SGE7j5mG/xGV//o14SLvfu3Ve85z2f+6//5r/4wlcWEIpypPwbWxXujOUtd9Iyx6+gcGiFI45bXsk53rvxxeHcvv6T5bPadyn3v59HLX/O23wroPAWXTm+O65fd9J9hs31sNgajj0A6+s5h+Najq4c2/mscD4qHFX57xPgONHl8cvMtZPXlwrG4uhqMRWV5YoURwblSAogkCuhMAqbCqCAsjGF0azk3G2VkjaUQsQKCptCrigCqUMlh0NRFCFcmCI5AiqMIsk1IqWfBSyI1LFhymdM0zFMbyTHNxRUGsiKKmgg5PNHBJBCEDSXiRQwFN44p+UyVCqkElEZnIcVhCWwaigoBwFKBcogGTkcv9KQFZEQQcrlw6kYAkpYKSfHpbShqILcrOq0Xua/Y48KCQkv94ewy2AzdbR0YLC1I1mxQjgvFd6SinCwwmEVoDhiCCEYDJijZwDGgEPMZy9AGRjK8NmHDWweL+cyEAdMyICCEP7LNAaMAUMSzkuDIQYyCIRDRwYGYMC8hVQ412VB2FT4798BT2B2ClA28NFQAMoKkBAlHGv5CAr/96hwbs15aT/To+XzqADhkLLCfyUHKrxFDYdV7mxtbDv3FcJ/Mx8tgBCEYw+bJmwGwv8FKkdQzlF4S1j+fzccNYT/isN5rZyzw4cLusZDQzm2rYq2nXPlKMrRlLcI/1doPzf/Vx7+K1bOjza2Kcp/Dz1vOLch/Hdj+Xn//7z/f97/P+//n/f/z/v/5/3/8/7/ef//vP9/3v8/j2FWUDggnlcBADAPBJ0BKgAEAAQ+MRiKQ6IhoSIk0hnoQAYJY277Od1vH3ak/G3r/3OZG/c+X5O/9fmseQdPdoX/S/+p5afVx4I/pO9Q+gB6RsWZih2b5S/VrEAPmL2XuAPwP/UD/gaqz+Gf6gWkByEPmLebf/FfgBeY3K+s/xH96/Zv9//EjkPxL91/tn+Z/zP94/cf5qeL+qPyT9o/yP+w/tf7t/b/+//73+O/KTyE6w/8f+v/MX4MPN/13/k/3z/Uftx85f8R/0/8P/qv2H+jf6b/6/+U/f36Av1a/5v+O/z37efTH/b/+f/S+6P+9f7//uf6T/U//D5Bfz7+6f9P/Efvn8tH+n/8X+z/3f/m+Un9T/0X/e/0X+J///0B/zT+2/8X88Pmx/2n/i/33wV/3//df+z8//oI/kP9q/2351/v/9u3/K/8v+w/23/v+jH+pf6z/0f6j/af/r/WfYR/L/6x/vv2c/+n+y//////AD/g/+P2AP+F/8Pcd/gH9K/3f7Z//b5r/Iv87/d/xl/ar5mfM/3X/M/5P/Pf4z+8+o/mD+Gfw/+h/5H+T/an5JMx/YZ9Pf6/0ffjf+f/iPbf/bf+n/O+Q/zK/3/UI/PP7J6HH4f7Uf8Dwt92/1H/i/2PsHe6n2v/qf5H1Zvpf+v/mfVn9a/zH/h/zHwA/zv+x/9r1o/3X/r/1P7/+oh9o/1P/v/0n+h+Qb+Yf2H/of5j/Pft79PX9x/7v9F/t/3o9vX6f/m//Z/pP9x8hP81/tX/k/xn+n98X/2e3X9tv/j/qv3/+jL9cf/F+1H///7BGY9+j3+VDF+2pYx79Hv8qGL9tSxj36Pf5UMX7aljHv0e/yoYv21LGPfo9/lQxftqWMe/R7/Khi/bUsY9+j3+VDF+2pYx79Hv8qGL9tSxj36GwvxYT8se/R7/Khi/bUsY9+j3+VDF+2pYx79Hv8qGL9tSw5b5rYXPPtOAItIaqV5iBIo9hqvL8qGL9tSxj36Pf5UMX7aljHv0e/yoYv21LGPfoUPdeo++ANzg/cBektDfhD+lZILGPfo9/lQxftqWMe/R7/Khi/bUsY9+j3+VDF+2nrP2ow6nfynVHBBiaBo/oR6QBltb6D8se/R7/Khi/bUsY9+j3+VDF+2pYx79Hv8qGL9g9uufTCCL/rGvpA5o4u2gZVy1flY6OGL9tSxj36Pf5UMX7aljHv0e/yoYv21LGPfo914DA+1gKVac6zDnHrQkFq8B/46hYwxftqWMe/R7/Khi/bUsY9+j3+VDF+2pYx79HcF7jwJLJj3Q1y6/SzKfR/i7eBsaUY3rdhVnTATfhVLoydn9vhWAmpMJE9NI21LGPfo9/lQxftqWMe/R7/Khi/bUsY9+hQ1od4QX6BSNI4wBEdNiAJBc0cxjnJ282EfK4TajhmDdyBAqjoOCq4uGSzaljHv0e/yoYv21LGPfo9/lQxftqWMe/R22/zi18+rKcWnrP6q8wJBtwKXpnaMyn3QfRL9w3bE2b+dUbmNHT69mxyY20FYfBd3SUMJuOU7bpqtb2hVgH/EXo9/lQxftqWMe/R7/Khi/bUsY9+j3+NivYPQe4ctDiqud7HTDF3XulHb+ZIVIycUE8JUDpOacYb4F7+OFw6ftOsXCtyjdji2bSuUihKE65bGiMi84yhL4Y+HhtNLu+lx5kd27bbej3+VDF+2pYx79Hv8qGL9tSxj36PfxddgoO1+12A/o3bJVIT+bLcgrZOK6UBHAAVwe737CLXsZ3g+hbSi+Gd2z20IuSy+d16GXd8BaAi4sRwNxDfg1IupbI8EYj2FBJDvw2cEPY1dN+WPfo9/lQxftqWMe/R7/Khi/bUpGySrk1+YhnkCxs8GgBQ3oqGIDj6wIMUEamv6371QkiWt1LT24GluZoy3++H78R7Ec2tO8GaCBUPy7Pjxw2wqZOOBKjcbAgFOtzh1wdE80TftqWMe/R7/Khi/bUsY9+j3+VDF+2MpP5mUscz3tASHIOfpB3Q2rlmaHYGu8kXOTSB8kNGeMEYHjPRPQ/mWTMqbYCacZ+RQBkX3aRx+UPW3aCDIghkaAfPdjQZTUjf1McuNypYx79Hv8qGL9tSxj36Pf5UMX7alIsllgwQ2lnkLqezZ+UDmyAsXS3evOsA7Qsc5Z7O3e3beZ1F0Fz9wVLDhu725w5QUEhB8EvsfdiSir37NNIweahnkoh9GA3yVCUkc0xItW1LGPfo9/lQxftqWMe/R7/Khi/bTzywGHp4QwHEpctIUtkouXuq7oNiyNmUI40J8XjZRvJFMaKZdak9dAp8giDf/0ertH3Ka5vteFW2Q6uhaz0uX57noovFfDZpb8asfTjYlqZRgsX7aljHv0e/yoYv21LGPfo9/lQXpdLqoR5kV/El+Rt2qnSWguPP4wnPQSxzRCbm7OZp6DD22jseFbCSOjiQdbwR/n77yQSkKxK2oTB7om7nEVDVwOOw3LZjNNM9nlm1LGPfo9/lQxftqWMe/R7/Khi/bJg8UV8spRYR86uVUFn4Xzmj/LuxGyQcOtHPD0KRodDMU7e73vyVGc8/STrQ6lrENWNqFcCxUdr54Z383u90INGlT8KvhT2DNCxftqWMe/R7/Khi/bUsY9+j3+VBcA/O8KfDZ7dKidQNf4g90E7ZYlmw22QkAPQeglOmcoRHT/RRfF1sQyxntKBVF+toOlVXu8iw/3fuMlp3sKi2c9ze8DqXo9/lQxftqWMe/R7/Khi/bUsY993GN4JAddp2WeXfSE/OW1hlb+VJIkHN7SiEX3om2K6tS9aUNj4Yv36Qul/5z+ZjLTaX1CocNREUvR7/Khi/bUsY9+j3+VDF+2pYx77wqLfI11ncM3la1WNLsfPT2+Z4xyXNnU+0qvNav2wUuZZrRx3OkK8SqblN+AYGljfLnc1NGqWMe/R7/Khi/bUsY9+j3+VDF+2oSnh34LEtTCVtDPA9SX9h9SZt2BqyuZyG7GKT3fwINaHJ6MlF0Oj4F/qqhgihOYBz87dEpedc8h2L9tSxj36Pf5UMX7aljHv0e/yoYrsscrDnoqA63xfEe2Y+Gp1oupZ/U8ez4MH52QLa42b8nHCAmG4EGlnnuwvjQFqL8Nry/Khi/bUsY9+j3+VDF+2pYx79He/CxH+rVy8Xqt+utIgiA1kvnX9RGJsgdpg/r57bySohMDYjYlA/UiwhQysAJzCuNmQ2pYx79Hv8qGL9tSxj36Pf5UMX7alhgeUFJH7hNZfqOpKz/sSQSXye7rdDjeJOCshDuCaFYhZm0XRu78iJ6YnFjfpaAJTK2R6H8i6hVZxhi/bUsY9+j3+VDF+2pYx79Hv8qGL9svRDaPZN+pnTsFcw/thTctVg2NiVmpyjfo6FPlgOcfTHwt3KvAhX3hevKUJBX/7ZyvbL87bJywjDuhJo4MylYTTVh//sc3v8qGL9tSxj36Pf5UMX7aljHv0e/yoX/25QjcIyOO8SpR5KhXAdYHYFmO2wFaVKC9fSssWpAAj3w6HnAjJMl4jdRp6R0diRK5/UCx/6CAeA6JU7Crab9b0oNfzwvwFEV8qGL9tSxj36Pf5UMX7aljHv0e/yoYv21CUaWzTJjYYWitCyzgUOUDvXNj4ukP7ag42e4Dd07pxNpVR1BK6ThEPMNG5Pp3xxYCi6UXgwScJf8LKRdQ0t7/Khi/bUsY9+j3+VDF+2pYx79Hv8qGL9tSn4Snh+l4FfE/W7FaN/OL9qzbYQxvZhpqrnT2g17ujQQWUvAEPhthLgkNua7Si5mTFwf/qulDg17etn56eoonEbaljHv0e/yoYv21LGPfo9/lQxftqWMe/R7zCbavOvK82stncCsI4OOX0g8/MaxaYxZcVOpZlKBMv21LGPfo9/lQxftqWMe/R7+LsCACOCPG+xuRByXcdX3gDWAvcmlDV+YIJeIXhzIvh9pxEJguZTAT0zJCa0CJaf/krFEletIctXApsQ7mNgUAGxFKlskt5wOE4Odz3CI+FFYpnuEO7VPTQ1O7h+7Cysz3YS5MXCJO4dvF/HiL5s5WyNtSxj36Pf5UMX7alhhr+d5u6IeflAD9fz+rgoMHtHN0LxmRIdTPiJn+JBiEKVU/oVAqQrQ9Niec89DkiAnP8Ko+K/yjavWBltoJnbrSjMtz9NWloZVaZwGZLC0JUol0qDiPBWhtJcEvAV4uSqrDKi3n3s9vjQtNj9YplEPvzQjSU6IKxc8idMKH6khKqe0FbrtXrllIb7a93OihQsa9BKaoYv21LGPfo9/lQxZ45RrK5JnADRQVcqSgy/WJLFjAVrQ/34ZEfLaBOPxRkL4oiku06Uhq8++IP7EKxcBpfbqr5Dj9SZeQDs9rcTimNEYCSfZlpgpIUcYFUzmGMn45PNr8ECZvHz+/ttrDgolIT+JbOvFEYGgNyxdsQHXp1vsTP6zLyWmGnWNIgAzZOKRpZP4Ez49d1olSTjq5Vy6m5yFwuLKFuefgxo2Y5OVglNUMX7aljHv0e/yoYGDArFSxwFWG2c1343BvC5aLO5/lK3cv63FFed8dqDanrUCdWLF1Bk4FK39H+sol7A7ErmrvXpLBbgEtVBiFX9zO5b6P/rf3vKbMi/X/bynG3pfgZxlJANOyyADQu3NIuj3W2ek6y1xm84ysJ01Hm9yp68ip1502VB/6/wMDBP4CHKdZ/q39/Vg4woWSN7DU+TmXz9LNJJgO9z5O/ghkvrp2L9tSxj36Pf5UMX7alhwaSXEao05YI+9+xntVZDJevyqom2FC9SIRSYvdch1hwhLe6rGWE9HMNUlVQ04SbAWuTqIcUff+Bp51VrcmYj1K6jdnnwz7gG/LHolAbl//xMCm6z5AAb3a6Jd/ub7AIB4Ba7eivcFo46WtUYQ+AbHVDIDrJfLKx5VKetPJjCjhCS1xYKggyK92SaAN/U5c3/w+8C4+ir9m6cRTyELZSVRZyyYx79Hv8qGL9tSxj36PdiBurPycZ1DbIulAmklr6mrTZ9wtnyX1CXNV+sckt/kd7kE5MAY9YpfLRtevMXMz6a10piN4/MBmkwdGNsXUIOgrjk3r4QeMMXU8S8fwrg7SZRh/I7knwmfwAXAzw+CPxU5B0w+VmESE/lclo6KcCnaDHWkwDCylVmPRt+jCDuKFN74Xb2eTvDL64TTKMyt7QteuQOB6zTwrZvimeaW9Ty4vR7/Khi/bUsY9+j3+LGmRmRCKiwTbNy1t6DrYzACwhmd7248kf+nsGAkjEW7tXj3IQbCPaX/1UWDsI89l7L6GTlSUwTBXCPN4gOVMOUN10x1RUZ2V3Yhe27h51EG9lx1YVxzB8luXO/9CSVDKGfQL7c2E1O9gZspNE4xqvFG224Jwv0A3S33Y6m7C2D7zzgkAm92MnLxQxjyfGcH5MZQacizZbdaulcoerK2Gdm4rCnVQXQWL9tSxj36Pf5UMX7alNvWEKvJilR/ecRmgTqgUQZ+OYbBjTvOcpCjjouuY2N4TwmgM0ZoKQEtnuOaEvOwmO7ALcVW+Q4//yKQNk4ZDrQnYE2m7iIyFKKV/I+5WxcSSFKY9rUb+DJ1UX5iTMQVq2PbQCVDZl50it1FuLvzpZMQm6BClzFnC/zMDfOgsX7aljHv0e/yoYv21LGPCBdYfZ89F/vnrm3jwie/q0//805cXlCSBX3iO8izCNBDQir7Z8lgMBcxwRqt1ChSepYwk01ijM1UCxftqWMe/R7/Khi/bUsY9+j3+UklfWbawcrm82cawlKyqoUQ5peMEZNUXpABm1pCN4yoYv2xrNFNl2BIvyoLlItN0PENuPSXFQ/MDScnSY+4+4wV8MX7aljHv0e/yoYv21LGF/2yrTAFDaw2pPAqWX8QlqrJUU/+a3OQFzeNVPqgNfsOKp5dn03eK68ePfgxNpxTupPAqWXIq2dDT7nN3HKM940XkX5e7C+VbFKpiXQ7MG93YruE3+vElbqXG7zLSlH7T0CM+HnFPaM2gSgLbdGzVsxvEL5flQxftqWMe/R7/KPO9acbIJas1rQ4p/U7hCWH6R+GVeJ7oeCtx6KpBSFC6T8vBQUm9jRcppcW3sGxcrb4CfbWo65c8dTRk6b206inmt65bJQvb1VRexEDUvbgVa6eDBtkfKmMmPGwWJcz/xSon7h7QFs8k4eUpK3BgCMcou4xeMKu1XQVpTD7ikh56AY3is2YHwtRJUYpCCz5ZTi3CTT31gEfsmB5trQteo+7urKaT+pzFXA3sSmqGL9tSxj36Pf45kKMS7ChZBTro5bhhQNO2x2vZplWxmbF1VcXQ8h/Xw04ozLyDO8Tmmv9V/RC0sz33HcILCttMW+IrdsFoaDkC9VfbSsEZcUTq1V8y4FntBI9jzeTQq9UiT5pFDoPTh9K/GXpOQMHl2gz6khYlZ+7fH7mG29A+js48uGzOsd5EtLrBQsJAgNl1n0nRpxzMZG7VZHZSf1FrLpr866DiniCmse/R7/Khi/bUsY9+CICrXQN6369zVrAN8KkeDLEFEVn4rh1TCz7/PvZKaaJgCLcbSpXDJbtby36yzOk4jhoHvvw9ZT809xVVY5kapH1QSVgbeA9lg6nJ1lOsg5Bx+YB8LFhEh0AYE8TA/76iL62rxL6ubFGveIpW8Kzb5yYDByb5NzBTXIcoeA5t7yf3psiYARXaaN9Z4yJFBoCNSLcciUDT6Opd7z0JHrWkylTq8qItqWMe/R7/Khi/bUsYOZYZROLevy42zn7OSOCF5qqQ3bRYmJW0qfQPt5vc8B5TafDb4xMqsslcYrrr/C799m9o3k0BMHrLdG0vvbn8qugSlIRtC5szBaWPYSkT6kJNFy5PSekDla5lgrP/04ZigU5vj22fHIsd0CgUtfLndN7/cJuE6IbF7FZI17MfZe2hUcHCg7D2V4/RRhrnqyDybC4x1SacsyCHm7sMQ4Q7Ufou8Br96qx4dcbwEpSO+BXHHv0e/yoYv21LGPfo91w7gFpXPpE1sEByktfSx1mDEADydi4RtDoduekZoWa41bxVk8JhfJW7gK1UFOA7q6s/AoBt2laiMh+D9/HiXadGBMjVlu7RB9NoGTFBCsEBHLZSc6SfGso6wszcmYXbTytAODThvc5aP1MecicENrHmLCJbpb057RJPvm27WYArfqzUzNQSG7xsh1SVAqeVHxVO9EZz/yA8/Wj8wsz/Ez2LN6JaVDF+2pYx79Hv8qGL6wh5P8MIBjt3TEtXOLZdxqSHOv8FpyNvPQJQMVkYRE24zN13OdCFEAAIcPOewab84YR1wWltFOSyUeQbi/GbNOttspQhicTUuZCrcT1zosfP6ATqmf9DwpK/Ol6KV9QvcsR/EETC8EP0FWGrzgxZ8E5W6JbwmDohTxdhpjzlaWNVtsLFaRW4xzymkhyr/ECH8Tn1tDC15flQxftqWMe/R7/KgxR0c67hb53753nTKQ80XqvpmmLr3vtB/dtubaTrL23ujknlLFyVvLmkPJ8pflKSM1Wrv8JSiuXtqU/2rbu7yPjC0TRBnk162I9o4yEe/yoYv21KaBSnEV+siyR0XYwyzLVdqRhmqmKJLbEPWH/fOW2Q/9X+kjMOIoYHkfxSm7ZWwvJRDVnFHfrRMAsC7nJYSyzmOU/050l90KkixpvV1UGiyshYvO6bsLPJP7e3AVB+s2UBXJJScXSD4RItxvMo0KmhEMCwz4j/QEGVEt7QQtehD/ggboZhvT4rZtC27+f7/UbbKvz6dDUAFaIqf3+/ZG2pTwla8Ee0XoswRmb1eZL/XsZvRvlHeRQ2AHe/MDByX/ZR6HaGmJKoLzDf12GOCorBHJhxcCfSGWaTRt3GJSP+6JYOWXQUUnzmkbxpnXQ4/VZYsfqW/NfS3YqmV+G5gPVTn6U8PRzzKtnSBi1pxvjQgFifff2pZUzk/+A77yJyDo4idccedzMMweHBYRdJhbgId6SsiKLaWvE3L1s/YCqhr8Aqxp+83aoU79R7T9b0IBdmGjwe9+u6CNd/mUGBK8f74EZQAwCR4GJRc2wgHqscCl/fOTlC3K6OlCc+12nzaytyRbd2lICM2xDvyZDoDpxrhQBxn/2E8P3JLaljHgd3VWJAtjla0cf+ij1kJEpKYCD1IJLdqmgTsJIMKaSJK/L3u8UNb+qv3It8Ikbl1CxKYYmpQ7zwYFpuKeWntluLOXEF9+wt5TuuiW9h9IoH7oaco5WO2BlPIrR3K3fDbv91j/dVArD8RqDq7Pe8SvRuc8M/gHk/YFGvM9sZoOGrbwYYG7SJ2/2TwzUNJcQKSqnrs3V3Rwwt19Ef72KnDYBa57NTwn+ZuGAoGkJnBgtmfTCMiCdrC15Ft5HnITQONmfk5h8ncEouY8bN8HJZmQIMW+WoZMfF4avQx0UpfFfmerGR2cebWq+Lo2AaJQQb/yoYsYIAgi0VVouk8mMuiKGltVqYUi6yom9pGZM0ZBwRuCAzelkmvg7l5KA7E8fNpcCC7tyPLvTjPiD7lxCYbdjwBgAMMvwcT9+VVuUQNBvD27q2WiYHLmooUEFjch7ksnvKE3R4GgBnbocBH3co2Zw4820njTF2hl8tZvrjchOmD5AeRAADI/uZXvC4kE5lrPJzC/iYHGf1xkx+kBjBWq/tbszRRqI+um5LgAUT5O7Ivp1/PesZKDgioFRsJgAkO3EUZI0dGtQ+Eks0Mh0DwZkvKCy9IxNzvyeCvunXldCNCy35Y9yiZfsyn/2hWFLlOZdlgNRN7CeeTB85JW+fiqoWIcGciHB4LPb74d79WRrDHdHn3G/ezX0+i0q8G+OoKRy77XShP2iCQcQfzzmjJw3OgYvI2O6+sSwYuR1DD/rZJTd5194LbeYScRsz/DUfcbi0LuciGVCZrhbrevkoq4XlofqmJlVAGqkZv5eh9CkFA6ADwTmXHn29kfy6WA1dMFU0cQEVcsz/Laj5hqgoppoSjnimTTXJVO1eG+aCx0nqDrc9zWx+D428/vPJ0L1mYl6O0jcESNQ9XhmB80gXlmXyFIKZ3iMOMpt9FkuTr0SYHv0e8yhiQgu3NkLVkdaRMpY5LYZQP5MzDgmiKJ7wnbWfPTQl/GvhlS0Co/zrzWDz2NxmfDvYPDR65gVdIwyA6Bubdzu3x6VtxMskcVf2D9gG/Iiap+h43/9En1XMJtJnS4y1kwkAhA5+O8gEYoeXU3oBSNSsRU7B3CHTrcsI+lhHoaIwhhJnO/Fd7SDweA+H55QgsWmYDm3jdHSREWTn01sZjfVtEAYr+V9PnpBUk4Rr5FFfAKNsbYT3TW/bTQ4I2IsHjC1gvSjvwzB6TX1LiJ4mZeWN0fStTF2G7eTSPv4v7g1m7T7z1i2wFKhvACg7T7SINXckQ2v25GYI0xlxJcXo9/N3dloB8dq4kWhVKP/Gwi4S6X4BY+Kv8lC57PQ0CzVK9P4rZl6p6v48pIAb+zxBmFfcKrO+io0PtDTJSHzjDSAw1T8Yj/obll0WFaiPW3SCKE7f9ZVXMBHZwdhLmGHvkQ4dNmeBlHMpXXxArauXHAUOF8rhGoWQi+YCnKlTdXOf1olJ21SYz3g9xybjAQ+rys16pdoEkCW4uLEjnFJUoJkXs2AzBMKRB8DF18qGL9tSxhOnKrOIS6faW/9d2HADXSjqjgdtHq8I5PklyhFiqNFa+xfrFGNW0ZnscBRNIhZ7VRngOqfIiqStf5lW2k8rbPGJAU+7ddLcOY+VjZVFKnMpz+qKZyk/+pgM4o0bBmKul+Vcc4P+7rayUpJ/u2BXQtyRlUguqV7zI8Gd8079rEirdb1YqLdrl5gENeJ0dCV1kIAn+7Y8B8TmFfZTn1if8lTEfutR4pRHviIc3zIDKXr48cTfeORirw+VR7mMt6Xsd6r99wCJeqCSlSyjdoXf4tmc5N06ipE3tkbaljHv0e6+4iinBVq1PUFRvfRb+V/4WF8fpggbZKphhyZTldOCwjhUruuFx7GqwDaQuhuwz10fg5NOQ7h9uhdzx52rzMk0f5KC7jmPXawSbwD5A0oLiu3V3OAS7uGwDI0NDEBBeAuH8VfSQX8r6NMlgpsg8FWe2j+rAKkukPhZlnm+ZjUhcr2jLOlYY80ZGZtVX3FPMk+NR5cSbYrAI886PMR/u43LR+CLPLrKlQNwQnWh51j+bXaxfzKrf4P1yx79Hv8qGL6r0x5b61wmPnkRF2gJKbbH8UnTg0Umh6trmsG99/8lM+l7QdiJGLjctyO9DrGah5UPmdvbM8ygegh/6U02xOSt5SOAybLvAdGtDSJAlUyg4gNyxe4ChS606BYY445MFdw/HIUju1BCwdG/5+0mH+i2zZLjc8D0F9Xr3eBkhhfyA+S/ZZIw+lqScFAZr7/Aed9HizXduCB2SsUPCmE5y2Pp+/9LlihtHFL/5wGIHGaiLvhE7QfwZq1ugF+VDF+2pYx12rdJpdMy5OH8zPBCP3xLsy9qMPTQMgScHdbDXlrHoFfhCPVmMohVbpo2I3LGh2H2yVSbdm6n84rY4MQW/H4vHvWDp1vkFYGM/Vdjl+Ou6T4dIWkIfkDrraTNwI9lqn4NO+3PWsV72t2iPkYCRuhJr0nvnQa4vsnq6dGGSSVTyYlQ1CW2InG+lD9LG56pJ+1t0urb75Wj3rfrFuJrjEtAsYpjle/ZOJmtncmmo4KvVau7BJjuPHGdB37aljHv0e/ml4d5KuWwsTDdlmP7ncyW/XivaL6KOF5sZVdBZC9TnxeeqFkbaufjNzPi8iBOoGo5Y3B9G/wlYKY9dwT3rne8/NpPmgJIZ3WXhgYPdq1DhAhpNr8kCRBjALih9R+Zs+PW6WwOV6AvsPpelonp4zcVQWsn9IYM2BPI4+uOJXIKCBrOqKjc6TPgaOFPI0Ceit4qYURToHrfuji2ZXai17XAFX78RkvLKP5SbaeI72jxLmEa2NV2cJ2bjYSH6GevPMdJhptbIC7BRCRt/2HZ89dWRgtpboyi7WFDLrdkab8se/R7/KhgcRn1CfhERSR1T05l9NhIYx7ixXmFAJRif88iPY41n+yzWdry/Khi/bUsY9+j3+VDF+2pYx79Hv8qGL9tSxj36Pf5UMX7aljHv0e/yoYv21LGPfo9/lQxftqWMe/R7/Khi/bUsY990AD+/voMAAAAAAAAAAAAA2yulGhA29YAKO8pkPb+lop6OGgSHjgAAAADrbvXb2offP0YgMBmUrz34fWkxHknlZv1cupTlqATlSKUEJJhW6UvuuFeIB8UBn8Xy49GsvUkYd2vUFy9Epz+GsMxGpI6Fjb9rWKwf3k7+Z+UAfmqZVX1ECz3YT5HFWEoPrjojwWq+bJ8W6P7VKnSRRmsnu8O5S2G6jtZkslAbawP5qw7AqYMADs6CMDxOr4ngNdXUHhARplMnChxdPaKIkT8k14+COIBnmSXv5DsTatjPT6Zm+kt/sQF+B1lGzDIb2QF7Y3LPTU+Lm8xpAAAAAB/FiYlCm7QMOcVHctVtn1uwJ021mt5j6SAT8wwuvZ51EIQdJGJ5tjLxk5tj7mtTImQrYyJsGG8r6v7cWX5tTulu9bwusdA2u6e4Jt8A8gQ9mJ/4qGiTrB/VmsONJbd0bw5I4CmESxJX4CP/DEDRF38zMyBNbzsaLmleOzkPpFPnLQpDgrko99T72qiJG2mRdDRyqDmykUVEO8Kzva/15qBFHTKDbYLG2J8GgLM44Zv0t4getCZmIB0+z7w9UOwKFSHWGkkdYRwL/lMq1y/8yAAAAW5hIc1DKMDTvJ1j1U1kSUrJsEBGj5svFsadtFM9wbdIlVDFVryEK5SgXFXqOnW3N0YWa26UlIzYTqJBsW9SZ7iAwGZQwvocYWdYeHcVBhB1zSGGB68kD0hOAG9jRtf77Sug5Pi1639sIiNK3QZkEMg7cz2cDRqzX3EpAmnF1xuZ4VOkdLnwymhvNR6eKNnVeBavXdYB4/+VGD3BwaCywq7ZvfzYYbue7703rITaCEg0sPAqBdCN4bRdTOoiknZ/T0OYstKWhC9Gxj7dxVc99lNRBCIXq3jax/lYbChIFj6FS9FzZeYYAAAAdeY2mb4lJpNKPLx8hCwyOJp7BAmpk1ylB6eKr4cYkfZtZjHSo8nrFqYMYfrSnx3KsVp4MZ9ft2zUkgjAhIPjDMQWujCSn1RcDATkEvuyTqeYCAO7LRp4Vr4pdfq8brdAmYmqPsPtBhsdkEirdZYqcZbg2bggQtHWC79Ou6EeZbQqhXWPX+X3LA9wUZOJTjfUbstdmHZUAgut9j2+WzBRkHAqtGtilnQ+t1ZbHmXDo9oyh1PA2r+5YvFIyRR+xoxCGi+flzxRCTwJ7jrS1X4NsgU8PTnMLXg2ktu9Sgt4T75mAWqKCecUtG9yhgQAAABbPJZzgbMgs3y4LYTwqraJtIwM2VPCGGEIT+Atig3Ce1dZoCl9sSDGWb6tr42YZJUmuj4IuitTnMt2ps0r+JUQjZVAPDZbDNZ0kV+0/l2By6fMRdnjrelhBiqJF/poc3mSY4RgCA9OrnmpFcRPvnWtChvJci8X22QZBfCBe2RLONwLLBenhbBhzZ3E+69HmCGVtE+pZ5P6UVHv/vIvLdbzFFCBOzgRYVbBFuZC75855gvKodeKrdFSDvgtYS3ECGRUL3NYTXq7k3E2epJONHGwjhzugk+d34ob77MQZxjby2mPgqtix+mZ1dlyF+srxrX4AAABy/zFfdqjpB7H/uWKkYgcnap0NeNbN3AP/HIzIZztbuCILBc2dhzoWrBzbWEiY781kMijPw6uHtXIbd/FsXHTk96X6hyyjctT+Hky18ohVO7ap3e7760BT9P837XlmQwf76f06sf+nJjD78C2G3yBdbR5ApYg2hHm4Nxep/mX39blxHUNcQpjqCcVjiBVvBsevhyainjIzPDivPkDr77yxGuOAZDkT09dHvi90n7OkUVmMB9PD0IrXTLFAAcSmIV2MT67kZHkjoOYDAvNtOjHpYH1hI0YD6VKfLhbwNpJJsSLseD4QgcfKwuNKXvBBgVB9EP/5aBi7dex8vEvCeoJlRapZt0xAc6I7eqhCsJj/TWJ4YgtQZqYnEzDKqJ7HIYUEUqfvP1bdWXachXoqpWYH3m3t/4Ukzpz9scZ7UyDtbSwmT45Ye0CrJs0WFRfCsNVXtHjWlZ2rhRnf0l+tmtlk4DRnwXxJ+Td8Q4Fg5Y4KRr3vF+CtWsHZ2TuGpUpDI5YQo0PJSBAWJTHyfziVOZQM5ICvMkBax1N4FxeYekCr0brSXtGeVILuru5Ltvs0iXCnWHLHCq6Of7CfkROk/5fjDYfPHY6jW8zohOSN3+MOqKneZnLBUqaiSOlHfAfsvn+xcE+i19BNBvJH4b96zMohiMTDfYAfXK1U11eTuflwP9MuqzAG8beoG6DwDU6uwpoOjHMJfXx2jcAMAqdQmGDoEgAAAOX+Ce2vFZ+kNBEszSoAtnldyCisO6Hjclxl+7O5o+0ufNZw9NeSsVA4gL9TSKkx4CZBt7JgZA+XFSKNeQFD6IoVLkP9sX5opDFFa8acDplZKdI3fs5svKAe/q7zb09KdsRM4eOY97zqgpyOI9ccQ3EdDPxUv0jLe7N/DqEAYqIivUOY24wdKOg7rz8XMUaoQ5c5zB7EcJEWmFfkm0BW+H9TM2kdlE8oG5rweN9nysrSu+/rrPqzt/HEU7VGd84OWTDXNs9RqGWkrsy3jjPM4i1x5hKxFWykggoDH3V4LRSqxKHPLpEFjtGVGHabjtkNuRuN+QcT8l4kUle2c9enEBcDbEEtB2vndfuYbiF1yATgoGmE9y0Aj22cx6nb9qidFYi0uKHNOUdxg1tpun3QvUdPqFf/u+0jNiRq806ISMmoFeOh0DnhmGRQlejroatxnls1OSLZKKpv8Axy+tmreOqqOurvmcK2va9zZ5wqJJCcPq4mfN608iR6+3JnmehFvw8Q0YcA/vSdXVhKWjbtI7DSEF/2M06ZlFVBGGYn5g5e3nnWghdgvEfZcbBgKzf/aitQC78+OwCJd0NuNzxhEbeaodi7JHehrV0IPPGuUOBCXYvcXIF+8okbWVQu9p3yPqlUMsqSdvdid9glN8yo27YrUDjKFh33uyxWc5G861fyWkMUUoxideoSgVPSQdyxZ2J5/+d+ZFWy2n9baQbwR7QraqdKJqw2AJMA8UIPaQTomnaSI49uQ/IJ103YYjOUadKyblKdBelz/3/5T6ORpJY/E/ktG+IpbDd/NZDoAAAAAdm2fvMxrYaaOlJKSAiOiZADa4q3/AD3jKDFM3EYwykXHsnIIOd7qdUuiPM7+n+GB3vQrRsXCHNZbTrLWhHs4S4uYYHTFSdZ72oFjj0ac+tzlBfMg4pd7eGe2OGr8TB8DO/TCbsUSyHBfYHo4HVfcKl8Wj1TL8tTy8iP9QwMBmhDkVaoyp9wXnuXZPasrsv5Nc33aZeE9cGSiYXtii2H4OHhJ2Q2Bb7KruTNVNVmK0eX5BNDVOhCljOzUF+aUftRq8/vmopTQeQ3rWPtFQrTVqdV5NfHAuUOI0xiS97agnZvuo+X4iArW3ED1L8/t0SuXlqKJ/vLXCS2QG/5J9xl4iEikZjMF3sCmjiM4q6/K1lpu4MP5neaYiwXGimvOUC7cFJVIezDXCq5Mb6NL+abwmmC0/h4HnsrxBaw+cQqxR0DgdVVCldFZDophoqsxJeyMGQYpTzleGQ8q618k8o8JNVbgzuCx+Zud2mGK5ztCfoCYcrJIesP0pK8fBINkRMQRoPHk8IBLTmakB7vs7zHC/LIXN+7CkPmxqxl4c/TLUcnVCTdSl9ostEmnpLhIcLPHH9fkj9SaEzoFFx7bJCCiQ/De0fgV/pwPc1ujN/bTKoi0IEmagQAFq/97f/5EYOFJEwhoGpopC6WnokeVbu3lp3o0QbrlnFgNqCzKAWQN/kKoDTIqCJBwNt0N7zoOFt0zh7KeLReIyorgP5lkkFqHQOwBShq2vLbmJsgr4aAoqVU7uHLMjxyLP5KAnpf22Ij8dBLM1dN+H97khz4OhUaiJbJiM8peYD/hUvbFSmiwpoO5ZFFujw2mpejplHgG4lXz/N9R3FLPRvjMX6uzTq0PgwAAAAk9Dq3rhbuP4Gww9Vmf4ODn6qzqWOYUqFbkuU7qArlgtPhyq0jn4Wywe7itExhyPjdh74GJPUGc6IJ3qNVKjbTX9sKmvaaWmqViauxHkq6xfsD9auT/9Z4ASRSITuTgPmaIQQVIiwgUqliD8LpUvY0IHnsfty+Oglv/k/6c3+mr+sbMwXnUctCVv/AtaUWRtP+NNEyTiCq/ngqoik1go01AHDYExbJfqjx+XX3CnDThG4MRd18bP+69CQByCPTVcGWcdMWwjp3vlc2m2uB3sAuxxVwxvvsI2iqyYZkm6RRrOS6Voll2MxaApMs06IhxpkKtGf5HSSBGVSfRRMv9bRjphEL3M41UmSEScBo5NpcInB5T+534/WBNtfVONOrrEDuOoiHxQ/QwfY8jKNBnNCoMlLaX3pTxmpACPv7Bm8TTri93X2vNeF09gmJqHq6EfPpK9xY4ZS72ArDCShXwfG6Uy4fwJVwsBkg7+voKa1n5x30gFtkfbi3Mq77DhdKdu0RsHH0KFQHMTrVH2U2lWEWVoL9JwEnES0k3gQJhDlVwrvdS6k/G0AzkHz5hB9yix98d/9SzSIVwAR+qQd+gl3CJwx/J/HkooADWbmeqy5E4jX/gSB+tsogbw4SInUNVCC9M/dLY5URQD8ejHMLxiwL2aq9rretlos5V0EWJWRXYt6ghkt5OTn/mlqDkfe3CJ+3sRT++ti6H9EAE8ZTw4nfJ7VOUmqhXa5zwpvWoUKLprcDv92h+XLUMHz3W0MnYrAT0KEc6M82EDZRaiQ9pGHnJHIS6X/vb2lf+kQv3qiCGokFeQeJUhrA8sjrlbCDptNFvX63KBmSHMEcTZ2UzBk7Q4bw2O9Pvrnak7timyl/ZeCtzjlVJehexCit6/q7tD0KSMMh4g7BrhHpmi/KL/Ytbtip28P/nW2bkesaT3TtJtsvcv863JRpl+h/WPr8OvmQKAwEG7b5K2sXVurzBqOQ88MhxU51gVPQXYAAACj8upifoj/H8b4poYfhLL8v5tMqmuSco6qjwUUliBFsLgH3tyzNhIGBJeMVO24G+sfT4FeESOH5fE5xnFARJUj3WN0o531zzvlq5aOl2FVxIZTCf6010ht0xlliLpsXFvluZfep6NORB7Qryg+tp56N6pRrhc+U4nmjqeglYw69TWv2g8nDZfIYyLDHTUlA+ZLkvEMddghmH2jib0fpAIYUacxoWf54dsE6S1UD98iBlBLPBcnRxzBJPve133zupnGa9Sv/opPKlHk97XhZ70banzHgul6iYDX+W8WqdyTmyqkom76sZ2wcnpd1qI9W+qMadcPVHKhfxzdBJ09umD9V/DAcLtUhkFsTU25f2bEvBo3dnyq67q8Q8bIFXqOoyQcu7Y1hDiQlSShDcOttwl2gqGxeflds4XyJxde2lWB12vPMFj2UaNV1LSj14ubZ+lJQDj6gcFdyaCt5BVzFco1GYf1ZY5cGYv8KysPuBigtqVNrpxoSYLL7RkaFEVDfHo7UGS/vBfYAA3p7z9QTS4h6Iu5yhkZWVCczOsxPCFIRL3A33/xa6Oj2iyyY+uiTjcJKn4ftuz/ScLof75t6fYOTiFCOPYXxQFTNemlQXeDNTD70rylsM4Boj4ncsPB7Bh+tXMd+eRkJrdFFnoCQvs7twggn8hmz9Wef7v09BtxeVqNzrnMOnfxsqslYHAnojwDTtwoeKFYpedjSd7XPXoqh6CbzmyyJRXSnDoVmCw42Qi73LUmzEG995dsnc3mnLFB2x2pEWsCTmyjkiGuua9Ykp21tNsufXcQ1uwHimnzeUMJGID8tV76TSrsvlTmFTJ4nOfBPnhYDKw0P0+0xVnpcR/dxqVR2jtfuS6r1/s++uU6BtUd+DF9EYWjVEB8CdQ/gXXag/89JDmAYjM2DYg4PI3vYwSGPXMtwkzAPafFIBgIoA5hT1slKSjm/52hnWAEvSfZbCHpPFMz37rujTOrW28HOZ02Mt6igmgqMYlDMMLCTOWpekS9GPfczmo85Pgz7MkzS59XkzMCXdMLR8irURt8i5mB6sJGf2zXJr5sjKzbeV4lAlNIvJD5fdxoBQMmnGIrjdOoTW7LGRHG7wezQdoVaorP2yRYpJfPFec7/iIxuEVbg3JAeDYUviV6fH/ZLl8+pF63D+fvmwnNhMwen3iIjypVm7xWtW+y0raDlh7LHr/P4MPUa3papNmZ7VhtV5bZwm/uW7XyeDpo9YNHVh+oFUtabVBZT8+CZZ3hM3apOLuSZ3WlGcfkJ9m0KK4P44e9LXgnJJPg17p/nA6mlS/BlWyldMvo76fJxKHnkicoCm/FlsTLd1unyh13A+D40jN9dc5D0A58v0jNf+9yLSUhQroTuWexNLyEwKcFPjeM5TVwpu9UAAABb4jmp9PZVVpKonR91zqfNnbEMnovsT95osmRRtVJFltwxSJ/Gj8EFHXsH4Fzn2Sq2Fz1Yp+ACZdLVwR8yUab2agvZQxdLgM4VNsuJXkgA8uhD4g2yVVDA4vzzom0YqlMRO3g8ZiN/AEudmw4C0ml/v8N21mRpyAF985WNEsIhVEe/GfEDWNbU3OvKfLa/krgAZl7cmEScXHrj0j6XnHoLnjoozTZjaOE2PxJavKixGMFUWKde4F6CA8bYi11U+HtBhjFHL9J6aU2GdEUwK9OS0j+kGgVGKCEVCgpgB1QUl6Ah24SnaCSQfvKa6SS8KzUQaNZe0+6TAOb3nD5ZPKB4CjmH/CHsCHWf5zpI0NEmJjUGjJQyuMCo1LsCkktXyDpvGZUL2ZsRhpQvqvOE9RKJmCPBSDfvEgvWmCzWiY/nqB+JX/RJFinDiXL/7byF87+0b16opcKGRBppcWiWTv6/KplhVpxeadkALsZpt4KdWAwA5zQZq50bnIMHJQ3cf7UyeWYqDFh5nS6r5G7K5khRv6rNOxGlMwvcyf0/PzZOJdcG2qC5CQA6ESEebO1jjKFGVnmRQEWg6aZEtiVfTRqWxkhqLXL6Sdyh+ThlSzXCh7UqVo1/EM0sgkLAUvuADaCDJBzuyoQe1hfTex4p9D5/MjqPYW1kMwpVWB/79nXh8DLlRj0SkvqUHMRK6Sn3eRNyy/i8MMOzAkA0nPr7NdXTHBc4/T3EEbbXkHpJfTE4ky6EOS/ZK4my5vYXqIO2VRyWXIsYP+XF7gCi8bGNg3Zi0BMjZv3YH8PjwpvS7P93gmWbyj75G2+tzvRzvdVnK+QQtPkbquXJ2lQYQjSl38T0eI9IL+VkyAolad0KJyhPK361+aYfRar1+xnbKq17lmggDSJrzZK4yM2DjVpR6WzLLSHXGPZHnbUbvP7AtQZrw7simk0U4b8pixac2UQuFEmeoD1AXlBTAuAtuuR1oBHUDi4ORTqPjowKkfJ+jHI8zeBtBr42o9HxJiGWhYFQgbai+7JlYxcSc4cOKg9u5CCrooyBB+7TW8ljdf0q9Kf8hWoGC1hR3uQf85QQYCWiySdyMP9QEZo/JoqOKGpU9YW3VFgPEX9ZXAHsRgN+1iB8w69m1dcvpy2dilQT9Ot1bNVKZqylDbnVXw6PnETv0vViNEOnT9s5nYOkTxjQAABQrzVM42RRfkyamW2o7/OC4gfErZuHammZYUV1tfnXuHRGtpE/ZevGRItlo7S5rjT+ENSfi0c5NE5V5K/FlEGvSLShwWc1n7Q4GES/ev2s2oEraYVoKRapLbsGZ96QOaYir+u9PCvIY8rPSa+6mHbYPJLM+kTkD6IYUxNNexaTF372cj4er5DKOstnQ5bZa2uIrN8fq9qlf8vHgC93KAL6RtrhDkWI9bm8GxUyKPkyIB4vMyxl6OqNZM74SXisXFjjIb+Mj0IlUOOekyWt6ZV6orOZD0W5iMnHpyevuuxFDBC3/cOcjGdf6UWDwjHFGtQ2u87TNCfpW4HuJrpWMAzj6tapzdR/AXzwe4CHHYVPZB78SIKqmRjg5xmoUhcEHxGHFkPaVwJlVVLGzNDr4d8g8PJfOY7o4K06LQ26wqvCghx3d7mzQrUO2lAvhFVwScO/epdRV+XbQUUYfANQsHAuzcNQEhTBR9T9GhYgaVnyVi9AQZGJAWPJVR+AH32axtxqRrvo5x1a0SEYR6ZQ4tbyuUWrykZIpD4oez+jFknd/NufGVKu0L08PHkRoOcDxXnRd4pfvlXNAd0ejiVXOudlrJTShlGlRodj/w4aaeehL8qSU73ZH3m5DaVrsOdzkUUMzYcS7Qx6CKawI0x3cJuXGcP2o2Jrl0WPMNY3YQyddKuYkkoHEZZtA8lF37MvsA1wOQY/kXMmtJ15Z7wE+LZAbFvgpUgmhreBpga1K6BbNNz2ii2xtr9huDkhMMlvTNke0K08rueH8nS9an/d1Qlgeydx68OTwc+RflH0STy0Y4uvSc7e/U4pVpjyqIn52T7I3kPFX3ewQiPlnwGARpLN7TStkFGuFUJXu5Q0wOKSapt+BcutAAaSpQj6pyA/dntEdCHfod6gsznAxLbSeauRq1TJq6JeGG/IzkheRALZai/+yqmKCMO8BHF26ZRmMprFvwXGl93TuOED+99USnQSaWjY2lGdzWam7Cz58YOLGe+D2sIJdJhV8IjIR9iOgpFc3l1QB+Eve4rtVKKk/C8czLnrqi06CYB9DMZ5jkpN+aaTwI7tISI1fZcC2kJYh/nJDlE8cU+KfoLyLN/1iNhmenS6/o10OTr2amColTaq/JhNBYwMsIp5FV0gAABcfJHRGWoxVW9THBBISw+SfZ4Rnkr446wOduKG4+Z+b2mRExGDXLRM1gOipfGa/+ogu5lzasf0ercCshB38vTKe/cjBVGBKTFzIyUfGTdMI6S++O422k+er4cPHTdJZWtYeA2C1iAaYaYtO+ef4E0IVyT8WOAeVkFIfj5Qx55R62mR9QCt+GT3jUC7S4jm+gDYAkbuFPVbEObw97XoGTBe37kbinC8oxHO478aFvdTJWcR4aJ4qFV5i2eT+FFsS9l8OWDI+eqYWHojJmzhlx0W3sTdJOtnFlrBQlDqh+G6Op+wDMQQaIaJc25mewS4ItsnZfywXnuLc0IAdOp7oEOKTfso/wVThpIKgiwddusSFhtbskovNErOR25toS1GbOmf9+iixsW1lP3wxcimx8Cm5Ba9iVwi8miAf16ALGQXB/t93zwno0TGD2dubD4+CXVoBhbyqebT5tB03Iywy14XknpjWWVkOvFsulgC09B9PlMawbVLCmOLsxrAlzmsC0YU8raDR6mrVpEVEsEIdozmUfXRA92DAOm1S9K6FlLTYid8OFw0IP32ZJ+e9TQNHrfWMo0fdngvBGBojdKzbNkFKqOmwtT2qsUIcfYeV2dgji2ZwTU4DJM9LnqDFbPDp1d3HCNNtUGgX1ulsNEwr2wVHkVGR89RKW0m18A2tZqmERse7ae5pTQ9m3pq6GflwRGrwCgR8LOJdM7sbMH0vNVnhdTTP3cR0eqgEhyCmmM4M34YjB6X+sgiRN3QhgSTcTO5wZ7JCFQ2HwVNhUcAzxld9JWMj+2cWbn4MapVhYlbD/rzKVlP1yVFeDC3zn8r+Heoe+q5pNDVu3hwb3JQQ+oVspf/2Nz0glPya4pSMMmHU9gWRwtUnXhw7wm0PnqzS/BptDnLvj/7RoXD50jGyiZi3W8hoC2Oc2m2xyjHSZf8Ek/MozXnPPDuO0ce/m9c1eGUb1yr6o7An0Tt727FwSh1fKfYJWUsAijjOI90qo5TG/72mWmObgOLpuqZz2BFnRdX7W24eC92//TGw55BRG//CG1gQwdkLaTVxiuhzJISgAAR4+bFDWXlWQk53v2+1chpqNqfzFM9JUtJ7ZsK74gL/szvDF6oLnx549Jg7Wm5ulm/5lnFqGTDnjrrXTaxpBtKCkDtanRPvcVstuqR0EhlZcllQLgM9gv/jeOZZO0zs7ZFnGSwtCLIpyfNif5BHaPOxYJYp60QK3eaeVzBSfD2w08LCCVZPaZmqDMZfkcJLieEarzanTkU0pHfVZUNEeqBNEW29Oi0LIhLfzrw5jAMwoQf3Q/aTh/H36Z7QRqoWH2vT9NhsYlLZ4WUNghpoU1u2M5o4Gdk//C0fjbYXtplX+tT+FcCHz38ieHaQgO1VW4I74PE6uhmimycdoCGexwhr5cFC5QE4kThfMdvw5f2cQ7E26CufEyThYaH+pCXNrRt+iUc6mMp1+6g8+/bO2X3kL1lNB2lV15Br/gT67i0H0GB7kMd2Ade5EmGHRK52zz31/95hh6zb4zF1F2XiAJtxmtNbKAvQUuNoFLucBLBwmVAq4wqxiKjoUydOAUo6f1Qkg+IL2PlmgZi4BgmbDbqxcYKvt0/2gP9jQU7USuo3mTvSOSw/6in4DJvdY0BhXE5eeFQiQ9fT7/ZyNos2n6659Fxpfr2abWZkmR9nuOFRCRu5fdVu7OF4yZEb3e0jS4trdyFVtquRuv3g9MWaRGVGOyZ2gkmGtiPhKLEWX3JY1Do2rjpetxyKP+0HkWbk78DNkCu+HrZzrW6+PhiYgN1zd1PybNs38e0MYa1kP5B5rLWWo49VUtPj+aXwGm9cYPQpT9AovjJ7EiWeFHajZvET13ZMPdAy9f6XFjCaHslV2BnMQnBjRh0vcQESs4OONjhFqR/tPz8U6tvNAaZ5151BWg6CPuQ8Hu/61r1h9o2ZibLcIaw9DgggLueSTqeGcjXhfVrToNQO7uHiScF+ViU+dTuGKapiwqKu0U64HhRtC76DpWNvrp0Id1R8G5vkcpxykhMFbiYiSaoNxROUChvRozQMX1OOdlm2NnWAbki47e0ssfmJB4Z9KJjEUjiG8xMOPoeNUFd5Xs6QA+aexGT8VgAAAF4PBsg8mvpe8l/8t4rdmrToZzcDBKu3a5Z2HyRTNfHTDaK/UBpcMA/GgUMeHS9Rxm2X5G80YSXZaqbLm/gG4OjVgkimZ8L6JxPcgmc+7BcgimWzo3G8jR5n85Io6kxHoul9HgHFZ8/Ppdu0rTWOHED5KxXypjg34Ik77IDH0py7vCxBAVXYjW0AJ5+ClJ6wLkedqpY4rcQLu7nT8weM4pW3H5Xm4s2B8X4o0s5JuhP6qTq5o9Ua0OBFk5ZXDMdEGWuElSBdTMl6CNpc847w7nePsNAYWQDVW8uCNXxn15DpomaDOkkkjJ8qEmFDR81Me8XpbIjdJPbv/qmk+wau4hfA7kPEU1YIyrhjG4fmB9ZmyWlfaRrtUIAIMPC2cEuyVLZt8rIqccQniB8ff+VooaVyscoVVLGGNxt7VDnBQP2TkNekA0DXRyeFl4CdnwX5bUGoNo3blDtwyMOSkxAQcDrr2lEuxi8LHZWA+woq4BZKQqpW4TjJZCVnnL1zpfjlnxI5/1ZfaYvXB12py6BXoKj85iHN7Q2YAjgkcxTi07PsU/52rgAZp3C4MHFretZpmBuYpCKVYG55HoHtbe1UFW7Bg3GDfc6v2a7V9dy3cb/9495q72fF8EZxgPtUSlJ+iXJ1Uegvz3v4yY4EFtS9ohQHu2ZjEmx2Q1F/hoj36GjzwlsriMZFjwe1GH40LPCopfdexDCHWkeiBzUQcDiv39sNFwBKkYZsKeceylsoBTwF47H02wHHXI+eucTe1OuOPog2qqYyTEW6dtUD5j0Z208JgDuMJAzy8CzwZg1ysORttaR+qgljk8qbqpqbNgQUu2JenIILx3sVngTNljhAh3yREC5QUJYV7/DAFTQC4MUAsXqVXIIooP0yRrT+MwdeyZFC8ormAoVqi2HcdT5+qGs/J8DZSSngg+SpKln6fV5SIcw8coTi17voAAAC4pswqzi3e6oTGc8Qdsw2s9vxB6aJzGPgQcV+G0A2q8n4bSiVuw8V+BKhkx9vMGCgXyAqly/WBqdX7L4ysueUSo23+sHK6oNgtgd/RmBtLaC1EFiY0vj9IawyPfJh00th08c8ydbwfW0/4cBHHKMU6tBJnfHFcF6jVtIM2YHEAimDxzUsYJtcWxBz9ApxeLI8sE7OmdrnhFv5KKKclxYC/W+uuDdtQbB50+DSLvHdftcL+2o0EMI4klxtPes6Ghgqeecx7D4cLgHyk2QwxnVN7f9kZYtftrrhBSCv/M6vcmTRXMXOD6yEvgS5mbyPVqHZRCMF40jUQ9B7rxh39tTGR0VlxOnoDgPV5zqym2BbG5eLFPo7B+tD2Xgc1LApreyzm/lnX9FumzszTbuAgmKQ2Xp0crDUykLr78leLkDp1HSO8BQQhVUTzO6xCK+MGvtcqwFCEdRD1CA5/3rDIRPXdvOk4VFhy0qgZeSPa35A6W671pd+a3zYubYVwrVfSJkaBPc1VESEkqDmChplP8CVzMOZlo/wK1/UjfFSmrA4aR5H2J/6ZPp0jeG7k6i3Rh++AGla45Zo625u0DpLu9neRU9Ko1LGDz71DkSXSQd92RllU3DiNsSAgacdCNOWn/9zxNAGh30JeEbYHrNGAGR/WmDmAajhXbowo1aUWnoe4tnFLkmJ1NZnVshHoKkQGwxg34yfKi3nypJqcYqu1DwZX7TZUakVast2NIlxpldChZ16s/u3ifSaBRSW7mapCMSzITmfq6v5DuI6dSXkLj0I5rdBtITKUL0mJBa0FjZu5K94mSC3eWhtIaWYSHK/hAGxtxVJGxf0LtPlN7tyIJRFtUr27wyV9wVBmRendrD/K/19No4AACQlEGY8LoCTqUxoHTkfMLPpDaa1XAEz5T/Ii1UbsPv0oMNXfCEz/98wgQl4Df9p0UR1c2kC7bWQUnZ0TaIjbkJcV581bflEBtSxFf4XDZbGUei+mBSPSjBQzr9KXTepclLyB3QtnsN8QBozyG9948R2pTLeOhsscY0vehY3H1slJtQdSsO8J33H6gTnFvteqqPbHgusrT1EAMW4EgL3VyeWWSTFnyQHKugKDQPRaMLUldMvJ4L0xVp6593QSQ4CaF/x0ZesdtiSQIB6K4tQTb64fDslykWqUOejgpQaR29za1m+8tNVWbmJHLjrTtT8NIXw9dCpHMIoHOIR1ZSW6jVH6rmy5gV5x/z3RyfREITs5V1wA3yPf9nJcP4iA9ItFJpSKYqiAz7cks7ZsMgvgqZ2kNwRL+90EaVEcusX/kJuotEP0u6+QsugcZ64JDSNsrSPl1eSMbShcOxJsmPOzAEzCuqOp9dLq9bgVnoz/r8CkxYCE1Ung+g+PEXSDaa9DEO7JO4wYHgFLacEIVFZPP0G/OWclZyFgIzlMofF3QfzxYtU9prtFbb/o9qvcPn8nsLtBBRzZr18cuQGJS1ixIx+D5DxUWV2gBlL/FDk7fUzxhojYGHAZ79ep+WNP4y7oGYIQ71fhv4VEoGUkaor3QMGyocDXwo/xIf1LLJToSfiv92nQ0qX7rU8w6VrI22Os8nyTQaJmFZmmtAK0YuYdPk+5f33a31Ax+nZoUPR7wneEQn+ppYf5Dfl/IFEZycivcPuFZ7833FmbbU5psE+Z/agqQmdyYrcr7rz5MJdY1EDh1TQLy+SH5ILQmtorNRSQwrtIXAniqaDpVbeFwL5TmO8gNRfqUFVn7SiSxMi2kEczruM2kKSrRIafkT/F04FQBCIMYAACetrnAsgR5d5cn2BvV5gO4Tdmy8i+9gKHOCgbwCbtaQsYiJhN4rPqVChJqsGrw/9NM2tIZELn6WVpyJarIQkGYKhCWQTheh7vnJvwcL/n6EN8diOY6mKeRIYj1vi4W24e9zAfR5QEB7ZXrK//lEGbxCZrNkzJVAPDqq74uYf+MDybVMqBKeCSQB+/vQgbSn6SFrgcOmEHCTepqn2Izit3mGGdoNYx2/XmXd+o4f6Upb8nyhxRaiYdMIy8bnpXjLcvd+Ud1WVd17HzU9ihLRfh369LVbw2jm1Zh6KfezPjTIMdK65x+gH85R6n4lXJnFhzU9dixngmXi4mtC+poPQAZo2YbI7nwsvHtUYPy72PdMCgqhsYIN2Angl0+odRP0OohBdrlP0yKZyC3GK3PQ2pTm/L2nTm8UZOte5FYT9JskfcDDyDnwZlFr+Da5u59i43j++ti305GczbkT4vPMitXi3SsGBdYpJOs4lt5zmG1aufzYgcJggCpS+bRTjAnQUknwvNKplySy7Rl3U1/CbtL8bl3OUVAvoJXSB8okXToIe+fuZsw5Z4wKGCvnt3nX/GkgOwPCmrBuo7hDhthbZJlvNI+D67mR1DHrv6V4d0A4w5bXOCiRG3OuqLoFCFvYtS/8/WkvA6DqtC7WqVMb1nzgAUOhJanQwQvWpjLzme3MrZUJDNYMavL3Tmq3WRGsFe1ORg3plY4dckaz183T9wvVmxd7ksJqn4ABZoIcXBM0D0eYQXU+/OdwmXSWHgS+9ilw0tEz8MUj9I3lXsnWmEQXxe6IwQXzYuFE/y6kgwILXNV1YIm0lh2BI3cMVYb7kxsIlfEBmxRJWrm5zBhPgAAHXxfpRp2UDqZKlAxdPrDLWtKn23FFjkEmDrUfUJKsQ/7wD4+DT04C/RRlIKp+BuhmgNYxHsAxWJFB3WiHKoDdtyCF8dgQlA86XJX7DDnHLW4eJT9lEkKQSalx+utiBNVyF2MIqv0aytQZxuzBTmRKIHfpvT2bpBuj5brLNIlhP1RIfw/ylNHLyDH7E7AmRA4LIaxv8Mi6F6aBl38uesuWAGI5x7biQT9tKVEctKl3AljKA50sdmZrgaAZTDz0IBCMHcAApgTr6dkRsb8cznDUwBqHzLOsOQKJp93c8jJ9XXahrLXonux2EtQjRidAs7U6pp3UAvXzf8tyD473Q4M1XwsZVTE+eKwSgEmGqNdZUFgkKhQJF9la4tRts9lZrMIDZbF166SepNH7vXZgDgdkjwqsvypy/LfsVJ81lGQB0GkOALuwxm7DI0rIpwa0b7P4IGDXzsyWfkvglgv+h8g3W3MdrhKiUNJlcwVwDyuRvw+8lHOPbI6egbc+VIVRBPinZ0A/HgRR6H7E6Qu0jw8Sv55znlJ7w9HrQLP6zOy+0ZOH5HgiTxt5ID05mmgDAsM3DcX5pd226JWdJiCh84+EeLBwsecUSM42nAGq49S2mDEoFStJ/8QAkygHymScuhS6EQSOG0lVGaAhjjh9pjMgG/awChOVG/A8eCHvujZ3lvtu5IGggTgUiUgIHwSumv+O46otWCOGAdv8mkbTtROP9+lwRtLnLSGO11SwOdN6WqSswdUbF7IkMWZFQa7/BxBDGdKppqr8UWhMnJGEDcKNmmsI/GZFy/P08PsHFnn2EOSEZksZM3ej25fQoWeIm+rgej0CJ2oDBWlIQML+yJLgoGycM5IAAAALzV2EXl+rmSVcIkmIXARGGREpNH/ZQoeAhXKFae7ylB24chP5i5+qV2VZ/Jf0Q881EmFsZc2UaVmxABRVxe/Gv2FMfFwN1H8+L7bdDn84gaeU0v6p+4a0iKxUSJ0ydCPKMB+VT+UMWziFMlkQfjXYjGz4OmaQK4S0fVfOcSwqlEkfuw7vor40AiIl8IroGjOGq+JsBMwy7QFwjHb5HJy2VPpCWZnlbSM+p0BA0kqr2bgLUVLVUyV0jLTFvtVo50M5fupZRgZajeMNIx/cBp9KUeoKQ7IHShsT144BQN8H38nQrHyzaFVaFrRnF6XfxFobvjrbKWBiw0f1E8np+ApnTtyqsLwK2UvTKsRSU4OLsggXbP1+fkf7AFdAEEOyIxQaK4hjAyoKvMspGSfiGiRYszScW5joAp8uHdWW69D+StqlQuvfeW96iHIAj0QxFWjN+1UW6O++X0BNAgnNMTfDCfO2rBwmFxpG/JnFjC6BBtQMSDBuzpg/YYj6SFzZW3g1S4CMjfA02oZwBpLTOIghcHqbC6/8WinSEldaMREOh5kxCRR9pxtM78mugP539zhjW5Ao8E9VuqoqLHvQWQpUy1kaF7uZEYPBSFtEP7lMclBRQVjAYKAVGIuPCkWAA4bn5TgsFwfyk7fCCyKPfC2Rny8Opr6W8za6IppNGAHiX1PVJpNjlHvxfL0ad9grFXQp+W78oORPI/xD7JReFt7d0kHFPex+rIiELAUwIGK0NaJRHVxSWW3G7sB8l1wkAAABaOFjNhc8rMoiDonbW1mYatxF0cKjeGJ6kcaYDsO5OyTZL5FxTveM9EIZ00Jz6VPNkvoI5T0IBsk3QB5cOSjCvoedDrlRPcR4tKdzf1eoE0cMMilQ91AxT6KsNEgaGHe+qMiIoi3aHVG9orx+gN5g7D6VOSGjx4BWoZytmM6IwJQjRxZdRN+kJnyIpfXpLMGE4zwRCub2OkMDdq6VOPovCGnaWaeBAJblpqLfdG7hPfJIDs0Jy7pvqnCra0mrhLEW1k1xV2T250Kqv9OQm1eQgrMX4lRBn9BTmeNzcz9DDLuMoSqBN/tcPiDY5Auu6mpW4OXTYfTUOLJ22AiVVYROJt/cpTSAQqaIkEbEngry3G7EnFd7vPlMNRyR3/e8QO1t13mbTcUil80foad4y2glLwuPI3Qh4HKeTZklToiLNQAXxjy//xrBI2RC4JWI39n+mCx+D9xNoRbttc1pcrSr1wggtSCb+lQ6eRGKBaZgVfPpWe41nQDMcyBxOybvyHE8Y7luv5CYdV40mb3RNnM9SD8E22pFbDl6sq1w9Qw7+YkQwiYfyA/kWbpM4Kg87qnnFT/XSCmNO8KcXb6xw9wYQxmUXMsuWLZBTnWgB+En+go4ebP0bgqzJ985I2AVrTDm2wPWudv8idQnUfFoVw4coXFw499hEY99+1xCCFvdgBiZeBbEeDuCUBxaiDJ9CPqymwN2NkUG3XfCFd4C6DbAfY/d4TiEy2PHRued2rVsGbkQG438B0V0BgAd/cnez1YAAAAp7lzJmhtQuv4ZviY9AZ4wJT/uH7dy0q5xZhSYvkGasgWB8NSr31LL8A4AygJKinjiAvwUBRnIobLsOCysxbJ1ibE/X1aE8BSB7KTeD2fNtha+WDshi/xyGn23bB0N54DJm0WNXBAY7vIPlXdr+WIwExKd0Tm2uuc4fz/1iuC9P/Q25qxU9d0E+V1M5N5xyTr1XRKI+TnwBnKxUP37I/y9l85etGo75+YRgafW0Rg23r4xRaSnzzJ0EwZPuQoA5iHvt5b/rKgR9bYarzfY30RlF1xg65UsJUy9Meg2dHo56+jDvCBgXEK52U+XtgKGayA5cTgBiygmJegxW4SayARX9EQhK/a8mEzN6+Ghd15U6wuBQl3pVPi3NK600h8JmUzc8St8hR0TlRZ2EpIwDAyPDiHfI+K4rXKcIWF5qydhBakWdoYse1Vw9fYGKSa0C1fmIljAC6YwA1PJNA4lPMTy02pzZzFdQ8leQ3SdetHd2hxCOJ7U5pvCKsnp4YMTYZJr+GTAEcx5D/m8OTw2eV043vX9hALNH4XkfTNl02+vo0yy1A25seAZ2F9yvabQLfUhaWQHhApWsHNQ2lFxk8OPp+vFa/lBvftgb59l/NsdYutzifelGfRfjOOenWo5FB4r3OTB2j1gqg//4Ql11wMSnhc8KAkZh4fCyvpgWNesNCWHWAAAAOdtk7Pd/dhT6S2hZRfjALmnch+TBW7u3GJjkHNJP+IMyO0pJXlN2ohX3LMDAPb6F0/GLJ97+Vj3f86WcDmEvbuN7RJ9aUcyA3RWvDs5HM8tTZJokRo2QM6lvNsv+6CYXGVe7+OZHG1QeJFjYLEHm8nJ/L94dBl6AzOg5AWJQb4R1XvGkGRkHzvly+jjLvDrb84Bs0RrwX42DuMDesbWRFfwoRB9xJ9TEzVQruwyNMIQ1jOzcOBZPZZnPFHfzNkqeMC6Ry1qX6Q6ypdVa5wsWTxksbe6H0gvUV6z+4gyx7cUK+MfF3E3xC5aa5gnpspMNvqSP2WG6VpU4eN0fDPYgp9yqYaJFtDUkOqPVW/qUh9ad6z+yMAd7mZnUHnVe+Md/sx9Mb3gJw60w/dfHrXngKaKwmPKMdqmMyVgbdM+zE063sX9GWVGuuoEf0jePdMcOqhprfqA5GAw7kkUWX3V8yWDa8E5m0UpCCftVroGnwNTPQj1+XZ0eCnFn5MWbEFvld2XZFb8XSpPh6F+VmdD0LxhrAj6/9HzdRaZjQLepjO+iPAjxaLImHCIcGrh0ubwTxEt5aiiATxNwQXUmiDx4p8aNdvMO0mwY3tQ8lXM1ExJnAFhiMQ9SS5whVWVsNH2cianP87BRkNNOf4S2ziXo7wtfssvU52SlKe3nrt6HVsKRHHoVWY+Z7RYMWALtP+D9Hs7oxrmlmyD4PGZLgg9imtaPXb3PABkJSaL8aLcmgMLHYeiJkYuiIrm0oQ8hMvGRsoAAAAE6X1LhnRVvtP/KLg+zBMcuZXQDz/zuSkYXDCAzLHUsj3RGwlNbILz5Ju7bXUkavyjvCRpw2917G+Xnr9S5f3kEP9SWjc5briTTIQWX3AOAtx4iAYwC7HIK4z23EXYxo69aPPOOPHSahpizQTI2PFegJnmnWdCi6GKPAZ4vPXIhXCjhRzWDkl7ur4yEGX0UEd310BD8uL+rWGSDaxKcXZXdb3/HDUK3GkRDhs5fQI7DROVFggmaGkieD1WLJHbTiCCa/lRMD9dRkMhjk73UypdO4vLOLEkCaqVSpF9GQOL6CzzuhJwhMXVSIllXcfUGdJFGxmw3uQ3nLASUbpNCOuUNXGLUCT7usmSYft7fVMEPcmhY2rstqD+/G9uKpM8ylqGosTA6iTLa5HnUkCTC3df7OCCUznjhliMer6lckI0pLifScc9HrOcsjLsGE5UmGLetfU/E19EfS1sIup9A954wgj9eu3VDJitkfIKQj2Uq5XjrgXk85jX5cz8Yvczn9APdb/ejHXUkJnt3msb0qbH3dOH3woxd6wlPfs8gagChcoAndwIPYC70j6/2OBHDnNeXlSbR7OrIaZtylq92R059yYBP7PfAS6Psp3y2Xu/rScbULtnUbtrc3GMBHKUsHlKemsyFagVALcuWuDrU5ooFc1ls0QfHuGA61RjQ3qhd6dpH5zOVp5dKboYWX1s47R/h9paaZSKReLo+ux/xB6eGRze2IS9YWSLYUxiEt74zidM+8iex6b8Dr9tM/zdjiXoxbv479YGX9lo49hK0ysh84ZZ2Jxu2C/oWxbxO2QAAAIsidzh1iyLMR88RbCeQWE4c01k/6o68tttl6iYUcdWSYMLvh+SOOZjLeBaF5gCRQI5uayPJQjsZO64sO1RA6uQncQLf+loM4UzixVS4IhcOKGc8pUMD2ucfhSk5YcgTUsApu3Z9/G+NTRXo6Ec/P5uz8OVuceNfsxfyK+2n9WXJ/kM7HEtcZclUyQy1Oa7r2qESmDNy7HI0llmQwUq4LMAwr/a4Ne6746ZerGt4qDwPx9cX/jQl42Dr8oDyl+Cuu+0nDtdmc+RWkIV6xe/w849Bgs5ItHczvZn08ud3plmyrrHWFV3lfDSc1xFO9G+t129s12TM1OuntSxL2g9Kvi+jRGhNCqBdHz+dBh651H0b817quLyXgB7/yCYlL9N/Ue+ve/KXtha2b0V/XraFRl5gI2qUDIAv9KdRU25VDXt/FbAUYS+UqQMyDKA8Jo1T6WBBolkirurNlBX1qw478HL0YLDj4H3vDge3r5dOU/yyA0Z+7DHA4GY7gYmfEzP4rbcOs6MIx0CIbcxDebzJ6caA4o397IqVjp/6J0H+dlyCFASQ1KA+NG3Bc1s9QgC16dT8EI3qVykG+uYoY5Ae5yJo5F/gS7rzampuH6opfmUkpbAAbjFanIJ+Phu94bOBX1pm+MpwOUU15Z05o9xAsmvcuA1Hb91nqohaL9cUozWBAjAkuauFmDNywDpmiTF4JQkL8p4NaO+HjK1/hEiuJ5NxnXjhZ2Rwb6eQuZMa5GtGDJBdUEru0uEHS2hTpYlUMnQnCYTNbNiKCJObJ7ov8Qgn4+ft7mX8Ich4RJetTDv4UpI+lHpgE9c6Bx8VxcHetmtKL0GSVd+7gL+TJXcKwmOF7IOsg5vvlpNPIbQV3ucGw4mYx7p+PS5KUc1FlSz1C3j9L6NH+Do1UwT/d/Jw6WwVbh4bzHz9bMLdlnkMO5t/MpsFgsBXRV6ZomA5rmp7nMj+BJEYQr1kwAAAAAAAkqY8uaPm1ggNpGnXfbrcnUkZh+LZAdSSOH0aEnJEtpMnG5b4yAmXiaXKEAmCru33UD2iY0Vz8suoqrU6mNMahfvVbMZ7HlgUDRw9aiFRwc/5WQr/wkAJFYSdF+l5zMj/XJQntGIZAys+evgn1t/89DR68Yv8Y1FJI5THwof8hCMjV869WTBhyqQgJ4KY0CtekJyINY0XWxUfNWlYVIwcYsVuP3ZETg85NNCKZVSAkM6/ULHAQ6JiC6O8qjJkLh6hNU5NO2d6QB88K4WVozxWwM6kpQ7m2DyP+OB+V0wkrs26krWFjk5QFM4gz4w7FimUEcJfQxyYOOzhGDn4wuIGR/KAcnI0XxBeywBaVcDvEX08jgxiAQ4HCJOWffnCHvXhOUAq73QfR6aNQlkOgoj/jtSiXiKgaQWAemIDpVOBLWEyw2pxi+FNmDdV+dlhxZZ32r8yX/CuM3wB0+WF4tPNrBYHaJkRADNZZFC+BmmifHuhxsQYSuICtfT/xhhQmp47/YuheOU+bjrCyfsKkl4gKV5pOX0/sgc2VKSgHQSE2vhYTxtI2GShx1jVt6fJ3SrvoYb4pytD+n/FwPB8TPRjVJXGG9aKWcZuroeAhdZw8rz/Z995NAq1gTTnTDjy4NFIWjRIyiNPLFvdbMT5s1rJVWTXtswoC2TWL8FryHhQwywYlFHt+KR5/aAI0C0sdrFc9huIjYQOEruAFNMkUSuq4VxNYEHP2EWOwaAf8IIHWcIF8FxgBLQF7mHLwvHrjZtjJtnQwLBcDnuIkfoafK+tNfKVw3byJOyPqr3IANfy23YwvQDn3xqTwH5bd62+AgNwRISdDNVHLuj1OXbtVCswA8kGKAlBsK/hmXMNOZxLlrC0zzyX5DDgxUwmZa8pYeg2Z1RbGPh5VvbWT8A8pPteL07HKkhwzicL0hF6FYUkk8SQpzmZLEclbeOLwEye12+ATxQVl7yNuaIAAAAAC+vADS6Uxm13ACp3ZBx26pEhTW2KutH7qaROu/RAIgjgA8uGuIaDjX9y4aC8qSUvlfGs4nM9D9wuCblkjEAExA63ay40wrN71OaSuCzM/0HsVi6ZXOGEaT4a4OYfT+t7AQ5kRdesyRRgNPkcxiAiy1GzEQmWV4KKRXlOQPF2hCSk9+vneAK5cNeAMjBJ0ZtEspwGM7izS67VgEdyrrljBQfZslhohM4jy2n7pbhaAjlWeJg5sOCd+eFb9+D8fP5rZu1s5em+Xd/1CPD6VoWXZvNghc7dRP9v5ozZJxi8wTV2Rv20q3ZZAJIiH+TMVvTh57g2p/tx1K2lU+vCOEKrjcjBTHPotn3xX9lAqUdPCFQ00WxsfRyg6GSBB5as/x3LSf/qC9s+7oov/ljPndv2PZ0FJ2QhCTq9rjO9599mwI+VMvPd3rPZv7324xLC99NyA6H6hVC16jjRpmgTMM0EOueHm73eAyhKhUxBD3hrRxeEPzi67E4HKhLmzuQPBNq2e4E/k0BUTyEQQzHR8aAbhGgrOIo+rsBrhS857JK1QXwtJ66DAbyxJuCob+9Xtvl6kuEkVrtK+ypIkFQGYL2z6QxtN5vkaAP7j0qxZiHjtAmPdDk3AinbZU5zhM5817Qb7NK1Eh6EBOnhMal3j9HYR/yXYdlontaTDrCaULzeXCp2zE3iu2tDGtCRurwvF5xxdhYR9uZL0XV8LmIq9X8ZrJDtUK49NCuvhBnq93Pckrq4fCJP3xN3/6ydvVTawU+XtAc0xAiT0M1pObUlv7MycLoG2l62B8ScQnXrOotOgkuvrIFQtMpLxYWBcGRmPAuXSyQjrwnnjinPLFl/h8F+BHSk9wiVbmT+MDeZHsh32uz0vTxmP9dhHmdw6q20ki+JR8fUeCuP7NypiUYuGT3lcPVxN0vVdHm/h1mhRzdZjHwfawSra4lQ+ZQ5l5wgMJW3itGo+G0/4BAkjAD02Sd8e+wv2/iQJlKHlTtnLwm105lP1NzP7HgnxCJ1gAAAAAAAAAAElfd3vTRHMBmFckfuGwQ6KYCDn9xxu6fcgvXxheRlxZbGwSpPLMCAoFeiyBFtUIbBq0/g50NrvhVqSf0ia7iTb5uOzQwyk7JrHSZsdgshs/+6Sd4pJcXkM7J+tgifKa3z771QVai2p1ncsLIppg5xkBwV/hL4EvsKuBXg6nxv1aFC5rnmCQbPHOd/qeE+uANEu5baDMSn2cLXMWBFqh00t9zTqkKKgMk02gNPCvjTPeLTrmwzAapdRvdOoKnrJJ4qeAYuYRonSgbPna4ewIuGd2e717eSd5YoPC5CHGj+bEaVYtx1iVVXvUlQx5GkSo+w1et6kH0ELs1gqAAA6zrMfS9DcChYaTMChozFPKK+HZ4rMq8cY7l2jsD6D9qZJCXvEZutqCWUWOrHl88vkvJaIS7w12pcWasDmg2W4spB54eqjR7ytNz5AxsJwOy9ZCZnbQHPKJmHwFbeaX4lh9a1gmyTJ8pFOcD4A+/4c6xQs1XBtoKBbv37cgJ0DvdhXbtXvSN9Y000BOXh3uhngBrqIKbVDWJuiiHBha++nlX+CA2ETuCQXzHYuGTchLFCTJ59H9x2i+dtybjfyzbO5Vim8fhDOqcIW3AVLWZsC4J47KrHBCcA0JjMPyZXhlgaXAVjWoQHTNgvgd9vmRm1pbJlkNZOG799qB8dgARtVIg22aA15BV+Uma0TIgi6J72KzQ8DYDzJJheZivTo17TFA70dYIiFPeO/cDp5jsPdGmYW8+LI1llAoVzp8CEXyuS8IVI9QAn1AIr15JkA0bP14ytSg8XJGZW2QaKkn/pLijUwfKAwchU3opfKUHEiGPdhaIpwGEzBKkGnCSMFyvpAugNjY3T5jUPaMtJ24FU8mr2bLgcvoEdd8jmqJSQi/QbTyiryFiTkIv9Y1g7zmaXZryDV/UE+6irlLWPCGgBFbf7BI0juf7v2lQbG38kjHVb/cKoYG2uuXCVB8DR+5FfudVvkrMUxsjrfExlkd3I89Sp7Wkoybw0XEz2CXw2J8ibhtvI8RAxbup9YwSBelcp3qz0lDmUX4aC+HQDMhtGyAUS2ammM+c/US2QWEokoTuaOMpbq2OFoQuZd4lLmY9fP8wueXuh71v0cAH+6PmMlE4GanAxCFJvEsPSSCWfCXfsTLEI4VJ7LptyBY3Ii6gtzutSGgmaqfRjcn7wsleNrz1FxAfF1l//Y49X60/2GKt8dgT2b0uQb0PYK2L/E3lf3nhc4MArZpBe/iHOKq5n92L565YfNSn26XCYOwpKIcMIDhYZUFOG1SKWyyIdzPsxlzZbxbb+h9Vi/BkHFs6M9iCxP1wCX/2evtaYDOFSpa12D/0ODS5yirYi5cMsY5LpluYjQgmgWWdzl9STdfpUWrWhfJXIv9MH47Lu/tDYxH+YxemmM4Zto8EFcDfaySLutNbiFrfU3+p+nqKE0ryIcdLqHBdgcrjTlhMQE6Qz4EE3Fj8b/mlnwRsqn1X7aiIZfhDpauOXu4HIMF9dgYazcooz96UYlOrOSfugNWMvQ5mRP4avV4GQJrQC+Pwh6xe6Q1HFCaNFcLAMkSaw/6Zq8lBUE02kwLCOpJ3dWSFcjoZQ86cG0iW7ek0uOchEmxJ+IOh2sGfsnWEuVEBsR9b0mRd2zORxRS3Sq1SJAQD+YFs+vgkxRdkiJUfa4EiT0rFvzEKr3GtvCO3NCeIZkzSsHGXyo885VAXdRtN13n6DiaSwWZnRZuaA+2HP3mwIA0QURqBKIeA1fh3g35Jttog/AxLrpQviieMaXWeDv6mMF5deKstY9Qh7LPC1LeUhtYD2A2NuPwdPOGrumKMTO2+M3HDK+PNlOQTKSqmAldtmVnc0O/wLC6xz3beVsFEZVB+0R+51mkkstBheeQt4FPiyNpexm8QNmMJHxWTjl7Zz/eKRTVTx6b/a1Wu8r0T0TAAUJZrSfTtQpdqNJG2piqxlhM+VfBtwy+E7dvXYgjr02jX3mKQab6Nlerb+ftaFMK+PeTdBZrFBYV/ZFJszZ5ClwthKvJLJTKvhBc+oVeG80sNZE96azkEoxuT8zeDSOcgPBI0KkQOm/r69OwZpkVm6honaoL/skv4ifll2wfs1rk9fwO9n3k/P1MbSc2S+DvnlbRFirUQhUaKDVit9uoVI8vG/nNN7eT4oumYr/AS4b3OgAQ9NoD4EChFRtqzieGHPCiDnxnNDmNkYTBJP+DTG2w0IRcLko0WlQaelOHyH0rJ7JUZpWyjll9HGZdmHJ30zWP/TrUVMn9Al+HoBXrXzAwIAsgym0r/WoBpcfUbl+PIJ0+BRKXhk+mrAAXxXCNOJj2MsF2sas+9AJKM1vjZC3OH6HaeqNxMURzFFzxwFP8slQj0DEtA6MJvqADhoxaCvsS0HpF1Vg7sSxVGHBsU+uyXH6zimG0tVaZSaznZEoVYrQIucn1gTHUv5dgWAj/UGP4T7wir4puI1LX9tJmgHSJuw7kiZaQd9paVyaTi+vXmaAv1ESJ2jM29HuLxXPyiCKy9xqMcQHDu5jdMeGEyPHi8yWO6cSSRYjYsrpgMnT6zXShSgq4hth0y4R0L5n/JsdLdEax258yPNLCyvt3i6i7l9uXVBwWQltlTKWlxxfhgaMam3A9C39gFjz/we21HYT4SHVQGQSbYdpulM67DMamWEW/K3T6zXS3nt169lvN8XP7m+cjjzWm2+fG/2QcIo8eBwHmbXp8x4bKPfjwnBawOD/sA2Xl62trsLXEanvtFMFi1ERdQ5enWwlLdWnFc4aj7cGKtPvXxeeV74CscBTvpbIvVnXZ24ijQyf8QiVqw7xTvpYxcN8W3gKcHaex1IEaNrWHoUaNTL7cuemJUDcJFYxmv7hcyqT816o/jVBsiY71oxS2IfuMlSu5xfaiLY5C/Z6FpAWiX1Qg9Rh/U7ESxkRUMFy9hAdXliHY8DbE+EEKEwmx6m2aXpNQk2oCVa8fPBxgnnQH55iU9DgrR3uqgP9sMWu6BiQ74o9aT00Z26Pkoxq8W7FIAXyS4lq8+6/tajQzcs1lpEkwGCCPNurVYHiyWyzIWCAT3/N4AB/wERwYQYZrYy0uIx5NiZTS5EOfO5nWn3EQVfTpRAPQOx3GurpMZcRq05t3FiyiIA7GcxIRYVlA0TJWDLOfzdKV5Xkv9rge/ULFaSiD6COlF6x418iKvxfgQCO5U6d1UbHDdfFNoZQLDIFdKxc/Kjq9TPwDub21P4HwBuKqJYu5dKsCkLOfrd2/eEDvhM2MkQR6fIlkS9l05SbSTaX123h9yyKxVJrG7jBX6kjGTKZueLLg1Lu537nddbwvSafkVrDFs3ARbr5XKCdxnL3IV68DVI/N1u4nJ2cIvCpLGN8ggEdv4Vhuyr/wmKFcVGZIcQcbyenc77oPVmINdTy7KChuDjCPQy28fNL7KLvxlOmtsSLC0GHtpvLYhagP6vFwNHP8VWZzXV+uplhET99rFE5pSzc0URVKtQmkY0uB1EknbFm3MBusekKrKeFYST4GwEVYjkwuBDxiXetrOzqwyCL11lI5ZehTpSFiyUmlj0qN8Qkm2nR4plEYA4iH3mWGK9wqY8wEI4DRFzvSYQpG6Zlw/gyiLxLasPYXIBTlrKKyQ8xqqDyKjj+D8C87fLAnzIyuaYP9fKVJIUtPRpDS4N1Gw166sdE3HVlilkMlCXj9IjEPKXh87GcarCL1m4aFDRjBm1aetFmLMRuTVjPNaTsted4/PrLT1/+48FMANji3jWTJn514BzF3BawI9gG2xAgai5iarU1mVdFebcc1xj7/xR93lM7nncAJqYmG6PANeIKzheRqlhcMBpYov7GIpXMG1M0//6CwE8lTSpAcDIFibmXTsXxwB6Y2dBlRPArntFa8vcab5AAgc4CzFb6G2oDPbeIBp3+IBs86RzHwSS0ZvD2hxxfLBV/tSSy6hyu+Ix6u3gDIGRNUQvoSPXJ1q0FrjikHONmGTjwYY2NIKZDr1mNxIW4TljBeGfO+6oyudHiXuEK7gH/yJ0mHrxcEIqxtbLm0up8zKLdjNUeLu0Zvk2qGXvvaUyh1ZMVoao4uZLOJm7LDvi9biFSUQLssszYnIpj3A+hwKXnJxQ4cvU/ynJC8luNcSTFClAxAdgfJW7deSZ47DPXbIeoxbXwgkci2CRjb8LztjcJC3QlqEpe4TpryRxKbfFuTrek1WZzZ/McfdGx34mV6ugsKvpegx1I/WxgJqf/tmXecS1Bq2XqjpfnRmqbiXyD3mceHwkzqI5Gjk9IQ31+KFWAAZwDgjsL4XI11+o4jt56FCTjPWlnNraZnojrWsJVrlFd2Gnm0vI3wBkM2oKksE5jRFXFlxUHA5IndT5FhoEQRfEPs60r+6Ih4MvBdlan+D5B/s6BhvXCIcPXkQVmM2wBpwMWDU1iLo4Nn3v+O+QEg2kAq+i3Nk8v5M+QflWRc9edhJIKSgFYNYAoIBcmjsrHrvt1el/bLzKpzGz2DFkE1T9SmaWSPh4ZAT9n+A/TkDQmCmxFIKWEAXsBDi6NrrM6bMxNY2Ux7k3BussJLtLzPQ33NROUvGU5OWln1oEBVP844wcQJkSnPDp4g6+U6yrlWz5prc0TZWCtOLmjkDLL8XGYlXMPErEFO9af9L99cKsYCFOKxNDhDLzUpejqEAc2b3yOf9lhJYPwAA9dAegah3miE6VZB0p6DsyM3hDIORDnRAQF44koMC3IkNs8h2WwmYdu/rsn1ZtgeOcjUA1ycMTK2LgVzvTefa6Iq2DLM2cvNM1OKyuGAHWNvHgcMME8J9PR3Egs42CY852z/W84G3twl8U7WvI4Tgg9gxnD8GBemVrmVAngPht7/yGkc+KMhAj7BnHXAfN4OzcUzxW+GINoW1oWonLfEIipU8hHLlQSBp3UWcZ39XJsNNwIULW0ENKd9gMPDdFsmhYgAiFzwgGKJqnkJ8T9ljHMAA0QbznDP39WOt5k9Y8ZhMV838ebr2xuSJkNFgHGK2alvVYFIZxIuWnoDAK8wupsVp3zTS8DE0QrcM3w8mxOXp3LbkK1Ffyzkt3AmqpT6YMdZs9ZRxgK9C5h5Xgn8RK8+iv3pNdwWidv7QVDF/IsNBnK7J2m1JfOIxHABVALZmMX28GRQYT0uiX/rTCl7Rkr1vYt4Lb9wezc1e3w5RR+hijpBjajk+K/FEbqD4V1qSxSnk+rVrJmoO1Avm71+mc4jn/MYvzB4DM4xQE1LYvJpXWU3pTBZ4uGZXmYrtfcvs/GmPE62SDzC8hSSX3vwPJ3OwYBpKMiaIlHGlxmLNi8wwg9lGzw3NYC4x8TaOLCWGN/uoZGzO81s12B6f64kYwi5QtCRqS2ZIlzL1VEjdEanyLzRsxU5d6NDHFpMWyf6p+MTl3YEOp5YTaI45WseeT5CImS7V3Un07F2qFV5k0+1AGbe35UbXlXQ2T7jAbNfp5PjA0WfQWqk4qfSidM3kxrachEF841iqS9cdp9Cu4nJwk+MnI4x7fzO2rjTxaUUTL/FE+lriklkHDomAlxjleqpwCWKiOVax5HDoXcuxeYEOAGEWNH7jsFUxdy0Dst0mmVFDgBlqXmSpVH6jDt8lSOCe0YbeMp4mWGjbveCsVQXRp7QBDy6rBnrH73dH37XM3me8cFHgshXVuanP4TmZ+UvE1vDpxXjKQ/ySLt0mava4GpYfoOCvofRoFJed/IaYH2dduKX/xjLjL5DZVyYe3emhn2FQ2/79VG7KCpqQOm1qQpEnMQs4ZdhdL07PN7rHpIui5k0Xi5okmF6wG6v8Z7dJFDEcoDhTwMRqJNzLK8q6nix4+DpUPsh5A3UWizeiTdWnM2utDrLgM5BadGC6RhVEAfYD9TnJleuuvBpKSCI5RSGBPIfMaO7JKBMZ87nyAiuA+nA76JsqHEAgG844y1n6o8IFzqWpHmZJC1O72gHfUUZxOmw+ZQMVlTZsEMtz5vba/vD6C32Uo/MrFgdZEhVFAXhscyi/Rz8Zl72kiyWj+AUDfiLe89oLdwydOTczCZFVOVs2zazvhhN5qN1IoHNgLw5z+Z/RRKg8Si3tibFEwo5iQhduR7GVzEZD9qMYdzQhnjuks372s6zsGtc5QsBWtbb7HCsjsv4F3QbHyqc2cZ7HMLh75b6HtD5cyXS7QHl5/3x6OtGhsDxnDKQ+3eUwJMcT+FrmytSHQneITrQjFQfLGXTPHaLSSK34lVyk+WuhSypZpOho74DZOmXiofZErH6vDICE1KmhXnssJvBvP9evKvITkDx4/6wMEDsLbIgUY3wwKyp3NXhov6iRkW3PGEURK0ie/UulJ0vQbxzQJy64/vnARSkHPpcuXbfSHIVvJRhdZe5nKZQTOUO10uLkRWwsSOF/50VaxFvNrm7O7ilEXUFudjb4PCootu97E1Agy8xYCMSF78KPm2nPbl5JVBAAAADZULq7y3+BtOUXXCfF5gJ3D0g5Jupl4toezwmuK2+X2Vxr5LhrIFdJZxVYgi54RNW+wX6on0xE1Dgukr8XNBI/2LxvhjkwEN4aIWQUjVapKs+5q1R8VdyoUqtmB0pavL60uyjfIhsb1Jjap0GNZvI6DwMWWz/kucqy/+GC2U9DSF8mn6jAEAEgyyz8lHe30d1QQUUS6n33UlyL7M1ck/ptoQcibCZJQwb8hqUwfjUTrocXh/wIIRBQ+gZjTCdUpwC9c5trGs1fX46LuJa/Dk7AhDb1QOoyFel1ahLMAaRCSGBmg14oQh45qKYPbSK2iTTHQL+VSe7jV+0qPP2toidPsApAADgDDa2cxmnpbcAzxoLv88VDLBeV2hYvRcyPJMdCZTc+00qP34fNkrgAHdnKpIbO+ckzw7VZTXfgltFMaGUd8GgrlWkqJh4oD59W0wlpmLIRNAAacZOTgMmZJba9JtSZWpK//8T1EvAkYtqeuuBfPUcspof622qffd41qhzQ8fAqIPrpqsQ7tpweZOA6QnZ+l4GogrZBiRHt+K30lKFJp18q777Wxsbmf9R4hCxfeUORz/16vvXccKUlL1kwtnjzkRy9ttfDW1uBpI1THeHTy0zrEapg9pQchb5CnidaPj43ZF4XmAMKao9WxHA2NutkmQk8AY0+71ywQiJr5d5mICXaMlWaEGwBkSf/iG529s5ClA3NyKTtdfIZpfpxP2Efe08+iOL6NOnTvTZ434n2TatN0NK57ouYZzgrtFl34GQXmokKCfpCoezDyCOtUfvjWJbis1nfm4PGca2y2+EXEXqn00u9F5mWyHBdTxGnYofbes7gAecRhkNV/sViOifSe4JEe2n10H6TwRviVDRZpOvuCox0Xm6+hpdLUAvP6tUmCQ5lJRgSqm0b4cN5GswEbzs4Ml0Vf+USpuH+t5jQesqY6keD579YDtHHqhkQjWHDKgFDiGjextzwB1RVX8/0PTMya6w7bW4mquWQuLG48Lrtd78Lz5aGHRDQLz8RDZ4fx5VNWaN83o5+Z6Ojyd1QBykeCRpFKcVmDozT4H4E4EmFAJVUTdUZCnl27TpzY9pQq+3GBSh9JiVEVdAEIeVGIaVCbI6jlMssnaMBkEFKzegosCk985H4j3oOtEVYloumHQSkbHnOxEGUOSHfuJUtSoEYtgFyHLGtl0GhwMUaZ2cFCCCkcwbRc/3x7c5hNwNatjnFNxVwMdIlflEpOr0xFE+RF51xMGeqlulpheZKG12aHQPR2/pN/stRRofolQBM6txuUwjOhyEh57MaylKhIdHnM6+SMZkIL2nc4mtIInB8YuTejsy5hvxSo/MZiUrhry+0omf2cRNOLLYGvk7/t2NNQK6BAc248h7WPhECzUFUi3uoKrGzIZhq4i/occ+Vpfuygi22HGT3H71mRioj32FWgnPruZyQ9M0Qjs2o/QwLH2WuMwfLCz/JHN+gkVprSNKQOaRBZcMQy1c9L7nUnlUd4hfR8dmMARsjU/s014apWdUgQS/XNg0eZdA0rz5EP6VBihU5oj9Z3scuLjkYhqNte6P5Ap8DReJQc7ueD7pHWubyUH9bR490qIfmwFASH1M8uxtkKX1cHPj/VkLG3Cqg3yM/Wx5Z1P0FLJt+7zOnIqtNNYsNV/piiCeIA0YuM1q2S/C/7wLfoPYRZZXgjNxcCz3mk1Y/xH8Kjo6NagQvECX9m48Ycan+qTiF1HfbLRXx6ems3srdpCSpdDOWrIYIt2I51mutCouSpOC1ycIOJsQ1uX43lVrlHML+AEhD5P6q1gVhNoRmWvr2gSBOlaD8oiYOgua4eK2xIOWs3QODCozKstjzyIRa4oAnf45SVc0X7vKryLr15OiJUAtS0UYjBhJ8RGx560SechPJ8ODpMXoWSd6bjoHfK8Jj3E9u4UB4wNdnaqv2CiMrFqnnp2DmuH/KSafSa4fQetmVS1WMB9v5iRcF1Q9KzAdPIQiqMN1Y9ogvnHPnVgZv70TZIUBKS10ndfq9WKg77UfY/3i6Db8axRJofKQ9Z3CvqH0TXiCxrp7f0zE82xznf84MLJT3EG1nP6knb8Yzq1sf5F5EFJaKWZmvxOnDOcaYPCQLwmYSkzm9dbtvKjoGNO0iUPZN1aT+impu6WGDTLBygp2Ze1oldjV18iTbr6XwYqA2DVnuNiFIBra7jL6WCDsesltx3c51T8tvYvwTTGop57lWmrxl9NgvtD6AgYeoM0/0NaB+AatcA3ZzZ/raCuOAhFMiK+qUOHKfY7ZGtGUQowIxamKWbVDH0K04K2n3FxfCMycIaUmaGE/HIpe0kgOOVwmvnyKBhwAVlhXPx8KkkNVTPU0ADju1TnAsyl87qhvbme49oto8mXfeS/kAN59Oexl/1Bx17KPuMltyzVKv2ly8JVjtZXdqOeP96PbGUqM8ZFMkO2EWnigYnZWno5ZfhL6VBhHoiTji0cCyKf8q8G/aAAAAUu0DJbFihxOIMgMX6MeqT190zmVXjBiM9DyUKN5YCgTvjnWRlYVpvsHllzjQH9WFdJwdVZaggDOdvgIHFJY60FhJMx92mATevy1PVFOxNZa2ZGJW5NvofD2JZlOTU/HzeX3dmQWoDoJraF77ST28y0FuzD6a5jw/veUNcYa6+bQ7cMnhLGmhmFcYgGKaTxTzMjJkSQZtqsturUmUGvPrXiYMLLZ+jqCpkv1JI+yqFbFxrdJVnB1Db0A6iDr696iDnx18DTT+Bj/hXUdS7b+6lRlznvhbwg7NXySXQxroywbVvnA1SwfFFqUu8KDSyv+mkohrl84cop2K6CHg3It9JMp5F3Ml5afFzfM83IROcjDc9fQ3lDRqgcc3pF0uAKfGkXMydB1T6jTA9y32TKGpcnvxJy++Gx9NMNf5YCcw5Z0oNuDVCY76dk9jBa6zUSRzissbnNqi63SPkMlwfWSDy51AwK3U2QYFWSLajgSQrClKCHMvoXNB8OzGjlVSSZvQx3UIfRXf42EamDWOpNBr9d1DJF9ghMoYa8KdCzrEdJnvOITwCixqy6Cj0lYtDFPRyn/sOvli/DsRGOCKt7lXPUcCFhlqIY5cNhVj3Amg84emPm1J3UfHC9X44mBh3TEpUeuvAHyOfB2ijkvhFFF2oFbObUDP2pJv8Trs+KWS1da31XlLFzTL438ngPrC0Ksvf86eWOGLrd+X3AbgwbfX0qEH6MuZC9i44rsuxd9BZDHiKCWxaP85wW9wdVGpNxXsusqMRn6VOhiKnHv1yk/jtBXT6Hm7dAqqP+TL0k5kyFxt4pCEK/JODz6m25Z4FxQNE+TvPz2bbFgL4vX3U81K9K06Jgyx1FTFWcHd0sbuTPxspO5NSd5UtZf3Nebs7D8lvajHamy7gGVSyfXFrTsJ91o6MbbmWD73MvDsQs72yBSYf9AIl3cjkTNENFtvgdXyfgaZwjwJm+cwFvQYH2mM2yseEOAo582U97CXCSGVPaEvDRvz9yd4wrKE8kYv/UG0FgpkUHfB6M/JLQx53Q6jWwS2z1hrSEozoqeOd0n+2aNltiZolVJqP6gs7lyMybR+2ZvPcdsJAf+6TGlOgU2hIDre7lOio8mf/pXPtcrmBcCwsLtLhJXhbmqrgWyesRdryzOW67hXaPSBwJsh00QA3grNJ78R/7n5dCkGD6Ji28DuYD9TGM6c026MqqT41OHBAj8DiQzFE2jQo2GxNyYrEJgjqjSo3LEs2CB5aR0WoQ5XiW86+31biFcuK9D4TE61C9SHYFtCLhF5suFhbifRHGUgAE4KpF7Va1EavOS8Zdg9InnxXNyHyM10aCnyW9JadbjKkkQQoTU+4uOYdd6xySUKu2HLsGe2fh0iJZGHtLLo6j9aDDcADiEmnl3dme+kozDwr4CYs9LhjWT74kc0/1racbMJt+wSGGuStcsvd9x81dgf5VzQcQUezC7OpY6m/OcHU5hh4K8lgUN6HwXfDnc9KCgqQkYHUoqpKbcwrRdWna0DYL+pxTalRS/K6F5uAbJsyZOPmLHLKcabNFyuLTZd+JOryZFLULVnN3UvZYb86Z2OqeSF5pwQtaC7YMjUjb9ztAMgMjS3TOQdwnwR1cYdcys5Tf02vppc4Z0OK/3bKzaJsDPJyXQwYwsMB5yquAknqETAHV/YX9xERr8iqJ+nVyI9UZHdOs3tCtyFD4fbkd82XLd/FLe622W6vhLyty1b0oZGhUfEYRYLmXH9t6N7FoazE2UVWGQRjfBgwde7bPMKGs22L+plDavEFlCJyE+WZ1Dfm6LCK2+EyBm8TqRYwPS59B40ZXnb5i6tQGmmjKXnrupDggUGX60bixl0eiRJVg5+Zwan7M6gOsCkC9ynjInN/0oty72mDczAnh0z1l7Q1a9rr8hB6mimMTbd0R0Bu5ejaL2rSwUzvw3y20ES7WmcV9DaePCQiCJUxxANpXnWMyg07FOwyAoVuPhsS8ayvds1yYxRFwW0QoOnXuUfunX9rTx3Ak8lWJLiBNSerkMKUQbIOnUzLvGb7bAYqAUbaFkOYCku2LjEGz3MzXaAwHe2x2L5MDyZNGE+vCqMdFoquVB3XU9weZw6ME23K7DUOnwZISmga/WZk9pRsxvE1lrMriAJA+4t7VQy8egT3I3U2wehjLvo1HsOlxYl3x8zP9mcP8MJjWwLjMPTatOX1/Ni2eXa7y3+1uJ9UHCL1wkKKhigp8XT8hgImxpXVWlsY5QakNhQwBX7HyeyM/EwNSIcfFcfRzJefKATIexNaWYhYpJw35keKdcKgJXWQhqBXivHjBa2eXyH15xdOzysdAZTBJCO2WG/lhJf55BRKdJ3grAxD4ge95E6H8cxBncv67KZbkJBXTSQsniRzNsJeEDLBEUkXhTkjRuR2yCA4Flfr1TovlcYDHipw4zI5Dad/KpdpRUpKwgpGbn4PQJoQu80IyQZIbGNyIgjbIvASxxjKrfQK9FipHJoM+rT0JcbtSodTw9PsZjlwW040mGWUCJ7S4DdTgAQjeXQOO6YFjVQAAF7k2pctDiwXv+Ks2DWeXzNREgR9At/NBZ4aFfms3JVJ03EaD6x4fcykqfczPNfNtU7sk5aneAjKpE+KdHnId6AWKIenhyl482ihHj53SkZ4gyn8ifWgAw5QwDuPwrOZ+9k+2N4KcHPOfifQuKgziTUkel4oogNZp+4WIE95Yrrohl3aRWMs5x0C5j2d787zdQAuJ5WwcN70tQT3wsN/NEyPl0voww0Zn+rlfDb0Cq0sNxBLUOK6avsD4EDG7c/ViHHO/rA4ccFAjAmFsILxw4o2OqTiovM37zH1ssepMwy8O8B0zQotkEnl+/pY+E+rnJi/uKsds3oYgXLXlT5sOQ2YzoJBMAQb2omzAakW+1AVVH1akF7eoyTgX92j680X+a0YYcO7oaZxl8msTVXdDX5WgIs+LmhCA3YkpnD8A4VIOMkz9hvYmWVSe2Ayqn0eC0egvUxVN7F8dDW3qK0SKGqtDo9NZ/XjirswRnbCRDzcdl0xOdK/4WccD2C2cy03EuPUqVYbXCcM+hLPpc8J/gP4wm2YeywDjKuBy/xJ9pR4j6lHm/Q7+6FKpvYCmt0+tLnPP7wfLeLlAaeCdC3DRPywd8e/5W2f3wbOZMMY9bhtFiApSb7RI6nolAbFtRfjUklrTVq2sEo1tHj9l2l+aUNC24cmd+ek6qyxoUw5TZ5GwS4PRT3NecxzUfcG2hejljsISlCpqmCRgbCIX3c/idZXk91hZ4T2K9PiiXwFlvNKFyf11azNkaHkQ12qgqYG0pHB2cZioIU4bH7Sryt+GVAQalEkqzxk4Fj/NPfwlI+22ygRllosCxfSGa023mS5V9LwESyiJnI7/V7u2yUEb59FCdPwhwIlPeW/I+eMOPIw+2GamQFZo0Q1f1K81Zz6rEfdS3UgRF8cguPrTH5XbrDemqpebWu9+PA3fI74x8uDHP/wiqEgMGYUr+V/r7eY6qUd5gzA4Cwa/wYX99uV0exk/5DTTpeti4n8t4OMzXDnRdLTIIUswW7AN4T/tABCJm4WdbnMVMmU88sSU4WHR8v0a/AqJTs7bAt5zYSWL9xOVIJ6Hx/3wH1WWrmLeNMtmMgZDpnL2tBoUtFJRaWWGOv69iU9WyXHGIVWMTsfrLjEoOgdNcNaWZuyW5o0dU2kovutEswkAy9f7VOXAS04RVJxGNa/EAhHAKvADY5TLV15IKsM9ZO6l0D+uIXMyxtfztoTG6hK/zlK8FJcQyTl+lgnIAzeUwuXu5JF7ONlIatjO40pas3hgv/wCBwgYkbguGXdypgT06+V2zV/2IwVpKA9Se652otlYkV8neHm93eAcO67s7+4if985+LnvuC7UW+uNk4v3TBuEha4ggy3yzXjLVPVww7MlM7MqpMm74XXzDasyWmvfdavaWZb2IB2QdCU9NAChPaYXi2byxL5jXEnfLZC7c038x5NlPSdNEYT64y9RhtPuKmVEhKQApkgLJxfK7B8276RzKVHJe4aHoQAxYh3gnemL8KYsjyC9OevZCxR5/JvN8HpMWwAQal4nou9hZPjLbAtjPWJvl7vGVwOObJjmijK1cFO8S1mEpst6JWhHjtwyWpgDFsKPZBEwgbvpimbGsI2/fDxf8h75ABtNn3byFMM4RsNdLJZTym+sgy9ms2aYxt/4UEEAinHLA/6X5/k4elpenPheUHPx1aziAE63FzbcyyxGWtzWcDOYKX0AKhiZN/nHsUVQDUP1dQWP7oQH4OgB76Ve8cbY5jAG7YnreBzzcF/w8EVg25LoYCY8azOOaJyctIjtgLqxmpX9AGqhDFsQ68q4Q2JflJCIuGAC+3+FkPZnl8Ux3ftsqUFFzqVdJ9NHGAPYCcMgKRiYLuL7wMNo4O676T6/jI1HAogyaWRn9v5Y/VcWzlEqekM1D7d2Hu3MaP6TQxD6qy4tH6iyHIHAuSl6f/P2skdxqGQfCh8N3lLZPq/c6CGfqm3fOMzVHVRMg7ns4L7KleBUt6C4DIDZcfanYitiDTTXtX67bwkqTu5v2qiv3ud/U7VAhEqxfZF79eNUELFNUyeSUhzhTf3TFdAqX3rMmWdKqbyjlD5xuidGxbVPwS1dxsMsxHvM93bueFOMoCKGfou+4+PBnCrZ1fimaYO3q5KVPxkhqTjMMS+Fgq1kpPdkDJw/bzn9ZPreI74EWSZ5m535xPMlMJN+3V7hjChTuGs6WC8oxlum8KWkSca6/TciNeBEx1G65A64leZUw988dVKC0bueRT8o6TdfXNjlNMSrxhGrQVArMa7vOsiwjkWk5TF6eEMDVD3CDSS3vvZ1BfbsZ05tBEKQBU4eRKIT+lVAN3BUuWFqKi8JRla8wOyu1g3vp2uPpzNNtOqLq55V2HhX8Z6/m/Pp3jnDCbHgkrkPqGs6SUX8VpdhK652NYzZ2G53/PxUsNf36IANOAeVD3OUxLEz528GcOLfSzDuewmPnGUK34wq+o2mnL14EVOFS/8kovbIODEJwchiQ+5Jc3mv7yPLsJRPC+cP8LrU0eJZeCuc1EwSOEAuLxDwbgawbhCdn72bXm6St5JjNPJD3Ux2iyqYRUS3PlxUZROv5UkyhjZ6MbQQWoHKxNmt6TCNqhJ6cXDNe3IbqoARwty1dBrrTgNKYatZrpoHz/KXoS14fZy536WJ3sTnnWH67cMsIT7KdywfD0MRLr5D2iaId+u4X54W8D6DLdFx8KJslp4w5bx3/IY1ej+sETAEqff/7TfxZh09FFslI9a4xEGbAg+UbEkRls/TmffacRGUqotz/HneTvgye8otnMa/lPMxAjlQaYJ8t1qpX5jNShntzsAZmnJngcfJyHSO5oVP/CoSYu9+CxQCsAgUlHJkDehDYqTOFzrS9JWJmV4/4OWP8FWVWiagP1pr1sgzdhxwBug7LaHYXV921OcMlnp+KXitZBvkpg8Y5PfGlVb01eXGFMX44bjvp31Y8AACATfLUOCdrzf7IH8K87SLF8jvbTGILVr2zL0wNoketUg0OYrJBV9NAtQTkOhWR/4IU2wgkROqwXjGYqAbrtbSQCA3t6Yj4v4XtqNgDdBRpaAZ/6AfZ98wYiN002I3KMWT72kipoDMWZz6u7kuID0+N/pe6XEm7/6d+pjvf38ylclhJu4RrgvkBS/B3rqXYe68tBdujXXr0lnlnRCrd+ug7fV00YViwvl3UHF+wx3H6rkGpRkCWqxdbSpLgf7Q4mEO3NsfmjsgasFIu3aJuJq9/MXW3ETIV6Bo5kfCsYwZb2TF2oPbBXB9RX8Q4c+W9yoyMNefs3dCwTkROzV80ZJ4LNONdg7rhIycz9Yp39o0rdFyxnM9RU37x1f6HD5D8fJhJ9Itv5M06c4a2WxSw7ZcEDD3MJduxC9OdlA2/Vrf7j34UfOLwftMb8QgOHQsc0M7VuCvRIycpwqR4kTnkmLAu7C2TOub8n3v0fCCZkDZkZ3XnKLlkCVuFE7YXnKYsLhaz1FMYzsmnwIff2MSjZS5zoLoa3mOr+4dkQxPUgE6uOT/BmVlPuIFbUlY+XQojQ8Wt4wqM5dQ07FGtKAXAKNyC1fFOdpTn+z4Cf4Usmg0jIQcKJhi0x4jfLWTwv6zCVg+kU8eyTw40HjhoHdN1fjWFshVBtgjSWWisAyP6UqlVk4zQwQr0bRxplxVeY2U6aYRl17WRb+Ft4NaJsdzn+FEGWWzE92EXeR2K7IK9mnRz5Uoda0p7AalS3qUdVRp2+HiVxd4RxGaGspeMRE4LU/zgTJFamh90+ofLAST7MibYVcRS0pNC9GxdOmg5p1yFSmwgPo7Sy88jKZMllgTL9ClQyKDFktI0J74P9+jkCo89BXFaYUCALsJG8CdYcLgkVdOH8j0qOQdfC/mp1k/OzMxgSD28T3EQSzkSC/lBlQMc+IUh6XTcNej3oHKhedGwdU5B052o9ZT4joDlULCBlZsqi6n4/LLE601f30YhmZ9RnMY2HdhGmclR0Rpqxy65E58ktYSkoRcRl+X6zWEIr1F7QTv39wasTr6lPs49N+fqfkmRWiqxu5CF06VrOPFESEKqZO4evEGEdjGrrbSmHlMYLeal7NGoimx5h9+5P/c5wODKy4ly1usgrTPamnZt5NgpTCMv+/nr2dsV3osOjpQf3SwRrSxK11BZJmR1skclAWYdTlqpvvd/LE941bMZQDoCm6gxbyAAiQZCzx+SMgLXdBRycwQAJszTB+LTR57RHkRoyo39IyNASueiiaf3bDp0pLyMyFAdQEeWh9lxavsNUSoUNTVFv8M9Vr/ZZQabZbkvnJcl0oa/jORupDWT4JyZtMUW8rYuIJ47IlZXEKCgyQXYHTN9l39nyVKFyuhhrNNnoH+sYd5LppRDxh/DOnn7ckLGDi7V9YAnqeaAht6qtKSLpaBFt4gnWwxIph7ih3pQlgxx8TQIxrtzcOOEJnb3lpzmHw/DQVye2ADr9AaeNrQslyQNEf1x64saYMm079MefvpURutYYj4UjdMKoMrpiSQBfvM7qg5qcIgV6oFXJMWql57LWPqQXOHEJQoV/CUTHsxkbZkrsuB4VoV5lE81CDzAq4/Xa/MRnDVpGSqlP7tZzVjQVOk3oNrtXO9uC6p8P+/jc2UYNaVvDPW9OHwtOh9k8QB3AQjF9rwBj3yL2KQXiiOCZPlKI4ISs9rd31inIncWYzBUNB5liZyVENCuZfDYxTOAP9Qp90zsPAoH1EQqLwOWoIjDq/qLEmxP7O8ZfhhJJbTAp0hpyQC/dmF40rpbX5SZCKDC5WS0B+owYd1YIrivfZWD6+zhXUDPUb0rL2Oc5U+9BcPMtfalc3KEHrzhvb2h2h0jECNIzqDd+Dsz/yCpuF7BUvcvKNEyDvM+f/3HskXACBPhXQxImoxq9llcXwNbzjdjtHM3Hduo+hI0LvLUQlhsuUEdOisxYL8E/8IcCGwXC957A5NU9EfkpyPF+pg5Dn99ZnrXdbQa78wJCBhpJTDbpMutvroFCYBEbise+gwzm6fG2IfNlc8UqU8c5dPRvMYYT8ANpRn2eszgjPgZZmbDF4ytpNPyhcZSdtkb0jvKpJXIVsNqwIwu6hPfMig/F5OQc3WXwkf6KKiiLZ70cEWJZspF+O62v7WwL+TeumF0pRGDfpWc5dtQ7/ob2iXvzm5ONcOwJezT9axOrx40nli0U9VYCujR/SxYt1YKt/gFyppnL0Pj+RviAkW8+55J66qmvdny/Fi1+eluBMDT+3O4AxeR+XdjrDutIS2kGiFu4QXX+mk53oJ9gfDXV448vvc/m1q/+eTVRk4uYhmD3wDqXCW0pPlgiQfMjTUXayUQbaXCTwep8T8Fp199d47aaujo6vlKhchc+pCpP3EClaz0+ZacRu7W1BCIcWhKqUug+OAewGe798I7ak/3pmSbmnWtzcwyK2Z+AyBZwCxgcPab/yauXERfRHO+Cgs+5qdLVViFf/VTySD9qIh7YxYDMhgAFpYLPCrIp89Ogdu6PsAgnOtY7N+GJyTqz+aKF1sokp7HF5r4zFBqvdPSZfeBT4I8OdcF8z1lGiFnGiQ+vbzcgqmGIB72oN2/OrT/7BkryRGli9RVAf1yeJQc68IY/lVj4qlM2ENhRcazaAaUU/Mc8LIKK1yjeYvQ8AgMZxCeI5JinAxBRRwM+kwd7g8N7QlaqKSOaHw1rJq6wkQivmMf/FhYCYJCa8jKnHj2Q7otsdS7emvK5wN8Ceky8F2pq3DSwHGoj4zZe+dqMmxjwHLeB5RKfczjJJ5QtN778Pr9HQPKXFKkpB82vj2S36nAMwIK/onjFjxU1fWgGA78IQ5UEowua1QBcDj35SmFMWbMFHMgGm/uynXBS6mPLnpVa3lxLeWNDzEptNTVizeyc0QtpI5frqsdoFwAABqtP/0Stj3XR9DZD+T9EMJZ5Rs2RrsKGwm1scKyzq/5yPEVr9Ts1nirdnOqnWuXZ7SAFuyaNZfvO0kR9j5eLEr6zPLOJx9nx9rF06CkMLv/qP+GbBK1pezF4JisxyF9oVXk/6gMqt3EJ2w34chhoFt/srQSQUtGUTm3MFuTpBZa+x7T71LlhcZnd1xXLOaXFnGd1492KWGK2cpw1L/d4yKiKtR3jG0PiX2ImQRP1I+WVgY28NjIfAzxliQQFmGZXNwpAkETiarNkfQ+ZrNaRfMCD7rEBhIgUivW7QGhl/85Thc/gFP4HCks+DKd+U538EoZoOi82axqfSaGLnupg3/DCt++ATFYG4ygwyWxTYZ2rCUGr+6OwP3FkICWJCV6pqEVADVndABQDgVUv4+iZm7ALEy3v87wGNxs8pC6VIIy6I7VbMJKxSPTR0OZxN3QyS787SDjZjHSOoQOImoGpLC/fBrQ1sAxaO8B6KPrz0WkjizGVtnSIu9/F3CJ0XZlcAq5FNyNV7PeBgeSSii+fVcMhIoaRPUWhikhwmLW54tKoYoyUmZJug6YIqd+aG19qUySYwV+DFO2Da1lfkLktC1hTtoPpqwW6VrINDFGCmmBK30OLMRayPNrfk6Gt+JlKk0eljJM4YMSj2caEOzsKkv3lkc6wEy4wR2HwbrZ76yPxSPRkq+PettEB8VR50d1h3Jf8AuHKnCPCmr34cBT7NuKNkwXuMkmtiFVI/VntNbalC6+VkMDAoXGJwk/2peWKjY/LwhMG3k0fhQIxUBRWjz3WfusArEMgVcXNmjg31o5EsKCd7c6wrLfBj4tCJgGXCfGG41zV54scqGfDVFlKS7Fwup0EmCcdgAhWeXzKYJIzccts92gIc0ascxGIBb07sLD4OK7gRdlfshOHZ63yjK/0qZCfuwHjnzbWnfVmUHdaIlsnch41jKTpJW0hexHe1PFsNsWfqWYhN4hwx0vi12+D7fFAaw40nnamhAs0QmjPG9wTSSQPGto/5EYVNjV7W6IDiYga3+5Jav+KtfUZ/BQgwfiCaQMAoKqxNoYcSx6BiA+zW5H3LwwrFITygRG5Y0wLqPjknLSnx17B2zF5sbUue4dFoGZN3EqMdzGfmASkhSgrcgubx+GnbPjRLgkBHulvo4sTd+ueqVTp+SXDLfSRRtfcOl8vXjh+45x377Z9YsSBACh0NhOxyQ1SH89PTSSyATaMyBUWlBq+Jlps2vzq07xEWx8eL9J6afe325iqJG8+RAW61m5RjbD2d4t45A0lCqPzANEiX3wrylX6yshLlmWt8a1HLvInsDNTrkKplK7njIJO9OzJVgp7BpfN1epKx9GmLKUtspwrH8A4uFgMFhc3ByARZn+aMO4XRxtAM2aU0G4ZiKvIIb5po8f4KKsC7MlxInkDpyLK38fV01+UpZ97eMkQ4kv+Ts2Dc5OUUdiwnuSrYYxVPgpGz0pNu+j6eWW4SmIZCbbqJWx2Cndn7TzQj87DvOK85B9Nh8M8J2ntp2ivtIdcR3tQnNe4yVza4kyQ6j6OgAJYYBL5vlCQNhpUn3Nc6e72fJIhzgvzLQCsK3SFnhmrwbxa/liYfLeb19Nt86wmm9DsP/NlCzdyhSsYPhKp9kcFWp5VT2bJlylFkw0EX7yhsY6C4LwVo7c6rVNyCqns/IiikLpMbJ4822E2/9ebcNuc6mqULg1QRa+ckooFFwAkQbiE7svreWeJR8hpx8R2+yQvxopUSdneziorjIyIxETH4OizIpjyK9UdxjdzYPSzXZ1wtZLeGmQ1iJU5izzrNgF2LBadA22Sy2sxqiVhlcVo/2y0rFu8igvrmcaIW4Ar9N+Y+iyEBCFl0DnCBSwOcPnudpLPFKDXC+OdJn2nJ/GKBsNoTOOHtj+hMZ+f/ncX3Lv+TDkOOjnacL4oqAYSa8gVVTRI4fOJ77+DJq8cQBgUOtmEbeEY7uptJAlmz27vQHoCPmjaVT9aX3/vyDu2y05CAph8xvsw+dC7GUadU1MFA7iukGiU2yIECq+GiL96KdscuZncdk/CWSoRooA6MxSVGgSV2aAXDliEMLd7z+uV3tSEbYEJzscedCjwfKfPtBOQUJn/mUCCqeGMP/LHqj+FCNV4+jJic/K4XEWF4dbGCpowtA+OJcb7rQggO+caZAKdDCAeQEuSmfZRdfil/ytJDXTU879ORHDub54WxA8eR2NmCUM/5qlB7yELbXDStNe8WUwpoxlIfsfNQ69X+o0g5tSWoQfvcqt4jZcwAZNmlYJXxB1D5SDIeNXsHF0NUkHf6KrFizuxwoD613dnFhMqbd9yUUYkA5doAvQ9vfEcr+Zjlp7OhjP6Q7n2TZs5EVsLBVJu5hKpBi/IApUAAAAAAAAAD6km2G0L6LKVKSHXZFfZJF8YgsH1mOd7gZJrncY/rrLSLGppDmjdnqxsJ6zkWH278OGs2rJmfhsC/TgPoXe8onBaK5D5pS+mEKBlHJI3XJ/8QBWYK+09GDB0mTi/fBIvKMHsB2hqGa/q/vNmQz9E2CJpawmzxU3jFZCwDYmmdcPgyrYdd0Nj42eHuXE61SIxAHx9bh9/O09X3lhRZdBfCcCJMp4TcJuKnD6fGdY61urzMuzTOq88amLEVaVmPBVY4ZDOJHJScdVq4NSldgAQa2bvtNxze781YJE7eJBTb12TIWYNvN5kezqZZBTUB2DizIIiMekJ9Mab0MmN6musuDWFBj0L5wUuy/AMHZr2lam88T3f/ulDzVDkHYk1gEEDpaade+sYlYO4UQZ0wAAAAACTV3iXJg8YmuXq/jG75X1kcMXqLyLu3Qo8oZ8hzjR2s4AfKOfOrQAVJk5GvJyv0J0RVm1tk6r4UI8G0nUKR19t2ZpGuXHFU7sqyKAQGz4yGKzx+wcyHoaSMIv87ZEbe2DBIg84HwWHQ9XB5BJ8If17UprlibVk7vt3jQWU+Ri/91b+TjST6iHHbZ5itNa6ToclkcdQkClYst0NUS+4l77W5jHxIvyVDV6rE/DaPXCLtgPkAa5mJrAtSmMYm0VqtN/Fdfca9gdvNi40TChUrsoEiYxqTXgfprfiQCOtKRjZsAYeJUrcZtmekcHlpiAqQEDyJQbIGd6y5aJOLLDF/3rztC0nnkS60wgfv8bXPx/Eo8XBrBdIabZAAAAAAAClqe+eJQgjYC/enZlsub2Epkjm2lNDeNxuIH+SmmotfxTft2hgyDL7ZBD7MT5Rqmif5UQBwHtXo6MWXNAkSXPAqzSgH5BbAhN/U37CbHVNWfNCZXfvZVGZm9bpPwUL3hcR2aR3F5VsCvNRXEAg0j39lLwut1PpfzHkTzuVjPaIOAdmjDjBcLMfvDqmZA2aLREW0NzA8gGRfDqeDGJiPrpuK7hBHcbrAJ/6u9DIM8lytRY+cojjYzm34nZSdmypuoORcMFjZDMJmTpTVy57l0FJto9JSw62KgLzAOHgLDb0JL7bQRGk0cY9/1V/Odo4gxl1BJxqOw5cXOiQKeXgpyujU0aAwVLbIr1SOGL+CChdEcV6SlOwxwkX/70qBeIgWOhGLlHRGrCXqCa3bilueAjlE7bAZW4oDYeYD6uwYXCrKQ4V4tbbJIW6vM7ikbIPOxSWcF0uD90dMOixc0NJo7yXKhrt7PTcPpZfa9rEaJgxprqhabz6eprl9ci0yAn5LeLN47V9N+uxYdZNi6lSB1dzAO8udG30lKWcKh66hNyH7S4NxNydb8GXvXM2kfCQX0+xcgTquitJ1Pbio5nSNm9MrqnR8lo8l0qazrtGOXV1YlfPefRiYydcptnv9KW8rVP0AG17PB18fek6xDeCM4y12894ehpS3wom0ImJoj7fIIPvBfoFQ5ahuNozh8epHkZ0zHSJI4xxmstudlg9cZw2EPK/X8As7SVwiSUcHo0T64a+z887C1IDUQ3rM0xGnpd4agPKD94AeF/8EeEbShy/LtQA5oaqbWwym/0O1ljZS09DxPNPU6H/aIkij2Dz31UQT8H59qFVbdSGc8eH+QkyEO77/KbhiqT5oxyDppFwIwt+P64VZb1x5n8AABBVSdpVUk3tHizZLYR2+oJU/K8NS/UrQHm6iJEu0dmqX432TYSmn9LXwtse+9lYO8RqNdzR281MHpXZqbOZksSEXuRSS1iYe/A6pUi4WuauP5gUFMiekyT22IXbEIHFXx8+CWtaJDddhr/iiut06Q2nfHI5CxKNN+SRoKeuy3PnVW0kfY6A8wwBsr5KO/0OFCIyoE5rRXsTHrWK8J2sytigA6kN0grocgdz1saxXSLWb6bkAcQYO0O4a1jhCpFIr4snBDadfKOxDeSkn433sV1ipbw89LFRgK7TxMcwqAqqSundc6gJayYt35LFLG4wi8HCKgvwyB0cAefdldOnQcwCVX0+CZMP9KoWIxRoOFwjs8DUKNqqaPnZFBJzjE60FkJ6DFlshMUNx5dP2/St6sRjangwiuvAIqgkFM2/VxEXYMCPIgKv6orT3KtFUB7QMFyWesR+1e1imDO0PNzbX+1zIcKPwKoYo9USoLH1MWEYE58Ima7hF8jaJAepPcOCMi4M/xiv6tz2S7h4a0sFU3Jmox0MTrNYQp5tnpkx5Dp4QEJuehPEnqmoD4wsTlpKoq4ZW8jSXC6BTAC+3lWQucot1Xu11TIbjkD0/4jAr+PeHvzO662NwAvr8AoI/GHe96PQEa5HIbxb1LRcNAujNz4aPABnvMXzYpS0Eqi951A6VQLLwFxRUXX51x2239j8Kl215Q0V5cP+JvexN+DBhuc1Zvz7GGYCCtijI6dwowljyXibDuuPwzuwJdbhBpIeVBeHdYD3ww+m/GfoBYFguLyM55YucMctoGW9gb9rF7iKBnv1bfxAk82WG5h7bff1QmQhG0gwsU0LSsVGNgoQmWdTyT2/7CFvocXko9o9RHluLOhIWwj7ncEj3Ravr1DlSuX3xT+KiUQ/nT3Bt/3Hm91me7PNzPcedFL90MW5gy/yb9JtRgWq0wqy/ciZJoFgmF6Do2GfYWunUMZ12xvJ2Ay5j8gOsob0h99F+zi87anOPLH8cKPV61tCtj/CfjJAG1Hzqd8dG3NVicUsJK8oE5OXjUPq9P5tJSTfYeqs/TvvegQsbnD8MvQfwixroXQbNOC83VLaemp9n4wwTC1qE6Mt0+9rg5xPpuXVRNRM/aIJRnDwCKtz/Mm5gWSMe7z8AmG4AVlqxemdSt8QXacnY9+I1bYuGWAtZPhm3WLBq5I/UZ5s/hqYS1Zs0lsm7nAN4qTi8VPzOudzjAokVeZj3mZhJ2NVLjD7JS59qmN1CZ0zkLRCM5hhDvYAn9MAHqXP9fkz1mA8/kXjMtsvPqZgNoEZAbInnqDzbeEIbp5bGlzGviqEpm6gzJaPtqtVKstRVQmHfB3ASEvYJa8BYCQXC0eUSkHvUlRh9uMp9LHORRiuYflo8lY6zeFemTuo9bdSmCflizZBqRw3R51BcknLa3PNJ988kkHFm3tTSSi9DPV2yNsj9Kqs053bqXGnTJUVvzqMhKgzhSbNl4cfXL1iLd3majcpyB40OCX4eQAn2cldXNL8MHH/SPP0RNMU8I8gK4d+Ldrqo1o5H3MGY205B9Djn8oAzIm5aBJme28GtYQFR/dFsFV3943gLsgICyUoeXL32t+uELaw1081qIE3YFScQM2Lxhx9cOIF9DwqiX+m7KrBS88kAiXkkmYSYSumAAntk6+wHlGXgclxuA+D3ayepI5C191Bv8EAcuEq+YO+FdgcRmHUFpvM978xMQfAQKALtlNp12Lw02Hua0pTAoyZBOEFifY8x1YNbWrneGqOyR3biCKbtyJzj67VcjjeoS8yLjBQKqivIKrZqpnJQFb8GctLkCADT2Uuzx97f/ThlD+Uzs5G2mqeTSvwjvjWDxfk2Jg2f4Ljem9LTzTWwx3c/a7zcCc4HzgUm/9XGXKprGV61EpNFacpHulAfhY0f0hSHh3n6LFgNiE58JMNBvPgEtPYmh2Trac6Wlnf6TL7x42eV2OgF81Rik06XG0Hgw9/BAkOJzu1KxFEz0EiyRTV8yT+Fhfncksx5VpD74sk6iTme8VVAWWTCz5cherQXbyBi5xMEqL68Qq9SMhNPbr1yvTOHyfGzETmCZzAnCjotf+aPezg/fzZPG3+UFS3ipuduKS9sYyAIbZ07ZzIypJMle/1C2RlYhcoBmGHz9BNy79PL9NPiRu3gicIsNGsEtwXFvkePQIjWoBc74aqNMxpKh7iChgWBS9dWAOohIsEJBvC4IhGL7d3p0sf17KnEO0Y8H+rATXRbNdIKYsGq514+43LAaXm6ZXyLL2XxC/C5l+hftJ7ASXgyV22axxjDWv45lIhWTvNh658ia2D2jbiQX5uV+QiKFtfcc95881BakJP+0buuUgQLaxeM+uBv9AoWl0NxaSh41Bl9gDvCjSa0lnvb7oqhOr+n+B5DymftRpc6irsVhunxCTsZ34fwFcyfwDNBl/jX2bYGngSNES3lF8qGKhOUZQmcLMB5UPUSPYvrk8RHRRsnAQDBDqJvJqAtmWAA6bi26gopd5T1GYterjZLlw5cX8OxwFnHFgI0JGCQ+Ayfb3vAE+4tUr66H5Te+qenogkWTlG0wqKdcQCbujdifa0mYcNm/fhl5a6KtXYXfw47HnlcxyqpXsCBpWKIzHc7uRThQJZFOY4nwhrQ/DtQRT0Z9hOPpaZYDZA7ANiO69DrwQNooryevuyG3IOqvKRLS728a/r15SjQx9L7M5Z9wu+vo1TkEvW8H1vGQQ+aWqdD5d4B63i6pVuK2IijUBK4MW5J3NsIngTJgJ5dQ/caPW+0HT9keMzObX44OnL6sVQ2nREMnGX9f9IjJK5yYXCDBzH7MG2ecgvru8TA0i+HNdljMi87oa1n2LQYDbuv+SLIFqB2Fen3JsFH+RDDKdZLeF/HDV3kEulsdfmsYY+dVqJ9W376dIAWAg3uekhi3fJi/qe/ZI5hhw06BByYFoTM/Q1NSBu4DIoeb/YvQVDdERIQDa1lfPxx8h4KJkTWcQxZXQs1K+JIZE0/Ypugde+tbkkn6CITT8hVb0TzavSLlCb32815UICZqa6ySOIoRhmV2UM4y5ZnotoheTL75QuATcQOF/+EAQPq3DCczztAqOuayAAErPrT23DKG+uD5AtzCmgPuoAuy0noFWse1AUFhM6u51aRM95gkcXpRMmkpJW+UmFveonHBfm+uIOQ5aHBKjm1AluoROJBtBWkF6Y4qB8hj9lJxfQ7uh6sEdGeBf16iQ7uR2KFdMo6DTwhZBqVg9zgewX4VCWAhioBP22Aci7eBCjj2rJnzocOreSBDAQgwcD8qseopVE8U/rgG7/7dutmJYJB9XKPMOFtPrqRzuoMKYOUIRGCtdgSPLH5TfcrqiLiyz+CWC4YMt3qP2Cvrb+6WwMpaBNQfTPlrfpK1u/Kg7Pr/pPuKxedX/OSMrYI3pMITnJ04xRY/lDifzidmWotBO09SF2DBt6YMivx9b2BRgtC8bo/tVN5rDRzEjfhkK9V6WUcLRSqkWYsB0dBWMfZs3JrAB9wvGnJ4fMIQIicOb+U3e6eFztDoikkrx6dwlpyJfpCTsX/MehGujTRYXFCIY1hR9A5Z1xgapLJvdqcmvRDW49KlDNYQXwrL0G9TCx2Nvp+u7vuZV6y188a3uoqebLttSpBQDzuxza+dFYWjcjo0cmEUukD4mXSCWSDfqGyFLO5Pa3xeI5KDDcOIF0aCv1bo6TmPpclmmKAQKs2KPgk+pZhuvq7ofLCU4FUKmB23FZ6W5ivZyp7SNAfNkeTxyeHsjadSG6e45GZAJod4OEOKT6NZ/WIfXRsQV3aZVBrTniDIOQAItG+BL10bcCxdUHn1uKDQfPoS9X7c3ALe+MmUa/ArCkxiVilmzbrcvhLzSWxWdUIvn+Euq2CcPMf0zv9T+xKE+4R81a28s9wr5LdbjDCq+BrvdB2C6XJK7oPD8Pecbg619YkPu1HPXWpLjsrc0aWh78V2tr8gqetsTN9aEG2Ttatcjby9EbpJ2+FwG64R1FmJZBvw7VO2o3MWSrQg2BcP1yF26W76cwo2WDFSgXiDbnsXbauQqI+FH3hHzB49LF1ZLM5ucVrfkeph5Oy/e1E1v6h3BQDZNAzwteB1ynnlfD0BbCgRf+h66tpKXnMyizLDKHOkq8UFSeQM7Oonst+AK7sMcVM2aMEBenZgAnXjhhl/UsMVaAPDjaFG3giZMNwVeZ/nDJa33yNWp6D/BYXQ9y0kD2wiswm9tH/O67ytv8qbL47HRBFoQWza4abkqL5D9/VDAMJdMvpd87FhcY5o4AWOSV7xR28NIKSanC0N/z0TIz9+D4AtbgrLIkxnO6XnXBqEi/KYifK+4tH+xCqxCNoREaMb4V1Pl8EDtgNZHT2V2XY6C4xb2+nU/mgbQdrM9hC4NdZ4E58lumhEY9OqukxQ+TpTxhnVUnPwPBRSYpFJp1+0kVahbXqE4H05fDyYyQF/1vOw4SKQIMI4KSeEEv+jdisZdNg25n0XpZnWcGCEf4z0OCoprAdykAOYRa6BOoNMyBeyMer+/jiAMgA7ENefuRIC8JJmG5ZvimhWbgIL6nu+RQBDGswhy8XpXddLT2WkYhYtOZR6Vuo7wWAS/81wdc5y1wuhKtXup3wrKDeT3AxgxVStHFSCNGCqIqQ919Lc7Jzf08lKTt0CwRF14tOja8JmmgJ5w4fPQA6ghEJu0PV018yhYHvVRILaIMJd8Ci/pasdYTyd9n0J9T4w9Ioogc4H0hYyMsrvmuvOZuUfdmdB8/ZGD34yZZkdWu2PN+Cu3dV/BcOIGDgp6RqKyY+zqK39njHocU0zKsuDAO9OtfvCfrUi10OfMjLfs2DckvjObVwvC8uXWi672FX+A3IKAIC3Ahl3kesmcHc5o1K/yUyPpq5Av7HjbiS0wgJP+ICgAtgxjN8SV7CIVJqg3CjkQ5uQeLYr8bpNc1A5ZSLJ9fGQG9aJHWdyvJ8vHF8DS1Bh0ucf6mzE+BSWE6XvKWxn6PfpMIMAF88VQnOdpnmTi2ezkdVv7DZJ06kmhD0p66XenA65wN3FCkGt+9cDM55QbvT+4yTLS630YakqAe99RA3DbjYw0ekdZ2Lz2EDX4Azyzwhdni5dsK0cjyG3McNVTt+IZ4uWpDK4QvdtJBWQ2nh4JyoM2AnY4t7QyO9sBZz4chMMDzWAH3mpkc59zObewcH3SbYjq15sn2tHtd/8Ta7kIwU6aAIyV0uV0XTbiqoecYH1Og0uPNqdG3athjA/5/PbOo90HW+3m2RlmZgK7Wj8QWseQiR/m/d/zq/F/j+DKroLKBSFxPjLqpXT+Gch4A0T4zTfUA7GHeiW4ORtDrkTeQoMibLVoAYPyVWgBUWpA1nOC+d4ZI2/eBkJw+oA5uiKEbUbhZ6QpuTbprQR01QzXlAU7gdpEu8ZVDHbz8chm/49fV9Yzqq9AvcHkp+ipo6NXiZWXXuvot+ocs4gmgXOiDi4eSCERADs08lI85K8MweWMyDVQOwan5/zV+6kY3aSHX5ZKALDCt7P0DT5KuxKceOtejnbmRFmYzNe17sxy6fdwkJJmrn1vJJ66Vyd3tEi0QVf8QRadcvBnbI4RPkhaWUZUlQ/QI8FL0PwKwEowhYjcs21VgwDQ5tCjF+dx954AZT4PC+lqs+PaPKV2Oylze1MLTXxsYKyDKg+9pFq7XFHIexveGajXfIGwufYv+7G7JiAG9/AnuR3PHpcm3Np+ggki2fepDLPpeUvpixQ2PQO62qjA5+3KAQJFuC7mDHPwIJY7mgQ/ilpS6WOXKhyv5pkLQCPGjAFw1ZJ1OoPwiaJ0Engh5+rEy9Y44D3EQHrtukhCH5rQp4x7lZvV0jp5BZRuhIFn4evQ0pES928EhJ6fJOIhF1OyFnoBXzOkAyLnvE7NhzDlf3qogR3qCjx7XA/N/iL/0gAhQMNDx+Qn2bLhpgTNkWP1CpqnhAxFrxZfFYDNm2+m9G1VZtUQjeDzTkgQulLV5UEEkzyIg+arwQ5AetYxemH2CmR3RsBeoESbwI7b+/S6dsmGmHXl6FutdngoxbI0wp5YmpMxu77v6VW2AXV2Uk1ieQ+FFhGYHQgL3syXLJxQ20GvV2izPHCqOYYwU8WWcrON9KtIbqskD8ahzSburshihdDi01IiGbn1mKAPWB8N2zXUtQCyz8sT8Mufk83+M48KRNjzlU0gCRmLeg7AvUKIM1IT6x5KqBUx7Y1WYEeJuzNk4fD31hFs+xQqgfQlnK4J5T6KiIOyuLuE0QpN6No7TlnoB2UiXyuOsHwAlNX46yXNFIBxI7MiGw+ZPSuXrLWz+irSyfOd9bgGI5b61vKS4uS40y2EyzDSLA7mbG3F45njOgaskeluhHQBuhbpgqBL1ypCpZ6/KF+10n0I4QDNOlcu20Ahw+HYPQMHBNuhrOs3Bm76yMtMGZo3vU6dOWubhy+RrY+Bjx23jvzimo/RQZdPo7NU3n/TyfXVTyWhZlTPG50SYqMgMN9oKT0xwEX/Ktyita7AMLiPTMcLDaM0AstjTe9USZq6IPcTpp774HNyBOtCCVd6u4C81pN2es97nyUEdXIlW0f4LLJFpFODsLsYhOOmg7gX15kCuRu8x5hAAC5CmNmHQFhBvah97BkMGQmEIGxWFnqUXWrVXxF160XMD8q3OuhTc5LU/Bt5wGkFVEflq46A7fbQXqVgmAZt6ijZTqsFeqL9grZA09s5jw0rFU62+TyfJzSNt+3ajBGCbdYfFW3lbyRezohroqqeem0jXDWtyA/NAuFH7lgtVrVK+aNJ03bF5sMGhBK0BjReiTtoXpZm4bUdQV4jxUpZ28X7oTA50wdXyLHIOyvArayBXa5gJ4xztDuxv6l1hg3RiSQAB77HAQFzBBD5gq4KLvBgSohTS/zK3yIM9DzeOEvehn9sJ7Gu/6mOS7hWSNklH78VrmzPcfAHW1UQXQqHlvYR9KhlUueoH0RXTZLVV0atfE75qJuploOsRoYq1Cwar3Tb4B0rMdx7mBcP4bOejcQ/+34OCfekgVAtmR/8uzd4YGYYECnzGFmUJp3O5Udw4VH45GUF/RUHw0sdCIS6jJCeGnW3TboU/qo4WEDJgjsHDP+ZQN0GcGWsVGFoLTh41bs6cU/ISdjpyh9+zDfK9+VySz0cp7qeFKNlqwebgiTJWWitq7+ZQ/yTE4/IfjyL8YsdQxPtKU8vfHO1cc3EfraT19WJtIbq4ZfRMRrtGK7yPTAw7vQxETs1tK1/mQinGz3TE8HAv3Utx9/p6ajcspYSoO5jIWcHaWnFyme3pXbaBtFYR1k3hDLpzGG1hxOIV1kKzEIJ7M2YRAjQehl+39lhp5bXpijGlY7DaTm1d7ZeVArYMWJDVpAP3C8smVRprOUPRg71W0oqO7L61rPW5g0CpzabPVjq5TQf5T1e+75vMb3mOG07ERrJfPF+W3+SNnbhDFIK4JQBp+hyglFE0zdLqO9BonTxMER7IjI4h2uN5UCDIIfjOx7SKsR8Lir/a2DnrB1vj4QKYsfdPxrqElhnIDRq4g+sUp8v7rH34xtGdT6vg2W4S/7V2i8G2KEHybiHeIW1S+f7XhUcrWV52xBoff+yNmYmHFsBWpLevDvttnf0MTDmCR+QSIA1ygA3S4sNcakBmIjbrGxyM1JNTw9woKyRgO+YodoO/WtxfzA9szynF36VgZWX6B+vh+MidlJNY0/ENAgZsMforNHqBm91JdXn6NJ25dBbnGywzfw0VTZhgpljQ61eybxC7xykdSPVMFLysdqpl96qXymEEp8A7Js7BqT+Ad+lMh6JhNZK5ydaFDFcvuOyB1reyV8mANAYAWPn2i4Ro6OOn6wM8fi7Rvp9G+d2DZdm8K+Wrx2WyzrK01Zy9aZ153r7xH9148ZfTIGosIbwNS78zlFqM+pBSsnDB1M/1WJnfC0DKwMR4xS5hyAbnANJFeDcrVK8/g+zPNMw4hz4lU8B08arniaVUkMEzBKmvQqVmQurCweOGTTxWZOfml7npGUWlFmrlBNUCac1+/WrX/meWs/t88eb6AOPgE+804UiGXEZBZCMJmF4sH8h87N7Ij9A+Ey1RcAiB6fGW3YGgLPdbo310aehP0AnVuoSa3zOc3xHt7jrDlSkctfNh3hUdZ2zB//wFwwKXggdLtn0JLv91rIhINc8GahgvHjpsaBZeUzk+Lg6zaJmVdBJOV9fLp75aHMh9Zq/T/lsuB0abTuZbbyFz1Bv0zauwt2KqWH8Uof/pZmMgmI4DXPxCFsUYBzcTTB/fsJdWBgxxXWhMGOO1d1H0ngn2+PJtTwKyp/yYl6LEgG8mlcDouYoLT3REUsS093XfpEybU10GXXqMo/ccIZHB2INNYKqA9WKUsa9tFSdZFXpOd4uvtvq1UYE1afdNbwgxYHFZMA+uExUn2P6bEvI1DnsCZZz/lTtD3FNI99ct0ao+qgI9vxPflVjYWMSPIxbHg1P9G1YSbLL3jCoWcUYjvPnCx+w1d44yMFzR9LUXlSD6PrpHfZAI7OrBouTUnOY3ZtCVLLnVkWzG+j+cowav+H7DjXwhVgcypy6ZcZ6AshWYLIKL8hi4O8wkkNRYM4wD6YErFGtLNaNEGL6465sFBb31MCSZuQvcc/HsxqtecrH/T3sVmiFe/9s1yF0tWUqbK65pqfeTDXxYBisF1Jm1ueVORaiJYwUGlyeWVXLaa49h7X6mETsr0qiCMA2M3+kc3z9+D3DOAmZ3ifUiOQOGkQ6ya67/+UQaoFuhnb+vwT2LiNQqz3KARBioMmUFoHJLnmQyzFPZA6WgURJ2x44eNLWnHyUPLfpe6NUFL5PC7c1VdwRbsbo4bQN4WzkzbRuBe5x+nRq0qyMrGwMaEycKVBhY3dCjGGSH4FC/L8LtP0miDC37VDW8bAvwN423Hucc0TRgnExw1I6+KveiihcKgyTvNhymZod9tcCveowQYo7glYyWrJ9b3f84ClLL6eP8MHENfGIOr6t5fywJMbn9U7STr+JPjONYtFjzC9uFJkHjk9jsXH1lQeMwrq3QfkQPv+6fAs5lPK/L8b4PZmV9Faxhdxt3uL9hXUlTCwwjIEeFi8/wO4GVAMn/x+EmsMbP9P0HJoOp/L1lSHT+mvrM/S7o2hnEG9/gXBg2akZ8PoVUmFr5rXzU6dsfML8zQay1BSbZ777IBCFyr4KlsPwWujJvVYpXr1kqNmOFVOMkCfm+zCC6joNR8i4Cuio3dk70VQezTaSTzHgxj8mlmqxPZNMcQzgvjZW7QfHJrnl3H19cCq++KWGjooZi8+JiiVANtuKtrISlTIXKRrNSDz5MZ5aZSJBbD8XRBa8qFVwG+A3DS3eiMnYLjvZnLC9YjYkOGQhkotwHZwztMEDQPhrjjc4qIDVj2wb7kU1jwAIBtSVYMGvJOGdwF8UIbxaVN1GicJc0q2gofi/C2VOkUdWmFdJ4I1XhpIp8b0VmnjlX44eFhqqzd5tifvLmBBR3K5QIyZBkZvWx96Le/WMZ1CdgBxmO44UEKXM6HxeLro8o20prKfVUencCMjQY7k/+3CZHF5PMEihaxcsmPYqxdZS0hBaiowkXxzhXO+U6ToVZizSx76rwGBRGLdplJGMFdHXzDMHF8dHpK0zBqCYMZdnAHLrmsgA16gY2z6fwSv6V+BKJvhxEIlm18UK4W5JD4tHzKM3E6ehT48N6ubhjJis0H5DSV9lgbLUDIJQEIJEje9ksw32l9ye6Nhw8mDi6m2+EQsSyYtKduj6j+iEdUSUrZyavKwa6YrXl/nS56Fow8sE9AfPbgeOFrJAMBmGJFt/Mi987fA3gEsRw7LwGZbwfHKGmqglvzZEBLCoVX3CTW01ItigKPX3r1GJtPl2/zPHjP8Rjw/hQUmCgov5jW1DreoxIC7zmKAf4luvojzQ3Aw4fnQc4tpSsyPCqXBkzZ5L0f8/H2ZNT4xB+nAlNfigganUwXEDZjAZ0oapDqwKS/N2BHGlPtOW/zS5e/UEEPeSIz4EDsPEDfs/f0LZt5cPHxgZlda0cZBg1WW2s+PRqPIgcS+EEerIVkS+WXG1b97YZ5Edg/3rbRvNYR57w46ueKCVSEh78tIq4TwPn0SJ76h5KZp2pDCZ39e1c/iw0WZaS2lSiAU6KyuBKEGO3XQgSwbSVCVkOGQazMX+927EOVAB3wMuobO3NPDIPzZ5wqjZ122+HSGgHKOn4kXs8oRGKWw785bj+8n3fwuc64OOkzN0fnt7UMZH5G+yqQYlrfuDs2xCA+EtUPA0p2IPSssO3c0lCZyf3mmnjV4z2u8CeQ/ND5kVb0Pyqjk3fsooQnxWO5hveA8l+DPahN6oAKWNVwFtjo2fF5pBEyC1ciyEpZ97hRdx1oDxSTP7/N1oUM4p7w2aDpG5jrA9RhqvHggcnV3joRtNDX7qtKqSWv4pQ8K3uQxH7RJcvN6MC+cD8p5trDHVGw5hFNAsOdQ6fOVPyh3cvRuxuiRUlgzoXizvygab8Mwshg0K3Xa7qzmwHb8Njv2KKijcQKAP03Iz73lowr1QE3Qd60wgbdGLAFgYnF6K591qlrv+CMst80cRaFPKULJMmGrib3Im3QjRjtAhWiJ5y2PR6Wpx1MAIa9AeXP/joZ1+1Hf8uxik4RlzvcRGXgBhAzBGsexpxUib5WBUXa1sblGQ9VJvWa0IKahhCu6ePZcByBg1qPJYCc0CYrUfFjda2uUuZks6RvH6m7crInXkYLFhlQeY3KjErixqYeQJ8WVOIqzvRhRocqQC8VoQQKElJqbqlY4u5C6nlzSX90S5A5YCV25a5hYBrvNCCgi5zmMaYIPbT2y7H0G0739RYrxKzHgBdGfEXaXj5IZmcSKZOFugIJALj1QY7tRSK0Raxh8osL8vdTlePtc7/FsWIFyTd68mc32wV6H/KjDAMcj31NKfe7Dr80PHXNsmlRisFDIJAonNxq59O+6Lmo+3UW2F7ZNhrjuj99M7Qk4inHlc2B+JDrF23bO4sLAluiCcONSNrYeceIWiQJ+JJQoBVdEP0TRfP8zrBt0fDGSzJmbo45TReJciHQcc1f+mpa7U3mNodHdC8Ob4vxid1Kjr/N1WkB5zYW1iZsTgnYNmpEwVBw/uWs0U9V8jwB3H57ehgxsUq+ZMFNSgoYIOzgyREZZwH3bRp7KkfTZ3zdrcd9s9LIJtimwxwVOGN8gssvJ/c7WhcLUJ5bhpzvAi8c7sO0D81CUlT1k3Xkmv9wCaV/oT0SM2nMd1Ikzs4yvU7Y5Tctrt7kZhbCS4Q3IoJxcftH50KLB52maK7vrTlF45JPNHfcHvz89FyXZQoX8NDUckj898LJLxUjUr89u6Me49/moFdE9M9QNMilwwsSIf+ZugCjfEHhrlR5dh58IRySUTUP/E4wUTlvrlBtcvJU11TyQEaRMzzM1/vidXcmPPUhDXhGfmXvizlABO+kcF+TTjXiHvZ11CUzh4+C06AED8Po3VC2jvckrx0C8mkLvXWTEn7brIB7wBDE3lu5YinhumaSbgn3AYWx568NM8FYS9Pvts1upRdf8VWPrzAwnmeotPRx4CZB1esacUHOMocogb1clkGZ18+ga+thUB2Mn7cslaPerVqcANHrgcrnMtiDXNT7f+yZvjRnoRyCvLD3WnTyugUPpjwbJW0r+PDaTcl0GwZIea2vVv40FEvDpqWAAi2QQmo4Xj4WgM2wFZqzTCpAoYJBFgahWpWbTuPly6ijSyuadPlvol5+1OArI0vuJDoKlEGnULKhzYLH15TeOjR/V1aElQht3tY5FjKu3owHIL3FGvq4bH44gKfvHekj9gX7Kcd5vrV0qA+86XxwdKNtOq7VL9NnSs3b14QgvuafrLJYiGY10gFBmEEXeVWQkj/GKkClbNVZlFDkUHDy3tG8RMq74Z03HkukqzXNSo/97vuH6ifoq8vxJPyLD0zVg2DOVS7xu/HJuLpcdPm2ymATIdNX7ZktwtkyXfe9kXLVope1jhw6jasWsvUbhAIOYmaqV+X+3MsjKpyfZ4MzpY9eYAqPqki8nW53BUbGLO6QEIclKevN2e/CkO0BP6yhMFWz3bg3agQ1L488UKYjIzFjbisSMIXBbmj4intD2nnkBJ5hT6ON3sYLaNOYSQhN9EDHH2q+t22yT0sn1/cGVscFBBGzH+2/mke8OlxmM7qqLTrE5XOs/yzQzCh4wfH4IKFqg0KI0j54NEgY4E7UslwqKEP/UfA3Tf5SsWjWdMBfivwMOk48QjXAPx2YDIRLIA6HjIdjElljorY3KUg71iAAGzIJ1UldcTTtOTYZlo6UNqlYSb+1WJkNofBqWnet0GeXmcknvYw9JgwNliLQImQPKBmfq/9PyxXBBJw13Fzh2OjrZ2poThmegiVG0Wg/yYgUdlqhW8s2xYMApxp6A7eSnp7p5Jr+RDLg9evmZ4JiiEZhd5wOZ2NK179NJdna2kIkCwDuGlxZy2xbLOkKEdcFnuSQYgn7uVJ9Y2TB26XAeeJWAbnRKQZDs89lnR/M7gl4OIoESIy7LEhxD6LSbcTfZA4JdAvZxdQF0+x7Jl1cb6xWgmYMUwKq9Nx/Zu9Y0MjTYPbHFT6zsJ0ZxpegaotuJYkykAa1sINSH5PJaXswNyzhLy6UsVz+LXpaZodmh/jTzCbW1z25sLDF053Y83NKY0Wr0TsE2/JVe/A5nehvSFWHyjtDI4+AA7GhWb4Z238Zp9yLYn+ahU6rG63FwIem7oj51d8FNRxcwDmiVjNAtWqw7IQXZsxuSvRCtDOmVQpqvbPyUGZKDcwue8Yvu72lXdGe/4Ye2gwA/2HMYA90ZRngSiKmGQsgUIYwVL6b5/6EbfQrS3ELE2yKUPBWUXsZF95B2FQ8onbmRfM063X6C+hoP2XUcCJzKkEW3dvRHBhZaVKpXOs15kY5OqcA4qe3TnlEELo4xtiMhVNOlaL5KsUvv0YJJyOKC3UC7FK7AMbfLVOp/BtWm/fQGljIGS3Krsy/aMh5Sb80/UQspB1bQO/+wznorZNEneu4QTh2RWSNR2+Bsin5XZ3UjZURrb4HVVQQwfi6/Wor1f4vcOZOd5hH6Y2ear+T8fH9c7pjatDuxCwy+0gVBzAv3JloP8qK1yuL4HWFzX0YO1rf6ANVTyJ9bc/uGaNbMxzp2s35YFeR9AmLhxlVOb8DBOT7j8Pd4mI35mtcfKiaFQfbpJHTqu4p10GfwTcQ5Jau5DXzcrIiSzXaUjl9BlqFEcXGRAexDii47e5U+orp6JifgoGNwOJsZcDe+wWY+gMN53yhrE3ZzWigwUV+e7CKzOGq+JeohI1JXLaO7c8AjL3WinyepM6kESnrS/UysyR/QS/IUIC7ZrEQLwt48a3uXTJDP1GSoTEssBOkbxWXT4a9EyJW6C7vx3glVHb7a7YaxyM4Azw6fIDLiFu1ideyaCTxJ/f+jMWGi3VCh7UtUES3ACuRsAVT1pfcrqX/EfwBhCB9auEa6+0VcXilVGAId10oKJduQHuFPQkyHMpaZXe9dnWXQevaisUFYegrxuOwSXa3lv7xlwjGiWuVVXOIGOOO25dXSw2AOrLdJ84aQFEJkxfjWa4XvniNkO1jT0R2DcKdVdBxrk/GT8R7imE67tYlm6vqZty0nBBUQYy9fU8W+7g7MLrNi6d+KYEQmdrl765nHHWVNCQ2vQETu5aR7BIZSGj+j8795sc+zu+a0TVqon0CPyX8MOl5M1JVDzcwqJIUcjNvdKQ0j723zL/DBr2So3vAavhCBe+A7AVnfO8BJh8MDKjolINpgsjLBtjMhkZDSK08cw3YWKStiI+hLvX/uhKS87TDvo+I68jd+eoPSO1Ny4IGhlTNvno0A744u8uYF6G0XS+rSNnvz306DTS2fmgIDoyonS3GCBCTNE0yTzqZyDOZ2Pvm4B0CaR3HYTDYWbbz/TFg+ldol1GqEnhTJw+Xg5vVvR9yAXOflTLmpEexg9xwFw80JYkV1E9mvnzwxwbhbA8L0Z86XEcuoWr0baRNgF2VeMBzKgLQMG0kE/OXpbEUYVsn3OMbY5G9E8IW0NfVoQU0FUm+TVwKxA/nvpbqU4JXGEa7EzQxgtoy+wWHBjSvA2eM/Xbw2G1AzyZ3ulhIp3t+jwgd4zB1hBlB85RrWIIie0eNGyxL1Dh4H9yl2TXECxekseu68YDyNZnGXgdiyiB2Y5b23tLQ1PY2VSde9R8EFeDZWzGp2EbpI9RF7nudhniEsGwX+O5W1THKzkEwdK14zEOG3KZXVgyN0h4DuUB4Olxu8BWOe5Fs8QYOKtqOfRkSg8OGeYPl0YWhAvjO1fKOGxjXCjPOj7aeyW+ud7eFWq2I4TDXT4uxn7kyq6dL7ADrwh5aVd3nZgorbXktbcuGM7HSlUNEtvL9USntSYwBnhQyuQ0AC0p9qo2VMqGGbyWSucjsyj4jlf271BP6ZibVsBONs09YpP+V0auwlBVBYMDuKIor6w7wKUh40DGJNNhjjk86m61QQPw34hsRXtcC1VNb57AbFIhuVznp/4ugzAjGKfzRzWrujiCg0a9QSV626bPrN0yaMt4aX6cyLhYoMcNJE02OA43sZxPy7DHy3qQmDPitSPBFo68O69Q0FmYHMGWWN62wKuIpVd9NXoVc/lXHZJ6Hy+RMQZpXIiVULLd33R8ofNpXQg0i/mVlFpfo0j3FKwT83RKeJTxWxhbjG+D/xdfYFEM3Cmd5xugIKybPATWbf1Aqc3PICyQklHofTeaxE+iJky6/br/CsdeIUACP0IdUB5F+Zb5cWDToscXDjiWCxmP2Leu9UXBKq0zyyEu2EfCJCb5g6h0JCqVdGIJMWqqZLRLuNYZqcyQhIMSqUAqR9wlWbBaEAlDjdP6jca2z4RfW+Ocvvc8qXqcrinuk/ZzA01+qaOh8hiP+gnjov8LN/ahhjQsrIIl91IU4Cw4aBOoQfhqznWu/aFcHemfYqh4mswac7BHfGgRp5xSbMXjUOFjgzhOZ4LDZ4sYbGeJe0kxaq1V8wsKmWOKXC4ndFhpdWEBczQPBSXEwwW0qH9l5h7xqel/Xnt3t9g1ERMtlg0IfzCupGPMAcldtEOna/NLuxEb15T2WpEa1EsdqjGMJV6SYG7QwAqx8uG65tZyUyB2rqubTnPyBLUsJ9u7jPPGcIzZBYBD72sQnQ9dymcYHMfTkKTj895GHVSM8jJ7cedlIVXfm3U7yYz8gzomTkvFeHZ9O7k+HcHjXUJrcSqf1FfgIkcjLExzPokW3xCM5xkHtDh/nu/2tLS6wAArGRD7xPwDgIk+6CGwGIf03LQfE4tZgu3CPQsiKYOfz/uhyZVjsMFH6KPKpv1rSasQ+olBPlNxvP6J+G8V4TSL0O2jEPJ0uPLq/ieEMqYAWtrdLR/nEFK8TPP4eqKXGnDqWYwzdTkThSmMFBvt0jP8uECCMsfTlIyIYIgEPbFpilJtpuvrmVcZPvvgzZY2ruy0bnfnTH2rk4SBTxq5crCW5g4rZb9UPSZsPbrmLYFW/AiRsdVwMioipvdHSyPdbkEwVKgqW6RRrnBtC8gyz+59hSLRA7B7XUmVngNBjGnoUrKD4ZHX3fxYleuvEnIvGbnB/1Gvl/E5jwFA90Sw1BFITFOwhpZqUkjlTDwnLLWBUPrdXtE+v1STupzP4fGfFNnUiS8OQK5iiRoymX6zWu9l1i8dFqylPR8Nn2RKXOiSb9+bPuetENw3iQPFg56KntjnKEfJ6L/b4lI2x7NfIyp/kMvaNqejy6dQ045MM7q8OIpLD9Pbio4rtrbH+l6hq9vC1b4ZE1ZTo5FEODuiDUTfbRLaVsAZYZlxx02BJUQaSvsgeQhHQHSbBjqZ7dUUnhUxQWW85a4eNu3hIIBlqX4m7w195RcUbJ1WXdKUOliIkfyApJApiUVieJP6nsiiHbDoZPUo/vCTlSdfY6LQlPLvgWz/+/6PaB4zaVw0a9L86xHm6aMul4iE+OWVViFmsqkE+6grWDirmmMFhFJ8i3sDcy9cvu/nF34AxdQt94jKoM1hDMVpzbnKftfVzbWh3NCRqjXAOae00bDMqSgWkFg88+d9Y62QSk1cUJGAgpPXCNATVT8EVex4K5YTAEmyAttzVSt5OLYpVI6zWFq5kOmCc9UFcT5RlxJYwjiWvwoFf0slJidmnZF4PY9fRQAo0/cazcDA1vuR7GvmPkDZ2nchn2u6qCiL+m4kp9S+PY/SeocmShDkGFsS8+/hnlwLWtKuxuwMqAGDoeXy13mVZjRzRoHdEu2Dvc0R5JsgeY/tDdXrsYbV/PmewNPiI6DIsCIECuAFRC+YE3BkF8n0idGoxf2Nj4qX0BCHRPnPElhXJAqwWZs9smqOaUnptGoECpjE27ERBT2eL711KlUwocBUXgcqF/AfAzUkIjqKx621ZA47Ry+DIZZ9FUr9SONMF25v2f6Lp00I4JZ74XSkgLWfzrT7c39BLPEZ+HRjYv286prChEdFicwjQIiaXgtAbD1eFEwyFIlb2ioSOUqCXYtsRmjOmbmjF0JVnL4wAomdMz85sOZ7vqzq39OnYOvI8T9SEYno9tS6sjR6pnmG0rBEagRFCG/1lAd5p9W1ARd2tOhIZcol3z8uNbi9bWC+9M2CdQWLY+3IyQItH/LWCQmWCAzOds+3OVqGBiKnqq+EA1DYTzcxXleXMUhtoVErZtq77aaXi/Fi7pGbLv43WAt2+jRN2w5zOHLFRdLKmQSx+0OUoi+yxpoUDF+l+ri0tQFPUtOMr4j1S4AkHmSfZWBa+wmwb6NUsETXxl3y1TIekxwF3PRH5n2jzIhYcHzmvTU6HnFdKgxN7kK2FMO8DEAgtY6v7CjkC5lqFHDypIaOA3HkV2ZYck0+DGDrhTLpUFFiFKyWNIi4sgW/7vogh1b4J1FEKqjkRGsRiwS9FTL6K7cuAoot1Yb6DODwGmsBhg26WQoVwVXJEmfWysK5gfjT+YRiDG5ks1PP/WQyWHAbYiX97dVVdBM+wVDkd12Olw/mSQlUg5UHHaQCslzAWsSPP6xfOUsl9N2XLpReOHVZhphVP6trJrQ/dB/N3o93gTTXGvns5ct7bgsMXjnmmDSiM7VAJ/8YCPILbrgLgegI03OT8IlU7zr2gshl32TE484iRmmYA5NpObukxKHVeuiCTXJHNqSsgu/NQMWK9KppbF6Y62ZrBuSMGGpdNB7kdB6N+IGJyaTFgf05hdIsJ3jDheOTtAK6pPcH5yXblkAp01IPoLHJa2VcA8ygfV8Rdy3XS8v/+6fe6ySZ1o6+yE6mVeaarH7nKN8kmjF2jjm0nXn+9OXSS+jzq8WYoAIUDd6/mbOgxYaQsw0xeL31FZxdLcULKvPML3t9dtuoy7Z2wwXOgkkIlytLOSRuzaODYnjE7KCCyNB64uXt5khlrMR/Pl8y4rR9yq4rEsIXKT25FHLKWQhlegMKABIvQmFYn+3FMP4tsRPPiAvbAib2sXIGj/BAZBvQ9IqzoM0nYSTRzYbFwexhfWpUzRxa/JCKgLmlk5dL4+VNspwtRSeUGxvQRCCt4oXTh1RORKxJleeurrtDaZZaLeSOtYvLEKr1I+rlaX8Ub6x+IL/SeUAkayKm366TCNuIOYgnG3w50uH/rX2fUoIa5ZZ4NAD3wypt8KW0ZokMBMwdWrTXhqho3c38aZuiEfCS0Km/ARDIkp97AU7/+XIAYuPqcoslHhNALNW2nv3jzydN7DFhyGV6ELuKBF7K8NlxTvnIINP0PeMTPRcepCJnSq90SyxkKn/hF/BQZXuLjUDWjxQNx6s5RuT3gEROd1+5WfXoJXXhhstXqbOzUS/qMXMTXi1WZCSUV0hRtsfBCkQQeLnrGslZrqRZf4ota6kt4DP4/6r1iqUppotkvcMHvTH+bcaJzr4QQtzilcjCQmvXbMB4D+NRN9hGSY5OsJG1YQJZZwSQk2bJe8nolKCZOuNlmEwMU8w6qMZqWyOR1wh02/UOzBMLZNA4FY49v0MahYienvowL5/uo5Vf3uQpNoOPfWvrfuRX/GMpvdiq3sTm/qJxYBELcc3AXqvGh3zcBG227R727r59YWVl3k+eWVWZ5/bkvbl15NRCEMGFFQkQLF0Ru/wpEU957evNY84zquQs9hFWwnmjFF/JUb8B09oNClHB/VdD6SXiGXL8w8oqC5Ivkzdmv4WkKWHkoU+GKx4yN3Rb2aPf5QS28w/VYEJWAAAIr3voRpd1EwSEJvjVr974YFYYc/x6Rs0Syprnj4fsU95FbUx7J6Qn+Mi/bW6p/jVw3yYIuYvxvKm9y9YWmPS5evECC8E8V1u6QZY45kGYQBj8OL7yu7VoG+xgilroTvxk+SgX01u2sBz4lcBHrtiShmzIoG06bXBuHuPi4DqmrAF3TnP5QIH6+GeY93ROl5NfUPcAASErNnY+6zeurl4wpoNis2p9KYsSFfbzfg14wVjYU71rAS3+UA+Qt+hfyTd7nle5VTjS5i9v4bXIf1nSVHzGVgPzczCD7pl4cmQ7FIQSnl2rCe6hBABWnNKPP+KQAUk4jIJxRBEVRgeToruXhg07dHaRTcTQIANgSj8h1wWnmiSMjkOwZ7fyv6br6G7AAABZS2NJi88d92EF4WgUandIOLwqfFypoF70mbTI1z5YaF69FStmAoIQqPLf6WaOu5AgWE4EaB21FJcRrKemArOm+AAvHtV6KbRikM/JVpLTYRoEaLu8RahsqikWAt1MHvkmrt/a3FmNhXbYg9BAMBPBI5nG1iORWm6bVpkAAAAAuyqUXZZHnm9Yv8Ba4JTd2c68AAFIQ5+se8PRp5Sk+sJUATORS9Agrq3Q4cfpB3bgHScZSRhoajtRFkl7bCEjANRl6ZyeAkrYfpnboW9yl73H0Dbt9WyBiFwwbByhr6H5rYdP1MZhEdqgMKS9NMGKcBay5TgkPwXNmob3T6UzAL6XOQNptiDeU+nmc/pjvPPURsyZqYlgOPraIy9nYELgNULC3d9f8F4XKbU8vk3pZK6O8/Aiafv/SBQjXV/JmmOn2a010ZoCkXMim+IgwJmU4Nja/Ii/iPEQxQLQkwK/5zSjm34QFd+xW8uyD5ImuEWS4+/SqkT+6Q8gsZfFJi5GgMXUExvoPIp/A4Gp93LxkESAInut/mkPB+0nIi0q1jDPJpscV3RcC3Mwqb1Oom7Vn/2C0+o7uj0FqmVNcxovuqJf1RewiXcMrxugL3mpfFuXu8gIDISLhkHB6LUGT/FnssnzBT2svjuUg9U6VKhtktzWLuPQZm4n+FYyHtrac8HS4i53RiIWOTvgAvdPLd7p/Brr+bav80O+KQU6soYbjBoApaCLTVKciqWTpbO9RknJ4ljNH3xzC3atN9THZb3iejzSoS1prIVDBMw7vjBtRqDj/T0Wogz3sq09riY1sj7xGOq5TbalwVG5Q2YbWcj2iHGhR+Z7kP89wqdNeIF+OABtJy1Tsgjy7CfPXeoWBhmxbt9XoBto+fDSqslf7G4dUw1rzJJXv7Nh/Hz9iwNEVkpt8R3jLtoDoTblNtIcFsv3IZwTic9OZriTLJ/7kTd545+q3fbkmwpbMucOMG5/Tc6TTt3MwXjaIbuelF5W0yo5elfoTkYqiTZsmEVFNDdxsnjKzS/49P7lS30Wmfex7cpnjPxq475Lwk9lciwq3PW2kmnsXKBxh37iYSr6URB9oNIAOGyt87CxGLHG/a7eC0HlmOOrOcndFVdb+UldhMBX96Kaypv71/f7h4LxYE6bbndoq2uOOJr787jGw2Rc8vfkay2kFYT/cJf5TxSmjvVB84Hz4+9JeMdNBwzK3WKMyJhTIo9Kq6QCj231JI0WazSYSms6Xu+Zw1RAbXk3hujYru2KTRiNLlyAG9Q0tzrm6doyGraJOFTNSIi543gogqUPxy+m1aLtALdbjveU1qbVM3nPNdcX6NvyAvWh4nBK4TrQChfPrAWFcDgQU/1X6hArpgXg3VuU/gGxE3RXEskC99Uql9y8R852sDLZEOnbqc1wGneRQesMsP91rzu9o2OFwU2Z40y2McLT8C1SgI+t+cCrW3w3ttwpN5jIvW5VaRz8LZXx/Nuu74ULcAAh0zS+Uncjlrcd8l4S7Vk564x/+M88ig6lGhd136rP4m/zdCHfseVepwkFiPC45dKQfp6Fv8MtPuKVIf5rUajlQXjPUwBCeOH4Q+mIffO/elu53CLRU2XConNE37MGoNXjOTPSXuoy1/z8pir5n1CPJR71QNxr/cUgIfh4cvn+/E4SdLSOvtB4sAwfhLm2uYL8JMG3dLCmOLkmn0qVkSkgl0TC0Zc/9LO4ZkGcITMv5T0jLgCrG4FyrrBrhA5eu7aMl25HbV6C7AIQgqYg2ON3Xt1l56S4r02Ix3dq4e/ATPbJbc16vhLtrY4g08OfjQ+RIZywyV/e7kTW5c0/lSL0+hXlqT4bHv5V0btluOoXyiHgzuAdS9KQfljMPZSlpHpYY1nfDffkP3yYCQXWOhzUTIlP7e/UPRdykDosRevM3Ep+Xqw5yRalxL0Hz7ZYtq3FhqL5d/TxZjh+yddvSCHBpqPinffbXgfBtghpjVdaixiaScAOrv7XzBlT4ktI7eMa28Y1To/CGlyMyESvZvOw864QBWsV/IfPDqO+5VoavsVmqIpY1hnqRbExR+Phpoaj+0t0RSF7/lnXUL+t8UrJGwBzgp0x98R1vQn10uWwMrMMoGj6MBdr+sdJ0aP/vPZ1jmCNlItyrLcN8X1BnvUJ4cNaLEk1JgfegscQZcbVY8JXcCjlViHS6BX6eQP41rn1mB4zd+5cfgNEU5hvcl/azOCcMy4o3C6IvzJBtqPTMLlOC3HahTGygCwTOuD+/3kww94wJKt9RXZL4qaoL37C9It8IFZX6xRgBLyJ/Gu+CXFvZUsDuP3Jnob1u51X+MSJry9U0Hn3772CMoNuHK3vdo8vqgDUYOkX3pmqlxjmLo1uShBTz7wVvhTVwuN6KRfuoiQFW6Pp74rG09s5xC2Mn5dJYxjGvh/Kt2qvFksPS2yIAR+2+ZcYC3vX8cbqDbf4coTSXBtAGqBLTlYEBfElYERE6KWqjVYzB8Wz5U/XQceZDROQbldEr6Q1d9DRgk6FSO0K/WB9vO8sdbPqg2+q/lYEjntZV0SLolQZYG26iT156qnz/ECiQtk4Sy3BwiR0bfAa9amzip4Hiv8vlMQvDis+x5gQx0qZqfqFXvr98jVBBhJr1HoaXnRu6xI4941Z3CubTEy2GvexZ+gCXn9Sa3ESxw3+ECEMHYlP+rR1EC89ifrrWV/Lm9kuzdZpKFRQBAWZFb5C4GykZ6isW9GlRUipSVpuQT3D7dj0P8Neijg+Nz44KitRzPQYf6ymPk3yvaRWdstYEfNUo+z/eufM+DQ6MwpNKE/RiF4yEAvQCvjEVIYgQ6WAA3GXGVNLfIdMUOw/AcDsZHoIsCOkJgw4nsBztbgqmRaIcZ8icAH5K5NfGhsjsagBw1pzQKmrj5qfU8OoTz6v7NojWGodpKx6ZgW1OtzdJIBC4XklovYc2fxaxkuQzeW5IkxcYbW4tCT20N1RpKw22ShiEfRQ2s7gljpNmKdQfzJcJHB8QG3TyxYgff17QyqUIYrDPULo7nG2tKC8XEr4/kiVcAxyTq0GMV4xRcfHunt3YVDcWYoZ0ehVA0XuNmsiKZw4CqgP13N/XYYbwz1y0lHGH5g53zQbw0olrsPKGm2giyInOni6C9mzvaPiZpmoXkT8ZjQm9fPjRWet1/h38v0vAHBpRNNjp8bWY4OPuJo2SoH2GUsOkkj31sXsX0wDcS/H+82Oayb3AZYHJzr07glAiRguselNUSppo/7bnMEkrlaERJXMu5aGCaQw8KL8FD9lEwcr7QD4XcUdiU1Xm6aAKbMYaW06IHFD/9kXzug+EYYvIv8N28s+e+s2tgy6KymOMGQK/N9tdb0v3ScrHPeQwARwBSpEYgl+X3jvNjWTJi5OM51bLhoX995Ch8qRG5IOoxf3wgqmnzfZeZqrfMUE5B3kQ4tU8TLmjU89DlKcVsuRzAD4wPVzmmzrCG7wZxIHuxvKO3SuhL22QI4FHdyD7755ULHel1+pzK2IuGPOUTrM9Q5ogEyeCA6ZDamXNqWe7ljgbQeLSiDUn+nhfVk6f8IBfrn8qiiKfIkpkZK5HURSD6XoXWTUZ9+d3AKawVQcHwZTLzP6Ni0GcctDTwsF26kJxkEoPz0ZWHP+bu2CupvvJO6dmpm1MZerQnlG5/sv1AGyrRHmQQ11SBVeZppnOG0M4uzIBSGX0qw+MzU+gfM13YkOeaJM9xSAWNYFy7ileSTcWzk1CTC0+Lvhv+8Z1tX+zU58pWuwUPMtRZliOGhSCPFiZ960Z6BMuVgzLyJJ0bIKvCYziRZuWa6HsVhg3eZxAfIz8v0z03pOR4ZM87zgvn6tOxG8ebnGijsv+ppWZ564GdhXfLh2oLuT8PCLvyYfyi31gN2Xoohp65cKviT8ceX6L+xAXRpLyzFdjnn01qLjYtNOq9hEySr2/d/Gs1s34pp5zWpZq/Z9Bg2s+JYh9ireyvUOH2SDmznhH95nqPky1vvCemdyU4JcNedFsbXFW6KdVFeeIFu0WkHrXCKpsfw/tGetmUC4JNy8EM9K7PfHHjS5gGJ5CS+PIRYBYobluEo1hspLYeFfcGCapI7v1lAeGQ9DX/vkx6TsYDoU5YF/uPJ0+YVJ+GyQrjm1BN6a0GIcOGM/MDnEOJT6dItl8t3wpvv9cZk7setZQsEaQXlmJuvw9TJi7spLzq9L5ofKH97vpIuV4J1bL5Or1phWg2qDnfp2Ptaw5v3jYucjzajX/sIbIRoRaPOJzuoRv682Tib8lMqFR41OyVIX0mLVMPAsza7Fj8GKBGSQOJpsqa6Pm0MFMU8deDXsaVQrngait18zU+JEdsr7o2Jvqs4sTkrZ7sYGBoMicg0W2PvyUJ0oLpCNE6KGFbCkI7rCr4++r+4HyIJhHWCxQkJ6FhOPaVAK2+3CgW1yIUPEHOAfmPFh0V0dHhcx+2cg27crGe2TVjyXIxJuG9bYSgvl3ddi3zqG4OAf4V1yYLC6ywRMcvFEq+QfrOCNkmd1XoKSRf+H4XOF2/cvNfRhrPmX5Wpa++13naJSaTEcvIcAV/KNtuKYiuiyMKenNGAQBsncYoxrj5OUhrTAIun9RiE9iH7Wwijl/PoiPGzYHcUYKd+d77n+OLPRSJ+T9OMIFwouQ146jjYpyIvXUqb78DdXVCoe3d2bKbiUvitWbUWDJTbxLKOSlm2n29+bxqf7cFW8KtNvDRVJXokOtWmq6OIwxnYZ0IpSxA8BPxGgxvcY+jP99kFuPHiub7b3AYXaDsEmcxhnELQnN0y6uYSiXdpSD0PuhDdVQA/D1qnIfXOKny4lyooGBD3XVQishYH7y5JV4wZB8ByhRe/uJR/8/fGqbZc1Iwaz5x9z0SGyi4Aib1M4JpPBWoH/zd7oip2pgVsFP4oOtO2cZTYX2gZrq4LyoeiZOxh2X/zuOwsTjp6zop6vkoOVN8S9uL4VQQuSWRF4oaS4i6KGpu7HITM0lHIQs1R4HabfSNzMrL6FTTw1kZTmMVnrQb8NQHpydZI9i1/zliZWD/Kqv5ZWIpqr885Kx/FeH8MLKmhgoXyFCNXAtCjhvenifvV3Az7q9xX8J7MfIRMn4JXx8VLtOeEnMMuo1T7JT67FdBvLJPZQZijgsHEJdIkxSOamY41l76YDTGmws3iOHYljnrvSRqfK7Rb6pDb0WmxP6EYazSvG3XF5RLgxCyclr8/FqCz08Uet8b4hcU6KIH8mTnCBNHMUFzdv9svqTf+L5IjG336q9iFGog0bR7ijLgtjA5ZVPwQzyEuWz+MI2uJn88OglrXJfA4z2MDqhYmZscOUtY2CITObXqkoLX1n0E1rVrbn5JFQYD6oqqhbJQJeywI45WlLZlgFERIMQzYJSHp1ZvHfCbeNEu07AKzpiI3GU+DcxNPyTO17335byN6EAq9CP0zFciJr3d1a2pR5oE5jEtj3YOZ4bnJg5qFsTHHas4TJVV0acgg76H4OwmoJEAOEfPG9jvZJ87CLi5szOWOmHn36uDIMOkQFz8s8UtctYG6jcQISwAuKAOesUYg9Md821NFFRrgP7YZgky+TzH8oG14QayS7QlOPW0Kzz0gOl16DqvDvGqZe5XBF3XCB1NjpNuxxJ+piqKRYDc2y6LPiK1uzio8+zrPZquywNWmrCIB9aMogQ0O1gLsJaNuEhN6Pr5qmsCAinyjXe1mDUcY2HytkGUd/3lRhOyx3JHw6bJbXyNAb8l8oadqRl5gMyMs8C0m2CovaYQSE/wArKIQGFghd/CaASu27BCOaE3WBxf4cVC9ZlP+SMahb0FQa8sRpXpddGnWNUO23TTMTRUvIY2EpsSpJeNded/alpUwFTzwGM8TXkHqIbCAJx2pfILeP+e9kz7YPNoRha+0PCsYRe7Y3RKlyENk3a6ZDm+JSaTSjy6tta3BYngFtHnXikkIMhptd5LATgXxGflCo3x+jhSyMFNgPN2YTFwrOWPP3W3DvvDpWyPGwmReUJbKqV83CqRMbbh0Ujd9aL4ZNtfq7D3LMmvfIOJjSo4nAQ7A9L/mRUZN6r1SvnRR0jK9LV9Aj4CZDirtk9CTozBLxPGkX1r/TCfmrDW+hMTlDMbKPvYgiJ9Id4YU4szSwdTRAkE8yAwc+3xKoDHlGizky8g0qyV7qwGCV5xqCepbpqRe+HcF+0fL2FB7U3lh0LgK2FUHkgch5/LYHsY1LnydwJhwGxAm4uVTq5GkQ/601aI/RR/Kwzgy6CR9+79/kTNMdb/Y0ykTi5ezftv/85Vgs9YJlhta8Ed3O2Yaq75rXhlorMbCs5XF6F887av91K2kbf7tVRq9aMp8Mj34u4vaYlPwmaC2A4nCaxxlK+yQdhH/Jfhhsy31ZSFlyzPLx72QJlDX1UvYqv8Na+WTEIRN5rvtO2kA4X6LEjyDFajiEEZ0bccIvJuBpcYt9lQ+ccdbTVfDjb7wtf9Zs74NdxWeCoPkj6xs/CHJreE47WR8DpHDl5o5XWrzyYNg4ml80gIW4I67bMdeE5rg7TqQprs16qVKlIrFpBCBN7g3BGCOK1W4CPeUpSXUvar/RhD4nLJ1S95WZ07Et5ujOx0MTxdYDTzCZtzINJCJGYHoW2Jp3ypx3xXMYt1e3yWTxyXl5gtNGp6CViRajovaRvp+OnLc48U6S/LzhmLo0F0qBPq4Q7tFe8V0fAWEyEgUv66O3hWRTcaM6knDJFdzBNKn0Zlkj/8yWmHNbVkrJBruw7owHxWSHf9rHF/cWz1qesnfrjG0Ogwl5Hm+L7f6T1VKeLiaXO325JA3v6mxjxx89k3PWKZT0l5TFdVR0PiAXinmrtOqXHVLfJNCCnowW7MHdIAsTRL9kuJkvnmFOLYrqKinCsMARjOUpFILgTwIVdq9etBzCEbKJpLiPKLfwOOCD28VJVvYe9aFEp2p7guMwzhOI+bik2csQWj8Xh7FI+pfNFkYZJPmuvmHK6Sy93m6vYnA1RF3nB/ic+zKNxhRoUJppPUU4N57j19n8qOfAdpcdjVovulhpIY9qhCaDL5Fd0TPr2yICBdSj3JRn52nWxz7mG/CfKnK2jZTzJ7K0k7VOTJRH3NLMkU+yLTYT78IO8RizTO6v+7aRbX7Y11QL2/pKqFov6nqoXmc5cax94niWAPrIpJTa+6FYhHHVvpZkxy51vjWRy4yLH9GVqcx5pGZ2AywBf4UkE/gqxxgYR5ojI26fsRxBMAK565cs7xVnTWIqbho5Y7ogr1rCuNiFD7vtib3ylnrFxIMm0IwdGCScWMZt5q2yS0r5fRpVdOGZRi6t/5YxSqnuKS4ENu7vwhJNsAh9x5sOtjtFb7E2/VYKdj4zBhIehQnp3K/tgB82dVor0+fvK5eMpu3f5m0O+kCy28w+ZtjgPb6duWXlpWgVyAWHVpgTTw8MmwTQmhF71Zdrfg7yKjMylu+YA8BDgCSxM6+tdW1kRaei/VUT1oRMS5kQSjZAAOzdUSI/LhdVXKiYzz+m5qwJkp6kNOPxXxIgMoFX38QsJbBRKl7CzQEe6udJOqY+2mXwhb97346OcJ3ahecEpttvPOGDacJPW33oZZrjiykH+DhgZQQLlSdE8ZDYu4Y49w9Rv8dEajJuUOWLTbNqZlRSyjRDXPg5Z0l1Y+Bsk56ACKDxqkPjj40gV6K6qfJiyzDUXqfeIIsnYcYawrMfYP7ipJC8rG2feanOnqr4hN660MLzRxzxL8YZOLrgkAAPXBmaG1f8ciNiKDIeYX1I8oZ/dnkndWvi8q8c88UwzXSo2KFm2WVdcv4E51or8P1ox+p7Gfw5SzUJSa6cUfsBS2niOYfGvTcYuX8r2F3x8SmSMwVPSEdjeJSD2DB5SDuZXkP2qL3CtmfNoLf9UWAp9MlloXfc7GFv2BYW5Bi5vDv4bLddAdZKOE8oSZ+jNQZ+aB8tYz4kwVnVw/Me3tGQfFxPnL5/ooCQo2Yf1nJAQkiGgtP1uUveOFNLxmNLROfd4n2Ka+/EKlJAVmDR2iXeoIA+FtWeQY7o6DLQZ3YVvi6WbeDxiwvz08JfjIo/5iARWaEfgUCVChgIBTldm1fheRgYHWtos8QpLX8yFgEcFnt/HjS1CWtmw/AFDMgPFnk6ZWW2+b7NiMNkRLULBlL4eXhookIhvApwkZCoWnARHXSc32QPa5IxNVmpVoCSfFfT9sC72w1t3aEVh97OoGNnZZhRbes3bbjOFLCRh8DFKR+oZer1arDKx0WYPAHaMLEORPuLaCSB7kcbJTMRyjdzdH6fI6SWoeJdJZpIqOcBew8cV8KMw3Mhui9hojzu3I4N95queVle194iH3GF//WSE04zo2zyBx1QaUcqC3Mg4QbO2R0IApYt6QxLn5R4XY4h7/NNPWUKdiGlDLDL6LxeFX8komYAhGuykCRzS0j3Eby4/r2gNZSmmZjvSwtFZfXzV7hVHdpZtG16GLlNnxaTPTlmx6NkDlXMaa3QfXRjAQJWZe2m+HVe9fkBM5iXUmY7T+6ZyWTPeyifcXyn2fi0eYL2pI1RQ/aXXHhCsgjK9q1TcrrZhhk57Ckeo9jGPr4HNf9WMs10ENrcX12B21CnEmDVwdAvc2r4zRx2Z/4ndxx9sqmvWItjIRjOSHKz6HyOfgI4cnAzHMJSrVTYpRlyR9DrVGrapkOnRGV/SzvxkLbfix742f1uihnT7iN7IUMKwWbErkLgyONgNfotw/AHBqHTp8GqTXLsMV0gjTnqN5mm2XSyHVDkwrOEL8lxcPRFgxV6i+FK0Fiaa40fNd64r4UjBiXwyIEuiFRhP/aR4t9XRJTkXwMGlU/hvBkxfYBPPbIEun9yuOHTcrrFjaPRsl84UAPuRMRa7noVQYzLdvG8KNIpRbu5MW1msE2WjZXpRcyfX96WVgRfka62TKqwx26/LQjMwHTKqlmvU0gra2/uiA9tgH0d80v9MrY3DjNqYt+I09nXOnIRPYzrs3/jQFbADtX6ckKPtpaY41gKLblgBG8rC4PYhcv4hu36IIA6lJFnnluvkyxVnEoLTYLkfU7gQKSSRl92jD+uCVGiNpEhtyEKjPL5nVfeDfcWqx/SNiuKCv1FFSRuEiD1/xKNvQYGxlHp0SGq93jVN4jmWPp8mmPTnBpUr13NYNCaBfsSpelZmbw4iXi2BJOrcs8gxHkYu8bSftZuDe6p4TekYWGq33t17b0j3llG1j69WJLu5QY1rr34RUcHZ1bRXLNY0dzGSnOcuAJjT8YubnHWytpLt5uFAix0z2oL/7QcvwtWZDhE7KPYS1eRlS3Gw5WuAd4V07kS9boI4faf44nkKB+eqsW6s032ML3YlJJKbVQaloKu9pNwFOR0nfdOYbjxhO7/wuYDBD+e0YRu3+xOMfijeZKOQjZ6T2onqDTd4YmFWSoq7mhdUtA42EsbNyb/I4GFhL5Le+C8vnM1ioDpOwP/EfA01yiOjMd+NSHnAYdgIFmDNEk6k6SplZQQTOTGaQZAHxE947O1mMHdhLaKsUeCMk9D3EHKnuhdTg5w8X+sDZ1Ma1oZc2eE8+5gesU6Xtdf0Xofvz+fAmQEYLtEl5m+wf2P9nZS8hpABJU6J1YT7dbI7nJUIUPMEAC1XyPy+bMP8GdbINFb3CnDSf8e7Bf2phQtvF7UjkjAu1uyLj8yBXHrUfLxALQN2A7YjqC+Pi+JZlW41tEXgHahBK38hefyI7YFE7ZUp9S9fFKAZn37mAyuzLiapKOwVsAOlH/fqzch/C1ZW173CL0ed0sE+KEoeahBxuLA4xktI8W05eh0TZQQGD84NZcxocLEUgmEZLVYZEwAD837jkolfus1yYXJ2zk4Tunx+h6F/uWlUcBNGFQNmb9bdhLNeODXgOdx369yGolBoA865gV6+z2huWLSJzrmDHZgct7EQbgDgOgRr7cnXxJ5Lxozr6jO/IsN/V6cAh4oFifC3byLTsScVT+rC1NAU1GO/EtiCSI9kiKK5y+oVx1rHZF0WEozX0dCkfmefUWc1BYg5Ll08Xm+95WgaipZLeFcyl0U4uiXWVq8cTKGNCNCv15vvn2putDRpudghgl35mAezipJZ7iInYWCKeDH9G4PEuDI7xfsdqG8teA7Zx39XzBsrFJkJ/tmtRzvheaADTJrS7lg5vt0glL+Ln6dF9yoEy0mZ8yFVdzAw6CeazCqwUYZuPK7p+XYpVNh3LXbwrG6oEWtEAcN4vltI7+AdIXtO9oVqG5pwZvTpRLf8M1LQ1p34SSLnFBE8w/fERPOAqSqGksI0e39lEtPguipf0NDOOfZeCaSs6UfrpzouWSGcpTHBcoaq/rufLYNBLg7M52RLmiQWEs6tgvrBRzZpF3Oe0DtYaPNSRGSsrUj5XBn1aIfNGHUL6VJnVYlxj02tzcq6cRuOEEcf0Zl0MUBioJWdsMXC8/tb4SBPjeQOkFxFpkjsudJy8xpSeiY2/8T1cDypDgpfw69tcpfrvrkQO7/PfkJJQSbL8/LIabmdWj3pPieKhnqKg6bKlrVe4mrm3MFRpXn2+xX0rwYosVnMcOUhK5y3g7epVRzJt7qY5xsVMox5/fuEDRuMHyDCuyeVv+IJp6fV06sT1CppOeJR5FSpneHFPWNz+TYCQg9PtRQEc2QPhJ4yTkQTLtYGHa5kJKW7AZka9y+cNurRLYLAM+IkY5nbV8C4JAEzzvJ8x1V09EQmUWFq/Rj4ylvk8yodRXvvVgYUo4foYNXbdruVYP4A/4cBHAQLX7WB7jWD84e059CN2Y1RtP2fhqceBhDU7ZKzAO0nFP/S9NdPAofgq5K3pkP12AWpI8QIU+f26h5OkE5Jg9UmzG1t+SKqeClX9YJ1hYqfPD/UKLi0yy2nQMBcXTLT8+yjoMApOBBZqw7aeLTUi2ML0/7Mn6bYa2QMZZpFjYQYLdZH5Juy5XbFL4mungp3Bpj8vAQyd0F73SmHFCsI73NwDz+DpmF4Diux2dqF1OQLaByQX2mAqP9QzAGiz+NGieRn1UEZP7qSt3tUAN7eG+Fd9kKU1k29Yy6pj210tRhtrFqZEBWs7mmz3t8G6le+wK8iRRHeU5I7igDOclT5rhbZokNL3OcoqM9HOMKfpkYdD8MygQDXNsN6uFhOt6ZmbL3Lo01n1/CBB4nYmvNB3UO1Lk+tfSueltDC2aYKD6F7AdCjeoZQjkc7+6/4AO7ap1RS7MlSfP5pECMp40Qxv6GCP0YlUeeDKZNnfoZQweETBo4oXK7WNOucTSiD1FvjzyRfgGSS0tNUY8WS6/GwyKgQnK/w9yiCDcykM9YwqdYtNvyBHhz7Qyepo8OTZmnF9ftJI8anEmyspP0lY8VxcYIjLeJGMS3tFJqL146Yv/zKmpTgwreshyltRZfKEGu8k+5Fw22vlo53Cc61GSxoYZBUf2bT9z1rvNiLKTCLaPIQrgUjsKXppZKCMX3hp8hip01tSQiYyaUSAmMzX/tt7v3o255aDhpNd/lCdJ6LPIo+x50tQyCHUqTspawt/B3JrRn4R4CwMesiCj2Ug/WoYgO2L4b/EjSbbCGm+4X+cFDFrRG2em09PIeKzDqCmAPWvmOuRlokkpM5wQqWSR4uGB3VSmQglQzQnNAUKuZygjC5YZih+h5wpxZKX4E9yAzEo/+MJEX4UVAOtA/TCd6qZrvc1niO3I8llnYt6jOSIoykJ5tdt7nRvB0msdNxP5u63RAFjBwrxtlbMXlkShwIafNJmY8cM84RSzcq96ZeFj9v914xeVtwUlYL0cFr1HwbOtden5U1PrRCpkV0jnEaIPqGIsAjJhFLfmflpKT2gHxend3pL6Zifnl5xBQWLFzy1tXxi6JsT+gHXuVm2NKCIKIsRdDZYwXPO6tBE0WxBwkstJ4RYPMknOtpPDmaw2dY2lsLrg/cAQb3aK60w4YR3HDuuJxsidKRqatlD1MNDrhnQA+LX204Ws66naOrhofJUlRIAULCFNdb/b3unfLRPs7I/2d3cWfLvyK0xnkcLvj3hASW+tFuCcaJLrMoetk8aC/4gNzQjJ5nOuMMMrf/dK21YfUCFST8QNEhhsBP1JEWvBQydJuQe3zadICmuW3/htwAz37usI3oUENgu00vIS3/Rr6eAgE4+/LgrmBmn5eAL+n9PBFhtLgeEr5m7Rwn0h3IYRIgh+GVnn2Cn97FvGjd8EpKsE/tC/mVOyTXXrSMGzVZDuPOaJAGA8mJKh3roetHIQiedKpnd+dUdum8GIC3I8BigwI1XEoBJgGkAixJQ5o01/MvG14MqZFSbxTNIiJV3hjukgTv0N/fp6aqYhehjS87ylnSOinNj77eJ30UD4S7JDE/SkTABMdgTNsDep0CaGnU+QIJ7Z7a3wQrB0adqEjHGi/111KXAA7tlJmVxG+zSyxNglIHz2WvBIt5IisvWShs9xyLkC/Z0XS7e6EmEpuKAqHhpqGGZrBiTlT2nnKMKQU+FY+gYqEKMDaPJIHVD9fWVR8+qrWHuXOtOYeY/6+t1YzJRsywWyyWK28aMs64JkgEnVXEjttVxQtmpoB2pxXQt62o0NlFny4A8KyaX15V7rKdAJI9fcOBzNtclsLjroB577AfhRNznFLkHUD+P6YHgWVM6YIo7AXMUZENoW4hEU2zEGjDiTuhGjMAENqAscspHDie6/HcVQOXqX+CvuVyQrkea8zJGVUMH4HPUd2bVaRTybuWXhokz4J9DuS9L1O0pryqRoyaBfJgKW0lj8hodVtTzaAkOnA6MiZe6eVb7YNsbIw/MP+05nDD+09WuGEiEIqMMAq1pcdbBeyGdNW3DAHqhvczQowNz1bnZe5tjkGq6jjMqiwLpIUJiJ4JqwgscVr/Ng9EVBR2OtWoRvyyCwRDgj37ac1toz24weoR+5My44ble+GcrXyqy4pOBQ4NKWEZJuEnvtrzrvGTao5AT0yyjkeO8pBjKTw6BpYTxUAAFpAPhhFUPfc9lj2a6MqYUBAl48V+x+ogFL6U/kQIaSkXRyl04FeQ5CFWWvho06fB+z2xTZzpoIzAZhHnbZesVwqh/5iUy9bDxc0r2PQlFe/+xBx9lLgBE9J89WFiZ14Onm8xDQquzb5Lzpqbdor+yN0WQhoGbNuYrWPqdEWa8Ry6VxdbizBp73Qb6G1KQlEs/IU8ALcnHzEy7GtL1ALwRGmPhrrwCwsk9FsMqeqChLna9BJXs/uOGXrMG8Z1tjBtdyRMJHfCT0MUsvAC3qlEFVZxqj9XwB75mCoUjouYjj86E3EfDyP7ALDjABADltNVKQ0zkSCskA85pdEJbF+Gc0ZbFvxJLmKTqrw4CUb9cYPOM6ZOyqgIzFXs5Acr74InVCYkkxo1cXHnCH6xpsPN1zG8wNCW3fHlgW/CB6O1rtnoFlTNMXIxz0wduNzxSRiwyl9dQffZxjPJChZhuRaaBmVGKMPpTG4zNiW/MDnrShfik4bt/kk6xE4LtL7CwovsynlGtJdoKQmlN/ED+KJpJPC+1LQ4dWgi72CKlgHjsNPnwLeQ/gJ8Z9m8sksxfDut7GmXWVv4i0K7OtZW+xGtlS1LmfO6Xb1y7Na4yTeli5UoqfBHN57OK2xg3MgmZ2LqWandLP6pinIPLOYpP1adY65PQsEDlRnI4p1hfw2g/ywgLMYB60KGtl37k1bQKhkwY1XG4r4y/vSHJHjDNxoernQaSb8JjZjNQAcWB+9pznRvJCTm7/5b2chohJzjCduB3jjLxdPjZCXZxvaSHvl6ndDj9XKRNz2+8TOLlh3+T9kaRYZFwn0GVV8+iOXzkbAvdP4OmumBL91csNLbOnj5MSCnQy03XVYKXumkvTe4leBhi/kHA6ihDJJ9+PquVGuPVEQAyytxUDwU6DsuFqR886/JyUP7rZbYCXUYuhum9Go3UCO20RczvdIQkgUDL1db3pZT+gSdTOi3P9Jm/F2Hqm5UyQH3xMKhWIuA1i2F+qDzd8sMfWJBmcJzHKcFznWBL0BLdGfb1I+YozFXouMA1moVsykgo+gRdZXouOJCkKJ7oiOCyaRLhWbtgABJBR5L6kPdBp4K8FxLNYyMmZ/azv2Lwd+vjGwosUgVxsrxqAmPDfYPdh9T59+5sCTsIaAlYHqUZClchhtpQhjektf0W6SXwX3w2EC7urganPdYtb3fzvPfMXBD/ZRaw5sCssZMnZydw/20yyhRuTQNt+zjfx2jPnl6wFAg+Xbo1CDfk1TAzZg9HC8FH0p2G4sDdRpHKu3gb87C9gU6w0oOpY1DHB5ZN7gvXK22kbXwBKAe80cy4i+fehUBWCl9yVMiyRqJbDDxBgfopFmK6c3l6paVo1q+VF5h3MVP+bAPxbiji5sF+tGvk0u3WBz80vUmkH3l4gd3NyXCu9GysW7WO55esbFUtKmyMePq3aPEGI/h/fyA9TPKrqiL8DZ26M1JXbRmHfQFTUqAHPFEm3cL7G6fjGY2G5yIy3ROzJmI3eAXvej8kCJfPqEIR9950xMW3ptcPxj8HbN6Q1g1dhs2TKrIxiYZ9iKTUmpg0yuRyV9xWLtk/woiL3OFyJrdabqULzNidXKAk0A6Mn121XiBOwlgLoQX+VAdn7H8qFkD8aDx1IOYZxM5KQD0kkowpdiV5CM+yYR7vpiaF5oiaHZW1aLjNayv8c/I0vrgK6t77XQasVYP02U7qAbhgPaBAdKn7IA/UmpXtS9mkBIeb3mFSdlA/CUhu7Z1MpJv1eOgeq4+8EEHAr0S2Zyc4qAqU+x+WE9Mw5DtAUTbN5n4q0H3SIbRd2nw4xmLCzx2ctmRFyYAkauU4kUKuh92RmoGA9D09o+9yih8npz26HPCGOvIv1NQdgi5X3vhQpFJBXOmgZpWUvhIACedBQk9pjddKoVuM/BhUMzEldsFJGnTGOqhY8FKJTbTlJKFOVd8I/S18pLQUqUm2G18jSJhW+uWjWjz8lPwofaBz1gZR69at9sIqtuJjCAbQzoasfWH1o54cxgDvwiqFFY0KvG8+Drw+WmhTUrXPThcK0dPyQaRkkFjy6Q9jowHqT6kxzxUwTx4c0LYN6H4YXuaBZMzaN+73nNlsfHyrnYuYhSR1qTti2S/x4TeSZBItE+c7ZnxF0lUWl5r4cDfFzGaLjiHzHek1eSeT7WnkYR819dvBS/ygPwZp+M8T/lQC4eMv7kPn+4uYCZ09V3cHEco5P5gww9ysZktDYImSk2TQvUhEwQgl8IG4XIAdK7T5DL+yFuStc57i50QVTX7yLUpKfFoA9gqNx2DXtQ5Uf5B3hBAL3B9/JcodMKde79x8WmqY7mCiaOApvuZopTDlqSsQQVK9M669goDZAOE08Vp1EBOyEuy0ebzNlrJilutnZhQeg09o93KNIBP1L0avpU2ghx2GlFAgzIPfoHSY+bQ/C36SpVa56LXbvH9lVjD130cEg4Wf/OZjgWr62IryQOQDBx8s509AjQ+25gh1AJpEu0g7n3rUZwEawXyX2JpstBFwa1gXxS1+L2hqdNFzpEx7QpmOuYAqAn0skJhb7doR//e5AHztjm4tUym2JzmgNapnCPTu/1rS+/OcBlLYOgqUiV2AG9YCTtae2wdsYCa20mT5mCQCi8hCvqnNhnAAZRObk8Pf82PHo66WbeMNf18636bRd9a7gkjvDd1scdWVvC48KDHsZkNkYgpuSih/TfYDAdEeubv5QmTmHecW8Rw0+YXg9KeEPuFZXa147U/FTtJ5cjl21EcjrSaAuUNhT10UMRMCt8HtrGT0VmyNGtPVm4CjIZJMEqAhJhZUNqRwNR8D+SBft2Pb/ce6NQD/ObtQMS/a0fG3c87mdzhOwmfOjBdrSrPIHrr1qHA6ohhfPy9cw/eRfq8Eql1+PQykQOxMcYwLfNkiCj96PtT3N01VywP0x1GBxYeQK3aIMXIJIvGfDOq6F50cI7AeaN+NjKUJ/y6aeuw5dmkBAJyxIfWgO48DGNwBfurwk6XQZXM4yngH0Fg0nXnv3aHDxF50Y0jAQxOUG8GKq9vWW29gWff2QcL3Ar5gOae6gipcYRQ/rZpVmivrzW1sZ1eu95TiWtLOtdynKVrEFTae3VSf0QA4WY83he3SHIRsqT9Yr7m9tf0X4jWYlNc/nYZ9KlMDjv8oTpDAd58C6+MjZmBadiEZusPib6hv97oyLicInjl6nyw4o+IoZ52wyH/13Q6nQyGS31IWuiL3q5fjogKR/4Lzb1dgtYWGUN3jasHdHVnWB/z1IB++APkJEZFapop2qeK9sPr/wbA7sd5UILk6yaeMg98PTZ0kff4pxzsZwWxJ6N6UJA13mNpqWeqZHmPiMHvQN2iP9ZAICcYMIAlO+fJnmO5bptXi5j3AnGhmn/bn8Hv1D+4Ea+q/VJRrzLO6tVIBCg6sC91eU39YUIwdxtcz99Hou5pBStKgfLWxGUpj5VmkQTkyfrZMaJkuEvGhZVnuQJ9fkkhycPFLyVoroY9e6UO4t7NxXwkpQ/LVcGREg57GfcO5pcIIuK5YPVsYGKwwTZp5hJR1RKBdpEZu2OEL6ZlJJ7nbKRlOoGAPzwBwAcZCStdGV5Zy2CBiKfp8hqibkBjm8oLrAYwSfulseFLLCB9nA43oqtxT3SmxtlGwsBhZVixNkSn9KlN2cw2MIR6g7a9sfxQRboj3S2Mu8e4oO8fbSmoUtlSveqbw4L4qzCOx10bkA2q2tnJ9J1Tuk2pyAQVX2JvgZQnSYb3Iuha6PjrafBMaUFqn+gB7ihgJ2T9tzGGnhUrp7N+Z6XFTt9e3LUYsIZo7HsRMgjWlHW4HXQXZGO2IaDK941wN9T7BuCHnkF/CtSgsygPEfwSLnGN6WP8CrPYftx2TAhdsTh++umNplpg08Q3dmMbjjk2XMjvMoVIHji8QDdcP/z5OaJFeX18/ajUOJ++faRLDk9228kewh+i3SQJXz5MPeAr8RERCQL2mVaJzeRuVo/202qAYekuGlqXRhocuPgJP9f5dRB4seTLN20kGQltbxRZb51l2x2pkEeQtlckR0UAP8mCmYpL5UB4IwBJxLPPtwBhmCAnO8fWbyjcNgGjivVFDopDh7noBRqvIeYMVycKMUGzzSjGBpGnZUJBcfvjSeX4q8dxfKq/itPlAJ88Xw00c2fwnd88sfxobs6gpaK/J/oGFdBGY1gc1Remo5YKiK1Dilqh4tqWd0Rt+PfNy8ACuYdm0ftU9xz34v4iP1HO0l5B5f7EFPQMKpiRZmLjyB2n7kuZKEW+BtjjPFuEtD3gve6BA49grlNUxA3zQ233Q+XxPxgOnugKmauJq/5ofWmkanHGf0pHy69u4PctgsGz07mlAirkZGfZiIjs9i9SxIGsfkMFLnREt7hfOjZD1R5NcfnbWfB0O9sXlGgu9kt49zi+t1IB8+F4rKN+zxqS//C2A3I2fkL6sBPpG/8dDgyR30ZbevxKtGcox2IMjnE2Yj3+mlerYHvIGku2v4friKWxmB1p57QqwTqFkwnI+ve+Ve3PCeFKYl5KjJe+KxkQdvEnbCnbNoh5s8wPXEMZlGdZcwldGNl+6JfB8lCg0uWaoZ5QG5r/qBPvx1kA4+Xvf/qtYNFB4muPlukkfa5Thk0Qk1ggGnQ9qWfjuZTz7UkzGT98srsvTR8njEUE9f4KD55HGO7qvC0FnZO/XUOq7m0j0b3SM5GRBiGq/AGKMqhX60j7h4/QYeasXPkOX2MlDjRSJbx4qfsuJLaaZQuIvtDOLp4SQ72egLEI5gvpt0eAGTF9/htlMX9dJF3cKLycJLTWSErwgejbXDVHI/xXu4PX7IMtR1QtVYy7ma9/BZoEDIjw1oHfmyBQQO884k31JS6YuPk5Ck9lr1cUi89gKLI3WOqmNGgrFMj5dJz+Y104/PF7h2WtUw0ro+jUga1DlLQJiFne34wTlF89d5p4QbnWfPI0ICu3OphaoJS5VH7MqDFSkpe3a0GnuAWo2hQ+0yytL25cS7MgUOTcjMTlVzKvmj2Iwn0Kt61x8ECVdMozEA1PiY4BsUtp0SRRFXiRyc1d9RwDXpC/EJexMZohtd1PGQm5Bv86jdFUq7dVyaXnJnbrBTm+R/tp5Yk1XPy9A8zGzzblHsUx2dY1TqEfz9x3xnWGDaBllBYTrVU0Rsb3WVC3gw/tn8TdETIkK05IgENBxCKaKthJbJFD5BGhDBPAJVKQBshJZWT8ljoX+viS10DQQ/j1IfIrSsdSpG/BhkdlwiB//p7k9+R3ioiKfpKIHsLnbfeDrOa4gTMFwcaplAivZvLfBnf2uOsghW6iGZ1siQ4fNZPpdmJxgiNy/gp38IjuecIFN6UdtPJMLBOwvVksUmwf3vATtLVaLMO0WoigwnnZtcvF3jidxAwLvM0zLD7+45KILJEC6XE2fME6BZn8nhijecw8ShwurWQCENss2CHZSMdMrELtf/KEKTPukRFZuF2v+gQw/shusuX5W2Hejz9W/pUWBezsuRJSG6cb8YyjSCJGPIucPKpTf5XmgWUjriiX68/OfAROBNl26UmPEhurzyU5QshzZWgjJ5+TsVROGDz6VdWjgQSuZOKksyH0ezJKVi8zHGKA7lQHPJODIpSVhmzVgLglauzpOVrppHAeIsXR/7sGoaFoaIBgtRE/2M4WBthgBYcLlj6QXy/TexuUi+oKJSRJvDqsmGrpM6fNzQerV0X7EnESyQzs8XgnPBSvMX/bmmbP3/5IaA5DglII539lrGtWyZ1dE4Rsrq+uMxb3bDmupxgImUH6Mh3yQ7r27HIqfECX1QEv2BQJto1fn7D4jEkCzQadqHRHQTmxSa8s/VzsMm335+mvoxGuKNgLcfwv5pCcqlu22LmgD1Qdhk+eHVqIA3XnnVUYBnFUPSkul3N1Ye3ROWjXKCdIzQCxpD9UNNGOnJRma7mhVFCPT6uogjg+6uRhyXy3lpbHD1FiovJskujapF5I/MGjUA32G72YUD5HCVxIDU1BWwbewj3KjPHYx96rMJyzUWYGhmeQlcpqXZdw5Bz6PkoNHT8AmCRNyQWr4G1xFg1GA4sRxAXQtvulkPCSq2SLwn4BWrc1rwz55/6qHTq/2yTMjyMLFJfHcIMWWWARCDv7ym+rkadtV30WhhGi48BOhpouCHxACn7WLZqjhoMcUEvQzbnnpPPEdsjsIfnRBr9jMq29ExmKQwv+D+eGRUxFGAub0MBPBevZX5ERV8bS+8sWVXmxQge8qjP5/4WVhz6l1DXLTQqlSFMcEdc4AAeZRyv026uRgX2GBlhrkAucg6k7Q+x/blOL8Yze0utIGWIFycmgyMkaxQCt2LNsX2n57GYC2hsl6/cZr/AbqC3umSEBZj28RArDrzOaoBaoH0U3W90Kwx+A8VYE/iyqzKIFWOdiT7XFUWLVVl9WZf/9UVTgUZ+NaMLagPqLtcFGhxILtxoI5UBODCahIYy3pkdZpRH+Qaep5Mf9Wn/J7jmUYH1kilfi2LbuXptObEwZb2BjifQsbguVJR5rPJ9LrK3AJdKvTUDgVw4xtRAtLrByV4Ti43s2CkgFHqw+cOPbRoeIfX6f7EDO7Ya+aOFQO9nNPJoeGE5sF7OTDY5HUI/KtdpYAokg5Y8pa/ss6jHQO07JLP0gnks50NxUms7gg5gNjYEZmLvfoVBowCb+SMwKdtmvRA6E0cYWO6RpP2fnqvL+VWiYSm7C3aQlImsslpPeMwCIqDYUsFoq6dN6ib/M3w+CiUHuIidPkVgx7l9eBADq/MuDGX/ppz9BK7IEA2OfkmFe32ng7/vY9rOYI8yY4v/lKyq+6Kbz/ZnP8GZWoakwzlpUAVkFaGareQBZkezenGYIVZjDYBygCtMJYE+ycCdpfrxV2tDY9CszYE+iRvGO0R1amMyUFhAU2BzWlP4ACG7ExRzyfl5PDl80lwBIXAUwwvXBPKeV26BoVHMwEY1+7ur0rcmjaqpTLS1FFIyUbR2cEk+J/uVXb+2Cuz6BOUVKZUlBU5wGX6tieTlgvLQXKmisgkh3v+ZLHBEW21OSU1gQ1fvqgAkVaT7pyUFXwy61T8qVKl5cXy/OxXECL9l7mNZ/OX7yUmCor3PUOPu1gNgIeHj8pUqObJBc/KXiVJfDDcyiyrn6FfGfrs34VE3kmwZFjw10SbEANaTM/7VaY77oKa8DR9RaBegCqlOklFieUOg/ZjAPkYXXv6H7DZ6P+oLbXvAMpnG6ZEvNn46+H8VTIbePCUFNZzS5CCNCkKIYOn+5MzlIZcZKZJEz+zSZ9/1witSS6G9Lk/wBo2cMQyMzkxyhcIJ/MzwG4Lboz4nzBj81pg5tW6CLd6S9+AJDhCUP7gmAPnvpxxsnB516I+ES4JhA0XjO4oWciPtJCzwqkQKeOoxvF60BnxoM+t9TWxK2nEgcMxU3IQtpnwfNRoOgm9OULFyZUGyToVZcCdOZSViuMQFwgnPRg6X7XJ1X7FJptPe0ihoXaI/PgEumD0kbsQM9bV/OwOmiax4fwXBpVCsoYAV3yFgBkQbSzktZ/XToxml8erZiCEzkQ5zW4ZwX7FiJrIjjwqd63/63Xsb5/Oi8famFFcfmPr9Ta5DAdzwHyDAJlnAAXVI+U/YTxynItVgKruHkUI9qM8n3EStuMjt1/dgE4RfvYQRUmOWjxvfjL/WSFL8KCkxryQyXJGr5Wv5Me6QtjZXPuxMw0PWAY3GmjiWxd75sbIsVueOYEl6YeQ/fxehWDfyD5Q0h4wcpvvoaqx/GvHDTBIvoeW8Id4JZYgT/p3gAK1TOLWXaiJGbDR2/XWCeSwRIbe/VParkgdnLoMoklYQl2V9owEak5LvaxUR/KzieL4EfyXr282YcLOybOhoraQZQgw75uappAMx+Rlm/fvQdvXVEW0NZlwpxkS62fF4L8G1wIjTIrAk1dMpGBx32cR3VGlGQRCHVZtpIwfWpW2VN6EZp9aCwbuqgLzzH452NT4j5UXFDV/qHcOB1mgpe+xRDOXgW2R8t4bHtMR2txOk1bwbhzrnlmydUTwJO99Di8QJsdpa2PQXKfgeqvzhxvDhoPcf4Syba6zOYtnUFzyoHPTToy8YseuRbGWdJAWmEoTHGgBQd2UgxVh777VedtaDlqoRMG86U1cbdIXJdcHSXzt9Rd2HHfezOIwdkFzJ9hoxNwSJNO+raTcoVfHpoFfs7slyGG0K1LEWXXnFrvgSFJcGQj4cNqXITrpFG1oJWcjy0lgT8+IeiIuSLWdA0RXDZwd9svTMCgpjTieTIrhH7B9/XaIeaqR+ACEdKkqets9iHKjOP+4uZR1SaK0lQKfuCZC24av4tf+tS6pR3HkETyE7IVgzc1Av/lZ0Ar8mF0P5TDLaVwb4LQEBFjG05bgCMDHeUaQzOs7pwHQ86GTVkdSFH9U3P8EzxGM7vwzGdROb5MtbMfXJnrdYKTjO0oI5GqiKtkB4clNPmwdCW/LecLoNjkLlwpsaE8PXIgOAoEm7Ey/KgvYqY5X8FQVKZeNRDt0lavN3blG410qRhFwkaAibVzKLNWMGgKWUsa8d4SPuzmLlpX6G5h/tekJIK6Hy0LiQzzB8CDUYO3s5HKbajDxtkJFs4elkiTFLSjMopA4qw62+615wZm3otzWSXwxSpetDbOxg5BnY5yq+rUFy2BEpN3Ju14wWZtJeKD/CzOC8QLMTD8qr/BXJ0t1XKcWpRtjad8L4kqtdE9ZjFHbr+I2hlE7LCVrepgcqxeJnGkh58XIWAfN1Mw49XWzG+civi7Iub/hHgJTDu7yPdo9hysoeiskdpbxPe5e9zDRkRyjECokQru12p3opVQIBAZB91j/V5dS3r/LWntfh68/rF8c4si/VyTAQJSuzXAHQx31pcveHlCmUtEY6l4aRRi0/z8G5qRNno4+1eECRhBowBeFJUG1N9JqQ2sUxq6POtuOg+yNdke+Z5BX7+6JD2u7dVgrFErj3xpKeJA4hYsvDMqtH7dws8Y0UCZHcv1gUfHsgv8XjQbQHwsvO0SJuejc2N6vma2/uuSzWpH/TZNE9Vyohl2i97HxCnMIlkIu6CpdOaiKusktv8onG5CHewyC9wVuR0B7yN22Kv/KAFRBO4dnrMxzplsfOSTVoaS1x57PCfK1B/MjJREKVRDGZG+D5DMXFf2Z55e8ZoHlhnZhSo7H18bqCFJRhQkq6KZxJm/B7lexNNtrvvK/tl10esny48/nWauFNOLHlrAlyW1rpc54CcBcerNzy6xSIuEJlF5oKYVxzdzIeRclax7qJM2ApjfeDG3indDMIdKhAfTGPCBWdO18jwqN3mJ+5PFzZ/YONPiMmAg4s2EG/lDaNvEyPoI6yF4GWGoEJqfPbl6bImL/oMu+Uaqjc3Uxm/9GRXBF8WAQbVuLfWrpeC39REolpxFvpIWlJ2dvcFpLd1EzRf/QkpDRRuj8L+0G5tVnhy5LBA0Do3kTPtvEEjxFY87K+DvXSySsXV+/7sGGBEQcDgf31stCxbKhtecGFU2QBzQu1zAUOfdDYXA9fcN0Ln09M2uIcXAeBqruuSWg3qmtHxiXvroJrXI6yWPpEu98jZzlnvgHpk4/2F9JI8TtcTV0FW0XrnTuytSBn2EqzDhuhXeE1a1bYjyJxrJTZ7dHuiQo6GlCLjSWfZb1+0U7H7FhzaEdDIkBuJXAMZbdiFFLxIlgVHcdaS6FqgFqk9+j3NmdGHnzv0uzKLqu19jj2ql++A6xu9SJVNYq1GWe7XZQ3At8SUhEEsVVvE3ZKneB6466ejYpAcreK+6Ms+GuAdC4fCswB+mv49eivUrutLCIOhK/MdJ9tjL5We32UsMmYkmPKkracO5eIToilCeH2O7AnJkg267vLk+gEdOpuR6AG4NWlgpUmbdT/BWXTHhcRamymswudVr2yNJcne4A6Hfp0HYJBwVZRPmHCgOWDA+vA2egJ2lvWOs4tQShp+xdHLWQQIv/457j/mjW2yvUjP8K62yqCBMJ5oHMMH7WF6sxb3OB2C88txiXHzuJ0LVAeEHLGgQPnGb8364dN9kQQ8/vJ/mM96eTIrOPqW3D6EknM0mPZ+nvYQERzLdn/yP6bm5VXA2vPpJfzOEmfZ86INFRzUNEGORwI6plKvJuWzP0HURhHKQBmCg7YYETyJG9t7Q0JQVPYygNx5NmLLssfJ3ec6aqkCQ9ml5MOTZYWFwvBqCJ+2rVS7j47H5sP5VYBsEFqxH5TJXEDz904Yy6QeNvrYjfKS4llUPP0ZtGO3GOHTlHJP99cbvcMvKJzesWCt2n9lAga2GkSmvqKCfbYtTFzJPZsh98o5eKZ18x4xl0c7x/8qGExS5eUwwKpsRHxZKkCAAPRyb0c0qlCIOrpNVXk5QgZMOYrW4dG4h7XfmRhBC0qgsIDQet9bJi1WOGEZLCsYo9DDPtWn8XH0qblYuyoBctom+m9Tmkuurb+5ZtTY+Hwoy+r9M1f018Jrp15Jb9db92oxvjVMAHlBSQrygpo2q1V7rJlelTuQ8OlEIRzhdX6qsX2HGVfGJHe4RoDKZ/+VZrrY4llHkrQWKMzgpCo194OzP1NLj9OCTJS7avhOMzjTO0IhJIA9dUEr10IkIaWDHJUTArjBaKU24aTpT4QxHdT1I7/cfymExQndKUEPCl3eZB775SBuvTLsgBmSb8KzXJncpHzcnvw/Gg/7Xyiabhuw146c8d421JPu30bV/E78scZO2GoVTSOQEu/67zpaYEnyNeZCtW0U5p8MJmkbF/dyojvO9E762bQU9S9riEWmo05PBxkDUh/d6CR78D3NDbJdsNxtwSxFMYbZ2/pns3ioRzp3uP4tVl1ozNK69/WJe7sKMRVBSOEhgCx80/ghcR01tmhAnrpB9eK7ya4B0+1I3nu/dUHqr5n7KSzsw2poJjA+X1x9iofBydeWNoM1bQwnwjVRhFXG+fcGi9t4VSqtvGeUfPqx51Bq+Vntkspb7huATiubrJc0L3jQXeaYUwcug6DOPuo/+f8WxiU6RG1AkRf/lhzw/OkabIx6E52KJpPbrC09G4BsXLwo2mzLkVX+MVacgfNmbMxaS7cfjbI/FZQhzKnNco7TfsYbJBx4XOn1MsDQOkOHew42xD5IU3NH2vJOVnllAbU5jrU1vaHqD997y1QsHMqaNTydeArfQp8rm/3IAz8/P99z9FOErWxqSBXYgAATLX/879ns3E7vah964SGDIoqBMqJmeBI0hb84zC+/Yrs+wxD2FIPTTv4ex2D3/P4JWDSMXRiJ/xAqDrkzeBvblN7KeZHPkYhTWxc4HwJp40ZRQVHnWmQ4ZONYXKDfERsywJrTI2mR6qcZBWLvsjaBc83t/mtdCvVg3Nhv+XHo37246PKc89B251wWLX1fk3l78Y0lLW6E8xzKyI4Y7zPMstH9C2TYl9CXfhsGub+6W7obOqcv3aggzDzKaJ5jC6HmVbWRrFwFgBAutlin9xwY9UsWotdWRbt102HOeLaUax3oyzM89tiAwknXbOwSa4I5UUC+/k3Ib3Vg6+zH8pu/aGSqWGnQZT6d3fnd3yinmpUpVANt26F6ONCmHqpzdg+aaMQgid+UWLq8vL9/bDH7IZMdH4qm8/C+nZzY4fWzKCpG6L6mvnn6noMzy88OH7ugg1dn5XlOA3QAASAiUdWr2JZLwuNHGzVRQXMlSW3cmqmEngHnCi2b5UGcEb3bRjtbrwdZae8+wVk59q01hxgOWmG8MdUsFnst33cLfn2cbA0VLsRwNA5PDO12wsUvW9JLa9ELPu/1NMvBIxTpLbLSLY6vhzkSWC39xmcSHPqkv7YrIza6XAwRkMF1KPgACiwiQxhOgRI/J9gOtc4TPuldApgPwO+jsIsqgThspjQSZ+FNvPka+I7aufn3BnSgcVMamPh7ainHY9JIJRSfD57bAYI/iw7Fiwy+Dr7F/vzAhwunsegR7b+BOLRZJg+NeUGoEXMiHw8xSXyCOQBRaRDR98oVGtyMGDF5QFig4fvcMa9rHWoWI5h4fr12nsGfzvZAj18q2uQGvtOdn2//87dP4f0DGiX8f+SZ8juNdDsrZDWe0dvPvgI0PQsZgbKZo5EmASwk8gCk4xO044+Nq+N+FYBV0I60ik1zQCCxJ5j04Nvn5Cpa2XmSoGve+PZGV9ut0Lp3HJRfTjIYWZYOY1bvZim6cfvdoOHUk2+5bfelG5kPS/UhgM4l9Er82tggr7qWbvm1EHoGXHH37vsJLK9MZijV95sooxXgO/9ErHGNdmpjRIgg9PWKNjODGhQhRr6kO/TuvAhLeRvz8KWm/2djU2ieT8/bB7OVvefYOXIxrvwe/YW7AI1F72ZxmlLgmCWHoQ8Hbym3ueS8TXIhI3Ddo1Pez/ABb0x9tj6ohNBBSgqGwiYo2zkDFvfESppvtrV/BM0aIzRoTM+uiqZLwOrfuUY3S5Up5uF+WX2FOX0gqyinhAGzibleY2sINIbK1gDxwivXxfNamjTlwsucqelncf0Z7uA4WS1nAPDaSa3lF+UPO2w8/+7WqNwP6VHvqJyDMuqQc4aINZIW/rW41WaNcq5tIilAHBjMfU9OM8NsypkUHWL8tg+PepFtZ4fqY7o9ghXKMEaMlQcRZRPLEjoPVhd0WmS7DtuDVU//A+UQjbWD2m+t0HKZ22ax8Y1krf4dGUTA89AmTivcK/GGT7czKDyjbiTXnwA7r4sdY6jFRnkIPvMZPmhIYMhLd66QZHMUBHOaAIZr2WyrwojQeosORYmNfz3jJkiZ+E0/mQYQzDWzLM4TuuStDLltSL8t+qnTc6GfZZlaPnG1O3XsuLfT7NBy+OwhtUzwJFFEwMaqLhINqVzRObmadr5rug6GNENDZTKqi0Lk7QJXP6R7iC6/0hwn4/6N/SanIAzoUrcz124ptc3SQhjPM/LbydvxJ4PuEHghu+aa1KdFzPMfAqGKJ93QR5dLC1/CkM83IRTQCdC8pN5cGAS3eNpxZASAqkcXJZZvxMXD4ONS0TrwykZzCKZAp9CctEmtLD9C0X443jnhetd6HBGQ3Ge8Lbg+YqAfHVcMPQd8lmvbBqXIO/MXUYXXOoRPLojFY/ZB8dC6o/Z7W+9fseQJA2KpC8yY1Rt/CYdy7IiL2URFKu/SSwx0IyQMPpXo0qYMTvTfMAybgyjXnot+XMJVJUITi1IToFej5EGnHc2kw0MxdoGoG5sFZ2GWJBvntgiHNunInHJIj7E7dWmiOmbi/F+nDQcLWsW126mUYDv4fPYreY4Bt47VajCEg0MSpXNVqp2qut5YCHMz3txGdI3dHRDTEpVRM29wQW14omdtT0Hn2JYgMC9NlROythypFOzButg0Czc2n1/FZWPIKT633fqMQjOZstBCEJT8KDqwvsN3YjgzpuVhLlVvpqdRqzg6P/wKcEnUUuEOTKfxWG/nSY9rfzpJTINNfuaBlyTq3R9e0CCC1Q+5NESksxjZ+dvX/S3Jw+Yz73MeJY2EsAEFIAKXR8IlRhYTMP7hDnRCid6e6FwVRUjWw6ek6oo5BSIniOSwOZr0ifspCWONW7xtigyLjTCwNgybsieM8YXdqsCJlLCy0BLyTVcTNWq/3M16y/pcaAnVHB3CoSVI8bhUBktAN8KbRb1fTuKzPlciSgYc6Kp5W5n0EtVM9sH5jg3MuXxIYAD2LQ2Zy4T75Uphe8ai0j7QayE6A7t5SeloHDxELFdA5kO7HSwjZ5VR8552u6r6GWb9tnQx6UrG/R2J3Y9RZ8A/MKV43PGv/FJzUdynAxqyZ3x4ME2YJNhu5sBHiykwgguQCv8zGXq3ALpknOo8mb13VxjXzdM3kHAcCtd9ibHebVxiNnl8fBAxzu2l77e+MOIOQyoSZWCmeOuKFzdwspgjqbqRaheSMNz6TAJSPFmHg0BlxedPM7CVIs574W+f2df72rrFtIh3Bb05c1B6hQHTaLMIXQMNSwH+Uomir/UjcEnay1igMPbnePdH1CmqDHWl+rBUgWt3u3PlNPpgCk4S6TqbSqpsrta1UV6pez2n4r2GCA+wVbJsgeT3Yp37U+WWHOthK9uSvuYK8rCeZesAzo7TNiRqZbsjiF4tcguioLED8d4DrJU+ooa3hX6K2we8JITLCCiTeFujczb2bWFrTxBhPAfY1JjiwUJjPgV31SsUHXuqOI7yMXXE+G1ut1IV6JDbIVmYHhDYGvBZIwEOZbtNvAjdgg0HCCCDiFku2g9+/A/l30Vej9tOS2rYAuP3fW7OSOJknOjpeUsCkMQeKpqohw+Oqu7D4QxIqJ60fQ/5ByLkzqfCU1Ju1c56izI1ukyn7DxFM6W4iUc3lNkyRmsunGsuUnKgJlGUPa0NtaiFs44CtRDiQZsYtrZfCaneNWbD/TIZ3KvPTN/l3Y0piRRIzyMN5kw58a22Qeo5vw/6oLNOhPfA3LoA/aKxVTZUTJceaaWVwgEU+Zz76aEHcFVhPwu6Efk1Mo9lGZAD/N9FR9vR7+WPfbSN1YYokKDmGyCJLcgv7qIaAW7UVxFnMo+tDM8vXPQyReDaqbqj/WaeoI6MGfMKc28a12uCiwv+ysDbpwXmxHDHxzUJUD1JiZPKlR8ncgD52psq54inP8RpGGtp6ZlCKjw+YWvc0TTJTWPbFuKQu4A6OMy4EZzqzk0uKrpkc3eKqSZaBARv/23L4Q1ePlW6yp0Yvl4L9XfOdpD5CV4o0mQrgl2At4DwKEaM6p/pyGKbUD7rQBBwP99akc5qBGonIEgsbW7ZazLH9FtWykep5J6Tv1phrNybb95WDDAHyz70elFc1dHqor10lw9SYpcvn/6X81jOx1zTVsdQiReCC3hRDhA++8s0ovSiLYZ7y67DkaLIqlTGoqfbmX7ML58j9gXklKLrPPi4OGhAmyiVKfwH6/8cNcPRETXRyb5fm3svI65thsyH0p/aHjn8O/AhAZYyrtbd0eQ7PO8eSYHLfGd/uMoZ1S2TdvlULgoFb0fPHDe33pBKVqH7nBYKd6zGF97rxCuICPBjLMERtsqP+fXg84rgElrp47dhSyNX/X1FRsm2Upg6F4Kqs+CdtddIzdWZGtJ9k//LmqEkaKMNA+a4V8HmKNqrk2N2E6HbuMkX80aniOf2PPcfHsBeTU/5PYNAq/rPFk+avA6VWJGYP4O2vxmrSWyxCTgA5qKGQGEjgq+S/Di6V/mSVyQvdKcfSDGcse5IwqrDYynmzD9eJx/tc0/ErmmbINdQKRL92yqq8xUW55Us+ySDV0sCFrelg0PiwNRGtIOYQUFdY1/KGK5vBd2ZRNNEwdvUZanysITCDqkEJvAx0WFeIPIGa/U+QSwVyLCALvDpXgn9EfHH/bEYr+Aqiz+qkIS5hT1N5KG/stAD85GBE1dN9+T0PaSIr2ZLOHFnbrr/w32z6RqXadkvNluVZ7SwNvPL0azTIP5+sxK7IAjAybkn2EyCviEvDqorZDEbtsF8BLFCh4Av/E5JDRASk/5vMPCCGh7cQvK4heyeTfEe6poryT9fNjxW0fz09RFLBGR5BJSLTciR2IV/EYU+TB7BY+Y+kuON+o3IkRC2844wlu2XKzFf7MnN9fXPtSOVtQtay1P7koEYAEtNqaGFpYQ8alkvTibbpfOWXArgzjwx8+lW9fbezhG62O69MTpc71wDEKE2b4Cc2CFS8MyXhBvJve0W9lQczziQWf+Vkih33ztUrG/n8khmpZpA+tRbAdAoP7RrrIqPYCPiOL0CLkDA1FjWKL5kvUVVqjtd4eeID5XSOS/MhC5A2MkW/OX+WVUNhGX5a86cdqm0CfXRxeiq+CmZJJOJ7Mc0HfEBGsRA6OkOEXDn9hbmB5p8aY/QNC1ud7AFEZ7VHbt/J8p6SugM4jzxqK1tHCVBpVZhz9akfBNJuQpd8mhbB+uVhQ9HCT5TnFputotDz3FDfUQCeA1mOJ4a4729PWdL+xOTlacCpHkhE2hRqP6zlg1ZU6HxtX3/Wkk9kF2c8bwa7AbsaS4jLUBASXptPT1U5I5oZFC31TycXR65/A72aZNeAQqTy+mOHpOzTuwkgQU8pIHcj8YXXftHmEBO5MMAUhOfBLqqOSaT0ycmXmcPunBVDlGxJrmF/KBat3uSAL+mapsSIGViuQ4/p5eEO2KY/54/HGi4L/noDerpm97Zq9VukxBd3i5HBh4ZwG2jMG1ifZQLTcw6GKGO8VIrH6WXKhYFAwWd9Enko+9X9uz6o/frxY1Xxy+ViZE7ZDoFc1ASO1EOSV+X6U5UQ3LU/eC0dlnoKFCsgxcKZZ3mzGfLxlkGH5b7Tr+r7qowXJjSkY/HDok7M7jx1DTnvaRTWve6uSpUAxuyvYvypHbCyPSQO874A1NJdNy5fcrjUutdj7TNU6FcSQBDMw8VG1/1qJu/ztJ+/U4lQAEB4XUwmeGK75fo8UCYuSaMD8syJlhVDVZlCiLUpsP476iF+VOJnDvYojg3WtVxCN95ye0olwXcreFHnYYvff7Fg98VJ3cLURnFyn7Y19daZfUu2CuJg/YljXhDpmY69xBrHNrwhT8/5pRPOfFXH7vfbH4OJesoaM8Dxj+3RwHqhmcpFrNE59konzOil8DIovvhFgB89tQW5/cATq/54x+zNnonZ8CMwv2ajyFo/mAf1qglL/j3naXRsD0pH64XApnzlkdwr3RQU2kWHzyenHFuhNg2qzV3xRitHG87/7hxtjrXmi7sm275mq89zhxf1uTwAErKoo5soZlvy1oqnewd/24oBmiQh9wyEBznznFTDlafODGnxKhPwZdiEhisBwLlYesi8b73fVknZAqsRtymeJv3boF+c+7Z6tEyJ22aL+Fg9/X8LUK7cPi5553yfF7QWblQ7wZqLc1DDeLrl7LXVH6wWJR33NQIwMGEWmdsw1CF6V3p8QGZc1oeLZ1XDDHH/orxQ2p3wjlz/A82iNfZKS2TMM+TXEW0BS8OGlEzA5StmATlvvJ0HBcPwxywaihfyBg0PDbY1ysDmCLo6nytvkyVQ35H55TIfD77sh1PmXQPDJuBRMMqdRVOq1EBMXYoYAeaZ/Qi0LEm5A1fO6Ec3F8WujAgllrf5Q1mggAgBeSZSf6mJ2LeG+FelNhptqQxiEAO2w199/CjSGq4zHOWt7vwaUSyXrg2i03EC72uX41U690QZGakHUIUaFeE8C0aq8d4RqKqfElKCOQgE06tMeRiLzA7m8zDCeC3z68mI1oM7PvmnCuGpI9zXKUs6pumhG7/huaZRu4lJxQ17JFdK93u87+d2rv/5j8H+vyq/+FFOwUgtQn+bi2eSKQZNmSYkLtR2I+YthcGgezO9QAiaKgHKS1CNZnPVqhsr9jJkgdBR9CeGCjNS9+rY4cXH7MD3k2vlPRunZY+4nA4/VKv0QCtexPF8psp7wc+fyhoMCu7sdJOVLchuEguEU9ZqpUpOD465hwqM1rt02IZ8hEDqSr0KwU8YwNB8/iiCD9Mo8nCf7ZxLJPgmPEgiv1LhLnJRrG7Rku1sJELQarlqV1Hq6HSNpcKTI7g6ybYIkz3ViMKgbkqH1uRoff/tp974gxGWfUmIwxPop/Jzo979qr/eksQCXAc2Apk9nzPZy4nLTFRDhcBRw4PGmtovIvMCpB+DFqyNcKFE675kwwtlQXPgmbRCX6pn3GSv/ZoREM6dLdD5LEE2tx2zLxTLuwJffKUhTys6ubjoieyWnwAMq0oYpwIiGtWGFKktbM64dEayIqU4F5ydJmkZcM1YIDn3UZBifVAWHbxaOlB841HCMaw8UQfMWLH/Ng/DE5F65wjMbNkwt+Ats8rSn5I0Z+VWP4jysitx+ZEI2cZK5FWy+KWf7Fg5vMeqCeL+1GytfH+VtJD13ydNFLIY9yesqOInNFxR2t0+pgaj5ipdeY4DOvZpXnx+FrcPyjiL/G4e67e4Si/hJ3jxTr9ICTpfm6Z2OE/LKqL5eVrbUk1TsZ/NQyKt766Rcc32XeVpc1hNDaG77YDp5E9xpa1h6SvVuyVmx2uomn7FTMDM7u9XFNxJVZ3H1FQ9ZLd1x9BesUGPp7UdVtwqb1KJR3TuVjsZrz6opx3wfD9k3YtlZQsio2U61rwFZCg4g8F5aCu4GveFTD2/fQYYNvNftbcXIfdUUEKx+A1lq1npygQxD57GzXoUWAqrJe20b48tbDjBwxPpYTImMTHQv3pRraR7LYHellmu7tNFKMDL7/F08PfQak7605nAgx42B9TXi6x5kC3a1FpvINw1AHnNH1lX06d+upHIjkXNEyAAvYgZ7+tE5JhMlDrU55pQ20aidnPJYVo1vbZQGNNuclaZMfgq0Je6bVYaHm0Nl62uOP1Bg0AuqTrvpskox/zvetoL25bvcPKIzIwx1WM/YkQbhE86MkdF9VDhlSjjZNaRicHnwvrS/mB+c4tkfzODALE21EfPSNdKHxwK5PPc8vc9nhBNgEGl3diIuUElZZb58zKmmuE1KebOCTDaVhLujp1A/Sd8JLAKqN/7Ogwq4kbUzC4KYGcjR9tbgrrkUKetN/CoYAi6obOj2sFVC+17i+XZhGyf3/Eyahg/MsTfHPq8+YHyLb+Wm+tlKbiLDYVmwCQ3U4ZrlSaF3mnTUZreU99H520+Itmwtm+WaTQL6dlpXt3R1yrHq1+gghyZKxWA9bUtoySEck7ZcvFrcKCP0AyNu1WKo8g6m0ensw8/cH93vMg5fJ+/4o4cNV5hVUnaqSW0SaRBK+CXhIeks4Cf24LTOhWTSHRRmBCpH3sL/7huIehl0pq+VR+4ZWG1d3ibVEiculF4Jw1HpctOIEGV3XFgGq2fJPEgLfNjjWqSUHpOjozwsC48lAQa2pGhlx/slG+m2lM77cMqT7F3y3SopA1FPQ3y9PeJzvJt5Elo4Xkt46jmtcsT08S3xoGox31vj5LoPrqd09yOexKUBC42ZaAn0m5BW+Ji24W7YtZIjj9QXSPuTt2K0sgwadwSDmqjOuAL1Ys8Wyvcftm1OXGH/JB8zRfKfFsznj3zHKUcFPMyfZYWMJ+brszPswGnxFOoJXuLqjEcpModmTPOp46CN6vcLLBb4BOu6rubjA/cbwbL2jRWhi2VJFN7ldhUdPaqBfYChlL41E90i71in83ZZMIAALGXCfLFM/SwIpxc7YF+pjiY4BniuYf5EKqg8khg+lTsxXCQ4VPcG2hGwE1o4zzM5e3qfwXfZe3BlCzy0lQ5ou6Fzj0hfjPGHKASFAQ9py/MJIhGvDyAngpUB7KyNLlGyGr8q3MdjOWR/lBvA/e1MIwyshxF7ejrdHoFiNbeKUkxvogQ7cZs8ETvAeALu5P7LMbinRtguxFX4ckiJKqiIgrqi23YV+VWAcyAJ1IUXpXIn1cTY+naA8MMFN/N140mDbNJY+WP01Ej+GfragsFrKjv+3/prPW8tg30uhnq5eHVY/R0YFFxxcK1YWHj6nSQsXG1bhTkK+bgn0J89/jS7CS9652OBSfxYEV03/mxpXE/aRJReVsFqx+mJz8hWWgrbXx6J7ztT2i24J35tLMPb9XvKTdLoofoV1HLlt4xnBKSDrpiEj+j4lVZVVqj/cpIpLh9FcuZJ+3vB8db4Xy0tmaql+GfO1xEHyZErYuV3wIvZ8rUnxXOhmvYXLRpCUU4rbKBivvpFQinhgc4GOpiWhQ3SoITslcaZ1Jgm1r535c1OEh08RSlM5qW/9dUjIG3TLVyWC60asa4nIs23tPe3rdKT5MirQiKZIppmaximM0RSgiEAVDPUPtXK/v/H/VEv568T93/Z0MD/gNVhhJusHutiEvc4M7Y4/m1soOXLGn0LXNlKjjIK7zUcm4XF/KNZNdyz9e8IbDAJYrhM6b2kdZeS0aQdReA6VFvtGsy+B7vBWkCS42sQB+VdIPw+pEeECOp5dRoOxJz3pNYapmo9iRXEc61W/dV+6lcu2bsKNWgnaSprR1XGAkuGYjh4KWvG6MsSsRWADjnbw1mIetN6nP27nLR9cWHaMSfcZcH/nizilMgCdheoLob1GvZlealMXHySpiUuK2/EHs/IH7UCARoySOAGj45Sj7kmOsPQoWnjv0vbWwXmq/O1lTAuaYAcrP2p9IdKiDclDWv9SUiLBDeueGeOKGQsr/xqzFiqioAxEbAQOYjsmD1LdyvMsdxdTbT4wpM/pwpggQCIv9rciUvNWjlRI4lGBbVYgY+YwOASpjhvyT/mIhejph/DxJafUW/Y0vNmK5IPhRF6Gc9tXRthghc49y6PGKER7sv6UAQKYX2AuF1X18wh7tL+/et4Jan5f9nZbDSO/MEJkyXUFV0n/7G6FDBT+UxCer9u7sXgeAPcl9BnF/Y85Swzi+M7oOJttfbf+qIWYHL6aAtyWu+G4ncSaPFNZ/k6OlcPzqJJuw+Dvm9/RH9qqKgMz0Ga5n5r/fWdSbKSyXDMNLEGLgueBWwyyTdduvoZ3WDPpvRKP2wH02PBJfb8S/TjO7PsXCq0pwYzj79S/4wOhAHkv9SM305KOV2RWHpx7XezKLWQsZ9RurYJ4+JEACSnYFN29GC/odv3xp/YhPbfRgZSdNSkxJ1m5Qz/NX+UEq/fQ+EWlhtzuEy1EC5ROoAXw/2d/h+vcLKvIvsH7vHKwwNNXFIBv33SOk0THcxbepHDI5kjA5J632pEfwvgFay/iBOwumSiL0EQGKNkw/AsjiYWvPs3EVkFutz6QZ/NkqoOO3owasn5jjxU6QyM9ox9TzUhwReD8IaSO1Hu2M4zwpCR07ucMNw9WR1/BTmrg3acdjzdRIhqH7Cbo6k8if4KAGWXOwsvxRq0v5hwUaAuTBuwqoo6qWrt7O1iLA26PbWoNY6faC0hVgMsKqvzryYt+NFReJt4dN/sa+R+bCAWWDRTBKMCTQlyAi4/ZlbiNCpCVD5TvIM6qrHb+559qxhyhLnQ1j0J3PxVrNBN0UQpjdrvWwJiAGMPmY569Kg9ZtYUohdD0qoqLX3aIdWnjoU+fH+k03fGrrxssYJ6cqvuhzsfk1UcD3NWhSsW4YgqrEc32dBJY2JKd7GjR7hquDO1+bPgTND5HkZk0OhzIC5wW2cFl226PZ+VM97LqUwx/VLtySlRhUXszz9u2wfjXDDgbVluJ8FPKaWTjXp8N2/+Kh5CC9wpAiqq/6Gq1KJa8xDU3Ee7/HvBlGrPUB8IGjQ7EE30HdzXMBDWwARr2Srmm0kNF3vvRfIjs5k8tGdu2pQLPVMdfywjRONld/Nqv54SNq2oIaXiEilYVIQggEsr87+nWCqAXHzgtOrOCWF8Lc5jFpquZsZJdcSQ8MJt4HmI+Nt2f3KZSjHp742lmaQkBTvM/plzdO/HeN0Kjb+mKQoqahMnspHVyaav46Bbtph8Hhd2H9XOtwQUoJv1VNRLTcJ1B1ps1+V+tMRY5atC4rkSPnTagcKfgHnVs6bbmvMDti8EepZ6mC2+nMy9L0tTE758s6k7rSub8CsYMG0mrVDz+MVFyD0/4VqFuTatJ5uvUWwr+pilkDeZgeLYOdZ2n9yi8FjB2DGs/D0fKs/XN26iMxnBrGYs6oItBdV8UDoubDg2Hre6bmE3HNtWl73gUuU1aOJSrUbXTYqeZnvdE7/0oTQDMZOxRQUs4ZXVLY1vOXvY3mUopeppuvo15krTTqbcMIX6s+sFzzXOgYGlk0IFtnZj4PRovbga9zTyVvuHU4CvTvQxn1haDRu4AO7naG0Mj+x+xfBCgohzvhj/5YXw2CUI4zdpSwlV0NHTEzK3APQAHH7rqqVgsjr1UO8AG4brauIGjNBfM4OgENDf79jN6z7ddiQDFtsFenE8iAVTKN50g81pah0Jr5cyrYxMGcIDapDzBnwKXapFeEc72woQlX4K45vx++h+3pwpthcrgq907BCx889pTKXQc4OYLekQG119k5uqcI6RirsOONPH59cljEgVT3h12VKdM3VT6z24/rIJ+Cf47Fo6xZgBDtvjkqK5jKZnUiHiOkA09YXdMmRALLJKu7R5VXyrk86qcUQonLZAKdH1VzrzTQEdh0BpWzkD9JYJxdrHsyZbdDR8RS3bZIZRBatqtAK/036DIdsg/SjQM8Wu0NSjWWI+Yp/fdrORrId8/CahdQaKuGo0CwFwM5yWRMxXbWztklOkKfACc6NKRLNzbZ86D34wUVHNgY0XCTyZmQJhLo1bdzJD3uu3Nu5wPtvXpQzT8yLAyXPSROzSP/z+X26KGJ87Pd6TX77XbpUaIZRJQV6QE3Q1iRoqopYbnDRndmTFvOcy8WcTZXYv1FiKB3u7b8sNXD7t6SubCwqvVqCfCuEaiaz/12bAKhUJtHyZDZVfjDG6Qr2cfwXViZ4+sHigQapbaAd0sPPr7Rc1bKLlBmtrQWrYvaPfIVb+yuF/ISgv37dG9hNHlpPDSgSJChPH5VMQ83jXF4r3Cq8mkbhddUW9JS70DGRZUWrdVXrxo2o/qh+sab4Re1LaSxs3ysKKH+wWADpyrl3Etxfd9RwCFnXTt4ymwrCMqF9z87enkqGwu/Wwjjt0EO46D8sAmf+NDLaNyPzSMNmr+fx0uS19uy5QRfYd2B/9R5zyt/soTIpuVQCXn5+/k9ajYlxxvsdAPkH3XIbF37KkBjGmt4jfZgMN2Iq9O7J1jsJ02PMHJHJ11GaHQWbw5D3OkccZJhiprzJ6tdYTcFAb8eBASZyFnRnwhUwObHQ/2graFesfq9QQdxbtV7SSWqFSCvrM8shSQABwRHB2oa+OP2C+Hi9/OzHAO1ryuL8tzXTUoi1U9Jh5MBmC7/m+l4Yi2SCd8+7HgGZ4nsICcGjHQlk0Y3RbLG7SV9l4mOTtkP8IpFfzi4IsMjvyjhuFksRMijpIRyi8ZxeBcDK/w6gmsCo/7bBPC62Aaimsl5+Ad4oKXEljD1rspOvv0Q9eIfPrrXLWJS/t8upgJidvvBjVFfa+rIndHDQDkA6TMX0fPZ1e/S0RUoU9kcq9x76Sf4F9nDQoamLL2Mu0Pl8u1TN5oRWa5QFqkX4BzHCAgu+IxfBoA2ycZI6pstHSTuKGrxHOf1H0UUORkU4lPmg+huT7NGPfo6jYBd66qa6ADKXylfwLZa/w8O3vz4WKHJmarKpdOuu2/zKnY6Vdp3nHQvTpFfRlZcL19KpRsOGKgwghMAAD7U5/b2AY0dmlphSupVJsM1jqlAuz+k75vgcibIFj6pIRCYUTepsMq/KaqTfnCle+eDThsQ7T3iXz7Mvp1KX7FJBULiHI/qQ4g5bBa91HOPPqBR7MlurwEDaCxl8y9xMTeYijATZhGcHjgpXhm9Js8g/fVtCczGTYy3ls8r1JTKQ8yvIFr+Y8sUEcqug3NoWVw72L/TPWhII0Qb0lLM4C1OueqgAALT244CJXLKVpOQO9TT/tYjjKaQpNV7o4D41xE7tP1c95Oa2I+5JHklnuRFTJ36iFxsHhV8cvUSuG6pK/xig4i1tBrKxd5jIneaYFbPB8P/kGU6DUWQP8WAZf3Jych6NlGBNlfNlKOXHmNCUCX75GQ6XSi1NPnTZZ+6iAGXfWaeDK3H3C6rU0iQKQFznDP3b8izC6d2XbmAzplIyskyFR7a0SVGnpFCW8+W5IwMu9NHIod56bdrK+WzxsYn5FDDziDlw4qRq7XRZf2OaVULPKCKEsNyDf6RrvE9EZCbsxfCZ+jopYGuvyjZ5EdbI/Q4EhWlkVDuomQ/8JK0b2oNyfTt/urjHb6d/e9qUkq022zhjCONNeMcbf26LrTbWfVqU87k9HHxiH7A2ALg3FF2lXpYgURKuQZLJ59cD+zQlOOqcTIQKZZOnSIm5y/qPImJe26x0Ns422hRidFkeqls7Cm0w/plO6YvPcCdAUEYDYc4odwVI+KUtTASrOWx5prAEVFPHgzaOxiVULteEOBvL3Lnt35W7nF9+hks5KpDfHkgmLgptHep3TekRqiKQf6lt9R2D9mdq9Fqn1n+KS6VE+kivdxEARC/pKC1JZ9THkwKGsw1e46O4iGQxtr0EjEWmvzdsHbq1otZ4Mxn67FkmmMh4zG42sBaTYbOBE9BqliPQm+O1Zmy9KYnLK6psDHs18mf9UAyLSwYUp3vYjbIjAi5oBeKKKDFdWqiWWHV9W7nJXUG2hM3C3K9q2ryJcPx05fcWvrRpWJPHn0xhq7VFGhQORqaQAUrRqrgB4oUKP94OCt0MpEcngwA6jFGR4/o93ScwsX57QETEBg7O5BmFdmxd/MfAjdsM0C7mR0I3qdOHAkDwkCkffGcVqune5d1gQPqujGC8J+fWZmRATOp/pM+PNt43sUsFFl0VTMri0az4v5awsMxW6ke83vKGPTE6qWRcueW5ED7GtkRNiSVsBMVZmAhCvDvrB84tRHxqMLGeTkWGdiUZZ3++4/xx79tstzDPt6p7VHNJGO9MaGcd924kpyefyy6nveO3QOpJRJ3+kCVsLiJNTbJBPmKAp8DC47Qk/9wBI01s8PVbKwdQlOCFdTC26f6RrK3bkub8VSpE8Cfaer9Dht77wSLVoDirmXNe8LUPcV0CYTpc71wDRWUFlhwxF5y+VSThvu4S3r6u9JSGij/neWcL9LdZjBZ2TnmvNc/1yKt6SAXKPQz8Om8mJoTLM29fdde8IW7AU/Ed3vfmWT1zJBpQHZhFsZ/ku2M6wrXq88MzUlICfgj3meXmo/1mmv4XB9mI5EAYVz6fO2eDQhpgKmVjcdZUSPvdO+CBNDBGXUuQ2RiGP+Ebzi7oWng1SZ/9VT8x0w3iMA7eKncyTPBIvO+tUQW9hJ+S63qj7ryLVQBqVGXcYTn+gTDSBToON5GVXZxoB3M6VcofAD7kZXiLZmRyOkGTstN6uVJbQosxfdwqcrW7hmAlKH+KQOeYocdGWGFCWXIGl1DwDhZjjVwUIdR9lYWTm+W30g6ox1fAqA1gS6VGvquRzsvX3c4fCuQ+HyUMiQgLk/BWZ3+/MEi2jrTdnfvUwaOJzgHvmV++hjMw6cmnxCNaa6FrztRjC8/dvS0H7zHDrSiIP3zy7bMK6DEl1WD3YBuTHxWwoaTVBCFrSQxdno+dJX9ZQ4+KmX6kFsCD+Gch4tNJj70jHiTdB+os515kis6OljqC9CcnIPFiYBJ2o8fK9BGfc1fh3BXexbRHIL2TFhS08DJ334c/EHI3FzMWFzTlj5w7LEUJ+m9mmgULa5nKCCVRgXzQK/uHFdmOrjd1tVg8xeR/a9tfKlMK5euSI/FC0JFrrFq9ymE6ko3Fm/8WZMZMNS/MI/AXPUgV4x74TPjE7w54Ee/E1vnQsrhWR82SjtTLbYKMm+F8h30SYf12bwTuOVedeEEaCFPhyWQskAid7eIFF0xgbEhfY+Pfib4+sh3xke8Zg67rZA02+DMkR6Yw60tzm0PMRF39/KHVcNaEwfEHP8tWu9QUgHjUhsFnbVltG5/rB+1gernyAQW18hjQ539+iDM8YOZnm8VuukSrOG17jtFQlUN5q7bWW48eK4eGh2vqteS2h8ZCghuKmWl914YmKpAVYcsiXBVSr3v4jf3EaXpSPojXCyI6Ivwxzhguhn6+K5MxIguChaURL4QAHUZOp0rARhdWQLM18xoxVM7pWwXmmkMytSVpbX3X0iNb5pohDayhxO9KaO0wz/DLDI4Fr91do0fJmgu0oa5P0c35enZeJNeEmaIgBS+h8UJyJTVt9dkgRUdpoqq3XtE49f26mqG1QlXERifoNXCCd3gdLrWCPyN27S2DuTwlJLlg4ABB4UtHiLHoacVfPqUYW2iAweftgxRG8K3gRdEaBr1ZFk+NYjMBj+Km7xszriSAjNVPA2x9/mrgGi8XwHfMoBxTivwK/kTaQ0d8lRYILsEU1pWPxHN5lYD23w5cjQoMkEg3aXE8XAjt7KpcHSjbZLK6Em1HJ/eO1Ps9UpXWkLEEHDmszQp11R5C/sxKgiP4iUOfXB3RP2J1Dcge0kJyoNfIf3W510S9+8hGQSlGlgtlOIDBm96GxAfavwabijSmB/drKCjZqUqUopzYu6FXy9Sg7ktNgnzeg2r1hUNJfhZ1MzNDfFtHpkoTyCuGUepYJv6Wtk1e0Tb+vhQ8Ba1p64v+NjR/dXI66t6XXXGh+TLW9qg/eQmnNwjv02GutrXUQPqgrDc939EXXy9uGrABAOf6WY5EwhAvBWxVmRr5A4eztJ276bC/MN9Cf/rbE7Hvtp3lxi4/rphS36itj4fa7+GiyS4bz6Gsf1qtOXuzLBqWemoWDGGIzll9i4zWkJXVAMzRTHM73QrjgTE3YITG5KcKiCpi++5A4uCA6JjyRpGnxASunO7awU5AMad9l4+lF6mQ68IdRVj8GUzDsbcDizIxB5X+NRpGVPpAHFLcmA6yVhSgBfMGOH/FTuWumPkdHw+tbeV6DJYLH0Af/iFK3hzRA4ko3L4CVbY04XUwzcVBLsPcWBfTYMEOqEaS7V/0fs+mconGjSz3dpVoFjtj3wsV+lC3gQ7s84xa3CGWuG6hoHdngq78yXjJjUetG+P+/RNtRW85HKkJRlmj/Y4Z/2V9pWruYdovDk99NnLfLmdLL8yTL+6+HGIDw/Haj2bAS4Vq9J0iK9bnluMa7dJwSH3qe64RiUv8MhgAhwTcFKN4+4KiKvC/Y1qW6ujez8wP5K/sqBED5etH+OzXw4gEhIF1CHBj9X14/aCKrak7pieChnA5qS7N38lQEN5t5ivISMJDldv/Mzu4XioPybkwjYSptdDenEEYA30SzjL7VWLSCGj0vfLLCwA39Ft5PNp7rckS16dxs8UpPRM+FEEZkfU32C+iDv2LJWc5YD1QuSkUiA7H45JAHhFdhmJQvjNmLJeBWVg8/+dZpmmzSGdCYQs90qL4IBPi5kwPJfv8begnGeGQH/5aRgtJ69wBAhr4aSZ5qJ7YqCQWIXlkLaxQbkz40VjkwJfz+rqJmR9K268bd/diZFheEMiHxAuneP/9KnPQ/LcDXwsCcS57Z6gyKr7CLsf+TV9iq1Cn4AZ4Z5j7bpmnmRWQ4zeGVTz8cKOb6ZHB3+NF5cxp7KlGIM/vvVP96BctavmcJH4OORupwL1eT5xvnPm3HxG9KDp1tK4ueXzWfViM2qK+peYdsJpCSL/OqlxMhPIBlrctuT5FSNnZRi89IU4/HA3LXRg4EUPZSvysAJ7EsAOlH4iREeGiBXUDn2fVYDePLGIMSUx2c6yuIz8eIGKximkFF3hDKinDvhjzpu5SgXwb1MHqiNji0lFDkmB2OK4WaDu4PNEM8rzZJC3LtePRa1FrFCY4DOtHZhQY24b32MPMDb/vWck4niwjWo4M3sl3eLO0RV3mtKd4/9eb1QMk2DJ/MBI1WD+uHQty/6wmXNHqfwZkLrFmhFTJh3lMFUxgSMoRWHx7Y8g977tQlxqCzOi1CGvilaJE1o+e+Xthb7roS0J2GtVw0yDGbw2igSuBL4Sv/onfNWksg4ZSplfNuGlorapQ/GOsj81ckhMcBfeABNT8n7/AGj8zJgen4Plx6tzgIJqUbIna4Wp8Hrk11IY5KhUA/Suw29u71+2tpWIrdTvY33B1DGwdudvzRzcAFTj4RPz9s6NYyrjMwAGXflKeM/HrLQNtmTLd9Ic1xNyI86UARBr11vjE2lwwHpEZhM4u2aEWYLk8YA6XlR2BZ8cRm0e6g5vliD5ODgXc/bOW9YZnUBBnqlxfmdPaNRxm0A12av2gxPJdqxWMIu4/e3UUmMeaiOyEjc1xB7mE9H2IK0M/1pmd1aRlW6jieFKCqc4PlIzk35D2Lp4r88PBV4kz851SmI1SSmO8MtmWMk4VjcAR8vm1TWS99UXMC9DYce8Uw9Bh7MRJPCrp4Eud78vgkVrHzn9lyWTWbiWOCKqBYvge4BOAuyCC3F5nMXyEG4Of+ld+SXHHePZP/M5Z28yVoqZdTqz5tk8shuWwPXV3sBGWc0ESARo6U4a9RBCUdU86iW3vuGqrTmz+Pjp+HLEYCyugYP8usHX25QnYTnz5zUNydnpBC1uJ0u5XtyR6XJEYxMY+Q4y3gUxLMknSWqIZDGh6iSE4DWyNrmxLjVKo+BPn02yhz+yzICe3Sddq5fR2V1Kk4fiaC458iWUKdefCaNW3J1oheTRJoGjFOzhzYUB7GhydEvVns4TBD5c0fSpbxlb1+aqx0T0/9VUX9gIP5LVDjBXMWc3x6pUcxOEY+1kZKLQvDp3COOiaAEC7QbLfU2DKwb1Ht+OhYxzQTimvqwUSgxK+KUz8nlrq4BT//X1hLUWf7htlHLLU581C5B+w5oEtIY0HVHaDGNHXn1g5QSFmcerFn4wSgAE2wa6M9nb62ZcOhoW1VQSPDzyCntFkK26JFXn5eoG9fUtIayMs9v4PeDiqBSb/TM4JsVmovWJFeTB8s321J8D/B+f0ap33vwBVElg7th80mMTiokTLtMEJh6q/6IKfgk+uy+RigVeB9+a8NM/wojxss4EibvpVEYliDQldy1pzdBMLu1RPUy0K+K2xat3FpSTf7eJf8/In4i8hiitqfUquIKwuQwFIx+zdjW3xCibU5befncT1+h9DgAZvE9Lro4grpgF5esu57uypqKG3kPfSAwGx82iAlrGfAjZHxBdfThU/Asqdr310nEXd84JRy+GD+rfXrhU40lJBJb27lbltlZ6Gkg7IYifo3aKiYDHQEcs0ZeLjoTG9LfAWdpHyWLWT3gEIQsGDmWUxLiTieWqkGC5ZlCy4Z5WwCpMsJfbK14vQYjdJ3r0lZwwKD0zXo26nTYXCaExA04gu7NpkdxM9ynPKutKW2TuLlxMuQ4dW63pMZdl1O9Jx+y8vZ52L9TkMDmYwWE8VA6wokKzBXPSUkRDHzIWdopZaFC36YYO/jY92MrdjNmXJSYn8XkNs01l4TZ4snlESS5Y/bkKb5KUsVihxtfpEoBrjgcVate21r+P5EXz1o5W9+4gun7RRhPLrPwWg7xBFY7wjqY1cIVheOOgZN4o4GOWfYg/tax15/FNwxGPy1KuaCFl40J/a0cjuNEsbbYy6p0L+3VTXdwvIzZ2QCXTJYqNUzb5k4+VbqUmEh4MOddqYlYTnnepqVr2ENkU2DVbjZKM9og+hBfIkuh+OldSk+iwX7v2PjmZLLECl4xxvxXUbc3nadERbKx+nvPjF5Ot+1LvXnvU9C3IkHOJDaDakkSG4MEPMW5R+2EJJEC0NsAsx3ch9U3jRCd8s3xfOKQEXa5oOk+MHiWFLWHlhsgFBZDmC0I2Il1HpbrUJ/zveWhOSLtL3SMoLIEYFwoNQ8/nucgMF5NyU9atZclwmf5KUTHsLFbcac2QcDjNNslypK4NJdaindnzhOTRpBqQ/zczc1Ta2wNX0yakLepWeUu2OO/HhsAvfgsQ7ZfAYI+06lpDoG97rbQutcbg6F530UXIu+v3KPOSlSmoOTv5HnYLVXMTf0ng9joebZcb1i1ElWsPhRoqDPmfI1xWVVf0E5SdgoKDua+i9K0tGFG5ft2b9T8rAszsQwfqRc618czvUJQAedFfOUmZZy2S2JuYfsy63Xyr0J9dJdofsKrQaTMCyP5kcpRek9wfO+Nf7W5CdCGYrrXcHPMlOzyiiHTe6klAwmSpQ/qxGUOaVHE8QKL+wsNEQMXJLOr+ow5yRohcpir4l9QqZnG28JS7gdvIIGHX/lQTocT2nGnIgSFgG9HPUZ2OrzPKl9gV//Komd2MvBI7JbxyhxpLXSxdau+iseEAC9U0u1qvo0DKDTLM4Aats3jg/RTLLLuz0doVYV+Fr2ss3ihd419WM+Uftr6weaS9i/Jf4O2usIJKAUcV/EJK3862RMHpiYL7jHwkAnsfNunaWrMblO0ezYwzYhBKRL7KhyC+hBysaNB+HKk5rb5Fh6xtc9wKBolF87uLa8i7GJRtZMJA3H/EolRs8Ab3mVD6SHYp2esVaZ+NarDrBwGWmIUcBZJbQMEmFUhuKOYI9ztJvSOusuAobUKu38eNh4ia71dkWldq1Omf7Dc1UFotE11OnQwgKh5ztX5Nw1RwSPUgyw5+NmJROEB+MCUFWj6c1DUE6RhoHjIAYYlHuKLY1Ipr92OUGHtrJ7KQx4J2tkehEyJU+dB1heDEq4O7rlhN3pM6pm0fJ4exsU+I31gGNmjm29qKAprevgdHBVvwOMScOJkcnTjpOAcm3Sl6qQxXI7+4XBmfMt3j5CrVDxZOH67t806KRY1Z6OB/riUUYfR9KR+zQ7QQGfRGngrzqZwQwbExjXQS3Cr95XBu9ToqZENU2XPsfsi5QWSe7GUvSi6NaoKy7IH0MjQm52Lpm9j4P3qH+xe2sKjf6A+TppQ0ryJ+8EFLk3MSgkFUdgC6v7NzO5boQbrgXW9NmbpPY1/35D/dsJB3rZDDwYTxgyNtZAQSxKuC3RoYjzpMUoFsYqaoI8JQacc5r8aDZwNlgIuO3PwrU6/NaN58/qFCiBgLxxja5OqGH/+oIU+8wGB9Jw1XqJrT/qCp5LY21Hs1KYk+e72EWC3/Cu8OEGnUwx/v84FnIwCEVFWQz2OG9k8IqPtQqaXy8SRHsO4Q3x9wuVPGeWYYy73VCf1+C4oO1n4Pe3YNyA8VbbZ/wyXNERkllLKhFlUxW/bWIVEJz6bIyIVfc+jgUO0ze6dN/aCLB+0H+A5JbKAvmpD7YxwlNH2GnfY74I9Yl2mlCo/xOLQqb4zSII83dqsPl8RWqQJqxCaa8srEqYa7nDWur357oMcG0BjnGejoyqJBN4jnwuXIgE1KPsGKsATE6FRd7foMVWYINOFofDIg7WHwy0A8RtGV4DIaS7a5Swk6XSXwDtKUCrFUHtcRwXzGF5gQBHPCD0V2kcc4K2QHbW+cj+yD2/YaMxDF9zlxbDqAzYggoVsQQQ5PPSYsBoHeVZfMFl5jCFm70DfNT2Ob3mlaUA5nSznMNfhDwdksiXSwt0jsP9P9KdSM5LqOoKRnCrDL8GqivktDcqe8g5i73Rkd6quFh854zHhjpPIJu0lCsYwf6X+C5BYmqzbQb07EkaeJKJ2mv9G9GlpKNwCoiUgktcGXu+HsyLZ43i81k0vO4DoLVtwU42D3CLKg21O2uzs8gqMEmJnzycrFi6nVN1TD2DBhvrSq0ta4AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA==';

function buildOtsApplicationFormHTML(row, d){
  const loanAmt = Number(row[C.SANCT_LIM]) || 0;
  const loanDate = d.acctOpenDate ? otsAppFormatDMonY(d.acctOpenDate) : '';
  const outstanding = d.osatots;
  const otsAmt = d.otsamt;
  const tokenAmt = d.tokenamt;
  const restDateStr = d.restdate ? otsAppFormatDMonY(d.restdate) : '';
  const dateStr = d.tokendate ? otsAppFormatDMonY(d.tokendate) : otsAppFormatDMonY(new Date());
  const purpose = d.purpose || '.....................';
  const district = d.district || '';

  return '<div style="font-family:\'Noto Sans Devanagari\',\'Mangal\',Arial,sans-serif; width:800px; padding:40px 44px; color:#111; font-size:15px; line-height:1.7; background:#fff;">'
    + '<div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:10px;">'
    +   '<div style="font-size:13px;">'
    +     '<div style="text-decoration:underline; font-weight:bold;">ऋणी से समझौता हेतु आवेदन-पत्र</div>'
    +     '<div style="margin-top:10px;">शाखा प्रबन्धक</div>'
    +     '<div>उत्तर प्रदेश ग्रामीण बैंक</div>'
    +     '<div>शाखा - ' + esc(row[C.SOL_DESC]||'') + ' (' + esc(row[C.SOL_ID]||'') + ')</div>'
    +     '<div>जिला - ' + esc(district) + '</div>'
    +   '</div>'
    +   '<img src="' + OTS_APP_LOGO_DATA_URI + '" alt="Uttar Pradesh Gramin Bank" style="height:100px;width:auto;display:block;">'
    + '</div>'
    + '<div style="margin:16px 0 8px 0;">महोदय,</div>'
    + '<div>विषय: समझौता प्रस्ताव श्री/श्रीमती/म॰ <b>' + esc(row[C.NAME]||'') + '</b> खाता संख्या <b>' + esc(row[C.ACCT_NO]||'') + '</b></div>'
    + '<div style="margin-top:10px;">उक्त सन्दर्भ में आपको अवगत कराना है कि आप द्वारा ऋण खाते के निपटान हेतु दी जाने वाली सुविधा का हम लाभ प्राप्त करना चाहते हैं। अतः आपसे सन्दर्भित ऋण खाते के निपटान हेतु निम्न अनुरोध है:</div>'
    + '<div style="margin-top:12px;">'
    +   otsAppPoint(1, 'यह कि हमने / उपरोक्त ऋणी ने आपकी शाखा से ' + esc(purpose) + ' उद्देश्य हेतु रु. ' + loanAmt.toFixed(2) + '/- का बैंक ऋण दिनांक ' + loanDate + ' को लिया था। वर्तमान में खाते में रु. ' + outstanding.toFixed(2) + '/- एवं ब्याज इत्यादि अवशेष है।')
    +   otsAppPoint(2, 'इस आवेदन के साथ हमारे द्वारा अपने (ऋणियों के) पैन कार्ड एवं आधार कार्ड की प्रति संलग्न की जा रही है (Not applicable for Deceased)।')
    +   otsAppPoint(3, 'यह कि उक्त बकाया/ऋणों के निपटान हेतु हम रु. ' + otsAmt.toFixed(2) + '/- (' + otsAppHindiRupeesOnly(otsAmt) + ') देकर समझौता आधारित निपटान चाहते हैं, जिसका न्यूनतम धनराशि जो कि रु. ' + tokenAmt.toFixed(2) + '/- (' + otsAppHindiRupeesOnly(tokenAmt) + ') होता है, इस समझौता प्रस्ताव के साथ ही शाखा में जमा किया जा रहा है।')
    +   otsAppPoint(4, 'यह कि समझौता आधारित शेष निपटान राशि हम एकमुश्त दिनांक ' + restDateStr + ' तक जमा करने हेतु वचनबद्ध हैं।', true)
    +   otsAppPoint(5, 'यह कि निपटान राशि रु. ' + otsAmt.toFixed(2) + '/- हम स्वयं के स्रोत से जमा करेंगे,', true)
    +   otsAppPoint(6, 'निपटान राशि पर बैंक द्वारा निर्धारित वसूली प्रक्रिया/कलेक्शन चार्जेस, यथायोग्य देने हेतु सहमत हैं।')
    +   otsAppPoint(7, 'यह कि उपरोक्त वर्णित बिन्दुओं के अनुसार यदि उक्त निपटान राशि का कुल भुगतान समय से नहीं होता है, तो बैंक को यह अधिकार होगा कि सम्पूर्ण ऋण राशि की वसूली निर्धारित नियम एवं शर्तों के अनुरूप करे।')
    +   otsAppPoint(8, 'समझौता पश्चात् हमारे द्वारा 12 माह से पहले बैंक की किसी भी शाखा में किसी भी ऋण हेतु आवेदन प्रस्तुत नहीं किया जाएगा।', true)
    + '</div>'
    + '<div style="font-size:12px; margin-top:8px;">नोट: यह स्वीकृति 3 माह तक वैध रहेगी तथा 3 माह के अंदर पूर्ण समझौता न जमा होने पर उक्त स्वीकृति स्वतः निरस्त मानी जाएगी।</div>'
    + '<div style="margin-top:24px;">भवदीय,</div>'
    + '<div style="display:flex; justify-content:space-between; margin-top:64px;">'
    +   '<div style="font-size:14px;">'
    +     '<div>नाम: <b>' + esc(row[C.NAME]||'') + '</b></div>'
    +     '<div>पता: ' + esc(row[C.ADDR]||'') + '</div>'
    +     '<div>मो. नं.: ' + esc(row[C.PHONE]||'') + '</div>'
    +     '<div>दिनांक: ' + dateStr + '</div>'
    +   '</div>'
    +   '<div style="font-size:14px; text-align:right;">'
    +     '<div style="margin-bottom:56px;">हस्ताक्षर</div>'
    +     '<div>शाखा प्रबंधक</div>'
    +     '<div>शाखा: ' + esc(row[C.SOL_DESC]||'') + '</div>'
    +   '</div>'
    + '</div>'
    + '</div>';
}

let __otsAppRow = null;
let __otsAppLetterHtml = null;
// Alok, 2026-09-27: "agar data already fetch ho raha hai to ye ban hi
// jayega agar ismain data nahi milta to all required data fill karne k
// liye aaye aur application generate ho jaye but show tab hi kare jab
// data available na ho" -- when a search finds a real account, the
// existing auto-fill flow (below) is unchanged. When it finds nothing,
// this now offers a manual, type-everything-yourself fallback instead of
// just a dead-end "not found" message -- shown ONLY in that case, never
// alongside a real match.
let __otsAppManualMode = false;
function otsAppLockedRows(){
  const solId = loggedInSolId();
  if(!solId) return DATA.npa.rows;
  return DATA.npa.rows.filter(r=>String(r[C.SOL_ID])===String(solId));
}
function otsAppFindAccount(query){
  const q = String(query||'').trim().toLowerCase();
  if(!q) return null;
  const rows = otsAppLockedRows();
  return rows.find(r=>String(r[C.ACCT_NO]||'').toLowerCase()===q)
      || rows.find(r=>String(r[C.ACCT_NO]||'').toLowerCase().includes(q))
      || rows.find(r=>String(r[C.NAME]||'').toLowerCase().includes(q))
      || null;
}
function otsAppSearch(){
  const input = document.getElementById('otsAppSearchInput');
  const q = input ? input.value : '';
  const row = otsAppFindAccount(q);
  __otsAppRow = row;
  __otsAppLetterHtml = null;
  __otsAppManualMode = !row && !!q.trim();
  const statusEl = document.getElementById('otsAppSearchStatus');
  if(statusEl) statusEl.innerHTML = __otsAppManualMode ? '<div class="upload-status warn">No matching account found in your branch — fill in the details manually below.</div>' : '';
  renderOtsApplicationDetail();
}
window.otsAppSearch = otsAppSearch;
// Auto-suggest Branch Name/District off a typed Sol ID (BRANCH_LIST/
// BRANCH_META, same reference data the auto-fill card's own lookup
// uses) -- only fills a field the user hasn't already typed into,
// never overwrites something they've already entered themselves.
function otsAppManualSolIdChanged(){
  const solIdEl = document.getElementById('otsAppManualSolId');
  if(!solIdEl) return;
  const solId = Number((solIdEl.value||'').trim());
  if(!solId) return;
  const branchEl = document.getElementById('otsAppManualBranch');
  const distEl = document.getElementById('otsAppManualDistrict');
  const listEntry = BRANCH_LIST.find(([,nid])=>Number(nid)===solId);
  if(branchEl && !branchEl.value.trim() && listEntry) branchEl.value = listEntry[2];
  const meta = BRANCH_META[solId];
  if(distEl && !distEl.value.trim() && meta && meta.district) distEl.value = meta.district;
}
window.otsAppManualSolIdChanged = otsAppManualSolIdChanged;
function otsAppAddDays(date, days){ const d = new Date(date.getTime()); d.setDate(d.getDate()+days); return d; }
// Shared by both the auto-fill path and the manual-entry path below --
// reads the same 5 settlement fields (ids unchanged either way) and
// returns null (after alerting) if Outstanding is missing/invalid.
function otsAppReadSettlementFields(){
  const outstandingEl = document.getElementById('otsAppOutstanding');
  const outstandingRaw = (outstandingEl && outstandingEl.value || '').replace(/,/g,'').trim();
  const outstanding = Number(outstandingRaw);
  if(!outstandingRaw || isNaN(outstanding) || outstanding<=0){
    alert('Please enter a valid Outstanding amount.');
    return null;
  }
  const purpose = (document.getElementById('otsAppPurpose').value||'').trim();
  const otsAmt = Number((document.getElementById('otsAppOtsAmt').value||'').replace(/,/g,'')) || 0;
  const tokenAmt = Number((document.getElementById('otsAppTokenAmt').value||'').replace(/,/g,'')) || 0;
  const tokenDateVal = document.getElementById('otsAppTokenDate').value;
  let tokenDate = null;
  if(tokenDateVal){
    const parts = tokenDateVal.split('-');
    tokenDate = new Date(Number(parts[0]), Number(parts[1])-1, Number(parts[2]));
  }
  let restDate = null;
  if(tokenDate) restDate = (otsAmt===tokenAmt) ? tokenDate : otsAppAddDays(tokenDate, 89);
  return { outstanding, purpose, otsAmt, tokenAmt, tokenDate, restDate };
}
function otsAppRenderPreview(){
  const wrap = document.getElementById('otsAppPreviewWrap');
  if(!wrap || !__otsAppLetterHtml) return;
  wrap.innerHTML = '<div class="card">'
    + '<div style="font-weight:800;margin-bottom:10px;">Generated Application Form</div>'
    + '<div style="overflow-x:auto;border:1px solid var(--line);border-radius:10px;">' + __otsAppLetterHtml + '</div>'
    + '<div style="display:flex;gap:10px;flex-wrap:wrap;margin-top:14px;">'
    +   '<button type="button" class="export-xl-btn" onclick="otsAppPrint()">🖨 Print</button>'
    +   '<button type="button" class="export-xl-btn" onclick="otsAppSavePdf()">⬇ Save as PDF</button>'
    +   '<button type="button" class="export-xl-btn" onclick="otsAppSharePdf()">📤 Share on WhatsApp</button>'
    + '</div>'
    + '</div>';
  wrap.scrollIntoView({behavior:'smooth', block:'start'});
}
function otsAppGenerate(){
  if(__otsAppManualMode){ otsAppGenerateManual(); return; }
  const row = __otsAppRow;
  if(!row) return;
  const f = otsAppReadSettlementFields();
  if(!f) return;
  const acctOpenDate = toDate(row[C.OPN_DT]);
  const district = (BRANCH_META[Number(row[C.SOL_ID])]||{}).district || '';
  __otsAppLetterHtml = buildOtsApplicationFormHTML(row, {
    acctOpenDate, osatots: f.outstanding, otsamt: f.otsAmt, tokenamt: f.tokenAmt,
    tokendate: f.tokenDate, restdate: f.restDate, purpose: f.purpose, district
  });
  otsAppRenderPreview();
}
window.otsAppGenerate = otsAppGenerate;
// Manual-entry fallback: builds a synthetic C-indexed row (same shape the
// letter builder already expects from a real DATA.npa.rows entry) out of
// hand-typed fields, so buildOtsApplicationFormHTML() needs no changes at
// all -- it can't tell the difference between this and a real row.
function otsAppGenerateManual(){
  const req = (id, label) => {
    const el = document.getElementById(id);
    const v = (el && el.value || '').trim();
    if(!v){ alert('Please enter ' + label + '.'); if(el) el.focus(); return undefined; }
    return v;
  };
  const name = req('otsAppManualName', 'the Name'); if(name===undefined) return;
  const acctNo = req('otsAppManualAcctNo', 'the Account No.'); if(acctNo===undefined) return;
  const branch = req('otsAppManualBranch', 'the Branch Name'); if(branch===undefined) return;
  const address = req('otsAppManualAddress', 'the Address'); if(address===undefined) return;
  const loanAmtRaw = req('otsAppManualLoanAmt', 'the Loan Amount'); if(loanAmtRaw===undefined) return;
  const loanDateVal = document.getElementById('otsAppManualLoanDate').value;
  if(!loanDateVal){ alert('Please enter the Loan Date.'); return; }
  const solId = (document.getElementById('otsAppManualSolId').value||'').trim();
  const district = (document.getElementById('otsAppManualDistrict').value||'').trim();
  const mobile = (document.getElementById('otsAppManualMobile').value||'').trim();
  const loanAmt = Number(loanAmtRaw.replace(/,/g,'')) || 0;
  const ldp = loanDateVal.split('-');
  const acctOpenDate = new Date(Number(ldp[0]), Number(ldp[1])-1, Number(ldp[2]));

  const f = otsAppReadSettlementFields();
  if(!f) return;

  const row = [];
  row[C.NAME] = name;
  row[C.ACCT_NO] = acctNo;
  row[C.SOL_DESC] = branch;
  row[C.SOL_ID] = solId;
  row[C.ADDR] = address;
  row[C.PHONE] = mobile;
  row[C.SANCT_LIM] = loanAmt;
  __otsAppRow = row; // otsAppFileBase()/otsAppSharePdf() read this same way for both modes

  __otsAppLetterHtml = buildOtsApplicationFormHTML(row, {
    acctOpenDate, osatots: f.outstanding, otsamt: f.otsAmt, tokenamt: f.tokenAmt,
    tokendate: f.tokenDate, restdate: f.restDate, purpose: f.purpose, district
  });
  otsAppRenderPreview();
}
window.otsAppGenerateManual = otsAppGenerateManual;
function otsAppPrint(){
  if(!__otsAppLetterHtml) return;
  document.getElementById('printArea').innerHTML = __otsAppLetterHtml;
  printWithPageSize('size:A4;margin:12mm');
}
window.otsAppPrint = otsAppPrint;
/* Shared by Save-as-PDF and WhatsApp Share below -- rasterizes the letter
   off-screen via #printArea (same trick shareOtsPdf() above already uses:
   #printArea is display:none outside @media print, so it's floated
   on-screen at -9999px just long enough for html2canvas to capture it). */
async function otsAppRenderCanvas(){
  await Promise.all([ensureHtml2Canvas(), ensureJsPDF()]);
  const printEl = document.getElementById('printArea');
  const prevCss = printEl.style.cssText;
  printEl.innerHTML = __otsAppLetterHtml;
  printEl.style.cssText = 'display:block;position:fixed;left:-9999px;top:0;width:900px;background:#fff;padding:0;z-index:-1';
  if(document.fonts && document.fonts.ready){ try{ await document.fonts.ready; }catch(e){} }
  await new Promise(r=>setTimeout(r, 60));
  let canvas;
  try{
    canvas = await html2canvas(printEl, { scale:2, backgroundColor:'#ffffff' });
  } finally {
    printEl.style.cssText = prevCss;
    printEl.innerHTML = '';
  }
  return canvas;
}
/* Filename includes both the customer's name AND their branch name, per
   Alok's own request, 2026-09-27 -- makes a saved/shared file identifiable
   without opening it, unlike the OTS Calculator's own name-only filename. */
function otsAppFileBase(){
  const row = __otsAppRow || {};
  const safeName = String(row[C.NAME]||'borrower').replace(/[\\/:*?"<>|]/g,' ').replace(/\s+/g,' ').trim().slice(0,40);
  const safeBranch = String(row[C.SOL_DESC]||'').replace(/[\\/:*?"<>|]/g,' ').replace(/\s+/g,' ').trim().slice(0,30);
  return 'Application_Form_' + safeName + (safeBranch ? ('_' + safeBranch) : '');
}
async function otsAppSavePdf(){
  if(!__otsAppLetterHtml) return;
  try{
    const canvas = await otsAppRenderCanvas();
    const blob = canvasToPdfBlob(canvas);
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = otsAppFileBase() + '.pdf';
    document.body.appendChild(a); a.click(); document.body.removeChild(a);
    setTimeout(()=>URL.revokeObjectURL(url), 30000);
  }catch(err){
    console.error(err);
    alert('Could not prepare the PDF. Please try again.');
  }
}
window.otsAppSavePdf = otsAppSavePdf;
async function otsAppSharePdf(){
  if(!__otsAppLetterHtml) return;
  try{
    const canvas = await otsAppRenderCanvas();
    const blob = canvasToPdfBlob(canvas);
    const row = __otsAppRow || {};
    const shareText = (row[C.NAME]||'—') + ' — ' + (row[C.SOL_DESC]||'—') + ' — Application Form';
    await shareFileOrFallback(blob, otsAppFileBase() + '.pdf', 'application/pdf', shareText);
  }catch(err){
    console.error(err);
    alert('Could not prepare the PDF to share. Please try again.');
  }
}
window.otsAppSharePdf = otsAppSharePdf;
// Shared by both the auto-fill card (below) and the manual-entry card --
// the 5 fields that are ALWAYS typed by hand, in either mode. hintHtml is
// the small note under Outstanding; only the auto-fill path has a ledger
// figure to reference, so the manual path passes ''.
function otsAppSettlementFieldsCardHtml(hintHtml){
  return '<div class="card">'
    + '<div style="font-weight:800;margin-bottom:12px;">Please enter these details</div>'
    + '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px;">'
    +   '<div><label style="display:block;font-size:12px;font-weight:700;color:var(--sub);margin-bottom:5px;">Outstanding as on Date of OTS <span style="color:#d1425a;">*</span></label>'
    +   '<input type="text" inputmode="decimal" id="otsAppOutstanding" class="dash-select" style="width:100%;min-width:0" placeholder="e.g. 245000">'
    +   (hintHtml ? ('<div style="font-size:11px;color:var(--sub);margin-top:4px;">' + hintHtml + '</div>') : '')
    + '</div>'
    +   '<div><label style="display:block;font-size:12px;font-weight:700;color:var(--sub);margin-bottom:5px;">Purpose of Loan</label>'
    +   '<input type="text" id="otsAppPurpose" class="dash-select" style="width:100%;min-width:0" placeholder="e.g. पशुपालन हेतु"></div>'
    +   '<div><label style="display:block;font-size:12px;font-weight:700;color:var(--sub);margin-bottom:5px;">OTS / Compromise Amount</label>'
    +   '<input type="text" inputmode="decimal" id="otsAppOtsAmt" class="dash-select" style="width:100%;min-width:0" placeholder="e.g. 140000"></div>'
    +   '<div><label style="display:block;font-size:12px;font-weight:700;color:var(--sub);margin-bottom:5px;">Token Amount</label>'
    +   '<input type="text" inputmode="decimal" id="otsAppTokenAmt" class="dash-select" style="width:100%;min-width:0" placeholder="e.g. 25000"></div>'
    +   '<div><label style="display:block;font-size:12px;font-weight:700;color:var(--sub);margin-bottom:5px;">Token Date</label>'
    +   '<input type="date" id="otsAppTokenDate" class="dash-select" style="width:100%;min-width:0"></div>'
    + '</div>'
    + '<div style="margin-top:16px;">'
    +   '<button type="button" class="export-xl-btn" onclick="otsAppGenerate()">✓ Generate Application Form</button>'
    + '</div>'
    + '</div>';
}
function otsAppManualField(id, label, type, placeholder, required){
  return '<div><label style="display:block;font-size:12px;font-weight:700;color:var(--sub);margin-bottom:5px;">' + esc(label) + (required?' <span style="color:#d1425a;">*</span>':'') + '</label>'
    + '<input type="' + type + '" id="' + id + '" class="dash-select" style="width:100%;min-width:0"'
    + (placeholder?(' placeholder="'+esc(placeholder)+'"'):'')
    + (id==='otsAppManualSolId'?' oninput="otsAppManualSolIdChanged()"':'')
    + '></div>';
}
function otsAppManualFormHtml(){
  return '<div class="card">'
    + '<div style="font-weight:800;margin-bottom:4px;">Enter Account Details Manually</div>'
    + '<div style="font-size:12px;color:var(--sub);margin-bottom:12px;">This account was not found in the loaded NPA data — fill in these details yourself to still generate the Application Form.</div>'
    + '<div style="display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:12px;">'
    +   otsAppManualField('otsAppManualName','Name','text','',true)
    +   otsAppManualField('otsAppManualAcctNo','Account No.','text','',true)
    +   otsAppManualField('otsAppManualSolId','Sol ID','text','e.g. 9270',false)
    +   otsAppManualField('otsAppManualBranch','Branch Name','text','',true)
    +   otsAppManualField('otsAppManualDistrict','District','text','',false)
    +   otsAppManualField('otsAppManualAddress','Address','text','',true)
    +   otsAppManualField('otsAppManualMobile','Mobile No.','text','',false)
    +   otsAppManualField('otsAppManualLoanAmt','Loan Amount','text','e.g. 200000',true)
    +   otsAppManualField('otsAppManualLoanDate','Loan Date','date','',true)
    + '</div>'
    + '</div>'
    + otsAppSettlementFieldsCardHtml('')
    + '<div id="otsAppPreviewWrap"></div>';
}
function renderOtsApplicationDetail(){
  const wrap = document.getElementById('otsAppDetail');
  if(!wrap) return;
  if(__otsAppManualMode){ wrap.innerHTML = otsAppManualFormHtml(); return; }
  const row = __otsAppRow;
  if(!row){ wrap.innerHTML = ''; return; }
  const meta = BRANCH_META[Number(row[C.SOL_ID])] || {};
  const acctOpenDate = toDate(row[C.OPN_DT]);
  wrap.innerHTML = '<div class="card">'
    + '<div style="font-weight:800;margin-bottom:4px;">Account Details (Auto-Filled from Records)</div>'
    + '<div class="info-grid">'
    +   '<div><div class="k">Name</div><div class="v">' + (esc(row[C.NAME])||'—') + '</div></div>'
    +   '<div><div class="k">Account No.</div><div class="v">' + (esc(row[C.ACCT_NO])||'—') + '</div></div>'
    +   '<div><div class="k">Branch</div><div class="v">' + (esc(row[C.SOL_DESC])||'—') + ' (' + (esc(row[C.SOL_ID])||'—') + ')</div></div>'
    +   '<div><div class="k">District</div><div class="v">' + (esc(meta.district)||'—') + '</div></div>'
    +   '<div><div class="k">Mobile No.</div><div class="v">' + (esc(row[C.PHONE])||'—') + '</div></div>'
    +   '<div><div class="k">Loan Amount / Date</div><div class="v">' + fmtINR2(Number(row[C.SANCT_LIM])||0) + ' · ' + fmtDate(acctOpenDate) + '</div></div>'
    +   '<div><div class="k">Outstanding (as per records)</div><div class="v">' + fmtINR2(Number(row[C.OUTBAL])||0) + '</div></div>'
    + '</div>'
    + '</div>'
    + otsAppSettlementFieldsCardHtml('Records show ' + fmtINR2(Number(row[C.OUTBAL])||0) + ' — enter the actual figure as on the settlement date.')
    + '<div id="otsAppPreviewWrap"></div>';
}
function renderOtsApplicationView(){
  const el = document.getElementById('otsApplicationArea');
  if(!el || el.dataset.wired) return;
  el.innerHTML = '<div class="card">'
    + '<div style="display:flex;gap:10px;flex-wrap:wrap;align-items:center;">'
    +   '<input type="text" id="otsAppSearchInput" class="dash-select" style="flex:1;min-width:220px" placeholder="Account No. or Customer Name" onkeydown="if(event.key===\'Enter\'){otsAppSearch();}">'
    +   '<button type="button" class="export-xl-btn" onclick="otsAppSearch()">Find</button>'
    + '</div>'
    + '<div id="otsAppSearchStatus"></div>'
    + '</div>'
    + '<div id="otsAppDetail"></div>';
  el.dataset.wired = '1';
  renderOtsApplicationDetail();
}

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
// Reverse lookup, Sol ID -> canonical branch name, straight off BRANCH_LIST
// (guaranteed complete for every valid Sol ID). Alok, 2026-09-25: "9283
// kholne par sari branches show ho rahe hain" -- Sol 9283 (Hathras Service
// Branch) has zero rows of its own in the NPA book / KCC Overdue / PNPA
// (a back-office branch with no live loan accounts), so every branch-lock
// resolver below, which only ever looked for the Sol ID's branch name
// among the branch names actually PRESENT in that dataset, silently fell
// back to '' (empty = "Regional Office (all branches)") for exactly this
// case -- the one branch-name source that can never come up empty for a
// valid Sol ID is this master list itself, used as the last-resort
// fallback so a real Sol ID always locks to SOME specific branch (even one
// with zero rows in a given table -- correctly showing "your branch, zero
// accounts" there, never "show everyone's").
const SOL_TO_BRANCH_NAME = Object.fromEntries(BRANCH_LIST.map(([,newId,name])=>[String(newId), name]));
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
// initLedgerAcctNumbers() MUST run before the prefill call below -- it
// captures each account's own real 10-digit suffix (e.g. "0015181213")
// off the pristine "XXXX..." markup into el.dataset.acctSuffix, which
// onLedgerSolIdInput() then reads back. Calling the prefill first (as this
// block briefly did) overwrites every account's textContent down to just
// the 4-digit Sol ID *before* its suffix is ever captured -- by the time
// initLedgerAcctNumbers() ran, textContent.slice(4) had nothing left to
// read, so every account's dataset.acctSuffix silently became '' and every
// row showed only the Sol ID (confirmed against a real screenshot, all 6
// rows reading "9321" instead of "9321" + their own suffix, 2026-09-27).
initLedgerAcctNumbers();
(function prefillLedgerSolId(){
  const solId = loggedInSolId();
  if(solId) onLedgerSolIdInput(solId);
})();
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
// custId is optional -- passed by callers whose own rows carry one (KCC
// Overdue, PNPA) so an account absent from the NPA book can still resolve
// via DATA.customerAddressMap, the one address source that isn't scoped to
// "currently in the NPA book" (see that map's own init comment).
function addressForAcctNo(acctNo, custId){
  const rows = DATA.npa && DATA.npa.rows;
  if(__addrByAcctCache.forRows !== rows){
    const m = new Map();
    (rows||[]).forEach(r=>{ if(r[C.ADDR]) m.set(normId(r[C.ACCT_NO]), r[C.ADDR]); });
    __addrByAcctCache = { forRows: rows, map: m };
  }
  const key = normId(acctNo);
  const custKey = custId ? normId(custId) : '';
  return __addrByAcctCache.map.get(key)
    || DATA.addressList[key] || DATA.addressListByCustomer[key]
    || (custKey ? DATA.customerAddressMap[custKey] : '')
    || '';
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
  // Falls back to BRANCH_LIST's own canonical name when this Sol ID has no
  // row of its own in the NPA book (e.g. a Service Branch with no live
  // loan accounts) -- see SOL_TO_BRANCH_NAME's own comment for why this
  // matters: without it, this silently resolved to '' ("Regional Office",
  // i.e. every branch), defeating the whole point of Sol-based login.
  __recoveryBranchName = row ? (row[C.SOL_DESC]||'') : (SOL_TO_BRANCH_NAME[String(solId)] || '');
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
   miniature), and the branch's own FY-end NPA target -- now a real figure
   once NPA-DASHBOARD's "Branch/Region NPA Target" upload has been
   Published (DATA.branchTargets, same shared cross-origin data this
   portal already reads everything else from), shown as "Not set yet"
   only when that upload genuinely hasn't happened for this branch. The
   final month in the uploaded target series is used (not hardcoded to
   "March 2027") so a future fiscal year's re-upload with a different
   final month still shows correctly with no code change here. These are
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
  const bTargets = solId && DATA.branchTargets && DATA.branchTargets.branches ? DATA.branchTargets.branches[String(solId)] : null;
  const finalTarget = bTargets && bTargets.targets && bTargets.targets.length ? bTargets.targets[bTargets.targets.length-1] : null;
  const targetTileHtml = (finalTarget && finalTarget.rupees!=null)
    ? (()=>{ const gap = todayNpa - finalTarget.rupees; const improved = gap<=0;
        return `<div class="npa-target-tile"><span class="lbl">Target for ${esc(finalTarget.label)}</span><span class="val">${fmtCr(finalTarget.rupees)}</span><span class="val-sub" style="color:${improved?'var(--green)':'var(--red)'}">${improved?'▼':'▲'} ${fmtCr(Math.abs(gap))}</span></div>`; })()
    : `<div class="npa-target-tile"><span class="lbl">Target for March 2027</span><span class="val muted">Not set yet</span></div>`;
  return `<div class="npa-target-strip">
    <div class="npa-target-tile"><span class="lbl">Today's NPA</span><span class="val">${fmtCr(todayNpa)}</span></div>
    ${marGapHtml}
    ${targetTileHtml}
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
const PC = {REGION:0, BRANCH:1, SCHEME:2, ACCT:3, NAME:4, OS:5, CADU:6, LIMIT:7, REVIEW:8, REASON:9, CUSTNPADATE:10, CUST_ID:11};
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
  // Falls back to BRANCH_LIST's own canonical name when this Sol ID's
  // branch has no rows of its own in this dataset (e.g. no PNPA slippage
  // this period) -- see SOL_TO_BRANCH_NAME's own comment.
  __pnpaLockedBranch = allBranches.find(b=>String(KCCOV_BRANCH_SOL[String(b).toUpperCase()])===String(solId)) || SOL_TO_BRANCH_NAME[String(solId)] || '';
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
// 'all'|'kcc'|'nonkcc' -- deliberately independent of pnpaSlipTab (picking
// KCC then switching to This Month keeps showing just KCC accounts for
// that period, rather than silently resetting the scheme choice).
let pnpaSlipSchemeFilter = 'all';
function setPnpaSlipTab(tab){ pnpaSlipTab = tab; renderPnpaSlipView(); }
window.setPnpaSlipTab = setPnpaSlipTab;
function setPnpaSlipScheme(scheme){ pnpaSlipSchemeFilter = scheme; renderPnpaSlipView(); }
window.setPnpaSlipScheme = setPnpaSlipScheme;
// The Today hero cards jump to the Today tab AND set the scheme filter in
// one render, per the existing documented behavior ("Clicking any card
// jumps to the Today tab") -- calling setPnpaSlipTab()+setPnpaSlipScheme()
// separately would trigger two renders back to back.
function pnpaSlipTodayCardClick(scheme){ pnpaSlipTab = 'today'; pnpaSlipSchemeFilter = scheme; renderPnpaSlipView(); }
window.pnpaSlipTodayCardClick = pnpaSlipTodayCardClick;

// PNPA rows carry no per-row address of their own -- resolved via the
// shared addressForAcctNo() (see its own comment near normId()), which
// looks at the main NPA book's own already-merged addresses first, falling
// back to the Customer-ID-keyed DATA.customerAddressMap (via custId, when
// PNPA_DATA has one) since PNPA accounts don't exist in the NPA book.
function pnpaAddressFor(acctNo, custId){
  return addressForAcctNo(acctNo, custId);
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
    list = list.filter(r=>pnpaAddressFor(r[PC.ACCT], r[PC.CUST_ID]).toLowerCase().includes(q));
  }
  if(!list.length) return `<div class="empty-state"><p>${esc(emptyMessage || 'No accounts slipped in this period.')}</p></div>`;
  const remarks = getPnpaRemarks();
  const objs = list.map(r=>({
    acctNo:r[PC.ACCT], name:r[PC.NAME], address:pnpaAddressFor(r[PC.ACCT], r[PC.CUST_ID]), os:r[PC.OS], cadu:r[PC.CADU],
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
// Alok, 2026-09-27: "in par click karne se kuch nahi ho raha jabki in par
// click karne par respective accounts ki list hi niche aani chahiye" --
// these 3 chips (and the matching 3 cards in pnpaTodayHeroBlocks below)
// used to be pure display, no scheme filter existed anywhere in this view
// at all -- clicking KCC/Non-KCC never narrowed the table below to just
// that scheme. Wired to the new pnpaSlipSchemeFilter state below; `active`
// marks whichever scheme is currently selected (not just the ever-blue
// `.total` styling, which was being mistaken for a selected state).
function pnpaSlipSummaryChips(totals, activeScheme){
  const chip = (key, cls, label, t) => `<div class="pnpa-slip-chip clickable${cls?' '+cls:''}${activeScheme===key?' active':''}" onclick="setPnpaSlipScheme('${key}')"><span class="lbl">${label}</span><span class="cnt">${t.cnt.toLocaleString('en-IN')} A/C</span><span class="amt">${fmtCr(t.amt)}</span></div>`;
  return `<div class="pnpa-slip-summary">
    ${chip('kcc','','KCC',totals.kcc)}
    ${chip('nonkcc','','Non-KCC',totals.nonkcc)}
    ${chip('all','total','Total',totals.all)}
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
function pnpaTodayHeroBlocks(todayTotals, pnpaSlipTab, activeScheme){
  const isActive = (scheme) => pnpaSlipTab==='today' && activeScheme===scheme;
  return `<div class="pnpa-today-hero">
    <div class="pnpa-today-hero-head">Today's Slippage</div>
    <div class="pnpa-today-hero-row">
      <div class="pnpa-today-card total clickable${isActive('all')?' active':''}" onclick="pnpaSlipTodayCardClick('all')">
        <span class="lbl">Total Slippage</span>
        <span class="cnt">${todayTotals.all.cnt.toLocaleString('en-IN')} A/C</span>
        <span class="amt">${fmtCr(todayTotals.all.amt)}</span>
      </div>
      <div class="pnpa-today-card kcc clickable${isActive('kcc')?' active':''}" onclick="pnpaSlipTodayCardClick('kcc')">
        <span class="lbl">KCC Slippage</span>
        <span class="cnt">${todayTotals.kcc.cnt.toLocaleString('en-IN')} A/C</span>
        <span class="amt">${fmtCr(todayTotals.kcc.amt)}</span>
      </div>
      <div class="pnpa-today-card nonkcc clickable${isActive('nonkcc')?' active':''}" onclick="pnpaSlipTodayCardClick('nonkcc')">
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
  const filteredList = pnpaSlipSchemeFilter==='all' ? active.list : active.list.filter(r=>pnpaSlipSchemeKey(r)===pnpaSlipSchemeFilter);
  const schemeLabel = {all:'', kcc:'KCC ', nonkcc:'Non-KCC '}[pnpaSlipSchemeFilter];
  const emptyMessages = {
    today: `No ${schemeLabel}accounts slipped today.`,
    week: __pnpaWeeklyFailed ? 'No Weekly PNPA file uploaded yet.' : `No ${schemeLabel}accounts in this week's slippage file.`,
    month: __pnpaMonthlyFailed ? 'No Monthly PNPA file uploaded yet.' : `No ${schemeLabel}accounts in this month's slippage file.`,
  };
  el.innerHTML = `
    ${pnpaTodayHeroBlocks(agg.today.totals, pnpaSlipTab, pnpaSlipSchemeFilter)}
    ${tabsHtml}
    <div class="bank-filter-row">
      <input type="text" id="pnpaAddressFilterInput" class="dash-select" placeholder="Filter by Address…" value="${esc(pnpaAddressFilter)}" style="max-width:220px">
    </div>
    ${pnpaSlipSummaryChips(active.totals, pnpaSlipSchemeFilter)}
    ${renderPnpaSlipTable(filteredList, emptyMessages[pnpaSlipTab])}
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
const KC = {BRANCH:0, SCHEME:1, ACCT:2, NAME:3, OS:4, CADU:5, LIMIT:6, REVIEW:7, CUSTNPADATE:8, FY:9, CATEGORY:10, SMA:11, REASON:12, CUST_ID:13};
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
// which looks at the main NPA book's own already-merged addresses first,
// falling back to the Customer-ID-keyed DATA.customerAddressMap (via
// custId, when KCC_OVERDUE_DATA has one) since KCC Overdue accounts don't
// exist in the NPA book at all. Alok, 2026-09-25: "fir kcc overdue main
// kyun nahi aa raha" (why isn't it showing in KCC Overdue).
function kccovAddressFor(acctNo, custId){
  return addressForAcctNo(acctNo, custId);
}
function kccovFilteredRows(d){
  let rows = d.rows;
  if(kccovBranchFilter) rows = rows.filter(r=>r[KC.BRANCH]===kccovBranchFilter);
  if(kccovFyFilter) rows = rows.filter(r=>r[KC.FY]===kccovFyFilter);
  if(kccovAddressFilter){
    const q = kccovAddressFilter.trim().toLowerCase();
    rows = rows.filter(r=>kccovAddressFor(r[KC.ACCT], r[KC.CUST_ID]).toLowerCase().includes(q));
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
  // Falls back to BRANCH_LIST's own canonical name when this Sol ID's
  // branch has no rows of its own in this dataset (e.g. no KCC Overdue
  // accounts this period) -- see SOL_TO_BRANCH_NAME's own comment. Without
  // this, the branch lock silently fell through to an open, all-branches
  // picker instead of a locked label for a real Sol ID.
  const lockedKccBranch = lockedSolId ? (allBranches.find(b=>String(KCCOV_BRANCH_SOL[String(b).toUpperCase()])===String(lockedSolId)) || SOL_TO_BRANCH_NAME[String(lockedSolId)] || null) : null;
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
    acctNo:r[KC.ACCT], name:r[KC.NAME], address:kccovAddressFor(r[KC.ACCT], r[KC.CUST_ID]), os:r[KC.OS], cadu:r[KC.CADU], limit:r[KC.LIMIT],
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
  const list = rows.map(r=>({ acctNo:r[KC.ACCT], name:r[KC.NAME], address:kccovAddressFor(r[KC.ACCT], r[KC.CUST_ID]), os:r[KC.OS], cadu:r[KC.CADU], limit:r[KC.LIMIT], custNpaDate:r[KC.CUSTNPADATE], fy:r[KC.FY], category:r[KC.CATEGORY], sma:r[KC.SMA] }));
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
  return rows.map(r=>({ acctNo:r[KC.ACCT], name:r[KC.NAME], address:kccovAddressFor(r[KC.ACCT], r[KC.CUST_ID]), os:r[KC.OS], cadu:r[KC.CADU], limit:r[KC.LIMIT], custNpaDate:r[KC.CUSTNPADATE], fy:r[KC.FY], category:r[KC.CATEGORY], sma:r[KC.SMA] }));
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
    if(view==='otsapplication') renderOtsApplicationView();
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
