/* Login screen — Sol ID acts as both the branch identity and the 4-digit
   PIN (e.g. 9270 = Agsauli). This is a "simple login + auto-filter" model,
   not real access control: the underlying data is decrypted the same way
   for every branch (see DATA_DECRYPT_KEY below), and any Sol ID is
   guessable/public (printed on every report/letterhead) -- the point is
   "show me my own branch by default", not "keep other branches' data
   secret from a determined user". Confirmed with Alok directly. */
(function () {
  // MUST stay in sync with the production site's js/splash.js CORRECT_PIN
  // -- data/latest.json etc. are encrypted against that one value only. If
  // Alok ever rotates the production PIN, this constant has to be updated
  // here too, by hand, in this separate repo.
  const DATA_DECRYPT_KEY = '9269';

  // [Sol ID, Branch name] -- copied from BRANCH_LIST in the production
  // repo's js/app.js. Only the fields this login screen needs (not the
  // full BRANCH_LIST/BRANCH_META, which also carry old Sol codes/address/
  // district and only exist after data has loaded, too late for login).
  const SOL_BRANCHES = [[9269,"R O Hathras"],[9270,"Agsauli"],[9271,"Bamnai"],[9272,"Bandhnoo"],[9273,"Baraus"],[9274,"Bastoi"],[9275,"Bisawar"],[9276,"Chandpa"],[9277,"Chhonda Gadua"],[9278,"Devinagar"],[9279,"Eihan"],[9280,"Hathras Agra Road"],[9281,"Hathras Aligarh Road"],[9282,"Mursan Gate"],[9283,"Hathras Service Branch"],[9284,"Hatisa Bhagwantpur"],[9285,"Jarera"],[9286,"Komari"],[9287,"Kota"],[9288,"Ladpur"],[9289,"Mahow"],[9290,"Meetai"],[9291,"Mendu"],[9292,"Mughal Garhi"],[9293,"Mursan"],[9294,"Parsara"],[9295,"Pora"],[9296,"Purdil Nagar"],[9297,"Ratibhanpur"],[9298,"Ruheri"],[9299,"Sadabad"],[9300,"Sahpau"],[9301,"Salempur"],[9302,"Sasni"],[9303,"Sikandra Rao"],[9304,"Tuksan"],[9305,"Wazidpur"],[9306,"Adarshnagar"],[9307,"Hasayan"],[9308,"Jaleser Road"],[9309,"Naugaon"],[9310,"Bajna"],[9311,"Baldev"],[9312,"Bati"],[9313,"Damodarpura"],[9314,"Farah"],[9315,"Goverdhan"],[9316,"Maant"],[9317,"Mathura City"],[9318,"Laxmi Nagar"],[9319,"Pali Kheda"],[9320,"Raya"],[9321,"Ronchi Bangar"],[9322,"Sonai"],[9323,"Tarsi"],[9324,"Vrindavan"],[9325,"Jajan Patti"]];
  const SOL_BY_ID = Object.fromEntries(SOL_BRANCHES.map(([id, name]) => [String(id), name]));

  function isValidSolId(v) { return Object.prototype.hasOwnProperty.call(SOL_BY_ID, v); }

  const screen = document.getElementById('splashScreen');
  if (!screen || screen.classList.contains('skip')) return;

  const wrap = document.getElementById('splashPinWrap');
  const errorEl = document.getElementById('splashPinError');
  const cellsEl = document.getElementById('splashPinCells');
  const padEl = document.getElementById('splashPad');
  const loginBtn = document.getElementById('splashLoginBtn');
  if (!wrap || !cellsEl || !padEl) return;

  const reduceMotion = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;

  const cells = Array.from(cellsEl.querySelectorAll('.splash-cell'));
  let value = '';
  let locked = false; // true while a wrong Sol ID is shaking, or after unlock

  function paint() {
    cells.forEach((c, i) => {
      c.textContent = value[i] ? '•' : '';
      c.classList.toggle('filled', !!value[i]);
    });
  }
  function setError(text, ok) {
    errorEl.textContent = text || ' ';
    errorEl.classList.toggle('ok', !!ok);
  }
  function unlock(solId) {
    locked = true;
    setError('Welcome, ' + SOL_BY_ID[solId], true);
    try {
      sessionStorage.setItem('upgb-splash-unlocked', '1');
      // Always the shared decrypt secret, regardless of which Sol ID
      // logged in -- see the comment on DATA_DECRYPT_KEY above.
      sessionStorage.setItem('upgb-splash-pin', DATA_DECRYPT_KEY);
      sessionStorage.setItem('upgb-sol-id', solId); // branch-scoping for the rest of the session
    } catch (e) {}
    // app.js's data fetch needs the decrypt key to be present, but app.js
    // finishes executing well before a human finishes entering 4 digits
    // here -- this event is how app.js knows to wait for an actual unlock
    // instead of racing it (see its own listener, registered only when no
    // key is in sessionStorage yet at startup). Same contract as the
    // production site's splash.js -- app.js needed zero changes for this.
    try { window.dispatchEvent(new CustomEvent('upgb-pin-unlocked')); } catch (e) {}
    setTimeout(() => {
      screen.classList.add('unlocked');
      setTimeout(() => { screen.style.display = 'none'; }, 700);
    }, reduceMotion ? 0 : 350);
  }
  function reject() {
    locked = true;
    setError('Sol ID not recognised — try again');
    wrap.classList.add('shake');
    setTimeout(() => {
      wrap.classList.remove('shake');
      value = ''; paint(); locked = false;
    }, 420);
  }
  function shakeIncomplete() {
    wrap.classList.add('shake');
    setTimeout(() => wrap.classList.remove('shake'), 420);
  }
  function submit() {
    if (locked || value.length !== 4) { shakeIncomplete(); return; }
    if (isValidSolId(value)) unlock(value); else reject();
  }
  function push(d) {
    if (locked || value.length >= 4) return;
    setError('');
    value += d;
    paint();
    if (value.length === 4) setTimeout(submit, 170);
  }
  function back() {
    if (locked) return;
    setError('');
    value = value.slice(0, -1);
    paint();
  }

  ['1', '2', '3', '4', '5', '6', '7', '8', '9', '', '0', '⌫'].forEach(k => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'splash-key' + (k === '⌫' || k === '' ? ' ghost' : '');
    b.textContent = k;
    if (k === '') { b.disabled = true; b.style.visibility = 'hidden'; }
    b.setAttribute('aria-label', k === '⌫' ? 'Delete last digit' : k);
    b.addEventListener('click', () => { k === '⌫' ? back() : push(k); });
    padEl.appendChild(b);
  });

  if (loginBtn) loginBtn.addEventListener('click', submit);

  // A physical keyboard still works -- the on-screen pad exists so a phone
  // does not raise its own keyboard over the sheet, not to replace typing.
  document.addEventListener('keydown', e => {
    if (screen.classList.contains('unlocked') || screen.classList.contains('skip')) return;
    if (e.key >= '0' && e.key <= '9') { e.preventDefault(); push(e.key); }
    else if (e.key === 'Backspace') { e.preventDefault(); back(); }
    else if (e.key === 'Enter') { e.preventDefault(); submit(); }
  });

  paint();
  setError('');

  // Same one-time write-in animation as the production site's splash.js --
  // org name/title/subtitle in the hero, then the credit line + signature
  // below the login button, in DOM order.
  (function playCredit() {
    const segs = Array.from(document.querySelectorAll('.splash-wipe')).map(wipeEl => {
      const box = wipeEl.closest('.splash-linewrap, .splash-credit-sigwrap');
      const tipEl = box ? box.querySelector('.splash-pen-tip') : null;
      const chars = wipeEl.textContent ? wipeEl.textContent.length : 22;
      const duration = Math.max(600, chars * 32);
      return { wipeEl, tipEl, duration };
    });
    if (reduceMotion) {
      segs.forEach(seg => seg.wipeEl.classList.add('play'));
      return;
    }
    let t = 260;
    segs.forEach(seg => {
      seg.wipeEl.style.animationDuration = seg.duration + 'ms';
      if (seg.tipEl) seg.tipEl.style.animationDuration = seg.duration + 'ms';
      setTimeout(() => {
        seg.wipeEl.classList.add('play');
        if (seg.tipEl) seg.tipEl.classList.add('play');
      }, t);
      t += seg.duration + 160;
    });
  })();

  // Sign-out affordance for a shared branch PC -- clears all 3 session
  // keys and reloads so the next person sees the login screen fresh. Not
  // wired to a visible button yet by default; app.js's own sign-out/menu
  // wiring (once ported in M2) should call this. Exposed on window so it
  // can be reached from a menu item added later.
  window.upgbSignOut = function () {
    try {
      sessionStorage.removeItem('upgb-splash-unlocked');
      sessionStorage.removeItem('upgb-splash-pin');
      sessionStorage.removeItem('upgb-sol-id');
    } catch (e) {}
    location.reload();
  };
})();
