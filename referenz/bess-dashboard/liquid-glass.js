/**
 * liquid-glass.js — parallax for hero glass orbs, purely decorative.
 * Adds no state, doesn't touch app logic.
 */
(() => {
  "use strict";
  const hero = document.getElementById("hero");
  const blobA = document.querySelector(".hero-blob-a");
  const blobB = document.querySelector(".hero-blob-b");
  if (!hero || !blobA || !blobB) return;

  let targetX = 0, targetY = 0, curX = 0, curY = 0;
  let scrollOffset = 0;
  let rafId = null;
  let scrollActive = false;
  let scrollIdleTimer = null;
  let promoted = false;

  // will-change haelt einen filter: blur(50px)-Layer als GPU-Textur. Das lohnt
  // nur, solange tatsaechlich animiert wird -- dauerhaft im Stylesheet stand es
  // frueher fuer alle 10 Blobs im Dokument, von denen hier ohnehin nur das
  // erste Paar bewegt wird.
  function promote(on) {
    if (on === promoted) return;
    promoted = on;
    blobA.style.willChange = on ? "transform" : "auto";
    blobB.style.willChange = on ? "transform" : "auto";
  }

  function maybeDemote() {
    if (rafId === null && !scrollActive) promote(false);
  }

  function applyTransform() {
    blobA.style.transform = `translate(${curX * 30}px, ${curY * 30 - scrollOffset}px)`;
    blobB.style.transform = `translate(${curX * -22}px, ${curY * -22 - scrollOffset * 0.7}px)`;
  }

  function ensureTicking() {
    if (rafId === null && !document.hidden) {
      promote(true);
      rafId = requestAnimationFrame(tick);
    }
  }

  hero.addEventListener("mousemove", (e) => {
    const rect = hero.getBoundingClientRect();
    targetX = (e.clientX - rect.left - rect.width / 2) / rect.width;
    targetY = (e.clientY - rect.top - rect.height / 2) / rect.height;
    ensureTicking();
  });

  hero.addEventListener("mouseleave", () => {
    targetX = 0; targetY = 0;
    ensureTicking();
  });

  const tabs = document.getElementById("page-tabs");
  const header = document.querySelector(".app-header");
  let wasScrolled = null;

  // EIN Scroll-Listener statt frueher drei. Der Scroll-Parallax wird hier
  // direkt geschrieben, statt die rAF-Schleife anzustossen: targetX/Y aendern
  // sich beim Scrollen nicht, die Schleife lief also waehrend des gesamten
  // Scrollens ohne Grund weiter. Entfallen ist ausserdem ein toter Style-Write
  // (--scrollY auf blobA), den kein Stylesheet je gelesen hat.
  window.addEventListener("scroll", () => {
    scrollOffset = window.scrollY * 0.15;
    scrollActive = true;
    promote(true);
    clearTimeout(scrollIdleTimer);
    scrollIdleTimer = setTimeout(() => { scrollActive = false; maybeDemote(); }, 200);

    // Laeuft die Schleife, schreibt sie den Transform im naechsten Frame
    // ohnehin -- dann hier nicht doppelt schreiben.
    if (rafId === null) applyTransform();

    // Nur bei echtem Zustandswechsel anfassen: der Toggle loest auf zwei
    // Sticky-Leisten ueber die volle Breite Style-Recalcs aus.
    const scrolled = window.scrollY > 12;
    if (scrolled !== wasScrolled) {
      wasScrolled = scrolled;
      if (tabs) tabs.classList.toggle("is-scrolled", scrolled);
      if (header) header.classList.toggle("is-scrolled", scrolled);
    }
  }, { passive: true });

  document.addEventListener("visibilitychange", () => {
    if (document.hidden && rafId !== null) {
      cancelAnimationFrame(rafId);
      rafId = null;
      promote(false);
    } else {
      ensureTicking();
    }
  });

  const SETTLE_EPSILON = 0.001;

  function tick() {
    curX += (targetX - curX) * 0.06;
    curY += (targetY - curY) * 0.06;
    // scrollOffset kommt aus dem Scroll-Listener, statt hier window.scrollY zu
    // lesen -- das war ein erzwungener Layout-Read direkt nach den
    // Style-Writes der Vorzeile.
    applyTransform();

    const settled = Math.abs(targetX - curX) < SETTLE_EPSILON && Math.abs(targetY - curY) < SETTLE_EPSILON;
    rafId = settled ? null : requestAnimationFrame(tick);
    if (settled) maybeDemote();
  }

  applyTransform();
})();
