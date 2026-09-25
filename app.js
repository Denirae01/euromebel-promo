(function () {
  "use strict";
  var C = window.PROMO_CONFIG || {};
  var STORE_KEY = "em_promo_" + (C.campaign || "default");

  // ---------- подставляем настройки из config.js ----------
  function setAll(sel, fn) { document.querySelectorAll(sel).forEach(fn); }
  if (C.prize) setAll("[data-prize]", function (el) { el.textContent = C.prize; });
  if (C.prizeScope) setAll("[data-scope]", function (el) { el.textContent = C.prizeScope; });
  if (C.catalogUrl) setAll("[data-catalog]", function (el) { el.href = C.catalogUrl; });
  if (C.privacyUrl) setAll("[data-privacy]", function (el) { el.href = C.privacyUrl; });
  if (C.phone) setAll("[data-phone]", function (el) {
    el.textContent = C.phone;
    el.href = "tel:" + C.phone.replace(/[^\d+]/g, "");
  });

  // ---------- карты ----------
  var tpl = document.getElementById("cardFaces");
  var cardsBox = document.querySelector(".cards");
  var cards = Array.prototype.slice.call(document.querySelectorAll(".card"));
  var opened = 0;
  var openedEl = document.getElementById("opened");
  var statusEl = document.getElementById("status");
  var form = document.getElementById("claim");
  var done = document.getElementById("done");

  cards.forEach(function (card, i) {
    card.appendChild(tpl.content.cloneNode(true));
    card.querySelector("[data-prize]").textContent = C.prize || "−20%";
    card.addEventListener("click", function () {
      if (card.classList.contains("is-flipped")) return;
      card.classList.add("is-flipped");
      card.disabled = true;
      card.setAttribute("aria-label", "Карта " + (i + 1) + ": " + (C.prize || ""));
      opened++;
      openedEl.textContent = opened;
      if (opened === cards.length) win();
    });
  });

  function win() {
    cardsBox.classList.add("win");
    setTimeout(function () {
      statusEl.hidden = true;
      form.hidden = false;
      document.getElementById("email").focus({ preventScroll: true });
      form.scrollIntoView({ behavior: "smooth", block: "nearest" });
    }, 900);
  }

  function flipAll() {
    cards.forEach(function (c) { c.classList.add("is-flipped"); c.disabled = true; });
    opened = cards.length;
    statusEl.hidden = true;
  }

  function showDone(email) {
    flipAll();
    form.hidden = true;
    done.hidden = false;
    document.getElementById("doneEmail").textContent = email;
    if (C.validUntil) {
      document.getElementById("validLine").textContent =
        " Промокод действует до " + C.validUntil + " на euromebel.kz и в салонах EuroMebel.";
    }
  }

  // Уже участвовал с этого устройства — сразу показываем финал
  try {
    var saved = localStorage.getItem(STORE_KEY);
    if (saved) showDone(saved);
  } catch (e) {}

  // ---------- отправка email ----------
  var errEl = document.getElementById("error");
  var btn = document.getElementById("submitBtn");
  var EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/;

  function showError(msg) { errEl.textContent = msg; errEl.hidden = false; }

  function utm() {
    var p = new URLSearchParams(location.search), out = {};
    ["utm_source", "utm_medium", "utm_campaign", "utm_content", "utm_term", "gclid", "fbclid"].forEach(function (k) {
      if (p.get(k)) out[k] = p.get(k);
    });
    return out;
  }

  form.addEventListener("submit", function (e) {
    e.preventDefault();
    errEl.hidden = true;
    var email = form.email.value.trim().toLowerCase();
    if (!EMAIL_RE.test(email)) { showError("Проверьте адрес почты"); form.email.focus(); return; }
    if (form.company.value) { showDone(email); return; } // бот

    if (!C.apiUrl) {
      console.warn("[promo] apiUrl не задан в config.js — демо-режим, email не сохранён");
      showDone(email);
      return;
    }

    btn.disabled = true;
    btn.textContent = "Отправляем…";
    fetch(C.apiUrl, {
      method: "POST",
      // text/plain — без CORS-preflight, так Apps Script принимает запрос
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify({
        action: "claim",
        email: email,
        campaign: C.campaign,
        page: location.href.split("?")[0],
        utm: utm(),
        company: form.company.value
      })
    })
      .then(function (r) { return r.json(); })
      .then(function (res) {
        if (!res || !res.ok) throw new Error((res && res.error) || "server");
        try { localStorage.setItem(STORE_KEY, email); } catch (e) {}
        showDone(email);
      })
      .catch(function (err) {
        var map = {
          invalid_email: "Проверьте адрес почты",
          no_codes: "Промокоды закончились — следите за новыми акциями!",
          rate_limited: "Слишком много попыток, попробуйте через минуту"
        };
        showError(map[err.message] || "Не получилось отправить. Попробуйте ещё раз.");
      })
      .finally(function () {
        btn.disabled = false;
        btn.textContent = "Получить промокод";
      });
  });
})();
