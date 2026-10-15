/**
 * custom-select.js — ersetzt die native Options-Liste durch ein Glass-Dropdown.
 *
 * Das <select> bleibt im DOM und die alleinige Quelle der Wahrheit: bestehender
 * Code liest weiter .value/.options und hoert auf "change". Die Liste wird bei
 * jedem Oeffnen frisch aus den Optionen gerendert, damit dynamisch nachgeladene
 * Projekte ohne Registrierung mitkommen.
 */
(() => {
  "use strict";

  const CHEVRON = `<svg class="cs-chevron" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M6 9l6 6 6-6" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"/></svg>`;
  const CHECK = `<svg class="cs-check" viewBox="0 0 24 24" fill="none" aria-hidden="true"><path d="M5 12.5l4.5 4.5L19 7.5" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"/></svg>`;

  let openInstance = null;
  let idCounter = 0;

  class CustomSelect {
    constructor(select) {
      this.select = select;
      this.id = `cs-${++idCounter}`;

      this.wrapper = document.createElement("div");
      this.wrapper.className = "cs";

      this.trigger = document.createElement("button");
      this.trigger.type = "button";
      this.trigger.className = "cs-trigger";
      this.trigger.setAttribute("aria-haspopup", "listbox");
      this.trigger.setAttribute("aria-expanded", "false");
      this.trigger.innerHTML = `<span class="cs-value"></span>${CHEVRON}`;

      this.menu = document.createElement("div");
      this.menu.className = "cs-menu";
      this.menu.id = `${this.id}-menu`;
      this.menu.setAttribute("role", "listbox");
      this.menu.hidden = true;

      select.parentNode.insertBefore(this.wrapper, select);
      this.wrapper.appendChild(select);
      this.wrapper.appendChild(this.trigger);
      document.body.appendChild(this.menu);

      select.classList.add("cs-native");
      select.setAttribute("tabindex", "-1");
      select.setAttribute("aria-hidden", "true");

      this.bind();
      this.syncTrigger();
    }

    bind() {
      this.trigger.addEventListener("click", () => this.toggle());
      this.trigger.addEventListener("keydown", (e) => this.onTriggerKey(e));
      this.menu.addEventListener("keydown", (e) => this.onMenuKey(e));

      this.menu.addEventListener("click", (e) => {
        const option = e.target.closest(".cs-option");
        if (option && !option.classList.contains("is-disabled")) {
          this.commit(option.dataset.value);
        }
      });

      // Optionen werden per innerHTML ersetzt; disabled wird per Attribut gesetzt.
      new MutationObserver(() => {
        this.syncTrigger();
        if (this.isOpen) this.render();
      }).observe(this.select, { childList: true, subtree: true, attributes: true });

      this.select.addEventListener("change", () => this.syncTrigger());

      // .value = x loest kein change-Event aus, wird aber zum Sync der
      // Projekt-Auswahl zwischen den Tabs genutzt.
      const proto = Object.getPrototypeOf(this.select);
      const nativeValue = Object.getOwnPropertyDescriptor(proto, "value");
      Object.defineProperty(this.select, "value", {
        configurable: true,
        get: () => nativeValue.get.call(this.select),
        set: (v) => {
          nativeValue.set.call(this.select, v);
          this.syncTrigger();
        }
      });
    }

    get isOpen() {
      return !this.menu.hidden;
    }

    syncTrigger() {
      const selected = this.select.options[this.select.selectedIndex];
      const label = this.wrapper.querySelector(".cs-value");
      label.textContent = selected ? selected.text.trim() : "";
      label.classList.toggle("is-placeholder", !selected);
      this.trigger.disabled = this.select.disabled;
      this.wrapper.classList.toggle("is-disabled", this.select.disabled);
    }

    render() {
      const options = Array.from(this.select.options);
      this.menu.innerHTML = options.map((opt, i) => `
        <div class="cs-option${opt.selected ? " is-selected" : ""}${opt.disabled ? " is-disabled" : ""}"
             role="option"
             aria-selected="${opt.selected}"
             data-value="${escapeAttr(opt.value)}"
             data-index="${i}"
             tabindex="-1">
          <span class="cs-option-label">${escapeHtml(opt.text.trim())}</span>
          ${CHECK}
        </div>
      `).join("") || `<div class="cs-empty">Keine Einträge</div>`;
      this.position();
    }

    position() {
      const rect = this.trigger.getBoundingClientRect();
      const gap = 8;
      const margin = 12;
      this.menu.style.minWidth = `${rect.width}px`;
      this.menu.style.left = `${rect.left}px`;

      // Erst messen, dann entscheiden ob nach unten oder oben geklappt wird.
      this.menu.style.top = "0px";
      this.menu.style.maxHeight = "";
      const height = this.menu.offsetHeight;
      const below = window.innerHeight - rect.bottom - gap - margin;
      const above = rect.top - gap - margin;
      const flip = height > below && above > below;

      this.menu.classList.toggle("is-above", flip);
      this.menu.style.maxHeight = `${Math.max(120, flip ? above : below)}px`;
      this.menu.style.top = flip
        ? `${Math.max(margin, rect.top - gap - Math.min(height, above))}px`
        : `${rect.bottom + gap}px`;

      const overflowRight = rect.left + this.menu.offsetWidth - (window.innerWidth - margin);
      if (overflowRight > 0) {
        this.menu.style.left = `${Math.max(margin, rect.left - overflowRight)}px`;
      }
    }

    open() {
      if (this.isOpen || this.select.disabled) return;
      if (openInstance) openInstance.close();
      openInstance = this;

      this.menu.hidden = false;
      this.render();
      this.trigger.setAttribute("aria-expanded", "true");
      this.trigger.setAttribute("aria-controls", this.menu.id);
      this.wrapper.classList.add("is-open");
      requestAnimationFrame(() => this.menu.classList.add("is-visible"));

      const active = this.menu.querySelector(".cs-option.is-selected") || this.menu.querySelector(".cs-option:not(.is-disabled)");
      if (active) {
        active.focus();
        active.scrollIntoView({ block: "nearest" });
      }
    }

    close({ focusTrigger = false } = {}) {
      if (!this.isOpen) return;
      this.menu.classList.remove("is-visible");
      this.trigger.setAttribute("aria-expanded", "false");
      this.wrapper.classList.remove("is-open");
      if (openInstance === this) openInstance = null;

      const hide = () => { this.menu.hidden = true; };
      this.menu.addEventListener("transitionend", hide, { once: true });
      setTimeout(hide, 250);

      if (focusTrigger) this.trigger.focus();
    }

    toggle() {
      this.isOpen ? this.close({ focusTrigger: true }) : this.open();
    }

    commit(value) {
      const changed = this.select.value !== value;
      this.select.value = value;
      this.syncTrigger();
      this.close({ focusTrigger: true });
      if (changed) {
        this.select.dispatchEvent(new Event("input", { bubbles: true }));
        this.select.dispatchEvent(new Event("change", { bubbles: true }));
      }
    }

    onTriggerKey(e) {
      if (["ArrowDown", "ArrowUp", "Enter", " "].includes(e.key)) {
        e.preventDefault();
        this.open();
      }
    }

    onMenuKey(e) {
      const items = Array.from(this.menu.querySelectorAll(".cs-option:not(.is-disabled)"));
      if (!items.length) return;
      const current = items.indexOf(document.activeElement);

      const focusAt = (i) => {
        const next = items[Math.max(0, Math.min(items.length - 1, i))];
        next.focus();
        next.scrollIntoView({ block: "nearest" });
      };

      switch (e.key) {
        case "ArrowDown": e.preventDefault(); focusAt(current + 1); break;
        case "ArrowUp": e.preventDefault(); focusAt(current - 1); break;
        case "Home": e.preventDefault(); focusAt(0); break;
        case "End": e.preventDefault(); focusAt(items.length - 1); break;
        case "Enter":
        case " ":
          e.preventDefault();
          if (current >= 0) this.commit(items[current].dataset.value);
          break;
        case "Escape":
          e.preventDefault();
          this.close({ focusTrigger: true });
          break;
        case "Tab":
          this.close();
          break;
      }
    }
  }

  function escapeHtml(str) {
    const div = document.createElement("div");
    div.textContent = str;
    return div.innerHTML;
  }

  function escapeAttr(str) {
    return escapeHtml(str).replace(/"/g, "&quot;");
  }

  function upgrade(root = document) {
    root.querySelectorAll("select:not(.cs-native)").forEach((select) => {
      if (select.multiple || select.size > 1) return;
      new CustomSelect(select);
    });
  }

  document.addEventListener("click", (e) => {
    if (!openInstance) return;
    if (openInstance.menu.contains(e.target) || openInstance.wrapper.contains(e.target)) return;
    openInstance.close();
  });

  document.addEventListener("focusin", (e) => {
    if (!openInstance) return;
    if (openInstance.menu.contains(e.target) || openInstance.wrapper.contains(e.target)) return;
    openInstance.close();
  });

  // Wie beim nativen Select: Scrollen schliesst die Liste, statt sie loszuloesen.
  window.addEventListener("scroll", () => openInstance?.close(), true);
  window.addEventListener("resize", () => openInstance?.close());

  function init() {
    upgrade();
    new MutationObserver((records) => {
      for (const record of records) {
        for (const node of record.addedNodes) {
          if (node.nodeType !== 1) continue;
          if (node.matches("select")) upgrade(node.parentNode);
          else upgrade(node);
        }
      }
    }).observe(document.body, { childList: true, subtree: true });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
