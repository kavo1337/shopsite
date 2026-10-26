(function () {
  const data = window.CATALOG;
  const catsEl = document.getElementById("cats");
  const subsEl = document.getElementById("subs");
  const catalogEl = document.getElementById("catalog");
  const searchEl = document.getElementById("search");
  const statusEl = document.getElementById("status");

  let activeCat = data.categories[0] ? data.categories[0].id : "";

  function formatPrice(price) {
    if (price === null || price === undefined || price === "") {
      return { text: "по запросу", ask: true };
    }
    return { text: Number(price).toLocaleString("ru-RU") + " ₽", ask: false };
  }

  function cardHtml(item) {
    const price = formatPrice(item.price);
    const photo = item.photo
      ? `<img src="${item.photo}" alt="">`
      : `<span class="empty">нет фото</span>`;
    const bits = [];
    if (item.article) bits.push(item.article);
    if (item.color) bits.push(item.color);
    if (item.material) bits.push(item.material);
    const meta = bits.join(" · ");
    return `<article class="card" data-text="${escapeAttr(
      (item.name + " " + meta).toLowerCase()
    )}">
      <div class="photo">${photo}</div>
      <div class="body">
        <div class="name">${escapeHtml(item.name)}</div>
        ${meta ? `<div class="meta">${escapeHtml(meta)}</div>` : ""}
        <div class="price${price.ask ? " ask" : ""}">${price.text}</div>
      </div>
    </article>`;
  }

  function escapeHtml(s) {
    return String(s || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function escapeAttr(s) {
    return escapeHtml(s).replace(/'/g, "&#39;");
  }

  function renderCats() {
    catsEl.innerHTML = data.categories
      .map(
        (c) =>
          `<button type="button" data-id="${c.id}" class="${
            c.id === activeCat ? "active" : ""
          }">${escapeHtml(c.title)}</button>`
      )
      .join("");
  }

  function currentCategory() {
    return data.categories.find((c) => c.id === activeCat) || data.categories[0];
  }

  function renderSubs(query) {
    const cat = currentCategory();
    const showAll = Boolean(query);
    if (showAll) {
      subsEl.innerHTML = "";
      return;
    }
    const useful = (cat.sections || []).filter((s) => s.title && s.title !== cat.title);
    if (useful.length < 2) {
      subsEl.innerHTML = "";
      return;
    }
    const hash = decodeURIComponent((location.hash || "").replace(/^#/, ""));
    subsEl.innerHTML = useful
      .map(
        (s) =>
          `<a href="#${s.id}" class="${s.id === hash ? "active" : ""}">${escapeHtml(
            s.title
          )}</a>`
      )
      .join("");
  }

  function renderCatalog(query) {
    const q = (query || "").trim().toLowerCase();
    let html = "";
    let shown = 0;
    const cats = q ? data.categories : [currentCategory()];

    cats.forEach((cat) => {
      (cat.sections || []).forEach((section) => {
        const items = (section.items || []).filter((item) => {
          if (!q) return true;
          const hay = [item.name, item.article, item.color, item.material]
            .join(" ")
            .toLowerCase();
          return hay.includes(q);
        });
        if (!items.length) return;
        shown += items.length;
        const title = q ? cat.title + " · " + section.title : section.title;
        html += `<section class="section" id="${section.id}">
          <h2>${escapeHtml(title)}</h2>
          ${section.note ? `<p class="note">${escapeHtml(section.note)}</p>` : ""}
          <div class="grid">${items.map(cardHtml).join("")}</div>
        </section>`;
      });
    });

    catalogEl.innerHTML = html || "<p class='status'>Ничего не найдено.</p>";
    const total = data.categories.reduce(
      (n, c) => n + c.sections.reduce((m, s) => m + s.items.length, 0),
      0
    );
    statusEl.textContent = q
      ? `Найдено: ${shown} из ${total}`
      : `${currentCategory().title}: ${shown} позиций`;
  }

  function scrollActiveChip(nav) {
    const active = nav.querySelector(".active");
    if (!active || typeof active.scrollIntoView !== "function") return;
    active.scrollIntoView({ inline: "center", block: "nearest", behavior: "smooth" });
  }

  function paint() {
    const q = searchEl.value;
    renderCats();
    renderSubs(q);
    renderCatalog(q);
    scrollActiveChip(catsEl);
    scrollActiveChip(subsEl);
  }

  catsEl.addEventListener("click", (e) => {
    const btn = e.target.closest("button");
    if (!btn) return;
    activeCat = btn.dataset.id;
    searchEl.value = "";
    if (history.replaceState) history.replaceState(null, "", location.pathname);
    paint();
    window.scrollTo({ top: 0, behavior: "smooth" });
  });

  subsEl.addEventListener("click", (e) => {
    const link = e.target.closest("a");
    if (!link) return;
    subsEl.querySelectorAll("a").forEach((a) => a.classList.toggle("active", a === link));
    scrollActiveChip(subsEl);
  });

  searchEl.addEventListener("input", () => paint());

  paint();
})();
