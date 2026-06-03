/* Lightweight search / sort / pagination for any <table class="sortable">. */
(function () {
  const PAGE_SIZE = 25;

  function cellText(row, index) {
    return (row.cells[index]?.textContent || "").trim();
  }
  function comparable(value) {
    const normalized = value.replace(/,/g, "");
    const number = Number(normalized);
    if (!Number.isNaN(number) && normalized !== "") return number;
    const date = Date.parse(value);
    if (!Number.isNaN(date)) return date;
    return value.toLowerCase();
  }

  function enhance(table) {
    const tbody = table.tBodies[0];
    if (!tbody) return;
    const rows = Array.from(tbody.rows).filter(r => !r.querySelector(".empty"));
    if (rows.length <= 1) return;

    let page = 1, sortIndex = null, sortDir = 1, query = "";
    const controls = document.createElement("div");
    controls.className = "table-tools";
    controls.innerHTML = `
      <label class="table-search"><span>Search</span>
        <input type="search" placeholder="Filter table..." aria-label="Search table"></label>
      <div class="table-pager">
        <button type="button" data-action="prev">Prev</button>
        <span data-role="page"></span>
        <button type="button" data-action="next">Next</button>
      </div>`;
    table.parentNode.insertBefore(controls, table);
    const search = controls.querySelector("input");
    const pageLabel = controls.querySelector("[data-role='page']");
    const prev = controls.querySelector("[data-action='prev']");
    const next = controls.querySelector("[data-action='next']");

    function filteredRows() {
      let out = rows;
      if (query) out = out.filter(r => r.textContent.toLowerCase().includes(query));
      if (sortIndex !== null) {
        out = out.slice().sort((a, b) => {
          const av = comparable(cellText(a, sortIndex)), bv = comparable(cellText(b, sortIndex));
          if (av < bv) return -1 * sortDir;
          if (av > bv) return 1 * sortDir;
          return 0;
        });
      }
      return out;
    }
    function render() {
      const out = filteredRows();
      const totalPages = Math.max(1, Math.ceil(out.length / PAGE_SIZE));
      page = Math.min(Math.max(1, page), totalPages);
      const start = (page - 1) * PAGE_SIZE;
      const visible = new Set(out.slice(start, start + PAGE_SIZE));
      rows.forEach(r => { r.hidden = !visible.has(r); });
      pageLabel.textContent = `${page} / ${totalPages} (${out.length} rows)`;
      prev.disabled = page <= 1; next.disabled = page >= totalPages;
    }
    search.addEventListener("input", () => { query = search.value.trim().toLowerCase(); page = 1; render(); });
    prev.addEventListener("click", () => { page -= 1; render(); });
    next.addEventListener("click", () => { page += 1; render(); });

    Array.from(table.tHead?.rows[0]?.cells || []).forEach((th, index) => {
      th.classList.add("sortable-head");
      th.tabIndex = 0; th.title = "Sort";
      function sort() {
        if (sortIndex === index) sortDir *= -1; else { sortIndex = index; sortDir = 1; }
        Array.from(table.tHead.rows[0].cells).forEach(h => h.removeAttribute("aria-sort"));
        th.setAttribute("aria-sort", sortDir === 1 ? "ascending" : "descending");
        page = 1; render();
      }
      th.addEventListener("click", sort);
      th.addEventListener("keydown", e => { if (e.key === "Enter" || e.key === " ") { e.preventDefault(); sort(); } });
    });
    render();
  }

  document.addEventListener("DOMContentLoaded", () => {
    document.querySelectorAll("table.sortable").forEach(enhance);
  });
})();
