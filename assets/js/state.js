let state = {
  orderData: null,
  incomeData: null,
  costData: [],
  costSearchQueries: [],
  costCurrentPage: 1,
  costItemsPerPage: 20,
  results: [],
  summary: [],
  filterMode: 'all',
  fileOrderLoaded: false,
  fileIncomeLoaded: false,
  editingId: null,
  shopStats: null,
  currentUser: null,
  currentPage: 1,
  itemsPerPage: 100,
  resultSearchQuery: '',
  summaryCurrentPage: 1,
  summaryItemsPerPage: 100,
  summarySearchQuery: '',
  adsData: [],
  adsCurrentPage: 1,
  adsItemsPerPage: 50,
  adsSearchQuery: '',
  charts: {},
  sort: {
    cost:    { col: null, dir: 'asc' },
    result:  { col: null, dir: 'asc' },
    summary: { col: null, dir: 'asc' },
    ads:     { col: null, dir: 'asc' }
  },
  incomeOverrides: JSON.parse(localStorage.getItem('shopee_income_overrides') || '{}'),
  theme: localStorage.getItem('torque_theme') || 'light',
};

// ─── Column Sort ──────────────────────────────────────────────
function sortTable(table, col) {
  const s = state.sort[table];
  if (s.col === col) {
    s.dir = s.dir === 'asc' ? 'desc' : 'asc';
  } else {
    s.col = col;
    s.dir = 'asc';
  }

  // Update header visual
  document.querySelectorAll(`th[id^="${table}-th-"]`).forEach(th => {
    th.classList.remove('sort-asc', 'sort-desc');
    const icon = th.querySelector('.sort-icon');
    if (icon) icon.textContent = '⇅';
  });
  const activeTh = document.getElementById(`${table}-th-${col}`);
  if (activeTh) {
    activeTh.classList.add(s.dir === 'asc' ? 'sort-asc' : 'sort-desc');
    const icon = activeTh.querySelector('.sort-icon');
    if (icon) icon.textContent = s.dir === 'asc' ? '▲' : '▼';
  }

  // Re-render the correct table
  if (table === 'cost')    { state.costCurrentPage = 1; renderCostTable(); }
  if (table === 'result')  { state.currentPage = 1; renderResultTable(); }
  if (table === 'summary') { state.summaryCurrentPage = 1; renderSummaryTable(); }
  if (table === 'ads')     { state.adsCurrentPage = 1; renderAdsTable(); }
}
