function changeResultPerPage(val) {
  state.itemsPerPage = parseInt(val, 10);
  state.currentPage = 1;
  renderResultTable();
}

function filterOrders(type) {
  state.filterMode = type;
  state.currentPage = 1;
  
  // Update UI button state
  document.querySelectorAll('.flex-gap .btn').forEach(b => {
    b.classList.remove('active');
    b.style.borderColor = 'var(--border)';
    b.style.background = 'var(--surface)';
    b.style.color = 'var(--text)';
  });
  
  const btn = document.querySelector(`.flex-gap .btn[onclick*="'${type}'"]`);
  if (btn) {
    btn.classList.add('active');
    btn.style.borderColor = 'var(--primary)';
    btn.style.background = 'var(--primary-light)';
    btn.style.color = 'var(--primary)';
  }
  
  renderResultTable();
}

function filterResultTable(){
  state.resultSearchQuery = document.getElementById('result-search').value.trim();
  state.currentPage = 1;
  renderResultTable();
}

function renderResultTable(){
  const tbody = document.getElementById('result-tbody');
  if(!tbody) return;
  tbody.innerHTML = '';
  
  let data = state.results;
  if(state.filterMode==='no-cost') data = data.filter(r => r.missingCost && !r.isCancelled);
  else if(state.filterMode==='success') data = data.filter(r => !r.isCancelled && !r.missingCost);
  else if(state.filterMode==='cancelled') data = data.filter(r => r.isCancelled);
  else if(state.filterMode==='shipping') data = data.filter(r => r.status.includes('การจัดส่ง'));
  else if(state.filterMode==='toship') data = data.filter(r => r.status.includes('ที่ต้องจัดส่ง'));

  if(state.resultSearchQuery) {
    const q = state.resultSearchQuery.toLowerCase();
    // Find all order IDs that have at least one item matching the search
    const matchingOrderIds = new Set(
      state.results.filter(r => {
        const fullText = `${r.orderId} ${r.product} ${r.variant} ${r.sku} ${r.status}`.toLowerCase();
        return fullText.includes(q);
      }).map(r => r.orderId)
    );
    // Filter the current data (which might already have mode filters) to only these orders
    data = data.filter(r => matchingOrderIds.has(r.orderId));
  }

  // Pagination logic based on unique orders
  // Sort by order-level value — need unique order list first
  const rs = state.sort.result;
  let allUniqueOrders;
  if (rs.col) {
    const dir = rs.dir === 'asc' ? 1 : -1;
    const orderMap = {};
    data.filter(r => r.isFirst).forEach(r => {
      orderMap[r.orderId] = r;
    });
    const sortVal = (r) => {
      if (rs.col === 'orderId') return (r.orderId||'').toLowerCase();
      if (rs.col === 'feePct') {
        const sp = r.orderSellingPrice || r.orderSalePrice || 1;
        return (Math.abs(r.commFee||0)+Math.abs(r.servFee||0)+Math.abs(r.transFee||0))/sp;
      }
      if (rs.col === 'product') return (r.product||'').toLowerCase();
      if (rs.col === 'qty') return r.qty||0;
      if (rs.col === 'salePrice') return r.salePrice||0;
      if (rs.col === 'shopeeFee') return Math.max(0,(r.orderSalePrice||0)-(r.income||0));
      if (rs.col === 'income') return r.income||0;
      if (rs.col === 'net') return r.net||0;
      return 0;
    };
    const orderedIds = Object.keys(orderMap).sort((a, b) => {
      const va = sortVal(orderMap[a]);
      const vb = sortVal(orderMap[b]);
      if (typeof va === 'string') return va.localeCompare(vb, 'th') * dir;
      return (va - vb) * dir;
    });
    allUniqueOrders = orderedIds;
  } else {
    allUniqueOrders = [...new Set(data.map(r => r.orderId))];
  }
  const totalOrders = allUniqueOrders.length;
  const totalAvailableOrders = new Set(state.results.map(r => r.orderId)).size;
  const resultCountEl = document.getElementById('result-count');
  if (resultCountEl) {
    if (totalOrders === totalAvailableOrders) {
      resultCountEl.innerText = `ทั้งหมด ${totalOrders.toLocaleString()} รายการ`;
    } else {
      resultCountEl.innerText = `พบ ${totalOrders.toLocaleString()} จาก ${totalAvailableOrders.toLocaleString()} รายการ`;
    }
  }
  const totalPages = Math.ceil(totalOrders / state.itemsPerPage) || 1;
  
  if (state.currentPage > totalPages) state.currentPage = totalPages;
  if (state.currentPage < 1) state.currentPage = 1;
  
  const startIdx = (state.currentPage - 1) * state.itemsPerPage;
  const pageOrders = new Set(allUniqueOrders.slice(startIdx, startIdx + state.itemsPerPage));
  
  const pageData = data.filter(r => pageOrders.has(r.orderId))
    .sort((a, b) => {
      // Keep original item order within same orderId, matching sorted order list
      const ai = allUniqueOrders.indexOf(a.orderId);
      const bi = allUniqueOrders.indexOf(b.orderId);
      if (ai !== bi) return ai - bi;
      return 0;
    });

  if (pageData.length === 0) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="12" style="text-align:center; padding:32px; color:var(--text-muted); font-size: 14px; background:#fbfbfb;">🔍 ไม่พบข้อมูลที่ค้นหา</td>`;
    tbody.appendChild(tr);
  }

  pageData.forEach(r=>{
    const netStr = r.net!==null ? '฿'+Math.round(r.net).toLocaleString() : '—';
    const netStyle = r.net!==null ? (r.net>=0?'color:var(--green);font-weight:700':'color:var(--red);font-weight:700') : 'color:var(--text-muted)';
    const badge = r.isCancelled 
      ? '<span class="badge red">ยกเลิก</span>' 
      : r.missingCost 
        ? '<span class="badge red">ขาด cost</span>' 
        : '<span class="badge green">OK</span>';
    
    const tr = document.createElement('tr');
    let html = '';

    // --- Payment channel color mapping ---
    const payColorMap = {
      'SPayLater': { bg: '#fff3e0', color: '#e65100', label: 'SPayLater' },
      'Mobile Banking': { bg: '#e3f2fd', color: '#1565c0', label: 'Mobile Banking' },
      'QR': { bg: '#f3e5f5', color: '#6a1b9a', label: 'QR พร้อมเพย์' },
      'บัตรเครดิต': { bg: '#f1f8e9', color: '#388e3c', label: 'บัตรเครดิต/เดบิต' },
      'ShopeePay': { bg: '#fff8e1', color: '#f57f17', label: 'ShopeePay' },
      'ยอดเงิน ShopeePay': { bg: '#fff8e1', color: '#f57f17', label: 'ShopeePay' },
      'เก็บเงินปลายทาง': { bg: '#e8f5e9', color: '#2e7d32', label: 'COD' },
    };
    const payKey2 = Object.keys(payColorMap).find(k => (r.paymentChannel||'').includes(k)) || '';
    const payStyle = payKey2 ? payColorMap[payKey2] : { bg: '#f5f5f5', color: '#555', label: r.paymentChannel || '-' };
    const payLabel = payStyle.label.length > 14 ? payStyle.label.substring(0,13)+'…' : payStyle.label;
    const payBadge = `<span style="display:inline-block;font-size:10px;font-weight:600;padding:3px 8px;border-radius:20px;background:${payStyle.bg};color:${payStyle.color};white-space:nowrap;">${payLabel}</span>`;

    if(r.isFirst) {
      // --- Fee % calculation ---
      const sp = r.orderSellingPrice || r.orderSalePrice || 1;
      const commPct  = (Math.abs(r.commFee  || 0) / sp * 100);
      const servPct  = (Math.abs(r.servFee  || 0) / sp * 100);
      const transPct = (Math.abs(r.transFee || 0) / sp * 100);
      const totalPct = commPct + servPct + transPct;

      const pctParts = [];
      if (commPct > 0) pctParts.push(`<span title="ค่าคอมมิชชั่น">คอม ${commPct.toFixed(2)}%</span>`);
      if (servPct > 0) pctParts.push(`<span title="ค่าบริการ">บริการ ${servPct.toFixed(2)}%</span>`);
      if (transPct > 0) pctParts.push(`<span title="ค่าธุรกรรม">ธุรกรรม ${transPct.toFixed(2)}%</span>`);

      const pctBreakdown = pctParts.length > 0
        ? `<div style="margin-top:4px;display:flex;flex-wrap:wrap;gap:3px;">${
            pctParts.map(p => `<span style="font-size:9px;padding:1px 5px;border-radius:10px;background:#fef2f2;color:#ef4444;">${p}</span>`).join('')
          }</div>`
        : '';

      html += `<td rowspan="${r.rowSpan}" style="padding:10px 12px;vertical-align:middle;" class="border-right">
        <div style="font-family:monospace;font-size:11px;color:var(--text-muted);letter-spacing:0.5px;">${r.orderId}</div>
      </td>`;
      html += `<td rowspan="${r.rowSpan}" style="padding:8px;vertical-align:middle;text-align:center;">${badge}</td>`;
      html += `<td rowspan="${r.rowSpan}" style="padding:8px 10px;vertical-align:middle;text-align:center;">${payBadge}</td>`;
      html += `<td rowspan="${r.rowSpan}" style="padding:8px 10px;vertical-align:middle;text-align:center;">
        <div style="font-size:13px;font-weight:700;color:var(--red);">${totalPct > 0 ? totalPct.toFixed(2)+'%' : (r.feePct || '-')}</div>
        ${pctBreakdown}
      </td>`;
    }

    // Product + SKU cell
    const skuChip = r.sku ? `<span style="display:inline-block;margin-top:2px;font-size:9px;font-weight:700;padding:1px 6px;background:#eff6ff;color:#1d4ed8;border-radius:4px;letter-spacing:0.5px;">${r.sku}</span>` : '';
    html += `
        <td style="padding:6px 8px;vertical-align:middle;">
          <div style="font-size:11px;font-weight:500;color:var(--text);max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${r.product}">${r.product}</div>
          ${skuChip}
        </td>
        <td style="padding:6px 8px;vertical-align:middle;font-size:11px;color:var(--text-muted);max-width:100px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${r.variant}">${r.variant}</td>
        <td style="padding:8px;vertical-align:middle;text-align:center;">
          <span style="display:inline-block;font-size:13px;font-weight:700;min-width:24px;">${r.qty}</span>
        </td>
        <td style="padding:8px 12px;vertical-align:middle;text-align:right;font-size:12px;color:var(--text-muted);">฿${(r.salePrice||0).toLocaleString()}</td>
    `;

    if(r.isFirst) {
      // Shopee fee breakdown - Sum of specific items
      const feeItems = [
        { label: 'ค่าคอมมิชชั่น', val: (r.commFee || 0) + (r.commAMSFee || 0) },
        { label: 'ค่าบริการ', val: r.servFee || 0 },
        { label: 'ค่าธรรมเนียมโครงสร้างพื้นฐาน', val: r.platFee || 0 },
        { label: 'ค่าธุรกรรมการชำระเงิน', val: r.transFee || 0 },
        { label: 'หักค่าจัดส่ง', val: r.shipDeduct || 0 }
      ].filter(f => Math.abs(f.val) > 0.01);

      const totalDeduction = feeItems.reduce((sum, item) => sum + Math.abs(item.val), 0);

      const feeBreakdown = feeItems.length > 0
        ? `<div style="margin-top:6px;border-top:1px dashed #fca5a5;padding-top:4px;">` +
          feeItems.map(f =>
            `<div style="display:flex;justify-content:space-between;align-items:center;font-size:10px;color:#9ca3af;line-height:1.8;gap:8px;">
              <span>${f.label}</span>
              <span style="color:#f87171;font-weight:500;">-฿${Math.abs(f.val).toLocaleString(undefined,{minimumFractionDigits:2, maximumFractionDigits:2})}</span>
            </div>`
          ).join('') + `</div>`
        : '';

      html += `<td rowspan="${r.rowSpan}" style="padding:10px 12px;vertical-align:top;min-width:140px;" class="border-right">
        <div style="display:flex;align-items:center;justify-content:flex-end;gap:4px;">
          <span style="font-size:13px;font-weight:700;color:var(--red);">-฿${totalDeduction.toLocaleString(undefined,{minimumFractionDigits:2, maximumFractionDigits:2})}</span>
        </div>
        ${feeBreakdown}
      </td>`;

      const incomeVal = r.income || 0;
      let incomeHtml = '';
      if (r.isOverridden) {
        incomeHtml = `
          <div style="font-weight:700;color:var(--blue);">฿${Math.round(incomeVal).toLocaleString()}</div>
          <div style="font-size:9px;color:var(--text-muted);">(แก้ไขแล้ว)</div>
        `;
      } else if (r.isEstimated) {
        incomeHtml = `
          <div style="font-weight:700;color:var(--amber);">฿${Math.round(incomeVal).toLocaleString()}</div>
          <div style="font-size:9px;color:var(--amber);font-weight:600;">(ประมาณ - คลิกแก้)</div>
        `;
      } else {
        incomeHtml = `<div style="font-weight:700;">฿${Math.round(incomeVal).toLocaleString()}</div>`;
      }

      html += `<td rowspan="${r.rowSpan}" style="padding:10px 12px;vertical-align:middle;text-align:right;font-size:13px;cursor:pointer;" class="border-right" onclick="editIncomeOverride('${r.orderId}', ${incomeVal})">
        ${incomeHtml}
      </td>`;
    }

    // Cost display
    let costDisplay;
    if (r.itemCostTotal !== null) {
      costDisplay = `<span style="font-size:12px;color:var(--text-muted);">฿${(r.itemCostTotal||0).toLocaleString()}</span>`;
    } else if (!r.isCancelled) {
      const sku2 = (r.sku||'').replace(/'/g,"\\'");
      const prod2 = (r.product||'').replace(/'/g,"\\'");
      const vari2 = (r.variant||'').replace(/'/g,"\\'");
      costDisplay = `<button class="btn sm" style="padding:2px 8px;font-size:10px;border-color:#fca5a5;color:var(--red);background:#fff5f5;" onclick="prefillVariant('${sku2}','${prod2}','${vari2}')">+ เพิ่มต้นทุน</button>`;
    } else {
      costDisplay = '<span style="color:#ccc;">—</span>';
    }

    html += `<td style="padding:8px 12px;vertical-align:middle;text-align:right;">${costDisplay}</td>`;

    if(r.isFirst){
      html += `
        <td rowspan="${r.rowSpan}" style="padding:10px 12px;vertical-align:middle;text-align:right;font-size:14px;${netStyle}" class="border-right">${netStr}</td>
      `;
    }

    tr.innerHTML = html;
    tbody.appendChild(tr);
  });
  
  renderPaginationData(totalPages);
}

function renderPaginationData(totalPages) {
  const container = document.getElementById('result-pagination');
  if(!container) return;
  if(totalPages <= 1) { container.innerHTML = ''; return; }
  
  let html = `<button class="page-btn" ${state.currentPage === 1 ? 'disabled' : ''} onclick="changePage(${state.currentPage - 1})">❮ Prev</button>`;
  
  let start = Math.max(1, state.currentPage - 2);
  let end = Math.min(totalPages, start + 4);
  if(end - start < 4) start = Math.max(1, end - 4);
  
  if(start > 1) {
    html += `<button class="page-btn" onclick="changePage(1)">1</button>`;
    if(start > 2) html += `<span style="color:var(--text-muted);font-weight:600;">...</span>`;
  }
  
  for(let i=start; i<=end; i++){
    html += `<button class="page-btn ${i === state.currentPage ? 'active' : ''}" onclick="changePage(${i})">${i}</button>`;
  }
  
  if(end < totalPages){
    if(end < totalPages - 1) html += `<span style="color:var(--text-muted);font-weight:600;">...</span>`;
    html += `<button class="page-btn" onclick="changePage(${totalPages})">${totalPages}</button>`;
  }
  
  html += `<button class="page-btn" ${state.currentPage === totalPages ? 'disabled' : ''} onclick="changePage(${state.currentPage + 1})">Next ❯</button>`;
  container.innerHTML = html;
}

function changePage(p) {
  state.currentPage = p;
  renderResultTable();
}

function filterSummaryTable(){
  state.summarySearchQuery = document.getElementById('summary-search').value.trim();
  state.summaryCurrentPage = 1;
  renderSummaryTable();
}

function renderSummaryTable(){
  const tbody = document.getElementById('summary-tbody');
  if(!tbody) return;
  tbody.innerHTML = '';
  
  let data = state.summary;
  if(state.summarySearchQuery) {
    const q = state.summarySearchQuery.toLowerCase();
    data = data.filter(r => {
      const fullText = `${r.sku} ${r.product} ${r.variant} ${r.title}`.toLowerCase();
      return fullText.includes(q);
    });
  }

  // Sort
  const ss = state.sort.summary;
  if (ss.col) {
    const dir = ss.dir === 'asc' ? 1 : -1;
    data = [...data].sort((a, b) => {
      if (ss.col === 'sku') return ((a.sku||'').localeCompare(b.sku||'', 'th')) * dir;
      if (ss.col === 'product') return ((a.product||'').localeCompare(b.product||'', 'th')) * dir;
      if (ss.col === 'title') return ((a.title||'').localeCompare(b.title||'', 'th')) * dir;
      if (ss.col === 'qty') return (a.qty - b.qty) * dir;
      if (ss.col === 'revenue') return (a.revenue - b.revenue) * dir;
      if (ss.col === 'cost') return (a.cost - b.cost) * dir;
      if (ss.col === 'profit') {
        const pa = (a.qty > 0 ? a.revenue/a.qty : 0) - (a.qty > 0 ? a.cost/a.qty : 0);
        const pb = (b.qty > 0 ? b.revenue/b.qty : 0) - (b.qty > 0 ? b.cost/b.qty : 0);
        return (pa - pb) * dir;
      }
      return 0;
    });
  }

  const totalItems = data.length;
  const totalAvailableItems = state.summary.length;
  const summaryCountEl = document.getElementById('summary-count');
  if (summaryCountEl) {
    if (totalItems === totalAvailableItems) {
      summaryCountEl.innerText = `ทั้งหมด ${totalItems.toLocaleString()} รายการ`;
    } else {
      summaryCountEl.innerText = `พบ ${totalItems.toLocaleString()} จาก ${totalAvailableItems.toLocaleString()} รายการ`;
    }
  }
  const totalPages = Math.ceil(totalItems / state.summaryItemsPerPage) || 1;
  
  if (state.summaryCurrentPage > totalPages) state.summaryCurrentPage = totalPages;
  if (state.summaryCurrentPage < 1) state.summaryCurrentPage = 1;

  const startIdx = (state.summaryCurrentPage - 1) * state.summaryItemsPerPage;
  const pageData = data.slice(startIdx, startIdx + state.summaryItemsPerPage);

  if (pageData.length === 0) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="8" style="text-align:center; padding:32px; color:var(--text-muted); font-size: 14px; background:#fbfbfb;">🔍 ไม่พบข้อมูลที่ค้นหา</td>`;
    tbody.appendChild(tr);
  }

  pageData.forEach(r => {
    const avgRev = r.qty > 0 ? r.revenue / r.qty : 0;
    const avgCost = r.qty > 0 ? r.cost / r.qty : 0;
    const avgNet = avgRev - avgCost;
    
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="font-size:12px;color:var(--blue);font-family:monospace;font-weight:600">${r.sku || '-'}</td>
      <td style="font-size:13px;font-weight:500;">${r.product}</td>
      <td style="font-size:12px;color:var(--text-muted);">${r.variant || '-'}</td>
      <td style="text-align:right;font-weight:600">${r.qty}</td>
      <td style="text-align:right">฿${Math.round(r.revenue).toLocaleString()}</td>
      <td style="text-align:right">฿${Math.round(r.cost).toLocaleString()}</td>
      <td style="text-align:right;color:var(--text-muted)">฿${Math.round(avgCost).toLocaleString()}</td>
      <td style="text-align:right;color:var(--blue);font-weight:500">฿${Math.round(avgRev).toLocaleString()}</td>
      <td style="text-align:right;font-weight:600;color:${avgNet>=0?'var(--green)':'var(--red)'}">฿${Math.round(avgNet).toLocaleString()}</td>
    `;
    tbody.appendChild(tr);
  });
  
  renderSummaryPagination(totalPages);
}

function renderSummaryPagination(totalPages) {
  const container = document.getElementById('summary-pagination');
  if(!container) return;
  if(totalPages <= 1) { container.innerHTML = ''; return; }
  
  let html = `<button class="page-btn" ${state.summaryCurrentPage === 1 ? 'disabled' : ''} onclick="changeSummaryPage(${state.summaryCurrentPage - 1})">❮ Prev</button>`;
  
  let start = Math.max(1, state.summaryCurrentPage - 2);
  let end = Math.min(totalPages, start + 4);
  if(end - start < 4) start = Math.max(1, end - 4);
  
  if(start > 1) {
    html += `<button class="page-btn" onclick="changeSummaryPage(1)">1</button>`;
    if(start > 2) html += `<span style="color:var(--text-muted);font-weight:600;">...</span>`;
  }
  
  for(let i=start; i<=end; i++){
    html += `<button class="page-btn ${i === state.summaryCurrentPage ? 'active' : ''}" onclick="changeSummaryPage(${i})">${i}</button>`;
  }
  
  if(end < totalPages){
    if(end < totalPages - 1) html += `<span style="color:var(--text-muted);font-weight:600;">...</span>`;
    html += `<button class="page-btn" onclick="changeSummaryPage(${totalPages})">${totalPages}</button>`;
  }
  
  html += `<button class="page-btn" ${state.summaryCurrentPage === totalPages ? 'disabled' : ''} onclick="changeSummaryPage(${state.summaryCurrentPage + 1})">Next ❯</button>`;
  container.innerHTML = html;
}

function changeSummaryPage(p) {
  state.summaryCurrentPage = p;
  renderSummaryTable();
}

function changeSummaryPerPage(val) {
  state.summaryItemsPerPage = parseInt(val, 10);
  state.summaryCurrentPage = 1;
  renderSummaryTable();
}

function switchTab(name){
  // Sidebar UI Update
  document.querySelectorAll('.nav-item').forEach(btn => btn.classList.remove('active'));
  const activeBtn = document.querySelector(`.nav-item[onclick*="${name}"]`);
  if (activeBtn) activeBtn.classList.add('active');

  // Content Update
  document.querySelectorAll('.tab-content').forEach(tc => tc.classList.remove('active'));
  const targetContent = document.getElementById('tab-' + name);
  if (targetContent) targetContent.classList.add('active');

  // Page Info Update
  const titles = {
    'import': { t: 'นำเข้าไฟล์', s: 'อัพโหลดไฟล์จาก Shopee และเตรียมการคำนวณ', i: 'upload_file' },
    'cost': { t: 'จัดการต้นทุน', s: 'เพิ่มและแก้ไขต้นทุนสินค้าเพื่อคำนวณกำไร', i: 'inventory_2' },
    'result': { t: 'ผลลัพธ์รายออเดอร์', s: 'ตรวจสอบรายละเอียดกำไรรายออเดอร์ที่คำนวณแล้ว', i: 'receipt_long' },
    'summary': { t: 'สรุปรายสินค้า', s: 'วิเคราะห์ภาพรวมกำไรและยอดขายแยกตาม SKU', i: 'analytics' },
    'ads': { t: 'วิเคราะห์โฆษณา', s: 'เจาะลึกประสิทธิภาพ Shopee Ads และความคุ้มค่า', i: 'ads_click' },
    'dashboard': { t: 'Dashboard Insights', s: 'ข้อมูลเชิงกลยุทธ์ผ่าน BCG Matrix และ Pricing Intelligence', i: 'dashboard' }
  };
  
  const info = titles[name] || { t: name, s: '', i: 'star' };
  document.getElementById('current-page-title').innerText = info.t;
  document.getElementById('current-page-sub').innerText = info.s;
  const iconEl = document.getElementById('current-page-icon');
  if (iconEl) iconEl.innerText = info.i;

  // Show/Hide PDF button in Results or Dashboard or Summary
  const pdfBtn = document.getElementById('btn-export-pdf');
  if (['result', 'summary', 'dashboard'].includes(name) && state.results.length > 0) {
    pdfBtn.style.display = 'flex';
  } else {
    pdfBtn.style.display = 'none';
  }
}

function toggleSidebar() {
  const sidebar = document.querySelector('.sidebar');
  const icon = document.querySelector('.sidebar-toggle .material-symbols-rounded');
  if (!sidebar) return;
  
  sidebar.classList.toggle('collapsed');
  
  if (sidebar.classList.contains('collapsed')) {
    icon.innerText = 'side_navigation';
  } else {
    icon.innerText = 'menu_open';
  }
  
  // Refresh charts after transition to ensure they fit the new container width
  setTimeout(() => {
    if (state.charts) {
      Object.values(state.charts).forEach(chart => {
        if (chart && typeof chart.updateOptions === 'function') {
          chart.render();
        }
      });
    }
  }, 350);
}

function toggleTheme() {
  state.theme = state.theme === 'light' ? 'dark' : 'light';
  document.documentElement.setAttribute('data-theme', state.theme);
  localStorage.setItem('torque_theme', state.theme);
  
  const icon = document.querySelector('.theme-toggle .material-symbols-rounded');
  icon.innerText = state.theme === 'dark' ? 'light_mode' : 'dark_mode';
  
  // Refresh charts for theme colors if on dashboard
  if (document.getElementById('tab-dashboard').classList.contains('active')) {
    renderDashboard();
  }
}

function showSkeletons(parentId) {
  const container = document.getElementById(parentId);
  const template = document.getElementById('skeleton-card');
  container.innerHTML = '';
  for (let i = 0; i < 3; i++) {
    container.appendChild(template.content.cloneNode(true));
  }
}

function filterByQuickStats(type, el) {
  // Clear active state from all metrics
  document.querySelectorAll('.metric').forEach(m => m.classList.remove('active'));
  if (el) el.classList.add('active');

  const searchInput = document.getElementById('result-search');
  
  if (type === 'all') {
    state.filterMode = 'all';
    searchInput.value = '';
  } else if (type === 'success') {
    state.filterMode = 'success';
    searchInput.value = '';
  } else if (type === 'cancelled') {
    state.filterMode = 'cancelled';
    searchInput.value = '';
  } else if (type === 'shipping') {
    state.filterMode = 'shipping';
    searchInput.value = '';
  } else if (type === 'toship') {
    state.filterMode = 'toship';
    searchInput.value = '';
  }
  
  state.currentPage = 1;
  filterResultTable();
}

function exportResult(){
  if(!state.results.length){ showWarningMessage('ไม่มีข้อมูล', 'ไม่มีข้อมูลให้ดาวน์โหลด'); return; }
  const rows = state.results.map(r=>({
    'Order ID': r.orderId,
    'สถานะ': r.status,
    'SKU': r.sku,
    'สินค้า': r.product,
    'Variant': r.variant,
    'จำนวน': r.qty,
    'ยอดขาย/ชิ้น': r.salePrice,
    'Income โอน (ต่อออเดอร์)': r.isFirst ? r.income : '',
    'ต้นทุน (รวมตามจำนวน)': r.itemCostTotal,
    'Net กำไร (ต่อออเดอร์)': r.isFirst ? r.net : ''
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Orders');
  XLSX.writeFile(wb, `Torque_Profit_Report_${getFormattedDateStr()}.xlsx`);
}

function exportSummary(){
  if(!state.summary.length){ showWarningMessage('ไม่มีข้อมูล', 'ไม่มีข้อมูลให้ดาวน์โหลด'); return; }
  const rows = state.summary.map(r=>({
    'SKU': r.sku,
    'สินค้า/ตัวเลือก': r.title,
    'จำนวนที่ขายได้ (ชิ้น)': r.qty,
    'รายได้ประมาณ (฿)': Math.round(r.revenue),
    'ต้นทุนรวม (฿)': Math.round(r.cost),
    'เฉลี่ยต้นทุนต่อชิ้น (฿)': Math.round(r.qty>0?r.cost/r.qty:0),
    'เฉลี่ยยอดขายต่อชิ้น (฿)': Math.round(r.qty>0?r.revenue/r.qty:0),
    'เฉลี่ยกำไรต่อชิ้น (฿)': Math.round(r.qty>0?(r.revenue-r.cost)/r.qty:0)
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Summary');
  XLSX.writeFile(wb, `Torque_Summary_Report_${getFormattedDateStr()}.xlsx`);
}

function editIncomeOverride(orderId, currentVal) {
  const result = state.results.find(r => r.orderId === orderId);
  if (result && result.isCancelled) {
    Swal.fire({
      icon: 'warning',
      title: 'ออเดอร์นี้ถูกยกเลิก',
      text: 'ไม่สามารถแก้ไขรายรับของออเดอร์ที่ยกเลิกแล้วได้ครับ',
      confirmButtonColor: 'var(--primary)'
    });
    return;
  }

  Swal.fire({
    title: 'แก้ไขรายรับ (Income)',
    html: `
      <div style="text-align:left; padding:10px;">
        <div style="font-size:12px;color:rgba(0,0,0,0.4);margin-bottom:4px;">หมายเลขคำสั่งซื้อ</div>
        <div style="font-size:16px;font-weight:700;margin-bottom:15px;color:var(--text);">${orderId}</div>
        <div style="font-size:12px;color:rgba(0,0,0,0.4);margin-bottom:8px;">ระบุยอดเงินที่ได้รับจริง (฿)</div>
      </div>
    `,
    input: 'number',
    inputValue: currentVal,
    inputAttributes: { step: '0.01', style: 'border-radius:12px; padding:12px; font-size:18px; font-weight:700;' },
    showCancelButton: true,
    confirmButtonText: 'บันทึกข้อมูล',
    cancelButtonText: 'ยกเลิก',
    confirmButtonColor: '#f97316', // Torque Primary
    cancelButtonColor: '#94a3b8',
    background: '#fff',
    padding: '24px',
    customClass: {
      popup: 'swal2-borderless',
      title: 'swal2-title-main',
      input: 'swal2-input-premium'
    }
  }).then((result) => {
    if (result.isConfirmed) {
      const newVal = parseFloat(result.value);
      if (isNaN(newVal)) return;
      
      state.incomeOverrides[orderId] = newVal;
      localStorage.setItem('shopee_income_overrides', JSON.stringify(state.incomeOverrides));
      
      processFiles(true);
      Swal.fire({
        icon: 'success',
        title: 'อัปเดตสำเร็จ',
        text: `บันทึกรายรับของออเดอร์ ${orderId} แล้ว`,
        timer: 1500,
        showConfirmButton: false
      });
    }
  });
}

function changeAdsPerPage(val) {
  state.adsItemsPerPage = parseInt(val, 10);
  state.adsCurrentPage = 1;
  renderAdsTable();
}

function filterAdsTable() {
  state.adsSearchQuery = document.getElementById('ads-search').value.trim();
  state.adsCurrentPage = 1;
  renderAdsTable();
}

function renderAdsTable() {
  const tbody = document.getElementById('ads-tbody');
  if (!tbody) return;
  tbody.innerHTML = '';
  
  let data = state.adsData;
  if (state.adsSearchQuery) {
    const q = state.adsSearchQuery.toLowerCase();
    data = data.filter(d => 
      (d.name && d.name.toLowerCase().includes(q)) ||
      (d.productId && d.productId.toLowerCase().includes(q))
    );
  }

  // Sort
  const as = state.sort.ads;
  if (as.col) {
    const dir = as.dir === 'asc' ? 1 : -1;
    data = [...data].sort((a, b) => {
      let va = a[as.col], vb = b[as.col];
      if (typeof va === 'string') return va.localeCompare(vb, 'th') * dir;
      return (va - vb) * dir;
    });
  }

  const totalItems = data.length;
  const totalAvailableItems = state.adsData.length;
  const countEl = document.getElementById('ads-count');
  if (countEl) {
    if (totalItems === totalAvailableItems) {
      countEl.innerText = `ทั้งหมด ${totalItems.toLocaleString()} รายการ`;
    } else {
      countEl.innerText = `พบ ${totalItems.toLocaleString()} จากทั้งหมด ${totalAvailableItems.toLocaleString()} รายการ`;
    }
  }

  const totalPages = Math.ceil(totalItems / state.adsItemsPerPage) || 1;
  if (state.adsCurrentPage > totalPages) state.adsCurrentPage = totalPages;
  if (state.adsCurrentPage < 1) state.adsCurrentPage = 1;

  const startIdx = (state.adsCurrentPage - 1) * state.adsItemsPerPage;
  const pageData = data.slice(startIdx, startIdx + state.adsItemsPerPage);

  if (pageData.length === 0) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="9" style="text-align:center; padding:32px; color:var(--text-muted); font-size: 14px; background:#fbfbfb;">🔍 ไม่พบข้อมูลโฆษณาที่ค้นหา</td>`;
    tbody.appendChild(tr);
  }

  pageData.forEach(d => {
    const cpc = d.clicks > 0 ? d.adSpend / d.clicks : 0;
    const cr = d.clicks > 0 ? (d.orders / d.clicks) * 100 : 0;
    
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="font-size:12px; max-width:250px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;" title="${d.name}">${d.name}</td>
      <td style="font-size:11px; color:var(--text-muted); font-family:monospace;">${d.productId}</td>
      <td style="text-align:right">${d.clicks.toLocaleString()}</td>
      <td style="text-align:right">${d.orders.toLocaleString()}</td>
      <td style="text-align:right; color:var(--red); font-weight:600">฿${d.adSpend.toLocaleString()}</td>
      <td style="text-align:right; color:var(--green)">฿${d.adRevenue.toLocaleString()}</td>
      <td style="text-align:right; font-weight:700; color:${d.roas >= 5 ? 'var(--green)' : d.roas >= 2 ? 'var(--amber)' : 'var(--red)'}">${d.roas.toFixed(2)}</td>
      <td style="text-align:right; color:var(--text-muted)">฿${cpc.toFixed(2)}</td>
      <td style="text-align:right; color:var(--blue)">${cr.toFixed(2)}%</td>
    `;
    tbody.appendChild(tr);
  });
  
  renderAdsPagination(totalPages);
}

function renderAdsPagination(totalPages) {
  const container = document.getElementById('ads-pagination');
  if (!container) return;
  if (totalPages <= 1) { container.innerHTML = ''; return; }
  
  let html = `<button class="page-btn" ${state.adsCurrentPage === 1 ? 'disabled' : ''} onclick="changeAdsPage(${state.adsCurrentPage - 1})">❮ Prev</button>`;
  for (let i = 1; i <= totalPages; i++) {
    if (i === 1 || i === totalPages || (i >= state.adsCurrentPage - 2 && i <= state.adsCurrentPage + 2)) {
      html += `<button class="page-btn ${i === state.adsCurrentPage ? 'active' : ''}" onclick="changeAdsPage(${i})">${i}</button>`;
    } else if (i === state.adsCurrentPage - 3 || i === state.adsCurrentPage + 3) {
      html += `<span style="color:var(--text-muted); padding:0 4px;">...</span>`;
    }
  }
  html += `<button class="page-btn" ${state.adsCurrentPage === totalPages ? 'disabled' : ''} onclick="changeAdsPage(${state.adsCurrentPage + 1})">Next ❯</button>`;
  container.innerHTML = html;
}

function changeAdsPage(p) {
  state.adsCurrentPage = p;
  renderAdsTable();
}

function exportAds() {
  if (!state.adsData.length) { showWarningMessage('ไม่มีข้อมูล', 'ไม่มีข้อมูลโฆษณาให้ดาวน์โหลด'); return; }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(state.adsData.map(d => ({
    'ชื่อโฆษณา': d.name,
    'รหัสสินค้า': d.productId,
    'จำนวนคลิก': d.clicks,
    'การสั่งซื้อ': d.orders,
    'ค่าโฆษณา': d.adSpend,
    'ยอดขาย': d.adRevenue,
    'ROAS': d.roas
  }))), 'Ads_Analytics');
  XLSX.writeFile(wb, `Torque_Ads_Report_${new Date().toISOString().slice(0,10)}.xlsx`);
}

// MOBILE MENU LOGIC
document.addEventListener('DOMContentLoaded', () => {
  const toggleBtn = document.getElementById('mobile-menu-toggle');
  const closeBtn = document.getElementById('mobile-sidebar-close');
  const sidebar = document.querySelector('.sidebar');
  const overlay = document.getElementById('sidebar-overlay');
  
  if (sidebar && overlay) {
    const toggleMenu = () => {
      sidebar.classList.toggle('mobile-open');
      overlay.classList.toggle('active');
    };
    
    if (toggleBtn) toggleBtn.addEventListener('click', toggleMenu);
    if (closeBtn) closeBtn.addEventListener('click', toggleMenu);
    overlay.addEventListener('click', toggleMenu);
    
    // Close menu when clicking nav items on mobile
    document.querySelectorAll('.nav-item').forEach(item => {
      item.addEventListener('click', () => {
        if (window.innerWidth <= 1024) {
          sidebar.classList.remove('mobile-open');
          overlay.classList.remove('active');
        }
      });
    });
  }
});
