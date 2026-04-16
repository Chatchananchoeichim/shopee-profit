function importCostTable(event){
  const file = event.target.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb = XLSX.read(e.target.result, {type:'array'});
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(sheet, {defval:''});
      let importedCount = 0;
      raw.forEach(r => {
        const sku = String(r['SKU'] || r['sku'] || '').trim();
        const product = String(r['ชื่อสินค้า'] || r['Product'] || '').trim();
        const variant = String(r['ชื่อตัวเลือก'] || r['Variant'] || '').trim();
        let costStr = r['ต้นทุน (฿)'] || r['ต้นทุน'] || r['Cost'] || r['cost'];
        const cost = parseFloat(String(costStr).replace(/,/g, ''));
        if(!isNaN(cost) && cost > 0 && (sku || product || variant)) {
          const existing = state.costData.find(d => (sku && d.sku === sku) || (!sku && d.variant === variant));
          if(existing){
             existing.cost = cost;
             if(sku) existing.sku = sku;
             if(product) existing.product = product;
          } else {
             state.costData.unshift({ id: Date.now() + importedCount, sku, product, variant, cost });
          }
          importedCount++;
        }
      });
      saveCostsToLocal();
      if(importedCount > 0) {
        Swal.fire('นำเข้าข้อมูลสำเร็จ', `อิมพอร์ตข้อมูลต้นทุนแล้ว จำนวน ${importedCount} รายการ`, 'success');
      }
    } catch(err) {
      Swal.fire('เกิดข้อผิดพลาด', "ไม่สามารถอ่านไฟล์ได้ กรุณาตรวจสอบฟอร์แมต Excel ให้ถูกต้อง: " + err.message, 'error');
    }
    event.target.value = '';
  };
  reader.readAsArrayBuffer(file);
}

function saveCostsToLocal() {
  localStorage.setItem('torque_cost_data', JSON.stringify(state.costData));
  
  if (dbRef) {
      const dataToSave = (state.costData && state.costData.length > 0) ? state.costData : null;
      dbRef.set(dataToSave).catch((error) => {
        Swal.fire('ข้อผิดพลาดจากฐานข้อมูล', "ไม่สามารถบันทึกข้อมูลขึ้น Firebase ได้: อาจจะเกิดจากสิทธิ์การใช้งานของ Database\n\n(" + error.message + ")", 'error');
        if(document.getElementById('sync-status-text')) document.getElementById('sync-status-text').innerText = "🔴 ออฟไลน์ (บันทึกไม่สำเร็จ)";
     });
  }
  populateCostFilterDropdown();
}

function resetDefaultCosts(){
  Swal.fire({
    title: 'ยืนยันการล้างข้อมูล?',
    text: "คุุณต้องการล้างข้อมูลต้นทุนทั้งหมดใช่ไหม? ข้อมูลที่ถูกลบไปแล้วจะไม่สามารถกู้คืนได้",
    icon: 'warning',
    showCancelButton: true,
    confirmButtonColor: '#e03131',
    confirmButtonText: 'ใช่, ล้างข้อมูลทั้งหมด',
    cancelButtonText: 'ยกเลิก'
  }).then((result) => {
    if(result.isConfirmed) {
      state.costData = [];
      if(dbRef) dbRef.set([]);
      renderCostTable();
      Swal.fire('สำเร็จ', 'ข้อมูลต้นทุนทั้งหมดถูกลบเรียบร้อยแล้ว', 'success');
    }
  });
}

function initCostMap(){
  // Setup Realtime Database listener
  dbRef.on('value', (snapshot) => {
      const data = snapshot.val();
      if(data) {
         state.costData = data;
      } else {
         state.costData = [];
      }
      renderCostTable();
      populateCostFilterDropdown();
      if(document.getElementById('sync-status-text')) {
        document.getElementById('sync-status-text').innerText = 'ออนไลน์ซิงก์เรียบร้อยแล้ว (Cloud)';
      }
      if(document.getElementById('sync-status-dot')) {
        document.getElementById('sync-status-dot').style.background = 'var(--green)';
      }
  }, (error) => {
      console.warn("Permission denied or error fetching init costs:", error);
  });
}

function renderCostTable(){
  const tbody = document.getElementById('cost-tbody');
  if(!tbody) return;
  tbody.innerHTML = '';
  
  // Filter by search query & dropdown
  let filteredData = state.costData;

  if(state.costSearchQueries && state.costSearchQueries.length > 0) {
    filteredData = filteredData.filter(r => {
      // return true if any of the multiple selected queries matches this row
      return state.costSearchQueries.some(q => 
        (r.sku && r.sku.toLowerCase().includes(q)) || 
        (r.product && r.product.toLowerCase().includes(q)) ||
        (r.variant && r.variant.toLowerCase().includes(q))
      );
    });
  }

  // Sort
  const cs = state.sort.cost;
  if (cs.col) {
    const dir = cs.dir === 'asc' ? 1 : -1;
    filteredData = [...filteredData].sort((a, b) => {
      let va = (a[cs.col] || '').toString().toLowerCase();
      let vb = (b[cs.col] || '').toString().toLowerCase();
      if (cs.col === 'cost') { va = parseFloat(a.cost)||0; vb = parseFloat(b.cost)||0; }
      if (typeof va === 'number') return (va - vb) * dir;
      return va.localeCompare(vb, 'th') * dir;
    });
  }

  // Pagination logic
  const totalItems = filteredData.length;
  const totalAvailableItems = state.costData.length;
  const costCountEl = document.getElementById('cost-count');
  if (costCountEl) {
    if (totalItems === totalAvailableItems) {
      costCountEl.innerText = `(ทั้งหมด ${totalItems.toLocaleString()} รายการ)`;
    } else {
      costCountEl.innerText = `(พบ ${totalItems.toLocaleString()} จาก ${totalAvailableItems.toLocaleString()} รายการ)`;
    }
  }
  const totalPages = Math.ceil(totalItems / state.costItemsPerPage) || 1;
  
  if (state.costCurrentPage > totalPages) state.costCurrentPage = totalPages;
  if (state.costCurrentPage < 1) state.costCurrentPage = 1;
  
  const startIdx = (state.costCurrentPage - 1) * state.costItemsPerPage;
  const pageData = filteredData.slice(startIdx, startIdx + state.costItemsPerPage);

  if (pageData.length === 0) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="5" style="text-align:center; padding:32px; color:var(--text-muted); font-size: 14px; background:#fbfbfb;">🔍 ไม่พบข้อมูลที่คุณค้นหา</td>`;
    tbody.appendChild(tr);
  }

  pageData.forEach(row => {
    const tr = document.createElement('tr');
    if (state.editingId === row.id) {
      tr.innerHTML = `
        <td style="padding-right:4px"><input type="text" id="edit-sku-${row.id}" value="${row.sku || ''}"></td>
        <td style="padding-right:4px"><input type="text" id="edit-prod-${row.id}" value="${row.product || ''}"></td>
        <td style="padding-right:4px"><input type="text" id="edit-var-${row.id}" value="${row.variant || ''}"></td>
        <td style="padding-right:4px"><input type="number" id="edit-cost-${row.id}" value="${row.cost}"></td>
        <td style="display:flex;gap:4px;justify-content:flex-end">
          <button class="btn primary sm" onclick="saveCost(${row.id})">✔</button>
          <button class="btn sm" onclick="cancelEdit()">✖</button>
        </td>
      `;
    } else {
      tr.innerHTML = `
        <td style="color:var(--blue);cursor:pointer;font-weight:500" ondblclick="editCost(${row.id})">${row.sku || '—'}</td>
        <td style="cursor:pointer" ondblclick="editCost(${row.id})">${row.product || '—'}</td>
        <td style="cursor:pointer" ondblclick="editCost(${row.id})"><strong>${row.variant || '(ทุก variant)'}</strong></td>
        <td style="text-align:right;font-weight:600;cursor:pointer" ondblclick="editCost(${row.id})">฿${row.cost.toLocaleString()}</td>
        <td style="display:flex;gap:12px;justify-content:flex-end">
          <button onclick="editCost(${row.id})" style="background:none;border:none;cursor:pointer;color:var(--text-muted);font-size:16px" title="แก้ไข">✎</button>
          <button onclick="duplicateCost(${row.id})" style="background:none;border:none;cursor:pointer;color:var(--text-muted);font-size:16px" title="คัดลอก (Duplicate)">⧉</button>
          <button onclick="deleteCost(${row.id})" style="background:none;border:none;cursor:pointer;color:var(--text-muted);font-size:20px;line-height:0.8" title="ลบ">×</button>
        </td>
      `;
    }
    tbody.appendChild(tr);
  });
  
  renderCostPagination(totalPages);
}

function filterCostTable() {
  const vals = $('#cost-search').val() || [];
  state.costSearchQueries = vals.map(v => v.trim().toLowerCase()).filter(v=>v);
  state.costCurrentPage = 1;
  renderCostTable();
}

let isSelect2Initialized = false;

function populateCostFilterDropdown() {
  const $select = $('#cost-search');
  if(!$select.length) return;
  
  if (!isSelect2Initialized) {
    $select.select2({
      placeholder: "ค้นหาและเลือกสินค้า (เลือกได้หลายรายการ)...",
      allowClear: true,
      width: '100%'
    });
    $select.on('change', function() {
      filterCostTable();
    });
    isSelect2Initialized = true;
  }
  
  const currentValues = $select.val() || [];
  
  const options = new Set();
  state.costData.forEach(r => {
    if(r.sku) options.add(r.sku);
    if(r.product) options.add(r.product);
  });
  
  const sortedOptions = Array.from(options).sort((a,b) => a.localeCompare(b));
  
  $select.empty();
  sortedOptions.forEach(opt => {
    $select.append(new Option(opt, opt, false, false));
  });
  
  $select.val(currentValues);
  $select.trigger('change.select2');
}

function renderCostPagination(totalPages) {
  const container = document.getElementById('cost-pagination');
  if(!container) return;
  if(totalPages <= 1) { container.innerHTML = ''; return; }
  
  let html = `<button class="page-btn" ${state.costCurrentPage === 1 ? 'disabled' : ''} onclick="changeCostPage(${state.costCurrentPage - 1})">❮ Prev</button>`;
  
  let start = Math.max(1, state.costCurrentPage - 2);
  let end = Math.min(totalPages, start + 4);
  if(end - start < 4) start = Math.max(1, end - 4);
  
  if(start > 1) {
    html += `<button class="page-btn" onclick="changeCostPage(1)">1</button>`;
    if(start > 2) html += `<span style="color:var(--text-muted);font-weight:600;">...</span>`;
  }
  
  for(let i=start; i<=end; i++){
    html += `<button class="page-btn ${i === state.costCurrentPage ? 'active' : ''}" onclick="changeCostPage(${i})">${i}</button>`;
  }
  
  if(end < totalPages){
    if(end < totalPages - 1) html += `<span style="color:var(--text-muted);font-weight:600;">...</span>`;
    html += `<button class="page-btn" onclick="changeCostPage(${totalPages})">${totalPages}</button>`;
  }
  
  html += `<button class="page-btn" ${state.costCurrentPage === totalPages ? 'disabled' : ''} onclick="changeCostPage(${state.costCurrentPage + 1})">Next ❯</button>`;
  container.innerHTML = html;
}

function changeCostPage(p) {
  state.costCurrentPage = p;
  renderCostTable();
}

function changeCostPerPage(val) {
  state.costItemsPerPage = parseInt(val, 10);
  state.costCurrentPage = 1;
  renderCostTable();
}

function editCost(id){
  state.editingId = id;
  renderCostTable();
}

function cancelEdit(){
  state.editingId = null;
  renderCostTable();
}

function saveCost(id){
  const s = document.getElementById('edit-sku-'+id).value.trim();
  const p = document.getElementById('edit-prod-'+id).value.trim();
  const v = document.getElementById('edit-var-'+id).value.trim();
  const c = parseFloat(document.getElementById('edit-cost-'+id).value);
  if(isNaN(c)) { 
    Swal.fire('ข้อมูลไม่ถูกต้อง', 'กรุณากรอกราคาต้นทุนเป็นตัวเลขให้ถูกต้อง', 'warning'); 
    return; 
  }
  
  const row = state.costData.find(r => r.id === id);
  if(row){
    row.sku = s;
    row.product = p;
    row.variant = v;
    row.cost = c;
    saveCostsToLocal();
  }
  state.editingId = null;
  renderCostTable();
  processFiles(true); // Sync results
}

function duplicateCost(id){
  const row = state.costData.find(r => r.id === id);
  if(row){
    Swal.fire({
      title: 'คัดลอกข้อมูล?',
      text: `ต้องการคัดลอกข้อมูลต้นทุน ${row.sku || row.variant || row.product} ใช่หรือไม่?`,
      icon: 'question',
      showCancelButton: true,
      confirmButtonText: 'ใช่, คัดลอก',
      cancelButtonText: 'ยกเลิก'
    }).then((result) => {
      if(result.isConfirmed) {
        const newId = Date.now();
        const newRow = { ...row, id: newId };
        const index = state.costData.findIndex(r => r.id === id);
        state.costData.splice(index + 1, 0, newRow);
        state.editingId = newId;
        saveCostsToLocal();
        renderCostTable();
        processFiles(true); // Sync results
      }
    });
  }
}

function addCostRow(){
  const s = document.getElementById('new-sku').value.trim();
  const p = document.getElementById('new-product').value.trim();
  const v = document.getElementById('new-variant').value.trim();
  const c = parseFloat(document.getElementById('new-cost').value);
  if(!s && !v && !p) { Swal.fire('ข้อมูลไม่ครบ', 'กรุณากรอก SKU หรือชื่อตัวเลือก หรือ ชื่อสินค้า อย่างน้อยหนึ่งอย่าง', 'warning'); return; }
  if(isNaN(c)) { Swal.fire('ข้อผิดพลาด', 'กรุณากรอกต้นทุนให้ถูกต้อง', 'warning'); return; }
  
  Swal.fire({
    title: 'ยืนยันการบันทึกต้นทุน',
    html: `เพิ่มข้อมูล: <b>${s || v || p}</b><br>ราคาต้นทุน: <b>฿${c}</b>`,
    icon: 'question',
    showCancelButton: true,
    confirmButtonText: 'บันทึก',
    cancelButtonText: 'ยกเลิก'
  }).then((result) => {
    if(result.isConfirmed) {
      state.costData.unshift({ id: Date.now(), sku: s, product: p, variant: v, cost: c });
      saveCostsToLocal();
      document.getElementById('new-sku').value='';
      document.getElementById('new-product').value='';
      document.getElementById('new-variant').value='';
      document.getElementById('new-cost').value='';
      renderCostTable();
      processFiles(true); // Sync results in background
      Swal.fire({ title: 'สำเร็จ', text: 'เพิ่มข้อมูลต้นทุนสำเร็จ', icon: 'success', timer: 1500, showConfirmButton: false });
    }
  });
}

function deleteCost(id){
  const row = state.costData.find(r => r.id === id);
  if(!row) return;
  const itemName = [row.sku, row.product, row.variant].filter(Boolean).join(' / ') || 'บรรทัดนี้';
  
  Swal.fire({
    title: 'ยืนยันการลบ?',
    html: `ลบข้อมูลต้นทุน: <b>${itemName}</b><br>ข้อมูลนี้จะไม่สามารถกู้คืนได้!`,
    icon: 'warning',
    showCancelButton: true,
    confirmButtonColor: '#e03131',
    confirmButtonText: 'ใช่, ลบเลย',
    cancelButtonText: 'ยกเลิก'
  }).then((result) => {
    if(result.isConfirmed) {
      state.costData = state.costData.filter(r => r.id !== id);
      saveCostsToLocal();
      renderCostTable();
      processFiles(true); // Sync results
      Swal.fire({ title: 'ลบแล้ว!', text: 'ลบข้อมูลต้นทุนเรียบร้อย', icon: 'success', timer: 1500, showConfirmButton: false });
    }
  });
}

function lookupCost(productName, variantName, sku){
  const s = (sku || '').toString().trim().toLowerCase();
  const v = (variantName || '').toString().trim().toLowerCase();
  const p = (productName || '').toString().trim().toLowerCase();

  // 1. Priority: Exact SKU Match
  if (s && s !== '-' && s !== 'none' && s !== 'n/a') {
    const match = state.costData.find(r => (r.sku || '').toString().trim().toLowerCase() === s);
    if (match) return match.cost;
  }

  // 2. Priority: Product Title + Variant Match
  // Look for a record where the stored product title is a substring of the Shopee product name
  for (const r of state.costData) {
    const rP = (r.product || '').toString().trim().toLowerCase();
    const rV = (r.variant || '').toString().trim().toLowerCase();

    // Check if product title matches (Stored title must be part of Shopee title)
    // If the stored product name is empty, we only match if SKU or Variant is unique
    const productTitleMatch = rP && p.includes(rP);

    if (productTitleMatch) {
      // a) Exact variant match (including both empty)
      if (rV === v) return r.cost;

      // b) Stored variant is wildcard "(ทุก variant)" or empty, and Shopee variant is empty
      const isStoredWildcard = !rV || rV === "(ทุก variant)" || rV === "any variant";
      if (isStoredWildcard && v === "") return r.cost;

      // c) Partial variant match (e.g. stored "Blue" matches Shopee "Light Blue")
      if (rV && v && v.includes(rV) && rV.length > 1) return r.cost;
    }
  }

  // 3. Fallback: Variant Match only (If the variant name is unique enough)
  if (v && v !== "" && v.length > 2) {
    const variantOnlyMatch = state.costData.find(r => {
      const rV = (r.variant || '').toString().trim().toLowerCase();
      return rV === v;
    });
    if (variantOnlyMatch) return variantOnlyMatch.cost;
  }

  return null;
}

function exportCostTable(){
  const rows = state.costData.map(r=>({
    'SKU': r.sku, 'สินค้า': r.product, 'Variant': r.variant, 'ต้นทุน': r.cost
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Cost');
  XLSX.writeFile(wb, `Torque_CostTable_${getFormattedDateStr()}.xlsx`);
}
