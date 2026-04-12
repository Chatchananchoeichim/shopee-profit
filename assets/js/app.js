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
  currentUser: null,
  currentPage: 1,
  itemsPerPage: 100,
  resultSearchQuery: '',
  summaryCurrentPage: 1,
  summaryItemsPerPage: 100,
  summarySearchQuery: '',
  charts: {},
  sort: {
    cost:    { col: null, dir: 'asc' },
    result:  { col: null, dir: 'asc' },
    summary: { col: null, dir: 'asc' }
  }
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
}

// Initialize Firebase automatically with the hardcoded config
const firebaseConfig = {
  apiKey: "AIzaSyAmtmEQEFEQx4zf2Q1bbeTnhM3iPwPdnh4",
  authDomain: "torque-fd288.firebaseapp.com",
  databaseURL: "https://torque-fd288-default-rtdb.asia-southeast1.firebasedatabase.app",
  projectId: "torque-fd288",
  storageBucket: "torque-fd288.firebasestorage.app",
  messagingSenderId: "1040884189232",
  appId: "1:1040884189232:web:4a729a7f01a89df4da45f5"
};

if (!firebase.apps.length) {
    firebase.initializeApp(firebaseConfig);
}
let dbRef = firebase.database().ref('shopee_cost_data');

// Authenticate via Email/Password
function login() {
  const email = document.getElementById('auth-email').value.trim();
  const pass = document.getElementById('auth-pass').value.trim();
  if(!email || !pass) {
    Swal.fire('ข้อมูลไม่ครบ', 'กรุณากรอกอีเมลและรหัสผ่าน', 'warning'); 
    return;
  }
  
  const btn = document.getElementById('login-btn');
  btn.disabled = true;
  btn.innerText = "กำลังเข้าสู่ระบบ...";

  // Set persistence to SESSION (cleared when browser closes)
  firebase.auth().setPersistence(firebase.auth.Auth.Persistence.SESSION)
    .then(() => {
      return firebase.auth().signInWithEmailAndPassword(email, pass);
    })
    .catch((error) => {
      Swal.fire('เข้าสู่ระบบไม่สำเร็จ', "สาเหตุ: " + error.message, 'error');
      btn.disabled = false;
      btn.innerText = "เข้าสู่ระบบ";
    });
}

function logout() {
  firebase.auth().signOut();
}

// Idle Timeout Logic
let inactivityTimer;
const INACTIVITY_LIMIT_MS = 30 * 60 * 1000; // 30 minutes

function resetInactivityTimer() {
  if (state.currentUser) {
    clearTimeout(inactivityTimer);
    inactivityTimer = setTimeout(() => {
      Swal.fire('เซสชันหมดอายุ', 'ไม่มีการใช้งานเป็นเวลานาน (30 นาที) กรุณาเข้าสู่ระบบใหม่เพื่อความปลอดภัย', 'warning').then(() => {
        logout();
      });
    }, INACTIVITY_LIMIT_MS);
  }
}

// Track Auth State changes
firebase.auth().onAuthStateChanged((user) => {
  if (user) {
    state.currentUser = user;
    document.getElementById('auth-container').style.display = 'none';
    document.getElementById('app-container').style.display = 'block';
    document.getElementById('user-email').innerText = user.email;
    initCostMap(); // Only fetch data if authenticated
    
    // Start idle timer and attach listeners
    resetInactivityTimer();
    ['mousemove', 'mousedown', 'keydown', 'touchstart'].forEach(evt => {
      document.addEventListener(evt, resetInactivityTimer);
    });
  } else {
    state.currentUser = null;
    clearTimeout(inactivityTimer);
    document.getElementById('auth-container').style.display = 'flex';
    document.getElementById('app-container').style.display = 'none';
    if(dbRef) dbRef.off(); // stop listening
  }
});


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
       Swal.fire('ข้อผิดพลาดจากฐานข้อมูล', "ไม่สามารถบันทึกข้อมูลขึ้น Firebase ได้: อาจจะเกิดจากสิทธิ์การใช้งานของ Database\\n\\n(" + error.message + ")", 'error');
       document.getElementById('sync-status').innerHTML = "🔴 ออฟไลน์ (บันทึกไม่สำเร็จ)";
    });
  }
  populateCostFilterDropdown();
}

function toggleFirebaseSetup() {}
function connectFirebase(isAuto = false) { }
function disconnectFirebase() { }

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
      document.getElementById('sync-status').innerText = '● ออนไลน์ซิงก์เรียบร้อยแล้ว (Cloud)';
      document.getElementById('sync-status').style.color = 'var(--green)';
  }, (error) => {
      console.warn("Permission denied or error fetching init costs:", error);
  });
}

function renderCostTable(){
  const tbody = document.getElementById('cost-tbody');
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

function changeResultPerPage(val) {
  state.itemsPerPage = parseInt(val, 10);
  state.currentPage = 1;
  renderResultTable();
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
      Swal.fire({ title: 'ลบแล้ว!', text: 'ลบข้อมูลต้นทุนเรียบร้อย', icon: 'success', timer: 1500, showConfirmButton: false });
    }
  });
}

function lookupCost(productName, variantName, sku){
  if(sku){
    for(const r of state.costData){
      if(r.sku && r.sku.toLowerCase() === sku.toLowerCase()) return r.cost;
    }
  }
  for(const r of state.costData){
    if(r.variant && variantName && variantName.toLowerCase() === r.variant.toLowerCase()) return r.cost;
  }
  for(const r of state.costData){
    if(r.variant && variantName && variantName.toLowerCase().includes(r.variant.toLowerCase()) && r.variant.length > 2){
      if(!r.product || productName.toLowerCase().includes(r.product.toLowerCase().split('/')[0].trim())){
        return r.cost;
      }
    }
  }
  return null;
}

function switchTab(name){
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelector(`.tabs .tab[onclick*="${name}"]`).classList.add('active');
  document.querySelectorAll('.tab-content').forEach(tc=>tc.classList.remove('active'));
  document.getElementById('tab-'+name).classList.add('active');
}

function handleFile(input, type){
  const file = input.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = e=>{
    try{
      const wb = XLSX.read(e.target.result, {type:'array'});
      if(type==='order'){
        const sheetName = wb.SheetNames.find(n=>n.toLowerCase().includes('order')) || wb.SheetNames[0];
        state.orderData = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {defval:''});
        state.fileOrderLoaded = true;
        document.getElementById('order-label').textContent = '✓ '+file.name;
        document.getElementById('drop-order').classList.add('has-file');
      } else {
        const sheetName = wb.SheetNames.find(n=>n.toLowerCase().includes('income')) || wb.SheetNames[0];
        const raw = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {defval:'', header:1});
        let headerIdx = raw.findIndex(r=>(r||[]).some(c=>String(c).includes('หมายเลขคำสั่งซื้อ')));
        if(headerIdx<0) headerIdx = 5;
        const headers = raw[headerIdx] || [];
        state.incomeData = raw.slice(headerIdx+1).map(r=>{
          const obj={};
          headers.forEach((h,i)=>{ obj[String(h).trim()]=r[i]; });
          return obj;
        }).filter(r=>r['หมายเลขคำสั่งซื้อ']);
        state.fileIncomeLoaded = true;
        document.getElementById('income-label').textContent = '✓ '+file.name;
        document.getElementById('drop-income').classList.add('has-file');
      }
      checkReady();
    } catch(err){ Swal.fire('เกิดข้อผิดพลาดในการอ่านไฟล์', err.message, 'error'); }
  };
  reader.readAsArrayBuffer(file);
}

function checkReady(){
  document.getElementById('btn-process').disabled = !(state.fileOrderLoaded && state.fileIncomeLoaded);
}

function loadSampleData(){
  state.orderData = [
    {'หมายเลขคำสั่งซื้อ':'26030105B', 'เวลาที่สั่งซื้อ': '2024-03-26 10:00', 'สถานะการสั่งซื้อ':'สำเร็จแล้ว','ชื่อสินค้า':'SHAD TR41','ชื่อตัวเลือก':'Polypropylene,ถาด','จำนวน': 1, 'เลข SKU อ้างอิง':'TR41-PP-T', 'ยอดชำระเงิน':3700},
    {'หมายเลขคำสั่งซื้อ':'2603011T9', 'เวลาที่สั่งซื้อ': '2024-03-26 14:00', 'สถานะการสั่งซื้อ':'สำเร็จแล้ว','ชื่อสินค้า':'SOMAN S3','ชื่อตัวเลือก':'S3+ไมค์ก้าน','จำนวน': 2, 'เลข SKU อ้างอิง':'S3-MIC1', 'ยอดชำระเงิน':1580},
    {'หมายเลขคำสั่งซื้อ':'2603011T10', 'เวลาที่สั่งซื้อ': '2024-03-27 15:30', 'สถานะการสั่งซื้อ':'สำเร็จแล้ว','ชื่อสินค้า':'SOMAN X7','ชื่อตัวเลือก':'Carbon','จำนวน': 1, 'เลข SKU อ้างอิง':'X7-CARB', 'ยอดชำระเงิน':4500},
    {'หมายเลขคำสั่งซื้อ':'26030362P', 'เวลาที่สั่งซื้อ': '2024-03-27 09:00', 'สถานะการสั่งซื้อ':'ยกเลิกแล้ว','ชื่อสินค้า':'Soman M10','ชื่อตัวเลือก':'ขาวมุก','จำนวน': 1, 'เลข SKU อ้างอิง':'', 'ยอดชำระเงิน':0},
  ];
  state.incomeData = [
    {'หมายเลขคำสั่งซื้อ':'26030105B','จำนวนเงินทั้งหมดที่โอนแล้ว (฿)':2562},
    {'หมายเลขคำสั่งซื้อ':'2603011T9','จำนวนเงินทั้งหมดที่โอนแล้ว (฿)':1540},
    {'หมายเลขคำสั่งซื้อ':'2603011T10','จำนวนเงินทั้งหมดที่โอนแล้ว (฿)':4100},
  ];
  state.fileOrderLoaded = true;
  state.fileIncomeLoaded = true;
  document.getElementById('order-label').textContent = '✓ ข้อมูลตัวอย่าง (พร้อม SKU + จำนวนชิ้น)';
  document.getElementById('income-label').textContent = '✓ ข้อมูลตัวอย่าง';
  document.getElementById('drop-order').classList.add('has-file');
  document.getElementById('drop-income').classList.add('has-file');
  checkReady();
  document.getElementById('import-status').innerHTML = '<div class="alert success">✓ โหลดข้อมูลตัวอย่างแล้ว กด "คำนวณเลย" ได้เลย</div>';
}

function processFilesWithLoader() {
  const btn = document.getElementById('btn-process');
  const orgText = btn.innerText;
  btn.innerText = "กำลังวิเคราะห์ข้อมูล...";
  btn.disabled = true;
  
  setTimeout(() => {
    processFiles();
    btn.innerText = orgText;
    btn.disabled = false;
  }, 150);
}

function processFiles(){
  if(!state.orderData || !state.incomeData){ 
    Swal.fire('ข้อมูลไม่ครบถ้วน', 'กรุณาอัปโหลดไฟล์ทั้ง Order และ Income จากระบบของ Shopee ก่อนกดคำนวณครับ', 'warning'); 
    return; 
  }

  const incomeLookup = {};
  state.incomeData.forEach(row=>{
    const idKey = Object.keys(row).find(k=>k.includes('คำสั่งซื้อ')) || '';
    const amtKey = Object.keys(row).find(k=>k.includes('โอนแล้ว')) || '';
    const payKey = Object.keys(row).find(k=>k.includes('ช่องทางการชำระเงินของผู้ซื้อ')) || '';
    const pctKey = Object.keys(row).find(k=>k.includes('ค่าธรรมเนียม (%)')) || '';

    // Exact-match fee columns from Shopee Income CSV
    const rowKeys = Object.keys(row);
    const commKey    = rowKeys.find(k => k === 'ค่าคอมมิชชั่น') || '';
    const commAMSKey = rowKeys.find(k => k === 'ค่าคอมมิชชั่น AMS') || '';
    const servKey    = rowKeys.find(k => k === 'ค่าบริการ') || '';
    const platKey    = rowKeys.find(k => k.includes('โครงสร้างพื้นฐาน')) || '';
    const transKey   = rowKeys.find(k => k.includes('ค่าธุรกรรมการชำระเงิน')) || '';
    const shipDeductKey = rowKeys.find(k => k.includes('ค่าจัดส่งที่ Shopee ชำระโดยชื่อของคุณ')) || '';

    const orderId = String(row[idKey]||'').trim();
    const amount = parseFloat(row[amtKey]||0)||0;
    if(orderId) {
      incomeLookup[orderId] = {
        amount,
        payment: String(row[payKey]||''),
        feePct: String(row[pctKey]||''),
        commFee:    parseFloat(row[commKey]||0)||0,
        commAMSFee: parseFloat(row[commAMSKey]||0)||0,
        servFee:    parseFloat(row[servKey]||0)||0,
        platFee:    parseFloat(row[platKey]||0)||0,
        transFee:   parseFloat(row[transKey]||0)||0,
        shipDeduct: parseFloat(row[shipDeductKey]||0)||0
      };
    }
  });

  const orders = {};
  const sample = state.orderData[0]||{};
  const keys = Object.keys(sample);
  const findKey = (...patterns) => keys.find(k=>patterns.some(p=>k.includes(p)))||'';
  
  const orderIdKey = findKey('หมายเลขคำสั่งซื้อ');
  const statusKey = findKey('สถานะการสั่งซื้อ');
  const productKey = findKey('ชื่อสินค้า');
  const variantKey = findKey('ชื่อตัวเลือก');
  const sellingPriceKey = keys.find(k => k.trim() === 'ราคาขาย') || findKey('ราคาขาย');
  const paidKey = findKey('ยอดชำระเงิน','ยอดชำระ');
  const qtyKey = findKey('จำนวน');
  const skuKey = findKey('SKU', 'sku', 'อ้างอิง');
  const dateKey = findKey('เวลาที่สั่งซื้อ', 'วันที่สั่งซื้อ', 'เวลาชำระเงิน', 'วันเวลา');
  const payChannelKey = findKey('ช่องทางการชำระเงิน');
  const feePctKey = findKey('ค่าธรรมเนียม (%)');

  state.orderData.forEach(row=>{
    const orderId = String(row[orderIdKey]||'').trim();
    if(!orderId) return;
    
    if(!orders[orderId]) {
       const inc = incomeLookup[orderId] || {};
       orders[orderId] = {
         orderId,
         date: String(row[dateKey]||''),
         status: String(row[statusKey]||''),
         items: [],
         income: inc.amount || 0,
         paymentChannel: inc.payment || String(row[payChannelKey]||''),
         feePct: inc.feePct || String(row[feePctKey]||''),
         commFee:    inc.commFee    || 0,
         commAMSFee: inc.commAMSFee || 0,
         servFee:    inc.servFee    || 0,
         platFee:    inc.platFee    || 0,
         transFee:   inc.transFee   || 0,
         shipDeduct: inc.shipDeduct || 0
       };
    }
    const product = String(row[productKey]||'');
    const variant = String(row[variantKey]||'');
    const sku = String(row[skuKey]||'');
    const qty = parseInt(row[qtyKey]||1)||1;
    // ราคาขาย = selling price per item (ใช้คูณจำนวนเพื่อได้ยอดขายที่ถูกต้อง)
    const sellingPrice = parseFloat(row[sellingPriceKey]||0) || parseFloat(row[paidKey]||0) || 0;
    const paidPrice = parseFloat(row[paidKey]||0) || sellingPrice;
    
    // sellingPriceTotal สะสมไว้ที่ order สำหรับคำนวณ %
    orders[orderId].sellingPriceTotal = (orders[orderId].sellingPriceTotal||0) + (sellingPrice * qty);
    orders[orderId].items.push({ product, variant, sku, qty, salePrice: paidPrice });
  });

  const results = [];
  const skuSummaryMap = {};
  const timeSeriesMap = {};
  let totalIncome=0, totalCost=0, totalNet=0, successCount=0;
  const missingVariants = new Map();

  Object.values(orders).forEach(o => {
    const isCancelled = o.status.includes('ยกเลิก');
    let orderCost = 0;
    let missingCost = false;
    let orderSalePrice = 0;

    o.items.forEach(item => {
      let unitCost = null;
      if(!isCancelled){
        unitCost = lookupCost(item.product, item.variant, item.sku);
        if(unitCost === null){
          missingCost = true;
          const k = item.sku || item.variant || '(ไม่มี variant)';
          missingVariants.set(k, { sku: item.sku, variant: item.variant, product: item.product });
        } else {
          orderCost += (unitCost * item.qty);
        }
      }
      item.unitCost = unitCost;
      orderSalePrice += item.salePrice;
    });

    const net = (!isCancelled && !missingCost) ? o.income - orderCost : null;

    let shortDate = 'ไม่ระบุ';
    if(o.date) {
        const d = o.date.split(' ')[0];
        if(d.length > 5) shortDate = d;
    }

    if(!isCancelled && o.income > 0){
      successCount++;
      totalIncome += o.income;
      
      if(!timeSeriesMap[shortDate]) timeSeriesMap[shortDate] = { revenue: 0, profit: 0 };
      timeSeriesMap[shortDate].revenue += o.income;
      
      if(!missingCost){ 
        totalCost += orderCost; 
        totalNet += net; 
        timeSeriesMap[shortDate].profit += net;
      }
      
      o.items.forEach(item => {
        let key = item.sku ? item.sku : (item.variant || item.product);
        if(!skuSummaryMap[key]) {
          skuSummaryMap[key] = { sku: item.sku, title: item.product + (item.variant?' ('+item.variant+')':''), qty: 0, revenue: 0, cost: 0 };
        }
        let ratio = orderSalePrice > 0 ? (item.salePrice / orderSalePrice) : (1 / o.items.length);
        let itemRevenue = o.income * ratio;
        
        skuSummaryMap[key].qty += item.qty;
        skuSummaryMap[key].revenue += itemRevenue;
        if(item.unitCost !== null && !isCancelled) {
          skuSummaryMap[key].cost += (item.unitCost * item.qty);
        }
      });
    }

    o.items.forEach((item, idx) => {
      results.push({
        orderId: o.orderId,
        paymentChannel: o.paymentChannel,
        feePct: o.feePct,
        commFee:    o.commFee,
        commAMSFee: o.commAMSFee,
        servFee:    o.servFee,
        platFee:    o.platFee,
        transFee:   o.transFee,
        shipDeduct: o.shipDeduct,
        orderSalePrice: orderSalePrice,
        // ราคาขาย (selling price) รวมทั้ง order สำหรับคำนวณ %
        orderSellingPrice: o.sellingPriceTotal || orderSalePrice,
        status: o.status,
        product: item.product,
        variant: item.variant,
        sku: item.sku,
        qty: item.qty,
        salePrice: item.salePrice,
        income: o.income,
        unitCost: item.unitCost,
        itemCostTotal: item.unitCost !== null ? item.unitCost * item.qty : null,
        orderCostTotal: orderCost,
        net: net,
        isCancelled,
        missingCost,
        isFirst: idx === 0,
        rowSpan: o.items.length
      });
    });
  });

  state.results = results;
  state.summary = Object.values(skuSummaryMap).sort((a,b) => b.qty - a.qty);
  
  state.timeSeries = Object.keys(timeSeriesMap).map(k => ({
     date: k, revenue: timeSeriesMap[k].revenue, profit: timeSeriesMap[k].profit
  })).sort((a,b) => a.date.localeCompare(b.date));

  state.currentPage = 1;

  document.getElementById('r-orders').textContent = successCount.toLocaleString();
  document.getElementById('r-income').textContent = Math.round(totalIncome).toLocaleString();
  document.getElementById('r-cost').textContent = Math.round(totalCost).toLocaleString();
  document.getElementById('r-net').textContent = Math.round(totalNet).toLocaleString();

  // --- Fee Summary Bar ---
  const feeBar = document.getElementById('fee-summary-bar');
  if (feeBar) {
    let fs_comm = 0, fs_serv = 0, fs_trans = 0;
    results.forEach(r => {
      if (r.isFirst && !r.isCancelled) {
        fs_comm  += Math.abs(r.commFee  || 0);
        fs_serv  += Math.abs(r.servFee  || 0);
        fs_trans += Math.abs(r.transFee || 0);
      }
    });
    const fs_total = fs_comm + fs_serv + fs_trans;
    const grossSales = results.filter(r=>r.isFirst && !r.isCancelled).reduce((s,r)=>(s + (r.orderSellingPrice||r.orderSalePrice||0)),0);
    const fs_pct = grossSales > 0 ? (fs_total / grossSales * 100) : 0;
    const fmt = v => '\u0e3f' + Math.round(v).toLocaleString();
    document.getElementById('fee-total-comm').textContent  = fmt(fs_comm);
    document.getElementById('fee-total-serv').textContent  = fmt(fs_serv);
    document.getElementById('fee-total-trans').textContent = fmt(fs_trans);
    document.getElementById('fee-total-all').textContent   = fmt(fs_total);
    document.getElementById('fee-total-pct').textContent   = `\u0e04\u0e34\u0e14\u0e40\u0e1b\u0e47\u0e19 ${fs_pct.toFixed(2)}% \u0e02\u0e2d\u0e07\u0e22\u0e2d\u0e14\u0e02\u0e32\u0e22\u0e17\u0e31\u0e49\u0e07\u0e2b\u0e21\u0e14`;
    feeBar.style.display = 'block';
  }

  renderResultTable();
  renderSummaryTable();

  const mv = document.getElementById('missing-variants');
  if(missingVariants.size===0){
    mv.innerHTML = '<div class="alert success">✓ ทุกรายการมีข้อมูลต้นทุนครบถ้วน</div>';
  } else {
    let mvHtml = '<div class="alert warning">⚠ พบ '+missingVariants.size+' รายการออเดอร์ที่ยังไม่มีต้นทุน</div>';
    missingVariants.forEach((data, k) => {
      const displayStr = data.sku ? `SKU: ${data.sku}` : `Variant: ${data.variant || data.product}`;
      mvHtml += `<div class="missing-row flex-between">
        <span style="font-size:12px">${displayStr}</span>
        <button class="btn sm" onclick="prefillVariant('${(data.sku||'').replace(/'/g,"\\'")}','${(data.product||'').replace(/'/g,"\\'")}','${(data.variant||'').replace(/'/g,"\\'")}')">+ เพิ่ม Cost</button>
      </div>`;
    });
    mv.innerHTML = mvHtml;
  }

  document.getElementById('result-empty').style.display='none';
  document.getElementById('result-content').style.display='block';
  document.getElementById('summary-empty').style.display='none';
  document.getElementById('summary-content').style.display='block';
  document.getElementById('dashboard-empty').style.display='none';
  document.getElementById('dashboard-content').style.display='block';
  
  renderDashboard();
  switchTab('result');
}

function renderDashboard() {
  if(state.charts.marginTier) state.charts.marginTier.destroy();
  if(state.charts.timeSeries) state.charts.timeSeries.destroy();
  if(state.charts.bcg) state.charts.bcg.destroy();
  if(state.charts.topProfit) state.charts.topProfit.destroy();
  if(state.charts.salesQty) state.charts.salesQty.destroy();
  if(state.charts.costBreakdown) state.charts.costBreakdown.destroy();

  let totalRevenue = state.summary.reduce((a,b)=>a+b.revenue, 0);
  let totalCost = state.summary.reduce((a,b)=>a+b.cost, 0);
  let totalNet = totalRevenue - totalCost;
  let overallMargin = totalRevenue > 0 ? (totalNet / totalRevenue) * 100 : 0;

  // Margin Tiers & BCG Data
  let tierA = 0, tierB = 0, tierC = 0, lossCount = 0;
  let skuProfitCount = 0;
  let starCount = 0;
  let sumRev = 0, sumProfit = 0;
  const bcgData = [];

  state.summary.forEach(s => {
    const profit = s.revenue - s.cost;
    const marginPct = s.revenue > 0 ? (profit / s.revenue) * 100 : 0;
    
    if(marginPct >= 30) tierA++;
    else if(marginPct >= 15) tierB++;
    else if(profit > 0) tierC++;
    
    if(profit > 0) skuProfitCount++;
    if(profit < 0) lossCount++;
    
    if(s.qty > 0) {
      sumRev += s.revenue;
      sumProfit += profit;
      bcgData.push({
        x: s.revenue, y: profit, r: Math.max(8, Math.min(s.qty * 1.5, 30)),
        label: s.sku || s.title.substring(0, 15)
      });
    }
  });

  const avgRev =  bcgData.length ? sumRev / bcgData.length : 0;
  const avgProfit = bcgData.length ? sumProfit / bcgData.length : 0;

  bcgData.forEach(item => { if(item.x >= avgRev && item.y >= avgProfit) starCount++; });

  document.getElementById('d-margin').innerText = overallMargin.toFixed(1) + '%';
  document.getElementById('d-star-count').innerText = starCount + ' SKUs';
  document.getElementById('d-profit-skus').innerText = skuProfitCount + ' SKUs';
  document.getElementById('d-loss-skus').innerText = lossCount + ' SKUs';

  // 1. Cost & Fee Breakdown Chart
  let sumComm = 0, sumCommAMS = 0, sumServ = 0, sumPlat = 0, sumTrans = 0, sumShip = 0;
  state.results.forEach(r => {
    if (r.isFirst && !r.isCancelled) {
       sumComm    += Math.abs(r.commFee    || 0);
       sumCommAMS += Math.abs(r.commAMSFee || 0);
       sumServ    += Math.abs(r.servFee    || 0);
       sumPlat    += Math.abs(r.platFee    || 0);
       sumTrans   += Math.abs(r.transFee   || 0);
       sumShip    += Math.abs(r.shipDeduct || 0);
    }
  });

  const ctxCost = document.getElementById('costBreakdownChart').getContext('2d');
  state.charts.costBreakdown = new Chart(ctxCost, {
    type: 'pie',
    data: {
      labels: [
        'ต้นทุนสินค้า',
        'ค่าคอมมิชชั่น',
        'ค่าบริการ',
        'ค่าธุรกรรม (ชำระเงิน)'
      ],
      datasets: [{
        data: [totalCost, sumComm, sumServ, sumTrans],
        backgroundColor: ['#2b8a3e','#e67700','#f03e3e','#1c7ed6'],
        borderWidth: 2, borderColor: '#fff'
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { 
        legend: { position: 'right', labels: { font: { size: 10 }, boxWidth: 14 } },
        tooltip: {
          callbacks: {
            label: ctx => ` ${ctx.label}: ฿${Math.round(ctx.raw).toLocaleString()}`
          }
        }
      }
    }
  });

  // 2. Margin Tier Chart
  const ctxTier = document.getElementById('marginTierChart').getContext('2d');
  state.charts.marginTier = new Chart(ctxTier, {
    type: 'doughnut',
    data: {
      labels: ['Tier A (>30%)', 'Tier B (15-30%)', 'Tier C (<15%)', 'ขาดทุน'],
      datasets: [{
        data: [tierA, tierB, tierC, lossCount],
        backgroundColor: ['#2b8a3e', '#74b816', '#fab005', '#e03131'],
        borderWidth: 2, borderColor: '#fff'
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { position: 'right' } }
    }
  });

  // 2. Seasonality Chart (Time Series)
  const ctxTime = document.getElementById('timeSeriesChart').getContext('2d');
  state.charts.timeSeries = new Chart(ctxTime, {
    type: 'line',
    data: {
      labels: (state.timeSeries || []).map(d => d.date),
      datasets: [
        {
          type: 'bar', label: 'ยอดขาย (Revenue)',
          data: (state.timeSeries || []).map(d => d.revenue),
          backgroundColor: '#a5d8ff', borderRadius: 4
        },
        {
          type: 'line', label: 'กำไร (Profit)',
          data: (state.timeSeries || []).map(d => d.profit),
          borderColor: '#2b8a3e', backgroundColor: '#2b8a3e',
          borderWidth: 3, tension: 0.3
        }
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      scales: { y: { beginAtZero: true } },
      plugins: { legend: { position: 'top' } }
    }
  });

  // 3. BCG Matrix
  const scatterColors = bcgData.map(d => {
    if(d.x >= avgRev && d.y >= avgProfit) return 'rgba(250, 176, 5, 0.8)'; // Star
    if(d.x >= avgRev && d.y < avgProfit) return 'rgba(28, 126, 214, 0.8)'; // Cash Cow
    if(d.x < avgRev && d.y >= avgProfit) return 'rgba(190, 75, 219, 0.8)'; // Question Mark
    return 'rgba(134, 142, 150, 0.7)'; // Dog
  });

  const ctxBCG = document.getElementById('bcgMatrixChart').getContext('2d');
  state.charts.bcg = new Chart(ctxBCG, {
    type: 'bubble',
    data: {
      datasets: [{
        label: 'สินค้า', data: bcgData,
        backgroundColor: scatterColors,
        borderColor: '#fff', borderWidth: 1
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: function(ctx) {
              const d = ctx.raw;
              return `${d.label} | ขาย: ฿${Math.round(d.x).toLocaleString()} | กำไร: ฿${Math.round(d.y).toLocaleString()}`;
            }
          }
        }
      },
      scales: {
        x: { title: { display: true, text: 'ยอดขายสุทธิ' }, grid: { color: (ctx) => ctx.tick.value === 0 ? '#333' : '#eee' } },
        y: { title: { display: true, text: 'กำไร (Profit)' }, grid: { color: (ctx) => ctx.tick.value === 0 ? '#333' : '#eee' } }
      }
    }
  });

  // 4. Top Profit
  const sortedByProfit = [...state.summary].sort((a,b) => (b.revenue-b.cost) - (a.revenue-a.cost)).slice(0, 5);
  const ctxTopProfit = document.getElementById('topProfitChart').getContext('2d');
  state.charts.topProfit = new Chart(ctxTopProfit, {
    type: 'bar',
    data: {
      labels: sortedByProfit.map(r => r.sku || r.title.substring(0,18)+'..'),
      datasets: [{
        data: sortedByProfit.map(r => r.revenue - r.cost),
        backgroundColor: '#2b8a3e', borderRadius: 4
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { y: { beginAtZero: true } }
    }
  });

  // 5. Top Qty
  const sortedByQty = [...state.summary].sort((a,b) => b.qty - a.qty).slice(0, 10);
  const ctxSalesQty = document.getElementById('salesQtyChart').getContext('2d');
  state.charts.salesQty = new Chart(ctxSalesQty, {
    type: 'bar',
    data: {
      labels: sortedByQty.map(r => r.sku || r.title.substring(0,18)+'..'),
      datasets: [{
        data: sortedByQty.map(r => r.qty),
        backgroundColor: '#1c7ed6', borderRadius: 4
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: { y: { beginAtZero: true } }
    }
  });
}

function filterResultTable(){
  state.resultSearchQuery = document.getElementById('result-search').value.trim();
  state.currentPage = 1;
  renderResultTable();
}

function renderResultTable(){
  const tbody = document.getElementById('result-tbody');
  tbody.innerHTML = '';
  
  let data = state.results;
  if(state.filterMode==='no-cost') data = data.filter(r => r.missingCost && !r.isCancelled);
  if(state.filterMode==='success') data = data.filter(r => !r.isCancelled && !r.missingCost);

  if(state.resultSearchQuery) {
    const q = state.resultSearchQuery.toLowerCase();
    data = data.filter(r => 
      (r.orderId && r.orderId.toLowerCase().includes(q)) ||
      (r.product && r.product.toLowerCase().includes(q)) ||
      (r.sku && r.sku.toLowerCase().includes(q))
    );
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
      ? '<span class="badge gray">ยกเลิก</span>' 
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
      html += `<td rowspan="${r.rowSpan}" style="padding:8px 10px;vertical-align:middle;text-align:center;">${payBadge}</td>`;
      html += `<td rowspan="${r.rowSpan}" style="padding:8px 10px;vertical-align:middle;text-align:center;">
        <div style="font-size:13px;font-weight:700;color:var(--red);">${totalPct > 0 ? totalPct.toFixed(2)+'%' : (r.feePct || '-')}</div>
        ${pctBreakdown}
      </td>`;
    }

    // Product + SKU cell
    const skuChip = r.sku ? `<span style="display:inline-block;margin-top:2px;font-size:9px;font-weight:700;padding:1px 6px;background:#eff6ff;color:#1d4ed8;border-radius:4px;letter-spacing:0.5px;">${r.sku}</span>` : '';
    html += `
        <td style="padding:10px 12px;vertical-align:middle;">
          <div style="font-size:12px;font-weight:500;color:var(--text);max-width:170px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;" title="${r.product}">${r.product}</div>
          ${skuChip}
        </td>
        <td style="padding:8px 10px;vertical-align:middle;font-size:12px;color:var(--text-muted);max-width:120px;">${r.variant}</td>
        <td style="padding:8px;vertical-align:middle;text-align:center;">
          <span style="display:inline-block;font-size:13px;font-weight:700;min-width:24px;">${r.qty}</span>
        </td>
        <td style="padding:8px 12px;vertical-align:middle;text-align:right;font-size:12px;color:var(--text-muted);">฿${(r.salePrice||0).toLocaleString()}</td>
    `;

    if(r.isFirst) {
      // Shopee fee breakdown
      const shopeeFee = Math.max(0, (r.orderSalePrice||0) - (r.income||0));
      const feeItems = [
        { label: 'คอมมิชชั่น', val: r.commFee },
        { label: 'ค่าบริการ', val: r.servFee },
        { label: 'ธุรกรรม', val: r.transFee }
      ].filter(f => Math.abs(f.val||0) > 0);

      const feeBreakdown = feeItems.length > 0
        ? `<div style="margin-top:6px;border-top:1px dashed #fca5a5;padding-top:4px;">` +
          feeItems.map(f =>
            `<div style="display:flex;justify-content:space-between;align-items:center;font-size:10px;color:#9ca3af;line-height:1.8;gap:8px;">
              <span>${f.label}</span>
              <span style="color:#f87171;font-weight:500;">-฿${Math.abs(f.val).toLocaleString(undefined,{maximumFractionDigits:0})}</span>
            </div>`
          ).join('') + `</div>`
        : '';

      html += `<td rowspan="${r.rowSpan}" style="padding:10px 12px;vertical-align:top;min-width:140px;" class="border-right">
        <div style="display:flex;align-items:center;justify-content:flex-end;gap:4px;">
          <span style="font-size:13px;font-weight:700;color:var(--red);">-฿${shopeeFee.toLocaleString(undefined,{maximumFractionDigits:0})}</span>
        </div>
        ${feeBreakdown}
      </td>`;

      html += `<td rowspan="${r.rowSpan}" style="padding:10px 12px;vertical-align:middle;text-align:right;font-size:13px;font-weight:600;" class="border-right">฿${(r.income||0).toLocaleString()}</td>`;
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
        <td rowspan="${r.rowSpan}" style="padding:8px;vertical-align:middle;text-align:center;">${badge}</td>
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
  tbody.innerHTML = '';
  
  let data = state.summary;
  if(state.summarySearchQuery) {
    const q = state.summarySearchQuery.toLowerCase();
    data = data.filter(r => 
      (r.sku && r.sku.toLowerCase().includes(q)) ||
      (r.title && r.title.toLowerCase().includes(q))
    );
  }

  // Sort
  const ss = state.sort.summary;
  if (ss.col) {
    const dir = ss.dir === 'asc' ? 1 : -1;
    data = [...data].sort((a, b) => {
      if (ss.col === 'sku') return ((a.sku||'').localeCompare(b.sku||'', 'th')) * dir;
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
      <td style="font-size:13px">${r.title}</td>
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

function filterOrders(type){
  state.filterMode = type;
  state.currentPage = 1;
  renderResultTable();
}

function prefillVariant(sku, product, variant){
  document.getElementById('new-sku').value = sku;
  document.getElementById('new-product').value = product;
  document.getElementById('new-variant').value = variant;
  switchTab('cost');
  document.getElementById('new-cost').focus();
}

function getFormattedDateStr() {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth()+1).padStart(2,'0')}${String(d.getDate()).padStart(2,'0')}`;
}

function exportResult(){
  if(!state.results.length){ Swal.fire('ไม่มีข้อมูล', 'ไม่มีข้อมูลให้ดาวน์โหลด', 'info'); return; }
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
  if(!state.summary.length){ Swal.fire('ไม่มีข้อมูล', 'ไม่มีข้อมูลให้ดาวน์โหลด', 'info'); return; }
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
  XLSX.writeFile(wb, `Torque_Sales_Summary_${getFormattedDateStr()}.xlsx`);
}

function exportCostTable(){
  const rows = state.costData.map(r=>({
    'SKU': r.sku, 'สินค้า': r.product, 'Variant': r.variant, 'ต้นทุน': r.cost
  }));
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), 'Cost');
  XLSX.writeFile(wb, `Torque_CostTable_${getFormattedDateStr()}.xlsx`);
}

['drop-order','drop-income'].forEach(id=>{
  const el = document.getElementById(id);
  if(!el) return;
  el.addEventListener('dragover', e=>{e.preventDefault(); el.style.borderColor='var(--primary)';});
  el.addEventListener('dragleave', ()=>el.style.borderColor='');
  el.addEventListener('drop', e=>{
    e.preventDefault(); el.style.borderColor='';
    const file = e.dataTransfer.files[0];
    if(!file) return;
    const type = id==='drop-order'?'order':'income';
    handleFile({files:[file]}, type);
  });
});
