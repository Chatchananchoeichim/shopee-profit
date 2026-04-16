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
    } catch(err){ 
      showErrorMessage('เกิดข้อผิดพลาดในการอ่านไฟล์', 'ไม่สามารถเปิดไฟล์ Excel ได้ครับ กรุณาตรวจสอบว่าไฟล์ไม่ใช่ไฟล์ที่เสียหรือมีการตั้งรหัสผ่านไว้<br><br>Error: ' + err.message); 
    }
  };
  reader.readAsArrayBuffer(file);
}

function handleAdsFile(input) {
  const file = input.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, {type: 'array'});
      const sheetName = wb.SheetNames[0];
      const sheet = wb.Sheets[sheetName];
      
      // Convert to JSON (raw array of arrays to find header)
      const raw = XLSX.utils.sheet_to_json(sheet, {header: 1, defval: ''});
      
      let headerIdx = raw.findIndex(row => (row || []).some(cell => String(cell).includes('รหัสสินค้า') || String(cell).includes('ชื่อโฆษณา')));
      
      if (headerIdx === -1) {
        showErrorMessage('รูปแบบไฟล์ Ads ไม่ถูกต้อง', 'ไม่พบหัวตารางที่ต้องการ (รหัสสินค้า หรือ ชื่อโฆษณา) ในไฟล์นี้ครับ');
        return;
      }
      
      const headers = raw[headerIdx];
      const dataRows = raw.slice(headerIdx + 1);
      
      state.adsData = dataRows.map(row => {
        const obj = {};
        headers.forEach((h, i) => {
          if (h) obj[String(h).trim()] = row[i];
        });
        
        const cleanNum = (v) => String(v || 0).replace(/,/g, '');
        
        return {
          name: obj['ชื่อโฆษณา'] || '',
          productId: obj['รหัสสินค้า'] || '',
          clicks: parseInt(cleanNum(obj['จำนวนคลิก'])) || 0,
          orders: parseInt(cleanNum(obj['การสั่งซื้อ'])) || 0,
          adSpend: parseFloat(cleanNum(obj['ค่าโฆษณา'])) || 0,
          adRevenue: parseFloat(cleanNum(obj['ยอดขาย'])) || 0,
          roas: parseFloat(cleanNum(obj['ยอดขาย/รายจ่าย (ROAS)'])) || 0,
          raw: obj
        };
      }).filter(d => d.productId || d.name);
      
      document.getElementById('ads-label').innerHTML = '✓ ' + file.name;
      document.getElementById('drop-ads').classList.add('has-file');
      
      renderAdsTable();
      if (state.results.length > 0) renderDashboard();
      
      document.getElementById('ads-empty').style.display = 'none';
      document.getElementById('ads-content').style.display = 'block';
      
    } catch(err) {
      showErrorMessage('เกิดข้อผิดพลาดในการอ่านไฟล์ Ads', 'ไม่สามารถอ่านไฟล์ได้: ' + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

function handleStatsFile(input) {
  const file = input.files[0];
  if(!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const data = new Uint8Array(e.target.result);
      const wb = XLSX.read(data, {type: 'array'});
      const sheet = wb.Sheets[wb.SheetNames[0]];
      const raw = XLSX.utils.sheet_to_json(sheet, {defval: ''});
      
      if (!raw.length) {
        showErrorMessage('ไฟล์ว่างเปล่า', 'ไม่พบข้อมูลในไฟล์ Shop Stats นี้ครับ');
        return;
      }

      const row = raw[0]; 
      const keys = Object.keys(row);
      
      const findVal = (variants) => {
        const k = keys.find(k => variants.some(v => String(k).includes(v)));
        return k ? row[k] : null;
      };

      const visitors = findVal(['ผู้เยี่ยมชม', 'ผู้เข้าชม', 'Visitors', 'ӹǹ']);
      const cr       = findVal(['อัตราการซื้อ', 'Conversion Rate', 'ѵҡëԹ']);
      const aov      = findVal(['ยอดขายเฉลี่ย', 'Basket Size', 'ʹµͤ觫']);
      const repeat   = findVal(['ซื้อซ้ำ', 'Repeat', 'ѵҡáѺҫͫ']);

      const cleanNumStr = (v) => {
        if (v === null || v === undefined) return '0';
        return String(v).replace(/,/g, '').replace(/"/g, '').trim() || '0';
      };

      state.shopStats = {
        visitors: cleanNumStr(visitors),
        cr: String(cr || '0%'),
        aov: cleanNumStr(aov),
        repeat: String(repeat || '0%')
      };

      document.getElementById('stats-label').innerHTML = '✓ ' + file.name;
      document.getElementById('drop-stats').classList.add('has-file');

      if (state.results.length > 0) renderDashboard();
      
    } catch(err) {
      showErrorMessage('เกิดข้อผิดพลาดในการอ่านไฟล์', 'ไม่สามารถอ่านไฟล์ XLSX ได้: ' + err.message);
    }
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
  Swal.fire({
    title: 'กำลังประมวลผลข้อมูล...',
    html: `
      <div style="text-align:center;">
        <lottie-player src="https://assets5.lottiefiles.com/packages/lf20_vnikbeve.json" 
          background="transparent" speed="1" style="width: 200px; height: 200px; margin: 0 auto;" loop autoplay>
        </lottie-player>
        <p style="font-size:14px; color:var(--text-muted); margin-top:10px;">ระบบกำลังคำนวณกำไรสุทธิและวิเคราะห์ข้อมูล Ads ให้คุณ...</p>
      </div>
    `,
    showConfirmButton: false,
    allowOutsideClick: false,
    timer: 2000,
    timerProgressBar: true,
    didOpen: () => {
      Swal.showLoading();
      setTimeout(() => {
        processFiles();
        Swal.close();
      }, 1800);
    },
    customClass: {
      popup: 'swal2-borderless'
    }
  });
}

function processFiles(skipTabSwitch = false){
  try {
    if(!state.orderData || !state.incomeData){ 
      if(!skipTabSwitch) Swal.fire('ข้อมูลไม่ครบถ้วน', 'กรุณาอัปโหลดไฟล์ทั้ง Order และ Income จากระบบของ Shopee ก่อนกดคำนวณครับ', 'warning'); 
      return; 
    }

    const incomeLookup = {};
    const mandatoryOrderKeys = ['หมายเลขคำสั่งซื้อ', 'สถานะการสั่งซื้อ', 'ราคาขาย', 'ยอดชำระเงิน'];
    const missingKeys = [];

    const keys = Object.keys(state.orderData[0] || {});
    const findKey = (...variants) => {
      const found = keys.find(k => variants.some(v => k.includes(v)));
      if (!found) {
        // Only mark mandatory ones that we absolutely need
        const isMandatory = mandatoryOrderKeys.some(m => variants.includes(m));
        if (isMandatory) missingKeys.push(variants[0]);
      }
      return found || '';
    };

    const orderIdKey = findKey('หมายเลขคำสั่งซื้อ');
    const statusKey = findKey('สถานะการสั่งซื้อ', 'สถานะ');
    const productKey = findKey('ชื่อสินค้า');
    const variantKey = findKey('ชื่อตัวเลือก');
    const sellingPriceKey = keys.find(k => k.trim() === 'ราคาขาย') || findKey('ราคาขาย');
    const paidKey = findKey('ยอดชำระเงิน','ยอดชำระ');
    const qtyKey = findKey('จำนวน');
    const skuKey = findKey('SKU', 'sku', 'อ้างอิง');
    const dateKey = findKey('วันที่ทำการสั่งซื้อ', 'เวลาที่สั่งซื้อ', 'วันที่สั่งซื้อ', 'เวลาชำระเงิน', 'เวลาการชำระเงิน', 'วันเวลา', 'Order Creation', 'เวลาเริ่มต้น');
    const payChannelKey = findKey('ช่องทางการชำระเงิน');
    const feePctKey = findKey('ค่าธรรมเนียม (%)');

    // New: Extract fees from Order file as fallback
    const orderCommKey  = findKey('ค่าคอมมิชชั่น');
    const orderTransKey = findKey('Transaction Fee', 'ค่าธุรกรรม');
    const orderServKey  = findKey('ค่าบริการ');
    const orderPaidKey  = findKey('ราคาสินค้าที่ชำระโดยผู้ซื้อ (THB)');

    if (missingKeys.length > 0) {
      showErrorMessage('รูปแบบข้อมูลไม่ถูกต้อง', `ไฟล์ Order ที่อัปโหลดขาดคอลัมน์สำคัญดังนี้: <br><br><b>${missingKeys.join(', ')}</b><br><br>กรุณาใช้ไฟล์ต้นฉบับจาก Shopee Seller Center (ห้ามแก้ไขหัวตารางไฟล์ Excel)`);
      return;
    }

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
    state.orderData.forEach(row=>{
      const orderId = String(row[orderIdKey]||'').trim();
      if(!orderId) return;

      const rowStatus = String(row[statusKey]||'');
      const isCancelledRow = rowStatus.includes('ยกเลิก');
      
      if(!orders[orderId]) {
         const inc = incomeLookup[orderId] || {};
         
         // Fallback from order file if income not found (e.g. In Transit)
         const ordComm  = parseFloat(row[orderCommKey]||0)||0;
         const ordTrans = parseFloat(row[orderTransKey]||0)||0;
         const ordServ  = parseFloat(row[orderServKey]||0)||0;
         // We use "ราคาสินค้าที่ชำระโดยผู้ซื้อ (THB)" as a closer estimate to net payout if income file missing
         const ordIncome = isCancelledRow ? 0 : (parseFloat(row[orderPaidKey]||0)||0);

         orders[orderId] = {
           orderId,
           date: String(row[dateKey]||''),
           status: rowStatus,
           items: [],
           income: state.incomeOverrides[orderId] !== undefined ? state.incomeOverrides[orderId] : (inc.amount || ordIncome || 0),
           paymentChannel: inc.payment || String(row[payChannelKey]||''),
           feePct: inc.feePct || String(row[feePctKey]||''),
           commFee:    isCancelledRow ? 0 : (inc.commFee    || ordComm),
           commAMSFee: isCancelledRow ? 0 : (inc.commAMSFee || 0),
           servFee:    isCancelledRow ? 0 : (inc.servFee    || ordServ),
           platFee:    isCancelledRow ? 0 : (inc.platFee    || 0),
           transFee:   isCancelledRow ? 0 : (inc.transFee   || ordTrans),
           shipDeduct: isCancelledRow ? 0 : (inc.shipDeduct || 0),
           isEstimated: (inc.amount === undefined && state.incomeOverrides[orderId] === undefined && !isCancelledRow),
           isOverridden: state.incomeOverrides[orderId] !== undefined && !isCancelledRow,
           isCancelled: isCancelledRow
         };
      }
      const product = String(row[productKey]||'');
      const variant = String(row[variantKey]||'');
      const sku = String(row[skuKey]||'');
      const qty = parseInt(row[qtyKey]||1)||1;
      // ราคาขาย = selling price per item (ใช้คูณจำนวนเพื่อได้ยอดขายที่ถูกต้อง)
      const sellingPrice = isCancelledRow ? 0 : (parseFloat(row[sellingPriceKey]||0) || parseFloat(row[paidKey]||0) || 0);
      const paidPrice = isCancelledRow ? 0 : (parseFloat(row[paidKey]||0) || sellingPrice);
      
      // sellingPriceTotal สะสมไว้ที่ order สำหรับคำนวณ %
      orders[orderId].sellingPriceTotal = (orders[orderId].sellingPriceTotal||0) + (sellingPrice * qty);
      orders[orderId].items.push({ 
        product, variant, sku, qty, 
        salePrice: paidPrice,         // Paid amount for this line
        unitSellingPrice: sellingPrice // Advertised unit price
      });
    });

    const results = [];
    const skuSummaryMap = {};
    const timeSeriesMap = {};
     let totalIncome=0, totalCost=0, totalNet=0, successCount=0;
     let totalGrossSales=0, cancelledCount=0, otherCount=0, allOrdersCount=0;
     const missingVariants = new Map();

     const allOrdersValue = Object.values(orders);
     allOrdersCount = allOrdersValue.length;

     allOrdersValue.forEach(o => {
       // Enhanced Success Match
       const isSuccess = ['สำเร็จ', 'จัดส่งสำเร็จ', 'ผู้ซื้อได้รับสินค้า', 'การจัดส่ง'].some(s => o.status.includes(s));
       const isCancelled = o.status.includes('ยกเลิก');
       
       if (isSuccess) { /* already handled in loop below but we'll use flags */ }
       else if (isCancelled) cancelledCount++;
       else otherCount++;

       if(!isCancelled) {
         totalGrossSales += (o.sellingPriceTotal || 0);
       }

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

      if(isSuccess && o.income > 0){
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
          
          // Use Selling Price Total (before discounts) as the ratio for fair distribution
          const itemSellingTotal = (item.unitSellingPrice || 0) * item.qty;
          const totalOrderSelling = o.sellingPriceTotal || orderSalePrice || 1;
          let ratio = itemSellingTotal / totalOrderSelling;
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
          isEstimated: o.isEstimated,
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

    document.getElementById('r-total-orders').textContent = allOrdersCount.toLocaleString();
    document.getElementById('r-sub-success').textContent = successCount.toLocaleString();
    document.getElementById('r-sub-cancelled').textContent = cancelledCount.toLocaleString();
    document.getElementById('r-sub-other').textContent = otherCount.toLocaleString();
    document.getElementById('r-gross-sales').textContent = Math.round(totalGrossSales).toLocaleString();

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
    if (!skipTabSwitch) switchTab('result');
  } catch (err) {
    console.error(err);
    showErrorMessage('เกิดข้อผิดพลาดระหว่างคำนวณ', 'ระบบพบปัญหาเกี่ยวกับรูปแบบข้อมูลในไฟล์ครับ ไม่สามารถคำนวณต่อได้<br><br>Error: ' + err.message);
  }
}
