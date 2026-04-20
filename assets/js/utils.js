function switchTab(name){
  document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
  document.querySelector(`.tabs .tab[onclick*="${name}"]`).classList.add('active');
  document.querySelectorAll('.tab-content').forEach(tc=>tc.classList.remove('active'));
  document.getElementById('tab-'+name).classList.add('active');
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

function showErrorMessage(title, details) {
  Swal.fire({
    icon: 'error',
    title: `<span style="font-family:'Plus Jakarta Sans',sans-serif; font-weight:800; color:#ef4444;">${title}</span>`,
    html: `
      <div style="text-align:left; font-size:13px; color:#475569; background:#f8fafc; padding:16px; border-radius:12px; border:1px solid #e2e8f0; line-height:1.6;">
        <div style="font-weight:700; color:#ef4444; margin-bottom:8px; display:flex; align-items:center; gap:6px;">
          <svg style="width:16px; height:16px;" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
          รายละเอียดข้อผิดพลาด
        </div>
        ${details}
      </div>
    `,
    padding: '2.5rem',
    heightAuto: false,
    customClass: { popup: 'swal2-borderless' }
  });
}

function showWarningMessage(title, text) {
  Swal.fire({
    icon: 'warning',
    title: `<span style="font-family:'Plus Jakarta Sans',sans-serif; font-weight:800; color:#f59e0b;">${title}</span>`,
    text: text,
    confirmButtonText: 'ตกลง',
    confirmButtonColor: '#f59e0b',
    heightAuto: false,
    customClass: { popup: 'swal2-borderless' }
  });
}

function showSuccessMessage(title, text) {
  Swal.fire({
    icon: 'success',
    title: `<span style="font-family:'Plus Jakarta Sans',sans-serif; font-weight:800; color:#10b981;">${title}</span>`,
    text: text,
    timer: 2000,
    showConfirmButton: false,
    heightAuto: false,
    customClass: { popup: 'swal2-borderless' }
  });
}
async function exportProPDF() {
  if(!state.results || !state.results.length){ showWarningMessage('ไม่มีข้อมูล', 'ไม่มีข้อมูลให้ดาวน์โหลด'); return; }

  Swal.fire({
    title: 'กำลังสร้างรายงาน PDF...',
    html: 'กำลังดึงข้อมูลและจัดการหน้ากระดาษ (Text-based PDF)...',
    allowOutsideClick: false,
    didOpen: () => { Swal.showLoading(); }
  });

  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('l', 'mm', 'a4');

    // Fetch and load Thai font
    const fontUrl = 'https://raw.githubusercontent.com/google/fonts/main/ofl/sarabun/Sarabun-Regular.ttf';
    const resp = await fetch(fontUrl);
    if (!resp.ok) throw new Error("Failed to fetch font: " + resp.status);
    const buffer = await resp.arrayBuffer();
    
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    const base64Font = window.btoa(binary);

    doc.addFileToVFS("Sarabun-Regular.ttf", base64Font);
    doc.addFont("Sarabun-Regular.ttf", "Sarabun", "normal");
    doc.setFont("Sarabun");

    doc.setFontSize(16);
    doc.text(`รายงานผลประกอบการ Shopee - Torque`, 14, 15);
    doc.setFontSize(10);
    doc.text(`วันที่ดึงข้อมูล: ${new Date().toLocaleDateString('th-TH')}`, 14, 22);

    const textVal = (id) => document.getElementById(id) ? document.getElementById(id).textContent : '-';
    
    // Draw Summary KPI Box
    doc.autoTable({
      body: [
        [
          `ออเดอร์ (ทั้งหมด)\n${textVal('r-total-orders')} รายการ`,
          `ยอดขายรวม\n฿${textVal('r-gross-sales')}`,
          `ออเดอร์สำเร็จ\n${textVal('r-sub-success')} รายการ`
        ],
        [
          `รายรับโอนแล้ว\n฿${textVal('r-income')}`,
          `ต้นทุนรวม\n฿${textVal('r-cost')}`,
          `กำไรสุทธิ NET\n฿${textVal('r-net')}`
        ],
        [
          {
            content: `สรุปค่าธรรมเนียม SHOPEE สะสม:\nคอมมิชชั่น: ${textVal('fee-total-comm')}  |  บริการ: ${textVal('fee-total-serv')}  |  ธุรกรรม: ${textVal('fee-total-trans')}   ==>   รวมทั้งหมด: ${textVal('fee-total-all')} (${textVal('fee-total-pct')})`,
            colSpan: 3,
            styles: { fillColor: [254, 242, 242], textColor: [220, 38, 38], fontStyle: 'bold' }
          }
        ]
      ],
      startY: 26,
      theme: 'grid',
      styles: { font: 'Sarabun', fontSize: 11, halign: 'center', valign: 'middle', cellPadding: 3, lineColor: [226, 232, 240], lineWidth: 0.2 },
      columnStyles: {
        0: { cellWidth: 89, fontStyle: 'bold', textColor: [71, 85, 105] },
        1: { cellWidth: 89, fontStyle: 'bold', textColor: [217, 119, 6] },
        2: { cellWidth: 89, fontStyle: 'bold', textColor: [37, 99, 235] }
      },
      didParseCell: function(data) {
        if (data.row.index === 1) {
          if(data.column.index === 0) data.cell.styles.textColor = [37, 99, 235]; // blue
          if(data.column.index === 1) data.cell.styles.textColor = [220, 38, 38]; // red
          if(data.column.index === 2) data.cell.styles.textColor = [22, 163, 74]; // green
        }
      }
    });

    const orderHeaders = [['Order ID', 'สถานะ', 'ช่องทางชำระ', '% หัก', 'SKU', 'สินค้า', 'จำนวน', 'ยอดขาย', 'Shopee หัก', 'Income', 'ต้นทุนรวม', 'Net กำไร']];
    const orderData = [];
    
    const exportResults = state.currentExportResult || state.results;
    exportResults.forEach(r => {
      let feePct = '';
      let shopeeFee = '';
      if (r.isFirst) {
        const sp = r.orderSellingPrice || r.orderSalePrice || 1;
        const totalFee = Math.abs(r.commFee||0) + Math.abs(r.servFee||0) + Math.abs(r.transFee||0);
        feePct = (totalFee / sp * 100).toFixed(2) + '%';
        const deduct = Math.abs(r.commFee||0) + Math.abs(r.servFee||0) + Math.abs(r.platFee||0) + Math.abs(r.transFee||0) + Math.abs(r.shipDeduct||0) + Math.abs(r.commAMSFee||0);
        shopeeFee = deduct.toLocaleString(undefined, {minimumFractionDigits:2});
      }
      
      const isCancel = r.isCancelled;
      // Make sure EVERYTHING is a string to prevent getTextWidth error in jspdf-autotable
      orderData.push([
        String(r.isFirst ? (r.orderId || '') : ''),
        String(r.isFirst ? (r.status || '') : ''),
        String(r.isFirst ? (r.paymentChannel || '') : ''),
        String(r.isFirst ? feePct : ''),
        String(r.sku || '-'),
        String(r.product || '-'),
        String(r.qty || 0),
        String((r.salePrice||0).toLocaleString()),
        String(r.isFirst ? shopeeFee : ''),
        String(r.isFirst ? (r.income||0).toLocaleString() : ''),
        String((r.itemCostTotal||0).toLocaleString()),
        String(r.isFirst ? (r.net||0).toLocaleString() : '')
      ]);
    });

    doc.autoTable({
      head: orderHeaders,
      body: orderData,
      startY: doc.lastAutoTable.finalY + 8,
      styles: { font: 'Sarabun', fontSize: 7, cellPadding: 1.5, lineColor: [226, 232, 240], lineWidth: 0.1 },
      headStyles: { fillColor: [249, 115, 22], textColor: [255, 255, 255], fontStyle: 'bold' },
      columnStyles: {
        0: { cellWidth: 26 },
        1: { cellWidth: 20 },
        5: { cellWidth: 40 }
      },
      didParseCell: function(data) {
        // Red text for cancelled orders if we can identify them (we can check the status column or data index)
        const rowIndex = data.row.index;
        const resultItem = state.results[rowIndex];
        if (resultItem && resultItem.isCancelled) {
          data.cell.styles.textColor = [220, 38, 38];
          data.cell.styles.fillColor = [254, 242, 242];
        }
      }
    });

    const exportSummary = state.currentExportSummary || state.summary;
    if (exportSummary && exportSummary.length > 0) {
      doc.addPage();
      doc.setFontSize(14);
      doc.text(`สรุปรายสินค้า (Product Summary)`, 14, 15);
      
      const summaryHeaders = [['SKU', 'สินค้า/ตัวเลือก', 'จำนวนขาย', 'รายได้', 'ต้นทุน', 'กำไรสุทธิ']];
      // Typecast everything to string here too
      const summaryData = exportSummary.map(r => [
        String(r.sku || '-'),
        String(r.title || '-'),
        String(r.qty || 0),
        String(Math.round(r.revenue || 0).toLocaleString()),
        String(Math.round(r.cost || 0).toLocaleString()),
        String(Math.round((r.revenue || 0) - (r.cost || 0)).toLocaleString())
      ]);

      doc.autoTable({
        head: summaryHeaders,
        body: summaryData,
        startY: 20,
        styles: { font: 'Sarabun', fontSize: 8, cellPadding: 2, lineColor: [226, 232, 240], lineWidth: 0.1 },
        headStyles: { fillColor: [59, 130, 246], textColor: [255, 255, 255], fontStyle: 'bold' }
      });
    }

    doc.save(`Torque_Report_${getFormattedDateStr()}.pdf`);
    Swal.fire({ icon: 'success', title: 'ดาวน์โหลดสำเร็จ', text: 'ระบบดึงข้อมูลดิบจัดทำเป็น PDF เรียบร้อยแล้ว (ขนาดไฟล์เล็กลงมาก)', timer: 2500, showConfirmButton: false });

  } catch (err) {
    console.error(err);
    Swal.fire('Error', 'ไม่สามารถสร้าง PDF ได้ในขณะนี้ (อาจเป็นปัญหาด้าน Network ตอนโหลด Font)', 'error');
  }
}

function exportFullExcel() {
  if(!state.results || !state.results.length){ showWarningMessage('ไม่มีข้อมูล', 'ไม่มีข้อมูลให้ดาวน์โหลด'); return; }
  
  const wb = XLSX.utils.book_new();
  const headers = ['Order ID', 'สถานะ', 'ช่องทางชำระ', '% หัก', 'SKU', 'สินค้า', 'Variant', 'จำนวน', 'ยอดขาย', 'Shopee หัก(฿)', 'Income โอน', 'ต้นทุนรวม(ชิ้นนี้)', 'Net กำไร(ออเดอร์)'];
  const merges = [];
  let currentRow = 1;

  // 1. Orders Sheet
  const exportResults = state.currentExportResult || state.results;
  const orderRows = exportResults.map(r => {
    let feePct = '';
    let shopeeFee = '';
    if (r.isFirst) {
      const sp = r.orderSellingPrice || r.orderSalePrice || 1;
      const totalFee = Math.abs(r.commFee||0) + Math.abs(r.servFee||0) + Math.abs(r.transFee||0);
      feePct = (totalFee / sp * 100).toFixed(2) + '%';
      const deduct = Math.abs(r.commFee||0) + Math.abs(r.servFee||0) + Math.abs(r.platFee||0) + Math.abs(r.transFee||0) + Math.abs(r.shipDeduct||0) + Math.abs(r.commAMSFee||0);
      shopeeFee = deduct;

      if (r.rowSpan > 1) {
        const endRow = currentRow + r.rowSpan - 1;
        [0, 1, 2, 3, 9, 10, 12].forEach(col => merges.push({ s: { r: currentRow, c: col }, e: { r: endRow, c: col } }));
      }
    }
    currentRow++;
    return {
      'Order ID': r.orderId, 'สถานะ': r.status, 'ช่องทางชำระ': r.isFirst ? (r.paymentChannel || '') : '',
      '% หัก': feePct, 'SKU': r.sku, 'สินค้า': r.product, 'Variant': r.variant, 'จำนวน': r.qty,
      'ยอดขาย': r.salePrice, 'Shopee หัก(฿)': shopeeFee, 'Income โอน': r.isFirst ? r.income : '',
      'ต้นทุนรวม(ชิ้นนี้)': r.itemCostTotal, 'Net กำไร(ออเดอร์)': r.isFirst ? r.net : '', '_isCancelled': r.isCancelled
    };
  });
  
  const cleanOrderRows = orderRows.map(r => { const n = {...r}; delete n._isCancelled; return n; });
  const wsOrders = XLSX.utils.json_to_sheet(cleanOrderRows, { header: headers });
  if (merges.length > 0) wsOrders['!merges'] = merges;
  
  const range = XLSX.utils.decode_range(wsOrders['!ref']);
  for (let c = 0; c <= range.e.c; c++) {
    const cellAddr = XLSX.utils.encode_cell({r: 0, c});
    if (wsOrders[cellAddr]) wsOrders[cellAddr].s = { fill: { fgColor: { rgb: "F1F5F9" } }, font: { bold: true, color: { rgb: "334155" } }, alignment: { horizontal: "center", vertical: "center" } };
  }
  const alignTopCols = [0, 1, 2, 3, 9, 10, 12];
  orderRows.forEach((rObj, rIdx) => {
    const rowNum = rIdx + 1;
    for (let c = 0; c <= range.e.c; c++) {
      const cellAddr = XLSX.utils.encode_cell({r: rowNum, c});
      if (!wsOrders[cellAddr]) wsOrders[cellAddr] = { v: '', t: 's' };
      if (!wsOrders[cellAddr].s) wsOrders[cellAddr].s = {};
      if (alignTopCols.includes(c)) wsOrders[cellAddr].s.alignment = { vertical: "top" };
      if (rObj._isCancelled) { wsOrders[cellAddr].s.fill = { fgColor: { rgb: "FEF2F2" } }; wsOrders[cellAddr].s.font = { color: { rgb: "991B1B" } }; }
    }
  });
  XLSX.utils.book_append_sheet(wb, wsOrders, 'All_Orders');

  // 2. Summary Sheet
  const exportSummary = state.currentExportSummary || state.summary;
  if (exportSummary && exportSummary.length > 0) {
    const summaryRows = exportSummary.map(r=>({
      'SKU': r.sku, 'สินค้า/ตัวเลือก': r.title, 'จำนวนที่ขายได้ (ชิ้น)': r.qty, 'รายได้ประมาณ (฿)': Math.round(r.revenue),
      'ต้นทุนรวม (฿)': Math.round(r.cost), 'เฉลี่ยต้นทุนต่อชิ้น (฿)': Math.round(r.qty>0?r.cost/r.qty:0),
      'เฉลี่ยยอดขายต่อชิ้น (฿)': Math.round(r.qty>0?r.revenue/r.qty:0), 'เฉลี่ยกำไรต่อชิ้น (฿)': Math.round(r.qty>0?(r.revenue-r.cost)/r.qty:0)
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), 'Product_Summary');
  }

  // 3. Ads Sheet
  if (state.adsData && state.adsData.length > 0) {
    const adsRows = state.adsData.map(d => ({
      'ชื่อโฆษณา': d.name, 'รหัสสินค้า': d.productId, 'จำนวนคลิก': d.clicks, 'การสั่งซื้อ': d.orders,
      'ค่าโฆษณา': d.adSpend, 'ยอดขาย': d.adRevenue, 'ROAS': d.roas
    }));
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(adsRows), 'Ads_Analytics');
  }

  XLSX.writeFile(wb, `Torque_Full_Report_${getFormattedDateStr()}.xlsx`);
}
