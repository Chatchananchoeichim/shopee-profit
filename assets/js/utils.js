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
    title: `<span style="font-family:'Outfit',sans-serif; font-weight:800; color:#ef4444;">${title}</span>`,
    html: `
      <div style="text-align:left; font-size:13px; color:#475569; background:#f8fafc; padding:16px; border-radius:12px; border:1px solid #e2e8f0; line-height:1.6;">
        <div style="font-weight:700; color:#ef4444; margin-bottom:8px; display:flex; align-items:center; gap:6px;">
          <svg style="width:16px; height:16px;" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>
          รายละเอียดข้อผิดพลาด
        </div>
        ${details}
      </div>
      <div style="margin-top:15px; font-size:12px; color:var(--text-muted); text-align:center;">
        ลองตรวจสอบไฟล์อีกครั้ง หรือห้ามแก้ไขโครงสร้างไฟล์ Excel ของ Shopee ครับ
      </div>
    `,
    confirmButtonText: 'ตกลง รับทราบ',
    confirmButtonColor: '#f97316',
    customClass: {
      popup: 'swal2-borderless',
      confirmButton: 'swal2-confirm'
    }
  });
}
