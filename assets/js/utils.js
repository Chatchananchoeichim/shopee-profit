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
  const { jsPDF } = window.jspdf;
  const doc = new jsPDF('p', 'mm', 'a4');
  
  Swal.fire({
    title: 'กำลังสร้างรายงาน PDF...',
    html: 'กรุณารอสักครู่ ระบบกำลังจัดเตรียมหน้ากระดาษและกราฟวิเคราะห์',
    allowOutsideClick: false,
    didOpen: () => { Swal.showLoading(); }
  });

  try {
    const canvas = await html2canvas(document.getElementById('tab-container-root'), {
      scale: 2,
      useCORS: true,
      logging: false,
      backgroundColor: state.theme === 'dark' ? '#0f172a' : '#ffffff'
    });
    
    const imgData = canvas.toDataURL('image/png');
    const imgProps = doc.getImageProperties(imgData);
    const pdfWidth = doc.internal.pageSize.getWidth();
    const pdfHeight = (imgProps.height * pdfWidth) / imgProps.width;
    
    doc.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
    doc.save(`Torque_Profit_Report_${getFormattedDateStr()}.pdf`);
    
    Swal.fire({ icon: 'success', title: 'ดาวน์โหลดสำเร็จ', text: 'บันทึกรายงาน PDF เรียบร้อยแล้ว', timer: 2000, showConfirmButton: false });
  } catch (err) {
    console.error(err);
    Swal.fire('Error', 'ไม่สามารถสร้าง PDF ได้ในขณะนี้', 'error');
  }
}
