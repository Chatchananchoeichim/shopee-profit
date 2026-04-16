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
