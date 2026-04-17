const FILES={ritm:null,incident:null,macm:null};
const STEPS=["ritm","incident","macm"];
let currentStep=1;

function initWizard(s){
  currentStep=s||1;
  showStep(currentStep);
  setupDropZones();
  setupFileInputs();
  if(currentStep>=6) loadSummary();
}

function showStep(s){
  document.querySelectorAll(".upload-section").forEach(e=>e.classList.add("d-none"));
  if(s===1)      document.getElementById("upload-ritm").classList.remove("d-none");
  else if(s===2) document.getElementById("upload-incident").classList.remove("d-none");
  else if(s===3) document.getElementById("upload-macm").classList.remove("d-none");
  else if(s===4) document.getElementById("upload-team-efforts").classList.remove("d-none");
  else if(s===5) document.getElementById("upload-reference").classList.remove("d-none");
  else           { document.getElementById("upload-complete").classList.remove("d-none"); loadSummary(); }
}

function setupFileInputs(){
  STEPS.forEach(t=>{
    const inp=document.getElementById("file-"+t);
    if(inp) inp.addEventListener("change",()=>{ if(inp.files[0]) onFileSel(t,inp.files[0]); });
  });
  // Reference file input
  const refInp=document.getElementById("ref-file");
  if(refInp) refInp.addEventListener("change",()=>{
    if(!refInp.files[0]) return;
    const f=refInp.files[0];
    document.getElementById("ref-filename").textContent=f.name;
    document.getElementById("ref-size").textContent=fmtSz(f.size);
    document.getElementById("ref-file-info").classList.remove("d-none");
    const btn=document.getElementById("btn-upload-ref"); if(btn) btn.disabled=false;
  });
}

function onFileSel(t,f){
  FILES[t]=f;
  document.getElementById(t+"-filename").textContent=f.name;
  document.getElementById(t+"-size").textContent=fmtSz(f.size);
  document.getElementById(t+"-file-info").classList.remove("d-none");
  const btn=document.getElementById("btn-upload-"+t); if(btn) btn.disabled=false;
}

function fmtSz(b){ if(b<1024)return b+" B"; if(b<1048576)return(b/1024).toFixed(1)+" KB"; return(b/1048576).toFixed(1)+" MB"; }

function setupDropZones(){
  STEPS.forEach(t=>{
    const z=document.getElementById("drop-"+t); if(!z)return;
    z.addEventListener("dragover",e=>{e.preventDefault();z.classList.add("drag-over");});
    z.addEventListener("dragleave",()=>z.classList.remove("drag-over"));
    z.addEventListener("drop",e=>{
      e.preventDefault(); z.classList.remove("drag-over");
      const f=e.dataTransfer.files[0]; if(!f)return;
      const dt=new DataTransfer(); dt.items.add(f);
      document.getElementById("file-"+t).files=dt.files;
      onFileSel(t,f);
    });
  });
  // Reference drop zone
  const rz=document.getElementById("drop-reference"); if(!rz)return;
  rz.addEventListener("dragover",e=>{e.preventDefault();rz.classList.add("drag-over");});
  rz.addEventListener("dragleave",()=>rz.classList.remove("drag-over"));
  rz.addEventListener("drop",e=>{
    e.preventDefault(); rz.classList.remove("drag-over");
    const f=e.dataTransfer.files[0]; if(!f)return;
    const dt=new DataTransfer(); dt.items.add(f);
    document.getElementById("ref-file").files=dt.files;
    document.getElementById("ref-filename").textContent=f.name;
    document.getElementById("ref-size").textContent=fmtSz(f.size);
    document.getElementById("ref-file-info").classList.remove("d-none");
    const btn=document.getElementById("btn-upload-ref"); if(btn) btn.disabled=false;
  });
}

async function uploadFile(t){
  const f=FILES[t]; if(!f){alert("Please select a file first.");return;}
  const btn=document.getElementById("btn-upload-"+t);
  const prog=document.getElementById(t+"-progress");
  const res=document.getElementById(t+"-result");
  btn.disabled=true; btn.innerHTML='<span class="spinner-border spinner-border-sm me-2"></span>Processing…';
  prog.classList.remove("d-none"); res.classList.add("d-none");
  const fd=new FormData(); fd.append("file",f);
  if(t==="macm"){
    const ml=document.getElementById("macm-row-label")?.value||"";
    if(ml) fd.append("macm_label",ml);
  }
  try{
    const resp=await fetch("/upload/"+t,{method:"POST",body:fd});
    prog.classList.add("d-none"); res.classList.remove("d-none");
    if(!resp.ok && resp.headers.get("content-type")?.includes("text/html")){
      res.innerHTML=renderErr(`Server error (HTTP ${resp.status}) — check server logs`);
      btn.disabled=false; btn.innerHTML='<i class="bi bi-upload me-2"></i>Retry';
      return;
    }
    const data=await resp.json();
    if(data.success){
      res.innerHTML=renderOK(t,data);
      updateIndicator(t,data.total);
      setTimeout(()=>{ currentStep=data.next_step; showStep(currentStep); },1500);
    } else {
      res.innerHTML=renderErr(data.error||"Upload failed.");
      btn.disabled=false; btn.innerHTML='<i class="bi bi-upload me-2"></i>Retry';
    }
  } catch(e){
    prog.classList.add("d-none"); res.classList.remove("d-none");
    res.innerHTML=renderErr("Network error: "+e.message);
    btn.disabled=false; btn.innerHTML='<i class="bi bi-upload me-2"></i>Retry';
  }
}

function renderOK(t,d){
  const warn=(d.errors&&d.errors.length)?`<div class="mt-1 text-warning small">${d.errors.join("; ")}</div>`:"";
  return`<div class="result-ok"><div class="d-flex align-items-center gap-2 mb-1">
    <i class="bi bi-check-circle-fill text-success fs-5"></i>
    <strong>${d.ticket_type} — ${d.total.toLocaleString()} tickets found</strong></div>
    ${warn}</div>`;
}
function renderErr(msg){ return`<div class="result-err"><i class="bi bi-x-circle-fill text-danger me-2"></i><strong>Error:</strong> ${msg}</div>`; }

function updateIndicator(type,count){
  const n={ritm:1,incident:2,macm:3}[type]; if(!n)return;
  const circ=document.getElementById("step"+n+"-circle");
  const cntEl=document.getElementById("step"+n+"-count");
  const conn=document.getElementById("conn"+n);
  const lbl=document.getElementById("step"+n+"-label");
  if(circ){circ.classList.remove("disabled");circ.classList.add("done");}
  if(cntEl) cntEl.textContent=count.toLocaleString()+" tickets";
  if(conn) conn.classList.add("active");
  if(lbl){lbl.classList.remove("text-muted");lbl.classList.add("text-success");}
  if(n<6){
    const nc=document.getElementById("step"+(n+1)+"-circle");
    const nl=document.getElementById("step"+(n+1)+"-label");
    if(nc) nc.classList.remove("disabled");
    if(nl){nl.classList.remove("text-muted");nl.classList.add("text-primary");}
  }
}

function goBack(t){const p={incident:1,macm:2,team_efforts:3,reference:4}[t]||1; currentStep=p; showStep(p);}
function skipRef(){currentStep=6; showStep(6);}
function skipTeamEfforts(){currentStep=5; showStep(5);}

// ── Team Efforts upload ─────────────────────────────────────────
(function(){
  document.addEventListener("DOMContentLoaded",function(){
    const inp=document.getElementById("te-file");
    if(!inp) return;
    function _teShowFiles(files){
      if(!files||!files.length) return;
      const info=document.getElementById("te-file-info");
      const names=Array.from(files).map(f=>f.name);
      document.getElementById("te-filename").textContent=
        names.length===1?names[0]:`${names.length} files: ${names.join(", ")}`;
      if(info) info.classList.remove("d-none");
      document.getElementById("btn-upload-te").disabled=false;
    }
    inp.addEventListener("change",function(){ _teShowFiles(this.files); });
    // Drop zone
    const dz=document.getElementById("drop-team-efforts");
    if(dz){
      dz.addEventListener("dragover",e=>{e.preventDefault();dz.classList.add("drag-over");});
      dz.addEventListener("dragleave",()=>dz.classList.remove("drag-over"));
      dz.addEventListener("drop",e=>{
        e.preventDefault(); dz.classList.remove("drag-over");
        if(!e.dataTransfer.files.length) return;
        inp.files=e.dataTransfer.files;
        _teShowFiles(e.dataTransfer.files);
      });
    }
  });
})();

async function uploadTeamEfforts(){
  const inp=document.getElementById("te-file");
  const st=document.getElementById("te-status");
  const btn=document.getElementById("btn-upload-te");
  if(!inp||!inp.files.length){if(st)st.textContent="Select at least one file.";return;}
  btn.disabled=true;
  btn.innerHTML='<span class="spinner-border spinner-border-sm me-2"></span>Processing\u2026';
  if(st) st.textContent="";
  const fd=new FormData();
  Array.from(inp.files).forEach(f=>fd.append("file",f));
  try{
    const r=await fetch("/upload/team_efforts",{method:"POST",body:fd});
    const d=await r.json();
    if(d.success){
      if(st) st.innerHTML=`<span class="text-success fw-semibold">\u2713 ${d.total.toLocaleString()} entries \u2014 <strong>${d.jira_count.toLocaleString()} JIRA</strong>, ${d.ritm_count} RITM, ${d.incident_count} Incident, ${d.macm_count} MACM</span>`;
      // Mark step 4 done in progress bar
      const circ=document.getElementById("step4-circle");
      const conn=document.getElementById("conn4");
      const lbl=document.getElementById("step4-label");
      const cnt=document.getElementById("step4-count");
      if(circ){circ.classList.remove("disabled");circ.classList.add("done");}
      if(conn) conn.classList.add("active");
      if(lbl){lbl.classList.remove("text-muted");lbl.classList.add("text-success");}
      if(cnt) cnt.textContent=d.total.toLocaleString()+" entries";
      // Activate step 5
      const nc=document.getElementById("step5-circle");
      const nl=document.getElementById("step5-label");
      if(nc) nc.classList.remove("disabled");
      if(nl){nl.classList.remove("text-muted");nl.classList.add("text-primary");}
      btn.innerHTML='<i class="bi bi-check-circle me-2"></i>Uploaded';
      setTimeout(()=>{ currentStep=5; showStep(5); },1200);
    } else {
      if(st) st.innerHTML=`<span class="text-danger">\u2717 ${d.error||"Upload failed"}</span>`;
      btn.disabled=false;
      btn.innerHTML='<i class="bi bi-upload me-2"></i>Upload &amp; Continue';
    }
  }catch(e){
    if(st) st.innerHTML=`<span class="text-danger">Error: ${e.message}</span>`;
    btn.disabled=false;
    btn.innerHTML='<i class="bi bi-upload me-2"></i>Upload &amp; Continue';
  }
}

async function loadSummary(){
  try{
    const r=await fetch("/api/summary"); const d=await r.json();
    const c=document.getElementById("summary-counts");
    if(c){
      const items=[{l:"RITMs",v:d.ritm_total,col:"primary"},{l:"Incidents",v:d.incident_total,col:"warning"},{l:"MACM",v:d.macm_total,col:"success"}];
      c.innerHTML=items.filter(i=>i.v!==null).map(i=>`<div class="col-auto"><div class="count-chip">
        <span class="badge bg-${i.col} rounded-pill">${i.l}</span>
        <span class="fs-5">${(i.v||0).toLocaleString()}</span></div></div>`).join("");
    }
  }catch(e){}
  populateReportMonths();
  loadEmailConfig();
}

async function uploadRef(){
  const inp=document.getElementById("ref-file");
  const st=document.getElementById("ref-status");
  if(!inp||!inp.files[0]){if(st)st.textContent="Select a file first.";return;}
  const btn=document.getElementById("btn-upload-ref");
  if(btn){btn.disabled=true; btn.innerHTML='<span class="spinner-border spinner-border-sm me-1"></span>Uploading…';}
  if(st) st.textContent="";
  const fd=new FormData(); fd.append("file",inp.files[0]);
  try{
    const r=await fetch("/upload/reference",{method:"POST",body:fd});
    const d=await r.json();
    if(d.success){
      if(st){st.textContent="✓ "+d.message; st.className="text-success small fw-semibold";}
      // Mark step 5 done
      const circ=document.getElementById("step5-circle");
      const conn=document.getElementById("conn5");
      const lbl=document.getElementById("step5-label");
      const cnt=document.getElementById("step5-count");
      if(circ){circ.classList.remove("disabled");circ.classList.add("done");}
      if(conn) conn.classList.add("active");
      if(lbl){lbl.classList.remove("text-muted");lbl.classList.add("text-success");}
      if(cnt) cnt.textContent="Uploaded";
      setTimeout(()=>{ currentStep=6; showStep(6); },1200);
    } else {
      if(st){st.textContent="✗ "+(d.error||"Upload failed"); st.className="text-danger small";}
      if(btn){btn.disabled=false; btn.innerHTML='<i class="bi bi-upload me-2"></i>Upload &amp; Continue';}
    }
  }catch(e){
    if(st){st.textContent="Error: "+e.message; st.className="text-danger small";}
    if(btn){btn.disabled=false; btn.innerHTML='<i class="bi bi-upload me-2"></i>Upload &amp; Continue';}
  }
}

// ── Populate month checkboxes in Step 5 ────────────────────────
async function populateReportMonths(){
  const box=document.getElementById("monthCheckboxes"); if(!box)return;
  try{
    const r=await fetch("/api/months");
    if(!r.ok){box.innerHTML=`<span class="text-danger small">Error loading months (${r.status})</span>`;return;}
    const d=await r.json();
    const months=d.months||[];
    if(!months.length){box.innerHTML='<span class="text-muted small">No months found in uploaded data.</span>';return;}
    box.innerHTML=months.map(m=>`
      <div class="form-check form-check-inline">
        <input class="form-check-input month-cb" type="checkbox" id="cb_${m}" value="${m}">
        <label class="form-check-label fw-semibold" for="cb_${m}">${m}</label>
      </div>`).join("");
  }catch(e){
    box.innerHTML=`<span class="text-danger small">Failed to load months: ${e.message}</span>`;
  }
}

function selectAllMonths(check){
  document.querySelectorAll(".month-cb").forEach(cb=>cb.checked=check);
}

function getSelectedMonths(){
  return [...document.querySelectorAll(".month-cb:checked")].map(cb=>cb.value);
}

function downloadWithMonth(){
  const months=getSelectedMonths();
  const qs=months.length?`?months=${months.map(encodeURIComponent).join(",")}`:"";
  window.open("/download/report"+qs, "_blank");
}

async function loadEmailConfig(){
  try{
    const r=await fetch("/api/email-config"); const d=await r.json();
    const set=(id,v)=>{const el=document.getElementById(id);if(el)el.value=v||"";};
    set("smtpHost",d.smtp_host); set("smtpPort",d.smtp_port);
    set("smtpUser",d.smtp_user); set("smtpPassword","");
    set("smtpRecipients",d.recipients_raw);
  }catch(e){}
}

async function saveEmailConfig(){
  const get=(id)=>document.getElementById(id)?.value||"";
  const payload={
    smtp_host:get("smtpHost"), smtp_port:parseInt(get("smtpPort"))||587,
    smtp_user:get("smtpUser"), smtp_password:get("smtpPassword")||undefined,
    recipients_raw:get("smtpRecipients")
  };
  if(!payload.smtp_password) delete payload.smtp_password;
  try{
    const r=await fetch("/api/email-config",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify(payload)});
    const d=await r.json();
    const st=document.getElementById("email-cfg-status");
    if(st){st.textContent=d.success?"✓ Saved":"✗ Failed"; st.className=d.success?"text-success":"text-danger";}
  }catch(e){}
}

// Poll for background email job result every 2 s (max 30 s = covers 15 s SMTP timeout)
async function _pollEmailJob(jobId,btn,st,btnLabel){
  const MAX=30000,TICK=2000; let elapsed=0;
  return new Promise(resolve=>{
    const t=setInterval(async()=>{
      elapsed+=TICK;
      try{
        const r=await fetch(`/api/email-status/${jobId}`);
        const d=await r.json();
        if(d.status==="done"){
          clearInterval(t);
          if(st){st.textContent=d.success?"✓ "+d.message:"✗ "+(d.error||"Failed");st.className=d.success?"text-success":"text-danger";}
          if(btn){btn.disabled=false;btn.innerHTML=btnLabel;}
          resolve(d);
        }
      }catch(_){}
      if(elapsed>=MAX){
        clearInterval(t);
        if(st){st.textContent="✗ Timed out — check your SMTP / network settings";st.className="text-danger";}
        if(btn){btn.disabled=false;btn.innerHTML=btnLabel;}
        resolve({success:false});
      }
    },TICK);
  });
}

async function sendEmailReport(){
  const btn=document.getElementById("btn-send-email");
  const st=document.getElementById("email-status");
  const months=getSelectedMonths();
  const lbl='<i class="bi bi-envelope me-2"></i>Send Email Only';
  if(btn){btn.disabled=true;btn.innerHTML='<span class="spinner-border spinner-border-sm me-1"></span>Sending…';}
  if(st){st.textContent="Connecting to mail server…";st.className="text-muted";}
  try{
    const r=await fetch("/api/send-email",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({months})});
    const d=await r.json();
    if(d.job_id){await _pollEmailJob(d.job_id,btn,st,lbl);}
    else{if(st){st.textContent=d.success?"✓ "+d.message:"✗ "+(d.error||"Failed");st.className=d.success?"text-success":"text-danger";}if(btn){btn.disabled=false;btn.innerHTML=lbl;}}
  }catch(e){
    if(st){st.textContent="✗ Network error — server unreachable";st.className="text-danger";}
    if(btn){btn.disabled=false;btn.innerHTML=lbl;}
  }
}

async function downloadAndSend(){
  const btn=document.getElementById("btn-dl-send");
  const st=document.getElementById("email-status");
  const months=getSelectedMonths();
  const lbl='<i class="bi bi-send me-2"></i>Download & Send Email';
  if(btn){btn.disabled=true;btn.innerHTML='<span class="spinner-border spinner-border-sm me-1"></span>Sending…';}
  if(st){st.textContent="Downloading and connecting to mail server…";st.className="text-muted";}
  const qs=months.length?`?months=${months.map(encodeURIComponent).join(",")}`:"";
  window.open("/download/report"+qs,"_blank");
  try{
    const r=await fetch("/api/send-email",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({months})});
    const d=await r.json();
    if(d.job_id){await _pollEmailJob(d.job_id,btn,st,lbl);}
    else{if(st){st.textContent=d.success?"✓ Email sent — "+d.message:"✗ Download done, email failed: "+(d.error||"Unknown error");st.className=d.success?"text-success":"text-danger";}if(btn){btn.disabled=false;btn.innerHTML=lbl;}}
  }catch(e){
    if(st){st.textContent="✗ Download done but email error: network issue";st.className="text-danger";}
    if(btn){btn.disabled=false;btn.innerHTML=lbl;}
  }
}


async function testEmailConnection(){
  const btn=document.getElementById("btn-test-email");
  const st=document.getElementById("email-cfg-status");
  if(btn){btn.disabled=true;btn.innerHTML='<span class="spinner-border spinner-border-sm me-1"></span>Testing…';}
  if(st){st.textContent="Connecting…";st.className="text-muted";}
  try{
    const r=await fetch("/api/test-email",{method:"POST"});
    const d=await r.json();
    if(d.job_id){
      const MAX=20000,TICK=2000;let elapsed=0;
      await new Promise(resolve=>{
        const t=setInterval(async()=>{
          elapsed+=TICK;
          try{
            const r2=await fetch(`/api/email-status/${d.job_id}`);
            const d2=await r2.json();
            if(d2.status==="done"){
              clearInterval(t);
              if(st){st.textContent=d2.success?"✓ Connection OK — test email sent":"✗ "+d2.error;st.className=d2.success?"text-success":"text-danger";}
              resolve();
            }
          }catch(_){}
          if(elapsed>=MAX){clearInterval(t);if(st){st.textContent="✗ Timed out — check SMTP settings";st.className="text-danger";}resolve();}
        },TICK);
      });
    } else {
      if(st){st.textContent=d.success?"✓ OK":"✗ "+(d.error||"Failed");st.className=d.success?"text-success":"text-danger";}
    }
  }catch(e){
    if(st){st.textContent="✗ Network error";st.className="text-danger";}
  }finally{
    if(btn){btn.disabled=false;btn.innerHTML='<i class="bi bi-plug me-1"></i>Test Connection';}
  }
}
