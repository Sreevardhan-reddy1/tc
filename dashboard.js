const PAL=["#0d6efd","#6f42c1","#d63384","#fd7e14","#ffc107","#198754","#20c997","#0dcaf0","#6c757d","#343a40","#0a58ca","#5a23c8","#ab296a","#ca6510","#cc9a06","#146c43","#13795b","#087990","#4d545a","#1a1d20"];
const TC={ritm:{bg:"rgba(13,110,253,.8)",br:"#0d6efd"},incident:{bg:"rgba(255,193,7,.85)",br:"#ffc107"},macm:{bg:"rgba(25,135,84,.8)",br:"#198754"}};

// ── State ──────────────────────────────────────────────────────
let _RAW={};
let _activeMonth="all";
let _charts={};              // id → Chart instance
let _dupesChartInst=null;    // duplicates chart instance
let _workTypesChartInst=null;        // work types doughnut
let _workTypesMonthlyChartInst=null; // work types monthly bar

// ── Helpers ────────────────────────────────────────────────────
function mergeKeys(...os){const s=new Set();os.forEach(o=>Object.keys(o||{}).forEach(k=>s.add(k)));return[...s].sort();}
function trunc(s,n){return s.length>n?s.slice(0,n-1)+"…":s;}
function esc(s){return String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}
function initials(n){const p=n.trim().split(/\s+/);return p.length>=2?(p[0][0]+p[p.length-1][0]).toUpperCase():n.slice(0,2).toUpperCase();}

function destroyChart(id){
  if(_charts[id]){_charts[id].destroy();delete _charts[id];}
  // also clear canvas
  const el=document.getElementById(id);
  if(el){const p=el.parentNode;p.removeChild(el);const c=document.createElement("canvas");c.id=id;c.style.maxHeight=el.style.maxHeight||"300px";p.appendChild(c);}
}

// ── Init ───────────────────────────────────────────────────────
function initDashboard(data){
  _RAW=data;
  buildMonthTabs();
  renderAll("all");
}

// ── Month tabs ─────────────────────────────────────────────────
function buildMonthTabs(){
  const R=_RAW.ritm||{},I=_RAW.incident||{},M=_RAW.macm||{};
  const months=mergeKeys(R.by_month,I.by_month,M.by_month);
  const container=document.getElementById("monthTabs"); if(!container)return;
  container.innerHTML=
    `<button class="btn btn-sm btn-primary month-tab me-1 mb-1" onclick="setMonth('all')" data-m="all">All Months</button>`
    +months.map(m=>`<button class="btn btn-sm btn-outline-secondary month-tab me-1 mb-1" onclick="setMonth('${m}')" data-m="${m}">${m}</button>`).join("");
}

function setMonth(m){
  _activeMonth=m;
  document.querySelectorAll(".month-tab").forEach(b=>{
    const active=b.dataset.m===m;
    b.className="btn btn-sm "+(active?"btn-primary":"btn-outline-secondary")+" month-tab me-1 mb-1";
  });
  // Update filter badge
  const lbl=document.getElementById("filterLabel");
  if(lbl) lbl.textContent=m==="all"?"All Months":m;
  renderAll(m);
}

// ── Get view-data for selected month ──────────────────────────
function viewOf(src, month){
  if(!src) return {total:0,by_team:{},by_assignee:{},by_month:{},by_team_assignee:{}};
  if(month==="all") return src;
  return {
    total:        src.by_month?.[month]||0,
    by_team:      src.by_month_team?.[month]||{},
    by_assignee:  src.by_month_assignee?.[month]||{},
    by_month:     src.by_month||{},          // monthly trend always full
    by_team_assignee: src.by_team_assignee||{}
  };
}

// ── Render everything ──────────────────────────────────────────
function renderAll(month){
  const R=viewOf(_RAW.ritm,month),I=viewOf(_RAW.incident,month),M=viewOf(_RAW.macm,month);
  const rt=R.total||0,it=I.total||0,mt=M.total||0;

  // KPIs — instant set (no animation on filter change to avoid flicker)
  ["kpi-grand-total","kpi-ritm","kpi-incident","kpi-macm"].forEach(id=>{
    const el=document.getElementById(id); if(el) el.textContent="—";
  });
  animN("kpi-grand-total",rt+it+mt); animN("kpi-ritm",rt); animN("kpi-incident",it); animN("kpi-macm",mt);

  // Destroy + redraw charts
  destroyChart("chartDoughnut"); renderDoughnut(rt,it,mt);

  // Monthly trend: always full data, highlight selected month
  destroyChart("chartMonthly");
  renderMonthly(_RAW.ritm||{},_RAW.incident||{},_RAW.macm||{}, month);

  // Tables + async cards
  buildIncRitmMacmSection();
  buildMonthlyTable();
  buildWorkTypesCard(month);
  buildTop10SlowSection();
  buildDuplicatesTable(month);
}

// ── Animated counter ───────────────────────────────────────────
function animN(id,target){
  const el=document.getElementById(id); if(!el)return;
  const dur=600,start=performance.now();
  function step(ts){const p=Math.min((ts-start)/dur,1);el.textContent=Math.round(p*target).toLocaleString();if(p<1)requestAnimationFrame(step);}
  requestAnimationFrame(step);
}

// ── Chart 1: Doughnut ──────────────────────────────────────────
function renderDoughnut(r,i,m){
  const ctx=document.getElementById("chartDoughnut"); if(!ctx)return;
  _charts["chartDoughnut"]=new Chart(ctx,{type:"doughnut",data:{labels:["RITM","Incident","MACM"],
    datasets:[{data:[r,i,m],backgroundColor:[TC.ritm.bg,TC.incident.bg,TC.macm.bg],
    borderColor:[TC.ritm.br,TC.incident.br,TC.macm.br],borderWidth:2,hoverOffset:10}]},
    options:{responsive:true,plugins:{legend:{position:"bottom"},
    tooltip:{callbacks:{label:c=>` ${c.label}: ${c.parsed.toLocaleString()}`}}}}});
}

// ── Chart 2: Team stacked bar ──────────────────────────────────
function renderTeamBar(R,I,M){
  const ctx=document.getElementById("chartTeam"); if(!ctx)return;
  const teams=mergeKeys(R.by_team,I.by_team,M.by_team);
  const sorted=teams.map(t=>({t,tot:(R.by_team?.[t]||0)+(I.by_team?.[t]||0)+(M.by_team?.[t]||0)}))
    .sort((a,b)=>b.tot-a.tot).slice(0,15);
  if(!sorted.length){
    ctx.closest(".card-body").innerHTML='<p class="text-muted text-center py-4">No team data for this period.</p>'; return;
  }
  _charts["chartTeam"]=new Chart(ctx,{type:"bar",data:{labels:sorted.map(x=>trunc(x.t,30)),datasets:[
    {label:"RITM",    data:sorted.map(x=>R.by_team?.[x.t]||0),backgroundColor:TC.ritm.bg,    borderColor:TC.ritm.br,    borderWidth:1},
    {label:"Incident",data:sorted.map(x=>I.by_team?.[x.t]||0),backgroundColor:TC.incident.bg,borderColor:TC.incident.br,borderWidth:1},
    {label:"MACM",    data:sorted.map(x=>M.by_team?.[x.t]||0),backgroundColor:TC.macm.bg,    borderColor:TC.macm.br,    borderWidth:1}]},
    options:{indexAxis:"y",responsive:true,plugins:{legend:{position:"top"},tooltip:{mode:"index",intersect:false}},
    scales:{x:{stacked:true,grid:{color:"#f0f0f0"}},y:{stacked:true,ticks:{font:{size:11}}}}}});
}

// ── Chart 3: Monthly trend (always full, highlight selected) ───
function renderMonthly(R,I,M,selectedMonth){
  const ctx=document.getElementById("chartMonthly"); if(!ctx)return;
  const months=mergeKeys(R.by_month,I.by_month,M.by_month);
  if(!months.length){
    ctx.closest(".card-body").innerHTML='<p class="text-muted text-center py-4">No date data found in uploaded files.</p>'; return;
  }
  // Point radius: highlight the selected month
  const ptRadius=(type)=>months.map(m=>(m===selectedMonth&&selectedMonth!=="all")?8:3);
  const ptBg=(br)=>months.map(m=>(m===selectedMonth&&selectedMonth!=="all")?"#fff":br);

  _charts["chartMonthly"]=new Chart(ctx,{type:"bar",
    data:{labels:months,datasets:[
      {type:"line",label:"RITM",    data:months.map(m=>R.by_month?.[m]||0),borderColor:TC.ritm.br,    backgroundColor:"rgba(13,110,253,.08)",fill:true,tension:.35,pointRadius:ptRadius("r"),pointBackgroundColor:ptBg(TC.ritm.br),    pointBorderWidth:2,order:1},
      {type:"line",label:"Incident",data:months.map(m=>I.by_month?.[m]||0),borderColor:TC.incident.br,backgroundColor:"rgba(255,193,7,.08)",fill:true,tension:.35,pointRadius:ptRadius("i"),pointBackgroundColor:ptBg(TC.incident.br),pointBorderWidth:2,order:2},
      {type:"line",label:"MACM",    data:months.map(m=>M.by_month?.[m]||0),borderColor:TC.macm.br,    backgroundColor:"rgba(25,135,84,.08)",fill:true,tension:.35,pointRadius:ptRadius("m"),pointBackgroundColor:ptBg(TC.macm.br),    pointBorderWidth:2,order:3}
    ]},
    options:{responsive:true,plugins:{legend:{position:"top"},tooltip:{mode:"index",intersect:false},
      annotation: selectedMonth!=="all" ? {} : {}
    },
    scales:{x:{grid:{color:"#f0f0f0"}},y:{beginAtZero:true,grid:{color:"#f0f0f0"}}}}});
}

// ── Chart 4: Top assignees ─────────────────────────────────────
function renderAssignee(R,I,M){
  const ctx=document.getElementById("chartAssignee"); if(!ctx)return;
  const all=mergeKeys(R.by_assignee,I.by_assignee,M.by_assignee);
  const sorted=all.map(a=>({a,tot:(R.by_assignee?.[a]||0)+(I.by_assignee?.[a]||0)+(M.by_assignee?.[a]||0)}))
    .sort((a,b)=>b.tot-a.tot).slice(0,12);
  if(!sorted.length){
    ctx.closest(".card-body").innerHTML='<p class="text-muted text-center py-4">No assignee data for this period.</p>'; return;
  }
  _charts["chartAssignee"]=new Chart(ctx,{type:"bar",data:{labels:sorted.map(x=>trunc(x.a,22)),
    datasets:[{label:"Tickets",data:sorted.map(x=>x.tot),backgroundColor:sorted.map((_,i)=>PAL[i%PAL.length]),borderWidth:1}]},
    options:{indexAxis:"y",responsive:true,plugins:{legend:{display:false},
    tooltip:{callbacks:{label:c=>` ${c.parsed.x} tickets`}}},
    scales:{x:{beginAtZero:true,grid:{color:"#f0f0f0"}},y:{ticks:{font:{size:10}}}}}});
}

// ── Team table ─────────────────────────────────────────────────
function buildTeamTable(R,I,M){
  const tbody=document.getElementById("teamTbody"); if(!tbody)return;
  const teams=mergeKeys(R.by_team,I.by_team,M.by_team);
  const rows=teams.map(t=>{const r=R.by_team?.[t]||0,i=I.by_team?.[t]||0,m=M.by_team?.[t]||0;return{t,r,i,m,tot:r+i+m};}).sort((a,b)=>b.tot-a.tot);
  if(!rows.length){tbody.innerHTML=`<tr><td colspan="5" class="text-center text-muted py-3">No data for this period.</td></tr>`;return;}
  tbody.innerHTML=rows.map(x=>`<tr>
    <td class="fw-semibold">${esc(x.t)}</td>
    <td class="text-center"><span class="badge bg-primary rounded-pill">${x.r.toLocaleString()}</span></td>
    <td class="text-center"><span class="badge bg-warning text-dark rounded-pill">${x.i.toLocaleString()}</span></td>
    <td class="text-center"><span class="badge bg-success rounded-pill">${x.m.toLocaleString()}</span></td>
    <td class="text-center fw-bold">${x.tot.toLocaleString()}</td></tr>`).join("");
}

// ── Assignee table ─────────────────────────────────────────────
function buildAssigneeTable(R,I,M){
  const tbody=document.getElementById("assigneeTbody"); if(!tbody)return;
  const people=mergeKeys(R.by_assignee,I.by_assignee,M.by_assignee);
  const rows=people.map(a=>{const r=R.by_assignee?.[a]||0,i=I.by_assignee?.[a]||0,m=M.by_assignee?.[a]||0;return{a,r,i,m,tot:r+i+m};}).sort((a,b)=>b.tot-a.tot);
  if(!rows.length){tbody.innerHTML=`<tr><td colspan="5" class="text-center text-muted py-3">No data for this period.</td></tr>`;return;}
  tbody.innerHTML=rows.map(x=>`<tr>
    <td><div class="d-flex align-items-center gap-2">
      <div class="rounded-circle bg-secondary text-white d-flex align-items-center justify-content-center" style="width:28px;height:28px;font-size:.7rem;flex-shrink:0">${initials(x.a)}</div>
      <span>${esc(x.a)}</span></div></td>
    <td class="text-center"><span class="badge bg-primary rounded-pill">${x.r.toLocaleString()}</span></td>
    <td class="text-center"><span class="badge bg-warning text-dark rounded-pill">${x.i.toLocaleString()}</span></td>
    <td class="text-center"><span class="badge bg-success rounded-pill">${x.m.toLocaleString()}</span></td>
    <td class="text-center fw-bold">${x.tot.toLocaleString()}</td></tr>`).join("");
}

// ── Monthly summary table (always full, not filtered) ──────────
function buildMonthlyTable(){
  const tbody=document.getElementById("monthlyTbody"); if(!tbody)return;
  const R=_RAW.ritm||{}, I=_RAW.incident||{}, M=_RAW.macm||{};
  const months=mergeKeys(R.by_month,I.by_month,M.by_month);
  if(!months.length){tbody.innerHTML=`<tr><td colspan="5" class="text-muted text-center py-3">No monthly data found.</td></tr>`;return;}
  const rows=months.map(mo=>{
    const r=R.by_month?.[mo]||0,i=I.by_month?.[mo]||0,m=M.by_month?.[mo]||0;
    return{mo,r,i,m,tot:r+i+m};
  });
  tbody.innerHTML=rows.map(x=>`<tr class="${_activeMonth===x.mo?"table-active fw-bold":""}">
    <td>${esc(x.mo)}</td>
    <td class="text-center"><span class="badge bg-primary rounded-pill">${x.r.toLocaleString()}</span></td>
    <td class="text-center"><span class="badge bg-warning text-dark rounded-pill">${x.i.toLocaleString()}</span></td>
    <td class="text-center"><span class="badge bg-success rounded-pill">${x.m.toLocaleString()}</span></td>
    <td class="text-center fw-bold">${x.tot.toLocaleString()}</td></tr>`).join("");
  // Totals row
  const tr=rows.reduce((a,x)=>a+x.r,0),ti=rows.reduce((a,x)=>a+x.i,0),tm=rows.reduce((a,x)=>a+x.m,0);
  tbody.innerHTML+=`<tr class="table-dark fw-bold">
    <td>TOTAL</td>
    <td class="text-center">${tr.toLocaleString()}</td>
    <td class="text-center">${ti.toLocaleString()}</td>
    <td class="text-center">${tm.toLocaleString()}</td>
    <td class="text-center">${(tr+ti+tm).toLocaleString()}</td></tr>`;
}

// ── Work Types Report ──────────────────────────────────────────
// Normalise any month key → YYYY-MM for consistent sorting
function _toYYYYMM(m){
  if(!m) return "";
  const iso=m.match(/^(\d{4})-(\d{2})$/);
  if(iso) return m;
  const MON={Jan:"01",Feb:"02",Mar:"03",Apr:"04",May:"05",Jun:"06",Jul:"07",Aug:"08",Sep:"09",Oct:"10",Nov:"11",Dec:"12"};
  const p=m.split("-");
  if(p.length===2 && MON[p[0]]) return `${p[1]}-${MON[p[0]]}`;
  return m;
}

async function buildWorkTypesCard(month){
  const wrap=document.getElementById("workTypesTableWrap");
  const badge=document.getElementById("wtMonthBadge");
  const monthlyWrap=document.getElementById("wtMonthlyChartWrap");
  if(!wrap) return;
  const mo=month||"all";
  const label=mo==="all"?"All Months":mo;
  if(badge) badge.textContent=label;
  if(_workTypesChartInst){_workTypesChartInst.destroy();_workTypesChartInst=null;}
  if(_workTypesMonthlyChartInst){_workTypesMonthlyChartInst.destroy();_workTypesMonthlyChartInst=null;}
  try{
    const res=await fetch("/api/work-types?month="+encodeURIComponent(mo));
    const d=await res.json();
    const wt=d.work_types||[];
    const total=wt.reduce((s,w)=>s+w.count,0);

    // ── table ─────────────────────────────────────────────────
    let html=`<div class="table-responsive"><table class="table table-sm table-bordered align-middle mb-0">
<thead class="table-dark"><tr>
  <th style="width:2rem">#</th>
  <th>Work Type</th>
  <th>Short Description</th>
  <th class="text-center" style="width:5rem">Count</th>
  <th class="text-center" style="width:4rem">%</th>
</tr></thead><tbody>`;
    wt.forEach((w,i)=>{
      const pct=total?((w.count/total)*100).toFixed(1):"0.0";
      html+=`<tr>
  <td class="text-muted small">${i+1}</td>
  <td><span class="me-1" style="display:inline-block;width:10px;height:10px;border-radius:50%;background:${esc(w.color)}"></span>${esc(w.label)}</td>
  <td class="text-muted small">${esc(w.description)}</td>
  <td class="text-center fw-bold">${w.count.toLocaleString()}</td>
  <td class="text-center text-muted">${pct}%</td>
</tr>`;
    });
    html+=`<tr class="table-secondary fw-bold">
  <td colspan="3">Total</td>
  <td class="text-center">${total.toLocaleString()}</td>
  <td class="text-center">100%</td>
</tr></tbody></table></div>`;
    if(!d.has_team_efforts){
      html+=`<p class="text-muted small mt-2 mb-0"><i class="bi bi-info-circle me-1"></i>Upload a <strong>Team Efforts sheet</strong> on the upload page to include JIRA (Operational-JIRA) tickets.</p>`;
    }
    wrap.innerHTML=html;

    // ── doughnut chart ────────────────────────────────────────
    const ctx=document.getElementById("chartWorkTypes");
    if(ctx && wt.some(w=>w.count>0)){
      _workTypesChartInst=new Chart(ctx,{
        type:"doughnut",
        data:{
          labels:wt.map(w=>w.label),
          datasets:[{data:wt.map(w=>w.count),backgroundColor:wt.map(w=>w.color),borderWidth:2,hoverOffset:8}]
        },
        options:{
          responsive:true,
          plugins:{
            legend:{position:"bottom",labels:{font:{size:11},boxWidth:12}},
            title:{display:true,text:"Work Types \u2014 "+label,font:{size:13}},
            tooltip:{callbacks:{label:c=>` ${c.label}: ${c.parsed.toLocaleString()}`}}
          }
        }
      });
    }

    // ── monthly stacked bar chart ─────────────────────────────
    if(monthlyWrap){
      const R=_RAW.ritm||{}, I=_RAW.incident||{}, M=_RAW.macm||{};
      const J=_RAW.jira||{};
      // Collect all months from all sources, normalise to YYYY-MM for sorting
      const allM=mergeKeys(R.by_month,I.by_month,M.by_month,J.by_month||{});
      const sortedM=allM.slice().sort((a,b)=>_toYYYYMM(a)<_toYYYYMM(b)?-1:1);
      if(sortedM.length>0){
        monthlyWrap.style.display="";
        // Build per-month counts for each type
        function getM(src,mk){
          const bm=src.by_month||{};
          if(mk in bm) return bm[mk]||0;
          // try both formats
          const alt=_toYYYYMM(mk);
          return bm[alt]||bm[mk]||0;
        }
        // For JIRA: by_month keys are Mon-YYYY; convert sortedM (YYYY-MM) to match
        function getJira(mk){
          const bm=J.by_month||{};
          if(mk in bm) return bm[mk]||0;
          // mk is YYYY-MM, jira uses Mon-YYYY
          const p=mk.split("-");
          if(p.length===2){
            const MON=["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
            const abbr=`${MON[parseInt(p[1])-1]}-${p[0]}`;
            return bm[abbr]||0;
          }
          return 0;
        }
        const ctx2=document.getElementById("chartWorkTypesMonthly");
        if(ctx2){
          _workTypesMonthlyChartInst=new Chart(ctx2,{
            type:"bar",
            data:{
              labels:sortedM,
              datasets:[
                {label:"RITM",     data:sortedM.map(m=>getM(R,m)), backgroundColor:"rgba(13,110,253,.8)",  borderRadius:3},
                {label:"Incident", data:sortedM.map(m=>getM(I,m)), backgroundColor:"rgba(255,193,7,.85)",  borderRadius:3},
                {label:"MACM",     data:sortedM.map(m=>getM(M,m)), backgroundColor:"rgba(25,135,84,.8)",   borderRadius:3},
                {label:"JIRA",     data:sortedM.map(m=>getJira(m)),backgroundColor:"rgba(253,126,20,.85)", borderRadius:3}
              ]
            },
            options:{
              responsive:true,
              plugins:{
                legend:{position:"top",labels:{font:{size:11},boxWidth:12}},
                title:{display:true,text:"Monthly Work Type Breakdown",font:{size:13}},
                tooltip:{mode:"index",intersect:false}
              },
              scales:{
                x:{stacked:true,ticks:{font:{size:10},maxRotation:40}},
                y:{stacked:true,beginAtZero:true,ticks:{stepSize:10},
                   title:{display:true,text:"Tickets"}}
              }
            }
          });
        }
      } else {
        monthlyWrap.style.display="none";
      }
    }
  }catch(e){
    if(wrap) wrap.innerHTML=`<div class="alert alert-warning mb-0">Could not load work types: ${e.message}</div>`;
  }
}

// ── Top 10 Slowest Tickets ─────────────────────────────────────
let _top10RitmChart=null, _top10IncChart=null;

async function buildTop10SlowSection(){
  const ritmWrap=document.getElementById("top10RitmWrap");
  const incWrap =document.getElementById("top10IncidentWrap");
  if(!ritmWrap || !incWrap) return;

  const loading=`<div class="text-center py-3 text-muted small">
    <div class="spinner-border spinner-border-sm me-1"></div> Loading\u2026</div>`;
  ritmWrap.innerHTML=loading; incWrap.innerHTML=loading;

  try{
    const res=await fetch("/api/top10-slow");
    const d=await res.json();
    _renderTop10Panel(ritmWrap,"top10RitmChart",  d.ritm||[],    "RITM",    "#0d6efd");
    _renderTop10Panel(incWrap, "top10IncChart",   d.incident||[], "Incident","#ffc107");
  }catch(e){
    const err=`<div class="alert alert-warning mb-0 small">Could not load: ${e.message}</div>`;
    ritmWrap.innerHTML=err; incWrap.innerHTML=err;
  }
}

function _renderTop10Panel(wrap, chartId, tickets, label, color){
  if(!tickets||!tickets.length){
    wrap.innerHTML=`<p class="text-muted small text-center py-3">
      No duration data. Upload a ${label} file that has both an opened and a closed/resolved date column.</p>`;
    return;
  }

  // Sort by duration desc (already sorted from API, but be safe)
  tickets=tickets.slice().sort((a,b)=>(b.duration_days||0)-(a.duration_days||0));
  const maxDur=tickets[0].duration_days||1;

  // ── Table ───────────────────────────────────────────────────
  let html=`<div class="table-responsive">
<table class="table table-sm table-hover align-middle mb-2" style="font-size:.82rem">
<thead><tr style="background:#f0f4ff">
  <th style="width:2rem">#</th>
  <th>Ticket #</th>
  <th>Short Description</th>
  <th class="text-end" style="width:7rem">Duration</th>
</tr></thead><tbody>`;

  tickets.forEach((t,idx)=>{
    const dur   = t.duration_days||0;
    const pct   = Math.round(dur/maxDur*100);
    const num   = esc(t.number||"—");
    const desc  = esc((t.short_description||"No description").slice(0,70));
    const barClr= label==="RITM" ? "#0d6efd" : "#ffc107";
    const txtClr= label==="RITM" ? "#fff"     : "#212529";
    html+=`<tr>
  <td class="text-muted">${idx+1}</td>
  <td><span class="fw-semibold" style="color:${barClr}">${num}</span></td>
  <td>
    <div class="text-truncate" style="max-width:260px" title="${desc}">${desc}</div>
    <div class="progress mt-1" style="height:5px;border-radius:3px">
      <div class="progress-bar" style="width:${pct}%;background:${barClr}"></div>
    </div>
  </td>
  <td class="text-end">
    <span class="badge rounded-pill" style="background:${barClr};color:${txtClr};font-size:.78rem">
      ${dur.toLocaleString()} day${dur===1?"":"s"}
    </span>
  </td>
</tr>`;
  });
  html+=`</tbody></table></div>`;

  // ── Horizontal bar chart ────────────────────────────────────
  html+=`<canvas id="${chartId}" style="max-height:220px"></canvas>`;
  wrap.innerHTML=html;

  const ctx=document.getElementById(chartId); if(!ctx) return;
  const isRitm=label==="RITM";
  if(isRitm){ if(_top10RitmChart){_top10RitmChart.destroy();_top10RitmChart=null;} }
  else       { if(_top10IncChart) {_top10IncChart.destroy(); _top10IncChart=null;} }

  const chartInst=new Chart(ctx,{
    type:"bar",
    data:{
      labels:tickets.map(t=>{
        const n=t.number||"?";
        const d=(t.short_description||"").slice(0,28);
        return `${n} — ${d}`;
      }),
      datasets:[{
        label:"Duration (days)",
        data:tickets.map(t=>t.duration_days||0),
        backgroundColor: isRitm ? "rgba(13,110,253,.75)" : "rgba(255,193,7,.85)",
        borderColor:     isRitm ? "#0d6efd"              : "#ffc107",
        borderWidth:1,
        borderRadius:4
      }]
    },
    options:{
      indexAxis:"y",
      responsive:true,
      plugins:{
        legend:{display:false},
        tooltip:{callbacks:{label:c=>" "+c.parsed.x.toLocaleString()+" days"}}
      },
      scales:{
        x:{beginAtZero:true,grid:{color:"#f0f0f0"},
           title:{display:true,text:"Days to Resolve",font:{size:10}}},
        y:{ticks:{font:{size:9},maxRotation:0}}
      }
    }
  });
  if(isRitm) _top10RitmChart=chartInst;
  else       _top10IncChart =chartInst;
}

// ── Duplicates Report ──────────────────────────────────────────
let _dupesMonthlyChartInst=null;
let _dupCachedData=null,_dupMonths=[];

async function buildDuplicatesTable(month){
  const wrap=document.getElementById("dupesWrap"); if(!wrap)return;
  wrap.innerHTML=`<div class="text-center py-3"><div class="spinner-border spinner-border-sm text-danger"></div> Loading\u2026</div>`;
  try{
    const res=await fetch("/api/duplicates?month=all");
    const d=await res.json();
    _dupCachedData=d;
    // Build sorted months list from monthly_summary
    const ms=d.monthly_summary||{};
    _dupMonths=Object.keys(ms).sort((a,b)=>_monthSortKey(a)-_monthSortKey(b));
    // Init dropdowns once (dataset.initialized guards re-init on global tab change)
    const sd=document.getElementById("dupStartMonth");
    const ed=document.getElementById("dupEndMonth");
    if(sd && !sd.dataset.initialized){
      sd.dataset.initialized="1";
      const opts=_dupMonths.map(m=>`<option value="${esc(m)}">${esc(m)}</option>`).join("");
      sd.innerHTML=`<option value="">-- Start --</option>`+opts;
      ed.innerHTML=`<option value="">-- End --</option>`+opts;
      // Default: From=first 2025 month, To=present month
      sd.value=_dupMonths.find(_is2025)||_dupMonths[0]||"";
      const ps=_presentSortKey();
      let endVal=_dupMonths[_dupMonths.length-1]||"";
      for(let i=_dupMonths.length-1;i>=0;i--){
        if(_monthSortKey(_dupMonths[i])<=ps){endVal=_dupMonths[i];break;}
      }
      ed.value=endVal;
    }
    _renderDuplicates();
  }catch(e){
    wrap.innerHTML=`<div class="alert alert-warning mb-0">Could not load duplicates: ${e.message}</div>`;
  }
}

function filterDuplicates(){ _renderDuplicates(); }

function resetDupFilter(){
  const sd=document.getElementById("dupStartMonth");
  const ed=document.getElementById("dupEndMonth");
  if(sd && ed){
    sd.value=_dupMonths.find(_is2025)||_dupMonths[0]||"";
    const ps=_presentSortKey();
    let endVal=_dupMonths[_dupMonths.length-1]||"";
    for(let i=_dupMonths.length-1;i>=0;i--){
      if(_monthSortKey(_dupMonths[i])<=ps){endVal=_dupMonths[i];break;}
    }
    ed.value=endVal;
  }
  _renderDuplicates();
}

function _renderDuplicates(){
  const wrap=document.getElementById("dupesWrap"); if(!wrap||!_dupCachedData)return;
  if(_dupesChartInst){_dupesChartInst.destroy();_dupesChartInst=null;}
  if(_dupesMonthlyChartInst){_dupesMonthlyChartInst.destroy();_dupesMonthlyChartInst=null;}
  const start=document.getElementById("dupStartMonth")?.value||"";
  const end  =document.getElementById("dupEndMonth")?.value||"";
  const allGroups=_dupCachedData.duplicates||[];
  const fullMS=_dupCachedData.monthly_summary||{};
  // Filter groups: include group if any of its by_month keys falls within the range
  let groups=allGroups;
  let monthlySummary=fullMS;
  if(start||end){
    groups=allGroups.filter(g=>{
      const gm=Object.keys(g.by_month||{});
      if(!gm.length)return true;
      return gm.some(m=>{
        const k=_monthSortKey(m);
        return (!start||k>=_monthSortKey(start))&&(!end||k<=_monthSortKey(end));
      });
    });
    monthlySummary={};
    Object.keys(fullMS).forEach(m=>{
      const k=_monthSortKey(m);
      if((!start||k>=_monthSortKey(start))&&(!end||k<=_monthSortKey(end)))
        monthlySummary[m]=fullMS[m];
    });
  }
  const _chips=document.getElementById("dupMonthChips");
  if(!groups.length){
    if(_chips) _chips.innerHTML='';
    wrap.innerHTML=`<p class="text-muted text-center py-3">No duplicates found for the selected date range.</p>`;
    return;
  }
  const totalTix=groups.reduce((s,g)=>s+g.count,0);
  // Sorted months for monthly chart
  const chartMonths=Object.keys(monthlySummary).sort((a,b)=>_monthSortKey(a)-_monthSortKey(b));
  // Update per-month count chips in header
  if(_chips){
    _chips.innerHTML=chartMonths.map(m=>`<span class="badge bg-white text-danger border border-danger" style="font-size:.78em">${esc(m)}<span class="ms-1 fw-bold">${monthlySummary[m]}</span></span>`).join('');
  }
  // ── Summary header ────────────────────────────────────────────
  let html=`<div class="d-flex align-items-center gap-3 mb-3 flex-wrap">
  <span class="text-muted">Found <strong>${groups.length}</strong> duplicate group(s) \u00b7
    <strong>${totalTix}</strong> tickets involved</span>
  <span class="badge bg-primary">RITM</span><span class="text-muted small">+</span>
  <span class="badge bg-warning text-dark">Incident</span><span class="text-muted small">+</span>
  <span class="badge bg-success">MACM</span>
</div>`;
  // ── Two-column chart area ─────────────────────────────────────
  html+=`<div class="row g-3 mb-3">
  <div class="${chartMonths.length?'col-md-6':'col-12'}">
    <div class="card border-0 shadow-sm p-2">
      <canvas id="dupesChart" style="max-height:240px"></canvas>
    </div>
  </div>
  ${chartMonths.length?`<div class="col-md-6">
    <div class="card border-0 shadow-sm p-2">
      <canvas id="dupesMonthlyChart" style="max-height:240px"></canvas>
    </div>
  </div>`:''}
</div>`;
  // ── Table ─────────────────────────────────────────────────────
  html+=`<div class="table-responsive"><table class="table table-sm table-hover table-bordered align-middle mb-0">
<thead class="table-dark"><tr>
  <th style="width:2rem">#</th>
  <th>Short Description</th>
  <th class="text-center">Months</th>
  <th class="text-center">RITM</th>
  <th class="text-center">INC</th>
  <th class="text-center">MACM</th>
  <th class="text-center">Count</th>
</tr></thead><tbody>`;
  groups.forEach((g,idx)=>{
    const rBadge=g.ritm_count?`<span class="badge bg-primary">${g.ritm_count}</span>`:`<span class="text-muted">\u2014</span>`;
    const iBadge=g.incident_count?`<span class="badge bg-warning text-dark">${g.incident_count}</span>`:`<span class="text-muted">\u2014</span>`;
    const mBadge=(g.macm_count||0)?`<span class="badge bg-success">${g.macm_count}</span>`:`<span class="text-muted">\u2014</span>`;
    const gMonths=g.by_month||{};
    // Show only month badges within the selected range
    const filteredGMonths=Object.keys(gMonths).filter(m=>{
      const k=_monthSortKey(m);
      return (!start||k>=_monthSortKey(start))&&(!end||k<=_monthSortKey(end));
    }).sort((a,b)=>_monthSortKey(a)-_monthSortKey(b));
    const monthBadges=filteredGMonths.map(mo2=>`<span class="badge bg-info text-dark me-1" style="font-size:.72em">${esc(mo2)}<span class="ms-1">${gMonths[mo2]}</span></span>`).join("");
    html+=`<tr>
  <td class="text-muted small">${idx+1}</td>
  <td>${esc(g.description)}</td>
  <td>${monthBadges||'<span class="text-muted">\u2014</span>'}</td>
  <td class="text-center">${rBadge}</td>
  <td class="text-center">${iBadge}</td>
  <td class="text-center">${mBadge}</td>
  <td class="text-center"><span class="badge bg-danger">${g.count}</span></td>
</tr>`;
  });
  html+=`</tbody></table></div>`;
  wrap.innerHTML=html;
  // ── Chart 1: top groups stacked bar ──────────────────────────
  const top=groups.slice(0,12);
  const descLabels=top.map(g=>g.description.length>28?g.description.slice(0,25)+'\u2026':g.description);
  _dupesChartInst=new Chart(document.getElementById("dupesChart"),{
    type:'bar',
    data:{
      labels:descLabels,
      datasets:[
        {label:'RITM',    data:top.map(g=>g.ritm_count),     backgroundColor:'rgba(13,110,253,.80)', borderRadius:3},
        {label:'Incident',data:top.map(g=>g.incident_count), backgroundColor:'rgba(255,193,7,.85)',  borderRadius:3},
        {label:'MACM',    data:top.map(g=>g.macm_count||0),  backgroundColor:'rgba(25,135,84,.80)',  borderRadius:3}
      ]
    },
    options:{responsive:true,plugins:{legend:{position:'top'},
      title:{display:true,text:'Top Duplicate Groups',font:{size:12}}},
      scales:{x:{stacked:true,ticks:{font:{size:9},maxRotation:40}},
              y:{stacked:true,beginAtZero:true,ticks:{stepSize:1}}}}
  });
  // ── Chart 2: monthly duplicate count ─────────────────────────
  if(chartMonths.length){
    _dupesMonthlyChartInst=new Chart(document.getElementById("dupesMonthlyChart"),{
      type:'bar',
      data:{
        labels:chartMonths,
        datasets:[{
          label:'Duplicate Tickets',
          data:chartMonths.map(m=>monthlySummary[m]||0),
          backgroundColor:'rgba(220,53,69,.75)',
          borderColor:'rgba(220,53,69,1)',
          borderWidth:1,
          borderRadius:4
        }]
      },
      options:{responsive:true,plugins:{legend:{display:false},
        title:{display:true,text:'Duplicates by Month',font:{size:12}}},
        scales:{x:{ticks:{font:{size:10}}},
                y:{beginAtZero:true,ticks:{stepSize:1},
                   title:{display:true,text:'Duplicate Tickets'}}}}
    });
  }
}

// ── Drill-down ─────────────────────────────────────────────────
let _drillData={};
function buildTeamSelect(R,I,M,month){
  _drillData={R,I,M,month};
  const sel=document.getElementById("teamDrillSel"); if(!sel)return;
  const teams=mergeKeys(R.by_team,I.by_team,M.by_team).sort();
  sel.innerHTML=`<option value="">— Pick a team —</option>`+teams.map(t=>`<option value="${esc(t)}">${esc(t)}</option>`).join("");
  document.getElementById("drillOut").innerHTML="";
}

function renderDrill(){
  const team=document.getElementById("teamDrillSel")?.value;
  const type=document.getElementById("typeDrillSel")?.value||"all";
  const out=document.getElementById("drillOut"); if(!out)return;
  if(!team){out.innerHTML="";return;}
  const {R,I,M}=_drillData;
  const srcMap={ritm:R,incident:I,macm:M};
  const types=type==="all"?["ritm","incident","macm"]:[type];
  const amap={};
  types.forEach(t=>{
    // Use by_team_assignee from full raw data (it's not month-filtered)
    const fullSrc=_RAW[t]||{};
    const ta=_activeMonth==="all"
      ? fullSrc.by_team_assignee?.[team]||{}
      : (fullSrc.by_month_team?.[_activeMonth]?.[team] ? {} : {});  // simplified
    // Fall back to current view's by_team for member names
    Object.entries(srcMap[t]?.by_assignee||{}).forEach(([a,c])=>{if(c>0)amap[a]=(amap[a]||0)+c;});
  });
  const rows=Object.entries(amap).sort((a,b)=>b[1]-a[1]);
  if(!rows.length){out.innerHTML=`<p class="text-muted">No detail for this selection.</p>`;return;}
  const total=rows.reduce((s,[,c])=>s+c,0);
  out.innerHTML=`<table class="table table-sm"><thead class="table-light"><tr><th>Assignee</th><th class="text-center">Tickets</th><th>Share</th></tr></thead><tbody>`
    +rows.map(([a,c])=>`<tr><td>${esc(a)}</td><td class="text-center fw-bold">${c}</td>
    <td style="min-width:130px"><div class="progress" style="height:12px;border-radius:6px"><div class="progress-bar bg-primary" style="width:${Math.round(c/total*100)}%">${Math.round(c/total*100)}%</div></div></td></tr>`).join("")
    +`</tbody><tfoot class="table-light"><tr><td class="fw-bold">Total</td><td class="text-center fw-bold">${total}</td><td></td></tr></tfoot></table>`;
}

// ── Table filter & sort ────────────────────────────────────────
function filterTable(tid,q){
  const qL=q.toLowerCase();
  document.querySelectorAll("#"+tid+" tbody tr").forEach(r=>{r.style.display=r.textContent.toLowerCase().includes(qL)?"":"none";});
}
function sortTable(tid,ci){
  const tbl=document.getElementById(tid); if(!tbl)return;
  const tbody=tbl.querySelector("tbody");
  const rows=[...tbody.querySelectorAll("tr:not(.table-dark)")];
  const asc=tbl.dataset.sc===String(ci)&&tbl.dataset.sd==="asc";
  tbl.dataset.sc=ci; tbl.dataset.sd=asc?"desc":"asc";
  rows.sort((a,b)=>{
    const ta=a.cells[ci]?.textContent.trim()||"",tb=b.cells[ci]?.textContent.trim()||"";
    const na=parseFloat(ta.replace(/,/g,"")),nb=parseFloat(tb.replace(/,/g,""));
    if(!isNaN(na)&&!isNaN(nb))return asc?nb-na:na-nb;
    return asc?tb.localeCompare(ta):ta.localeCompare(tb);
  });
  rows.forEach(r=>tbody.insertBefore(r,tbody.querySelector(".table-dark")));
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
  // Open in new tab so errors are visible and the wizard page stays open
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

async function testEmailConnection(){
  const btn=document.getElementById("btn-test-email");
  const st=document.getElementById("email-cfg-status");
  if(btn){btn.disabled=true;btn.innerHTML='<span class="spinner-border spinner-border-sm me-1"></span>Testing…';}
  if(st){st.textContent="";st.className="";}
  try{
    const r=await fetch("/api/test-email",{method:"POST"});
    const d=await r.json();
    if(d.job_id){
      // poll
      const MAX=20000,TICK=2000;let elapsed=0;
      await new Promise(resolve=>{
        const t=setInterval(async()=>{
          elapsed+=TICK;
          try{const r2=await fetch(`/api/email-status/${d.job_id}`);const d2=await r2.json();
            if(d2.status==="done"){clearInterval(t);
              if(st){st.textContent=d2.success?"✓ Connection OK — test email sent":"✗ "+d2.error;st.className=d2.success?"text-success":"text-danger";}
              resolve();}}catch(_){}
          if(elapsed>=MAX){clearInterval(t);if(st){st.textContent="✗ Timed out";st.className="text-danger";}resolve();}
        },TICK);
      });
    } else {
      if(st){st.textContent=d.success?"✓ OK":"✗ "+(d.error||"Failed");st.className=d.success?"text-success":"text-danger";}
    }
  }catch(e){if(st){st.textContent="✗ Network error";st.className="text-danger";}}
  finally{if(btn){btn.disabled=false;btn.innerHTML='<i class="bi bi-plug me-1"></i>Test Connection';}}
}

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

// ── INC / RITM / MACM — date-range filter ─────────────────────
let _irmAllMonths=[];

// Handles both "YYYY-MM" (2025-09) and "Mon-YYYY" (Sep-2025)
function _monthSortKey(m){
  if(!m) return 0;
  const iso=m.match(/^(\d{4})-(\d{2})$/);
  if(iso) return parseInt(iso[1])*100+parseInt(iso[2]);   // YYYY-MM path
  const MON={Jan:1,Feb:2,Mar:3,Apr:4,May:5,Jun:6,Jul:7,Aug:8,Sep:9,Oct:10,Nov:11,Dec:12};
  const [mon,yr]=m.split("-");
  return (parseInt(yr)||0)*100+(MON[mon]||0);             // Mon-YYYY path
}

// Returns true if key belongs to year 2025 (handles both formats)
function _is2025(m){
  return /^2025-/.test(m) || /-2025$/.test(m);
}

// Current month as sort key (format-independent comparison)
function _presentSortKey(){
  const now=new Date();
  return now.getFullYear()*100+(now.getMonth()+1);
}

function buildIncRitmMacmSection(){
  const R=_RAW.ritm||{}, I=_RAW.incident||{}, M=_RAW.macm||{};
  // Sort chronologically regardless of key format
  const allMonths=mergeKeys(R.by_month,I.by_month,M.by_month)
    .sort((a,b)=>_monthSortKey(a)-_monthSortKey(b));
  _irmAllMonths=allMonths;

  const sd=document.getElementById("irmStartMonth");
  const ed=document.getElementById("irmEndMonth");
  if(sd&&allMonths.length){
    const opts=allMonths.map(m=>`<option value="${esc(m)}">${esc(m)}</option>`).join("");
    sd.innerHTML=`<option value="">-- Start --</option>`+opts;
    ed.innerHTML=`<option value="">-- End --</option>`+opts;

    // From: first 2025 month in data; fallback to earliest overall
    sd.value=allMonths.find(_is2025)||allMonths[0];

    // To: last data month that is <= today
    const ps=_presentSortKey();
    let endVal=allMonths[allMonths.length-1];
    for(let i=allMonths.length-1;i>=0;i--){
      if(_monthSortKey(allMonths[i])<=ps){endVal=allMonths[i];break;}
    }
    ed.value=endVal;
  }
  filterIncRitmMacm();
}

function filterIncRitmMacm(){
  const start=document.getElementById("irmStartMonth")?.value||"";
  const end  =document.getElementById("irmEndMonth")?.value||"";
  let filtered=_irmAllMonths;
  if(start) filtered=filtered.filter(m=>_monthSortKey(m)>=_monthSortKey(start));
  if(end)   filtered=filtered.filter(m=>_monthSortKey(m)<=_monthSortKey(end));
  _renderIrmTable(filtered);
}

function resetIrmFilter(){
  const sd=document.getElementById("irmStartMonth");
  const ed=document.getElementById("irmEndMonth");
  if(sd) sd.value=_irmAllMonths.find(_is2025)||_irmAllMonths[0]||"";
  if(ed){
    const ps=_presentSortKey();
    let endVal=_irmAllMonths[_irmAllMonths.length-1]||"";
    for(let i=_irmAllMonths.length-1;i>=0;i--){
      if(_monthSortKey(_irmAllMonths[i])<=ps){endVal=_irmAllMonths[i];break;}
    }
    ed.value=endVal;
  }
  filterIncRitmMacm();
}

function _renderIrmTable(months){
  const R=_RAW.ritm||{}, I=_RAW.incident||{}, M=_RAW.macm||{};
  const tbody=document.getElementById("incRitmMacmTbody"); if(!tbody)return;

  if(!months.length){
    tbody.innerHTML=`<tr><td colspan="5" class="text-center text-muted py-3">No data for selected range.</td></tr>`;
    destroyChart("chartIncRitmMacm"); return;
  }

  const rows=months.map(mo=>{
    const inc=I.by_month?.[mo]||0, ritm=R.by_month?.[mo]||0, macm=M.by_month?.[mo]||0;
    return{mo,inc,ritm,macm,tot:inc+ritm+macm};
  });
  const tInc=rows.reduce((a,x)=>a+x.inc,0);
  const tRitm=rows.reduce((a,x)=>a+x.ritm,0);
  const tMacm=rows.reduce((a,x)=>a+x.macm,0);

  tbody.innerHTML=rows.map(x=>`<tr>
    <td class="text-end fw-semibold px-3">${esc(x.mo)}</td>
    <td class="text-center">${x.inc.toLocaleString()}</td>
    <td class="text-center">${x.ritm.toLocaleString()}</td>
    <td class="text-center">${x.macm.toLocaleString()}</td>
    <td class="text-center fw-bold">${x.tot.toLocaleString()}</td>
  </tr>`).join("")
  +`<tr style="background:#1565c0;color:#fff;font-weight:bold">
    <td class="px-3">TicketWiseTotal</td>
    <td class="text-center">${tInc.toLocaleString()}</td>
    <td class="text-center">${tRitm.toLocaleString()}</td>
    <td class="text-center">${tMacm.toLocaleString()}</td>
    <td class="text-center">—</td>
  </tr>`;

  const _dataLabels={
    id:"_dlIRM",
    afterDatasetsDraw(chart){
      const{ctx}=chart;
      chart.data.datasets.forEach((_ds,di)=>{
        chart.getDatasetMeta(di).data.forEach((bar,ji)=>{
          const v=chart.data.datasets[di].data[ji];
          if(!v)return;
          ctx.save();
          ctx.font="bold 9px sans-serif";
          ctx.fillStyle="#222";
          ctx.textAlign="center";
          ctx.textBaseline="bottom";
          ctx.fillText(v.toLocaleString(), bar.x, bar.y-2);
          ctx.restore();
        });
      });
    }
  };

  destroyChart("chartIncRitmMacm");
  const ctx=document.getElementById("chartIncRitmMacm"); if(!ctx)return;
  _charts["chartIncRitmMacm"]=new Chart(ctx,{
    type:"bar",
    plugins:[_dataLabels],
    data:{
      labels:months,
      datasets:[
        {label:"INC",           data:rows.map(x=>x.inc), backgroundColor:"rgba(158,158,158,0.75)",borderColor:"#9e9e9e",borderWidth:1,borderRadius:3},
        {label:"RITM",          data:rows.map(x=>x.ritm),backgroundColor:"rgba(33,33,33,0.82)",  borderColor:"#212121",borderWidth:1,borderRadius:3},
        {label:"MACM",          data:rows.map(x=>x.macm),backgroundColor:"rgba(41,182,246,0.78)",borderColor:"#29b6f6",borderWidth:1,borderRadius:3},
        {label:"MonthWiseTotal",data:rows.map(x=>x.tot), backgroundColor:"rgba(96,125,139,0.65)",borderColor:"#607d8b",borderWidth:1,borderRadius:3}
      ]
    },
    options:{
      responsive:true,
      layout:{padding:{top:22}},
      plugins:{
        legend:{position:"bottom",labels:{usePointStyle:true,padding:14,font:{size:11}}},
        tooltip:{mode:"index",intersect:false}
      },
      scales:{
        x:{grid:{color:"#f0f0f0"},ticks:{font:{size:11}}},
        y:{beginAtZero:true,grid:{color:"#f0f0f0"}}
      }
    }
  });
}
