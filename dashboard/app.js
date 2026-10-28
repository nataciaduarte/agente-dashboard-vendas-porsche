const data=window.PORSCHE_DATA||[];
const Q=window.DATA_QUALITY||{};
const $=id=>document.getElementById(id);
const colors=['#d5001c','#0d0d0f','#b68b42','#74716b','#bcb8b0','#e4e1dc'];
const countBy=(arr,key)=>arr.reduce((a,r)=>(a[r[key]]=(a[r[key]]||0)+1,a),{});
const entries=o=>Object.entries(o).sort((a,b)=>b[1]-a[1]);
const money=n=>new Intl.NumberFormat('pt-BR',{style:'currency',currency:'USD',maximumFractionDigits:0}).format(n);
const num=n=>new Intl.NumberFormat('pt-BR').format(n);
const plural=n=>`${num(n)} venda${n===1?'':'s'}`;

function fillSelect(id,key,sort='alpha'){
  const vals=[...new Set(data.map(r=>r[key]))].filter(Boolean);
  vals.sort(sort==='desc'?(a,b)=>b-a:(a,b)=>String(a).localeCompare(String(b)));
  vals.forEach(v=>$(id).add(new Option(v,v)));
}
fillSelect('modelFilter','model');fillSelect('yearFilter','year','desc');fillSelect('cityFilter','city');fillSelect('paymentFilter','payment');

function selected(){return {model:$('modelFilter').value,year:$('yearFilter').value,city:$('cityFilter').value,payment:$('paymentFilter').value,from:$('dateFrom').value,to:$('dateTo').value}}
function filtered(){const f=selected();return data.filter(r=>(!f.model||r.model===f.model)&&(!f.year||r.year===f.year)&&(!f.city||r.city===f.city)&&(!f.payment||r.payment===f.payment)&&(!f.from||(r.date&&r.date>=f.from))&&(!f.to||(r.date&&r.date<=f.to)))}
function render(){
  const all=filtered(), cancelled=all.filter(r=>r.status==='Cancelled').length, rows=all.filter(r=>r.status!=='Cancelled'), total=rows.reduce((s,r)=>s+r.price,0), models=entries(countBy(rows,'model'));
  $('salesKpi').textContent=num(rows.length);$('salesNote').textContent=cancelled?`${num(cancelled)} cancelada${cancelled>1?'s':''} fora do total`:'nenhuma cancelada no recorte';$('revenueKpi').textContent=money(total);$('ticketKpi').textContent=rows.length?money(total/rows.length):'—';
  const tied=models.filter(m=>models[0]&&m[1]===models[0][1]);$('leaderKpi').textContent=tied.length>1?tied.map(m=>m[0]).join(' e '):(models[0]?.[0]||'—');$('leaderNote').textContent=models[0]?`${plural(models[0][1])}${tied.length>1?' cada, empatados':''} · ${Math.round(models[0][1]/rows.length*100)}% do total`:'Sem dados';
  const f=selected();$('periodLabel').textContent=f.from||f.to?`${f.from||'início'} — ${f.to||'hoje'}`:'Todos os períodos';
  renderTrend(rows);renderCity(rows);renderYear(rows);renderPayment(rows);renderStates(rows);
}
let trendMetric='sales';
function renderTrend(rows){
  const monthly={};rows.filter(r=>r.date).forEach(r=>{const m=r.date.slice(0,7);monthly[m]??={sales:0,revenue:0};monthly[m].sales++;monthly[m].revenue+=r.price});
  const points=Object.entries(monthly).sort(([a],[b])=>a.localeCompare(b)).map(([month,v])=>({month,value:v[trendMetric]}));
  if(!points.length){$('trendChart').innerHTML='<div class="trend-empty">Nenhuma data válida neste recorte.</div>';$('trendDelta').textContent='—';$('trendSummary').textContent='Ajuste os filtros para visualizar a evolução.';return}
  const W=1100,H=220,pad={l:34,r:20,t:24,b:34},max=Math.max(...points.map(p=>p.value),1),span=Math.max(points.length-1,1);
  const xy=points.map((p,i)=>({x:pad.l+i*(W-pad.l-pad.r)/span,y:pad.t+(max-p.value)/max*(H-pad.t-pad.b),...p}));
  const line=xy.map((p,i)=>`${i?'L':'M'}${p.x.toFixed(1)},${p.y.toFixed(1)}`).join(' '),area=`${line} L${xy.at(-1).x},${H-pad.b} L${xy[0].x},${H-pad.b} Z`;
  const labels=xy.map((p,i)=>{const show=points.length<=10||i%Math.ceil(points.length/8)===0||i===points.length-1;if(!show)return'';const d=new Date(`${p.month}-02T00:00:00`),label=d.toLocaleDateString('pt-BR',{month:'short',year:'2-digit'});return `<text class="trend-label" x="${p.x}" y="${H-8}" text-anchor="middle">${label}</text>`}).join('');
  const dots=xy.map(p=>`<circle class="trend-point" cx="${p.x}" cy="${p.y}" r="4"><title>${p.month}: ${trendMetric==='sales'?num(p.value):money(p.value)}</title></circle>${points.length<=8?`<text class="trend-value" x="${p.x}" y="${p.y-10}" text-anchor="middle">${trendMetric==='sales'?num(p.value):money(p.value)}</text>`:''}`).join('');
  $('trendChart').innerHTML=`<svg viewBox="0 0 ${W} ${H}" preserveAspectRatio="none" role="img"><defs><linearGradient id="trendFill" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#d5001c" stop-opacity=".18"/><stop offset="1" stop-color="#d5001c" stop-opacity="0"/></linearGradient></defs>${[0,.5,1].map(t=>`<line class="trend-grid" x1="${pad.l}" y1="${pad.t+t*(H-pad.t-pad.b)}" x2="${W-pad.r}" y2="${pad.t+t*(H-pad.t-pad.b)}"/>`).join('')}<path class="trend-area" d="${area}"/><path class="trend-line" d="${line}"/>${dots}${labels}</svg>`;
  const sum=points.reduce((a,p)=>a+p.value,0),avg=sum/points.length;$('trendDelta').textContent=trendMetric==='sales'?`${avg.toFixed(1).replace('.',',')} vendas por mês com venda`:`${money(avg)} por mês com venda`;$('trendDelta').style.color='#0d0d0f';$('trendSummary').textContent=`${points[0].month} → ${points.at(-1).month} · pico em ${points.reduce((a,b)=>a.value>b.value?a:b).month}`;
}
function renderCity(rows){
  const cities=entries(countBy(rows,'city')).slice(0,6);const topModels=entries(countBy(rows,'model')).slice(0,6).map(x=>x[0]);
  $('cityModelChart').innerHTML=cities.map(([city,total])=>{const cityRows=rows.filter(r=>r.city===city);const counts=countBy(cityRows,'model');return `<div class="stack-row"><span class="city">${city}</span><div class="stack">${topModels.map((m,i)=>`<i title="${m}: ${counts[m]||0}" style="width:${(counts[m]||0)/total*100}%;background:${colors[i]}"></i>`).join('')}</div><b>${total}</b></div>`}).join('')+`<div class="legend">${topModels.map((m,i)=>`<span><i style="background:${colors[i]}"></i>${m}</span>`).join('')}</div>`;
}
function renderYear(rows){const years=entries(countBy(rows,'year')).slice(0,5),max=years[0]?.[1]||1;$('yearHero').innerHTML=years[0]?`<strong>${years[0][0]}</strong><span>${plural(years[0][1])}<br>no recorte atual</span>`:'<strong>—</strong>';$('yearBars').innerHTML=years.map(([y,c])=>`<div class="mini-row"><b>${y}</b><div class="mini-track"><i style="width:${c/max*100}%"></i></div><span>${c}</span></div>`).join('')}
function renderPayment(rows){const vals=entries(countBy(rows,'payment')),total=rows.length||1;let cursor=0;const stops=vals.map(([k,v],i)=>{const start=cursor;cursor+=v/total*100;return `${colors[i%colors.length]} ${start}% ${cursor}%`});$('paymentChart').innerHTML=`<div class="donut" style="background:conic-gradient(${stops.join(',')||'#eee 0 100%'})"></div><div class="pay-legend">${vals.map(([k,v],i)=>`<div><span><i style="background:${colors[i%colors.length]}"></i>${k}</span><b>${Math.round(v/total*100)}%</b></div>`).join('')}</div>`}
function renderStates(rows){const states=entries(countBy(rows,'state')),grand=rows.length||1;$('insightCount').textContent=`${states.length} estados`;$('insights').innerHTML=states.map(([state,total])=>{const sr=rows.filter(r=>r.state===state),lead=entries(countBy(sr,'model'))[0],revenue=sr.reduce((s,r)=>s+r.price,0),share=Math.round(total/grand*100);return `<div class="insight"><h3>${state}</h3><strong>${plural(total)}</strong><p><span class="share">${share}%</span> do volume · receita ${money(revenue)}</p><p>Modelo líder: ${lead?.[0]||'—'} (${plural(lead?.[1]||0)})</p></div>`}).join('')||'<p>Não há dados para este recorte.</p>'}
['modelFilter','yearFilter','cityFilter','paymentFilter','dateFrom','dateTo'].forEach(id=>$(id).addEventListener('change',render));$('reset').addEventListener('click',()=>{['modelFilter','yearFilter','cityFilter','paymentFilter','dateFrom','dateTo'].forEach(id=>$(id).value='');render()});document.querySelectorAll('[data-trend]').forEach(btn=>btn.addEventListener('click',()=>{trendMetric=btn.dataset.trend;document.querySelectorAll('[data-trend]').forEach(b=>b.classList.toggle('active',b===btn));render()}));render();

['Received','Impossible','Future','Cancelled','Used'].forEach(k=>{$('q'+k).textContent=num(Q[k.toLowerCase()]??0)});$('qRef').textContent=Q.reference?Q.reference.split('-').reverse().join('/'):'—';
