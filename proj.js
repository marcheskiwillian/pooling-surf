/* proj.js — Tendência de consumo para os próximos 6 meses.
   Linhas: histórico da base ativa por plano (app Benchmark de MVNOs), ajustado à contagem da planilha SURF, respeitando o filtro de MVNOs.
   GB por linha: média dos meses selecionados, com variação mensal editável. Faixa = menor e maior GB/linha observados. */
(function(){
'use strict';
var KEY='poolsurf.proj.v1';
function ls(k,v){ try{ if(v===undefined){ var r=localStorage.getItem(k); return r?JSON.parse(r):null; } if(v===null) localStorage.removeItem(k); else localStorage.setItem(k,JSON.stringify(v)); }catch(e){ return null; } }
function $(id){ return document.getElementById(id); }
function css(n){ return getComputedStyle(document.documentElement).getPropertyValue(n).trim(); }
var ST=Object.assign({base:'3m',manual:0,gpl:0},ls(KEY)||{});
var N=6;

function addMes(m,k){ var p=m.split('-'), y=+p[0], mo=+p[1]-1+k; y+=Math.floor(mo/12); mo=((mo%12)+12)%12; return y+'-'+String(mo+1).padStart(2,'0'); }
/* série histórica de linhas (base ativa A–F) do recorte atual */
function historico(M){
  var ML=window.MVNO_LINHAS; if(!ML) return null;
  var exc=M.exc||[], ks=['A','B','C','D','E','F'];
  return ML.meses.map(function(m,i){
    var tot=0; ks.forEach(function(k){ var t=(ML.total[k]||[])[i]||0, x=0; exc.forEach(function(e){ x+=((ML.mix[e]||{})[k]||[])[i]||0; }); tot+=Math.max(0,t-x); });
    return {m:m,l:tot};
  });
}
function taxa(serie,n){ /* crescimento médio mensal composto dos últimos n meses */
  var s=serie.slice(-(n+1)); if(s.length<2||!s[0].l) return 0; return Math.pow(s[s.length-1].l/s[0].l,1/(s.length-1))-1;
}
function calc(M){
  var A=window.PoolApp, S=A.SIM, T=M.T;
  var H=historico(M)||[];
  /* fator de ajuste: linhas da planilha SURF ÷ base ativa, nos meses em comum */
  var fat=[]; M.meses.forEach(function(mm){ var h=H.filter(function(x){return x.m===mm.m;})[0]; if(h&&h.l) fat.push(mm.l/h.l); });
  var aj=fat.length?fat.reduce(function(a,b){return a+b;},0)/fat.length:1;
  var HA=H.map(function(x){ return {m:x.m,l:x.l*aj}; });
  var t3=taxa(H,3), t6=taxa(H,6);
  var tx=ST.base==='6m'?t6:ST.base==='man'?ST.manual/100:t3;
  var gpls=M.meses.map(function(x){return x.gpl;}), g0=T.l?T.g/T.l:0;
  var gmin=Math.min.apply(null,gpls), gmax=Math.max.apply(null,gpls);
  var last=M.meses[M.meses.length-1], l0=last?last.l:T.l, m0=last?last.m:(H.length?H[H.length-1].m:'2026-08');
  var custoLinha=T.l?T.cost/T.l:0, commit=S.c/100*T.g;
  var P=[];
  for(var k=1;k<=N;k++){
    var l=l0*Math.pow(1+tx,k), gpl=g0*Math.pow(1+ST.gpl/100,k);
    var gb=l*gpl, lo=l*gmin*Math.pow(1+ST.gpl/100,k), hi=l*gmax*Math.pow(1+ST.gpl/100,k);
    var pool=A.fm(S)*l+S.p*Math.max(gb,commit); if(S.cap) pool=Math.min(pool,custoLinha*l);
    P.push({m:addMes(m0,k),l:l,gpl:gpl,gb:gb,lo:lo,hi:hi,fixo:custoLinha*l,pool:pool});
  }
  var hist=M.meses.map(function(x){ var pool=A.fm(S)*x.l+S.p*Math.max(x.g,commit); if(S.cap) pool=Math.min(pool,x.c); return {m:x.m,l:x.l,gb:x.g,gpl:x.gpl,fixo:x.c,pool:pool}; });
  return {H:HA,hist:hist,P:P,tx:tx,t3:t3,t6:t6,g0:g0,gmin:gmin,gmax:gmax,aj:aj,commit:commit};
}
function html(M){
  var A=window.PoolApp, nf=A.nf, Rk=A.Rk, pct=A.pct, ml=A.mesLbl, C=calc(M);
  var sg=function(x){ return (x>0?'+':x<0?'−':'')+nf(Math.abs(x*100),1)+'%'; };
  var Pn=C.P[N-1], h0=C.hist[C.hist.length-1]||{l:0,gb:0};
  var h='<section class="card" id="proj"><div class="sec-hd"><h2>Tendência para os próximos 6 meses</h2>'+
    '<p class="muted small">Projeção de '+ml(C.P[0].m)+' a '+ml(Pn.m)+' a partir do recorte atual. Linhas seguem o ritmo de crescimento da base; o consumo por linha parte da média observada. A faixa sombreada vai do menor ao maior GB/linha dos meses selecionados.</p></div>'+
    '<div class="gctl" style="grid-template-columns:1.4fr 1fr">'+
    '<div class="ctl"><span class="gl" style="margin-bottom:2px">Crescimento de linhas</span><div class="seggrp">'+
      [['3m','Ritmo dos últimos 3 meses ('+sg(C.t3)+'/mês)'],['6m','Últimos 6 meses ('+sg(C.t6)+'/mês)'],['man','Manual']].map(function(x){return '<button class="chip" data-pb="'+x[0]+'" aria-pressed="'+(ST.base===x[0])+'">'+x[1]+'</button>';}).join('')+
      (ST.base==='man'?'<input id="p-man" type="number" step="0.5" value="'+ST.manual+'" style="width:80px;font:inherit;font-family:var(--mono);padding:4px 8px;border:1px solid var(--axis);border-radius:6px;background:var(--surface-2);color:var(--ink)"> <span class="small">%/mês</span>':'')+'</div>'+
      '<span class="hint">Base ativa por plano do app Benchmark de MVNOs (fev a ago), sem as MVNOs desmarcadas.</span></div>'+
    '<div class="ctl"><label for="p-gpl">Variação do consumo por linha (%/mês)</label><input id="p-gpl" type="number" step="0.5" value="'+ST.gpl+'" style="width:110px;font:inherit;font-family:var(--mono);padding:6px 10px;border:1px solid var(--axis);border-radius:8px;background:var(--surface-2);color:var(--ink)"><span class="hint">Só 3 meses de consumo: pouco para medir tendência. Padrão 0% = média de '+nf(C.g0,2)+' GB/linha.</span></div></div>'+
    '<div class="kpis" style="grid-template-columns:repeat(4,1fr)">'+
      '<div class="kpi"><span class="v">'+nf(h0.l)+' → '+nf(Pn.l)+'</span><span class="l">linhas: '+ml(C.hist[C.hist.length-1].m)+' → '+ml(Pn.m)+'</span></div>'+
      '<div class="kpi"><span class="v">'+nf(Pn.gb/1000,1)+' mil GB</span><span class="l">consumo em '+ml(Pn.m)+' (faixa '+nf(Pn.lo/1000,1)+'–'+nf(Pn.hi/1000,1)+' mil)</span></div>'+
      '<div class="kpi"><span class="v">'+Rk(Pn.fixo)+'</span><span class="l">fatura SURF no modelo fixo em '+ml(Pn.m)+'</span></div>'+
      '<div class="kpi"><span class="v">'+Rk(Pn.pool)+'</span><span class="l">fatura no pool (cenário do simulador) em '+ml(Pn.m)+'</span></div></div>'+
    '<div class="pgrid"><div><h3>Linhas por mês</h3><div class="chartbox" id="pc-l"></div></div><div><h3>GB consumidos por mês</h3><div class="chartbox" id="pc-g"></div></div><div><h3>Fatura SURF: fixo × pool</h3><div class="chartbox" id="pc-f"></div></div></div>'+
    '<div class="legend"><span><i style="background:var(--used)"></i>Realizado</span><span><i style="background:transparent;border:2px dashed var(--used);height:6px"></i>Projeção</span><span><i style="background:var(--used);opacity:.18"></i>Faixa de consumo</span><span><i style="background:var(--contract)"></i>Modelo fixo</span></div>'+
    '<div class="tbl"><table style="min-width:760px"><thead><tr><th>Mês</th><th>Linhas</th><th>GB/linha</th><th>GB consumidos</th><th>Faixa de GB</th><th>Fatura fixa</th><th>Fatura pool</th><th>Diferença</th></tr></thead><tbody>'+
    C.hist.map(function(x){ return '<tr><td>'+ml(x.m)+' <span class="sub">real</span></td><td>'+nf(x.l)+'</td><td>'+nf(x.gpl,2)+'</td><td>'+nf(x.gb)+'</td><td>—</td><td>'+Rk(x.fixo)+'</td><td>'+Rk(x.pool)+'</td><td class="'+(x.pool<=x.fixo?'ok':'hot')+'">'+(x.pool<=x.fixo?'−':'+')+Rk(Math.abs(x.pool-x.fixo))+'</td></tr>'; }).join('')+
    C.P.map(function(x){ return '<tr class="proj"><td>'+ml(x.m)+' <span class="sub">projeção</span></td><td>'+nf(x.l)+'</td><td>'+nf(x.gpl,2)+'</td><td>'+nf(x.gb)+'</td><td>'+nf(x.lo)+'–'+nf(x.hi)+'</td><td>'+Rk(x.fixo)+'</td><td>'+Rk(x.pool)+'</td><td class="'+(x.pool<=x.fixo?'ok':'hot')+'">'+(x.pool<=x.fixo?'−':'+')+Rk(Math.abs(x.pool-x.fixo))+'</td></tr>'; }).join('')+
    '</tbody></table></div>'+
    '<p class="small muted">Fatura fixa = custo médio por linha da tabela SURF atual × linhas projetadas (mesmo mix de planos). Fatura pela proposta = custos fixos por linha × linhas + R$/GB × GB consumidos, respeitando o compromisso mínimo (~'+nf(C.commit/1000,1)+' mil GB/mês) e o teto do simulador. Linhas históricas ajustadas pela diferença entre a base ativa e a contagem da planilha SURF (fator '+nf(C.aj,3)+'). É uma projeção de tendência, não uma previsão de vendas.</p></section>';
  return h;
}
/* gráfico de linha com histórico sólido + projeção tracejada (+ faixa opcional) */
function chart(box,series,opt){
  if(!box) return; var W=box.clientWidth||360, H=210, pl=58, pr=12, pt=14, pb=28, A=window.PoolApp;
  var all=[]; series.forEach(function(s){ s.pts.forEach(function(p){ all.push(p.v); if(p.hi!=null) all.push(p.hi); if(p.lo!=null) all.push(p.lo); }); });
  var mx=Math.max.apply(null,all)*1.08, mn=opt.zero?Math.min(0,Math.min.apply(null,all)*1.1):(Math.min.apply(null,all)<0?Math.min.apply(null,all)*1.1:Math.min.apply(null,all)*0.9);
  var step=niceStep(mx-mn,4); mn=Math.floor(mn/step)*step; mx=Math.ceil(mx/step)*step;
  var n=opt.labels.length, x=function(i){return pl+(W-pl-pr)*(i+.5)/n;}, y=function(v){return pt+(H-pt-pb)*(1-(v-mn)/(mx-mn));};
  var s='<svg width="'+W+'" height="'+H+'" viewBox="0 0 '+W+' '+H+'" role="img" aria-label="'+opt.aria+'">';
  for(var v=mn;v<=mx+1e-9;v+=step) s+='<line x1="'+pl+'" x2="'+(W-pr)+'" y1="'+y(v)+'" y2="'+y(v)+'" stroke="'+css('--grid')+'"/><text x="'+(pl-6)+'" y="'+(y(v)+4)+'" text-anchor="end">'+opt.fmtAxis(v)+'</text>';
  var sep=opt.split; if(sep>0&&sep<n) s+='<line x1="'+((x(sep-1)+x(sep))/2)+'" x2="'+((x(sep-1)+x(sep))/2)+'" y1="'+pt+'" y2="'+(H-pb)+'" stroke="'+css('--axis')+'" stroke-dasharray="3 3"/><text x="'+((x(sep-1)+x(sep))/2+4)+'" y="'+(pt+10)+'" style="font-size:10px">projeção →</text>';
  var every=(W-pl-pr)/n>46?1:2; opt.labels.forEach(function(l,i){ if(every===1||i%2===((n-1)%2)) s+='<text x="'+x(i)+'" y="'+(H-8)+'" text-anchor="middle">'+l+'</text>'; });
  series.forEach(function(se){
    var c=css(se.cor);
    var band=se.pts.filter(function(p){return p.lo!=null;});
    if(band.length>1){ s+='<path d="'+band.map(function(p,i){return (i?'L':'M')+x(p.i)+' '+y(p.hi);}).join(' ')+' '+band.slice().reverse().map(function(p){return 'L'+x(p.i)+' '+y(p.lo);}).join(' ')+' Z" fill="'+c+'" opacity=".15"/>'; }
    var real=se.pts.filter(function(p){return !p.proj;}), pro=se.pts.filter(function(p,i){return p.proj;});
    if(real.length) s+='<path d="'+real.map(function(p,i){return (i?'L':'M')+x(p.i)+' '+y(p.v);}).join(' ')+'" fill="none" stroke="'+c+'" stroke-width="2"/>';
    if(pro.length){ var pp=(real.length?[real[real.length-1]]:[]).concat(pro); s+='<path d="'+pp.map(function(p,i){return (i?'L':'M')+x(p.i)+' '+y(p.v);}).join(' ')+'" fill="none" stroke="'+c+'" stroke-width="2" stroke-dasharray="6 4"/>'; }
    se.pts.forEach(function(p){ s+='<circle cx="'+x(p.i)+'" cy="'+y(p.v)+'" r="'+(p.proj?3.5:4)+'" fill="'+(p.proj?css('--surface'):c)+'" stroke="'+c+'" stroke-width="2"><title>'+se.nome+' · '+opt.labels[p.i]+': '+opt.fmtTip(p.v)+(p.lo!=null?' (faixa '+opt.fmtTip(p.lo)+' a '+opt.fmtTip(p.hi)+')':'')+'</title></circle>'; });
  });
  box.innerHTML=s+'</svg>';
}
function niceStep(r,n){ if(r<=0) return 1; var raw=r/n, p=Math.pow(10,Math.floor(Math.log10(raw))); return [1,2,2.5,5,10].map(function(m){return m*p;}).find(function(v){return raw<=v;}); }
function draw(M){
  var A=window.PoolApp, nf=A.nf, ml=A.mesLbl, C=calc(M);
  /* linhas: histórico longo (base ativa ajustada) + projeção */
  var lm=C.H.map(function(x){return x.m;}), iLast=lm.length;
  var lab=lm.concat(C.P.map(function(x){return x.m;})).map(ml);
  chart($('pc-l'),[{nome:'Linhas',cor:'--used',pts:C.H.map(function(x,i){return {i:i,v:x.l};}).concat(C.P.map(function(x,k){return {i:iLast+k,v:x.l,proj:true};}))}],
    {labels:lab,split:iLast,aria:'Linhas por mês com projeção',fmtAxis:function(v){return nf(v/1000,1)+' mil';},fmtTip:function(v){return nf(v)+' linhas';}});
  var hm=C.hist.length, lab2=C.hist.map(function(x){return ml(x.m);}).concat(C.P.map(function(x){return ml(x.m);}));
  chart($('pc-g'),[{nome:'GB consumidos',cor:'--used',pts:C.hist.map(function(x,i){return {i:i,v:x.gb};}).concat(C.P.map(function(x,k){return {i:hm+k,v:x.gb,lo:x.lo,hi:x.hi,proj:true};}))}],
    {labels:lab2,split:hm,aria:'GB consumidos por mês com projeção',fmtAxis:function(v){return nf(v/1000,0)+' mil';},fmtTip:function(v){return nf(v)+' GB';}});
  chart($('pc-f'),[{nome:'Fatura fixa',cor:'--contract',pts:C.hist.map(function(x,i){return {i:i,v:x.fixo};}).concat(C.P.map(function(x,k){return {i:hm+k,v:x.fixo,proj:true};}))},
                   {nome:'Fatura pool',cor:'--used',pts:C.hist.map(function(x,i){return {i:i,v:x.pool};}).concat(C.P.map(function(x,k){return {i:hm+k,v:x.pool,proj:true};}))}],
    {labels:lab2,split:hm,zero:true,aria:'Fatura SURF fixa e pool com projeção',fmtAxis:function(v){return nf(v/1000,0)+' mil';},fmtTip:function(v){return 'R$ '+nf(v);}});
}
function bind(M){
  var A=window.PoolApp;
  document.querySelectorAll('[data-pb]').forEach(function(b){ b.onclick=function(){ ST.base=b.dataset.pb; ls(KEY,ST); A.rerender(); }; });
  var pm=$('p-man'); if(pm) pm.addEventListener('change',function(){ var v=parseFloat(String(pm.value).replace(',','.')); if(isFinite(v)){ ST.manual=v; ls(KEY,ST); A.rerender(); } });
  var pg=$('p-gpl'); if(pg) pg.addEventListener('change',function(){ var v=parseFloat(String(pg.value).replace(',','.')); if(isFinite(v)){ ST.gpl=v; ls(KEY,ST); A.rerender(); } });
  draw(M);
}
function reset(){ ST={base:'3m',manual:0,gpl:0}; ls(KEY,null); }
window.PoolProj={html:html,bind:bind,draw:draw,calc:calc,reset:reset,chart:chart,get ST(){return ST;}};
})();
