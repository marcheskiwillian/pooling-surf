/* novos.js — Cenário "só linhas novas": o Plano por Consumo vale apenas para contratos novos.
   Novas MVNOs entram todo mês e ativam chips; a base atual fica na tabela fixa e fora desta conta.
   Mix de planos, consumo por linha, preço ao MVNO e custo do pool vêm da Nova grade e do simulador. */
(function(){
'use strict';
var KEY='poolsurf.novos.v1';
function ls(k,v){ try{ if(v===undefined){ var r=localStorage.getItem(k); return r?JSON.parse(r):null; } if(v===null) localStorage.removeItem(k); else localStorage.setItem(k,JSON.stringify(v)); }catch(e){ return null; } }
function $(id){ return document.getElementById(id); }
function esc(s){ return String(s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];}); }
var SC0=[
  {id:'pes',nome:'Pessimista',cor:'--s8',mv:3,ch:100,dur:12},
  {id:'rea',nome:'Realista',cor:'--s1',mv:4,ch:120,dur:12},
  {id:'oti',nome:'Otimista',cor:'--s3',mv:5,ch:150,dur:12}];
function ST0(){ return {hz:12,churn:null,sc:SC0.map(function(s){return Object.assign({},s);})}; }
var ST=(function(){ var s=ls(KEY), b=ST0(); if(!s) return b; b.hz=s.hz||12; b.churn=s.churn!=null?s.churn:null;
  (s.sc||[]).forEach(function(x){ var t=b.sc.filter(function(y){return y.id===x.id;})[0]; if(t){ ['mv','ch','dur'].forEach(function(k){ if(x[k]!=null) t[k]=+x[k]; }); } }); return b; })();

function addMes(m,k){ var p=m.split('-'), y=+p[0], mo=+p[1]-1+k; y+=Math.floor(mo/12); mo=((mo%12)+12)%12; return y+'-'+String(mo+1).padStart(2,'0'); }

/* médias por linha no mix de planos da base (recorte atual de MVNOs) */
function medias(M){
  var A=window.PoolApp, S=A.SIM, G=window.PoolGrade?window.PoolGrade.calc(M):null, l=0, u=0, pr=0, cf=0;
  if(G) G.rows.forEach(function(r){ l+=r.l; u+=r.l*r.uN; pr+=r.l*r.g.p; cf+=r.l*r.c0; });
  else M.planos.forEach(function(p){ l+=p.l; u+=p.g; pr+=p.l*p.mv; cf+=p.l*p.custo; });
  var pv=G?G.pv:null, pp=G?G.pp:0.65;
  var churnAuto=pv?(pp*pv.mP+(1-pp)*pv.mN)*100:0.9;
  return {u:l?u/l:0, preco:l?pr/l:0, fixa:l?cf/l:0, recL:A.recL(S), oneL:A.oneL(S), p:S.p, mb:S.mb, churnAuto:churnAuto, pp:pp};
}
/* referência: crescimento líquido real da base fora da ISUPER (app Benchmark de MVNOs) */
function historico(){
  var ML=window.MVNO_LINHAS; if(!ML||!ML.mix) return null; var n=ML.meses.length-1; if(n<1) return null;
  var tot0=0, tot1=0, best=null;
  Object.keys(ML.mix).forEach(function(mv){ if(/ISUPER/i.test(mv)) return; var a=0, b=0;
    Object.keys(ML.mix[mv]).forEach(function(k){ var arr=ML.mix[mv][k]||[]; a+=arr[0]||0; b+=arr[n]||0; });
    tot0+=a; tot1+=b; var g=(b-a)/n; if(!best||g>best.g) best={mv:mv,g:g,a:a,b:b}; });
  return {m0:ML.meses[0],m1:ML.meses[n],n:n,tot0:tot0,tot1:tot1,mes:(tot1-tot0)/n,best:best};
}
function calc(M){
  var A=window.PoolApp, md=medias(M), H=Math.max(1,ST.hz|0), ch=(ST.churn!=null?ST.churn:md.churnAuto)/100;
  var last=M.meses[M.meses.length-1], m0=last?last.m:'2026-08';
  var out=ST.sc.map(function(s){
    var L=0, acM=0, acE=0, mvs=0, P=[];
    for(var k=1;k<=H;k++){
      var cresc=Math.min(k,Math.max(1,s.dur))*s.mv, ativ=cresc*s.ch;   /* MVNOs ainda em fase de ativação × chips/mês */
      mvs=k*s.mv; L=L*(1-ch)+ativ;
      var gb=L*md.u, rec=L*(md.recL)+md.p*gb, one=ativ*md.oneL, consumo=rec+one;
      var fixa=L*md.fixa, receita=L*md.preco, margem=receita-consumo;
      acM+=margem; acE+=fixa-consumo;
      P.push({k:k,m:addMes(m0,k),mvs:mvs,ativ:ativ,l:L,gb:gb,consumo:consumo,one:one,fixa:fixa,receita:receita,margem:margem,acM:acM,acE:acE});
    }
    return {s:s,P:P,end:P[P.length-1]};
  });
  return {md:md,H:H,ch:ch,out:out,m0:m0};
}
function html(M){
  var A=window.PoolApp, nf=A.nf, R=A.R, Rk=A.Rk, ml=A.mesLbl, C=calc(M), md=C.md;
  var mk=function(i,k,st,lab){ return '<input type="number" min="0" step="'+st+'" data-nv="'+i+'" data-nk="'+k+'" value="'+ST.sc[i][k]+'" aria-label="'+lab+' '+ST.sc[i].nome+'" style="width:72px">'; };
  var h='<section class="card" id="novos"><div class="sec-hd"><h2>Cenário só linhas novas</h2>'+
    '<p class="muted small">O Plano por Consumo vale só para contratos novos: novas MVNOs entram todo mês e ativam chips; a base atual continua na tabela fixa e fica fora desta conta. Cada linha nova segue o mix de planos da base, o consumo por linha e o preço da Nova grade, e o custo do pool da nossa proposta ('+R(md.p,2)+'/GB).</p></div>';
  h+='<div class="gctl" style="grid-template-columns:minmax(420px,2fr) 1fr 1fr">'+
    '<div class="tbl" style="margin:0"><table class="gtab"><thead><tr><th>Cenário</th><th>Novas MVNOs/mês</th><th>Chips por MVNO/mês</th><th>Meses ativando</th></tr></thead><tbody>'+
      ST.sc.map(function(s,i){ return '<tr><td><i style="display:inline-block;width:10px;height:10px;border-radius:2px;background:var('+s.cor+');margin-right:6px"></i>'+s.nome+'</td><td>'+mk(i,'mv','1','Novas MVNOs por mês')+'</td><td>'+mk(i,'ch','10','Chips por MVNO por mês')+'</td><td>'+mk(i,'dur','1','Meses de ativação por MVNO')+'</td></tr>'; }).join('')+
    '</tbody></table></div>'+
    '<div class="ctl"><span class="gl" style="margin-bottom:2px">Horizonte</span><div class="seggrp">'+[12,24].map(function(n){return '<button class="chip" data-nh="'+n+'" aria-pressed="'+(C.H===n)+'">'+n+' meses</button>';}).join('')+'</div>'+
      '<button class="chip" id="nv-reset" style="margin-top:8px">Voltar aos cenários padrão</button></div>'+
    '<div class="ctl"><label for="nv-churn">Churn mensal (%)</label><input id="nv-churn" type="number" min="0" step="0.1" value="'+nf(C.ch*100,2).replace(',','.')+'" style="width:96px"><span class="hint">'+(ST.churn!=null?'Digitado. Calculado: '+nf(md.churnAuto,2)+'%':'Da tabela de retenção: '+nf(md.pp*100,0)+'% portadas a 0,50% e o resto a 1,65% ao mês')+'</span></div></div>';
  var HI=historico();
  if(HI) h+='<p class="small" style="margin:4px 0 10px;padding:8px 12px;border-left:3px solid var(--axis);background:var(--surface-2)">Referência real ('+ml(HI.m0)+' a '+ml(HI.m1)+', base ativa sem a ISUPER): '+nf(HI.tot0)+' → '+nf(HI.tot1)+' linhas, <b>'+(HI.mes>=0?'+':'')+nf(HI.mes)+' linhas líquidas/mês somando todas as MVNOs</b>. A MVNO que mais cresceu ('+esc(HI.best.mv)+') ganhou ~'+nf(HI.best.g)+' linhas/mês. No mês 1, o cenário pessimista ativa '+nf(C.out[0].P[0].ativ)+' chips; no mês '+C.H+', '+nf(C.out[0].P[C.H-1].ativ)+'.</p>';
  h+='<div class="kpis" style="grid-template-columns:repeat(3,1fr)">'+C.out.map(function(o){ var e=o.end;
    return '<div class="kpi" style="border-top:3px solid var('+o.s.cor+')"><span class="l" style="font-weight:600;color:var(--ink)">'+o.s.nome+' · '+o.s.mv+' MVNOs/mês × '+nf(o.s.ch)+' chips</span>'+
      '<span class="v">'+nf(e.l)+' linhas</span><span class="l">em '+ml(e.m)+' ('+nf(e.mvs)+' MVNOs novas)</span>'+
      '<span class="v" style="font-size:18px;margin-top:6px">'+Rk(e.margem)+'/mês</span><span class="l">margem bruta ISPX em '+ml(e.m)+' · acumulada '+Rk(e.acM)+'</span>'+
      '<span class="v" style="font-size:18px;margin-top:6px;color:var(--good-ink)">'+(e.acE>=0?'−':'+')+Rk(Math.abs(e.acE))+'</span><span class="l">pago à SURF no período vs. as mesmas linhas na tabela fixa</span></div>'; }).join('')+'</div>';
  h+='<div class="pgrid" style="grid-template-columns:repeat(2,1fr)"><div><h3>Linhas ativas no pool novo</h3><div class="chartbox" id="nv-l"></div></div><div><h3>Margem bruta ISPX por mês</h3><div class="chartbox" id="nv-m"></div></div></div>'+
    '<div class="legend">'+C.out.map(function(o){return '<span><i style="background:var('+o.s.cor+')"></i>'+o.s.nome+'</span>';}).join('')+'</div>';
  var marcos=[3,6,12,18,24].filter(function(k){return k<=C.H;});
  h+='<div class="tbl"><table style="min-width:980px"><thead><tr><th>Mês</th><th>Cenário</th><th>Linhas</th><th>GB/mês</th><th>Fatura SURF (consumo)</th><th>Mesmas linhas na tabela fixa</th><th>Receita ISPX</th><th>Margem/mês</th><th>Margem acumulada</th></tr></thead><tbody>'+
    marcos.map(function(k){ return C.out.map(function(o,j){ var x=o.P[k-1];
      return '<tr'+(j===0?' style="border-top:2px solid var(--axis)"':'')+'><td>'+(j===0?ml(x.m)+' <span class="sub">mês '+k+'</span>':'')+'</td><td><i style="display:inline-block;width:8px;height:8px;border-radius:2px;background:var('+o.s.cor+');margin-right:6px"></i>'+o.s.nome+'</td><td>'+nf(x.l)+'</td><td>'+nf(x.gb)+'</td><td>'+Rk(x.consumo)+'<div class="sub">'+Rk(x.one)+' de ativação</div></td><td>'+Rk(x.fixa)+'</td><td>'+Rk(x.receita)+'</td><td class="'+(x.margem>=0?'ok':'hot')+'">'+Rk(x.margem)+'</td><td><b>'+Rk(x.acM)+'</b></td></tr>'; }).join(''); }).join('')+
    '</tbody></table></div>'+
    '<p class="small muted">Por linha nova (mix da base atual'+(M.exc&&M.exc.length?', sem as MVNOs desmarcadas':'')+'): consumo '+nf(md.u,2)+' GB/mês · preço ao MVNO '+R(md.preco,2)+' (Nova grade) · recorrentes '+R(md.recL,2)+' (SIM, TFF, voz e SMS) · ativação '+R(md.oneL,2)+' paga no mês em que o chip entra (TFI + portabilidade em '+nf(md.pp*100,0)+'% das linhas) · tabela fixa equivalente '+R(md.fixa,2)+'. '+
    'Linhas do mês = linhas do mês anterior × (1 − churn) + ativações. Ativações = MVNOs ainda ativando × chips por MVNO; cada MVNO ativa pelo número de meses informado e depois só mantém a base. A base atual (sócios com margem zero para a ISPX e a ISUPER, que está saindo) fica fora desta conta. Cenários realista e otimista são hipóteses: ajuste os números.</p></section>';
  return h;
}
function draw(M){
  var A=window.PoolApp, nf=A.nf, Rk=A.Rk, ml=A.mesLbl, C=calc(M), PJ=window.PoolProj;
  if(!PJ||!PJ.chart) return;
  var lab=C.out[0].P.map(function(x){return ml(x.m);});
  PJ.chart($('nv-l'),C.out.map(function(o){ return {nome:o.s.nome,cor:o.s.cor,pts:o.P.map(function(x,i){return {i:i,v:x.l};})}; }),
    {labels:lab,split:0,zero:true,aria:'Linhas ativas no pool novo por cenário',fmtAxis:function(v){return v>=1000?nf(v/1000,0)+' mil':nf(v);},fmtTip:function(v){return nf(v)+' linhas';}});
  PJ.chart($('nv-m'),C.out.map(function(o){ return {nome:o.s.nome,cor:o.s.cor,pts:o.P.map(function(x,i){return {i:i,v:x.margem};})}; }),
    {labels:lab,split:0,zero:false,aria:'Margem bruta mensal por cenário',fmtAxis:function(v){return nf(v/1000,0)+' mil';},fmtTip:function(v){return Rk(v);}});
}
function bind(M){
  var A=window.PoolApp;
  document.querySelectorAll('[data-nv]').forEach(function(inp){ inp.addEventListener('change',function(){ var v=parseFloat(String(inp.value).replace(',','.')); if(!isFinite(v)||v<0) return; ST.sc[+inp.dataset.nv][inp.dataset.nk]=v; ls(KEY,ST); A.rerender(); }); });
  document.querySelectorAll('[data-nh]').forEach(function(b){ b.onclick=function(){ ST.hz=+b.dataset.nh; ls(KEY,ST); A.rerender(); }; });
  var c=$('nv-churn'); if(c) c.addEventListener('change',function(){ var v=parseFloat(String(c.value).replace(',','.')); if(isFinite(v)&&v>=0){ ST.churn=v; ls(KEY,ST); A.rerender(); } });
  var r=$('nv-reset'); if(r) r.onclick=function(){ ST=ST0(); ls(KEY,null); A.rerender(); };
  draw(M);
}
function reset(){ ST=ST0(); ls(KEY,null); }
window.PoolNovos={html:html,bind:bind,draw:draw,calc:calc,reset:reset,get ST(){return ST;}};
})();
