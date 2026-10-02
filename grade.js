/* grade.js — Nova grade de planos: GB, preço ao MVNO e custo do pool editáveis.
   Custo real por linha = consumo médio × R$/GB do pool (+ custo fixo por linha, se houver). */
(function(){
'use strict';
var KEY_G='poolsurf.grade.v1', KEY_SC='poolsurf.gradesc.v1';
function ls(k,v){ try{ if(v===undefined){ var r=localStorage.getItem(k); return r?JSON.parse(r):null; } if(v===null) localStorage.removeItem(k); else localStorage.setItem(k,JSON.stringify(v)); }catch(e){ return null; } }
function $(id){ return document.getElementById(id); }
function esc(s){ return String(s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];}); }

var PRESETS={
  inter:{nome:'Intermediária',g:{A:[8,22.90],B:[13,26.90],C:[20,33.90],D:[27,39.90],E:[40,54.90],F:[65,74.90]}},
  mod:{nome:'Moderada',g:{A:[8,23.90],B:[12,26.90],C:[20,34.90],D:[25,41.90],E:[40,57.90],F:[60,79.90]}},
  agr:{nome:'Agressiva',g:{A:[10,22.90],B:[16,26.90],C:[25,33.90],D:[35,39.90],E:[50,54.90],F:[80,74.90]}}
};
var G=ls(KEY_G)||clonePreset('inter');
var SC=ls(KEY_SC)||'base';
function clonePreset(id){ var o={}; Object.keys(PRESETS[id].g).forEach(function(k){ o[k]={gb:PRESETS[id].g[k][0],p:PRESETS[id].g[k][1]}; }); return o; }
var LS=window.LINHA_STATS||null;
var BO0={A:2,B:3,C:4,D:10,E:10,F:15};          /* bônus de portabilidade (GB) sugerido para a nova grade */
function bo(k){ var g=G[k]; return g&&g.bo!=null?+g.bo:(BO0[k]||0); }
function pp(){ var s=window.PoolApp&&window.PoolApp.SIM; return Math.min(1,Math.max(0,((s&&s.portPct!=null)?+s.portPct:65)/100)); }
var KEY_PV='poolsurf.portval.v1';
var PV0={cp:4.4,cn:13.9,per:9,hz:24,fonte:'Retenção por origem do número, 26/12/2025 a 25/09/2026 (Total ISPX: 5.625 portadas, 3.100 próprias)'};
var PV=Object.assign({},PV0,ls(KEY_PV)||{});

/* acréscimo médio de consumo da franquia por linha, pelo cenário, a partir do consumo linha a linha */
function incremento(k,gb){
  if(!LS||!LS.planos[k]) return 0;
  var P=LS.planos[k], n=P.u.length, s=0;
  for(var i=0;i<n;i++){ var u=P.u[i], cap=P.caps[P.ci[i]], uc=Math.min(u,cap), nu;
    if(SC==='base') nu=Math.min(uc,gb);
    else { nu=Math.min(u,gb); if(SC==='pior' && u>=0.9*cap && u<=cap && gb>cap) nu=gb; }
    s+=nu-uc; }
  return s/n;
}
function recargaGB(k,gb){ /* GB/linha que hoje é recarga paga e passa a caber na franquia nova */
  if(!LS||!LS.planos[k]) return 0; var P=LS.planos[k], s=0;
  for(var i=0;i<P.u.length;i++){ var cap=P.caps[P.ci[i]]; s+=Math.max(0,Math.min(P.u[i],gb)-cap); } return s/P.u.length;
}
function ratio(k){ return LS&&LS.planos[k]?LS.planos[k].ratio:1; }

var KEY_GF='poolsurf.gradefixo.v1', KEY_U='poolsurf.gradeuso.v1', KEY_UB='poolsurf.gradeusob.v1';
var UO=ls(KEY_U)||{};                           /* consumo médio por linha digitado à mão */
var UB=ls(KEY_UB)||'franquia';                  /* base do consumo: 'franquia' (sem recargas) ou 'total' */

function portVal(T){
  var mpl=T.l?T.mN/T.l:0, per=Math.max(1,+PV.per||1), hz=Math.max(1,+PV.hz||1);
  var cm=function(c){ c=Math.min(99.9,Math.max(0,+c||0))/100; return 1-Math.pow(1-c,1/per); };
  var vida=function(m){ return m>0?(1-Math.pow(1-m,hz))/m:hz; };
  var mP=cm(PV.cp), mN=cm(PV.cn), eP=vida(mP), eN=vida(mN), vP=mpl*eP, vN=mpl*eN, d=vP-vN;
  var boM=T.boC, boL=T.lP?T.boC/T.lP:0;
  return {mpl:mpl,per:per,hz:hz,mP:mP,mN:mN,r12P:Math.pow(1-mP,12),r12N:Math.pow(1-mN,12),eP:eP,eN:eN,vP:vP,vN:vN,d:d,boM:boM,boL:boL,conv:d>0?boM/d:0};
}
function calc(M){
  var A=window.PoolApp, S=A.SIM, PC=window.PoolComp, pGB=S.p, GF=A.fm(S);
  var q=pp(), BASEC=!!(PC&&PC.BASE==='com'), ss=PC&&PC.COMP?PC.series(M,true):null, rivals=ss?ss.filter(function(s){return s.op!=='ISPX';}):[];
  var rows=M.planos.map(function(p){
    var g=G[p.k]||{gb:p.gb,p:p.mv}, b=bo(p.k), gbC=g.gb+b, gbE=g.gb+q*b, b0=p.port||0, base0=p.base!=null?p.base:p.gb-b0, gb0E=base0+q*b0;
    var uTot=p.l?p.g/p.l:0, uAuto=UB==='total'?uTot:uTot*ratio(p.k);
    var iS=incremento(p.k,g.gb), iC=incremento(p.k,gbC), incBo=q*(iC-iS);
    var manual=UO[p.k]!=null, uBase=manual?UO[p.k]:uAuto, inc=(1-q)*iS+q*iC, uN=Math.max(0,uBase+inc);
    var dados=pGB*uN, custo=GF+dados, m0=p.mv-p.custo, mN=g.p-custo;
    var gbX=BASEC?gbC:g.gb, cpg=g.p/g.gb, cpgC=g.p/gbC, cpgX=g.p/gbX, cpg0=p.mv/base0, cpg0C=p.mv/p.gb, comp=null;
    if(rivals.length){ comp=rivals.map(function(s){ var c=s.pts.reduce(function(b,x){ var d=Math.abs(x.gb-gbX), bd=b?Math.abs(b.gb-gbX):1e9; return (!b||d<bd||(d===bd&&x.gb<b.gb))?x:b; },null);
        return c?{s:s,p:c,delta:(c.custo/c.gb)/cpgX-1}:null; }).filter(Boolean).sort(function(a,b){return a.delta-b.delta;}); }
    return {k:p.k,l:p.l,gb0:p.gb,base0:base0,port0:b0,gb0E:gb0E,p0:p.mv,c0:p.custo,g:g,bo:b,gbC:gbC,gbE:gbE,gbX:gbX,uAuto:uAuto,uBase:uBase,manual:manual,inc:inc,incBo:incBo,boC:pGB*incBo,uN:uN,dados:dados,custo:custo,m0:m0,mN:mN,cpg:cpg,cpgC:cpgC,cpgX:cpgX,cpg0:cpg0,cpg0C:cpg0C,comp:comp,
      rec:(1-q)*recargaGB(p.k,g.gb)+q*recargaGB(p.k,gbC)};
  });
  var T={l:0,m0:0,mN:0,eco:0,rev0:0,revN:0,gb0:0,gbN:0,uN:0,rec:0,c0:0,cN:0,boC:0,lP:0};
  rows.forEach(function(r){ T.l+=r.l; T.m0+=r.l*r.m0; T.mN+=r.l*r.mN; T.eco+=r.l*(r.p0-r.g.p); T.rev0+=r.l*r.p0; T.revN+=r.l*r.g.p;
    T.gb0+=r.l*r.gb0E; T.gbN+=r.l*r.gbE; T.boC+=r.l*r.boC; T.lP+=r.l*q; T.uN+=r.l*r.uN; T.rec+=r.l*r.rec; T.c0+=r.l*r.c0; T.cN+=r.l*r.custo; });
  return {rows:rows,T:T,p:pGB,f:GF,pp:q,pv:portVal(T)};
}
function html(M){
  var A=window.PoolApp, nf=A.nf, R=A.R, Rk=A.Rk, pct=A.pct, C=calc(M), T=C.T;
  var sg=function(x,d){ return (x>0?'+':x<0?'−':'')+nf(Math.abs(x),d===undefined?1:d); };
  var h='<section class="card" id="grade"><div class="sec-hd" style="flex-direction:row;justify-content:space-between;align-items:flex-start;gap:14px;flex-wrap:wrap"><div style="display:flex;flex-direction:column;gap:4px"><h2>Nova grade de planos</h2>'+
    '<p class="muted small">Parte 1: a oferta ao MVNO (franquia e preço de venda). Parte 2: o custo real da ISPX, que é o consumo médio de cada linha × o custo do GB no pool, e a margem bruta que sobra.</p></div>'+
    '<div class="seggrp"><span class="gl">Carregar grade</span>'+Object.keys(PRESETS).map(function(id){return '<button class="chip" data-gp="'+id+'">'+PRESETS[id].nome+'</button>';}).join('')+'<button class="chip" data-gp="atual">Grade atual</button></div></div>';
  h+='<div class="kpis" style="grid-template-columns:repeat(4,1fr)">'+
    '<div class="kpi"><span class="v">'+Rk(T.m0)+' → '+Rk(T.mN)+'</span><span class="l">margem bruta ISPX/mês: hoje → nova</span></div>'+
    '<div class="kpi"><span class="v" style="color:var(--good-ink)">'+(T.eco>=0?'−':'+')+Rk(Math.abs(T.eco))+'</span><span class="l">no que as MVNOs pagam/mês ('+sg(T.rev0?-T.eco/T.rev0*100:0)+'%)</span></div>'+
    '<div class="kpi"><span class="v">'+sg((T.gbN/T.gb0-1)*100,0)+'%</span><span class="l">GB de franquia entregues às MVNOs (média com '+nf(C.pp*100,0)+'% portadas)</span></div>'+
    '<div class="kpi"><span class="v">'+Rk(T.cN)+'</span><span class="l">custo real SURF no pool/mês (hoje '+Rk(T.c0)+')</span></div></div>';

  /* parte 1 */
  h+='<div class="gpart"><div class="gph"><span class="gnum">1</span><div><h3>Oferta ao MVNO</h3><p class="muted small">Edite a franquia e o preço de venda. A comparação é com o plano atual e com o concorrente mais próximo.</p></div></div>'+
    '<div class="tbl"><table class="gtab" style="min-width:1180px"><thead><tr><th>Plano</th><th>Hoje (sem + bônus)</th><th>Franquia sem portab. (GB)</th><th>Bônus portab. (GB)</th><th>Com portab.</th><th>Preço ao MVNO</th><th>Preço vs hoje</th><th>GB vs hoje (sem · com)</th><th>R$/GB hoje → novo</th><th>Confronto mais difícil</th></tr></thead><tbody>'+
    C.rows.map(function(r){ var dP=r.p0?r.g.p/r.p0-1:0, dG=r.base0?r.g.gb/r.base0-1:0, dGC=r.gb0?r.gbC/r.gb0-1:0, w=r.comp&&r.comp[0];
      return '<tr><td>'+r.k+'</td><td>'+r.base0+' + '+r.port0+' = '+r.gb0+' GB · '+R(r.p0,2)+'</td>'+
        '<td><input type="number" min="1" step="1" data-gk="'+r.k+'" data-gf="gb" value="'+r.g.gb+'" aria-label="Franquia sem portabilidade plano '+r.k+'"></td>'+
        '<td><input type="number" min="0" step="1" data-gk="'+r.k+'" data-gf="bo" value="'+r.bo+'" aria-label="Bônus de portabilidade plano '+r.k+'"><div class="sub">'+(r.g.gb?'+'+nf(r.bo/r.g.gb*100,0)+'%':'')+'</div></td>'+
        '<td><b>'+nf(r.gbC)+' GB</b></td>'+
        '<td><input type="number" min="0" step="0.10" data-gk="'+r.k+'" data-gf="p" value="'+r.g.p.toFixed(2)+'" aria-label="Preço ao MVNO plano '+r.k+'"></td>'+
        '<td class="'+(dP<0?'ok':dP>0?'hot':'')+'">'+sg(dP*100)+'% <span class="sub">('+(r.g.p-r.p0>=0?'+':'−')+R(Math.abs(r.g.p-r.p0),2)+')</span></td>'+
        '<td class="'+(dG>0?'ok':dG<0?'hot':'')+'">'+sg(dG*100,0)+'% <span class="sub">('+sg(r.g.gb-r.base0,0)+' GB)</span> · '+sg(dGC*100,0)+'% <span class="sub">('+sg(r.gbC-r.gb0,0)+' GB)</span></td>'+
        '<td>sem: '+R(r.cpg0,2)+' → <b>'+R(r.cpg,2)+'</b><div class="sub">com: '+R(r.cpg0C,2)+' → '+R(r.cpgC,2)+'</div></td>'+
        '<td>'+(w?'<span class="'+(w.delta>=0?'ok':'hot')+'">'+esc(w.s.nome)+' '+nf(w.p.gb)+' GB · '+R(w.p.custo,2)+' · '+(w.delta>0?'+':'')+pct(w.delta,0)+'</span>':'—')+'</td></tr>'; }).join('')+
    '</tbody><tfoot><tr><td>Carteira</td><td></td><td></td><td></td><td></td><td></td><td class="'+(T.eco>0?'ok':'')+'">'+sg(T.rev0?-T.eco/T.rev0*100:0)+'%</td><td>'+sg((T.gbN/T.gb0-1)*100,0)+'%</td><td></td><td></td></tr></tfoot></table></div>'+
    '<p class="small muted">Bônus de portabilidade: GB a mais para linha portada ('+nf(C.pp*100,0)+'% da base, premissa editável na proposta SURF). Hoje: A +2, B +2, C +5, D +5, E +5, F +10. Confronto mais difícil (base '+(r0base())+'): entre os planos concorrentes de tamanho mais próximo, o de menor R$/GB. Verde = a ISPX é mais barata por GB; vermelho = o concorrente é. % = quanto o GB do concorrente custa a mais (+) ou a menos (−) que o da ISPX.</p></div>';

  /* parte 2 */
  h+='<div class="gpart"><div class="gph"><span class="gnum">2</span><div><h3>Custo real e margem bruta da ISPX</h3><p class="muted small">Custo real por linha = consumo médio × R$/GB da proposta + custos fixos por linha (SIM, TFF, TFI e portabilidade diluídas, voz e SMS). Margem bruta = preço ao MVNO − custo real. Abaixo de cada margem, a diferença contra a margem de hoje (verde: maior; vermelho: menor). O preço ao MVNO pode ser editado aqui ou na Parte 1.</p></div></div>'+
    '<div class="gctl"><div class="ctl"><label for="g-mb">Dados: R$ por MB (nossa proposta)</label><input id="g-mb" type="number" step="0.00001" min="0" value="'+(+A.SIM.mb).toFixed(5)+'"><span class="hint">= '+R(C.p,2)+'/GB. SURF propôs R$ '+(+A.SURF0.mb).toFixed(4)+'/MB.</span></div>'+
    '<div class="ctl"><span class="gl" style="margin-bottom:2px">Custos fixos por linha/mês</span><b style="font-family:var(--mono);font-size:16px">'+R(C.f,2)+'</b><span class="hint">SIM '+R(A.SIM.sim,2)+' + TFF '+R(A.SIM.tff/12,2)+' + TFI/portabilidade '+R(A.oneL(A.SIM)/Math.max(1,A.SIM.meses),2)+' + voz e SMS. Edite na tabela da proposta.</span></div>'+
    '<div class="ctl"><span class="gl" style="margin-bottom:2px">Consumo médio por linha</span><div class="seggrp">'+
      [['franquia','Só franquia'],['total','Total, com recargas']].map(function(x){return '<button class="chip" data-gu="'+x[0]+'" aria-pressed="'+(UB===x[0])+'">'+x[1]+'</button>';}).join('')+
      '<button class="chip" id="g-ureset"'+(Object.keys(UO).length?'':' hidden')+'>Voltar ao calculado</button></div>'+
      '<div class="seggrp" style="margin-top:6px"><span class="gl">Cenário</span>'+[['base','Igual ao de hoje'],['real','Realista'],['pior','Pior caso']].map(function(x){return '<button class="chip" data-gs="'+x[0]+'" aria-pressed="'+(SC===x[0])+'">'+x[1]+'</button>';}).join('')+'</div>'+
      '<span class="hint">'+(SC==='base'?'Cada linha consome o mesmo que hoje.':SC==='real'?'Quem hoje compra recarga passa a usar a franquia nova até o consumo que já tem.':'Além do realista, quem está entre 90% e 100% da franquia passa a usar a franquia nova inteira.')+'</span></div></div>'+
    '<div class="tbl"><table class="gtab" style="min-width:1440px"><thead><tr><th>Plano</th><th>Linhas</th><th>Consumo médio (GB/linha)</th><th>× Custo do GB</th><th>= Custo dos dados</th>'+(C.f?'<th>+ Fixos/linha</th>':'')+'<th>Custo real/linha</th><th>Franquia sem portab. (GB)</th><th>Preço ao MVNO</th><th>R$/GB ao MVNO</th><th>Margem bruta/linha</th><th>Margem hoje</th><th>Margem %</th><th>Margem/mês</th><th>Concorrentes (plano mais próximo)</th></tr></thead><tbody>'+
    C.rows.map(function(r){
      var cc=(r.comp||[]).slice(0,3).map(function(c){ return '<div class="cc '+(c.delta>=0?'ok':'hot')+'">'+esc(c.s.nome)+': '+nf(c.p.gb)+' GB · '+R(c.p.custo,2)+' <span class="sub">('+R(c.p.custo/c.p.gb,2)+'/GB, '+(c.delta>0?'+':'')+pct(c.delta,0)+')</span></div>'; }).join('');
      return '<tr><td>'+r.k+'<div class="sub">'+r.g.gb+' GB</div></td><td>'+nf(r.l)+'</td>'+
        '<td><input type="number" min="0" step="0.01" data-gu-k="'+r.k+'" value="'+r.uBase.toFixed(2)+'" class="'+(r.manual?'man':'')+'" aria-label="Consumo médio plano '+r.k+'">'+(r.inc>0.005?'<span class="sub"> +'+nf(r.inc,2)+' no cenário</span>':'')+(r.manual?'<span class="sub"> (digitado; calculado '+nf(r.uAuto,2)+')</span>':'')+'</td>'+
        '<td>'+R(C.p,2)+'</td><td>'+R(r.dados,2)+'</td>'+(C.f?'<td>'+R(C.f,2)+'</td>':'')+'<td><b>'+R(r.custo,2)+'</b></td><td><input type="number" min="1" step="1" data-gk="'+r.k+'" data-gf="gb" value="'+r.g.gb+'" aria-label="Franquia plano '+r.k+' (custo real)"><div class="sub">+'+nf(r.bo)+' portab. = '+nf(r.gbC)+' · hoje '+r.base0+'+'+r.port0+'</div></td><td><input type="number" min="0" step="0.10" data-gk="'+r.k+'" data-gf="p" value="'+r.g.p.toFixed(2)+'" aria-label="Preço ao MVNO plano '+r.k+' (custo real)"><div class="sub">hoje '+R(r.p0,2)+'</div></td><td><b>'+R(r.cpg,2)+'</b><div class="sub">hoje '+R(r.cpg0,2)+'</div></td>'+
        '<td><b class="'+(r.mN<0?'hot':'ok')+'">'+(r.mN<0?'−':'')+R(Math.abs(r.mN),2)+'</b><div class="sub '+(r.mN>=r.m0?'ok':'hot')+'">'+(r.mN>=r.m0?'+':'−')+R(Math.abs(r.mN-r.m0),2)+' vs hoje</div></td><td>'+R(r.m0,2)+'</td><td>'+pct(r.g.p?r.mN/r.g.p:0,0)+'</td><td>'+Rk(r.l*r.mN)+'</td><td class="ccol">'+(cc||'—')+'</td></tr>'; }).join('')+
    '</tbody><tfoot><tr><td>Carteira</td><td>'+nf(T.l)+'</td><td>'+nf(T.l?T.uN/T.l:0,2)+'</td><td></td><td>'+Rk(T.cN-C.f*T.l)+'</td>'+(C.f?'<td>'+Rk(C.f*T.l)+'</td>':'')+'<td>'+Rk(T.cN)+'</td><td>'+nf(T.l?T.gbN/T.l:0,1)+' GB <span class="sub">média</span></td><td>'+Rk(T.revN)+'</td><td>'+R(T.gbN?T.revN/T.gbN:0,2)+'</td><td><b class="'+(T.mN>=T.m0?'ok':'hot')+'">'+Rk(T.mN)+'</b></td><td>'+Rk(T.m0)+'</td><td>'+pct(T.revN?T.mN/T.revN:0,0)+'</td><td><b>'+Rk(T.mN)+'</b></td><td></td></tr></tfoot></table></div>'+
    '<p class="small muted">Concorrentes: os 3 planos de tamanho mais próximo com menor R$/GB. % = quanto o GB do concorrente custa a mais (+, verde: ISPX mais barata) ou a menos (−, vermelho) que o GB do plano ISPX.</p>';
  var pe=T.uN?(T.revN-T.m0-C.f*T.l)/T.uN:0;
  h+='<p class="small muted">Exemplo do Plano A: '+nf(C.rows[0].uN,2)+' GB × '+R(C.p,2)+' = '+R(C.rows[0].dados,2)+(C.f?' + '+R(C.f,2)+' de custos fixos':'')+'; '+R(C.rows[0].g.p,2)+' − '+R(C.rows[0].custo,2)+' = margem de '+R(C.rows[0].mN,2)+' por linha. '+
    'Para esta grade não perder a margem de hoje ('+Rk(T.m0)+'), o GB do pool pode custar no máximo <b>'+R(pe,2)+'</b>.'+(' Custo do bônus de portabilidade no pool: <b>'+(T.boC>0.5?Rk(T.boC)+'/mês':'R$ 0')+'</b>'+(T.boC>0.5?' ('+R(C.pv.boL,2)+' por linha portada), já incluído no custo real':' neste cenário (o consumo não passa da franquia de hoje)')+'.')+(T.rec>0.5?' Cerca de <b>'+nf(T.rec)+' GB/mês</b> que hoje são recarga paga passam a caber na franquia nova (receita de recarga não descontada).':'')+'</p>'+
    '<p class="small muted">Consumo "só franquia": média da carteira sem o que foi consumido com recarga extra, ajustada pelo consumo linha a linha'+(LS?' ('+LS.marcas.length+' MVNOs, sem a ISUPER)':'')+'. "Total": média da carteira com recargas (ex.: Plano A 2,10 GB). A média da amostra linha a linha (2,34 GB no A) vale só para essas '+(LS?LS.marcas.length:'')+' MVNOs. Margem hoje = preço atual − custo atual na tabela SURF. '+'Custos fixos por linha: '+R(C.f,2)+'/mês pela nossa proposta.'+'</p></div>'+partePort(C)+'</section>';
  return h;
}
function r0base(){ var PC=window.PoolComp; return PC&&PC.BASE==='com'?'com portabilidade':'sem portabilidade'; }
function partePort(C){
  var A=window.PoolApp, nf=A.nf, R=A.R, Rk=A.Rk, pct=A.pct, v=C.pv;
  var inp=function(k,st,lab,suf){ return '<div class="ctl"><label for="pv-'+k+'">'+lab+'</label><input id="pv-'+k+'" type="number" min="0" step="'+st+'" data-pv="'+k+'" value="'+PV[k]+'"><span class="hint">'+suf+'</span></div>'; };
  var h='<div class="gpart"><div class="gph"><span class="gnum">3</span><div><h3>Valor da portabilidade</h3><p class="muted small">Linha portada cancela menos. Quanto vale cada linha que o bônus convence a portar, contra o custo do bônus no pool. Fonte: '+esc(PV.fonte)+'. Churn = cancelamentos no período.</p></div></div>'+
    '<div class="gctl">'+inp('cp','0.1','Churn portadas no período (%)','Total ISPX: 4,4% (sócios 4,5%, clientes 3,6%)')+inp('cn','0.1','Churn números próprios no período (%)','Total ISPX: 13,9% (sócios 10,7%, clientes 32,7%)')+
      inp('per','1','Meses do período','26/12/2025 a 25/09/2026 = 9 meses')+inp('hz','1','Horizonte de análise (meses)','Prazo em que a margem por linha é somada')+
      '<div class="ctl" style="justify-content:flex-end"><button class="chip" id="pv-reset">Voltar aos dados da tabela</button></div></div>';
  h+='<div class="tbl"><table class="gtab"><thead><tr><th>Origem do número</th><th>Churn no período</th><th>Churn mensal equivalente</th><th>Ativas após 12 meses</th><th>Meses ativos no horizonte</th><th>Margem ISPX no horizonte</th></tr></thead><tbody>'+
    '<tr><td>Portada</td><td>'+nf(PV.cp,1)+'%</td><td>'+nf(v.mP*100,2)+'%</td><td>'+pct(v.r12P,0)+'</td><td>'+nf(v.eP,1)+' de '+v.hz+'</td><td><b class="ok">'+R(v.vP,2)+'</b></td></tr>'+
    '<tr><td>Número próprio</td><td>'+nf(PV.cn,1)+'%</td><td>'+nf(v.mN*100,2)+'%</td><td>'+pct(v.r12N,0)+'</td><td>'+nf(v.eN,1)+' de '+v.hz+'</td><td><b>'+R(v.vN,2)+'</b></td></tr>'+
    '</tbody><tfoot><tr><td>Diferença por linha</td><td></td><td></td><td></td><td>'+nf(v.eP-v.eN,1)+' meses</td><td><b class="ok">+'+R(v.d,2)+'</b></td></tr></tfoot></table></div>';
  h+='<div class="kpis" style="grid-template-columns:repeat(3,1fr)">'+
    '<div class="kpi"><span class="v">+'+R(v.d,2)+'</span><span class="l">margem a mais por linha que troca número próprio por portada ('+v.hz+' meses, '+R(v.mpl,2)+'/linha/mês da nova grade)</span></div>'+
    '<div class="kpi"><span class="v">'+(v.boM>0.5?Rk(v.boM):'R$ 0')+'</span><span class="l">custo do bônus no pool por mês'+(v.boM>0.5?' ('+R(v.boL,2)+' por linha portada)':' neste cenário')+'</span></div>'+
    '<div class="kpi"><span class="v">'+(v.boM>0.5?nf(Math.ceil(v.conv*10)/10,1):'0')+'</span><span class="l">linhas convertidas por mês para o bônus se pagar</span></div></div>'+
    '<p class="small muted">Cuidado: correlação não é causa. Quem porta já decidiu ficar; o bônus converte quem está em dúvida, e esse cliente tende a ficar entre os dois churns. Mesmo se ele valer metade da diferença, bastam '+(v.boM>0.5?nf(Math.ceil(v.conv*2*10)/10,1):'0')+' conversões por mês. O churn mensal equivalente supõe cancelamento constante ao longo do período.</p></div>';
  return h;
}
function bind(){
  var A=window.PoolApp;
  document.querySelectorAll('[data-gk]').forEach(function(inp){ inp.addEventListener('change',function(){
    var v=parseFloat(String(inp.value).replace(',','.')), k=inp.dataset.gk, f=inp.dataset.gf;
    if(!isFinite(v)||v<0||(v===0&&f!=='bo')){ inp.value=f==='bo'?bo(k):G[k][f]; return; } if(!G[k]) G[k]={gb:v,p:0}; G[k][f]=v; ls(KEY_G,G); A.rerender(); }); });
  document.querySelectorAll('[data-pv]').forEach(function(inp){ inp.addEventListener('change',function(){ var v=parseFloat(String(inp.value).replace(',','.')); if(!isFinite(v)||v<0) return; PV[inp.dataset.pv]=v; ls(KEY_PV,PV); A.rerender(); }); });
  var pvr=$('pv-reset'); if(pvr) pvr.onclick=function(){ PV=Object.assign({},PV0); ls(KEY_PV,null); A.rerender(); };
  document.querySelectorAll('[data-gu-k]').forEach(function(inp){ inp.addEventListener('change',function(){
    var v=parseFloat(String(inp.value).replace(',','.')); if(!isFinite(v)||v<0) return; UO[inp.dataset.guK]=v; ls(KEY_U,UO); A.rerender(); }); });
  var gp=$('g-p'), gf=$('g-f'), ur=$('g-ureset');
  var gmb=$('g-mb'); if(gmb) gmb.addEventListener('change',function(){ var v=parseFloat(String(gmb.value).replace(',','.')); if(isFinite(v)&&v>=0) A.setSim({mb:v}); });
  if(ur) ur.onclick=function(){ UO={}; ls(KEY_U,null); A.rerender(); };
  document.querySelectorAll('[data-gu]').forEach(function(b){ b.onclick=function(){ UB=b.dataset.gu; ls(KEY_UB,UB); A.rerender(); }; });
  document.querySelectorAll('[data-gs]').forEach(function(b){ b.onclick=function(){ SC=b.dataset.gs; ls(KEY_SC,SC); A.rerender(); }; });
  document.querySelectorAll('[data-gp]').forEach(function(b){ b.onclick=function(){
    if(b.dataset.gp==='atual'){ G={}; A.TAB.forEach(function(t){ G[t.k]={gb:t.base!=null?t.base:t.gb,bo:t.port||0,p:t.mv}; }); } else G=clonePreset(b.dataset.gp);
    ls(KEY_G,G); A.rerender(); A.toast('Grade '+(b.dataset.gp==='atual'?'atual':PRESETS[b.dataset.gp].nome)+' carregada.'); }; });
}
function reset(){ G=clonePreset('inter'); SC='base'; UO={}; UB='franquia'; PV=Object.assign({},PV0); [KEY_G,KEY_SC,KEY_GF,KEY_U,KEY_UB,KEY_PV].forEach(function(k){ls(k,null);}); }
window.PoolGrade={html:html,bind:bind,calc:calc,reset:reset,get PV(){return PV;},get SC(){return SC;},get UB(){return UB;}};
})();
