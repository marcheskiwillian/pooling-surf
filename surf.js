/* surf.js — Proposta da SURF (cobrança por consumo) × nossa contraproposta.
   Os valores da SURF ficam fixos como referência; a coluna "Nossa proposta" é editável e alimenta todo o app. */
(function(){
'use strict';
function $(id){ return document.getElementById(id); }
var ROWS=[
  {k:'sim',   item:'Mensalidade SIM Card', desc:'Custo mensal por SIM Card', per:'mensal, por linha', dec:2, un:'R$'},
  {k:'tff',   item:'FISTEL - TFF', desc:'Taxa de funcionamento anual por SIM Card', per:'anual, por linha (÷12 no mês)', dec:2, un:'R$'},
  {k:'tfi',   item:'Taxa de Instalação (TFI)', desc:'Taxa única de instalação por SIM Card', per:'única, por linha ativada', dec:2, un:'R$'},
  {k:'port',  item:'Portabilidade', desc:'Custo para portabilidade de número', per:'única, por portabilidade', dec:2, un:'R$'},
  {k:'mb',    item:'Dados', desc:'Consumo de dados por megabyte', per:'por MB consumido', dec:5, un:'R$'},
  {k:'voz',   item:'Voz', desc:'Consumo de voz por minuto', per:'por minuto', dec:3, un:'R$'},
  {k:'sms',   item:'SMS', desc:'Envio de mensagens SMS por unidade', per:'por SMS', dec:2, un:'R$'}
];
var PREM=[
  {k:'portPct', item:'Ativações com portabilidade', desc:'Parte das vendas de chip que chega por portabilidade', dec:0, un:'%'},
  {k:'meses',   item:'Diluir custos únicos em', desc:'TFI e portabilidade rateadas no custo mensal da linha', dec:0, un:'meses'},
  {k:'mbgb',    item:'MB por GB', desc:'Conversão usada pela SURF (1.024 no arquivo de consumo)', dec:0, un:'MB'},
  {k:'min',     item:'Minutos de voz por linha/mês', desc:'Sem dado de consumo de voz. Informe o valor medido', dec:0, un:'min', warn:true},
  {k:'nsms',    item:'SMS por linha/mês', desc:'Sem dado de consumo de SMS. Informe o valor medido', dec:0, un:'SMS', warn:true}
];
function f(v,d){ return (+v||0).toLocaleString('pt-BR',{minimumFractionDigits:d,maximumFractionDigits:d}); }
function custoPlano(S,u){ var A=window.PoolApp; return A.fm(S)+S.mb*S.mbgb*u; }
function calc(M){
  var A=window.PoolApp, S=A.SIM, S0=A.syncP(Object.assign({},A.SURF0,{min:S.min,nsms:S.nsms,portPct:S.portPct,meses:S.meses,mbgb:S.mbgb})), T=M.T;
  var planos=M.planos.map(function(p){ var u=p.l?p.g/p.l:0; return {k:p.k,l:p.l,u:u,hoje:p.custo,surf:custoPlano(S0,u),nossa:custoPlano(S,u),preco:p.mv}; });
  var tot=function(key){ return planos.reduce(function(a,p){return a+p.l*p[key];},0); };
  var beMB=function(s){ var fx=A.fm(s)*T.l; return T.g?(T.cost-fx)/T.g/(s.mbgb||1024):0; };
  return {S:S,S0:S0,planos:planos,hoje:T.cost,surf:tot('surf'),nossa:tot('nossa'),rev:T.rev,beSurf:beMB(S0),beNossa:beMB(S),T:T};
}
function html(M){
  var A=window.PoolApp, R=A.R, Rk=A.Rk, nf=A.nf, pct=A.pct, C=calc(M), S=C.S, S0=C.S0;
  var dv=function(a,b){ var d=b?a/b-1:0; return '<span class="'+(d<-0.0005?'ok':d>0.0005?'hot':'muted')+'">'+(d>0?'+':'')+pct(d,1)+'</span>'; };
  var h='<section class="card" id="surfprop"><div class="sec-hd" style="flex-direction:row;justify-content:space-between;align-items:flex-start;gap:14px;flex-wrap:wrap"><div style="display:flex;flex-direction:column;gap:4px"><h2>Proposta da SURF × nossa proposta</h2>'+
    '<p class="muted small">Cobrança por consumo: não há pacote de GB comprado. A coluna "Nossa proposta" é editável e alimenta o simulador, a nova grade, a projeção e o PPT.</p></div>'+
    '<button class="btn" id="sp-reset">Voltar aos valores da SURF</button></div>';
  h+='<div class="tbl"><table class="gtab sptab" style="min-width:860px"><thead><tr><th>Item</th><th style="text-align:left">Descrição</th><th style="text-align:left">Cobrança</th><th>SURF propôs</th><th>Nossa proposta</th><th>Diferença</th></tr></thead><tbody>'+
    ROWS.map(function(r){ return '<tr><td>'+r.item+'</td><td style="text-align:left;font-family:inherit">'+r.desc+'</td><td style="text-align:left;font-family:inherit" class="small muted">'+r.per+'</td>'+
      '<td>R$ '+f(A.SURF0[r.k],r.dec)+'</td><td><input type="number" min="0" step="'+Math.pow(10,-r.dec)+'" data-sp="'+r.k+'" value="'+(+S[r.k]||0).toFixed(r.dec)+'" aria-label="'+r.item+' nossa proposta"></td><td>'+dv(S[r.k],A.SURF0[r.k])+'</td></tr>'; }).join('')+
    '</tbody></table></div>';
  h+='<div class="gph" style="margin-top:4px"><div><h3>Premissas</h3><p class="muted small">Valem para as duas colunas.</p></div></div><div class="tbl"><table class="gtab sptab" style="min-width:760px"><tbody>'+
    PREM.map(function(r){ return '<tr'+(r.warn&&!(+S[r.k])?' class="spwarn"':'')+'><td>'+r.item+'</td><td style="text-align:left;font-family:inherit">'+r.desc+'</td><td><input type="number" min="0" step="1" data-sp="'+r.k+'" value="'+(+S[r.k]||0)+'" aria-label="'+r.item+'"> <span class="small">'+r.un+'</span></td></tr>'; }).join('')+
    '</tbody></table></div>';
  if(!(+S.min)||!(+S.nsms)) h+='<p class="small" style="color:var(--warn-ink);background:var(--warn-bg);padding:8px 12px;border-radius:8px">Voz e SMS estão zerados porque não temos o consumo por linha. Os planos têm voz ilimitada: cada 100 minutos por linha custam '+R(100*S.voz,2)+'/linha/mês pela proposta ('+Rk(100*S.voz*C.T.l)+'/mês na carteira). Peça à SURF o consumo de voz e SMS por linha.</p>';
  /* composição do custo por linha */
  var uMed=C.T.l?C.T.g/C.T.l:0;
  h+='<div class="kpis" style="grid-template-columns:repeat(4,1fr)">'+
    '<div class="kpi"><span class="v">'+Rk(C.hoje)+'</span><span class="l">fatura hoje (tabela fixa por plano)</span></div>'+
    '<div class="kpi alert"><span class="v">'+Rk(C.surf)+'</span><span class="l">fatura pela proposta da SURF ('+(C.surf>=C.hoje?'+':'−')+pct(Math.abs(C.surf/C.hoje-1))+' vs hoje)</span></div>'+
    '<div class="kpi"><span class="v" style="color:'+(C.nossa<=C.hoje?'var(--good-ink)':'var(--crit-ink)')+'">'+Rk(C.nossa)+'</span><span class="l">fatura pela nossa proposta ('+(C.nossa>=C.hoje?'+':'−')+pct(Math.abs(C.nossa/C.hoje-1))+' vs hoje)</span></div>'+
    '<div class="kpi"><span class="v">R$ '+f(C.beNossa,4)+'</span><span class="l">R$/MB que empata com a fatura de hoje (com os fixos da nossa proposta)</span></div></div>';
  h+='<div class="g2e"><div><h3>Custo por linha/mês</h3><div class="tbl"><table class="gtab"><thead><tr><th>Componente</th><th>SURF propôs</th><th>Nossa proposta</th></tr></thead><tbody>'+
    [['SIM Card',S0.sim,S.sim],['TFF (anual ÷ 12)',S0.tff/12,S.tff/12],['TFI + portabilidade ÷ '+S.meses+' meses',A.oneL(S0)/Math.max(1,S0.meses),A.oneL(S)/Math.max(1,S.meses)],['Voz ('+nf(S.min)+' min)',S0.voz*S.min,S.voz*S.min],['SMS ('+nf(S.nsms)+')',S0.sms*S.nsms,S.sms*S.nsms],
     ['<b>Fixo por linha</b>',A.fm(S0),A.fm(S)],['Dados ('+nf(uMed,2)+' GB médios × R$/GB)',S0.p*uMed,S.p*uMed],['<b>Total por linha</b>',A.fm(S0)+S0.p*uMed,A.fm(S)+S.p*uMed]]
      .map(function(r){return '<tr><td>'+r[0]+'</td><td>'+R(r[1],2)+'</td><td>'+R(r[2],2)+'</td></tr>';}).join('')+
    '</tbody></table></div><p class="small muted">R$/GB: SURF '+R(S0.p,2)+' · nossa '+R(S.p,2)+' (R$/MB × '+nf(S.mbgb)+'). Hoje o custo médio por linha na tabela fixa é '+R(C.T.l?C.hoje/C.T.l:0,2)+'.</p></div>'+
    '<div><h3>Custo por linha em cada plano, com o consumo atual</h3><div class="tbl"><table class="gtab"><thead><tr><th>Plano</th><th>GB/linha</th><th>Hoje (tabela)</th><th>Proposta SURF</th><th>Nossa proposta</th><th>Preço ao MVNO</th></tr></thead><tbody>'+
    C.planos.map(function(p){ return '<tr><td>'+p.k+'</td><td>'+nf(p.u,2)+'</td><td>'+R(p.hoje,2)+'</td><td class="'+(p.surf>p.hoje?'hot':'ok')+'">'+R(p.surf,2)+'</td><td class="'+(p.nossa>p.hoje?'hot':'ok')+'">'+R(p.nossa,2)+'</td><td>'+R(p.preco,2)+(p.surf>p.preco?' <span class="sub hot">SURF acima do preço</span>':'')+'</td></tr>'; }).join('')+
    '</tbody></table></div><p class="small muted">Consumo total por linha (inclui o que foi usado com recarga extra, que também é cobrado por MB). Vermelho = mais caro que a tabela fixa de hoje.</p></div></div>';
  h+='<p class="small muted">Custos únicos (TFI e portabilidade) diluídos em '+S.meses+' meses por linha, com '+nf(S.portPct)+'% das ativações por portabilidade: '+R(A.oneL(S),2)+' por linha ativada. Para empatar com a fatura de hoje pela proposta da SURF, o MB teria de custar R$ '+f(C.beSurf,4)+' (R$ '+f(C.beSurf*S.mbgb,2)+'/GB).</p></section>';
  return h;
}
function bind(){
  var A=window.PoolApp;
  document.querySelectorAll('[data-sp]').forEach(function(inp){ inp.addEventListener('change',function(){
    var v=parseFloat(String(inp.value).replace(',','.')); if(!isFinite(v)||v<0){ inp.value=A.SIM[inp.dataset.sp]; return; }
    var o={}; o[inp.dataset.sp]=v; A.setSim(o); }); });
  var b=$('sp-reset'); if(b) b.onclick=function(){ var S=A.SIM, o={}; ['sim','tff','tfi','port','mb','voz','sms'].forEach(function(k){ o[k]=A.SURF0[k]; }); A.setSim(o); A.toast('Nossa proposta voltou aos valores da SURF.'); };
}
window.PoolSurf={html:html,bind:bind,calc:calc};
})();
