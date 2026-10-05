/* app.js — Pooling SURF · ISPX
   Lê a planilha de consumo da SURF, cruza com a tabela de custo por plano
   e simula a compra em pool. 100% navegador: nada sai deste computador. */
(function(){
'use strict';

var KEY_EXC='poolsurf.exc.v1';
var KEY_DATA='poolsurf.payload.v1', KEY_TAB='poolsurf.tab.v1', KEY_SIM='poolsurf.sim.v2', KEY_MES='poolsurf.meses.v1';

/* tabela padrão: custo SURF informado pelo Wagner; preço MVNO = tabela oficial ISPX */
var TAB0=[
  {k:'A',base:4, port:2, gb:6,  custo:20, mv:24.90},
  {k:'B',base:8, port:2, gb:10, custo:25, mv:28.00},
  {k:'C',base:10,port:5, gb:15, custo:32, mv:36.50},
  {k:'D',base:15,port:5, gb:20, custo:40, mv:44.90},
  {k:'E',base:25,port:5, gb:30, custo:60, mv:67.00},
  {k:'F',base:40,port:10,gb:50, custo:79, mv:85.00}];
/* proposta recebida da SURF (valores de referência; a 'nossa proposta' começa igual e é editável) */
var SURF0={mb:0.0061,mbgb:1024,sim:1.45,tff:14,tfi:26.83,port:4,portPct:65,meses:12,voz:0.026,min:0,sms:0.08,nsms:0};
var SIM0=Object.assign({},SURF0,{c:0,g:0,cap:false});
/* f = ativação por linha, paga uma única vez; entra no custo mensal diluída em 'meses' */
function syncP(s){ s.p=(+s.mb||0)*(+s.mbgb||1024); return s; }
function recL(s){ return (+s.sim||0)+(+s.tff||0)/12+(+s.voz||0)*(+s.min||0)+(+s.sms||0)*(+s.nsms||0); }   /* recorrente por linha/mês */
function oneL(s){ return (+s.tfi||0)+((+s.portPct||0)/100)*(+s.port||0); }                              /* único por linha ativada */
function fm(s){ return recL(s)+oneL(s)/Math.max(1,+s.meses||12); }                                       /* fixo por linha/mês */
var MESN=['jan','fev','mar','abr','mai','jun','jul','ago','set','out','nov','dez'];

/* ---------- utilidades ---------- */
function ls(k,v){ try{ if(v===undefined){ var r=localStorage.getItem(k); return r?JSON.parse(r):null; } if(v===null) localStorage.removeItem(k); else localStorage.setItem(k,JSON.stringify(v)); }catch(e){ return null; } }
function clone(o){ return JSON.parse(JSON.stringify(o)); }
function nf(v,d){ d=d||0; return (isFinite(v)?v:0).toLocaleString('pt-BR',{minimumFractionDigits:d,maximumFractionDigits:d}); }
function R(v,d){ return 'R$ '+nf(v,d||0); }
function Rk(v){ return 'R$ '+nf(v/1000,1).replace(/,0$/,'')+' mil'; }
function pct(v,d){ return nf(v*100,d===undefined?1:d)+'%'; }
function mesLbl(m){ var p=String(m).split('-'); return (MESN[(+p[1])-1]||p[1])+'/'+String(p[0]).slice(2); }
function sum(a){ return a.reduce(function(s,x){return s+x;},0); }
function avg(a){ return a.length?sum(a)/a.length:0; }
function css(n){ return getComputedStyle(document.documentElement).getPropertyValue(n).trim(); }
function $(id){ return document.getElementById(id); }
function esc(s){ return String(s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];}); }
function toast(msg){ var t=$('toast'); t.textContent=msg; t.hidden=false; clearTimeout(toast._t); toast._t=setTimeout(function(){t.hidden=true;},4200); }

/* ---------- estado ---------- */
var ORIG=window.POOL_DATA||null;
var DATA=ls(KEY_DATA)||ORIG;
var TAB=ls(KEY_TAB)||clone(TAB0);
var SAVEDSIM=ls(KEY_SIM); var SIM=syncP(Object.assign(clone(SIM0), SAVEDSIM||{}));
/* premissa de portabilidade: 70% → 65% (base real: 64,5% portadas na tabela de retenção). Migra uma vez. */
if(!ls('poolsurf.pp65.v1')){ if(SAVEDSIM && +SIM.portPct===70){ SIM.portPct=65; ls(KEY_SIM,SIM); } ls('poolsurf.pp65.v1',1); }
var MSEL=null;
var EXC=ls(KEY_EXC)||[];           /* MVNOs desconsideradas (nomes normalizados) */
var ML=window.MVNO_LINHAS||null;   /* linhas por MVNO e plano (app Benchmark de MVNOs) */
var MVOPEN=false;
var ALIAS={'HOMENET':'HNET MOVEL'};
function normMv(n){ var s=String(n||'').toUpperCase().replace(/\s+/g,' ').trim(); return ALIAS[s]||s; }

/* ---------- importação ---------- */
function mesDe(v){
  if(v instanceof Date) return v.getFullYear()+'-'+String(v.getMonth()+1).padStart(2,'0');
  if(typeof v==='number' && v>20000 && v<80000){ var d=new Date(Math.round((v-25569)*86400000)); return d.getUTCFullYear()+'-'+String(d.getUTCMonth()+1).padStart(2,'0'); }
  var s=String(v==null?'':v).trim();
  var m=s.match(/^(\d{4})[-\/](\d{1,2})/); if(m) return m[1]+'-'+m[2].padStart(2,'0');
  m=s.match(/^(\d{1,2})[-\/](\d{4})/); if(m) return m[2]+'-'+m[1].padStart(2,'0');
  return s||null;
}
function add(o,a,b,v){ (o[a]=o[a]||{}); o[a][b]=(o[a][b]||0)+v; }

function parseWorkbook(wb, nome){
  var base=null, hdr=null, rows=null;
  wb.SheetNames.forEach(function(sn){
    if(base) return;
    var r=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:null});
    for(var i=0;i<Math.min(8,r.length);i++){
      var h=(r[i]||[]).map(function(x){return String(x==null?'':x).trim().toLowerCase();});
      if(h.indexOf('gb_total')>=0 && h.indexOf('no_plano_linha')>=0 && h.indexOf('dt_mes')>=0){ base=sn; hdr=h; rows=r.slice(i+1); break; }
    }
  });
  if(!base) throw new Error('Não achei a aba de consumo. Ela precisa ter as colunas dt_mes, no_plano_linha e gb_total (como a aba "Base").');
  var ix=function(n){ return hdr.indexOf(n); };
  var iM=ix('dt_mes'), iMa=ix('marca'), iP=ix('no_plano_linha'), iT=ix('no_plano_tipo_consumido'), iS=ix('no_msisdn_status'), iG=ix('gb_total');
  var P={gbPlano:{},gbTipo:{},gbMarca:{},gbStatus:{},gbMP:{},gbMT:{},registros:0};
  var meses={};
  rows.forEach(function(r){
    if(!r || r[iM]==null || r[iM]==='') return;
    var m=mesDe(r[iM]); if(!m) return;
    var g=+r[iG]; if(!isFinite(g)) g=0;
    meses[m]=1; P.registros++;
    add(P.gbPlano, r[iP]==null||r[iP]===''?'(sem plano)':String(r[iP]).trim(), m, g);
    if(iT>=0) add(P.gbTipo, r[iT]==null?'(vazio)':String(r[iT]).trim(), m, g);
    if(iMa>=0){ var mk=r[iMa]==null?'(vazio)':String(r[iMa]).trim(); add(P.gbMarca, mk, m, g);
      var pn=r[iP]==null||r[iP]===''?'(sem plano)':String(r[iP]).trim(); P.gbMP[mk]=P.gbMP[mk]||{}; add(P.gbMP[mk], pn, m, g);
      if(iT>=0){ P.gbMT[mk]=P.gbMT[mk]||{}; add(P.gbMT[mk], r[iT]==null?'(vazio)':String(r[iT]).trim(), m, g); } }
    if(iS>=0) add(P.gbStatus, r[iS]==null?'(vazio)':String(r[iS]).trim(), m, g);
  });
  P.meses=Object.keys(meses).sort();
  if(!P.meses.length) throw new Error('A aba "'+base+'" não tem linhas com mês preenchido.');

  /* linhas por plano: bloco "QUANTIDADE DE CLIENTES" (aba Por plano) */
  var linhas=null;
  wb.SheetNames.forEach(function(sn){
    if(linhas) return;
    var r=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:null});
    for(var i=0;i<r.length;i++){
      var row=r[i]||[];
      var hit=row.some(function(x){ return /quantidade de clientes/i.test(String(x||'')); });
      if(!hit) continue;
      var L={};
      for(var j=i+1;j<Math.min(i+15,r.length);j++){
        var rr=r[j]||[]; var m=String(rr[0]||'').trim().match(/^plano\s+([A-Z])$/i);
        if(!m){ if(Object.keys(L).length) break; else continue; }
        var nums=rr.slice(1).filter(function(x){return typeof x==='number';});
        var k=m[1].toUpperCase(); L[k]={};
        P.meses.forEach(function(mm,q){ if(nums[q]!=null) L[k][mm]=nums[q]; });
      }
      if(Object.keys(L).length){ linhas=L; break; }
    }
  });
  P.linhasFonte='planilha';
  if(!linhas){
    linhas={}; var ok=false;
    if(DATA && DATA.linhas){ Object.keys(DATA.linhas).forEach(function(k){ linhas[k]={}; P.meses.forEach(function(mm){ if(DATA.linhas[k][mm]!=null){ linhas[k][mm]=DATA.linhas[k][mm]; ok=true; } }); }); }
    P.linhasFonte= ok?'anterior':'ausente';
  }
  P.linhas=linhas;
  P.fonte=nome; P.aba=base; P.importadaEm=new Date().toISOString();
  return P;
}

/* ---------- motor ---------- */
function gbDoNome(n){ var m=String(n).match(/(\d+)\s*GB/i); return m?+m[1]:null; }
function sub(dst,src){ Object.keys(src).forEach(function(k){ if(!dst[k]) return; Object.keys(src[k]).forEach(function(m){ dst[k][m]=Math.max(0,(dst[k][m]||0)-src[k][m]); }); }); }
function view(exc){
  if(!exc||!exc.length) return {D:DATA,info:null};
  var set={}; exc.forEach(function(x){set[x]=1;});
  var D={gbPlano:clone(DATA.gbPlano),gbTipo:clone(DATA.gbTipo||{}),gbMarca:{},linhas:clone(DATA.linhas||{}),meses:DATA.meses};
  var semGB=!DATA.gbMP;
  Object.keys(DATA.gbMarca||{}).forEach(function(mk){ if(set[normMv(mk)]){ if(DATA.gbMP&&DATA.gbMP[mk]) sub(D.gbPlano,DATA.gbMP[mk]); if(DATA.gbMT&&DATA.gbMT[mk]) sub(D.gbTipo,DATA.gbMT[mk]); } else D.gbMarca[mk]=DATA.gbMarca[mk]; });
  var semLin=exc.filter(function(x){ return !(ML&&ML.mix[x]); }), mesFora=[];
  Object.keys(D.linhas).forEach(function(k){ Object.keys(D.linhas[k]).forEach(function(m){
    var i=ML?ML.meses.indexOf(m):-1; if(i<0){ if(mesFora.indexOf(m)<0) mesFora.push(m); return; }
    var tot=(ML.total[k]||[])[i]||0, x=0; exc.forEach(function(e){ x+=((ML.mix[e]||{})[k]||[])[i]||0; });
    D.linhas[k][m]=tot?D.linhas[k][m]*(1-Math.min(1,x/tot)):D.linhas[k][m];
  }); });
  return {D:D,info:{semGB:semGB,semLin:semLin,mesFora:mesFora}};
}
function compute(exc){
  var V=view(exc===undefined?EXC:exc), D=V.D, ms=MSEL.slice().sort();
  var planos=TAB.map(function(t){
    var nomes=Object.keys(D.gbPlano).filter(function(n){ return gbDoNome(n)===t.gb; });
    var Lm=ms.map(function(m){ return (D.linhas[t.k]||{})[m]; }).filter(function(x){return x!=null;});
    var Gm=ms.map(function(m){ return sum(nomes.map(function(n){ return D.gbPlano[n][m]||0; })); });
    var o=Object.assign({},t,{nomes:nomes,l:avg(Lm),g:avg(Gm),Lm:ms.map(function(m){return (D.linhas[t.k]||{})[m]||0;}),Gm:Gm});
    o.cap=o.l*o.gb; o.cost=o.l*o.custo; o.rev=o.l*o.mv;
    return o;
  });
  var T={l:0,g:0,cap:0,cost:0,rev:0};
  planos.forEach(function(p){ T.l+=p.l; T.g+=p.g; T.cap+=p.cap; T.cost+=p.cost; T.rev+=p.rev; });
  var mapeados={}; planos.forEach(function(p){ p.nomes.forEach(function(n){mapeados[n]=1;}); });
  var legado=Object.keys(D.gbPlano).filter(function(n){return !mapeados[n];}).map(function(n){
    return {nome:n, g:sum(ms.map(function(m){return D.gbPlano[n][m]||0;}))};
  }).filter(function(x){return x.g>0;});
  var totalGB=sum(Object.keys(D.gbPlano).map(function(n){ return sum(ms.map(function(m){return D.gbPlano[n][m]||0;})); }));
  var meses=ms.map(function(m,i){
    var l=sum(planos.map(function(p){return p.Lm[i];})), g=sum(planos.map(function(p){return p.Gm[i];})), c=sum(planos.map(function(p){return p.Lm[i]*p.custo;}));
    return {m:m,l:l,g:g,c:c,gpl:l?g/l:0};
  });
  var tipos=Object.keys(D.gbTipo||{}).map(function(t){ return {nome:t, v:ms.map(function(m){return D.gbTipo[t][m]||0;})}; })
    .filter(function(t){return sum(t.v)>0;}).sort(function(a,b){return sum(b.v)-sum(a.v);});
  var marcas=Object.keys(D.gbMarca||{}).map(function(t){ return {nome:t, v:sum(ms.map(function(m){return D.gbMarca[t][m]||0;}))}; })
    .filter(function(t){return t.v>0;}).sort(function(a,b){return b.v-a.v;});
  var semLinhas=planos.filter(function(p){return !p.l;}).map(function(p){return p.k;});
  return {planos:planos,T:T,meses:meses,tipos:tipos,marcas:marcas,legado:legado,totalGB:totalGB,ms:ms,semLinhas:semLinhas,info:V.info,exc:(exc===undefined?EXC:exc).slice()};
}
function bill(M,gpl,s){ var T=M.T; var used=gpl*T.l; var billed=Math.max(used,s.c/100*T.g); var b=fm(s)*T.l+s.p*billed; if(s.cap) b=Math.min(b,T.cost); return b; }
/* limite da nova grade: maior R$/GB que mantém a margem em R$ de todos os planos */
function gradeLim(M){
  var PG=window.PoolGrade; if(!PG) return null; var C=PG.calc(M), lim=null, k='';
  C.rows.forEach(function(r){ if(!r.l||!(r.uN>0)) return; var pmax=(r.g.p-C.f-r.m0)/r.uN; if(lim===null||pmax<lim){ lim=pmax; k=r.k; } });
  if(lim===null) return null;
  var ganho=function(p){ var t=0; C.rows.forEach(function(r){ t+=r.l*(r.g.p-C.f-p*r.uN); }); return t-C.T.m0; };
  return {p:lim,k:k,C:C,ganho:ganho};
}
function presets(M){
  var T=M.T, F=fm(SIM), gb=SIM.mbgb||1024, L=gradeLim(M);
  if(L && L.p>0){ /* degraus abaixo do limite da nova grade; R$/MB arredondado para baixo para não passar do limite */
    var mk=function(id,nome,d,c){ var mb=Math.max(0.0001,Math.floor(L.p*(1-d)/gb*100000)/100000), p=mb*gb; return {id:id,nome:nome,alvo:d,mb:mb,p:p,c:c,base:'grade',ganho:L.ganho(p),k:L.k,lim:L.p}; };
    return [mk('ab','Abertura',.30,90),mk('al','Alvo',.15,90),mk('li','Limite',0,85)];
  }
  function mb(sv){ return Math.max(0.0001, Math.round(((T.cost*(1-sv))-F*T.l)/T.g/gb*10000)/10000); }
  return [{id:'ab',nome:'Abertura',alvo:.30,mb:mb(.30),p:mb(.30)*gb,c:90,base:'fatura'},
          {id:'al',nome:'Alvo',alvo:.20,mb:mb(.20),p:mb(.20)*gb,c:90,base:'fatura'},
          {id:'li',nome:'Limite',alvo:.10,mb:mb(.10),p:mb(.10)*gb,c:85,base:'fatura'}];
}
/* termos da proposta à SURF, alinhados à nova grade (compartilhado com o PPT) */
function propTerms(M){
  var T=M.T, o=simOut(M,SIM), GF=fm(SIM), gb=SIM.mbgb||1024, PG=window.PoolGrade, C=PG?PG.calc(M):null;
  var fx=function(v,d){ return (+v||0).toLocaleString('pt-BR',{minimumFractionDigits:d,maximumFractionDigits:d}); };
  /* 1 estrutura: fixos abertos item a item */
  var dil=Math.max(1,+SIM.meses||12), vz=(+SIM.voz||0)*(+SIM.min||0), sm=(+SIM.sms||0)*(+SIM.nsms||0);
  var partes=['SIM '+R(SIM.sim,2),'TFF '+R(SIM.tff,2)+'/ano ('+R(SIM.tff/12,2)+')',
    'TFI '+R(SIM.tfi,2)+' + portabilidade '+R(SIM.port,2)+' em '+nf(SIM.portPct,0)+'% das linhas, diluídas em '+dil+' meses ('+R(oneL(SIM)/dil,2)+')'];
  partes.push((+SIM.min||0)>0?'voz '+nf(SIM.min,0)+' min × R$ '+fx(SIM.voz,3)+' ('+R(vz,2)+')':'voz sem premissa de consumo (R$ 0,00)');
  partes.push((+SIM.nsms||0)>0?'SMS '+nf(SIM.nsms,0)+' × R$ '+fx(SIM.sms,2)+' ('+R(sm,2)+')':'SMS sem premissa de consumo (R$ 0,00)');
  /* limite da nova grade: maior R$/GB que mantém a margem em R$ de todos os planos */
  var L=gradeLim(M), lim=L?L.p:null, kLim=L?L.k:'', PRq=presets(M);
  var estr=['Custos fixos por linha: <b>'+R(GF,2)+'/mês</b> = '+partes.join(' + '),
    'Dados: <b>R$ '+fx(SIM.mb,SIM.mb*10000%1>1e-6?5:4)+'/MB</b> (≈ '+R(SIM.p,2)+'/GB)'+(lim!==null?(SIM.p<=lim+1e-9?' · dentro do limite da nova grade':' · <b>acima</b> do limite da nova grade'):''),
    'Pool único de dados para todas as marcas e faixas','Sem ativação para linhas que já estão na base hoje','A ISPX mantém a venda por plano fixo às MVNOs'];
  var preco=[];
  if(PRq[0].base==='grade'){
    PRq.forEach(function(q){ preco.push(q.nome+': <b>'+R(q.p,2)+'/GB</b> (R$ '+fx(q.mb,5)+'/MB)'+(q.alvo?' · −'+(q.alvo*100)+'% do limite':'')+' · margem '+(q.ganho>=0?'+':'−')+Rk(Math.abs(q.ganho))+'/mês vs hoje'); });
    preco.push('Limite = maior preço que mantém a margem em R$ de todos os planos da nova grade (quem trava: plano '+kLim+')');
  } else if(lim!==null){
    preco.push('<b>Nenhum preço de dados fecha:</b> com os fixos atuais, o plano '+kLim+' da nova grade já fica abaixo da margem de hoje');
  } else PRq.forEach(function(q){ preco.push(q.nome+': <b>'+R(q.p,2)+'/GB</b> (−'+(q.alvo*100)+'% na fatura)'); });
  preco.push('Escada de desconto por volume acima do consumo atual');
  var cGB=SIM.c/100*T.g;
  var prot=['<b>Piso para a SURF:</b> compromisso de '+SIM.c+'% do consumo (~'+nf(cGB/1000,1)+' mil GB/mês)','<b>Teto para a ISPX:</b> fatura nunca acima da tabela fixa atual','Excedente no mesmo preço/GB','12 meses com revisão semestral','Relatório diário de consumo por linha e marca'];
  var vant=['Paga pelo que a base usa. Hoje '+pct(T.cap?1-T.g/T.cap:0,0)+' da franquia é paga e não usada'];
  if(C){ var CT=C.T; vant.push('MVNOs pagam '+(CT.eco>=0?'−':'+')+Rk(Math.abs(CT.eco))+'/mês ('+(CT.rev0?(CT.eco>=0?'−':'+')+nf(Math.abs(CT.eco/CT.rev0*100),1)+'%':'')+') e levam '+(CT.gb0?(CT.gbN>=CT.gb0?'+':'−')+nf(Math.abs(CT.gbN/CT.gb0-1)*100,0)+'%':'')+' de GB na franquia');
    vant.push('Margem da ISPX em R$: '+Rk(CT.m0)+' → '+Rk(CT.mN)+'/mês (mantida plano a plano, não em %)'); }
  vant.push('Liberdade para criar planos novos sem renegociar tabela','Para a SURF: receita previsível, contrato mais longo, incentivo a crescer a base');
  var risc=['A ISPX passa a vender fixo e comprar variável. Se o uso subir, a margem cai'];
  if(C){ var gbN=C.T.gbN, worstN=GF*T.l+SIM.p*Math.max(gbN,cGB);
    risc.push('Pior caso sem teto (todos usando 100% da nova franquia): ~'+Rk(worstN)+'/mês'); }
  else risc.push('Pior caso sem teto (todos usando 100%): ~'+Rk(o.worstRaw)+'/mês');
  if(M.marcas[0]) risc.push('Concentração: '+pct(M.marcas[0].v/M.totalGB,0)+' do tráfego vem de uma marca');
  risc.push('Canibalização: linhas migrando para o plano mais barato com a mesma franquia. Precisa de regra de migração');
  if(C){ var perde=[]; C.rows.forEach(function(r){ (r.comp||[]).forEach(function(c){ if(c.delta<0) perde.push('plano '+r.k+' × '+c.s.nome+' ('+pct(c.delta,0).replace('-','−')+')'); }); });
    if(perde.length) risc.push('Concorrente mais barato por GB: '+perde.join('; ')); }
  risc.push((+SIM.min||0)>0||(+SIM.nsms||0)>0?'Voz e SMS são premissa ('+nf(SIM.min||0,0)+' min e '+nf(SIM.nsms||0,0)+' SMS por linha), sem dado real de consumo':'Voz e SMS sem dado de consumo: o custo deles está zerado no modelo');
  risc.push('Take-or-pay: se a base encolher, o piso é pago mesmo assim','Exige monitoramento de consumo em tempo real');
  return {estr:estr,preco:preco,prot:prot,vant:vant,risc:risc,lim:lim,kLim:kLim};
}
function fitLinha(){ /* reta custo = a + b·GB sobre a tabela atual */
  var xs=TAB.map(function(t){return t.gb;}), ys=TAB.map(function(t){return t.custo;});
  var mx=avg(xs), my=avg(ys), num=0, den=0;
  xs.forEach(function(x,i){ num+=(x-mx)*(ys[i]-my); den+=(x-mx)*(x-mx); });
  var b=den?num/den:0; return {a:my-b*mx, b:b};
}
function simOut(M,s){
  var T=M.T, gpl0=T.l?T.g/T.l:0, gpl=gpl0*(1+s.g/100);
  var b=bill(M,gpl,s), be=s.p?(T.cost/T.l-fm(s))/s.p:0;
  var worstRaw=fm(s)*T.l+s.p*Math.max(T.cap, s.c/100*T.g);
  return {gpl0:gpl0,gpl:gpl,bill:b,sav:T.cost?(b-T.cost)/T.cost:0,be:be,worst:bill(M,T.l?T.cap/T.l:0,s),worstRaw:worstRaw,
          mNow:T.rev-T.cost, mNew:T.rev-b, commitGB:s.c/100*T.g};
}

/* ---------- tooltip ---------- */
var tip;
function bindTip(el,html){
  el.addEventListener('mousemove',function(e){ tip.innerHTML=typeof html==='function'?html(e):html; tip.hidden=false;
    tip.style.left=Math.min(e.clientX+14,innerWidth-270)+'px'; tip.style.top=(e.clientY+14)+'px'; });
  el.addEventListener('mouseleave',function(){ tip.hidden=true; });
}

/* ---------- render ---------- */
var TIPCOR=['--s1','--s2','--s3','--s4','--ink-3'];
function render(){
  var app=$('app');
  if(!DATA){ app.innerHTML='<div class="card empty">Nenhuma planilha carregada. Use <b>Importar planilha</b> para abrir o arquivo de consumo da SURF.</div>'; $('stamp').innerHTML=''; return; }
  if(!MSEL){ var sv=ls(KEY_MES); MSEL=(sv||[]).filter(function(m){return DATA.meses.indexOf(m)>=0;}); if(!MSEL.length) MSEL=DATA.meses.slice(); }
  renderMeses(); renderMv();
  var M=compute(); window.PoolApp._M=M;
  var o=simOut(M,SIM), T=M.T, fit=fitLinha(), PR=presets(M);
  var imp=new Date(DATA.importadaEm);
  $('stamp').innerHTML='Fonte: <b>'+esc(DATA.fonte||'—')+'</b><br>'+nf(DATA.registros)+' registros · '+DATA.meses.length+' meses · importada em '+imp.toLocaleDateString('pt-BR')+', '+imp.toLocaleTimeString('pt-BR',{hour:'2-digit',minute:'2-digit'});
  $('hd-periodo').textContent='ISPX · COMPRA DE DADOS SURF · '+mesLbl(M.ms[0]).toUpperCase()+(M.ms.length>1?' A '+mesLbl(M.ms[M.ms.length-1]).toUpperCase():'');
  $('scope').textContent='Médias mensais de '+M.ms.length+' '+(M.ms.length>1?'meses':'mês')+(EXC.length?' · sem '+EXC.map(mvLabel).join(', '):'');
  $('dstatus').innerHTML=DATA===ORIG?'Dados <b>originais</b> do app':'Planilha <b>importada</b>';

  var alerta='';
  if(M.semLinhas.length) alerta='<div class="card" style="background:var(--warn-bg);border-color:transparent;color:var(--warn-ink)">Sem quantidade de linhas para o(s) plano(s) '+M.semLinhas.join(', ')+' no período. Custo e uso desses planos ficam zerados. A planilha precisa do bloco “QUANTIDADE DE CLIENTES” na aba “Por plano”.</div>';
  if(DATA.linhasFonte==='anterior') alerta+='<div class="card" style="background:var(--warn-bg);border-color:transparent;color:var(--warn-ink)">A planilha importada não trouxe a quantidade de linhas por plano. Usei as linhas da carga anterior para os meses coincidentes.</div>';

  var h='';
  if(M.info){ var I=M.info;
    if(I.semGB) alerta+='<div class="card" style="background:var(--warn-bg);border-color:transparent;color:var(--warn-ink)">Esta planilha foi importada numa versão anterior do app e não guarda o consumo por MVNO. Importe a planilha de novo para o filtro de MVNOs funcionar.</div>';
    if(I.semLin.length) alerta+='<div class="card" style="background:var(--warn-bg);border-color:transparent;color:var(--warn-ink)">Sem contagem de linhas para '+I.semLin.map(mvLabel).join(', ')+' no app Benchmark de MVNOs: só o consumo (GB) delas foi retirado, as linhas continuam no total.</div>';
    if(I.mesFora.length) alerta+='<div class="card" style="background:var(--warn-bg);border-color:transparent;color:var(--warn-ink)">Sem linhas por MVNO para '+I.mesFora.map(mesLbl).join(', ')+': nesses meses as linhas não foram ajustadas.</div>';
  }
  h+=alerta;
  if(EXC.length) h+=impacto(M);
  h+='<section class="kpis">'+
    kpi(nf(T.l),'linhas/mês (A–F)')+kpi(nf(T.cap),'GB contratados/mês')+kpi(nf(T.g),'GB usados/mês')+
    kpi(pct(T.cap?T.g/T.cap:0),'da franquia usada',true)+kpi(Rk(T.cost),'pago à SURF/mês')+kpi(R(T.g?T.cost/T.g:0,2),'custo por GB usado',true)+'</section>';

  h+='<section class="thesis"><span class="eyebrow">Tese da negociação</span>'+
    '<p class="big">A ISPX usa '+pct(T.cap?T.g/T.cap:0,0)+' da franquia que compra. O GB contratado sai a '+R(T.cap?T.cost/T.cap:0,2)+', mas o GB usado sai a '+R(T.g?T.cost/T.g:0,2)+'.</p>'+
    '<p>São ~'+nf((T.cap-T.g)/1000,0)+' mil GB/mês pagos e não usados. Um pool troca a franquia fixa por consumo real, com um piso de compromisso que dá previsibilidade à SURF.</p></section>';

  h+='<div class="g2"><section class="card"><div class="sec-hd"><h2>Franquia contratada × consumo real</h2><p class="muted small">GB por mês, média do período. Cinza = franquia total (base + portabilidade); azul = consumo.</p></div>'+
    '<div class="legend"><span><i style="background:var(--contract)"></i>Contratado</span><span><i style="background:var(--used)"></i>Usado</span></div><div class="bullets" id="bullets"></div></section>'+
    '<section class="card"><div class="sec-hd"><h2>Tendência mensal</h2><p class="muted small">Consumo médio por linha (GB/linha/mês, planos A–F).</p></div><div class="chartbox" id="trend"></div>'+
    '<div class="tbl"><table><thead><tr><th>Mês</th><th>Linhas</th><th>GB usados</th><th>GB/linha</th><th>Pago SURF</th></tr></thead><tbody>'+
    M.meses.map(function(m){return '<tr><td>'+mesLbl(m.m)+'</td><td>'+nf(m.l)+'</td><td>'+nf(m.g)+'</td><td>'+nf(m.gpl,2)+'</td><td>'+R(m.c)+'</td></tr>';}).join('')+
    '</tbody></table></div></section></div>';

  if(window.PoolProj) h+=window.PoolProj.html(M);

  h+='<section class="card"><div class="sec-hd"><h2>Raio-X por plano</h2><p class="muted small">Médias mensais. Custo SURF e preço MVNO são editáveis. As mudanças recalculam tudo e ficam salvas neste computador.</p></div>'+
    '<div class="tbl"><table style="min-width:900px"><thead><tr><th>Plano</th><th>GB franquia</th><th>Custo SURF</th><th>Preço MVNO</th><th>Linhas</th><th>GB usados</th><th>GB/linha</th><th>Uso</th><th>Pago SURF</th><th>R$/GB contratado</th><th>R$/GB usado</th><th>% do custo</th></tr></thead><tbody>'+
    M.planos.map(function(p,i){ return '<tr><td>'+p.k+'</td><td>'+p.base+'+'+p.port+' = '+p.gb+'</td>'+
      '<td><input type="number" step="0.01" min="0" data-i="'+i+'" data-f="custo" value="'+p.custo+'" aria-label="Custo SURF plano '+p.k+'"></td>'+
      '<td><input type="number" step="0.01" min="0" data-i="'+i+'" data-f="mv" value="'+p.mv+'" aria-label="Preço MVNO plano '+p.k+'"></td>'+
      '<td>'+nf(p.l)+'</td><td>'+nf(p.g)+'</td><td>'+nf(p.l?p.g/p.l:0,2)+'</td><td>'+pct(p.cap?p.g/p.cap:0)+'</td><td>'+R(p.cost)+'</td><td>'+R(p.custo/p.gb,2)+'</td>'+
      '<td class="'+(p.g&&p.cost/p.g>7?'hot':'')+'">'+R(p.g?p.cost/p.g:0,2)+'</td><td>'+pct(T.cost?p.cost/T.cost:0)+'</td></tr>'; }).join('')+
    '</tbody><tfoot><tr><td>Total</td><td></td><td></td><td></td><td>'+nf(T.l)+'</td><td>'+nf(T.g)+'</td><td>'+nf(T.l?T.g/T.l:0,2)+'</td><td>'+pct(T.cap?T.g/T.cap:0)+'</td><td>'+R(T.cost)+'</td><td>'+R(T.cap?T.cost/T.cap:0,2)+'</td><td>'+R(T.g?T.cost/T.g:0,2)+'</td><td>100%</td></tr></tfoot></table></div>'+
    '<p class="muted small">Pela tabela atual, cada plano custa ≈ '+R(fit.a,2)+' fixos + '+R(fit.b,2)+' por GB (reta que melhor explica os 6 preços). É uma inferência sobre a tabela, não um número da SURF.</p></section>';

  if(window.PoolSurf) h+=window.PoolSurf.html(M);
  h+='<section class="card"><div class="sec-hd"><h2>Simulador de proposta</h2><p class="muted small">Custos fixos por linha da proposta (SIM, TFF, TFI e portabilidade diluídas, voz e SMS) + dados pagos por MB consumido. Os botões (Abertura, Alvo e Limite) usam o preço-limite da Nova grade: Limite é o maior R$/MB que mantém a margem em R$ de todos os planos, Alvo fica 15% abaixo e Abertura 30% abaixo; compromisso mínimo e teto são condições que podemos pedir. O PPT exporta o cenário desta tela.</p></div>'+
    '<div class="sim"><div class="ctrls"><div class="seggrp" id="presets">'+PR.map(function(q){return '<button class="chip" data-p="'+q.id+'" title="'+(q.base==='grade'?(q.alvo?'−'+(q.alvo*100)+'% do limite da nova grade':'Limite da nova grade: margem em R$ mantida em todos os planos (trava no plano '+q.k+')'):'−'+(q.alvo*100)+'% na fatura de hoje')+'">'+q.nome+' '+(q.base==='grade'?R(q.p,2)+'/GB':'−'+(q.alvo*100)+'%')+'</button>';}).join('')+'</div>'+
    '<div class="ctl"><label>Custos fixos por linha/mês <output>'+R(fm(SIM),2)+'</output></label><span class="hint">SIM, TFF, TFI/portabilidade diluídas, voz e SMS. Edite na tabela da proposta acima.</span></div>'+
    ctl('p','Preço do GB (R$/MB × '+SIM.mbgb+')',1,8,0.01,'Hoje o GB contratado custa de '+R(Math.min.apply(null,TAB.map(function(t){return t.custo/t.gb;})),2)+' a '+R(Math.max.apply(null,TAB.map(function(t){return t.custo/t.gb;})),2))+
    ctl('c','Compromisso mínimo',0,110,5,'% do consumo médio atual pago mesmo sem uso (take-or-pay)')+
    ctl('g','Variação do consumo/linha',-30,150,5,'Teste de estresse: e se a base passar a usar mais?')+
    '<label class="toggle" for="s-cap"><input id="s-cap" type="checkbox"'+(SIM.cap?' checked':'')+'> Teto: fatura nunca passa a tabela fixa atual</label></div>'+
    '<div style="display:flex;flex-direction:column;gap:14px;min-width:0"><div class="outs">'+
    '<div class="out"><span class="v" id="o-bill"></span><span class="l">fatura pool/mês</span></div>'+
    '<div class="out" id="o-savb"><span class="v" id="o-sav"></span><span class="l">vs. hoje ('+Rk(T.cost)+')</span></div>'+
    '<div class="out"><span class="v" id="o-be"></span><span class="l">GB/linha em que o pool empata</span></div>'+
    '<div class="out" id="o-worstb"><span class="v" id="o-worst"></span><span class="l">pior caso (100% de uso)</span></div></div>'+
    '<div class="chartbox" id="simchart"></div>'+
    '<div class="legend"><span><i style="background:var(--contract)"></i>Modelo fixo atual</span><span><i style="background:var(--used)"></i>Pool proposto</span><span><i style="background:var(--s2);border-radius:50%"></i>Consumo simulado</span></div>'+
    '<p class="small muted" id="o-margin"></p></div></div></section>';

  if(window.PoolGrade) h+=window.PoolGrade.html(M);
  if(window.PoolNovos) h+=window.PoolNovos.html(M);
  if(window.PoolComp) h+=window.PoolComp.html(M);

  h+='<div class="g2e"><section class="card"><div class="sec-hd"><h2>O que compõe o tráfego</h2><p class="muted small">GB por tipo de consumo registrado na SURF.</p></div>'+
    '<div class="legend">'+M.tipos.map(function(t,i){return '<span><i style="background:var('+TIPCOR[Math.min(i,4)]+')"></i>'+esc(cap1(t.nome))+'</span>';}).join('')+'</div><div class="stack" id="stack"></div>'+
    '<p class="small muted">'+tipoNota(M)+'</p></section>'+
    '<section class="card"><div class="sec-hd"><h2>Concentração por marca</h2><p class="muted small">Participação no consumo do período ('+nf(M.totalGB)+' GB).</p></div><div id="mvno" style="display:flex;flex-direction:column;gap:7px"></div>'+
    '<p class="small muted">'+(M.marcas[0]?'A '+esc(M.marcas[0].nome)+' responde por <b>'+pct(M.marcas[0].v/M.totalGB)+'</b> do tráfego. No pool, o comportamento de uma marca grande move a fatura inteira.':'')+'</p></section></div>';

  var PT=propTerms(M), li=function(a){ return a.map(function(x){return '<li>'+x+'</li>';}).join(''); };
  h+='<section class="card"><div class="sec-hd"><h2>Proposta sugerida para levar à SURF</h2><p class="muted small">Os valores saem do simulador (nossa proposta) e da Nova grade: o preço-limite é o maior R$/GB que mantém a margem em R$ de todos os planos da nova grade.</p></div><div class="terms">'+
    '<div class="term"><span class="eyebrow">1 · Estrutura</span><ul>'+li(PT.estr)+'</ul></div>'+
    '<div class="term"><span class="eyebrow">2 · Preço</span><ul>'+li(PT.preco)+'</ul></div>'+
    '<div class="term"><span class="eyebrow">3 · Proteções</span><ul>'+li(PT.prot)+'</ul></div></div></section>';
  h+='<section class="card pr"><div style="display:flex;flex-direction:column;gap:8px"><h3 class="good">Vantagens do pooling</h3><ul>'+li(PT.vant)+'</ul></div>'+
    '<div style="display:flex;flex-direction:column;gap:8px"><h3 class="bad">Riscos a controlar</h3><ul>'+li(PT.risc)+'</ul></div></section>';

  h+='<section class="card"><h3>Notas sobre os dados</h3><ul class="notes">'+
    '<li>Linhas por plano: bloco “QUANTIDADE DE CLIENTES” da aba “Por plano”. A coluna qtd_registros da aba de consumo conta registros, não linhas.</li>'+
    '<li>Filtro de MVNOs: o consumo sai direto da planilha por marca. As linhas de cada plano são reduzidas pela participação da MVNO na base ativa daquele plano no mês, segundo o app Benchmark de MVNOs'+(ML?' ('+esc(ML.fonte)+')':'')+'. HOMENET e HNET MOVEL são tratadas como a mesma MVNO.</li>'+
    (M.legado.length?'<li>Planos sem faixa na tabela ('+M.legado.map(function(x){return esc(x.nome);}).join(', ')+') somam '+nf(sum(M.legado.map(function(x){return x.g;})))+' GB no período ('+pct(sum(M.legado.map(function(x){return x.g;}))/M.totalGB)+') e ficaram fora do raio-X.</li>':'')+
    '<li>A franquia contratada usa o total com portabilidade. Se parte das linhas não tem o bônus, a franquia real é menor e o uso é maior.</li>'+
    '<li>A margem ISPX supõe que todas as linhas pagam o “Preço MVNO” da tabela acima.</li></ul></section>';

  app.innerHTML=h;
  drawBullets(M); drawTrend(M); drawStack(M); drawMarcas(M);
  if(window.PoolProj) window.PoolProj.bind(M);
  if(window.PoolSurf) window.PoolSurf.bind(M);
  if(window.PoolComp){ window.PoolComp.draw(M); window.PoolComp.bind(M); }
  if(window.PoolGrade) window.PoolGrade.bind();
  if(window.PoolNovos) window.PoolNovos.bind(M);
  bindTab(); bindSim(M);
  updSim(M);
}
function cap1(s){ s=String(s).toLowerCase(); return s.charAt(0).toUpperCase()+s.slice(1); }
function kpi(v,l,al){ return '<div class="kpi'+(al?' alert':'')+'"><span class="v">'+v+'</span><span class="l">'+l+'</span></div>'; }
function ctl(id,lab,mi,ma,st,hint){ return '<div class="ctl"><label for="s-'+id+'">'+lab+' <output id="o-'+id+'"></output></label><input id="s-'+id+'" type="range" min="'+mi+'" max="'+ma+'" step="'+st+'" value="'+SIM[id]+'"><span class="hint">'+hint+'</span></div>'; }
function tipoNota(M){
  var tot=sum(M.tipos.map(function(t){return sum(t.v);})); if(!tot) return '';
  var b=M.tipos.filter(function(t){return /bonus/i.test(t.nome);})[0], sd=M.tipos.filter(function(t){return /sem debito/i.test(t.nome);})[0];
  var s=[];
  if(b) s.push('<b>'+pct(sum(b.v)/tot,0)+'</b> do tráfego é “'+esc(cap1(b.nome))+'”');
  if(sd) s.push('<b>'+pct(sum(sd.v)/tot,0)+'</b> é “'+esc(cap1(sd.nome))+'”');
  return (s.length?s.join(' e ')+'. ':'')+'Antes de fechar, é preciso confirmar com a SURF quais tipos entram no pool.';
}
function mvGrupos(){ /* marcas da planilha agrupadas pelo nome normalizado */
  var g={}; Object.keys(DATA.gbMarca||{}).forEach(function(mk){ var k=normMv(mk); g[k]=g[k]||{k:k,nomes:[],gb:0}; g[k].nomes.push(mk);
    g[k].gb+=sum(MSEL.map(function(m){return DATA.gbMarca[mk][m]||0;})); });
  return Object.keys(g).map(function(k){return g[k];}).filter(function(x){return x.gb>0||EXC.indexOf(x.k)>=0;}).sort(function(a,b){return b.gb-a.gb;});
}
function mvLabel(k){ var g=Object.keys(DATA&&DATA.gbMarca||{}).filter(function(mk){return normMv(mk)===k;}); return g.length>1?g.join(' / '):(g[0]||k); }
function renderMv(){
  var el=$('seg-mv'); if(!el) return;
  var G=mvGrupos(), tot=sum(G.map(function(x){return x.gb;}))||1, ativos=G.filter(function(x){return EXC.indexOf(x.k)<0;}).length;
  el.innerHTML='<span class="gl">MVNOs</span><button class="chip'+(EXC.length?' on':'')+'" id="mvbtn" aria-expanded="'+MVOPEN+'">'+(EXC.length?ativos+' de '+G.length+' · sem '+esc(EXC.map(mvLabel).join(', ')):'Todas ('+G.length+')')+' ▾</button>'+
    (G[0]?'<button class="chip" id="mvtop">'+(EXC.length===1&&EXC[0]===G[0].k?'Incluir '+esc(mvLabel(G[0].k)):'Sem '+esc(mvLabel(G[0].k)))+'</button>':'')+
    (EXC.length?'<button class="chip" id="mvall">Todas</button>':'');
  var pn=$('mvpanel');
  pn.hidden=!MVOPEN;
  if(MVOPEN){
    pn.innerHTML='<div class="mvhd"><b>Considerar na análise</b><span class="muted small">Desmarque para simular sem a MVNO. % = participação no consumo do período.</span>'+
      '<span style="margin-left:auto;display:flex;gap:6px"><button class="chip" id="mvmk">Marcar todas</button><button class="chip" id="mvcl">Fechar</button></span></div><div class="mvlist">'+
      G.map(function(x){ var on=EXC.indexOf(x.k)<0, hasL=!!(ML&&ML.mix[x.k]);
        return '<label class="mvitem"><input type="checkbox" data-mv="'+esc(x.k)+'"'+(on?' checked':'')+'><span class="nm">'+esc(mvLabel(x.k))+(hasL?'':' <span class="muted" title="Sem contagem de linhas no app Benchmark de MVNOs">*</span>')+'</span><span class="pc">'+pct(x.gb/tot)+'</span><span class="bar"><i style="width:'+(x.gb/G[0].gb*100)+'%"></i></span></label>'; }).join('')+
      '</div>'+(G.some(function(x){return !(ML&&ML.mix[x.k]);})?'<p class="muted small">* sem contagem de linhas: ao desmarcar, só o consumo sai.</p>':'');
    pn.querySelectorAll('[data-mv]').forEach(function(c){ c.onchange=function(){ var k=c.dataset.mv, i=EXC.indexOf(k);
      if(c.checked){ if(i>=0) EXC.splice(i,1); } else { if(G.length-EXC.length<=1){ c.checked=true; toast('Deixe pelo menos uma MVNO.'); return; } if(i<0) EXC.push(k); }
      ls(KEY_EXC,EXC); render(); }; });
    $('mvmk').onclick=function(){ EXC=[]; ls(KEY_EXC,EXC); render(); };
    $('mvcl').onclick=function(){ MVOPEN=false; renderMv(); };
  }
  $('mvbtn').onclick=function(){ MVOPEN=!MVOPEN; renderMv(); };
  if($('mvtop')) $('mvtop').onclick=function(){ EXC=(EXC.length===1&&EXC[0]===G[0].k)?[]:[G[0].k]; ls(KEY_EXC,EXC); render(); };
  if($('mvall')) $('mvall').onclick=function(){ EXC=[]; ls(KEY_EXC,EXC); render(); };
}
function impacto(M){
  var A=compute([]), oA=simOut(A,SIM), oM=simOut(M,SIM);
  var rows=[['Linhas/mês',nf(A.T.l),nf(M.T.l)],['GB usados/mês',nf(A.T.g),nf(M.T.g)],['Franquia usada',pct(A.T.cap?A.T.g/A.T.cap:0),pct(M.T.cap?M.T.g/M.T.cap:0)],
    ['GB por linha',nf(A.T.l?A.T.g/A.T.l:0,2),nf(M.T.l?M.T.g/M.T.l:0,2)],['Pago à SURF/mês',Rk(A.T.cost),Rk(M.T.cost)],['Custo por GB usado',R(A.T.g?A.T.cost/A.T.g:0,2),R(M.T.g?M.T.cost/M.T.g:0,2)],
    ['Fatura pool (cenário atual)',Rk(oA.bill),Rk(oM.bill)],['Economia do pool',(oA.sav<=0?'−':'+')+nf(Math.abs(oA.sav*100),1)+'%',(oM.sav<=0?'−':'+')+nf(Math.abs(oM.sav*100),1)+'%'],
    ['Empate do pool (GB/linha)',nf(oA.be,2),nf(oM.be,2)]];
  var dU=(M.T.cap?M.T.g/M.T.cap:0)-(A.T.cap?A.T.g/A.T.cap:0);
  return '<section class="card" style="border-color:var(--accent)"><div class="sec-hd"><h2>Impacto do recorte: sem '+esc(M.exc.map(mvLabel).join(', '))+'</h2>'+
    '<p class="muted small">Tudo abaixo já está recalculado sem '+(M.exc.length>1?'essas MVNOs':'essa MVNO')+'. A tabela compara com a carteira inteira no mesmo cenário do simulador.</p></div>'+
    '<div class="tbl"><table style="min-width:480px"><thead><tr><th>Indicador</th><th>Todas as MVNOs</th><th>Recorte</th></tr></thead><tbody>'+
    rows.map(function(r){return '<tr><td>'+r[0]+'</td><td>'+r[1]+'</td><td>'+r[2]+'</td></tr>';}).join('')+'</tbody></table></div>'+
    '<p class="small muted">Sem '+(M.exc.length>1?'elas':'ela')+', o uso da franquia '+(Math.abs(dU)<0.0005?'fica igual':(dU<0?'cai':'sobe')+' '+nf(Math.abs(dU*100),1)+' ponto'+(Math.abs(dU*100)>=1.95?'s':'')+' percentual'+(Math.abs(dU*100)>=1.95?'is':''))+' ('+pct(A.T.cap?A.T.g/A.T.cap:0)+' → '+pct(M.T.cap?M.T.g/M.T.cap:0)+').</p></section>';
}
function renderMeses(){
  var g=$('seg-mes'); g.innerHTML='<span class="gl">Meses</span>'+DATA.meses.map(function(m){return '<button class="chip sm" data-m="'+m+'" aria-pressed="'+(MSEL.indexOf(m)>=0)+'">'+mesLbl(m)+'</button>';}).join('');
  g.querySelectorAll('[data-m]').forEach(function(b){ b.onclick=function(){
    var m=b.dataset.m, i=MSEL.indexOf(m);
    if(i>=0){ if(MSEL.length===1){ toast('Deixe pelo menos um mês selecionado.'); return; } MSEL.splice(i,1); } else MSEL.push(m);
    ls(KEY_MES,MSEL); render(); }; });
}
function drawBullets(M){
  var el=$('bullets'), mx=Math.max.apply(null,M.planos.map(function(p){return p.cap;}))||1;
  M.planos.forEach(function(p){ var r=document.createElement('div'); r.className='brow';
    r.innerHTML='<div class="nm">Plano '+p.k+'<span>'+p.gb+' GB · '+nf(p.l)+' linhas</span></div><div class="btrack" style="width:'+Math.max(1,p.cap/mx*100)+'%"><div class="bfill" style="width:'+(p.cap?Math.min(100,p.g/p.cap*100):0)+'%"></div></div><div class="pct">'+pct(p.cap?p.g/p.cap:0,0)+'</div>';
    el.appendChild(r); bindTip(r.querySelector('.btrack'),'<div>Plano '+p.k+' ('+p.gb+' GB)</div>Contratado <b>'+nf(p.cap)+' GB</b><br>Usado <b>'+nf(p.g)+' GB</b><br>Não usado <b>'+nf(p.cap-p.g)+' GB</b>'); });
}
function drawTrend(M){
  var box=$('trend'), W=box.clientWidth||400, H=170, pl=36, pr=16, pt=20, pb=26, n=M.meses.length;
  var y1=Math.max(2,Math.ceil(Math.max.apply(null,M.meses.map(function(m){return m.gpl;}))*1.2));
  var x=function(i){return pl+(W-pl-pr)*(i+.5)/n;}, y=function(v){return pt+(H-pt-pb)*(1-v/y1);};
  var s='<svg width="'+W+'" height="'+H+'" viewBox="0 0 '+W+' '+H+'" role="img" aria-label="GB por linha por mês">';
  var step=y1<=6?2:Math.ceil(y1/4);
  for(var v=0;v<=y1;v+=step) s+='<line x1="'+pl+'" x2="'+(W-pr)+'" y1="'+y(v)+'" y2="'+y(v)+'" stroke="'+css('--grid')+'"/><text x="'+(pl-6)+'" y="'+(y(v)+4)+'" text-anchor="end">'+v+'</text>';
  var d=M.meses.map(function(m,i){return (i?'L':'M')+x(i)+' '+y(m.gpl);}).join(' ');
  if(n>1) s+='<path d="'+d+' L'+x(n-1)+' '+y(0)+' L'+x(0)+' '+y(0)+' Z" fill="'+css('--used')+'" opacity=".12"/><path d="'+d+'" fill="none" stroke="'+css('--used')+'" stroke-width="2"/>';
  M.meses.forEach(function(m,i){ s+='<circle cx="'+x(i)+'" cy="'+y(m.gpl)+'" r="5" fill="'+css('--used')+'" stroke="'+css('--surface')+'" stroke-width="2"/><text x="'+x(i)+'" y="'+(y(m.gpl)-10)+'" text-anchor="middle" style="fill:'+css('--ink')+';font-family:var(--mono)">'+nf(m.gpl,2)+'</text><text x="'+x(i)+'" y="'+(H-8)+'" text-anchor="middle">'+mesLbl(m.m)+'</text>'; });
  box.innerHTML=s+'</svg>';
}
function drawStack(M){
  var el=$('stack');
  M.ms.forEach(function(m,i){ var tot=sum(M.tipos.map(function(t){return t.v[i];}))||1; var r=document.createElement('div'); r.className='srow';
    r.innerHTML='<div>'+mesLbl(m)+'</div><div class="sbar">'+M.tipos.map(function(t,j){return t.v[i]>0?'<div data-j="'+j+'" style="flex:'+t.v[i]+';background:var('+TIPCOR[Math.min(j,4)]+')"></div>':'';}).join('')+'</div>';
    el.appendChild(r);
    r.querySelectorAll('.sbar div').forEach(function(dv){ var j=+dv.dataset.j; bindTip(dv,'<div>'+esc(cap1(M.tipos[j].nome))+' · '+mesLbl(m)+'</div><b>'+nf(M.tipos[j].v[i])+' GB</b> ('+pct(M.tipos[j].v[i]/tot)+')'); });
  });
}
function drawMarcas(M){
  var el=$('mvno'), top=M.marcas.slice(0,10), rest=M.marcas.slice(10);
  var list=top.map(function(x){return {nome:x.nome,v:x.v};});
  if(rest.length) list.push({nome:'Outras '+rest.length+' marcas',v:sum(rest.map(function(x){return x.v;})),other:true});
  var mx=list.length?Math.max.apply(null,list.map(function(x){return x.v;})):1;
  list.forEach(function(x){ var r=document.createElement('div'); r.className='mrow'+(x.other?' other':'');
    r.innerHTML='<div>'+esc(x.nome)+'</div><div><div class="mbar" style="width:'+(x.v/mx*100)+'%"></div></div><div class="pct">'+pct(x.v/M.totalGB)+'</div>';
    el.appendChild(r); bindTip(r.querySelector('.mbar'),'<div>'+esc(x.nome)+'</div><b>'+nf(x.v)+' GB</b> no período<br>~'+nf(x.v/M.ms.length)+' GB/mês'); });
}
function bindTab(){
  document.querySelectorAll('td input[data-f]').forEach(function(inp){
    inp.addEventListener('change',function(){ var v=parseFloat(String(inp.value).replace(',','.')); if(!isFinite(v)||v<0){ inp.value=TAB[+inp.dataset.i][inp.dataset.f]; return; }
      TAB[+inp.dataset.i][inp.dataset.f]=v; ls(KEY_TAB,TAB); render(); toast('Tabela atualizada e salva neste computador.'); });
  });
}
function bindSim(M){
  ['p','c','g'].forEach(function(k){ $('s-'+k).addEventListener('input',function(){ SIM[k]=+this.value; if(k==='p') SIM.mb=SIM.p/(SIM.mbgb||1024); ls(KEY_SIM,SIM); updSim(M); }); });
  ['p'].forEach(function(k){ $('s-'+k).addEventListener('change',function(){ render(); }); });
  $('s-cap').addEventListener('change',function(){ SIM.cap=this.checked; ls(KEY_SIM,SIM); updSim(M); });
  document.querySelectorAll('#presets [data-p]').forEach(function(b){ b.onclick=function(){ var q=presets(M).filter(function(x){return x.id===b.dataset.p;})[0];
    SIM.mb=q.mb; syncP(SIM); SIM.c=q.c; SIM.g=0; ls(KEY_SIM,SIM); render(); }; });
  var rt; window.onresize=function(){ clearTimeout(rt); rt=setTimeout(function(){ drawTrend(M); drawSim(M); if(window.PoolComp) window.PoolComp.draw(M); if(window.PoolNovos) window.PoolNovos.draw(M); if(window.PoolProj) window.PoolProj.draw(M); },150); };
}
function updSim(M){
  var o=simOut(M,SIM);
  $('o-p').textContent=R(SIM.p,2)+' (R$ '+(SIM.mb||0).toLocaleString('pt-BR',{minimumFractionDigits:4,maximumFractionDigits:5})+'/MB)'; $('o-c').textContent=SIM.c+'%'; $('o-g').textContent=(SIM.g>0?'+':'')+SIM.g+'%';
  $('o-bill').textContent=Rk(o.bill);
  $('o-sav').textContent=(o.sav<=0?'−':'+')+nf(Math.abs(o.sav*100),1)+'%';
  $('o-savb').className='out '+(o.sav<-0.001?'good':o.sav>0.001?'bad':'');
  $('o-be').textContent=o.be>0?nf(o.be,2)+' GB':'—';
  $('o-worst').textContent=Rk(o.worst); $('o-worstb').className='out '+(o.worst>M.T.cost+1?'bad':'');
  var T=M.T;
  $('o-margin').innerHTML='Margem bruta ISPX sobre as MVNOs: hoje <b>'+Rk(o.mNow)+'/mês ('+pct(T.rev?o.mNow/T.rev:0)+')</b> → neste cenário, com os preços de hoje às MVNOs, <b>'+Rk(o.mNew)+'/mês ('+pct(T.rev?o.mNew/T.rev:0)+')</b>'+(function(){ var G=window.PoolGrade?window.PoolGrade.calc(M):null; return G?'; com a nova grade (MVNOs pagando menos), <b>'+Rk(G.T.mN)+'/mês</b>':''; })()+'. O pool empata com o fixo em '+nf(o.be,2)+' GB/linha, '+nf(o.gpl0?(o.be/o.gpl0-1)*100:0,0)+'% acima do consumo atual ('+nf(o.gpl0,2)+').';
  var PR=presets(M);
  document.querySelectorAll('#presets [data-p]').forEach(function(b){ var q=PR.filter(function(x){return x.id===b.dataset.p;})[0]; b.classList.toggle('on',Math.abs(q.mb-SIM.mb)<1e-7&&q.c===SIM.c&&SIM.g===0); });
  drawSim(M);
  if(window.PoolComp && window.PoolComp.POOL) window.PoolComp.draw(M);
}
function drawSim(M){
  var box=$('simchart'); if(!box) return; var T=M.T, o=simOut(M,SIM);
  var W=box.clientWidth||600, H=260, pl=58, pr=16, pt=14, pb=32;
  var xm=Math.max(10,Math.ceil(Math.max(o.gpl,o.be,T.l?T.cap/T.l:0)*1.05/2)*2); xm=Math.min(xm,Math.max(14,Math.ceil(o.gpl*1.3)));
  var ymax=Math.max(T.cost,bill(M,xm,Object.assign({},SIM,{cap:false})))*1.1;
  var st=Math.pow(10,Math.floor(Math.log10(ymax/5))); var steps=[1,2,2.5,5,10]; var stp=st; for(var i=0;i<steps.length;i++){ if(ymax/(st*steps[i])<=6){ stp=st*steps[i]; break; } }
  var ym=Math.ceil(ymax/stp)*stp;
  var x=function(v){return pl+(W-pl-pr)*v/xm;}, y=function(v){return pt+(H-pt-pb)*(1-v/ym);};
  var s='<svg width="'+W+'" height="'+H+'" viewBox="0 0 '+W+' '+H+'" role="img" aria-label="Custo mensal por consumo médio por linha">';
  for(var v=0;v<=ym+1;v+=stp) s+='<line x1="'+pl+'" x2="'+(W-pr)+'" y1="'+y(v)+'" y2="'+y(v)+'" stroke="'+css('--grid')+'"/><text x="'+(pl-6)+'" y="'+(y(v)+4)+'" text-anchor="end">'+(v?nf(v/1000)+' mil':'0')+'</text>';
  var xs=xm<=16?2:Math.ceil(xm/8);
  for(var g=0;g<=xm;g+=xs) s+='<text x="'+x(g)+'" y="'+(H-12)+'" text-anchor="middle">'+g+'</text>';
  s+='<text x="'+(W-pr)+'" y="'+(H-1)+'" text-anchor="end">GB por linha/mês</text>';
  s+='<line x1="'+x(0)+'" x2="'+x(xm)+'" y1="'+y(T.cost)+'" y2="'+y(T.cost)+'" stroke="'+css('--contract')+'" stroke-width="2" stroke-dasharray="6 4"/>';
  var d=''; for(var q=0;q<=xm+1e-9;q+=xm/140) d+=(q?'L':'M')+x(q).toFixed(1)+' '+y(bill(M,q,SIM)).toFixed(1)+' ';
  s+='<path d="'+d+'" fill="none" stroke="'+css('--used')+'" stroke-width="2"/>';
  var bx=o.bill; s+='<circle cx="'+x(Math.min(o.gpl,xm))+'" cy="'+y(bx)+'" r="6" fill="'+css('--s2')+'" stroke="'+css('--surface')+'" stroke-width="2"/>';
  s+='<text x="'+Math.min(x(o.gpl)+10,W-130)+'" y="'+(y(bx)+18)+'" style="fill:'+css('--ink')+'">'+nf(o.gpl,2)+' GB → '+Rk(bx)+'</text>';
  s+='<text x="'+x(xm)+'" y="'+(y(T.cost)-6)+'" text-anchor="end">Fixo hoje '+Rk(T.cost)+'</text>';
  s+='<line id="xh" y1="'+pt+'" y2="'+(H-pb)+'" stroke="'+css('--ink-3')+'" opacity="0"/><rect id="hit" x="'+pl+'" y="'+pt+'" width="'+(W-pl-pr)+'" height="'+(H-pt-pb)+'" fill="transparent"/>';
  box.innerHTML=s+'</svg>';
  var hit=box.querySelector('#hit'), xh=box.querySelector('#xh');
  bindTip(hit,function(e){ var r=box.getBoundingClientRect(); var v=Math.max(0,Math.min(xm,(e.clientX-r.left-pl)/(W-pl-pr)*xm)); xh.setAttribute('x1',x(v)); xh.setAttribute('x2',x(v)); xh.setAttribute('opacity',.5);
    var b=bill(M,v,SIM); return '<div>'+nf(v,1)+' GB/linha ('+(o.gpl0?nf((v/o.gpl0-1)*100,0):0)+'% vs hoje)</div>Pool <b>'+Rk(b)+'</b><br>Fixo <b>'+Rk(T.cost)+'</b><br>Diferença <b>'+(b<=T.cost?'−':'+')+Rk(Math.abs(b-T.cost))+'</b>'; });
  hit.addEventListener('mouseleave',function(){ xh.setAttribute('opacity',0); });
}

/* ---------- ações ---------- */
function importar(file){
  var fr=new FileReader();
  fr.onload=function(){
    try{
      var wb=XLSX.read(new Uint8Array(fr.result),{type:'array',cellDates:true});
      var P=parseWorkbook(wb,file.name);
      DATA=P; MSEL=null; if(!ls(KEY_SIM)) SAVEDSIM=null; ls(KEY_MES,null); ls(KEY_DATA,P); render();
      toast('Planilha importada: '+P.meses.length+' meses, '+nf(P.registros)+' registros.'+(P.linhasFonte==='ausente'?' Atenção: sem linhas por plano.':''));
    }catch(e){ toast('Não deu para importar: '+e.message); console.error(e); }
  };
  fr.onerror=function(){ toast('Não consegui ler o arquivo. Ele está aberto no Excel?'); };
  fr.readAsArrayBuffer(file);
}
function init(){
  tip=$('tip');
  $('btnImport').onclick=function(){ $('fileInput').click(); };
  $('fileInput').onchange=function(){ if(this.files[0]) importar(this.files[0]); this.value=''; };
  $('btnReset').onclick=function(){ DATA=ORIG; TAB=clone(TAB0); SIM=clone(SIM0); SAVEDSIM=null; MSEL=null; EXC=[]; [KEY_DATA,KEY_TAB,KEY_SIM,KEY_MES,KEY_EXC].forEach(function(k){ls(k,null);}); if(window.PoolComp) window.PoolComp.reset(); if(window.PoolGrade) window.PoolGrade.reset(); if(window.PoolProj) window.PoolProj.reset(); if(window.PoolNovos) window.PoolNovos.reset(); render(); toast('Voltei para a planilha e a tabela originais.'); };
  $('btnPpt').onclick=function(){ if(!DATA){ toast('Importe uma planilha antes.'); return; } var b=this; b.disabled=true; var t=b.lastChild.textContent; b.lastChild.textContent=' Gerando…';
    Promise.resolve().then(function(){ return window.PoolPPT.exportar(); }).then(function(n){ toast('PPT gerado: '+n); }).catch(function(e){ toast('Falha ao gerar o PPT: '+e.message); console.error(e); })
    .then(function(){ b.disabled=false; b.lastChild.textContent=t; }); };
  var mq=matchMedia('(prefers-color-scheme: dark)'); if(mq.addEventListener) mq.addEventListener('change',render);
  render();
}
window.PoolApp={ propTerms:function(M){return propTerms(M||compute());}, fm:function(s){return fm(s||SIM);}, recL:recL, oneL:oneL, syncP:syncP, SURF0:SURF0, rerender:function(){render();}, setSim:function(o){ Object.keys(o).forEach(function(k){SIM[k]=o[k];}); if(o.p!=null&&o.mb==null) SIM.mb=o.p/(SIM.mbgb||1024); syncP(SIM); SAVEDSIM=true; ls(KEY_SIM,SIM); render(); }, toast:toast, parseWorkbook:parseWorkbook, compute:function(){return compute();}, simOut:function(M){return simOut(M,SIM);}, presets:presets, bill:bill, fit:fitLinha,
  get SIM(){return SIM;}, get EXC(){return EXC;}, mvLabel:function(x){return mvLabel(x);}, computeAll:function(){return compute([]);}, get DATA(){return DATA;}, get TAB(){return TAB;}, mesLbl:mesLbl, nf:nf, R:R, Rk:Rk, pct:pct, cap1:cap1, _M:null };
if(document.readyState==='loading') document.addEventListener('DOMContentLoaded',init); else init();
})();
