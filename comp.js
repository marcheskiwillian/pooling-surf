/* comp.js — Custo ao MVNO: ISPX × concorrentes (TIP, EAI, B2B…)
   Compara o CUSTO que cada fornecedor cobra do MVNO contra o pacote de dados. */
(function(){
'use strict';
var KEY_COMP='poolsurf.comp.v1', KEY_CB='poolsurf.compbase.v1', KEY_CH='poolsurf.comphide.v1', KEY_CP='poolsurf.comppool.v1';
function ls(k,v){ try{ if(v===undefined){ var r=localStorage.getItem(k); return r?JSON.parse(r):null; } if(v===null) localStorage.removeItem(k); else localStorage.setItem(k,JSON.stringify(v)); }catch(e){ return null; } }
function $(id){ return document.getElementById(id); }
function esc(s){ return String(s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];}); }
function css(n){ return getComputedStyle(document.documentElement).getPropertyValue(n).trim(); }

var ORIG=window.COMP_DATA||null;
var COMP=ls(KEY_COMP)||ORIG;
var BASE=ls(KEY_CB)||'sem';          /* 'sem' = sem portabilidade (base da planilha); 'com' = GB total */
var HIDE=ls(KEY_CH)||{};
var POOL=ls(KEY_CP)===true;
var CMPV=ls('poolsurf.compcmp.v1')||'grade';   /* faixa a faixa: 'grade' (nova grade) ou 'atual' */

/* ---------- leitura da planilha ---------- */
function linhaDe(nome){ return String(nome).replace(/\b\d+\b/g,'').replace(/\s*-\s*/g,' ').replace(/\s+/g,' ').trim(); }
function parse(wb,nome){
  var out=null;
  wb.SheetNames.forEach(function(sn){
    if(out) return;
    var r=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:null});
    for(var i=0;i<Math.min(8,r.length);i++){
      var h=(r[i]||[]).map(function(x){return String(x==null?'':x).trim().toLowerCase();});
      var iO=h.indexOf('operadora'), iC=h.indexOf('custo');
      var iS=h.findIndex(function(x){return /total sem portab/.test(x);}), iT=h.indexOf('gb total');
      if(iO<0||iC<0||(iS<0&&iT<0)) continue;
      var rows=[];
      r.slice(i+1).forEach(function(x){
        if(!x||!x[iO]) return; var n=String(x[iO]).trim(), c=+x[iC];
        var gs=iS>=0?+x[iS]:NaN, gt=iT>=0?+x[iT]:NaN;
        if(!isFinite(gs)||gs<=0) gs=gt; if(!isFinite(gt)||gt<=0) gt=gs;
        if(!isFinite(c)||c<=0||!isFinite(gs)||gs<=0) return;
        var ln=linhaDe(n), op=ln.split(' ')[0].toUpperCase();
        if(op==='ISPX') return;              /* a ISPX vem da tabela do Raio-X */
        rows.push({nome:n,op:op,linha:ln,gbSem:gs,gbTot:gt,custo:c});
      });
      if(rows.length){ out={fonte:nome,aba:sn,importadaEm:new Date().toISOString(),linhas:rows}; break; }
    }
  });
  if(!out) throw new Error('Não achei as colunas Operadora, CUSTO e GB (Total sem portab / GB Total).');
  return out;
}

/* ---------- séries ---------- */
var OPCOR={ISPX:'--s1',EAI:'--s2',TIP:'--s3',B2B:'--s7'};
var EXTRA=['--s4','--s5','--s8','--s6'];
var SHAPE={ISPX:'circle',EAI:'square',TIP:'triangle',B2B:'diamond'};
var DASH=['','7 4','2 3','10 3 2 3'];
function rotulo(l){ return l.replace(/\bVivo\b/g,'VIVO').replace(/^TIP Multi TIM \+ VIVO$/,'TIP Multi (TIM+VIVO)').replace(/^EAI TIM \+ VIVO$/,'EAI (TIM+VIVO)'); }
function series(M,semGrade){
  var A=window.PoolApp, TAB=A.TAB, S=A.SIM, gk=BASE==='sem'?'gbSem':'gbTot';
  var ss=[];
  ss.push({id:'ISPX',op:'ISPX',nome:'ISPX',cor:'--s1',shape:'circle',dash:'',hero:true,
    pts:TAB.map(function(t){ return {nome:'ISPX Plano '+t.k,gb:BASE==='sem'?t.base:t.gb,custo:t.mv,k:t.k}; })});
  if(POOL && M){
    ss.push({id:'ISPXPOOL',op:'ISPX',nome:'ISPX com pool (margem mantida)',cor:'--s1',shape:'circle',dash:'5 4',hollow:true,
      pts:M.planos.filter(function(p){return p.l>0;}).map(function(p){ var gpl=p.g/p.l, pc=window.PoolApp.fm(S)+S.p*gpl, eco=p.custo-pc;
        return {nome:'ISPX Plano '+p.k+' com pool',gb:BASE==='sem'?p.base:p.gb,custo:Math.max(0,p.mv-eco),k:p.k,eco:eco}; })});
  }
  if(window.PoolGrade && M && !semGrade){
    ss.push({id:'ISPXG',op:'ISPX',nome:'ISPX nova grade',cor:'--s5',shape:'circle',dash:'',hero:true,
      pts:window.PoolGrade.calc(M).rows.map(function(r){ return {nome:'ISPX nova grade Plano '+r.k,gb:BASE==='sem'?r.g.gb:r.gbC,custo:r.g.p,k:r.k}; })});
  }
  var lines={}, ord=[];
  (COMP?COMP.linhas:[]).forEach(function(r){ if(!lines[r.linha]){ lines[r.linha]={op:r.op,pts:[]}; ord.push(r.linha); } lines[r.linha].pts.push({nome:r.nome,gb:r[gk],custo:r.custo}); });
  var opCount={}, extra=0, opColor={};
  ord.forEach(function(l){
    var o=lines[l], n=opCount[o.op]||0; opCount[o.op]=n+1;
    if(!opColor[o.op]) opColor[o.op]=OPCOR[o.op]||EXTRA[(extra++)%EXTRA.length];
    ss.push({id:l,op:o.op,nome:rotulo(l),cor:opColor[o.op],shape:SHAPE[o.op]||'circle',dash:DASH[n%DASH.length],
      pts:o.pts.sort(function(a,b){return a.gb-b.gb;})});
  });
  ss.forEach(function(s){ s.pts.forEach(function(p){ p.cpg=p.custo/p.gb; }); s.pts.sort(function(a,b){return a.gb-b.gb;}); });
  return ss;
}
function nearest(s,gb){ return s.pts.reduce(function(best,p){ var d=Math.abs(p.gb-gb), bd=best?Math.abs(best.gb-gb):1e9; return (!best||d<bd||(d===bd&&p.gb<best.gb))?p:best; },null); }
function faixas(ss,id){
  var isp=(id&&ss.filter(function(s){return s.id===id;})[0])||ss[0], rivals=ss.filter(function(s){return s.op!=='ISPX';});
  return isp.pts.map(function(ip){
    return {ispx:ip, cells:rivals.map(function(s){ var c=nearest(s,ip.gb); return {s:s,p:c,delta:c?c.cpg/ip.cpg-1:null}; })};
  });
}

/* ---------- gráfico (SVG com paleta explícita: serve para tela e PPT) ---------- */
function palTela(){ var o={}; ['--ink','--ink-2','--ink-3','--grid','--surface','--s1','--s2','--s3','--s4','--s5','--s6','--s7','--s8'].forEach(function(k){o[k]=css(k);}); return o; }
var PAL_PPT={'--ink':'#0c1418','--ink-2':'#4c5b64','--ink-3':'#7d8c95','--grid':'#e4ebf0','--surface':'#ffffff','--s1':'#2a78d6','--s2':'#eb6834','--s3':'#1baf7a','--s4':'#eda100','--s5':'#e87ba4','--s6':'#008300','--s7':'#4a3aa7','--s8':'#e34948'};
function niceStep(max,n){ var raw=max/n, p=Math.pow(10,Math.floor(Math.log10(raw))), m=[1,2,2.5,5,10].find(function(x){return raw<=x*p;}); return m*p; }
function marker(shape,x,y,r,fill,stroke,sw){
  if(shape==='square') return '<rect x="'+(x-r*.85)+'" y="'+(y-r*.85)+'" width="'+(r*1.7)+'" height="'+(r*1.7)+'" fill="'+fill+'" stroke="'+stroke+'" stroke-width="'+sw+'"/>';
  if(shape==='triangle') return '<path d="M'+x+' '+(y-r*1.1)+' L'+(x+r)+' '+(y+r*.75)+' L'+(x-r)+' '+(y+r*.75)+' Z" fill="'+fill+'" stroke="'+stroke+'" stroke-width="'+sw+'"/>';
  if(shape==='diamond') return '<path d="M'+x+' '+(y-r*1.15)+' L'+(x+r*1.05)+' '+y+' L'+x+' '+(y+r*1.15)+' L'+(x-r*1.05)+' '+y+' Z" fill="'+fill+'" stroke="'+stroke+'" stroke-width="'+sw+'"/>';
  return '<circle cx="'+x+'" cy="'+y+'" r="'+r+'" fill="'+fill+'" stroke="'+stroke+'" stroke-width="'+sw+'"/>';
}
/* kind: 'custo' (R$ × GB) ou 'cpg' (R$/GB × GB) */
function chartSvg(ss,kind,pal,W,H,leg){
  var vis=ss.filter(function(s){return !HIDE[s.id];});
  var all=[]; vis.forEach(function(s){ all=all.concat(s.pts); });
  var fk=kind==='cpg'?'cpg':'custo';
  var xmax=Math.max(10,Math.max.apply(null,all.map(function(p){return p.gb;}).concat([10])));
  var xs=niceStep(xmax,8); xmax=Math.ceil(xmax/xs)*xs;
  var ymax=Math.max.apply(null,all.map(function(p){return p[fk];}).concat([1]));
  var CAP=kind==='cpg'&&ymax>8?8:null; if(CAP) ymax=CAP;
  var ys=niceStep(ymax,5); ymax=Math.ceil(ymax*1.04/ys)*ys;
  var pl=kind==='cpg'?48:56, pr=18, pt=leg?40:14, pb=38;
  var x=function(v){return pl+(W-pl-pr)*v/xmax;}, y=function(v){return pt+(H-pt-pb)*(1-Math.min(v,ymax)/ymax);};
  var tf='font-family:IBM Plex Sans,Segoe UI,Arial,sans-serif;font-size:11px';
  var s='<svg xmlns="http://www.w3.org/2000/svg" width="'+W+'" height="'+H+'" viewBox="0 0 '+W+' '+H+'" role="img" aria-label="'+(kind==='cpg'?'Custo por GB':'Custo ao MVNO')+' por pacote de dados">';
  for(var v=0;v<=ymax+1e-9;v+=ys){ s+='<line x1="'+pl+'" x2="'+(W-pr)+'" y1="'+y(v)+'" y2="'+y(v)+'" stroke="'+pal['--grid']+'"/><text x="'+(pl-6)+'" y="'+(y(v)+4)+'" text-anchor="end" fill="'+pal['--ink-3']+'" style="'+tf+'">'+(kind==='cpg'?'R$ '+v.toLocaleString('pt-BR',{maximumFractionDigits:1}):'R$ '+v.toLocaleString('pt-BR'))+'</text>'; }
  for(var g=0;g<=xmax+1e-9;g+=xs) s+='<text x="'+x(g)+'" y="'+(H-pb+16)+'" text-anchor="middle" fill="'+pal['--ink-3']+'" style="'+tf+'">'+g+'</text>';
  s+='<text x="'+(W-pr)+'" y="'+(H-4)+'" text-anchor="end" fill="'+pal['--ink-3']+'" style="'+tf+'">GB do pacote ('+(BASE==='sem'?'sem portabilidade':'com portabilidade')+')</text>';
  var order=vis.slice().sort(function(a,b){return (a.hero?1:0)-(b.hero?1:0);});
  order.forEach(function(se){
    var c=pal[se.cor]||se.cor, d=se.pts.map(function(p,i){return (i?'L':'M')+x(p.gb).toFixed(1)+' '+y(p[fk]).toFixed(1);}).join(' ');
    s+='<path d="'+d+'" fill="none" stroke="'+c+'" stroke-width="'+(se.hero?3:1.8)+'"'+(se.dash?' stroke-dasharray="'+se.dash+'"':'')+' opacity="'+(se.hero?1:.9)+'"/>';
  });
  order.forEach(function(se){
    var c=pal[se.cor]||se.cor;
    se.pts.forEach(function(p,i){ s+='<g data-s="'+esc(se.id)+'" data-i="'+i+'" style="cursor:default">'+marker(se.shape,x(p.gb),y(p[fk]),se.hero?6:4.6,se.hollow?pal['--surface']:c,se.hollow?c:pal['--surface'],se.hollow?2:1.5)+'</g>'; });
  });
  if(CAP){ var cl=[]; vis.forEach(function(se){ se.pts.forEach(function(p){ if(p[fk]>ymax) cl.push(p); }); });
    if(cl.length){ var mxv=Math.max.apply(null,cl.map(function(p){return p[fk];})), mng=Math.min.apply(null,cl.map(function(p){return p.gb;})), mxg=Math.max.apply(null,cl.map(function(p){return p.gb;}));
      s+='<text x="'+(x(mxg)+12)+'" y="'+(y(ymax)+12)+'" fill="'+pal['--ink-2']+'" style="'+tf+';font-size:10.5px">↑ '+cl.length+' plano'+(cl.length>1?'s':'')+' de '+mng+'–'+mxg+' GB acima de R$ '+ymax+'/GB (até R$ '+mxv.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})+')</text>'; } }
  if(leg){ var lx=pl; vis.forEach(function(se){ var c=pal[se.cor]||se.cor; var tw=se.nome.length*6.2+44; if(lx+tw>W-pr){ return; }
    s+='<line x1="'+lx+'" x2="'+(lx+24)+'" y1="14" y2="14" stroke="'+c+'" stroke-width="'+(se.hero?3:2)+'"'+(se.dash?' stroke-dasharray="'+se.dash+'"':'')+'/>'+marker(se.shape,lx+12,14,4.2,se.hollow?pal['--surface']:c,se.hollow?c:pal['--surface'],1.2)+
      '<text x="'+(lx+30)+'" y="18" fill="'+pal['--ink-2']+'" style="'+tf+'">'+esc(se.nome)+'</text>'; lx+=tw; }); }
  /* rótulo direto só na ISPX */
  vis.filter(function(s){return s.id==='ISPX'||s.id==='ISPXG';}).forEach(function(isp){ isp.pts.forEach(function(p){ var lx=x(p.gb), ly=y(p[fk]);
    s+='<text x="'+lx+'" y="'+(ly-11)+'" text-anchor="middle" fill="'+pal['--ink']+'" style="'+tf+';font-weight:600;paint-order:stroke;stroke:'+pal['--surface']+';stroke-width:3px">'+p.k+'</text>'; }); });
  return s+'</svg>';
}

/* ---------- tela ---------- */
function html(M){
  var A=window.PoolApp, R=A.R, pct=A.pct, nf=A.nf;
  if(!COMP) return '<section class="card" id="comp"><div class="sec-hd"><h2>Custo ao MVNO: ISPX × concorrentes</h2></div><p class="muted">Nenhuma tabela de concorrentes carregada. Use <b>Importar concorrentes</b>.</p><div><button class="btn" id="btnComp">Importar concorrentes</button></div></section>';
  var ss=series(M), rivals=ss.filter(function(s){return s.op!=='ISPX';});
  var temG=ss.some(function(s){return s.id==='ISPXG';}); if(!temG) CMPV='atual';
  var cont=function(id){ var cl=[]; faixas(ss,id).forEach(function(f){ f.cells.forEach(function(c){ if(c.p) cl.push(c.delta); }); }); return {n:cl.length,mais:cl.filter(function(d){return d<0;}).length}; };
  var K0=cont('ISPX'), KG=temG?cont('ISPXG'):null;
  var F=faixas(ss,CMPV==='grade'?'ISPXG':'ISPX'), cells=[]; F.forEach(function(f){ f.cells.forEach(function(c){ if(c.p) cells.push(c.delta); }); });
  var mais=K0.mais;
  var serM=function(id){ var s=ss.filter(function(x){return x.id===id;})[0]; return s?s.pts.reduce(function(a,p){return a+p.custo;},0)/s.pts.reduce(function(a,p){return a+p.gb;},0):0; };
  var ispMed=ss[0].pts.reduce(function(a,p){return a+p.custo;},0)/ss[0].pts.reduce(function(a,p){return a+p.gb;},0);
  var med=rivals.map(function(s){ return {s:s,v:s.pts.reduce(function(a,p){return a+p.custo;},0)/s.pts.reduce(function(a,p){return a+p.gb;},0)}; }).sort(function(a,b){return a.v-b.v;});
  var h='<section class="sec" id="comp"><div class="card">'+
    '<div class="sec-hd" style="flex-direction:row;justify-content:space-between;align-items:flex-start;gap:14px;flex-wrap:wrap"><div style="display:flex;flex-direction:column;gap:4px"><h2>Custo ao MVNO: ISPX × concorrentes</h2>'+
    '<p class="muted small">O que cada fornecedor cobra do MVNO (coluna CUSTO) contra o tamanho do pacote de dados. '+esc(COMP.fonte||'')+' · '+COMP.linhas.length+' planos de concorrentes.</p></div>'+
    '<button class="btn" id="btnComp">Importar concorrentes</button></div>'+
    '<div class="segbar"><div class="seggrp"><span class="gl">Base de GB</span>'+
      '<button class="chip" data-cb="sem" aria-pressed="'+(BASE==='sem')+'">Sem portabilidade</button><button class="chip" data-cb="com" aria-pressed="'+(BASE==='com')+'">Com portabilidade</button></div>'+
      '<label class="toggle" for="cpool"><input type="checkbox" id="cpool"'+(POOL?' checked':'')+'> Mostrar ISPX com pool (cenário do simulador)</label></div>'+
    '<div class="kpis" style="grid-template-columns:repeat(4,1fr)">'+
      '<div class="kpi"><span class="v"><span style="color:var(--crit-ink)">'+K0.mais+'</span>'+(KG?' → <span style="color:'+(KG.mais<K0.mais?'var(--good-ink)':'var(--crit-ink)')+'">'+KG.mais+'</span>':'')+' de '+K0.n+'</span><span class="l">comparações em que o concorrente é mais barato por GB'+(KG?': tabela atual → nova grade':'')+'</span></div>'+
      '<div class="kpi"><span class="v">'+R(ispMed,2)+(KG?' → '+R(serM('ISPXG'),2):'')+'</span><span class="l">R$/GB médio ISPX'+(KG?': tabela atual → nova grade':'')+'</span></div>'+
      '<div class="kpi"><span class="v">'+R(med[0].v,2)+'</span><span class="l">R$/GB médio mais baixo: '+esc(med[0].s.nome)+'</span></div>'+
      '<div class="kpi"><span class="v">'+R(med[med.length-1].v,2)+'</span><span class="l">R$/GB médio mais alto: '+esc(med[med.length-1].s.nome)+'</span></div></div>'+
    '<div class="seggrp" id="cleg"><span class="gl">Séries</span>'+ss.map(function(s){ return '<button class="chip schip" data-sid="'+esc(s.id)+'" aria-pressed="'+(!HIDE[s.id])+'" style="display:inline-flex;align-items:center;gap:6px">'+legSvg(s)+esc(s.nome)+'</button>'; }).join('')+'</div>'+
    '<div class="g2e"><div style="display:flex;flex-direction:column;gap:6px;min-width:0"><h3>Custo ao MVNO por pacote</h3><p class="muted small">Quanto mais à direita e mais baixo, mais dados por menos dinheiro.</p><div class="chartbox" id="c-custo"></div></div>'+
    '<div style="display:flex;flex-direction:column;gap:6px;min-width:0"><h3>Custo por GB</h3><p class="muted small">R$ por GB do pacote. Mostra quem ganha escala nos planos grandes.</p><div class="chartbox" id="c-cpg"></div></div></div>'+
    '</div>';
  /* tabela faixa a faixa */
  h+='<div class="card"><div class="sec-hd"><h2>Faixa a faixa: '+(CMPV==='grade'?'nova grade ISPX':'tabela atual ISPX')+' × plano mais próximo de cada concorrente</h2>'+(temG?'<div class="seggrp" style="margin-top:6px"><span class="gl">Comparar</span><button class="chip" data-cmp="grade" aria-pressed="'+(CMPV==='grade')+'">Nova grade</button><button class="chip" data-cmp="atual" aria-pressed="'+(CMPV==='atual')+'">Tabela atual</button></div>':'')+'<p class="muted small">Comparação pelo R$/GB, já que os pacotes não têm o mesmo tamanho. <span style="color:var(--good-ink);font-weight:600">Verde</span> = concorrente mais caro (ISPX competitiva); <span style="color:var(--crit-ink);font-weight:600">vermelho</span> = concorrente mais barato.</p></div>'+
    '<div class="tbl"><table style="min-width:980px"><thead><tr><th>ISPX</th>'+rivals.map(function(s){return '<th style="text-align:center">'+esc(s.nome)+'</th>';}).join('')+'</tr></thead><tbody>'+
    F.map(function(f){ var ip=f.ispx;
      return '<tr><td>Plano '+ip.k+'<div style="font-family:var(--mono);font-weight:400;font-size:12px;color:var(--ink-2)">'+nf(ip.gb)+' GB · '+R(ip.custo,2)+'<br>'+R(ip.cpg,2)+'/GB</div></td>'+
        f.cells.map(function(c){ if(!c.p) return '<td>—</td>'; var good=c.delta>0.005, bad=c.delta<-0.005;
          return '<td style="text-align:center;background:'+(good?'var(--good-bg)':bad?'var(--crit-bg)':'transparent')+'"><div style="font-size:12px;color:var(--ink-2)">'+nf(c.p.gb)+' GB · '+R(c.p.custo,2)+'</div><div style="font-weight:600;color:'+(good?'var(--good-ink)':bad?'var(--crit-ink)':'var(--ink)')+'">'+R(c.p.cpg,2)+'/GB · '+(c.delta>0?'+':'')+pct(c.delta,0)+'</div></td>'; }).join('')+'</tr>'; }).join('')+
    '</tbody></table></div>'+
    '<p class="small muted">'+leitura(F,rivals)+'</p>'+
    '<p class="small muted">Base “sem portabilidade” é a mesma da coluna Custo/GB da planilha. Nos planos TIP TIM ela já inclui a recorrência. A tabela atual vem do Raio-X (coluna Preço MVNO); a nova grade vem da seção Nova grade (franquia e preço ao MVNO, GB sem depender de portabilidade). “ISPX com pool” mostra o preço ao MVNO se a economia do pool em cada plano (custos fixos por linha + R$/GB × consumo real da faixa) for repassada, mantendo a margem por linha. É uma simulação.</p></div></section>';
  return h;
}
function legSvg(s){ var c='var('+s.cor+')';
  return '<svg width="26" height="12" viewBox="0 0 26 12" aria-hidden="true" style="flex:none"><line x1="1" x2="25" y1="6" y2="6" stroke="'+c+'" stroke-width="'+(s.hero?3:2)+'"'+(s.dash?' stroke-dasharray="'+s.dash+'"':'')+'/>'+
    marker(s.shape,13,6,4,s.hollow?'var(--surface)':c,s.hollow?c:'var(--surface)',1.2).replace(/var\(--surface\)/g,'var(--surface)')+'</svg>'; }
function leitura(F,rivals){
  var A=window.PoolApp, pct=A.pct;
  var porRival=rivals.map(function(s){ var ds=F.map(function(f){ var c=f.cells.filter(function(x){return x.s===s;})[0]; return c&&c.p?c.delta:null; }).filter(function(d){return d!=null;});
    return {s:s, mais:ds.filter(function(d){return d<-0.005;}).length, n:ds.length, med:ds.reduce(function(a,b){return a+b;},0)/(ds.length||1)}; });
  var dom=porRival.filter(function(r){return r.mais===r.n&&r.n;}).map(function(r){return r.s.nome;});
  var perde=porRival.filter(function(r){return r.mais===0&&r.n;}).map(function(r){return r.s.nome;});
  var t=[];
  if(dom.length) t.push('Mais baratos que a ISPX em todas as faixas: <b>'+dom.join(', ')+'</b>.');
  if(perde.length) t.push('A ISPX é mais barata por GB em todas as faixas contra: <b>'+perde.join(', ')+'</b>.');
  var big=F.map(function(f){ var w=f.cells.filter(function(c){return c.p;}).sort(function(a,b){return a.delta-b.delta;})[0]; return {k:f.ispx.k,w:w}; }).filter(function(x){return x.w&&x.w.delta<0;}).sort(function(a,b){return a.w.delta-b.w.delta;})[0];
  if(big) t.push('Maior desvantagem: Plano '+big.k+' contra '+big.w.s.nome+' ('+pct(big.w.delta,0)+' no R$/GB).');
  return t.join(' ');
}
var SS=[];
function draw(M){
  var box1=$('c-custo'), box2=$('c-cpg'); if(!box1) return;
  SS=series(M); var pal=palTela();
  [[box1,'custo'],[box2,'cpg']].forEach(function(b){
    var W=b[0].clientWidth||520; b[0].innerHTML=chartSvg(SS,b[1],pal,W,320);
    b[0].querySelectorAll('g[data-s]').forEach(function(g){
      g.addEventListener('mousemove',function(e){ var se=SS.filter(function(s){return s.id===g.dataset.s;})[0], p=se.pts[+g.dataset.i], A=window.PoolApp;
        var tip=$('tip'); tip.innerHTML='<div>'+esc(p.nome)+'</div><b>'+A.nf(p.gb)+' GB</b> · custo <b>'+A.R(p.custo,2)+'</b><br>'+A.R(p.cpg,2)+' por GB'+(p.eco!=null?'<br>Economia do pool: '+A.R(p.eco,2)+'/linha':'');
        tip.hidden=false; tip.style.left=Math.min(e.clientX+14,innerWidth-270)+'px'; tip.style.top=(e.clientY+14)+'px'; });
      g.addEventListener('mouseleave',function(){ $('tip').hidden=true; });
    });
  });
}
function bind(M){
  var b=$('btnComp'); if(b) b.onclick=function(){ $('compFile').click(); };
  document.querySelectorAll('[data-cb]').forEach(function(x){ x.onclick=function(){ BASE=x.dataset.cb; ls(KEY_CB,BASE); window.PoolApp.rerender(); }; });
  document.querySelectorAll('[data-sid]').forEach(function(x){ x.onclick=function(){ var id=x.dataset.sid; if(HIDE[id]) delete HIDE[id]; else HIDE[id]=1; ls(KEY_CH,HIDE); x.setAttribute('aria-pressed',!HIDE[id]); draw(M); }; });
  document.querySelectorAll('[data-cmp]').forEach(function(x){ x.onclick=function(){ CMPV=x.dataset.cmp; ls('poolsurf.compcmp.v1',CMPV); window.PoolApp.rerender(); }; });
  var cp=$('cpool'); if(cp) cp.onchange=function(){ POOL=cp.checked; ls(KEY_CP,POOL); window.PoolApp.rerender(); };
}
function importar(file){
  var fr=new FileReader();
  fr.onload=function(){ try{ var wb=XLSX.read(new Uint8Array(fr.result),{type:'array'}); COMP=parse(wb,file.name); HIDE={}; ls(KEY_CH,null); ls(KEY_COMP,COMP); window.PoolApp.rerender(); window.PoolApp.toast('Concorrentes importados: '+COMP.linhas.length+' planos.'); }
    catch(e){ window.PoolApp.toast('Não deu para importar: '+e.message); } };
  fr.readAsArrayBuffer(file);
}
function reset(){ COMP=ORIG; BASE='sem'; HIDE={}; POOL=false; [KEY_COMP,KEY_CB,KEY_CH,KEY_CP].forEach(function(k){ls(k,null);}); }

/* PNG para o PPT */
function png(kind,M){
  return new Promise(function(ok){
    var W=1000,H=470, svg=chartSvg(series(M),kind,PAL_PPT,W,H,true);
    var img=new Image(); img.onload=function(){ var sc=2, cv=document.createElement('canvas'); cv.width=W*sc; cv.height=H*sc; var cx=cv.getContext('2d'); cx.fillStyle='#fff'; cx.fillRect(0,0,cv.width,cv.height); cx.drawImage(img,0,0,cv.width,cv.height);
      try{ ok({data:cv.toDataURL('image/png'),ar:W/H}); }catch(e){ ok(null); } };
    img.onerror=function(){ ok(null); }; img.src='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(svg);
  });
}
document.addEventListener('DOMContentLoaded',function(){
  var inp=document.createElement('input'); inp.type='file'; inp.id='compFile'; inp.accept='.xlsx,.xlsm,.xls'; inp.hidden=true; document.body.appendChild(inp);
  inp.onchange=function(){ if(inp.files[0]) importar(inp.files[0]); inp.value=''; };
});
window.PoolComp={get CMPV(){return CMPV;},leitura:leitura,html:html,draw:draw,bind:bind,reset:reset,series:series,faixas:faixas,png:png,parse:parse,
  get COMP(){return COMP;}, get BASE(){return BASE;}, get POOL(){return POOL;}, get HIDE(){return HIDE;}};
})();
