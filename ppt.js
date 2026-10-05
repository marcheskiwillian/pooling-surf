/* ppt.js — exporta o cenário ativo como deck executivo .pptx (gráficos nativos, editáveis) */
(function(){
'use strict';
var C={ink:'0C1418',ink2:'4C5B64',ink3:'7D8C95',line:'DBE4EA',soft:'F6F9FB',page:'FFFFFF',accent:'1A5F9E',
  used:'2A78D6',contract:'C7D2DA',s1:'2A78D6',s2:'EB6834',s3:'1BAF7A',s4:'EDA100',good:'046804',crit:'A32323',dark:'0C1418'};
var HF='Archivo', BF='IBM Plex Sans', MF='IBM Plex Mono';

function exportar(){
  var PC=window.PoolComp, M0=window.PoolApp.compute();
  if(!PC||!PC.COMP) return build(null);
  return Promise.all([PC.png('custo',M0),PC.png('cpg',M0)]).then(build);
}
function build(IMG){
  var PC=window.PoolComp;
  var A=window.PoolApp, M=A.compute(), T=M.T, S=A.SIM, o=A.simOut(M), PR=A.presets(M), fit=A.fit();
  var nf=A.nf, R=A.R, Rk=A.Rk, pct=A.pct, ml=A.mesLbl;
  var EXC=A.EXC, REC=EXC.length?' · sem '+EXC.map(A.mvLabel).join(', '):'';
  var per=ml(M.ms[0])+(M.ms.length>1?' a '+ml(M.ms[M.ms.length-1]):'');
  var pptx=new PptxGenJS(); pptx.layout='LAYOUT_16x9'; pptx.title='Pooling SURF · ISPX'; pptx.company='ISPX';
  var nSl=0;
  function slide(tit,sub){
    var s=pptx.addSlide(); nSl++; s.background={color:C.page};
    s.addText('ISPX · POOLING SURF · '+per.toUpperCase()+REC.toUpperCase(),{x:.5,y:.28,w:7,h:.25,fontFace:MF,fontSize:8,color:C.ink3,charSpacing:2});
    s.addText(String(nSl),{x:9,y:.28,w:.5,h:.25,fontFace:MF,fontSize:8,color:C.ink3,align:'right'});
    s.addText(tit,{x:.5,y:.55,w:9,h:.55,fontFace:HF,fontSize:22,bold:true,color:C.ink});
    if(sub) s.addText(sub,{x:.5,y:1.08,w:9,h:.4,fontFace:BF,fontSize:11,color:C.ink2});
    return s;
  }
  function tiles(s,items,y,h){
    var w=9/items.length;
    items.forEach(function(it,i){
      s.addShape(pptx.ShapeType.rect,{x:.5+i*w,y:y,w:w-.08,h:h||.95,fill:{color:C.soft},line:{color:C.line,width:.75}});
      s.addText(it[0],{x:.6+i*w,y:y+.1,w:w-.2,h:.45,fontFace:MF,fontSize:items.length>5?12:(items.some(function(x){return String(x[0]).length>12;})?13:16),color:it[2]||C.ink});
      s.addText(it[1],{x:.6+i*w,y:y+.52,w:w-.25,h:.35,fontFace:BF,fontSize:9,color:C.ink2});
    });
  }
  function table(s,head,rows,opt){
    var hdr=head.map(function(t,i){return {text:t,options:{bold:true,color:C.ink2,fontSize:8.5,fontFace:BF,align:i?'right':'left',fill:{color:C.soft}}};});
    var body=rows.map(function(r,ri){ return r.map(function(t,i){ return {text:String(t),options:{fontFace:i?MF:BF,fontSize:opt.fs||9,color:C.ink,bold:!i||(opt.boldLast&&ri===rows.length-1),align:i?'right':'left'}}; }); });
    s.addTable([hdr].concat(body),Object.assign({x:.5,y:1.6,w:9,border:{type:'solid',pt:.5,color:C.line},rowH:.28,margin:.04},opt));
  }
  function bullets(s,list,x,y,w,h,fs){
    s.addText(list.map(function(t){return {text:t,options:{bullet:{indent:12},paraSpaceAfter:5}};}),{x:x,y:y,w:w,h:h,fontFace:BF,fontSize:fs||11,color:C.ink,valign:'top'});
  }

  /* 1 capa */
  var s=pptx.addSlide(); nSl++; s.background={color:C.dark};
  s.addText('ISPX · COMPRA DE DADOS SURF',{x:.6,y:1.2,w:8,h:.3,fontFace:MF,fontSize:10,color:'9FB0BA',charSpacing:3});
  s.addText('Pooling SURF',{x:.6,y:1.6,w:8.8,h:.9,fontFace:HF,fontSize:40,bold:true,color:'FFFFFF'});
  s.addText('Consumo real × franquia contratada e proposta de compra em pool',{x:.6,y:2.5,w:8.8,h:.5,fontFace:BF,fontSize:16,color:'D6E0E6'});
  if(EXC.length) s.addText('Recorte: desconsiderando '+EXC.map(A.mvLabel).join(', '),{x:.6,y:3.1,w:8.8,h:.4,fontFace:BF,fontSize:14,color:'F0B54A'});
  s.addText('Período: '+per+' · Fonte: '+(A.DATA.fonte||'—')+' · Gerado em '+new Date().toLocaleDateString('pt-BR'),{x:.6,y:4.6,w:8.8,h:.3,fontFace:MF,fontSize:9,color:'9FB0BA'});

  /* 2 resumo */
  s=slide('Resumo executivo','Médias mensais do período, planos A–F');
  tiles(s,[[nf(T.l),'linhas/mês'],[nf(T.cap),'GB contratados/mês'],[nf(T.g),'GB usados/mês'],[pct(T.cap?T.g/T.cap:0),'da franquia usada',C.crit],[Rk(T.cost),'pago à SURF/mês'],[R(T.g?T.cost/T.g:0,2),'custo por GB usado',C.crit]],1.55);
  s.addShape(pptx.ShapeType.rect,{x:.5,y:2.75,w:9,h:2.4,fill:{color:C.dark},line:{color:C.dark}});
  s.addText('TESE DA NEGOCIAÇÃO',{x:.75,y:2.9,w:8,h:.25,fontFace:MF,fontSize:8,color:'9FB0BA',charSpacing:2});
  s.addText('A ISPX usa '+pct(T.cap?T.g/T.cap:0,0)+' da franquia que compra. O GB contratado sai a '+R(T.cap?T.cost/T.cap:0,2)+', mas o GB usado sai a '+R(T.g?T.cost/T.g:0,2)+'.',{x:.75,y:3.2,w:8.5,h:.9,fontFace:HF,fontSize:18,bold:true,color:'FFFFFF',valign:'top'});
  s.addText('Nossa proposta: custos fixos de '+R(A.fm(S),2)+' por linha/mês + R$ '+(S.mb||0).toLocaleString('pt-BR',{minimumFractionDigits:4,maximumFractionDigits:4})+'/MB ('+R(S.p,2)+'/GB), compromisso de '+S.c+'%. Fatura estimada de '+Rk(o.bill)+'/mês ('+(o.sav<=0?'−':'+')+nf(Math.abs(o.sav*100),1)+'% vs. hoje).',{x:.75,y:4.15,w:8.5,h:.8,fontFace:BF,fontSize:12,color:'D6E0E6',valign:'top'});

  /* proposta SURF x nossa */
  if(window.PoolSurf){
    var SP=window.PoolSurf.calc(M), S0=SP.S0;
    s=slide('Proposta da SURF × nossa proposta','Cobrança por consumo, sem pacote de GB · premissas: '+nf(S.portPct)+'% das ativações com portabilidade, custos únicos diluídos em '+S.meses+' meses');
    var it=[['Mensalidade SIM Card','mensal/linha','sim',2],['FISTEL - TFF','anual/linha','tff',2],['Taxa de Instalação (TFI)','única/linha','tfi',2],['Portabilidade','única/portab.','port',2],['Dados','por MB','mb',4],['Voz','por minuto','voz',3],['SMS','por SMS','sms',2]];
    table(s,['Item','Cobrança','SURF propôs','Nossa proposta'],it.map(function(r){ return [r[0],r[1],'R$ '+(+S0[r[2]]).toLocaleString('pt-BR',{minimumFractionDigits:r[3],maximumFractionDigits:r[3]}),'R$ '+(+S[r[2]]).toLocaleString('pt-BR',{minimumFractionDigits:r[3],maximumFractionDigits:r[3]})]; }),{x:.5,y:1.5,w:4.6,fs:8.5,rowH:.3});
    [[Rk(SP.hoje),'fatura hoje',C.ink],[Rk(SP.surf),'proposta SURF',SP.surf>SP.hoje?C.crit:C.good],[Rk(SP.nossa),'nossa proposta',SP.nossa>SP.hoje?C.crit:C.good]].forEach(function(t,i){ var x=5.3+i*1.42;
      s.addShape(pptx.ShapeType.rect,{x:x,y:1.5,w:1.36,h:.8,fill:{color:C.soft},line:{color:C.line,width:.75}});
      s.addText(t[0],{x:x+.06,y:1.55,w:1.26,h:.38,fontFace:MF,fontSize:12,color:t[2]}); s.addText(t[1],{x:x+.06,y:1.93,w:1.26,h:.3,fontFace:BF,fontSize:8,color:C.ink2}); });
    table(s,['Plano','GB/linha','Hoje','SURF','Nossa'],SP.planos.map(function(q){return ['Plano '+q.k,nf(q.u,2),R(q.hoje,2),R(q.surf,2),R(q.nossa,2)];}),{x:5.3,y:2.45,w:4.2,fs:8,rowH:.26});
    s.addText('Voz e SMS: '+nf(S.min)+' min e '+nf(S.nsms)+' SMS por linha/mês'+(!(+S.min)?' (sem dado de consumo; informar)':'')+'. Empate com a fatura de hoje: R$ '+SP.beSurf.toFixed(4)+'/MB com os fixos da SURF.',{x:.5,y:4.95,w:9,h:.4,fontFace:BF,fontSize:8,color:C.ink2});
  }

  /* impacto do recorte */
  if(EXC.length){
    var MA=A.computeAll(), SI=A.SIM, oA=(function(){ var T0=MA.T, g0=T0.l?T0.g/T0.l:0; var gpl=g0*(1+SI.g/100); var b=A.bill(MA,gpl,SI); return {bill:b,sav:T0.cost?(b-T0.cost)/T0.cost:0,be:SI.p?(T0.cost/T0.l-A.fm(SI))/SI.p:0}; })();
    s=slide('Impacto do recorte: sem '+EXC.map(A.mvLabel).join(', '),'Carteira inteira × recorte, no mesmo cenário do simulador');
    var TA=MA.T, sv=function(x){return (x<=0?'−':'+')+nf(Math.abs(x*100),1)+'%';};
    table(s,['Indicador','Todas as MVNOs','Recorte'],[
      ['Linhas/mês',nf(TA.l),nf(T.l)],['GB usados/mês',nf(TA.g),nf(T.g)],['Franquia usada',pct(TA.cap?TA.g/TA.cap:0),pct(T.cap?T.g/T.cap:0)],
      ['GB por linha',nf(TA.l?TA.g/TA.l:0,2),nf(T.l?T.g/T.l:0,2)],['Pago à SURF/mês',Rk(TA.cost),Rk(T.cost)],['Custo por GB usado',R(TA.g?TA.cost/TA.g:0,2),R(T.g?T.cost/T.g:0,2)],
      ['Fatura pool (cenário)',Rk(oA.bill),Rk(o.bill)],['Economia do pool',sv(oA.sav),sv(o.sav)],['Empate do pool (GB/linha)',nf(oA.be,2),nf(o.be,2)]],{x:1.5,w:7,fs:10.5});
    s.addText('Linhas da MVNO retiradas pela participação dela na base ativa de cada plano (app Benchmark de MVNOs). Consumo retirado direto da planilha da SURF.',{x:.5,y:4.85,w:9,h:.45,fontFace:BF,fontSize:8.5,color:C.ink2});
  }

  /* 3 franquia x consumo */
  s=slide('Franquia contratada × consumo real','GB por mês por plano: o uso fica perto de '+pct(T.cap?T.g/T.cap:0,0)+' em todas as faixas');
  s.addChart(pptx.ChartType.bar,[
    {name:'Contratado',labels:M.planos.map(function(p){return 'Plano '+p.k+' ('+p.gb+' GB)';}),values:M.planos.map(function(p){return Math.round(p.cap);})},
    {name:'Usado',labels:M.planos.map(function(p){return 'Plano '+p.k+' ('+p.gb+' GB)';}),values:M.planos.map(function(p){return Math.round(p.g);})}],
    {x:.5,y:1.5,w:6.2,h:3.8,barDir:'bar',barGrouping:'clustered',chartColors:[C.contract,C.used],showLegend:true,legendPos:'b',legendFontSize:9,
     catAxisLabelFontSize:9,valAxisLabelFontSize:8,valAxisLabelFormatCode:'#,##0',valGridLine:{color:'E4EBF0',size:.5},catAxisOrientation:'maxMin',
     showValue:true,dataLabelFontSize:7,dataLabelFormatCode:'#,##0',dataLabelColor:C.ink2});
  table(s,['Plano','Uso'],M.planos.map(function(p){return ['Plano '+p.k,pct(p.cap?p.g/p.cap:0)];}).concat([['Total',pct(T.cap?T.g/T.cap:0)]]),{x:7,y:1.6,w:2.5,boldLast:true});

  /* 4 raio-x */
  s=slide('Raio-X por plano','Custo SURF pela tabela atual; R$/GB usado mostra quanto custa cada GB consumido de fato');
  table(s,['Plano','Franquia','Custo SURF','Linhas','GB usados','GB/linha','Uso','Pago SURF','R$/GB contr.','R$/GB usado','% custo'],
    M.planos.map(function(p){return ['Plano '+p.k,p.gb+' GB',R(p.custo,2),nf(p.l),nf(p.g),nf(p.l?p.g/p.l:0,2),pct(p.cap?p.g/p.cap:0),R(p.cost),R(p.custo/p.gb,2),R(p.g?p.cost/p.g:0,2),pct(T.cost?p.cost/T.cost:0)];})
    .concat([['Total','','',nf(T.l),nf(T.g),nf(T.l?T.g/T.l:0,2),pct(T.cap?T.g/T.cap:0),R(T.cost),R(T.cap?T.cost/T.cap:0,2),R(T.g?T.cost/T.g:0,2),'100%']]),{boldLast:true,fs:8});
  s.addText('Pela tabela atual, cada plano custa ≈ '+R(fit.a,2)+' fixos + '+R(fit.b,2)+' por GB (reta que melhor explica os 6 preços). Inferência sobre a tabela, não número informado pela SURF.',{x:.5,y:4.55,w:9,h:.6,fontFace:BF,fontSize:9.5,color:C.ink2});

  /* 5 tendência */
  s=slide('Tendência mensal','Consumo por linha estável: argumento para o compromisso mínimo');
  s.addChart(pptx.ChartType.line,[{name:'GB/linha',labels:M.meses.map(function(m){return ml(m.m);}),values:M.meses.map(function(m){return +m.gpl.toFixed(2);})}],
    {x:.5,y:1.5,w:4.6,h:3.6,chartColors:[C.used],lineSize:2,lineDataSymbol:'circle',lineDataSymbolSize:8,showValue:true,dataLabelFormatCode:'0.00',dataLabelPosition:'t',dataLabelFontSize:9,
     valAxisMinVal:0,valAxisLabelFontSize:8,catAxisLabelFontSize:9,valGridLine:{color:'E4EBF0',size:.5},showLegend:false});
  table(s,['Mês','Linhas','GB usados','GB/linha','Pago SURF'],M.meses.map(function(m){return [ml(m.m),nf(m.l),nf(m.g),nf(m.gpl,2),R(m.c)];}),{x:5.3,y:1.6,w:4.2});

  /* 6 simulador */
  s=slide('Cenário simulado de pooling','Custos fixos por linha + dados por MB consumido'+(S.cap?' e teto':''));
  tiles(s,[[R(A.fm(S),2),'fixos por linha/mês'],[R(S.p,2),'R$/GB (R$ '+(S.mb||0).toLocaleString('pt-BR',{minimumFractionDigits:4,maximumFractionDigits:4})+'/MB)'],[S.c+'%','compromisso mínimo'],[(S.g>0?'+':'')+S.g+'%','variação de consumo'],[Rk(o.bill),'fatura pool/mês'],[(o.sav<=0?'−':'+')+nf(Math.abs(o.sav*100),1)+'%','vs. hoje ('+Rk(T.cost)+')',o.sav<=0?C.good:C.crit]],1.5,.9);
  var xs=[],fx=[],pl=[]; var xm=Math.max(10,Math.ceil(Math.max(o.gpl,o.be)*1.4));
  for(var v=0;v<=xm;v+=1){ xs.push(String(v)); fx.push(Math.round(T.cost)); pl.push(Math.round(A.bill(M,v,S))); }
  s.addChart(pptx.ChartType.line,[{name:'Modelo fixo atual',labels:xs,values:fx},{name:'Pool proposto',labels:xs,values:pl}],
    {x:.5,y:2.5,w:5.8,h:2.9,chartColors:['9AA8B3',C.used],lineSize:2,lineDataSymbol:'none',valAxisLabelFormatCode:'#,##0',valAxisLabelFontSize:8,catAxisLabelFontSize:7,
     valGridLine:{color:'E4EBF0',size:.5},showLegend:true,legendPos:'b',legendFontSize:9,catAxisTitle:'GB por linha/mês',showCatAxisTitle:true,catAxisTitleFontSize:8,valAxisMinVal:0});
  bullets(s,['Consumo atual: '+nf(o.gpl0,2)+' GB/linha','Pool empata com o fixo em '+nf(o.be,2)+' GB/linha ('+nf(o.gpl0?(o.be/o.gpl0-1)*100:0,0)+'% acima de hoje)',
    'Pior caso (100% de uso): '+Rk(o.worst)+(S.cap?' com teto; '+Rk(o.worstRaw)+' sem teto':''),
    'Margem ISPX sobre as MVNOs: '+Rk(o.mNow)+' → '+Rk(o.mNew)+'/mês','Compromisso: ~'+nf(o.commitGB/1000,1)+' mil GB/mês'],6.5,2.55,3,2.9,10);

  /* 7 cenários */
  s=slide('Cenários de negociação','Custos fixos da nossa proposta ('+R(A.fm(S),2)+'/linha/mês) e R$/MB '+(PR[0].base==='grade'?'a partir do limite da nova grade':'calculado para cada meta de economia'));
  var PG0=window.PoolGrade?window.PoolGrade.calc(M):{T:{m0:0}};
  var rows=PR.map(function(q){ var ss=Object.assign({},S,{mb:q.mb,p:q.p,c:q.c,g:0,cap:true}); var b=A.bill(M,o.gpl0,ss); var be=(T.cost/T.l-A.fm(ss))/q.p;
    return [q.nome,R(A.fm(S),2),R(q.p,2)+' (R$ '+q.mb.toFixed(5)+'/MB)',q.c+'%',Rk(b),(b<=T.cost?'−':'+')+nf(Math.abs((b-T.cost)/T.cost*100),1)+'%',nf(be,2)+' GB',q.base==='grade'?Rk(q.ganho+PG0.T.m0):Rk(T.rev-b)]; });
  table(s,['Cenário','Fixos/linha','R$/GB pool','Compromisso','Fatura/mês','vs. hoje','Empate (GB/linha)',PR[0].base==='grade'?'Margem ISPX/mês (nova grade)':'Margem ISPX/mês'],rows,{});
  bullets(s,['Abrir a conversa no cenário de Abertura, mirar o Alvo e não passar do Limite.'].concat(PR[0].base==='grade'?['Limite = maior R$/GB que mantém a margem em R$ de todos os planos da nova grade (trava no plano '+PR[2].k+'). Alvo = −15% e Abertura = −30% do limite.']:[]).concat(['Acima do preço-limite, o modelo fixo compensa pela simplicidade e pela ausência de risco de consumo.','Hoje: fatura de '+Rk(T.cost)+'/mês e margem ISPX de '+Rk(o.mNow)+'/mês.']),.5,3.2,9,1.8,10);

  /* projeção 6 meses */
  if(window.PoolProj){
    var PJ=window.PoolProj.calc(M), labs=PJ.hist.map(function(x){return ml(x.m);}).concat(PJ.P.map(function(x){return ml(x.m)+'*';}));
    s=slide('Tendência para os próximos 6 meses','Linhas crescendo '+(PJ.tx>=0?'+':'')+nf(PJ.tx*100,1)+'%/mês · consumo por linha '+(window.PoolProj.ST.gpl>=0?'+':'')+nf(window.PoolProj.ST.gpl,1)+'%/mês · * = projeção');
    s.addChart(pptx.ChartType.line,[{name:'GB consumidos',labels:labs,values:PJ.hist.map(function(x){return Math.round(x.gb);}).concat(PJ.P.map(function(x){return Math.round(x.gb);}))},
      {name:'Faixa alta',labels:labs,values:PJ.hist.map(function(x){return Math.round(x.gb);}).concat(PJ.P.map(function(x){return Math.round(x.hi);}))},
      {name:'Faixa baixa',labels:labs,values:PJ.hist.map(function(x){return Math.round(x.gb);}).concat(PJ.P.map(function(x){return Math.round(x.lo);}))}],
      {x:.5,y:1.45,w:4.4,h:2.3,chartColors:[C.used,'9AA8B3','9AA8B3'],lineSize:2,lineDataSymbolSize:5,valAxisLabelFormatCode:'#,##0',valAxisLabelFontSize:7,catAxisLabelFontSize:7,valGridLine:{color:'E4EBF0',size:.5},showLegend:true,legendPos:'b',legendFontSize:7,showTitle:true,title:'GB consumidos/mês',titleFontSize:9});
    s.addChart(pptx.ChartType.line,[{name:'Fatura fixa',labels:labs,values:PJ.hist.map(function(x){return Math.round(x.fixo);}).concat(PJ.P.map(function(x){return Math.round(x.fixo);}))},
      {name:'Fatura pool',labels:labs,values:PJ.hist.map(function(x){return Math.round(x.pool);}).concat(PJ.P.map(function(x){return Math.round(x.pool);}))}],
      {x:5.1,y:1.45,w:4.4,h:2.3,chartColors:['9AA8B3',C.used],lineSize:2,lineDataSymbolSize:5,valAxisMinVal:0,valAxisLabelFormatCode:'#,##0',valAxisLabelFontSize:7,catAxisLabelFontSize:7,valGridLine:{color:'E4EBF0',size:.5},showLegend:true,legendPos:'b',legendFontSize:7,showTitle:true,title:'Fatura SURF/mês (R$)',titleFontSize:9});
    table(s,['Mês','Linhas','GB/linha','GB consumidos','Faixa de GB','Fatura fixa','Fatura pool'],PJ.P.map(function(x){return [ml(x.m),nf(x.l),nf(x.gpl,2),nf(x.gb),nf(x.lo)+'–'+nf(x.hi),Rk(x.fixo),Rk(x.pool)];}),{y:3.85,fs:8,rowH:.2});
  }

  /* nova grade */
  if(window.PoolGrade){
    var GC=window.PoolGrade.calc(M), GT=GC.T, scN={base:'consumo igual ao de hoje',real:'cenário realista',pior:'pior caso'}[window.PoolGrade.SC];
    var cell=function(t,i,col,b){ return {text:String(t),options:{fontSize:8,fontFace:i?MF:BF,bold:!!b||!i,color:col||C.ink,align:i?'right':'left'}}; };
    var head=function(a){ return a.map(function(t,i){return {text:t,options:{bold:true,fontSize:8,fontFace:BF,color:C.ink2,align:i?'right':'left',fill:{color:C.soft}}};}); };
    s=slide('Nova grade · 1. Oferta ao MVNO','Franquia e preço de venda ao MVNO, comparados à grade atual e ao concorrente mais próximo');
    tiles(s,[[(GT.eco>=0?'−':'+')+Rk(Math.abs(GT.eco)),'no que as MVNOs pagam/mês',C.good],[(GT.gbN/GT.gb0-1>=0?'+':'')+nf((GT.gbN/GT.gb0-1)*100,0)+'%','GB de franquia entregues (média, '+nf(GC.pp*100,0)+'% portadas)']],1.45,.8);
    s.addTable([head(['Plano','Hoje (sem + bônus)','Nova (sem + bônus)','Preço vs hoje','GB vs hoje (sem · com)','R$/GB sem portab.','Confronto mais difícil'])].concat(GC.rows.map(function(r){ var dP=r.p0?r.g.p/r.p0-1:0, dG=r.base0?r.g.gb/r.base0-1:0, dGC=r.gb0?r.gbC/r.gb0-1:0, w=r.comp&&r.comp[0];
      return [cell(r.k,0),cell(r.base0+'+'+r.port0+' GB · '+R(r.p0,2),1),cell(r.g.gb+'+'+nf(r.bo)+' = '+nf(r.gbC)+' GB · '+R(r.g.p,2),2,C.ink,true),cell((dP>0?'+':'')+pct(dP),3,dP<=0?C.good:C.crit),cell((dG>=0?'+':'')+pct(dG,0)+' · '+(dGC>=0?'+':'')+pct(dGC,0),4,dG>=0?C.good:C.crit),
        cell(R(r.cpg0,2)+' → '+R(r.cpg,2),5),cell(w?(w.s.nome+' '+nf(w.p.gb)+' GB '+(w.delta>0?'+':'')+pct(w.delta,0)):'—',6,w?(w.delta>=0?C.good:C.crit):C.ink)]; })),
      {x:.5,y:2.4,w:9,colW:[.5,1.3,1.6,.9,1.1,1.2,2.4],border:{type:'solid',pt:.5,color:C.line},rowH:.28,margin:.04});
    s.addText('Bônus = GB a mais para linha portada. Confronto mais difícil: o concorrente de menor R$/GB entre os planos de tamanho mais próximo. % = quanto o GB dele custa a mais (+) ou a menos (−) que o da ISPX.',{x:.5,y:4.95,w:9,h:.4,fontFace:BF,fontSize:8,color:C.ink2});
    s=slide('Nova grade · 2. Custo real e margem bruta','Custo real = consumo médio × '+R(GC.p,2)+'/GB no pool'+(GC.f?' + '+R(GC.f,2)+'/linha de custos fixos':'')+' · '+scN);
    s.addTable([head(['Plano','Linhas','Consumo (GB)','Custo real','Preço MVNO · R$/GB','Margem/linha','Margem hoje','Margem/mês','Concorrentes (mais próximos)'])].concat(GC.rows.map(function(r){
      var cc=(r.comp||[]).slice(0,2).map(function(c){ return c.s.nome+' '+nf(c.p.gb)+' GB '+R(c.p.custo,2)+' ('+(c.delta>0?'+':'')+pct(c.delta,0)+')'; }).join('\n');
      var w=r.comp&&r.comp[0];
      return [cell(r.k+' · '+r.g.gb+'+'+nf(r.bo)+' GB',0),cell(nf(r.l),1),cell(nf(r.uN,2),2),cell(R(r.custo,2),3,C.ink,true),cell(R(r.g.p,2)+' · '+R(r.cpg,2),4),cell(R(r.mN,2),5,r.mN>=0?(r.mN>=r.m0?C.good:C.crit):C.crit,true),cell(R(r.m0,2),6),cell(Rk(r.l*r.mN),7),cell(cc||'—',8,w?(w.delta>=0?C.good:C.crit):C.ink)]; }))
      .concat([[cell('Carteira',0),cell(nf(GT.l),1),cell(nf(GT.l?GT.uN/GT.l:0,2),2),cell(Rk(GT.cN),3,C.ink,true),cell(Rk(GT.revN),4),cell(Rk(GT.mN),5,GT.mN>=GT.m0?C.good:C.crit,true),cell(Rk(GT.m0),6),cell(Rk(GT.mN),7,C.ink,true),cell('',8)]]),
      {x:.5,y:1.65,w:9,colW:[.85,.6,.65,.75,1.1,.8,.75,.8,2.7],border:{type:'solid',pt:.5,color:C.line},rowH:.36,margin:.04,valign:'middle'});
    s.addText('Consumo médio por linha '+(window.PoolGrade.UB==='total'?'total, com recargas':'só da franquia, sem recargas extras')+'. Margem hoje = preço atual − custo na tabela SURF atual. Concorrentes: os 2 planos de tamanho mais próximo com menor R$/GB; % = quanto o GB deles custa a mais (+) ou a menos (−) que o da ISPX.',{x:.5,y:5.05,w:9,h:.4,fontFace:BF,fontSize:7.5,color:C.ink2});
    var V=GC.pv, PVv=window.PoolGrade.PV;
    s=slide('Valor da portabilidade','Linha portada cancela menos: quanto vale cada linha que o bônus convence a portar');
    tiles(s,[['+'+R(V.d,2),'margem a mais por linha convertida ('+V.hz+' meses)',C.good],[V.boM>0.5?Rk(V.boM):'R$ 0','custo do bônus no pool/mês'],[V.boM>0.5?nf(Math.ceil(V.conv*10)/10,1):'0','conversões/mês para o bônus se pagar']],1.45,.8);
    s.addTable([head(['Origem do número','Churn no período','Churn mensal eq.','Ativas após 12 meses','Meses ativos ('+V.hz+')','Margem ISPX ('+V.hz+' meses)'])].concat([
      [cell('Portada',0),cell(nf(PVv.cp,1)+'%',1),cell(nf(V.mP*100,2)+'%',2),cell(pct(V.r12P,0),3),cell(nf(V.eP,1),4),cell(R(V.vP,2),5,C.good,true)],
      [cell('Número próprio',0),cell(nf(PVv.cn,1)+'%',1),cell(nf(V.mN*100,2)+'%',2),cell(pct(V.r12N,0),3),cell(nf(V.eN,1),4),cell(R(V.vN,2),5,C.ink,true)]]),
      {x:.5,y:2.45,w:9,colW:[1.9,1.3,1.3,1.5,1.4,1.6],border:{type:'solid',pt:.5,color:C.line},rowH:.32,margin:.05});
    bullets(s,['Fonte: '+PVv.fonte+'. Margem de '+R(V.mpl,2)+'/linha/mês da nova grade.',
      'Argumento para a MVNO: linha portada fica mais. EAI, TIP Multi e B2B não dão bônus de portabilidade; TIP TIM e TIP VIVO dão.',
      'Correlação não é causa: quem porta já decidiu ficar. Mesmo valendo metade da diferença, bastam '+(V.boM>0.5?nf(Math.ceil(V.conv*2*10)/10,1):'0')+' conversões/mês.'],.5,3.65,9,1.6,10);
  }

  /* só linhas novas */
  if(window.PoolNovos){
    var NV=window.PoolNovos.calc(M), mk=[3,6,12,24].filter(function(k){return k<=NV.H;});
    s=slide('Cenário só linhas novas','Plano por Consumo só para contratos novos · R$ '+nf(NV.md.p,2)+'/GB · churn '+nf(NV.ch*100,2)+'%/mês · base atual fora da conta');
    tiles(s,NV.out.map(function(o){ return [Rk(o.end.margem)+'/mês',o.s.nome+': '+nf(o.end.l)+' linhas em '+ml(o.end.m)+' (acum. '+Rk(o.end.acM)+')']; }),1.45,.8);
    var rowsN=[]; mk.forEach(function(k){ NV.out.forEach(function(o,j){ var x=o.P[k-1];
      rowsN.push([j===0?ml(x.m)+' (mês '+k+')':'',o.s.nome+' ('+o.s.mv+'×'+nf(o.s.ch)+')',nf(x.l),Rk(x.consumo),Rk(x.fixa),Rk(x.receita),Rk(x.margem),Rk(x.acM)]); }); });
    table(s,['Mês','Cenário (MVNOs/mês × chips)','Linhas','Fatura SURF consumo','Mesmas linhas na fixa','Receita ISPX','Margem/mês','Margem acum.'],rowsN,{y:2.4,fs:8,rowH:.24});
  }

  /* concorrentes */
  if(IMG && PC && PC.COMP){
    var ssC=PC.series(M), FC=PC.faixas(ssC,PC.CMPV==='grade'?'ISPXG':'ISPX'), riv=ssC.filter(function(x){return x.op!=='ISPX';});
    var bas=PC.BASE==='sem'?'GB sem portabilidade':'GB com portabilidade';
    [[0,'Custo ao MVNO: ISPX × concorrentes','Custo cobrado do MVNO por tamanho de pacote ('+bas+')'],[1,'Custo por GB: ISPX × concorrentes','R$ por GB do pacote ('+bas+'): quem ganha escala nos planos grandes']].forEach(function(z){
      var im=IMG[z[0]]; s=slide(z[1],z[2]); if(!im) return;
      var bw=9, bh=bw/im.ar; if(bh>3.3){ bh=3.3; bw=bh*im.ar; }
      s.addImage({data:im.data,x:.5,y:1.45,w:bw,h:bh});
      var txt=PC.leitura(FC,riv).replace(/<[^>]+>/g,'');
      s.addText(z[0]===0?txt:'Cada ponto é um plano. A ISPX é a linha grossa azul com rótulos A–F. Pacotes pequenos têm R$/GB alto em todos os fornecedores. Pontos acima do teto do eixo estão indicados no gráfico.',{x:.5,y:4.8,w:9,h:.55,fontFace:BF,fontSize:9.5,color:C.ink2,valign:'top'});
    });
    s=slide('Faixa a faixa: '+(PC.CMPV==='grade'?'nova grade':'tabela atual')+' ISPX × plano mais próximo','R$/GB do plano concorrente mais próximo e diferença contra a ISPX ('+bas+')');
    var hd=[{text:'ISPX',options:{bold:true,fontSize:8,fontFace:BF,color:C.ink2,fill:{color:C.soft}}}].concat(riv.map(function(x){return {text:x.nome,options:{bold:true,fontSize:7.5,fontFace:BF,color:C.ink2,align:'center',fill:{color:C.soft}}};}));
    var bd=FC.map(function(f){ var ip=f.ispx;
      return [{text:'Plano '+ip.k+'\n'+nf(ip.gb)+' GB · '+R(ip.custo,2)+'\n'+R(ip.cpg,2)+'/GB',options:{fontSize:7.5,fontFace:BF,bold:true,color:C.ink}}].concat(f.cells.map(function(c){
        if(!c.p) return {text:'—',options:{align:'center'}}; var good=c.delta>0.005, bad=c.delta<-0.005;
        return {text:nf(c.p.gb)+' GB · '+R(c.p.custo,2)+'\n'+R(c.p.cpg,2)+'/GB · '+(c.delta>0?'+':'')+pct(c.delta,0),options:{fontSize:7.5,fontFace:MF,align:'center',color:good?C.good:bad?C.crit:C.ink,fill:{color:good?'E6F5E6':bad?'FAE7E7':'FFFFFF'}}}; })); });
    s.addTable([hd].concat(bd),{x:.5,y:1.5,w:9,colW:[1.25].concat(riv.map(function(){return 7.75/riv.length;})),border:{type:'solid',pt:.5,color:C.line},rowH:.5,margin:.03,valign:'middle'});
    s.addText('Verde = concorrente mais caro por GB (ISPX competitiva). Vermelho = concorrente mais barato.',{x:.5,y:5.1,w:9,h:.3,fontFace:BF,fontSize:8.5,color:C.ink2});
  }

  /* 8 tráfego + marcas */
  s=slide('Composição do tráfego e concentração','O que entra no pool e quem move o consumo');
  if(M.tipos.length) s.addChart(pptx.ChartType.bar,M.tipos.map(function(t){return {name:A.cap1(t.nome),labels:M.ms.map(ml),values:t.v};}),
    {x:.5,y:1.5,w:4.4,h:3.7,barDir:'bar',barGrouping:'percentStacked',catAxisOrientation:'maxMin',valAxisLabelFormatCode:'0%',chartColors:[C.s1,C.s2,C.s3,C.s4,'9AA8B3'],showLegend:true,legendPos:'b',legendFontSize:8,catAxisLabelFontSize:9,valAxisLabelFontSize:8,valGridLine:{color:'E4EBF0',size:.5}});
  var top=M.marcas.slice(0,8), rest=M.marcas.slice(8); var ml2=top.map(function(x){return x.nome;}), mv2=top.map(function(x){return +(x.v/M.totalGB*100).toFixed(1);});
  if(rest.length){ ml2.push('Outras '+rest.length); mv2.push(+(rest.reduce(function(a,x){return a+x.v;},0)/M.totalGB*100).toFixed(1)); }
  s.addChart(pptx.ChartType.bar,[{name:'% do tráfego',labels:ml2,values:mv2}],{x:5.1,y:1.5,w:4.4,h:3.7,barDir:'bar',chartColors:[C.used],catAxisOrientation:'maxMin',showValue:true,dataLabelFormatCode:'0.0"%"',dataLabelFontSize:8,
    catAxisLabelFontSize:8.5,valAxisHidden:true,valGridLine:{style:'none'},showLegend:false});

  /* 9 proposta */
  s=slide('Proposta sugerida para levar à SURF','Estrutura de term sheet');
  var PT=A.propTerms(M), tx=function(a){ return a.map(function(x){ return x.replace(/<[^>]+>/g,''); }); };
  var cols=[['1 · ESTRUTURA',tx(PT.estr)],['2 · PREÇO',tx(PT.preco)],['3 · PROTEÇÕES',tx(PT.prot)]];
  cols.forEach(function(c,i){ var x=.5+i*3.03; s.addShape(pptx.ShapeType.rect,{x:x,y:1.55,w:2.9,h:3.6,fill:{color:C.page},line:{color:C.line,width:.75}});
    s.addText(c[0],{x:x+.15,y:1.65,w:2.6,h:.3,fontFace:MF,fontSize:9,color:C.accent,charSpacing:2}); bullets(s,c[1],x+.1,2,2.7,3.1,c[1].join('').length>600?8:c[1].join('').length>350?9:10.5); });

  /* 10 vantagens e riscos */
  s=slide('Vantagens e riscos','O que ganhamos e o que precisamos controlar');
  s.addText('Vantagens',{x:.5,y:1.5,w:4.3,h:.35,fontFace:HF,fontSize:14,bold:true,color:C.good});
  bullets(s,tx(PT.vant),.5,1.9,4.3,3.4,10);
  s.addText('Riscos a controlar',{x:5.2,y:1.5,w:4.3,h:.35,fontFace:HF,fontSize:14,bold:true,color:C.crit});
  bullets(s,tx(PT.risc),5.2,1.9,4.3,3.4,tx(PT.risc).join('').length>650?8.5:10);

  /* 11 notas */
  s=slide('Notas sobre os dados','');
  bullets(s,['Linhas por plano vêm do bloco “QUANTIDADE DE CLIENTES” da aba “Por plano”. A coluna qtd_registros conta registros, não linhas.',
    M.legado.length?'Planos sem faixa na tabela ('+M.legado.map(function(x){return x.nome;}).join(', ')+') ficaram fora do raio-X: '+pct(M.legado.reduce(function(a,x){return a+x.g;},0)/M.totalGB)+' do tráfego.':'Todos os planos da planilha foram associados a uma faixa da tabela.',
    'A franquia contratada usa o total com portabilidade. Se parte das linhas não tem o bônus, o uso real é maior.',
    'A taxa por linha de referência é inferida da tabela SURF e deve ser substituída pelos números da proposta da SURF.',
    'A margem ISPX supõe que todas as linhas pagam o preço MVNO da tabela oficial.'],.5,1.3,9,3.8,12);

  var nome='Pooling SURF - '+per.replace(/\//g,'-')+'.pptx';
  return pptx.writeFile({fileName:nome}).then(function(){ return nome; });
}
window.PoolPPT={exportar:exportar};
})();
