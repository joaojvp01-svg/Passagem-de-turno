import "./style.css";
import "@fontsource/fira-sans/400.css";
import "@fontsource/fira-sans/600.css";
import "@fontsource/fira-sans/700.css";
import "@fontsource/fira-sans-condensed/600.css";
import "@fontsource/fira-sans-condensed/700.css";
import "@fontsource/sora/700.css";
import "@fontsource/sora/800.css";
import html2canvas from "html2canvas";
import * as XLSX from "xlsx";
import LOGO_O from "./assets/logo-laranja.png";
import LOGO_W from "./assets/logo-branco.png";
import PUP from "./assets/pet-filhote.png";
import DUO from "./assets/pet-dupla.png";

const fmt = (n,d=0)=> n==null||isNaN(n) ? "—" : n.toLocaleString("pt-BR",{minimumFractionDigits:d,maximumFractionDigits:d});
const t = kg => fmt(kg/1000,1);
const title = s => s.toLowerCase().replace(/(^|\s)\S/g,m=>m.toUpperCase()).replace(/\b(Da|De|Do|Dos|Das|E)\b/g,m=>m.toLowerCase());
const short = s => { const p = title(s).split(" ").filter(w=>!/^(da|de|do|dos|das|e)$/.test(w)); return p.length>1 ? p[0]+" "+p[p.length-1] : p[0]; };

// ---------- dados ----------
// O cadastro da equipe e os relatórios entram pelo próprio site (abas Equipe e Relatórios)
// e ficam salvos no navegador. Nenhum nome de colaborador fica no código.
let sep = [];   // [nome, pedidos, itens, peso_picking, peso_onda, usuario]
let conf = [];  // [nome, pedidos, itens, peso_total, usuario]
let equipe = []; // [nome, função, turno "1"|"2"|"3"|""]
const motivos = ["Falta injustificada","Abono chefia","Atestado médico","Presença parcial"];

// campos manuais (null = não informado); carros pendentes e total pendente são calculados
let M = {};
let ausencias = [];

// mesma ordem e mesmos nomes dos quadros da planilha antiga (e da página 1 e 2 do painel)
const FORM = [
  ["VD", [["vd_prog","Total programado (kg)"],["vd_carrosProg","Carros programados"],["vd_carrosExp","Carros expedidos"]]],
  ["VET", [["vet_prog","Total programado"],["vet_blocSep","Blocado separado"],["vet_blocPend","Blocado pendente"],["vet_blocConf","Blocados conferidos"],["vet_rotasSep","Rotas separadas"],["vet_rotasPend","Rotas pendentes"],["vet_rotasConf","Rotas conferidas"],["vet_carrosProg","Carros programados"],["vet_carrosExp","Carros expedidos"]]],
  ["Operação", [["op_prog","Carretas programadas"],["op_rec","Carretas recebidas"],["op_conf","Carretas conferidas"],["op_arm","Carretas armazenadas"],["op_recre","REC RE"],["op_missoes","Missões pendentes"],["op_estoque","Estoque em toneladas"],["op_posicoes","Posições vazias"],["av_porto","PLT de avaria Porto"],["av_dourado","PLT de avaria Dourado"],["hidraulicos","Total de hidráulicos"]]]
];

// ---------- helpers de render ----------
const $ = id => document.getElementById(id);
const semAc = x => String(x).normalize("NFD").replace(/[\u0300-\u036f]/g,"").toUpperCase().trim();
function totals(){
  const pick = sep.reduce((a,r)=>a+r[3],0), onda = sep.reduce((a,r)=>a+r[4],0), itens = sep.reduce((a,r)=>a+r[2],0), ped = sep.reduce((a,r)=>a+r[1],0);
  const confKg = conf.reduce((a,r)=>a+r[3],0), confIt = conf.reduce((a,r)=>a+r[2],0);
  return {pick,onda,sepKg:pick+onda,itens,ped,confKg,confIt};
}

function render(){
  const T = totals();
  // alertas: o que conferir antes de enviar (não aparecem nas imagens)
  const al = [];
  if(M.op_conf>M.op_rec) al.push(`<b>Recebimento:</b> ${fmt(M.op_conf)} carretas conferidas para ${fmt(M.op_rec)} recebidas. Se há saldo do turno anterior, registre à parte; se não, corrija o lançamento.`);
  if(T.confKg>T.sepKg*1.15 && T.sepKg) al.push(`<b>Conferência acima da separação:</b> ${t(T.confKg)} t conferidas contra ${t(T.sepKg)} t separadas. Normal se a conferência incluiu volume separado antes.`);
  if(M.vd_prog==null) al.push(`<b>VD sem total programado:</b> sem ele a página 1 não mostra o pendente nem o percentual separado.`);
  if(!sep.length) al.push(`<b>Sem relatório de separação</b> neste turno: suba a planilha na aba Relatórios.`);
  if(!daTurno().length) al.push(`<b>Nenhum colaborador cadastrado no ${turno}º turno:</b> importe a equipe na aba Equipe para o absenteísmo funcionar.`);
  $("alerts").innerHTML = al.length ? `<div class="callout"><div><span class="lbl" style="color:var(--warn)">Antes de enviar, confira</span><ul>${al.map(a=>`<li>${a}</li>`).join("")}</ul></div></div>` : "";
  drawAbs();
  renderSep();
  renderShot();
  salvarDia();
}

// ---------- modo foto ----------
let turno = "1";
const hojeISO = () => { const d=new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`; };
const dataISO = () => $("dataTurno").value || hojeISO();
const dataBR = (o) => { const [y,m,d]=dataISO().split("-").map(Number); return new Date(y,m-1,d).toLocaleDateString("pt-BR",o); };
const MOT = {"Falta injustificada":["var(--bad)","var(--bad-soft)","#c2372f","#f9dedb"],"Atestado médico":["var(--warn)","var(--warn-soft)","#a86b12","#fbf0d9"],"Abono chefia":["var(--info)","var(--info-soft)","#2f6aa8","#dfeaf6"],"Presença parcial":["var(--muted)","var(--surface-2)","#646b74","#eceef0"]};
const funcOf = n => (equipe.find(e=>e[0]===n)||[])[1]||"";
const daTurno = () => equipe.filter(e=>e[2]===turno);
function fillAbsNome(){
  const l=daTurno().map(e=>e[0]).sort((a,b)=>a.localeCompare(b,"pt-BR"));
  $("absNome").innerHTML = l.length ? l.map(n=>`<option>${n}</option>`).join("") : `<option value="">Nenhum colaborador cadastrado neste turno</option>`;
  $("absAdd").disabled = !l.length;
}
const GRUPOS = [["Conferentes",f=>/CONFERENTE/.test(semAc(f))],["Aux. de expedição",f=>/AUXILIAR DE EXPEDICAO/.test(semAc(f))],["Op. de empilhadeira",f=>/EMPILHADEIRA/.test(semAc(f))],["Outras funções",f=>!/CONFERENTE|AUXILIAR DE EXPEDICAO|EMPILHADEIRA/.test(semAc(f))]];
function absStats(){
  const q=daTurno().length, n=ausencias.length, parc=ausencias.filter(a=>a.mot==="Presença parcial").length;
  const cont={}; motivos.forEach(m=>cont[m]=0); ausencias.forEach(a=>cont[a.mot]++);
  const pct = q ? (n-parc)/q*100 : 0;
  return {q,n,parc,pres:q-n+parc,cont,pct};
}

function renderShot(){
  const T = totals(), A = absStats();
  const nv = v => v==null||isNaN(v) ? null : fmt(v);
  const now=new Date(), ger=`Gerado em ${now.toLocaleDateString("pt-BR")} às ${now.toLocaleTimeString("pt-BR",{hour:"2-digit",minute:"2-digit"})}`;
  const dataTxt = dataBR({weekday:"long",day:"2-digit",month:"2-digit",year:"numeric"});
  const head = (img,sub) => `<div class="x-h"><img class="lg" src="${LOGO_W}" alt="PremieRpet"><div class="t">STATUS ${turno}º T</div><div class="d">${sub?sub+" · ":""}${dataTxt} · Estab. 40</div><div class="sw"></div><img class="pets" src="${img}" alt=""></div>`;
  const foot = pn => `<div class="x-f"><b>Passagem de turno · ${turno}º T · Página ${pn} de 3</b><span>${ger}</span></div>`;
  const tag = (c,x) => `<span class="x-tag ${c}">${x}</span>`;
  const T1 = (k,v,sub="",cls="") => { const na = v==null||v===""; return `<div class="x-tile ${cls} ${na?"na":""}"><div class="k">${k}</div><div class="v">${na?"—":v}</div>${sub?`<div class="s">${sub}</div>`:(na?`<div class="s">não informado</div>`:"")}</div>`; };
  const G = (cells,n=4) => { while(cells.length%n) cells.push('<div class="x-tile empty"></div>'); return `<div class="x-g" style="grid-template-columns:repeat(${n},minmax(0,1fr))">${cells.join("")}</div>`; };
  const P = (lbl,done,tot,txt) => { if(!tot) return `<div class="x-prog"><span>${lbl}</span><div class="tr"></div><span class="sub">sem programado</span></div>`; const pc=Math.min(100,done/tot*100); return `<div class="x-prog"><span>${lbl}</span><div class="tr"><i class="${pc>=100?"ok":""}" style="width:${pc}%"></i></div><span><b>${fmt(pc)}%</b> · ${txt}</span></div>`; };
  const pend = (prog,feito,lanc) => { // pendente calculado quando há programado e realizado
    if(prog!=null && feito!=null){ const c=Math.max(0,prog-feito); return {v:c, sub: lanc!=null && lanc!==c ? `lançado: ${fmt(lanc)}` : (c? `${fmt(prog)} − ${fmt(feito)}` : "em dia"), cls: lanc!=null&&lanc!==c ? "warn" : (c?"warn":"ok")}; }
    return {v:lanc, sub: lanc===0?"em dia":"", cls: lanc>0?"warn":(lanc===0?"ok":"")};
  };
  const pctOf = (a,b) => b ? `${fmt(a/b*100)}% do ${"sep."}` : "";

  // ---- página 1: VD e VET (ordem da planilha antiga)
  const vdPend = M.vd_prog!=null ? Math.max(0,M.vd_prog-T.sepKg) : null;
  const vdCar = pend(M.vd_carrosProg, M.vd_carrosExp, M.vd_carrosPend);
  const vetCar = pend(M.vet_carrosProg, M.vet_carrosExp, M.vet_carrosPend);
  const vd = G([
    T1("Total programado", nv(M.vd_prog), M.vd_prog!=null?"kg":""),
    T1("Total separado", T.sepKg?fmt(T.sepKg):"0", "kg", "hl"),
    T1("Total pendente", nv(vdPend), vdPend!=null ? "prog. − sep." : "", vdPend>0?"warn":(vdPend===0?"ok":"")),
    T1("Separação picking", fmt(T.pick), T.sepKg?`${fmt(T.pick/T.sepKg*100)}% do sep.`:""),
    T1("Separação onda", fmt(T.onda), T.sepKg?`${fmt(T.onda/T.sepKg*100)}% do sep.`:""),
    T1("SKU's separados", fmt(T.itens)),
    T1("Total conferido", fmt(T.confKg), "kg"),
    T1("Carros programados", nv(M.vd_carrosProg)),
    T1("Carros expedidos", nv(M.vd_carrosExp), M.vd_carrosProg&&M.vd_carrosExp!=null?`${fmt(M.vd_carrosExp)} de ${fmt(M.vd_carrosProg)}`:"", M.vd_carrosProg&&M.vd_carrosExp>=M.vd_carrosProg?"ok":""),
    T1("Carros pendentes", nv(vdCar.v), vdCar.sub, vdCar.cls)]);
  const vetBlPend = M.vet_blocPend;
  const vet = G([
    T1("Total programado", nv(M.vet_prog)),
    T1("Blocado separado", nv(M.vet_blocSep)),
    T1("Blocado pendente", nv(vetBlPend), vetBlPend===0?"em dia":"", vetBlPend>0?"warn":(vetBlPend===0?"ok":"")),
    T1("Blocados conferidos", nv(M.vet_blocConf), M.vet_blocSep?`${fmt(M.vet_blocConf||0)} de ${fmt(M.vet_blocSep)}`:"", M.vet_blocSep&&M.vet_blocConf>=M.vet_blocSep?"ok":""),
    T1("Rotas separadas", nv(M.vet_rotasSep)),
    T1("Rotas pendentes", nv(M.vet_rotasPend), "", M.vet_rotasPend>0?"warn":""),
    T1("Rotas conferidas", nv(M.vet_rotasConf)),
    T1("Carros programados", nv(M.vet_carrosProg)),
    T1("Carros expedidos", nv(M.vet_carrosExp)),
    T1("Carros pendentes", nv(vetCar.v), vetCar.sub, vetCar.cls)]);
  const vdTag = (M.vd_prog==null&&!T.sepKg) ? tag("na","Sem lançamento") : ((vdPend>0||vdCar.v>0) ? tag("warn","Com pendências") : tag("ok","Em dia"));
  const vetVazio = ["vet_prog","vet_blocSep","vet_blocConf","vet_rotasSep","vet_rotasConf","vet_carrosProg","vet_carrosExp"].every(k=>M[k]==null);
  const vetTag = vetVazio ? tag("na","Sem lançamento") : ((vetCar.v>0||M.vet_blocPend||M.vet_rotasPend) ? tag("warn","Com pendências") : tag("ok","Em dia"));
  $("pg1").innerHTML = head(PUP) + `<div class="x-b c2">
    <div class="x-sec"><div class="x-band"><b>VD<small>VENDA DIRETA</small></b>${vdTag}</div>${P("Separado",T.sepKg,M.vd_prog,`${t(T.sepKg)} de ${t(M.vd_prog||0)} t`)}${vd}</div>
    <div class="x-sec"><div class="x-band"><b>VET</b>${vetTag}</div>${P("Carros expedidos",M.vet_carrosExp||0,M.vet_carrosProg,`${fmt(M.vet_carrosExp||0)} de ${fmt(M.vet_carrosProg)}`)}${vet}</div>
  </div>` + foot(1);

  // ---- página 2: operação, absenteísmo, observações
  const falta=(M.op_prog||0)-(M.op_rec||0);
  const recW = M.op_conf>M.op_rec ? "warn" : (M.op_prog&&M.op_rec>=M.op_prog?"ok":"");
  const op = G([
    T1("Carretas programadas", nv(M.op_prog)),
    T1("Carretas recebidas", nv(M.op_rec), M.op_conf>M.op_rec?`há ${fmt(M.op_conf)} conferidas`:(M.op_prog?`${fmt(M.op_rec||0)} de ${fmt(M.op_prog)}`:""), recW),
    T1("Carretas conferidas", nv(M.op_conf)),
    T1("Carretas armazenadas", nv(M.op_arm)),
    T1("REC RE", nv(M.op_recre)),
    T1("Missões pendentes", nv(M.op_missoes), M.op_missoes>0?"fila p/ próximo turno":(M.op_missoes===0?"fila zerada":""), M.op_missoes>0?"warn":(M.op_missoes===0?"ok":"")),
    T1("Estoque em toneladas", nv(M.op_estoque)),
    T1("Posições vazias", nv(M.op_posicoes)),
    T1("Absenteísmo", fmt(A.n), `${A.pres} de ${A.q} presentes`, "hl"),
    T1("PLT de avaria Porto", nv(M.av_porto), M.av_porto===0?"sem avaria":"", M.av_porto>0?"warn":(M.av_porto===0?"ok":"")),
    T1("PLT de avaria Dourado", nv(M.av_dourado), M.av_dourado===0?"sem avaria":"", M.av_dourado>0?"warn":(M.av_dourado===0?"ok":"")),
    T1("Total de hidráulicos", nv(M.hidraulicos))]);
  const opTag = M.op_prog==null ? tag("na","Sem lançamento") : (falta>0 ? tag("warn","Recebimento pendente") : tag("ok","Recebimento completo"));
  const rows = ausencias.map(a=>{ const c=MOT[a.mot]; return `<div class="x-row"><b>${title(a.nome)}</b><span>${funcOf(a.nome)}</span><i style="background:${c[3]};color:${c[2]}">${a.mot}</i></div>`; }).join("");
  const barra = motivos.map(m=>A.cont[m]?`<i style="width:${A.cont[m]/Math.max(1,A.q)*100}%;background:${MOT[m][2]}"></i>`:"").join("");
  const chips = motivos.filter(m=>A.cont[m]).map(m=>`<span style="background:${MOT[m][3]};color:${MOT[m][2]}"><b>${A.cont[m]}</b> ${m}</span>`).join("");
  const ab = `<div class="x-ab"><div class="hd"><span>Absenteísmo · detalhe</span></div>
    <div class="x-ab-b"><div class="ab-top"><span class="ab-n ${A.n?"":"zero"}">${fmt(A.pct,1)}%</span><div class="ab-m"><b>de absenteísmo</b><span>quadro do turno: ${A.q}</span>${A.parc?`<span>${A.parc} com presença parcial</span>`:""}<div class="ab-bar">${barra}</div><div class="ab-chips">${chips}</div></div></div>
    <div class="ab-lst">${A.n ? rows : `<div class="ab-ok">Equipe completa. Nenhuma ausência registrada.</div>`}</div></div></div>`;
  const o=$("obsTxt").value.trim();
  const obs = `<div class="x-blk"><div class="x-blk-h">Observações do turno</div><div class="x-obs">${o?o.replace(/[&<>]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[c])):'<span class="na">Sem observações.</span>'}</div></div>`;
  $("pg2").innerHTML = head(DUO) + `<div class="x-b"><div class="x-sec"><div class="x-band"><b>OPERAÇÃO</b>${opTag}</div>${P("Carretas armazenadas",M.op_arm||0,M.op_prog,`${fmt(M.op_arm||0)} de ${fmt(M.op_prog)}`)}${op}</div>${ab}${obs}</div>` + foot(2);

  // ---- página 3: produtividade
  const sepR=[...sep].map(r=>({n:r[0],it:r[2],ped:r[1],p:r[3],o:r[4],tot:r[3]+r[4]})).sort((a,b)=>b.tot-a.tot);
  const confR=[...conf].map(r=>({n:r[0],it:r[2],ped:r[1],kg:r[3]})).sort((a,b)=>b.kg-a.kg);
  const confPed=conf.reduce((a,r)=>a+r[1],0);
  const top5 = (arr,val,sub,sep) => { const mx=Math.max(1,...arr.map(val)); return arr.length ? arr.slice(0,5).map((r,i)=>`<div class="x-rk"><span class="pos ${i===0?"first":""}">${i+1}º</span><div class="who"><b>${title(r.n)}</b><small>${sub(r)}</small></div><div class="tr">${sep?`<i style="width:${r.p/mx*100}%;background:#ff6900"></i><i style="width:${r.o/mx*100}%;background:#2b2724"></i>`:`<i style="width:${val(r)/mx*100}%;background:#2b2724"></i>`}</div><span class="vv">${t(val(r))} t</span></div>`).join("") : `<div class="x-obs"><span class="na">Sem relatório para este turno. Suba a planilha na aba Relatórios.</span></div>`; };
  $("pg3").innerHTML = head(PUP,"Produtividade") + `<div class="x-b c2">
    <div class="x-sec"><div class="x-band"><b>SEPARAÇÃO</b>${tag(sep.length?"ok":"na",sep.length?`${sep.length} pessoas`:"Sem relatório")}</div>
      ${G([T1("Peso separado",T.sepKg?fmt(T.sepKg):"0","kg","hl"),T1("Picking",fmt(T.pick),T.sepKg?`${fmt(T.pick/T.sepKg*100)}% do total`:""),T1("Onda",fmt(T.onda),T.sepKg?`${fmt(T.onda/T.sepKg*100)}% do total`:""),
           T1("Pedidos",fmt(T.ped)),T1("Itens",fmt(T.itens)),T1("Média por pessoa",sep.length?fmt(T.sepKg/sep.length):null,"kg")],3)}
      <div class="x-blk"><div class="x-blk-h">5 destaques da separação <span class="lg2"><i style="background:#ff6900"></i>picking <i style="background:#2b2724"></i>onda</span></div>${top5(sepR,r=>r.tot,r=>`${fmt(r.it)} itens · ${fmt(r.ped)} pedidos`,true)}</div></div>
    <div class="x-sec"><div class="x-band"><b>CONFERÊNCIA</b>${tag(conf.length?"ok":"na",conf.length?`${conf.length} conferentes`:"Sem relatório")}</div>
      ${G([T1("Peso conferido",T.confKg?fmt(T.confKg):"0","kg","hl"),T1("Pedidos",fmt(confPed)),T1("Itens",fmt(T.confIt)),
           T1("Média por conferente",conf.length?fmt(T.confKg/conf.length):null,"kg"),T1("Itens por pedido",confPed?fmt(T.confIt/confPed,1):null)],3)}
      <div class="x-blk"><div class="x-blk-h">5 destaques da conferência</div>${top5(confR,r=>r.kg,r=>`${fmt(r.it)} itens · ${fmt(r.ped)} pedidos`,false)}</div></div>
  </div>` + foot(3);
  fitShot();
}
function fitShot(){
  [1,2,3].forEach(n=>{ const st=$("st"+n), w=st.clientWidth; if(!w) return;
    const sc=Math.min(1.25,w/800); $("fit"+n).style.transform=`scale(${sc})`; st.style.height=($("pg"+n).offsetHeight*sc+2)+"px"; });
}
window.addEventListener("resize",fitShot);

const MIME={png:"image/png",xlsx:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",csv:"text/csv"};
const ehCelular = () => matchMedia("(pointer: coarse)").matches;
function baixarArquivo(nome,blob){
  const url=URL.createObjectURL(blob); const a=document.createElement("a");
  a.href=url; a.download=nome; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(()=>URL.revokeObjectURL(url),10000);
}
// devolve "compartilhado", "baixado" ou "cancelado"
async function entregar(arquivos, titulo){
  const files=arquivos.map(f=>new File([f.blob],f.nome,{type:MIME[f.nome.split(".").pop()]||f.blob.type}));
  if(ehCelular() && navigator.canShare && navigator.canShare({files})){
    try{ await navigator.share({files,title:titulo}); return "compartilhado"; }
    catch(e){ if(e && e.name==="AbortError") return "cancelado"; }
  }
  files.forEach((f,i)=>setTimeout(()=>baixarArquivo(f.name,f),i*400));
  return "baixado";
}
async function shotBlob(id){
  const canvas = await html2canvas($(id),{scale:2,backgroundColor:"#ffffff",useCORS:true,
    onclone:d=>{ d.querySelectorAll(".pg-fit").forEach(f=>f.style.transform="none"); d.querySelectorAll(".pg-stage").forEach(f=>f.style.height="auto"); }});
  return await new Promise(r=>canvas.toBlob(r,"image/png"));
}
async function saveShot(pages){
  const msg=$("saveMsg"); msg.className="save-msg"; msg.textContent="Gerando imagens…";
  const arquivos=[];
  try{ for(const n of pages) arquivos.push({nome:`status-${turno}T-${dataISO()}-p${n}.png`, blob: await shotBlob("pg"+n)}); }
  catch(e){ msg.textContent="Não consegui gerar as imagens. Tente de novo ou tire um print das páginas."; return; }
  const r = await entregar(arquivos, `Status ${turno}º turno · ${dataBR({day:"2-digit",month:"2-digit"})}`);
  if(r==="cancelado"){ msg.textContent="Envio cancelado."; return; }
  msg.className="save-msg ok";
  msg.textContent = r==="compartilhado" ? "Imagens enviadas para o compartilhamento." : (arquivos.length>1 ? `${arquivos.length} imagens baixadas.` : "Imagem baixada.");
}

// ---------- turnos e dados salvos por data ----------
const blankM = () => Object.fromEntries(FORM.flatMap(([,f])=>f.map(([k])=>[k,null])));
const blankDia = () => ({"1":{M:blankM(),sep:[],conf:[],aus:[],obs:"",fonte:""},"2":{M:blankM(),sep:[],conf:[],aus:[],obs:"",fonte:""},"3":{M:blankM(),sep:[],conf:[],aus:[],obs:"",fonte:""}});
const chaveDia = d => "passagemTurno.dia." + d;
let S = blankDia();
function carregarDia(d){
  let v=null; try{ v=JSON.parse(localStorage.getItem(chaveDia(d))||"null"); }catch(e){}
  const b=blankDia();
  if(v) ["1","2","3"].forEach(k=>{ if(v[k]) b[k]={...b[k],...v[k],M:{...b[k].M,...(v[k].M||{})}}; });
  return b;
}
function salvarDia(){ try{ localStorage.setItem(chaveDia(dataISO()), JSON.stringify(S)); }catch(e){} }
function vincular(){ // aponta as variáveis de trabalho para o turno atual
  const x=S[turno]; M=x.M; sep=x.sep; conf=x.conf; ausencias=x.aus; sepFonte=x.fonte||""; $("obsTxt").value=x.obs||"";
  document.querySelectorAll("#formGrid input[data-k]").forEach(i=>{ i.value = M[i.dataset.k]??""; });
  sepUndo=null; $("sepStatus").innerHTML=""; $("absMsg").textContent="";
  $("saveMsg").className="save-msg"; $("saveMsg").textContent="No celular, abre o compartilhamento (WhatsApp). No computador, baixa as imagens.";
}
function setTurno(tn){
  turno=tn;
  document.querySelectorAll("#turnoSeg button").forEach(b=>b.setAttribute("aria-pressed",b.dataset.t===tn));
  $("turnoLbl").textContent=`${tn}º turno`;
  vincular(); fillAbsNome(); render();
}
function setData(){ S=carregarDia(dataISO()); vincular(); fillAbsNome(); render(); }

// ---------- formulário ----------
function buildForm(){
  $("formGrid").innerHTML = FORM.map(([g,fields])=>`<fieldset><legend>${g}</legend>${fields.map(([k,l])=>`<div class="f"><label for="in_${k}">${l}</label><input type="number" min="0" step="any" id="in_${k}" data-k="${k}" value="${M[k]??""}" placeholder="—"></div>`).join("")}</fieldset>`).join("");
  $("formGrid").addEventListener("input",e=>{
    const k=e.target.dataset.k; if(!k) return;
    const v=e.target.value.trim(); M[k]= v==="" ? null : Number(v); render();
  });
  fillAbsNome();
  $("absMotSel").innerHTML = motivos.map(m=>`<option>${m}</option>`).join("");
  $("absAdd").addEventListener("click",()=>{
    const nome=$("absNome").value; if(!nome) return;
    if(ausencias.some(a=>a.nome===nome)){ $("absMsg").textContent=`${short(nome)} já está na lista deste turno.`; return; }
    ausencias.push({nome,mot:$("absMotSel").value}); $("absMsg").textContent=`${short(nome)} registrado(a): ${$("absMotSel").value.toLowerCase()}.`; render();
  });
}
function drawAbs(){
  const A=absStats();
  $("absN").textContent=A.n; $("absN").className="n num"+(A.n?"":" zero");
  $("absMeta").innerHTML=`<b>${A.n===1?"ausência":"ausências"}</b><span>${A.pres} de ${A.q} presentes · ${fmt(A.pct,1)}% de absenteísmo</span>`;
  $("absBar").innerHTML = motivos.map(m=>A.cont[m]?`<i style="width:${A.cont[m]/Math.max(1,A.q)*100}%;background:${MOT[m][0]}"></i>`:"").join("");
  $("absMot").innerHTML = motivos.map(m=>`<div class="mot" style="--sbg:${MOT[m][1]};--sfg:${MOT[m][0]};${A.cont[m]?"":"opacity:.45"}"><span class="n">${A.cont[m]}</span><span class="l">${m}</span></div>`).join("");
  $("absFunc").innerHTML = `<span class="lbl" style="margin-top:4px">Presentes por função</span>` + GRUPOS.map(([g,fn])=>{
    const tot=daTurno().filter(e=>fn(e[1])).length; if(!tot) return "";
    const fora=ausencias.filter(a=>a.mot!=="Presença parcial" && fn(funcOf(a.nome))).length; const pr=tot-fora;
    return `<div class="func"><span>${g}</span><div class="tr"><i style="width:${pr/tot*100}%;${fora?'background:var(--warn)':''}"></i></div><span class="num">${pr}/${tot}</span></div>`; }).join("");
  $("absTable").innerHTML = A.n
    ? `<div class="abs-tbl-wrap"><table class="abs-tbl"><thead><tr><th>Colaborador</th><th>Função</th><th>Motivo</th><th></th></tr></thead><tbody>${ausencias.map((a,i)=>`<tr><td><b>${title(a.nome)}</b></td><td>${funcOf(a.nome)}</td><td><span class="mbadge" style="--sbg:${MOT[a.mot][1]};--sfg:${MOT[a.mot][0]}">${a.mot}</span></td><td style="text-align:right"><button class="linkbtn" type="button" data-i="${i}">Remover</button></td></tr>`).join("")}</tbody></table></div>`
    : `<div class="abs-empty"><span><b>Nenhuma ausência registrada neste turno.</b> Use o formulário para lançar faltas, atestados, abonos e presença parcial.</span></div>`;
}
$("absTable").addEventListener("click",e=>{
  const i=e.target.dataset.i; if(i==null) return; ausencias.splice(+i,1); render();
});

// ---------- relatório de separação ----------
let sepFonte="", sepUndo=null, sepView="sep", sepSort={k:"tot",d:-1};
function funcWMS(nome){
  const p=semAc(nome).split(/\s+/); const f=p[0], l=p[p.length-1];
  let hit=equipe.find(e=>{ const q=semAc(e[0]).split(/\s+/); return q[0]===f && q[q.length-1]===l; });
  if(!hit){ // nomes abreviados no WMS: maior número de palavras em comum (mínimo 2)
    const tk=new Set(p.filter(w=>w.length>2)); let best=0;
    equipe.forEach(e=>{ const sc=semAc(e[0]).split(/\s+/).filter(w=>w.length>2&&tk.has(w)).length; if(sc>best){best=sc;hit=e;} });
    if(best<2) hit=null;
  }
  return hit ? hit[1] : "";
}
const toNum = v => { if(typeof v==="number") return v; let x=String(v??"").trim(); if(!x) return 0; if(/,\d+$/.test(x)) x=x.replace(/\./g,"").replace(",","."); return parseFloat(x.replace(/[^\d.\-]/g,""))||0; };
function lerRelatorio(rows){
  // devolve {tipo:"sep"|"conf", linhas} ou null
  const H = r => r.map(c=>semAc(c).replace(/\s+/g,"_"));
  for(let i=0;i<Math.min(rows.length,15);i++){
    const h=H(rows[i]); const ix=n=>h.indexOf(n);
    const iN=ix("NOME"); if(iN<0) continue;
    const iU=ix("USUARIO"), iP=ix("NO_PEDIDOS")>=0?ix("NO_PEDIDOS"):ix("PEDIDOS"), iI=ix("NO_ITENS")>=0?ix("NO_ITENS"):ix("ITENS");
    const body=rows.slice(i+1).filter(r=>String(r[iN]??"").trim() && !/^total/i.test(String(r[iN])));
    if(ix("PESO_PICKING")>=0 || ix("PESO_ONDA")>=0){
      const a=ix("PESO_PICKING"), b=ix("PESO_ONDA");
      return {tipo:"sep", linhas: body.map(r=>[String(r[iN]).trim(), toNum(r[iP]), toNum(r[iI]), a>=0?toNum(r[a]):0, b>=0?toNum(r[b]):0, iU>=0?String(r[iU]).trim():""])};
    }
    if(ix("PESO_TOTAL")>=0 || ix("PESO")>=0){
      const a=ix("PESO_TOTAL")>=0?ix("PESO_TOTAL"):ix("PESO");
      return {tipo:"conf", linhas: body.map(r=>[String(r[iN]).trim(), toNum(r[iP]), toNum(r[iI]), toNum(r[a]), iU>=0?String(r[iU]).trim():""])};
    }
  }
  return null;
}
async function lerSep(file){
  if(!file) return;
  const st=$("sepStatus"); st.innerHTML=`<div class="prev">Lendo ${esc(file.name)}…</div>`;
  try{
    const achados=[];
    if(/\.csv$/i.test(file.name)){
      const txt=await file.text(); const l0=txt.split("\n")[0]; const sc=(l0.match(/;/g)||[]).length>(l0.match(/,/g)||[]).length?";":(l0.includes("\t")?"\t":",");
      const r=lerRelatorio(txt.split(/\r?\n/).filter(x=>x.trim()).map(x=>x.split(sc).map(c=>c.replace(/^"|"$/g,"").trim()))); if(r) achados.push(r);
    } else {
            const wb=XLSX.read(await file.arrayBuffer(),{type:"array"});
      for(const sn of wb.SheetNames){ const r=lerRelatorio(XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:""})); if(r && !achados.some(a=>a.tipo===r.tipo)) achados.push(r); }
    }
    if(!achados.length){ st.innerHTML=`<div class="prev err">Não encontrei um relatório de separação ou conferência nesse arquivo. A linha de cabeçalho precisa ter NOME e PESO_PICKING / PESO_ONDA (separação) ou PESO_TOTAL (conferência).</div>`; return; }
    sepUndo={sep:sep.map(r=>[...r]),conf:conf.map(r=>[...r]),fonte:sepFonte};
    const partes=[];
    achados.forEach(a=>{ if(a.tipo==="sep"){ sep=a.linhas; partes.push(`separação com ${a.linhas.length} colaboradores`); } else { conf=a.linhas; partes.push(`conferência com ${a.linhas.length} conferentes`); } });
    sepFonte=file.name; S[turno].sep=sep; S[turno].conf=conf; S[turno].fonte=sepFonte;
    if(!achados.some(a=>a.tipo===sepView)) sepView=achados[0].tipo;
    st.innerHTML=`<div class="prev"><span class="ok-msg">Importado: ${partes.join(" e ")}. Painel do ${turno}º turno atualizado.</span><span><button class="linkbtn" type="button" id="sepDesf">Desfazer importação</button></span></div>`;
    render();
  }catch(e){ st.innerHTML=`<div class="prev err">Não consegui ler esse arquivo. Confira se é .xlsx ou .csv e se não está protegido por senha.</div>`; }
  $("sepFile").value="";
}
$("sepStatus").addEventListener("click",e=>{ if(e.target.id!=="sepDesf"||!sepUndo) return; sep=sepUndo.sep; conf=sepUndo.conf; sepFonte=sepUndo.fonte; S[turno].sep=sep; S[turno].conf=conf; S[turno].fonte=sepFonte; sepUndo=null; $("sepStatus").innerHTML=`<div class="prev"><span class="sub">Importação desfeita.</span></div>`; render(); });
$("sepFile").addEventListener("change",e=>lerSep(e.target.files[0]));
(()=>{ const dz=$("sepDrop"); ["dragenter","dragover"].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add("over");})); ["dragleave","drop"].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove("over");})); dz.addEventListener("drop",e=>lerSep(e.dataTransfer.files[0])); })();
$("sepView").addEventListener("click",e=>{ const b=e.target.closest("[data-v]"); if(!b) return; sepView=b.dataset.v; sepSort={k:sepView==="sep"?"tot":"kg",d:-1}; renderSep(); });
$("sepBusca").addEventListener("input",renderSep);
$("sepZerar").addEventListener("click",()=>{
  const zb=$("sepZerar"), nome = sepView==="sep" ? "separação" : "conferência";
  if(!zb.classList.contains("sure")){ zb.classList.add("sure"); zb.textContent=`Confirmar: zerar ${nome} do ${turno}º turno`; return; }
  sepUndo={sep:sep.map(r=>[...r]),conf:conf.map(r=>[...r]),fonte:sepFonte};
  if(sepView==="sep") sep=[]; else conf=[];
  if(!sep.length && !conf.length) sepFonte="";
  S[turno].sep=sep; S[turno].conf=conf; S[turno].fonte=sepFonte;
  $("sepStatus").innerHTML=`<div class="prev"><span class="ok-msg">${nome[0].toUpperCase()+nome.slice(1)} do ${turno}º turno zerada. Painel atualizado.</span><span><button class="linkbtn" type="button" id="sepDesf">Desfazer</button></span></div>`;
  render();
});
document.addEventListener("click",e=>{ const zb=$("sepZerar"); if(zb.classList.contains("sure") && e.target!==zb){ zb.classList.remove("sure"); zb.textContent = sepView==="sep" ? "Zerar separação" : "Zerar conferência"; } });
$("sepTbl").addEventListener("click",e=>{ const b=e.target.closest("button[data-k]"); if(!b) return; const k=b.dataset.k; sepSort = sepSort.k===k ? {k,d:-sepSort.d} : {k,d:(k==="nome"||k==="func")?1:-1}; renderSep(); });

function renderSep(){
  $("sepTitle").textContent=`Relatórios do ${turno}º turno`;
  document.querySelectorAll("#sepView button").forEach(b=>b.setAttribute("aria-pressed",b.dataset.v===sepView));
  $("sepFonteTxt").textContent = (sep.length||conf.length) ? `Fonte: ${sepFonte||"lançamento manual"}` : "";
  const isSep = sepView==="sep";
  const zb=$("sepZerar"); zb.classList.remove("sure"); zb.textContent = isSep ? "Zerar separação" : "Zerar conferência"; zb.disabled = !(isSep?sep.length:conf.length);
  const base = isSep ? sep.map(r=>({nome:r[0],usr:r[5]||"",ped:r[1],it:r[2],pick:r[3],onda:r[4],tot:r[3]+r[4]}))
                     : conf.map(r=>({nome:r[0],usr:r[4]||"",ped:r[1],it:r[2],kg:r[3]}));
  base.forEach(r=>{ r.func=funcWMS(r.nome); r.ipp = r.ped? r.it/r.ped : 0; });
  const totKg = base.reduce((a,r)=>a+(isSep?r.tot:r.kg),0), totPed=base.reduce((a,r)=>a+r.ped,0), totIt=base.reduce((a,r)=>a+r.it,0);
  const pick=isSep?base.reduce((a,r)=>a+r.pick,0):0, onda=isSep?base.reduce((a,r)=>a+r.onda,0):0;
  const pc = x => totKg? fmt(x/totKg*100,0)+"%" : "—";
  $("sepKpi").innerHTML = (isSep ? [
    ["Peso separado",t(totKg),"t"],["Picking",t(pick),`t · ${pc(pick)}`],["Onda",t(onda),`t · ${pc(onda)}`],
    ["Pedidos",fmt(totPed),""],["Itens",fmt(totIt),""],["Média por pessoa",base.length?t(totKg/base.length):"—",`t · ${base.length} pessoas`]
  ] : [
    ["Peso conferido",t(totKg),"t"],["Pedidos",fmt(totPed),""],["Itens",fmt(totIt),""],
    ["Conferentes",fmt(base.length),""],["Média por conferente",base.length?t(totKg/base.length):"—","t"],["Itens por pedido",totPed?fmt(totIt/totPed,1):"—",""]
  ]).map(([l,v,u])=>`<div><span class="lbl">${l}</span><span class="v">${v}${u?`<small>${u}</small>`:""}</span></div>`).join("");

  const q=semAc($("sepBusca").value||"");
  let rows=base.filter(r=>!q || semAc(r.nome+" "+r.usr+" "+r.func).includes(q));
  const k=sepSort.k; rows.sort((a,b)=> (typeof a[k]==="string" ? a[k].localeCompare(b[k],"pt-BR") : a[k]-b[k]) * sepSort.d);
  const cols = isSep
    ? [["nome","Colaborador","l"],["func","Função","l"],["ped","Pedidos"],["it","Itens"],["pick","Picking (kg)"],["onda","Onda (kg)"],["tot","Total (kg)"],["share","% do total"]]
    : [["nome","Conferente","l"],["func","Função","l"],["ped","Pedidos"],["it","Itens"],["ipp","Itens/pedido"],["kg","Peso (kg)"],["share","% do total"]];
  const th = c => c[0]==="share" ? `<th>${c[1]}</th>` : `<th class="${c[2]||""}" ${sepSort.k===c[0]?`aria-sort="${sepSort.d>0?"ascending":"descending"}"`:""}><button type="button" data-k="${c[0]}">${c[1]}${sepSort.k===c[0]?(sepSort.d>0?" ▲":" ▼"):""}</button></th>`;
  const kgf = v => v ? `<span class="n">${fmt(v)}</span>` : `<span class="n zero">0</span>`;
  const maxV = Math.max(1,...base.map(r=>isSep?r.tot:r.kg));
  const bar = r => { const w = totKg ? (isSep?r.tot:r.kg)/totKg*100 : 0; return `<div class="share"><span>${fmt(w,1)}%</span><span class="tr">${isSep?`<i style="width:${r.pick/maxV*100}%;background:var(--brand)"></i><i style="width:${r.onda/maxV*100}%;background:var(--bar2)"></i>`:`<i style="width:${r.kg/maxV*100}%;background:var(--bar2)"></i>`}</span></div>`; };
  const tr = (r,i) => `<tr><td class="l nm"><b>${title(r.nome)}</b><small>${r.usr||""}</small></td><td class="l">${r.func||'<span class="fn-miss">não está no cadastro</span>'}</td><td><span class="n">${fmt(r.ped)}</span></td><td><span class="n">${fmt(r.it)}</span></td>` +
     (isSep ? `<td>${kgf(r.pick)}</td><td>${kgf(r.onda)}</td><td><span class="n"><b>${fmt(r.tot)}</b></span></td>` : `<td><span class="n">${fmt(r.ipp,1)}</span></td><td><span class="n"><b>${fmt(r.kg)}</b></span></td>`) + `<td>${bar(r)}</td></tr>`;
  $("sepTbl").innerHTML = base.length
    ? `<thead><tr>${cols.map(th).join("")}</tr></thead><tbody>${rows.map(tr).join("") || `<tr><td class="l" colspan="${cols.length}">Nenhum resultado para a busca.</td></tr>`}</tbody>`
    : `<tbody><tr><td class="l" style="padding:22px 14px">Nenhum relatório de ${isSep?"separação":"conferência"} no ${turno}º turno. Suba a planilha no quadro acima.</td></tr></tbody>`;
}

// ---------- apoio (pop-up) ----------
let apoioVolta=null;
function abrirApoio(){ apoioVolta=document.activeElement; $("apoio").hidden=false; $("apoioX").focus(); }
function fecharApoio(){ $("apoio").hidden=true; if(apoioVolta&&apoioVolta.focus) apoioVolta.focus(); }
$("btnApoio").addEventListener("click",abrirApoio);
$("apoioX").addEventListener("click",fecharApoio);
$("apoio").addEventListener("click",e=>{ if(e.target.id==="apoio") fecharApoio(); });
document.addEventListener("keydown",e=>{ if(e.key==="Escape" && !$("apoio").hidden) fecharApoio(); });

// ---------- equipe / cadastro ----------
const LSK="passagemTurno.equipe";
function saveEq(){ try{ localStorage.setItem(LSK, JSON.stringify(equipe)); }catch(e){} }
(function(){ try{ const v=JSON.parse(localStorage.getItem(LSK)||"null"); if(Array.isArray(v)&&(!v.length||Array.isArray(v[0]))) equipe=v; }catch(e){} })();
let eqFiltro="todos", pendImport=null;
const TLBL={"1":"1º turno","2":"2º turno","3":"3º turno","":"Sem turno"};
const normTurno = v => { const m=String(v??"").match(/[123]/); if(m) return m[0]; const t=String(v??"").toLowerCase(); if(/prim/.test(t)) return "1"; if(/seg/.test(t)) return "2"; if(/terc/.test(t)) return "3"; return ""; };
const esc = x => String(x).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
function eqChanged(){ saveEq(); fillAbsNome(); renderEquipe(); render(); }

function renderEquipe(){
  const cnt={"1":0,"2":0,"3":0,"":0}; equipe.forEach(e=>cnt[e[2]||""]++);
  const cards=[["todos","Todos",equipe.length,"colaboradores"],["1","1º turno",cnt["1"],"colaboradores"],["2","2º turno",cnt["2"],"colaboradores"],["3","3º turno",cnt["3"],"colaboradores"],["sem","Sem turno",cnt[""],cnt[""]?"defina o turno":"todos com turno"]];
  $("eqTurnos").innerHTML = cards.map(([k,l,n,sub])=>`<button type="button" class="eq-t ${k==="sem"&&n?"warn":""}" data-f="${k}" aria-pressed="${eqFiltro===k}"><span class="lbl">${l}</span><span class="v num">${n}</span><span class="sub">${sub}</span></button>`).join("");
  $("eqFiltroTxt").textContent = {todos:"Todos os turnos","1":"1º turno","2":"2º turno","3":"3º turno",sem:"Sem turno definido"}[eqFiltro];
  const q=($("eqBusca").value||"").toLowerCase();
  const rows=equipe.map((e,i)=>[e,i]).filter(([e])=> (eqFiltro==="todos" || (eqFiltro==="sem"? !e[2] : e[2]===eqFiltro)) && (!q || (e[0]+" "+e[1]).toLowerCase().includes(q)))
    .sort((a,b)=>a[0][0].localeCompare(b[0][0],"pt-BR"));
  $("eqBody").innerHTML = rows.length ? rows.map(([e,i])=>`<tr>
     <td><input type="text" id="eqn_${i}" data-i="${i}" data-c="0" value="${esc(e[0])}" aria-label="Nome"></td>
     <td><input type="text" id="eqf_${i}" data-i="${i}" data-c="1" value="${esc(e[1])}" list="funcList" aria-label="Função"></td>
     <td class="tc"><select id="eqt_${i}" data-i="${i}" data-c="2" aria-label="Turno">${["1","2","3",""].map(t=>`<option value="${t}" ${e[2]===t?"selected":""}>${TLBL[t]}</option>`).join("")}</select></td>
     <td style="text-align:right"><button class="linkbtn" type="button" data-del="${i}">Remover</button></td></tr>`).join("")
    : `<tr><td colspan="4" class="sub" style="padding:14px 10px">Nenhum colaborador neste filtro.</td></tr>`;
  $("funcList").innerHTML=[...new Set(equipe.map(e=>e[1]).filter(Boolean))].sort().map(f=>`<option value="${esc(f)}">`).join("");
}
$("eqTurnos").addEventListener("click",e=>{ const b=e.target.closest("[data-f]"); if(b){ eqFiltro=b.dataset.f; renderEquipe(); }});
$("eqBusca").addEventListener("input",renderEquipe);
$("eqBody").addEventListener("change",e=>{
  const i=e.target.dataset.i, c=e.target.dataset.c; if(i==null||c==null) return;
  const old=equipe[+i][0], v=e.target.value.trim();
  if(c==="0"){ if(!v){ e.target.value=old; return; } ausencias.forEach(a=>{ if(a.nome===old) a.nome=v; }); }
  equipe[+i][+c]=v; eqChanged();
});
$("eqBody").addEventListener("click",e=>{
  const i=e.target.dataset.del; if(i==null) return;
  const row=e.target.closest("tr");
  if(e.target.dataset.sure!=="1"){ e.target.dataset.sure="1"; e.target.textContent="Confirmar remoção"; e.target.style.color="var(--bad)"; return; }
  equipe.splice(+i,1); eqChanged();
});
$("btnAddPes").addEventListener("click",()=>{
  const n=$("nNome").value.trim(), f=$("nFunc").value.trim(), t=$("nTurno").value;
  if(!n){ $("addMsg").textContent="Informe o nome."; return; }
  if(equipe.some(e=>e[0].toLowerCase()===n.toLowerCase())){ $("addMsg").textContent="Esse nome já está no cadastro."; return; }
  equipe.push([n,f,t]); $("nNome").value=""; $("nFunc").value=""; $("addMsg").textContent=`${short(n)} incluído(a) no ${TLBL[t]}.`; eqChanged();
});

// importação
function lerLinhas(rows){
  const cel = c => String(c).trim().toLowerCase();
  const isN = c => /^(nome|colaborador)(es)?( completo)?$/.test(cel(c));
  const isX = c => /^(fun[cç][aã]o|cargo|turno)/.test(cel(c));
  let hi = rows.findIndex(r=>r.some(isN) && r.some(isX));
  if(hi<0) hi = rows.findIndex(r=>r.some(isN));
  if(hi<0) hi = rows.findIndex(r=>r.some(c=>/nome|colaborador/i.test(String(c))) && r.filter(c=>String(c).trim()).length>1); if(hi<0) return {erro:"Não encontrei a coluna de nome. A primeira linha precisa ter um cabeçalho com “Nome” ou “Colaborador”."};
  const h=rows[hi].map(c=>String(c).trim().toLowerCase());
  const iN=Math.max(h.findIndex(c=>/^(nome|colaborador)/.test(c)), h.findIndex(c=>/nome|colaborador/.test(c))), iF=h.findIndex(c=>/fun[cç]|cargo/.test(c)), iT=h.findIndex(c=>/turno/.test(c));
  const out=[]; const vistos=new Set();
  rows.slice(hi+1).forEach(r=>{ const n=String(r[iN]??"").trim(); if(!n||vistos.has(n.toLowerCase())) return; vistos.add(n.toLowerCase()); out.push([n, iF>=0?String(r[iF]??"").trim():"", iT>=0?normTurno(r[iT]):""]); });
  if(!out.length) return {erro:"O arquivo tem cabeçalho, mas nenhuma linha com nome abaixo dele."};
  return {lista:out, semTurnoCol:iT<0, semFuncCol:iF<0};
}
function previewImport(res,nome){
  if(res.erro){ pendImport=null; $("impPrev").innerHTML=`<div class="prev err">${res.erro}</div>`; return; }
  pendImport=res.lista;
  const c={"1":0,"2":0,"3":0,"":0}; res.lista.forEach(e=>c[e[2]]++);
  const novos=res.lista.filter(e=>!equipe.some(x=>x[0].toLowerCase()===e[0].toLowerCase())).length;
  $("impPrev").innerHTML=`<div class="prev"><span><b>${res.lista.length} colaboradores</b> em ${esc(nome)}: 1º turno ${c["1"]} · 2º turno ${c["2"]} · 3º turno ${c["3"]}${c[""]?` · <b style="color:var(--warn)">${c[""]} sem turno</b>`:""}.</span>
   ${res.semTurnoCol?`<span class="sub">A planilha não tem coluna Turno: defina o turno de cada um na tabela abaixo depois de importar.</span>`:""}
   <span class="sub">${novos} ainda não estão no cadastro atual (${equipe.length} pessoas).</span>
   <div class="row-actions" style="margin-top:0"><button class="btn" type="button" id="impRep">Substituir cadastro</button><button class="btn alt" type="button" id="impAdd">Adicionar ao cadastro</button><button class="linkbtn" type="button" id="impCan">Cancelar</button></div></div>`;
}
$("impPrev").addEventListener("click",e=>{
  if(!pendImport) return;
  if(e.target.id==="impRep"){ equipe=pendImport.map(x=>[...x]); }
  else if(e.target.id==="impAdd"){ pendImport.forEach(x=>{ const k=equipe.find(y=>y[0].toLowerCase()===x[0].toLowerCase()); if(k){ if(x[1]) k[1]=x[1]; if(x[2]) k[2]=x[2]; } else equipe.push([...x]); }); }
  else if(e.target.id!=="impCan") return;
  const msg = e.target.id==="impCan" ? "Importação cancelada." : `Cadastro atualizado: ${equipe.length} colaboradores.`;
  pendImport=null; $("impPrev").innerHTML=`<div class="prev"><span class="${e.target.id==="impCan"?"sub":"ok-msg"}">${msg}</span></div>`; $("eqFile").value="";
  if(e.target.id!=="impCan"){ eqFiltro="todos"; eqChanged(); }
});
async function lerArquivo(file){
  if(!file) return;
  $("dropTxt").innerHTML=`<b>${esc(file.name)}</b>`;
  try{
    let rows;
    if(/\.csv$/i.test(file.name)){
      const txt=await file.text(); const sepC=(txt.split("\n")[0].match(/;/g)||[]).length>(txt.split("\n")[0].match(/,/g)||[]).length?";":",";
      rows=txt.split(/\r?\n/).filter(l=>l.trim()).map(l=>l.split(sepC).map(c=>c.replace(/^"|"$/g,"").trim()));
    } else {
            const wb=XLSX.read(await file.arrayBuffer(),{type:"array"});
      // usa a primeira aba que tenha coluna de nome
      rows=null;
      for(const sn of wb.SheetNames){ const r=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:""}); if(r.some(x=>x.some(c=>/nome|colaborador/i.test(String(c))))){ rows=r; break; } }
      if(!rows) rows=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{header:1,defval:""});
    }
    previewImport(lerLinhas(rows),file.name);
  }catch(err){ previewImport({erro:"Não consegui ler esse arquivo. Confira se é .xlsx ou .csv e se não está protegido por senha."},file.name); }
}
$("eqFile").addEventListener("change",e=>lerArquivo(e.target.files[0]));
const dz=$("drop");
["dragenter","dragover"].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add("over");}));
["dragleave","drop"].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove("over");}));
dz.addEventListener("drop",e=>lerArquivo(e.dataTransfer.files[0]));

// downloads de planilhas
function planilha(linhas){
  const ws=XLSX.utils.aoa_to_sheet([["Nome","Função","Turno"],...linhas]); ws["!cols"]=[{wch:36},{wch:30},{wch:8}];
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Colaboradores");
  return new Blob([XLSX.write(wb,{bookType:"xlsx",type:"array"})],{type:MIME.xlsx});
}
function baixar(nome,linhas){
  baixarArquivo(`${nome}.xlsx`, planilha(linhas));
  const msg=$("eqDl"); msg.className="ok-msg"; msg.textContent=`${nome}.xlsx baixado.`;
}
$("btnModelo").addEventListener("click",()=>baixar("modelo-colaboradores",[["Maria da Silva","Conferente I","1"],["João Souza","Auxiliar de Expedição I","2"]]));
$("btnExport").addEventListener("click",()=>{ if(!equipe.length){ $("eqDl").className="sub"; $("eqDl").textContent="O cadastro está vazio."; return; } baixar("cadastro-colaboradores",equipe.map(e=>[e[0],e[1],e[2]])); });

// ---------- abas ----------
const tabs=[["t-painel","v-painel"],["t-lanc","v-lanc"],["t-sep","v-sep"],["t-eq","v-eq"]];
function show(id){ tabs.forEach(([t,v])=>{ $(t).setAttribute("aria-selected",t===id); $(v).hidden = t!==id; }); if(id==="t-painel") fitShot(); }
tabs.forEach(([t])=>$(t).addEventListener("click",()=>show(t)));
const h=location.hash.slice(1); if(h==="lancamento") show("t-lanc"); if(h==="relatorios") show("t-sep"); if(h==="equipe") show("t-eq"); if(h==="apoio") abrirApoio();

$("dataTurno").value=hojeISO();
$("dataTurno").addEventListener("change",()=>{ if(!$("dataTurno").value) $("dataTurno").value=hojeISO(); setData(); });
$("turnoSeg").addEventListener("click",e=>{ const b=e.target.closest("button[data-t]"); if(b && b.dataset.t!==turno) setTurno(b.dataset.t); });
$("obsTxt").addEventListener("input",()=>{ S[turno].obs=$("obsTxt").value; renderShot(); salvarDia(); });
$("btnSaveAll").addEventListener("click",()=>saveShot([1,2,3]));
$("btnSave1").addEventListener("click",()=>saveShot([1]));
$("btnSave2").addEventListener("click",()=>saveShot([2]));
$("btnSave3").addEventListener("click",()=>saveShot([3]));
$("logoTopo").src=LOGO_O;
buildForm(); renderEquipe(); setData();
