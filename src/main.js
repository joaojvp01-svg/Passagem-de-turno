import "./style.css";
import "@fontsource/fira-sans/400.css";
import "@fontsource/fira-sans/600.css";
import "@fontsource/fira-sans/700.css";
import "@fontsource/fira-sans-condensed/600.css";
import "@fontsource/fira-sans-condensed/700.css";
import "@fontsource/sora/700.css";
import "@fontsource/sora/800.css";
import * as XLSX from "xlsx";
import LOGO_O from "./assets/logo-laranja.png";
import LOGO_W from "./assets/logo-branco.png";
import PUP from "./assets/pet-filhote.png";
import DUO from "./assets/pet-dupla.png";
import { MIME, baixarArquivo, entregar, imagemDe } from "./compartilhar.js";
import "./planejamento.js";

const fmt = (n,d=0)=> n==null||isNaN(n) ? "—" : n.toLocaleString("pt-BR",{minimumFractionDigits:d,maximumFractionDigits:d});
const t = kg => fmt(kg/1000,1);
const title = s => s.toLowerCase().replace(/(^|\s)\S/g,m=>m.toUpperCase()).replace(/\b(Da|De|Do|Dos|Das|E)\b/g,m=>m.toLowerCase());
const short = s => { const p = title(s).split(" ").filter(w=>!/^(da|de|do|dos|das|e)$/.test(w)); return p.length>1 ? p[0]+" "+p[p.length-1] : p[0]; };

// ---------- dados ----------
// O cadastro da equipe e os relatórios entram pelo próprio site (abas Equipe e Relatórios)
// e ficam salvos no navegador. Nenhum nome de colaborador fica no código.
let sep = [];   // [nome, pedidos|null, itens, peso_picking, peso_onda, usuario, peso_semi_onda]
let conf = [];  // [nome, pedidos, itens, peso_total, usuario]
let mov = null; // movimentação de estoque: {pick, onda, semi, n, ini, fim, outros, usuarios:{LOGIN:{pick,onda,semi,n}}, fonte}
let equipe = []; // [nome, função, turno "1"|"2"|"3"|"", usuário WMS]
const motivos = ["Falta injustificada","Abono chefia","Atestado médico","Presença parcial","Férias","Afastamento"];
// Férias e afastamento ficam numa lista à parte e não entram no percentual de absenteísmo (saem do quadro disponível).
const FORA_QUADRO = new Set(["Férias","Afastamento"]);
// Ausências de mais de um dia (ou com início anterior) valem em todos os dias do período, no turno em que foram lançadas.
const KAF = "passagemTurno.afast";
let afast = []; try{ afast = JSON.parse(localStorage.getItem(KAF)||"[]"); }catch(e){}
const saveAf = () => { try{ localStorage.setItem(KAF, JSON.stringify(afast)); }catch(e){} };
const isoAdd = (iso,n) => { const d=new Date(iso+"T12:00:00"); d.setDate(d.getDate()+n); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`; };
const diasEntre = (a,b) => Math.round((new Date(b+"T12:00:00")-new Date(a+"T12:00:00"))/864e5);
const ddmm = iso => iso.slice(8,10)+"/"+iso.slice(5,7);
function ausDoTurno(){ // lista do dia: lançamentos de um dia + períodos ativos
  const hoje=dataISO();
  const curtas = ausencias.map((a,i)=>({...a, ini:hoje, dias:1, n:1, i}));
  const longas = afast.filter(a=>a.turno===turno && a.ini<=hoje && diasEntre(a.ini,hoje)<a.dias).map(a=>({...a, n:diasEntre(a.ini,hoje)+1, af:a.id}));
  return [...curtas, ...longas];
}

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
const semiDe = r => r[6]||0;
const totSep = r => r[3]+r[4]+semiDe(r);
// Totais do VD: a movimentação de estoque manda nos pesos de picking, onda e semi-onda;
// sem ela, valem os pesos do relatório de produtividade.
function totals(){
  const pick = mov ? mov.pick : sep.reduce((a,r)=>a+r[3],0), onda = mov ? mov.onda : sep.reduce((a,r)=>a+r[4],0), semi = mov ? mov.semi : sep.reduce((a,r)=>a+semiDe(r),0);
  const itens = sep.length ? sep.reduce((a,r)=>a+r[2],0) : (mov ? mov.n : 0);
  const ped = !sep.length || sep.some(r=>r[1]==null) ? null : sep.reduce((a,r)=>a+r[1],0);
  const confKg = conf.reduce((a,r)=>a+r[3],0), confIt = conf.reduce((a,r)=>a+r[2],0);
  return {pick,onda,semi,sepKg:pick+onda+semi,itens,ped,confKg,confIt};
}

function render(){
  const T = totals();
  // alertas: o que conferir antes de enviar (não aparecem nas imagens)
  const al = [];
  if(M.op_conf>M.op_rec) al.push(`<b>Recebimento:</b> ${fmt(M.op_conf)} carretas conferidas para ${fmt(M.op_rec)} recebidas. Se há saldo do turno anterior, registre à parte; se não, corrija o lançamento.`);
  if(T.confKg>T.sepKg*1.15 && T.sepKg) al.push(`<b>Conferência acima da separação:</b> ${t(T.confKg)} t conferidas contra ${t(T.sepKg)} t separadas. Normal se a conferência incluiu volume separado antes.`);
  if(M.vd_prog==null) al.push(`<b>VD sem total programado:</b> sem ele a página 1 não mostra o pendente nem o percentual separado.`);
  if(!mov) al.push(`<b>Sem Movimentação de Estoque</b> neste turno: os pesos de picking, onda e semi-onda vêm dela (aba Relatórios).`);
  if(!sep.length && !conf.length) al.push(`<b>Sem relatório de separação ou conferência</b> neste turno: suba na aba Relatórios para a página 3 mostrar pedidos e conferência.`);
  if(!daTurno().length) al.push(`<b>Nenhum colaborador cadastrado no ${turno}º turno:</b> importe a equipe na aba Equipe para o absenteísmo funcionar.`);
  $("alerts").innerHTML = al.length ? `<div class="callout"><div><span class="lbl" style="color:var(--warn)">Antes de enviar, confira</span><ul>${al.map(a=>`<li>${a}</li>`).join("")}</ul></div></div>` : "";
  drawAbs();
  renderSep();
  renderShot();
  renderExemplo();
  salvarDia();
}

// ---------- modo foto ----------
let turno = "1";
const hojeISO = () => { const d=new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}-${String(d.getDate()).padStart(2,"0")}`; };
const dataISO = () => $("dataTurno").value || hojeISO();
const dataBR = (o) => { const [y,m,d]=dataISO().split("-").map(Number); return new Date(y,m-1,d).toLocaleDateString("pt-BR",o); };
const MOT = {"Férias":["#2f8a6b","#dcf1e8","#1f7a5a","#dcf1e8"],"Afastamento":["#7a5aa8","#ece4f6","#6a4a98","#ece4f6"],"Falta injustificada":["var(--bad)","var(--bad-soft)","#c2372f","#f9dedb"],"Atestado médico":["var(--warn)","var(--warn-soft)","#a86b12","#fbf0d9"],"Abono chefia":["var(--info)","var(--info-soft)","#2f6aa8","#dfeaf6"],"Presença parcial":["var(--muted)","var(--surface-2)","#646b74","#eceef0"]};
const funcOf = n => (equipe.find(e=>e[0]===n)||[])[1]||"";
const daTurno = () => equipe.filter(e=>e[2]===turno);
function fillAbsNome(){
  const l=daTurno().map(e=>e[0]).sort((a,b)=>a.localeCompare(b,"pt-BR"));
  $("absNome").innerHTML = l.length ? l.map(n=>`<option>${n}</option>`).join("") : `<option value="">Nenhum colaborador cadastrado neste turno</option>`;
  $("absAdd").disabled = !l.length;
}
const GRUPOS = [["Conferentes",f=>/CONFERENTE/.test(semAc(f))],["Aux. de expedição",f=>/AUXILIAR DE EXPEDICAO/.test(semAc(f))],["Op. de empilhadeira",f=>/EMPILHADEIRA/.test(semAc(f))],["Outras funções",f=>!/CONFERENTE|AUXILIAR DE EXPEDICAO|EMPILHADEIRA/.test(semAc(f))]];
function absStats(){
  const L=ausDoTurno(), ab=L.filter(a=>!FORA_QUADRO.has(a.mot)), fq=L.length-ab.length;
  const q=Math.max(0,daTurno().length-fq), n=ab.length, parc=ab.filter(a=>a.mot==="Presença parcial").length;
  const cont={}; motivos.forEach(m=>cont[m]=0); L.forEach(a=>cont[a.mot]=(cont[a.mot]||0)+1);
  const pct = q ? (n-parc)/q*100 : 0;
  return {q,n,parc,pres:q-n+parc,cont,pct,fq,L,ab,fora:L.filter(a=>FORA_QUADRO.has(a.mot))};
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
    T1("Separação semi onda", fmt(T.semi), T.sepKg?`${fmt(T.semi/T.sepKg*100)}% do sep.`:""),
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
  const linha = a=>{ const c=MOT[a.mot]||MOT["Falta injustificada"]; return `<div class="x-row"><b>${title(a.nome)}</b><span>${funcOf(a.nome)}</span><i style="background:${c[3]};color:${c[2]}">${a.mot}</i><em>${a.dias>1?`${a.n} // ${a.dias}<small>até ${ddmm(isoAdd(a.ini,a.dias-1))}</small>`:"1 dia"}</em></div>`; };
  const rows = A.ab.map(linha).join("") + (A.fora.length ? `<div class="ab-sub">Férias e afastamento · ${A.fora.length}</div>`+A.fora.map(linha).join("") : "");
  const barra = motivos.filter(m=>!FORA_QUADRO.has(m)).map(m=>A.cont[m]?`<i style="width:${A.cont[m]/Math.max(1,A.q)*100}%;background:${MOT[m][2]}"></i>`:"").join("");
  const chips = motivos.filter(m=>A.cont[m] && !FORA_QUADRO.has(m)).map(m=>`<span style="background:${MOT[m][3]};color:${MOT[m][2]}"><b>${A.cont[m]}</b> ${m}</span>`).join("");
  const ab = `<div class="x-ab"><div class="hd"><span>Absenteísmo · detalhe</span></div>
    <div class="x-ab-b"><div class="ab-top"><span class="ab-n ${A.n?"":"zero"}">${fmt(A.pct,1)}%</span><div class="ab-m"><b>de absenteísmo</b><span>quadro do turno: ${A.q}${A.fq?` (fora ${A.fq} em férias/afastamento)`:""}</span>${A.parc?`<span>${A.parc} com presença parcial</span>`:""}<div class="ab-bar">${barra}</div><div class="ab-chips">${chips}</div></div></div>
    <div class="ab-lst">${A.L.length ? (A.n?"":`<div class="ab-ok">Nenhuma ausência no turno.</div>`)+rows : `<div class="ab-ok">Equipe completa. Nenhuma ausência registrada.</div>`}</div></div></div>`;
  const o=$("obsTxt").value.trim();
  const obs = `<div class="x-blk"><div class="x-blk-h">Observações do turno</div><div class="x-obs">${o?o.replace(/[&<>]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;"}[c])):'<span class="na">Sem observações.</span>'}</div></div>`;
  $("pg2").innerHTML = head(DUO) + `<div class="x-b"><div class="x-sec"><div class="x-band"><b>OPERAÇÃO</b>${opTag}</div>${P("Carretas armazenadas",M.op_arm||0,M.op_prog,`${fmt(M.op_arm||0)} de ${fmt(M.op_prog)}`)}${op}</div>${ab}${obs}</div>` + foot(2);

  // ---- página 3: produtividade de todos, separando auxiliares (picking) e operadores (onda e semi-onda)
  const ps=pessoasSep();
  const aux=ps.filter(p=>p.grupo==="aux").sort((a,b)=>b.tot-a.tot), ops=ps.filter(p=>p.grupo==="op").sort((a,b)=>b.tot-a.tot);
  const confR=[...conf].map(r=>({n:r[0],it:r[2],ped:r[1],kg:r[3]})).sort((a,b)=>b.kg-a.kg);
  const confPed=conf.reduce((a,r)=>a+r[1],0);
  const nomeCurto = (n,u) => { if(u && semAc(n)===semAc(u)) return String(u).toUpperCase(); const x=title(n); return x.length>26 ? short(n) : x; };
  const lista = (arr, val, sub, barras) => { const mx=Math.max(1,...arr.map(val)); return arr.length ? arr.map((r,i)=>`<div class="x-rk2"><span class="pos ${i===0?"first":""}">${i+1}º</span><div class="who"><b>${nomeCurto(r.nome||r.n, r.usr)}</b><div class="tr">${barras(r,mx)}</div></div><span class="vv">${t(val(r))} t<small>${sub(r)}</small></span></div>`).join("") : `<div class="x-obs"><span class="na">Sem dados neste turno.</span></div>`; };
  const subSep = r => r.temProd ? (r.ped==null ? `${fmt(r.it)} itens` : `${fmt(r.it)} itens · ${fmt(r.ped)} ped.`) : `${fmt(r.mv??r.it)} movim.`;
  const pesoOp = r => r.onda+r.semi;
  $("pg3").innerHTML = head(PUP,"Produtividade") + `<div class="x-b">
    <div class="p3-top">
      <div class="x-sec"><div class="x-band"><b>SEPARAÇÃO</b>${tag(ps.length?"ok":"na",ps.length?`${aux.length} aux. · ${ops.length} op.`:"Sem relatório")}</div>
        ${G([T1("Peso separado",T.sepKg?fmt(T.sepKg):"0","kg","hl"),T1("Picking",fmt(T.pick),T.sepKg?`${fmt(T.pick/T.sepKg*100)}% do total`:""),T1("Onda",fmt(T.onda),T.sepKg?`${fmt(T.onda/T.sepKg*100)}% do total`:""),T1("Semi-onda",fmt(T.semi),T.sepKg?`${fmt(T.semi/T.sepKg*100)}% do total`:"")],4)}</div>
      <div class="x-sec"><div class="x-band"><b>CONFERÊNCIA</b>${tag(conf.length?"ok":"na",conf.length?`${conf.length} conferentes`:"Sem relatório")}</div>
        ${G([T1("Peso conferido",T.confKg?fmt(T.confKg):"0","kg","hl"),T1("Pedidos",fmt(confPed)),T1("Itens",fmt(T.confIt))],3)}</div>
    </div>
    <div class="p3-cols ${aux.length>12?"":"curta"}">
      <div class="x-blk p3-aux ${aux.length>12?"duas":""}"><div class="x-blk-h">Auxiliares · separação de picking <span class="cnt">${aux.length}</span></div><div class="p3-lst" style="--linhas:${Math.ceil(aux.length/2)}">${lista(aux,r=>r.tot,subSep,(r,mx)=>`<i style="width:${r.pick/mx*100}%;background:#ff6900"></i><i style="width:${(r.onda+r.semi)/mx*100}%;background:#2b2724"></i>`)}</div></div>
      <div class="p3-dir"><div class="x-blk"><div class="x-blk-h">Operadores · onda e semi-onda <span class="cnt">${ops.length}</span></div>${lista(ops,pesoOp,r=>`onda ${t(r.onda)} · semi ${t(r.semi)}`,(r,mx)=>`<i style="width:${r.onda/mx*100}%;background:#2b2724"></i><i style="width:${r.semi/mx*100}%;background:#a49b93"></i>`)}</div>
      <div class="x-blk"><div class="x-blk-h">Conferentes <span class="cnt">${confR.length}</span></div>${lista(confR,r=>r.kg,r=>`${fmt(r.it)} itens · ${fmt(r.ped)} ped.`,(r,mx)=>`<i style="width:${r.kg/mx*100}%;background:#2f6aa8"></i>`)}</div></div>
    </div>
    <div class="p3-leg"><span><i style="background:#ff6900"></i>picking</span><span><i style="background:#2b2724"></i>onda</span><span><i style="background:#a49b93"></i>semi-onda</span><span><i style="background:#2f6aa8"></i>conferência</span></div>
  </div>` + foot(3);
  fitShot();
}
function fitShot(){
  [1,2,3].forEach(n=>{ const st=$("st"+n), w=st.clientWidth; if(!w) return;
    const sc=Math.min(1.25,w/800); $("fit"+n).style.transform=`scale(${sc})`; st.style.height=($("pg"+n).offsetHeight*sc+2)+"px"; });
}
window.addEventListener("resize",fitShot);

const shotBlob = id => imagemDe($(id));
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
const blankTurno = () => ({M:blankM(),sep:[],conf:[],mov:null,aus:[],obs:"",fonte:"",fonteConf:""});
const blankDia = () => ({"1":blankTurno(),"2":blankTurno(),"3":blankTurno()});
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
  const x=S[turno]; M=x.M; sep=x.sep; conf=x.conf; mov=x.mov||null; ausencias=x.aus; sepFonte=x.fonte||""; confFonte=x.fonteConf||x.fonte||""; imp.sep.fonte=sepFonte; imp.conf.fonte=confFonte; imp.sep.pend=imp.conf.pend=imp.sep.undo=imp.conf.undo=null; $("obsTxt").value=x.obs||"";
  document.querySelectorAll("#formGrid input[data-k]").forEach(i=>{ i.value = M[i.dataset.k]??""; });
  sepUndo=null; movUndo=null; $("sepStatus").innerHTML=""; $("confStatus").innerHTML=""; $("movStatus").innerHTML=""; $("absMsg").textContent=""; $("absIni").value=dataISO(); $("absDias").value=1; $("absFimTxt").textContent="";
  $("saveMsg").className="save-msg"; $("saveMsg").textContent="No celular, abre o compartilhamento (WhatsApp). No computador, baixa as imagens.";
}
function setTurno(tn){
  turno=tn;
  document.querySelectorAll("#turnoSeg button").forEach(b=>b.setAttribute("aria-pressed",b.dataset.t===tn));
  $("turnoLbl").textContent=`${tn}º turno`;
  vincular(); fillAbsNome(); render();
}
function setData(){ S=carregarDia(dataISO()); vincular(); fillAbsNome(); render(); }

// ---------- exemplo com dados fictícios (para conhecer o site antes de usar de verdade) ----------
const KEX="passagemTurno.exemplo";
const EX_EQUIPE=[["Colaborador Exemplo 01","Conferente I","1"],["Colaborador Exemplo 02","Conferente I","1"],["Colaborador Exemplo 03","Conferente II","1"],["Colaborador Exemplo 04","Auxiliar de Expedição I","1"],["Colaborador Exemplo 05","Auxiliar de Expedição I","1"],["Colaborador Exemplo 06","Auxiliar de Expedição I","1"],["Colaborador Exemplo 07","Auxiliar de Expedição I","1"],["Colaborador Exemplo 08","Auxiliar de Expedição I","1"],["Colaborador Exemplo 09","Operador de Empilhadeira I","1"],["Colaborador Exemplo 10","Operador de Empilhadeira II","1"],["Colaborador Exemplo 11","Conferente I","1"],["Colaborador Exemplo 12","Auxiliar de Expedição I","1"]];
const EX_SEP=[["COLABORADOR EXEMPLO 04",310,478,9194,0,"EXEMPLO04"],["COLABORADOR EXEMPLO 10",10,24,0,10741,"EXEMPLO10"],["COLABORADOR EXEMPLO 09",2,14,0,8153,"EXEMPLO09"],["COLABORADOR EXEMPLO 05",5,75,7082,0,"EXEMPLO05"],["COLABORADOR EXEMPLO 06",3,44,5968,0,"EXEMPLO06"],["COLABORADOR EXEMPLO 07",5,58,4483,0,"EXEMPLO07"],["COLABORADOR EXEMPLO 08",217,285,3208,0,"EXEMPLO08"],["COLABORADOR EXEMPLO 12",2,83,2849,0,"EXEMPLO12"]];
const EX_CONF=[["COLABORADOR EXEMPLO 01",6,171,14447,"EXEMPLO01"],["COLABORADOR EXEMPLO 02",45,401,13262,"EXEMPLO02"],["COLABORADOR EXEMPLO 03",5,127,8562,"EXEMPLO03"],["COLABORADOR EXEMPLO 11",310,478,9195,"EXEMPLO11"]];
const EX_M={vd_prog:107395,vd_carrosProg:16,vd_carrosExp:16,vet_prog:null,vet_blocSep:3,vet_blocPend:0,vet_blocConf:3,vet_rotasSep:0,vet_rotasPend:0,vet_rotasConf:0,vet_carrosProg:13,vet_carrosExp:12,op_prog:4,op_rec:4,op_conf:4,op_arm:4,op_recre:2,op_missoes:25,op_estoque:3195,op_posicoes:653,av_porto:0,av_dourado:0,hidraulicos:25};
const emExemplo = () => { try{ return localStorage.getItem(KEX)===dataISO()+"|"+turno; }catch(e){ return false; } };
function renderExemplo(){
  const vazio = !equipe.length && !sep.length && !conf.length && Object.values(M).every(v=>v==null);
  $("exemplo").innerHTML = emExemplo()
    ? `<div class="ex-card"><span><b>Você está vendo um exemplo com dados fictícios.</b> Quando for usar de verdade, limpe o exemplo e importe a equipe e o relatório do dia.</span><button class="btn alt" type="button" id="exLimpar">Limpar exemplo</button></div>`
    : (vazio ? `<div class="ex-card"><span><b>Primeira vez aqui?</b> Carregue um exemplo com dados fictícios para ver como as 3 páginas ficam preenchidas.</span><button class="btn" type="button" id="exCarregar">Carregar exemplo</button></div>` : "");
}
$("exemplo").addEventListener("click",e=>{
  if(e.target.id==="exCarregar"){
    equipe=EX_EQUIPE.map(r=>[...r]); saveEq();
    Object.assign(M,EX_M); sep.push(...EX_SEP.map(r=>[...r])); conf.push(...EX_CONF.map(r=>[...r]));
    ausencias.push({nome:"Colaborador Exemplo 06",mot:"Falta injustificada"},{nome:"Colaborador Exemplo 11",mot:"Atestado médico"});
    S[turno].obs="Exemplo: rota 12 aguardando liberação do cliente."; S[turno].fonte="exemplo fictício"; sepFonte="exemplo fictício";
    try{ localStorage.setItem(KEX,dataISO()+"|"+turno); }catch(err){}
    vincular(); fillAbsNome(); renderEquipe(); render();
  }
  if(e.target.id==="exLimpar"){
    equipe=equipe.filter(r=>!/^Colaborador Exemplo \d+$/.test(r[0])); saveEq();
    S[turno]=blankTurno();
    try{ localStorage.removeItem(KEX); }catch(err){}
    vincular(); fillAbsNome(); renderEquipe(); render();
  }
});

// ---------- formulário ----------
function buildForm(){
  $("formGrid").innerHTML = FORM.map(([g,fields])=>`<fieldset><legend>${g}</legend>${fields.map(([k,l])=>`<div class="f"><label for="in_${k}">${l}</label><input type="number" min="0" step="any" id="in_${k}" data-k="${k}" value="${M[k]??""}" placeholder="—"></div>`).join("")}</fieldset>`).join("");
  $("formGrid").addEventListener("input",e=>{
    const k=e.target.dataset.k; if(!k) return;
    const v=e.target.value.trim(); M[k]= v==="" ? null : Number(v); render();
  });
  fillAbsNome();
  $("absMotSel").innerHTML = motivos.map(m=>`<option>${m}</option>`).join("");
  const fimTxt = () => { const ini=$("absIni").value||dataISO(), d=Math.max(1,Math.round(+$("absDias").value||1));
    $("absFimTxt").textContent = d>1 ? `De ${ddmm(ini)} a ${ddmm(isoAdd(ini,d-1))}. Aparece em todos os dias do período, no ${turno}º turno.` : (ini!==dataISO()?`Só no dia ${ddmm(ini)}.`:""); };
  ["absIni","absDias"].forEach(id=>$(id).addEventListener("input",fimTxt));
  $("absIni").value=dataISO(); fimTxt();
  $("absAdd").addEventListener("click",()=>{
    const nome=$("absNome").value; if(!nome) return;
    const mot=$("absMotSel").value, ini=$("absIni").value||dataISO(), dias=Math.max(1,Math.min(365,Math.round(+$("absDias").value||1))), hoje=dataISO();
    const fim=isoAdd(ini,dias-1);
    const choca = afast.some(a=>a.nome===nome && a.turno===turno && !(isoAdd(a.ini,a.dias-1)<ini || a.ini>fim)) || (ini<=hoje && fim>=hoje && ausencias.some(a=>a.nome===nome));
    if(choca){ $("absMsg").textContent=`${short(nome)} já tem ausência lançada nesse período.`; return; }
    if(dias===1 && ini===hoje) ausencias.push({nome,mot});
    else { afast.push({id:Date.now().toString(36), nome, mot, ini, dias, turno}); saveAf(); }
    $("absMsg").textContent = `${short(nome)} registrado(a): ${mot.toLowerCase()}` + (dias>1?`, ${dias} dias (${ddmm(ini)} a ${ddmm(fim)}).`:".") + (ini>hoje?` Começa em ${ddmm(ini)}.`:"");
    $("absDias").value=1; $("absIni").value=hoje; fimTxt(); render();
  });
}
function drawAbs(){
  const A=absStats();
  $("absN").textContent=A.n; $("absN").className="n num"+(A.n?"":" zero");
  $("absMeta").innerHTML=`<b>${A.n===1?"ausência":"ausências"}</b><span>${A.pres} de ${A.q} presentes · ${fmt(A.pct,1)}% de absenteísmo${A.fq?` · ${A.fq} em férias/afastamento (fora do quadro)`:""}</span>`;
  $("absBar").innerHTML = motivos.filter(m=>!FORA_QUADRO.has(m)).map(m=>A.cont[m]?`<i style="width:${A.cont[m]/Math.max(1,A.q)*100}%;background:${MOT[m][0]}"></i>`:"").join("");
  $("absMot").innerHTML = motivos.map(m=>`<div class="mot" style="--sbg:${MOT[m][1]};--sfg:${MOT[m][0]};${A.cont[m]?"":"opacity:.45"}"><span class="n">${A.cont[m]}</span><span class="l">${m}</span></div>`).join("");
  $("absFunc").innerHTML = `<span class="lbl" style="margin-top:4px">Presentes por função</span>` + GRUPOS.map(([g,fn])=>{
    const tot=daTurno().filter(e=>fn(e[1])).length; if(!tot) return "";
    const fora=A.L.filter(a=>a.mot!=="Presença parcial" && fn(funcOf(a.nome))).length; const pr=tot-fora;
    return `<div class="func"><span>${g}</span><div class="tr"><i style="width:${pr/tot*100}%;${fora?'background:var(--warn)':''}"></i></div><span class="num">${pr}/${tot}</span></div>`; }).join("");
  const tr = a=>{ const c=MOT[a.mot]||MOT["Falta injustificada"]; const fim=isoAdd(a.ini,a.dias-1);
    return `<tr><td><b>${title(a.nome)}</b></td><td>${funcOf(a.nome)}</td><td><span class="mbadge" style="--sbg:${c[1]};--sfg:${c[0]}">${a.mot}</span></td>
    <td class="dias">${a.n===0?`<b>${a.dias} ${a.dias>1?"dias":"dia"}</b><small>começa ${ddmm(a.ini)}</small>`:a.dias>1?`<b>${a.n} // ${a.dias}</b><small>${ddmm(a.ini)} a ${ddmm(fim)}</small>`:"1 dia"}</td>
    <td style="text-align:right">${a.af?`<button class="linkbtn" type="button" data-af="${a.af}">${a.ini>=dataISO()?"Remover":"Encerrar"}</button>`:`<button class="linkbtn" type="button" data-i="${a.i}">Remover</button>`}</td></tr>`; };
  const prox = afast.filter(a=>a.turno===turno && a.ini>dataISO());
  $("absTable").innerHTML = (A.L.length || prox.length)
    ? `<div class="abs-tbl-wrap"><table class="abs-tbl"><thead><tr><th>Colaborador</th><th>Função</th><th>Motivo</th><th>Dias</th><th></th></tr></thead><tbody>${A.ab.map(tr).join("")}${A.fora.length?`<tr class="abs-sec"><td colspan="5">Férias e afastamento</td></tr>`+A.fora.map(tr).join(""):""}${prox.length?`<tr class="abs-sec"><td colspan="5">Programadas (começam depois de ${ddmm(dataISO())})</td></tr>`+prox.map(a=>tr({...a,n:0,af:a.id})).join(""):""}</tbody></table></div>`
    : `<div class="abs-empty"><span><b>Nenhuma ausência registrada neste turno.</b> Use o formulário para lançar faltas, atestados, abonos, presença parcial, férias e afastamentos.</span></div>`;
}
$("absTable").addEventListener("click",e=>{
  const af=e.target.dataset.af;
  if(af){ const k=afast.findIndex(a=>a.id===af); if(k<0) return; const a=afast[k], hoje=dataISO();
    if(a.ini>=hoje) afast.splice(k,1); else a.dias=diasEntre(a.ini,hoje); // encerra ontem; o histórico fica
    saveAf(); $("absMsg").textContent = a.ini>=hoje ? `${short(a.nome)} removido(a).` : `${short(a.nome)}: ausência encerrada em ${ddmm(isoAdd(hoje,-1))}.`; render(); return; }
  const i=e.target.dataset.i; if(i==null) return; ausencias.splice(+i,1); render();
});

// ---------- relatório de separação ----------
let sepFonte="", confFonte="", sepUndo=null, sepView="sep", sepSort={k:"tot",d:-1};
const porLogin = usr => usr ? equipe.find(e=>e[3] && semAc(e[3])===semAc(usr)) : null;
// login do WMS -> nome, aprendido sempre que entra um relatório de separação com USUARIO e NOME
const KLOG="passagemTurno.logins";
let logins = (()=>{ try{ return JSON.parse(localStorage.getItem(KLOG)||"{}")||{}; }catch(e){ return {}; } })();
function aprenderLogins(linhas){ let n=0; linhas.forEach(r=>{ const u=semAc(r[5]||""); if(u && r[0] && semAc(r[0])!==u && logins[u]!==r[0]){ logins[u]=r[0]; n++; } }); if(n) try{ localStorage.setItem(KLOG,JSON.stringify(logins)); }catch(e){} }
const nomeDoLogin = u => { const e=porLogin(u); return e ? e[0] : (logins[semAc(u)] || u); };
function funcWMS(nome, usr){
  const pl=porLogin(usr); if(pl) return pl[1];
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
// ---------- Relatórios: dois envios separados ----------
// 1) Movimentação de Estoque (WMS): uma linha por movimentação. Dá os totais do VD
//    (separação de picking, onda e semi-onda) e o que cada usuário movimentou.
// 2) Relatório de produtividade: separação (NOME, NO_PEDIDOS, NO_ITENS, PESO_PICKING, PESO_ONDA)
//    e conferência (NOME, NO_PEDIDOS, ITENS, PESO_TOTAL) por colaborador.
const H_ = r => r.map(c=>semAc(c).replace(/\s+/g,"_"));
const tipoMov = d => /SEMI/.test(d) ? "semi" : (/^ONDA/.test(d) ? "onda" : (/SEPARACAO DE PICKING/.test(d) ? "pick" : null));
function lerMovimentacao(rows){
  const hi = rows.findIndex(r=>{ const h=H_(r); return h.includes("DESCRICAO") && h.includes("USUARIO") && h.includes("PESO") && (h.includes("TRANSACAO")||h.includes("DATA_TRANSACAO")); });
  if(hi<0) return null;
  const h=H_(rows[hi]), iD=h.indexOf("DESCRICAO"), iU=h.indexOf("USUARIO"), iP=h.indexOf("PESO"), iT=h.indexOf("DATA_TRANSACAO");
  const m={pick:0,onda:0,semi:0,n:0,outros:{},usuarios:{},ini:null,fim:null};
  for(const r of rows.slice(hi+1)){
    const d=semAc(r[iD]), u=String(r[iU]??"").trim().toUpperCase(); if(!d||!u) continue;
    const kg=toNum(r[iP]), tp=tipoMov(d);
    const dt = r[iT] instanceof Date ? r[iT].toISOString() : null; if(dt){ if(!m.ini||dt<m.ini) m.ini=dt; if(!m.fim||dt>m.fim) m.fim=dt; }
    if(!tp){ const nm=String(r[iD]).trim().toLowerCase(); m.outros[nm]=(m.outros[nm]||0)+kg; continue; }
    const x = m.usuarios[u] || (m.usuarios[u]={pick:0,onda:0,semi:0,n:0}); x[tp]+=kg; x.n++; m[tp]+=kg; m.n++;
  }
  return m.n ? m : null;
}
function lerProdutividade(rows){
  // devolve {tipo:"sep"|"conf", linhas, dias:{"aaaa-mm-dd": linhas}|null} ou null
  for(let i=0;i<Math.min(rows.length,15);i++){
    const h=H_(rows[i]); const ix=n=>h.indexOf(n);
    const iN=ix("NOME"); if(iN<0) continue;
    const iU=ix("USUARIO"), iP=ix("NO_PEDIDOS")>=0?ix("NO_PEDIDOS"):ix("PEDIDOS"), iI=ix("NO_ITENS")>=0?ix("NO_ITENS"):ix("ITENS"), iDt=ix("DATA");
    const body=rows.slice(i+1).filter(r=>String(r[iN]??"").trim() && !/^total/i.test(String(r[iN])));
    let tipo=null, montar;
    if(ix("PESO_PICKING")>=0 || ix("PESO_ONDA")>=0){
      const a=ix("PESO_PICKING"), b=ix("PESO_ONDA"), c=[ix("PESO_SEMI_ONDA"),ix("PESO_SEMIONDA"),ix("SEMI_ONDA")].find(k=>k>=0)??-1;
      tipo="sep"; montar = r=>[String(r[iN]).trim(), toNum(r[iP]), toNum(r[iI]), a>=0?toNum(r[a]):0, b>=0?toNum(r[b]):0, iU>=0?String(r[iU]).trim().toUpperCase():"", c>=0?toNum(r[c]):0];
    } else if(ix("PESO_TOTAL")>=0){
      const a=ix("PESO_TOTAL");
      tipo="conf"; montar = r=>[String(r[iN]).trim(), toNum(r[iP]), toNum(r[iI]), toNum(r[a]), iU>=0?String(r[iU]).trim().toUpperCase():""];
    } else continue;
    // soma por pessoa (o mesmo login pode aparecer mais de uma vez no mesmo dia)
    const somar = lista => { const acc=new Map(); lista.forEach(r=>{ const x=montar(r); const k=x[tipo==="sep"?5:4]||semAc(x[0]); const y=acc.get(k); if(!y) acc.set(k,x); else for(const n of (tipo==="sep"?[1,2,3,4,6]:[1,2,3])) y[n]=(y[n]??0)+(x[n]??0); }); return [...acc.values()]; };
    let dias=null;
    if(iDt>=0){ const g={}; body.forEach(r=>{ const d=diaISO(r[iDt]); if(d) (g[d]=g[d]||[]).push(r); }); if(Object.keys(g).length){ dias={}; Object.entries(g).forEach(([d,l])=>dias[d]=somar(l)); } }
    return {tipo, linhas: somar(body), dias};
  }
  return null;
}
async function tabelasDe(file){
  if(/\.csv$/i.test(file.name)){
    const txt=await file.text(); const l0=txt.split("\n")[0]; const sc=(l0.match(/;/g)||[]).length>(l0.match(/,/g)||[]).length?";":(l0.includes("\t")?"\t":",");
    return [txt.split(/\r?\n/).filter(x=>x.trim()).map(x=>x.split(sc).map(c=>c.replace(/^"|"$/g,"").trim()))];
  }
  const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true});
  return wb.SheetNames.map(sn=>XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:""}));
}
const hm = d => d ? new Date(d).toLocaleString("pt-BR",{day:"2-digit",month:"2-digit",hour:"2-digit",minute:"2-digit"}) : "";
const desfazer = id => `<span><button class="linkbtn" type="button" id="${id}">Desfazer</button></span>`;

// --- movimentação
let movUndo=null;
async function lerMov(file){
  if(!file) return;
  const st=$("movStatus"); st.innerHTML=`<div class="prev">Lendo ${esc(file.name)}…</div>`;
  try{
    let m=null, outro=false;
    for(const rows of await tabelasDe(file)){ m = m || lerMovimentacao(rows); if(!m && lerProdutividade(rows)) outro=true; }
    if(!m){ st.innerHTML=`<div class="prev err">${outro?"Esse arquivo é de separação ou conferência: use o campo correspondente.":"Não encontrei as colunas DESCRIÇÃO, USUÁRIO e PESO do relatório de Movimentação de Estoque."}</div>`; $("movFile").value=""; return; }
    movUndo={mov};
    mov={...m, fonte:file.name}; S[turno].mov=mov;
    const outros=Object.entries(m.outros);
    st.innerHTML=`<div class="prev"><span class="ok-msg">Importado: ${fmt(m.n)} movimentações de separação, de ${hm(m.ini)} a ${hm(m.fim)}. Picking ${t(m.pick)} t · onda ${t(m.onda)} t · semi-onda ${t(m.semi)} t.</span>${outros.length?`<span class="sub">Fora da separação: ${outros.map(([k,v])=>`${k} ${t(v)} t`).join(", ")}.</span>`:""}${desfazer("movDesf")}</div>`;
    if(!sep.length) sepView="mov";
    render();
  }catch(e){ st.innerHTML=`<div class="prev err">Não consegui ler esse arquivo. Confira se é .xlsx ou .csv e se não está protegido por senha.</div>`; }
  $("movFile").value="";
}
$("movStatus").addEventListener("click",e=>{ if(e.target.id!=="movDesf"||!movUndo) return; mov=movUndo.mov; S[turno].mov=mov; movUndo=null; $("movStatus").innerHTML=`<div class="prev"><span class="sub">Importação desfeita.</span></div>`; render(); });
$("movFile").addEventListener("change",e=>lerMov(e.target.files[0]));

// --- Relatório de Separação e Conferência: campos separados.
// Os relatórios diários do WMS trazem o mês inteiro (uma linha por pessoa e por dia, coluna DATA);
// o site usa só o dia escolhido no topo. Se esse dia não estiver no arquivo, usa o mais recente e avisa.
const diaISO = v => {
  if(v instanceof Date) return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,"0")}-${String(v.getDate()).padStart(2,"0")}`;
  const m=String(v??"").trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/); if(!m) return null;
  const y=m[3].length===2?"20"+m[3]:m[3]; return `${y}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;
};
const diaBR = iso => { const [y,m,d]=iso.split("-"); return `${d}/${m}`; };
const imp = { sep:{fonte:"",undo:null,pend:null}, conf:{fonte:"",undo:null,pend:null} };
const ROT = { sep:"Relatório de Separação", conf:"Conferência" };
const IDS = { sep:{st:"sepStatus",file:"sepFile",drop:"sepDrop"}, conf:{st:"confStatus",file:"confFile",drop:"confDrop"} };
function aplicarRel(tipo, linhas, fonte){
  if(tipo==="sep"){ sep=linhas; S[turno].sep=sep; S[turno].fonte=fonte; sepFonte=fonte; aprenderLogins(linhas); }
  else { conf=linhas; S[turno].conf=conf; S[turno].fonteConf=fonte; confFonte=fonte; aprenderLogins(linhas.map(r=>[r[0],0,0,0,0,r[4]])); }
}
function statusRel(tipo){
  const p=imp[tipo].pend, st=$(IDS[tipo].st); if(!p) return;
  const n=p.linhas.length, quem = tipo==="sep"?"colaboradores":"conferentes";
  const sel = p.dias ? `<label class="dia-sel"><span class="sub">Dia usado</span><select data-tipo="${tipo}">${Object.keys(p.dias).sort().reverse().map(d=>`<option value="${d}" ${d===p.dia?"selected":""}>${diaBR(d)}</option>`).join("")}</select></label>` : "";
  const aviso = p.dias && p.dia!==dataISO() ? `<span class="sub" style="color:var(--warn)">O arquivo não tem o dia ${diaBR(dataISO())}; usei ${diaBR(p.dia)}.</span>` : "";
  st.innerHTML=`<div class="prev"><span class="ok-msg">Importado: ${n} ${quem}${p.dias?` em ${diaBR(p.dia)} (o arquivo tem ${Object.keys(p.dias).length} dias)`:""}. Painel do ${turno}º turno atualizado.</span>${aviso}${sel}<span><button class="linkbtn" type="button" data-desf="${tipo}">Desfazer</button></span></div>`;
}
async function importarRel(tipo, file){
  if(!file) return;
  const ids=IDS[tipo], st=$(ids.st); st.innerHTML=`<div class="prev">Lendo ${esc(file.name)}…</div>`;
  try{
    const achados=[]; let outro=null;
    for(const rows of await tabelasDe(file)){
      const r=lerProdutividade(rows);
      if(r && r.tipo===tipo) achados.push(r); else if(r) outro=r.tipo; else if(lerMovimentacao(rows)) outro="mov";
    }
    if(!achados.length){
      const dica = outro==="mov" ? "Esse arquivo é a Movimentação de Estoque." : outro ? `Esse arquivo é de ${outro==="sep"?"separação: use o campo Relatório de Separação":"conferência: use o campo Conferência"}.` :
        (tipo==="sep" ? "Não encontrei o relatório de separação. O cabeçalho precisa ter NOME e PESO_PICKING / PESO_ONDA." : "Não encontrei o relatório de conferência. O cabeçalho precisa ter NOME e PESO_TOTAL.");
      st.innerHTML=`<div class="prev err">${dica}</div>`; $(ids.file).value=""; return;
    }
    // aba com dados por dia (a base bruta) tem prioridade sobre tabelas dinâmicas
    const r = achados.find(a=>a.dias) || achados.sort((a,b)=>b.linhas.length-a.linhas.length)[0];
    imp[tipo].undo = { linhas:(tipo==="sep"?sep:conf).map(x=>[...x]), fonte:imp[tipo].fonte };
    const dia = r.dias ? (r.dias[dataISO()] ? dataISO() : Object.keys(r.dias).sort().pop()) : null;
    imp[tipo].pend = { dias:r.dias, dia, linhas: r.dias ? r.dias[dia] : r.linhas, arquivo:file.name };
    imp[tipo].fonte = file.name + (dia ? ` · ${diaBR(dia)}` : "");
    aplicarRel(tipo, imp[tipo].pend.linhas, imp[tipo].fonte);
    sepView=tipo; statusRel(tipo); render();
  }catch(e){ st.innerHTML=`<div class="prev err">Não consegui ler esse arquivo. Confira se é .xlsx ou .csv e se não está protegido por senha.</div>`; }
  $(ids.file).value="";
}
["sep","conf"].forEach(tipo=>{
  const ids=IDS[tipo];
  $(ids.file).addEventListener("change",e=>importarRel(tipo,e.target.files[0]));
  $(ids.st).addEventListener("change",e=>{ const s=e.target.closest("select[data-tipo]"); const p=imp[tipo].pend; if(!s||!p||!p.dias) return;
    p.dia=s.value; p.linhas=p.dias[p.dia]; imp[tipo].fonte=`${p.arquivo} · ${diaBR(p.dia)}`; aplicarRel(tipo,p.linhas,imp[tipo].fonte); statusRel(tipo); render(); });
  $(ids.st).addEventListener("click",e=>{ if(e.target.dataset.desf!==tipo || !imp[tipo].undo) return; const u=imp[tipo].undo;
    aplicarRel(tipo,u.linhas,u.fonte); imp[tipo].fonte=u.fonte; imp[tipo].undo=null; imp[tipo].pend=null; $(ids.st).innerHTML=`<div class="prev"><span class="sub">Importação desfeita.</span></div>`; render(); });
});
const lerSep = f => importarRel("sep", f), lerConf = f => importarRel("conf", f);
[["sepDrop",lerSep],["confDrop",lerConf],["movDrop",lerMov]].forEach(([id,fn])=>{ const dz=$(id); ["dragenter","dragover"].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.add("over");})); ["dragleave","drop"].forEach(ev=>dz.addEventListener(ev,e=>{e.preventDefault();dz.classList.remove("over");})); dz.addEventListener("drop",e=>fn(e.dataTransfer.files[0])); });

// --- quem é auxiliar (separação de picking) e quem é operador (onda e semi-onda)
function grupoSep(p){
  const f=semAc(p.func||"");
  if(/EMPILHADEIRA|OPERADOR/.test(f)) return "op";
  if(/AUXILIAR/.test(f)) return "aux";
  return (p.onda+p.semi) > p.pick ? "op" : "aux";
}
// Pessoas da separação. Com a Movimentação de Estoque, os pesos de cada um vêm dela
// (a mesma fonte dos totais do VD, então a soma bate); pedidos e itens vêm do relatório
// de produtividade, pelo login. Sem movimentação, vale só o relatório de produtividade.
function pessoasSep(){
  let ps;
  if(mov){
    const porUsr={}; sep.forEach(r=>{ if(r[5]) porUsr[semAc(r[5])]=r; });
    ps = Object.entries(mov.usuarios||{}).map(([u,x])=>{ const r=porUsr[u]; return {nome:r?r[0]:nomeDoLogin(u), usr:u, ped:r?r[1]:null, it:r?r[2]:x.n, mv:x.n, pick:x.pick, onda:x.onda, semi:x.semi, temProd:!!r}; });
  } else ps = sep.map(r=>({nome:r[0], usr:r[5]||"", ped:r[1], it:r[2], pick:r[3], onda:r[4], semi:semiDe(r), temProd:true}));
  ps.forEach(p=>{ p.tot=p.pick+p.onda+p.semi; p.func=funcWMS(p.nome,p.usr); p.grupo=grupoSep(p); });
  return ps;
}

// --- zerar (vale para o relatório que está na tela)
const NOMEV={sep:"separação",conf:"conferência",mov:"movimentação"};
$("sepZerar").addEventListener("click",()=>{
  const zb=$("sepZerar"), nome = NOMEV[sepView];
  if(!zb.classList.contains("sure")){ zb.classList.add("sure"); zb.textContent=`Confirmar: zerar ${nome} do ${turno}º turno`; return; }
  if(sepView==="mov"){ movUndo={mov}; mov=null; S[turno].mov=null; $("movStatus").innerHTML=`<div class="prev"><span class="ok-msg">Movimentação do ${turno}º turno zerada.</span>${desfazer("movDesf")}</div>`; }
  else {
    const tipo=sepView; imp[tipo].undo={ linhas:(tipo==="sep"?sep:conf).map(x=>[...x]), fonte:imp[tipo].fonte }; imp[tipo].pend=null;
    aplicarRel(tipo, [], ""); imp[tipo].fonte="";
    $(IDS[tipo].st).innerHTML=`<div class="prev"><span class="ok-msg">${nome[0].toUpperCase()+nome.slice(1)} do ${turno}º turno zerada. Painel atualizado.</span><span><button class="linkbtn" type="button" data-desf="${tipo}">Desfazer</button></span></div>`;
  }
  render();
});
document.addEventListener("click",e=>{ const zb=$("sepZerar"); if(zb.classList.contains("sure") && e.target!==zb){ zb.classList.remove("sure"); zb.textContent = `Zerar ${NOMEV[sepView]}`; } });
$("sepView").addEventListener("click",e=>{ const b=e.target.closest("[data-v]"); if(!b) return; sepView=b.dataset.v; sepSort={k:sepView==="conf"?"kg":"tot",d:-1}; renderSep(); });
$("sepBusca").addEventListener("input",renderSep);
$("sepTbl").addEventListener("click",e=>{ const b=e.target.closest("button[data-k]"); if(!b) return; const k=b.dataset.k; sepSort = sepSort.k===k ? {k,d:-sepSort.d} : {k,d:(k==="nome"||k==="func"||k==="grupo")?1:-1}; renderSep(); });

function renderSep(){
  $("sepTitle").textContent=`Relatórios do ${turno}º turno`;
  document.querySelectorAll("#sepView button").forEach(b=>b.setAttribute("aria-pressed",b.dataset.v===sepView));
  const fonte = sepView==="mov" ? (mov&&mov.fonte) : (sepView==="conf" ? confFonte : sepFonte);
  $("sepFonteTxt").textContent = fonte ? `Fonte: ${fonte}` : "";
  const zb=$("sepZerar"); zb.classList.remove("sure"); zb.textContent = `Zerar ${NOMEV[sepView]}`; zb.disabled = !(sepView==="mov" ? mov : (sepView==="sep"?sep.length:conf.length));
  const isConf = sepView==="conf";
  let base;
  if(sepView==="mov") base = mov ? Object.entries(mov.usuarios).map(([u,x])=>({nome:nomeDoLogin(u),usr:u,ped:null,it:x.n,pick:x.pick,onda:x.onda,semi:x.semi,tot:x.pick+x.onda+x.semi})) : [];
  else if(sepView==="sep") base = sep.map(r=>({nome:r[0],usr:r[5]||"",ped:r[1],it:r[2],pick:r[3],onda:r[4],semi:semiDe(r),tot:totSep(r)}));
  else base = conf.map(r=>({nome:r[0],usr:r[4]||"",ped:r[1],it:r[2],kg:r[3]}));
  base.forEach(r=>{ r.func=funcWMS(r.nome, r.usr); r.ipp = r.ped? r.it/r.ped : 0; if(!isConf) r.grupo = grupoSep(r)==="op" ? "Operador" : "Auxiliar"; });
  const totKg = base.reduce((a,r)=>a+(isConf?r.kg:r.tot),0), totPed=base.some(r=>r.ped==null)?null:base.reduce((a,r)=>a+r.ped,0), totIt=base.reduce((a,r)=>a+r.it,0);
  const soma = k => base.reduce((a,r)=>a+(r[k]||0),0);
  const pc = x => totKg? fmt(x/totKg*100,0)+"%" : "—";
  const nAux = base.filter(r=>r.grupo==="Auxiliar").length, nOp = base.filter(r=>r.grupo==="Operador").length;
  $("sepKpi").innerHTML = (!isConf ? [
    ["Peso separado",t(totKg),"t"],["Picking",t(soma("pick")),`t · ${pc(soma("pick"))}`],["Onda",t(soma("onda")),`t · ${pc(soma("onda"))}`],["Semi-onda",t(soma("semi")),`t · ${pc(soma("semi"))}`],
    [sepView==="mov"?"Movimentações":"Itens",fmt(totIt),""],["Pessoas",fmt(base.length),`${nAux} aux. · ${nOp} op.`]
  ] : [
    ["Peso conferido",t(totKg),"t"],["Pedidos",fmt(totPed),""],["Itens",fmt(totIt),""],
    ["Conferentes",fmt(base.length),""],["Média por conferente",base.length?t(totKg/base.length):"—","t"],["Itens por pedido",totPed?fmt(totIt/totPed,1):"—",""]
  ]).map(([l,v,u])=>`<div><span class="lbl">${l}</span><span class="v">${v}${u?`<small>${u}</small>`:""}</span></div>`).join("");

  const q=semAc($("sepBusca").value||"");
  let rows=base.filter(r=>!q || semAc(r.nome+" "+r.usr+" "+r.func+" "+(r.grupo||"")).includes(q));
  const k=sepSort.k; rows.sort((a,b)=> (typeof a[k]==="string" ? a[k].localeCompare(b[k],"pt-BR") : ((a[k]??-1)-(b[k]??-1))) * sepSort.d);
  const cols = isConf
    ? [["nome","Conferente","l"],["func","Função","l"],["ped","Pedidos"],["it","Itens"],["ipp","Itens/pedido"],["kg","Peso (kg)"],["share","% do total"]]
    : [["nome","Colaborador","l"],["grupo","Grupo","l"],["func","Função","l"],...(sepView==="sep"?[["ped","Pedidos"]]:[]),["it",sepView==="mov"?"Movim.":"Itens"],["pick","Picking (kg)"],["onda","Onda (kg)"],["semi","Semi-onda (kg)"],["tot","Total (kg)"],["share","% do total"]];
  const th = c => c[0]==="share" ? `<th>${c[1]}</th>` : `<th class="${c[2]||""}" ${sepSort.k===c[0]?`aria-sort="${sepSort.d>0?"ascending":"descending"}"`:""}><button type="button" data-k="${c[0]}">${c[1]}${sepSort.k===c[0]?(sepSort.d>0?" ▲":" ▼"):""}</button></th>`;
  const kgf = v => v ? `<span class="n">${fmt(v)}</span>` : `<span class="n zero">0</span>`;
  const maxV = Math.max(1,...base.map(r=>isConf?r.kg:r.tot));
  const bar = r => { const w = totKg ? (isConf?r.kg:r.tot)/totKg*100 : 0; return `<div class="share"><span>${fmt(w,1)}%</span><span class="tr">${!isConf?`<i style="width:${r.pick/maxV*100}%;background:var(--brand)"></i><i style="width:${r.onda/maxV*100}%;background:var(--bar2)"></i><i style="width:${r.semi/maxV*100}%;background:var(--faint)"></i>`:`<i style="width:${r.kg/maxV*100}%;background:var(--bar2)"></i>`}</span></div>`; };
  const tr = r => `<tr><td class="l nm"><b>${title(r.nome)}</b><small>${r.usr||""}</small></td>` + (isConf?"":`<td class="l"><span class="pl-badge" style="--c:${r.grupo==="Operador"?"#2f6aa8":"#ff6900"}">${r.grupo}</span></td>`) + `<td class="l">${r.func||'<span class="fn-miss">não está no cadastro</span>'}</td>` +
     (isConf ? `<td><span class="n">${fmt(r.ped)}</span></td><td><span class="n">${fmt(r.it)}</span></td><td><span class="n">${fmt(r.ipp,1)}</span></td><td><span class="n"><b>${fmt(r.kg)}</b></span></td>`
             : `${sepView==="sep"?`<td><span class="n">${fmt(r.ped)}</span></td>`:""}<td><span class="n">${fmt(r.it)}</span></td><td>${kgf(r.pick)}</td><td>${kgf(r.onda)}</td><td>${kgf(r.semi)}</td><td><span class="n"><b>${fmt(r.tot)}</b></span></td>`) + `<td>${bar(r)}</td></tr>`;
  const vazio = {sep:"Nenhum relatório de separação", conf:"Nenhum relatório de conferência", mov:"Nenhuma movimentação de estoque"}[sepView];
  $("sepTbl").innerHTML = base.length
    ? `<thead><tr>${cols.map(th).join("")}</tr></thead><tbody>${rows.map(tr).join("") || `<tr><td class="l" colspan="${cols.length}">Nenhum resultado para a busca.</td></tr>`}</tbody>`
    : `<tbody><tr><td class="l" style="padding:22px 14px">${vazio} no ${turno}º turno. Suba o arquivo no campo acima.</td></tr></tbody>`;
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
     <td class="tl"><input type="text" id="eql_${i}" data-i="${i}" data-c="3" value="${esc(e[3]||"")}" placeholder="—" aria-label="Usuário WMS"></td>
     <td class="tc"><select id="eqt_${i}" data-i="${i}" data-c="2" aria-label="Turno">${["1","2","3",""].map(t=>`<option value="${t}" ${e[2]===t?"selected":""}>${TLBL[t]}</option>`).join("")}</select></td>
     <td style="text-align:right"><button class="linkbtn" type="button" data-del="${i}">Remover</button></td></tr>`).join("")
    : `<tr><td colspan="5" class="sub" style="padding:14px 10px">Nenhum colaborador neste filtro.</td></tr>`;
  $("funcList").innerHTML=[...new Set(equipe.map(e=>e[1]).filter(Boolean))].sort().map(f=>`<option value="${esc(f)}">`).join("");
}
$("eqTurnos").addEventListener("click",e=>{ const b=e.target.closest("[data-f]"); if(b){ eqFiltro=b.dataset.f; renderEquipe(); }});
$("eqBusca").addEventListener("input",renderEquipe);
$("eqBody").addEventListener("change",e=>{
  const i=e.target.dataset.i, c=e.target.dataset.c; if(i==null||c==null) return;
  const old=equipe[+i][0], v=e.target.value.trim();
  if(c==="0"){ if(!v){ e.target.value=old; return; } ausencias.forEach(a=>{ if(a.nome===old) a.nome=v; }); afast.forEach(a=>{ if(a.nome===old) a.nome=v; }); saveAf(); }
  equipe[+i][+c]= c==="3" ? v.toUpperCase() : v; eqChanged();
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
  const iN=Math.max(h.findIndex(c=>/^(nome|colaborador)/.test(c)), h.findIndex(c=>/nome|colaborador/.test(c))), iF=h.findIndex(c=>/fun[cç]|cargo/.test(c)), iT=h.findIndex(c=>/turno/.test(c)), iL=h.findIndex(c=>/usu[aá]rio|login/.test(c));
  const out=[]; const vistos=new Set();
  rows.slice(hi+1).forEach(r=>{ const n=String(r[iN]??"").trim(); if(!n||vistos.has(n.toLowerCase())) return; vistos.add(n.toLowerCase()); out.push([n, iF>=0?String(r[iF]??"").trim():"", iT>=0?normTurno(r[iT]):"", iL>=0?String(r[iL]??"").trim().toUpperCase():""]); });
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
  else if(e.target.id==="impAdd"){ pendImport.forEach(x=>{ const k=equipe.find(y=>y[0].toLowerCase()===x[0].toLowerCase()); if(k){ if(x[1]) k[1]=x[1]; if(x[2]) k[2]=x[2]; if(x[3]) k[3]=x[3]; } else equipe.push([...x]); }); }
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
            const wb=XLSX.read(await file.arrayBuffer(),{type:"array",cellDates:true});
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
  const ws=XLSX.utils.aoa_to_sheet([["Nome","Função","Turno","Usuário WMS"],...linhas]); ws["!cols"]=[{wch:36},{wch:30},{wch:8},{wch:14}];
  const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Colaboradores");
  return new Blob([XLSX.write(wb,{bookType:"xlsx",type:"array"})],{type:MIME.xlsx});
}
function baixar(nome,linhas){
  baixarArquivo(`${nome}.xlsx`, planilha(linhas));
  const msg=$("eqDl"); msg.className="ok-msg"; msg.textContent=`${nome}.xlsx baixado.`;
}
$("btnModelo").addEventListener("click",()=>baixar("modelo-colaboradores",[["Maria da Silva","Conferente I","1","MSILVA"],["João Souza","Auxiliar de Expedição I","2","JSOUZA"]]));
$("btnExport").addEventListener("click",()=>{ if(!equipe.length){ $("eqDl").className="sub"; $("eqDl").textContent="O cadastro está vazio."; return; } baixar("cadastro-colaboradores",equipe.map(e=>[e[0],e[1],e[2],e[3]||""])); });

// ---------- abas ----------
const tabs=[["t-painel","v-painel"],["t-lanc","v-lanc"],["t-sep","v-sep"],["t-plan","v-plan"],["t-eq","v-eq"]];
function show(id){ tabs.forEach(([t,v])=>{ $(t).setAttribute("aria-selected",t===id); $(v).hidden = t!==id; }); if(id==="t-painel") fitShot(); }
tabs.forEach(([t])=>$(t).addEventListener("click",()=>show(t)));
const h=location.hash.slice(1); if(h==="lancamento") show("t-lanc"); if(h==="relatorios") show("t-sep"); if(h==="equipe") show("t-eq"); if(h==="planejamento") show("t-plan"); if(h==="apoio") abrirApoio();

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
