"use strict";

/* ===== 初期データ（経費管理表_申告用.xlsx から引き継ぎ） ===== */
const SEED = [
["2025/09/10","消耗品費","MARUTO ネイルニッパー",4780,"その他","",100,"施術用器具"],
["2025/10/25","消耗品費","エプロン（撥水加工）",1385,"その他","",100,"ワークウェア"],
["2025/11/14","消耗品費","施術用ユニフォーム（上下セット等）",17325,"その他","",100,"クラシコで購入"],
["2025/11/25","消耗品費","ネイルマシン（ウラワG3）",55000,"その他","",100,"本格始動に不可欠な高額備品"],
["2025/11/27","消耗品費","フットゾンデ（エキスカベーター）",3980,"その他","",100,"施術用器具"],
["2025/12/02","器具備品","フットケア用フレーザー（4点）",15675,"銀行振込","",100,"パン産商で購入。精密器具。"],
["2025/12/10","消耗品費","超音波洗浄器（シチズン）",7164,"その他","",100,"衛生管理用"],
["2025/12/10","消耗品費","爪やすり（ニグロ）",1310,"その他","",100,"施術用器具"],
["2025/12/15","消耗品費","ネイル集塵機（コードレス）",4680,"その他","",100,"衛生管理用"],
["2025/12/15","消耗品費","爪やすり（ニグロ）×2",2400,"その他","",100,"施術用器具"],
["2025/12/15","消耗品費","かかとヤスリ（グリーンベル）",565,"その他","",100,"施術用器具"],
["2026/01/01","消耗品費","足湯用たらい（野田琺瑯）",4290,"その他","",100,"フットバス用備品"],
["2026/01/02","交際接待費","食事代（情報交換）",5709,"現金","",100,"パートナーと仕事内容の情報交換"],
["2026/01/04","研修費","巻き爪ケアディプロマセミナー",11880,"その他","",100,"技術習得のための講習代"],
["2026/01/04","消耗品費","フラットケアジェルセット",11000,"その他","",100,"施術用消耗品"],
["2026/01/10","交際接待費","食事代（打ち合わせ）",1850,"現金","",100,"事業上の打ち合わせ"],
["2026/01/11","消耗品費","サロン用デスク（IKEA）",19990,"クレジットカード","",100,"IKEA神戸で購入。事務・カウンセリング用。"],
["2026/01/11","交際接待費","食事代（情報交換）",2530,"現金","",100,"パートナーと仕事内容の情報交換"],
["2026/01/12","消耗品費","リクライニングチェア（フットレスト一体型）",61020,"その他","",100,"サロンのメイン家具（100%事業用）"],
["2026/01/12","交際接待費","食事代（情報交換）",3760,"現金","",100,"パートナーと仕事内容の情報交換"],
["2026/01/15","消耗品費","UV/LEDジェルネイルライト",2980,"その他","",100,"施術用器具"],
["2026/01/24","交際接待費","食事代（作戦会議）",7500,"現金","",100,"パートナーと仕事の見通しについて作戦会議"],
["2026/01/25","交際接待費","食事代（作戦会議）",2030,"現金","",100,"パートナーと仕事の見通しについて作戦会議"],
["2026/01/25","交際接待費","円城白菜（引越し挨拶品）",200,"楽天ペイ","",100,"階下への引越し挨拶（近隣対策）"],
["2026/02/08","消耗品費","キャッシュトレー",2970,"クレジットカード","",100,"FRANK 暮らしの道具で購入。受付用。"],
["2026/02/21","交際接待費","食事代（すり合わせ）",4580,"現金","",100,"パートナーと仕事のすり合わせ"],
["2026/02/22","交際接待費","食事代（すり合わせ）",6017,"現金","",100,"パートナーと仕事のすり合わせ"],
["2026/02/24","消耗品費","サロン用生花代",434,"現金","",100,"サロン空間の装飾"],
["2026/03/04","消耗品費","ダイヤモンドビット・各種一式",19556,"その他","",100,"WSPTジャパンで購入。施術用器具。"],
["2026/03/05","消耗品費","LCN アンティセプト（爪美容液）",4780,"その他","",100,"ケア用品"],
["2026/03/10","消耗品費","キャスター付きスチールワゴン",13200,"その他","",100,"収納家具"],
["2026/03/10","消耗品費","今治5重ガーゼケット",4266,"その他","",100,"お客様用リネン"],
["2026/03/14","消耗品費","フットケア施術者用チェア",2580,"クレジットカード","",100,"ZOZOTOWNで購入"],
["2026/03/17","研修費","靴とインソールの研究室 講習代",1500,"クレジットカード","",100,"フットケア業務の知識習得"],
["2026/03/28","広告宣伝費","PR撮影用寝具一式",17758,"クレジットカード","",60,"PR撮影演出用（副業・SNS関連）按分60%"],
["2026/03/28","消耗品費","サロン用延長コード",5420,"クレジットカード","",100,"機器・照明等の電源用"],
["2026/03/28","消耗品費","経費整理用ファイル",1610,"クレジットカード","",100,"領収書等の事務管理用"],
["2026/03/28","消耗品費","足拭き取りタオル",398,"現金","",100,"施術用消耗品"],
["2026/03/28","交際接待費","食事代（作戦会議）",3660,"現金","",100,"パートナーと仕事の話（作戦会議）"]
];

const DEFAULT_CATS = ["消耗品費","器具備品","研修費","交際接待費","広告宣伝費","旅費交通費","通信費","地代家賃","水道光熱費","支払手数料","新聞図書費","雑費"];

const DEFAULT_RULES = [
["セミナー","研修費"],["講習","研修費"],["研修","研修費"],["講座","研修費"],
["食事","交際接待費"],["会食","交際接待費"],["手土産","交際接待費"],["挨拶","交際接待費"],
["広告","広告宣伝費"],["PR","広告宣伝費"],["撮影","広告宣伝費"],["チラシ","広告宣伝費"],["名刺","広告宣伝費"],
["電車","旅費交通費"],["バス","旅費交通費"],["タクシー","旅費交通費"],["交通","旅費交通費"],["駐車","旅費交通費"],["ガソリン","旅費交通費"],
["通信","通信費"],["スマホ","通信費"],["携帯","通信費"],["Wi-Fi","通信費"],["サーバー","通信費"],["ドメイン","通信費"],
["家賃","地代家賃"],["電気","水道光熱費"],["水道","水道光熱費"],["ガス","水道光熱費"],
["振込手数料","支払手数料"],["手数料","支払手数料"],
["書籍","新聞図書費"],["本","新聞図書費"],["雑誌","新聞図書費"],
["ニッパー","消耗品費"],["やすり","消耗品費"],["ヤスリ","消耗品費"],["タオル","消耗品費"],["ジェル","消耗品費"]
];

/* ===== 状態管理 ===== */
const LS_KEY="keihi-data-v1", LS_RULES="keihi-rules-v1", LS_CATS="keihi-cats-v1", LS_INCOME="keihi-income-v1";
let rows, rules, cats, incomes;

function load(){
  try{
    rows = JSON.parse(localStorage.getItem(LS_KEY)) || null;
    rules = JSON.parse(localStorage.getItem(LS_RULES)) || null;
    cats = JSON.parse(localStorage.getItem(LS_CATS)) || null;
    incomes = JSON.parse(localStorage.getItem(LS_INCOME)) || null;
  }catch(e){ rows=rules=cats=incomes=null; }
  if(!rows) rows = SEED.map(r=>({date:r[0],cat:r[1],name:r[2],amount:r[3],pay:r[4],receipt:r[5],ratio:r[6],memo:r[7]}));
  if(!rules) rules = DEFAULT_RULES.map(r=>({kw:r[0],cat:r[1]}));
  if(!cats) cats = DEFAULT_CATS.slice();
  if(!incomes) incomes = [];
}
function save(){
  try{
    localStorage.setItem(LS_KEY, JSON.stringify(rows));
    localStorage.setItem(LS_RULES, JSON.stringify(rules));
    localStorage.setItem(LS_CATS, JSON.stringify(cats));
    localStorage.setItem(LS_INCOME, JSON.stringify(incomes));
  }catch(e){ console.error("保存に失敗しました", e); }
}

/* ===== ユーティリティ ===== */
const $ = s=>document.querySelector(s);
const yen = n=>"¥"+Math.round(n).toLocaleString("ja-JP");
const keisan = r=>Math.round(r.amount * (r.ratio||100)/100);
function guessCat(name){
  if(!name) return null;
  for(const r of rules){ if(r.kw && name.includes(r.kw)) return r.cat; }
  return null;
}
function esc(s){ return String(s??"").replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c])); }

/* ===== タブ切替 ===== */
document.querySelectorAll("nav button").forEach(b=>{
  b.addEventListener("click",()=>{
    document.querySelectorAll("nav button").forEach(x=>x.classList.remove("on"));
    document.querySelectorAll(".tab").forEach(x=>x.classList.remove("on"));
    b.classList.add("on");
    $("#tab-"+b.dataset.tab).classList.add("on");
    render();
  });
});

/* ===== カテゴリselect構築 ===== */
function buildCatSelects(){
  const opts = cats.map(c=>`<option>${esc(c)}</option>`).join("");
  $("#f-cat").innerHTML = opts;
  $("#filter-cat").innerHTML = `<option value="">すべてのカテゴリ</option>`+opts;
}

/* ===== 入力フォーム ===== */
$("#f-name").addEventListener("input",e=>{
  const g = guessCat(e.target.value);
  if(g && cats.includes(g)){
    $("#f-cat").value = g;
    $("#f-cat").classList.remove("match-flash"); void $("#f-cat").offsetWidth;
    $("#f-cat").classList.add("match-flash");
  }
});
$("#btn-add").addEventListener("click",()=>{
  const d=$("#f-date").value, a=Number($("#f-amount").value), n=$("#f-name").value.trim();
  if(!d || !n || !a || a<=0){ alert("日付・品名・金額（1円以上）を入力してください"); return; }
  rows.push({
    date:d.replace(/-/g,"/"), cat:$("#f-cat").value, name:n, amount:a,
    pay:$("#f-pay").value, receipt:$("#f-receipt").value.trim(),
    ratio:Math.min(100,Math.max(1,Number($("#f-ratio").value)||100)),
    memo:$("#f-memo").value.trim()
  });
  rows.sort((x,y)=>x.date.localeCompare(y.date));
  save(); clearForm(); render();
  alert("追加しました（明細一覧タブで確認できます）");
});
function clearForm(){
  ["f-date","f-amount","f-name","f-receipt","f-memo"].forEach(id=>$("#"+id).value="");
  $("#f-ratio").value=100;
}
$("#btn-clear").addEventListener("click",clearForm);

/* ===== CSV取り込み ===== */
function parseCSV(text){
  const out=[]; let row=[],cell="",q=false;
  for(let i=0;i<text.length;i++){
    const c=text[i];
    if(q){ if(c==='"'){ if(text[i+1]==='"'){cell+='"';i++;} else q=false; } else cell+=c; }
    else if(c==='"') q=true;
    else if(c===','){ row.push(cell); cell=""; }
    else if(c==='\n'||c==='\r'){ if(c==='\r'&&text[i+1]==='\n')i++; row.push(cell); out.push(row); row=[];cell=""; }
    else cell+=c;
  }
  if(cell!==""||row.length){ row.push(cell); out.push(row); }
  return out.filter(r=>r.some(c=>c.trim()!==""));
}
$("#csv-file").addEventListener("change",e=>{
  const f=e.target.files[0]; if(!f) return;
  const reader=new FileReader();
  reader.onload=()=>{
    try{
      const recs=parseCSV(reader.result.replace(/^\uFEFF/,""));
      if(recs.length<2){ alert("データ行が見つかりませんでした"); return; }
      let added=0, guessed=0;
      for(let i=1;i<recs.length;i++){
        const r=recs[i];
        const amount=Number(String(r[3]||"").replace(/[¥,，円\s]/g,""));
        if(!r[0]||!amount) continue;
        let cat=(r[1]||"").trim();
        if(!cat){ cat=guessCat(r[2]||"")||"雑費"; guessed++; }
        if(!cats.includes(cat)) cats.push(cat);
        rows.push({
          date:r[0].trim().replace(/-/g,"/"), cat, name:(r[2]||"").trim(), amount,
          pay:(r[4]||"その他").trim()||"その他", receipt:(r[5]||"").trim(),
          ratio:Math.min(100,Math.max(1,Number(r[6])||100)), memo:(r[7]||"").trim()
        });
        added++;
      }
      rows.sort((x,y)=>x.date.localeCompare(y.date));
      save(); render(); buildCatSelects();
      alert(`${added}件を取り込みました${guessed?`（うち${guessed}件はカテゴリを自動仕分け）`:""}`);
    }catch(err){ alert("読み込みに失敗しました。CSVの形式をご確認ください。"); }
    e.target.value="";
  };
  reader.readAsText(f);
});

/* ===== 収入 ===== */
$("#btn-add-income").addEventListener("click",()=>{
  const d=$("#i-date").value, a=Number($("#i-amount").value);
  if(!d || !a || a<=0){ alert("入金日と金額（1円以上）を入力してください"); return; }
  incomes.push({date:d.replace(/-/g,"/"), type:$("#i-type").value, desc:$("#i-desc").value.trim(), amount:a});
  incomes.sort((x,y)=>x.date.localeCompare(y.date));
  save();
  $("#i-date").value=""; $("#i-amount").value=""; $("#i-desc").value="";
  render();
  alert("収入を追加しました");
});
function renderIncome(){
  const tbody=$("#income-table tbody"); tbody.innerHTML="";
  if(!incomes.length){ tbody.innerHTML=`<tr><td colspan="6" class="empty">収入がまだ登録されていません</td></tr>`; return; }
  let total=0;
  incomes.forEach((r,i)=>{
    total+=r.amount;
    const tr=document.createElement("tr");
    tr.innerHTML=`<td>${i+1}</td><td>${esc(r.date)}</td><td><span class="cat-pill">${esc(r.type)}</span></td>
      <td>${esc(r.desc)}</td><td class="num">${yen(r.amount)}</td>
      <td><button class="del-x" data-i="${i}" aria-label="削除">✕</button></td>`;
    tbody.appendChild(tr);
  });
  const tr=document.createElement("tr"); tr.className="total-row";
  tr.innerHTML=`<td colspan="4">合計（${incomes.length}件）</td><td class="num">${yen(total)}</td><td></td>`;
  tbody.appendChild(tr);
  tbody.querySelectorAll(".del-x").forEach(b=>b.addEventListener("click",()=>{
    const r=incomes[Number(b.dataset.i)];
    if(confirm(`「${r.desc||r.type}」(${yen(r.amount)}) を削除しますか？`)){ incomes.splice(Number(b.dataset.i),1); save(); render(); }
  }));
}

/* ===== 一覧描画 ===== */
function renderList(){
  const filter=$("#filter-cat").value;
  const tbody=$("#list-table tbody"); tbody.innerHTML="";
  const view=rows.map((r,i)=>({r,i})).filter(o=>!filter||o.r.cat===filter);
  if(!view.length){ tbody.innerHTML=`<tr><td colspan="10" class="empty">該当する経費がありません</td></tr>`; return; }
  let totalA=0,totalK=0;
  view.forEach((o,n)=>{
    const r=o.r; totalA+=r.amount; totalK+=keisan(r);
    const tr=document.createElement("tr");
    tr.innerHTML=`<td>${n+1}</td><td>${esc(r.date)}</td><td><span class="cat-pill">${esc(r.cat)}</span></td>
      <td>${esc(r.name)}${r.receipt?`<div class="muted">領収書: ${esc(r.receipt)}</div>`:""}</td>
      <td class="num">${yen(r.amount)}</td><td class="num">${r.ratio}%</td><td class="num">${yen(keisan(r))}</td>
      <td>${esc(r.pay)}</td><td class="muted">${esc(r.memo)}</td>
      <td><button class="del-x" data-i="${o.i}" aria-label="削除">✕</button></td>`;
    tbody.appendChild(tr);
  });
  const tr=document.createElement("tr"); tr.className="total-row";
  tr.innerHTML=`<td colspan="4">合計（${view.length}件）</td><td class="num">${yen(totalA)}</td><td></td><td class="num">${yen(totalK)}</td><td colspan="3"></td>`;
  tbody.appendChild(tr);
  tbody.querySelectorAll(".del-x").forEach(b=>b.addEventListener("click",()=>{
    const r=rows[Number(b.dataset.i)];
    if(confirm(`「${r.name}」(${yen(r.amount)}) を削除しますか？`)){ rows.splice(Number(b.dataset.i),1); save(); render(); }
  }));
}
$("#filter-cat").addEventListener("change",renderList);

/* ===== 集計 ===== */
function aggregate(){
  const byCat={}, byMonth={};
  let totalA=0,totalK=0;
  for(const r of rows){
    const k=keisan(r); totalA+=r.amount; totalK+=k;
    (byCat[r.cat]??=({n:0,a:0,k:0})); byCat[r.cat].n++; byCat[r.cat].a+=r.amount; byCat[r.cat].k+=k;
    const m=r.date.slice(0,7);
    (byMonth[m]??=({n:0,a:0,k:0})); byMonth[m].n++; byMonth[m].a+=r.amount; byMonth[m].k+=k;
  }
  return {byCat,byMonth,totalA,totalK};
}
function renderSummary(){
  const {byCat,byMonth,totalA,totalK}=aggregate();
  const totalIncome = incomes.reduce((s,r)=>s+r.amount,0);
  const shotoku = totalIncome - totalK;
  const shotokuColor = shotoku<0 ? "background:var(--danger-soft)" : "";
  const shotokuVal = shotoku<0 ? "color:var(--danger)" : "";
  $("#stat-cards").innerHTML=`
    <div class="stat"><div class="lbl">収入合計</div><div class="val">${yen(totalIncome)}</div></div>
    <div class="stat"><div class="lbl">経費計上額（按分後）</div><div class="val">${yen(totalK)}</div></div>
    <div class="stat" style="${shotokuColor}"><div class="lbl">差引所得（収入−経費）</div><div class="val" style="${shotokuVal}">${shotoku<0?"−":""}${yen(Math.abs(shotoku))}</div></div>
    <div class="stat"><div class="lbl">登録件数</div><div class="val">経費${rows.length}・収入${incomes.length}</div></div>`;
  const catKeys=Object.keys(byCat).sort((a,b)=>byCat[b].k-byCat[a].k);
  $("#cat-table tbody").innerHTML = catKeys.map(c=>{
    const v=byCat[c];
    return `<tr><td><span class="cat-pill">${esc(c)}</span></td><td class="num">${v.n}</td><td class="num">${yen(v.a)}</td><td class="num">${yen(v.k)}</td><td class="num">${totalK?Math.round(v.k/totalK*100):0}%</td></tr>`;
  }).join("") + `<tr class="total-row"><td>合計</td><td class="num">${rows.length}</td><td class="num">${yen(totalA)}</td><td class="num">${yen(totalK)}</td><td class="num">100%</td></tr>`;
  const mKeys=Object.keys(byMonth).sort();
  $("#month-table tbody").innerHTML = mKeys.map(m=>{
    const v=byMonth[m];
    return `<tr><td>${m.replace("/","年")}月</td><td class="num">${v.n}</td><td class="num">${yen(v.a)}</td><td class="num">${yen(v.k)}</td></tr>`;
  }).join("") + `<tr class="total-row"><td>合計</td><td class="num">${rows.length}</td><td class="num">${yen(totalA)}</td><td class="num">${yen(totalK)}</td></tr>`;
}

/* ===== ルール編集 ===== */
function renderRules(){
  const wrap=$("#kw-list"); wrap.innerHTML="";
  rules.forEach((r,i)=>{
    const div=document.createElement("div"); div.className="kw-row";
    const opts=cats.map(c=>`<option ${c===r.cat?"selected":""}>${esc(c)}</option>`).join("");
    div.innerHTML=`<input type="text" value="${esc(r.kw)}" placeholder="キーワード" data-i="${i}" data-f="kw">
      <select data-i="${i}" data-f="cat">${opts}</select>
      <button class="del-x" data-i="${i}" aria-label="このルールを削除">✕</button>`;
    wrap.appendChild(div);
  });
  wrap.querySelectorAll("input,select").forEach(el=>el.addEventListener("change",()=>{
    rules[Number(el.dataset.i)][el.dataset.f]=el.value; save();
  }));
  wrap.querySelectorAll(".del-x").forEach(b=>b.addEventListener("click",()=>{
    rules.splice(Number(b.dataset.i),1); save(); renderRules();
  }));
  const cl=$("#cat-list"); cl.innerHTML="";
  cats.forEach((c,i)=>{
    const div=document.createElement("div"); div.className="kw-row";
    div.innerHTML=`<input type="text" value="${esc(c)}" data-i="${i}">
      <button class="del-x" data-i="${i}" aria-label="このカテゴリを削除">✕</button>`;
    cl.appendChild(div);
  });
  cl.querySelectorAll("input").forEach(el=>el.addEventListener("change",()=>{
    const old=cats[Number(el.dataset.i)], nv=el.value.trim();
    if(!nv) return;
    cats[Number(el.dataset.i)]=nv;
    rows.forEach(r=>{ if(r.cat===old) r.cat=nv; });
    rules.forEach(r=>{ if(r.cat===old) r.cat=nv; });
    save(); buildCatSelects(); render();
  }));
  cl.querySelectorAll(".del-x").forEach(b=>b.addEventListener("click",()=>{
    const c=cats[Number(b.dataset.i)];
    if(rows.some(r=>r.cat===c)){ alert(`「${c}」は明細で使用中のため削除できません`); return; }
    cats.splice(Number(b.dataset.i),1);
    save(); buildCatSelects(); renderRules();
  }));
}
$("#btn-add-rule").addEventListener("click",()=>{ rules.unshift({kw:"",cat:cats[0]}); save(); renderRules(); });
$("#btn-reset-rules").addEventListener("click",()=>{
  if(confirm("仕分けルールを初期状態に戻しますか？")){ rules=DEFAULT_RULES.map(r=>({kw:r[0],cat:r[1]})); save(); renderRules(); }
});
$("#btn-add-cat").addEventListener("click",()=>{
  const nv=prompt("追加するカテゴリ名（勘定科目）"); if(!nv||!nv.trim()) return;
  if(cats.includes(nv.trim())){ alert("既に存在します"); return; }
  cats.push(nv.trim()); save(); buildCatSelects(); renderRules();
});

/* ===== CSV出力 ===== */
$("#btn-csv").addEventListener("click",()=>{
  const head=["日付","カテゴリ","品名・内容","金額","支払方法","領収書No","按分率","備考","経費計上額"];
  const q=v=>{ v=String(v??""); return /[",\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v; };
  const lines=[head.join(",")];
  rows.forEach(r=>lines.push([r.date,r.cat,r.name,r.amount,r.pay,r.receipt,r.ratio,r.memo,keisan(r)].map(q).join(",")));
  const blob=new Blob(["\uFEFF"+lines.join("\r\n")],{type:"text/csv;charset=utf-8"});
  const a=document.createElement("a");
  a.href=URL.createObjectURL(blob);
  a.download="経費明細_"+new Date().toISOString().slice(0,10)+".csv";
  a.click(); URL.revokeObjectURL(a.href);
});

$("#btn-csv-income").addEventListener("click",()=>{
  const q=v=>{ v=String(v??""); return /[",\n]/.test(v)?'"'+v.replace(/"/g,'""')+'"':v; };
  const lines=["入金日,区分,摘要,金額"];
  incomes.forEach(r=>lines.push([r.date,r.type,r.desc,r.amount].map(q).join(",")));
  const blob=new Blob(["\uFEFF"+lines.join("\r\n")],{type:"text/csv;charset=utf-8"});
  const a=document.createElement("a");
  a.href=URL.createObjectURL(blob);
  a.download="収入明細_"+new Date().toISOString().slice(0,10)+".csv";
  a.click(); URL.revokeObjectURL(a.href);
});

/* ===== PDF（印刷）出力 ===== */
$("#btn-pdf").addEventListener("click",()=>{
  const {byCat,totalA,totalK}=aggregate();
  const catKeys=Object.keys(byCat).sort((a,b)=>byCat[b].k-byCat[a].k);
  const dates=rows.map(r=>r.date).concat(incomes.map(r=>r.date)).sort();
  const totalIncome=incomes.reduce((s,r)=>s+r.amount,0);
  const shotoku=totalIncome-totalK;
  let html=`<h1>収支明細書（事業所得・雑所得 申告用）</h1>
  <p>対象期間：${dates[0]||"—"} 〜 ${dates[dates.length-1]||"—"}　／　作成日：${new Date().toLocaleDateString("ja-JP")}</p>
  <h2>収支サマリー</h2>
  <table><tr><th>収入合計</th><th>経費計上額（按分後）</th><th>差引所得</th></tr>
  <tr><td style="text-align:right">${yen(totalIncome)}</td><td style="text-align:right">${yen(totalK)}</td><td style="text-align:right"><b>${shotoku<0?"−":""}${yen(Math.abs(shotoku))}</b></td></tr></table>`;
  if(incomes.length){
    html+=`<h2>収入明細</h2><table><tr><th>No.</th><th>入金日</th><th>区分</th><th>摘要</th><th>金額</th></tr>`;
    incomes.forEach((r,i)=>{ html+=`<tr><td>${i+1}</td><td>${esc(r.date)}</td><td>${esc(r.type)}</td><td>${esc(r.desc)}</td><td style="text-align:right">${yen(r.amount)}</td></tr>`; });
    html+=`<tr><td colspan="4"><b>合計</b></td><td style="text-align:right"><b>${yen(totalIncome)}</b></td></tr></table>`;
  }
  html+=`<h2>カテゴリ別集計（経費）</h2>
  <table><tr><th>カテゴリ（勘定科目）</th><th>件数</th><th>支払総額</th><th>経費計上額（按分後）</th></tr>`;
  catKeys.forEach(c=>{ const v=byCat[c];
    html+=`<tr><td>${esc(c)}</td><td style="text-align:right">${v.n}</td><td style="text-align:right">${yen(v.a)}</td><td style="text-align:right">${yen(v.k)}</td></tr>`; });
  html+=`<tr><td><b>合計</b></td><td style="text-align:right"><b>${rows.length}</b></td><td style="text-align:right"><b>${yen(totalA)}</b></td><td style="text-align:right"><b>${yen(totalK)}</b></td></tr></table>
  <h2>経費明細一覧</h2>
  <table><tr><th>No.</th><th>日付</th><th>カテゴリ</th><th>品名・内容</th><th>金額</th><th>按分</th><th>計上額</th><th>支払方法</th><th>備考</th></tr>`;
  rows.forEach((r,i)=>{
    html+=`<tr><td>${i+1}</td><td>${esc(r.date)}</td><td>${esc(r.cat)}</td><td>${esc(r.name)}</td>
    <td style="text-align:right">${yen(r.amount)}</td><td style="text-align:right">${r.ratio}%</td>
    <td style="text-align:right">${yen(keisan(r))}</td><td>${esc(r.pay)}</td><td>${esc(r.memo)}</td></tr>`;
  });
  html+=`</table><p style="margin-top:4mm">※ 経費計上額は按分率適用後の金額。領収書等の原本は別途保管。</p>`;
  $("#print-area").innerHTML=html;
  window.print();
});

/* ===== 初期化 ===== */
$("#btn-wipe").addEventListener("click",()=>{
  if(confirm("経費・収入の全データを削除します。よろしいですか？")&&confirm("本当に削除しますか？（この操作は取り消せません）")){
    rows=[]; incomes=[]; save(); render();
  }
});

function render(){ renderList(); renderSummary(); renderRules(); renderIncome(); }
load(); buildCatSelects(); render();
