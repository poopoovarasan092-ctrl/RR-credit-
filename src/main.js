import * as XLSX from 'xlsx';
import './style.css';

const DB='rr-credit-v1', STORE='records';
let db;
const openDB=()=>new Promise((resolve,reject)=>{
 const req=indexedDB.open(DB,1);
 req.onupgradeneeded=()=>req.result.createObjectStore(STORE,{keyPath:'id'});
 req.onsuccess=()=>resolve(req.result); req.onerror=()=>reject(req.error);
});
const all=()=>new Promise((resolve,reject)=>{const r=db.transaction(STORE).objectStore(STORE).getAll();r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error)});
const put=x=>new Promise((resolve,reject)=>{const r=db.transaction(STORE,'readwrite').objectStore(STORE).put(x);r.onsuccess=()=>resolve();r.onerror=()=>reject(r.error)});
const del=id=>new Promise((resolve,reject)=>{const r=db.transaction(STORE,'readwrite').objectStore(STORE).delete(id);r.onsuccess=()=>resolve();r.onerror=()=>reject(r.error)});
const id=()=>crypto.randomUUID();
const money=n=>'₹'+Number(n||0).toLocaleString('en-IN',{maximumFractionDigits:2});
let records=[],tab='dashboard';
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const customers=()=>records.filter(x=>x.type==='customer');
const tx=()=>records.filter(x=>x.type==='transaction');
const savings=()=>records.filter(x=>x.type==='saving');
const balance=cid=>tx().filter(x=>x.customerId===cid).reduce((a,x)=>a+(x.kind==='credit'?+x.amount:-x.amount),0);
const totalDue=()=>customers().reduce((a,c)=>a+Math.max(0,balance(c.id)),0);
const fmtDate=d=>new Date(d).toLocaleDateString('en-IN');
function shell(){
 document.querySelector('#app').innerHTML=`
 <header><div class="brand"><div class="logo">RR</div><div><b>RR Credit</b><small>RR Mobiles · Offline ledger</small></div></div><span class="offline"><i></i> Offline ready</span></header>
 <nav>${[['dashboard','⌂','Home'],['customers','♙','Customers'],['transactions','⇄','Entries'],['savings','▣','Savings'],['reports','▤','Reports']].map(([k,ic,t])=>`<button data-tab="${k}" class="${tab===k?'active':''}"><span>${ic}</span>${t}</button>`).join('')}</nav>
 <main id="view"></main><footer>Data stays on this phone. Export a backup regularly.</footer>`;
 document.querySelectorAll('[data-tab]').forEach(b=>b.onclick=()=>{tab=b.dataset.tab;render()});
}
function render(){
 shell(); const v=document.querySelector('#view');
 if(tab==='dashboard') v.innerHTML=`
 <div class="welcome"><div><p class="eyebrow">YOUR SHOP LEDGER</p><h1>Good day 👋</h1><p>Keep customer accounts and savings in one place.</p></div><button class="primary" id="quick">＋ Add entry</button></div>
 <section class="stats"><article><small>Total outstanding</small><strong>${money(totalDue())}</strong></article><article><small>Customers</small><strong>${customers().length}</strong></article><article><small>Total savings</small><strong>${money(savings().reduce((a,x)=>a+ +x.amount,0))}</strong></article></section>
 <section class="panel"><div class="panelhead"><h2>Recent entries</h2><button class="link" id="alltx">View all</button></div>${tx().sort((a,b)=>b.date.localeCompare(a.date)).slice(0,6).map(entryRow).join('')||'<p class="empty">No entries yet. Add your first customer or credit entry.</p>'}</section>`;
 if(tab==='customers') v.innerHTML=`<div class="pagehead"><div><p class="eyebrow">CUSTOMER ACCOUNTS</p><h1>Customers</h1></div><button class="primary" id="addcust">＋ Add customer</button></div><div class="panel">${customers().map(c=>`<div class="customer"><div class="avatar">${esc(c.name[0]?.toUpperCase())}</div><div class="grow"><b>${esc(c.name)}</b><small>${esc(c.phone||'No phone number')}</small></div><strong>${money(balance(c.id))}</strong><button class="small" data-ledger="${c.id}">Ledger</button></div>`).join('')||'<p class="empty">No customers added yet.</p>'}</div>`;
 if(tab==='transactions') v.innerHTML=`<div class="pagehead"><div><p class="eyebrow">CREDIT & COLLECTIONS</p><h1>Entries</h1></div><button class="primary" id="addtx">＋ Add entry</button></div><div class="panel">${tx().sort((a,b)=>b.date.localeCompare(a.date)).map(entryRow).join('')||'<p class="empty">No transactions yet.</p>'}</div>`;
 if(tab==='savings') v.innerHTML=`<div class="pagehead"><div><p class="eyebrow">PERSONAL / SHOP SAVINGS</p><h1>Savings</h1></div><button class="primary" id="addsave">＋ Add saving</button></div><div class="stats"><article><small>Total recorded savings</small><strong>${money(savings().reduce((a,x)=>a+ +x.amount,0))}</strong></article><article><small>Records</small><strong>${savings().length}</strong></article></div><div class="panel">${savings().sort((a,b)=>b.date.localeCompare(a.date)).map(x=>`<div class="entry"><div><b>${esc(x.note||'Savings deposit')}</b><small>${fmtDate(x.date)}</small></div><strong class="green">+ ${money(x.amount)}</strong></div>`).join('')||'<p class="empty">No savings recorded yet.</p>'}</div>`;
 if(tab==='reports') v.innerHTML=`<div class="pagehead"><div><p class="eyebrow">YOUR BUSINESS RECORDS</p><h1>Reports & backup</h1></div></div><div class="panel"><h2>Excel report</h2><p>Export a workbook containing Dashboard, Customers, Transactions, Collections and Savings sheets.</p><button class="primary" id="excel">⇩ Download Excel report</button><hr><h2>Local backup</h2><p>Save all RR Credit records to a JSON file. Keep it somewhere safe.</p><button class="secondary" id="backup">Download data backup</button><label class="filebtn">Restore from backup<input id="restore" type="file" accept=".json,application/json"></label><p class="muted">Google Drive automatic backup needs Google sign-in and permission. It is not connected in this first build.</p></div>`;
 bind();
}
function entryRow(x){const c=customers().find(c=>c.id===x.customerId);return `<div class="entry"><div><b>${esc(c?.name||'Customer')}</b><small>${x.kind==='credit'?'Credit given':'Payment received'} · ${fmtDate(x.date)} ${x.note?'· '+esc(x.note):''}</small></div><strong class="${x.kind==='credit'?'red':'green'}">${x.kind==='credit'?'+':'−'} ${money(x.amount)}</strong></div>`}
function modal(title,body){const m=document.createElement('div');m.className='modalback';m.innerHTML=`<form class="modal"><div class="panelhead"><h2>${title}</h2><button type="button" class="close">×</button></div>${body}<button class="primary full" type="submit">Save</button></form>`;document.body.append(m);m.querySelector('.close').onclick=()=>m.remove();return [m,m.querySelector('form')]}
function bind(){
 document.querySelector('#quick')?.addEventListener('click',()=>{tab='transactions';render();document.querySelector('#addtx')?.click()});
 document.querySelector('#alltx')?.addEventListener('click',()=>{tab='transactions';render()});
 document.querySelector('#addcust')?.addEventListener('click',()=>{const [m,f]=modal('Add customer',`<label>Customer name<input name="name" required></label><label>Mobile number<input name="phone" inputmode="numeric"></label>`);f.onsubmit=async e=>{e.preventDefault();await put({id:id(),type:'customer',...Object.fromEntries(new FormData(f))});m.remove();await refresh() }});
 document.querySelectorAll('[data-ledger]').forEach(b=>b.onclick=()=>{const c=customers().find(x=>x.id===b.dataset.ledger);tab='transactions';render();const add=document.querySelector('#addtx');add?.click();setTimeout(()=>{const s=document.querySelector('[name=customerId]');if(s)s.value=c.id},0)});
 document.querySelector('#addtx')?.addEventListener('click',()=>{if(!customers().length){alert('Add a customer first.');return}const opts=customers().map(c=>`<option value="${c.id}">${esc(c.name)}</option>`).join('');const [m,f]=modal('Add ledger entry',`<label>Customer<select name="customerId">${opts}</select></label><label>Entry type<select name="kind"><option value="credit">Credit given (customer owes)</option><option value="payment">Payment received</option></select></label><label>Amount (₹)<input name="amount" type="number" min="0.01" step="0.01" required></label><label>Date<input name="date" type="date" value="${new Date().toISOString().slice(0,10)}" required></label><label>Note<input name="note" placeholder="e.g. Mobile display repair"></label>`);f.onsubmit=async e=>{e.preventDefault();const d=Object.fromEntries(new FormData(f));await put({id:id(),type:'transaction',...d});m.remove();await refresh()}});
 document.querySelector('#addsave')?.addEventListener('click',()=>{const [m,f]=modal('Record savings',`<label>Amount (₹)<input name="amount" type="number" min="0.01" step="0.01" required></label><label>Date<input name="date" type="date" value="${new Date().toISOString().slice(0,10)}" required></label><label>Note<input name="note" placeholder="Daily savings"></label>`);f.onsubmit=async e=>{e.preventDefault();await put({id:id(),type:'saving',...Object.fromEntries(new FormData(f))});m.remove();await refresh()}});
 document.querySelector('#excel')?.addEventListener('click',exportExcel);
 document.querySelector('#backup')?.addEventListener('click',()=>download(JSON.stringify(records,null,2),'rr-credit-backup.json','application/json'));
 document.querySelector('#restore')?.addEventListener('change',async e=>{const file=e.target.files[0];if(!file)return;try{const arr=JSON.parse(await file.text());if(!Array.isArray(arr))throw Error();for(const x of arr)if(x.id&&x.type)await put(x);await refresh();alert('Backup restored.')}catch{alert('Invalid backup file.')}});
}
function download(data,name,type){const a=document.createElement('a');const url=URL.createObjectURL(new Blob([data],{type}));a.href=url;a.download=name;a.click();URL.revokeObjectURL(url)}
function exportExcel(){
 const wb=XLSX.utils.book_new(), add=(name,rows)=>XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),name);
 add('Dashboard',[{Metric:'Total outstanding',Value:totalDue()},{Metric:'Customers',Value:customers().length},{Metric:'Total savings',Value:savings().reduce((a,x)=>a+ +x.amount,0)}]);
 add('Customers',customers().map(c=>({Name:c.name,Phone:c.phone||'',Outstanding:balance(c.id)})));
 add('Transactions',tx().map(x=>({Date:x.date,Customer:customers().find(c=>c.id===x.customerId)?.name||'',Type:x.kind==='credit'?'Credit given':'Payment received',Amount:+x.amount,Note:x.note||''})));
 add('Collections',tx().filter(x=>x.kind==='payment').map(x=>({Date:x.date,Customer:customers().find(c=>c.id===x.customerId)?.name||'',Amount:+x.amount,Note:x.note||''})));
 add('Savings',savings().map(x=>({Date:x.date,Amount:+x.amount,Note:x.note||''})));
 XLSX.writeFile(wb,'RR_Credit_Report.xlsx');
}
async function refresh(){records=await all();render()}
(async()=>{db=await openDB();await refresh();if('serviceWorker'in navigator)navigator.serviceWorker.register('/sw.js').catch(()=>{});})();
