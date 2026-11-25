import * as XLSX from 'xlsx';
import './style.css';

const DB='rr-credit-v2', STORE='records';
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
let records=[],tab='dashboard',selectedCustomerId=null;
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const customers=()=>records.filter(x=>x.type==='customer');
const tx=()=>records.filter(x=>x.type==='transaction');
const savings=()=>records.filter(x=>x.type==='saving');
const balance=cid=>tx().filter(x=>x.customerId===cid).reduce((a,x)=>a+(x.kind==='credit'?+x.amount:-x.amount),0);
const totalDue=()=>customers().reduce((a,c)=>a+Math.max(0,balance(c.id)),0);
const fmtDate=d=>new Date(d).toLocaleDateString('en-IN');
function shell(){
 document.querySelector('#app').innerHTML=`
 <header><div class="brand"><img class="brand-logo" src="/rr-credit-logo.png" alt="RR Credit"><div><b>RR Credit</b><small>RR Mobiles · Offline ledger</small></div></div><span class="offline"><i></i> Offline ready</span></header>
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
 if(tab==='customers') v.innerHTML=selectedCustomerId?customerLedgerView():`<div class="pagehead"><div><p class="eyebrow">CUSTOMER ACCOUNTS</p><h1>Customers</h1></div><button class="primary" id="addcust">＋ Add customer</button></div><div class="panel">${customers().map(c=>`<div class="customer"><div class="avatar">${esc(c.name[0]?.toUpperCase())}</div><div class="grow"><b>${esc(c.name)}</b><small>${esc(c.phone||'No phone number')}</small></div><strong>${money(balance(c.id))}</strong><button class="small" data-ledger="${c.id}">Ledger</button><button class="small delete-customer" data-delete-customer="${c.id}" title="Delete customer">Delete</button></div>`).join('')||'<p class="empty">No customers added yet.</p>'}</div>`;
 if(tab==='transactions') v.innerHTML=`<div class="pagehead"><div><p class="eyebrow">CREDIT & COLLECTIONS</p><h1>Entries</h1></div><button class="primary" id="addtx">＋ Add entry</button></div><div class="panel">${tx().sort((a,b)=>b.date.localeCompare(a.date)).map(entryRow).join('')||'<p class="empty">No transactions yet.</p>'}</div>`;
 if(tab==='savings') v.innerHTML=`<div class="pagehead"><div><p class="eyebrow">PERSONAL / SHOP SAVINGS</p><h1>Savings</h1></div><button class="primary" id="addsave">＋ Add saving</button></div><div class="stats"><article><small>Total recorded savings</small><strong>${money(savings().reduce((a,x)=>a+ +x.amount,0))}</strong></article><article><small>Records</small><strong>${savings().length}</strong></article></div><div class="panel">${savings().sort((a,b)=>b.date.localeCompare(a.date)).map(x=>`<div class="entry"><div><b>${esc(x.note||'Savings deposit')}</b><small>${fmtDate(x.date)}</small></div><strong class="green">+ ${money(x.amount)}</strong></div>`).join('')||'<p class="empty">No savings recorded yet.</p>'}</div>`;
 if(tab==='reports') v.innerHTML=`<div class="pagehead"><div><p class="eyebrow">YOUR BUSINESS RECORDS</p><h1>Reports & backup</h1></div></div><div class="panel"><h2>Excel report</h2><p>Export a workbook containing Dashboard, Customers, Transactions, Collections and Savings sheets.</p><button class="primary" id="excel">⇩ Download Excel report</button><hr><h2>Local backup</h2><p>Save all RR Credit records to a JSON file. Keep it somewhere safe.</p><button class="secondary" id="backup">Download data backup</button><label class="filebtn">Restore from backup<input id="restore" type="file" accept=".json,application/json"></label><p class="muted">Offline-only: no cloud sync. Create a manual backup only when you choose.</p></div>`;
 bind();
}
function entryRow(x){
 const c=customers().find(c=>c.id===x.customerId);
 const given=x.kind==='credit';
 const time=x.time?new Date(`2000-01-01T${x.time}`).toLocaleTimeString('en-IN',{hour:'2-digit',minute:'2-digit',hour12:true}):'Time not recorded';
 return `<div class="entry"><div><b>${esc(c?.name||'Customer')}</b><small>${given?'Given':'Received'} · ${fmtDate(x.date)} · ${time}${x.note?' · '+esc(x.note):''}</small></div><strong class="${given?'red':'green'}">${given?'↑':'↓'} ${money(x.amount)}</strong></div>`;
}
function customerLedgerView(){
 const c=customers().find(x=>x.id===selectedCustomerId);
 if(!c){selectedCustomerId=null;return '<p class="empty">Customer not found.</p>'}
 const items=tx().filter(x=>x.customerId===c.id);
 const chronological=[...items].sort((a,b)=>
  `${a.date||'0000-00-00'}T${a.time||'00:00'}`.localeCompare(`${b.date||'0000-00-00'}T${b.time||'00:00'}`)||(a.id||'').localeCompare(b.id||'')
 );
 let running=0;
 const withBalances=chronological.map(x=>{running+=x.kind==='credit'?+x.amount:-x.amount;return {...x,running}});
 const grouped={};withBalances.forEach(x=>{(grouped[x.date||'No date']??=[]).push(x)});
 const rawPhone=(c.phone||'').replace(/\D/g,'');
 const phone=rawPhone.length===10?'91'+rawPhone:rawPhone;
 const days=Object.entries(grouped).sort((a,b)=>b[0].localeCompare(a[0]));
 const givenTotal=items.filter(x=>x.kind==='credit').reduce((a,x)=>a+ +x.amount,0);
 const receivedTotal=items.filter(x=>x.kind==='payment').reduce((a,x)=>a+ +x.amount,0);
 const currentBalance=balance(c.id);
 const balanceLabel=currentBalance>0?'Balance Due':currentBalance<0?'Advance received':'Settled';
 const balanceClass=currentBalance<0?'green':currentBalance>0?'red':'settled';
 return `<div class="ledger-top">
 <button class="link back-ledger" id="backcustomers">← Customers</button>
 <div class="ledger-person"><div class="avatar big">${esc(c.name[0]?.toUpperCase())}</div><div><h1>${esc(c.name)}</h1><small>${esc(c.phone||'No mobile number')}</small></div></div>
 <div class="ledger-actions">${phone?`<a class="action-call" href="tel:${esc(phone)}">☎ Call</a><button class="action-wa" type="button" id="whatsappBusinessRemind">◉ Remind</button>`:''}<button class="due-btn" id="setdue">▣ ${c.dueDate?'Due '+fmtDate(c.dueDate):'Due Date'}</button></div>
 <div class="balance-card"><small>${balanceLabel}</small><strong class="${balanceClass}">${currentBalance<0?'+ ':''}${money(Math.abs(currentBalance))}</strong></div>
 <div class="ledger-totals"><div><small>↓ Received</small><b>${money(receivedTotal)}</b></div><div><small>↑ Given</small><b>${money(givenTotal)}</b></div></div></div>
 <div class="ledger-list">${days.length?days.map(([date,rows])=>{
  const ordered=[...rows].sort((a,b)=>(b.time||'00:00').localeCompare(a.time||'00:00')||(b.id||'').localeCompare(a.id||''));
  return `<section class="ledger-day"><div class="day-label">${date==='No date'?'No date':fmtDate(date)}</div>
   ${ordered.map(x=>{
    const given=x.kind==='credit',time=x.time?new Date(`2000-01-01T${x.time}`).toLocaleTimeString('en-IN',{hour:'2-digit',minute:'2-digit',hour12:true}):'Time not recorded';
    const due=x.running,runningText=due>0?`${money(due)} Due`:due<0?`${money(Math.abs(due))} Advance`:'Settled';
    const runningClass=due<0?'green':due>0?'red':'settled';
    return `<div class="ledger-transaction ${given?'given':'received'}"><div class="transaction-box">
     <div class="transaction-heading"><b>${given?'↑ Given':'↓ Received'}</b><time>◷ ${time}</time></div>
     <strong class="transaction-amount ${given?'red':'green'}">${money(x.amount)}</strong>
     ${x.note?`<small class="transaction-note">${esc(x.note)}</small>`:''}
     <small class="transaction-balance">Balance: <b class="${runningClass}">${runningText}</b></small>
    </div></div>`;
   }).join('')}
  </section>`;
 }).join(''):'<p class="empty">No transactions yet. Add the first entry below.</p>'}</div>
 <div class="ledger-add-actions"><button class="add-received" id="addReceived">↓ Add Received</button><button class="add-given" id="addGiven">↑ Add Given</button></div>`;
}
function modal(title,body){const m=document.createElement('div');m.className='modalback';m.innerHTML=`<form class="modal"><div class="panelhead"><h2>${title}</h2><button type="button" class="close">×</button></div>${body}<button class="primary full" type="submit">Save</button></form>`;document.body.append(m);m.querySelector('.close').onclick=()=>m.remove();return [m,m.querySelector('form')]}
function bind(){
 document.querySelector('#quick')?.addEventListener('click',()=>{tab='transactions';render();document.querySelector('#addtx')?.click()});
 document.querySelector('#alltx')?.addEventListener('click',()=>{tab='transactions';render()});
 document.querySelector('#addcust')?.addEventListener('click',()=>{const [m,f]=modal('Add customer',`<label>Customer name<input name="name" required></label><label>Mobile number<input name="phone" inputmode="numeric"></label>`);f.onsubmit=async e=>{e.preventDefault();await put({id:id(),type:'customer',...Object.fromEntries(new FormData(f))});m.remove();await refresh() }});
 document.querySelectorAll('[data-ledger]').forEach(b=>b.onclick=()=>{selectedCustomerId=b.dataset.ledger;render()});
 document.querySelectorAll('[data-delete-customer]').forEach(b=>b.onclick=async()=>{const customer=customers().find(c=>c.id===b.dataset.deleteCustomer);if(!customer)return;const currentBalance=balance(customer.id);if(currentBalance!==0){const amount=money(Math.abs(currentBalance));const direction=currentBalance>0?'outstanding due':'advance balance';alert(`Cannot delete ${customer.name} yet. This customer has ${amount} ${direction}. First settle the balance to ₹0, then tap Delete again.`);return;}if(!confirm(`Delete ${customer.name}? This will permanently delete this customer's Given and Received transaction history from this phone. This cannot be undone.`))return;for(const entry of tx().filter(x=>x.customerId===customer.id)){await del(entry.id);}await del(customer.id);await refresh();});
 document.querySelector('#backcustomers')?.addEventListener('click',()=>{selectedCustomerId=null;render()});
 document.querySelector('#addReceived')?.addEventListener('click',()=>openTransactionForm(selectedCustomerId,'payment'));
 document.querySelector('#addGiven')?.addEventListener('click',()=>openTransactionForm(selectedCustomerId,'credit'));
 document.querySelector('#whatsappBusinessRemind')?.addEventListener('click',()=>{
  const c=customers().find(x=>x.id===selectedCustomerId); if(!c)return;
  const digits=(c.phone||'').replace(/\D/g,'');
  const phone=digits.length===10?'91'+digits:digits;
  if(!phone){alert('Add a valid customer mobile number first.');return;}
  const message=`Hello ${c.name}, your outstanding balance at RR Mobiles is ${money(Math.max(0,balance(c.id)))}. Please make the payment. Thank you.`;
  const intent=`intent://send?phone=${phone}&text=${encodeURIComponent(message)}#Intent;scheme=whatsapp;package=com.whatsapp.w4b;end`;
  let switched=false;
  const onVisibility=()=>{if(document.visibilityState==='hidden')switched=true;};
  document.addEventListener('visibilitychange',onVisibility,{once:true});
  window.location.href=intent;
  setTimeout(()=>{document.removeEventListener('visibilitychange',onVisibility);if(!switched&&document.visibilityState==='visible')alert('WhatsApp Business did not open. Please install WhatsApp Business and try again.');},1800);
 });
 document.querySelector('#setdue')?.addEventListener('click',()=>{const c=customers().find(x=>x.id===selectedCustomerId);const [m,f]=modal('Set customer due date',`<label>Due date<input name="dueDate" type="date" value="${c.dueDate||''}" required></label>`);f.onsubmit=async e=>{e.preventDefault();c.dueDate=new FormData(f).get('dueDate');await put(c);m.remove();await refresh()}});
 document.querySelector('#addtx')?.addEventListener('click',()=>openTransactionForm());
 document.querySelector('#addsave')?.addEventListener('click',()=>{const [m,f]=modal('Record savings',`<label>Amount (₹)<input name="amount" type="number" min="0.01" step="0.01" required></label><label>Date<input name="date" type="date" value="${new Date().toISOString().slice(0,10)}" required></label><label>Note<input name="note" placeholder="Daily savings"></label>`);f.onsubmit=async e=>{e.preventDefault();await put({id:id(),type:'saving',...Object.fromEntries(new FormData(f))});m.remove();await refresh()}});
 document.querySelector('#excel')?.addEventListener('click',exportExcel);
 document.querySelector('#backup')?.addEventListener('click',()=>download(JSON.stringify(records,null,2),'rr-credit-backup.json','application/json'));
 document.querySelector('#restore')?.addEventListener('change',async e=>{const file=e.target.files[0];if(!file)return;try{const arr=JSON.parse(await file.text());if(!Array.isArray(arr))throw Error();for(const x of arr)if(x.id&&x.type)await put(x);await refresh();alert('Backup restored.')}catch{alert('Invalid backup file.')}});
}

function openTransactionForm(customerId=null,defaultKind=null){
 if(!customers().length){alert('Add a customer first.');return}
 const opts=customers().map(c=>`<option value="${c.id}" ${c.id===customerId?'selected':''}>${esc(c.name)}</option>`).join('');
 const now=new Date(),pad=n=>String(n).padStart(2,'0');
 const today=`${now.getFullYear()}-${pad(now.getMonth()+1)}-${pad(now.getDate())}`;
 const currentTime=`${pad(now.getHours())}:${pad(now.getMinutes())}`;
 const [m,f]=modal('Add ledger entry',`<label>Customer<select name="customerId" ${customerId?'disabled':''}>${opts}</select></label>${customerId?`<input type="hidden" name="customerId" value="${customerId}">`:''}<label>Entry type<select name="kind"><option value="credit" ${defaultKind==='credit'?'selected':''}>↑ Given (customer owes)</option><option value="payment" ${defaultKind==='payment'?'selected':''}>↓ Received (customer paid)</option></select></label><label>Amount (₹)<input name="amount" type="number" min="0.01" step="0.01" required></label><label>Date<input name="date" type="date" value="${today}" required></label><label>Time<input name="time" type="time" value="${currentTime}" required></label><label>Description<input name="note" placeholder="e.g. Mobile display repair / cash received"></label>`);
 f.onsubmit=async e=>{e.preventDefault();const d=Object.fromEntries(new FormData(f));await put({id:id(),type:'transaction',...d});m.remove();await refresh()};
}

function download(data,name,type){const a=document.createElement('a');const url=URL.createObjectURL(new Blob([data],{type}));a.href=url;a.download=name;a.click();URL.revokeObjectURL(url)}
function exportExcel(){
 const wb=XLSX.utils.book_new(), add=(name,rows)=>XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),name);
 add('Dashboard',[{Metric:'Total outstanding',Value:totalDue()},{Metric:'Customers',Value:customers().length},{Metric:'Total savings',Value:savings().reduce((a,x)=>a+ +x.amount,0)}]);
 add('Customers',customers().map(c=>({Name:c.name,Phone:c.phone||'',Outstanding:balance(c.id)})));
 add('Transactions',tx().map(x=>({Date:x.date,Time:x.time||'',Customer:customers().find(c=>c.id===x.customerId)?.name||'',Type:x.kind==='credit'?'Credit given':'Payment received',Amount:+x.amount,Note:x.note||''})));
 add('Collections',tx().filter(x=>x.kind==='payment').map(x=>({Date:x.date,Time:x.time||'',Customer:customers().find(c=>c.id===x.customerId)?.name||'',Amount:+x.amount,Note:x.note||''})));
 add('Savings',savings().map(x=>({Date:x.date,Amount:+x.amount,Note:x.note||''})));
 XLSX.writeFile(wb,'RR_Credit_Report.xlsx');
}
async function refresh(){records=await all();render()}
(async()=>{db=await openDB();await refresh();if('serviceWorker'in navigator)navigator.serviceWorker.register('/sw.js').catch(()=>{});})();
