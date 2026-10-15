function publicDataV6(){
 const cash=rows_(APP.sheets.cash).filter(r=>active_(r,9));
 const goods=rows_(APP.sheets.goods).filter(r=>active_(r,11)&&String(r.v[10]).toUpperCase()==='PUBLIC');
 const expense=rows_(APP.sheets.expenses).filter(r=>active_(r,9));
 const contributed=new Set(cash.map(r=>String(r.v[2]||'').trim()).filter(Boolean));
 const cashTotal=cash.reduce((s,r)=>s+Number(r.v[5]||0),0),expenseTotal=expense.reduce((s,r)=>s+Number(r.v[4]||0),0);
 const monthKey=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM');
 const monthMap={},catMap={};
 cash.forEach(r=>{const k=dateV6_(r.v[1]).slice(0,7);if(!monthMap[k])monthMap[k]={cash:0,expense:0};monthMap[k].cash+=Number(r.v[5]||0);});
 expense.forEach(r=>{const k=dateV6_(r.v[1]).slice(0,7);if(!monthMap[k])monthMap[k]={cash:0,expense:0};monthMap[k].expense+=Number(r.v[4]||0);const cat=String(r.v[2]||'Khác');catMap[cat]=(catMap[cat]||0)+Number(r.v[4]||0);});
 const eligible=roster_().length;
 const latestExpense=expense.map(r=>({id:String(r.v[0]),date:dateV6_(r.v[1]),category:String(r.v[2]),description:String(r.v[3]),amount:Number(r.v[4]||0)})).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,5);
 return {
   cash:{total:cashTotal,contributors:contributed.size,eligible,average:contributed.size?Math.round(cashTotal/contributed.size/1000)*1000:0},
   summary:{expense:expenseTotal,balance:cashTotal-expenseTotal,expenseCount:expense.length,goodsCount:goods.length,thisMonthExpense:(monthMap[monthKey]||{}).expense||0},
   latestExpense,
   monthly:Object.keys(monthMap).sort().slice(-12).map(k=>({month:k,cash:monthMap[k].cash,expense:monthMap[k].expense})),
   categories:Object.keys(catMap).map(k=>({name:k,value:catMap[k]})).sort((a,b)=>b.value-a.value),
   goods:goods.map(r=>({id:String(r.v[0]),date:dateV6_(r.v[1]),donor:String(r.v[3]||''),product:String(r.v[4]||''),quantity:Number(r.v[5]||0),unit:String(r.v[6]||''),note:String(r.v[9]||''),hasImage:!!r.v[8]})).sort((a,b)=>b.date.localeCompare(a.date)).slice(0,60),
   updated:Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm')
 };
}
