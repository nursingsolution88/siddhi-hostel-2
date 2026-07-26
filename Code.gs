
const CONFIG = {
  SPREADSHEET_ID: "14XDuf-Gs-RsbtI4hSs9Jsv2iaw7mhoVRqXAuIKTMNy0",
  HOSTEL_NAME: "Siddhi Hostel",
  OWNER_EMAIL: "", // apni email yahan likhen
  TOTAL_ROOMS: 18,
  BEDS: ["A","B","C"],
  TIMEZONE: "Asia/Kolkata"
};

const H = {
  Students:["id","name","fatherName","motherName","phone","parentPhone","whatsapp","email","room","bedNo","joiningDate","monthlyFee","securityDeposit","previousDue","nextDue","active","aadhaar","collegeName","course","semester","address","emergencyName","emergencyPhone","bloodGroup","medicalNotes","photoUrl","lastDueEmailDate","lastOverdueEmailDate","createdAt"],
  Payments:["receiptId","studentId","studentName","room","bedNo","amount","mode","note","monthsCleared","oldDueDate","nextDueDate","balanceDue","date","createdAt"],
  DueList:["studentId","name","room","bedNo","email","joiningDate","dueDate","monthsDue","monthlyFee","previousDue","totalDue","status"],
  RoomStatus:["room","bedNo","studentId","studentName","phone","joiningDate","nextDue","status"],
  MonthlyReport:["month","totalStudents","totalCollection","pendingAmount","occupancyPercent","vacantBeds","generatedAt"]
};

function doGet(){ setup(); return HtmlService.createHtmlOutputFromFile("index").setTitle(CONFIG.HOSTEL_NAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); }
function book(){ return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID); }
function tab(n){ return book().getSheetByName(n); }
function setup(){
  Object.keys(H).forEach(n=>{
    let s=tab(n); if(!s)s=book().insertSheet(n);
    const row=s.getRange(1,1,1,H[n].length).getValues()[0];
    if(H[n].some((x,i)=>row[i]!==x)){ s.clear(); s.getRange(1,1,1,H[n].length).setValues([H[n]]).setFontWeight("bold"); s.setFrozenRows(1); }
  });
  rebuildRooms(); rebuildDues(); installTrigger();
  return "Setup complete";
}
function installTrigger(){
  const has=ScriptApp.getProjectTriggers().some(t=>t.getHandlerFunction()==="dailyAutomation");
  if(!has) ScriptApp.newTrigger("dailyAutomation").timeBased().everyDays(1).atHour(8).create();
}
function dailyAutomation(){ rebuildDues(); sendAutomaticReminders(); sendOwnerSummary(); monthlyReport(); }

function api(action,p={}){
  try{
    setup();
    if(action==="dashboard") return ok(dashboard());
    if(action==="students") return ok(rows("Students"));
    if(action==="rooms") return ok(rebuildRooms());
    if(action==="dues") return ok(rebuildDues());
    if(action==="payments") return ok(rows("Payments").reverse());
    if(action==="addStudent") return ok(addStudent(p));
    if(action==="payment") return ok(payment(p));
    if(action==="reminder") return ok(reminder(p.studentId));
    if(action==="left") return ok(markLeft(p.id));
    throw new Error("Unknown action");
  }catch(e){ return {success:false,error:e.message}; }
}
function ok(data){return {success:true,data};}
function val(v){ return v instanceof Date?Utilities.formatDate(v,CONFIG.TIMEZONE,"yyyy-MM-dd"):v; }
function rows(n){
  const a=tab(n).getDataRange().getValues(); if(a.length<2)return [];
  return a.slice(1).filter(r=>r.some(x=>x!=="")).map(r=>Object.fromEntries(a[0].map((h,i)=>[h,val(r[i])])));
}
function append(n,o){ tab(n).appendRow(H[n].map(k=>o[k]??"")); }
function update(n,key,id,p){
  const a=tab(n).getDataRange().getValues(), hd=a[0], c=hd.indexOf(key);
  for(let i=1;i<a.length;i++) if(String(a[i][c])===String(id)){Object.keys(p).forEach(k=>{const j=hd.indexOf(k);if(j>=0)tab(n).getRange(i+1,j+1).setValue(p[k]);});return;}
  throw new Error("Record not found");
}
function date(v){
  if(v instanceof Date&&!isNaN(v))return new Date(v.getFullYear(),v.getMonth(),v.getDate());
  const m=String(v||"").match(/^(\d{4})-(\d{2})-(\d{2})$/); if(m)return new Date(+m[1],+m[2]-1,+m[3]);
  const d=new Date(v); if(isNaN(d))throw new Error("Invalid date: "+v); return new Date(d.getFullYear(),d.getMonth(),d.getDate());
}
function f(d){return Utilities.formatDate(d,CONFIG.TIMEZONE,"yyyy-MM-dd");}
function plusMonth(v,n=1){const d=date(v),day=d.getDate(),x=new Date(d.getFullYear(),d.getMonth()+n,1),last=new Date(x.getFullYear(),x.getMonth()+1,0).getDate();x.setDate(Math.min(day,last));return x;}
function occupied(room,bed,ignore=""){return rows("Students").some(s=>String(s.active).toLowerCase()==="yes"&&String(s.room)===String(room)&&String(s.bedNo).toUpperCase()===String(bed).toUpperCase()&&String(s.id)!==String(ignore));}

function addStudent(p){
  ["name","phone","room","bedNo","joiningDate","monthlyFee"].forEach(k=>{if(!p[k])throw new Error(k+" required");});
  if(occupied(p.room,p.bedNo))throw new Error("Room/Bed already occupied");
  const j=date(p.joiningDate), id="SH"+Utilities.formatDate(new Date(),CONFIG.TIMEZONE,"yyyyMMddHHmmss");
  const o={id,name:p.name,fatherName:p.fatherName||"",motherName:p.motherName||"",phone:p.phone,parentPhone:p.parentPhone||"",whatsapp:p.whatsapp||p.phone,email:p.email||"",room:String(p.room),bedNo:String(p.bedNo).toUpperCase(),joiningDate:f(j),monthlyFee:+p.monthlyFee||0,securityDeposit:+p.securityDeposit||0,previousDue:+p.previousDue||0,nextDue:f(plusMonth(j)),active:"Yes",aadhaar:p.aadhaar||"",collegeName:p.collegeName||"",course:p.course||"",semester:p.semester||"",address:p.address||"",emergencyName:p.emergencyName||"",emergencyPhone:p.emergencyPhone||"",bloodGroup:p.bloodGroup||"",medicalNotes:p.medicalNotes||"",photoUrl:p.photoUrl||"",lastDueEmailDate:"",lastOverdueEmailDate:"",createdAt:new Date()};
  append("Students",o); rebuildRooms(); rebuildDues(); return o;
}
function markLeft(id){update("Students","id",id,{active:"No"});rebuildRooms();rebuildDues();return true;}

function rebuildRooms(){
  const s=tab("RoomStatus"); s.getRange(2,1,Math.max(1,s.getMaxRows()-1),H.RoomStatus.length).clearContent();
  const st=rows("Students").filter(x=>String(x.active).toLowerCase()==="yes"), out=[];
  for(let r=1;r<=CONFIG.TOTAL_ROOMS;r++)CONFIG.BEDS.forEach(b=>{const x=st.find(z=>String(z.room)===String(r)&&String(z.bedNo).toUpperCase()===b);out.push(x?[r,b,x.id,x.name,x.phone,x.joiningDate,x.nextDue,"Occupied"]:[r,b,"","","","","","Vacant"]);});
  s.getRange(2,1,out.length,out[0].length).setValues(out);
  return out.map(r=>Object.fromEntries(H.RoomStatus.map((h,i)=>[h,r[i]])));
}
function dueOf(s,t=new Date()){
  const d=date(s.nextDue),today=date(t); if(d>today)return null;
  let m=0,c=new Date(d);while(c<=today&&m<120){m++;c=plusMonth(c);}
  const fee=+s.monthlyFee||0,prev=+s.previousDue||0;
  return {studentId:s.id,name:s.name,room:s.room,bedNo:s.bedNo,email:s.email,joiningDate:s.joiningDate,dueDate:f(d),monthsDue:m,monthlyFee:fee,previousDue:prev,totalDue:m*fee+prev,status:d<today?"Overdue":"Due Today"};
}
function rebuildDues(){
  const s=tab("DueList");s.getRange(2,1,Math.max(1,s.getMaxRows()-1),H.DueList.length).clearContent();
  const out=rows("Students").filter(x=>String(x.active).toLowerCase()==="yes").map(x=>dueOf(x)).filter(Boolean);
  if(out.length)s.getRange(2,1,out.length,H.DueList.length).setValues(out.map(o=>H.DueList.map(k=>o[k]??"")));
  return out;
}
function payment(p){
  const s=rows("Students").find(x=>String(x.id)===String(p.studentId));if(!s)throw new Error("Student not found");
  const amt=+p.amount||0;if(amt<=0)throw new Error("Enter valid amount");
  const d=dueOf(s)||{monthsDue:0,totalDue:+s.previousDue||0},fee=+s.monthlyFee||0;
  let cleared=fee?Math.floor(Math.max(0,amt-(+s.previousDue||0))/fee):0;if(amt>=d.totalDue&&d.monthsDue>cleared)cleared=d.monthsDue;
  const nd=plusMonth(s.nextDue,cleared), bal=Math.max(0,d.totalDue-amt), rid="RCP"+Utilities.formatDate(new Date(),CONFIG.TIMEZONE,"yyyyMMddHHmmss");
  append("Payments",{receiptId:rid,studentId:s.id,studentName:s.name,room:s.room,bedNo:s.bedNo,amount:amt,mode:p.mode||"Cash",note:p.note||"",monthsCleared:cleared,oldDueDate:s.nextDue,nextDueDate:f(nd),balanceDue:bal,date:f(new Date()),createdAt:new Date()});
  update("Students","id",s.id,{nextDue:f(nd),previousDue:bal});
  rebuildDues();sendReceipt(s,{receiptId:rid,amount:amt,mode:p.mode||"Cash",nextDueDate:f(nd),balanceDue:bal});
  return {receiptId:rid,nextDueDate:f(nd),balanceDue:bal};
}
function sendReceipt(s,r){
  if(!s.email)return;
  MailApp.sendEmail(s.email,CONFIG.HOSTEL_NAME+" Payment Receipt - "+r.receiptId,
`Dear ${s.name},

Payment received successfully.
Receipt: ${r.receiptId}
Room/Bed: ${s.room}/${s.bedNo}
Amount: ₹${r.amount}
Mode: ${r.mode}
Next Due: ${r.nextDueDate}
Balance: ₹${r.balanceDue}

Regards,
${CONFIG.HOSTEL_NAME}`);
}
function reminder(id){
  const s=rows("Students").find(x=>String(x.id)===String(id));if(!s||!s.email)throw new Error("Student email not available");
  const d=dueOf(s);if(!d)throw new Error("Student is not due");
  MailApp.sendEmail(s.email,CONFIG.HOSTEL_NAME+" Fee Reminder",`Dear ${s.name},

Your hostel fee is ${d.status.toLowerCase()}.
Room/Bed: ${s.room}/${s.bedNo}
Due Date: ${d.dueDate}
Total Due: ₹${d.totalDue}

Please deposit the fee.

Regards,
${CONFIG.HOSTEL_NAME}`);
  return true;
}
function sendAutomaticReminders(){rebuildDues().forEach(d=>{try{reminder(d.studentId)}catch(e){}});}
function dashboard(){
  const st=rows("Students").filter(x=>String(x.active).toLowerCase()==="yes"),du=rebuildDues(),pa=rows("Payments"),mo=Utilities.formatDate(new Date(),CONFIG.TIMEZONE,"yyyy-MM"),beds=CONFIG.TOTAL_ROOMS*CONFIG.BEDS.length;
  return {totalRooms:CONFIG.TOTAL_ROOMS,totalBeds:beds,totalStudents:st.length,occupiedBeds:st.length,vacantBeds:beds-st.length,dueStudents:du.filter(x=>x.status==="Due Today").length,overdueStudents:du.filter(x=>x.status==="Overdue").length,pendingAmount:du.reduce((a,x)=>a+(+x.totalDue||0),0),monthlyCollection:pa.filter(x=>String(x.date).startsWith(mo)).reduce((a,x)=>a+(+x.amount||0),0),recentPayments:pa.slice(-5).reverse(),todayDue:du.slice(0,5)};
}
function sendOwnerSummary(){if(!CONFIG.OWNER_EMAIL)return;const d=dashboard();MailApp.sendEmail(CONFIG.OWNER_EMAIL,CONFIG.HOSTEL_NAME+" Daily Summary",`Students: ${d.totalStudents}
Vacant Beds: ${d.vacantBeds}
Due Today: ${d.dueStudents}
Overdue: ${d.overdueStudents}
Pending: ₹${d.pendingAmount}
Monthly Collection: ₹${d.monthlyCollection}`);}
function monthlyReport(){
  const d=dashboard(),month=Utilities.formatDate(new Date(),CONFIG.TIMEZONE,"yyyy-MM"),o={month,totalStudents:d.totalStudents,totalCollection:d.monthlyCollection,pendingAmount:d.pendingAmount,occupancyPercent:Math.round(d.occupiedBeds/d.totalBeds*100),vacantBeds:d.vacantBeds,generatedAt:new Date()};
  append("MonthlyReport",o);return o;
}
