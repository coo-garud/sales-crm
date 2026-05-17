// Sales Command Centre — Apps Script v3
// Auto-creates 6 sheets: Leads, FollowUps, Bookings, TestDrives, Stock, Users

const LH=["Lead ID","Created DT","Location","Source","Customer Name","Phone","Alt Phone","Model Interest","Variant","Color Pref","Budget","Finance/Cash","Salesperson","Status","Interest Level","First Contact DT","Last Contact DT","Next Followup DT","Followup Count","Lost Reason","VoC Notes","Customer Area","Customer City","Customer Expected Delivery"];
const FH=["FU ID","Lead ID","DateTime","Salesperson","Method","Outcome","Status After","Interest After","Notes","Next Followup DT"];
const BH=["Booking ID","Lead ID","Booking DT","Location","Customer Name","Phone","Model","Variant","Color","Booking Amount","Payment Mode","Lead Source","Exchange","Old Car Make","Old Car Model","Old Car Year","Exchange Value","In-house Insurance","VC Number","Discount Amount","Customer Expected Delivery","In Stock","Stock Ref","Stockyard","Expected Arrival","Actual Delivery","Planned Delivery","Status","Salesperson","Notes"];
const TH=["TD ID","Lead ID","DateTime","Customer Name","Phone","Location","Model","Salesperson","Post-TD Interest","Notes"];
const SH=["Stock ID","Chassis No","Model","Variant","Color","Added Date","Current Location","Status","Allocated To","Booking ID","Notes"];
const UH=["Username","Password","Role","Display Name"];
const DEFAULT_USERS=[
  ["masteradmin","Admin@123","masteradmin","Master Admin"],
  ["manager",    "Manager@123","manager",  "Manager"],
  ["salesperson","Sales@123","salesperson","Sales Staff"]
];

function doGet(e){return handleReq(e);}
function doPost(e){return handleReq(e);}
function handleReq(e){
  const p=(e&&e.parameter)?e.parameter:{};let res;
  try{
    switch(p.action){
      case "getAll":res=getAll();break;case "addLead":res=addLead(JSON.parse(p.data));break;
      case "addFollowUp":res=addFollowUp(JSON.parse(p.data));break;case "addBooking":res=addBooking(JSON.parse(p.data));break;
      case "updateBooking":res=updateBooking(JSON.parse(p.data));break;case "addStock":res=addStock(JSON.parse(p.data));break;case "bulkAddStock":res=bulkAddStock(JSON.parse(p.data));break;
      case "updateStock":res=updateStock(JSON.parse(p.data));break;
      case "getUsers":res=getUsers();break;
      case "saveUser":res=saveUser(JSON.parse(p.data));break;
      case "deleteUser":res=deleteUser(JSON.parse(p.data));break;
      default:res={status:"ok",msg:"Sales CRM API v3 running"};
    }
  }catch(err){res={status:"error",msg:err.toString()};}
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}
function getOrCreate(name,headers){
  const ss=SpreadsheetApp.getActiveSpreadsheet();let sh=ss.getSheetByName(name);
  if(!sh){sh=ss.insertSheet(name);const r=sh.getRange(1,1,1,headers.length);r.setValues([headers]);r.setFontWeight("bold").setBackground("#003A6B").setFontColor("#FFFFFF");sh.setFrozenRows(1);}
  return sh;
}
function sheetToArr(sh,headers){
  const last=sh.getLastRow();if(last<2)return[];
  const tz=Session.getScriptTimeZone();
  return sh.getRange(2,1,last-1,headers.length).getValues().filter(r=>r.some(c=>c!=="")).map((r,i)=>{
    const obj={rowIndex:i+2};headers.forEach((h,j)=>{let v=r[j];if(v instanceof Date)v=Utilities.formatDate(v,tz,"yyyy-MM-dd HH:mm");obj[h]=(v===null||v===undefined)?"":String(v);});return obj;
  });
}
function getAll(){return{status:"ok",leads:sheetToArr(getOrCreate("Leads",LH),LH),followups:sheetToArr(getOrCreate("FollowUps",FH),FH),bookings:sheetToArr(getOrCreate("Bookings",BH),BH),testdrives:sheetToArr(getOrCreate("TestDrives",TH),TH),stock:sheetToArr(getOrCreate("Stock",SH),SH)};}
function addLead(d){
  const sh=getOrCreate("Leads",LH);const tz=Session.getScriptTimeZone();
  d["Lead ID"]="LEAD-"+Utilities.formatDate(new Date(),tz,"yyyyMMddHHmmss");d["Status"]=d["Status"]||"Active";d["Interest Level"]=d["Interest Level"]||"Warm";d["Followup Count"]="0";
  sh.appendRow(LH.map(h=>d[h]||""));return{status:"ok",leadId:d["Lead ID"]};
}
function addFollowUp(d){
  const fsh=getOrCreate("FollowUps",FH);const lsh=getOrCreate("Leads",LH);
  const tz=Session.getScriptTimeZone();const now=Utilities.formatDate(new Date(),tz,"yyyy-MM-dd HH:mm");
  d["FU ID"]="FU-"+Utilities.formatDate(new Date(),tz,"yyyyMMddHHmmss");fsh.appendRow(FH.map(h=>d[h]||""));
  if(d["Is TD"]==="yes"){
    const tsh=getOrCreate("TestDrives",TH);const td={};
    td["TD ID"]="TD-"+Utilities.formatDate(new Date(),tz,"yyyyMMddHHmmss");td["Lead ID"]=d["Lead ID"];td["DateTime"]=now;
    td["Customer Name"]=d["Customer Name"]||"";td["Phone"]=d["Phone"]||"";td["Location"]=d["Location"]||"";
    td["Model"]=d["TD Model"]||"";td["Salesperson"]=d["Salesperson"]||"";td["Post-TD Interest"]=d["TD Interest"]||"";td["Notes"]=d["Notes"]||"";
    tsh.appendRow(TH.map(h=>td[h]||""));
  }
  const lid=d["Lead ID"];const lr=lsh.getLastRow();
  if(lr>=2){const ids=lsh.getRange(2,1,lr-1,1).getValues().flat();const ix=ids.indexOf(lid);
    if(ix>=0){const rn=ix+2;const li=h=>LH.indexOf(h)+1;
      if(!lsh.getRange(rn,li("First Contact DT")).getValue())lsh.getRange(rn,li("First Contact DT")).setValue(now);
      lsh.getRange(rn,li("Last Contact DT")).setValue(now);
      if(d["Next Followup DT"])lsh.getRange(rn,li("Next Followup DT")).setValue(d["Next Followup DT"]);
      if(d["Status After"])lsh.getRange(rn,li("Status")).setValue(d["Status After"]);
      if(d["Interest After"])lsh.getRange(rn,li("Interest Level")).setValue(d["Interest After"]);
      if(d["Notes"])lsh.getRange(rn,li("VoC Notes")).setValue(d["Notes"]);
      if(d["Status After"]==="Lost"&&d["Lost Reason"])lsh.getRange(rn,li("Lost Reason")).setValue(d["Lost Reason"]);
      const cnt=parseInt(lsh.getRange(rn,li("Followup Count")).getValue())||0;lsh.getRange(rn,li("Followup Count")).setValue(cnt+1);
    }
  }
  return{status:"ok",msg:"Follow-up logged"};
}
function addBooking(d){
  const sh=getOrCreate("Bookings",BH);const tz=Session.getScriptTimeZone();
  d["Booking ID"]="BK-"+Utilities.formatDate(new Date(),tz,"yyyyMMddHHmmss");d["Status"]=d["Status"]||"Booked";
  sh.appendRow(BH.map(h=>d[h]||""));
  if(d["Lead ID"]){const lsh=getOrCreate("Leads",LH);const lr=lsh.getLastRow();if(lr>=2){const ids=lsh.getRange(2,1,lr-1,1).getValues().flat();const ix=ids.indexOf(d["Lead ID"]);if(ix>=0)lsh.getRange(ix+2,LH.indexOf("Status")+1).setValue("Booked");}}
  return{status:"ok",bookingId:d["Booking ID"]};
}
function updateBooking(d){
  const sh=getOrCreate("Bookings",BH);const ri=parseInt(d.rowIndex);if(!ri||ri<2)throw new Error("Invalid row");
  ["Status","In Stock","Stock Ref","Stockyard","Expected Arrival","Actual Delivery","Planned Delivery","Notes"].forEach(col=>{const ci=BH.indexOf(col)+1;if(ci>0&&d[col]!==undefined&&d[col]!=="")sh.getRange(ri,ci).setValue(d[col]);});
  return{status:"ok"};
}
function bulkAddStock(d){
  const sh=getOrCreate("Stock",SH);const tz=Session.getScriptTimeZone();const today=Utilities.formatDate(new Date(),tz,"yyyy-MM-dd");
  const existing=sh.getLastRow()>=2?sh.getRange(2,1,sh.getLastRow()-1,2).getValues().map(r=>String(r[1]).trim().toUpperCase()):[];
  let added=0,skipped=0;
  (d.items||[]).forEach(item=>{
    const cn=String(item["Chassis No"]||"").trim().toUpperCase();
    if(!cn){skipped++;return;}
    if(existing.includes(cn)){skipped++;return;}
    const row=SH.map(h=>item[h]||"");row[0]="STK-"+Date.now()+"-"+added;row[5]=today;
    sh.appendRow(row);existing.push(cn);added++;
  });
  return{status:"ok",added:added,skipped:skipped};
}
function addStock(d){
  const sh=getOrCreate("Stock",SH);const tz=Session.getScriptTimeZone();
  d["Stock ID"]="STK-"+Utilities.formatDate(new Date(),tz,"yyyyMMddHHmmss");d["Added Date"]=Utilities.formatDate(new Date(),tz,"yyyy-MM-dd");d["Status"]=d["Status"]||"Available";
  sh.appendRow(SH.map(h=>d[h]||""));return{status:"ok",stockId:d["Stock ID"]};
}
function updateStock(d){
  const sh=getOrCreate("Stock",SH);const ri=parseInt(d.rowIndex);if(!ri||ri<2)throw new Error("Invalid row");
  ["Status","Current Location","Allocated To","Booking ID","Notes"].forEach(col=>{const ci=SH.indexOf(col)+1;if(ci>0&&d[col]!==undefined)sh.getRange(ri,ci).setValue(d[col]);});
  return{status:"ok"};
}
function getUsersSheet(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();let sh=ss.getSheetByName("Users");
  if(!sh){
    sh=ss.insertSheet("Users");
    const r=sh.getRange(1,1,1,UH.length);r.setValues([UH]);r.setFontWeight("bold").setBackground("#003A6B").setFontColor("#FFFFFF");sh.setFrozenRows(1);
    DEFAULT_USERS.forEach(row=>sh.appendRow(row));
  }
  return sh;
}
function getUsers(){
  const sh=getUsersSheet();const last=sh.getLastRow();if(last<2)return{status:"ok",users:[]};
  const rows=sh.getRange(2,1,last-1,UH.length).getValues();
  const users=rows.filter(r=>r[0]).map(r=>({username:String(r[0]),password:String(r[1]),role:String(r[2]),display:String(r[3]||r[0])}));
  return{status:"ok",users};
}
function saveUser(d){
  const sh=getUsersSheet();const last=sh.getLastRow();let found=false;
  if(last>=2){const ids=sh.getRange(2,1,last-1,1).getValues().flat().map(String);const ix=ids.indexOf(String(d.username));
    if(ix>=0){sh.getRange(ix+2,1,1,4).setValues([[d.username,d.password,d.role,d.display||d.username]]);found=true;}}
  if(!found)sh.appendRow([d.username,d.password,d.role,d.display||d.username]);
  return{status:"ok"};
}
function deleteUser(d){
  const sh=getUsersSheet();const last=sh.getLastRow();if(last<2)return{status:"ok"};
  const ids=sh.getRange(2,1,last-1,1).getValues().flat().map(String);const ix=ids.indexOf(String(d.username));
  if(ix>=0)sh.deleteRow(ix+2);
  return{status:"ok"};
}
