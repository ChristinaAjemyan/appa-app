import{useState,useMemo,useEffect,useRef}from"react";
import{calcStorage as supabaseStorage}from"./calcStorage";
import * as XLSX from "xlsx";

const COMPANIES=["Nairi","Ingo","Liga","Sil","Rego"];
const ALL_COMPANIES=["Nairi","Ingo","Liga","Sil","Rego","Armenia"];
const ARM_GROUPS=["1-9","10-14","15-25"];
const MONTHS_RU=["Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"];
const EXCEPTION_BRANDS=["opel","mercedes","mercedes-benz","bmw","nissan"];
const _now=new Date();
const CURRENT_MONTH=`${_now.getFullYear()}-${String(_now.getMonth()+1).padStart(2,"0")}`;
const MIN_MONTH="2023-01";const MAX_MONTH="2030-12";
const EXPIRY_OPTIONS=[];
for(let y=2025;y<=2031;y++)for(let m=1;m<=12;m++)EXPIRY_OPTIONS.push(`${y}-${String(m).padStart(2,"0")}`);

const DEFAULT_RATES={
  officeRates:{Nairi:16,Ingo:7,Liga:5,Sil:10,Rego:5},
  armeniaOffice:{"1-9":11,"10-14":10,"15-25":7},
  agentRates:{Nairi:8,Ingo:3,Liga:2,Sil:5,Rego:2},
  armeniaAgent:{"1-9":6,"10-14":5,"15-25":3},
  armeniaShort:{officeRate:37,agentRate:20},
  nairiRegion:{bm1_7:10,bm8_25:5},
  agentOverrides:{},
};
const DEFAULT_VOL_RATES={rates:[]};
const DEFAULT_EXCEPTIONS=[
  {id:"e1",company:"Nairi",enabled:true,excludedAgents:[],conditions:[{field:"region",op:"eq",value:"YR"},{field:"bm",op:"eq",value:"10"},{field:"term",op:"eq",value:"L"}]},
  {id:"e2",company:"Sil",enabled:true,excludedAgents:[],conditions:[{field:"bm",op:"between",value:"11",value2:"14"}]},
  {id:"e3",company:"Rego",enabled:true,excludedAgents:[],conditions:[{field:"region",op:"eq",value:"YR"}]},
  {id:"e4",company:"Rego",enabled:true,excludedAgents:[],conditions:[{field:"bm",op:"gte",value:"8"},{field:"power",op:"lte",value:"141"}]},
  {id:"e5",company:"Rego",enabled:true,excludedAgents:[],conditions:[{field:"bm",op:"gte",value:"10"},{field:"power",op:"gte",value:"141"}]},
  {id:"e6",company:"Nairi",enabled:true,type:"brand_whitelist",excludedAgents:[],conditions:[{field:"bm",op:"between",value:"10",value2:"17"}]},
];
const FIELDS=[{key:"region",label:"Регион",type:"string"},{key:"bm",label:"БМ",type:"number"},{key:"power",label:"Мощность",type:"number"},{key:"term",label:"Срок",type:"string"},{key:"car",label:"Марка авто",type:"string"}];
const OPS_NUM=[{key:"eq",label:"="},{key:"gte",label:"≥"},{key:"lte",label:"≤"},{key:"between",label:"от–до"}];
const OPS_STR=[{key:"eq",label:"="},{key:"neq",label:"≠"},{key:"contains",label:"содержит"}];

const SEED_AGENTS=(()=>{
  const d=[
    ["768-01","Հakob","Vkhkryan","1606-01","","","","",""],
    ["768-03","Ռobert","Bichakhchyan (112)","1606-03","","","","",""],
    ["768-05","Lusine","Mosoyan (12+66)","1606-05","541-01","","","",""],
    ["768-07","Mkrtich","Baghdasaryan","1606-07","","310-71-10","797-25","",""],
    ["768-11","Srap","Khachatryan","1606-11","","","","",""],
    ["768-19","Grigor","Hovhannisyan (Natella)","1606-168","","","","",""],
    ["768-20","David","Minasyan","1606-20","","","","",""],
    ["768-24","Manvel","Nahapetyan","1606-24","","","","",""],
    ["768-27","Gagik","Vanoyan","1606-27","","","","",""],
    ["768-40","Levon","Vardanyan","1606-40","3671-07","310-71-11","797-26","A50-M3-40-G12","13021-09"],
    ["768-53","Vladimir","Shahinyan","1606-53","","","","",""],
    ["768-74","Armen","Khanoyan","1606-74","","310-71-09","797-23","A50-M3-40-G9","13021-07"],
    ["768-101","Nelli","Harutyunyan","1606-101","3671-01","310-71-01","797-16","A50-M3-40-G5","13021-01"],
    ["768-105","Viktorya","Vkhkryan","1606-105","3671-05","310-71-02","797-17","A50-M3-40-G3","13021-02"],
    ["768-106","Gayane","Aleksanyan","1606-106","3671-02","310-71-03","797-18","A50-M3-40-G111","13021-04"],
    ["768-115","Annman","Ghevondyan (Shant)","1606-115","","","","",""],
    ["768-116","Aghuniq","Petrosyan","1606-116","","","","",""],
    ["768-118","Hovhannes","Torosyan","1606-118","","","","",""],
    ["768-121","Artur","Safaryan","1606-121","","","","",""],
    ["768-122","Hovhannes","Djanoyan","1606-122","","","","",""],
    ["768-123","Tehmine","Kocharyan","1606-123","","","","",""],
    ["768-125","Narine","Arzumanyan","1606-125","101-1","101-1","101-1","A50-M3-40-G4","101-1"],
    ["768-127","Armen","Simonyan","1606-127","","","","",""],
    ["768-128","GURGEN","AREVIKYAN","1606-128","","310-71-08","797-22","A50-M3-40-G8","13021-06"],
    ["768-130","RITA","GRIGORYAN","1606-130","","","","",""],
    ["768-131","HRIPSIME","TOROSYAN","1606-131","","","","",""],
    ["768-132","Zhanna","Gasparyan","1606-132","","","797-27","","13021-12"],
    ["768-133","Romik","Nazaretyan (+145+164)","1606-133","","310-71-07","797-21","","13021-10"],
    ["768-135","Artak","Khachatryan","1606-135","","","","",""],
    ["768-137","ARTYOM","AVETISYAN (Artak)","1606-137","","","","",""],
    ["768-138","Zarzand","Brsoyan","1606-138","","","","",""],
    ["768-139","Aghasi","Sahakyan (Artak)","1606-139","","","","",""],
    ["768-142","Varduhi","Yayloyan","1606-142","","","","",""],
    ["768-144","Vram","Ayvazyan","1606-144","","","","",""],
    ["768-145","Amalya","Yephremyan","1606-145","","","","",""],
    ["768-146","Sargis","Hamazaspyan","1606-146","","","797-24","A50-M3-40-G10",""],
    ["768-147","Anushiq","Sargsyan","1606-147","","","","",""],
    ["768-148","Lusine","Ghazaryan","1606-148","","","","",""],
    ["768-149","Varduhi","Tadevosyan","1606-149","","","","",""],
    ["768-150","MARGARIT","MIKOYAN","1606-150","","","","",""],
    ["768-151","Sveta","Hasoyan","1606-151","","","","",""],
    ["768-152","Ashot","Martirosyan","1606-152","","","","",""],
    ["768-153","VARDAN","GHAZARYAN","1606-153","","","","",""],
    ["768-154","SUSANNA","GRIGORYAN","1606-154","","","","",""],
    ["768-155","ARINA","METSOYAN","1606-155","","","","",""],
    ["768-156","KARINE","TAMAZYAN (Shaboyan)","1606-156","","","","",""],
    ["768-157","ARKADI","TADEVOSYAN","1606-157","","","","",""],
    ["768-158","ARTAK","KHACHATRYAN (EDGARI)","1606-158","","","","",""],
    ["768-159","VAHAN","AGHABABYAN","1606-159","","310-71-13","","",""],
    ["768-160","Hovhannes","Hovhannisyan","1606-160","","","","",""],
    ["768-161","MARUSYA","GRIGORYAN","1606-161","","","","",""],
    ["768-164","Grigor","Gasparyan-Roman","1606-164","","","","",""],
    ["768-165","Yelena","","1606-165","","","","",""],
    ["768-166","ARTYOM","SARGSYAN","1606-166","","","","",""],
    ["768-170","ARMENUI","MGDESYAN","1606-170","","","","",""],
    ["768-171","TATYEVIK","KHACHATRYAN (Narine)","1606-171","","","","",""],
    ["768-172","Mariam","Avagyan (Yulia)","1606-172","","","","",""],
    ["768-173","Andranik","Hovhannisyan","1606-173","","","","",""],
    ["768-175","Ashkhen","Galoyan","1606-175","","","","",""],
    ["768-176","Rustam","Karapetyan (Nelli Reso)","1606-176","","","","",""],
    ["768-177","Aghavni","Stepanyan (Yelena)","1606-177","","","","",""],
    ["1606-178","Smbat","Nersisyan","1606-178","","","","",""],
    ["768-178","Mkrtich","Abajyan","1606-178-01","","","","",""],
    ["768-179","Hovhannes","Gevorgyan (Vanadzor)","1606-179","","","","",""],
    ["768-180","Vahan","Manukyan","1606-180","","","","",""],
    ["768-181","Naira","Tovmasyan","1606-181","","310-71-04","797-04","A50-M3-40-G1",""],
    ["768-183","Tigran","Manukyan","1606-183","","","","",""],
    ["768-184","Gevorg","Harutyunyan Shiro","1606-184","","","","",""],
    ["768-185","Karlen","Harutyunyan","1606-185","","","","",""],
    ["768-186","Nelli","Nikoyan-Maga","1606-186","3671-03","310-71-05","797-19","A50-M3-40-G2","13021-03"],
    ["768-187","ANIA","IGITYAN","1606-187","3671-111","310-71-06","797-20","A50-M3-40-G13","13021-111"],
    ["768-188","Hovhannes","Hovhannisyan JB","1606-188","","","","",""],
    ["768-190","Narine","Arzumanyan (2)","1606-190","3671-10","310-71-15","797-28","A50-M3-40-G14","13021-13"],
    ["768-191","Taghuhi","Gevorgyan","1606-191","","310-71-14","797-29","A50-M3-40-G15",""],
    ["768-192","Susanna","Harutyunyan (Rozik)","1606-192","","","797-30","",""],
  ];
  return Object.fromEntries(d.map(([ic,nm,sr,n,i,s,r,l,a])=>[`ag-${ic}`,{name:nm,surname:sr,internalCode:ic,codes:{Nairi:n,Ingo:i,Sil:s,Rego:r,Liga:l,Armenia:a}}]));
})();

function parseDate(raw){
  if(raw==null||raw==="")return null;
  if(raw instanceof Date)return isNaN(raw.getTime())?null:raw;
  const num=Number(raw);
  if(!isNaN(num)&&num>10000&&num<200000){const d=new Date(Date.UTC(1899,11,30)+num*86400000);return isNaN(d.getTime())?null:d;}
  const s=String(raw).trim();if(!s)return null;
  const m1=s.match(/^(\d{1,2})[.\-\/](\d{1,2})[.\-\/](\d{4})$/);if(m1)return new Date(+m1[3],+m1[2]-1,+m1[1]);
  const m2=s.match(/^(\d{4})[.\-\/](\d{1,2})[.\-\/](\d{1,2})$/);if(m2)return new Date(+m2[1],+m2[2]-1,+m2[3]);
  const d=new Date(s);return isNaN(d.getTime())?null:d;
}
const fmtDate=d=>!d?"":String(d.getDate()).padStart(2,"0")+"."+String(d.getMonth()+1).padStart(2,"0")+"."+d.getFullYear();
const getMonthKey=d=>!d?null:`${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,"0")}`;
const fmtMonth=key=>{if(!key)return"";const[y,m]=key.split("-").map(Number);return MONTHS_RU[m-1]+" "+y;};
const prevMo=key=>{const[y,m]=key.split("-").map(Number);return m===1?`${y-1}-12`:`${y}-${String(m-1).padStart(2,"0")}`;};
const nextMo=key=>{const[y,m]=key.split("-").map(Number);return m===12?`${y+1}-01`:`${y}-${String(m+1).padStart(2,"0")}`;};

const detectCo=name=>{const n=(name||"").toLowerCase();if(n.includes("nairi")||n.includes("наири"))return"Nairi";if(n.includes("ingo")||n.includes("инго"))return"Ingo";if(n.includes("liga")||n.includes("лига"))return"Liga";if(n.includes("sil")||n.includes("сил"))return"Sil";if(n.includes("rego")||n.includes("рего"))return"Rego";if(n.includes("armenia")||n.includes("армения"))return"Armenia";return null;};
const normReg=raw=>{
  const s=(raw||"").toLowerCase().trim();
  if(s==="yr"||s.includes("ереван")||s.includes("yerevan"))return"YR";
  if(s==="kt"||s.includes("котайк")||s.includes("kotayk")||s.includes("կոտայք"))return"KT";
  return(raw||"").toUpperCase().slice(0,10);
};
const isExcBrand=car=>EXCEPTION_BRANDS.some(b=>(car||"").toLowerCase().includes(b));

const HMAP={
  agentCode:["agentcode","agentCode"],
  policyNum:["policynum","policyNum"],
  amount:["amount"],
  region:["region"],
  bm:["bm"],
  power:["power"],
  term:["term"],
  car:["car"],
  insuredName:["insuredname","insuredName"],
  startDate:["startdate","startDate"],
  endDate:["enddate","endDate"],
  carPlate:["carplate","carPlate"],
  phone:["phone"],
  productName:["productname","productName"],
};

function detectCols(headers,dataRows){
  const map={};
  headers.forEach((h,i)=>{
    const n=(h||"").toLowerCase().trim();
    for(const[f,vs]of Object.entries(HMAP)){
      if(!(f in map)&&vs.some(v=>n.includes(v.toLowerCase())))map[f]=i;
    }
  });
  if(!("amount" in map)&&dataRows&&dataRows.length>0){
    const testRows=dataRows.slice(0,Math.min(10,dataRows.length));
    let bestCol=-1,bestScore=0;
    headers.forEach((_,i)=>{
      const score=testRows.filter(r=>{const v=parseFloat(String(r[i]||"").replace(/\s/g,"").replace(",","."));return !isNaN(v)&&v>100;}).length;
      if(score>bestScore){bestScore=score;bestCol=i;}
    });
    if(bestCol>=0&&bestScore>=2)map["amount"]=bestCol;
  }
  if((!("startDate" in map)||!("endDate" in map))&&dataRows&&dataRows.length>0){
    const testRows=dataRows.slice(0,Math.min(5,dataRows.length));
    const isDateLike=v=>{
      if(v instanceof Date)return true;
      const s=String(v||"").trim();
      if(/^\d{1,2}[.\-\/]\d{1,2}[.\-\/]\d{4}$/.test(s))return true;
      if(/^\d{4}[.\-\/]\d{1,2}[.\-\/]\d{1,2}$/.test(s))return true;
      const n=Number(s);return !isNaN(n)&&n>40000&&n<60000;
    };
    const dateCols=[];
    headers.forEach((_,i)=>{
      if(Object.values(map).includes(i))return;
      const score=testRows.filter(r=>isDateLike(r[i])).length;
      if(score>=2)dateCols.push({i,score});
    });
    dateCols.sort((a,b)=>b.score-a.score);
    if(!("startDate" in map)&&dateCols.length>0)map["startDate"]=dateCols[0].i;
    if(!("endDate" in map)&&dateCols.length>1)map["endDate"]=dateCols[1].i;
  }
  return map;
}

function parseCoRows(rows,company){
  if(!rows.length)return[];
  const cm=detectCols(rows[0],rows.slice(1));
  return rows.slice(1).filter(r=>r.some(c=>c!=="")).map((row,idx)=>{
    const getRaw=f=>cm[f]!==undefined?row[cm[f]]:null;
    const get=f=>{const v=getRaw(f);if(v instanceof Date)return v.toISOString();return String(v||"").trim();};
    const getCode=f=>{const raw=String(getRaw(f)||"").trim();return raw.split(/\s+-\s+/)[0].replace(/\s+/g,"").trim();};
    let sd=parseDate(getRaw("startDate"));
    let ed=parseDate(getRaw("endDate"));
    const termRaw=get("term");
    const termNum=parseInt(termRaw)||0;
    let days=null;
    if(sd&&ed){days=Math.round((ed.getTime()-sd.getTime())/86400000);}
    else if(company==="Armenia"&&termNum>0){days=termNum;if(sd)ed=new Date(sd.getTime()+termNum*86400000);}
    const termAuto=days!==null?(days<88?"SH":"L"):termRaw||"";
    const amt=parseFloat(String(get("amount")).replace(/\s/g,"").replace(",","."))||0;
    return{_id:`${company}-${idx}`,company,agentCode:getCode("agentCode"),policyNum:get("policyNum"),
      amount:amt,region:normReg(get("region")),bm:parseInt(get("bm"))||0,
      power:parseInt(get("power"))||0,term:termAuto,days,car:get("car"),
      insuredName:get("insuredName"),carPlate:getCode("carPlate"),phone:get("phone"),
      startDate:sd?sd.toISOString():null,endDate:ed?ed.toISOString():null,
      startDateFmt:fmtDate(sd),endDateFmt:fmtDate(ed)};
  }).filter(p=>p.policyNum||p.amount>0);
}

function parseVolRows(rows){
  if(!rows.length)return[];
  const cm=detectCols(rows[0],rows.slice(1));
  return rows.slice(1).filter(r=>r.some(c=>c!=="")).map((row,idx)=>{
    const get=f=>cm[f]!==undefined?String(row[cm[f]]||"").trim():"";
    return{_id:`vol-${idx}`,productName:get("productName"),company:get("company")||"",
      agentCode:get("agentCode"),policyNum:get("policyNum"),
      amount:parseFloat(String(get("amount")).replace(/\s/g,"").replace(",","."))||0,
      insuredName:get("insuredName")};
  }).filter(p=>p.amount>0||p.productName);
}

function parseOfficeRows(rows){
  if(rows.length<2)return{profit:0,rows:[],fileName:""};
  const headers=rows[0].map(h=>String(h).toLowerCase().trim());
  const profitIdx=headers.findIndex(h=>h.includes("прибыль")||h.includes("profit"));
  const companyIdx=headers.findIndex(h=>h.includes("компани")||h.includes("company"));
  const amountIdx=headers.findIndex(h=>h.includes("страховая")||h.includes("сумма")||h.includes("amount"));
  const parsed=rows.slice(1).map((r,i)=>({
    _id:"off-"+i,
    company:companyIdx>=0?String(r[companyIdx]||""):"",
    amount:parseFloat(String(r[amountIdx>=0?amountIdx:0]||"").replace(/\s/g,"").replace(",","."))||0,
    profit:profitIdx>=0?parseFloat(String(r[profitIdx]||"").replace(/\s/g,"").replace(",","."))||0:0,
  })).filter(r=>r.profit!==0||r.amount!==0);
  return{profit:parsed.reduce((s,r)=>s+r.profit,0),rows:parsed};
}

function checkCond(v,cond){
  const sv=String(v||"");const n=Number(v);const cv=Number(cond.value);const cv2=Number(cond.value2);
  if(cond.op==="eq")return sv===String(cond.value);if(cond.op==="neq")return sv!==String(cond.value);
  if(cond.op==="contains")return sv.toLowerCase().includes(String(cond.value).toLowerCase());
  if(cond.op==="gte")return n>=cv;if(cond.op==="lte")return n<=cv;if(cond.op==="between")return n>=cv&&n<=cv2;
  return false;
}
function checkExc(p,excepts,agentUid){
  return excepts.filter(e=>e.enabled&&e.company===p.company).some(exc=>{
    if(exc.excludedAgents&&exc.excludedAgents.length&&exc.excludedAgents.includes(agentUid))return false;
    if(exc.type==="brand_whitelist"){
      if(!isExcBrand(p.car))return false;
      if(p.term==="SH")return false;
      const regionOk=!p.region||p.region!=="YR";
      const bmCond=exc.conditions.find(c=>c.field==="bm");
      return regionOk&&(bmCond?checkCond(p.bm,bmCond):true);
    }
    return exc.conditions.every(c=>checkCond(p[c.field],c));
  });
}
const condLabel=c=>{
  const f={region:"Регион",bm:"БМ",power:"Мощность",term:"Срок",car:"Марка"}[c.field]||c.field;
  if(c.op==="eq")return f+" = "+c.value;if(c.op==="neq")return f+" ≠ "+c.value;
  if(c.op==="contains")return f+' содержит "'+c.value+'"';
  if(c.op==="gte")return f+" ≥ "+c.value;if(c.op==="lte")return f+" ≤ "+c.value;
  if(c.op==="between")return f+" "+c.value+"–"+c.value2;return"";
};
const excReason=(p,excepts,agentUid)=>{
  const exc=excepts.find(e=>e.enabled&&e.company===p.company&&!(e.excludedAgents&&e.excludedAgents.includes(agentUid))&&!(e.type==="brand_whitelist"&&(!isExcBrand(p.car)||p.term==="SH"))&&e.conditions.every(c=>checkCond(p[c.field],c)));
  if(!exc)return"—";const parts=exc.conditions.map(condLabel);if(exc.type==="brand_whitelist")parts.push("Марка: "+p.car);return parts.join(" + ");
};

const getArmGroup=bm=>bm<=9?"1-9":bm<=14?"10-14":"15-25";
const getAgentRate=(p,agentUid,rates)=>{
  if(agentUid&&rates.agentOverrides&&rates.agentOverrides[agentUid]&&rates.agentOverrides[agentUid][p.company]!==undefined)
    return rates.agentOverrides[agentUid][p.company];
  if(p.company==="Armenia"){
    const isShort=p.term==="SH"||(p.days!=null&&p.days<88);
    if(isShort)return(rates.armeniaShort&&rates.armeniaShort.agentRate)||20;
    return(rates.armeniaAgent&&rates.armeniaAgent[getArmGroup(p.bm)])||0;
  }
  if(p.company==="Nairi"&&(p.region==="YR"||p.region==="KT")){
    return p.bm>=1&&p.bm<=7?((rates.nairiRegion&&rates.nairiRegion.bm1_7)||10):((rates.nairiRegion&&rates.nairiRegion.bm8_25)||5);
  }
  return(rates.agentRates&&rates.agentRates[p.company])||0;
};
const getOfficeRate=(p,rates)=>{
  if(p.company==="Armenia"){
    const isShort=p.term==="SH"||(p.days!=null&&p.days<88);
    if(isShort)return(rates.armeniaShort&&rates.armeniaShort.officeRate)||37;
    return(rates.armeniaOffice&&rates.armeniaOffice[getArmGroup(p.bm)])||0;
  }
  return(rates.officeRates&&rates.officeRates[p.company])||0;
};

function exportToExcel(agentData,effVol,agentDir,totals,excepts){
  const wb=XLSX.utils.book_new();
  const gn=uid=>{const a=agentDir[uid];return a?(a.name+" "+a.surname).trim():uid||"";};
  const s1=[["Агент","768-код","Всего","Зачётные","Доход офиса","Выплачено агенту","Прибыль офиса"],
    ...agentData.map(a=>[a.uid,agentDir[a.uid]&&agentDir[a.uid].internalCode||"",a.totalSales,a.validSales,a.totalOffice,a.totalAgent,a.profit]),
    [],["ИТОГО","",agentData.reduce((s,a)=>s+a.totalSales,0),agentData.reduce((s,a)=>s+a.validSales,0),totals.office,totals.agent,totals.profit]];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(s1),"Сводка");
  const s2=[["Агент","768-код","№ полиса","Компания","Марка","Рег.номер","Страхователь","Телефон","Начало","Окончание","Дней","Срок","Сумма","Регион","БМ","Мощность","Статус","% Агент","Ком. Агент"]];
  agentData.forEach(a=>{
    const ic=(agentDir[a.uid]&&agentDir[a.uid].internalCode)||"";
    a.policies.forEach(p=>s2.push([gn(a.uid),ic,p.policyNum,p.company,p.car,p.carPlate,p.insuredName,p.phone,p.startDateFmt,p.endDateFmt,p.days!=null?p.days:"",p.term,p.amount,p.region,p.bm,p.power,p.exception?"Исключение":"Зачёт",(p.agentRate||0)+"%",p.agentComm]));
  });
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(s2),"Полисы");
  const s3=[["Агент","№ полиса","Компания","Марка","Сумма","Регион","БМ","Причина"]];
  agentData.forEach(a=>a.policies.filter(p=>p.exception).forEach(p=>s3.push([gn(a.uid),p.policyNum,p.company,p.car,p.amount,p.region,p.bm,excReason(p,excepts,a.uid)])));
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(s3),"Исключения");
  if(effVol.length>0){
    const s4=[["Агент","Продукт","Компания","Страхователь","Сумма","% О","Ком.О","% А","Ком.А"]];
    effVol.forEach(v=>s4.push([gn(v.agentUid),v.productName,v.company,v.insuredName,v.amount,(v.officeRate||0)+"%",v.officeComm,(v.agentRate||0)+"%",v.agentComm]));
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(s4),"Добровольные");
  }
  XLSX.writeFile(wb,"Комиссии.xlsx");
}

const th={padding:"7px 10px",fontWeight:600,fontSize:12,whiteSpace:"nowrap",color:"#374151",borderBottom:"2px solid #e5e7eb",textAlign:"left",background:"#f9fafb"};
const td={padding:"7px 10px",fontSize:13,borderBottom:"1px solid #f0f0f0"};
const inp={border:"1px solid #d1d5db",borderRadius:5,padding:"3px 7px",fontSize:13,boxSizing:"border-box"};
const btn=(bg,col,ex)=>({padding:"5px 12px",background:bg||"#2563eb",color:col||"#fff",border:"none",borderRadius:6,cursor:"pointer",fontSize:12,fontWeight:600,...(ex||{})});
const fmt=n=>Number(n).toLocaleString("ru-RU")+" ֏";
const genUid=()=>Math.random().toString(36).slice(2,8);

function RatesPanel({rates,onSave,agentDir}){
  const[local,setLocal]=useState(()=>JSON.parse(JSON.stringify(rates)));
  const set=(path,val)=>setLocal(prev=>{const next=JSON.parse(JSON.stringify(prev));const keys=path.split(".");let o=next;for(let i=0;i<keys.length-1;i++)o=o[keys[i]];o[keys[keys.length-1]]=val;return next;});
  const ni=(path,val,w)=><input type="number" value={val||0} onChange={e=>set(path,parseFloat(e.target.value)||0)} style={{...inp,width:w||65,textAlign:"center"}}/>;
  const addOv=id=>{if(!id||local.agentOverrides[id]!==undefined)return;setLocal(p=>({...p,agentOverrides:{...p.agentOverrides,[id]:Object.fromEntries(ALL_COMPANIES.map(c=>[c,0]))}}));};
  const rmOv=id=>setLocal(p=>{const next=JSON.parse(JSON.stringify(p));delete next.agentOverrides[id];return next;});
  const setOv=(id,c,val)=>setLocal(p=>({...p,agentOverrides:{...p.agentOverrides,[id]:{...p.agentOverrides[id],[c]:parseFloat(val)||0}}}));
  const[selAg,setSelAg]=useState("");
  const agWithout=Object.entries(agentDir).filter(([id])=>!local.agentOverrides[id]);
  return(
    <div style={{fontSize:13}}>
      <h4 style={{margin:"0 0 6px",color:"#1d4ed8"}}>% офиса</h4>
      <div style={{overflowX:"auto",marginBottom:16}}><table style={{borderCollapse:"collapse"}}><thead><tr>{[...COMPANIES,"Арм.1-9","Арм.10-14","Арм.15-25"].map(c=><th key={c} style={th}>{c}</th>)}</tr></thead><tbody><tr>{COMPANIES.map(c=><td key={c} style={td}>{ni("officeRates."+c,local.officeRates[c])}</td>)}{ARM_GROUPS.map(g=><td key={g} style={td}>{ni("armeniaOffice."+g,local.armeniaOffice[g])}</td>)}</tr></tbody></table></div>
      <h4 style={{margin:"0 0 4px",color:"#16a34a"}}>% агента — базовые</h4>
      <div style={{overflowX:"auto",marginBottom:16}}><table style={{borderCollapse:"collapse"}}><thead><tr>{[...COMPANIES,"Арм.1-9","Арм.10-14","Арм.15-25"].map(c=><th key={c} style={th}>{c}</th>)}</tr></thead><tbody><tr>{COMPANIES.map(c=><td key={c} style={td}>{ni("agentRates."+c,local.agentRates[c])}</td>)}{ARM_GROUPS.map(g=><td key={g} style={td}>{ni("armeniaAgent."+g,local.armeniaAgent[g])}</td>)}</tr></tbody></table></div>
      <h4 style={{margin:"0 0 4px",color:"#dc2626"}}>Armenia — короткий срок (до 87 дней)</h4>
      <p style={{margin:"0 0 8px",fontSize:11,color:"#6b7280"}}>Перекрывает стандартные ставки Armenia если срок меньше 88 дней.</p>
      <div style={{overflowX:"auto",marginBottom:16}}><table style={{borderCollapse:"collapse"}}><thead><tr><th style={th}>% офиса</th><th style={th}>% агента</th></tr></thead><tbody><tr><td style={td}>{ni("armeniaShort.officeRate",(local.armeniaShort&&local.armeniaShort.officeRate)||37)}</td><td style={td}>{ni("armeniaShort.agentRate",(local.armeniaShort&&local.armeniaShort.agentRate)||20)}</td></tr></tbody></table></div>
      <h4 style={{margin:"0 0 4px",color:"#d97706"}}>Nairi — регион YR / KT</h4>
      <p style={{margin:"0 0 8px",fontSize:11,color:"#6b7280"}}>Применяется если регион YR или KT. Индивидуальная ставка агента перекрывает это правило.</p>
      <div style={{overflowX:"auto",marginBottom:16}}><table style={{borderCollapse:"collapse"}}><thead><tr><th style={th}>БМ 1–7 (% агента)</th><th style={th}>БМ 8–25 (% агента)</th></tr></thead><tbody><tr><td style={td}>{ni("nairiRegion.bm1_7",(local.nairiRegion&&local.nairiRegion.bm1_7)||10)}</td><td style={td}>{ni("nairiRegion.bm8_25",(local.nairiRegion&&local.nairiRegion.bm8_25)||5)}</td></tr></tbody></table></div>
      <h4 style={{margin:"0 0 6px",color:"#7c3aed"}}>% агента — индивидуальные</h4>
      {Object.entries(local.agentOverrides).map(([id,r])=>{const a=agentDir[id];const nm=a?(a.name+" "+a.surname).trim():id;return(
        <div key={id} style={{border:"1px solid #e5e7eb",borderRadius:8,marginBottom:8,overflow:"hidden"}}>
          <div style={{background:"#f5f3ff",padding:"6px 12px",display:"flex",justifyContent:"space-between",alignItems:"center"}}><span style={{fontWeight:600,color:"#6d28d9",fontSize:13}}>{"👤 "+nm}</span><button onClick={()=>rmOv(id)} style={btn("#fff1f2","#dc2626",{border:"1px solid #fca5a5",fontSize:11})}>✕</button></div>
          <div style={{padding:"8px 12px",overflowX:"auto"}}><table style={{borderCollapse:"collapse"}}><thead><tr>{ALL_COMPANIES.map(c=><th key={c} style={{...th,background:"white"}}>{c}</th>)}</tr></thead><tbody><tr>{ALL_COMPANIES.map(c=><td key={c} style={td}><input type="number" value={r[c]||0} onChange={e=>setOv(id,c,e.target.value)} style={{...inp,width:62,textAlign:"center"}}/></td>)}</tr></tbody></table></div>
        </div>
      );})}
      {agWithout.length>0&&<div style={{display:"flex",gap:8,alignItems:"center",marginBottom:12}}><select value={selAg} onChange={e=>setSelAg(e.target.value)} style={{...inp,padding:"4px 8px",fontSize:12}}><option value="">Выберите агента...</option>{agWithout.map(([id,a])=><option key={id} value={id}>{a.name+" "+a.surname}</option>)}</select><button onClick={()=>{addOv(selAg);setSelAg("");}} style={btn("#7c3aed")} disabled={!selAg}>+ Добавить</button></div>}
      <div style={{display:"flex",gap:8}}><button onClick={()=>onSave(local)} style={btn("#16a34a")}>💾 Сохранить</button><button onClick={()=>setLocal(JSON.parse(JSON.stringify(DEFAULT_RATES)))} style={btn("#f3f4f6","#374151")}>Сбросить</button></div>
    </div>
  );
}

function VolRatesPanel({volRates,onSave}){
  const[local,setLocal]=useState(()=>JSON.parse(JSON.stringify(volRates)));
  const add=()=>setLocal(p=>({...p,rates:[...p.rates,{id:genUid(),name:"",officeRate:0,agentRate:0}]}));
  const rm=id=>setLocal(p=>({...p,rates:p.rates.filter(r=>r.id!==id)}));
  const upd=(id,f,v)=>setLocal(p=>({...p,rates:p.rates.map(r=>r.id===id?{...r,[f]:v}:r)}));
  return(
    <div style={{fontSize:13}}>
      {local.rates.length===0&&<p style={{color:"#9ca3af"}}>Нет продуктов.</p>}
      {local.rates.map(r=>(
        <div key={r.id} style={{display:"flex",gap:8,alignItems:"center",marginBottom:8,flexWrap:"wrap"}}>
          <input value={r.name} onChange={e=>upd(r.id,"name",e.target.value)} placeholder="Название продукта" style={{...inp,width:200,padding:"4px 8px"}}/>
          <label style={{display:"flex",alignItems:"center",gap:4,fontSize:12,color:"#6b7280"}}>Офис %<input type="number" value={r.officeRate} onChange={e=>upd(r.id,"officeRate",parseFloat(e.target.value)||0)} style={{...inp,width:60,textAlign:"center",marginLeft:4}}/></label>
          <label style={{display:"flex",alignItems:"center",gap:4,fontSize:12,color:"#6b7280"}}>Агент %<input type="number" value={r.agentRate} onChange={e=>upd(r.id,"agentRate",parseFloat(e.target.value)||0)} style={{...inp,width:60,textAlign:"center",marginLeft:4}}/></label>
          <button onClick={()=>rm(r.id)} style={btn("#fff1f2","#dc2626",{border:"1px solid #fca5a5"})}>✕</button>
        </div>
      ))}
      <div style={{display:"flex",gap:8,marginTop:8}}><button onClick={add} style={btn("#6366f1")}>+ Добавить</button><button onClick={()=>onSave(local)} style={btn("#16a34a")}>💾 Сохранить</button></div>
    </div>
  );
}

function ExceptionsPanel({exceptions,onSave,agentDir}){
  const[local,setLocal]=useState(()=>JSON.parse(JSON.stringify(exceptions)));
  const[expanded,setExpanded]=useState(null);
  const[newCond,setNewCond]=useState({field:"region",op:"eq",value:"",value2:""});
  const toggle=id=>setLocal(p=>p.map(e=>e.id===id?{...e,enabled:!e.enabled}:e));
  const rmExc=id=>setLocal(p=>p.filter(e=>e.id!==id));
  const rmCond=(eid,ci)=>setLocal(p=>p.map(e=>e.id===eid?{...e,conditions:e.conditions.filter((_,i)=>i!==ci)}:e));
  const addExc=()=>{const e={id:genUid(),company:"Nairi",enabled:true,conditions:[],excludedAgents:[]};setLocal(p=>[...p,e]);setExpanded(e.id);};
  const setCo=(id,co)=>setLocal(p=>p.map(e=>e.id===id?{...e,company:co}:e));
  const addCond=id=>{
    if(!newCond.value)return;
    const c={field:newCond.field,op:newCond.op,value:newCond.value};
    if(newCond.op==="between")c.value2=newCond.value2;
    setLocal(p=>p.map(e=>e.id===id?{...e,conditions:[...e.conditions,c]}:e));
    setNewCond({field:"region",op:"eq",value:"",value2:""});
  };
  const togExcl=(eid,aUid)=>setLocal(p=>p.map(e=>{if(e.id!==eid)return e;const cur=e.excludedAgents||[];return{...e,excludedAgents:cur.includes(aUid)?cur.filter(x=>x!==aUid):[...cur,aUid]};}));
  const fldType=f=>(FIELDS.find(x=>x.key===f)||{type:"string"}).type;
  const ops=f=>fldType(f)==="number"?OPS_NUM:OPS_STR;
  return(
    <div style={{fontSize:13}}>
      {local.length===0&&<p style={{color:"#9ca3af"}}>Нет исключений.</p>}
      {local.map(exc=>(
        <div key={exc.id} style={{border:"1px solid #e5e7eb",borderRadius:8,marginBottom:10,overflow:"hidden"}}>
          <div style={{display:"flex",alignItems:"center",gap:10,padding:"8px 12px",background:exc.enabled?"white":"#f9fafb",flexWrap:"wrap"}}>
            <label style={{display:"flex",alignItems:"center",gap:5,cursor:"pointer"}}><input type="checkbox" checked={exc.enabled} onChange={()=>toggle(exc.id)} style={{width:15,height:15}}/><span style={{fontSize:11,color:exc.enabled?"#374151":"#9ca3af"}}>{exc.enabled?"Активно":"Выкл."}</span></label>
            <select value={exc.company} onChange={e=>setCo(exc.id,e.target.value)} style={{...inp,padding:"3px 6px",fontSize:12}}>{ALL_COMPANIES.map(c=><option key={c}>{c}</option>)}</select>
            {exc.type==="brand_whitelist"&&<span style={{background:"#fef3c7",color:"#92400e",borderRadius:12,padding:"2px 8px",fontSize:11}}>🚗 Opel/Mercedes/BMW/NISSAN</span>}
            <div style={{flex:1,display:"flex",flexWrap:"wrap",gap:5}}>{exc.conditions.length===0?<span style={{color:"#9ca3af",fontSize:12}}>нет условий</span>:exc.conditions.map((c,i)=><span key={i} style={{background:"#dbeafe",color:"#1d4ed8",borderRadius:12,padding:"2px 8px",fontSize:12}}>{condLabel(c)}</span>)}</div>
            {(exc.excludedAgents||[]).length>0&&<span style={{fontSize:11,color:"#16a34a",background:"#f0fdf4",padding:"2px 8px",borderRadius:12}}>{"✓ Белый список: "+exc.excludedAgents.length}</span>}
            <div style={{display:"flex",gap:6,marginLeft:"auto"}}>
              <button onClick={()=>setExpanded(expanded===exc.id?null:exc.id)} style={btn("#f3f4f6","#374151")}>{expanded===exc.id?"Скрыть":"✎ Изменить"}</button>
              <button onClick={()=>rmExc(exc.id)} style={btn("#fff1f2","#dc2626",{border:"1px solid #fca5a5"})}>✕</button>
            </div>
          </div>
          {expanded===exc.id&&(
            <div style={{padding:"12px",background:"#f9fafb",borderTop:"1px solid #e5e7eb"}}>
              <div style={{marginBottom:8,display:"flex",flexWrap:"wrap",gap:6}}>{exc.conditions.map((c,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:4,background:"white",border:"1px solid #e5e7eb",borderRadius:6,padding:"4px 8px",fontSize:12}}><span>{condLabel(c)}</span><button onClick={()=>rmCond(exc.id,i)} style={{background:"none",border:"none",cursor:"pointer",color:"#dc2626",fontSize:14}}>×</button></div>)}{exc.conditions.length===0&&<span style={{color:"#9ca3af",fontSize:12}}>Нет условий</span>}</div>
              <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center",marginBottom:14}}>
                <select value={newCond.field} onChange={e=>setNewCond({field:e.target.value,op:"eq",value:"",value2:""})} style={{...inp,padding:"4px 6px",fontSize:12}}>{FIELDS.map(f=><option key={f.key} value={f.key}>{f.label}</option>)}</select>
                <select value={newCond.op} onChange={e=>setNewCond(p=>({...p,op:e.target.value}))} style={{...inp,padding:"4px 6px",fontSize:12}}>{ops(newCond.field).map(o=><option key={o.key} value={o.key}>{o.label}</option>)}</select>
                <input value={newCond.value} onChange={e=>setNewCond(p=>({...p,value:e.target.value}))} placeholder="Значение" style={{...inp,width:110,padding:"4px 7px",fontSize:12}}/>
                {newCond.op==="between"&&<><span style={{fontSize:12,color:"#6b7280"}}>до</span><input value={newCond.value2} onChange={e=>setNewCond(p=>({...p,value2:e.target.value}))} placeholder="До" style={{...inp,width:70,padding:"4px 7px",fontSize:12}}/></>}
                <button onClick={()=>addCond(exc.id)} style={btn()}>+ Добавить</button>
              </div>
              <div style={{borderTop:"1px solid #e5e7eb",paddingTop:12}}>
                <div style={{fontWeight:600,fontSize:12,marginBottom:6}}>Агенты, для которых исключение НЕ действует:</div>
                {Object.keys(agentDir).length===0?<p style={{color:"#9ca3af",fontSize:12,margin:0}}>Сначала добавьте агентов.</p>:
                <div style={{display:"flex",flexWrap:"wrap",gap:6}}>{Object.entries(agentDir).map(([aUid,a])=>{const checked=(exc.excludedAgents||[]).includes(aUid);return(<label key={aUid} style={{display:"flex",alignItems:"center",gap:4,cursor:"pointer",background:checked?"#dcfce7":"white",border:"1px solid "+(checked?"#86efac":"#e5e7eb"),borderRadius:20,padding:"3px 10px",fontSize:12}}><input type="checkbox" checked={checked} onChange={()=>togExcl(exc.id,aUid)} style={{width:12,height:12}}/>{a.name+" "+a.surname}</label>);})}</div>}
              </div>
            </div>
          )}
        </div>
      ))}
      <div style={{display:"flex",gap:8,marginTop:12}}><button onClick={addExc} style={btn("#6366f1")}>+ Новое исключение</button><button onClick={()=>onSave(local)} style={btn("#16a34a")}>💾 Сохранить</button><button onClick={()=>setLocal(JSON.parse(JSON.stringify(DEFAULT_EXCEPTIONS)))} style={btn("#f3f4f6","#374151")}>Сбросить</button></div>
    </div>
  );
}

const valR=r=>{try{return r&&r.officeRates&&r.agentRates;}catch{return false;}};
const valE=e=>{try{return Array.isArray(e)&&e.every(x=>x.id&&x.company&&typeof x.enabled==="boolean"&&Array.isArray(x.conditions));}catch{return false;}};
const valD=d=>{try{return d&&typeof d==="object"&&!Array.isArray(d);}catch{return false;}};
const valV=v=>{try{return v&&Array.isArray(v.rates);}catch{return false;}};

export default function App(){
  const[tab,setTab]=useState("commissions");
  const[selMonth,setSelMonth]=useState(CURRENT_MONTH);
  const[uploadedFiles,setUploadedFiles]=useState([]);
  const[storedPols,setStoredPols]=useState([]);
  const[volSession,setVolSession]=useState([]);
  const[storedVol,setStoredVol]=useState([]);
  const[officeSession,setOfficeSession]=useState(null);
  const[storedOffice,setStoredOffice]=useState({profit:0,rows:[]});
  const[dbPols,setDbPols]=useState([]);
  const[dbLoaded,setDbLoaded]=useState(false);
  const[expiryF,setExpiryF]=useState(CURRENT_MONTH);
  const[activeAgent,setActiveAgent]=useState(null);
  const[agentDir,setAgentDir]=useState({});
  const[rates,setRates]=useState(DEFAULT_RATES);
  const[volRates,setVolRates]=useState(DEFAULT_VOL_RATES);
  const[exceptions,setExceptions]=useState(DEFAULT_EXCEPTIONS);
  const[panel,setPanel]=useState(null);
  const[showBackup,setShowBackup]=useState(false);
  const[savedOk,setSavedOk]=useState(false);
  const[newName,setNewName]=useState("");const[newSur,setNewSur]=useState("");const[newIC,setNewIC]=useState("");
  const[newCodes,setNewCodes]=useState(Object.fromEntries(ALL_COMPANIES.map(c=>[c,""])));
  const[editUid,setEditUid]=useState(null);
  const[uploadingFor,setUploadingFor]=useState(null);
  const fileRef=useRef();const backRef=useRef();const officeFileRef=useRef();

  useEffect(()=>{(async()=>{
    try{const r=await supabaseStorage.get("agentDirectory").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(valD(p))setAgentDir(p);}else{setAgentDir(SEED_AGENTS);supabaseStorage.set("agentDirectory",JSON.stringify(SEED_AGENTS)).catch(()=>{});}}catch{setAgentDir(SEED_AGENTS);}
    try{const r=await supabaseStorage.get("ratesConfig").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(valR(p))setRates(p);}}catch{}
    try{const r=await supabaseStorage.get("volRates").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(valV(p))setVolRates(p);}}catch{}
    try{const r=await supabaseStorage.get("exceptionsConfig").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(valE(p))setExceptions(p);}}catch{}
  })();},[]);

  useEffect(()=>{
    setUploadedFiles([]);setVolSession([]);setOfficeSession(null);setActiveAgent(null);
    (async()=>{
      try{const r=await supabaseStorage.get("month:"+selMonth).catch(()=>null);if(r&&r.value){const d=JSON.parse(r.value);setStoredPols(d.policies||[]);setStoredVol(d.voluntary||[]);}else{setStoredPols([]);setStoredVol([]);}}catch{setStoredPols([]);setStoredVol([]);}
      try{const r=await supabaseStorage.get("officeStore:"+selMonth).catch(()=>null);if(r&&r.value)setStoredOffice(JSON.parse(r.value));else setStoredOffice({profit:0,rows:[]});}catch{setStoredOffice({profit:0,rows:[]});}
    })();
  },[selMonth]);

  const saveDir=d=>{setAgentDir(d);supabaseStorage.set("agentDirectory",JSON.stringify(d)).catch(()=>{});};
  const saveRates=r=>{setRates(r);supabaseStorage.set("ratesConfig",JSON.stringify(r)).catch(()=>{});};
  const saveVR=r=>{setVolRates(r);supabaseStorage.set("volRates",JSON.stringify(r)).catch(()=>{});};
  const saveExcs=e=>{setExceptions(e);supabaseStorage.set("exceptionsConfig",JSON.stringify(e)).catch(()=>{});};
  const getName=id=>{const a=agentDir[id];return a?(a.name+" "+a.surname).trim():"";};

  const addAgent=()=>{
    if(!newName.trim())return;
    const id=genUid();
    saveDir({...agentDir,[id]:{name:newName.trim(),surname:newSur.trim(),internalCode:newIC.trim(),codes:{...newCodes}}});
    setNewName("");setNewSur("");setNewIC("");setNewCodes(Object.fromEntries(ALL_COMPANIES.map(c=>[c,""])));
  };
  const rmAgent=id=>{const d={...agentDir};delete d[id];saveDir(d);};

  const codeLookup=useMemo(()=>{
    const map={};
    Object.entries(agentDir).forEach(([aUid,agent])=>{
      Object.entries(agent.codes||{}).forEach(([co,code])=>{if(code&&code.trim())map[co+":"+code.replace(/\s+/g,"").trim()]=aUid;});
    });
    return map;
  },[agentDir]);

  const handleSlotClick=co=>{setUploadingFor(co);fileRef.current.click();};
  const handleFile=e=>{
    const file=e.target.files[0];if(!file)return;
    const co=uploadingFor||detectCo(file.name)||"";
    const reader=new FileReader();
    reader.onload=evt=>{
      const wb=XLSX.read(evt.target.result,{type:"array",cellDates:true});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""}).filter(r=>r.some(c=>c!=="")).map(r=>r.map(c=>c instanceof Date?c:String(c)));
      if(co==="voluntary"){setVolSession(parseVolRows(rows));}
      else{setUploadedFiles(prev=>[...prev.filter(f=>f.company!==co),{id:genUid(),fileName:file.name,company:co,rows}]);}
      setActiveAgent(null);
    };
    reader.readAsArrayBuffer(file);
    e.target.value="";setUploadingFor(null);
  };
  const removeFile=co=>{if(co==="voluntary")setVolSession([]);else setUploadedFiles(p=>p.filter(f=>f.company!==co));};

  const handleOfficeFile=e=>{
    const file=e.target.files[0];if(!file)return;
    const reader=new FileReader();
    reader.onload=evt=>{
      const wb=XLSX.read(evt.target.result,{type:"array",cellDates:true});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""}).filter(r=>r.some(c=>c!=="")).map(r=>r.map(c=>c instanceof Date?c:String(c)));
      const result=parseOfficeRows(rows);
      result.fileName=file.name;
      setOfficeSession(result);
    };
    reader.readAsArrayBuffer(file);
    e.target.value="";
  };

  const officeData=officeSession||storedOffice;

  const sessionPols=useMemo(()=>{
    try{
      const result=[];
      uploadedFiles.forEach(({rows,company})=>{
        if(!company||!rows.length)return;
        parseCoRows(rows,company).forEach(p=>{
          const agentUid=codeLookup[company+":"+p.agentCode]||null;
          result.push({...p,agentUid,exception:checkExc(p,exceptions,agentUid)});
        });
      });
      return result;
    }catch{return[];}
  },[uploadedFiles,exceptions,codeLookup]);

  const effPols=useMemo(()=>{
    if(uploadedFiles.length>0)return sessionPols;
    return storedPols.map(p=>{const aUid=p.agentUid||codeLookup[p.company+":"+p.agentCode]||null;return{...p,agentUid:aUid,exception:checkExc(p,exceptions,aUid)};});
  },[uploadedFiles,sessionPols,storedPols,exceptions,codeLookup]);

  const effVol=useMemo(()=>{
    const vols=volSession.length>0?volSession:storedVol;
    return vols.map(v=>{
      const aUid=codeLookup[v.company+":"+v.agentCode]||null;
      const vr=(volRates.rates||[]).find(r=>r.name===v.productName);
      const oR=vr?vr.officeRate:0;const aR=vr?vr.agentRate:0;
      return{...v,agentUid:aUid,officeRate:oR,agentRate:aR,officeComm:Math.round(v.amount*oR/100),agentComm:Math.round(v.amount*aR/100)};
    });
  },[volSession,storedVol,codeLookup,volRates]);

  const warnings=useMemo(()=>{
    const w=[];
    const noA=effPols.filter(p=>!p.agentUid);
    if(noA.length){const codes=[...new Set(noA.map(p=>p.company+":"+p.agentCode))].slice(0,5);w.push(noA.length+" полисов без агента: "+codes.join(", "));}
    const nums=effPols.map(p=>p.policyNum).filter(Boolean);
    const dup=[...new Set(nums.filter((n,i)=>nums.indexOf(n)!==i))];
    if(dup.length)w.push("Дубли ("+dup.length+"): "+dup.slice(0,5).join(", ")+(dup.length>5?"…":""));
    return w;
  },[effPols]);

  const agentData=useMemo(()=>{
    try{
      const map={};
      effPols.forEach(p=>{const k=p.agentUid||("__raw__"+p.company+":"+p.agentCode);if(!map[k])map[k]=[];map[k].push(p);});
      return Object.entries(map).map(([key,pols])=>{
        const valid=pols.filter(p=>!p.exception);
        const totalSales=pols.reduce((s,p)=>s+p.amount,0);
        const validSales=valid.reduce((s,p)=>s+p.amount,0);
        const enriched=pols.map(p=>{
          if(p.exception)return{...p,officeRate:0,agentRate:0,officeComm:0,agentComm:0};
          const oR=getOfficeRate(p,rates);
          const aR=getAgentRate(p,key,rates);
          return{...p,officeRate:oR,agentRate:aR,officeComm:Math.round(p.amount*oR/100),agentComm:Math.round(p.amount*aR/100)};
        });
        const totalOffice=enriched.reduce((s,p)=>s+p.officeComm,0);
        const totalAgent=enriched.reduce((s,p)=>s+p.agentComm,0);
        return{uid:key,policies:enriched,totalSales,validSales,totalOffice,totalAgent,profit:totalOffice-totalAgent};
      });
    }catch{return[];}
  },[effPols,rates]);

  const totals=useMemo(()=>({
    office:agentData.reduce((s,a)=>s+a.totalOffice,0)+effVol.reduce((s,v)=>s+v.officeComm,0),
    agent:agentData.reduce((s,a)=>s+a.totalAgent,0)+effVol.reduce((s,v)=>s+v.agentComm,0),
    profit:agentData.reduce((s,a)=>s+a.profit,0)+effVol.reduce((s,v)=>s+v.officeComm-v.agentComm,0),
  }),[agentData,effVol]);

  const allExcs=useMemo(()=>agentData.flatMap(a=>a.policies.filter(p=>p.exception).map(p=>({...p,agentUid:a.uid}))),[agentData]);
  const detail=agentData.find(a=>a.uid===activeAgent);

  const saveMonth=()=>{
    const pols=agentData.flatMap(a=>a.policies);
    supabaseStorage.set("month:"+selMonth,JSON.stringify({policies:pols,voluntary:effVol})).catch(()=>{});
    if(officeSession)supabaseStorage.set("officeStore:"+selMonth,JSON.stringify(officeSession)).catch(()=>{});
    setSavedOk(true);setTimeout(()=>setSavedOk(false),2500);
  };

  const loadDB=()=>{
    setDbLoaded(false);
    supabaseStorage.list("month:").then(res=>{
      if(!res||!res.keys||!res.keys.length){setDbPols([]);setDbLoaded(true);return;}
      const all=[];let done=0;
      res.keys.forEach(key=>{
        supabaseStorage.get(key).then(r=>{
          if(r&&r.value){try{const d=JSON.parse(r.value);const mk=key.replace("month:","");(d.policies||[]).forEach(p=>all.push({...p,_monthKey:mk}));}catch{}}
          done++;if(done===res.keys.length){setDbPols(all);setDbLoaded(true);}
        }).catch(()=>{done++;if(done===res.keys.length){setDbPols(all);setDbLoaded(true);}});
      });
    }).catch(()=>{setDbPols([]);setDbLoaded(true);});
  };

  const filteredDB=useMemo(()=>{
    if(!expiryF)return dbPols;
    return dbPols.filter(p=>{if(!p.endDate)return false;const d=new Date(p.endDate);return!isNaN(d.getTime())&&getMonthKey(d)===expiryF;});
  },[dbPols,expiryF]);

  const importBackup=e=>{
    const f=e.target.files[0];if(!f)return;
    const r=new FileReader();
    r.onload=evt=>{try{const d=JSON.parse(evt.target.result);if(d.agentDir&&valD(d.agentDir))saveDir(d.agentDir);if(d.rates&&valR(d.rates))saveRates(d.rates);if(d.volRates&&valV(d.volRates))saveVR(d.volRates);if(d.exceptions&&valE(d.exceptions))saveExcs(d.exceptions);alert("Восстановлено.");}catch{alert("Ошибка файла.");}};
    r.readAsText(f);e.target.value="";
  };
  const agName=a=>{const n=getName(a.uid);if(n)return n;if(a.uid.startsWith("__raw__"))return a.uid.replace("__raw__","")+" (нет в справочнике)";return a.uid;};

  const exportPayroll=(agentData,agentDir,month)=>{
    const wb=XLSX.utils.book_new();
    const gn=uid=>{const a=agentDir[uid];return a?(a.name+" "+a.surname).trim():uid||"";};
    const header1=["Имя агента","768-код",...ALL_COMPANIES.flatMap(c=>[c,""]),"Итого","Подпись"];
    const header2=["","", ...ALL_COMPANIES.flatMap(()=>["кол-во","сумма"]),"",""];
    const rows=[header1,header2];
    agentData.filter(a=>a.totalAgent>0).forEach(a=>{
      const ic=(agentDir[a.uid]&&agentDir[a.uid].internalCode)||"";
      const row=[gn(a.uid),ic];
      ALL_COMPANIES.forEach(c=>{
        const pols=a.policies.filter(p=>p.company===c&&!p.exception);
        const sum=pols.reduce((s,p)=>s+p.agentComm,0);
        row.push(pols.length>0?pols.length:"",sum>0?sum:"");
      });
      row.push(a.totalAgent,"");
      rows.push(row);
    });
    const ws=XLSX.utils.aoa_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb,ws,"Начисления "+fmtMonth(month));
    XLSX.writeFile(wb,"Начисления_"+month+".xlsx");
  };
  const hasData=agentData.length>0||effVol.length>0;
  const panels=[["agents","👤 Агенты",Object.keys(agentDir).length],["rates","⚙️ Ставки",null],["volrates","📦 Доброволь.",volRates.rates.length],["exceptions","🚫 Исключения",exceptions.filter(e=>e.enabled).length]];
  const backupJson=JSON.stringify({version:6,agentDir,rates,volRates,exceptions},null,2);

  return(
    <div style={{fontFamily:"system-ui,sans-serif",padding:20,maxWidth:1400,margin:"0 auto",color:"#111"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
        <h2 style={{margin:0,fontSize:20}}>Калькулятор комиссий</h2>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {panels.map(([id,label,count])=>(
            <button key={id} onClick={()=>setPanel(panel===id?null:id)} style={{...btn(panel===id?"#eff6ff":"#f9fafb",panel===id?"#1d4ed8":"#374151"),border:"1px solid #d1d5db",fontWeight:500}}>
              {label}{count!=null?" ("+count+")":""}
            </button>
          ))}
          {hasData&&<button onClick={()=>exportToExcel(agentData,effVol,agentDir,totals,exceptions)} style={btn("#16a34a")}>⬇ Excel</button>}
          <button onClick={()=>setShowBackup(p=>!p)} style={btn("#7c3aed")}>💾 Резерв</button>
          <button onClick={()=>backRef.current.click()} style={btn("#f3f4f6","#374151",{border:"1px solid #d1d5db"})}>📂 Восстановить</button>
          <input ref={backRef} type="file" accept=".json" onChange={importBackup} style={{display:"none"}}/>
        </div>
      </div>

      <div style={{display:"flex",borderBottom:"2px solid #e5e7eb",marginBottom:16,gap:0}}>
        {[["commissions","💰 Комиссии"],["policydb","📋 База полисов"],["officesales","🏢 Продажи офиса"],["payroll","📝 Начисления"]].map(([id,label])=>(
          <button key={id} onClick={()=>{setTab(id);if(id==="policydb")loadDB();}}
            style={{...btn(tab===id?"#2563eb":"transparent",tab===id?"#fff":"#6b7280"),borderRadius:"6px 6px 0 0",fontSize:14,padding:"9px 22px",marginBottom:"-2px",border:"2px solid "+(tab===id?"#2563eb":"transparent"),fontWeight:tab===id?700:400}}>
            {label}
          </button>
        ))}
      </div>

      {panel&&(
        <div style={{border:"1px solid #e5e7eb",borderRadius:8,padding:16,marginBottom:16,background:"#fafafa"}}>
          {panel==="agents"&&(
            <div>
              <h3 style={{margin:"0 0 10px",fontSize:15}}>Справочник агентов</h3>
              <div style={{background:"#eff6ff",border:"1px solid #bfdbfe",borderRadius:8,padding:"10px 14px",marginBottom:12,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:8}}>
                <span style={{fontSize:12,color:"#1d4ed8"}}>📋 База 768-XX: 74 агента</span>
                <div style={{display:"flex",gap:8}}>
                  <button onClick={()=>saveDir({...SEED_AGENTS,...agentDir})} style={btn("#1d4ed8")}>⬇ Загрузить / дополнить</button>
                  <button onClick={()=>{if(window.confirm("Заменить весь справочник базой 768-XX?"))saveDir({...SEED_AGENTS});}} style={btn("#f3f4f6","#374151",{border:"1px solid #d1d5db"})}>Заменить</button>
                </div>
              </div>
              <div style={{background:"white",border:"1px solid #e5e7eb",borderRadius:8,padding:12,marginBottom:12}}>
                <div style={{fontWeight:600,fontSize:13,marginBottom:8}}>Новый агент</div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:8}}>
                  <input value={newIC} onChange={e=>setNewIC(e.target.value)} placeholder="Внутр. код" style={{...inp,width:110,padding:"4px 8px",fontWeight:600,color:"#6366f1"}}/>
                  <input value={newName} onChange={e=>setNewName(e.target.value)} placeholder="Имя" style={{...inp,width:120,padding:"4px 8px"}}/>
                  <input value={newSur} onChange={e=>setNewSur(e.target.value)} placeholder="Фамилия" style={{...inp,width:140,padding:"4px 8px"}}/>
                </div>
                <div style={{fontSize:12,color:"#6b7280",marginBottom:6}}>Коды по компаниям:</div>
                <div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:10}}>
                  {ALL_COMPANIES.map(c=><div key={c} style={{display:"flex",flexDirection:"column",gap:3}}><span style={{fontSize:11,color:"#6b7280",fontWeight:600}}>{c}</span><input value={newCodes[c]||""} onChange={e=>setNewCodes(p=>({...p,[c]:e.target.value}))} placeholder="Код" style={{...inp,width:80,padding:"3px 6px",fontSize:12}}/></div>)}
                </div>
                <button onClick={addAgent} style={btn()}>Добавить</button>
              </div>
              {Object.keys(agentDir).length===0?<p style={{color:"#9ca3af",fontSize:13}}>Справочник пуст.</p>:
                <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead><tr><th style={{...th,color:"#6366f1"}}>Внутр. код</th><th style={th}>Имя</th><th style={th}>Фамилия</th>{ALL_COMPANIES.map(c=><th key={c} style={th}>{c}</th>)}<th style={th}></th></tr></thead>
                  <tbody>{Object.entries(agentDir).map(([aUid,ag])=>{
                    const ic=(ag.internalCode)||"";const cd=(ag.codes)||{};const name=ag.name||"";const surname=ag.surname||"";
                    return(<tr key={aUid}>{editUid===aUid?<>
                      <td style={td}><input defaultValue={ic} id={"ic-"+aUid} style={{...inp,width:80,padding:"3px 7px",fontWeight:600,color:"#6366f1"}}/></td>
                      <td style={td}><input defaultValue={name} id={"n-"+aUid} style={{...inp,width:90,padding:"3px 7px"}}/></td>
                      <td style={td}><input defaultValue={surname} id={"s-"+aUid} style={{...inp,width:100,padding:"3px 7px"}}/></td>
                      {ALL_COMPANIES.map(c=><td key={c} style={td}><input defaultValue={cd[c]||""} id={"c-"+aUid+"-"+c} style={{...inp,width:75,padding:"3px 6px",fontSize:12}}/></td>)}
                      <td style={td}>
                        <button onClick={()=>{saveDir({...agentDir,[aUid]:{internalCode:document.getElementById("ic-"+aUid).value,name:document.getElementById("n-"+aUid).value,surname:document.getElementById("s-"+aUid).value,codes:Object.fromEntries(ALL_COMPANIES.map(c=>[c,document.getElementById("c-"+aUid+"-"+c).value]))}});setEditUid(null);}} style={{...btn("#16a34a"),marginRight:6}}>✓</button>
                        <button onClick={()=>setEditUid(null)} style={btn("#f3f4f6","#374151")}>✕</button>
                      </td>
                    </>:<>
                      <td style={{...td,fontWeight:700,color:ic?"#6366f1":"#d1d5db"}}>{ic||"—"}</td>
                      <td style={td}>{name}</td><td style={td}>{surname}</td>
                      {ALL_COMPANIES.map(c=><td key={c} style={{...td,color:cd[c]?"#111":"#d1d5db",fontSize:12}}>{cd[c]||"—"}</td>)}
                      <td style={td}><button onClick={()=>setEditUid(aUid)} style={{...btn("#f3f4f6","#374151"),marginRight:6}}>✎</button><button onClick={()=>rmAgent(aUid)} style={{...btn("#fff1f2","#dc2626"),border:"1px solid #fca5a5"}}>✕</button></td>
                    </>}</tr>);
                  })}</tbody>
                </table></div>}
            </div>
          )}
          {panel==="rates"&&<div><h3 style={{margin:"0 0 12px",fontSize:15}}>Ставки</h3><RatesPanel rates={rates} onSave={saveRates} agentDir={agentDir}/></div>}
          {panel==="volrates"&&<div><h3 style={{margin:"0 0 12px",fontSize:15}}>Ставки добровольных продуктов</h3><VolRatesPanel volRates={volRates} onSave={saveVR}/></div>}
          {panel==="exceptions"&&<div><h3 style={{margin:"0 0 12px",fontSize:15}}>Исключения</h3><ExceptionsPanel exceptions={exceptions} onSave={saveExcs} agentDir={agentDir}/></div>}
        </div>
      )}

      {showBackup&&(
        <div style={{border:"1px solid #7c3aed",borderRadius:8,padding:16,marginBottom:16,background:"#faf5ff"}}>
          <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:8}}>
            <strong style={{fontSize:14,color:"#6d28d9"}}>💾 Резервная копия</strong>
            <button onClick={()=>setShowBackup(false)} style={{background:"none",border:"none",cursor:"pointer",fontSize:18,color:"#9ca3af"}}>×</button>
          </div>
          <textarea readOnly value={backupJson} onClick={e=>e.target.select()} style={{width:"100%",height:140,fontFamily:"monospace",fontSize:11,padding:8,border:"1px solid #d1d5db",borderRadius:6,boxSizing:"border-box",background:"white"}}/>
          <p style={{fontSize:11,color:"#9ca3af",margin:"6px 0 0"}}>Нажми на поле → Ctrl+C</p>
        </div>
      )}

      {tab==="commissions"&&(
        <div>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12,background:"#f1f5f9",borderRadius:8,padding:"10px 14px",flexWrap:"wrap"}}>
            <button onClick={()=>setSelMonth(prevMo(selMonth))} disabled={selMonth<=MIN_MONTH} style={{...btn("#fff","#374151",{border:"1px solid #d1d5db",fontSize:18,padding:"3px 10px"}),opacity:selMonth<=MIN_MONTH?0.4:1}}>‹</button>
            <span style={{fontWeight:700,fontSize:16,minWidth:160,textAlign:"center"}}>{fmtMonth(selMonth)}</span>
            <button onClick={()=>setSelMonth(nextMo(selMonth))} disabled={selMonth>=MAX_MONTH} style={{...btn("#fff","#374151",{border:"1px solid #d1d5db",fontSize:18,padding:"3px 10px"}),opacity:selMonth>=MAX_MONTH?0.4:1}}>›</button>
            <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
              {savedOk&&<span style={{color:"#16a34a",fontSize:12,fontWeight:600}}>✓ Сохранено</span>}
              {(storedPols.length>0||storedVol.length>0)&&uploadedFiles.length===0&&volSession.length===0&&<span style={{fontSize:11,color:"#6b7280"}}>{"📥 "+storedPols.length+" полисов из хранилища"}</span>}
              <button onClick={saveMonth} style={btn("#16a34a",undefined,{fontSize:11})}>💾 Сохранить месяц</button>
            </div>
          </div>

          <div style={{border:"1px solid #e5e7eb",borderRadius:10,padding:14,marginBottom:12,background:"#fafafa"}}>
            <div style={{fontWeight:600,fontSize:14,marginBottom:10}}>{"📂 Файлы — "+fmtMonth(selMonth)}</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(155px,1fr))",gap:8}}>
              {[...ALL_COMPANIES,"voluntary"].map(co=>{
                const isVol=co==="voluntary";
                const file=uploadedFiles.find(f=>f.company===co);
                const volActive=isVol&&(volSession.length>0||storedVol.length>0);
                const active=file||volActive;
                const label=isVol?"📦 Доброволь.":co;
                const rowCount=isVol?(volSession.length||storedVol.length):file?(file.rows.length-1):0;
                return(
                  <div key={co} onClick={()=>handleSlotClick(co)}
                    style={{border:"2px dashed "+(active?"#86efac":"#d1d5db"),borderRadius:8,padding:10,background:active?"#f0fdf4":"white",minHeight:78,cursor:"pointer",display:"flex",flexDirection:"column",gap:4}}>
                    <div style={{fontWeight:700,fontSize:13,color:active?"#16a34a":"#6b7280"}}>{label}</div>
                    {active?(
                      <div>
                        {file&&<div style={{fontSize:11,color:"#6b7280",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{file.fileName}</div>}
                        <div style={{fontSize:11,color:"#16a34a"}}>{rowCount+" строк"}</div>
                        <button onClick={ev=>{ev.stopPropagation();removeFile(co);}} style={{...btn("#fff1f2","#dc2626",{border:"1px solid #fca5a5"}),fontSize:11,padding:"2px 6px",marginTop:4,width:"fit-content"}}>✕</button>
                      </div>
                    ):<div style={{fontSize:11,color:"#9ca3af",marginTop:"auto"}}>Нажмите для загрузки</div>}
                  </div>
                );
              })}
            </div>
            <input ref={fileRef} type="file" accept=".xlsx,.xls" onChange={handleFile} style={{display:"none"}}/>
          </div>

          {warnings.map((w,i)=><div key={i} style={{background:"#fffbeb",border:"1px solid #fcd34d",borderRadius:8,padding:"10px 14px",marginBottom:10,fontSize:13,color:"#92400e"}}>{"⚠ "+w}</div>)}

          {hasData&&(
            <div>
              <div style={{display:"flex",gap:16,flexWrap:"wrap",background:"#111827",color:"#fff",borderRadius:10,padding:"14px 20px",marginBottom:16}}>
                {[["Доход офиса",fmt(totals.office),"#93c5fd"],["Выплачено агентам",fmt(totals.agent),"#fcd34d"],["Прибыль офиса",fmt(totals.profit),"#4ade80"]].map(([l,v,c])=>(
                  <div key={l} style={{minWidth:160}}><div style={{fontSize:11,opacity:.55,marginBottom:2}}>{l}</div><div style={{fontSize:18,fontWeight:700,color:c}}>{v}</div></div>
                ))}
              </div>

              {agentData.length>0&&(
                <div style={{overflowX:"auto",borderRadius:8,border:"1px solid #e5e7eb",marginBottom:16}}>
                  <table style={{width:"100%",borderCollapse:"collapse"}}>
                    <thead><tr>{["Агент","768-код","Полисов всего","Не зачтено","Зачётные полисы","Всего продаж","Зачётные продажи","Выплачено агенту","Доход офиса","Прибыль офиса",""].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
                    <tbody>{agentData.map(a=>{
                      const hasOv=!!(rates.agentOverrides&&rates.agentOverrides[a.uid]);
                      const excCnt=a.policies.filter(p=>p.exception).length;
                      const validCnt=a.policies.length-excCnt;
                      const intCode=(agentDir[a.uid]&&agentDir[a.uid].internalCode)||"";
                      return(
                        <tr key={a.uid} style={{background:activeAgent===a.uid?"#eff6ff":"white"}}>
                          <td style={{...td,fontWeight:600}}>{agName(a)}{hasOv&&<span style={{marginLeft:5,fontSize:10,background:"#f5f3ff",color:"#7c3aed",borderRadius:10,padding:"1px 5px"}}>инд.</span>}</td>
                          <td style={{...td,color:"#6366f1",fontWeight:600,fontSize:12}}>{intCode||"—"}</td>
                          <td style={{...td,textAlign:"center"}}>{a.policies.length}</td>
                          <td style={{...td,textAlign:"center",color:excCnt>0?"#dc2626":"#9ca3af",fontWeight:excCnt>0?600:400}}>{excCnt>0?excCnt:"—"}</td>
                          <td style={{...td,textAlign:"center"}}>{validCnt}</td>
                          <td style={td}>{fmt(a.totalSales)}</td>
                          <td style={td}>{fmt(a.validSales)}</td>
                          <td style={td}>{fmt(a.totalAgent)}</td>
                          <td style={td}>{fmt(a.totalOffice)}</td>
                          <td style={{...td,fontWeight:700,color:a.profit>=0?"#16a34a":"#dc2626"}}>{fmt(a.profit)}</td>
                          <td style={td}><button onClick={()=>setActiveAgent(activeAgent===a.uid?null:a.uid)} style={{...btn("#f9fafb","#374151"),border:"1px solid #d1d5db",fontWeight:400,fontSize:11}}>{activeAgent===a.uid?"Скрыть":"Детали"}</button></td>
                        </tr>
                      );
                    })}</tbody>
                  </table>
                </div>
              )}

              {effVol.length>0&&(
                <div style={{border:"1px solid #ddd6fe",borderRadius:8,overflow:"hidden",marginBottom:16}}>
                  <div style={{background:"#f5f3ff",padding:"10px 16px",fontWeight:600,fontSize:14,color:"#6d28d9"}}>{"📦 Добровольные — "+effVol.length}</div>
                  <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}>
                    <thead><tr>{["Продукт","Компания","Агент","Страхователь","Сумма","% О","Ком.О","% А","Ком.А"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
                    <tbody>{effVol.map(v=><tr key={v._id}><td style={{...td,fontWeight:600}}>{v.productName}</td><td style={td}>{v.company}</td><td style={td}>{getName(v.agentUid)||v.agentCode}</td><td style={td}>{v.insuredName}</td><td style={td}>{fmt(v.amount)}</td><td style={{...td,color:"#6b7280"}}>{v.officeRate+"%"}</td><td style={td}>{fmt(v.officeComm)}</td><td style={{...td,color:"#6b7280"}}>{v.agentRate+"%"}</td><td style={td}>{fmt(v.agentComm)}</td></tr>)}</tbody>
                  </table></div>
                </div>
              )}

              {detail&&(
                <div style={{border:"1px solid #e5e7eb",borderRadius:8,overflow:"hidden",marginBottom:16}}>
                  <div style={{background:"#f1f5f9",padding:"10px 16px",fontWeight:600,fontSize:14}}>{"Агент: "+agName(detail)}</div>
                  <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}>
                    <thead><tr>{["№ полиса","Компания","Марка","Рег.номер","Страхователь","Телефон","Начало","Окончание","Дней","Срок","Сумма","Регион","БМ","Мощн","Статус","% А","Ком.А"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
                    <tbody>{detail.policies.map(p=>{
                      const termColor=p.term==="SH"?"#d97706":p.term==="L"?"#2563eb":"#6b7280";
                      return(
                        <tr key={p._id} style={{background:p.exception?"#fff1f2":"white"}}>
                          <td style={td}>{p.policyNum}</td><td style={td}>{p.company}</td>
                          <td style={{...td,color:"#6b7280"}}>{p.car}</td>
                          <td style={{...td,fontSize:11}}>{p.carPlate}</td>
                          <td style={{...td,fontSize:11}}>{p.insuredName}</td>
                          <td style={{...td,fontSize:11}}>{p.phone}</td>
                          <td style={{...td,fontSize:11,color:"#6b7280"}}>{p.startDateFmt||"—"}</td>
                          <td style={{...td,fontSize:11,color:"#6b7280"}}>{p.endDateFmt||"—"}</td>
                          <td style={{...td,fontSize:11,textAlign:"center"}}>{p.days!=null?p.days:"—"}</td>
                          <td style={{...td,fontSize:12,textAlign:"center",fontWeight:700,color:termColor}}>{p.term||"—"}</td>
                          <td style={td}>{fmt(p.amount)}</td>
                          <td style={td}>{p.region||"—"}</td><td style={td}>{p.bm||"—"}</td><td style={td}>{p.power||"—"}</td>
                          <td style={td}>{p.exception?<span style={{color:"#dc2626",fontSize:12,fontWeight:600}}>Искл.</span>:<span style={{color:"#16a34a",fontSize:12}}>✓</span>}</td>
                          <td style={{...td,color:"#6b7280",fontSize:11}}>{p.agentRate+"%"}</td><td style={td}>{fmt(p.agentComm)}</td>
                        </tr>
                      );
                    })}</tbody>
                  </table></div>
                </div>
              )}

              {allExcs.length>0&&(
                <div style={{border:"1px solid #fecaca",borderRadius:8,overflow:"hidden",marginBottom:16}}>
                  <div style={{background:"#fff1f2",padding:"10px 16px",fontWeight:600,fontSize:14,color:"#dc2626"}}>{"⚠ Исключения — "+allExcs.length}</div>
                  <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse"}}>
                    <thead><tr>{["Агент","№ полиса","Компания","Марка","Сумма","Регион","БМ","Мощн","Срок","Причина"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
                    <tbody>{allExcs.map(p=><tr key={p._id} style={{background:"#fff7f7"}}>
                      <td style={{...td,fontWeight:600}}>{getName(p.agentUid)||p.agentCode}</td>
                      <td style={td}>{p.policyNum}</td><td style={td}>{p.company}</td><td style={td}>{p.car}</td>
                      <td style={td}>{fmt(p.amount)}</td><td style={td}>{p.region||"—"}</td><td style={td}>{p.bm}</td><td style={td}>{p.power}</td><td style={td}>{p.term}</td>
                      <td style={{...td,color:"#dc2626",fontSize:12}}>{excReason(p,exceptions,p.agentUid)}</td>
                    </tr>)}</tbody>
                  </table></div>
                </div>
              )}
            </div>
          )}
        </div>
      )}

      {tab==="policydb"&&(
        <div>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16,flexWrap:"wrap"}}>
            <div style={{fontWeight:600,fontSize:14}}>Полисы с окончанием в:</div>
            <select value={expiryF} onChange={e=>setExpiryF(e.target.value)} style={{...inp,padding:"5px 10px",fontSize:13}}>
              <option value="">— Все —</option>
              {EXPIRY_OPTIONS.map(mo=><option key={mo} value={mo}>{fmtMonth(mo)}</option>)}
            </select>
            <button onClick={loadDB} style={btn()}>🔄 Обновить</button>
            {dbLoaded&&<span style={{fontSize:12,color:"#6b7280"}}>{filteredDB.length+" полисов"+(expiryF?" ("+fmtMonth(expiryF)+")":"")}</span>}
          </div>
          {!dbLoaded&&<div style={{padding:30,textAlign:"center",color:"#9ca3af"}}>Загрузка...</div>}
          {dbLoaded&&filteredDB.length===0&&<div style={{padding:40,textAlign:"center",color:"#9ca3af",fontSize:14}}>{dbPols.length===0?"База пуста. Загрузите файлы и нажмите «💾 Сохранить месяц».":"Нет полисов с окончанием в "+fmtMonth(expiryF)+"."}</div>}
          {dbLoaded&&filteredDB.length>0&&(
            <div style={{overflowX:"auto",borderRadius:8,border:"1px solid #e5e7eb"}}>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead><tr>{["Месяц","№ полиса","Компания","Страхователь","Телефон","Марка","Рег.номер","Начало","Окончание","Агент","Регион","Сумма"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
                <tbody>{filteredDB.map((p,i)=>(
                  <tr key={i} style={{background:i%2===0?"white":"#fafafa"}}>
                    <td style={{...td,fontSize:11,color:"#6b7280"}}>{fmtMonth(p._monthKey)}</td>
                    <td style={td}>{p.policyNum}</td><td style={td}>{p.company}</td>
                    <td style={td}>{p.insuredName}</td><td style={{...td,fontSize:11}}>{p.phone}</td>
                    <td style={{...td,color:"#6b7280"}}>{p.car}</td><td style={{...td,fontSize:11}}>{p.carPlate}</td>
                    <td style={{...td,fontSize:11,color:"#6b7280"}}>{p.startDateFmt}</td>
                    <td style={{...td,fontSize:11,fontWeight:600,color:"#dc2626"}}>{p.endDateFmt}</td>
                    <td style={td}>{getName(p.agentUid)||p.agentCode}</td>
                    <td style={td}>{p.region}</td><td style={td}>{fmt(p.amount)}</td>
                  </tr>
                ))}</tbody>
              </table>
            </div>
          )}
        </div>
      )}

      {tab==="payroll"&&(
        <div>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16,background:"#f1f5f9",borderRadius:8,padding:"10px 14px",flexWrap:"wrap"}}>
            <button onClick={()=>setSelMonth(prevMo(selMonth))} disabled={selMonth<=MIN_MONTH} style={{...btn("#fff","#374151",{border:"1px solid #d1d5db",fontSize:18,padding:"3px 10px"}),opacity:selMonth<=MIN_MONTH?0.4:1}}>‹</button>
            <span style={{fontWeight:700,fontSize:16,minWidth:160,textAlign:"center"}}>{fmtMonth(selMonth)}</span>
            <button onClick={()=>setSelMonth(nextMo(selMonth))} disabled={selMonth>=MAX_MONTH} style={{...btn("#fff","#374151",{border:"1px solid #d1d5db",fontSize:18,padding:"3px 10px"}),opacity:selMonth>=MAX_MONTH?0.4:1}}>›</button>
            <div style={{marginLeft:"auto",display:"flex",gap:8}}>
              <button onClick={()=>exportPayroll(agentData,agentDir,selMonth)} style={btn("#16a34a")}>⬇ Экспорт Excel</button>
              <button onClick={()=>window.print()} style={btn("#6366f1")}>🖨 Печать</button>
            </div>
          </div>

          {agentData.filter(a=>a.validSales>0).length===0
            ?<div style={{padding:40,textAlign:"center",color:"#9ca3af",fontSize:14}}>{"Нет данных за "+fmtMonth(selMonth)+". Загрузите файлы в разделе Комиссии."}</div>
            :(
            <div style={{overflowX:"auto",borderRadius:8,border:"1px solid #e5e7eb"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                <thead>
                  <tr>
                    <th rowSpan={2} style={{...th,verticalAlign:"middle"}}>Имя агента</th>
                    <th rowSpan={2} style={{...th,verticalAlign:"middle",color:"#6366f1"}}>768-код</th>
                    {ALL_COMPANIES.map(c=><th key={c} colSpan={2} style={{...th,textAlign:"center",borderLeft:"1px solid #e5e7eb"}}>{c}</th>)}
                    <th rowSpan={2} style={{...th,verticalAlign:"middle",textAlign:"center",borderLeft:"2px solid #374151"}}>Итого</th>
                    <th rowSpan={2} style={{...th,verticalAlign:"middle",minWidth:120}}>Подпись</th>
                  </tr>
                  <tr>
                    {ALL_COMPANIES.map(c=>[
                      <th key={c+"-n"} style={{...th,fontSize:10,textAlign:"center",borderLeft:"1px solid #e5e7eb"}}>кол-во</th>,
                      <th key={c+"-s"} style={{...th,fontSize:10,textAlign:"center"}}>сумма</th>
                    ])}
                  </tr>
                </thead>
                <tbody>
                  {agentData.filter(a=>a.validSales>0).map(a=>{
                    const ic=(agentDir[a.uid]&&agentDir[a.uid].internalCode)||"";
                    return(
                      <tr key={a.uid} style={{borderBottom:"1px solid #f0f0f0"}}>
                        <td style={{...td,fontWeight:600}}>{agName(a)}</td>
                        <td style={{...td,color:"#6366f1",fontWeight:600,fontSize:11}}>{ic||"—"}</td>
                        {ALL_COMPANIES.map(c=>{
                          const pols=a.policies.filter(p=>p.company===c&&!p.exception);
                          const sum=pols.reduce((s,p)=>s+p.agentComm,0);
                          return[
                            <td key={c+"-n"} style={{...td,textAlign:"center",borderLeft:"1px solid #e5e7eb",color:pols.length>0?"#111":"#d1d5db"}}>{pols.length>0?pols.length:"—"}</td>,
                            <td key={c+"-s"} style={{...td,textAlign:"right",color:sum>0?"#111":"#d1d5db",fontSize:11}}>{sum>0?fmt(sum):"—"}</td>
                          ];
                        })}
                        <td style={{...td,fontWeight:700,textAlign:"right",borderLeft:"2px solid #374151",color:"#16a34a"}}>{fmt(a.totalAgent)}</td>
                        <td style={{...td,borderBottom:"1px solid #9ca3af",minWidth:120}}></td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
        <div>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16,background:"#f1f5f9",borderRadius:8,padding:"10px 14px",flexWrap:"wrap"}}>
            <button onClick={()=>setSelMonth(prevMo(selMonth))} disabled={selMonth<=MIN_MONTH} style={{...btn("#fff","#374151",{border:"1px solid #d1d5db",fontSize:18,padding:"3px 10px"}),opacity:selMonth<=MIN_MONTH?0.4:1}}>‹</button>
            <span style={{fontWeight:700,fontSize:16,minWidth:160,textAlign:"center"}}>{fmtMonth(selMonth)}</span>
            <button onClick={()=>setSelMonth(nextMo(selMonth))} disabled={selMonth>=MAX_MONTH} style={{...btn("#fff","#374151",{border:"1px solid #d1d5db",fontSize:18,padding:"3px 10px"}),opacity:selMonth>=MAX_MONTH?0.4:1}}>›</button>
            <div style={{marginLeft:"auto",display:"flex",gap:8}}>
              {savedOk&&<span style={{color:"#16a34a",fontSize:12,fontWeight:600}}>✓ Сохранено</span>}
              <button onClick={saveMonth} style={btn("#16a34a",undefined,{fontSize:11})}>💾 Сохранить месяц</button>
            </div>
          </div>
          <div style={{border:"1px solid #e5e7eb",borderRadius:10,padding:16,marginBottom:16,background:"#fafafa"}}>
            <div style={{fontWeight:600,fontSize:14,marginBottom:10}}>{"📂 Файл продаж офиса — "+fmtMonth(selMonth)}</div>
            <div onClick={()=>officeFileRef.current.click()}
              style={{border:"2px dashed "+(officeData.rows&&officeData.rows.length>0?"#86efac":"#d1d5db"),borderRadius:8,padding:16,cursor:"pointer",background:officeData.rows&&officeData.rows.length>0?"#f0fdf4":"white",textAlign:"center",marginBottom:8}}>
              {officeData.rows&&officeData.rows.length>0?(
                <div>
                  <div style={{fontWeight:600,color:"#16a34a",fontSize:14}}>{"✓ "+(officeData.fileName||"Файл загружен")}</div>
                  <div style={{fontSize:12,color:"#6b7280",marginTop:4}}>{officeData.rows.length+" строк"}</div>
                </div>
              ):(
                <div>
                  <div style={{fontSize:24,marginBottom:4}}>📂</div>
                  <div style={{color:"#6b7280",fontSize:14}}>Нажмите для загрузки файла</div>
                </div>
              )}
            </div>
            <input ref={officeFileRef} type="file" accept=".xlsx,.xls" onChange={handleOfficeFile} style={{display:"none"}}/>
            <p style={{fontSize:11,color:"#9ca3af",margin:0}}>Колонки: Компания, Страховая сумма, Прибыль</p>
          </div>
          {officeData.rows&&officeData.rows.length>0&&(
            <div>
              <div style={{display:"flex",gap:16,flexWrap:"wrap",background:"#111827",color:"#fff",borderRadius:10,padding:"14px 20px",marginBottom:16}}>
                <div style={{minWidth:200}}>
                  <div style={{fontSize:11,opacity:.55,marginBottom:2}}>{"Прибыль офиса за "+fmtMonth(selMonth)}</div>
                  <div style={{fontSize:22,fontWeight:700,color:"#4ade80"}}>{fmt(officeData.profit)}</div>
                </div>
                <div style={{minWidth:160}}>
                  <div style={{fontSize:11,opacity:.55,marginBottom:2}}>Строк в файле</div>
                  <div style={{fontSize:18,fontWeight:700,color:"#93c5fd"}}>{officeData.rows.length}</div>
                </div>
              </div>
              <div style={{overflowX:"auto",borderRadius:8,border:"1px solid #e5e7eb"}}>
                <table style={{width:"100%",borderCollapse:"collapse"}}>
                  <thead><tr>{["Компания","Страховая сумма","Прибыль"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
                  <tbody>{officeData.rows.map((r,i)=>(
                    <tr key={i} style={{background:i%2===0?"white":"#fafafa"}}>
                      <td style={td}>{r.company||"—"}</td>
                      <td style={td}>{fmt(r.amount)}</td>
                      <td style={{...td,fontWeight:600,color:"#16a34a"}}>{fmt(r.profit)}</td>
                    </tr>
                  ))}</tbody>
                  <tfoot><tr>
                    <td style={{...td,fontWeight:700,borderTop:"2px solid #e5e7eb"}}>ИТОГО</td>
                    <td style={{...td,fontWeight:700,borderTop:"2px solid #e5e7eb"}}>{fmt(officeData.rows.reduce((s,r)=>s+r.amount,0))}</td>
                    <td style={{...td,fontWeight:700,color:"#16a34a",borderTop:"2px solid #e5e7eb"}}>{fmt(officeData.profit)}</td>
                  </tr></tfoot>
                </table>
              </div>
            </div>
          )}
        </div>
      
    </div>
  );
}