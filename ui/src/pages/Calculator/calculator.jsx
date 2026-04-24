import{useState,useMemo,useEffect,useRef,Fragment}from"react";
import{calcStorage}from"./calcStorage";
import * as XLSX from "xlsx";
import XLSXStyle from "xlsx-js-style";
import JSZip from "jszip";

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

const POL_STATUSES=[{k:"",l:"Обычный"},{k:"taxi",l:"🚕 Такси"},{k:"ot",l:"ОТ"},{k:"ok",l:"ОК"},{k:"yr_kt",l:"YR-KT"},{k:"restricted",l:"С ограничением"}];
const fmtPolStatus=k=>(POL_STATUSES.find(s=>s.k===k)||POL_STATUSES[0]).l;

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
  productName:["productname","productName","Product","продукт","наименование","название","вид","тип","product","продукты"],
  company:["company","Company","компания","страховая","страховщик"],
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
  // Fallback: if company column not detected by header, scan data for company names
  let coIdx=cm.company;
  if(coIdx===undefined){
    const testRows=rows.slice(1,Math.min(6,rows.length));
    rows[0].forEach((_,i)=>{
      if(coIdx!==undefined)return;
      if(testRows.some(r=>!!detectCo(String(r[i]||""))))coIdx=i;
    });
  }
  return rows.slice(1).filter(r=>r.some(c=>c!=="")).map((row,idx)=>{
    const get=f=>cm[f]!==undefined?String(row[cm[f]]||"").trim():"";
    const company=coIdx!==undefined?String(row[coIdx]||"").trim():get("company");
    return{_id:`vol-${idx}`,productName:get("productName"),company,
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
  const[payrollIC,setPayrollIC]=useState("");
  const[payrollUID,setPayrollUID]=useState(null);
  const[showAllPayroll,setShowAllPayroll]=useState(false);
  const[officeCodes,setOfficeCodes]=useState([]);
  const[newOfficeCode,setNewOfficeCode]=useState("");
  const fileRef=useRef();const backRef=useRef();const officeFileRef=useRef();
  const[role,setRole]=useState("employee");
  const[pinModal,setPinModal]=useState(false);
  const[pinInput,setPinInput]=useState("");
  const[pinError,setPinError]=useState("");
  const[adminPin,setAdminPin]=useState(null);
  const[newPinA,setNewPinA]=useState("");
  const[newPinB,setNewPinB]=useState("");
  const[pinChangeMsg,setPinChangeMsg]=useState("");
  const[opCurrentMonth,setOpCurrentMonth]=useState([]);
  const[opPrevUnpaid,setOpPrevUnpaid]=useState([]);
  const[opLoaded,setOpLoaded]=useState(false);
  const[opFormOpen,setOpFormOpen]=useState(false);
  const[opEditPol,setOpEditPol]=useState(null);
  const[opFD,setOpFD]=useState({});
  const[opPayPol,setOpPayPol]=useState(null);
  const[opPayData,setOpPayData]=useState({});
  const[opSearch,setOpSearch]=useState("");
  const[opStatusFilter,setOpStatusFilter]=useState("all");
  const[cashDays,setCashDays]=useState({});
  const[cashLoaded,setCashLoaded]=useState(false);
  const[cashExpandDay,setCashExpandDay]=useState(null);
  const[cashCloseModal,setCashCloseModal]=useState(null);
  const[cashReopenModal,setCashReopenModal]=useState(null);
  const[cashMonthPols,setCashMonthPols]=useState([]);
  const[officeStaff,setOfficeStaff]=useState(["768-101","768-105","768-106"]);
  const[newStaffCode,setNewStaffCode]=useState("");

  useEffect(()=>{(async()=>{
    try{const r=await calcStorage.get("agentDirectory").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(valD(p))setAgentDir(p);}else{setAgentDir(SEED_AGENTS);calcStorage.set("agentDirectory",JSON.stringify(SEED_AGENTS)).catch(()=>{});}}catch{setAgentDir(SEED_AGENTS);}
    try{const r=await calcStorage.get("ratesConfig").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(valR(p))setRates(p);}}catch{}
    try{const r=await calcStorage.get("volRates").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(valV(p))setVolRates(p);}}catch{}
    try{const r=await calcStorage.get("exceptionsConfig").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(valE(p))setExceptions(p);}}catch{}
    try{const r=await calcStorage.get("appSettings").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(p&&p.adminPin)setAdminPin(p.adminPin);if(p&&Array.isArray(p.officeStaff)&&p.officeStaff.length)setOfficeStaff(p.officeStaff);}}catch{}
  })();},[]);

  useEffect(()=>{
    setUploadedFiles([]);setVolSession([]);setOfficeSession(null);setActiveAgent(null);setOfficeCodes([]);setNewOfficeCode("");
    (async()=>{
      try{const r=await calcStorage.get("month:"+selMonth).catch(()=>null);if(r&&r.value){const d=JSON.parse(r.value);setStoredPols(d.policies||[]);setStoredVol(d.voluntary||[]);}else{setStoredPols([]);setStoredVol([]);}}catch{setStoredPols([]);setStoredVol([]);}
      try{const r=await calcStorage.get("officeStore:"+selMonth).catch(()=>null);if(r&&r.value)setStoredOffice(JSON.parse(r.value));else setStoredOffice({profit:0,rows:[]});}catch{setStoredOffice({profit:0,rows:[]});}
      try{const r=await calcStorage.get("officeCodes:"+selMonth).catch(()=>null);if(r&&r.value)setOfficeCodes(JSON.parse(r.value));else setOfficeCodes([]);}catch{setOfficeCodes([]);}
    })();
  },[selMonth]);

  useEffect(()=>{if(role==="employee"&&tab==="commissions")setTab("policydb");},[role]);
  const saveDir=d=>{setAgentDir(d);calcStorage.set("agentDirectory",JSON.stringify(d)).catch(()=>{});};
  const saveOfficeCodes=codes=>{setOfficeCodes(codes);calcStorage.set("officeCodes:"+selMonth,JSON.stringify(codes)).catch(()=>{});};
  const addOfficeCode=()=>{const v=newOfficeCode.trim();if(!v)return;saveOfficeCodes([...officeCodes,v]);setNewOfficeCode("");};
  const removeOfficeCode=idx=>saveOfficeCodes(officeCodes.filter((_,i)=>i!==idx));
  const saveRates=r=>{setRates(r);calcStorage.set("ratesConfig",JSON.stringify(r)).catch(()=>{});};
  const saveVR=r=>{setVolRates(r);calcStorage.set("volRates",JSON.stringify(r)).catch(()=>{});};
  const saveExcs=e=>{setExceptions(e);calcStorage.set("exceptionsConfig",JSON.stringify(e)).catch(()=>{});};
  const isAdmin=role==="admin";
  const tryAdminLogin=()=>{
    const p=pinInput.trim();
    if(!p){setPinError("Введите PIN");return;}
    if(!adminPin){
      const s={adminPin:p,officeStaff};
      calcStorage.set("appSettings",JSON.stringify(s)).catch(()=>{});
      setAdminPin(p);setRole("admin");setPinModal(false);setPinInput("");setPinError("");
    } else {
      if(p===adminPin){setRole("admin");setPinModal(false);setPinInput("");setPinError("");}
      else setPinError("Неверный PIN");
    }
  };
  const saveOfficeStaff=(list)=>{setOfficeStaff(list);calcStorage.set("appSettings",JSON.stringify({adminPin:adminPin||"",officeStaff:list})).catch(()=>{});};
  const changePin=()=>{
    if(!newPinA.trim()){setPinChangeMsg("Введите новый PIN");return;}
    if(newPinA!==newPinB){setPinChangeMsg("PIN не совпадают");return;}
    const s={adminPin:newPinA.trim(),officeStaff};
    calcStorage.set("appSettings",JSON.stringify(s)).catch(()=>{});
    setAdminPin(newPinA.trim());setNewPinA("");setNewPinB("");
    setPinChangeMsg("✓ PIN изменён");setTimeout(()=>setPinChangeMsg(""),3000);
  };
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
      Object.entries(agent.codes||{}).forEach(([co,code])=>{
        if(code&&code.trim()){
          const c=code.replace(/\s+/g,"").trim();
          map[co+":"+c]=aUid;
          if(!map[c])map[c]=aUid; // без привязки к компании (первый найденный)
        }
      });
      if(agent.internalCode&&agent.internalCode.trim())map[agent.internalCode.trim()]=aUid;
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
      const _ac=(v.agentCode||"").replace(/\s+/g,"").trim();
      const aUid=codeLookup[v.company+":"+_ac]||codeLookup[_ac]
        ||Object.keys(agentDir).find(uid=>Object.values(agentDir[uid].codes||{}).some(c=>c&&c.replace(/\s+/g,"").trim()===_ac))
        ||null;
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
    calcStorage.set("month:"+selMonth,JSON.stringify({policies:pols,voluntary:effVol})).catch(()=>{});
    if(officeSession)calcStorage.set("officeStore:"+selMonth,JSON.stringify(officeSession)).catch(()=>{});
    setSavedOk(true);setTimeout(()=>setSavedOk(false),2500);
  };

  const loadDB=()=>{
    setDbLoaded(false);
    calcStorage.list("month:").then(res=>{
      if(!res||!res.keys||!res.keys.length){setDbPols([]);setDbLoaded(true);return;}
      const all=[];let done=0;
      res.keys.forEach(key=>{
        calcStorage.get(key).then(r=>{
          if(r&&r.value){try{const d=JSON.parse(r.value);const mk=key.replace("month:","");(d.policies||[]).forEach(p=>all.push({...p,_monthKey:mk}));}catch{}}
          done++;if(done===res.keys.length){setDbPols(all);setDbLoaded(true);}
        }).catch(()=>{done++;if(done===res.keys.length){setDbPols(all);setDbLoaded(true);}});
      });
    }).catch(()=>{setDbPols([]);setDbLoaded(true);});
  };

  const loadOfficeSales=async()=>{
    setOpLoaded(false);
    const mk=selMonth;
    try{const r=await calcStorage.get("officePol:"+mk).catch(()=>null);setOpCurrentMonth(r&&r.value?JSON.parse(r.value):[]);}catch{setOpCurrentMonth([]);}
    try{
      const res=await calcStorage.list("officePol:").catch(()=>({keys:[]}));
      const otherKeys=(res.keys||[]).filter(k=>k!=="officePol:"+mk);
      if(!otherKeys.length){setOpPrevUnpaid([]);setOpLoaded(true);return;}
      const results=await Promise.all(otherKeys.map(async key=>{try{const r=await calcStorage.get(key).catch(()=>null);return r&&r.value?JSON.parse(r.value):[];}catch{return[];}}));
      setOpPrevUnpaid(results.flat().filter(p=>!p.paid).sort((a,b)=>new Date(a.date)-new Date(b.date)));
    }catch{setOpPrevUnpaid([]);}
    setOpLoaded(true);
  };
  useEffect(()=>{if(tab==="officesales")loadOfficeSales();},[tab,selMonth]);

  const loadCashBook=async()=>{
    setCashLoaded(false);
    try{const r=await calcStorage.get("cashBook:"+selMonth).catch(()=>null);setCashDays(r&&r.value?JSON.parse(r.value):{});}catch{setCashDays({});}
    try{const r=await calcStorage.get("officePol:"+selMonth).catch(()=>null);setCashMonthPols(r&&r.value?JSON.parse(r.value):[]);}catch{setCashMonthPols([]);}
    setCashLoaded(true);
  };
  useEffect(()=>{if(tab==="cashbook")loadCashBook();},[tab,selMonth]);
  const closeCashDay=(date)=>{
    const pols=cashMonthPols.filter(p=>p.paid&&p.paidDate===date);
    const snapshot=pols.map(p=>({_id:p._id,insuredName:p.insuredName,polType:p.polType,productName:p.productName,car:p.car,carPlate:p.carPlate,policyNum:p.policyNum,company:p.company,phone:p.phone,date:p.date,amount:p.amount,discount:p.discount,paidAmount:p.paidAmount,paymentType:p.paymentType,paidDate:p.paidDate,agentUid:p.agentUid,comment:p.comment}));
    const cash=pols.filter(p=>p.paymentType==="cash").reduce((s,p)=>s+(p.paidAmount||0),0);
    const acba=pols.filter(p=>p.paymentType==="acba").reduce((s,p)=>s+(p.paidAmount||0),0);
    const ineco=pols.filter(p=>p.paymentType==="ineco").reduce((s,p)=>s+(p.paidAmount||0),0);
    const updated={...cashDays,[date]:{closed:true,closedAt:new Date().toISOString(),reopenedAt:(cashDays[date]||{}).reopenedAt||null,snapshot,totals:{cash,acba,ineco,total:cash+acba+ineco}}};
    setCashDays(updated);
    calcStorage.set("cashBook:"+selMonth,JSON.stringify(updated)).catch(()=>{});
    setCashCloseModal(null);
  };
  const reopenCashDay=(date)=>{
    const updated={...cashDays,[date]:{...cashDays[date],closed:false,reopenedAt:new Date().toISOString()}};
    setCashDays(updated);
    calcStorage.set("cashBook:"+selMonth,JSON.stringify(updated)).catch(()=>{});
    setCashReopenModal(null);
  };

  const saveOpMonth=(pols)=>{setOpCurrentMonth(pols);calcStorage.set("officePol:"+selMonth,JSON.stringify(pols)).catch(()=>{});};
  const initOpFD=()=>({polType:"osago",insuredName:"",phone:"",company:ALL_COMPANIES[0],policyNum:"",date:new Date().toISOString().slice(0,10),dateStart:"",dateEnd:"",car:"",carPlate:"",bm:"",region:"",power:"",term:"L",polStatus:"",amount:"",discount:"0",agentUid:"",comment:"",productName:"",payNow:false,paymentType:""});
  const addOfficePol=(fd)=>{const defaults=fd.paid?{}:{paid:false,paidAt:null,paidAmount:null,paymentType:null};const pol={_id:genUid(),_monthKey:selMonth,...fd,...defaults};saveOpMonth([...opCurrentMonth,pol]);};
  const saveEditPol=async(pol,updates)=>{
    if(pol._monthKey===selMonth){saveOpMonth(opCurrentMonth.map(p=>p._id===pol._id?{...p,...updates}:p));}
    else{
      const updated={...pol,...updates};
      const r=await calcStorage.get("officePol:"+pol._monthKey).catch(()=>null);
      const pols=r&&r.value?JSON.parse(r.value):[];
      calcStorage.set("officePol:"+pol._monthKey,JSON.stringify(pols.map(p=>p._id===pol._id?updated:p))).catch(()=>{});
      setOpPrevUnpaid(prev=>prev.map(p=>p._id===pol._id?updated:p));
    }
  };
  const acceptOpPayment=async(pol,payData)=>{
    const updated={...pol,...payData,paid:true,paidAt:new Date().toISOString()};
    if(pol._monthKey===selMonth){saveOpMonth(opCurrentMonth.map(p=>p._id===pol._id?updated:p));}
    else{
      const r=await calcStorage.get("officePol:"+pol._monthKey).catch(()=>null);
      const pols=r&&r.value?JSON.parse(r.value):[];
      calcStorage.set("officePol:"+pol._monthKey,JSON.stringify(pols.map(p=>p._id===pol._id?updated:p))).catch(()=>{});
      setOpPrevUnpaid(prev=>prev.filter(p=>p._id!==pol._id));
    }
    setOpPayPol(null);
  };
  const deleteOfficePol=async(pol)=>{
    if(!window.confirm("Удалить полис "+pol.insuredName+"?"))return;
    if(pol._monthKey===selMonth){saveOpMonth(opCurrentMonth.filter(p=>p._id!==pol._id));}
    else{
      const r=await calcStorage.get("officePol:"+pol._monthKey).catch(()=>null);
      const pols=r&&r.value?JSON.parse(r.value):[];
      calcStorage.set("officePol:"+pol._monthKey,JSON.stringify(pols.filter(p=>p._id!==pol._id))).catch(()=>{});
      setOpPrevUnpaid(prev=>prev.filter(p=>p._id!==pol._id));
    }
  };
  const openOpNew=()=>{setOpEditPol(null);setOpFD(initOpFD());setOpFormOpen(true);};
  const openOpEdit=(pol)=>{setOpEditPol(pol);setOpFD({polType:pol.polType||"osago",insuredName:pol.insuredName||"",phone:pol.phone||"",company:pol.company||ALL_COMPANIES[0],policyNum:pol.policyNum||"",date:pol.date||new Date().toISOString().slice(0,10),dateStart:pol.dateStart||"",dateEnd:pol.dateEnd||"",car:pol.car||"",carPlate:pol.carPlate||"",bm:pol.bm||"",region:pol.region||"",power:pol.power||"",term:pol.term||"L",polStatus:pol.polStatus||"",amount:String(pol.amount||""),discount:String(pol.discount||0),agentUid:pol.agentUid||"",comment:pol.comment||"",productName:pol.productName||"",payNow:false,paymentType:""});setOpFormOpen(true);};
  const openOpPay=(pol)=>{setOpPayPol(pol);setOpPayData({paidAmount:String(pol.amount-(pol.discount||0)),paymentType:"cash",paidDate:new Date().toISOString().slice(0,10)});};
  const submitOpForm=()=>{
    if(!opFD.insuredName||!opFD.insuredName.trim()||!opFD.amount)return;
    const today=new Date().toISOString().slice(0,10);
    const base={polType:opFD.polType||"osago",insuredName:opFD.insuredName.trim(),phone:(opFD.phone||"").trim(),company:opFD.company,policyNum:(opFD.policyNum||"").trim(),amount:parseFloat(opFD.amount)||0,discount:parseFloat(opFD.discount)||0,date:opFD.date,agentUid:opFD.agentUid||null,comment:(opFD.comment||"").trim()};
    const typeData=opFD.polType==="voluntary"
      ?{productName:(opFD.productName||"").trim()}
      :{dateStart:opFD.dateStart,dateEnd:opFD.dateEnd,car:(opFD.car||"").trim(),carPlate:(opFD.carPlate||"").trim(),bm:opFD.bm,region:opFD.region,power:opFD.power,term:opFD.term,polStatus:opFD.polStatus};
    const payData=(opFD.polType==="voluntary"&&opFD.payNow&&opFD.paymentType)
      ?{paid:true,paidAt:new Date().toISOString(),paidAmount:parseFloat(opFD.amount||0)-(parseFloat(opFD.discount)||0),paymentType:opFD.paymentType,paidDate:today}
      :{};
    const data={...base,...typeData,...payData};
    if(opEditPol)saveEditPol(opEditPol,data);else addOfficePol(data);
    setOpFormOpen(false);
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

  const polCols=["№ полиса","Компания","Страхователь","Марка","Рег.номер","Телефон","Начало","Окончание","Регион","БМ","Сумма","Ставка %","Начислено"];
  const polRow=p=>[p.policyNum,p.company,p.insuredName,p.car,p.carPlate,p.phone,p.startDateFmt,p.endDateFmt,p.region,p.bm,p.amount,p.exception?0:(p.agentRate||0),p.exception?0:p.agentComm];

  const exportAgentValid=(uid,pols,month)=>{
    const wb=XLSX.utils.book_new();
    const rows=[polCols,...pols.map(polRow)];
    rows.push(["ИТОГО","","","","","","","","","",pols.reduce((s,p)=>s+p.amount,0),"",pols.reduce((s,p)=>s+p.agentComm,0)]);
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),"Зачётные");
    const ic=(agentDir[uid]&&agentDir[uid].internalCode)||uid;
    XLSX.writeFile(wb,"Зачётные_"+ic+"_"+month+".xlsx");
  };

  const exportAgentExc=(uid,pols,month)=>{
    const wb=XLSX.utils.book_new();
    const rows=[polCols,...pols.map(polRow)];
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(rows),"Незачётные");
    const ic=(agentDir[uid]&&agentDir[uid].internalCode)||uid;
    XLSX.writeFile(wb,"Незачётные_"+ic+"_"+month+".xlsx");
  };

  const exportAgentAll=(uid,validPols,excPols,month)=>{
    const wb=XLSX.utils.book_new();
    const vRows=[polCols,...validPols.map(polRow)];
    vRows.push(["ИТОГО","","","","","","","","","",validPols.reduce((s,p)=>s+p.amount,0),"",validPols.reduce((s,p)=>s+p.agentComm,0)]);
    XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(vRows),"Зачётные");
    if(excPols.length>0){
      const eRows=[polCols,...excPols.map(polRow)];
      XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(eRows),"Незачётные");
    }
    const ic=(agentDir[uid]&&agentDir[uid].internalCode)||uid;
    XLSX.writeFile(wb,"Полисы_"+ic+"_"+month+".xlsx");
  };

  const buildAllPayrollRows=(agentData,effVol,agentDir,excludeCodes=[])=>{
    const icNum=ic=>{const m=(ic||"").match(/(\d+)$/);return m?parseInt(m[1]):99999;};
    const excSet=new Set(excludeCodes.map(c=>c.trim().toLowerCase()));
    const allUids=new Set([
      ...agentData.filter(a=>a.policies.length>0).map(a=>a.uid),
      ...effVol.filter(v=>v.agentUid).map(v=>v.agentUid),
    ]);
    return [...allUids].map(uid=>{
      const agData=agentData.find(a=>a.uid===uid);
      const volDebt=effVol.filter(v=>v.agentUid===uid).reduce((s,v)=>s+(v.amount-v.agentComm),0);
      const accrued=agData?agData.totalAgent:0;
      const polCount=agData?agData.policies.length:0;
      const net=accrued-volDebt;
      const ic=(agentDir[uid]&&agentDir[uid].internalCode)||"";
      const name=agentDir[uid]?(agentDir[uid].name+" "+agentDir[uid].surname).trim():uid;
      return{uid,name,ic,polCount,accrued,volDebt:Math.round(volDebt),net:Math.round(net)};
    })
    .filter(r=>!excSet.has(r.ic.trim().toLowerCase()))
    .sort((a,b)=>icNum(a.ic)-icNum(b.ic));
  };

  const exportAllPayrollXlsx=(agentData,effVol,agentDir,month)=>{
    const rows=buildAllPayrollRows(agentData,effVol,agentDir,officeCodes);
    const wb=XLSX.utils.book_new();
    const header=["Имя агента","768-код","Кол-во полисов","Начислено (ОСАГО)","К оплате офису (доброволь.)","К выплате","Примечание","Подпись"];
    const data=[header,...rows.map(r=>[r.name,r.ic,r.polCount,r.accrued,r.volDebt>0?r.volDebt:"",r.net,"",""])];
    const totPols=rows.reduce((s,r)=>s+r.polCount,0);
    const totAcc=rows.reduce((s,r)=>s+r.accrued,0);
    const totVol=rows.reduce((s,r)=>s+r.volDebt,0);
    const totNet=rows.reduce((s,r)=>s+r.net,0);
    data.push(["ИТОГО","",totPols,totAcc,totVol>0?totVol:"",totNet,"",""]);
    const ws=XLSX.utils.aoa_to_sheet(data);
    ws["!cols"]=[{wch:28},{wch:12},{wch:14},{wch:18},{wch:24},{wch:16},{wch:20},{wch:18}];
    XLSX.utils.book_append_sheet(wb,ws,"Начисления "+fmtMonth(month));
    XLSX.writeFile(wb,"Все_начисления_"+month+".xlsx");
  };

  const exportRenewalsZip=async(pols,expiryLabel)=>{
    const zip=new JSZip();
    const byAgent={};
    pols.forEach(p=>{const k=p.agentUid||"__unknown__";if(!byAgent[k])byAgent[k]=[];byAgent[k].push(p);});
    const HEADERS=["Месяц","№ полиса","Компания","Страхователь","Телефон","Марка","Рег. номер","Начало","Окончание","Агент","Регион","Сумма","БМ","Мощность"];
    const PHONE_COL=4;
    const br={style:"thin",color:{rgb:"AAAAAA"}};
    const borders={top:br,bottom:br,left:br,right:br};
    const sHdr={fill:{patternType:"solid",fgColor:{rgb:"E5E7EB"}},font:{bold:true,sz:10},border:borders,alignment:{horizontal:"center"}};
    const sHdrPhone={fill:{patternType:"solid",fgColor:{rgb:"FDE047"}},font:{bold:true,sz:10},border:borders,alignment:{horizontal:"center"}};
    const sCell={font:{sz:10},border:borders};
    const sPhone={fill:{patternType:"solid",fgColor:{rgb:"FEF9C3"}},font:{sz:10},border:borders};
    Object.entries(byAgent).forEach(([uid,agPols])=>{
      const ag=agentDir[uid];
      const name=ag?(ag.name+" "+ag.surname).trim():(getName(uid)||uid||"Неизвестно");
      const ic=(ag&&ag.internalCode)||"";
      const wb=XLSXStyle.utils.book_new();
      const ws={};
      HEADERS.forEach((h,c)=>{
        ws[XLSXStyle.utils.encode_cell({r:0,c})]={v:h,t:"s",s:c===PHONE_COL?sHdrPhone:sHdr};
      });
      agPols.forEach((p,ri)=>{
        const agName=getName(p.agentUid)||p.agentCode||"";
        const vals=[fmtMonth(p._monthKey),p.policyNum||"",p.company||"",p.insuredName||"",p.phone||"",p.car||"",p.carPlate||"",p.startDateFmt||"",p.endDateFmt||"",agName,p.region||"",p.amount||0,p.bm||"",p.power||""];
        vals.forEach((v,c)=>{
          const t=typeof v==="number"?"n":"s";
          ws[XLSXStyle.utils.encode_cell({r:ri+1,c})]={v,t,s:c===PHONE_COL?sPhone:sCell};
        });
      });
      ws["!ref"]=XLSXStyle.utils.encode_range({s:{r:0,c:0},e:{r:agPols.length,c:HEADERS.length-1}});
      ws["!cols"]=[{wch:13},{wch:17},{wch:10},{wch:26},{wch:17},{wch:20},{wch:13},{wch:12},{wch:12},{wch:24},{wch:8},{wch:14},{wch:6},{wch:8}];
      XLSXStyle.utils.book_append_sheet(wb,ws,"Продления");
      const out=XLSXStyle.write(wb,{bookType:"xlsx",type:"array"});
      const safe=s=>String(s).replace(/[/\\:*?"<>|]/g,"_");
      zip.file(`${safe(ic)}_${safe(name).slice(0,30)}.xlsx`,new Uint8Array(out));
    });
    const blob=await zip.generateAsync({type:"blob"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url;a.download=`Продления_${expiryLabel||"все"}.zip`;
    document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
  };

  const hasData=agentData.length>0||effVol.length>0;
  const panels=isAdmin?[["agents","👤 Агенты",Object.keys(agentDir).length],["rates","⚙️ Ставки",null],["volrates","📦 Доброволь.",volRates.rates.length],["exceptions","🚫 Исключения",exceptions.filter(e=>e.enabled).length],["access","🔐 Доступ",null]]:[];
  const backupJson=JSON.stringify({version:6,agentDir,rates,volRates,exceptions},null,2);

  return(
    <div style={{fontFamily:"system-ui,sans-serif",padding:20,maxWidth:1400,margin:"0 auto",color:"#111"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
        <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
          <h2 style={{margin:0,fontSize:20}}>Калькулятор комиссий</h2>
          {isAdmin
            ?<><span style={{background:"#fef3c7",border:"1px solid #fcd34d",borderRadius:6,padding:"4px 10px",fontSize:12,fontWeight:600,color:"#92400e"}}>🔑 Администратор</span>
               <button onClick={()=>{setRole("employee");setPanel(null);}} style={btn("#f3f4f6","#374151",{border:"1px solid #d1d5db",fontSize:12})}>Выйти</button></>
            :<><span style={{background:"#f0fdf4",border:"1px solid #86efac",borderRadius:6,padding:"4px 10px",fontSize:12,fontWeight:600,color:"#166534"}}>👤 Сотрудник</span>
               <button onClick={()=>setPinModal(true)} style={btn("#1d4ed8",undefined,{fontSize:12})}>🔐 Войти как администратор</button></>
          }
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {panels.map(([id,label,count])=>(
            <button key={id} onClick={()=>setPanel(panel===id?null:id)} style={{...btn(panel===id?"#eff6ff":"#f9fafb",panel===id?"#1d4ed8":"#374151"),border:"1px solid #d1d5db",fontWeight:500}}>
              {label}{count!=null?" ("+count+")":""}
            </button>
          ))}
          {isAdmin&&hasData&&<button onClick={()=>exportToExcel(agentData,effVol,agentDir,totals,exceptions)} style={btn("#16a34a")}>⬇ Excel</button>}
          {isAdmin&&<button onClick={()=>setShowBackup(p=>!p)} style={btn("#7c3aed")}>💾 Резерв</button>}
          {isAdmin&&<button onClick={()=>backRef.current.click()} style={btn("#f3f4f6","#374151",{border:"1px solid #d1d5db"})}>📂 Восстановить</button>}
          <input ref={backRef} type="file" accept=".json" onChange={importBackup} style={{display:"none"}}/>
        </div>
      </div>

      <div style={{display:"flex",borderBottom:"2px solid #e5e7eb",marginBottom:16,gap:0}}>
        {[["commissions","💰 Комиссии"],["policydb","📋 База полисов"],["officesales","🏢 Продажи офиса"],["cashbook","📒 Касса"],["payroll","📝 Начисления"]].filter(([id])=>isAdmin||id!=="commissions").map(([id,label])=>(
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
          {panel==="access"&&(
            <div style={{fontSize:13}}>
              <h3 style={{margin:"0 0 14px",fontSize:15}}>🔐 Управление доступом</h3>
              <div style={{maxWidth:420,marginBottom:20}}>
                <div style={{fontWeight:600,fontSize:13,marginBottom:10,color:"#374151"}}>
                  {adminPin?"Изменить PIN администратора":"Установить PIN администратора"}
                </div>
                {!adminPin&&<p style={{fontSize:12,color:"#6b7280",marginBottom:12,marginTop:0}}>PIN ещё не установлен. После сохранения при следующем входе потребуется этот PIN.</p>}
                <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap",marginBottom:8}}>
                  <input type="password" value={newPinA} onChange={e=>setNewPinA(e.target.value)} placeholder="Новый PIN" style={{...inp,width:150,padding:"7px 10px",letterSpacing:3,fontSize:16}}/>
                  <input type="password" value={newPinB} onChange={e=>setNewPinB(e.target.value)} onKeyDown={e=>e.key==="Enter"&&changePin()} placeholder="Повторите PIN" style={{...inp,width:150,padding:"7px 10px",letterSpacing:3,fontSize:16}}/>
                  <button onClick={changePin} style={btn("#7c3aed")}>Сохранить</button>
                </div>
                {pinChangeMsg&&<div style={{fontSize:13,color:pinChangeMsg.startsWith("✓")?"#16a34a":"#dc2626"}}>{pinChangeMsg}</div>}
              </div>
              <div style={{borderTop:"1px solid #e5e7eb",paddingTop:16}}>
              <div style={{fontWeight:600,fontSize:13,marginBottom:4,color:"#374151"}}>Операторы (в форме «Продажи офиса»)</div>
              <p style={{fontSize:12,color:"#6b7280",margin:"0 0 10px"}}>Внутренние коды сотрудников, доступных для выбора в форме ввода полиса.</p>
              {officeStaff.length>0&&(
                <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:10}}>
                  {officeStaff.map((code,i)=>{
                    const ag=Object.values(agentDir).find(a=>(a.internalCode||"").trim()===code.trim());
                    const label=ag?(ag.name+" "+ag.surname+" ("+code+")"):code;
                    return(
                      <span key={i} style={{display:"inline-flex",alignItems:"center",gap:5,background:"#eff6ff",border:"1px solid #bfdbfe",borderRadius:6,padding:"4px 10px",fontSize:12,color:"#1e40af",fontWeight:600}}>
                        {label}
                        <button onClick={()=>saveOfficeStaff(officeStaff.filter((_,j)=>j!==i))} style={{background:"none",border:"none",cursor:"pointer",color:"#6b7280",fontSize:14,padding:0,lineHeight:1}}>×</button>
                      </span>
                    );
                  })}
                </div>
              )}
              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                <input value={newStaffCode} onChange={e=>setNewStaffCode(e.target.value)} onKeyDown={e=>e.key==="Enter"&&(newStaffCode.trim()&&!officeStaff.includes(newStaffCode.trim()))&&(saveOfficeStaff([...officeStaff,newStaffCode.trim()]),setNewStaffCode(""))} placeholder="Внутренний код (напр. 768-101)" style={{...inp,width:220,padding:"6px 10px"}}/>
                <button onClick={()=>{if(newStaffCode.trim()&&!officeStaff.includes(newStaffCode.trim())){saveOfficeStaff([...officeStaff,newStaffCode.trim()]);setNewStaffCode("");}}} style={btn("#2563eb")}>+ Добавить</button>
              </div>
              </div>
            </div>
          )}
        </div>
      )}

      {showBackup&&isAdmin&&(
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
            <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(160px,1fr))",gap:8}}>
              {[...ALL_COMPANIES,"voluntary"].map(co=>{
                const isVol=co==="voluntary";
                const file=uploadedFiles.find(f=>f.company===co);
                const volSession_=isVol&&volSession.length>0;
                const volStored=isVol&&storedVol.length>0;
                const savedCount=isVol?storedVol.length:storedPols.filter(p=>p.company===co).length;
                const sessionCount=isVol?volSession.length:file?(file.rows.length-1):0;
                const hasSession=isVol?volSession_:!!file;
                const hasSaved=savedCount>0;
                const label=isVol?"📦 Доброволь.":co;
                // Определяем состояние
                const state=hasSession?"session":hasSaved?"saved":"empty";
                const styles={
                  session:{border:"2px solid #16a34a",bg:"#f0fdf4",labelColor:"#15803d",icon:"📄"},
                  saved:{border:"2px solid #3b82f6",bg:"#eff6ff",labelColor:"#1d4ed8",icon:"💾"},
                  empty:{border:"2px dashed #d1d5db",bg:"white",labelColor:"#6b7280",icon:""},
                }[state];
                return(
                  <div key={co} onClick={()=>handleSlotClick(co)}
                    style={{border:styles.border,borderRadius:8,padding:10,background:styles.bg,minHeight:84,cursor:"pointer",display:"flex",flexDirection:"column",gap:3,position:"relative"}}>
                    <div style={{fontWeight:700,fontSize:13,color:styles.labelColor,display:"flex",alignItems:"center",gap:4}}>
                      {styles.icon&&<span>{styles.icon}</span>}{label}
                    </div>
                    {state==="session"&&(
                      <div>
                        {file&&<div style={{fontSize:10,color:"#6b7280",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",maxWidth:"100%"}}>{file.fileName}</div>}
                        <div style={{fontSize:11,color:"#16a34a",fontWeight:600}}>{sessionCount+" строк"}</div>
                        {hasSaved&&<div style={{fontSize:10,color:"#3b82f6"}}>{"💾 сохр. "+savedCount}</div>}
                        <button onClick={ev=>{ev.stopPropagation();removeFile(co);}} style={{...btn("#fff1f2","#dc2626",{border:"1px solid #fca5a5"}),fontSize:10,padding:"2px 6px",marginTop:4,width:"fit-content"}}>✕ Убрать</button>
                      </div>
                    )}
                    {state==="saved"&&(
                      <div>
                        <div style={{fontSize:11,color:"#1d4ed8",fontWeight:600}}>{savedCount+" полисов"}</div>
                        <div style={{fontSize:10,color:"#6b7280",marginTop:2}}>Сохранено в базе</div>
                        <div style={{fontSize:10,color:"#3b82f6",marginTop:2}}>Нажмите для обновления</div>
                      </div>
                    )}
                    {state==="empty"&&(
                      <div style={{fontSize:11,color:"#9ca3af",marginTop:"auto"}}>Нажмите для загрузки</div>
                    )}
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
                    <thead><tr>{["Продукт","Компания","Агент","Страхователь","Сумма","% О","Ком.О","% А","Ком.А","К выплате"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
                    <tbody>{effVol.map(v=><tr key={v._id}><td style={{...td,fontWeight:600}}>{v.productName}</td><td style={td}>{v.company}</td><td style={td}>{getName(v.agentUid)||v.agentCode}</td><td style={td}>{v.insuredName}</td><td style={td}>{fmt(v.amount)}</td><td style={{...td,color:"#6b7280"}}>{v.officeRate+"%"}</td><td style={td}>{fmt(v.officeComm)}</td><td style={{...td,color:"#6b7280"}}>{v.agentRate+"%"}</td><td style={td}>{fmt(v.agentComm)}</td><td style={{...td,fontWeight:700,color:"#1d4ed8"}}>{fmt(v.amount-v.agentComm)}</td></tr>)}</tbody>
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
            {dbLoaded&&filteredDB.length>0&&(
              <button onClick={()=>exportRenewalsZip(filteredDB,expiryF?fmtMonth(expiryF):"все")} style={btn("#16a34a",undefined,{marginLeft:"auto"})}>⬇ Экспорт по агентам (ZIP)</button>
            )}
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

      {tab==="officesales"&&(()=>{
        const fmtPay=t=>({cash:"💵 Наличные",acba:"🏦 ACBA",ineco:"🏦 INECO"})[t]||t||"—";
        const fmtPolDate=d=>{if(!d)return"—";try{return fmtDate(parseDate(d));}catch{return d;}};
        const finp={...inp,border:"1.5px solid #9ca3af",padding:"8px 10px"};
        const flbl={fontSize:12,color:"#111827",marginBottom:4,fontWeight:500};
        const sortedAgents=Object.entries(agentDir).sort((a,b)=>{
          const ma=(a[1].internalCode||"").match(/(\d+)$/);const mb=(b[1].internalCode||"").match(/(\d+)$/);
          return (ma?parseInt(ma[1]):99999)-(mb?parseInt(mb[1]):99999);
        });
        const allUnpaid=[...opPrevUnpaid,...opCurrentMonth.filter(p=>!p.paid)].sort((a,b)=>new Date(a.date)-new Date(b.date));
        const currentPaid=opCurrentMonth.filter(p=>p.paid).sort((a,b)=>new Date(a.paidAt||0)-new Date(b.paidAt||0));
        const staffAgents=sortedAgents.filter(([,a])=>officeStaff.includes((a.internalCode||"").trim()));
        const tblH={...th,whiteSpace:"nowrap"};
        const actBtn=(label,bg,col,onClick)=><button onClick={onClick} style={{...btn(bg,col,{fontSize:11,padding:"3px 8px"}),marginRight:3}}>{label}</button>;
        const opSrch=opSearch.trim().toLowerCase();
        const matchesText=p=>!opSrch||(p.insuredName||"").toLowerCase().includes(opSrch)||(p.phone||"").includes(opSrch)||(p.policyNum||"").toLowerCase().includes(opSrch)||(p.car||"").toLowerCase().includes(opSrch)||(p.carPlate||"").toLowerCase().includes(opSrch);
        const matchesStatus=p=>opStatusFilter==="all"||(opStatusFilter==="paid"&&p.paid)||(opStatusFilter==="unpaid"&&!p.paid);
        const filterPol=p=>matchesText(p)&&matchesStatus(p);
        const calcTotals=pols=>({count:pols.length,paid:pols.filter(p=>p.paid).length,unpaid:pols.filter(p=>!p.paid).length,totalAmount:pols.reduce((s,p)=>s+(p.amount||0),0),totalNet:pols.reduce((s,p)=>s+(p.amount||0)-(p.discount||0),0),totalPaidAmt:pols.filter(p=>p.paid).reduce((s,p)=>s+(p.paidAmount||0),0)});
        const osagoList=opCurrentMonth.filter(p=>(p.polType||"osago")==="osago");
        const volList=opCurrentMonth.filter(p=>p.polType==="voluntary");

        return(
          <div>
            {/* Month nav + add */}
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16,background:"#f1f5f9",borderRadius:8,padding:"10px 14px",flexWrap:"wrap"}}>
              <button onClick={()=>setSelMonth(prevMo(selMonth))} disabled={selMonth<=MIN_MONTH} style={{...btn("#fff","#374151",{border:"1px solid #d1d5db",fontSize:18,padding:"3px 10px"}),opacity:selMonth<=MIN_MONTH?0.4:1}}>‹</button>
              <span style={{fontWeight:700,fontSize:16,minWidth:160,textAlign:"center"}}>{fmtMonth(selMonth)}</span>
              <button onClick={()=>setSelMonth(nextMo(selMonth))} disabled={selMonth>=MAX_MONTH} style={{...btn("#fff","#374151",{border:"1px solid #d1d5db",fontSize:18,padding:"3px 10px"}),opacity:selMonth>=MAX_MONTH?0.4:1}}>›</button>
              <select value={opStatusFilter} onChange={e=>setOpStatusFilter(e.target.value)} style={{...inp,padding:"6px 10px",fontSize:13,fontWeight:600,marginLeft:8}}>
                <option value="all">Все</option>
                <option value="unpaid">Неоплаченные</option>
                <option value="paid">Оплаченные</option>
              </select>
              <button onClick={openOpNew} style={btn("#2563eb",undefined,{marginLeft:"auto",fontSize:13,padding:"7px 16px"})}>+ Добавить полис</button>
            </div>

            {/* Search */}
            <div style={{marginBottom:12}}>
              <input value={opSearch} onChange={e=>setOpSearch(e.target.value)} placeholder="🔍 Поиск по имени, телефону, № полиса, марке авто, рег. номеру..." style={{...inp,width:"100%",padding:"7px 12px",boxSizing:"border-box"}}/>
            </div>

            {!opLoaded&&<div style={{padding:40,textAlign:"center",color:"#9ca3af"}}>Загрузка...</div>}

            {/* Unpaid from previous months */}
            {opLoaded&&opPrevUnpaid.filter(filterPol).length>0&&(
              <div style={{border:"2px solid #fcd34d",borderRadius:8,overflow:"hidden",marginBottom:20}}>
                <div style={{background:"#fffbeb",padding:"10px 16px",fontWeight:700,fontSize:14,color:"#92400e",display:"flex",alignItems:"center",gap:8}}>
                  <span>⚠ Неоплаченные из предыдущих месяцев — {opPrevUnpaid.length}{opSrch&&opPrevUnpaid.filter(filterPol).length!==opPrevUnpaid.length?" (показано "+opPrevUnpaid.filter(filterPol).length+")":""}</span>
                </div>
                <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                  <thead><tr style={{background:"#fef9c3"}}>{["Тип","Дата сост.","Месяц","Страхователь","Телефон","Компания","Продукт / Авто","Срок","Статус","№ полиса","Сумма","К оплате","Агент","Коммент.",""].map(h=><th key={h} style={tblH}>{h}</th>)}</tr></thead>
                  <tbody>{opPrevUnpaid.filter(filterPol).map((pol,i)=>(
                    <tr key={pol._id} style={{background:i%2===0?"white":"#fefce8",borderBottom:"1px solid #fde68a"}}>
                      <td style={{...td,whiteSpace:"nowrap"}}>{pol.polType==="voluntary"?<span style={{background:"#ede9fe",color:"#6d28d9",borderRadius:10,padding:"1px 7px",fontSize:10,fontWeight:600}}>🛡 Доброволь.</span>:<span style={{background:"#dbeafe",color:"#1e40af",borderRadius:10,padding:"1px 7px",fontSize:10,fontWeight:600}}>🚗 ОСАГО</span>}</td>
                      <td style={{...td,fontSize:11,whiteSpace:"nowrap"}}>{fmtPolDate(pol.date)}</td>
                      <td style={{...td,fontSize:11,color:"#6b7280",whiteSpace:"nowrap"}}>{fmtMonth(pol._monthKey)}</td>
                      <td style={{...td,fontWeight:600}}>{pol.insuredName}</td>
                      <td style={{...td,fontSize:11}}>{pol.phone||"—"}</td>
                      <td style={td}>{pol.company||"—"}</td>
                      <td style={{...td,fontSize:11,color:"#6b7280"}}>{pol.polType==="voluntary"?(pol.productName||"—"):(pol.car||"—")}</td>
                      <td style={{...td,textAlign:"center"}}>{pol.term?<span style={{background:pol.term==="L"?"#dbeafe":"#fef3c7",color:pol.term==="L"?"#1d4ed8":"#92400e",borderRadius:10,padding:"1px 7px",fontSize:11,fontWeight:600}}>{pol.term}</span>:"—"}</td>
                      <td style={{...td,fontSize:11}}>{pol.polStatus?<span style={{background:"#f3e8ff",color:"#6d28d9",borderRadius:10,padding:"1px 6px",fontSize:10,fontWeight:600}}>{fmtPolStatus(pol.polStatus)}</span>:"—"}</td>
                      <td style={{...td,fontSize:11}}>{pol.policyNum||"—"}</td>
                      <td style={{...td,textAlign:"right"}}>{fmt(pol.amount)}</td>
                      <td style={{...td,textAlign:"right",fontWeight:700,color:"#1d4ed8"}}>{fmt((pol.amount||0)-(pol.discount||0))}</td>
                      <td style={{...td,fontSize:11}}>{getName(pol.agentUid)||"—"}</td>
                      <td style={{...td,fontSize:11,color:"#6b7280",maxWidth:100,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{pol.comment||""}</td>
                      <td style={{...td,whiteSpace:"nowrap"}}>
                        {actBtn("✎","#f3f4f6","#374151",()=>openOpEdit(pol))}
                        {actBtn("✓ Оплата","#16a34a","#fff",()=>openOpPay(pol))}
                        {isAdmin&&actBtn("✕","#fff1f2","#dc2626",()=>deleteOfficePol(pol))}
                      </td>
                    </tr>
                  ))}</tbody>
                </table></div>
              </div>
            )}

            {/* Current month — ОСАГО */}
            {opLoaded&&(()=>{
              const list=osagoList.filter(filterPol);
              const t=calcTotals(osagoList);
              return(
              <div style={{border:"1px solid #dbeafe",borderRadius:8,overflow:"hidden",marginBottom:16}}>
                <div style={{background:"#eff6ff",padding:"10px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:6}}>
                  <span style={{fontWeight:700,fontSize:14,color:"#1e40af"}}>🚗 ОСАГО — {fmtMonth(selMonth)}</span>
                  <div style={{display:"flex",gap:16,fontSize:12,flexWrap:"wrap"}}>
                    <span style={{color:"#374151"}}>Полисов: <strong>{t.count}</strong></span>
                    <span style={{color:"#374151"}}>Сумма: <strong>{fmt(t.totalAmount)}</strong></span>
                    <span style={{color:"#374151"}}>К оплате: <strong>{fmt(t.totalNet)}</strong></span>
                    <span style={{color:"#16a34a"}}>Получено: <strong>{fmt(t.totalPaidAmt)}</strong></span>
                    {t.unpaid>0&&<span style={{color:"#dc2626"}}>Ожидает: <strong>{t.unpaid}</strong></span>}
                  </div>
                </div>
                {osagoList.length===0
                  ?<div style={{padding:24,textAlign:"center",color:"#9ca3af",fontSize:13}}>Нет полисов ОСАГО за {fmtMonth(selMonth)}</div>
                  :list.length===0
                  ?<div style={{padding:24,textAlign:"center",color:"#9ca3af",fontSize:13}}>Нет совпадений по фильтру</div>
                  :<div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead><tr style={{background:"#dbeafe"}}>{["Статус","Дата сост.","Страхователь","Телефон","Компания","Марка авто","Рег.номер","Срок","Ст-с","№ полиса","Сумма","К оплате","Оператор","Оплачено","Тип оплаты",""].map(h=><th key={h} style={tblH}>{h}</th>)}</tr></thead>
                    <tbody>{list.map((pol,i)=>(
                      <tr key={pol._id} style={{background:pol.paid?(i%2===0?"#f0fdf4":"#dcfce7"):(i%2===0?"white":"#fafafa"),borderBottom:"1px solid #e5e7eb"}}>
                        <td style={{...td,whiteSpace:"nowrap"}}>{pol.paid?<span style={{background:"#dcfce7",color:"#166534",borderRadius:12,padding:"2px 8px",fontSize:11,fontWeight:600}}>✓ Оплачен</span>:<span style={{background:"#fef9c3",color:"#92400e",borderRadius:12,padding:"2px 8px",fontSize:11,fontWeight:600}}>⏳ Ожидает</span>}</td>
                        <td style={{...td,fontSize:11,whiteSpace:"nowrap"}}>{fmtPolDate(pol.date)}</td>
                        <td style={{...td,fontWeight:600}}>{pol.insuredName}</td>
                        <td style={{...td,fontSize:11}}>{pol.phone||"—"}</td>
                        <td style={td}>{pol.company}</td>
                        <td style={{...td,fontSize:11,color:"#6b7280"}}>{pol.car||"—"}</td>
                        <td style={{...td,fontSize:11}}>{pol.carPlate||"—"}</td>
                        <td style={{...td,textAlign:"center"}}>{pol.term?<span style={{background:pol.term==="L"?"#dbeafe":"#fef3c7",color:pol.term==="L"?"#1d4ed8":"#92400e",borderRadius:10,padding:"1px 7px",fontSize:11,fontWeight:600}}>{pol.term}</span>:"—"}</td>
                        <td style={{...td,fontSize:11}}>{pol.polStatus?<span style={{background:"#f3e8ff",color:"#6d28d9",borderRadius:10,padding:"1px 6px",fontSize:10,fontWeight:600}}>{fmtPolStatus(pol.polStatus)}</span>:"—"}</td>
                        <td style={{...td,fontSize:11}}>{pol.policyNum||"—"}</td>
                        <td style={{...td,textAlign:"right"}}>{fmt(pol.amount)}</td>
                        <td style={{...td,textAlign:"right",fontWeight:700}}>{fmt((pol.amount||0)-(pol.discount||0))}</td>
                        <td style={{...td,fontSize:11}}>{getName(pol.agentUid)||"—"}</td>
                        <td style={{...td,fontSize:11,whiteSpace:"nowrap"}}>{pol.paid?<><div style={{fontWeight:600}}>{fmt(pol.paidAmount||0)}</div><div style={{fontSize:10,color:"#6b7280"}}>{fmtPolDate(pol.paidDate)}</div></>:"—"}</td>
                        <td style={{...td,fontSize:11,whiteSpace:"nowrap"}}>{pol.paid?fmtPay(pol.paymentType):"—"}</td>
                        <td style={{...td,whiteSpace:"nowrap"}}>
                          {!pol.paid&&actBtn("✎","#f3f4f6","#374151",()=>openOpEdit(pol))}
                          {!pol.paid&&actBtn("✓ Оплата","#16a34a","#fff",()=>openOpPay(pol))}
                          {isAdmin&&actBtn("✕","#fff1f2","#dc2626",()=>deleteOfficePol(pol))}
                        </td>
                      </tr>
                    ))}</tbody>
                  </table></div>
                }
              </div>
              );
            })()}

            {/* Current month — Добровольные */}
            {opLoaded&&(()=>{
              const list=volList.filter(filterPol);
              const t=calcTotals(volList);
              return(
              <div style={{border:"1px solid #e9d5ff",borderRadius:8,overflow:"hidden",marginBottom:20}}>
                <div style={{background:"#f5f3ff",padding:"10px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:6}}>
                  <span style={{fontWeight:700,fontSize:14,color:"#6d28d9"}}>🛡 Добровольные — {fmtMonth(selMonth)}</span>
                  <div style={{display:"flex",gap:16,fontSize:12,flexWrap:"wrap"}}>
                    <span style={{color:"#374151"}}>Полисов: <strong>{t.count}</strong></span>
                    <span style={{color:"#374151"}}>Сумма: <strong>{fmt(t.totalAmount)}</strong></span>
                    <span style={{color:"#374151"}}>К оплате: <strong>{fmt(t.totalNet)}</strong></span>
                    <span style={{color:"#16a34a"}}>Получено: <strong>{fmt(t.totalPaidAmt)}</strong></span>
                    {t.unpaid>0&&<span style={{color:"#dc2626"}}>Ожидает: <strong>{t.unpaid}</strong></span>}
                  </div>
                </div>
                {volList.length===0
                  ?<div style={{padding:24,textAlign:"center",color:"#9ca3af",fontSize:13}}>Нет добровольных полисов за {fmtMonth(selMonth)}</div>
                  :list.length===0
                  ?<div style={{padding:24,textAlign:"center",color:"#9ca3af",fontSize:13}}>Нет совпадений по фильтру</div>
                  :<div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead><tr style={{background:"#ede9fe"}}>{["Статус","Дата сост.","Страхователь","Телефон","Продукт","Компания","№ полиса","Сумма","К оплате","Оператор","Оплачено","Тип оплаты",""].map(h=><th key={h} style={tblH}>{h}</th>)}</tr></thead>
                    <tbody>{list.map((pol,i)=>(
                      <tr key={pol._id} style={{background:pol.paid?(i%2===0?"#f0fdf4":"#dcfce7"):(i%2===0?"white":"#fafafa"),borderBottom:"1px solid #e5e7eb"}}>
                        <td style={{...td,whiteSpace:"nowrap"}}>{pol.paid?<span style={{background:"#dcfce7",color:"#166534",borderRadius:12,padding:"2px 8px",fontSize:11,fontWeight:600}}>✓ Оплачен</span>:<span style={{background:"#fef9c3",color:"#92400e",borderRadius:12,padding:"2px 8px",fontSize:11,fontWeight:600}}>⏳ Ожидает</span>}</td>
                        <td style={{...td,fontSize:11,whiteSpace:"nowrap"}}>{fmtPolDate(pol.date)}</td>
                        <td style={{...td,fontWeight:600}}>{pol.insuredName}</td>
                        <td style={{...td,fontSize:11}}>{pol.phone||"—"}</td>
                        <td style={{...td,fontSize:11,color:"#6d28d9",fontWeight:500}}>{pol.productName||"—"}</td>
                        <td style={td}>{pol.company||"—"}</td>
                        <td style={{...td,fontSize:11}}>{pol.policyNum||"—"}</td>
                        <td style={{...td,textAlign:"right"}}>{fmt(pol.amount)}</td>
                        <td style={{...td,textAlign:"right",fontWeight:700}}>{fmt((pol.amount||0)-(pol.discount||0))}</td>
                        <td style={{...td,fontSize:11}}>{getName(pol.agentUid)||"—"}</td>
                        <td style={{...td,fontSize:11,whiteSpace:"nowrap"}}>{pol.paid?<><div style={{fontWeight:600}}>{fmt(pol.paidAmount||0)}</div><div style={{fontSize:10,color:"#6b7280"}}>{fmtPolDate(pol.paidDate)}</div></>:"—"}</td>
                        <td style={{...td,fontSize:11,whiteSpace:"nowrap"}}>{pol.paid?fmtPay(pol.paymentType):"—"}</td>
                        <td style={{...td,whiteSpace:"nowrap"}}>
                          {!pol.paid&&actBtn("✎","#f3f4f6","#374151",()=>openOpEdit(pol))}
                          {!pol.paid&&actBtn("✓ Оплата","#16a34a","#fff",()=>openOpPay(pol))}
                          {isAdmin&&actBtn("✕","#fff1f2","#dc2626",()=>deleteOfficePol(pol))}
                        </td>
                      </tr>
                    ))}</tbody>
                  </table></div>
                }
              </div>
              );
            })()}

            {/* Empty state */}
            {opLoaded&&osagoList.length===0&&volList.length===0&&opPrevUnpaid.length===0&&(
              <div style={{padding:48,textAlign:"center",color:"#9ca3af",fontSize:14,border:"2px dashed #e5e7eb",borderRadius:8}}>
                <div style={{fontSize:32,marginBottom:8}}>🏢</div>
                <div>Нет полисов. Нажмите «+ Добавить полис» для начала работы.</div>
              </div>
            )}

            {/* Policy form modal */}
            {opFormOpen&&(
              <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}
                onClick={e=>{if(e.target===e.currentTarget)setOpFormOpen(false);}}>
                <div style={{background:"white",borderRadius:12,padding:24,width:"100%",maxWidth:580,maxHeight:"92vh",overflowY:"auto",boxShadow:"0 20px 60px rgba(0,0,0,0.3)"}}>
                  <h3 style={{margin:"0 0 12px",fontSize:16}}>{opEditPol?"✎ Редактировать полис":"➕ Новый полис"}</h3>
                  {/* --- Тип продукта --- */}
                  <div style={{display:"flex",gap:0,marginBottom:16,borderRadius:8,overflow:"hidden",border:"2px solid #e5e7eb"}}>
                    {[["osago","🚗 ОСАГО"],["voluntary","🛡 Добровольный"]].map(([k,l])=>(
                      <button key={k} onClick={()=>setOpFD(p=>({...initOpFD(),polType:k,insuredName:p.insuredName,phone:p.phone,policyNum:p.policyNum,agentUid:p.agentUid,date:p.date,comment:p.comment}))}
                        style={{flex:1,padding:"9px 0",fontSize:13,fontWeight:600,cursor:"pointer",border:"none",background:opFD.polType===k?"#2563eb":"#f8fafc",color:opFD.polType===k?"#fff":"#6b7280",transition:"all .15s"}}>
                        {l}
                      </button>
                    ))}
                  </div>
                  {/* --- Страхователь --- */}
                  <div style={{fontSize:11,fontWeight:700,color:"#374151",marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>Страхователь</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
                    <div style={{gridColumn:"1/-1"}}>
                      <div style={flbl}>ФИО страхователя *</div>
                      <input value={opFD.insuredName||""} onChange={e=>setOpFD(p=>({...p,insuredName:e.target.value}))} placeholder="Имя Фамилия" style={{...finp,width:"100%",boxSizing:"border-box",fontSize:14}}/>
                    </div>
                    <div>
                      <div style={flbl}>№ полиса</div>
                      <input value={opFD.policyNum||""} onChange={e=>setOpFD(p=>({...p,policyNum:e.target.value}))} placeholder="Номер полиса" style={{...finp,width:"100%",boxSizing:"border-box"}}/>
                    </div>
                    <div>
                      <div style={flbl}>Телефон</div>
                      <input value={opFD.phone||""} onChange={e=>setOpFD(p=>({...p,phone:e.target.value}))} placeholder="+374..." style={{...finp,width:"100%",boxSizing:"border-box"}}/>
                    </div>
                  </div>
                  {/* === ОСАГО: Полис + Авто + Статус === */}
                  {opFD.polType==="osago"&&(<>
                  <div style={{fontSize:11,fontWeight:700,color:"#374151",marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>Полис</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
                    <div>
                      <div style={flbl}>Компания *</div>
                      <select value={opFD.company||""} onChange={e=>setOpFD(p=>({...p,company:e.target.value}))} style={{...finp,width:"100%",boxSizing:"border-box"}}>
                        {ALL_COMPANIES.map(c=><option key={c}>{c}</option>)}
                      </select>
                    </div>
                    <div>
                      <div style={flbl}>Оператор (принял полис)</div>
                      <select value={opFD.agentUid||""} onChange={e=>setOpFD(p=>({...p,agentUid:e.target.value}))} style={{...finp,width:"100%",boxSizing:"border-box"}}>
                        <option value="">— Не указан —</option>
                        {staffAgents.map(([id,a])=><option key={id} value={id}>{a.name+" "+a.surname+(a.internalCode?" ("+a.internalCode+")":"")}</option>)}
                      </select>
                    </div>
                    <div>
                      <div style={flbl}>Дата составления *</div>
                      <input type="date" value={opFD.date||""} onChange={e=>setOpFD(p=>({...p,date:e.target.value}))} style={{...finp,width:"100%",boxSizing:"border-box"}}/>
                    </div>
                    <div>
                      <div style={flbl}>Срок</div>
                      <div style={{display:"flex",gap:6}}>
                        {[["L","L (от 3 месяцев)"],["SH","SH (краткосрочный)"]].map(([k,l])=>(
                          <button key={k} onClick={()=>setOpFD(p=>({...p,term:k}))}
                            style={{...btn(opFD.term===k?"#1d4ed8":"#f3f4f6",opFD.term===k?"#fff":"#374151",{flex:1,border:"2px solid "+(opFD.term===k?"#1d4ed8":"#e5e7eb"),fontSize:12})}}>
                            {l}
                          </button>
                        ))}
                      </div>
                    </div>
                    <div>
                      <div style={flbl}>Дата вступления в силу</div>
                      <input type="date" value={opFD.dateStart||""} onChange={e=>setOpFD(p=>({...p,dateStart:e.target.value}))} style={{...finp,width:"100%",boxSizing:"border-box"}}/>
                    </div>
                    <div>
                      <div style={flbl}>Дата окончания</div>
                      <input type="date" value={opFD.dateEnd||""} onChange={e=>setOpFD(p=>({...p,dateEnd:e.target.value}))} style={{...finp,width:"100%",boxSizing:"border-box"}}/>
                    </div>
                  </div>
                  <div style={{fontSize:11,fontWeight:700,color:"#374151",marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>Транспортное средство</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
                    <div>
                      <div style={flbl}>Марка / модель авто</div>
                      <input value={opFD.car||""} onChange={e=>setOpFD(p=>({...p,car:e.target.value}))} placeholder="Toyota Camry..." style={{...finp,width:"100%",boxSizing:"border-box"}}/>
                    </div>
                    <div>
                      <div style={flbl}>Рег. номер</div>
                      <input value={opFD.carPlate||""} onChange={e=>setOpFD(p=>({...p,carPlate:e.target.value}))} placeholder="00 AA 000" style={{...finp,width:"100%",boxSizing:"border-box"}}/>
                    </div>
                    <div>
                      <div style={flbl}>Регион</div>
                      <select value={opFD.region||""} onChange={e=>setOpFD(p=>({...p,region:e.target.value}))} style={{...finp,width:"100%",boxSizing:"border-box",background:"white"}}>
                        <option value="">— выбрать —</option>
                        {["YR","AG","AR","AV","GH","LO","KT","SH","SY","VD","TV"].map(r=><option key={r} value={r}>{r}</option>)}
                      </select>
                    </div>
                    <div>
                      <div style={flbl}>БМ</div>
                      <select value={opFD.bm||""} onChange={e=>setOpFD(p=>({...p,bm:e.target.value}))} style={{...finp,width:"100%",boxSizing:"border-box",background:"white"}}>
                        <option value="">— выбрать —</option>
                        {Array.from({length:25},(_,i)=>i+1).map(n=><option key={n} value={n}>{n}</option>)}
                      </select>
                    </div>
                    <div>
                      <div style={flbl}>Мощность (л.с.)</div>
                      <input type="number" value={opFD.power||""} onChange={e=>setOpFD(p=>({...p,power:e.target.value}))} placeholder="л.с." style={{...finp,width:"100%",boxSizing:"border-box"}}/>
                    </div>
                  </div>
                  <div style={{fontSize:11,fontWeight:700,color:"#374151",marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>Статус полиса</div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:14}}>
                    {POL_STATUSES.map(({k,l})=>(
                      <button key={k} onClick={()=>setOpFD(p=>({...p,polStatus:k}))}
                        style={{...btn(opFD.polStatus===k?"#7c3aed":"#f3f4f6",opFD.polStatus===k?"#fff":"#374151",{border:"2px solid "+(opFD.polStatus===k?"#7c3aed":"#e5e7eb"),fontSize:12,padding:"5px 12px"})}}>
                        {l}
                      </button>
                    ))}
                  </div>
                  </>)}
                  {/* === Добровольный: Продукт + Компания + Оператор + Дата === */}
                  {opFD.polType==="voluntary"&&(<>
                  <div style={{fontSize:11,fontWeight:700,color:"#374151",marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>Продукт</div>
                  <div style={{marginBottom:14}}>
                    {(volRates.rates||[]).length>0
                      ?<select value={opFD.productName||""} onChange={e=>setOpFD(p=>({...p,productName:e.target.value}))} style={{...finp,width:"100%",boxSizing:"border-box",background:"white"}}>
                          <option value="">— выбрать продукт —</option>
                          {(volRates.rates||[]).map(r=><option key={r.name} value={r.name}>{r.name}</option>)}
                        </select>
                      :<input value={opFD.productName||""} onChange={e=>setOpFD(p=>({...p,productName:e.target.value}))} placeholder="Название продукта" style={{...finp,width:"100%",boxSizing:"border-box"}}/>
                    }
                  </div>
                  <div style={{fontSize:11,fontWeight:700,color:"#374151",marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>Полис</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
                    <div>
                      <div style={flbl}>Компания *</div>
                      <select value={opFD.company||""} onChange={e=>setOpFD(p=>({...p,company:e.target.value}))} style={{...finp,width:"100%",boxSizing:"border-box"}}>
                        {ALL_COMPANIES.map(c=><option key={c}>{c}</option>)}
                      </select>
                    </div>
                    <div>
                      <div style={flbl}>Оператор (принял полис)</div>
                      <select value={opFD.agentUid||""} onChange={e=>setOpFD(p=>({...p,agentUid:e.target.value}))} style={{...finp,width:"100%",boxSizing:"border-box"}}>
                        <option value="">— Не указан —</option>
                        {staffAgents.map(([id,a])=><option key={id} value={id}>{a.name+" "+a.surname+(a.internalCode?" ("+a.internalCode+")":"")}</option>)}
                      </select>
                    </div>
                    <div style={{gridColumn:"1/-1"}}>
                      <div style={flbl}>Дата составления *</div>
                      <input type="date" value={opFD.date||""} onChange={e=>setOpFD(p=>({...p,date:e.target.value}))} style={{...finp,width:"100%",boxSizing:"border-box"}}/>
                    </div>
                  </div>
                  </>)}
                  {/* === Финансы (общее) === */}
                  <div style={{fontSize:11,fontWeight:700,color:"#374151",marginBottom:6,textTransform:"uppercase",letterSpacing:.5}}>Финансы</div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:14}}>
                    <div>
                      <div style={flbl}>Страховая премия (AMD) *</div>
                      <input type="number" value={opFD.amount||""} onChange={e=>setOpFD(p=>({...p,amount:e.target.value}))} placeholder="0" style={{...finp,width:"100%",boxSizing:"border-box",textAlign:"right"}}/>
                    </div>
                    <div>
                      <div style={flbl}>Скидка (AMD)</div>
                      <input type="number" value={opFD.discount||"0"} onChange={e=>setOpFD(p=>({...p,discount:e.target.value}))} placeholder="0" style={{...finp,width:"100%",boxSizing:"border-box",textAlign:"right"}}/>
                    </div>
                    <div style={{gridColumn:"1/-1",background:"#f0fdf4",border:"1px solid #86efac",borderRadius:6,padding:"8px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <span style={{fontSize:12,color:"#166534"}}>К оплате клиентом:</span>
                      <span style={{fontSize:17,fontWeight:700,color:"#16a34a"}}>{fmt((parseFloat(opFD.amount)||0)-(parseFloat(opFD.discount)||0))}</span>
                    </div>
                  </div>
                  {/* === Добровольный: принятие оплаты сразу === */}
                  {opFD.polType==="voluntary"&&(
                  <div style={{marginBottom:14,border:"1px solid #e5e7eb",borderRadius:8,overflow:"hidden"}}>
                    <div style={{background:"#f8fafc",padding:"8px 14px",display:"flex",alignItems:"center",gap:10}}>
                      <input type="checkbox" id="payNowCb" checked={!!opFD.payNow} onChange={e=>setOpFD(p=>({...p,payNow:e.target.checked,paymentType:e.target.checked?(p.paymentType||"cash"):""}))} style={{width:16,height:16,cursor:"pointer"}}/>
                      <label htmlFor="payNowCb" style={{fontSize:13,fontWeight:600,color:"#374151",cursor:"pointer"}}>Принять оплату сразу</label>
                    </div>
                    {opFD.payNow&&(
                      <div style={{padding:"10px 14px"}}>
                        <div style={{fontSize:11,color:"#6b7280",marginBottom:6}}>Способ оплаты *</div>
                        <div style={{display:"flex",gap:8}}>
                          {[["cash","💵 Наличные"],["acba","🏦 ACBA"],["ineco","🏦 INECO"]].map(([k,l])=>(
                            <button key={k} onClick={()=>setOpFD(p=>({...p,paymentType:k}))}
                              style={{...btn(opFD.paymentType===k?"#1d4ed8":"#f3f4f6",opFD.paymentType===k?"#fff":"#374151",{flex:1,border:"2px solid "+(opFD.paymentType===k?"#1d4ed8":"#e5e7eb"),fontSize:12})}}>
                              {l}
                            </button>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                  )}
                  {/* --- Комментарий --- */}
                  <div style={{marginBottom:16}}>
                    <div style={flbl}>Комментарий</div>
                    <input value={opFD.comment||""} onChange={e=>setOpFD(p=>({...p,comment:e.target.value}))} onKeyDown={e=>e.key==="Enter"&&submitOpForm()} placeholder="Необязательно" style={{...finp,width:"100%",boxSizing:"border-box"}}/>
                  </div>
                  <div style={{display:"flex",gap:8}}>
                    <button onClick={submitOpForm} style={{...btn("#2563eb"),flex:1,padding:"10px",fontSize:14}}>{opEditPol?"Сохранить изменения":"Добавить полис"}</button>
                    <button onClick={()=>setOpFormOpen(false)} style={{...btn("#f3f4f6","#374151"),flex:1,padding:"10px",fontSize:14}}>Отмена</button>
                  </div>
                </div>
              </div>
            )}

            {/* Payment modal */}
            {opPayPol&&(
              <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
                <div style={{background:"white",borderRadius:12,padding:24,width:"100%",maxWidth:400,boxShadow:"0 20px 60px rgba(0,0,0,0.3)"}}>
                  <div style={{textAlign:"center",fontSize:28,marginBottom:6}}>✓</div>
                  <h3 style={{margin:"0 0 16px",fontSize:16,textAlign:"center"}}>Принять оплату</h3>
                  <div style={{background:"#f8fafc",borderRadius:8,padding:"10px 14px",marginBottom:16,fontSize:13}}>
                    <div style={{fontWeight:600,fontSize:14,marginBottom:4}}>{opPayPol.insuredName}</div>
                    <div style={{color:"#6b7280"}}>{opPayPol.company}{opPayPol.policyNum?" · "+opPayPol.policyNum:""}{opPayPol._monthKey&&opPayPol._monthKey!==selMonth?<span style={{marginLeft:8,background:"#fef9c3",color:"#92400e",borderRadius:10,padding:"1px 6px",fontSize:11}}>{"из "+fmtMonth(opPayPol._monthKey)}</span>:null}</div>
                    <div style={{marginTop:4,fontSize:12,color:"#374151"}}>Сумма: <strong>{fmt(opPayPol.amount)}</strong>{opPayPol.discount>0?<> — скидка <strong style={{color:"#dc2626"}}>{fmt(opPayPol.discount)}</strong></>:null}</div>
                  </div>
                  <div style={{marginBottom:12}}>
                    <div style={{fontSize:11,color:"#6b7280",marginBottom:4}}>Сумма к оплате (AMD)</div>
                    <div style={{width:"100%",padding:"10px 12px",fontSize:18,fontWeight:700,textAlign:"right",boxSizing:"border-box",background:"#eff6ff",color:"#1d4ed8",borderRadius:6,border:"1.5px solid #bfdbfe"}}>{fmt(parseFloat(opPayData.paidAmount)||0)}</div>
                  </div>
                  <div style={{marginBottom:12}}>
                    <div style={{fontSize:11,color:"#6b7280",marginBottom:6}}>Способ оплаты *</div>
                    <div style={{display:"flex",gap:8}}>
                      {[["cash","💵 Наличные"],["acba","🏦 ACBA"],["ineco","🏦 INECO"]].map(([k,l])=>(
                        <button key={k} onClick={()=>setOpPayData(p=>({...p,paymentType:k}))}
                          style={{...btn(opPayData.paymentType===k?"#1d4ed8":"#f3f4f6",opPayData.paymentType===k?"#fff":"#374151",{flex:1,border:"2px solid "+(opPayData.paymentType===k?"#1d4ed8":"#e5e7eb"),fontSize:12})}}>
                          {l}
                        </button>
                      ))}
                    </div>
                  </div>
                  <div style={{marginBottom:14}}>
                    <div style={{fontSize:11,color:"#6b7280",marginBottom:4}}>Дата оплаты</div>
                    <div style={{padding:"7px 10px",background:"#f1f5f9",borderRadius:6,border:"1.5px solid #e5e7eb",fontSize:14,color:"#374151",fontWeight:500}}>{opPayData.paidDate||new Date().toISOString().slice(0,10)}</div>
                  </div>
                  <div style={{background:"#fff7ed",border:"1px solid #fed7aa",borderRadius:6,padding:"8px 12px",marginBottom:14,fontSize:12,color:"#9a3412"}}>
                    ⚠ После подтверждения изменить данные оплаты будет невозможно
                  </div>
                  <div style={{display:"flex",gap:8}}>
                    <button onClick={()=>acceptOpPayment(opPayPol,{paidAmount:parseFloat(opPayData.paidAmount)||0,paymentType:opPayData.paymentType,paidDate:opPayData.paidDate})} style={{...btn("#16a34a"),flex:1,padding:"10px",fontSize:14}}>✓ Подтвердить</button>
                    <button onClick={()=>setOpPayPol(null)} style={{...btn("#f3f4f6","#374151"),flex:1,padding:"10px",fontSize:14}}>Отмена</button>
                  </div>
                </div>
              </div>
            )}
          </div>
        );
      })()}

      {tab==="cashbook"&&(()=>{
        const today=new Date().toISOString().slice(0,10);
        const fmtDay=d=>{try{const[y,m,dd]=d.split("-");return dd+"."+m+"."+y;}catch{return d;}};
        const fmtDatetime=iso=>{try{const d=new Date(iso);return d.toLocaleDateString("ru-RU")+", "+d.toLocaleTimeString("ru-RU",{hour:"2-digit",minute:"2-digit"});}catch{return"";}};
        const fmtTime=iso=>{try{return new Date(iso).toLocaleTimeString("ru-RU",{hour:"2-digit",minute:"2-digit"});}catch{return"";}};
        const payLabel=t=>({cash:"💵 Наличные",acba:"🏦 ACBA",ineco:"🏦 INECO"})[t]||t||"—";
        const polTypeBadge=p=>p.polType==="voluntary"
          ?<span style={{background:"#ede9fe",color:"#6d28d9",borderRadius:10,padding:"1px 6px",fontSize:10,fontWeight:600}}>🛡 Доброволь.</span>
          :<span style={{background:"#dbeafe",color:"#1e40af",borderRadius:10,padding:"1px 6px",fontSize:10,fontWeight:600}}>🚗 ОСАГО</span>;

        // Group live payments by paidDate
        const byDay={};
        cashMonthPols.filter(p=>p.paid&&p.paidDate).forEach(p=>{
          const d=p.paidDate;
          if(!byDay[d])byDay[d]={cash:0,acba:0,ineco:0,pols:[]};
          const amt=p.paidAmount||0;
          if(p.paymentType==="acba")byDay[d].acba+=amt;
          else if(p.paymentType==="ineco")byDay[d].ineco+=amt;
          else byDay[d].cash+=amt;
          byDay[d].pols.push(p);
        });
        if(today.startsWith(selMonth)&&!byDay[today])byDay[today]={cash:0,acba:0,ineco:0,pols:[]};

        // For closed days use snapshot totals; for open use live
        const getTotals=date=>{
          const cd=cashDays[date];
          if(cd&&cd.closed&&cd.totals)return cd.totals;
          const d=byDay[date]||{cash:0,acba:0,ineco:0};
          return{cash:d.cash,acba:d.acba,ineco:d.ineco,total:(d.cash+d.acba+d.ineco)};
        };
        const getPols=date=>{
          const cd=cashDays[date];
          if(cd&&cd.closed&&cd.snapshot)return cd.snapshot;
          return(byDay[date]||{pols:[]}).pols;
        };

        // All days: union of live byDay keys + closed cashDays keys
        const allDayKeys=new Set([...Object.keys(byDay),...Object.keys(cashDays)]);
        const days=[...allDayKeys].filter(d=>d.startsWith(selMonth)).sort().reverse();

        // Month totals (sum of all days — closed use snapshot, open use live)
        const mTotals=days.reduce((acc,d)=>{const t=getTotals(d);acc.cash+=t.cash;acc.acba+=t.acba;acc.ineco+=t.ineco;return acc;},{cash:0,acba:0,ineco:0});
        const mTotal=mTotals.cash+mTotals.acba+mTotals.ineco;

        const card=(label,val,color)=>(
          <div style={{flex:1,minWidth:140,background:"white",border:"1px solid #e5e7eb",borderRadius:8,padding:"12px 16px",boxShadow:"0 1px 3px rgba(0,0,0,.06)"}}>
            <div style={{fontSize:11,color:"#6b7280",marginBottom:4}}>{label}</div>
            <div style={{fontSize:20,fontWeight:700,color}}>{fmt(val)}</div>
          </div>
        );

        const fmtPolDate2=d=>{if(!d)return"—";try{const[y,m,dd]=d.split("-");return dd+"."+m+"."+y;}catch{return d;}};
        const polTypeLabel=p=>p.polType==="voluntary"?(p.productName||"Добровольный"):"ОСАГО";

        const polsTable=(pols,compact)=>(
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:compact?11:12}}>
            <thead><tr style={{background:"#f1f5f9"}}>{["Страхователь","Тип","Компания","№ полиса","Дата заключения","Сумма","Скидка","К оплате","Способ","Заметка"].map(h=><th key={h} style={{...th,fontSize:10,padding:"5px 8px"}}>{h}</th>)}</tr></thead>
            <tbody>{pols.map((p,i)=>(
              <tr key={p._id||i} style={{background:i%2===0?"white":"#f8fafc",borderBottom:"1px solid #f3f4f6"}}>
                <td style={{...td,fontWeight:600,whiteSpace:"nowrap"}}>{p.insuredName||"—"}</td>
                <td style={{...td,fontSize:11,whiteSpace:"nowrap"}}>{polTypeLabel(p)}</td>
                <td style={{...td,fontSize:11}}>{p.company||"—"}</td>
                <td style={{...td,fontSize:11}}>{p.policyNum||"—"}</td>
                <td style={{...td,fontSize:11,whiteSpace:"nowrap"}}>{fmtPolDate2(p.date)}</td>
                <td style={{...td,textAlign:"right"}}>{fmt(p.amount||0)}</td>
                <td style={{...td,textAlign:"right",color:"#dc2626"}}>{(p.discount||0)>0?fmt(p.discount):"—"}</td>
                <td style={{...td,textAlign:"right",fontWeight:700,color:"#16a34a"}}>{fmt(p.paidAmount||0)}</td>
                <td style={{...td,whiteSpace:"nowrap",fontSize:11}}>{payLabel(p.paymentType)}</td>
                <td style={{...td,fontSize:11,color:"#6b7280"}}>{p.comment||"—"}</td>
              </tr>
            ))}</tbody>
          </table>
        );

        const printCashReport=(date)=>{
          const pols=getPols(date);
          const t=getTotals(date);
          const isClosed=cashDays[date]&&cashDays[date].closed;
          const fmtAmd=v=>(v||0).toLocaleString("ru-RU")+" AMD";
          const pTypeLabel=p=>p.polType==="voluntary"?(p.productName||"Добровольный"):"ОСАГО";
          const rows=pols.map(p=>`<tr><td>${p.insuredName||"—"}</td><td>${pTypeLabel(p)}</td><td>${p.company||"—"}</td><td>${p.policyNum||"—"}</td><td>${fmtPolDate2(p.date)}</td><td style="text-align:right">${fmtAmd(p.amount)}</td><td style="text-align:right">${(p.discount||0)>0?fmtAmd(p.discount):"—"}</td><td style="text-align:right;font-weight:700">${fmtAmd(p.paidAmount)}</td><td>${{cash:"Наличные",acba:"ACBA",ineco:"INECO"}[p.paymentType]||"—"}</td><td>${p.comment||"—"}</td></tr>`).join("");
          const html=`<!DOCTYPE html><html><head><meta charset="utf-8"><title>Касса ${fmtDay(date)}</title><style>@page{size:landscape;margin:15mm}body{font-family:Arial,sans-serif;font-size:11px;padding:16px;color:#111}h2{margin:0 0 2px;font-size:16px}p.sub{margin:0 0 14px;color:#666;font-size:11px}table{width:100%;border-collapse:collapse;margin-bottom:16px}th,td{border:1px solid #ccc;padding:4px 7px;text-align:left}th{background:#f1f5f9;font-size:10px;font-weight:600}.totals{border:2px solid #374151;border-radius:4px;padding:10px 14px;max-width:320px}.totals .row{display:flex;justify-content:space-between;padding:3px 0;font-size:12px}.totals .grand{font-size:15px;font-weight:700;border-top:2px solid #374151;margin-top:5px;padding-top:7px}.footer{margin-top:16px;font-size:10px;color:#999}@media print{.no-print{display:none}}</style></head><body><h2>Кассовый отчёт</h2><p class="sub">${fmtDay(date)}${isClosed?" · Касса закрыта "+fmtDatetime(cashDays[date].closedAt):""}</p><table><thead><tr><th>Страхователь</th><th>Тип</th><th>Компания</th><th>№ полиса</th><th>Дата заключения</th><th>Сумма</th><th>Скидка</th><th>К оплате</th><th>Способ</th><th>Заметка</th></tr></thead><tbody>${rows||"<tr><td colspan='10' style='text-align:center;color:#999;padding:12px'>Нет записей</td></tr>"}</tbody></table><div class="totals"><div class="row"><span>💵 Наличные</span><span>${fmtAmd(t.cash)}</span></div><div class="row"><span>🏦 ACBA</span><span>${fmtAmd(t.acba)}</span></div><div class="row"><span>🏦 INECO</span><span>${fmtAmd(t.ineco)}</span></div><div class="row grand"><span>Итого</span><span>${fmtAmd(t.total||t.cash+t.acba+t.ineco)}</span></div></div><p class="footer">Распечатано: ${new Date().toLocaleString("ru-RU")}</p><script>window.onload=function(){window.print();}<\/script></body></html>`;
          const w=window.open("","_blank","width=900,height=650");
          if(w){w.document.write(html);w.document.close();}
        };

        return(
          <div>
            {/* Month nav */}
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16,background:"#f1f5f9",borderRadius:8,padding:"10px 14px",flexWrap:"wrap"}}>
              <button onClick={()=>setSelMonth(prevMo(selMonth))} disabled={selMonth<=MIN_MONTH} style={{...btn("#fff","#374151",{border:"1px solid #d1d5db",fontSize:18,padding:"3px 10px"}),opacity:selMonth<=MIN_MONTH?0.4:1}}>‹</button>
              <span style={{fontWeight:700,fontSize:16,minWidth:160,textAlign:"center"}}>{fmtMonth(selMonth)}</span>
              <button onClick={()=>setSelMonth(nextMo(selMonth))} disabled={selMonth>=MAX_MONTH} style={{...btn("#fff","#374151",{border:"1px solid #d1d5db",fontSize:18,padding:"3px 10px"}),opacity:selMonth>=MAX_MONTH?0.4:1}}>›</button>
            </div>
            {/* Month totals */}
            <div style={{display:"flex",gap:12,marginBottom:20,flexWrap:"wrap"}}>
              {card("💵 Наличные",mTotals.cash,"#374151")}
              {card("🏦 ACBA",mTotals.acba,"#1d4ed8")}
              {card("🏦 INECO",mTotals.ineco,"#0f766e")}
              {card("📊 Итого за месяц",mTotal,"#7c3aed")}
            </div>
            {!cashLoaded&&<div style={{padding:40,textAlign:"center",color:"#9ca3af"}}>Загрузка...</div>}
            {cashLoaded&&days.length===0&&(
              <div style={{padding:48,textAlign:"center",color:"#9ca3af",fontSize:14,border:"2px dashed #e5e7eb",borderRadius:8}}>
                <div style={{fontSize:32,marginBottom:8}}>📒</div>
                <div>Нет платежей за {fmtMonth(selMonth)}</div>
              </div>
            )}
            {/* Days table */}
            {cashLoaded&&days.length>0&&(
              <div style={{border:"1px solid #e5e7eb",borderRadius:8,overflow:"hidden"}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead>
                    <tr style={{background:"#f1f5f9"}}>
                      {["Дата","💵 Наличные","🏦 ACBA","🏦 INECO","Итого","Записей","Статус",""].map(h=><th key={h} style={{...th,padding:"10px 12px"}}>{h}</th>)}
                    </tr>
                  </thead>
                  <tbody>
                    {days.map(date=>{
                      const t=getTotals(date);
                      const pols=getPols(date);
                      const total=t.total||t.cash+t.acba+t.ineco;
                      const isClosed=!!(cashDays[date]&&cashDays[date].closed);
                      const wasReopened=!!(cashDays[date]&&cashDays[date].reopenedAt);
                      const isToday=date===today;
                      const isPast=date<today;
                      const isExpanded=cashExpandDay===date;
                      const canClose=!isClosed&&(isToday||(isAdmin&&isPast));
                      return(<Fragment key={date}>
                        <tr onClick={()=>setCashExpandDay(isExpanded?null:date)}
                          style={{background:isToday?"#fffbeb":isClosed?"#f0fdf4":"white",borderBottom:"1px solid #e5e7eb",cursor:"pointer"}}>
                          <td style={{...td,fontWeight:700,whiteSpace:"nowrap"}}>
                            <span style={{marginRight:6,fontSize:11,color:"#9ca3af"}}>{isExpanded?"▾":"▸"}</span>
                            {fmtDay(date)}
                            {isToday&&<span style={{marginLeft:6,background:"#fef9c3",color:"#92400e",borderRadius:10,padding:"1px 7px",fontSize:11,fontWeight:600}}>сегодня</span>}
                          </td>
                          <td style={{...td,textAlign:"right",color:t.cash>0?"#374151":"#d1d5db"}}>{t.cash>0?fmt(t.cash):"—"}</td>
                          <td style={{...td,textAlign:"right",color:t.acba>0?"#1d4ed8":"#d1d5db"}}>{t.acba>0?fmt(t.acba):"—"}</td>
                          <td style={{...td,textAlign:"right",color:t.ineco>0?"#0f766e":"#d1d5db"}}>{t.ineco>0?fmt(t.ineco):"—"}</td>
                          <td style={{...td,textAlign:"right",fontWeight:700}}>{total>0?fmt(total):"—"}</td>
                          <td style={{...td,textAlign:"center",color:"#6b7280"}}>{pols.length}</td>
                          <td style={{...td,whiteSpace:"nowrap"}}>
                            {isClosed
                              ?<><span style={{background:"#dcfce7",color:"#166534",borderRadius:12,padding:"2px 9px",fontSize:11,fontWeight:600}}>✓ Закрыта</span>
                                <span style={{marginLeft:5,fontSize:10,color:"#9ca3af"}}>{fmtTime(cashDays[date].closedAt)}</span>
                                {wasReopened&&<span style={{marginLeft:5,fontSize:10,color:"#d97706"}}>переоткрывалась</span>}</>
                              :<span style={{background:"#fef9c3",color:"#92400e",borderRadius:12,padding:"2px 9px",fontSize:11,fontWeight:600}}>⏳ Открыта</span>
                            }
                          </td>
                          <td style={{...td,whiteSpace:"nowrap"}} onClick={e=>e.stopPropagation()}>
                            {canClose&&<button onClick={()=>setCashCloseModal(date)} style={{...btn("#16a34a",undefined,{fontSize:11,padding:"3px 10px"}),marginRight:4}}>Закрыть кассу</button>}
                            {isClosed&&<button onClick={()=>printCashReport(date)} style={{...btn("#f3f4f6","#374151",{fontSize:11,padding:"3px 10px"}),marginRight:4}}>🖨 Печать</button>}
                            {isClosed&&isAdmin&&<button onClick={()=>setCashReopenModal(date)} style={btn("#fff7ed","#b45309",{fontSize:11,padding:"3px 10px",border:"1px solid #fcd34d"})}>🔓 Открыть</button>}
                          </td>
                        </tr>
                        {isExpanded&&(
                          <tr key={date+"-exp"}>
                            <td colSpan={8} style={{padding:"12px 16px",background:"#fafafa",borderBottom:"1px solid #e5e7eb"}}>
                              {isClosed&&<div style={{fontSize:11,color:"#6b7280",marginBottom:8}}>Снимок на момент закрытия · {fmtDatetime(cashDays[date].closedAt)}</div>}
                              {pols.length>0
                                ?<div style={{overflowX:"auto"}}>{polsTable(pols,true)}</div>
                                :<div style={{color:"#9ca3af",fontSize:12,padding:"8px 0"}}>Нет платежей за этот день</div>
                              }
                            </td>
                          </tr>
                        )}
                      </Fragment>);
                    })}
                  </tbody>
                </table>
              </div>
            )}

            {/* === Close-of-day modal === */}
            {cashCloseModal&&(()=>{
              const d=byDay[cashCloseModal]||{cash:0,acba:0,ineco:0,pols:[]};
              const total=d.cash+d.acba+d.ineco;
              return(
              <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.6)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
                <div style={{background:"white",borderRadius:12,width:"100%",maxWidth:680,maxHeight:"90vh",display:"flex",flexDirection:"column",boxShadow:"0 20px 60px rgba(0,0,0,0.35)"}}>
                  {/* Header */}
                  <div style={{padding:"18px 24px 14px",borderBottom:"1px solid #e5e7eb"}}>
                    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between"}}>
                      <h3 style={{margin:0,fontSize:16}}>📒 Закрыть кассу — {fmtDay(cashCloseModal)}</h3>
                      <button onClick={()=>setCashCloseModal(null)} style={{background:"none",border:"none",fontSize:20,cursor:"pointer",color:"#9ca3af"}}>×</button>
                    </div>
                    <p style={{margin:"4px 0 0",fontSize:12,color:"#6b7280"}}>Проверьте записи перед закрытием. После подтверждения изменения будут недоступны.</p>
                  </div>
                  {/* Policy list */}
                  <div style={{flex:1,overflowY:"auto",padding:"12px 24px"}}>
                    {d.pols.length===0
                      ?<div style={{padding:24,textAlign:"center",color:"#9ca3af",fontSize:13}}>Нет платежей за этот день</div>
                      :<div style={{overflowX:"auto"}}>{polsTable(d.pols,false)}</div>
                    }
                  </div>
                  {/* Totals */}
                  <div style={{padding:"12px 24px",background:"#f8fafc",borderTop:"1px solid #e5e7eb",borderBottom:"1px solid #e5e7eb"}}>
                    <div style={{display:"flex",gap:24,flexWrap:"wrap",fontSize:13}}>
                      <div><span style={{color:"#6b7280"}}>💵 Наличные: </span><strong>{fmt(d.cash)}</strong></div>
                      <div><span style={{color:"#6b7280"}}>🏦 ACBA: </span><strong style={{color:"#1d4ed8"}}>{fmt(d.acba)}</strong></div>
                      <div><span style={{color:"#6b7280"}}>🏦 INECO: </span><strong style={{color:"#0f766e"}}>{fmt(d.ineco)}</strong></div>
                      <div style={{marginLeft:"auto"}}><span style={{color:"#6b7280"}}>Итого: </span><strong style={{fontSize:16,color:"#7c3aed"}}>{fmt(total)}</strong></div>
                    </div>
                  </div>
                  {/* Warning + buttons */}
                  <div style={{padding:"12px 24px 18px"}}>
                    <div style={{background:"#fff7ed",border:"1px solid #fed7aa",borderRadius:6,padding:"7px 12px",marginBottom:12,fontSize:12,color:"#9a3412"}}>
                      ⚠ После закрытия сотрудник не сможет вносить изменения. Администратор может открыть кассу при необходимости.
                    </div>
                    <div style={{display:"flex",gap:8}}>
                      <button onClick={()=>closeCashDay(cashCloseModal)} style={{...btn("#16a34a"),flex:1,padding:"10px",fontSize:14}}>✓ Закрыть кассу</button>
                      <button onClick={()=>setCashCloseModal(null)} style={{...btn("#f3f4f6","#374151"),flex:1,padding:"10px",fontSize:14}}>Отмена</button>
                    </div>
                  </div>
                </div>
              </div>
              );
            })()}

            {/* === Reopen modal (admin only) === */}
            {cashReopenModal&&(
              <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
                <div style={{background:"white",borderRadius:12,padding:28,width:"100%",maxWidth:400,boxShadow:"0 20px 60px rgba(0,0,0,0.3)"}}>
                  <div style={{textAlign:"center",fontSize:36,marginBottom:8}}>🔓</div>
                  <h3 style={{margin:"0 0 10px",fontSize:16,textAlign:"center"}}>Открыть кассу</h3>
                  <p style={{textAlign:"center",fontSize:14,color:"#374151",margin:"0 0 16px"}}>за <strong>{fmtDay(cashReopenModal)}</strong></p>
                  <div style={{background:"#fffbeb",border:"1px solid #fcd34d",borderRadius:6,padding:"8px 12px",marginBottom:16,fontSize:12,color:"#92400e"}}>
                    После открытия сотрудник сможет принять новые платежи и повторно закрыть кассу. Снимок предыдущего закрытия будет очищен.
                  </div>
                  <div style={{display:"flex",gap:8}}>
                    <button onClick={()=>reopenCashDay(cashReopenModal)} style={{...btn("#d97706"),flex:1,padding:"10px",fontSize:14}}>🔓 Открыть</button>
                    <button onClick={()=>setCashReopenModal(null)} style={{...btn("#f3f4f6","#374151"),flex:1,padding:"10px",fontSize:14}}>Отмена</button>
                  </div>
                </div>
              </div>
            )}
          </div>
        );
      })()}

      {tab==="payroll"&&(()=>{
        const foundUID=payrollUID&&agentData.find(a=>a.uid===payrollUID)?payrollUID:null;
        const foundAgent=foundUID?agentData.find(a=>a.uid===foundUID):null;
        const validPols=foundAgent?foundAgent.policies.filter(p=>!p.exception):[];
        const excPols=foundAgent?foundAgent.policies.filter(p=>p.exception):[];
        return(
        <div>
          {/* Навигация по месяцу */}
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16,background:"#f1f5f9",borderRadius:8,padding:"10px 14px",flexWrap:"wrap"}}>
            <button onClick={()=>setSelMonth(prevMo(selMonth))} disabled={selMonth<=MIN_MONTH} style={{...btn("#fff","#374151",{border:"1px solid #d1d5db",fontSize:18,padding:"3px 10px"}),opacity:selMonth<=MIN_MONTH?0.4:1}}>‹</button>
            <span style={{fontWeight:700,fontSize:16,minWidth:160,textAlign:"center"}}>{fmtMonth(selMonth)}</span>
            <button onClick={()=>setSelMonth(nextMo(selMonth))} disabled={selMonth>=MAX_MONTH} style={{...btn("#fff","#374151",{border:"1px solid #d1d5db",fontSize:18,padding:"3px 10px"}),opacity:selMonth>=MAX_MONTH?0.4:1}}>›</button>
          </div>

          {/* Коды офиса */}
          <div style={{border:"1px solid #e5e7eb",borderRadius:8,padding:"12px 14px",marginBottom:20,background:"#fafafa"}}>
            <div style={{fontWeight:600,fontSize:13,marginBottom:10,color:"#374151"}}>Коды офиса — {fmtMonth(selMonth)}</div>
            {officeCodes.length>0&&(
              <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:10}}>
                {officeCodes.map((code,i)=>(
                  <span key={i} style={{display:"inline-flex",alignItems:"center",gap:5,background:"#eff6ff",border:"1px solid #bfdbfe",borderRadius:6,padding:"4px 10px",fontSize:13,color:"#1e40af",fontWeight:600}}>
                    {code}
                    <button onClick={()=>removeOfficeCode(i)} style={{background:"none",border:"none",cursor:"pointer",color:"#6b7280",fontSize:14,padding:0,lineHeight:1}}>×</button>
                  </span>
                ))}
              </div>
            )}
            <div style={{display:"flex",gap:8,alignItems:"center"}}>
              <input
                value={newOfficeCode}
                onChange={e=>setNewOfficeCode(e.target.value)}
                onKeyDown={e=>e.key==="Enter"&&addOfficeCode()}
                placeholder="Введите код офиса"
                style={{padding:"6px 10px",border:"1px solid #d1d5db",borderRadius:6,fontSize:13,minWidth:180,outline:"none"}}
              />
              <button onClick={addOfficeCode} style={btn("#2563eb",undefined,{fontSize:12})}>+ Добавить</button>
            </div>
          </div>

          {/* Фильтр по внутреннему коду агента */}
          <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:20,flexWrap:"wrap"}}>
            <input
              value={payrollIC}
              onChange={e=>setPayrollIC(e.target.value)}
              onKeyDown={e=>{if(e.key==="Enter"){const uid=Object.keys(agentDir).find(k=>agentDir[k].internalCode===payrollIC.trim());setPayrollUID(uid||"__notfound__");} }}
              placeholder="Внутренний код агента (напр. 768-40)"
              style={{padding:"7px 12px",border:"1px solid #d1d5db",borderRadius:6,fontSize:14,minWidth:260,outline:"none"}}
            />
            <button onClick={()=>{const uid=Object.keys(agentDir).find(k=>agentDir[k].internalCode===payrollIC.trim());setPayrollUID(uid||"__notfound__");}} style={btn("#2563eb")}>Показать</button>
            {payrollUID&&<button onClick={()=>{setPayrollUID(null);setPayrollIC("");}} style={btn("#6b7280",undefined,{fontSize:12})}>✕ Сбросить</button>}
          </div>

          {/* Состояния */}
          {!payrollUID&&<div style={{padding:40,textAlign:"center",color:"#9ca3af",fontSize:14}}>Введите внутренний код агента и нажмите «Показать»</div>}
          {payrollUID==="__notfound__"&&<div style={{padding:30,textAlign:"center",color:"#dc2626",fontSize:14,background:"#fef2f2",borderRadius:8}}>Агент с кодом «{payrollIC}» не найден в справочнике</div>}
          {foundAgent&&foundAgent.policies.length===0&&<div style={{padding:30,textAlign:"center",color:"#9ca3af",fontSize:14}}>Нет полисов за {fmtMonth(selMonth)}</div>}

          {foundAgent&&foundAgent.policies.length>0&&(
            <div>
              {/* Шапка агента */}
              <div style={{background:"#1e293b",color:"#fff",borderRadius:10,padding:"14px 20px",marginBottom:20,display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
                <div>
                  <div style={{fontSize:11,opacity:.6,marginBottom:2}}>Агент</div>
                  <div style={{fontWeight:700,fontSize:16}}>{agName(foundAgent)}</div>
                  <div style={{fontSize:12,opacity:.7,marginTop:2}}>{(agentDir[foundAgent.uid]&&agentDir[foundAgent.uid].internalCode)||""}</div>
                </div>
                <div style={{display:"flex",gap:16,flexWrap:"wrap"}}>
                  <div style={{textAlign:"center"}}>
                    <div style={{fontSize:11,opacity:.6}}>Зачётных</div>
                    <div style={{fontWeight:700,fontSize:18,color:"#4ade80"}}>{validPols.length}</div>
                  </div>
                  <div style={{textAlign:"center"}}>
                    <div style={{fontSize:11,opacity:.6}}>Незачётных</div>
                    <div style={{fontWeight:700,fontSize:18,color:"#f87171"}}>{excPols.length}</div>
                  </div>
                  <div style={{textAlign:"center"}}>
                    <div style={{fontSize:11,opacity:.6}}>Итого начислено</div>
                    <div style={{fontWeight:700,fontSize:18,color:"#fbbf24"}}>{fmt(foundAgent.totalAgent)}</div>
                  </div>
                </div>
                <button onClick={()=>exportAgentAll(foundAgent.uid,validPols,excPols,selMonth)} style={btn("#16a34a")}>⬇ Экспорт (все)</button>
              </div>

              {/* Таблица по компаниям */}
              <div style={{overflowX:"auto",borderRadius:8,border:"1px solid #e5e7eb",marginBottom:24}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead>
                    <tr style={{background:"#f8fafc"}}>
                      <th style={th}>Компания</th>
                      <th style={{...th,textAlign:"center"}}>Кол-во полисов</th>
                      <th style={{...th,textAlign:"right"}}>Сумма начислений</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ALL_COMPANIES.map(c=>{
                      const pols=foundAgent.policies.filter(p=>p.company===c&&!p.exception);
                      const sum=pols.reduce((s,p)=>s+p.agentComm,0);
                      if(pols.length===0)return null;
                      return(
                        <tr key={c} style={{borderBottom:"1px solid #f0f0f0"}}>
                          <td style={{...td,fontWeight:600}}>{c}</td>
                          <td style={{...td,textAlign:"center"}}>{pols.length}</td>
                          <td style={{...td,textAlign:"right",color:"#16a34a",fontWeight:600}}>{fmt(sum)}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot>
                    <tr style={{background:"#f1f5f9"}}>
                      <td style={{...td,fontWeight:700}}>ИТОГО</td>
                      <td style={{...td,textAlign:"center",fontWeight:700}}>{validPols.length}</td>
                      <td style={{...td,textAlign:"right",fontWeight:700,color:"#16a34a"}}>{fmt(foundAgent.totalAgent)}</td>
                    </tr>
                  </tfoot>
                </table>
              </div>

              {/* Незачётные полисы */}
              {excPols.length>0&&(
                <div style={{marginBottom:24}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8,flexWrap:"wrap",gap:8}}>
                    <div style={{fontWeight:700,fontSize:14,color:"#dc2626"}}>Незачётные полисы ({excPols.length})</div>
                    <button onClick={()=>exportAgentExc(foundAgent.uid,excPols,selMonth)} style={btn("#dc2626",undefined,{fontSize:12})}>⬇ Экспорт незачётных</button>
                  </div>
                  <div style={{overflowX:"auto",borderRadius:8,border:"1px solid #fca5a5"}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                      <thead><tr style={{background:"#fef2f2"}}>{["№ полиса","Компания","Страхователь","Марка","Рег.номер","Начало","Окончание","Регион","БМ","Сумма"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
                      <tbody>{excPols.map((p,i)=>(
                        <tr key={i} style={{background:i%2===0?"white":"#fff5f5",borderBottom:"1px solid #fee2e2"}}>
                          <td style={td}>{p.policyNum}</td>
                          <td style={td}>{p.company}</td>
                          <td style={td}>{p.insuredName}</td>
                          <td style={{...td,color:"#6b7280"}}>{p.car}</td>
                          <td style={{...td,fontSize:11}}>{p.carPlate}</td>
                          <td style={{...td,fontSize:11,color:"#6b7280"}}>{p.startDateFmt}</td>
                          <td style={{...td,fontSize:11}}>{p.endDateFmt}</td>
                          <td style={td}>{p.region}</td>
                          <td style={{...td,textAlign:"center"}}>{p.bm}</td>
                          <td style={{...td,textAlign:"right"}}>{fmt(p.amount)}</td>
                        </tr>
                      ))}</tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* Зачётные полисы */}
              {validPols.length>0&&(
                <div style={{marginBottom:24}}>
                  <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8,flexWrap:"wrap",gap:8}}>
                    <div style={{fontWeight:700,fontSize:14,color:"#16a34a"}}>Зачётные полисы ({validPols.length})</div>
                    <button onClick={()=>exportAgentValid(foundAgent.uid,validPols,selMonth)} style={btn("#16a34a",undefined,{fontSize:12})}>⬇ Экспорт зачётных</button>
                  </div>
                  <div style={{overflowX:"auto",borderRadius:8,border:"1px solid #86efac"}}>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                      <thead><tr style={{background:"#f0fdf4"}}>{["№ полиса","Компания","Страхователь","Марка","Рег.номер","Начало","Окончание","Регион","БМ","Сумма","Ставка %","Начислено"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
                      <tbody>{validPols.map((p,i)=>(
                        <tr key={i} style={{background:i%2===0?"white":"#f7fef9",borderBottom:"1px solid #d1fae5"}}>
                          <td style={td}>{p.policyNum}</td>
                          <td style={td}>{p.company}</td>
                          <td style={td}>{p.insuredName}</td>
                          <td style={{...td,color:"#6b7280"}}>{p.car}</td>
                          <td style={{...td,fontSize:11}}>{p.carPlate}</td>
                          <td style={{...td,fontSize:11,color:"#6b7280"}}>{p.startDateFmt}</td>
                          <td style={{...td,fontSize:11}}>{p.endDateFmt}</td>
                          <td style={td}>{p.region}</td>
                          <td style={{...td,textAlign:"center"}}>{p.bm}</td>
                          <td style={{...td,textAlign:"right"}}>{fmt(p.amount)}</td>
                          <td style={{...td,textAlign:"center",color:"#6366f1"}}>{(p.agentRate||0)+"%"}</td>
                          <td style={{...td,textAlign:"right",fontWeight:600,color:"#16a34a"}}>{fmt(p.agentComm)}</td>
                        </tr>
                      ))}</tbody>
                      <tfoot>
                        <tr style={{background:"#f0fdf4"}}>
                          <td colSpan={9} style={{...td,fontWeight:700,borderTop:"2px solid #86efac"}}>ИТОГО</td>
                          <td style={{...td,textAlign:"right",fontWeight:700,borderTop:"2px solid #86efac"}}>{fmt(validPols.reduce((s,p)=>s+p.amount,0))}</td>
                          <td style={{...td,borderTop:"2px solid #86efac"}}></td>
                          <td style={{...td,textAlign:"right",fontWeight:700,color:"#16a34a",borderTop:"2px solid #86efac"}}>{fmt(foundAgent.totalAgent)}</td>
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                </div>
              )}
            </div>
          )}
          {/* Кнопка общих начислений за месяц */}
          <div style={{marginTop:32,paddingTop:20,borderTop:"2px solid #e5e7eb"}}>
            <button
              onClick={()=>setShowAllPayroll(v=>!v)}
              style={{...btn(showAllPayroll?"#374151":"#1e293b",undefined,{fontSize:14,padding:"9px 20px"})}}>
              {showAllPayroll?"Скрыть":"📋 Все начисления за "+fmtMonth(selMonth)}
            </button>
          </div>

          {showAllPayroll&&(()=>{
            const rows=buildAllPayrollRows(agentData,effVol,agentDir,officeCodes);
            const totPols=rows.reduce((s,r)=>s+r.polCount,0);
            const totAcc=rows.reduce((s,r)=>s+r.accrued,0);
            const totVol=rows.reduce((s,r)=>s+r.volDebt,0);
            const totNet=rows.reduce((s,r)=>s+r.net,0);
            const fillGreen={background:"#dcfce7"};
            const fillRed={background:"#fee2e2"};
            const fillBlue={background:"#dbeafe"};
            const P="padding:8px 12px;border:1.5px solid #444;";

            const doPrint=()=>{
              const bodyRows=rows.map((r,i)=>{
                const bg=i%2===0?"#fff":"#f5f7fa";
                const netColor=r.net>=0?"#1d4ed8":"#dc2626";
                return`<tr style="background:${bg}">
                  <td style="${P}font-weight:600;">${r.name}</td>
                  <td style="${P}text-align:center;color:#4f46e5;font-weight:600;">${r.ic||"—"}</td>
                  <td style="${P}text-align:center;">${r.polCount}</td>
                  <td style="${P}text-align:right;background:#dcfce7;color:#16a34a;font-weight:600;-webkit-print-color-adjust:exact;print-color-adjust:exact;">${r.accrued.toLocaleString("ru-RU")} ֏</td>
                  <td style="${P}text-align:right;background:#fee2e2;color:${r.volDebt>0?"#dc2626":"#6b7280"};-webkit-print-color-adjust:exact;print-color-adjust:exact;">${r.volDebt>0?r.volDebt.toLocaleString("ru-RU")+" ֏":"0 ֏"}</td>
                  <td style="${P}text-align:right;background:#dbeafe;color:${netColor};font-weight:700;-webkit-print-color-adjust:exact;print-color-adjust:exact;">${r.net.toLocaleString("ru-RU")} ֏</td>
                  <td style="${P}min-width:80px;"> </td>
                  <td style="${P}min-width:130px;"> </td>
                </tr>`;
              }).join("");
              const html=`<!DOCTYPE html><html><head><meta charset="utf-8">
                <style>
                  @page { size: A4 landscape; margin: 8mm; }
                  body { font-family: Arial, sans-serif; margin: 0; }
                  h3 { text-align: center; font-size: 11pt; margin: 0 0 10px; }
                  table { width: 100%; border-collapse: collapse; }
                  th { padding: 8px 12px; border: 1.5px solid #333; background: #1e293b; color: #fff; white-space: nowrap; font-size: 8.5pt; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
                  td { font-size: 8.5pt; }
                </style>
              </head><body>
                <h3>Начисления агентам — ${fmtMonth(selMonth)}</h3>
                <table>
                  <thead><tr>
                    <th style="text-align:left;min-width:140px;">Имя агента</th>
                    <th style="min-width:80px;">768-код</th>
                    <th style="min-width:55px;">Кол-во полисов</th>
                    <th style="min-width:110px;background:#15803d!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;">Начислено (ОСАГО)</th>
                    <th style="min-width:120px;background:#b91c1c!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;">К оплате офису (доброволь.)</th>
                    <th style="min-width:110px;background:#1d4ed8!important;-webkit-print-color-adjust:exact;print-color-adjust:exact;">К выплате</th>
                    <th style="min-width:80px;">Примечание</th>
                    <th style="min-width:130px;">Подпись</th>
                  </tr></thead>
                  <tbody>${bodyRows}
                  <tr style="background:#e2e8f0;font-weight:bold;border-top:2.5px solid #374151;-webkit-print-color-adjust:exact;print-color-adjust:exact;">
                    <td style="${P}font-weight:700;border-top:2.5px solid #374151;">ИТОГО</td>
                    <td style="${P}border-top:2.5px solid #374151;"></td>
                    <td style="${P}text-align:center;font-weight:700;border-top:2.5px solid #374151;">${totPols}</td>
                    <td style="${P}text-align:right;background:#dcfce7;color:#15803d;font-weight:700;border-top:2.5px solid #374151;-webkit-print-color-adjust:exact;print-color-adjust:exact;">${totAcc.toLocaleString("ru-RU")} ֏</td>
                    <td style="${P}text-align:right;background:#fee2e2;color:#b91c1c;font-weight:700;border-top:2.5px solid #374151;-webkit-print-color-adjust:exact;print-color-adjust:exact;">${totVol>0?totVol.toLocaleString("ru-RU")+" ֏":"0 ֏"}</td>
                    <td style="${P}text-align:right;background:#dbeafe;color:#1d4ed8;font-weight:700;border-top:2.5px solid #374151;-webkit-print-color-adjust:exact;print-color-adjust:exact;">${totNet.toLocaleString("ru-RU")} ֏</td>
                    <td style="${P}border-top:2.5px solid #374151;"></td>
                    <td style="${P}border-top:2.5px solid #374151;"></td>
                  </tr></tbody>
                </table>
              </body></html>`;
              const w=window.open("","_blank","width=1000,height=700");
              if(!w)return;
              w.document.write(html);
              w.document.close();
              w.focus();
              setTimeout(()=>{w.print();},400);
            };

            return(
              <div>
                <div style={{display:"flex",alignItems:"center",gap:10,margin:"16px 0",flexWrap:"wrap"}}>
                  <span style={{fontWeight:700,fontSize:15}}>Все начисления — {fmtMonth(selMonth)}</span>
                  <button onClick={()=>exportAllPayrollXlsx(agentData,effVol,agentDir,selMonth)} style={btn("#16a34a",undefined,{fontSize:12})}>⬇ Excel</button>
                  <button onClick={doPrint} style={btn("#6366f1",undefined,{fontSize:12})}>🖨 Печать / PDF</button>
                </div>

                <div style={{overflowX:"auto",borderRadius:8,border:"1px solid #e5e7eb"}}>
                  <div style={{fontWeight:700,fontSize:13,padding:"10px 14px",borderBottom:"1px solid #e5e7eb",background:"#f8fafc",textAlign:"center"}}>
                    Начисления агентам — {fmtMonth(selMonth)}
                  </div>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead>
                      <tr style={{background:"#1e293b",color:"#fff"}}>
                        <th style={{...th,color:"#fff",background:"#1e293b",textAlign:"left",minWidth:160}}>Имя агента</th>
                        <th style={{...th,color:"#fff",background:"#1e293b",textAlign:"center",minWidth:90}}>768-код</th>
                        <th style={{...th,color:"#fff",background:"#1e293b",textAlign:"center",minWidth:70}}>Кол-во полисов</th>
                        <th style={{...th,color:"#fff",background:"#1e293b",textAlign:"right",...fillGreen,minWidth:110}}>Начислено (ОСАГО)</th>
                        <th style={{...th,color:"#fff",background:"#1e293b",textAlign:"right",...fillRed,minWidth:130}}>К оплате офису (доброволь.)</th>
                        <th style={{...th,color:"#fff",background:"#1e293b",textAlign:"right",...fillBlue,minWidth:110}}>К выплате</th>
                        <th style={{...th,color:"#fff",background:"#1e293b",minWidth:100}}>Примечание</th>
                        <th style={{...th,color:"#fff",background:"#1e293b",minWidth:130}}>Подпись</th>
                      </tr>
                    </thead>
                    <tbody>
                      {rows.map((r,i)=>(
                        <tr key={r.uid} style={{borderBottom:"1px solid #e5e7eb",background:i%2===0?"white":"#f8fafc"}}>
                          <td style={{...td,fontWeight:600}}>{r.name}</td>
                          <td style={{...td,textAlign:"center",color:"#6366f1",fontWeight:600}}>{r.ic||"—"}</td>
                          <td style={{...td,textAlign:"center"}}>{r.polCount}</td>
                          <td style={{...td,textAlign:"right",...fillGreen,fontWeight:600,color:"#16a34a"}}>{fmt(r.accrued)}</td>
                          <td style={{...td,textAlign:"right",...fillRed,color:r.volDebt>0?"#dc2626":"#6b7280"}}>{r.volDebt>0?fmt(r.volDebt):"0 ֏"}</td>
                          <td style={{...td,textAlign:"right",...fillBlue,fontWeight:700,color:r.net>=0?"#1d4ed8":"#dc2626"}}>{fmt(r.net)}</td>
                          <td style={td}></td>
                          <td style={{...td,borderBottom:"1px solid #9ca3af"}}></td>
                        </tr>
                      ))}
                    </tbody>
                    <tfoot>
                      <tr style={{background:"#f1f5f9",borderTop:"2px solid #374151"}}>
                        <td style={{...td,fontWeight:700}}>ИТОГО</td>
                        <td style={td}></td>
                        <td style={{...td,textAlign:"center",fontWeight:700}}>{totPols}</td>
                        <td style={{...td,textAlign:"right",...fillGreen,fontWeight:700,color:"#16a34a"}}>{fmt(totAcc)}</td>
                        <td style={{...td,textAlign:"right",...fillRed,fontWeight:700,color:"#dc2626"}}>{totVol>0?fmt(totVol):"0 ֏"}</td>
                        <td style={{...td,textAlign:"right",...fillBlue,fontWeight:700,color:"#1d4ed8"}}>{fmt(totNet)}</td>
                        <td style={td}></td>
                        <td style={td}></td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              </div>
            );
          })()}
        </div>
        );
      })()}

      {pinModal&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center"}}
          onClick={e=>{if(e.target===e.currentTarget){setPinModal(false);setPinInput("");setPinError("");}}}>
          <div style={{background:"white",borderRadius:14,padding:"32px 28px",width:300,boxShadow:"0 24px 64px rgba(0,0,0,0.3)"}}>
            <div style={{textAlign:"center",fontSize:32,marginBottom:8}}>🔐</div>
            <h3 style={{margin:"0 0 6px",fontSize:17,textAlign:"center"}}>Вход для администратора</h3>
            {!adminPin&&<p style={{fontSize:12,color:"#6b7280",textAlign:"center",margin:"0 0 16px"}}>Первый вход — введите новый PIN</p>}
            {adminPin&&<div style={{height:12}}/>}
            <input
              type="password"
              value={pinInput}
              onChange={e=>{setPinInput(e.target.value);setPinError("");}}
              onKeyDown={e=>e.key==="Enter"&&tryAdminLogin()}
              placeholder="Введите PIN"
              autoFocus
              style={{...inp,width:"100%",padding:"11px 12px",fontSize:20,textAlign:"center",letterSpacing:6,boxSizing:"border-box",marginBottom:6}}
            />
            {pinError&&<div style={{color:"#dc2626",fontSize:12,textAlign:"center",marginBottom:8}}>{pinError}</div>}
            <div style={{display:"flex",gap:8,marginTop:12}}>
              <button onClick={tryAdminLogin} style={{...btn(),flex:1,padding:"10px",fontSize:14}}>Войти</button>
              <button onClick={()=>{setPinModal(false);setPinInput("");setPinError("");}} style={{...btn("#f3f4f6","#374151"),flex:1,padding:"10px",fontSize:14}}>Отмена</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}