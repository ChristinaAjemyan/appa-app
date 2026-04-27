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
const DEFAULT_MGR_RATES={
  managerRates:{Nairi:16,Ingo:7,Liga:5,Sil:10,Rego:5},
  armeniaManager:{"1-9":11,"10-14":10,"15-25":7},
  operatorRates:{Nairi:[16,14,12,10,8,6],Ingo:[7,6,5,4,3,2],Liga:[5,4,3,3,2,2],Sil:[10,9,8,7,6,5],Rego:[5,4,3,3,2,2]},
  armeniaOperator:{"1-9":[11,10,9,8,7,6],"10-14":[10,9,8,7,6,5],"15-25":[7,6,5,4,3,2]},
  tierThresholds:[0,500000,1000000,1500000,2000000,2500000],
  tierFixes:[0,20000,30000,40000,50000,70000],
  operatorUids:[],
};
const DEFAULT_EXPENSE_ROWS=[
  {id:"sal1",cat:"Зарплаты",name:"Зарплата 1",amount:0,type:"static"},
  {id:"sal2",cat:"Зарплаты",name:"Зарплата 2",amount:0,type:"static"},
  {id:"sal3",cat:"Зарплаты",name:"Зарплата 3",amount:0,type:"static"},
  {id:"sal4",cat:"Зарплаты",name:"Зарплата 4",amount:0,type:"static"},
  {id:"sal5",cat:"Зарплаты",name:"Зарплата 5",amount:0,type:"static"},
  {id:"sal6",cat:"Зарплаты",name:"Зарплата 6",amount:0,type:"static"},
  {id:"sal7",cat:"Зарплаты",name:"Зарплата 7",amount:0,type:"static"},
  {id:"com1",cat:"Коммунальные",name:"Свет",amount:0,type:"static"},
  {id:"com2",cat:"Коммунальные",name:"Газ",amount:0,type:"static"},
  {id:"com3",cat:"Коммунальные",name:"Вода",amount:0,type:"static"},
  {id:"tel1",cat:"Связь",name:"Телефония офис",amount:0,type:"static"},
  {id:"tel2",cat:"Связь",name:"Телефония КолЦентр",amount:0,type:"static"},
  {id:"tax1",cat:"Налоги",name:"Подоходный налог",amount:0,type:"static"},
  {id:"tax2",cat:"Налоги",name:"Подоходный налог от аренды",amount:0,type:"static"},
  {id:"tax3",cat:"Налоги",name:"Налог на прибыль",amount:0,type:"static"},
  {id:"tax4",cat:"Налоги",name:"Соц. выплаты",amount:0,type:"static"},
  {id:"tax5",cat:"Налоги",name:"Налог на акцизы",amount:0,type:"static"},
  {id:"tax6",cat:"Налоги",name:"Штрафы",amount:0,type:"static"},
  {id:"tax7",cat:"Налоги",name:"Налог ИП Гагик",amount:0,type:"static"},
  {id:"tax8",cat:"Налоги",name:"Налог ИП Гриш",amount:0,type:"static"},
  {id:"tax9",cat:"Налоги",name:"Налог ИП Левон",amount:0,type:"static"},
  {id:"oth1",cat:"Прочее",name:"Аренда основного офиса",amount:0,type:"static"},
  {id:"oth2",cat:"Прочее",name:"Аренда офиса МРЭО",amount:0,type:"static"},
  {id:"oth3",cat:"Прочее",name:"Реклама",amount:0,type:"static"},
  {id:"oth4",cat:"Прочее",name:"Заправка принтеров",amount:0,type:"static"},
  {id:"oth5",cat:"Прочее",name:"Бумага и файлы",amount:0,type:"static"},
  {id:"oth6",cat:"Прочее",name:"Хоз. расходы",amount:0,type:"static"},
  {id:"oth7",cat:"Прочее",name:"Бухгалтерия",amount:0,type:"static"},
];
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
    ["768-01","Հակոբ","Վխկրյան","1606-01","","","","",""],
    ["768-03","Ռոբերտ","Բիչախչյան (112)","1606-03","","","","",""],
    ["768-05","Լուսինե","Մոսոյան (12+66)","1606-05","541-01","","","",""],
    ["768-07","Մկրտիչ","Բաղդասարյան","1606-07","","310-71-10","797-25","",""],
    ["768-11","Սրապ","Խաչատրյան","1606-11","","","","",""],
    ["768-19","Գրիգոր","Հովհաննիսյան (Նաթելլա)","1606-168","","","","",""],
    ["768-20","Դավիթ","Մինասյան","1606-20","","","","",""],
    ["768-24","Մանվել","Նահապետյան","1606-24","","","","",""],
    ["768-27","Գագիկ","Վանոյան","1606-27","","","","",""],
    ["768-40","Լեվոն","Վարդանյան","1606-40","3671-07","310-71-11","797-26","A50-M3-40-G12","13021-09"],
    ["768-53","Վլադիմիր","Շահինյան","1606-53","","","","",""],
    ["768-74","Արմեն","Խանոյան","1606-74","","310-71-09","797-23","A50-M3-40-G9","13021-07"],
    ["768-101","Նելլի","Հարությունյան","1606-101","3671-01","310-71-01","797-16","A50-M3-40-G5","13021-01"],
    ["768-105","Վիկտորյա","Վխկրյան","1606-105","3671-05","310-71-02","797-17","A50-M3-40-G3","13021-02"],
    ["768-106","Գայանե","Ալեքսանյան","1606-106","3671-02","310-71-03","797-18","A50-M3-40-G111","13021-04"],
    ["768-115","Աննման","Ղևոնդյան (Շանթ)","1606-115","","","","",""],
    ["768-116","Աղունիկ","Պետրոսյան","1606-116","","","","",""],
    ["768-118","Հովհաննես","Թորոսյան","1606-118","","","","",""],
    ["768-121","Արթուր","Սաֆարյան","1606-121","","","","",""],
    ["768-122","Հովհաննես","Ջանոյան","1606-122","","","","",""],
    ["768-123","Թեհմինե","Քոչարյան","1606-123","","","","",""],
    ["768-125","Նարինե","Արզումանյան","1606-125","101-1","101-1","101-1","A50-M3-40-G4","101-1"],
    ["768-127","Արմեն","Սիմոնյան","1606-127","","","","",""],
    ["768-128","ԳՈՒՐԳԵՆ","ԱՐևԻԿՅԱՆ","1606-128","","310-71-08","797-22","A50-M3-40-G8","13021-06"],
    ["768-130","ՌԻՏԱ","ԳՐԻԳՈՐՅԱՆ","1606-130","","","","",""],
    ["768-131","ՀՌԻՓՍԻՄԵ","ԹՈՐՈՍՅԱՆ","1606-131","","","","",""],
    ["768-132","Ժաննա","Գասպարյան","1606-132","","","797-27","","13021-12"],
    ["768-133","Ռոմիկ","Նազարեթյան (+145+164)","1606-133","","310-71-07","797-21","","13021-10"],
    ["768-135","Արտակ","Խաչատրյան","1606-135","","","","",""],
    ["768-137","ԱՐՏՅՈՄ","ԱՎԵՏԻՍՅԱՆ (Արտակ)","1606-137","","","","",""],
    ["768-138","Զարզանդ","Բրսոյան","1606-138","","","","",""],
    ["768-139","Աղասի","Սահակյան (Արտակ)","1606-139","","","","",""],
    ["768-142","Վարդուհի","Յայլոյան","1606-142","","","","",""],
    ["768-144","Վռամ","Այվազյան","1606-144","","","","",""],
    ["768-145","Ամալյա","Եփրեմյան","1606-145","","","","",""],
    ["768-146","Սարգիս","Համազասպյան","1606-146","","","797-24","A50-M3-40-G10",""],
    ["768-147","Անուշիկ","Սարգսյան","1606-147","","","","",""],
    ["768-148","Լուսինե","Ղազարյան","1606-148","","","","",""],
    ["768-149","Վարդուհի","Թադևոսյան","1606-149","","","","",""],
    ["768-150","ՄԱՐԳԱՐԻՏ","ՄԻԿՈՅԱՆ","1606-150","","","","",""],
    ["768-151","Սվետա","Հասոյան","1606-151","","","","",""],
    ["768-152","Աշոտ","Մարտիրոսյան","1606-152","","","","",""],
    ["768-153","ՎԱՐԴԱՆ","ՂԱԶԱՐՅԱՆ","1606-153","","","","",""],
    ["768-154","ՍՈՒՍԱՆՆԱ","ԳՐԻԳՈՐՅԱՆ","1606-154","","","","",""],
    ["768-155","ԱՐԻՆԱ","ՄԵԾՈՅԱՆ","1606-155","","","","",""],
    ["768-156","ԿԱՐԻՆԵ","ԹԱՄԱԶՅԱՆ (Շաբոյան)","1606-156","","","","",""],
    ["768-157","ԱՐԿԱԴԻ","ԹԱԴևՈՍՅԱՆ","1606-157","","","","",""],
    ["768-158","ԱՐՏԱԿ","ԽԱՉԱՏՐՅԱՆ (ԷԴԳԱՐԻ)","1606-158","","","","",""],
    ["768-159","ՎԱՀԱՆ","ԱՂԱԲԱԲՅԱՆ","1606-159","","310-71-13","","",""],
    ["768-160","Հովհաննես","Հովհաննիսյան","1606-160","","","","",""],
    ["768-161","ՄԱՐՈՒՍՅԱ","ԳՐԻԳՈՐՅԱՆ","1606-161","","","","",""],
    ["768-164","Գրիգոր","Գասպարյան-Ռոման","1606-164","","","","",""],
    ["768-165","Ելենա","","1606-165","","","","",""],
    ["768-166","ԱՐՏՅՈՄ","ՍԱՐԳՍՅԱՆ","1606-166","","","","",""],
    ["768-170","ԱՐՄԵՆՈՒՀԻ","ՄՂԴԵՍՅԱՆ","1606-170","","","","",""],
    ["768-171","ՏԱԹԵՎԻԿ","ԽԱՉԱՏՐՅԱՆ (Նարինե)","1606-171","","","","",""],
    ["768-172","Մարիամ","Ավագյան(Յուլյա)","1606-172","","","","",""],
    ["768-173","Անդրանիկ","Հովհաննիսյան","1606-173","","","","",""],
    ["768-175","Աշխեն","Գալոյան","1606-175","","","","",""],
    ["768-176","Ռուստամ","Կարապետյան (Նելլի Ռեսո)","1606-176","","","","",""],
    ["768-177","Աղավնի","Ստեփանյան(Ելենա)","1606-177","","","","",""],
    ["1606-178","Սմբատ","Ներսիսյան","1606-178","","","","",""],
    ["768-178","Մկրտիչ","Աբաջյան","1606-178-01","","","","",""],
    ["768-179","Հովհաննես","Գևորգյան (Վանաձոր)","1606-179","","","","",""],
    ["768-180","Վահան","Մանուկյան","1606-180","","","","",""],
    ["768-181","Նաիռա","Թովմասյան","1606-181","","310-71-04","797-04","A50-M3-40-G1",""],
    ["768-183","Տիգրան","Մանուկյան","1606-183","","","","",""],
    ["768-184","Գևորգ","Հարությունյան Շիրո","1606-184","","","","",""],
    ["768-185","Կարլեն","Հարությունյան","1606-185","","","","",""],
    ["768-186","Նելլի","Նիկոյան-Մագա","1606-186","3671-03","310-71-05","797-19","A50-M3-40-G2","13021-03"],
    ["768-187","ԱՆԻԱ","ԻԳԻԹՅԱՆ","1606-187","3671-111","310-71-06","797-20","A50-M3-40-G13","13021-111"],
    ["768-188","Հովհաննես","Հովհաննիսյան JB","1606-188","","","","",""],
    ["768-190","Նարինե","Արզումանյան","1606-190","3671-10","310-71-15","797-28","A50-M3-40-G14","13021-13"],
    ["768-191","Թագուհի","Գևորգյան","1606-191","","310-71-14","797-29","A50-M3-40-G15",""],
    ["768-192","Սուսաննա","Հարությունյան (Ռոզիկ)","1606-192","","","797-30","",""],
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
const getTierMgr=(sales,thresholds)=>{let t=1;for(let i=1;i<thresholds.length;i++)if(sales>=thresholds[i])t=i+1;return t;};
const getMgrPolicyRate=(p,cfg)=>{
  if(p.company==="Armenia"){const isShort=p.term==="SH"||(p.days!=null&&p.days<88);if(isShort)return(cfg.armeniaManager&&cfg.armeniaManager["1-9"])||0;return(cfg.armeniaManager&&cfg.armeniaManager[getArmGroup(p.bm)])||0;}
  return(cfg.managerRates&&cfg.managerRates[p.company])||0;
};
const getOpPolicyRate=(p,tier,cfg)=>{
  if(p.company==="Armenia"){const isShort=p.term==="SH"||(p.days!=null&&p.days<88);if(isShort)return 0;const grp=getArmGroup(p.bm);return(cfg.armeniaOperator&&cfg.armeniaOperator[grp]&&cfg.armeniaOperator[grp][tier-1])||0;}
  return(cfg.operatorRates&&cfg.operatorRates[p.company]&&cfg.operatorRates[p.company][tier-1])||0;
};
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

// ─── Office template download ────────────────────────────────────────────────
function downloadOfficeTemplate(){
  const wb=XLSXStyle.utils.book_new();
  const sHdr={font:{bold:true,color:{rgb:"FFFFFF"},sz:11},fill:{patternType:"solid",fgColor:{rgb:"1D4ED8"}},alignment:{horizontal:"center",wrapText:true},border:{bottom:{style:"thin",color:{rgb:"93C5FD"}}}};
  const sHint={font:{sz:10},fill:{patternType:"solid",fgColor:{rgb:"DBEAFE"}},alignment:{wrapText:true}};
  const sHintOpt={font:{sz:10},fill:{patternType:"solid",fgColor:{rgb:"F3F4F6"}},alignment:{wrapText:true}};
  const sEx={font:{italic:true,sz:10,color:{rgb:"92400E"}},fill:{patternType:"solid",fgColor:{rgb:"FEF9C3"}},alignment:{wrapText:true}};
  const sData={font:{sz:11},fill:{patternType:"solid",fgColor:{rgb:"FFFFFF"}}};
  const mkSheet=(headers,labels,hints,exRow)=>{
    const ws={};
    const range={s:{c:0,r:0},e:{c:headers.length-1,r:13}};
    headers.forEach((h,c)=>{
      ws[XLSXStyle.utils.encode_cell({r:0,c})]={v:h,t:"s",s:sHdr};
      ws[XLSXStyle.utils.encode_cell({r:1,c})]={v:labels[c]||h,t:"s",s:sHdr};
      ws[XLSXStyle.utils.encode_cell({r:2,c})]={v:hints[c]||"",t:"s",s:hints[c]&&hints[c].startsWith("*")?sHint:sHintOpt};
      ws[XLSXStyle.utils.encode_cell({r:3,c})]={v:exRow[c]||"",t:"s",s:sEx};
      for(let r=4;r<14;r++)ws[XLSXStyle.utils.encode_cell({r,c})]={v:"",t:"s",s:sData};
    });
    ws["!ref"]=XLSXStyle.utils.encode_range(range);
    ws["!cols"]=headers.map(()=>({wch:18}));
    return ws;
  };
  const h1=["polType","insuredName","phone","company","policyNum","date","dateStart","dateEnd","car","carPlate","bm","region","power","term","polStatus","amount","discount","agentUid","comment","paid","paymentType","paidAmount","paidDate"];
  const l1=["Тип","Страхователь","Телефон","Компания","№ полиса","Дата продажи","Дата начала","Дата конца","Марка авто","Гос. номер","КБМ","Регион","Мощность","Срок","Статус","Сумма","Скидка","Код агента","Комментарий","Оплачено","Способ оплаты","Оплач. сумма","Дата оплаты"];
  const hint1=["*osago","*Иванов Иван","*+37400000000","*Nairi / Ingo...","№ полиса","*ДД.ММ.ГГГГ","ДД.ММ.ГГГГ","ДД.ММ.ГГГГ","Toyota Camry","00 AA 000","1-25","YR/AG...","л.с.","*L или SH","статус","*число","число","код агента","текст","*TRUE/FALSE","cash/acba/ineco","число","ДД.ММ.ГГГГ"];
  const ex1=["osago","⚠ ПРИМЕР — не удалять","+37400000000","Nairi","AB123456","15.01.2024","15.01.2024","15.01.2025","Toyota Camry","00 AA 000","3","YR","105","L","","85000","0","768-101","","FALSE","","",""];
  const h2=["polType","insuredName","phone","company","policyNum","productName","date","amount","discount","agentUid","comment","paid","paymentType","paidAmount","paidDate"];
  const l2=["Тип","Страхователь","Телефон","Компания","№ полиса","Продукт","Дата продажи","Сумма","Скидка","Код агента","Комментарий","Оплачено","Способ оплаты","Оплач. сумма","Дата оплаты"];
  const hint2=["*voluntary","*Иванов Иван","*+37400000000","*компания","№ полиса","*название продукта","*ДД.ММ.ГГГГ","*число","число","код агента","текст","*TRUE/FALSE","cash/acba/ineco","число","ДД.ММ.ГГГГ"];
  const ex2=["voluntary","⚠ ПРИМЕР — не удалять","+37499000000","Nairi","VOL123","КАСКО","20.01.2024","150000","0","768-102","","FALSE","","",""];
  XLSXStyle.utils.book_append_sheet(wb,mkSheet(h1,l1,hint1,ex1),"АПРА (osago)");
  XLSXStyle.utils.book_append_sheet(wb,mkSheet(h2,l2,hint2,ex2),"Добровольные");
  XLSXStyle.writeFile(wb,"Шаблон_импорта_офис.xlsx");
}

// ─── OOXML chart helpers ──────────────────────────────────────────────────────
const escXml=s=>String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
const buildPieChartXml=(title,labels,values)=>`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${escXml(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
  <c:autoTitleDeleted val="0"/><c:plotArea><c:pieChart><c:varyColors val="1"/>
  <c:ser><c:idx val="0"/><c:order val="0"/>
  <c:cat><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="${labels.length}"/>${labels.map((l,i)=>`<c:pt idx="${i}"><c:v>${escXml(l)}</c:v></c:pt>`).join("")}</c:strCache></c:strRef></c:cat>
  <c:val><c:numRef><c:f>Sheet1!$B$1</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v,i)=>`<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></c:val>
  </c:ser><c:firstSliceAng val="0"/></c:pieChart></c:plotArea>
  <c:legend><c:legendPos val="r"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
const buildBarChartXml=(title,labels,values)=>`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${escXml(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>
  <c:autoTitleDeleted val="0"/><c:plotArea><c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:varyColors val="0"/>
  <c:ser><c:idx val="0"/><c:order val="0"/>
  <c:cat><c:strRef><c:f>Sheet1!$A$1</c:f><c:strCache><c:ptCount val="${labels.length}"/>${labels.map((l,i)=>`<c:pt idx="${i}"><c:v>${escXml(l)}</c:v></c:pt>`).join("")}</c:strCache></c:strRef></c:cat>
  <c:val><c:numRef><c:f>Sheet1!$B$1</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((v,i)=>`<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`).join("")}</c:numCache></c:numRef></c:val>
  </c:ser></c:barChart></c:plotArea>
  <c:legend><c:legendPos val="r"/></c:legend><c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
const buildDrawingXml=(rid,c1,r1,c2,r2)=>`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<xdr:twoCellAnchor><xdr:from><xdr:col>${c1}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r1}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>
<xdr:to><xdr:col>${c2}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${r2}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>
<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`;

const th={padding:"7px 10px",fontWeight:600,fontSize:12,whiteSpace:"nowrap",color:"#1e293b",borderBottom:"2px solid #cbd5e1",textAlign:"left",background:"#e2e8f0"};
const td={padding:"7px 10px",fontSize:13,borderBottom:"1px solid #e2e8f0",color:"#1e293b"};
const inp={border:"1.5px solid #94a3b8",borderRadius:8,padding:"4px 8px",fontSize:13,boxSizing:"border-box",color:"#1e293b",background:"#f8fafc"};
const btn=(bg,col,ex)=>({padding:"5px 12px",background:bg||"#2563eb",color:col||"#fff",border:"1.5px solid rgba(0,0,0,0.25)",borderRadius:8,cursor:"pointer",fontSize:12,fontWeight:700,...(ex||{})});
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

function MgrRatesPanel({cfg,onSave}){
  const[local,setLocal]=useState(()=>JSON.parse(JSON.stringify(cfg)));
  useEffect(()=>setLocal(JSON.parse(JSON.stringify(cfg))),[cfg]);
  const updArr=(arr,i,val,path)=>setLocal(p=>{const n=JSON.parse(JSON.stringify(p));const keys=path.split(".");let o=n;for(let k=0;k<keys.length;k++)o=o[keys[k]];o[i]=parseFloat(val)||0;return n;});
  const updObj=(path,val)=>setLocal(p=>{const n=JSON.parse(JSON.stringify(p));const keys=path.split(".");let o=n;for(let i=0;i<keys.length-1;i++)o=o[keys[i]];o[keys[keys.length-1]]=parseFloat(val)||0;return n;});
  const ni=(v,cb,w)=><input type="number" value={v||0} onChange={e=>cb(e.target.value)} style={{...inp,width:w||55,textAlign:"center"}}/>;
  const TIERS=["С1","С2","С3","С4","С5","С6"];
  return(
    <div style={{fontSize:13}}>
      <h4 style={{margin:"0 0 6px",color:"#7c3aed"}}>% менеджера (с каждого полиса оператора)</h4>
      <div style={{overflowX:"auto",marginBottom:14}}><table style={{borderCollapse:"collapse"}}><thead><tr>{[...COMPANIES,"Арм.1-9","Арм.10-14","Арм.15-25"].map(c=><th key={c} style={th}>{c}</th>)}</tr></thead><tbody><tr>{COMPANIES.map(c=><td key={c} style={td}>{ni((local.managerRates||{})[c],v=>updObj("managerRates."+c,v))}</td>)}{ARM_GROUPS.map(g=><td key={g} style={td}>{ni((local.armeniaManager||{})[g],v=>updObj("armeniaManager."+g,v))}</td>)}</tr></tbody></table></div>
      <h4 style={{margin:"0 0 4px",color:"#059669"}}>% оператора по ступеням (ОСАГО)</h4>
      <div style={{overflowX:"auto",marginBottom:14}}><table style={{borderCollapse:"collapse"}}><thead><tr><th style={th}>Ступень</th>{COMPANIES.map(c=><th key={c} style={th}>{c}</th>)}</tr></thead><tbody>{TIERS.map((t,i)=><tr key={t}><td style={{...td,fontWeight:600,color:"#374151"}}>{t}</td>{COMPANIES.map(c=><td key={c} style={td}>{ni(((local.operatorRates||{})[c]||[])[i],v=>updArr(local.operatorRates[c]||[],i,v,"operatorRates."+c))}</td>)}</tr>)}</tbody></table></div>
      <h4 style={{margin:"0 0 4px",color:"#059669"}}>% оператора Armenia по ступеням</h4>
      <div style={{overflowX:"auto",marginBottom:14}}><table style={{borderCollapse:"collapse"}}><thead><tr><th style={th}>Ступень</th>{ARM_GROUPS.map(g=><th key={g} style={th}>{"Арм."+g}</th>)}</tr></thead><tbody>{TIERS.map((t,i)=><tr key={t}><td style={{...td,fontWeight:600}}>{t}</td>{ARM_GROUPS.map(g=><td key={g} style={td}>{ni(((local.armeniaOperator||{})[g]||[])[i],v=>updArr((local.armeniaOperator||{})[g]||[],i,v,"armeniaOperator."+g))}</td>)}</tr>)}</tbody></table></div>
      <h4 style={{margin:"0 0 4px",color:"#d97706"}}>Пороги ступеней и фиксированная часть (AMD)</h4>
      <div style={{overflowX:"auto",marginBottom:14}}><table style={{borderCollapse:"collapse"}}><thead><tr><th style={th}>Параметр</th>{TIERS.map(t=><th key={t} style={th}>{t}</th>)}</tr></thead><tbody>
        <tr><td style={{...td,fontWeight:600}}>Порог</td>{(local.tierThresholds||[]).map((v,i)=><td key={i} style={td}>{ni(v,val=>{setLocal(p=>{const n=JSON.parse(JSON.stringify(p));n.tierThresholds[i]=parseFloat(val)||0;return n;});},80)}</td>)}</tr>
        <tr><td style={{...td,fontWeight:600}}>Фикс. часть</td>{(local.tierFixes||[]).map((v,i)=><td key={i} style={td}>{ni(v,val=>{setLocal(p=>{const n=JSON.parse(JSON.stringify(p));n.tierFixes[i]=parseFloat(val)||0;return n;});},80)}</td>)}</tr>
      </tbody></table></div>
      <div style={{display:"flex",gap:8}}><button onClick={()=>onSave(local)} style={btn("#16a34a")}>💾 Сохранить</button><button onClick={()=>setLocal(JSON.parse(JSON.stringify(DEFAULT_MGR_RATES)))} style={btn("#f3f4f6","#374151")}>Сбросить</button></div>
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
  const fileRef=useRef();const backRef=useRef();const officeFileRef=useRef();const importOfficeRef=useRef();
  const[importPending,setImportPending]=useState(null);
  const[importPreview,setImportPreview]=useState(null);
  const DEFAULT_EMPLOYEES=[
    {id:"emp1",name:"Сотрудник 1",pin:"111111",tabs:["policydb","officesales","cashbook","payroll"],viewOnly:false},
    {id:"emp2",name:"Сотрудник 2",pin:"222222",tabs:["policydb","officesales"],viewOnly:true},
  ];
  const[role,setRole]=useState(null);
  const[currentEmployee,setCurrentEmployee]=useState(null);
  const[employees,setEmployees]=useState(DEFAULT_EMPLOYEES);
  const[loginPin,setLoginPin]=useState("");
  const[loginError,setLoginError]=useState("");
  const[adminPin,setAdminPin]=useState("000000");
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
  const[managerConfig,setManagerConfig]=useState(DEFAULT_MGR_RATES);
  const[mgrDetail,setMgrDetail]=useState(null);
  const[showMgrSettings,setShowMgrSettings]=useState(false);
  const[mgrNewOp,setMgrNewOp]=useState("");
  const[officeExpenses,setOfficeExpenses]=useState({});
  const[officeExpLoaded,setOfficeExpLoaded]=useState(false);
  const[expNewName,setExpNewName]=useState("");
  const[lockedMonths,setLockedMonths]=useState({});
  const[monthSnapshot,setMonthSnapshot]=useState(null);
  const[searchQuery,setSearchQuery]=useState("");
  const[searchResults,setSearchResults]=useState(null);
  const[searchLoading,setSearchLoading]=useState(false);
  const[searchViewPol,setSearchViewPol]=useState(null);

  useEffect(()=>{(async()=>{
    try{const r=await calcStorage.get("agentDirectory").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(valD(p))setAgentDir(p);}else{setAgentDir(SEED_AGENTS);calcStorage.set("agentDirectory",JSON.stringify(SEED_AGENTS)).catch(()=>{});}}catch{setAgentDir(SEED_AGENTS);}
    try{const r=await calcStorage.get("ratesConfig").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(valR(p))setRates(p);}}catch{}
    try{const r=await calcStorage.get("volRates").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(valV(p))setVolRates(p);}}catch{}
    try{const r=await calcStorage.get("exceptionsConfig").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(valE(p))setExceptions(p);}}catch{}
    try{const r=await calcStorage.get("appSettings").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(p&&p.adminPin)setAdminPin(p.adminPin);if(p&&Array.isArray(p.officeStaff)&&p.officeStaff.length)setOfficeStaff(p.officeStaff);if(p&&Array.isArray(p.employees)&&p.employees.length)setEmployees(p.employees);}}catch{}
    try{const r=await calcStorage.get("managerConfig").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(p&&p.managerRates)setManagerConfig({...DEFAULT_MGR_RATES,...p,operatorUids:p.operatorUids||[]});}}catch{}
    try{const r=await calcStorage.get("officeExpenses").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(p&&typeof p==="object")setOfficeExpenses(p);}setOfficeExpLoaded(true);}catch{setOfficeExpLoaded(true);}
    try{const r=await calcStorage.get("lockedMonths").catch(()=>null);if(r&&r.value){const p=JSON.parse(r.value);if(p&&typeof p==="object")setLockedMonths(p);}}catch{}
  })();},[]);

  useEffect(()=>{
    setUploadedFiles([]);setVolSession([]);setOfficeSession(null);setActiveAgent(null);setOfficeCodes([]);setNewOfficeCode("");
    (async()=>{
      try{const r=await calcStorage.get("month:"+selMonth).catch(()=>null);if(r&&r.value){const d=JSON.parse(r.value);setStoredPols(d.policies||[]);setStoredVol(d.voluntary||[]);}else{setStoredPols([]);setStoredVol([]);}}catch{setStoredPols([]);setStoredVol([]);}
      try{const r=await calcStorage.get("officeStore:"+selMonth).catch(()=>null);if(r&&r.value)setStoredOffice(JSON.parse(r.value));else setStoredOffice({profit:0,rows:[]});}catch{setStoredOffice({profit:0,rows:[]});}
      try{const r=await calcStorage.get("officeCodes:"+selMonth).catch(()=>null);if(r&&r.value)setOfficeCodes(JSON.parse(r.value));else setOfficeCodes([]);}catch{setOfficeCodes([]);}
      try{const r=await calcStorage.get("monthSnapshot:"+selMonth).catch(()=>null);if(r&&r.value)setMonthSnapshot(JSON.parse(r.value));else setMonthSnapshot(null);}catch{setMonthSnapshot(null);}
    })();
  },[selMonth]);

  useEffect(()=>{
    if(role==="employee"&&currentEmployee){
      if(!currentEmployee.tabs.includes(tab))setTab(currentEmployee.tabs[0]||"policydb");
    }
  },[role,currentEmployee]);
  useEffect(()=>{
    if(tab!=="income"||!officeExpLoaded)return;
    setOfficeExpenses(prev=>{
      if(prev[selMonth])return prev;
      const months=Object.keys(prev).sort();
      const lastMo=months[months.length-1];
      const rows=lastMo
        ?prev[lastMo].filter(r=>r.type==="static").map(r=>({...r,id:"ex"+Math.random().toString(36).slice(2)}))
        :DEFAULT_EXPENSE_ROWS.map(r=>({...r}));
      const next={...prev,[selMonth]:rows};
      calcStorage.set("officeExpenses",JSON.stringify(next)).catch(()=>{});
      return next;
    });
  },[tab,selMonth,officeExpLoaded]);
  const saveDir=d=>{setAgentDir(d);calcStorage.set("agentDirectory",JSON.stringify(d)).catch(()=>{});};
  const saveOfficeCodes=codes=>{setOfficeCodes(codes);calcStorage.set("officeCodes:"+selMonth,JSON.stringify(codes)).catch(()=>{});};
  const addOfficeCode=()=>{const v=newOfficeCode.trim();if(!v)return;saveOfficeCodes([...officeCodes,v]);setNewOfficeCode("");};
  const removeOfficeCode=idx=>saveOfficeCodes(officeCodes.filter((_,i)=>i!==idx));
  const saveRates=r=>{setRates(r);calcStorage.set("ratesConfig",JSON.stringify(r)).catch(()=>{});};
  const saveVR=r=>{setVolRates(r);calcStorage.set("volRates",JSON.stringify(r)).catch(()=>{});};
  const saveExcs=e=>{setExceptions(e);calcStorage.set("exceptionsConfig",JSON.stringify(e)).catch(()=>{});};
  const isAdmin=role==="admin";
  const isViewOnly=!isAdmin&&currentEmployee?.viewOnly===true;
  const allowedTabs=isAdmin?["commissions","policydb","officesales","cashbook","payroll","manager","income","search"]:[...(currentEmployee?.tabs||[]),"search"];

  const saveAppSettings=(updates={})=>{
    const s={adminPin,employees,officeStaff,...updates};
    calcStorage.set("appSettings",JSON.stringify(s)).catch(()=>{});
  };
  const tryLogin=()=>{
    const p=loginPin.trim();
    if(!p){setLoginError("Введите PIN");return;}
    if(p===adminPin){setRole("admin");setCurrentEmployee(null);setLoginPin("");setLoginError("");return;}
    const emp=employees.find(e=>e.pin===p);
    if(emp){
      setRole("employee");setCurrentEmployee(emp);setLoginPin("");setLoginError("");
      return;
    }
    setLoginError("Неверный PIN");
  };
  const logout=()=>{setRole(null);setCurrentEmployee(null);setLoginPin("");setLoginError("");setPanel(null);};
  const saveOfficeStaff=(list)=>{setOfficeStaff(list);saveAppSettings({officeStaff:list});};
  const saveEmployees=(list)=>{setEmployees(list);saveAppSettings({employees:list});};
  const saveManagerConfig=cfg=>{setManagerConfig(cfg);calcStorage.set("managerConfig",JSON.stringify(cfg)).catch(()=>{});};
  const saveOfficeExpenses=data=>{setOfficeExpenses(data);calcStorage.set("officeExpenses",JSON.stringify(data)).catch(()=>{});};
  const changePin=()=>{
    if(!newPinA.trim()){setPinChangeMsg("Введите новый PIN");return;}
    if(newPinA!==newPinB){setPinChangeMsg("PIN не совпадают");return;}
    const np=newPinA.trim();
    setAdminPin(np);saveAppSettings({adminPin:np});
    setNewPinA("");setNewPinB("");
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

  const isLocked=lockedMonths[selMonth]===true;
  const effRates=isLocked&&monthSnapshot?monthSnapshot.rates:rates;
  const effVolRates=isLocked&&monthSnapshot?monthSnapshot.volRates:volRates;
  const effExceptions=isLocked&&monthSnapshot?monthSnapshot.exceptions:exceptions;
  const effAgentDir=isLocked&&monthSnapshot?monthSnapshot.agentDir:agentDir;

  const codeLookup=useMemo(()=>{
    const map={};
    Object.entries(effAgentDir).forEach(([aUid,agent])=>{
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
  },[effAgentDir]);

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
  },[uploadedFiles,effExceptions,codeLookup]);

  const effPols=useMemo(()=>{
    if(uploadedFiles.length>0)return sessionPols;
    return storedPols.map(p=>{const aUid=p.agentUid||codeLookup[p.company+":"+p.agentCode]||null;return{...p,agentUid:aUid,exception:checkExc(p,effExceptions,aUid)};});
  },[uploadedFiles,sessionPols,storedPols,effExceptions,codeLookup]);

  const effVol=useMemo(()=>{
    const vols=volSession.length>0?volSession:storedVol;
    return vols.map(v=>{
      const _ac=(v.agentCode||"").replace(/\s+/g,"").trim();
      const aUid=codeLookup[v.company+":"+_ac]||codeLookup[_ac]
        ||Object.keys(effAgentDir).find(uid=>Object.values(effAgentDir[uid].codes||{}).some(c=>c&&c.replace(/\s+/g,"").trim()===_ac))
        ||null;
      const vr=(effVolRates.rates||[]).find(r=>r.name===v.productName);
      const oR=vr?vr.officeRate:0;const aR=vr?vr.agentRate:0;
      return{...v,agentUid:aUid,officeRate:oR,agentRate:aR,officeComm:Math.round(v.amount*oR/100),agentComm:Math.round(v.amount*aR/100)};
    });
  },[volSession,storedVol,codeLookup,effVolRates,effAgentDir]);

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
          const oR=getOfficeRate(p,effRates);
          const aR=getAgentRate(p,key,effRates);
          return{...p,officeRate:oR,agentRate:aR,officeComm:Math.round(p.amount*oR/100),agentComm:Math.round(p.amount*aR/100)};
        });
        const totalOffice=enriched.reduce((s,p)=>s+p.officeComm,0);
        const totalAgent=enriched.reduce((s,p)=>s+p.agentComm,0);
        return{uid:key,policies:enriched,totalSales,validSales,totalOffice,totalAgent,profit:totalOffice-totalAgent};
      });
    }catch{return[];}
  },[effPols,effRates]);

  const totals=useMemo(()=>({
    office:agentData.reduce((s,a)=>s+a.totalOffice,0)+effVol.reduce((s,v)=>s+v.officeComm,0),
    agent:agentData.reduce((s,a)=>s+a.totalAgent,0)+effVol.reduce((s,v)=>s+v.agentComm,0),
    profit:agentData.reduce((s,a)=>s+a.profit,0)+effVol.reduce((s,v)=>s+v.officeComm-v.agentComm,0),
  }),[agentData,effVol]);

  const allExcs=useMemo(()=>agentData.flatMap(a=>a.policies.filter(p=>p.exception).map(p=>({...p,agentUid:a.uid}))),[agentData]);
  const detail=agentData.find(a=>a.uid===activeAgent);

  const runSearch=async()=>{
    const q=searchQuery.trim().toLowerCase();
    if(q.length<2)return;
    setSearchLoading(true);setSearchResults(null);
    const results=[];
    const norm=s=>String(s||"").toLowerCase().replace(/\s/g,"");
    const qn=norm(q);
    const match=p=>[p.insuredName,p.carPlate,p.policyNum,p.phone].some(v=>norm(v).includes(qn));
    try{
      const mk1=await calcStorage.list("month:").catch(()=>({keys:[]}));
      for(const key of(mk1.keys||[])){
        const r=await calcStorage.get(key).catch(()=>null);
        if(!r?.value)continue;
        const d=JSON.parse(r.value);
        const mk=key.replace("month:","");
        [...(d.policies||[]),...(d.voluntary||[])].forEach(p=>{if(match(p))results.push({...p,_source:"agent",_monthKey:mk});});
      }
      const mk2=await calcStorage.list("officePol:").catch(()=>({keys:[]}));
      for(const key of(mk2.keys||[])){
        const r=await calcStorage.get(key).catch(()=>null);
        if(!r?.value)continue;
        const pols=JSON.parse(r.value);
        const mk=key.replace("officePol:","");
        pols.forEach(p=>{if(match(p))results.push({...p,_source:"office",_monthKey:mk});});
      }
      results.sort((a,b)=>b._monthKey.localeCompare(a._monthKey));
      setSearchResults(results);
    }catch{setSearchResults([]);}
    setSearchLoading(false);
  };

  const saveMonth=()=>{
    const pols=agentData.flatMap(a=>a.policies);
    const snap={rates,volRates,exceptions,agentDir,managerConfig};
    calcStorage.set("month:"+selMonth,JSON.stringify({policies:pols,voluntary:effVol})).catch(()=>{});
    calcStorage.set("monthSnapshot:"+selMonth,JSON.stringify(snap)).catch(()=>{});
    if(officeSession)calcStorage.set("officeStore:"+selMonth,JSON.stringify(officeSession)).catch(()=>{});
    setSavedOk(true);setTimeout(()=>setSavedOk(false),2500);
  };

  const lockMonth=async()=>{
    if(!window.confirm("Закрыть "+fmtMonth(selMonth)+"?\nПосле закрытия ставки будут зафиксированы, а добавление/удаление полисов будет недоступно сотрудникам."))return;
    const snap={rates,volRates,exceptions,agentDir,managerConfig};
    await calcStorage.set("monthSnapshot:"+selMonth,JSON.stringify(snap)).catch(()=>{});
    const updated={...lockedMonths,[selMonth]:true};
    setLockedMonths(updated);setMonthSnapshot(snap);
    await calcStorage.set("lockedMonths",JSON.stringify(updated)).catch(()=>{});
  };

  const unlockMonth=async()=>{
    if(!window.confirm("Открыть "+fmtMonth(selMonth)+"?\nМесяц снова станет редактируемым."))return;
    const updated={...lockedMonths};delete updated[selMonth];
    setLockedMonths(updated);
    await calcStorage.set("lockedMonths",JSON.stringify(updated)).catch(()=>{});
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

  const handleImportOfficeFile=e=>{
    const f=e.target.files[0];if(!f)return;
    e.target.value="";
    const reader=new FileReader();
    reader.onload=evt=>{
      try{
        const wb=XLSX.read(new Uint8Array(evt.target.result),{type:"array"});
        const isValidDate=s=>{if(!s)return false;return/^\d{4}-\d{2}-\d{2}$/.test(s)&&!isNaN(new Date(s).getTime());};
        const parseDate=s=>{if(!s)return"";const m=String(s).match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);if(m)return`${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}`;return String(s).trim();};
        const parseBool=v=>{const s=String(v).trim().toLowerCase();return s==="true"||s==="1"||s==="yes";};
        const parseNum=v=>{const n=parseFloat(String(v).replace(/\s/g,"").replace(",","."));return isNaN(n)?0:n;};
        const allRows=[];
        [["АПРА (osago)","osago"],["Добровольные","voluntary"]].forEach(([sname,polType])=>{
          const ws=wb.Sheets[sname];if(!ws)return;
          const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
          const headers=(rows[0]||[]).map(h=>String(h).trim());
          for(let i=4;i<rows.length;i++){
            const row=rows[i];
            if(!row||row.every(c=>!String(c).trim()))continue;
            const obj={};
            headers.forEach((h,ci)=>{if(h)obj[h]=row[ci]!==undefined?String(row[ci]).trim():"";});
            if(!obj.insuredName&&!obj.policyNum)continue;
            const errs=[];
            if(!obj.insuredName)errs.push("Не заполнено имя");
            const parsedDate=parseDate(obj.date);
            if(!obj.date)errs.push("Не указана дата");
            else if(!isValidDate(parsedDate))errs.push("Неверный формат даты: "+obj.date);
            if(!obj.amount||parseNum(obj.amount)<=0)errs.push("Не указана сумма");
            if(polType==="osago"&&!obj.company)errs.push("Не указана компания");
            obj.polType=polType;
            obj.date=parsedDate;
            obj.dateStart=parseDate(obj.dateStart||"");
            obj.dateEnd=parseDate(obj.dateEnd||"");
            obj.paidDate=parseDate(obj.paidDate||"");
            obj.paid=parseBool(obj.paid);
            obj.amount=parseNum(obj.amount);
            obj.discount=parseNum(obj.discount||0);
            obj.paidAmount=obj.paid?parseNum(obj.paidAmount):null;
            obj.paymentType=obj.paid?(obj.paymentType||null):null;
            obj.paidAt=obj.paid&&obj.paidDate?(new Date(obj.paidDate).toISOString()):null;
            obj.power=parseInt(obj.power)||0;
            obj.agentUid=obj.agentUid||null;
            obj._id=genUid();
            obj._monthKey=selMonth;
            obj._rowNum=i+1;
            obj._sheetName=sname;
            obj._errors=errs;
            obj._valid=errs.length===0;
            allRows.push(obj);
          }
        });
        if(!allRows.length){alert("Данных для импорта не найдено. Данные должны начинаться с 5-й строки.");return;}
        setImportPreview({rows:allRows,month:selMonth});
      }catch(err){alert("Ошибка чтения файла: "+err.message);}
    };
    reader.readAsArrayBuffer(f);
  };

  const confirmImport=validPols=>{
    setImportPreview(null);
    if(opCurrentMonth.length>0){
      setImportPending({pols:validPols,month:selMonth});
    }else{
      saveOpMonth(validPols);
      alert("Импортировано "+validPols.length+" записей за "+fmtMonth(selMonth));
    }
  };

  const exportIncomeExcel=async()=>{
    const wb=new JSZip();
    const month=selMonth;
    const rows=officeExpenses[month]||[];
    const opPols=cashMonthPols||[];
    const totalExp=rows.reduce((s,r)=>s+(parseFloat(r.amount)||0),0);
    const osagoGross=opPols.filter(p=>(p.polType||"osago")==="osago").reduce((s,p)=>s+(p.amount||0),0);
    const volGross=opPols.filter(p=>p.polType==="voluntary").reduce((s,p)=>s+(p.amount||0),0);
    const totalGross=osagoGross+volGross;
    const agentCommissions=agentData.reduce((s,a)=>s+a.totalAgent,0)+effVol.reduce((s,v)=>s+v.agentComm,0);
    const netIncome=totalGross-agentCommissions-totalExp;

    const sTitle={font:{bold:true,sz:14,color:{rgb:"1E293B"}},fill:{patternType:"solid",fgColor:{rgb:"E0E7FF"}},alignment:{horizontal:"center"}};
    const sSec={font:{bold:true,sz:12,color:{rgb:"FFFFFF"}},fill:{patternType:"solid",fgColor:{rgb:"3730A3"}},alignment:{horizontal:"left"}};
    const sColHdr={font:{bold:true,sz:11,color:{rgb:"374151"}},fill:{patternType:"solid",fgColor:{rgb:"F0F4F8"}},alignment:{horizontal:"center"}};
    const sTxt=(rgb)=>({font:{sz:11},fill:{patternType:"solid",fgColor:{rgb:rgb||"FFFFFF"}},alignment:{horizontal:"left"}});
    const sNum=(rgb)=>({font:{sz:11},fill:{patternType:"solid",fgColor:{rgb:rgb||"FFFFFF"}},alignment:{horizontal:"right"}});

    const mkCell=(v,s,t)=>({v,t:t||"s",s});
    const mkRow=cells=>cells;

    // Sheet 1: Summary
    const s1Data=[
      [mkCell("Сводка — "+fmtMonth(month),sTitle,"s")],
      [],
      [mkCell("Показатель",sColHdr),mkCell("Сумма (AMD)",sColHdr)],
      [mkCell("Валовые продажи АПРА",sTxt()),mkCell(osagoGross,sNum(),"n")],
      [mkCell("Валовые продажи Добровольные",sTxt()),mkCell(volGross,sNum(),"n")],
      [mkCell("Итого валовые",sTxt("DBEAFE")),mkCell(totalGross,sNum("DBEAFE"),"n")],
      [],
      [mkCell("Комиссии агентов",sTxt("FEE2E2")),mkCell(agentCommissions,sNum("FEE2E2"),"n")],
      [mkCell("Расходы офиса",sTxt("FEE2E2")),mkCell(totalExp,sNum("FEE2E2"),"n")],
      [],
      [mkCell("Чистый доход",sTxt("DCFCE7")),mkCell(netIncome,sNum("DCFCE7"),"n")],
    ];
    const ws1=XLSXStyle.utils.aoa_to_sheet(s1Data.map(r=>r.map(c=>c&&c.v!==undefined?c:({v:"",t:"s",s:sTxt()}))));
    s1Data.forEach((row,ri)=>row.forEach((cell,ci)=>{if(cell&&cell.s){const addr=XLSXStyle.utils.encode_cell({r:ri,c:ci});if(ws1[addr])ws1[addr].s=cell.s;}}));
    ws1["!merges"]=[{s:{r:0,c:0},e:{r:0,c:1}}];
    ws1["!cols"]=[{wch:35},{wch:18}];

    // Sheet 2: Expenses
    const expRows=[[mkCell("Расходы — "+fmtMonth(month),sTitle)],[],[mkCell("Категория",sColHdr),mkCell("Название",sColHdr),mkCell("Сумма",sColHdr)],...rows.map(r=>[mkCell(r.cat||"",sTxt()),mkCell(r.name||"",sTxt()),mkCell(r.amount||0,sNum(),"n")]),[],[mkCell("ИТОГО",sTxt("FEE2E2")),mkCell("",sTxt("FEE2E2")),mkCell(totalExp,sNum("FEE2E2"),"n")]];
    const ws2=XLSXStyle.utils.aoa_to_sheet(expRows.map(r=>r.map(c=>c&&c.v!==undefined?c:({v:"",t:"s",s:sTxt()}))));
    expRows.forEach((row,ri)=>row.forEach((cell,ci)=>{if(cell&&cell.s){const addr=XLSXStyle.utils.encode_cell({r:ri,c:ci});if(ws2[addr])ws2[addr].s=cell.s;}}));
    ws2["!merges"]=[{s:{r:0,c:0},e:{r:0,c:2}}];
    ws2["!cols"]=[{wch:20},{wch:30},{wch:16}];

    // Sheet 3: Income sources
    const incSources=[["АПРА",osagoGross],["Добровольные",volGross]];
    const incRows=[[mkCell("Доходы — "+fmtMonth(month),sTitle)],[],[mkCell("Источник",sColHdr),mkCell("Сумма",sColHdr)],...incSources.map(([n,v])=>[mkCell(n,sTxt()),mkCell(v,sNum(),"n")])]
    const ws3=XLSXStyle.utils.aoa_to_sheet(incRows.map(r=>r.map(c=>c&&c.v!==undefined?c:({v:"",t:"s",s:sTxt()}))));
    incRows.forEach((row,ri)=>row.forEach((cell,ci)=>{if(cell&&cell.s){const addr=XLSXStyle.utils.encode_cell({r:ri,c:ci});if(ws3[addr])ws3[addr].s=cell.s;}}));
    ws3["!merges"]=[{s:{r:0,c:0},e:{r:0,c:1}}];
    ws3["!cols"]=[{wch:25},{wch:16}];

    // Build xlsx via XLSXStyle then inject charts via JSZip
    const xlsxBlob=XLSXStyle.write({SheetNames:["Сводка","Расходы","Доходы"],Sheets:{Сводка:ws1,Расходы:ws2,Доходы:ws3}},{type:"array",bookType:"xlsx"});
    const zip=await JSZip.loadAsync(xlsxBlob);

    // Add pie chart for income
    const pieXml=buildPieChartXml("Источники дохода",incSources.map(s=>s[0]),incSources.map(s=>s[1]));
    zip.file("xl/charts/chart1.xml",pieXml);
    zip.file("xl/drawings/drawing1.xml",buildDrawingXml("rId1",3,4,9,18));
    zip.file("xl/drawings/_rels/drawing1.xml.rels",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/></Relationships>`);

    // Add bar chart for expenses
    const expLabels=rows.map(r=>r.name||r.cat||"");
    const expVals=rows.map(r=>parseFloat(r.amount)||0);
    const barXml=buildBarChartXml("Расходы по категориям",expLabels,expVals);
    zip.file("xl/charts/chart2.xml",barXml);
    zip.file("xl/drawings/drawing2.xml",buildDrawingXml("rId1",3,4,9,18));
    zip.file("xl/drawings/_rels/drawing2.xml.rels",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart2.xml"/></Relationships>`);

    // Patch Content_Types.xml
    const ctXml=await zip.file("[Content_Types].xml").async("string");
    const ctPatched=ctXml.replace("</Types>",`<Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/xl/charts/chart2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/drawings/drawing2.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`);
    zip.file("[Content_Types].xml",ctPatched);

    const out=await zip.generateAsync({type:"uint8array"});
    const blob=new Blob([out],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url;a.download="Доходы_"+month+".xlsx";document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
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

  // Shared helper: build one operator worksheet. showMgr=true includes manager income columns.
  const _buildOpSheet=(r,cfg,agentDir,month,excepts,showMgr)=>{
    const br={style:"thin",color:{rgb:"D1D5DB"}};const borders={top:br,bottom:br,left:br,right:br};
    const sDark=(rgb,al)=>({fill:{patternType:"solid",fgColor:{rgb:rgb||"1E293B"}},font:{bold:true,sz:10,color:{rgb:"FFFFFF"}},border:borders,alignment:{horizontal:al||"left",wrapText:true}});
    const sLight=(rgb,al)=>({fill:{patternType:"solid",fgColor:{rgb:rgb||"E5E7EB"}},font:{bold:true,sz:10},border:borders,alignment:{horizontal:al||"center",wrapText:true}});
    const sC=(bg,bold)=>({fill:bg?{patternType:"solid",fgColor:{rgb:bg}}:{},font:{sz:10,bold:!!bold},border:borders,alignment:{horizontal:"left"}});
    const sN=(bg,bold)=>({fill:bg?{patternType:"solid",fgColor:{rgb:bg}}:{},font:{sz:10,bold:!!bold},border:borders,alignment:{horizontal:"right"}});
    const ac=(ws,row,c,v,s)=>{ws[XLSXStyle.utils.encode_cell({r:row,c})]={v,t:typeof v==="number"?"n":"s",s};};
    const gn=uid=>{const a=agentDir[uid];return a?(a.name+" "+a.surname).trim():uid||"";};
    const gc=uid=>{const a=agentDir[uid];return(a&&a.internalCode)||"";};
    const ws={};let rw=0;const mg=[];
    const VHDRS=showMgr
      ?["№ полиса","Компания","Страхователь","Марка","Рег. номер","Регион","БМ","Мощность","Срок","Сумма","% Мен.","Доход мен.","% Опер.","Выплата опер."]
      :["№ полиса","Компания","Страхователь","Марка","Рег. номер","Регион","БМ","Мощность","Срок","Сумма","% Опер.","Выплата опер."];
    const EHDRS=[...VHDRS,"Причина исключения"];
    const nc=EHDRS.length-1;
    // Header row
    ac(ws,rw,0,(gn(r.uid)||gc(r.uid)||r.uid)+" — "+fmtMonth(month),sDark("4C1D95","center"));mg.push({s:{r:rw,c:0},e:{r:rw,c:nc}});rw++;
    // Info row
    const infoBase=["Ступень: С"+r.tier,"Зачётные: "+r.validSales.toLocaleString("ru-RU")+" AMD","Фикс: "+r.fix.toLocaleString("ru-RU")+" AMD","Начислено: "+r.oi.toLocaleString("ru-RU")+" AMD","К выплате: "+(r.oi+r.fix).toLocaleString("ru-RU")+" AMD"];
    const infoMgr=showMgr?["Доход мен.: "+r.mi.toLocaleString("ru-RU")+" AMD","Прибыль: "+r.profit.toLocaleString("ru-RU")+" AMD"]:[];
    [...infoBase,...infoMgr].forEach((v,c)=>ac(ws,rw,c,v,sLight("EDE9FE","left")));rw++;
    // Valid policies
    const vp=r.policies.filter(p=>!p.exception);const ep=r.policies.filter(p=>p.exception);
    if(vp.length>0){
      ac(ws,rw,0,"ЗАЧЁТНЫЕ ПОЛИСЫ ("+vp.length+")",sDark("166534","center"));mg.push({s:{r:rw,c:0},e:{r:rw,c:nc}});rw++;
      VHDRS.forEach((h,c)=>ac(ws,rw,c,h,sLight("D1FAE5")));rw++;
      vp.forEach((p,i)=>{
        const bg=i%2===0?"FFFFFF":"F0FDF4";const mr=getMgrPolicyRate(p,cfg);const or=getOpPolicyRate(p,r.tier,cfg);
        const vals=showMgr
          ?[p.policyNum||"",p.company||"",p.insuredName||"",p.car||"",p.carPlate||"",p.region||"",p.bm||0,p.power||0,p.term||"",p.amount||0,mr,Math.round(p.amount*mr/100),or,Math.round(p.amount*or/100)]
          :[p.policyNum||"",p.company||"",p.insuredName||"",p.car||"",p.carPlate||"",p.region||"",p.bm||0,p.power||0,p.term||"",p.amount||0,or,Math.round(p.amount*or/100)];
        vals.forEach((v,c)=>ac(ws,rw,c,v,typeof v==="number"?sN(bg):sC(bg)));rw++;
      });
      const totAmt=vp.reduce((s,p)=>s+p.amount,0);
      const totRow=showMgr?["ИТОГО","","","","","","","","",totAmt,"",r.mi,"",r.oi]:["ИТОГО","","","","","","","","",totAmt,"",r.oi];
      totRow.forEach((v,c)=>ac(ws,rw,c,v,typeof v==="number"?sN("E5E7EB",true):sC("E5E7EB",true)));rw+=2;
    }
    if(ep.length>0){
      ac(ws,rw,0,"НЕЗАЧЁТНЫЕ ПОЛИСЫ ("+ep.length+")",sDark("991B1B","center"));mg.push({s:{r:rw,c:0},e:{r:rw,c:nc}});rw++;
      EHDRS.forEach((h,c)=>ac(ws,rw,c,h,sLight("FEE2E2")));rw++;
      ep.forEach((p,i)=>{
        const bg=i%2===0?"FFFFFF":"FFF7ED";const reason=excReason(p,excepts,r.uid);
        const vals=showMgr
          ?[p.policyNum||"",p.company||"",p.insuredName||"",p.car||"",p.carPlate||"",p.region||"",p.bm||0,p.power||0,p.term||"",p.amount||0,"—","—","—","—",reason]
          :[p.policyNum||"",p.company||"",p.insuredName||"",p.car||"",p.carPlate||"",p.region||"",p.bm||0,p.power||0,p.term||"",p.amount||0,"—","—",reason];
        vals.forEach((v,c)=>ac(ws,rw,c,v,typeof v==="number"?sN(bg):sC(bg)));rw++;
      });
    }
    ws["!ref"]=XLSXStyle.utils.encode_range({s:{r:0,c:0},e:{r:rw,c:nc}});
    ws["!cols"]=showMgr
      ?[{wch:16},{wch:10},{wch:26},{wch:18},{wch:13},{wch:8},{wch:6},{wch:8},{wch:6},{wch:14},{wch:8},{wch:14},{wch:8},{wch:14},{wch:35}]
      :[{wch:16},{wch:10},{wch:26},{wch:18},{wch:13},{wch:8},{wch:6},{wch:8},{wch:6},{wch:14},{wch:8},{wch:14},{wch:35}];
    ws["!merges"]=mg;
    return ws;
  };

  const _dlXlsx=(wb,name)=>{
    const out=XLSXStyle.write(wb,{bookType:"xlsx",type:"array"});
    const blob=new Blob([new Uint8Array(out)],{type:"application/octet-stream"});
    const url=URL.createObjectURL(blob);const a=document.createElement("a");
    a.href=url;a.download=name;document.body.appendChild(a);a.click();document.body.removeChild(a);URL.revokeObjectURL(url);
  };

  const _computeOpR=(agentData,cfg)=>agentData.filter(a=>(cfg.operatorUids||[]).includes(a.uid)).map(op=>{
    const thrs=cfg.tierThresholds||DEFAULT_MGR_RATES.tierThresholds;
    const fixes=cfg.tierFixes||DEFAULT_MGR_RATES.tierFixes;
    const tier=getTierMgr(op.validSales,thrs);const fix=fixes[tier-1]||0;
    let mi=0,oi=0;
    op.policies.filter(p=>!p.exception).forEach(p=>{mi+=Math.round(p.amount*getMgrPolicyRate(p,cfg)/100);oi+=Math.round(p.amount*getOpPolicyRate(p,tier,cfg)/100);});
    return{uid:op.uid,totalSales:op.totalSales,validSales:op.validSales,tier,fix,mi,oi,profit:mi-oi-fix,policies:op.policies};
  });

  const exportManagerXlsx=(agentData,cfg,agentDir,month,excepts)=>{
    const br={style:"thin",color:{rgb:"D1D5DB"}};const borders={top:br,bottom:br,left:br,right:br};
    const sDark=(rgb,al)=>({fill:{patternType:"solid",fgColor:{rgb:rgb||"1E293B"}},font:{bold:true,sz:10,color:{rgb:"FFFFFF"}},border:borders,alignment:{horizontal:al||"left",wrapText:true}});
    const sLight=(rgb,al)=>({fill:{patternType:"solid",fgColor:{rgb:rgb||"E5E7EB"}},font:{bold:true,sz:10},border:borders,alignment:{horizontal:al||"center",wrapText:true}});
    const sC=(bg,bold,al)=>({fill:bg?{patternType:"solid",fgColor:{rgb:bg}}:{},font:{sz:10,bold:!!bold},border:borders,alignment:{horizontal:al||"left"}});
    const sN=(bg,bold,col)=>({fill:bg?{patternType:"solid",fgColor:{rgb:bg}}:{},font:{sz:10,bold:!!bold,color:col?{rgb:col}:undefined},border:borders,alignment:{horizontal:"right"}});
    const ac=(ws,r,c,v,s)=>{ws[XLSXStyle.utils.encode_cell({r,c})]={v,t:typeof v==="number"?"n":"s",s};};
    const gn=uid=>{const a=agentDir[uid];return a?(a.name+" "+a.surname).trim():uid||"";};
    const gc=uid=>{const a=agentDir[uid];return(a&&a.internalCode)||"";};
    const opR=_computeOpR(agentData,cfg);
    const totMgr=opR.reduce((s,r)=>s+r.mi,0);const totOp=opR.reduce((s,r)=>s+r.oi,0);const totFix=opR.reduce((s,r)=>s+r.fix,0);const totProfit=opR.reduce((s,r)=>s+r.profit,0);
    const sT=(al)=>({fill:{patternType:"solid",fgColor:{rgb:"111827"}},font:{bold:true,sz:10,color:{rgb:"FFFFFF"}},border:borders,alignment:{horizontal:al||"right"}});
    const mkSummary=(showMgr)=>{
      const ws={};let rw=0;const mg=[];
      const ncols=showMgr?8:6;
      const title=showMgr?"Менеджерский отчёт — "+fmtMonth(month):"Отчёт по операторам — "+fmtMonth(month);
      ac(ws,rw,0,title,sDark("4C1D95","center"));mg.push({s:{r:rw,c:0},e:{r:rw,c:ncols}});rw++;
      const hdrs=showMgr
        ?["Оператор","Код","Всего продаж","Зачётные","Ступень","Фикс (AMD)","Доход мен. (AMD)","Выплата опер. (AMD)","Прибыль (AMD)"]
        :["Оператор","Код","Всего продаж","Зачётные","Ступень","Фикс (AMD)","Начислено (AMD)","К выплате (AMD)"];
      hdrs.forEach((h,c)=>ac(ws,rw,c,h,sDark("374151","center")));rw++;
      opR.forEach((r,i)=>{
        const bg=i%2===0?"FFFFFF":"F5F3FF";
        const row=showMgr
          ?[gn(r.uid),gc(r.uid),r.totalSales,r.validSales,"С"+r.tier,r.fix,r.mi,r.oi,r.profit]
          :[gn(r.uid),gc(r.uid),r.totalSales,r.validSales,"С"+r.tier,r.fix,r.oi,r.oi+r.fix];
        row.forEach((v,c)=>{
          const isProfit=showMgr&&c===8;
          const s=isProfit?sN(bg,true,r.profit>=0?"166534":"DC2626"):c>=2?sN(bg,c>=5):sC(bg,false,c===0?"left":"center");
          ac(ws,rw,c,v,s);
        });rw++;
      });
      const totRow=showMgr
        ?["ИТОГО","",opR.reduce((s,r)=>s+r.totalSales,0),opR.reduce((s,r)=>s+r.validSales,0),"",totFix,totMgr,totOp,totProfit]
        :["ИТОГО","",opR.reduce((s,r)=>s+r.totalSales,0),opR.reduce((s,r)=>s+r.validSales,0),"",totFix,totOp,totOp+totFix];
      totRow.forEach((v,c)=>ac(ws,rw,c,v,sT(c===0?"left":"right")));rw+=2;
      if(showMgr){
        ac(ws,rw,0,"Офис должен менеджеру:",sLight("FDE68A","left"));
        ac(ws,rw,1,totMgr,{fill:{patternType:"solid",fgColor:{rgb:"FDE68A"}},font:{bold:true,sz:12},border:borders,alignment:{horizontal:"right"}});
        mg.push({s:{r:rw,c:0},e:{r:rw,c:0}});
      }
      ws["!ref"]=XLSXStyle.utils.encode_range({s:{r:0,c:0},e:{r:rw,c:ncols}});
      ws["!cols"]=showMgr?[{wch:28},{wch:12},{wch:16},{wch:16},{wch:10},{wch:14},{wch:18},{wch:18},{wch:16}]:[{wch:28},{wch:12},{wch:16},{wch:16},{wch:10},{wch:14},{wch:16},{wch:16}];
      ws["!merges"]=mg;
      return ws;
    };
    // File 1: manager (with mgr income/profit)
    const wb1=XLSXStyle.utils.book_new();
    XLSXStyle.utils.book_append_sheet(wb1,mkSummary(true),"Сводка");
    opR.forEach(r=>{const sn=(gc(r.uid)||gn(r.uid)).slice(0,31)||("Op"+r.uid.slice(0,28));XLSXStyle.utils.book_append_sheet(wb1,_buildOpSheet(r,cfg,agentDir,month,excepts,true),sn);});
    _dlXlsx(wb1,"Менеджер_"+month+".xlsx");
    // File 2: operators (without mgr income/profit)
    const wb2=XLSXStyle.utils.book_new();
    XLSXStyle.utils.book_append_sheet(wb2,mkSummary(false),"Сводка");
    opR.forEach(r=>{const sn=(gc(r.uid)||gn(r.uid)).slice(0,31)||("Op"+r.uid.slice(0,28));XLSXStyle.utils.book_append_sheet(wb2,_buildOpSheet(r,cfg,agentDir,month,excepts,false),sn);});
    setTimeout(()=>_dlXlsx(wb2,"Операторы_"+month+".xlsx"),300);
  };

  const exportSingleOpXlsx=(r,cfg,agentDir,month,excepts)=>{
    const wb=XLSXStyle.utils.book_new();
    const gn=uid=>{const a=agentDir[uid];return a?(a.name+" "+a.surname).trim():uid||"";};
    const gc=uid=>{const a=agentDir[uid];return(a&&a.internalCode)||"";};
    const sn=(gc(r.uid)||gn(r.uid)).slice(0,31)||("Op"+r.uid.slice(0,28));
    XLSXStyle.utils.book_append_sheet(wb,_buildOpSheet(r,cfg,agentDir,month,excepts,false),sn);
    const name=gn(r.uid)||gc(r.uid)||r.uid;
    _dlXlsx(wb,name.slice(0,30)+"_"+month+".xlsx");
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

  if(role===null){
    return(
      <div style={{minHeight:"100vh",background:"#dde3ed",display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"system-ui,sans-serif"}}>
        <div style={{background:"white",borderRadius:16,padding:"40px 48px",boxShadow:"0 8px 32px rgba(0,0,0,0.12)",minWidth:320,textAlign:"center"}}>
          <div style={{fontSize:32,marginBottom:8}}>🏢</div>
          <h2 style={{margin:"0 0 4px",fontSize:22,color:"#1e293b"}}>INSURANCE MANAGER</h2>
          <div style={{color:"#64748b",fontSize:13,marginBottom:28}}>Введите PIN-код для входа</div>
          <input
            type="password"
            value={loginPin}
            onChange={e=>setLoginPin(e.target.value)}
            onKeyDown={e=>e.key==="Enter"&&tryLogin()}
            placeholder="••••••"
            maxLength={8}
            style={{...inp,width:"100%",fontSize:22,textAlign:"center",letterSpacing:6,padding:"10px 16px",marginBottom:12,boxSizing:"border-box"}}
            autoFocus
          />
          {loginError&&<div style={{color:"#dc2626",fontSize:13,marginBottom:10,fontWeight:600}}>{loginError}</div>}
          <button onClick={tryLogin} style={{...btn("#1d4ed8"),width:"100%",fontSize:15,padding:"10px 0"}}>Войти</button>
        </div>
      </div>
    );
  }

  return(
    <div style={{fontFamily:"system-ui,sans-serif",padding:20,maxWidth:1400,margin:"0 auto",color:"#1e293b",background:"#dde3ed",minHeight:"100vh"}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12,flexWrap:"wrap",gap:8}}>
        <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap"}}>
          <h2 style={{margin:0,fontSize:20}}>Калькулятор комиссий</h2>
          {isAdmin
            ?<><span style={{background:"#fef3c7",border:"1px solid #fcd34d",borderRadius:6,padding:"4px 10px",fontSize:12,fontWeight:600,color:"#92400e"}}>🔑 Администратор</span></>
            :<><span style={{background:"#f0fdf4",border:"1px solid #86efac",borderRadius:6,padding:"4px 10px",fontSize:12,fontWeight:600,color:"#166534"}}>👤 {currentEmployee?.name||"Сотрудник"}</span></>
          }
          <button onClick={logout} style={btn("#64748b",undefined,{fontSize:12})}>Выйти</button>
        </div>
        <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>
          {panels.map(([id,label,count])=>(
            <button key={id} onClick={()=>setPanel(panel===id?null:id)} style={{...btn(panel===id?"#eff6ff":"#f9fafb",panel===id?"#1d4ed8":"#374151"),border:"2px solid "+(panel===id?"#6366f1":"#94a3b8"),fontWeight:panel===id?700:500}}>
              {label}{count!=null?" ("+count+")":""}
            </button>
          ))}
          {isAdmin&&hasData&&<button onClick={()=>exportToExcel(agentData,effVol,agentDir,totals,exceptions)} style={btn("#16a34a")}>⬇ Excel</button>}
          {isAdmin&&<button onClick={()=>setShowBackup(p=>!p)} style={btn("#7c3aed")}>💾 Резерв</button>}
          {isAdmin&&<button onClick={()=>backRef.current.click()} style={btn("#f3f4f6","#374151",{border:"1px solid #d1d5db"})}>📂 Восстановить</button>}
          <input ref={backRef} type="file" accept=".json" onChange={importBackup} style={{display:"none"}}/>
        </div>
      </div>

      <div style={{display:"flex",marginBottom:16,gap:6,flexWrap:"wrap"}}>
        {[["commissions","💰 Комиссии","#1d4ed8","#dbeafe"],["policydb","📋 База полисов","#0f766e","#ccfbf1"],["officesales","🏢 Продажи офиса","#7c3aed","#ede9fe"],["cashbook","📒 Касса","#b45309","#fef3c7"],["payroll","📝 Начисления","#0369a1","#e0f2fe"],["manager","👔 Менеджер","#be185d","#fce7f3"],["income","📊 Доходы офиса","#15803d","#dcfce7"],["search","🔍 Поиск","#0f172a","#e2e8f0"]].filter(([id])=>allowedTabs.includes(id)).map(([id,label,activeCol,activeBg])=>(
          <button key={id} onClick={()=>{setTab(id);if(id==="policydb")loadDB();}}
            style={{padding:"8px 18px",background:tab===id?activeBg:"#e2e8f0",color:"#0f172a",border:tab===id?"2px solid "+activeCol:"2px solid #94a3b8",borderRadius:10,fontSize:13,fontWeight:700,cursor:"pointer",transition:"all 0.15s"}}>
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
              <div style={{borderTop:"1px solid #e5e7eb",paddingTop:16,marginBottom:20}}>
                <div style={{fontWeight:600,fontSize:13,marginBottom:10,color:"#374151"}}>👥 Сотрудники</div>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13,marginBottom:10}}>
                  <thead><tr>
                    <th style={{...th,textAlign:"left"}}>Имя</th>
                    <th style={th}>PIN</th>
                    <th style={th}>Доступные вкладки</th>
                    <th style={th}>Только просмотр</th>
                    <th style={th}></th>
                  </tr></thead>
                  <tbody>
                    {employees.map((emp,i)=>{
                      const TAB_LABELS={policydb:"📋 База",officesales:"🏢 Продажи",cashbook:"📒 Касса",payroll:"📝 Начисления"};
                      return(
                        <tr key={emp.id}>
                          <td style={td}><input value={emp.name} onChange={e=>{const l=[...employees];l[i]={...l[i],name:e.target.value};saveEmployees(l);}} style={{...inp,width:"100%"}}/></td>
                          <td style={td}><input type="password" value={emp.pin} onChange={e=>{const l=[...employees];l[i]={...l[i],pin:e.target.value};saveEmployees(l);}} style={{...inp,width:90,letterSpacing:3}}/></td>
                          <td style={{...td,fontSize:11}}>
                            <div style={{display:"flex",flexWrap:"wrap",gap:4}}>
                              {Object.entries(TAB_LABELS).map(([id,label])=>(
                                <label key={id} style={{display:"flex",alignItems:"center",gap:3,cursor:"pointer",userSelect:"none"}}>
                                  <input type="checkbox" checked={emp.tabs.includes(id)} onChange={e=>{const l=[...employees];const tabs=e.target.checked?[...emp.tabs,id]:emp.tabs.filter(t=>t!==id);l[i]={...l[i],tabs};saveEmployees(l);}}/>
                                  {label}
                                </label>
                              ))}
                            </div>
                          </td>
                          <td style={{...td,textAlign:"center"}}>
                            <input type="checkbox" checked={emp.viewOnly||false} onChange={e=>{const l=[...employees];l[i]={...l[i],viewOnly:e.target.checked};saveEmployees(l);}}/>
                          </td>
                          <td style={{...td,textAlign:"center"}}>
                            <button onClick={()=>{if(window.confirm("Удалить сотрудника "+emp.name+"?"))saveEmployees(employees.filter((_,j)=>j!==i));}} style={{background:"none",border:"none",cursor:"pointer",color:"#dc2626",fontSize:16}}>×</button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                <button onClick={()=>saveEmployees([...employees,{id:"emp"+Date.now(),name:"Новый сотрудник",pin:"000000",tabs:["policydb","officesales"],viewOnly:false}])} style={btn("#2563eb",undefined,{fontSize:12})}>+ Добавить сотрудника</button>
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
              {isLocked&&<span style={{background:"#fef3c7",border:"1px solid #fcd34d",borderRadius:6,padding:"3px 8px",fontSize:11,fontWeight:700,color:"#92400e"}}>🔒 Закрыт</span>}
              {(storedPols.length>0||storedVol.length>0)&&uploadedFiles.length===0&&volSession.length===0&&<span style={{fontSize:11,color:"#6b7280"}}>{"📥 "+storedPols.length+" полисов из хранилища"}</span>}
              <button onClick={saveMonth} style={btn("#16a34a",undefined,{fontSize:11})}>💾 Сохранить месяц</button>
              {isAdmin&&(isLocked
                ?<button onClick={unlockMonth} style={btn("#dc2626",undefined,{fontSize:11})}>🔒 Открыть месяц</button>
                :<button onClick={lockMonth} style={btn("#92400e",undefined,{fontSize:11})}>🔓 Закрыть месяц</button>
              )}
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
              {isLocked&&<span style={{background:"#fef3c7",border:"1px solid #fcd34d",borderRadius:6,padding:"3px 8px",fontSize:11,fontWeight:700,color:"#92400e"}}>🔒 Закрыт</span>}
              {isAdmin&&(isLocked
                ?<button onClick={unlockMonth} style={btn("#dc2626",undefined,{fontSize:11})}>🔒 Открыть месяц</button>
                :<button onClick={lockMonth} style={btn("#92400e",undefined,{fontSize:11})}>🔓 Закрыть месяц</button>
              )}
              {(!isLocked||isAdmin)&&!isViewOnly&&<button onClick={openOpNew} style={btn("#2563eb",undefined,{marginLeft:"auto",fontSize:13,padding:"7px 16px"})}>+ Добавить полис</button>}
              {isAdmin&&<>
                <button onClick={downloadOfficeTemplate} style={btn("#0891b2",undefined,{fontSize:13,padding:"7px 16px"})}>📋 Шаблон Excel</button>
                <button onClick={()=>importOfficeRef.current.click()} style={btn("#0f766e",undefined,{fontSize:13,padding:"7px 16px"})}>⬆ Импорт</button>
                <input ref={importOfficeRef} type="file" accept=".xlsx,.xls" onChange={handleImportOfficeFile} style={{display:"none"}}/>
                <span title={"Как заполнять шаблон:\n\nСтрока 1 — технические имена полей (не менять)\nСтрока 2 — названия колонок\nСтрока 3 — подсказки: синие = обязательные, серые = необязательные\nСтрока 4 — ПРИМЕР (жёлтая, не удалять)\nСтрока 5 и ниже — ваши данные\n\nФормат дат: ДД.ММ.ГГГГ (например: 15.01.2024)\nОплачено: TRUE или FALSE\nСпособ оплаты: cash / acba / ineco\nСрок (term): L — годовой, SH — краткосрочный"} style={{display:"inline-flex",alignItems:"center",justifyContent:"center",width:26,height:26,borderRadius:"50%",background:"#e0f2fe",color:"#0369a1",fontWeight:700,fontSize:13,cursor:"help",border:"1.5px solid #7dd3fc",userSelect:"none"}}>i</span>
              </>}
            </div>

            {/* Import preview panel */}
            {importPreview&&(()=>{
              const validRows=importPreview.rows.filter(r=>r._valid);
              const errorRows=importPreview.rows.filter(r=>!r._valid);
              return(
              <div style={{background:"#f8fafc",border:"2px solid #6366f1",borderRadius:10,padding:"16px 20px",marginBottom:12}}>
                <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:8,marginBottom:12}}>
                  <div style={{fontWeight:700,fontSize:15,color:"#1e293b"}}>Предпросмотр импорта — {fmtMonth(importPreview.month)}</div>
                  <div style={{display:"flex",gap:8,alignItems:"center",flexWrap:"wrap"}}>
                    <span style={{background:"#dcfce7",color:"#166534",borderRadius:20,padding:"3px 12px",fontSize:13,fontWeight:600}}>✓ Валидных: {validRows.length}</span>
                    {errorRows.length>0&&<span style={{background:"#fee2e2",color:"#991b1b",borderRadius:20,padding:"3px 12px",fontSize:13,fontWeight:600}}>✗ Ошибок: {errorRows.length}</span>}
                  </div>
                </div>
                <div style={{overflowX:"auto",maxHeight:360,overflowY:"auto",borderRadius:6,border:"1px solid #e2e8f0",marginBottom:12}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead><tr style={{background:"#e0e7ff",position:"sticky",top:0,zIndex:1}}>
                      {["Стр.","Лист","Статус","Страхователь","Дата","Компания","Сумма","№ полиса","Ошибки"].map(h=><th key={h} style={{...th,background:"#e0e7ff",color:"#3730a3",fontWeight:700,whiteSpace:"nowrap"}}>{h}</th>)}
                    </tr></thead>
                    <tbody>{importPreview.rows.map((r,i)=>{
                      const bg=r._valid?(i%2===0?"#f0fdf4":"#dcfce7"):(i%2===0?"#fff1f2":"#fee2e2");
                      return(
                      <tr key={r._id} style={{background:bg,borderBottom:"1px solid #e2e8f0"}}>
                        <td style={{...td,color:"#6b7280",textAlign:"center"}}>{r._rowNum}</td>
                        <td style={{...td,fontSize:11,color:"#6b7280",whiteSpace:"nowrap"}}>{r._sheetName}</td>
                        <td style={{...td,textAlign:"center"}}>{r._valid?<span style={{background:"#16a34a",color:"white",borderRadius:10,padding:"2px 8px",fontSize:11,fontWeight:600}}>✓ OK</span>:<span style={{background:"#dc2626",color:"white",borderRadius:10,padding:"2px 8px",fontSize:11,fontWeight:600}}>✗ Ошибка</span>}</td>
                        <td style={{...td,fontWeight:r._valid?600:400,color:r._valid?"#1e293b":"#991b1b"}}>{r.insuredName||"—"}</td>
                        <td style={{...td,whiteSpace:"nowrap"}}>{r.date||"—"}</td>
                        <td style={td}>{r.company||"—"}</td>
                        <td style={{...td,textAlign:"right"}}>{r.amount>0?r.amount.toLocaleString():"—"}</td>
                        <td style={{...td,fontSize:11,color:"#6b7280"}}>{r.policyNum||"—"}</td>
                        <td style={{...td,color:"#dc2626",fontSize:11}}>{r._errors.join("; ")||""}</td>
                      </tr>);
                    })}</tbody>
                  </table>
                </div>
                <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
                  {validRows.length>0
                    ?<button onClick={()=>confirmImport(validRows)} style={btn("#16a34a",undefined,{fontSize:13,padding:"7px 18px"})}>Импортировать валидные ({validRows.length})</button>
                    :<span style={{fontSize:13,color:"#dc2626",fontWeight:600}}>Нет строк без ошибок</span>
                  }
                  <button onClick={()=>setImportPreview(null)} style={btn("#6b7280",undefined,{fontSize:13,padding:"7px 18px"})}>Отмена</button>
                  {errorRows.length>0&&<span style={{fontSize:12,color:"#6b7280"}}>Строки с ошибками будут пропущены</span>}
                </div>
              </div>
              );
            })()}

            {/* Import confirmation modal */}
            {importPending&&(
              <div style={{background:"#fffbeb",border:"2px solid #f59e0b",borderRadius:10,padding:"16px 20px",marginBottom:12,display:"flex",flexDirection:"column",gap:10}}>
                <div style={{fontWeight:700,fontSize:14,color:"#92400e"}}>⚠ Данные за {fmtMonth(importPending.month)} уже существуют ({opCurrentMonth.length} зап.)</div>
                <div style={{fontSize:13,color:"#78350f"}}>Импортируется {importPending.pols.length} записей. Что сделать с существующими?</div>
                <div style={{display:"flex",gap:10,flexWrap:"wrap"}}>
                  <button onClick={()=>{saveOpMonth(importPending.pols);setImportPending(null);}} style={btn("#dc2626",undefined,{fontSize:13,padding:"7px 18px"})}>Заменить полностью</button>
                  <button onClick={()=>{saveOpMonth([...opCurrentMonth,...importPending.pols]);setImportPending(null);}} style={btn("#0f766e",undefined,{fontSize:13,padding:"7px 18px"})}>Добавить к существующим</button>
                  <button onClick={()=>setImportPending(null)} style={btn("#6b7280",undefined,{fontSize:13,padding:"7px 18px"})}>Отмена</button>
                </div>
              </div>
            )}

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
                        {!isViewOnly&&actBtn("✎","#f3f4f6","#374151",()=>openOpEdit(pol))}
                        {!isViewOnly&&actBtn("✓ Оплата","#16a34a","#fff",()=>openOpPay(pol))}
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
                          {!pol.paid&&!isViewOnly&&actBtn("✎","#f3f4f6","#374151",()=>openOpEdit(pol))}
                          {!pol.paid&&!isViewOnly&&actBtn("✓ Оплата","#16a34a","#fff",()=>openOpPay(pol))}
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
                          {!pol.paid&&!isViewOnly&&actBtn("✎","#f3f4f6","#374151",()=>openOpEdit(pol))}
                          {!pol.paid&&!isViewOnly&&actBtn("✓ Оплата","#16a34a","#fff",()=>openOpPay(pol))}
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
                      <div style={flbl}>№ полиса{isLocked&&opEditPol&&<span style={{marginLeft:6,fontSize:10,color:"#dc2626"}}>🔒</span>}</div>
                      <input value={opFD.policyNum||""} onChange={e=>setOpFD(p=>({...p,policyNum:e.target.value}))} placeholder="Номер полиса" disabled={isLocked&&!!opEditPol} style={{...finp,width:"100%",boxSizing:"border-box",...(isLocked&&opEditPol?{background:"#f1f5f9",color:"#94a3b8",cursor:"not-allowed"}:{})}}/>
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
                      <div style={flbl}>Страховая премия (AMD) *{isLocked&&opEditPol&&<span style={{marginLeft:6,fontSize:10,color:"#dc2626"}}>🔒</span>}</div>
                      <input type="number" value={opFD.amount||""} onChange={e=>setOpFD(p=>({...p,amount:e.target.value}))} placeholder="0" disabled={isLocked&&!!opEditPol} style={{...finp,width:"100%",boxSizing:"border-box",textAlign:"right",...(isLocked&&opEditPol?{background:"#f1f5f9",color:"#94a3b8",cursor:"not-allowed"}:{})}}/>
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

      {tab==="manager"&&(()=>{
        const cfg=managerConfig;
        const opResults=_computeOpR(agentData,cfg);
        const totMgr=opResults.reduce((s,r)=>s+r.mi,0);
        const totOp=opResults.reduce((s,r)=>s+r.oi,0);
        const totFix=opResults.reduce((s,r)=>s+r.fix,0);
        const totProfit=opResults.reduce((s,r)=>s+r.profit,0);
        const detailResult=mgrDetail?opResults.find(r=>r.uid===mgrDetail):null;
        const detailOp=mgrDetail?agentData.find(a=>a.uid===mgrDetail):null;
        const agentsNotOp=Object.entries(agentDir).filter(([id])=>!(cfg.operatorUids||[]).includes(id));
        return(
          <div>
            {/* Month nav */}
            <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:16,background:"#f1f5f9",borderRadius:8,padding:"10px 14px",flexWrap:"wrap"}}>
              <button onClick={()=>setSelMonth(prevMo(selMonth))} disabled={selMonth<=MIN_MONTH} style={{...btn("#fff","#374151",{border:"1px solid #d1d5db",fontSize:18,padding:"3px 10px"}),opacity:selMonth<=MIN_MONTH?0.4:1}}>‹</button>
              <span style={{fontWeight:700,fontSize:16,minWidth:160,textAlign:"center"}}>{fmtMonth(selMonth)}</span>
              <button onClick={()=>setSelMonth(nextMo(selMonth))} disabled={selMonth>=MAX_MONTH} style={{...btn("#fff","#374151",{border:"1px solid #d1d5db",fontSize:18,padding:"3px 10px"}),opacity:selMonth>=MAX_MONTH?0.4:1}}>›</button>
              <div style={{marginLeft:"auto",display:"flex",gap:8,alignItems:"center"}}>
                {isLocked&&<span style={{background:"#fef3c7",border:"1px solid #fcd34d",borderRadius:6,padding:"3px 8px",fontSize:11,fontWeight:700,color:"#92400e"}}>🔒 Закрыт</span>}
                {isAdmin&&(isLocked
                  ?<button onClick={unlockMonth} style={btn("#dc2626",undefined,{fontSize:11})}>🔒 Открыть месяц</button>
                  :<button onClick={lockMonth} style={btn("#92400e",undefined,{fontSize:11})}>🔓 Закрыть месяц</button>
                )}
                {opResults.length>0&&<button onClick={()=>exportManagerXlsx(agentData,cfg,agentDir,selMonth,exceptions)} style={btn("#16a34a",undefined,{fontSize:12})}>⬇ Excel</button>}
                <button onClick={()=>setShowMgrSettings(v=>!v)} style={{...btn(showMgrSettings?"#7c3aed":"#f3f4f6",showMgrSettings?"#fff":"#374151",{border:"1px solid #d1d5db"}),fontSize:12}}>⚙️ Настройки</button>
              </div>
            </div>
            {/* Settings panel */}
            {showMgrSettings&&(
              <div style={{border:"1px solid #d8b4fe",borderRadius:8,padding:16,marginBottom:16,background:"#faf5ff"}}>
                <h3 style={{margin:"0 0 14px",fontSize:15,color:"#6d28d9"}}>👔 Настройки менеджера</h3>
                {/* Operator selection */}
                <div style={{marginBottom:16}}>
                  <div style={{fontWeight:600,fontSize:13,marginBottom:8,color:"#374151"}}>Операторы в команде</div>
                  <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:8}}>
                    {(cfg.operatorUids||[]).map(uid=>{const a=agentDir[uid];const nm=a?(a.name+" "+a.surname).trim():uid;return(
                      <span key={uid} style={{display:"inline-flex",alignItems:"center",gap:5,background:"#ede9fe",border:"1px solid #c4b5fd",borderRadius:6,padding:"4px 10px",fontSize:12,color:"#5b21b6",fontWeight:600}}>
                        {nm}
                        <button onClick={()=>saveManagerConfig({...cfg,operatorUids:(cfg.operatorUids||[]).filter(x=>x!==uid)})} style={{background:"none",border:"none",cursor:"pointer",color:"#6b7280",fontSize:14,padding:0,lineHeight:1}}>×</button>
                      </span>
                    );})}
                    {(cfg.operatorUids||[]).length===0&&<span style={{color:"#9ca3af",fontSize:12}}>Нет операторов</span>}
                  </div>
                  <div style={{display:"flex",gap:8,alignItems:"center"}}>
                    <select value={mgrNewOp} onChange={e=>setMgrNewOp(e.target.value)} style={{...inp,padding:"4px 8px",fontSize:12,minWidth:220}}>
                      <option value="">Выберите агента...</option>
                      {agentsNotOp.map(([id,a])=><option key={id} value={id}>{a.name+" "+a.surname+" ("+(a.internalCode||id)+")"}</option>)}
                    </select>
                    <button onClick={()=>{if(!mgrNewOp)return;saveManagerConfig({...cfg,operatorUids:[...(cfg.operatorUids||[]),mgrNewOp]});setMgrNewOp("");}} disabled={!mgrNewOp} style={btn("#7c3aed")}>+ Добавить</button>
                  </div>
                </div>
                <div style={{borderTop:"1px solid #e5e7eb",paddingTop:14}}>
                  <div style={{fontWeight:600,fontSize:13,marginBottom:10,color:"#374151"}}>Ставки</div>
                  <MgrRatesPanel cfg={cfg} onSave={nc=>saveManagerConfig({...nc,operatorUids:cfg.operatorUids||[]})}/>
                </div>
              </div>
            )}
            {/* Summary cards */}
            {opResults.length>0&&(
              <div style={{display:"flex",gap:12,marginBottom:20,flexWrap:"wrap"}}>
                {[["Доход менеджера",totMgr,"#6d28d9","#f5f3ff","#ede9fe"],["Выплачено операторам",totOp,"#1d4ed8","#eff6ff","#dbeafe"],["Фикс. расходы",totFix,"#b45309","#fffbeb","#fde68a"],["Прибыль",totProfit,totProfit>=0?"#15803d":"#dc2626",totProfit>=0?"#f0fdf4":"#fff1f2",totProfit>=0?"#bbf7d0":"#fecaca"]].map(([l,v,col,bg,border])=>(
                  <div key={l} style={{background:bg,border:"1px solid "+border,borderRadius:10,padding:"12px 18px",minWidth:150}}>
                    <div style={{fontSize:11,color:"#6b7280",marginBottom:2}}>{l}</div>
                    <div style={{fontSize:17,fontWeight:700,color:col}}>{fmt(v)}</div>
                  </div>
                ))}
              </div>
            )}
            {/* No operators message */}
            {(cfg.operatorUids||[]).length===0&&(
              <div style={{padding:48,textAlign:"center",color:"#9ca3af",fontSize:14,border:"2px dashed #e5e7eb",borderRadius:8}}>
                <div style={{fontSize:32,marginBottom:8}}>👔</div>
                <div>Добавьте операторов в настройках</div>
              </div>
            )}
            {/* Operators table */}
            {opResults.length>0&&(
              <div style={{overflowX:"auto",borderRadius:8,border:"1px solid #e5e7eb",marginBottom:16}}>
                <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                  <thead><tr style={{background:"#1e293b",color:"#fff"}}>
                    {["Оператор","Код","Всего продаж","Зачётные","Ступень","Фикс","Доход мен.","Выплата опер.","Прибыль","",""].map(h=><th key={h} style={{...th,color:"#fff",background:"#1e293b",whiteSpace:"nowrap"}}>{h}</th>)}
                  </tr></thead>
                  <tbody>
                    {opResults.map((r,i)=>{
                      const a=agentDir[r.uid];const nm=a?(a.name+" "+a.surname).trim():r.uid;const ic=a&&a.internalCode||"";
                      return(
                        <tr key={r.uid} style={{background:i%2===0?"white":"#f8fafc",cursor:"pointer"}} onClick={()=>setMgrDetail(mgrDetail===r.uid?null:r.uid)}>
                          <td style={{...td,fontWeight:600,color:"#374151"}}>{nm}</td>
                          <td style={{...td,color:"#6366f1",fontSize:12,fontWeight:600}}>{ic||"—"}</td>
                          <td style={{...td,textAlign:"right"}}>{fmt(r.totalSales)}</td>
                          <td style={{...td,textAlign:"right",color:"#16a34a",fontWeight:600}}>{fmt(r.validSales)}</td>
                          <td style={{...td,textAlign:"center"}}><span style={{background:"#ede9fe",color:"#5b21b6",borderRadius:12,padding:"2px 8px",fontSize:12,fontWeight:700}}>{"С"+r.tier}</span></td>
                          <td style={{...td,textAlign:"right",color:"#d97706"}}>{fmt(r.fix)}</td>
                          <td style={{...td,textAlign:"right",color:"#7c3aed",fontWeight:600}}>{fmt(r.mi)}</td>
                          <td style={{...td,textAlign:"right",color:"#1d4ed8"}}>{fmt(r.oi)}</td>
                          <td style={{...td,textAlign:"right",fontWeight:700,color:r.profit>=0?"#16a34a":"#dc2626"}}>{fmt(r.profit)}</td>
                          <td style={{...td,color:"#6b7280",fontSize:11}}>{mgrDetail===r.uid?"▲":"▼"}</td>
                          <td style={td} onClick={e=>e.stopPropagation()}><button onClick={()=>exportSingleOpXlsx(r,cfg,agentDir,selMonth,exceptions)} style={btn("#16a34a",undefined,{fontSize:11,padding:"3px 8px"})} title="Экспорт оператора">⬇</button></td>
                        </tr>
                      );
                    })}
                    <tr style={{background:"#111827",color:"#fff"}}>
                      <td colSpan={4} style={{...td,fontWeight:700,color:"#fff",borderTop:"2px solid #374151"}}>ИТОГО</td>
                      <td style={{...td,borderTop:"2px solid #374151"}}></td>
                      <td style={{...td,textAlign:"right",color:"#fbbf24",fontWeight:700,borderTop:"2px solid #374151"}}>{fmt(totFix)}</td>
                      <td style={{...td,textAlign:"right",color:"#c4b5fd",fontWeight:700,borderTop:"2px solid #374151"}}>{fmt(totMgr)}</td>
                      <td style={{...td,textAlign:"right",color:"#93c5fd",fontWeight:700,borderTop:"2px solid #374151"}}>{fmt(totOp)}</td>
                      <td style={{...td,textAlign:"right",fontWeight:700,color:totProfit>=0?"#4ade80":"#f87171",borderTop:"2px solid #374151"}}>{fmt(totProfit)}</td>
                      <td style={{...td,borderTop:"2px solid #374151"}}></td>
                      <td style={{...td,borderTop:"2px solid #374151"}}></td>
                    </tr>
                  </tbody>
                </table>
              </div>
            )}
            {/* Detail view */}
            {detailResult&&detailOp&&(()=>{
              const thrs=cfg.tierThresholds||DEFAULT_MGR_RATES.tierThresholds;
              const tier=detailResult.tier;
              return(
                <div style={{border:"1px solid #d8b4fe",borderRadius:8,marginBottom:16,overflow:"hidden"}}>
                  <div style={{background:"#4c1d95",color:"#fff",padding:"12px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <span style={{fontWeight:700,fontSize:14}}>{getName(detailOp.uid)||detailOp.uid}</span>
                    <div style={{display:"flex",gap:12,flexWrap:"wrap",fontSize:12,opacity:.9}}>
                      <span>{"Ступень "+tier+" (зачётные ≥ "+fmt(thrs[tier-1])+")"}</span>
                      <span>{"Фикс: "+fmt(detailResult.fix)}</span>
                      <span>{"Прибыль: "+fmt(detailResult.profit)}</span>
                    </div>
                    <button onClick={()=>setMgrDetail(null)} style={{background:"none",border:"none",cursor:"pointer",color:"#fff",fontSize:18,padding:0}}>×</button>
                  </div>
                  {/* Зачётные */}
                  <div style={{overflowX:"auto",padding:0}}>
                    <div style={{background:"#166534",color:"#fff",padding:"6px 14px",fontSize:12,fontWeight:600}}>{"✓ Зачётные полисы ("+detailOp.policies.filter(p=>!p.exception).length+")"}</div>
                    <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                      <thead><tr style={{background:"#f0fdf4"}}>{["№ полиса","Компания","Страхователь","Марка","Рег.номер","Регион","БМ","Мощ-ть","Срок","Сумма","% Мен.","Доход мен.","% Опер.","Выплата опер."].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
                      <tbody>{detailOp.policies.filter(p=>!p.exception).map((p,i)=>{
                        const mr=getMgrPolicyRate(p,cfg);const or=getOpPolicyRate(p,tier,cfg);
                        const mc=Math.round(p.amount*mr/100);const oc=Math.round(p.amount*or/100);
                        return(
                          <tr key={i} style={{background:i%2===0?"white":"#f0fdf4",borderBottom:"1px solid #d1fae5"}}>
                            <td style={{...td,fontSize:11}}>{p.policyNum||"—"}</td>
                            <td style={td}>{p.company}</td>
                            <td style={td}>{p.insuredName}</td>
                            <td style={{...td,fontSize:11,color:"#6b7280"}}>{p.car||"—"}</td>
                            <td style={{...td,fontSize:11,color:"#6b7280"}}>{p.carPlate||"—"}</td>
                            <td style={{...td,textAlign:"center"}}>{p.region||"—"}</td>
                            <td style={{...td,textAlign:"center"}}>{p.bm||"—"}</td>
                            <td style={{...td,textAlign:"center"}}>{p.power||"—"}</td>
                            <td style={{...td,textAlign:"center"}}>{p.term||"—"}</td>
                            <td style={{...td,textAlign:"right"}}>{fmt(p.amount)}</td>
                            <td style={{...td,textAlign:"center",color:"#7c3aed"}}>{mr+"%"}</td>
                            <td style={{...td,textAlign:"right",color:"#7c3aed",fontWeight:600}}>{fmt(mc)}</td>
                            <td style={{...td,textAlign:"center",color:"#1d4ed8"}}>{or+"%"}</td>
                            <td style={{...td,textAlign:"right",color:"#1d4ed8"}}>{fmt(oc)}</td>
                          </tr>
                        );
                      })}</tbody>
                      <tfoot><tr style={{background:"#d1fae5",fontWeight:700}}>
                        <td colSpan={9} style={{...td,borderTop:"2px solid #86efac"}}>ИТОГО</td>
                        <td style={{...td,textAlign:"right",borderTop:"2px solid #86efac"}}>{fmt(detailOp.policies.filter(p=>!p.exception).reduce((s,p)=>s+p.amount,0))}</td>
                        <td style={{...td,borderTop:"2px solid #86efac"}}></td>
                        <td style={{...td,textAlign:"right",color:"#7c3aed",borderTop:"2px solid #86efac"}}>{fmt(detailResult.mi)}</td>
                        <td style={{...td,borderTop:"2px solid #86efac"}}></td>
                        <td style={{...td,textAlign:"right",color:"#1d4ed8",borderTop:"2px solid #86efac"}}>{fmt(detailResult.oi)}</td>
                      </tr></tfoot>
                    </table>
                  </div>
                  {/* Незачётные */}
                  {detailOp.policies.filter(p=>p.exception).length>0&&(
                    <div style={{overflowX:"auto"}}>
                      <div style={{background:"#991b1b",color:"#fff",padding:"6px 14px",fontSize:12,fontWeight:600}}>{"✕ Незачётные полисы ("+detailOp.policies.filter(p=>p.exception).length+")"}</div>
                      <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                        <thead><tr style={{background:"#fee2e2"}}>{["№ полиса","Компания","Страхователь","Марка","Рег.номер","Регион","БМ","Мощ-ть","Срок","Сумма","Причина исключения"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
                        <tbody>{detailOp.policies.filter(p=>p.exception).map((p,i)=>(
                          <tr key={i} style={{background:i%2===0?"white":"#fff7ed",borderBottom:"1px solid #fed7aa"}}>
                            <td style={{...td,fontSize:11}}>{p.policyNum||"—"}</td>
                            <td style={td}>{p.company}</td>
                            <td style={td}>{p.insuredName}</td>
                            <td style={{...td,fontSize:11}}>{p.car||"—"}</td>
                            <td style={{...td,fontSize:11}}>{p.carPlate||"—"}</td>
                            <td style={{...td,textAlign:"center"}}>{p.region||"—"}</td>
                            <td style={{...td,textAlign:"center",color:p.bm?"#dc2626":"#6b7280",fontWeight:p.bm?600:400}}>{p.bm||"—"}</td>
                            <td style={{...td,textAlign:"center",color:p.power?"#dc2626":"#6b7280",fontWeight:p.power?600:400}}>{p.power||"—"}</td>
                            <td style={{...td,textAlign:"center"}}>{p.term||"—"}</td>
                            <td style={{...td,textAlign:"right"}}>{fmt(p.amount)}</td>
                            <td style={{...td,fontSize:11,color:"#dc2626"}}>{excReason(p,exceptions,detailOp.uid)}</td>
                          </tr>
                        ))}</tbody>
                      </table>
                    </div>
                  )}
                </div>
              );
            })()}
          </div>
        );
      })()}

      {tab==="income"&&(()=>{
        const rows=officeExpenses[selMonth]||[];
        const osagoGross=agentData.reduce((s,a)=>s+a.totalOffice,0);
        const osagoAgentPay=agentData.reduce((s,a)=>s+a.totalAgent,0);
        const volGross=effVol.reduce((s,v)=>s+v.officeComm,0);
        const volAgentPay=effVol.reduce((s,v)=>s+v.agentComm,0);
        const salesGross=osagoGross+volGross;
        const agentPayTotal=osagoAgentPay+volAgentPay;
        const salesNet=salesGross-agentPayTotal;
        const opR=_computeOpR(agentData,managerConfig);
        const mgrTeamSales=opR.reduce((s,r)=>s+r.validSales,0);
        const totMgr=opR.reduce((s,r)=>s+r.mi,0);
        const totOp=opR.reduce((s,r)=>s+r.oi,0);
        const totFix=opR.reduce((s,r)=>s+r.fix,0);
        const totMgrPersonal=totMgr-totOp-totFix;
        const totalExpenses=rows.reduce((s,r)=>s+(Number(r.amount)||0),0);
        const netProfit=salesNet-totalExpenses;
        const updRow=(id,key,val)=>{const nr=rows.map(r=>r.id===id?{...r,[key]:val}:r);saveOfficeExpenses({...officeExpenses,[selMonth]:nr});};
        const delRow=id=>saveOfficeExpenses({...officeExpenses,[selMonth]:rows.filter(r=>r.id!==id)});
        const addDynRow=()=>{if(!expNewName.trim())return;saveOfficeExpenses({...officeExpenses,[selMonth]:[...rows,{id:"dyn"+Date.now()+Math.random().toString(36).slice(2),cat:"Доп.",name:expNewName.trim(),amount:0,type:"dynamic"}]});setExpNewName("");};
        const STATIC_CATS=["Зарплаты","Коммунальные","Связь","Налоги","Прочее"];
        const dynRows=rows.filter(r=>r.type==="dynamic");
        const catColors={"Зарплаты":"#eff6ff","Коммунальные":"#f0fdf4","Связь":"#fdf4ff","Налоги":"#fff7ed","Прочее":"#f8fafc"};
        const catBorders={"Зарплаты":"#bfdbfe","Коммунальные":"#bbf7d0","Связь":"#e9d5ff","Налоги":"#fed7aa","Прочее":"#e2e8f0"};
        const RowEl=({r})=>(
          <div style={{display:"flex",alignItems:"center",gap:8,padding:"5px 14px",borderBottom:"1px solid #f9fafb"}}>
            <input value={r.name} onChange={e=>updRow(r.id,"name",e.target.value)} style={{...inp,flex:1,padding:"3px 8px",fontSize:13,minWidth:0}}/>
            <input type="number" value={r.amount||""} onChange={e=>updRow(r.id,"amount",Number(e.target.value)||0)} placeholder="0" style={{...inp,width:120,padding:"3px 8px",fontSize:13,textAlign:"right"}}/>
            <button onClick={()=>delRow(r.id)} style={{background:"none",border:"none",cursor:"pointer",color:"#9ca3af",fontSize:18,padding:"0 4px",lineHeight:1}}>×</button>
          </div>
        );
        return(
          <div style={{maxWidth:960}}>
            <div style={{display:"flex",justifyContent:"flex-end",marginBottom:10}}>
              <button onClick={exportIncomeExcel} style={btn("#16a34a",undefined,{fontSize:12})}>⬇ Excel</button>
            </div>
            {/* Сводка */}
            <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:20}}>
              {[
                ["Валовой доход",salesGross,"#1d4ed8","#eff6ff","#bfdbfe"],
                ["Выплачено агентам",agentPayTotal,"#7c3aed","#f5f3ff","#ede9fe"],
                ["Расходы офиса",totalExpenses,"#b45309","#fffbeb","#fde68a"],
                ["Чистая прибыль",netProfit,netProfit>=0?"#15803d":"#dc2626",netProfit>=0?"#f0fdf4":"#fff1f2",netProfit>=0?"#bbf7d0":"#fecaca"],
              ].map(([l,v,col,bg,border])=>(
                <div key={l} style={{background:bg,border:"1px solid "+border,borderRadius:10,padding:"12px 18px",minWidth:160}}>
                  <div style={{fontSize:11,color:"#6b7280",marginBottom:2}}>{l}</div>
                  <div style={{fontSize:18,fontWeight:700,color:col}}>{fmt(v)}</div>
                </div>
              ))}
            </div>
            {/* Доходы от продаж */}
            <div style={{background:"white",border:"1px solid #e5e7eb",borderRadius:8,marginBottom:14,overflow:"hidden"}}>
              <div style={{background:"#1d4ed8",color:"white",padding:"8px 14px",fontWeight:600,fontSize:13}}>💰 Доходы от продаж — {fmtMonth(selMonth)}</div>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr>
                  <th style={{...th,textAlign:"left",paddingLeft:14}}>Источник</th>
                  <th style={th}>Брутто</th>
                  <th style={th}>Выплачено агентам</th>
                  <th style={th}>Нетто</th>
                </tr></thead>
                <tbody>
                  <tr>
                    <td style={{...td,paddingLeft:14}}>ОСАГО</td>
                    <td style={{...td,textAlign:"right"}}>{fmt(osagoGross)}</td>
                    <td style={{...td,textAlign:"right",color:"#dc2626"}}>−{fmt(osagoAgentPay)}</td>
                    <td style={{...td,textAlign:"right",fontWeight:600}}>{fmt(osagoGross-osagoAgentPay)}</td>
                  </tr>
                  <tr>
                    <td style={{...td,paddingLeft:14}}>Добровольные</td>
                    <td style={{...td,textAlign:"right"}}>{fmt(volGross)}</td>
                    <td style={{...td,textAlign:"right",color:"#dc2626"}}>−{fmt(volAgentPay)}</td>
                    <td style={{...td,textAlign:"right",fontWeight:600}}>{fmt(volGross-volAgentPay)}</td>
                  </tr>
                  <tr style={{background:"#f8fafc"}}>
                    <td style={{...td,paddingLeft:14,fontWeight:700}}>Итого</td>
                    <td style={{...td,textAlign:"right",fontWeight:700}}>{fmt(salesGross)}</td>
                    <td style={{...td,textAlign:"right",fontWeight:700,color:"#dc2626"}}>−{fmt(agentPayTotal)}</td>
                    <td style={{...td,textAlign:"right",fontWeight:700,color:"#15803d"}}>{fmt(salesNet)}</td>
                  </tr>
                </tbody>
              </table>
            </div>
            {/* Команда менеджера */}
            {opR.length>0&&(
              <div style={{background:"white",border:"1px solid #e5e7eb",borderRadius:8,marginBottom:14,overflow:"hidden"}}>
                <div style={{background:"#7c3aed",color:"white",padding:"8px 14px",fontWeight:600,fontSize:13}}>👔 Команда менеджера</div>
                <div style={{padding:"12px 14px"}}>
                  <div style={{display:"flex",gap:12,flexWrap:"wrap",marginBottom:10}}>
                    {[
                      ["Объём продаж команды",fmt(mgrTeamSales),"#374151"],
                      ["Бюджет менеджера (расход офиса)",fmt(totMgr),"#7c3aed"],
                      ["Выплачено операторам",fmt(totOp+totFix),"#1d4ed8"],
                      ["Доход менеджера лично",fmt(totMgrPersonal),totMgrPersonal>=0?"#15803d":"#dc2626"],
                    ].map(([l,v,c])=>(
                      <div key={l} style={{background:"#faf5ff",border:"1px solid #ede9fe",borderRadius:8,padding:"8px 14px",minWidth:155}}>
                        <div style={{fontSize:11,color:"#6b7280",marginBottom:2}}>{l}</div>
                        <div style={{fontSize:15,fontWeight:700,color:c}}>{v}</div>
                      </div>
                    ))}
                  </div>
                  <div style={{fontSize:12,color:"#6b7280",padding:"6px 10px",background:"#faf5ff",borderRadius:6,border:"1px solid #ede9fe"}}>
                    ℹ️ Бюджет менеджера входит в строку «Выплачено агентам» выше — операторы учтены в общем расчёте
                  </div>
                </div>
              </div>
            )}
            {/* Расходы */}
            <div style={{background:"white",border:"1px solid #e5e7eb",borderRadius:8,marginBottom:14,overflow:"hidden"}}>
              <div style={{background:"#b45309",color:"white",padding:"8px 14px",fontWeight:600,fontSize:13,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                <span>📋 Расходы офиса — {fmtMonth(selMonth)}</span>
                <span style={{fontSize:12,fontWeight:400,opacity:0.85}}>Итого: {fmt(totalExpenses)}</span>
              </div>
              {rows.length===0&&<div style={{padding:20,textAlign:"center",color:"#9ca3af",fontSize:13}}>Загрузка данных...</div>}
              {STATIC_CATS.map(cat=>{
                const catRows=rows.filter(r=>r.cat===cat&&r.type==="static");
                if(!catRows.length)return null;
                const catTotal=catRows.reduce((s,r)=>s+(Number(r.amount)||0),0);
                return(
                  <div key={cat} style={{borderBottom:"1px solid #f3f4f6"}}>
                    <div style={{background:catColors[cat]||"#f9fafb",borderBottom:"1px solid "+(catBorders[cat]||"#e5e7eb"),padding:"5px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <span style={{fontWeight:600,fontSize:12,color:"#374151"}}>{cat}</span>
                      <span style={{fontSize:12,color:"#6b7280",fontWeight:600}}>{fmt(catTotal)}</span>
                    </div>
                    {catRows.map(r=><RowEl key={r.id} r={r}/>)}
                  </div>
                );
              })}
              {dynRows.length>0&&(
                <div style={{borderBottom:"1px solid #f3f4f6"}}>
                  <div style={{background:"#f9fafb",borderBottom:"1px solid #e5e7eb",padding:"5px 14px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                    <span style={{fontWeight:600,fontSize:12,color:"#374151"}}>Дополнительно</span>
                    <span style={{fontSize:12,color:"#6b7280",fontWeight:600}}>{fmt(dynRows.reduce((s,r)=>s+(Number(r.amount)||0),0))}</span>
                  </div>
                  {dynRows.map(r=><RowEl key={r.id} r={r}/>)}
                </div>
              )}
              <div style={{padding:"10px 14px",display:"flex",gap:8,alignItems:"center",borderTop:"1px solid #f3f4f6"}}>
                <input value={expNewName} onChange={e=>setExpNewName(e.target.value)} placeholder="Название доп. расхода..." style={{...inp,flex:1,padding:"5px 8px",fontSize:13}} onKeyDown={e=>e.key==="Enter"&&addDynRow()}/>
                <button onClick={addDynRow} style={btn("#374151",undefined,{fontSize:13,padding:"5px 14px"})}>+ Добавить</button>
              </div>
            </div>
            {/* Итог */}
            <div style={{background:"white",border:"1px solid #e5e7eb",borderRadius:8,overflow:"hidden"}}>
              <div style={{background:"#111827",color:"white",padding:"8px 14px",fontWeight:600,fontSize:13}}>📊 Итоговый расчёт</div>
              <div style={{padding:"14px"}}>
                {[
                  ["Нетто от продаж",salesNet,"#15803d"],
                  ["Расходы офиса",-totalExpenses,"#dc2626"],
                ].map(([l,v,c])=>(
                  <div key={l} style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #f3f4f6",fontSize:14}}>
                    <span style={{color:"#374151"}}>{l}</span>
                    <span style={{fontWeight:600,color:c}}>{v>=0?"+":""}{fmt(v)}</span>
                  </div>
                ))}
                <div style={{display:"flex",justifyContent:"space-between",padding:"10px 0 0",fontSize:16}}>
                  <span style={{fontWeight:700}}>Чистая прибыль</span>
                  <span style={{fontWeight:800,color:netProfit>=0?"#15803d":"#dc2626",fontSize:18}}>{fmt(netProfit)}</span>
                </div>
              </div>
            </div>
          </div>
        );
      })()}

      {tab==="search"&&(
        <div style={{maxWidth:1100}}>
          <div style={{background:"white",borderRadius:10,padding:20,marginBottom:16,border:"1px solid #e2e8f0"}}>
            <div style={{fontWeight:700,fontSize:15,marginBottom:14,color:"#1e293b"}}>🔍 Глобальный поиск по всем месяцам</div>
            <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
              <input value={searchQuery} onChange={e=>setSearchQuery(e.target.value)} onKeyDown={e=>e.key==="Enter"&&runSearch()} placeholder="Имя, рег. номер, номер полиса, телефон..." style={{...inp,flex:1,minWidth:260,padding:"9px 14px",fontSize:14}} autoFocus/>
              <button onClick={runSearch} disabled={searchLoading||searchQuery.trim().length<2} style={{...btn("#1d4ed8"),padding:"9px 22px",fontSize:14,opacity:searchQuery.trim().length<2?0.5:1}}>{searchLoading?"Поиск...":"Найти"}</button>
              {searchResults!==null&&<button onClick={()=>{setSearchResults(null);setSearchQuery("");}} style={btn("#64748b",undefined,{fontSize:12})}>✕ Очистить</button>}
            </div>
            <div style={{fontSize:12,color:"#94a3b8",marginTop:8}}>Поиск по: страхователь, рег. номер, номер полиса, телефон — агентские и офисные продажи за все месяцы</div>
          </div>

          {searchLoading&&<div style={{textAlign:"center",padding:40,color:"#64748b",fontSize:14}}>⏳ Загрузка данных...</div>}

          {searchResults!==null&&!searchLoading&&(
            searchResults.length===0
              ?<div style={{textAlign:"center",padding:40,color:"#94a3b8",fontSize:14}}>Ничего не найдено</div>
              :<div style={{background:"white",borderRadius:10,border:"1px solid #e2e8f0",overflow:"hidden"}}>
                <div style={{padding:"10px 16px",background:"#f8fafc",borderBottom:"1px solid #e2e8f0",fontSize:13,color:"#64748b",fontWeight:600}}>
                  Найдено: {searchResults.length} {searchResults.length===1?"запись":searchResults.length<5?"записи":"записей"}
                </div>
                <div style={{overflowX:"auto"}}>
                  <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
                    <thead><tr>
                      <th style={th}>Месяц</th>
                      <th style={th}>Тип</th>
                      <th style={{...th,textAlign:"left"}}>Страхователь</th>
                      <th style={th}>Телефон</th>
                      <th style={th}>Рег. номер</th>
                      <th style={th}>№ полиса</th>
                      <th style={th}>Компания</th>
                      <th style={th}>Агент</th>
                      <th style={th}>Сумма</th>
                      <th style={th}>Оплата</th>
                      <th style={th}></th>
                    </tr></thead>
                    <tbody>
                      {searchResults.map((p,i)=>{
                        const agName=p.agentUid?(getName(p.agentUid)||p.agentUid):"—";
                        const isPaid=p._source==="office"?p.paid:null;
                        return(
                          <tr key={i} style={{background:i%2===0?"white":"#f8fafc",cursor:"pointer"}} onClick={()=>setSearchViewPol(p)}>
                            <td style={{...td,whiteSpace:"nowrap",fontWeight:600,color:"#1d4ed8"}}>{fmtMonth(p._monthKey)}</td>
                            <td style={{...td,whiteSpace:"nowrap"}}>
                              <span style={{background:p._source==="office"?"#ede9fe":"#dbeafe",color:p._source==="office"?"#6d28d9":"#1d4ed8",borderRadius:5,padding:"2px 6px",fontSize:11,fontWeight:600}}>
                                {p._source==="office"?"🏢 Офис":"💰 Агент"}
                              </span>
                            </td>
                            <td style={{...td,maxWidth:160,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontWeight:600}}>{p.insuredName||"—"}</td>
                            <td style={{...td,whiteSpace:"nowrap",fontFamily:"monospace",fontSize:11}}>{p.phone||"—"}</td>
                            <td style={{...td,whiteSpace:"nowrap",fontFamily:"monospace"}}>{p.carPlate||"—"}</td>
                            <td style={{...td,whiteSpace:"nowrap",fontFamily:"monospace",fontSize:11}}>{p.policyNum||"—"}</td>
                            <td style={td}>{p.company||"—"}</td>
                            <td style={{...td,fontSize:11,maxWidth:120,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{agName}</td>
                            <td style={{...td,textAlign:"right",whiteSpace:"nowrap",fontWeight:600}}>{p.amount?fmt(p.amount):"—"}</td>
                            <td style={{...td,whiteSpace:"nowrap"}}>
                              {isPaid===true&&<span style={{color:"#15803d",fontWeight:600}}>✓ {fmt(p.paidAmount||0)}</span>}
                              {isPaid===false&&<span style={{color:"#dc2626",fontWeight:600}}>✗ Не опл.</span>}
                              {isPaid===null&&<span style={{color:"#94a3b8"}}>—</span>}
                            </td>
                            <td style={{...td,whiteSpace:"nowrap"}}>
                              <button onClick={e=>{e.stopPropagation();setSearchViewPol(p);}} style={btn("#0f172a",undefined,{fontSize:11,padding:"3px 10px"})}>👁 Открыть</button>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
          )}
        </div>
      )}

      {searchViewPol&&(
        <div style={{position:"fixed",inset:0,background:"rgba(15,23,42,0.65)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16}} onClick={e=>{if(e.target===e.currentTarget)setSearchViewPol(null);}}>
          <div style={{background:"#f1f5f9",borderRadius:16,width:"100%",maxWidth:620,maxHeight:"92vh",overflowY:"auto",boxShadow:"0 32px 80px rgba(0,0,0,0.35)"}}>

            {/* Заголовок */}
            <div style={{background:"#1e293b",borderRadius:"16px 16px 0 0",padding:"18px 24px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div>
                <div style={{fontWeight:700,fontSize:17,color:"white"}}>{searchViewPol.insuredName||"—"}</div>
                <div style={{fontSize:12,color:"#94a3b8",marginTop:3}}>
                  {fmtMonth(searchViewPol._monthKey)}&nbsp;·&nbsp;
                  {searchViewPol._source==="office"?"🏢 Продажи офиса":"💰 Агентские продажи"}
                  {searchViewPol.company&&<>&nbsp;·&nbsp;{searchViewPol.company}</>}
                </div>
              </div>
              <button onClick={()=>setSearchViewPol(null)} style={{background:"rgba(255,255,255,0.1)",border:"none",cursor:"pointer",fontSize:20,color:"white",borderRadius:8,width:34,height:34,display:"flex",alignItems:"center",justifyContent:"center"}}>×</button>
            </div>

            {(()=>{
              const p=searchViewPol;
              const isOsago=p.polType!=="voluntary";
              const agName=p.agentUid?(getName(p.agentUid)||p.agentUid):"—";
              const fmtTerm=t=>t==="L"?"L (от 3 месяцев)":t==="SH"?"SH (краткосрочный)":t||"—";
              const fmtPay2=t=>t==="cash"?"Наличные":t==="acba"?"ACBA":t==="ineco"?"Ineco":t||"—";
              const v=x=>x!=null&&String(x).trim()?String(x).trim():"—";

              const Section=({title,color,children})=>(
                <div style={{margin:"14px 16px 0"}}>
                  <div style={{fontSize:11,fontWeight:700,color:color||"#64748b",textTransform:"uppercase",letterSpacing:.8,marginBottom:8,paddingLeft:4}}>{title}</div>
                  <div style={{background:"white",borderRadius:10,boxShadow:"0 1px 4px rgba(0,0,0,0.08)",overflow:"hidden"}}>{children}</div>
                </div>
              );
              const Row=({label,value,mono,bold,color})=>(
                <div style={{display:"flex",alignItems:"center",padding:"9px 14px",borderBottom:"1px solid #f1f5f9"}}>
                  <div style={{width:155,flexShrink:0,fontSize:12,color:"#64748b",fontWeight:600}}>{label}</div>
                  <div style={{fontSize:13,color:color||(value==="—"?"#cbd5e1":"#1e293b"),fontWeight:bold?700:500,fontFamily:mono?"monospace":"inherit"}}>{value}</div>
                </div>
              );

              return(
                <div style={{paddingBottom:20}}>

                  <Section title="Страхователь" color="#1d4ed8">
                    <Row label="ФИО" value={v(p.insuredName)} bold/>
                    <Row label="Телефон" value={v(p.phone)} mono/>
                    <Row label="Оператор / Агент" value={agName}/>
                    <Row label="Комментарий" value={v(p.comment)}/>
                  </Section>

                  <Section title="Полис" color="#6d28d9">
                    <Row label="Тип" value={isOsago?"🚗 ОСАГО":"🛡 Добровольный"}/>
                    {!isOsago&&<Row label="Продукт" value={v(p.productName)}/>}
                    <Row label="Компания" value={v(p.company)}/>
                    <Row label="№ полиса" value={v(p.policyNum)} mono bold/>
                    <Row label="Дата составления" value={v(p.date)}/>
                    {isOsago&&<>
                      <Row label="Срок" value={p.term?fmtTerm(p.term):"—"}/>
                      <Row label="Дата вступления в силу" value={v(p.dateStart||p.startDateFmt)}/>
                      <Row label="Дата окончания" value={v(p.dateEnd||p.endDateFmt)}/>
                      <Row label="Статус полиса" value={v(p.polStatus)}/>
                    </>}
                  </Section>

                  {isOsago&&(
                    <Section title="Транспортное средство" color="#0f766e">
                      <Row label="Марка / Модель" value={v(p.car)}/>
                      <Row label="Рег. номер" value={v(p.carPlate)} mono/>
                      <Row label="Регион" value={v(p.region)}/>
                      <Row label="КБМ (БМ)" value={v(p.bm)}/>
                      <Row label="Мощность" value={p.power?p.power+" л.с.":"—"}/>
                    </Section>
                  )}

                  <Section title="Финансы" color="#15803d">
                    <Row label="Страховая премия" value={p.amount?fmt(p.amount):"—"} bold color={p.amount?"#1d4ed8":undefined}/>
                    <Row label="Скидка" value={p.discount&&Number(p.discount)>0?fmt(p.discount):"—"}/>
                    <Row label="К оплате" value={p.amount!=null?fmt((Number(p.amount)||0)-(Number(p.discount)||0)):"—"} bold/>
                    {p._source==="office"
                      ?p.paid
                        ?<>
                          <Row label="Статус оплаты" value="✓ Оплачено" bold color="#15803d"/>
                          <Row label="Сумма оплаты" value={p.paidAmount?fmt(p.paidAmount):"—"} bold color="#15803d"/>
                          <Row label="Дата оплаты" value={v(p.paidDate)}/>
                          <Row label="Способ оплаты" value={fmtPay2(p.paymentType)}/>
                        </>
                        :<Row label="Статус оплаты" value="✗ Не оплачено" bold color="#dc2626"/>
                      :<Row label="Статус оплаты" value="— (агентский полис)" color="#94a3b8"/>
                    }
                  </Section>

                </div>
              );
            })()}
          </div>
        </div>
      )}

    </div>
  );
}