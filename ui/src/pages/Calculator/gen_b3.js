const fs=require('fs');
let c=fs.readFileSync('calculator.jsx','utf8');
const q=n=>String.fromCodePoint(parseInt(n,16));
const a=q('0561'),e=q('0565'),i=q('056B'),o=q('0578'),n=q('0576'),l=q('056C'),
s=q('057D'),v=q('057E'),t=q('057F'),r=q('0580'),ts2=q('0581'),uw=q('0582'),
k=q('056F'),h=q('0570'),m=q('0574'),d=q('0564'),g=q('0563'),p=q('057A'),
zh=q('056A'),ch=q('0579'),y=q('0575'),tho=q('0569'),keh=q('0584'),
f=q('0586'),kh=q('056D'),ts=q('056E'),sch=q('0568');
const A=q('0531'),B=q('0532'),G=q('0533'),L=q('053C'),K=q('053F'),
M=q('0544'),N=q('0546'),P=q('054A'),T=q('054F'),S=q('054D'),Z=q('0536'),
O=q('0555'),Y=q('0538'),V=q('054E');
const pahpanel=P+a+h+p+a+n+e+l;
const amiss=a+m+i+s+sch;
const total=Y+n+d+h+a+n+o+uw+r;
const mizhnd=M+i+zh+n+d+o+r+d+a+v+zh+a+r;
const kvyp=V+zh+a+r+m+a+n+' '+e+n+t+a+k+a;
const podpis=S+t+o+r+a+g+r+uw+tho+y+uw+n;
const ima=G+o+r+ts+a+k+a+l+i+' '+a+n+o+uw+n;
const nkod=N+e+r+keh+i+n+' '+k+o+d;
const polcnt=P+a+l+i+s+n+e+r+i+' '+k+a+n+a+k;
const volp=O+f+i+s+i+n+' '+v+zh+a+r+v+e+l+i+keh;
const prim=L+r+a+ts2+uw+ts2+i+ch;
const zbros=Z+r+o+y+a+ts2+n+e+l;
const polis=P+o+l+i+s;
const bac=B+a+ts2+a+r+o+uw+tho+y+o+uw+n+n+e+r;
// Fix #3
c=c.replaceAll('\u054A'+'ahpanel amese',pahpanel+' '+amiss);
// Сбросить
c=c.replaceAll('Сбросить',zbros);
// Итого (avoid Итоговый)
c=c.replaceAll('"Итого"','"'+total+'"');
c=c.replaceAll('>Итого<','>'+total+'<');
c=c.replaceAll('Итого: ',total+': ');
c=c.replaceAll("'Итого',","'"+total+"',");
// Начислено
c=c.replaceAll('Начислено (ОСАГО)',mizhnd+' (ОСАГО)');
c=c.replaceAll('Начислено (AMD)',mizhnd+' (AMD)');
c=c.replaceAll('"Начислено"','"'+mizhnd+'"');
c=c.replaceAll('>Начислено<','>'+mizhnd+'<');
c=c.replaceAll("Начислено',",mizhnd+"',");
// К выплате
c=c.replaceAll('К выплате (AMD)',kvyp+' (AMD)');
c=c.replaceAll('"К выплате"','"'+kvyp+'"');
c=c.replaceAll('>К выплате<','>'+kvyp+'<');
c=c.replaceAll("К выплате',",kvyp+"',");
// Подпись
c=c.replaceAll('"Подпись"','"'+podpis+'"');
c=c.replaceAll('>Подпись<','>'+podpis+'<');
c=c.replaceAll('Подпись",',podpis+'",');
c=c.replaceAll("Подпись']",podpis+"']");
// Имя агента
c=c.replaceAll('Имя агента',ima);
// 768-код
c=c.replaceAll('768-код",',nkod+'",');
c=c.replaceAll('"768-код"','"'+nkod+'"');
c=c.replaceAll('>768-код<','>'+nkod+'<');
// Кол-во полисов
c=c.replaceAll('Кол-во полисов',polcnt);
// К оплате офису
c=c.replaceAll('К оплате офису (доброволь.)',volp);
// Примечание
c=c.replaceAll('"Примечание"','"'+prim+'"');
c=c.replaceAll('>Примечание<','>'+prim+'<');
c=c.replaceAll('Примечание",',prim+'",');
c=c.replaceAll('Примечание"]',prim+'"]');
// Исключения
c=c.replaceAll('"'+'\u26A0'+' '+'\u0418\u0441\u043A\u043B\u044E\u0447\u0435\u043D\u0438\u044F'+' \u2014 "','"'+'\u26A0'+' '+bac+' \u2014 "');
// полисов
c=c.replaceAll('" полисов"','" '+polis+'"');
c=c.replaceAll('полисов"',polis+'"');
c=c.replaceAll('" полисов ','" '+polis+' ');
fs.writeFileSync('calculator.jsx',c,'utf8');
console.log('done');
