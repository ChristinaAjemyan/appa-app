// Шаблон справочника агентов.
// Скопируй этот файл в agents_seed.js и заполни реальными данными.
// agents_seed.js добавлен в .gitignore и не попадает в репозиторий.
//
// Формат строки: [внутр_код, имя, фамилия, Nairi, Ingo, Sil, Rego, Liga, Armenia]
const d=[
  // ["768-01","Имя","Фамилия","1606-01","","","","",""],
];
export const SEED_AGENTS=Object.fromEntries(d.map(([ic,nm,sr,n,i,s,r,l,a])=>[`ag-${ic}`,{name:nm,surname:sr,internalCode:ic,codes:{Nairi:n,Ingo:i,Sil:s,Rego:r,Liga:l,Armenia:a}}]));
