const {JSDOM}=require('jsdom');const fs=require('fs');const util=require('util');const crypto=require('crypto');const ExcelJS=require('exceljs');
const SITO=process.argv[2]||'../cantina/sito';
let html=fs.readFileSync(SITO+'/index.html','utf8').replace('<script src="exceljs.min.js"></script>','');
const VINO=process.argv[3]||'/mnt/user-data/uploads/Downloads/Vino.xlsx';
const server={files:{'Vino.xlsx':{sha:'s0',content:fs.readFileSync(VINO).toString('base64')}},puts:0};
function mockFetch(url,opt){opt=opt||{};const m=opt.method||'GET';const auth=opt.headers&&opt.headers.Authorization;
 const res=(st,b)=>Promise.resolve({status:st,ok:st>=200&&st<300,json:async()=>b,headers:{get:()=>null}});
 if(auth!=='Bearer good')return res(401,{});
 const rest=url.replace('https://api.github.com/repos/','');if(!rest.startsWith('me/data'))return res(404,{});
 const sub=rest.slice('me/data'.length);if(sub==='')return res(200,{});
 const f=decodeURIComponent(sub.replace('/contents/',''));
 if(m==='GET'){const x=server.files[f];return x?res(200,{sha:x.sha,content:x.content}):res(404,{});}
 if(m==='PUT'){const b=JSON.parse(opt.body);const x=server.files[f];if(x&&b.sha!==x.sha)return res(409,{});if(!x&&b.sha)return res(422,{});
  const sha=crypto.randomBytes(6).toString('hex');server.files[f]={sha,content:b.content};server.puts++;return res(200,{content:{sha}});}
 return res(405,{});}
const wait=ms=>new Promise(r=>setTimeout(r,ms));
function device(){
 const dom=new JSDOM(html,{runScripts:'dangerously',url:'https://x.test/',pretendToBeVisual:true,beforeParse(w){w.TextEncoder=util.TextEncoder;w.TextDecoder=util.TextDecoder;w.fetch=mockFetch;w.scrollTo=()=>{};w.ExcelJS=ExcelJS;w.Date=Date;w.Blob=Blob;}});
 const w=dom.window,errs=[];w.addEventListener('error',e=>errs.push(e.message));
 return {w,d:w.document,errs,ev:s=>w.eval(s),click:s=>{const e=w.document.querySelector(s);if(!e)throw new Error('missing '+s);e.click();},text:()=>w.document.body.textContent};
}
let fails=0;const ok=(c,m)=>{if(!c)fails++;console.log((c?'PASS ':'FAIL ')+m);};
async function setupSync(D){D.click('[data-gx="settings"]');D.d.getElementById('sy-repo').value='me/data';D.d.getElementById('sy-token').value='good';D.click('[data-gx="save"]');await wait(2500);}
const readRemote=async()=>{const wb=new ExcelJS.Workbook();await wb.xlsx.load(Buffer.from(server.files['Vino.xlsx'].content,'base64'));return wb;};
(async()=>{
 const A=device(),B=device();await wait(300);
 ok(/Nessun file Excel/.test(A.text()),'senza file: schermata iniziale');
 await setupSync(A);
 const dA=A.ev('S.data');ok(dA&&dA.rossi.length>50,'scaricato da repo: rossi='+(dA&&dA.rossi.length)+' nonrossi='+dA.nonrossi.length+' dolci='+dA.dolci.length+' magnum='+dA.magnum.length+' bevuti='+dA.bevuti.length);
 // confronto con python
 const py=JSON.parse(require('child_process').execSync('python3 py_read.py '+VINO,{maxBuffer:1e8}).toString());
 const mine=Object.fromEntries(Object.entries(dA).map(([k,v])=>[k,v.map(o=>Object.entries(o).filter(([kk])=>kk!=='row').map(([,x])=>x))]));
 let same=true;for(const k in py){if(JSON.stringify(py[k])!==JSON.stringify(mine[k])){same=false;console.log('diff in',k,JSON.stringify(py[k][0]),JSON.stringify(mine[k][0]));}}
 ok(same,'lettura identica al motore Python (tutti i fogli)');
 A.click('[data-tab="rossi"]');ok(A.d.querySelectorAll('.item').length===dA.rossi.length,'lista rossi a schermo');
 A.click('[data-tab="stats"]');const lastN=A.d.querySelectorAll('.last');ok(lastN.length===6,'stats: 6 ultime bottiglie');console.log('ultime:',[...lastN].map(x=>x.textContent.replace(/\s+/g,' ')).slice(0,3).join(' || '));lastN[0].click();ok(/Acquisto|Quantità/.test(A.d.getElementById('modal').textContent),'tocco apre il dettaglio');A.ev('S.modal=null;render()');
 // modifiche su A
 const r0=dA.rossi[0];
 await A.ev("run('api_update',['rossi',"+r0.row+","+JSON.stringify(r0.Nome)+",{'Quantità':"+(Number(r0['Quantità'])+1)+"}],'')");
 await A.ev("run('api_add',['nonrossi',{Nome:'Test Bianco',Cantina:'Cantina X',Anno:'2022',Acquisto:'2026-10',Prezzo:'12,5',Valore:'15',Quantità:'2',Cantinetta:'1',Colore:'Bianco'}],'')");
 const dd=A.ev('S.data');const rr=dd.rossi.find(x=>x.Nome===dA.rossi[3].Nome&&x.Anno===dA.rossi[3].Anno);
 await A.ev("run('api_drink',['rossi',"+rr.row+","+JSON.stringify(rr.Nome)+",true],'')");
 ok(A.ev('OPS.length')===3,'3 modifiche in sospeso');ok(!A.d.getElementById('badge').hidden&&A.d.getElementById('badge').textContent==='3','badge 3');
 ok(A.ev('S.data').bevuti.length===dA.bevuti.length+1,'bevuta registrata');
 ok(server.puts===0,'nessun invio automatico');
 A.click('[data-gx="sync"]');await wait(2500);
 ok(server.puts===1&&A.ev('OPS.length')===0,'invio con pulsante');
 const rw=await readRemote();const sh=rw.getWorksheet('Non rossi');let found=null;sh.eachRow((r,i)=>{if(String(r.getCell(1).value).trim()==='Test Bianco')found=r;});
 ok(found&&found.getCell(5).value===12.5&&found.getCell(9).value==='Bianco','riga nuova nel file con valori');
 // stile preservato
 const orig=new ExcelJS.Workbook();await orig.xlsx.load(fs.readFileSync(VINO));
 const canon=o=>o===null||typeof o!=='object'?JSON.stringify(o):Array.isArray(o)?'['+o.map(canon)+']':'{'+Object.keys(o).sort().map(k=>JSON.stringify(k)+':'+canon(o[k]))+'}';
 const cmp=(a,b,addr)=>canon(a.getCell(addr).style)===canon(b.getCell(addr).style);
 ok(cmp(orig.getWorksheet('Rossi'),rw.getWorksheet('Rossi'),'A2')&&cmp(orig.getWorksheet('Rossi'),rw.getWorksheet('Rossi'),'C5'),'stili celle conservati');
 ok(String(sh.getCell('J2').value&&sh.getCell('J2').value.formula||'').includes('SUM'),'formule conservate');
 ok(rw.getWorksheet('_sync')&&rw.getWorksheet('_sync').state==='veryHidden','foglio _sync nascosto');
 // B scarica
 await setupSync(B);ok(B.ev('S.data').nonrossi.some(x=>x.Nome==='Test Bianco'),'B riceve le modifiche di A');
 // modifiche concorrenti
 const bb=B.ev('S.data'),t=bb.nonrossi.find(x=>x.Nome==='Test Bianco');
 await A.ev("run('api_update',['nonrossi',"+A.ev('S.data').nonrossi.find(x=>x.Nome==='Test Bianco').row+",'Test Bianco',{'Quantità':5}],'')");
 await B.ev("run('api_add',['dolci',{Nome:'Dolce B',Prezzo:'10',Valore:'12',Quantità:'1'}],'')");
 await B.ev("run('api_add',['magnum',{Nome:'Magnum B',Prezzo:'50',Valore:'60',Quantità:'1'}],'')");
 B.click('[data-gx="sync"]');await wait(2500);
 A.click('[data-gx="sync"]');await wait(2500);
 const dA2=A.ev('S.data');ok(dA2.dolci.some(x=>x.Nome==='Dolce B')&&dA2.magnum.some(x=>x.Nome==='Magnum B')&&!dA2.dolci.some(x=>x.Nome==='Magnum B'),'merge: dolci e magnum nei blocchi giusti');
 ok(dA2.nonrossi.find(x=>x.Nome==='Test Bianco')['Quantità']===5,'merge: modifica di A mantenuta');
 B.click('[data-gx="sync"]');await wait(2500);
 ok(JSON.stringify(A.ev('S.data'))===JSON.stringify(B.ev('S.data')),'A e B allineati');
 // elimina
 const tb=A.ev('S.data').nonrossi.find(x=>x.Nome==='Test Bianco');
 await A.ev("run('api_delete',['nonrossi',"+tb.row+",'Test Bianco'],'')");A.click('[data-gx="sync"]');await wait(2500);
 B.click('[data-gx="sync"]');await wait(2500);ok(!B.ev('S.data').nonrossi.some(x=>x.Nome==='Test Bianco'),'eliminazione propagata');
 // conflitto: riga cambiata altrove
 const n0=B.ev('S.data').dolci.find(x=>x.Nome==='Dolce B');
 await B.ev("run('api_delete',['dolci',"+n0.row+",'Dolce B'],'')");
 await A.ev("run('api_update',['dolci',"+A.ev('S.data').dolci.find(x=>x.Nome==='Dolce B').row+",'Dolce B',{'Quantità':9}],'')");
 B.click('[data-gx="sync"]');await wait(2500);A.click('[data-gx="sync"]');await wait(2500);
 ok(!A.ev('S.data').dolci.some(x=>x.Nome==='Dolce B'),'modifica su riga eliminata altrove: scartata senza errori');
 // file Excel di partenza non modificato nel confronto dei dati originali
 ok(A.errs.length+B.errs.length===0,'nessun errore JS '+A.errs.concat(B.errs).join(';'));
 // download
 fs.writeFileSync('/tmp/claude-0/x/out.xlsx',Buffer.from(server.files['Vino.xlsx'].content,'base64'));console.log(fails?'FALLITI '+fails:'TUTTO OK');process.exit(0);
})();
