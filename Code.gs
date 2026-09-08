// Google Apps Script: осы кодты Google Sheet-ке байланысқан Apps Script жобасына салыңыз.
// Sheet атаулары: Users, Progress, Feedback
function doPost(e){
  const d=JSON.parse(e.postData.contents||'{}');
  try{
    if(d.action==='register') return out(register(d));
    if(d.action==='login') return out(login(d));
    if(d.action==='progress') return out(saveProgress(d));
    if(d.action==='dashboard') return out(dashboard());
    return out({ok:false,message:'Unknown action'});
  }catch(err){return out({ok:false,message:String(err)})}
}
function out(x){return ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON)}
function sh(name,headers){const ss=SpreadsheetApp.getActive();let s=ss.getSheetByName(name);if(!s){s=ss.insertSheet(name);s.appendRow(headers)}else if(s.getLastColumn()<headers.length){s.getRange(1,1,1,headers.length).setValues([headers])}return s}
function hash(v){const b=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,v);return b.map(x=>(x+256)%256).map(x=>x.toString(16).padStart(2,'0')).join('')}
function register(d){const s=sh('Users',['ID','Name','Class','Login','PasswordHash','Avatar','Created']);const vals=s.getDataRange().getValues();if(vals.slice(1).some(r=>r[3]===d.user))return {ok:false,message:'Бұл логин тіркелген'};s.appendRow([Utilities.getUuid(),d.name,d.class,d.user,hash(d.pass),d.avatar||'🧬',new Date()]);return {ok:true,message:'Тіркелу сәтті'} }
function login(d){const s=sh('Users',['ID','Name','Class','Login','PasswordHash','Avatar','Created']);const row=s.getDataRange().getValues().slice(1).find(r=>r[3]===d.user&&r[4]===hash(d.pass));if(!row)return {ok:false,message:'Логин немесе құпиясөз қате'};const headers=['Login','Lesson','Score','Mood','Feedback','Date','StageStart','StageMiddle','StageEnd','Duration','Descriptors'];const p=sh('Progress',headers).getDataRange().getValues().slice(1).filter(r=>r[0]===d.user);const progress={};p.forEach(r=>{const n=Number(r[1]),key=n>100?Math.floor(n/100)+'-'+(n%100-1):String(n-1);progress[key]={score:Number(r[2]),mood:r[3],feedback:r[4],date:r[5],stages:{start:Number(r[6]||0),middle:Number(r[7]||0),end:Number(r[8]||0)},duration:Number(r[9]||0),criteria:String(r[10]||'').split(' | ')}});return {ok:true,user:{id:row[0],name:row[1],class:row[2],user:row[3],avatar:row[5]||'🧬',role:'student',progress:progress}}}
function saveProgress(d){const headers=['Login','Lesson','Score','Mood','Feedback','Date','StageStart','StageMiddle','StageEnd','Duration','Descriptors'];const s=sh('Progress',headers);const rows=s.getDataRange().getValues();let found=0;for(let i=1;i<rows.length;i++){if(rows[i][0]===d.user&&Number(rows[i][1])===Number(d.lesson)){found=i+1;break}}const values=[d.user,d.lesson,d.score,d.mood,d.feedback,new Date(),d.stageStart||0,d.stageMiddle||0,d.stageEnd||0,d.duration||0,d.descriptors||''];if(found)s.getRange(found,1,1,values.length).setValues([values]);else s.appendRow(values);return {ok:true}}
function dashboard(){const users=sh('Users',['ID','Name','Class','Login','PasswordHash','Avatar','Created']).getDataRange().getValues().slice(1);const progress=sh('Progress',['Login','Lesson','Score','Mood','Feedback','Date']).getDataRange().getValues().slice(1);return {ok:true,students:users.map(u=>({name:u[1],class:u[2],user:u[3],avatar:u[5]||'🧬',progress:progress.filter(p=>p[0]===u[3])}))}}
