/** RENDEVO Barber & Salon — separate Google Apps Script backend.
 * Setup: create a NEW spreadsheet and a NEW Apps Script project. Paste this file.
 * Run setupRendevo() once, authorize, then Deploy > New deployment > Web app.
 * Execute as: Me. Access: Anyone. Put deployed /exec URL in barber/config.js.
 * Do not paste this into the medical project's Apps Script.
 */
const SETTINGS = {
  timezone: 'Africa/Algiers',
  spreadsheetId: '', // Optional: paste the NEW barber/salon spreadsheet ID. Blank uses the active spreadsheet.
  businessName: 'RENDEVO Barber & Salon',
  slotStepMinutes: 15,
  daysAhead: 60
};
const DEFAULT_SERVICES = [
  ['Coupe classique',30,800,'barber',true],
  ['Dégradé / fade',40,1200,'barber',true],
  ['Taille de barbe',20,500,'barber',true],
  ['Coupe + barbe',50,1500,'barber',true],
  ['Brushing',30,1000,'salon',true],
  ['Coupe femme',45,1500,'salon',true],
  ['Coloration',90,3500,'salon',true],
  ['Soin capillaire',45,1800,'salon',true]
];
function setupRendevo() {
  const ss = SETTINGS.spreadsheetId ? SpreadsheetApp.openById(SETTINGS.spreadsheetId) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Ouvrez Apps Script depuis une nouvelle feuille Google Sheets ou renseignez spreadsheetId.');
  PropertiesService.getScriptProperties().setProperty('RENDEVO_SPREADSHEET_ID', ss.getId());
  const services = ensureSheet_(ss, 'Services', ['ID','Nom','Durée (min)','Prix (DA)','Catégorie','Actif']);
  if (services.getLastRow() < 2) DEFAULT_SERVICES.forEach((s,i)=>services.appendRow(['S'+(i+1),s[0],s[1],s[2],s[3],s[4]]));
  const staff = ensureSheet_(ss, 'Équipe', ['ID','Nom','Rôle','Catégorie','Actif','Jours (0=dimanche)','Début','Fin','Pause début','Pause fin']);
  if (staff.getLastRow() < 2) {
    staff.appendRow(['P1','Professionnel 1','Barbier','barber',true,'1,2,3,4,5,6','09:00','18:00','13:00','14:00']);
    staff.appendRow(['P2','Professionnel 2','Coiffeur / coiffeuse','salon',true,'1,2,3,4,5,6','09:00','18:00','13:00','14:00']);
  }
  ensureSheet_(ss, 'Rendez-vous', ['Référence','Créé le','Statut','Date','Heure','Fin','Service ID','Prestation','Prix (DA)','Professionnel ID','Professionnel','Nom client','Téléphone','E-mail','Notes','Jeton annulation']);
  ensureSheet_(ss, 'Indisponibilités', ['Date','Début','Fin','Professionnel ID','Motif']);
  ensureSheet_(ss, 'Paramètres', ['Clé','Valeur']);
  const settings=ss.getSheetByName('Paramètres');
  if(settings.getLastRow()<2) [['Fuseau horaire',SETTINGS.timezone],['Pas de créneau (minutes)',SETTINGS.slotStepMinutes],['Délai max (jours)',SETTINGS.daysAhead]].forEach(r=>settings.appendRow(r));
  return {success:true,spreadsheetId:ss.getId(),message:'Structure RENDEVO créée. Configurez les vrais professionnels, horaires et prestations avant toute utilisation.'};
}
function ensureSheet_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);if(sh.getLastRow()===0)sh.appendRow(headers);return sh;}
function json_(data,callback){const body=JSON.stringify(data);if(callback&&/^[A-Za-z_$][\w$\.]{0,80}$/.test(callback))return ContentService.createTextOutput(callback+'('+body+');').setMimeType(ContentService.MimeType.JAVASCRIPT);return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);}
function doGet(e){const p=e&&e.parameter||{},cb=p.callback;try{let out;if(p.action==='config')out=getConfig_();else if(p.action==='availability')out={success:true,slots:getSlots_(p.date,p.serviceId,p.staffId)};else if(p.action==='book')out=book_(p);else if(p.action==='cancel')out=cancel_(p);else out={success:true,service:SETTINGS.businessName,message:'Backend disponible'};return json_(out,cb)}catch(err){return json_({success:false,message:err.message||'Erreur serveur'},cb)}}
function ss_(){const id=PropertiesService.getScriptProperties().getProperty('RENDEVO_SPREADSHEET_ID')||SETTINGS.spreadsheetId;if(id)return SpreadsheetApp.openById(id);const active=SpreadsheetApp.getActiveSpreadsheet();if(!active)throw new Error('Exécutez setupRendevo() depuis le projet lié à la nouvelle feuille.');return active;}
function rows_(name){const sh=ss_().getSheetByName(name);if(!sh||sh.getLastRow()<2)return [];const v=sh.getDataRange().getValues();return v.slice(1).filter(r=>r.some(x=>x!==''));}
function getConfig_(){return {success:true,services:rows_('Services').filter(r=>r[5]===true||String(r[5]).toLowerCase()==='true').map(r=>({id:String(r[0]),name:String(r[1]),duration:Number(r[2]),price:Number(r[3]),category:String(r[4])})),staff:rows_('Équipe').filter(r=>r[4]===true||String(r[4]).toLowerCase()==='true').map(r=>({id:String(r[0]),name:String(r[1]),role:String(r[2]),category:String(r[3])}))};}
function parseTime_(v){if(v instanceof Date)return v.getHours()*60+v.getMinutes();const m=String(v).match(/^(\d{1,2}):(\d{2})$/);if(!m)throw new Error('Horaire invalide. Utilisez HH:MM.');return Number(m[1])*60+Number(m[2]);}
function dateObj_(s){if(!/^\d{4}-\d{2}-\d{2}$/.test(String(s)))throw new Error('Date invalide.');const d=new Date(s+'T12:00:00');if(isNaN(d.getTime()))throw new Error('Date invalide.');return d;}
function fmtDate_(d){return Utilities.formatDate(d,SETTINGS.timezone,'yyyy-MM-dd');}
function getSlots_(date,serviceId,staffId){const d=dateObj_(date),today=fmtDate_(new Date());if(date<today)throw new Error('Choisissez une date future.');const max=new Date();max.setDate(max.getDate()+SETTINGS.daysAhead);if(d>max)throw new Error('Cette date est trop éloignée.');const svc=rows_('Services').find(r=>String(r[0])===String(serviceId)&&(r[5]===true||String(r[5]).toLowerCase()==='true'));if(!svc)throw new Error('Prestation introuvable.');const people=rows_('Équipe').filter(r=>(r[4]===true||String(r[4]).toLowerCase()==='true')&&(staffId==='any'||String(r[0])===String(staffId))&&(!r[3]||String(r[3])===String(svc[4])));const bookings=rows_('Rendez-vous').filter(r=>String(r[2])==='Confirmé'&&fmtDate_(dateObj_(fmtDate_(r[3] instanceof Date?r[3]:dateObj_(String(r[3])))))===date);const blocks=rows_('Indisponibilités');const day=d.getDay(),duration=Number(svc[2]);const result=new Set();people.forEach(p=>{const days=String(p[5]||'1,2,3,4,5,6').split(',').map(x=>Number(x.trim()));if(!days.includes(day))return;const start=parseTime_(p[6]||'09:00'),end=parseTime_(p[7]||'18:00'),pauseStart=p[8]?parseTime_(p[8]):-1,pauseEnd=p[9]?parseTime_(p[9]):-1;for(let t=start;t+duration<=end;t+=SETTINGS.slotStepMinutes){if(t<1440&&pauseStart>=0&&t<pauseEnd&&t+duration>pauseStart)continue;const finish=t+duration;const conflict=bookings.some(b=>String(b[9])===String(p[0])&&overlap_(t,finish,parseTime_(String(b[4] instanceof Date?Utilities.formatDate(b[4],SETTINGS.timezone,'HH:mm'):b[4])),parseTime_(String(b[5] instanceof Date?Utilities.formatDate(b[5],SETTINGS.timezone,'HH:mm'):b[5]))));const blocked=blocks.some(b=>{if(fmtDate_(b[0] instanceof Date?b[0]:dateObj_(String(b[0])))!==date||String(b[3]||'')&&String(b[3])!==String(p[0]))return false;return overlap_(t,finish,parseTime_(b[1]),parseTime_(b[2]));});if(!conflict&&!blocked)result.add(('0'+Math.floor(t/60)).slice(-2)+':'+('0'+t%60).slice(-2));}});return Array.from(result).sort();}
function overlap_(a,b,c,d){return a<d&&c<b;}
function book_(p){const lock=LockService.getScriptLock();lock.waitLock(10000);try{const service=rows_('Services').find(r=>String(r[0])===String(p.serviceId)&&(r[5]===true||String(r[5]).toLowerCase()==='true'));if(!service)throw new Error('Prestation indisponible.');const name=String(p.name||'').trim(),phone=String(p.phone||'').trim();if(name.length<2||name.length>100||phone.length<6||phone.length>30)throw new Error('Nom ou téléphone invalide.');const email=String(p.email||'').trim();if(email.length>150||email&& !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email))throw new Error('Adresse e-mail invalide.');if(String(p.notes||'').length>500)throw new Error('Notes trop longues.');const slots=getSlots_(p.date,p.serviceId,p.staffId);if(!slots.includes(String(p.time)))throw new Error('Ce créneau vient d’être pris. Choisissez un autre horaire.');const svc=service,people=rows_('Équipe').filter(r=>(r[4]===true||String(r[4]).toLowerCase()==='true')&&(!r[3]||String(r[3])===String(svc[4]))&&(p.staffId==='any'||String(r[0])===String(p.staffId)));const chosen=people.find(r=>{const days=String(r[5]||'1,2,3,4,5,6').split(',').map(Number);const d=dateObj_(p.date);const t=parseTime_(p.time);return days.includes(d.getDay())&&t+Number(svc[2])<=parseTime_(r[7]||'18:00')&&(!r[8]||!overlap_(t,t+Number(svc[2]),parseTime_(r[8]),parseTime_(r[9]||r[8])));});if(!chosen)throw new Error('Professionnel indisponible pour ce créneau.');const start=parseTime_(p.time),end=start+Number(svc[2]),time=n=>('0'+Math.floor(n/60)).slice(-2)+':'+('0'+n%60).slice(-2),ref='RDV-'+Utilities.getUuid().slice(0,8).toUpperCase(),token=Utilities.getUuid();ss_().getSheetByName('Rendez-vous').appendRow([ref,new Date(),'Confirmé',dateObj_(p.date),time(start),time(end),svc[0],svc[1],Number(svc[3]),chosen[0],chosen[1],name,phone,email,String(p.notes||''),token]);return {success:true,reference:ref,staffName:String(chosen[1]),cancelToken:token,message:'Rendez-vous confirmé.'};}finally{lock.releaseLock();}}
function cancel_(p){const ref=String(p.reference||''),token=String(p.token||'');if(!ref||!token)throw new Error('Référence ou jeton manquant.');const sh=ss_().getSheetByName('Rendez-vous');const data=sh.getDataRange().getValues();for(let i=1;i<data.length;i++){if(String(data[i][0])===ref&&String(data[i][15])===token){if(String(data[i][2])!=='Confirmé')throw new Error('Ce rendez-vous est déjà annulé ou terminé.');sh.getRange(i+1,3).setValue('Annulé');return {success:true,message:'Rendez-vous annulé. Le créneau est de nouveau disponible.'};}}throw new Error('Rendez-vous introuvable ou lien invalide.');}
