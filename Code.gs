// 인융인의 밤 입장 체크 서버. 운영진 Google 로그인(ID 토큰 검증) 후에만 판정한다.
// Script Properties(편집기에서만 설정): OWNER_EMAIL, CLIENT_ID, OID_HMAC_KEY, SHEET_ID, STAFF_EMAILS(선택).
// 학번·이름·QR 원문은 로그/응답에 남기지 않는다. 입장기록에는 복원 불가 HMAC 판정키만 저장한다.
const OID_TRIAL_ = Object.freeze({prefix:'OIT1:', challenge:180, session:1800,
  endpoint:'https://oauth2.googleapis.com/tokeninfo?id_token='});
const ADM_ = Object.freeze({tz:'Asia/Seoul', ledger:'입장기록', maxAge:600, skew:120,
  label:{ok:'입장',duplicate:'중복',wait:'대기',not_member:'비대상'}});
function oidCheck_(v) { if (!v) throw new Error('denied'); }
function oidDeny_() { return {status:'denied'}; }
function oidJson_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet(e) { return oidJson_(oidDeny_()); }
function doPost(e) {
  try {
    oidCheck_(e && !e.queryString && e.postData && /^(?:application\/json|text\/plain)(?:\s*;|$)/i.test(e.postData.type));
    const body=e.postData.contents;
    oidCheck_(typeof body==='string' && body.length<=8192);
    return oidJson_(oidDispatch_(JSON.parse(body)));
  } catch (_) { return oidJson_(oidDeny_()); }
}
function oidFields_(o,keys) {
  oidCheck_(o && typeof o==='object' && !Array.isArray(o) && Object.keys(o).sort().join('|')===keys.slice().sort().join('|') && keys.every(k=>Object.prototype.hasOwnProperty.call(o,k)));
}
function oidMac_(key,text) {
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(text,key,Utilities.Charset.UTF_8)).replace(/=+$/,'');
}
function oidEq_(a,b) {
  if (typeof a!=='string' || typeof b!=='string' || a.length!==b.length) return false;
  let d=0; for (let i=0;i<a.length;i++) d|=a.charCodeAt(i)^b.charCodeAt(i); return d===0;
}
function oidCfg_() {
  const props=PropertiesService.getScriptProperties(), owner=props.getProperty('OWNER_EMAIL'), client=props.getProperty('CLIENT_ID'), key=props.getProperty('OID_HMAC_KEY');
  oidCheck_(typeof owner==='string' && /^[a-z0-9]+(?:\.[a-z0-9]+)*@gmail\.com$/.test(owner) && typeof client==='string' && /^[0-9]+-[a-z0-9-]+\.apps\.googleusercontent\.com$/.test(client) && typeof key==='string' && /^[a-f0-9]{64}$/.test(key));
  const sheet=props.getProperty('SHEET_ID'); oidCheck_(typeof sheet==='string' && /^[A-Za-z0-9_-]{25,80}$/.test(sheet));
  const allowed=[owner].concat(String(props.getProperty('STAFF_EMAILS')||'').split(/[,;\s]+/).map(x=>x.trim().toLowerCase()).filter(x=>/^[^\s@]+@[^\s@]+\.[a-z]{2,}$/.test(x)));
  return {props:props,owner:owner,client:client,allowed:allowed,sheet:sheet,k:oidMac_(key,'trial-v1|'+client+'|'+owner),lk:oidMac_(key,'ledger|v1')};
}
function oidNow_() { return Math.floor(Date.now()/1000); }
function oidChallenge_(c,seed) {
  oidCheck_(typeof seed==='string' && /^[a-fA-F0-9]{64}$/.test(seed));
  const body='v1.'+oidNow_()+'.'+oidMac_(c.k,'issue|'+seed+'|'+Utilities.getUuid()).slice(0,22);
  return {status:'challenge',nonce:body+'.'+oidMac_(c.k,'nonce|'+body)};
}
function oidNonceEnd_(c,nonce,now) {
  oidCheck_(typeof nonce==='string' && /^v1\.[1-9][0-9]{9}\.[A-Za-z0-9_-]{22}\.[A-Za-z0-9_-]{43}$/.test(nonce));
  const n=nonce.split('.'), t=Number(n[1]), end=t+OID_TRIAL_.challenge;
  oidCheck_(t<=now && now<end && oidEq_(n[3],oidMac_(c.k,'nonce|'+n.slice(0,3).join('.'))));
  return end;
}
function oidPart_(segment) {
  const b=Utilities.base64DecodeWebSafe(segment+'='.repeat((4-segment.length%4)%4));
  const o=JSON.parse(Utilities.newBlob(b).getDataAsString('UTF-8'));
  oidCheck_(o && typeof o==='object' && !Array.isArray(o)); return o;
}
function oidNum_(x) {
  if (typeof x==='string') { oidCheck_(/^[1-9][0-9]{0,11}$/.test(x)); x=Number(x); }
  oidCheck_(Number.isSafeInteger(x) && x>0); return x;
}
function oidTimes_(iat,exp,now) { oidCheck_(iat<=now+30 && now<exp && iat<exp && exp-iat<=3900 && now-iat<=3900); }
function oidVerified_(c,token,nonce) {
  // Pass the exact, unmodified compact token. No redirects, fallback URLs, or retry on error.
  oidCheck_(typeof token==='string' && /^[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+$/.test(token));
  const url=OID_TRIAL_.endpoint+encodeURIComponent(token);
  oidCheck_(url.length<=2082); // Apps Script URL Fetch hard limit; never truncate a token.
  const r=UrlFetchApp.fetch(url,{method:'get',followRedirects:false,validateHttpsCertificates:true,muteHttpExceptions:true});
  oidCheck_(r.getResponseCode()===200);
  const text=r.getContentText(); oidCheck_(typeof text==='string' && text.length<=16384);
  const g=JSON.parse(text); oidCheck_(g && typeof g==='object' && !Array.isArray(g) && !('error' in g) && !('error_description' in g));
  const seg=token.split('.'), header=oidPart_(seg[0]), p=oidPart_(seg[1]);
  oidCheck_(header.alg==='RS256' && typeof header.kid==='string' && header.kid.length>0);
  for (const k of ['aud','azp','iss','sub','email']) oidCheck_(typeof g[k]==='string' && g[k]===p[k]);
  oidCheck_(g.aud===c.client && g.azp===c.client && (g.iss==='https://accounts.google.com' || g.iss==='accounts.google.com') && /^[!-~]{1,255}$/.test(g.sub) && c.allowed.indexOf(String(g.email).toLowerCase())>=0);
  oidCheck_((g.email_verified===true || g.email_verified==='true') && p.email_verified===true);
  const iat=oidNum_(g.iat), exp=oidNum_(g.exp);
  oidCheck_(Number.isSafeInteger(p.iat) && Number.isSafeInteger(p.exp) && p.iat===iat && p.exp===exp);
  oidTimes_(iat,exp,oidNow_());
  // Only now may the signed payload supply a nonce omitted by Google's tokeninfo.
  oidCheck_((!Object.prototype.hasOwnProperty.call(g,'nonce') || oidEq_(g.nonce,p.nonce)) && oidEq_(p.nonce,nonce));
  return {subject:oidMac_(c.k,'subject|'+g.sub),iat:iat,exp:exp};
}
function oidLock_(f) {
  const l=LockService.getScriptLock(); oidCheck_(l.tryLock(20000));
  try { return f(); } finally { l.releaseLock(); }
}
function oidKey_(type,id) { return OID_TRIAL_.prefix+type+':'+id; }
// This namespace only: records contain an expiry, and for sessions a keyed subject digest.
// Fail closed on a damaged record. Never delete unrelated configuration or state.
function oidRecord_(type,s) {
  const pat=type==='s'?/^([1-9][0-9]{9})\|([A-Za-z0-9_-]{43})$/:/^([1-9][0-9]{9})$/;
  const m=typeof s==='string' && s.match(pat); oidCheck_(m); return {end:Number(m[1]),subject:m[2]};
}
function oidClean_(props,now) {
  const all=props.getProperties();
  for (const k of Object.keys(all)) {
    const m=k.match(/^OIT1:([ns]):[A-Za-z0-9_-]{43}$/); if (!m) continue;
    const r=oidRecord_(m[1],all[k]); if (r.end<=now) props.deleteProperty(k);
  }
}
function oidLogin_(c,b) {
  oidNonceEnd_(c,b.nonce,oidNow_());
  const a=oidVerified_(c,b.credential,b.nonce);
  return oidLock_(()=>{
    const now=oidNow_(), end=oidNonceEnd_(c,b.nonce,now); oidTimes_(a.iat,a.exp,now);
    oidClean_(c.props,now);
    const used=oidKey_('n',oidMac_(c.k,'spent|'+b.nonce)); oidCheck_(c.props.getProperty(used)===null);
    const token=oidMac_(c.k,'session|'+b.nonce+'|'+Utilities.getUuid()), sk=oidKey_('s',oidMac_(c.k,'lookup|'+token));
    oidCheck_(c.props.getProperty(sk)===null);
    // Spend first. A failed subsequent write may burn a challenge, never re-enable it.
    c.props.setProperty(used,String(end));
    c.props.setProperty(sk,String(Math.min(now+OID_TRIAL_.session,a.exp))+'|'+a.subject);
    return {status:'ready',session:token};
  });
}
function admSession_(c,token) {
  oidCheck_(typeof token==='string' && /^[A-Za-z0-9_-]{43}$/.test(token));
  const sk=oidKey_('s',oidMac_(c.k,'lookup|'+token)), s=oidRecord_('s',c.props.getProperty(sk));
  oidCheck_(s.end>oidNow_()); return sk;
}
// 도서관 모바일 이용증 QR: "_KW" + 공백 + "0" + 학번10 + 발급시각14(yyyyMMddHHmmss, KST) + 숫자
function admParse_(raw) {
  if (typeof raw!=='string' || raw.length>120) return null;
  const m=/^_KW\s*0?(\d{10})(\d{14})\d{0,10}$/.exec(raw.trim()); if (!m) return null;
  const t=m[2], ms=Date.parse(t.slice(0,4)+'-'+t.slice(4,6)+'-'+t.slice(6,8)+'T'+t.slice(8,10)+':'+t.slice(10,12)+':'+t.slice(12,14)+'+09:00');
  if (!Number.isFinite(ms) || Utilities.formatDate(new Date(ms),ADM_.tz,'yyyyMMddHHmmss')!==t) return null;
  return {sid:m[1],ms:ms};
}
function admFresh_(ms) { const now=Date.now(); if (ms>now+ADM_.skew*1000) return 'invalid'; return now-ms>ADM_.maxAge*1000?'expired':''; }
function admHm_(v,fb) {
  const s=String(v||'').trim(); if (!s) return fb;
  const m=/^(\d{1,2}):([0-5]\d)(?::[0-5]\d)?$/.exec(s); oidCheck_(m && Number(m[1])<=23);
  return ('0'+m[1]).slice(-2)+':'+m[2];
}
function admHas_(sh,col,text) {
  const n=sh.getLastRow()-1;
  return n>0 && !!sh.getRange(2,col,n,1).createTextFinder(text).useRegularExpression(false).matchCase(true).matchEntireCell(true).findNext();
}
function admLedger_(ss,lk) {
  const marker='판정키_v1:'+oidMac_(lk,'marker').slice(0,16);
  let sh=ss.getSheetByName(ADM_.ledger);
  if (!sh) { sh=ss.insertSheet(ADM_.ledger); sh.appendRow(['시각(KST)','결과','학생회비','판정키(복원 불가)',marker]); sh.setFrozenRows(1); }
  else oidCheck_(sh.getRange('E1').getDisplayValue()===marker);
  return sh;
}
function admCount_(sh) { const n=sh.getLastRow()-1; return n>0?sh.getRange(2,2,n,1).createTextFinder(ADM_.label.ok).matchEntireCell(true).findAll().length:0; }
function admScan_(c,b) {
  admSession_(c,b.session); // 세션이 없거나 만료면 denied
  try {
    const q=admParse_(b.qr); if (!q) return {status:'invalid'};
    const early=admFresh_(q.ms); if (early) return {status:early};
    const l=LockService.getScriptLock(); if (!l.tryLock(20000)) return {status:'busy'};
    try {
      const stale=admFresh_(q.ms); if (stale) return {status:stale};
      const ss=SpreadsheetApp.openById(c.sheet), conf=ss.getSheetByName('설정'), roster=ss.getSheetByName('명단');
      oidCheck_(conf && roster);
      const now=new Date(), hm=Utilities.formatDate(now,ADM_.tz,'HH:mm');
      const paidAt=admHm_(conf.getRange('F3').getDisplayValue(),'16:30'), unpaidAt=admHm_(conf.getRange('F2').getDisplayValue(),'16:45');
      oidCheck_(paidAt<=unpaidAt);
      if (hm<paidAt) return {status:'closed',open:paidAt};
      const paid=admHas_(roster,1,q.sid), member=paid || admHas_(conf,1,q.sid.slice(4,7));
      const led=admLedger_(ss,c.lk), tag=oidMac_(c.lk,'sid|'+q.sid);
      let s;
      if (!member) s='not_member';
      else if (admHas_(led,4,'ok:'+tag)) s='duplicate';
      else if (!paid && hm<unpaidAt) s='wait';
      else s='ok';
      led.appendRow([Utilities.formatDate(now,ADM_.tz,'yyyy-MM-dd HH:mm:ss'),ADM_.label[s],member?(paid?'납부':'미납'):'',s+':'+tag]);
      SpreadsheetApp.flush();
      const out={status:s,count:admCount_(led)}; if (s==='wait') out.open=unpaidAt; return out;
    } finally { l.releaseLock(); }
  } catch (_) { return {status:'error'}; }
}
function admLogout_(c,b) { try { c.props.deleteProperty(admSession_(c,b.session)); } catch (_) {} return {status:'bye'}; }
function oidDispatch_(b) {
  try {
    const c=oidCfg_();
    if (b.action==='challenge') {oidFields_(b,['action','seed']); return oidChallenge_(c,b.seed);}
    if (b.action==='login') {oidFields_(b,['action','credential','nonce']); return oidLogin_(c,b);}
    if (b.action==='scan') {oidFields_(b,['action','session','qr']); return admScan_(c,b);}
    if (b.action==='logout') {oidFields_(b,['action','session']); return admLogout_(c,b);}
  } catch (_) { /* Deliberately omit details, provider exceptions can include a credential. */ }
  return oidDeny_();
}
