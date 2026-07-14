const SHEET_NAME = "Registrations";
const ORGANIZER_EMAIL = "vailvelocity@gmail.com";

const HEADERS = [
  "Registration ID","Timestamp","Submitted At (Local)","Season","Registration Fee",
  "Payment Status","Registration Status","Coach Applicant","Parent Name","Relationship",
  "Email","Phone","Home Address","Player Name","Preferred Name","Birthdate","Grade",
  "School Attending","School District","Closest School","Alternate School","Experience",
  "Jersey Size","Unavailable Nights","Placement Request","Coach Interest","Coach Experience",
  "Previous Coaching","Coach Academy Commitment","Background Check","Emergency Contact",
  "Emergency Phone","Medical Information","Waiver Agreement","Refund Agreement",
  "Media Permission","Electronic Signature","Signature Date","Team","Assigned Coach",
  "Practice Location","Game Schedule","Jersey Ordered","Payment Date","Stripe Checkout Session",
  "Stripe Payment Intent","Payment Confirmation Sent","Refund Status","Organizer Notes"
];

function doGet(){
  return ContentService.createTextOutput("Vail Velocity registration endpoint is active.");
}

function doPost(e){
  const lock=LockService.getScriptLock();
  try{
    lock.waitLock(30000);
    const s=getSheet_(), p=(e&&e.parameter)||{};
    const id=clean_(p.registration_id)||fallbackId_();
    if(findRow_(s,id)) return json_({ok:true,registrationId:id,duplicate:true});
    const row=[
      id,new Date(),clean_(p.submitted_at_local),clean_(p.season||"Fall 2026"),"135",
      "Pending","Awaiting Payment",coachStatus_(p.coach_interest),clean_(p.parent_name),clean_(p.relationship),
      clean_(p.email),clean_(p.phone),clean_(p.home_address),clean_(p.player_name),clean_(p.preferred_name),
      clean_(p.birthdate),clean_(p.grade),clean_(p.school_attending),clean_(p.school_district),
      clean_(p.closest_school),clean_(p.alternate_school),clean_(p.experience),clean_(p.jersey_size),
      clean_(p.unavailable_nights),clean_(p.placement_request),clean_(p.coach_interest),
      clean_(p.coach_experience),clean_(p.previous_coaching),clean_(p.coach_training),
      clean_(p.background_check),clean_(p.emergency_contact),clean_(p.emergency_phone),
      clean_(p.medical_information),agree_(p.waiver_agreement),agree_(p.refund_agreement),
      clean_(p.media_permission),clean_(p.electronic_signature),clean_(p.signature_date),
      "","","","","Not Ordered","","","","No","Not Requested",""
    ];
    s.appendRow(row);
    sendReceived_(p,id);
    sendOrganizer_(p,id);
    return json_({ok:true,registrationId:id});
  }catch(err){return json_({ok:false,message:String(err)});}
  finally{try{lock.releaseLock();}catch(_){}}
}

function installPaymentReconciliationTrigger(){
  ScriptApp.getProjectTriggers().forEach(t=>{if(t.getHandlerFunction()==="reconcileStripePayments") ScriptApp.deleteTrigger(t);});
  ScriptApp.newTrigger("reconcileStripePayments").timeBased().everyMinutes(5).create();
  reconcileStripePayments();
}

function reconcileStripePayments(){
  const key=PropertiesService.getScriptProperties().getProperty("STRIPE_SECRET_KEY");
  if(!key) throw new Error("Add STRIPE_SECRET_KEY in Project Settings → Script Properties.");
  const s=getSheet_(), pending=pendingMap_(s);
  if(!Object.keys(pending).length) return;
  let after="", pages=0;
  do{
    let url="https://api.stripe.com/v1/checkout/sessions?status=complete&limit=100";
    if(after) url+="&starting_after="+encodeURIComponent(after);
    const r=UrlFetchApp.fetch(url,{headers:{Authorization:"Bearer "+key},muteHttpExceptions:true});
    if(r.getResponseCode()<200||r.getResponseCode()>=300) throw new Error(r.getContentText());
    const data=JSON.parse(r.getContentText()), sessions=data.data||[];
    sessions.forEach(x=>{
      const id=clean_(x.client_reference_id);
      if(id&&pending[id]&&x.payment_status==="paid"){
        markPaid_(s,pending[id],x);
        delete pending[id];
      }
    });
    if(!data.has_more||!sessions.length) break;
    after=sessions[sessions.length-1].id;
    pages++;
  }while(Object.keys(pending).length&&pages<5);
}

function testPaymentReconciliationNow(){reconcileStripePayments();}
function sendTestEmail(){MailApp.sendEmail({to:ORGANIZER_EMAIL,subject:"Vail Velocity email test",htmlBody:"<p>Email is working.</p>",name:"Vail Velocity"});}

function markPaid_(s,row,x){
  s.getRange(row,6).setValue("Paid");
  s.getRange(row,7).setValue("Complete");
  s.getRange(row,44).setValue(new Date((x.created||0)*1000));
  s.getRange(row,45).setValue(x.id||"");
  s.getRange(row,46).setValue(x.payment_intent||"");
  if(String(s.getRange(row,47).getValue()).toLowerCase()!=="yes"){
    const v=s.getRange(row,1,1,HEADERS.length).getValues()[0];
    sendPaid_({id:v[0],parent:v[8],email:v[10],player:v[13],amount:((x.amount_total||13500)/100).toFixed(2)});
    s.getRange(row,47).setValue("Yes");
  }
}

function sendReceived_(p,id){
  if(!p.email) return;
  MailApp.sendEmail({
    to:p.email, replyTo:ORGANIZER_EMAIL, name:"Vail Velocity",
    subject:"Vail Velocity registration received — "+id,
    htmlBody:"<p>Hi "+h_(p.parent_name||"there")+",</p><p>We received the registration for <strong>"+h_(p.player_name)+"</strong>.</p><p><strong>Registration ID:</strong> "+h_(id)+"<br><strong>Payment status:</strong> Awaiting payment<br><strong>Fee:</strong> $135<br><strong>Season begins:</strong> Week of August 3, 2026<br><strong>Game jersey:</strong> Included</p><p>Your registration is complete after payment is confirmed.</p><p>Thank you,<br><strong>Vail Velocity</strong></p>"
  });
}

function sendPaid_(d){
  if(!d.email) return;
  MailApp.sendEmail({
    to:d.email, replyTo:ORGANIZER_EMAIL, name:"Vail Velocity",
    subject:"Vail Velocity payment confirmed — "+d.id,
    htmlBody:"<p>Hi "+h_(d.parent||"there")+",</p><p>Payment has been confirmed for <strong>"+h_(d.player)+"</strong>.</p><p><strong>Registration ID:</strong> "+h_(d.id)+"<br><strong>Amount paid:</strong> $"+h_(d.amount)+"<br><strong>Status:</strong> Complete</p><p>Team and practice information will be emailed after assignments are finalized.</p><p>Thank you,<br><strong>Vail Velocity</strong></p>"
  });
}

function sendOrganizer_(p,id){
  MailApp.sendEmail({to:ORGANIZER_EMAIL,name:"Vail Velocity Registration",subject:"New registration — "+clean_(p.player_name),htmlBody:"<p><strong>ID:</strong> "+h_(id)+"<br><strong>Player:</strong> "+h_(p.player_name)+"<br><strong>Parent:</strong> "+h_(p.parent_name)+"<br><strong>Email:</strong> "+h_(p.email)+"<br><strong>Closest school:</strong> "+h_(p.closest_school)+"<br><strong>Coach interest:</strong> "+h_(p.coach_interest)+"</p>"});
}

function getSheet_(){
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let s=ss.getSheetByName(SHEET_NAME)||ss.insertSheet(SHEET_NAME);
  if(s.getLastRow()===0){s.getRange(1,1,1,HEADERS.length).setValues([HEADERS]);s.setFrozenRows(1);s.getRange(1,1,1,HEADERS.length).setBackground("#0A132B").setFontColor("#FFFFFF").setFontWeight("bold");}
  return s;
}
function pendingMap_(s){const m={};if(s.getLastRow()<2)return m;s.getRange(2,1,s.getLastRow()-1,7).getValues().forEach((r,i)=>{if(r[0]&&String(r[5]).toLowerCase()!=="paid")m[r[0]]=i+2;});return m;}
function findRow_(s,id){if(s.getLastRow()<2)return 0;const f=s.getRange(2,1,s.getLastRow()-1,1).createTextFinder(id).matchEntireCell(true).findNext();return f?f.getRow():0;}
function coachStatus_(v){v=clean_(v).toLowerCase();if(!v||v.includes("no, not"))return"No";if(v.includes("head"))return"Head Coach";if(v.includes("assistant"))return"Assistant Coach";if(v.includes("maybe"))return"Maybe";return"Yes";}
function fallbackId_(){return"VV-"+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyyMMdd-HHmmss")+"-"+Math.random().toString(36).slice(2,7).toUpperCase();}
function agree_(v){return v?"Agreed":"Not Agreed";}
function clean_(v){return v==null?"":String(v).trim();}
function h_(v){return clean_(v).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/\"/g,"&quot;").replace(/'/g,"&#039;");}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
