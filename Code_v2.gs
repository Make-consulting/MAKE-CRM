// ============================================================
// MAKE CONSULTING — Apps Script v2
// ============================================================
const CLE_SECRETE = "MAKE2026$$";
const SHEET_ID = "12D4PMPBbMYjOiFtglAbZHm0SmlT7la944VqsOBajkb8";

function doGet(e) { globalThis._currentEvent = e; return handleRequest(e); }
function doPost(e) { globalThis._currentEvent = e; return handleRequest(e); }

function handleRequest(e) {
  try {
    const params = e.parameter || {};
    const postData = e.postData ? JSON.parse(e.postData.contents || "{}") : {};
    const data = Object.assign({}, params, postData);
    if (data.action === "recevoir_lead_site") return jsonResponse(recevoirLeadSite(data));
    if (data.cle !== CLE_SECRETE) return jsonResponse({ erreur: "Accès refusé" });
    switch (data.action || "") {
      case "lire_config":                    return jsonResponse(lireConfig());
      case "lire_kpi":                       return jsonResponse(lireKPI(parseInt(data.annee)||2026));
      case "lire_factures":                  {const _rf=lireOngletFiltre("Factures",data.annee,0);(_rf.factures||[]).forEach(f=>{f.jours=parseFloat(f["Jours passés"])||0;});return jsonResponse(_rf);}
      case "lire_devis":                     return jsonResponse(lireDevis(data.annee));
      case "lire_clients":                   return jsonResponse(lireClients());
      case "lire_prospects":                 return jsonResponse(lireOngletFiltre("Prospects",data.annee,1));
      case "lire_depenses":                  return jsonResponse(lireOngletFiltre("Dépenses",data.annee,0));
      case "lire_sdis":                      return jsonResponse(lireSDIS());
      case "lire_immobilier":                return jsonResponse(lireImmobilier());
      case "lire_locataires":                return jsonResponse(lireLocataires());
      case "lire_historique":                return jsonResponse(lireHistorique());
      case "lire_firetraining":              return jsonResponse(lireFiretraining());
      case "lire_missions":                  return jsonResponse(lireMissions(parseInt(data.annee)||2026));
      case "modifier_statut_facture":        return jsonResponse(modifierStatutFacture(data.numero,data.statut,data.date_paiement));
      case "modifier_statut_devis":          return jsonResponse(modifierStatutDevis(data.numero,data.statut,data.date_reponse,data.ligne));
      case "ajouter_facture":                return jsonResponse(ajouterFacture(JSON.parse(data.facture)));
      case "ajouter_devis":                  return jsonResponse(ajouterDevis(JSON.parse(data.devis)));
      case "ajouter_client":                 return jsonResponse(ajouterClient(JSON.parse(data.client)));
      case "modifier_client":                return jsonResponse(modifierClient(parseInt(data.ligne),JSON.parse(data.client)));
      case "ajouter_prospect":               return jsonResponse(ajouterProspect(JSON.parse(data.prospect)));
      case "ajouter_depense":                return jsonResponse(ajouterDepense(JSON.parse(data.depense)));
      case "ajouter_locataire":              return jsonResponse(ajouterLocataire(JSON.parse(data.locataire)));
      case "modifier_locataire":             return jsonResponse(modifierLocataire(parseInt(data.ligne),JSON.parse(data.locataire)));
      case "valider_paiement_loyer":         return jsonResponse(validerPaiementLoyer(data.mois_ligne,data.loyer,data.charges,data.date));
      case "modifier_parametre_bien":        return jsonResponse(modifierParametreBien(data.param,data.valeur));
      case "saisir_sdis_mois":               return jsonResponse(saisirSDISMois(data.mois_ligne,data.vacations,data.heures));
      case "ajouter_vente_firetraining":     return jsonResponse(ajouterVenteFiretraining(JSON.parse(data.vente)));
      case "generer_numero_facture":         return jsonResponse(genererNumeroFacture(parseInt(data.annee)||2026));
      case "generer_numero_devis":           return jsonResponse(genererNumeroDevis(parseInt(data.annee)||2026));
      case "modifier_statut_prospect":       return jsonResponse(modifierStatutProspect(parseInt(data.ligne),data.statut));
      case "ajouter_historique_prospect":    return jsonResponse(ajouterHistoriqueProspect(parseInt(data.ligne),data.type,data.commentaire));
      case "modifier_relance_prospect":      return jsonResponse(modifierRelanceProspect(parseInt(data.ligne),data.date_relance));
      case "ajouter_mission":               return jsonResponse(ajouterMission(JSON.parse(data.mission)));
      case "modifier_mission":              return jsonResponse(modifierMission(parseInt(data.ligne),JSON.parse(data.mission)));
      case "ajouter_commentaire_mission":   return jsonResponse(ajouterCommentaireMission(parseInt(data.ligne),data.commentaire));
      case "ajouter_au_calendrier":         return jsonResponse(ajouterAuCalendrier(parseInt(data.ligne)));
      default: return jsonResponse({ erreur: "Action inconnue : "+(data.action||"") });
    }
  } catch(err) { return jsonResponse({ erreur: err.message }); }
}

// ── Lead depuis formulaire site web ──────────────────────────
function recevoirLeadSite(data) {
  if (data.token !== "MAKE-SITE-2026") return { erreur: "Token invalide" };
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName("Prospects");
  if (!sheet) return { erreur: "Onglet Prospects introuvable" };
  const id = "PRO-" + String(sheet.getLastRow() - 1).padStart(3, "0");
  const annee = new Date().getFullYear();
  const entreprise = (data.organisation || "").trim() || "Particulier";
  const interlocuteur = ((data.prenom || "") + " " + (data.nom || "")).trim();
  const auj = new Date().toLocaleDateString("fr-FR");
  const relance = ajouterJoursOuvres(new Date(), 3).toLocaleDateString("fr-FR");
  sheet.appendRow([
    id, annee, entreprise, "", interlocuteur, "",
    data.telephone || "", data.email || "", data.objet || "",
    "Site web", "Lead", "B",
    auj, auj, relance,
    "Non", "", "", data.message || "", "[]"
  ]);
  try {
    const objet = data.objet || "—";
    const sujet = "🔔 Nouveau lead site — " + entreprise + " (" + objet + ")";
    const corps =
      "Nom : " + interlocuteur + "\n" +
      "Entreprise : " + entreprise + "\n" +
      "Email : " + (data.email || "—") + "\n" +
      "Téléphone : " + (data.telephone || "—") + "\n" +
      "Objet : " + objet + "\n" +
      "Message : " + (data.message || "—") + "\n" +
      "Date : " + auj + "\n" +
      "Lien direct Sheets : https://docs.google.com/spreadsheets/d/" + SHEET_ID;
    MailApp.sendEmail("msoileux.make@gmail.com", sujet, corps);
  } catch(e) {
    Logger.log("Erreur envoi mail notification lead : " + e.message);
  }
  return { succes: true, id };
}

function ajouterJoursOuvres(date, n) {
  const d = new Date(date);
  let ajoutes = 0;
  while (ajoutes < n) {
    d.setDate(d.getDate() + 1);
    const j = d.getDay();
    if (j !== 0 && j !== 6) ajoutes++;
  }
  return d;
}

function lireConfig() {
  const d = SpreadsheetApp.openById(SHEET_ID).getSheetByName("CONFIG").getDataRange().getValues();
  const c = {};
  for(let i=2;i<=11;i++) if(d[i]&&d[i][0]) c[String(d[i][0])]=d[i][1];
  const obj=[];
  for(let i=15;i<=20;i++) if(d[i]&&d[i][0]) obj.push({annee:d[i][0],statut:d[i][1],ca_cible:d[i][2],tjm_cible:d[i][3],jours_cibles:d[i][4],clients_cibles:d[i][5],taux_urssaf:d[i][6]});
  c.objectifs=obj; return c;
}

function lireKPI(annee) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const factData=ss.getSheetByName("Factures").getDataRange().getValues();
  const factH=factData[1];
  const ci={};factH.forEach((h,j)=>{if(h)ci[String(h)]=j;});
  const colHT=ci["Montant HT (€)"]!==undefined?ci["Montant HT (€)"]:8;
  const colSt=ci["Statut"]!==undefined?ci["Statut"]:15;
  const colCl=ci["Client"]!==undefined?ci["Client"]:4;
  const colTy=ci["Type"]!==undefined?ci["Type"]:6;
  const colJr=ci["Jours passés"]!==undefined?ci["Jours passés"]:-1;
  const colTVA_e=ci["TVA (€)"]!==undefined?ci["TVA (€)"]:10;
  let ca_f=0,ca_e=0,ca_p=0,jours_f=0,tva_collectee=0;
  const jPT={Formation:0,Consulting:0,"Ass. Admin":0,Autre:0};
  const jPC={},caPC={};
  factData.slice(2).forEach(r=>{
    if(String(r[0])!==String(annee)||!r[colHT])return;
    const ht=parseFloat(r[colHT])||0;
    ca_f+=ht;
    if(String(r[colSt])==="Payé")ca_e+=ht;
    if(String(r[colCl]).toUpperCase().includes("PROFORM"))ca_p+=ht;
    const j=colJr>=0?parseFloat(r[colJr])||0:0;
    jours_f+=j;
    const ty=String(r[colTy]||'');const cl=String(r[colCl]||'');
    if(ty.includes('Formation'))jPT.Formation+=j;
    else if(ty.includes('Admin'))jPT['Ass. Admin']+=j;
    else if(ty.includes('Consult'))jPT.Consulting+=j;
    else if(j>0)jPT.Autre+=j;
    if(cl){jPC[cl]=(jPC[cl]||0)+j;caPC[cl]=(caPC[cl]||0)+ht;}
    tva_collectee+=parseFloat(r[colTVA_e])||0;
  });
  const JOURS_AN=228;
  const taux_occupation=Math.round(jours_f/JOURS_AN*1000)/1000;
  const now=new Date();
  const jourN=Math.max(1,Math.ceil((now-new Date(now.getFullYear(),0,1))/86400000));
  const projection_annuelle=Math.round(jours_f/jourN*365);
  const jours_par_client=Object.keys(jPC).map(c=>({client:c,jours:Math.round(jPC[c]*10)/10,ca:Math.round(caPC[c]||0),tjm:jPC[c]>0?Math.round((caPC[c]||0)/jPC[c]):0})).sort((a,b)=>b.jours-a.jours);
  const sdisD=ss.getSheetByName("SDIS").getDataRange().getValues();
  let sdis_v=0,sdis_h=0;
  for(let i=9;i<=20;i++) if(sdisD[i]){sdis_v+=parseFloat(sdisD[i][1])||0;sdis_h+=parseFloat(sdisD[i][2])||0;}
  if(sdis_v===0) for(let i=3;i<=7;i++) if(sdisD[i]&&String(sdisD[i][0])===String(annee)){sdis_v=parseFloat(sdisD[i][1])||0;sdis_h=parseFloat(sdisD[i][2])||0;}
  const immoD=ss.getSheetByName("Immobilier").getDataRange().getValues();
  let lr=0,cr=0;
  for(let i=27;i<=38;i++) if(immoD[i]){lr+=parseFloat(immoD[i][3])||0;cr+=parseFloat(immoD[i][4])||0;}
  const cred=(parseFloat(immoD[5]?immoD[5][1]:0)||745)*12;
  const chR=(parseFloat(immoD[6]?immoD[6][1]:0)||365)+(parseFloat(immoD[7]?immoD[7][1]:0)||612)+(parseFloat(immoD[8]?immoD[8][1]:0)||122.5)+(parseFloat(immoD[9]?immoD[9][1]:0)||20)+(parseFloat(immoD[10]?immoD[10][1]:0)||0);
  const imp=parseFloat(immoD[11]?immoD[11][1]:0)||432;
  const inv=parseFloat(immoD[12]?immoD[12][1]:0)||0;
  const immo_net=Math.round((lr+cr-cred-chR-imp-inv)*100)/100;
  let dv=0;
  try{const s=ss.getSheetByName("Devis");if(s)s.getDataRange().getValues().slice(2).forEach(r=>{if(String(r[0])===String(annee)&&String(r[10])==="En attente")dv++;});}catch(e){}
  // Frais missions + missions à venir — utilise les en-têtes pour rester robuste
  let total_frais_missions=0,missions_a_venir=0;
  try{
    const misSheet=ss.getSheetByName("Missions");
    if(misSheet){
      const misData=misSheet.getDataRange().getValues();
      const misH=misData[1]||[];
      let fraisIdx=-1,anneeIdx=-1,forfaitIdx=-1,statutIdx=-1;
      misH.forEach((h,j)=>{
        if(String(h)==="Frais mission")fraisIdx=j;
        if(String(h)==="Année")anneeIdx=j;
        if(String(h)==="Forfait HT")forfaitIdx=j;
        if(String(h)==="Statut")statutIdx=j;
      });
      if(fraisIdx>=0&&anneeIdx>=0){
        misData.slice(2).forEach(r=>{
          if(String(r[anneeIdx])!==String(annee))return;
          total_frais_missions+=parseFloat(r[fraisIdx])||0;
          const st=String(r[statutIdx]||"");
          if(forfaitIdx>=0&&(st==="Confirmé"||st==="À venir"))missions_a_venir+=parseFloat(r[forfaitIdx])||0;
        });
      }
      total_frais_missions=Math.round(total_frais_missions);
      missions_a_venir=Math.round(missions_a_venir);
    }
  }catch(e){Logger.log("Missions KPI error: "+e.message);}
  let depenses_total=0;
  try{
    const depSheet=ss.getSheetByName("Dépenses");
    if(depSheet){
      const depData=depSheet.getDataRange().getValues();
      const depH=depData[1]||[];
      let depHtIdx=-1;
      depH.forEach((h,j)=>{if(String(h)==="Montant HT (€)")depHtIdx=j;});
      if(depHtIdx>=0)depData.slice(2).forEach(r=>{if(String(r[0])===String(annee))depenses_total+=parseFloat(r[depHtIdx])||0;});
      depenses_total=Math.round(depenses_total);
    }
  }catch(e){Logger.log("Dépenses KPI error: "+e.message);}
  const mois_ecoules=new Date().getMonth()+1;
  return{annee,ca_facture:Math.round(ca_f),ca_encaisse:Math.round(ca_e),ca_attente:Math.round(ca_f-ca_e),part_proform:ca_f>0?Math.round(ca_p/ca_f*100)/100:0,sdis:Math.round(sdis_v),sdis_heures:Math.round(sdis_h),immo_net,devis_attente:dv,jours_factures:Math.round(jours_f*10)/10,jours_par_type:jPT,jours_par_client,taux_occupation,projection_annuelle,frais_missions:total_frais_missions,depenses_total,missions_a_venir,mois_ecoules,tva_collectee:Math.round(tva_collectee)};
}

function lireOngletFiltre(nom,annee,colA) {
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName(nom);
  if(!sheet) return{erreur:"Onglet introuvable : "+nom};
  const data=sheet.getDataRange().getValues();const headers=data[1];const items=[];
  for(let i=2;i<data.length;i++){const r=data[i];if(!r[0]&&!r[1])continue;if(annee&&String(r[colA])!==String(annee))continue;const o={};headers.forEach((h,j)=>{if(h)o[String(h)]=r[j];});o._ligne=i+1;items.push(o);}
  const k=nom==="Factures"?"factures":nom==="Prospects"?"prospects":"depenses";
  return{[k]:items,total:items.length};
}

function lireDevis(annee) {
  const ss=SpreadsheetApp.openById(SHEET_ID);
  let sheet=ss.getSheetByName("Devis");
  if(!sheet){sheet=ss.insertSheet("Devis");const h=["Année","N° Devis","Date","Client","ID Client","Type","Description","Montant HT (€)","TVA (%)","TVA (€)","Statut","Date réponse","Notes"];sheet.getRange(1,1,1,h.length).setValues([h]);sheet.getRange(2,1,1,h.length).setValues([h]);}
  const data=sheet.getDataRange().getValues();if(data.length<3)return{devis:[]};
  const headers=data[1];const devis=[];
  for(let i=2;i<data.length;i++){if(!data[i][0])continue;if(annee&&String(data[i][0])!==String(annee))continue;const o={};headers.forEach((h,j)=>{if(h)o[String(h)]=data[i][j];});o._ligne=i+1;devis.push(o);}
  return{devis};
}

function lireClients() {
  const data=SpreadsheetApp.openById(SHEET_ID).getSheetByName("Clients").getDataRange().getValues();
  const headers=data[1];const clients=[];
  for(let i=2;i<data.length;i++){if(!data[i][0]||!data[i][1])continue;const o={};headers.forEach((h,j)=>{if(h)o[String(h)]=data[i][j];});o._ligne=i+1;clients.push(o);}
  return{clients};
}

function lireSDIS() {
  const data=SpreadsheetApp.openById(SHEET_ID).getSheetByName("SDIS").getDataRange().getValues();
  const annuel=[];
  for(let i=3;i<=7;i++) if(data[i]&&data[i][0]){const v=parseFloat(data[i][1])||0,h=parseFloat(data[i][2])||0;annuel.push({annee:data[i][0],vacations:v,cumul_heures:h,ratio_h24:h>0?Math.round(h/24*100)/100:0,eur_ratio:h>0?Math.round(v/(h/24)*100)/100:0});}
  const mensuel=[];let ch=0,cv=0;
  for(let i=9;i<=20;i++) if(data[i]&&data[i][0]){const v=parseFloat(data[i][1])||0,h=parseFloat(data[i][2])||0;ch+=h;cv+=v;mensuel.push({mois:data[i][0],vacations:v,heures_mois:h,cumul_heures:Math.round(ch*10)/10,ratio_h24:Math.round(ch/24*100)/100,eur_ratio:ch>0?Math.round(cv/(ch/24)*100)/100:0,statut:data[i][6]||"",notes:data[i][7]||"",_ligne:i+1});}
  const tv=mensuel.reduce((s,m)=>s+m.vacations,0),th=mensuel.reduce((s,m)=>s+m.heures_mois,0);
  return{annuel,mensuel,total_2026:{vacations:Math.round(tv),cumul_heures:Math.round(th*10)/10,ratio_h24:Math.round(th/24*100)/100,eur_ratio:th>0?Math.round(tv/(th/24)*100)/100:0}};
}

function lireImmobilier() {
  const data=SpreadsheetApp.openById(SHEET_ID).getSheetByName("Immobilier").getDataRange().getValues();
  const params={},paramLabels=[];
  for(let i=3;i<=12;i++) if(data[i]&&data[i][0]){params[String(data[i][0])]=data[i][1];paramLabels.push({label:String(data[i][0]),valeur:data[i][1],ligne:i+1});}
  const locataire={};
  for(let i=3;i<=8;i++) if(data[i]&&data[i][3])locataire[String(data[i][3])]=data[i][4];
  const mensuel=[];
  for(let i=27;i<=38;i++) if(data[i]&&data[i][0]){const lr=parseFloat(data[i][3])||0,cr=parseFloat(data[i][4])||0,lp=parseFloat(data[i][1])||890,cp=parseFloat(data[i][2])||50;mensuel.push({mois:data[i][0],loyer_prevu:lp,charges_prevues:cp,loyer_recu:lr,charges_recues:cr,date_paiement:data[i][5],statut:String(data[i][6]||""),delta:lr+cr-lp-cp,quittance:data[i][8],notes:data[i][9],_ligne:i+1});}
  const historique=[];
  for(let i=42;i<=46;i++) if(data[i]&&data[i][0])historique.push({annee:data[i][0],loyer_mensuel:data[i][1],total_loyers:data[i][2],charges_encaissees:data[i][3],credit:data[i][4],charges_reelles:data[i][5],impots:data[i][6],net:data[i][7]});
  return{params,paramLabels,locataire,mensuel,historique};
}

function lireLocataires() {
  const ss=SpreadsheetApp.openById(SHEET_ID);
  let sheet=ss.getSheetByName("Locataires");
  if(!sheet){sheet=ss.insertSheet("Locataires");const h=["ID","Bien","Nom complet","Email","Téléphone","Date entrée","Fin bail prévue","Dépôt garantie (€)","Date état des lieux","Compteur eau entrée (m³)","Compteur eau actuel (m³)","Compteur élec entrée (kWh)","Compteur élec actuel (kWh)","Statut","Notes"];sheet.getRange(1,1,1,h.length).setValues([h]);sheet.getRange(2,1,1,h.length).setValues([h]);sheet.getRange(3,1,1,h.length).setValues([["LOC-001","41 rue du Maréchal Joffre — Trilport","À renseigner","","","","",1780,"","","","","","Actuel",""]]);}
  const data=sheet.getDataRange().getValues();if(data.length<3)return{locataires:[]};
  const headers=data[1];const locataires=[];
  for(let i=2;i<data.length;i++){if(!data[i][0])continue;const o={};headers.forEach((h,j)=>{if(h)o[String(h)]=data[i][j];});o._ligne=i+1;locataires.push(o);}
  return{locataires};
}

function lireHistorique() {
  const ss=SpreadsheetApp.openById(SHEET_ID);
  const data=ss.getSheetByName("Historique_CA").getDataRange().getValues();
  const synthese=[];
  for(let i=2;i<=6;i++) if(data[i]&&data[i][0])synthese.push({annee:data[i][0],statut:data[i][1],ca:data[i][2],sdis:data[i][3],immo:data[i][4],revenu_consolide:data[i][5],urssaf:data[i][6],tjm:data[i][7]});
  let ca2026=0;
  ss.getSheetByName("Factures").getDataRange().getValues().slice(2).forEach(r=>{if(String(r[0])==="2026"&&r[8])ca2026+=parseFloat(r[8])||0;});
  const i26=synthese.findIndex(s=>String(s.annee)==="2026");
  if(i26>=0)synthese[i26].ca=Math.round(ca2026);else synthese.push({annee:2026,statut:"EI",ca:Math.round(ca2026),sdis:0,immo:0,revenu_consolide:Math.round(ca2026),urssaf:0,tjm:0});
  const headers=data.length>9?data[9]:[];const detail=[];
  for(let i=10;i<data.length;i++){if(!data[i][0])continue;const o={};headers.forEach((h,j)=>{if(h)o[String(h)]=data[i][j];});detail.push(o);}
  return{synthese,detail};
}

function lireFiretraining() {
  const data=SpreadsheetApp.openById(SHEET_ID).getSheetByName("FireTraining_MS").getDataRange().getValues();
  const ventes=[];
  for(let i=11;i<=60;i++) if(data[i]&&data[i][0])ventes.push({numero:data[i][0],annee:data[i][1],date:data[i][2],client:data[i][3],contact:data[i][4],email:data[i][5],telephone:data[i][6],nb_licences:data[i][7],prix_unitaire:data[i][8],total_ht:data[i][9],tva:data[i][10],total_ttc:data[i][11],facture:data[i][12],statut:data[i][13],notes:data[i][14],_ligne:i+1});
  return{ventes};
}

// ── MISSIONS ─────────────────────────────────────────────────
// Col : 1=ID, 2=Année, 3=Client, 4=ID Client, 5=Type, 6=Nom mission,
//       7=Statut, 8=Date début, 9=Date fin, 10=Lieu, 11=Horaires,
//       12=Jours prévus, 13=Jours réalisés, 14=Forfait HT (€),
//       15=Frais mission (€), 16=CA net frais (€), 17=Devis lié,
//       18=Facture liée, 19=Ajouté au calendrier, 20=Commentaires JSON
const MISSIONS_HEADERS = ["ID","Année","Client","ID Client","Type","Nom mission","Statut","Date début","Date fin","Lieu","Horaires","Jours prévus","Jours réalisés","Forfait HT","Frais mission","CA net frais","Devis lié","Facture liée","Ajouté au calendrier","Commentaires JSON"];

function getMissionsSheet() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName("Missions");
  if (!sheet) {
    sheet = ss.insertSheet("Missions");
    sheet.getRange(1, 1, 1, MISSIONS_HEADERS.length).setValues([MISSIONS_HEADERS]);
    sheet.getRange(2, 1, 1, MISSIONS_HEADERS.length).setValues([MISSIONS_HEADERS]);
  }
  return sheet;
}

function lireMissions(annee) {
  const sheet = getMissionsSheet();
  const data = sheet.getDataRange().getValues();
  if (data.length < 3) return { missions: [] };
  const headers = data[1];
  const missions = [];
  for (let i = 2; i < data.length; i++) {
    if (!data[i][0] && !data[i][1]) continue;
    if (annee && String(data[i][1]) !== String(annee)) continue; // col 2 = Année (index 1)
    const o = {};
    headers.forEach((h, j) => { if (h) o[String(h)] = data[i][j]; });
    o._ligne = i + 1;
    missions.push(o);
  }
  return { missions };
}

function ajouterMission(m) {
  const sheet = getMissionsSheet();
  const annee = m.annee || new Date().getFullYear();
  const allData = sheet.getDataRange().getValues();
  let max = 0;
  allData.slice(2).forEach(r => {
    if (String(r[1]) === String(annee)) { // col 2 = Année (index 1)
      const match = String(r[0] || "").match(/(\d+)$/); // col 1 = ID (index 0)
      if (match) max = Math.max(max, parseInt(match[1]));
    }
  });
  const id = "MIS-" + annee + "-" + String(max + 1).padStart(3, "0");
  const forfait = parseFloat(m.forfait_ht) || 0;
  const frais = parseFloat(m.frais_mission) || 0;
  const ca_net = Math.round((forfait - frais) * 100) / 100;
  sheet.appendRow([
    id,                                 // 1 : ID
    annee,                              // 2 : Année
    m.client || "",                     // 3 : Client
    m.id_client || "",                  // 4 : ID Client
    m.type || "",                       // 5 : Type
    m.nom_mission || "",                // 6 : Nom mission
    m.statut || "Demande",             // 7 : Statut
    m.date_debut || "",                 // 8 : Date début
    m.date_fin || "",                   // 9 : Date fin
    m.lieu || "",                       // 10 : Lieu
    m.horaires || "",                   // 11 : Horaires
    parseFloat(m.jours_prevus) || 0,    // 12 : Jours prévus
    parseFloat(m.jours_realises) || 0, // 13 : Jours réalisés
    forfait,                            // 14 : Forfait HT (€)
    frais,                              // 15 : Frais mission (€)
    ca_net,                             // 16 : CA net frais (€)
    m.num_devis || "",                  // 17 : Devis lié
    m.num_facture || "",                // 18 : Facture liée
    "Non",                              // 19 : Ajouté au calendrier
    "[]"                                // 20 : Commentaires JSON
  ]);
  return { succes: true, id };
}

function modifierMission(ligne, m) {
  const sheet = getMissionsSheet();
  const forfait = parseFloat(m.forfait_ht) || 0;
  const frais = parseFloat(m.frais_mission) || 0;
  const ca_net = Math.round((forfait - frais) * 100) / 100;
  [
    [3,  m.client],
    [4,  m.id_client],
    [5,  m.type],
    [6,  m.nom_mission],
    [7,  m.statut],
    [8,  m.date_debut],
    [9,  m.date_fin],
    [10, m.lieu],
    [11, m.horaires],
    [12, parseFloat(m.jours_prevus) || 0],
    [13, parseFloat(m.jours_realises) || 0],
    [14, forfait],
    [15, frais],
    [16, ca_net],
    [17, m.num_devis],
    [18, m.num_facture]
  ].forEach(([col, val]) => { if (val !== undefined) sheet.getRange(ligne, col).setValue(val); });
  return { succes: true };
}

function ajouterCommentaireMission(ligne, commentaire) {
  const sheet = getMissionsSheet();
  const raw = sheet.getRange(ligne, 20).getValue() || "[]"; // col 20 = Commentaires JSON
  let hist = [];
  try { hist = JSON.parse(raw); } catch(e) { hist = []; }
  let comm;
  try { comm = typeof commentaire === "string" ? JSON.parse(commentaire) : commentaire; }
  catch(e) { comm = { type: "Note", note: String(commentaire) }; }
  hist.push({
    date: new Date().toLocaleDateString("fr-FR"),
    type: comm.type || "Note",
    note: comm.note || ""
  });
  sheet.getRange(ligne, 20).setValue(JSON.stringify(hist)); // col 20 = Commentaires JSON
  return { succes: true };
}

function ajouterAuCalendrier(ligne) {
  const sheet = getMissionsSheet();
  const r = sheet.getRange(ligne, 1, 1, 20).getValues()[0];
  // index : 0=ID, 1=Année, 2=Client, 3=ID Client, 4=Type, 5=Nom mission,
  //         6=Statut, 7=Date début, 8=Date fin, 9=Lieu, 10=Horaires,
  //         11=Jours prévus, 12=Jours réalisés, 13=Forfait HT,
  //         14=Frais, 15=CA net, 16=Devis, 17=Facture,
  //         18=Ajouté au calendrier, 19=Commentaires JSON
  const type = String(r[4] || ""), nom = String(r[5] || ""), client = String(r[2] || "");
  const lieu = String(r[9] || ""), horaires = String(r[10] || "");
  const forfait = r[13] || 0;
  const titre = "[" + type + "] " + nom + " — " + client;

  function parseSheetDate(d) {
    if (!d) return new Date();
    if (d instanceof Date) return d;
    const s = String(d).trim();
    if (s.match(/^\d{4}-\d{2}-\d{2}$/)) {
      const [y, mo, dy] = s.split("-").map(Number);
      return new Date(y, mo - 1, dy);
    }
    if (s.includes("/")) {
      const p = s.split("/");
      if (p.length === 3) return new Date(parseInt(p[2]), parseInt(p[1]) - 1, parseInt(p[0]));
    }
    return new Date(s);
  }

  const dateDebut = parseSheetDate(r[7]);  // col 8 = Date début (index 7)
  const dateFin = parseSheetDate(r[8] || r[7]);  // col 9 = Date fin (index 8)
  const dateFinCal = new Date(dateFin);
  dateFinCal.setDate(dateFinCal.getDate() + 1);

  const description = "Forfait HT : " + forfait + " €\nLieu : " + lieu + "\nHoraires : " + horaires;
  const cal = CalendarApp.getDefaultCalendar();
  cal.createAllDayEvent(titre, dateDebut, dateFinCal, { description });
  sheet.getRange(ligne, 19).setValue("Oui"); // col 19 = Ajouté au calendrier
  return { succes: true };
}

// ── Factures / Devis ─────────────────────────────────────────
function ajouterFacture(f) {
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName("Factures");
  const ht=parseFloat(f.montant_ht)||0,tva=parseFloat(f.tva)||0,tva_e=Math.round(ht*tva/100*100)/100;
  sheet.appendRow([f.annee||new Date().getFullYear(),f.numero||"",f.mois||new Date().toLocaleDateString("fr-FR",{month:"long"}),f.date_emission||new Date().toLocaleDateString("fr-FR"),f.client||"",f.id_client||"",f.type||"",f.description||"",ht,tva,tva_e,Math.round((ht+tva_e)*100)/100,f.date_envoi||new Date().toLocaleDateString("fr-FR"),f.date_echeance||"","","En attente","",f.notes||""]);
  return{succes:true,numero:f.numero};
}

function ajouterDevis(d) {
  const ss=SpreadsheetApp.openById(SHEET_ID);
  let sheet=ss.getSheetByName("Devis");
  if(!sheet){sheet=ss.insertSheet("Devis");const h=["Année","N° Devis","Date","Client","ID Client","Type","Description","Montant HT (€)","TVA (%)","TVA (€)","Statut","Date réponse","Notes"];sheet.getRange(1,1,1,h.length).setValues([h]);sheet.getRange(2,1,1,h.length).setValues([h]);}
  const ht=parseFloat(d.montant_ht)||0,tva=parseFloat(d.tva)||0,tva_e=Math.round(ht*tva/100*100)/100;
  sheet.appendRow([d.annee||new Date().getFullYear(),d.numero||"",d.date||new Date().toLocaleDateString("fr-FR"),d.client||"",d.id_client||"",d.type||"",d.description||"",ht,tva,tva_e,"En attente","",d.notes||""]);
  return{succes:true,numero:d.numero};
}

function modifierStatutDevis(numero,statut,dateReponse,ligne) {
  const ss=SpreadsheetApp.openById(SHEET_ID);
  let sheet=ss.getSheetByName("Devis");
  if(!sheet)return{erreur:"Onglet Devis introuvable"};
  const data=sheet.getDataRange().getValues();
  for(let i=2;i<data.length;i++){
    if(String(data[i][1])===String(numero)||String(i+1)===String(ligne)){
      sheet.getRange(i+1,11).setValue(statut);
      if(dateReponse)sheet.getRange(i+1,12).setValue(dateReponse);
      return{succes:true,ligne:i+1};
    }
  }
  if(ligne&&parseInt(ligne)>2){
    sheet.getRange(parseInt(ligne),11).setValue(statut);
    if(dateReponse)sheet.getRange(parseInt(ligne),12).setValue(dateReponse);
    return{succes:true,ligne:parseInt(ligne)};
  }
  return{erreur:"Devis non trouvé : "+numero};
}

function modifierStatutFacture(numero,statut,datePmt) {
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName("Factures");
  const data=sheet.getDataRange().getValues();
  for(let i=2;i<data.length;i++){if(String(data[i][1])===String(numero)){sheet.getRange(i+1,16).setValue(statut);if(datePmt)sheet.getRange(i+1,15).setValue(datePmt);return{succes:true,ligne:i+1};}}
  return{erreur:"Facture non trouvée : "+numero};
}

function ajouterClient(c) {
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName("Clients");
  const id="CLI-"+String(sheet.getLastRow()-1).padStart(3,"0");
  sheet.appendRow([id,c.entreprise||"",c.statut_jur||"",c.siren||"",c.secteur||"",c.type_relation||"",c.interlocuteur||"",c.fonction||"",c.email||"",c.telephone||"",c.adresse||"",c.cp||"",c.ville||"",c.tva||0,c.delai_pmt||30,new Date().toLocaleDateString("fr-FR"),new Date().toLocaleDateString("fr-FR"),0,0,0,0,0,0,"Oui",c.notes||""]);
  return{succes:true,id};
}

function modifierClient(ligne,c) {
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName("Clients");
  [[2,c.entreprise],[3,c.statut_jur],[4,c.siren],[5,c.secteur],[6,c.type_relation],[7,c.interlocuteur],[8,c.fonction],[9,c.email],[10,c.telephone],[11,c.adresse],[12,c.cp],[13,c.ville],[14,c.tva],[15,c.delai_pmt],[25,c.notes]].forEach(([col,val])=>{if(val!==undefined)sheet.getRange(ligne,col).setValue(val);});
  sheet.getRange(ligne,17).setValue(new Date().toLocaleDateString("fr-FR"));
  return{succes:true};
}

function ajouterProspect(p) {
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName("Prospects");
  const id="PRO-"+String(sheet.getLastRow()-1).padStart(3,"0");
  sheet.appendRow([id,new Date().getFullYear(),p.entreprise||"",p.secteur||"",p.interlocuteur||"",p.fonction||"",p.telephone||"",p.email||"",p.type_prestation||"",p.source||"",p.statut||"À contacter",p.priorite||"B",new Date().toLocaleDateString("fr-FR"),"","","Non","","",p.commentaire||""]);
  return{succes:true,id};
}

function ajouterDepense(d) {
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName("Dépenses");
  const annee=d.annee||new Date().getFullYear();
  const id="DEP-"+annee+"-"+String(sheet.getLastRow()-1).padStart(3,"0");
  const ht=parseFloat(d.montant_ht)||0,tva=parseFloat(d.tva)||20,tva_e=Math.round(ht*tva/100*100)/100;
  sheet.appendRow([annee,id,d.date||new Date().toLocaleDateString("fr-FR"),d.categorie||"",d.sous_categorie||"",d.fournisseur||"",d.description||"",ht,tva,tva_e,ht+tva_e,d.remboursable||"Non",d.rembourse_par||"",d.notes||""]);
  return{succes:true,id};
}

function ajouterLocataire(l) {
  const ss=SpreadsheetApp.openById(SHEET_ID);
  const sheet=ss.getSheetByName("Locataires")||ss.insertSheet("Locataires");
  const id="LOC-"+String(Math.max(sheet.getLastRow()-1,1)).padStart(3,"0");
  sheet.appendRow([id,l.bien||"",l.nom||"",l.email||"",l.telephone||"",l.date_entree||"",l.fin_bail||"",parseFloat(l.depot_garantie)||0,l.date_etat_lieux||"",parseFloat(l.compteur_eau_entree)||0,parseFloat(l.compteur_eau_actuel)||0,parseFloat(l.compteur_elec_entree)||0,parseFloat(l.compteur_elec_actuel)||0,l.statut||"Actuel",l.notes||""]);
  return{succes:true,id};
}

function modifierLocataire(ligne,l) {
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName("Locataires");
  if(!sheet)return{erreur:"Onglet Locataires introuvable"};
  [[2,l.bien],[3,l.nom],[4,l.email],[5,l.telephone],[6,l.date_entree],[7,l.fin_bail],[8,l.depot_garantie],[9,l.date_etat_lieux],[10,l.compteur_eau_entree],[11,l.compteur_eau_actuel],[12,l.compteur_elec_entree],[13,l.compteur_elec_actuel],[14,l.statut],[15,l.notes]].forEach(([col,val])=>{if(val!==undefined)sheet.getRange(ligne,col).setValue(val);});
  return{succes:true};
}

function validerPaiementLoyer(ligne,loyer,charges,date) {
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName("Immobilier");
  sheet.getRange(ligne,4).setValue(parseFloat(loyer)||0);
  sheet.getRange(ligne,5).setValue(parseFloat(charges)||0);
  sheet.getRange(ligne,6).setValue(date||"");
  sheet.getRange(ligne,7).setValue("Payé");
  return{succes:true};
}

function modifierParametreBien(param,valeur) {
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName("Immobilier");
  const data=sheet.getDataRange().getValues();
  for(let i=3;i<=12;i++){if(data[i]&&String(data[i][0])===String(param)){sheet.getRange(i+1,2).setValue(parseFloat(valeur)||valeur);return{succes:true,ligne:i+1};}}
  return{erreur:"Paramètre non trouvé : "+param};
}

function saisirSDISMois(ligne,vac,h) {
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName("SDIS");
  sheet.getRange(ligne,2).setValue(parseFloat(vac)||0);
  sheet.getRange(ligne,3).setValue(parseFloat(h)||0);
  sheet.getRange(ligne,7).setValue("Saisi");
  return{succes:true};
}

function ajouterVenteFiretraining(v) {
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName("FireTraining_MS");
  const annee=new Date().getFullYear();
  const numero="FT-"+annee+"-"+String(Math.max(sheet.getLastRow()-10,1)).padStart(3,"0");
  const pu=parseFloat(v.prix_unitaire)||0,nb=parseInt(v.nb_licences)||1,ht=pu*nb,tva_e=Math.round(ht*0.20*100)/100;
  sheet.appendRow([numero,annee,new Date().toLocaleDateString("fr-FR"),v.client||"",v.contact||"",v.email||"",v.telephone||"",nb,pu,ht,20,ht+tva_e,"","À facturer",v.notes||""]);
  return{succes:true,numero};
}

function modifierStatutProspect(ligne,statut) {
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName("Prospects");
  if(!sheet)return{erreur:"Onglet Prospects introuvable"};
  sheet.getRange(ligne,11).setValue(statut);
  sheet.getRange(ligne,14).setValue(new Date().toLocaleDateString("fr-FR"));
  return{succes:true};
}

function ajouterHistoriqueProspect(ligne,type,commentaire) {
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName("Prospects");
  if(!sheet)return{erreur:"Onglet Prospects introuvable"};
  const raw=sheet.getRange(ligne,20).getValue()||"[]";
  let hist=[];
  try{hist=JSON.parse(raw);}catch(e){hist=[];}
  hist.push({date:new Date().toLocaleDateString("fr-FR"),type:type||"",commentaire:commentaire||""});
  sheet.getRange(ligne,20).setValue(JSON.stringify(hist));
  sheet.getRange(ligne,14).setValue(new Date().toLocaleDateString("fr-FR"));
  return{succes:true};
}

function modifierRelanceProspect(ligne,dateRelance) {
  const sheet=SpreadsheetApp.openById(SHEET_ID).getSheetByName("Prospects");
  if(!sheet)return{erreur:"Onglet Prospects introuvable"};
  let valeur=dateRelance||"";
  if(valeur.includes("-")){const p=valeur.split("-");if(p.length===3)valeur=p[2]+"/"+p[1]+"/"+p[0];}
  sheet.getRange(ligne,15).setValue(valeur);
  return{succes:true};
}

function genererNumeroFacture(annee) {
  const data=SpreadsheetApp.openById(SHEET_ID).getSheetByName("Factures").getDataRange().getValues();
  let max=0;
  data.slice(2).forEach(r=>{if(String(r[0])===String(annee)){const m=String(r[1]||"").match(/(\d+)$/);if(m)max=Math.max(max,parseInt(m[1]));}});
  return{numero:"FA"+annee+"-"+String(max+1).padStart(3,"0")};
}

function genererNumeroDevis(annee) {
  const ss=SpreadsheetApp.openById(SHEET_ID);
  const sheet=ss.getSheetByName("Devis");
  if(!sheet)return{numero:"DV"+annee+"-001"};
  let max=0;
  sheet.getDataRange().getValues().slice(2).forEach(r=>{if(String(r[0])===String(annee)){const m=String(r[1]||"").match(/(\d+)$/);if(m)max=Math.max(max,parseInt(m[1]));}});
  return{numero:"DV"+annee+"-"+String(max+1).padStart(3,"0")};
}

function jsonResponse(data) {
  const json=JSON.stringify(data);
  try{const ev=globalThis._currentEvent;if(ev&&ev.parameter&&ev.parameter.callback){const out=ContentService.createTextOutput(ev.parameter.callback+"("+json+")");out.setMimeType(ContentService.MimeType.JAVASCRIPT);return out;}}catch(e){}
  const out=ContentService.createTextOutput(json);out.setMimeType(ContentService.MimeType.JSON);return out;
}
