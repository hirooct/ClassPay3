/* =========================================================
 * ClassPay Phase 2
 * 会社申請UI / 政府会計連動 / ランキング
 * ========================================================= */

function setupClassPayPhase2(){
  setupClassPayPhase1();
  _ensureSheetWithHeader_(SHEETS.HOLDINGS || "Holdings", [
    "userId","shopId","shares","updatedAt"
  ]);
  _ensureSheetWithHeader_(SHEETS.GOVERNMENT_LEDGER || "GovernmentLedger", [
    "ledgerId","at","type","amount","balanceAfter","referenceId","note"
  ]);
  _ensureSheetWithHeader_(SHEETS.RETIREMENT_APPLICATIONS || "RetirementApplications", [
    "applicationId","submittedAt","shopId","shopName","userId","userName","reason","status","reviewedAt","reviewNote","reviewedBy","governmentStatus"
  ]);
  var snap = _getCompanySnapshotsSheet_();
  var required = ["snapshotAt","shopId","shopName","balance","previousBalance","growthAmount","growthRate","valueCreated"];
  var current = snap.getLastColumn() ? snap.getRange(1,1,1,snap.getLastColumn()).getValues()[0].map(String) : [];
  required.forEach(function(name){
    if (current.map(function(x){ return x.toLowerCase(); }).indexOf(name.toLowerCase()) < 0) {
      snap.getRange(1, snap.getLastColumn()+1).setValue(name);
      current.push(name);
    }
  });
  return {ok:true, version:"2.2", sheets:["CompanyMembers","CompanyApplications","RetirementApplications","Government","CompanySnapshots","Holdings",SHEETS.GOVERNMENT_LEDGER || "GovernmentLedger"]};
}

/** Shops列の並びが変わっていても、ヘッダー名に合わせて安全に追加する */
function _appendShopPhase2_(shopId,shopName,companyPass){
  const sh=_getShopsSheet_(), width=Math.max(sh.getLastColumn(),5);
  const headers=sh.getRange(1,1,1,width).getValues()[0];
  const m=_headerMap_(headers), row=new Array(width).fill("");
  const set=(name,value)=>{const c=m.idx(name);if(c<0)throw new Error("Shopsヘッダに "+name+" が必要です");row[c]=value;};
  set("shopid",shopId);set("shopname",shopName);set("balance",0);set("isactive",true);set("companypass",companyPass);
  sh.appendRow(row);
}

function _appendGovernmentLedger_(type, amount, referenceId, note){
  const g = _getGovernmentAccount_();
  const sh = _sheetByNameOrThrow_(SHEETS.GOVERNMENT_LEDGER || "GovernmentLedger");
  sh.appendRow([uuid_(),_fmtJst_(new Date()),String(type||""),Number(amount||0),g.balance,String(referenceId||""),String(note||"")]);
}

function _phase2GovernmentMove_(delta, type, note, referenceId){
  const balance = _changeGovernmentBalance_(delta);
  _appendGovernmentLedger_(type, delta, referenceId, note);
  return balance;
}

function api_phase2Dashboard(adminPass){
  _assertAdminPassValue_(adminPass);
  const g = _getGovernmentAccount_();
  const users = api_adminListUsersWithPin(adminPass);
  return {
    version:"2.0",
    government:{accountId:g.accountId,accountName:g.accountName,balance:g.balance},
    applications:api_adminListCompanyApplications(adminPass,"PENDING"),
    ranking:api_companyRanking(),
    users:users,
    studentRanking:users.filter(u=>u.isActive)
      .slice().sort((a,b)=>b.balance-a.balance||a.userId.localeCompare(b.userId))
      .map((u,i)=>({rank:i+1,userId:u.userId,name:u.name,balance:u.balance})),
    shops:api_adminListAllShops(adminPass),
    ledger:api_phase2GovernmentLedger(adminPass,50)
  };
}

function api_phase2GovernmentLedger(adminPass, limit){
  _assertAdminPassValue_(adminPass);
  limit=Math.min(Math.max(Number(limit||50),1),300);
  const sh=_sheetByNameOrThrow_(SHEETS.GOVERNMENT_LEDGER || "GovernmentLedger");
  const v=sh.getDataRange().getValues();
  if(v.length<2) return [];
  const m=_headerMap_(v[0]), ix=n=>m.idx(n);
  return v.slice(1).filter(r=>r[ix("ledgerid")]).slice(-limit).reverse().map(r=>({
    ledgerId:String(r[ix("ledgerid")]||""),at:String(r[ix("at")]||""),type:String(r[ix("type")]||""),
    amount:Number(r[ix("amount")]||0),balanceAfter:Number(r[ix("balanceafter")]||0),
    referenceId:String(r[ix("referenceid")]||""),note:String(r[ix("note")]||"")
  }));
}

function api_adminMoneyFlow(adminPass,limit){
  _assertAdminPassValue_(adminPass);
  limit=Math.min(Math.max(Number(limit||100),1),300);
  const users=api_adminListUsersWithPin(adminPass),shops=api_listShops(),government=_getGovernmentAccount_(),transactions=api_adminListTx(limit);
  const byType={};
  transactions.forEach(x=>{const type=String(x.type||"OTHER").toUpperCase();if(!byType[type])byType[type]={type,count:0,total:0};byType[type].count++;byType[type].total+=Number(x.amount||0);});
  return {
    summary:{userBalance:users.filter(x=>x.isActive).reduce((s,x)=>s+Number(x.balance||0),0),shopBalance:shops.reduce((s,x)=>s+Number(x.balance||0),0),governmentBalance:Number(government.balance||0),userCount:users.filter(x=>x.isActive).length,shopCount:shops.length},
    byType:Object.keys(byType).map(k=>byType[k]).sort((a,b)=>b.count-a.count),
    transactions
  };
}

function api_submitCompanyApplicationV2(payload){
  payload=payload||{};
  const presidentUserId=String(payload.presidentUserId||"").trim().toUpperCase();
  _assertPin_(presidentUserId,payload.pin);
  if(!/^\d{4,12}$/.test(String(payload.companyPass||"").trim())) throw new Error("会社PASSは4〜12桁の数字で入力してください");
  return api_submitCompanyApplication(payload);
}

function api_phase2CollectTax(adminPass,taxAmount,reason,mode){
  _assertAdminPassValue_(adminPass);
  const result=api_adminCollectTaxAll(adminPass,taxAmount,reason,mode);
  if(result.totalCollected>0) result.governmentBalance=_phase2GovernmentMove_(result.totalCollected,"TAX_IN",result.reason,"");
  else result.governmentBalance=_getGovernmentAccount_().balance;
  return result;
}

function api_phase2DistributeUsers(adminPass,amount,reason){
  _assertAdminPassValue_(adminPass);
  const result=api_adminDistributeAll(adminPass,amount,reason);
  result.governmentBalance=_phase2GovernmentMove_(-result.total,"DISTRIBUTION_OUT",result.reason,"");
  return result;
}

function api_phase2DistributeShops(adminPass,amount,reason){
  _assertAdminPassValue_(adminPass);
  const result=api_adminDistributeAllShops(adminPass,amount,reason);
  result.governmentBalance=_phase2GovernmentMove_(-result.total,"SHOP_SUPPORT_OUT",result.reason,"");
  return result;
}

function api_phase2UserInterest(adminPass,ratePercent,reason){
  _assertAdminPassValue_(adminPass);
  const rate=Number(ratePercent)/100;
  if(!Number.isFinite(rate)||rate<0||rate>1) throw new Error("利率は0〜100%で入力してください");
  const result=api_adminInterestOnBalance(adminPass,rate,reason||"児童口座利息",true);
  return result;
}

function api_phase2ShopInterest(adminPass,ratePercent,reason){
  _assertAdminPassValue_(adminPass);
  const rate=Number(ratePercent)/100;
  if(!Number.isFinite(rate)||rate<0||rate>1) throw new Error("利率は0〜100%で入力してください");
  const result=api_adminInterestOnShopBalance(adminPass,rate,reason||"会社口座利息");
  result.governmentBalance=_phase2GovernmentMove_(-result.totalInterest,"SHOP_INTEREST_OUT",reason||"会社口座利息","");
  return result;
}

function api_phase2AdjustGovernment(adminPass,amount,reason){
  _assertAdminPassValue_(adminPass);
  amount=Number(amount); reason=String(reason||"").trim();
  if(!Number.isFinite(amount)||amount===0) throw new Error("0以外の調整額を入力してください");
  if(!reason) throw new Error("調整理由を入力してください");
  const balance=_phase2GovernmentMove_(amount,"MANUAL_ADJUST",reason,"");
  return {ok:true,amount,balance,reason};
}

function _latestSnapshotMap_(){
  const sh=_getCompanySnapshotsSheet_(), v=sh.getDataRange().getValues();
  if(v.length<2) return {};
  const m=_headerMap_(v[0]), cShop=m.idx("shopid"), cBal=m.idx("balance"), cAt=m.idx("snapshotat"), cGrowth=m.idx("growthamount"), cRate=m.idx("growthrate"), cValue=m.idx("valuecreated"), out={};
  v.slice(1).forEach(r=>{
    const id=String(r[cShop]||"").trim().toUpperCase(); if(!id) return;
    const at=new Date(r[cAt]).getTime()||0;
    if(!out[id]||at>=out[id].time) out[id]={balance:Number(r[cBal]||0),time:at,growthAmount:cGrowth<0?0:Number(r[cGrowth]||0),growthRate:cRate<0?0:Number(r[cRate]||0),valueCreated:cValue<0?0:Number(r[cValue]||0)};
  });
  return out;
}

function _companyValueCreatedSince_(shopId,sinceMs){
  const sh=_getTxSheet_(), v=sh.getDataRange().getValues(); if(v.length<2) return 0;
  const m=_headerMap_(v[0]), cShop=m.idx("shopid"),cType=m.idx("type"),cAmount=m.idx("amount"),cAt=m.idx("at");
  return v.slice(1).reduce((sum,r)=>{
    if(String(r[cShop]||"").trim().toUpperCase()!==shopId) return sum;
    const t=new Date(r[cAt]).getTime()||0; if(t<sinceMs) return sum;
    const type=String(r[cType]||"").toUpperCase();
    return type==="PAY" ? sum+Math.max(0,Number(r[cAmount]||0)) : sum;
  },0);
}

/** 全会社分の価値創出を、Txシート1回の読込で集計する。 */
function _companyValueCreatedMap_(shops, latest){
  const out={};
  shops.forEach(s=>{out[s.shopId]=0;});
  const sh=_getTxSheet_(), v=sh.getDataRange().getValues();
  if(v.length<2) return out;
  const m=_headerMap_(v[0]), cShop=m.idx("shopid"),cType=m.idx("type"),cAmount=m.idx("amount"),cAt=m.idx("at");
  v.slice(1).forEach(r=>{
    const shopId=String(r[cShop]||"").trim().toUpperCase();
    if(!Object.prototype.hasOwnProperty.call(out,shopId)) return;
    const sinceMs=latest[shopId]?latest[shopId].time:0;
    const t=new Date(r[cAt]).getTime()||0;
    if(t<sinceMs || String(r[cType]||"").toUpperCase()!=="PAY") return;
    out[shopId]+=Math.max(0,Number(r[cAmount]||0));
  });
  return out;
}

function api_adminSaveCompanySnapshot(adminPass){
  _assertAdminPassValue_(adminPass);
  const previous=_latestSnapshotMap_(), shops=api_listShops(), values=_companyValueCreatedMap_(shops,previous), at=_fmtJst_(new Date()), rows=[];
  shops.forEach(s=>{
    const p=previous[s.shopId], prev=p?p.balance:s.balance, growth=s.balance-prev;
    const rate=prev===0 ? (growth>0?100:0) : growth/Math.abs(prev)*100;
    const value=values[s.shopId]||0;
    rows.push([at,s.shopId,s.shopName,s.balance,prev,growth,rate,value]);
  });
  if(rows.length){ const sh=_getCompanySnapshotsSheet_(); sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows); }
  return {ok:true,at,count:rows.length,ranking:api_companyRanking()};
}

function api_companyRanking(){
  const latest=_latestSnapshotMap_();
  const shops=api_listShops(), values=_companyValueCreatedMap_(shops,latest);
  return shops.map(s=>{
    const p=latest[s.shopId], prev=p?p.balance:s.balance, liveGrowth=s.balance-prev;
    const settled=!!p && liveGrowth===0;
    const growth=settled?p.growthAmount:liveGrowth;
    const growthRate=settled?p.growthRate:(prev===0?(growth>0?100:0):growth/Math.abs(prev)*100);
    const valueCreated=settled?p.valueCreated:(values[s.shopId]||0);
    return {shopId:s.shopId,shopName:s.shopName,balance:s.balance,previousBalance:prev,growthAmount:growth,growthRate,valueCreated};
  }).sort((a,b)=>b.balance-a.balance||b.valueCreated-a.valueCreated||a.shopId.localeCompare(b.shopId))
    .map((x,i)=>Object.assign({rank:i+1},x));
}
