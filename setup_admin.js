/* =========================================================
 * ClassPay Phase 2.2: 初期設定・登録・認証情報管理
 * ========================================================= */

function _ensureCoreSheets_(){
  _ensureColumns_(_ensureSheetWithHeader_(SHEETS.USERS,["userId","name","balance","isActive","pin"]),["userId","name","balance","isActive","pin"]);
  _ensureColumns_(_ensureSheetWithHeader_(SHEETS.SHOPS,["shopId","shopName","balance","isActive","companyPass","URL","totalShares","basePrice","k","priceMin","priceMax","stockActive","spread"]),["shopId","shopName","balance","isActive","companyPass","URL","totalShares","basePrice","k","priceMin","priceMax","stockActive","spread"]);
  _ensureColumns_(_ensureSheetWithHeader_(SHEETS.TX,["txId","at","type","userId","userName","shopId","shopName","amount","status","note","meta"]),["txId","at","type","userId","userName","shopId","shopName","amount","status","note","meta"]);
  _ensureColumns_(_ensureSheetWithHeader_(SHEETS.CONFIG,["key","value"]),["key","value"]);
  setupClassPayPhase2();
  setupClassPayV32();
  _ensureColumns_(_getCompanyMembersSheet_(),["shopId","userId","role","isActive","joinedAt","leftAt","leaveReason","leftBy"]);
  _ensureColumns_(_getCompanyApplicationsSheet_(),["applicationId","at","companyName","presidentUserId","memberUserIds","companyPass","activity","status","reviewedAt","reviewNote","shopId"]);
  _ensureColumns_(_ensureSheetWithHeader_(SHEETS.RETIREMENT_APPLICATIONS||"RetirementApplications",["applicationId","submittedAt","shopId","shopName","userId","userName","reason","status","reviewedAt","reviewNote","reviewedBy","governmentStatus"]),["applicationId","submittedAt","shopId","shopName","userId","userName","reason","status","reviewedAt","reviewNote","reviewedBy","governmentStatus"]);
  _ensureColumns_(_getGovernmentSheet_(),["accountId","accountName","balance","updatedAt"]);
  _ensureColumns_(_getCompanySnapshotsSheet_(),["snapshotAt","shopId","shopName","balance","previousBalance","growthAmount","growthRate","valueCreated"]);
  _ensureColumns_(SpreadsheetApp.getActive().getSheetByName(SHEETS.HOLDINGS||"Holdings"),["userId","shopId","shares","updatedAt"]);
  _ensureColumns_(SpreadsheetApp.getActive().getSheetByName(SHEETS.GOVERNMENT_LEDGER||"GovernmentLedger"),["ledgerId","at","type","amount","balanceAfter","referenceId","note"]);
  _ensureColumns_(_ensureSheetWithHeader_(SHEETS.MIGRATION_LOG||"MigrationLog",["migrationId","at","sourceSpreadsheetId","sourceVersion","targetVersion","status","backupSpreadsheetId","importedSheets","deletedSheets","validation","message"]),["migrationId","at","sourceSpreadsheetId","sourceVersion","targetVersion","status","backupSpreadsheetId","importedSheets","deletedSheets","validation","message"]);
}

function _ensureColumns_(sh,required){
  var width=Math.max(1,sh.getLastColumn()),headers=sh.getRange(1,1,1,width).getValues()[0],lower=headers.map(function(x){return String(x||"").trim().toLowerCase();});
  required.forEach(function(name){if(lower.indexOf(String(name).toLowerCase())<0){sh.getRange(1,sh.getLastColumn()+1).setValue(name);lower.push(String(name).toLowerCase());}});
  sh.setFrozenRows(1);return sh;
}

function api_adminEnsureAllSheets(adminPass){
  _assertAdminPassValue_(adminPass);
  _ensureCoreSheets_();
  _setConfigValue_("CLASS_PAY_VERSION","4.0.0");
  return api_setupStatus();
}

function _setConfigValue_(key,value){
  var sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.CONFIG);
  if(!sh) sh=_ensureSheetWithHeader_(SHEETS.CONFIG,["key","value"]);
  var v=sh.getDataRange().getValues();
  for(var i=1;i<v.length;i++){
    if(String(v[i][0]||"").trim()===key){sh.getRange(i+1,2).setValue(value);return;}
  }
  sh.appendRow([key,value]);
}

function api_setupStatus(){
  var ss=SpreadsheetApp.getActive(), config=ss.getSheetByName(SHEETS.CONFIG);
  var pass=config ? String(getConfig_("ADMIN_PASS","")).trim() : "";
  return {initialized:!!pass,appName:config?String(getConfig_("APP_NAME","ClassPay")):"ClassPay",version:config?String(getConfig_("CLASS_PAY_VERSION","4.0.0")):"4.0.0",missing:["Users","Shops","Tx","Config","Holdings","CompanyMembers","CompanyApplications","RetirementApplications","WeeklyReports","RecruitmentPostings","EmploymentApplications","CompanyAnnouncements","Government","CompanySnapshots","GovernmentLedger","ProductCatalog","ProductOrders","CompanyContracts","WeeklySettlements","MigrationLog"].filter(function(n){return !ss.getSheetByName(n);})};
}

/** URLを受け取った人が先に初期化しないよう、公開前に教師が直ちに設定する一度限りの処理 */
function api_initialSetup(payload){
  payload=payload||{};
  var status=api_setupStatus();
  if(status.initialized) throw new Error("初期設定は完了済みです。管理パスワードでログインしてください");
  var adminPass=String(payload.adminPass||"").trim();
  var confirmPass=String(payload.confirmPass||"").trim();
  if(adminPass.length<4) throw new Error("管理パスワードは4文字以上にしてください");
  if(adminPass!==confirmPass) throw new Error("確認用パスワードが一致しません");
  _ensureCoreSheets_();
  _setConfigValue_("APP_NAME",String(payload.appName||"ClassPay").trim()||"ClassPay");
  _setConfigValue_("ADMIN_PASS",adminPass);
  _setConfigValue_("BASE_URL",String(payload.baseUrl||"").trim());
  _setConfigValue_("INTEREST_ROUND",String(payload.interestRound||"FLOOR").toUpperCase());
  _setConfigValue_("STOCK_SPREAD",Number(payload.stockSpread||10)/100);
  _setConfigValue_("MIN_AMOUNT",Number(payload.minAmount||1));
  _setConfigValue_("MAX_AMOUNT",Number(payload.maxAmount||500));
  _setConfigValue_("SUPPORT_PER_POINT",Number(payload.supportPerPoint||10));
  _setConfigValue_("STOCK_WEEKLY_LIMIT",Number(payload.stockWeeklyLimit||3));
  _setConfigValue_("STOCK_HOLD_DAYS",Number(payload.stockHoldDays||7));
  _setConfigValue_("STOCK_MAX_OWNERSHIP_PERCENT",Number(payload.stockMaxOwnershipPercent||20));
  _setConfigValue_("STOCK_ACTIVITY_WEIGHT",Number(payload.stockActivityWeight||20)/100);
  _setConfigValue_("PUBLICITY_MODE","GOVERNMENT");
  _setConfigValue_("PRIVATE_PROVIDER_NAME","");
  _setConfigValue_("RECRUIT_MIN_BID",Number(payload.recruitMinBid||10));
  _setConfigValue_("RECRUIT_SLOTS",Number(payload.recruitSlots||3));
  _setConfigValue_("RECRUIT_DAYS",Number(payload.recruitDays||7));
  _setConfigValue_("ANNOUNCEMENT_FEE",Number(payload.announcementFee||5));
  _setConfigValue_("ANNOUNCEMENT_DAYS",Number(payload.announcementDays||7));
  _setConfigValue_("STOCK_DAILY_LIMIT",Number(payload.stockDailyLimit||2));
  _setConfigValue_("STOCK_REASON_REQUIRED","TRUE");
  _setConfigValue_("WEEKLY_TAX_AMOUNT",Number(payload.weeklyTaxAmount||0));
  _setConfigValue_("WEEKLY_USER_INTEREST_RATE",Number(payload.weeklyUserInterestRate||5));
  _setConfigValue_("WEEKLY_SHOP_INTEREST_RATE",Number(payload.weeklyShopInterestRate||0));
  _setConfigValue_("CLASS_PAY_VERSION","4.0.0");
  return {ok:true};
}

function api_adminSettings(adminPass){
  _assertAdminPassValue_(adminPass);
  return {appName:getConfig_("APP_NAME","ClassPay"),baseUrl:getConfig_("BASE_URL",""),interestRound:getConfig_("INTEREST_ROUND","FLOOR"),stockSpread:Number(getConfig_("STOCK_SPREAD",0.1))*100,minAmount:Number(getConfig_("MIN_AMOUNT",1)),maxAmount:Number(getConfig_("MAX_AMOUNT",500)),supportPerPoint:Number(getConfig_("SUPPORT_PER_POINT",10)),stockWeeklyLimit:Number(getConfig_("STOCK_WEEKLY_LIMIT",3)),stockDailyLimit:Number(getConfig_("STOCK_DAILY_LIMIT",2)),stockHoldDays:Number(getConfig_("STOCK_HOLD_DAYS",7)),stockMaxOwnershipPercent:Number(getConfig_("STOCK_MAX_OWNERSHIP_PERCENT",20)),stockActivityWeight:Number(getConfig_("STOCK_ACTIVITY_WEIGHT",0.2))*100,publicityMode:String(getConfig_("PUBLICITY_MODE","GOVERNMENT")),privateProviderName:String(getConfig_("PRIVATE_PROVIDER_NAME","")),recruitMinBid:Number(getConfig_("RECRUIT_MIN_BID",10)),recruitSlots:Number(getConfig_("RECRUIT_SLOTS",3)),recruitDays:Number(getConfig_("RECRUIT_DAYS",7)),announcementFee:Number(getConfig_("ANNOUNCEMENT_FEE",5)),announcementDays:Number(getConfig_("ANNOUNCEMENT_DAYS",7)),weeklyTaxAmount:Number(getConfig_("WEEKLY_TAX_AMOUNT",0)),weeklyUserInterestRate:Number(getConfig_("WEEKLY_USER_INTEREST_RATE",5)),weeklyShopInterestRate:Number(getConfig_("WEEKLY_SHOP_INTEREST_RATE",0)),classPayVersion:String(getConfig_("CLASS_PAY_VERSION","4.0.0"))};
}

function api_adminSaveSettings(adminPass,payload){
  _assertAdminPassValue_(adminPass);payload=payload||{};
  _setConfigValue_("APP_NAME",String(payload.appName||"ClassPay").trim()||"ClassPay");
  _setConfigValue_("BASE_URL",String(payload.baseUrl||"").trim());
  _setConfigValue_("INTEREST_ROUND",String(payload.interestRound||"FLOOR").toUpperCase());
  _setConfigValue_("STOCK_SPREAD",Number(payload.stockSpread||0)/100);
  _setConfigValue_("MIN_AMOUNT",Number(payload.minAmount||1));
  _setConfigValue_("MAX_AMOUNT",Number(payload.maxAmount||500));
  _setConfigValue_("SUPPORT_PER_POINT",Math.max(0,Number(payload.supportPerPoint||0)));
  _setConfigValue_("STOCK_WEEKLY_LIMIT",Math.max(1,Math.floor(Number(payload.stockWeeklyLimit||3))));
  _setConfigValue_("STOCK_DAILY_LIMIT",Math.max(1,Math.floor(Number(payload.stockDailyLimit||2))));
  _setConfigValue_("STOCK_HOLD_DAYS",Math.max(0,Math.floor(Number(payload.stockHoldDays||7))));
  _setConfigValue_("STOCK_MAX_OWNERSHIP_PERCENT",Math.max(1,Math.min(100,Number(payload.stockMaxOwnershipPercent||20))));
  _setConfigValue_("STOCK_ACTIVITY_WEIGHT",Math.max(0,Math.min(100,Number(payload.stockActivityWeight||20)))/100);
  _setConfigValue_("PUBLICITY_MODE",["GOVERNMENT","DISABLED","PRIVATE"].includes(String(payload.publicityMode||"GOVERNMENT").toUpperCase())?String(payload.publicityMode).toUpperCase():"GOVERNMENT");
  _setConfigValue_("PRIVATE_PROVIDER_NAME",String(payload.privateProviderName||"").trim());
  _setConfigValue_("RECRUIT_MIN_BID",Math.max(0,Math.floor(Number(payload.recruitMinBid||0))));
  _setConfigValue_("RECRUIT_SLOTS",Math.max(1,Math.floor(Number(payload.recruitSlots||3))));
  _setConfigValue_("RECRUIT_DAYS",Math.max(1,Math.floor(Number(payload.recruitDays||7))));
  _setConfigValue_("ANNOUNCEMENT_FEE",Math.max(0,Math.floor(Number(payload.announcementFee||0))));
  _setConfigValue_("ANNOUNCEMENT_DAYS",Math.max(1,Math.floor(Number(payload.announcementDays||7))));
  _setConfigValue_("WEEKLY_TAX_AMOUNT",Math.max(0,Math.floor(Number(payload.weeklyTaxAmount||0))));
  _setConfigValue_("WEEKLY_USER_INTEREST_RATE",Math.max(0,Math.min(100,Number(payload.weeklyUserInterestRate||0))));
  _setConfigValue_("WEEKLY_SHOP_INTEREST_RATE",Math.max(0,Math.min(100,Number(payload.weeklyShopInterestRate||0))));
  return {ok:true,settings:api_adminSettings(adminPass)};
}

function _writeByHeader_(sh,rowNumber,obj,required){
  var width=sh.getLastColumn(),headers=sh.getRange(1,1,1,width).getValues()[0],m=_headerMap_(headers),row=rowNumber?sh.getRange(rowNumber,1,1,width).getValues()[0]:new Array(width).fill("");
  required.forEach(function(k){if(m.idx(k)<0)throw new Error(sh.getName()+"ヘッダに "+k+" が必要です");});
  Object.keys(obj).forEach(function(k){var c=m.idx(k);if(c>=0)row[c]=obj[k];});
  if(rowNumber)sh.getRange(rowNumber,1,1,width).setValues([row]);else sh.appendRow(row);
}

function api_adminSaveUser(adminPass,payload){
  _assertAdminPassValue_(adminPass);payload=payload||{};
  var id=String(payload.userId||"").trim().toUpperCase(),name=String(payload.name||"").trim(),pin=_normalizePin_(payload.pin),balance=Number(payload.balance||0),active=payload.isActive!==false;
  if(!id||!name)throw new Error("児童IDと名前を入力してください");if(!/^\d{4}$/.test(pin))throw new Error("PINは4桁の数字にしてください");if(!Number.isFinite(balance))throw new Error("残高が不正です");
  var sh=_getUsersSheet_(),v=sh.getDataRange().getValues(),m=_headerMap_(v[0]),c=m.idx("userid"),rowNo=0;
  for(var i=1;i<v.length;i++)if(String(v[i][c]||"").trim().toUpperCase()===id){rowNo=i+1;break;}
  _writeByHeader_(sh,rowNo,{userid:id,name:name,balance:balance,isactive:active,pin:pin},["userid","name","balance","isactive","pin"]);
  return {ok:true,userId:id,updated:!!rowNo};
}

function api_adminBulkUsers(adminPass,text,defaultBalance){
  _assertAdminPassValue_(adminPass);var lines=String(text||"").split(/\r?\n/).map(function(x){return x.trim();}).filter(Boolean),results=[];
  lines.forEach(function(line,index){try{var p=line.split(/[\t,]/).map(function(x){return x.trim();});if(p.length<3)throw new Error("ID,名前,PINが必要です");results.push(api_adminSaveUser(adminPass,{userId:p[0],name:p[1],pin:p[2],balance:p[3]===""||p[3]===undefined?Number(defaultBalance||0):Number(p[3]),isActive:true}));}catch(e){results.push({ok:false,line:index+1,error:e.message});}});
  return {count:results.filter(function(x){return x.ok;}).length,failed:results.filter(function(x){return !x.ok;}),results:results};
}

function _endCompanyMemberRows_(shopId,userId,reason,leftBy,allowPresident){
  var sh=_getCompanyMembersSheet_(),v=sh.getDataRange().getValues();if(v.length<2)return 0;
  var m=_headerMap_(v[0]),cShop=m.idx("shopid"),cUser=m.idx("userid"),cRole=m.idx("role"),cActive=m.idx("isactive"),cLeft=m.idx("leftat"),cReason=m.idx("leavereason"),cBy=m.idx("leftby"),count=0,at=_fmtJst_(new Date());
  for(var i=1;i<v.length;i++){
    if(String(v[i][cShop]||"").trim().toUpperCase()!==shopId)continue;
    if(userId&&String(v[i][cUser]||"").trim().toUpperCase()!==userId)continue;
    if(cActive>=0&&!_isTrue_(v[i][cActive]))continue;
    if(!allowPresident&&String(v[i][cRole]||"").trim().toUpperCase()==="PRESIDENT")throw new Error("社長は先に別の児童へ変更するか、会社を停止してください");
    if(cActive>=0)sh.getRange(i+1,cActive+1).setValue(false);
    if(cLeft>=0)sh.getRange(i+1,cLeft+1).setValue(at);
    if(cReason>=0)sh.getRange(i+1,cReason+1).setValue(String(reason||"退社"));
    if(cBy>=0)sh.getRange(i+1,cBy+1).setValue(String(leftBy||""));
    count++;
  }
  return count;
}

function _deactivateCompanyMembers_(shopId,reason,leftBy){
  return _endCompanyMemberRows_(shopId,"",reason||"所属更新",leftBy||"ADMIN",true);
}

function _appendCompanyMember_(shopId,userId,role,joinedAt){
  var sh=_getCompanyMembersSheet_();
  _writeByHeader_(sh,0,{shopid:shopId,userid:userId,role:role,isactive:true,joinedat:joinedAt||_fmtJst_(new Date()),leftat:"",leavereason:"",leftby:""},["shopid","userid","role","isactive","joinedat"]);
}

function _reconcileCompanyMembers_(shopId,president,members){
  var desired={};
  if(president)desired[president]="PRESIDENT";
  Array.from(new Set(members||[])).forEach(function(id){if(id&&id!==president)desired[id]="MEMBER";});
  var sh=_getCompanyMembersSheet_(),v=sh.getDataRange().getValues(),m=_headerMap_(v[0]),cShop=m.idx("shopid"),cUser=m.idx("userid"),cRole=m.idx("role"),cActive=m.idx("isactive"),activeByUser={},kept={};
  v.slice(1).forEach(function(r){
    if(String(r[cShop]||"").trim().toUpperCase()!==shopId||(cActive>=0&&!_isTrue_(r[cActive])))return;
    var id=String(r[cUser]||"").trim().toUpperCase(),role=String(r[cRole]||"MEMBER").trim().toUpperCase();
    if(!activeByUser[id])activeByUser[id]=[];activeByUser[id].push(role);
  });
  Object.keys(activeByUser).forEach(function(id){var roles=activeByUser[id];if(desired[id]&&roles.length===1&&roles[0]===desired[id])kept[id]=true;else _endCompanyMemberRows_(shopId,id,"所属更新","ADMIN",true);});
  Object.keys(desired).forEach(function(id){if(!kept[id])_appendCompanyMember_(shopId,id,desired[id],_fmtJst_(new Date()));});
}

function api_adminSaveCompany(adminPass,payload){
  _assertAdminPassValue_(adminPass);payload=payload||{};
  var shopId=String(payload.shopId||"").trim().toUpperCase()||_nextShopId_(),name=String(payload.shopName||"").trim(),pass=String(payload.companyPass||"").trim(),balance=Number(payload.balance||0),active=payload.isActive!==false;
  if(!name)throw new Error("会社名を入力してください");if(!/^\d{4,12}$/.test(pass))throw new Error("会社PASSは4〜12桁の数字にしてください");if(!Number.isFinite(balance))throw new Error("会社残高が不正です");
  var sh=_getShopsSheet_(),v=sh.getDataRange().getValues(),m=_headerMap_(v[0]),c=m.idx("shopid"),rowNo=0;
  for(var i=1;i<v.length;i++)if(String(v[i][c]||"").trim().toUpperCase()===shopId){rowNo=i+1;break;}
  _writeByHeader_(sh,rowNo,{shopid:shopId,shopname:name,balance:balance,isactive:active,companypass:pass,stockactive:payload.stockActive!==false},["shopid","shopname","balance","isactive","companypass"]);
  var president=String(payload.presidentUserId||"").trim().toUpperCase(),members=(payload.memberUserIds||[]).map(function(x){return String(x||"").trim().toUpperCase();}).filter(Boolean);
  if(president&&!_findUser_(president))throw new Error("社長の児童が見つかりません");Array.from(new Set(members)).forEach(function(id){if(!_findUser_(id))throw new Error("社員が見つかりません："+id);});
  if(active)_reconcileCompanyMembers_(shopId,president,members);else _deactivateCompanyMembers_(shopId,"会社停止","ADMIN");
  return {ok:true,shopId:shopId,updated:!!rowNo};
}

function api_adminDeleteCompany(adminPass,shopId,reason){
  _assertAdminPassValue_(adminPass);
  return lockRun_(function(){
    shopId=String(shopId||"").trim().toUpperCase();reason=String(reason||"会社を停止").trim()||"会社を停止";
    var s=_findShop_(shopId);if(!s)throw new Error("会社が見つかりません");
    var sh=_getShopsSheet_(),v=sh.getDataRange().getValues(),m=_headerMap_(v[0]),cActive=m.idx("isactive");
    if(cActive<0)throw new Error("Shopsヘッダに isActive が必要です");
    sh.getRange(s.row,cActive+1).setValue(false);
    var ended=_deactivateCompanyMembers_(shopId,reason,"ADMIN");
    return {ok:true,shopId:shopId,endedMemberships:ended,message:"会社を停止しました。取引・株・所属履歴は保持されます"};
  });
}

function api_adminLeaveCompany(adminPass,userId,shopId,reason){
  _assertAdminPassValue_(adminPass);
  return lockRun_(function(){
    userId=String(userId||"").trim().toUpperCase();shopId=String(shopId||"").trim().toUpperCase();
    var count=_endCompanyMemberRows_(shopId,userId,reason||"管理者による退社","ADMIN",false);
    if(!count)throw new Error("有効な所属が見つかりません");
    return {ok:true,userId:userId,shopId:shopId};
  });
}

function api_adminCredentials(adminPass){
  _assertAdminPassValue_(adminPass);var users=api_adminListUsersWithPin(adminPass),sh=_getShopsSheet_(),v=sh.getDataRange().getValues(),shops=[];
  if(v.length>1){var m=_headerMap_(v[0]),cId=m.idx("shopid"),cName=m.idx("shopname"),cPass=m.idx("companypass"),cBal=m.idx("balance"),cActive=m.idx("isactive");v.slice(1).forEach(function(r){if(!r[cId])return;shops.push({shopId:String(r[cId]),shopName:String(r[cName]||""),companyPass:String(r[cPass]||""),balance:Number(r[cBal]||0),isActive:cActive<0?true:_isTrue_(r[cActive])});});}
  return {users:users,shops:shops};
}

function api_adminCompanyDetail(adminPass,shopId){
  _assertAdminPassValue_(adminPass);shopId=String(shopId||"").trim().toUpperCase();var s=_findShop_(shopId);if(!s)throw new Error("会社が見つかりません");
  var members=_membershipHistoryRows_("").filter(function(x){return x.shopId===shopId&&x.isActive;}),president=members.find(function(x){return x.role==="PRESIDENT";});
  return {shopId:s.shopId,shopName:s.shopName,balance:s.balance,isActive:s.active,companyPass:s.companyPass,presidentUserId:president?president.userId:"",memberUserIds:members.filter(function(x){return x.role!=="PRESIDENT";}).map(function(x){return x.userId;})};
}

function api_adminListAllShops(adminPass){
  _assertAdminPassValue_(adminPass);var sh=_getShopsSheet_(),v=sh.getDataRange().getValues();if(v.length<2)return [];
  var m=_headerMap_(v[0]),cId=m.idx("shopid"),cName=m.idx("shopname"),cBal=m.idx("balance"),cActive=m.idx("isactive");
  return v.slice(1).map(function(r){var id=String(r[cId]||"").trim().toUpperCase();return id?{shopId:id,shopName:String(r[cName]||""),balance:Number(r[cBal]||0),active:cActive<0?true:_isTrue_(r[cActive])}:null;}).filter(Boolean).sort(function(a,b){return (a.active===b.active?0:(a.active?-1:1))||a.shopId.localeCompare(b.shopId);});
}

function api_adminChangeAdminPassV2(adminPass,newPass,confirmPass){
  _assertAdminPassValue_(adminPass);newPass=String(newPass||"").trim();if(newPass.length<4)throw new Error("新しい管理パスワードは4文字以上にしてください");if(newPass!==String(confirmPass||"").trim())throw new Error("確認用パスワードが一致しません");_setConfigValue_("ADMIN_PASS",newPass);return {ok:true};
}
