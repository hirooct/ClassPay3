/* =========================================================
 * ClassPay 4.0: 新しいスプレッドシートへのデータ引き継ぎ
 * 旧スプレッドシートは読み取り専用で扱い、変更しない。
 * ========================================================= */

const CLASS_PAY_V4_VERSION = "4.0.0";

function _cpV4RequiredDataSheets_(){
  return [
    "Users","Shops","Tx","Holdings","CompanyMembers","CompanyApplications",
    "RetirementApplications","WeeklyReports","RecruitmentPostings",
    "EmploymentApplications","CompanyAnnouncements","Government",
    "CompanySnapshots","GovernmentLedger","ProductCatalog","ProductOrders",
    "CompanyContracts","WeeklySettlements"
  ];
}

/** 4.0では使用しないシート。ここにない未知のシートは絶対に自動削除しない。 */
function _cpV4ObsoleteSheets_(){
  return [
    {name:"RuleProposals",reason:"現在停止中の国会・ルール提案機能"},
    {name:"RuleVotes",reason:"現在停止中の国会・投票機能"},
    {name:"UpdateHistory",reason:"Apps Script API方式の旧Updater履歴"}
  ];
}

function _cpV4SheetKey_(value){
  return String(value||"").trim().toLowerCase().replace(/[\s_\-]/g,"");
}

function _cpV4SpreadsheetId_(input){
  var text=String(input||"").trim(),match=text.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  if(match) return match[1];
  if(/^[a-zA-Z0-9_-]{20,}$/.test(text)) return text;
  throw new Error("旧ClassPayのスプレッドシートURLまたはIDを入力してください");
}

function _cpV4OpenSource_(input){
  var id=_cpV4SpreadsheetId_(input),current=SpreadsheetApp.getActive().getId();
  if(id===current) throw new Error("現在のスプレッドシートは引き継ぎ元に指定できません");
  try{return SpreadsheetApp.openById(id);}catch(e){
    throw new Error("旧スプレッドシートを開けません。IDと閲覧権限を確認してください");
  }
}

function _cpV4ConfigMap_(ss){
  var sh=ss.getSheetByName("Config"),out={};
  if(!sh||sh.getLastRow()<2) return out;
  sh.getRange(2,1,sh.getLastRow()-1,Math.min(2,sh.getLastColumn())).getValues().forEach(function(r){
    var key=String(r[0]||"").trim();if(key)out[key]=r[1];
  });
  return out;
}

function _cpV4Metric_(ss,name){
  var sh=ss.getSheetByName(name),result={rows:0,balance:0};
  if(!sh||sh.getLastRow()<2) return result;
  var values=sh.getDataRange().getValues(),headers=values[0].map(_cpV4SheetKey_),balanceCol=headers.indexOf("balance");
  result.rows=values.length-1;
  if(balanceCol>=0)result.balance=values.slice(1).reduce(function(sum,row){return sum+(Number(row[balanceCol])||0);},0);
  return result;
}

function _cpV4Summary_(ss){
  var names=["Users","Shops","Tx","Holdings","CompanyMembers","Government"],out={};
  names.forEach(function(name){out[name]=_cpV4Metric_(ss,name);});
  return out;
}

function _cpV4MigrationLogRows_(){
  var sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MIGRATION_LOG||"MigrationLog");
  if(!sh||sh.getLastRow()<2)return [];
  var values=sh.getDataRange().getValues(),headers=values[0];
  return values.slice(1).reverse().slice(0,20).map(function(row){
    var item={};headers.forEach(function(h,i){item[String(h)]=row[i];});return item;
  });
}

function api_adminMigrationStatusV4(adminPass){
  _assertAdminPassValue_(adminPass);
  _ensureCoreSheets_();
  var ss=SpreadsheetApp.getActive();
  return {
    version:CLASS_PAY_V4_VERSION,
    obsolete:_cpV4ObsoleteSheets_().map(function(x){return {name:x.name,reason:x.reason,present:!!ss.getSheetByName(x.name)};}),
    history:_cpV4MigrationLogRows_()
  };
}

function api_adminMigrationPreviewV4(adminPass,sourceInput){
  _assertAdminPassValue_(adminPass);
  _ensureCoreSheets_();
  var source=_cpV4OpenSource_(sourceInput),sourceConfig=_cpV4ConfigMap_(source),target=SpreadsheetApp.getActive();
  var sourceNames=source.getSheets().map(function(sh){return sh.getName();});
  var dataSheets=_cpV4RequiredDataSheets_().map(function(name){
    var sh=source.getSheetByName(name);
    return {name:name,present:!!sh,rows:sh?Math.max(0,sh.getLastRow()-1):0};
  });
  var known=_cpV4RequiredDataSheets_().concat(["Config","MigrationLog"]).concat(_cpV4ObsoleteSheets_().map(function(x){return x.name;}));
  return {
    sourceSpreadsheetId:source.getId(),
    sourceName:source.getName(),
    sourceVersion:String(sourceConfig.CLASS_PAY_VERSION||"不明"),
    targetVersion:CLASS_PAY_V4_VERSION,
    dataSheets:dataSheets,
    missing:dataSheets.filter(function(x){return !x.present;}).map(function(x){return x.name;}),
    unknown:sourceNames.filter(function(name){return known.indexOf(name)<0;}),
    summary:_cpV4Summary_(source),
    obsolete:_cpV4ObsoleteSheets_().map(function(x){return {name:x.name,reason:x.reason,present:!!target.getSheetByName(x.name)};})
  };
}

function _cpV4Snapshot_(sh){
  return {sheet:sh,values:sh.getDataRange().getValues()};
}

function _cpV4RestoreSnapshot_(snapshot){
  var sh=snapshot.sheet,values=snapshot.values;
  sh.clearContents();
  if(values.length&&values[0].length)sh.getRange(1,1,values.length,values[0].length).setValues(values);
}

function _cpV4ReplaceByHeader_(source,target){
  var sourceValues=source.getDataRange().getValues();
  if(!sourceValues.length)return 0;
  var targetHeaders=target.getRange(1,1,1,target.getLastColumn()).getValues()[0];
  var sourceHeaders=sourceValues[0].map(_cpV4SheetKey_),rows=[];
  sourceValues.slice(1).forEach(function(sourceRow){
    rows.push(targetHeaders.map(function(header){
      var col=sourceHeaders.indexOf(_cpV4SheetKey_(header));
      return col>=0?sourceRow[col]:"";
    }));
  });
  if(target.getLastRow()>1)target.getRange(2,1,target.getLastRow()-1,target.getLastColumn()).clearContent();
  if(rows.length){
    if(target.getMaxRows()<rows.length+1)target.insertRowsAfter(target.getMaxRows(),rows.length+1-target.getMaxRows());
    target.getRange(2,1,rows.length,targetHeaders.length).setValues(rows);
  }
  return rows.length;
}

function _cpV4ReplaceConfig_(source,target){
  var current=_cpV4ConfigMap_(SpreadsheetApp.getActive()),old=_cpV4ConfigMap_(source),protectedKeys={
    ADMIN_PASS:true,BASE_URL:true,CLASS_PAY_VERSION:true,UPDATER_URL:true,DEPLOYMENT_ID:true,
    RELEASE_MANIFEST_URL:true,RELEASE_CHANNEL:true
  };
  var merged={};
  Object.keys(current).forEach(function(key){merged[key]=current[key];});
  Object.keys(old).forEach(function(key){if(!protectedKeys[key])merged[key]=old[key];});
  ["UPDATER_URL","DEPLOYMENT_ID","RELEASE_MANIFEST_URL","RELEASE_CHANNEL"].forEach(function(key){delete merged[key];});
  merged.CLASS_PAY_VERSION=CLASS_PAY_V4_VERSION;
  if(target.getLastRow()>1)target.getRange(2,1,target.getLastRow()-1,Math.max(2,target.getLastColumn())).clearContent();
  var rows=Object.keys(merged).sort().map(function(key){return [key,merged[key]];});
  if(rows.length)target.getRange(2,1,rows.length,2).setValues(rows);
  return rows.length;
}

function _cpV4Validation_(source,target){
  var before=_cpV4Summary_(source),after=_cpV4Summary_(target),checks=[],ok=true;
  ["Users","Shops","Tx","Holdings","CompanyMembers","Government"].forEach(function(name){
    if(!source.getSheetByName(name))return;
    var rowOk=before[name].rows===after[name].rows,balanceOk=Math.abs(before[name].balance-after[name].balance)<0.000001;
    if(!rowOk||!balanceOk)ok=false;
    checks.push({name:name,beforeRows:before[name].rows,afterRows:after[name].rows,beforeBalance:before[name].balance,afterBalance:after[name].balance,ok:rowOk&&balanceOk});
  });
  return {ok:ok,checks:checks};
}

function _cpV4BackupCurrent_(){
  var ss=SpreadsheetApp.getActive(),stamp=Utilities.formatDate(new Date(),"Asia/Tokyo","yyyyMMdd-HHmmss");
  return DriveApp.getFileById(ss.getId()).makeCopy("ClassPay_移行前バックアップ_"+stamp);
}

function _cpV4DeleteObsolete_(ss,names){
  var allow={};_cpV4ObsoleteSheets_().forEach(function(x){allow[x.name]=true;});
  var deleted=[];
  (names||[]).forEach(function(name){
    name=String(name||"");if(!allow[name])throw new Error("削除対象として許可されていないシートです: "+name);
    var sh=ss.getSheetByName(name);if(sh&&ss.getSheets().length>1){ss.deleteSheet(sh);deleted.push(name);}
  });
  return deleted;
}

function _cpV4AppendLog_(item){
  var sh=SpreadsheetApp.getActive().getSheetByName(SHEETS.MIGRATION_LOG||"MigrationLog");
  sh.appendRow([item.migrationId,item.at,item.sourceSpreadsheetId,item.sourceVersion,CLASS_PAY_V4_VERSION,item.status,item.backupSpreadsheetId,(item.importedSheets||[]).join(","),(item.deletedSheets||[]).join(","),JSON.stringify(item.validation||{}),item.message||""]);
}

function api_adminMigrateFromSpreadsheetV4(adminPass,sourceInput,options){
  _assertAdminPassValue_(adminPass);options=options||{};
  if(String(options.confirmText||"").trim()!=="引き継ぐ")throw new Error("確認欄に「引き継ぐ」と入力してください");
  return lockRun_(function(){
    _ensureCoreSheets_();
    var source=_cpV4OpenSource_(sourceInput),target=SpreadsheetApp.getActive(),sourceConfig=_cpV4ConfigMap_(source);
    var backup=_cpV4BackupCurrent_(),snapshots=[],imported=[],deleted=[],validation=null;
    var log={migrationId:"MIG-"+Utilities.formatDate(new Date(),"Asia/Tokyo","yyyyMMdd-HHmmss")+"-"+uuid_().slice(0,5).toUpperCase(),at:_fmtJst_(new Date()),sourceSpreadsheetId:source.getId(),sourceVersion:String(sourceConfig.CLASS_PAY_VERSION||"不明"),backupSpreadsheetId:backup.getId(),status:"RUNNING"};
    try{
      _cpV4RequiredDataSheets_().forEach(function(name){
        var oldSheet=source.getSheetByName(name),newSheet=target.getSheetByName(name);if(!oldSheet||!newSheet)return;
        snapshots.push(_cpV4Snapshot_(newSheet));_cpV4ReplaceByHeader_(oldSheet,newSheet);imported.push(name);
      });
      var configTarget=target.getSheetByName("Config");snapshots.push(_cpV4Snapshot_(configTarget));_cpV4ReplaceConfig_(source,configTarget);imported.push("Config");
      SpreadsheetApp.flush();
      validation=_cpV4Validation_(source,target);
      if(!validation.ok)throw new Error("件数または残高の照合に失敗しました");
      if(options.deleteObsolete===true)deleted=_cpV4DeleteObsolete_(target,options.obsoleteSheets||[]);
      _setConfigValue_("CLASS_PAY_VERSION",CLASS_PAY_V4_VERSION);
      log.status="SUCCESS";log.importedSheets=imported;log.deletedSheets=deleted;log.validation=validation;log.message="データ引き継ぎが完了しました";_cpV4AppendLog_(log);
      return {ok:true,migrationId:log.migrationId,backupSpreadsheetId:backup.getId(),backupUrl:backup.getUrl(),importedSheets:imported,deletedSheets:deleted,validation:validation,message:log.message};
    }catch(e){
      snapshots.reverse().forEach(function(snapshot){try{_cpV4RestoreSnapshot_(snapshot);}catch(ignore){}});SpreadsheetApp.flush();
      log.status="FAILED";log.importedSheets=imported;log.deletedSheets=deleted;log.validation=validation||{};log.message=e.message;_cpV4AppendLog_(log);
      throw new Error("引き継ぎを中止し、移行前の状態へ戻しました: "+e.message);
    }
  });
}

function api_adminCleanupObsoleteSheetsV4(adminPass,names,confirmText){
  _assertAdminPassValue_(adminPass);
  if(String(confirmText||"").trim()!=="削除する")throw new Error("確認欄に「削除する」と入力してください");
  return lockRun_(function(){
    var backup=_cpV4BackupCurrent_(),deleted=_cpV4DeleteObsolete_(SpreadsheetApp.getActive(),names||[]);
    var log={migrationId:"CLEAN-"+Utilities.formatDate(new Date(),"Asia/Tokyo","yyyyMMdd-HHmmss"),at:_fmtJst_(new Date()),sourceSpreadsheetId:"",sourceVersion:"",backupSpreadsheetId:backup.getId(),status:"SUCCESS",importedSheets:[],deletedSheets:deleted,validation:{},message:"不要シートを整理しました"};
    _cpV4AppendLog_(log);
    return {ok:true,deletedSheets:deleted,backupSpreadsheetId:backup.getId(),backupUrl:backup.getUrl(),message:deleted.length?deleted.length+"個のシートを削除しました":"削除対象のシートはありませんでした"};
  });
}
