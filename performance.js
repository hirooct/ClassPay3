/* ClassPay 3.1: 高速化基盤 */
var _CP_REQUEST_VALUES_ = null;

function _cpBeginRequest_(){
  _CP_REQUEST_VALUES_ = {};
}

function _cpEndRequest_(){
  _CP_REQUEST_VALUES_ = null;
}

/** 1回のAPI実行中は、同じシートを1度だけ読み込む。 */
function _cpSheetValues_(sh){
  if(!_CP_REQUEST_VALUES_) return sh.getDataRange().getValues();
  const key=String(sh.getSheetId());
  if(!_CP_REQUEST_VALUES_[key]) _CP_REQUEST_VALUES_[key]=sh.getDataRange().getValues();
  return _CP_REQUEST_VALUES_[key];
}

function _cpTimed_(name,fn){
  const started=Date.now();
  try{return fn();}
  finally{console.log("ClassPayPerf "+name+" "+(Date.now()-started)+"ms");}
}

/** 残高とマイページ情報を1回の通信で返す。読み取りだけなので排他ロックは使わない。 */
function api_studentDashboardV31(userId,pin,limit){
  _cpBeginRequest_();
  try{
    return _cpTimed_("studentDashboard",function(){
      return {
        version:"3.1",
        account:_apiBalanceRead_(userId,pin,limit),
        civic:api_userCivicDashboard(userId,pin)
      };
    });
  }finally{
    _cpEndRequest_();
  }
}

/** 管理画面の初回表示専用。重い一覧は各タブを開くまで取得しない。 */
function api_adminHomeV31(adminPass){
  _cpBeginRequest_();
  try{
    _assertAdminPassValue_(adminPass);
    return _cpTimed_("adminHome",function(){
      const users=_cpSheetValues_(_getUsersSheet_()),shops=_cpSheetValues_(_getShopsSheet_());
      const um=users.length?_headerMap_(users[0]):null,sm=shops.length?_headerMap_(shops[0]):null;
      const userCount=users.length<2?0:users.slice(1).filter(r=>r[um.idx("userid")]&&(um.idx("isactive")<0||_isTrue_(r[um.idx("isactive")]))).length;
      const shopCount=shops.length<2?0:shops.slice(1).filter(r=>r[sm.idx("shopid")]&&(sm.idx("isactive")<0||_isTrue_(r[sm.idx("isactive")]))).length;
      const summary=api_adminDashboardSummary(adminPass);
      return Object.assign({version:"3.1",governmentBalance:_getGovernmentAccount_().balance,userCount,shopCount},summary);
    });
  }finally{
    _cpEndRequest_();
  }
}

/** 各ページのBASE_URLと初期一覧を1通信で返す。 */
function api_pageBootstrapV31(page){
  page=String(page||"").toLowerCase();
  _cpBeginRequest_();
  try{
    return _cpTimed_("pageBootstrap:"+page,function(){
      const out={version:"3.1",baseUrl:getConfig_("BASE_URL","")};
      if(page==="balance"||page==="pay")out.options=api_getActivePayOptions();
      if(page==="shop")out.shops=api_shopList();
      return out;
    });
  }finally{
    _cpEndRequest_();
  }
}
