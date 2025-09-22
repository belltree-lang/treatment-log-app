function ping(){ return 'pong'; }

// 今プロジェクトで「外から呼べる（=トップレベル関数）」を一覧
function listGlobals_(){
  return Object.keys(this)
    .filter(k => typeof this[k] === 'function')
    .sort();
}