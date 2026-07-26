/* ============================================================
   发票自动生成系统 v1 · 本地优先 · 无服务器
   L4 主数据(IndexedDB) → L5 生成引擎(ExcelJS填模板副本) → L6 校验 → L7 导出
   支持 5 家物流商模板(各自字段映射)、数据库降级容错、渲染错误上屏。
   ============================================================ */
'use strict';

/* ---------- 存储层：IndexedDB，不可用时降级为内存(保证不空白) ---------- */
const DB_NAME = 'invoice_sys_v1', DB_VER = 3;
let DB = null, USE_DB = true;
const mem = { channels:[], skus:[], templates:[], records:[], warehouses:[] };
function openDB(){
  return new Promise((res)=>{
    try{
      const r = indexedDB.open(DB_NAME, DB_VER);
      r.onupgradeneeded = e=>{
        const db = e.target.result;
        ['channels','skus','templates','records','warehouses','boxspecs'].forEach(s=>{ if(!db.objectStoreNames.contains(s)) db.createObjectStore(s,{keyPath:'id'}); });
      };
      r.onsuccess = e=>{ DB=e.target.result; res(DB); };
      r.onerror = ()=>{ USE_DB=false; res(); };
    }catch(e){ USE_DB=false; res(); }
  });
}
function _idb(store,mode){ return DB.transaction(store,mode).objectStore(store); }
function getAll(store){
  if(!USE_DB) return Promise.resolve(mem[store]||[]);
  return new Promise((res,rej)=>{ const r=_idb(store,'readonly').getAll(); r.onsuccess=()=>res(r.result||[]); r.onerror=()=>rej(r.error); });
}
function put(store,val){
  if(!USE_DB){ const a=mem[store]||(mem[store]=[]); const i=a.findIndex(x=>x.id===val.id); if(i>=0)a[i]=val;else a.push(val); return Promise.resolve(val); }
  return new Promise((res,rej)=>{ const r=_idb(store,'readwrite').put(val); r.onsuccess=()=>res(val); r.onerror=()=>rej(r.error); });
}
function del(store,id){
  if(!USE_DB){ mem[store]=(mem[store]||[]).filter(x=>x.id!==id); return Promise.resolve(); }
  return new Promise((res,rej)=>{ const r=_idb(store,'readwrite').delete(id); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); });
}
function clear(store){
  if(!USE_DB){ mem[store]=[]; return Promise.resolve(); }
  return new Promise((res,rej)=>{ const r=_idb(store,'readwrite').clear(); r.onsuccess=()=>res(); r.onerror=()=>rej(r.error); });
}
const uid = ()=> Date.now().toString(36)+Math.random().toString(36).slice(2,7);

/* ---------- 5 家物流商模板字段映射(由 inspect_all.js 解析得到) ---------- */
const MAPPINGS = {
  '安速':{
    titleCell:'A1', titleText:'FBA订单（V3）',
    meta:{ fbaNo:'D2', amazonRef:'D4', shipMethod:'D3', warehouseCode:'D5', company:'E6', country:'D7', province:'D8', city:'D9', address:'D10', phone:'D11', zip:'D12', email:'E13', customs:'D14', vat:'E15', eori:'E16', vatName:'E17', vatAddr:'E18', customInfo:'E19' },
    item:{ boxNo:'A', nameCn:'G', nameEn:'H', qty:'I', declare:'K', material:'M', hs:'L', brand:'V', model:'W', boxWeight:'C', len:'D', wid:'E', hgt:'F', elec:'O', magnet:'P', img:'Q', imgUrl:'R', salePrice:'S', saleUrl:'T', currency:'Z', origin:'AA' },
    itemStartRow:21
  },
  '艾杜克':{
    meta:{ company:'B5', vat:'C7', warehouseCode:'I8', amazonRef:'I6', fbaNo:'I7' },
    item:{ nameEn:'A', nameCn:'B', boxNo:'C', brand:'D', hs:'E', qty:'G', boxWeight:'H', len:'I', declare:'J', img:'L', elec:'M', model:'O' },
    itemStartRow:14
  },
  '亦邦':{
    meta:{ company:'F1', amazonRef:'G1', boxWeight:'B1', len:'C1', wid:'D1', hgt:'E1' },
    item:{ boxNo:'A', boxWeight:'B', len:'C', wid:'D', hgt:'E', company:'F', amazonRef:'G', nameCn:'I', nameEn:'J', declare:'K', qty:'L', elec:'M', magnet:'O', sku:'P', hs:'Q', material:'R', purpose:'T', img:'V' },
    itemStartRow:2
  },
  '亚丰':{
    meta:{ fbaNo:'A1', shipMethod:'B2', warehouseCode:'B3', company:'B5', address:'B6', city:'B9', province:'B10', zip:'B11', country:'B12', phone:'B13', email:'C14', customs:'F6', vat:'G11', poNo:'B15' },
    item:{ boxNo:'A', boxWeight:'B', len:'C', wid:'D', hgt:'E', nameEn:'F', nameCn:'G', declare:'H', qty:'I', material:'J', hs:'K', purpose:'L', brand:'M', model:'N', saleUrl:'O', salePrice:'P', img:'Q', imgUrl:'R', prodWeight:'S', elec:'T', magnet:'U', asin:'V', fnsku:'W', sku:'X' },
    itemStartRow:19
  },
  '合联':{
    titleCell:'A1', titleText:'PACKING LIST',
    meta:{ fbaNo:'A3' },
    item:{ fbaNo:'A', boxNo:'B', nameEn:'C', nameCn:'D', hs:'E', boxCount:'F', qty:'G', sku:'H', declare:'I', boxWeight:'K', len:'L', wid:'M', hgt:'N', brand:'P', elec:'Q', img:'R', material:'S' },
    itemStartRow:5
  }
};
const TPL_FILES = {
  '安速':'安速发票模板.xlsx','艾杜克':'艾杜克发票模板.xlsx','亦邦':'亦邦发票模板.xlsx','亚丰':'亚丰发票模板.xlsx','合联':'合联发票模板.xlsx'
};

/* ---------- 种子数据(首次运行注入) ---------- */
async function seedIfEmpty(){
  const ch = await getAll('channels');
  // 迁移守卫:旧版仅 3 条废渠道(空国家/空仓库,与新版 58 条 ID 体系不同)。
  // 改为"按首条新渠道是否存在"判定,确保已缓存旧数据的浏览器也能补全 58 条;
  // 同时清理已知的 3 条旧版废渠道(否则会污染渠道下拉、破坏向导⑥校验)。
  const SEED_NEW_FIRST = 'ch_安速_00';
  const SEED_OLD_IDS = ['ch_ansu_us','ch_aiduk_sa','ch_yifeng_us'];
  if(!ch.some(c=>c.id===SEED_NEW_FIRST)){
    for(const id of SEED_OLD_IDS){ try{ await del('channels', id); }catch(e){} }
    await put('channels',{id:'ch_安速_00',物流商:'安速',渠道:'中运通达-广州DHL(不含油)',国家:'',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_01',物流商:'安速',渠道:'中运通达-广州联邦IP(不含油)',国家:'',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_02',物流商:'安速',渠道:'中运通达-大陆UPS红单小货(含油)',国家:'',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_03',物流商:'安速',渠道:'欧洲包税-空派快线(普货)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_04',物流商:'安速',渠道:'欧洲包税-空派慢线(普货)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_05',物流商:'安速',渠道:'欧洲包税-卡航',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_06',物流商:'安速',渠道:'欧洲包税-卡航卡派',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_07',物流商:'安速',渠道:'欧洲包税-海运',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_08',物流商:'安速',渠道:'欧洲包税-海运卡派',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_09',物流商:'安速',渠道:'加拿大包税-空派(普货)',国家:'加拿大',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_10',物流商:'亚丰',渠道:'亚丰-欧洲空运包税(UPS快线)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_11',物流商:'安速',渠道:'美国包税-空派(普货)',国家:'美国',VAT:'',EORI:'',注册名:'JW PEI INC',注册地址:'123 Oak Ave, CA, US',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_12',物流商:'安速',渠道:'美国包税-海派(美森正班CLX)',国家:'美国',VAT:'',EORI:'',注册名:'JW PEI INC',注册地址:'123 Oak Ave, CA, US',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_13',物流商:'安速',渠道:'美国包税-海派(美森加班MAX)',国家:'美国',VAT:'',EORI:'',注册名:'JW PEI INC',注册地址:'123 Oak Ave, CA, US',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_14',物流商:'安速',渠道:'美国包税-海派(盐田普船)',国家:'美国',VAT:'',EORI:'',注册名:'JW PEI INC',注册地址:'123 Oak Ave, CA, US',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_15',物流商:'安速',渠道:'美国包税-海卡(美森正班卡派)',国家:'美国',VAT:'',EORI:'',注册名:'JW PEI INC',注册地址:'123 Oak Ave, CA, US',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_16',物流商:'安速',渠道:'美国包税-海卡(美森加班卡派)',国家:'美国',VAT:'',EORI:'',注册名:'JW PEI INC',注册地址:'123 Oak Ave, CA, US',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_17',物流商:'安速',渠道:'美国包税-海卡(盐田普船卡派)',国家:'美国',VAT:'',EORI:'',注册名:'JW PEI INC',注册地址:'123 Oak Ave, CA, US',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_18',物流商:'安速',渠道:'加拿大包税-海派(限时达UPS派送)',国家:'加拿大',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_19',物流商:'安速',渠道:'加拿大包税-海派(限时达卡派)',国家:'加拿大',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_20',物流商:'安速',渠道:'加拿大包税-海派(定提UPS派送)',国家:'加拿大',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_21',物流商:'安速',渠道:'加拿大包税-海派(定提卡派)',国家:'加拿大',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_22',物流商:'安速',渠道:'加拿大包税-海派(海运UPS派送)',国家:'加拿大',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_23',物流商:'安速',渠道:'加拿大包税-海派(海运卡派)',国家:'加拿大',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_24',物流商:'安速',渠道:'澳洲包税-海运(悉尼代表)',国家:'澳洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_25',物流商:'安速',渠道:'澳洲包税-空运(普货)',国家:'澳洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_26',物流商:'安速',渠道:'澳洲包税-空运(带磁)',国家:'澳洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_27',物流商:'安速',渠道:'日本自税-海运快船ACP逆算(贴标)',国家:'日本',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_28',物流商:'安速',渠道:'欧洲VAT递延-空派(普货)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_29',物流商:'安速',渠道:'欧洲VAT递延-卡航',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_30',物流商:'安速',渠道:'欧洲VAT递延-海运',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_31',物流商:'安速',渠道:'英国VAT递延-空运快线',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_32',物流商:'安速',渠道:'英国VAT递延-空运慢线',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_33',物流商:'安速',渠道:'英国VAT递延-卡航',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_34',物流商:'安速',渠道:'英国VAT递延-海运',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_35',物流商:'安速',渠道:'英国VAT递延-海卡(BHX4/BHX8/LBA4)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_安速_36',物流商:'安速',渠道:'英国VAT递延-海卡(LPL2/LBA8/EMA3等)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_37',物流商:'亚丰',渠道:'美国包税-海卡(普船)',国家:'美国',VAT:'',EORI:'',注册名:'JW PEI INC',注册地址:'123 Oak Ave, CA, US',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_38',物流商:'亚丰',渠道:'美国空派快线(双清包税UPS)',国家:'美国',VAT:'',EORI:'',注册名:'JW PEI INC',注册地址:'123 Oak Ave, CA, US',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_39',物流商:'亚丰',渠道:'美国空派经济线(双清包税UPS)',国家:'美国',VAT:'',EORI:'',注册名:'JW PEI INC',注册地址:'123 Oak Ave, CA, US',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_40',物流商:'亚丰',渠道:'美国美森限时达(CLX双清包税)',国家:'美国',VAT:'',EORI:'',注册名:'JW PEI INC',注册地址:'123 Oak Ave, CA, US',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_41',物流商:'亚丰',渠道:'欧盟海运包税',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_42',物流商:'亚丰',渠道:'欧盟快铁包税',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_43',物流商:'亚丰',渠道:'欧盟卡航包税(UPS派送)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_44',物流商:'亚丰',渠道:'欧盟卡航包税(限时达)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_45',物流商:'亚丰',渠道:'欧洲自税递延-空运(普货特快限时达)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_46',物流商:'亚丰',渠道:'欧洲自税递延-空运(普货快线限时达)',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_47',物流商:'亚丰',渠道:'欧洲自税递延-海运卡派',国家:'欧洲',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_48',物流商:'亚丰',渠道:'英国空派(自税)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_49',物流商:'亚丰',渠道:'英国空派(包税)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_50',物流商:'亚丰',渠道:'英国海运(自税)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_51',物流商:'亚丰',渠道:'英国海运(包税)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_52',物流商:'亚丰',渠道:'英国卡航(自税)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_亚丰_53',物流商:'亚丰',渠道:'英国卡航(包税)',国家:'英国',VAT:'',EORI:'',注册名:'',注册地址:'',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_艾杜克_54',物流商:'艾杜克',渠道:'沙特空派',国家:'沙特',VAT:'',EORI:'',注册名:'JW PEI Direct',注册地址:'Riyadh, SA',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_艾杜克_55',物流商:'艾杜克',渠道:'沙特空运',国家:'沙特',VAT:'',EORI:'',注册名:'JW PEI Direct',注册地址:'Riyadh, SA',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_艾杜克_56',物流商:'艾杜克',渠道:'沙特空运包税',国家:'沙特',VAT:'',EORI:'',注册名:'JW PEI Direct',注册地址:'Riyadh, SA',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
    await put('channels',{id:'ch_合联_57',物流商:'合联',渠道:'沙特海运',国家:'沙特',VAT:'',EORI:'',注册名:'JW PEI Direct',注册地址:'Riyadh, SA',
      仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]});
  }
  const whs = await getAll('warehouses');
  if(whs.length===0){
    await put('warehouses',{id:'wh_sck8',代码:'SCK8',公司:'Amazon SCK8',省份:'CA',城市:'OAKLEY',地址:'4700 WILBUR AVE',邮编:'94561',电话:'0'});
    await put('warehouses',{id:'wh_edi4',代码:'EDI4',公司:'Amazon EDI4',省份:'TX',城市:'DALLAS',地址:'940 W BETHEL RD',邮编:'75201',电话:'0'});
    await put('warehouses',{id:'wh_ruh8',代码:'RUH8',公司:'Amazon RUH8',省份:'',城市:'Riyadh',地址:'RUH8',邮编:'',电话:'0'});
    await put('warehouses',{id:'wh_tcy1',代码:'TCY1',公司:'TCY1',省份:'CA',城市:'STOCKTON',地址:'2690 East Arch Airport Road',邮编:'95206',电话:'0'});
  }
  // SKU 主数据：从烘焙的 window.SKUS(商品申报信息同步) seed；版本号变更则重 seed
  // 声明价值来源 = 商品申报信息.成本价(人民币) ÷ 汇率(见 sync_master.py RATE)，带版本号可复验
  const SKUS_SEED_VER = 2;
  let skusSeeded = false;
  try { skusSeeded = localStorage.getItem('skus_seeded_ver') === String(SKUS_SEED_VER); } catch(e){}
  if(!skusSeeded){
    await clear('skus');
    for(const s of (window.SKUS||[])){ await put('skus', s); }
    try { localStorage.setItem('skus_seeded_ver', String(SKUS_SEED_VER)); } catch(e){}
  }
  // 箱型规格主数据：从烘焙的 window.BOX_SPECS(「SKU纸箱规格」飞书表同步) seed
  const BOXSPEC_SEED_VER = 1;
  let bsSeeded = false;
  try { bsSeeded = localStorage.getItem('boxspecs_seeded_ver') === String(BOXSPEC_SEED_VER); } catch(e){}
  if(!bsSeeded){
    await clear('boxspecs');
    const specs = window.BOX_SPECS || {};
    for(const b of Object.values(specs)){ await put('boxspecs', b); }
    try { localStorage.setItem('boxspecs_seeded_ver', String(BOXSPEC_SEED_VER)); } catch(e){}
  }
  // 预置 5 家模板（同目录 fetch，仅 http 下可用；file:// 失败则手动上传）
  // 改为"按缺失补全"：避免旧版残留(空/失败记录)导致新模板永远拉不下来。
  const tmpls = await getAll('templates');
  const have = new Set(tmpls.map(t=>t.id));
  for(const [key, file] of Object.entries(TPL_FILES)){
    if(have.has('tmpl_'+key)) continue; // 已有则跳过,不重复
    try{
      const r = await fetch('./'+file);
      if(r.ok){ const blob = await r.blob(); await put('templates',{id:'tmpl_'+key,物流商:key,渠道:'(通用)',名称:file,blob,状态:'ACTIVE',版本:1,创建日:new Date().toISOString().slice(0,10),mapping:MAPPINGS[key]}); }
    }catch(e){ /* 离线 file:// 下跳过，用户手动上传 */ }
  }
}
const COEFF = 0.3; // 推算系数: 申报价 = 成本 × 系数(标黄)

/* ---------- 后端代理状态检测 ---------- */
async function checkBackend(){
  const el=$('#backendStatus'); if(!el) return;
  const url = localStorage.getItem('backend_url') || 'http://localhost:3460';
  try{
    const r=await fetch(url+'/api/health', {signal:AbortSignal.timeout(3000)});
    const d=await r.json();
    if(d.ok){ el.textContent='后端: 已连接'; el.style.borderColor='var(--green)'; el.style.color='var(--green)'; }
    else throw new Error('not ok');
  }catch(e){
    el.textContent='后端: 未连接 (点击配置)'; el.style.borderColor='var(--warn)'; el.style.color='var(--warn)';
  }
  el.style.cursor='pointer';
  el.title='点击配置后端地址';
}
setTimeout(()=>{ checkBackend(); $('#backendStatus').onclick=()=>{
  const cur=localStorage.getItem('backend_url')||'http://localhost:3460';
  const v=prompt('请输入后端代理地址：\n默认 http://localhost:3460 即本机常驻后端（推荐，免配置）\n如需远程访问再填公网地址', cur);
  if(v&&v.trim()){ localStorage.setItem('backend_url', v.trim()); checkBackend(); }
}; }, 500);

/* ---------- 工具 ---------- */
const $ = (s,el=document)=>el.querySelector(s);
const $$ = (s,el=document)=>[...el.querySelectorAll(s)];
function el(html){ const t=document.createElement('template'); t.innerHTML=html.trim(); return t.content.firstElementChild; }
function esc(s){ return String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }
function main(){ return document.getElementById('main'); }

/* ---------- 视图路由 ---------- */
const VIEWS = { overview, wizard, channels, skus, templates, monitor };
function go(view){
  try{
    $$('.nav-btn').forEach(b=>b.classList.toggle('active',b.dataset.view===view));
    main().innerHTML=''; VIEWS[view]();
  }catch(e){ main().innerHTML='<div class="alert alert-err">渲染错误：'+esc(e.message)+'</div>'; console.error(e); }
}
$$('.nav-btn').forEach(b=> b.onclick=()=>go(b.dataset.view));

/* ============================================================
   架构总览
   ============================================================ */
function overview(){
  main().innerHTML = `
  <h2>发票系统骨架架构图 · v1</h2>
  <div class="sub">七层解耦 · 本地优先 · 质量第一。模板=排版层；值来自 L4 主数据（生成时反查填模板）；导出 Excel=默认交付。</div>
  <div class="grid7">
    ${layer('L1 入口','外部触发','入口A 查飞书单号 / 入口B 用户拖文件 / 手建装箱单','')}
    ${layer('L2 适配','可插拔','飞书表·文件·SP-API·聚水潭（解析防错, fail loud）','')}
    ${layer('L3 规范模型','转轴 · 质量命门','装箱单中枢(规范字段); 源忠实、推算值高亮','tag-core')}
    ${layer('L4 配置与资源','主数据权威源','渠道·收货人 / SKU主数据 / 模板库 / 仓库（本地 IndexedDB）','tag-src')}
    ${layer('L5 生成引擎','业务操作','源忠实填空白模板副本 + 装箱单 CRUD','')}
    ${layer('L6 校验反查','质量命门','溯源 / 反查 / 合理性 / 勾稽 / 人审闸门','tag-core')}
    ${layer('L7 交付','输出','导出 Excel(默认) + 发送(可选可跳过)','')}
  </div>
  <div class="card">
    <h3>本 v1 已落地能力</h3>
    <ul>
      <li><b>L4 主数据</b>：渠道·收货人（含仓库子表）、SKU 主数据（申报价带版本号）本地 CRUD，IndexedDB 持久化。</li>
      <li><b>反查机制</b>：选「物流商+渠道+仓库代码」→ 自动带出国家/VAT/地址（绿=主数据反查，白=手填，黄=推算）。</li>
      <li><b>L5 生成引擎</b>：ExcelJS 打开空白模板副本按映射填格，<b>保留原模板样式/合并/图片公式</b>。已支持 5 家模板（安速/艾杜克/亦邦/亚丰/合联）各自字段映射。</li>
      <li><b>L6 校验</b>：必填完整性 + 勾稽（箱数/数量/申报总值）+ 源忠实高亮。</li>
      <li><b>L7 交付</b>：导出填好的 .xlsx（默认，物流商可直接导入其系统）；发送为可选、可跳过。</li>
    </ul>
    <div class="hint">已知边界（v1）：图片自动嵌入未做（模板样例行 =DISPIMG 公式保留）；飞书单号直查/拖文件解析/SP-API 待接入 L2 适配层。</div>
  </div>`;
}
function layer(title,tag,desc,cls){
  const tagHtml = cls?`<span class="tag ${cls}">${tag}</span>`:`<span class="tag tag-src">${tag}</span>`;
  return `<div class="layer">${tagHtml}<h4>${title}</h4><small>${desc}</small></div>`;
}

/* ============================================================
   生成发票向导
   ============================================================ */
let W = null;
async function wizard(){
  const channels = await getAll('channels');
  const skus = await getAll('skus');
  const templates = (await getAll('templates')).filter(t=>t.状态!=='DISABLED');
  W = { step:1, channels, skus, boxspecs: await getAll('boxspecs'), templates, mode:'forward', handover:null,
        form:{ 物流商:'安速', 渠道:'美国包税-空派(普货)', 仓库代码:'SCK8', fbaNo:'', amazonRef:'', customs:'否', customInfo:'', items:[] },
        sources:{}, checks:null };
  renderWizard();
}
async function renderWizard(){
  const m = main();
  m.innerHTML = `
  <h2>生成发票向导</h2>
  <div class="sub">入口 → 反查收货人 → 物品明细 → 选模板预览 → 校验反查 → 人审交付。每一步标注数据来源（绿=主数据反查 / 白=手填 / 黄=推算）。</div>
  <div class="stepper">
    <span class="s ${W.step>=1?'active':''}" data-s="1">① 入口·装箱单</span>
    <span class="s ${W.step>=2?'active':''}" data-s="2">② 反查收货人</span>
    <span class="s ${W.step>=3?'active':''}" data-s="3">③ 物品明细</span>
    <span class="s ${W.step>=4?'active':''}" data-s="4">④ 选模板·预览</span>
    <span class="s ${W.step>=5?'active':''}" data-s="5">⑤ 校验反查</span>
    <span class="s ${W.step>=6?'active':''}" data-s="6">⑥ 人审·交付</span>
  </div>
  <div id="wstep"></div>`;
  $$('.stepper .s').forEach(s=> s.onclick=()=>{ const n=+s.dataset.s; if(n<W.step) {W.step=n; renderWizard();} });
  const box = $('#wstep');
  if(W.step===1) step1(box);
  else if(W.step===2) await step2(box);
  else if(W.step===3) step3(box);
  else if(W.step===4) step4(box);
  else if(W.step===5) step5(box);
  else if(W.step===6) await step6(box);
}
function step1(box){
  const f = W.form;
  const isRev = W.mode==='reverse';
  box.innerHTML = `
  <div class="card">
    <h3>① 入口与装箱单（中枢单据）</h3>
    <div class="hint">两种录入方式：<b>正着填</b>逐项手填；<b>倒着填</b>输入单号，从已同步的「FBA箱唛交接」索引一次性抓取整行信息（装箱清单需先同步）。</div>
    <div class="seg">
      <button class="seg-btn ${!isRev?'on':''}" id="m_forward">正着填（手动）</button>
      <button class="seg-btn ${isRev?'on':''}" id="m_reverse">倒着填（按单号搜）</button>
    </div>
    ${isRev ? reversePanelHTML() : forwardPanelHTML(f)}
  </div>
  <div style="margin-top:14px"><button class="btn" id="next1">下一步：反查收货人 →</button></div>`;
  $('#m_forward').onclick=()=>{W.mode='forward';renderWizard();};
  $('#m_reverse').onclick=()=>{W.mode='reverse';W.handover=null;renderWizard();};
  if(isRev){ bindReverse(); }
  else { bindForward(f); }
  $('#next1').onclick = ()=>{ if(!isRev) captureForward(f); W.step=2; renderWizard(); };
}
function forwardPanelHTML(f){
  return `
    <div class="row">
      <div><label>物流商</label><select id="f_物流商">${[...new Set(W.channels.map(c=>c.物流商))].map(o=>`<option ${o===f.物流商?'selected':''}>${o}</option>`).join('')}</select></div>
      <div><label>渠道</label><select id="f_渠道">${W.channels.filter(c=>c.物流商===f.物流商).map(c=>`<option ${c.渠道===f.渠道?'selected':''}>${c.渠道}</option>`).join('')}</select></div>
      <div><label>仓库代码</label><select id="f_仓库代码">${warehouseOptions(f)}</select></div>
    </div>
    <div class="row">
      <div><label>客户订单号(FBA号)</label><input id="f_fbaNo" value="${esc(f.fbaNo)}" placeholder="FBA19JMKBVJW"></div>
      <div><label>Amazon Reference ID</label><input id="f_amazonRef" value="${esc(f.amazonRef)}" placeholder="4F7O73TF"></div>
      <div><label>报关(否/是)</label><select id="f_customs"><option ${f.customs==='否'?'selected':''}>否</option><option ${f.customs==='是'?'selected':''}>是</option></select></div>
    </div>
    <label>自定义信息</label><input id="f_customInfo" value="${esc(f.customInfo)}" placeholder="可留空">`;
}
function bindForward(f){
  $('#f_物流商').onchange = e=>{ f.物流商=e.target.value; const ch=W.channels.find(c=>c.物流商===f.物流商); f.渠道=ch?ch.渠道:f.渠道; f.仓库代码=ch&&ch.仓库[0]?ch.仓库[0].代码:f.仓库代码; renderWizard(); };
  $('#f_渠道').onchange = e=>{ f.渠道=e.target.value; const ch=W.channels.find(c=>c.物流商===f.物流商&&c.渠道===f.渠道); f.仓库代码=ch&&ch.仓库[0]?ch.仓库[0].代码:f.仓库代码; renderWizard(); };
  $('#f_仓库代码').onchange = e=>{ f.仓库代码=e.target.value; };
  const cap = ()=>{ f.fbaNo=$('#f_fbaNo').value; f.amazonRef=$('#f_amazonRef').value; f.customs=$('#f_customs').value; f.customInfo=$('#f_customInfo').value; };
  ['#f_fbaNo','#f_amazonRef','#f_customs','#f_customInfo'].forEach(s=>$(s).onchange=cap);
}
function captureForward(f){ f.fbaNo=$('#f_fbaNo').value; f.amazonRef=$('#f_amazonRef').value; f.customs=$('#f_customs').value; f.customInfo=$('#f_customInfo').value; }
function reversePanelHTML(){
  const n=(window.HANDOVER_INDEX||[]).length;
  return `
    <div class="hint ok">数据源：本机已同步的「FBA箱唛交接」索引（${n} 票）。输入单号 → 搜索 → 确认即抓取整行信息。</div>
    <div class="row"><div style="flex:1"><label>内部单号 / 聚水潭 或 FBA货件号</label><input id="rev_q" placeholder="如 DAFA-EXP 或 FBA15K43YSGT" autocomplete="off"></div></div>
    <div id="rev_res" class="rev-res"></div>
    <div id="rev_confirm"></div>`;
}
function bindReverse(){
  const q=$('#rev_q'); const res=$('#rev_res');
  q.oninput=()=>{
    const key=(q.value||'').trim().toLowerCase();
    res.innerHTML=''; $('#rev_confirm').innerHTML='';
    if(key.length<2) return;
    const all=window.HANDOVER_INDEX||[];
    const hits=all.filter(r=> (r.internal_no&&r.internal_no.toLowerCase().includes(key)) || (r.fba_shipment&&r.fba_shipment.toLowerCase().includes(key)) ).slice(0,30);
    if(hits.length===0){ res.innerHTML='<div class="hint">未找到匹配。可检查单号，或切回「正着填」，或先执行数据同步。</div>'; return; }
    res.innerHTML = hits.map((r,i)=>`<div class="rev-item" data-i="${i}"><span><b>${esc(r.internal_no||'(无内部单号)')}</b> / ${esc(r.fba_shipment||'(无FBA号)')}</span><span class="muted">${esc(r.carrier||r.物流商||'?')} · ${esc(r.country||'?')} · ${esc(r.boxes||'?')}箱 · ${esc(r.packing_list||'无装箱清单')}</span></div>`).join('');
    res.querySelectorAll('.rev-item').forEach(el=> el.onclick=()=> showConfirm(hits[+el.dataset.i]) );
  };
}
function showConfirm(r){
  const c=$('#rev_confirm');
  c.innerHTML = `
    <div class="card" style="margin-top:10px">
      <div class="hint warn">请确认这是你要的单（防取错）：</div>
      <table class="kv">
        <tr><td>内部单号</td><td>${esc(r.internal_no)}</td></tr>
        <tr><td>FBA货件号</td><td>${esc(r.fba_shipment)}</td></tr>
        <tr><td>物流商</td><td>${esc(r.carrier||r.物流商)}</td></tr>
        <tr><td>国家</td><td>${esc(r.country)}</td></tr>
        <tr><td>空/海运</td><td>${esc(r.air_sea)}</td></tr>
        <tr><td>箱数</td><td>${esc(r.boxes)}</td></tr>
        <tr><td>取件方式(地址)</td><td>${esc(r.pickup_addr)}</td></tr>
        <tr><td>装箱清单</td><td>${esc(r.packing_list)}</td></tr>
        <tr><td>FNSKU信息</td><td>${esc(r.fnsku_file)}</td></tr>
        <tr><td>发票-物流填</td><td>${esc(r.invoice_drop)}</td></tr>
      </table>
      <div style="display:flex;gap:10px;margin-top:8px">
        <button class="btn" id="rev_ok">确认采用，抓取信息 →</button>
        <button class="btn secondary" id="rev_cancel">重选</button>
      </div>
    </div>`;
  $('#rev_cancel').onclick=()=>{ c.innerHTML=''; };
  $('#rev_ok').onclick=()=> applyHandover(r);
}
function applyHandover(r){
  const f=W.form; W.handover=r;
  f.fbaNo = r.fba_shipment || r.internal_no;
  if((r.carrier||r.物流商) && W.channels.some(c=>c.物流商===(r.carrier||r.物流商))) f.物流商=r.carrier||r.物流商;
  const sameCarrier = W.channels.filter(c=>c.物流商===f.物流商);
  const byCountry = sameCarrier.find(c=>c.国家 && r.country && c.国家.includes(r.country));
  f.渠道 = (byCountry||sameCarrier[0]||{渠道:f.渠道}).渠道;
  const whByC = { '美国':'SCK8', '沙特':'RUH8' };
  if(whByC[r.country]) f.仓库代码=whByC[r.country];
  f.customInfo = r.invoice_drop || f.customInfo;
  // 倒填核心：按货件号从已烘焙的装箱清单一次性拉箱内容，自动填入物品（点一下自动填充）
  const fid = r.fba_shipment || r.internal_no;
  W.packFbaId = fid;
  const pl = (window.PACKING_LISTS && window.PACKING_LISTS[fid]) || [];
  if(pl.length){ W.form.items = pl.map(x=>Object.assign({}, x)); W.plAutoFilled = pl.length; }
  else { W.plAutoFilled = 0; }
  W.step=2; renderWizard();
}
function warehouseOptions(f){
  const ch = W.channels.find(c=>c.物流商===f.物流商&&c.渠道===f.渠道);
  if(!ch) return '';
  return ch.仓库.map(w=>`<option ${w.代码===f.仓库代码?'selected':''}>${w.代码}</option>`).join('');
}
function lookupChannel(){ return W.channels.find(c=>c.物流商===W.form.物流商&&c.渠道===W.form.渠道) || null; }
async function lookupWarehouse(){ return (await getAll('warehouses')).find(w=>w.代码===W.form.仓库代码) || null; }
async function step2(box){
  const ch=lookupChannel(), wh=await lookupWarehouse();
  const src = W.sources = {
    shipMethod:{v:ch?ch.渠道:'',src:'channel'}, country:{v:ch?ch.国家:'',src:'channel'}, vat:{v:ch?ch.VAT:'',src:'channel'},
    eori:{v:ch?ch.EORI:'',src:'channel'}, vatName:{v:ch?ch.注册名:'',src:'channel'}, vatAddr:{v:ch?ch.注册地址:'',src:'channel'},
    warehouseCode:{v:W.form.仓库代码,src:'manual'}, company:{v:wh?wh.公司:'',src:'warehouse'}, province:{v:wh?wh.省份:'',src:'warehouse'},
    city:{v:wh?wh.城市:'',src:'warehouse'}, address:{v:wh?wh.地址:'',src:'warehouse'}, zip:{v:wh?wh.邮编:'',src:'warehouse'}, phone:{v:wh?wh.电话:'',src:'warehouse'},
    fbaNo:{v:W.form.fbaNo,src:'manual'}, amazonRef:{v:W.form.amazonRef,src:'manual'}, customs:{v:W.form.customs,src:'manual'}, customInfo:{v:W.form.customInfo,src:'manual'},
    title:{v:'',src:'template'}, poNo:{v:'',src:'manual'}
  };
  const srcClass = s => (s==='channel'||s==='warehouse'||s==='template') ? 'cell-src' : 'cell-manual';
  const srcLabel = s => ({channel:'主数据·渠道',warehouse:'主数据·仓库',manual:'人工填写',template:'模板固定',calc:'推算'}[s]||s);
  const FIELDS = ['fbaNo','amazonRef','poNo','shipMethod','warehouseCode','company','country','province','city','address','phone','zip','email','customs','vat','eori','vatName','vatAddr','customInfo'];
  const LABELS = {fbaNo:'客户订单号(FBA号)',amazonRef:'Amazon Reference ID',poNo:'PO Number',shipMethod:'运输方式',warehouseCode:'收件人(仓库代码)',company:'收件人公司',country:'国家',province:'收件省份',city:'收件城市',address:'收件地址',phone:'收件电话',zip:'邮编',email:'收件人email',customs:'报关(否/是)',vat:'VAT号',eori:'EORI',vatName:'VAT注册名',vatAddr:'VAT注册地址',customInfo:'自定义信息'};
  box.innerHTML = `
  <div class="card">
    <h3>② 反查收货人（生成时从 L4 主数据取值，不手敲）</h3>
    <div class="hint">所选：<b>${esc(W.form.物流商)} / ${esc(W.form.渠道)}</b>，仓库代码 <b>${esc(W.form.仓库代码)}</b>。绿底=主数据反查带出，白底=需人工填（传统贸易常见）。可直接改，但建议改「主数据页」以保证全量一致。</div>
    <table>
      <thead><tr><th>字段</th><th>取值</th><th>来源</th></tr></thead>
      <tbody>
        ${FIELDS.map(k=>{ const s=src[k]; if(!s) return ''; return `<tr><td>${LABELS[k]}</td><td class="${srcClass(s.src)}"><input data-meta="${k}" value="${esc(s.v)}"></td><td><span class="pill ${s.src==='manual'?'pill-gray':'pill-green'}">${srcLabel(s.src)}</span></td></tr>`; }).join('')}
      </tbody>
    </table>
  </div>
  <div style="margin-top:14px;display:flex;gap:10px"><button class="btn secondary" id="prev2">← 上一步</button><button class="btn" id="next2">下一步：物品明细 →</button></div>`;
  $$('[data-meta]').forEach(inp=> inp.oninput = e=>{ W.sources[e.target.dataset.meta].v=e.target.value; W.sources[e.target.dataset.meta].src='manual'; });
  $('#prev2').onclick=()=>{W.step=1;renderWizard();};
  $('#next2').onclick=()=>{W.step=3;renderWizard();};
}
function step3(box){
  if(W.form.items.length===0){ W.form.items.push({boxNo:'',sku:'8T026-12',nameCn:'',nameEn:'',qty:8,declare:'',material:'',hs:'',brand:'',model:'',boxWeight:'',len:'',wid:'',hgt:'',elec:'N',magnet:'N',saleUrl:'',cost:''}); }
  function addRow(){ W.form.items.push({boxNo:'',sku:'',nameCn:'',nameEn:'',qty:1,declare:'',material:'',hs:'',brand:'',model:'',boxWeight:'',len:'',wid:'',hgt:'',elec:'N',magnet:'N',saleUrl:'',cost:'',boxes:'',boxSpec:''}); renderWizard(); }
  function renderRows(){
    return W.form.items.map((it,i)=>{
      const sk = W.skus.find(s=>s.sku===it.sku);
      const normSku = s=>(s||'').replace(/@us$/i,'').trim();
      const bs = W.boxspecs.find(b=>normSku(b.sku)===normSku(it.sku)) || null;
      let declareSrc='manual', declareVal=it.declare;
      if(declareVal===''||declareVal==null){ if(sk && sk.申报价){ declareVal=sk.申报价; declareSrc='sku'; } else if(window.SKU_DECLARE && window.SKU_DECLARE[it.sku]){ declareVal=window.SKU_DECLARE[it.sku].d; declareSrc='sku'; } else if(it.cost!==''){ declareVal=(parseFloat(it.cost)*COEFF).toFixed(2); declareSrc='calc'; } }
      let dcls, pill, pillTxt;
      if(declareSrc==='calc'){ dcls='cell-calc'; pill='pill-yellow'; pillTxt='推算(成本×'+COEFF+')'; }
      else if(declareSrc==='sku'){ dcls='cell-src'; pill='pill-green'; pillTxt='SKU主数据'; }
      else if(declareVal!=='' && declareVal!=null){ dcls='cell-manual'; pill='pill-gray'; pillTxt='手填'; }
      else { dcls='cell-warn'; pill='pill-red'; pillTxt='⚠ 待手填'; }
      return `<tr>
        <td><input data-i="${i}" data-k="boxNo" value="${esc(it.boxNo)}" placeholder="箱号"></td>
        <td><input data-i="${i}" data-k="sku" value="${esc(it.sku)}" list="skuList" placeholder="SKU"></td>
        <td><input data-i="${i}" data-k="nameCn" value="${esc(it.nameCn||(sk?sk.中文品名:'')||(window.SKU_DECLARE&&window.SKU_DECLARE[it.sku]?window.SKU_DECLARE[it.sku].n:''))}" placeholder="中文"></td>
        <td><input data-i="${i}" data-k="nameEn" value="${esc(it.nameEn||(sk?sk.英文品名:''))}" placeholder="英文"></td>
        <td><input data-i="${i}" data-k="qty" value="${esc(it.qty)}" style="width:54px" placeholder="数量"></td>
        <td class="${dcls}"><input data-i="${i}" data-k="declare" value="${esc(declareVal)}" style="width:74px"><br><span class="pill ${pill}">${pillTxt}</span></td>
        <td><input data-i="${i}" data-k="material" value="${esc(it.material||(sk?sk.材质:''))}" placeholder="材质"></td>
        <td><input data-i="${i}" data-k="hs" value="${esc(it.hs||(sk?sk.HS:''))}" placeholder="HS"></td>
        <td><input data-i="${i}" data-k="brand" value="${esc(it.brand||(sk?sk.品牌:''))}" placeholder="品牌"></td>
        <td><input data-i="${i}" data-k="model" value="${esc(it.model||(sk?sk.型号:''))}" placeholder="型号"></td>
        <td><input data-i="${i}" data-k="boxes" value="${esc(it.boxes)}" style="width:46px" placeholder="箱数"></td>
        <td class="${(!it.boxSpec||it.boxSpec==='')&&bs&&bs.model?'cell-src':''}"><input data-i="${i}" data-k="boxSpec" value="${esc((!it.boxSpec||it.boxSpec==='')&&bs&&bs.model?bs.model:it.boxSpec)}" placeholder="箱规"></td>
        <td class="${(!it.boxWeight||it.boxWeight==='')&&bs&&(bs.weight!=null)?'cell-src':''}"><input data-i="${i}" data-k="boxWeight" value="${esc((!it.boxWeight||it.boxWeight==='')&&bs&&(bs.weight!=null)?bs.weight:it.boxWeight)}" style="width:54px" placeholder="箱重"></td>
        <td class="${(!it.len||it.len==='**')&&bs&&bs.l?'cell-src':''}"><input data-i="${i}" data-k="len" value="${esc((!it.len||it.len==='**')&&bs&&bs.l?bs.l:it.len)}" style="width:46px" placeholder="长"></td>
        <td class="${(!it.wid||it.wid==='**')&&bs&&bs.w?'cell-src':''}"><input data-i="${i}" data-k="wid" value="${esc((!it.wid||it.wid==='**')&&bs&&bs.w?bs.w:it.wid)}" style="width:46px" placeholder="宽"></td>
        <td class="${(!it.hgt||it.hgt==='**')&&bs&&bs.h?'cell-src':''}"><input data-i="${i}" data-k="hgt" value="${esc((!it.hgt||it.hgt==='**')&&bs&&bs.h?bs.h:it.hgt)}" style="width:46px" placeholder="高"></td>
        <td><input data-i="${i}" data-k="elec" value="${esc(it.elec)}" style="width:38px" placeholder="电"></td>
        <td><input data-i="${i}" data-k="magnet" value="${esc(it.magnet)}" style="width:38px" placeholder="磁"></td>
        <td><input data-i="${i}" data-k="saleUrl" value="${esc(it.saleUrl)}" placeholder="销售链接"></td>
        <td><button class="btn danger" data-del="${i}" style="padding:4px 8px">删</button></td>
      </tr>`;
    }).join('');
  }
  box.innerHTML = `
  <div class="card">
    <h3>③ 物品明细（装箱单行项目）</h3>
    <div class="hint">填 SKU 自动反查中文品名/材质/HS/品牌/型号，并带出<b>申报价</b>（绿=SKU主数据；黄=无主数据按成本×${COEFF}推算，需人审确认）。</div>
    <datalist id="skuList">${W.skus.map(s=>`<option value="${s.sku}">${s.中文品名}</option>`).join('')}</datalist>
    ${ W.handover ? packingBannerHTML() : '' }
    <div style="overflow:auto"><table>
      <thead><tr><th>箱号</th><th>SKU</th><th>中文</th><th>英文</th><th>数量</th><th>申报价(USD)</th><th>材质</th><th>HS</th><th>品牌</th><th>型号</th><th>箱数</th><th>箱规</th><th>箱重</th><th>长</th><th>宽</th><th>高</th><th>电</th><th>磁</th><th>销售链接</th><th></th></tr></thead>
      <tbody id="rows">${renderRows()}</tbody>
    </table></div>
    <button class="btn secondary" id="addRow" style="margin-top:10px">+ 添加一行</button>
  </div>
  <div style="margin-top:14px;display:flex;gap:10px"><button class="btn secondary" id="prev3">← 上一步</button><button class="btn" id="next3">下一步：选模板·预览 →</button></div>`;
  $$('#rows [data-i]').forEach(inp=> inp.oninput = e=>{ const i=+e.target.dataset.i, k=e.target.dataset.k; W.form.items[i][k]=e.target.value; if(k==='sku') renderWizard(); });
  $$('#rows [data-del]').forEach(b=> b.onclick=()=>{ W.form.items.splice(+b.dataset.del,1); renderWizard(); });
  $('#addRow').onclick=addRow;
  $('#prev3').onclick=()=>{W.step=2;renderWizard();};
  $('#next3').onclick=()=>{W.step=4;renderWizard();};
  if(W.handover){
    const fid = W.packFbaId || W.handover.fba_shipment || W.handover.internal_no;
    const rl=$('#pl_reload'); if(rl) rl.onclick=()=> loadPackingList(fid);
    const pf=$('#pl_file'); if(pf) pf.onchange=e=>{
      const file=e.target.files[0]; if(!file) return;
      const msg=$('#pl_msg'); msg.textContent='⏳ 正在解析 '+file.name+'...';
      const isXlsx = /\.(xlsx|xls)$/i.test(file.name);
      const rd=new FileReader();
      rd.onload=async()=>{
        try{
          let items = isXlsx ? await parsePackingXlsx(rd.result) : parsePackingList(rd.result);
          if(items.length){ W.form.items=items; msg.textContent='✅ 已上传并填入 '+items.length+' 行（'+file.name+'）。'; renderWizard(); }
          else msg.textContent='⚠️ 解析为空，请确认文件是有效的装箱清单（Excel xlsx 或 CSV）。';
        }catch(err){ console.error(err); msg.textContent='❌ 解析失败：'+(err.message||err); }
      };
      rd.onerror=()=>{ msg.textContent='❌ 文件读取失败'; };
      if(isXlsx) rd.readAsArrayBuffer(file); else rd.readAsText(file);
    };
    const ob=$('#pl_online'); if(ob) ob.onclick=()=> onlineFetch(fid);
  }
}
function packingBannerHTML(){
  const fid = W.packFbaId || (W.handover&&(W.handover.fba_shipment||W.handover.internal_no)) || '';
  const pl = (window.PACKING_LISTS && window.PACKING_LISTS[fid]) || [];
  const fn = W.handover ? W.handover.packing_list : '';
  if(pl.length){
    return `
    <div class="card" style="margin-top:10px;border-color:#2b6cb0">
      <div class="hint ok">✅ 已从装箱清单自动填入 <b>${pl.length}</b> 行物品（货件 ${esc(fid)}）。请核对品名/数量/申报价，无误即可继续。</div>
      <div style="display:flex;gap:10px;align-items:center;flex-wrap:wrap">
        <button class="btn" id="pl_reload">↻ 重新从装箱清单拉取</button>
        <span class="muted">源：${esc(fn||'(FBA箱唛交接表关联)')}</span>
      </div>
    </div>`;
  }
  return `
  <div class="card" style="margin-top:10px;border-color:#c53030">
    <div class="hint warn">⚠️ 本地未收录该货件（${esc(fid)}）的装箱清单内容。下方两个按钮：</div>
    <div style="display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin-top:10px">
      <label class="btn" id="pl_upload_btn" style="margin:0;font-size:15px;padding:10px 20px;background:#2b6cb0;color:#fff"><span style="font-size:17px;margin-right:6px">📤</span>上传装箱清单（Excel / CSV）<input type="file" id="pl_file" accept=".xlsx,.xls,.csv" style="display:none"></label>
      <button class="btn" id="pl_online" style="font-size:15px;padding:10px 20px;background:#38a169;color:#fff"><span style="font-size:17px;margin-right:6px">🌐</span>在线获取</button>
      <span id="pl_msg" class="muted" style="flex:1;min-width:200px"></span>
    </div>
    <div class="hint" style="margin-top:10px;font-size:12px;color:#888">
      <b>📤 上传</b>：直接选本机的 xlsx/xls/csv 装箱清单（无需另存为 CSV）。<br>
      <b>🌐 在线获取</b>：①系统已收录该货件号 → 秒级自动填入；②<b>未收录</b> → <b>点击后会自动调后端从飞书云文档拉取</b>。<br>
      <span style="color:#c53030">⚠️ 后端未启动？</span> 双击 <code>backend\install_autostart.bat</code> 一次性安装开机自启，<b>之后再也不需要管</b>。
    </div>
  </div>`;
}
function loadPackingList(fid){
  const pl = (window.PACKING_LISTS && window.PACKING_LISTS[fid]) || [];
  if(pl.length){
    W.form.items = pl.map(x=>{
      x = Object.assign({}, x);
      if((!x.declare||x.declare==='') && window.SKU_DECLARE && window.SKU_DECLARE[x.sku]){
        x.declare = window.SKU_DECLARE[x.sku].d;
        if(!x.nameCn) x.nameCn = window.SKU_DECLARE[x.sku].n;
      }
      return x;
    });
    renderWizard();
  } else { const m=$('#pl_msg'); if(m) m.textContent='本地未收录该装箱清单'; }
}
function parsePackingList(text){
  const lines=text.split(/\r?\n/).filter(l=>l.trim());
  if(lines.length<2) return [];
  const splitCsv=line=> line.split(',').map(s=>s.trim().replace(/^"|"$/g,''));
  const hdr=splitCsv(lines[0]).map(h=>h.toLowerCase());
  const map={'boxNo':['箱号','箱','carton','boxno','ctn'],'sku':['sku'],'nameCn':['中文品名','品名','名称','namecn'],'nameEn':['英文品名','英文名称','nameen'],'qty':['数量','qty','quantity'],'declare':['申报价','申报价值','declare','price'],'material':['材质','material'],'hs':['hs','海关编码'],'brand':['品牌','brand'],'model':['型号','model'],'boxWeight':['箱重','重量','weight'],'len':['长','length'],'wid':['宽','width'],'hgt':['高','height'],'elec':['带电','elec'],'magnet':['带磁','magnet'],'saleUrl':['销售链接','链接','url','saleurl']};
  const idx={};
  for(const f in map){ const i=hdr.findIndex(h=>map[f].some(k=>h.includes(k))); if(i>=0) idx[f]=i; }
  const out=[];
  for(let n=1;n<lines.length;n++){
    const c=splitCsv(lines[n]);
    if(c.every(x=>!x)) continue;
    const get=f=> idx[f]>=0 ? c[idx[f]] : '';
    out.push({boxNo:get('boxNo'),sku:get('sku'),nameCn:get('nameCn'),nameEn:get('nameEn'),qty:get('qty')||1,declare:get('declare'),material:get('material'),hs:get('hs'),brand:get('brand'),model:get('model'),boxWeight:get('boxWeight'),len:get('len'),wid:get('wid'),hgt:get('hgt'),elec:(get('elec')||'N').toUpperCase().startsWith('Y')?'Y':'N',magnet:(get('magnet')||'N').toUpperCase().startsWith('Y')?'Y':'N',saleUrl:get('saleUrl'),cost:''});
  }
  return out;
}

/* 在线获取：带 loading 进度条，先查缓存，后续扩展为后端代理 */
function onlineFetch(fid){
  const msg=$('#pl_msg'); if(!msg) return;
  const btn=$('#pl_online'); if(btn) btn.disabled=true;
  const steps=[
    {txt:'正在搜索货件号 <b>'+esc(fid)+'</b> 的装箱清单...'},
    {txt:'正在读取系统预装数据...'},
    {txt:'数据解析完成，正在填入表单...'}
  ];
  msg.innerHTML = '<div class="loading-wrap" id="loading_ui">'
    +'<div class="loading-header"><div class="spinner"></div><div class="loading-title">正在处理，请稍候...</div></div>'
    +'<div class="loading-bar-wrap"><div class="loading-bar" id="loading_bar"></div></div>'
    +'<div class="loading-steps" id="loading_steps">'
    +steps.map((s,i)=>'<div class="loading-step" data-i="'+i+'"><div class="dot"></div>'+s.txt+'</div>').join('')
    +'</div></div>';

  /* 阶段推进 */
  function setStep(i,state){
    const el=$('[data-i="'+i+'"]','#loading_steps');
    if(el){ el.className='loading-step '+state; }
  }
  function setBar(pct){
    const bar=$('#loading_bar');
    if(bar) bar.style.width=pct+'%';
  }
  function done(ok,html){
    const spinner=$('.spinner','#loading_ui');
    const title=$('.loading-title','#loading_ui');
    if(spinner) spinner.style.animation='none';
    if(title){
      title.innerHTML=ok ? '✅ 完成' : '⚠️ 未找到';
      title.style.color=ok ? 'var(--green)' : 'var(--warn)';
    }
    const wrap=$('#loading_ui');
    /* 2秒后替换成结果信息 */
    setTimeout(()=>{
      if(wrap) wrap.outerHTML=html;
      if(btn) btn.disabled=false;
    }, ok ? 800 : 2000);
  }

  /* 阶段1: 正在搜索（立即激活）*/
  setStep(0,'active');
  setBar(20);

  setTimeout(()=>{
    setStep(0,'done');
    setStep(1,'active');
    setBar(50);

    setTimeout(()=>{
      setStep(1,'done');

      /* 实际查数据 */
      const pl = (window.PACKING_LISTS && window.PACKING_LISTS[fid]) || [];
      if(pl.length){
        setStep(2,'active');
        setBar(80);
        setTimeout(()=>{
          W.form.items = pl.map(x=>{
            x = Object.assign({}, x);
            if((!x.declare||x.declare==='') && window.SKU_DECLARE && window.SKU_DECLARE[x.sku]){
              x.declare = window.SKU_DECLARE[x.sku].d;
              if(!x.nameCn) x.nameCn = window.SKU_DECLARE[x.sku].n;
            }
            return x;
          });
          setStep(2,'done');
          setBar(100);
          renderWizard();
          done(true,'<div class="hint ok" style="margin-top:10px">✅ 已从系统预装的装箱清单自动填入 <b>'+pl.length+'</b> 行（货件 '+esc(fid)+'）。请核对品名/数量/申报价。</div>');
        },300);
      } else {
        /* 未预装 → 尝试调本地后端代理（15秒超时,卡死自动放弃） */
        setStep(1,'done');
        setStep(2,'active');
        setBar(60);
        const backendUrl = localStorage.getItem('backend_url') || 'http://localhost:3460';
        fetch(backendUrl+'/api/fetch-packing-list', {
          method:'POST',
          headers:{'Content-Type':'application/json'},
          body:JSON.stringify({fid}),
          signal: AbortSignal.timeout(15000) // 15s 超时,死锁自动放弃
        }).then(r=>r.json()).then(data=>{
          if(data.ok && data.items && data.items.length>0){
            setStep(2,'done');
            setBar(100);
            W.form.items = data.items.map(x=>{
              if((!x.declare||x.declare==='') && window.SKU_DECLARE && window.SKU_DECLARE[x.sku]){
                x.declare = window.SKU_DECLARE[x.sku].d;
                if(!x.nameCn) x.nameCn = window.SKU_DECLARE[x.sku].n;
              }
              return x;
            });
            renderWizard();
            done(true,'<div class="hint ok" style="margin-top:10px">✅ 已通过后端代理从飞书云文档拉取 <b>'+data.items.length+'</b> 行（货件 '+esc(fid)+'）。请核对品名/数量/申报价。</div>');
          } else {
            throw new Error(data?.error||'后端返回空数据');
          }
        }).catch(e=>{
          setStep(2,'fail');
          setBar(100);
          let hint = '';
          if(e.name==='TimeoutError' || e.message.includes('timeout')){
            hint = '<div class="hint warn" style="margin-top:10px">⏱️ 后端15秒超时未响应。<br>① <b>双击 <code>backend\\start_backend.bat</code> 启动本地后端</b>（首次需本机 lark-cli 已授权）；<br>② 或运行 <code>backend\\install_autostart.bat</code> 注册开机自启，<b>之后完全不用管</b>；<br>③ 当前货件可直接点「<b>📤 上传装箱清单</b>」选本机 xlsx。</div>';
          } else if(e.message.includes('Failed to fetch')||e.message.includes('fetch')){
            hint = '<div class="hint warn" style="margin-top:10px">❌ 浏览器连不上本地后端。<br><b>最简单的解决方案：</b>双击 <code>backend\\install_autostart.bat</code> 注册开机自启，<b>之后什么都不用操作</b>，任何时间打开本页面点「在线获取」就秒响应。</div>';
          } else {
            hint = '<div class="hint warn" style="margin-top:10px">⚠️ <b>'+esc(fid)+'</b> 后端不可用（'+esc(e.message)+'）。<br>① 双击 <code>backend\\start_backend.bat</code> 启动；<br>② 或点「<b>📤 上传装箱清单</b>」选本机 xlsx（<b>已加强解析支持</b>）。</div>';
          }
          done(false, hint);
        });
      }
    }, 400);
  }, 300);
}

/* 带 loading 的 loadPackingList（从搜索结果自动填入也用同样的流程） */
function loadPackingList(fid){
  const pl = (window.PACKING_LISTS && window.PACKING_LISTS[fid]) || [];
  if(pl.length){
    W.form.items = pl.map(x=>{
      x = Object.assign({}, x);
      if((!x.declare||x.declare==='') && window.SKU_DECLARE && window.SKU_DECLARE[x.sku]){
        x.declare = window.SKU_DECLARE[x.sku].d;
        if(!x.nameCn) x.nameCn = window.SKU_DECLARE[x.sku].n;
      }
      return x;
    });
    renderWizard();
  } else { const m=$('#pl_msg'); if(m) m.textContent='本地未收录该装箱清单，请上传 CSV 或 Excel。'; }
}

/* 解析 Excel xlsx 装箱清单：用 ExcelJS 读，支持亚马逊 ONE_SKU 导出(格式A,按箱号/箱子名称区间展开)和通用按箱展开格式 */
async function parsePackingXlsx(arrayBuffer){
  if(typeof ExcelJS==='undefined') throw new Error('ExcelJS 未加载');
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(arrayBuffer);
  const ws = wb.worksheets[0];
  if(!ws) throw new Error('Excel文件无工作表');

  // 1. 扫描前10行,找到真正的表头行（包含"SKU"+"数量/箱数"或"MSKU"+"FNSKU"）
  let headerRow = 1;
  let headerStr = '';
  for(let r=1;r<=Math.min(ws.rowCount,10);r++){
    const row=ws.getRow(r);
    let txt=''; for(let c=1;c<=row.cellCount;c++) txt += String(row.getCell(c).value||'')+'|';
    const low=txt.toLowerCase();
    // 检测多种表头模式
    if((low.includes('sku')||low.includes('msku')) &&
       (low.includes('数量')||low.includes('qty')||low.includes('件数')||low.includes('箱数'))){
      headerRow=r; headerStr=txt; break;
    }
  }

  // 2. 读表头
  const hdr=ws.getRow(headerRow);
  const headers=[]; for(let c=1;c<=Math.max(hdr.cellCount,30);c++) headers.push(String(hdr.getCell(c).value||'').trim());
  const lower=headers.map(h=>h.toLowerCase());

  const findCol=cands=>{const i=lower.findIndex(h=>cands.some(k=>h===k||h.includes(k))); return i>=0?i+1:0;};
  const col={
    sku: findCol(['msku','sku','型号','产品型号']),
    fnsku: findCol(['fnsku']),
    nameCn: findCol(['申报中文名','中文品名','中文名称','品名','中文','名称']),
    nameEn: findCol(['英文品名','英文名称','商品名称','英文','nameen','english','product name']),
    qty: findCol(['发货量','已装量','数量','qty','quantity','商品总数']),
    qtyPerBox: findCol(['单箱数量','每箱数量','每箱件数']),
    boxes: findCol(['箱子总数','箱数','boxes','ctns','cartons']),
    boxSpec: findCol(['箱子型号','箱规','boxspec','box spec']),
    boxWeight: findCol(['箱子毛重','单箱毛重','箱重','重量','weight']),
    len: findCol(['箱子长度','长','length','l']),
    wid: findCol(['箱子宽度','宽','width','w']),
    hgt: findCol(['箱子高度','高','height','h']),
    boxNo: findCol(['箱号','boxno','box no','fba箱号']),
    boxLabel: findCol(['箱子名称','箱标签','boxlabel','label']),
    asin: findCol(['asin']),
    brand: findCol(['品牌','brand']),
  };

  // 3. 格式检测
  const hasWorkingWorkflow = col.asin>0 && col.boxNo>0 && !col.boxLabel; // 亚马逊"原厂包装发货"格式
  const hasAmazonOneSKU = col.qtyPerBox>0 && col.boxLabel>0; // 亚马逊 ONE_SKU_NO_PIC
  const hasGenericFormat = col.sku>0 && col.qty>0;

  if(!col.sku && !col.boxNo){
    throw new Error(`未识别表头。请确认 Excel 包含 SKU/MSKU/箱号/装箱清单 等列。当前表头: ${headers.slice(0,10).join(', ')}`);
  }

  const out=[];
  const getStr=(row,c)=> c>0 ? String(row.getCell(c).value||'').trim() : '';
  const getNum=(row,c)=> c>0 ? (parseFloat(row.getCell(c).value)||'') : '';

  // 解析 FBA 箱号区间: "FBA19J6FCXNKU000001～2；" → [箱号1,箱号2]
  const parseBoxRange=(str)=>{
    str=str.replace(/[；;]$/,'').trim();
    const m=str.match(/^(.+?)(\d+)～(\d+)$/);
    if(m){
      const prefix=m[1], start=parseInt(m[2]), end=parseInt(m[3]);
      const pad=m[2].length;
      const arr=[]; for(let i=start;i<=end;i++) arr.push(prefix+String(i).padStart(pad,'0'));
      return arr;
    }
    return str ? [str] : [];
  };
  // 解析箱标签: "P2 - B1～B2" → [B1,B2]（跳过托盘号）
  const parseLabelRange=(str)=>{
    if(!str) return [];
    const last = str.includes(' - ') ? str.split(' - ').pop() : str;
    const m=last.match(/^([A-Za-z]?)(\d+)\s*[～~\-]\s*([A-Za-z]?\d+)$/);
    if(m){
      const prefix=m[1]||'B';
      const start=parseInt(m[2]), end=parseInt(m[3].replace(/^[A-Za-z]/,''));
      const arr=[]; for(let i=start;i<=end;i++) arr.push(prefix+i);
      return arr;
    }
    if(last.match(/^[A-Za-z]?\d+$/)) return [last];
    return last ? [last] : [];
  };

  for(let r=headerRow+1;r<=ws.rowCount;r++){
    const row=ws.getRow(r);
    const vals=[]; for(let c=1;c<=Math.max(row.cellCount,30);c++) vals.push(row.getCell(c).value);
    // 跳过全空行
    if(vals.every(v=>v===null||v===undefined||v==='')) continue;

    const sku=getStr(row,col.sku);
    const fnsku=getStr(row,col.fnsku);
    const nameEn=getStr(row,col.nameEn);
    const nameCn=getStr(row,col.nameCn);
    const boxWeight=getNum(row,col.boxWeight);
    const len=getNum(row,col.len);
    const wid=getNum(row,col.wid);
    const hgt=getNum(row,col.hgt);

    const base={sku,fnsku,nameCn,nameEn,brand:getStr(row,col.brand)||'JW PEI',boxSpec:getStr(row,col.boxSpec),boxWeight,len,wid,hgt,material:'',hs:'',model:sku,elec:'N',magnet:'N',saleUrl:'',declare:'',cost:''};

    if(hasAmazonOneSKU){
      // 亚马逊 ONE_SKU_NO_PIC 格式: 一行一个SKU, 展开多箱
      const qtyPerBox=parseFloat(getStr(row,col.qtyPerBox))||0;
      const boxNos=parseBoxRange(getStr(row,col.boxNo));
      const boxLabels=parseLabelRange(getStr(row,col.boxLabel));
      const count=Math.max(boxNos.length, boxLabels.length, parseInt(getStr(row,col.boxes))||1);
      for(let k=0;k<count;k++){
        const item={...base, boxNo:boxNos[k]||`box${k+1}`, boxLabel:boxLabels[k]||'', qty:qtyPerBox};
        out.push(item);
      }
    } else if(hasWorkingWorkflow){
      // 亚马逊"原厂包装发货"格式: 行末"箱号"列是逗号分隔的箱号列表
      const qtyPerBox=parseFloat(getStr(row,col.qtyPerBox))||parseFloat(getStr(row,col.qty))||0;
      const boxNosStr=getStr(row,col.boxNo);
      const boxNos=boxNosStr ? boxNosStr.split(',').map(s=>s.trim()).filter(Boolean) : [];
      if(boxNos.length===0){
        // 没有箱号就当一行一箱
        out.push({...base, boxNo:'', boxLabel:'', qty:qtyPerBox||1});
      } else {
        for(const bn of boxNos){
          out.push({...base, boxNo:bn, boxLabel:bn, qty:qtyPerBox});
        }
      }
    } else {
      // 通用格式: 一行一箱,或一行多箱
      const qty=parseFloat(getStr(row,col.qty))||1;
      const boxNoStr=getStr(row,col.boxNo);
      // 如果箱号是逗号分隔,也展开
      if(boxNoStr && boxNoStr.includes(',')){
        for(const bn of boxNoStr.split(',').map(s=>s.trim()).filter(Boolean)){
          out.push({...base, boxNo:bn, boxLabel:getStr(row,col.boxLabel), qty});
        }
      } else {
        out.push({...base, boxNo:boxNoStr, boxLabel:getStr(row,col.boxLabel), qty});
      }
    }
  }

  if(out.length===0){
    throw new Error('解析后无数据行。请检查 Excel 格式。');
  }

  // 同步本地主数据(品名/HS/申报价/箱规格反查)
  try{
    const norm=s=>(s||'').replace(/@us$/i,'').trim();
    for(const it of out){
      const sk=(W.skus||[]).find(x=>x.sku===it.sku);
      if(sk){ it.nameCn = it.nameCn || sk.中文品名; it.nameEn = it.nameEn || sk.英文品名; it.hs = it.hs || sk.HS; it.material = it.material || sk.材质; it.declare = it.declare || sk.申报价; }
      if((!it.declare||it.declare==='') && window.SKU_DECLARE && window.SKU_DECLARE[it.sku]){
        it.declare = window.SKU_DECLARE[it.sku].d;
        if(!it.nameCn) it.nameCn = window.SKU_DECLARE[it.sku].n;
      }
      const bs=(W.boxspecs||[]).find(x=>norm(x.sku)===norm(it.sku));
      if(bs){ it.boxSpec = it.boxSpec || bs.model; if(!it.boxWeight) it.boxWeight = bs.weight; if(!it.len) it.len = bs.l; if(!it.wid) it.wid = bs.w; if(!it.hgt) it.hgt = bs.h; }
    }
  }catch(e){ console.warn('parsePackingXlsx 主数据同步跳过:', e); }
  return out;
}

function step4(box){
  const tmpls = W.templates;
  const defT = tmpls.find(t=>t.物流商===W.form.物流商) || tmpls[0];
  if(defT && !W.selTmpl) W.selTmpl = defT.id;
  box.innerHTML = `
  <div class="card">
    <h3>④ 选模板 · 预览映射</h3>
    <div class="hint">选一个 ACTIVE 模板（已内置 5 家各自字段映射）。下方展示「字段 → 取值来源」。绿=主数据反查，白=手填，黄=推算。模板只定格子位置，值来自 L4。</div>
    <label>模板（物流商）</label>
    <select id="selTmpl">${tmpls.length? tmpls.map(t=>`<option value="${t.id}" ${t.id===W.selTmpl?'selected':''}>${esc(t.物流商)} (v${t.版本||1}, ${t.状态})</option>`).join('') : '<option>（无可用模板，去「模板库」上传）</option>'}</select>
    <div id="mapPreview" style="margin-top:14px"></div>
  </div>
  <div style="margin-top:14px;display:flex;gap:10px"><button class="btn secondary" id="prev4">← 上一步</button><button class="btn" id="next4">下一步：校验反查 →</button></div>`;
  const preview = ()=>{
    const t = tmpls.find(x=>x.id===$('#selTmpl').value);
    if(!t){ $('#mapPreview').innerHTML='<div class="empty">无模板</div>'; return; }
    W.selTmpl = t.id;
    const mp = t.mapping||{};
    const rows = Object.keys(W.sources).filter(k=>mp.meta&&mp.meta[k]).map(k=>{
      const s=W.sources[k]; const cls=(s.src==='channel'||s.src==='warehouse'||s.src==='template')?'cell-src':'cell-manual';
      return `<tr><td>${k}</td><td class="${cls}">${esc(s.v)||'<span class=muted>—</span>'}</td><td>${s.src}</td><td>${mp.meta[k]}</td></tr>`;
    }).join('');
    $('#mapPreview').innerHTML = `<table><thead><tr><th>收货人字段</th><th>值</th><th>来源</th><th>模板格子</th></tr></thead><tbody>${rows||'<tr><td colspan=4 class=empty>该模板无对应收货人字段</td></tr>'}</tbody></table>
      <div class="hint">物品行将按模板第 ${mp.itemStartRow||'?'} 行起逐行填入（箱号/品名/数量/申报价/材质/HS/品牌/型号等，按该模板实际列映射）。</div>`;
  };
  $('#selTmpl').onchange=preview;
  preview();
  $('#prev4').onclick=()=>{W.step=3;renderWizard();};
  $('#next4').onclick=()=>{ if(!W.selTmpl){ alert('请先选择一个模板'); return;} W.step=5; renderWizard(); };
}
function step5(box){
  const checks = runChecks();
  W.checks = checks;
  const passAll = checks.every(c=>c.level!=='err');
  box.innerHTML = `
  <div class="card">
    <h3>⑤ 校验反查（质量命门）</h3>
    <div class="hint">独立重读对账：必填完整性 / 勾稽（箱数=物品行数、数量合计）/ 源忠实高亮。告警+阻断：有红色错误须先修。</div>
    ${checks.map(c=>{ const cls=c.level==='err'?'alert-err':(c.level==='warn'?'alert-warn':'alert-ok'); const icon=c.level==='err'?'⛔':(c.level==='warn'?'⚠️':'✅'); return `<div class="alert ${cls}">${icon} <b>${esc(c.name)}</b>：${esc(c.msg)}</div>`; }).join('')}
    <div style="margin-top:10px"><b>勾稽汇总：</b>物品行数=${W.form.items.length}，数量合计=${checks.reduce((a,c)=>a+(c.qtySum||0),0)}，申报总值=$${checks.reduce((a,c)=>a+(c.decSum||0),0).toFixed(2)}</div>
  </div>
  <div style="margin-top:14px;display:flex;gap:10px"><button class="btn secondary" id="prev5">← 上一步</button><button class="btn ${passAll?'':'secondary'}" id="next5" ${passAll?'':'disabled'}>${passAll?'下一步：人审·交付 →':'请先修复红色错误'}</button></div>`;
  $('#prev5').onclick=()=>{W.step=4;renderWizard();};
  $('#next5').onclick=()=>{ if(passAll){W.step=6;renderWizard();} };
}
function runChecks(){
  const out=[];
  const reqMeta=['country','company','address','zip','warehouseCode'];
  const missing = reqMeta.filter(k=> !W.sources[k] || !String(W.sources[k].v).trim());
  if(missing.length) out.push({level:'err',name:'必填完整性',msg:'以下字段为空：'+missing.join('、')});
  else out.push({level:'ok',name:'必填完整性',msg:'收货人关键字段均已填'});
  let qtySum=0, decSum=0, itemErr=0;
  W.form.items.forEach((it,i)=>{
    // 校验时反查 SKU 主数据（与 step3 渲染一致：原始空 → 用 sk.申报价/sk.中文品名）
    const sk = W.skus.find(s=>s.sku===it.sku);
    const effDeclare = (it.declare!==''&&it.declare!=null) ? it.declare : (sk && sk.申报价 ? sk.申报价 : (window.SKU_DECLARE && window.SKU_DECLARE[it.sku] ? window.SKU_DECLARE[it.sku].d : ''));
    const effNameCn = it.nameCn || (sk ? sk.中文品名 : (window.SKU_DECLARE && window.SKU_DECLARE[it.sku] ? window.SKU_DECLARE[it.sku].n : ''));
    if(!it.boxNo||!effNameCn||!it.qty||!(effDeclare!==''&&effDeclare!=null)) itemErr++;
    qtySum+=parseFloat(it.qty)||0;
    decSum+=(parseFloat(effDeclare)||0)*(parseFloat(it.qty)||0);
  });
  if(itemErr) out.push({level:'err',name:'物品必填',msg:`有 ${itemErr} 行缺字段。商品申报信息表未收录的 SKU 需手填申报价（详见物品表的红色"待手填"标记），或在飞书「商品申报信息」表补充这些 SKU 的成本价后重烤。`});
  else out.push({level:'ok',name:'物品必填',msg:`${W.form.items.length} 行物品均完整`});
  const boxes=[...new Set(W.form.items.map(it=>it.boxNo).filter(Boolean))];
  // qtySum/decSum 挂到该项上,让外层 reduce 能拿到(之前挂在 out 顶层是 bug)
  out.push({level:'ok',name:'勾稽·箱数',msg:`去重箱号 ${boxes.length} 个，物品行数 ${W.form.items.length} 行（逐箱多 SKU 属正常）`, qtySum, decSum});
  const calcRows = W.form.items.filter(it=>{
    const sk=W.skus.find(s=>s.sku===it.sku);
    const hasDeclare = (it.declare!==''&&it.declare!=null) || (sk&&sk.申报价);
    return !hasDeclare && it.cost!=='';
  }).length;
  if(calcRows) out.push({level:'warn',name:'推算申报价',msg:`${calcRows} 行无 SKU 主数据申报价，按成本×${COEFF}推算（标黄），需人审确认`});
  else out.push({level:'ok',name:'申报价来源',msg:'申报价均有 SKU 主数据支撑'});
  return out;
}
async function step6(box){
  const t = W.templates.find(x=>x.id===W.selTmpl);
  box.innerHTML = `
  <div class="card">
    <h3>⑥ 人审闸门 · 交付</h3>
    <div class="alert alert-warn">⚠️ <b>生成 ≠ 发送</b>。本系统未与物流商打通，默认交付=导出 Excel（物流商导入其系统）。发送为可选、可跳过，须先勾选人审确认。</div>
    <label style="margin-top:10px"><input type="checkbox" id="humanOk" style="width:auto;margin-right:8px">我已核对源数据、映射与勾稽结果，确认无误</label>
    <div id="deliverBtns" style="margin-top:14px;display:flex;gap:10px;opacity:.5;pointer-events:none">
      <button class="btn green" id="exportBtn">⬇ 导出 Excel 交付（默认）</button>
      <button class="btn secondary" id="sendBtn">✉ 发送给物流商（可选·未集成可跳过）</button>
    </div>
    <div id="genLog" style="margin-top:12px"></div>
  </div>
  <div style="margin-top:14px"><button class="btn secondary" id="prev6">← 上一步</button></div>`;
  $('#humanOk').onchange = e=>{ const on=e.target.checked; const b=$('#deliverBtns'); b.style.opacity=on?'1':'0.5'; b.style.pointerEvents=on?'auto':'none'; };
  $('#prev6').onclick=()=>{W.step=5;renderWizard();};
  $('#exportBtn').onclick = async ()=>{
    const log=$('#genLog'); log.innerHTML='<div class="alert alert-warn">⏳ 正在用 ExcelJS 填模板副本…</div>';
    try{
      const blob = await generateInvoice(t);
      downloadBlob(blob, `发票_${W.form.物流商}_${W.form.仓库代码}_${W.form.fbaNo||'draft'}.xlsx`);
      await put('records',{id:uid(),时间:new Date().toISOString(),物流商:W.form.物流商,渠道:W.form.渠道,仓库:W.form.仓库代码,fba:W.form.fbaNo,模板:t.id,状态:'DELIVERED(导出)'});
      log.innerHTML='<div class="alert alert-ok">✅ 已导出填好的 Excel（保留原模板样式/合并/图片公式）。可在「校验·监控」看记录。</div>';
    }catch(err){ log.innerHTML='<div class="alert alert-err">❌ 生成失败：'+esc(err.message)+'</div>'; }
  };
  $('#sendBtn').onclick = ()=>{ $('#genLog').innerHTML='<div class="alert alert-warn">ℹ️ 发送适配器未集成（物流商系统未打通）。本步可跳过，已导出 Excel 即可交付。</div>'; };
}
async function generateInvoice(tmpl){
  if(typeof ExcelJS==='undefined') throw new Error('ExcelJS 未加载');
  const wb = new ExcelJS.Workbook();
  const buf = await tmpl.blob.arrayBuffer();
  await wb.xlsx.load(buf);
  const ws = wb.getWorksheet(1);
  const M = tmpl.mapping;
  if(!M) throw new Error('该模板无字段映射');
  if(M.titleCell && M.titleText) ws.getCell(M.titleCell).value = M.titleText;
  // 收货人块
  if(M.meta) Object.entries(M.meta).forEach(([k,cell])=>{ const s=W.sources[k]; if(s && (s.v||s.v===0)) ws.getCell(cell).value = s.v; });
  // 物品行
  W.form.items.forEach((it,i)=>{
    const r = (M.itemStartRow||21) + i;
    if(M.item) Object.entries(M.item).forEach(([fld,col])=>{
      const v = it[fld];
      if(v||v===0){ const num = (fld==='qty'||fld==='declare'||fld==='boxWeight'||fld==='len'||fld==='wid'||fld==='hgt'||fld==='boxCount'||fld==='prodWeight'); ws.getCell(col+r).value = num?parseFloat(v):v; }
    });
    // 安速等带币种/原产地固定列
    if(M.item && M.item.currency) ws.getCell(M.item.currency+r).value='USD';
    if(M.item && M.item.origin) ws.getCell(M.item.origin+r).value='CN';
  });
  return await wb.xlsx.writeBuffer();
}
function downloadBlob(blob,name){
  const url=URL.createObjectURL(new Blob([blob]));
  const a=document.createElement('a'); a.href=url; a.download=name; a.click();
  setTimeout(()=>URL.revokeObjectURL(url),3000);
}

/* ============================================================
   渠道·收货人主数据 (L4)
   ============================================================ */
async function channels(){
  const list = await getAll('channels');
  main().innerHTML = `
  <h2>渠道·收货人主数据</h2>
  <div class="sub">L4 配置与资源层。这里维护「值」，生成时由向导反查，向导内只读。改一处、全量一致，易错录入收敛到此。</div>
  <div class="card">
    <button class="btn" id="addCh">+ 新增渠道</button>
    <table style="margin-top:12px"><thead><tr><th>物流商</th><th>渠道</th><th>国家</th><th>VAT</th><th>EORI</th><th>仓库</th><th></th></tr></thead>
    <tbody id="chBody">${list.map(c=>`<tr><td>${esc(c.物流商)}</td><td>${esc(c.渠道)}</td><td>${esc(c.国家)}</td><td>${esc(c.VAT)}</td><td>${esc(c.EORI)}</td><td>${esc((c.仓库||[]).map(w=>w.代码).join(', '))}</td>
      <td><button class="btn secondary" data-edit="${c.id}" style="padding:4px 8px">编辑</button> <button class="btn danger" data-del="${c.id}" style="padding:4px 8px">删</button></td></tr>`).join('')}</tbody></table>
  </div>
  <div id="chEditor"></div>`;
  $('#addCh').onclick=()=>editChannel(null);
  $$('[data-edit]').forEach(b=> b.onclick=()=>editChannel(b.dataset.edit));
  $$('[data-del]').forEach(b=> b.onclick=async()=>{ if(confirm('确认删除该渠道？')){ await del('channels',b.dataset.del); channels(); } });
}
async function editChannel(id){
  const all = await getAll('channels');
  const c = id? all.find(x=>x.id===id) : {id:uid(),物流商:'',渠道:'',国家:'',VAT:'',EORI:'',注册名:'',注册地址:'',仓库:[{代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}]};
  const box = $('#chEditor');
  box.innerHTML = `
  <div class="card" style="border-color:var(--accent)">
    <h3>${id?'编辑':'新增'}渠道</h3>
    <div class="row">
      <div><label>物流商</label><input id="c_物流商" value="${esc(c.物流商)}"></div>
      <div><label>渠道</label><input id="c_渠道" value="${esc(c.渠道)}"></div>
      <div><label>国家</label><input id="c_国家" value="${esc(c.国家)}"></div>
    </div>
    <div class="row">
      <div><label>VAT号</label><input id="c_VAT" value="${esc(c.VAT)}"></div>
      <div><label>EORI</label><input id="c_EORI" value="${esc(c.EORI)}"></div>
      <div><label>VAT注册名</label><input id="c_注册名" value="${esc(c.注册名)}"></div>
    </div>
    <label>VAT注册地址</label><input id="c_注册地址" value="${esc(c.注册地址)}">
    <h3 style="margin-top:18px">仓库子表（按仓库代码反查地址）</h3>
    <div id="whList"></div>
    <button class="btn secondary" id="addWh" style="margin-top:8px">+ 加仓库</button>
    <div style="margin-top:14px;display:flex;gap:10px"><button class="btn" id="saveCh">保存</button><button class="btn secondary" id="cancelCh">取消</button></div>
  </div>`;
  const renderWh = ()=>{
    $('#whList').innerHTML = (c.仓库||[]).map((w,i)=>`
      <div class="row" style="margin:6px 0;align-items:end">
        <div><label>代码</label><input data-w="${i}" data-f="代码" value="${esc(w.代码)}"></div>
        <div><label>公司</label><input data-w="${i}" data-f="公司" value="${esc(w.公司)}"></div>
        <div><label>省份</label><input data-w="${i}" data-f="省份" value="${esc(w.省份)}"></div>
        <div><label>城市</label><input data-w="${i}" data-f="城市" value="${esc(w.城市)}"></div>
        <div><label>地址</label><input data-w="${i}" data-f="地址" value="${esc(w.地址)}"></div>
        <div><label>邮编</label><input data-w="${i}" data-f="邮编" value="${esc(w.邮编)}"></div>
        <div><label>电话</label><input data-w="${i}" data-f="电话" value="${esc(w.电话)}"></div>
        <div><button class="btn danger" data-wdel="${i}" style="padding:7px 10px">×</button></div>
      </div>`).join('');
    $$('#whList [data-w]').forEach(inp=> inp.oninput=e=>{ c.仓库[+e.target.dataset.w][e.target.dataset.f]=e.target.value; });
    $$('#whList [data-wdel]').forEach(b=> b.onclick=()=>{ c.仓库.splice(+b.dataset.wdel,1); renderWh(); });
  };
  renderWh();
  $('#addWh').onclick=()=>{ c.仓库.push({代码:'',公司:'',省份:'',城市:'',地址:'',邮编:'',电话:''}); renderWh(); };
  $('#cancelCh').onclick=()=>{ box.innerHTML=''; };
  $('#saveCh').onclick=async()=>{
    c.物流商=$('#c_物流商').value; c.渠道=$('#c_渠道').value; c.国家=$('#c_国家').value; c.VAT=$('#c_VAT').value; c.EORI=$('#c_EORI').value; c.注册名=$('#c_注册名').value; c.注册地址=$('#c_注册地址').value;
    c.仓库=c.仓库.filter(w=>w.代码);
    await put('channels',c); box.innerHTML=''; channels();
  };
}

/* ============================================================
   SKU 主数据 (L4)
   ============================================================ */
async function skus(){
  const list = await getAll('skus');
  main().innerHTML = `
  <h2>SKU 主数据</h2>
  <div class="sub">申报价带版本号：变动追加新版本（生效日/原因），不覆盖；发票快照可复验。</div>
  <div class="card">
    <button class="btn" id="addSk">+ 新增 SKU</button>
    <table style="margin-top:12px"><thead><tr><th>SKU</th><th>中文品名</th><th>英文</th><th>材质</th><th>HS</th><th>品牌</th><th>型号</th><th>申报价</th><th>版本</th><th></th></tr></thead>
    <tbody>${list.map(s=>`<tr><td>${esc(s.sku)}</td><td>${esc(s.中文品名)}</td><td>${esc(s.英文品名)}</td><td>${esc(s.材质)}</td><td>${esc(s.HS)}</td><td>${esc(s.品牌)}</td><td>${esc(s.型号)}</td><td>${esc(s.申报价)}</td><td>${esc((s.版本||[]).length)}</td>
      <td><button class="btn secondary" data-edit="${s.id}" style="padding:4px 8px">编辑</button> <button class="btn danger" data-del="${s.id}" style="padding:4px 8px">删</button></td></tr>`).join('')}</tbody></table>
  </div>
  <div id="skEditor"></div>`;
  $('#addSk').onclick=()=>editSku(null);
  $$('[data-edit]').forEach(b=> b.onclick=()=>editSku(b.dataset.edit));
  $$('[data-del]').forEach(b=> b.onclick=async()=>{ if(confirm('确认删除？')){ await del('skus',b.dataset.del); skus(); } });
}
async function editSku(id){
  const all=await getAll('skus');
  const s=id?all.find(x=>x.id===id):{id:uid(),sku:'',中文品名:'',英文品名:'',材质:'',HS:'',品牌:'',型号:'',申报价:'',成本:'',版本:[],图片:''};
  const box=$('#skEditor');
  box.innerHTML=`
  <div class="card" style="border-color:var(--accent)">
    <h3>${id?'编辑':'新增'}SKU</h3>
    <div class="row">
      <div><label>SKU</label><input id="s_sku" value="${esc(s.sku)}"></div>
      <div><label>中文品名</label><input id="s_cn" value="${esc(s.中文品名)}"></div>
      <div><label>英文品名</label><input id="s_en" value="${esc(s.英文品名)}"></div>
    </div>
    <div class="row">
      <div><label>材质</label><input id="s_mat" value="${esc(s.材质)}"></div>
      <div><label>HS编码</label><input id="s_hs" value="${esc(s.HS)}"></div>
      <div><label>品牌</label><input id="s_br" value="${esc(s.品牌)}"></div>
      <div><label>型号</label><input id="s_md" value="${esc(s.型号)}"></div>
    </div>
    <div class="row">
      <div><label>申报价(USD)</label><input id="s_dec" value="${esc(s.申报价)}"></div>
      <div><label>成本(USD)</label><input id="s_cost" value="${esc(s.成本)}"></div>
      <div><label>版本原因</label><input id="s_reason" placeholder="如：成本上涨调申报价"></div>
    </div>
    <div style="margin-top:14px;display:flex;gap:10px"><button class="btn" id="saveSk">保存</button><button class="btn secondary" id="cancelSk">取消</button></div>
  </div>`;
  $('#cancelSk').onclick=()=>box.innerHTML='';
  $('#saveSk').onclick=async()=>{
    const newDec=$('#s_dec').value, oldDec=s.申报价;
    s.sku=$('#s_sku').value; s.中文品名=$('#s_cn').value; s.英文品名=$('#s_en').value; s.材质=$('#s_mat').value; s.HS=$('#s_hs').value; s.品牌=$('#s_br').value; s.型号=$('#s_md').value; s.成本=$('#s_cost').value;
    if(newDec!==oldDec){ s.版本=s.版本||[]; s.版本.push({v:(s.版本.length+1),值:newDec,生效日:new Date().toISOString().slice(0,10),原因:$('#s_reason').value||'更新'}); }
    s.申报价=newDec;
    await put('skus',s); box.innerHTML=''; skus();
  };
}

/* ============================================================
   模板库 (L4)
   ============================================================ */
async function templates(){
  const list = await getAll('templates');
  main().innerHTML = `
  <h2>模板库</h2>
  <div class="sub">上传空白模板 xlsx（存本地 IndexedDB）。模板=排版层；版本迭代可停用旧版（状态 ACTIVE/DISABLED/DEPRECATED），停模板≠丢数据。</div>
  <div class="card">
    <h3>上传新模板</h3>
    <div class="hint">v1 已内置 5 家物流商映射。上传同结构模板可直接复用；其它物流商映射待补（在 MAPPINGS 增加即可）。</div>
    <div class="row">
      <div><label>物流商</label><input id="t_物流商" value="安速"></div>
      <div><label>渠道</label><input id="t_渠道" value="美国包税海卡(正班)"></div>
      <div><label>空白模板 xlsx</label><input type="file" id="t_file" accept=".xlsx"></div>
    </div>
    <button class="btn" id="upTmpl" style="margin-top:10px">上传并入库</button>
    <div id="upLog" style="margin-top:8px"></div>
  </div>
  <div class="card">
    <h3>已有模板</h3>
    <table><thead><tr><th>物流商</th><th>渠道</th><th>版本</th><th>状态</th><th>创建日</th><th>操作</th></tr></thead>
    <tbody>${list.length?list.map(t=>`
      <tr><td>${esc(t.物流商)}</td><td>${esc(t.渠道)}</td><td>v${t.版本||1}</td><td><span class="pill ${t.状态==='ACTIVE'?'pill-green':(t.状态==='DISABLED'?'pill-gray':'pill-yellow')}">${t.状态}</span></td><td>${esc(t.创建日||'')}</td>
      <td>
        <button class="btn secondary" data-toggle="${t.id}" style="padding:4px 8px">${t.状态==='ACTIVE'?'停用':'启用'}</button>
        ${t.状态==='ACTIVE'?'<button class="btn secondary" data-deprec="'+t.id+'" style="padding:4px 8px">迭代弃用</button>':''}
        <button class="btn danger" data-del="${t.id}" style="padding:4px 8px">删</button>
      </td></tr>`).join(''):'<tr><td colspan=6 class="empty">暂无模板</td></tr>'}</tbody></table>
  </div>`;
  $('#upTmpl').onclick=async()=>{
    const f=$('#t_file').files[0];
    if(!f){ $('#upLog').innerHTML='<div class="alert alert-err">请选择 xlsx 文件</div>'; return; }
    const key=$('#t_物流商').value;
    const blob=new Blob([await f.arrayBuffer()],{type:f.type});
    const rec={id:uid(),物流商:key,渠道:$('#t_渠道').value,名称:f.name,blob,状态:'ACTIVE',版本:1,创建日:new Date().toISOString().slice(0,10),mapping:MAPPINGS[key]||MAPPINGS['安速']};
    await put('templates',rec); $('#upLog').innerHTML='<div class="alert alert-ok">✅ 已入库</div>'; templates();
  };
  $$('[data-toggle]').forEach(b=> b.onclick=async()=>{ const t=list.find(x=>x.id===b.dataset.toggle); t.状态=t.状态==='ACTIVE'?'DISABLED':'ACTIVE'; await put('templates',t); templates(); });
  $$('[data-deprec]').forEach(b=> b.onclick=async()=>{ const t=list.find(x=>x.id===b.dataset.deprec); t.状态='DEPRECATED'; await put('templates',t); templates(); });
  $$('[data-del]').forEach(b=> b.onclick=async()=>{ if(confirm('确认删除模板？')){ await del('templates',b.dataset.del); templates(); } });
}

/* ============================================================
   校验·监控 (L6)
   ============================================================ */
async function monitor(){
  const recs = await getAll('records');
  main().innerHTML = `
  <h2>校验·监控</h2>
  <div class="sub">质量规则总览 + 生成记录（溯源审计）。每条记录带时间/物流商/模板，可复验。</div>
  <div class="card">
    <h3>质量规则</h3>
    <table><thead><tr><th>规则</th><th>类型</th><th>说明</th></tr></thead><tbody>
      <tr><td>必填完整性</td><td><span class="pill pill-red">阻断</span></td><td>收货人关键字段 + 物品必填项为空则阻断</td></tr>
      <tr><td>勾稽·箱数/数量</td><td><span class="pill pill-red">阻断</span></td><td>箱号/数量合计与装箱单一致</td></tr>
      <tr><td>申报价来源</td><td><span class="pill pill-yellow">告警</span></td><td>无主数据按下推算(成本×${COEFF})需人审确认</td></tr>
      <tr><td>源忠实</td><td><span class="pill pill-green">提示</span></td><td>推算值标黄、主数据反查标绿，便于复核</td></tr>
      <tr><td>人审闸门</td><td><span class="pill pill-red">阻断</span></td><td>未勾选确认不得导出/发送</td></tr>
    </tbody></table>
  </div>
  <div class="card">
    <h3>生成记录（${recs.length}）</h3>
    ${recs.length?`<table><thead><tr><th>时间</th><th>物流商</th><th>渠道</th><th>仓库</th><th>FBA号</th><th>状态</th></tr></thead><tbody>
      ${recs.slice().reverse().map(r=>`<tr><td>${esc(r.时间)}</td><td>${esc(r.物流商)}</td><td>${esc(r.渠道)}</td><td>${esc(r.仓库)}</td><td>${esc(r.fba)}</td><td><span class="pill pill-green">${esc(r.状态)}</span></td></tr>`).join('')}</tbody></table>`
      :'<div class="empty">暂无生成记录，去「生成发票向导」产出第一票。</div>'}
  </div>`;
}

/* ---------- 启动 ---------- */
(async function init(){
  const status = document.getElementById('dbStatus');
  try{
    await openDB();
    await seedIfEmpty();
    status.textContent = USE_DB ? '存储: 本地 IndexedDB ✓' : '存储: 内存模式(IndexedDB不可用)';
    status.style.color = USE_DB ? 'var(--green)' : 'var(--warn)';
    go('overview');
  }catch(e){
    status.textContent = '存储: 异常，已降级';
    console.error(e);
    try{ go('overview'); }catch(_){}
    // 兜底：保证 main 永不空白（即使 overview 也抛错也能看到具体异常）
    if(main() && !main().innerHTML){
      main().innerHTML = '<div class="alert alert-err" style="margin:20px"><b>⚠️ 初始化异常，已降级运行</b><br>'+
        '<pre style="white-space:pre-wrap;color:#ff9;background:#222;padding:8px;margin-top:8px;border-radius:4px">'+
        esc((e&&e.stack)?e.stack+'\n\n[name:'+e.name+', message:'+e.message+']':String(e))+
        '</pre>'+
        '<div style="margin-top:10px;color:#bbb">常见原因：浏览器隐私模式禁用 IndexedDB / 旧版本 store schema 不兼容。请按 <b>Ctrl+Shift+R</b> 硬刷新一次；仍不行请打开开发者工具 (F12) 把 Console 错误发我。</div></div>';
    }
  }
})();
