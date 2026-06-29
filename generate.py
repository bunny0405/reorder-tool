"""
追加下單建議工具 generate.py
每週一 13:00 由 GitHub Actions 自動執行，或手動執行
產出 index.html 部署至 GitHub Pages
"""

import os, json, glob, re
from datetime import datetime, timedelta
import pandas as pd
import numpy as np

BASE_DIR  = os.path.dirname(os.path.abspath(__file__))
DATA_DIR  = os.path.join(BASE_DIR, 'data')
TODAY     = datetime.today().date()

# ════════════════════════════════════════════════════════
# 1. 工具函式
# ════════════════════════════════════════════════════════

def read_file(path):
    if path.endswith('.csv'):
        for enc in ['utf-8-sig','utf-8','cp950','big5']:
            try: return pd.read_csv(path, encoding=enc, dtype=str)
            except: continue
    return pd.read_excel(path, dtype=str)

def latest_file(pattern):
    files = glob.glob(os.path.join(DATA_DIR, pattern))
    if not files: return None
    def extract_date(f):
        m = re.search(r'(\d{8})', os.path.basename(f))
        return m.group(1) if m else '00000000'
    return sorted(files, key=extract_date)[-1]

def all_files(pattern):
    return sorted(glob.glob(os.path.join(DATA_DIR, pattern)))

def safe_int(v, default=0):
    try: return int(float(str(v).replace(',','')))
    except: return default

def safe_float(v, default=0.0):
    try: return float(str(v).replace(',',''))
    except: return default

def parse_date(v):
    if pd.isna(v) or str(v).strip() in ('','nan','None'): return None
    for fmt in ('%Y/%m/%d','%Y-%m-%d','%Y%m%d'):
        try: return datetime.strptime(str(v).strip(), fmt).date()
        except: pass
    return None

def natural_week_start(d):
    return d - timedelta(days=d.weekday())

# ════════════════════════════════════════════════════════
# 2. 載入資料
# ════════════════════════════════════════════════════════
print("📂 載入資料...")

# ── 門市基本資料 ──
df_store = read_file(os.path.join(DATA_DIR, '門市基本資料.xlsx'))
df_store.columns = [str(c).strip() for c in df_store.columns]
# 移除全空欄
df_store = df_store.dropna(axis=1, how='all')
# 確認倉庫編號欄（可能是A欄或B欄，找非空的）
if df_store.columns[0] in (None,'None','') or df_store.iloc[:,0].isna().all():
    df_store = df_store.iloc[:,1:]  # 跳過空A欄
df_store.columns = [str(c).strip() for c in df_store.columns]
# 標準化欄位名稱對應
col_map = {}
for i,c in enumerate(df_store.columns):
    col_map[c] = i
store_id_col = '倉庫編號' if '倉庫編號' in df_store.columns else df_store.columns[0]
store_type_col = '門市型態' if '門市型態' in df_store.columns else None
store_line_col = '販售支線' if '販售支線' in df_store.columns else None

df_store[store_id_col] = df_store[store_id_col].astype(str).str.strip()
if store_line_col:
    df_store['_支線'] = df_store[store_line_col].fillna('')
else:
    df_store['_支線'] = ''

OUTLET_STORES = {'AS51','AS63','AT08'}
IP_STORES     = {'AS62','AT10'}

# 建立店鋪支線查詢字典
store_line_map = {}
store_type_map = {}
for _, row in df_store.iterrows():
    sid = str(row[store_id_col]).strip()
    store_line_map[sid] = str(row['_支線']).strip()
    if store_type_col:
        store_type_map[sid] = str(row[store_type_col]).strip()

all_store_ids = set(store_line_map.keys()) - {'nan','None',''}
print(f"  門市數：{len(all_store_ids)}")

# ── 商品主檔 ──
df_prod = read_file(os.path.join(DATA_DIR, '新品基本資料.xlsx'))
df_prod.columns = [str(c).strip() for c in df_prod.columns]
# A欄=新品牌/OT，G欄=IP
col_brand_a = df_prod.columns[0]
col_ip_g    = df_prod.columns[6] if len(df_prod.columns)>6 else None
col_end_date = '最後販售日' if '最後販售日' in df_prod.columns else None

df_prod['商品編號'] = df_prod['商品編號'].astype(str).str.strip()
df_prod['款號']     = df_prod['款號'].astype(str).str.strip()
df_prod['launch_date'] = df_prod['最新上架日期'].apply(parse_date)
df_prod['end_date']    = df_prod[col_end_date].apply(parse_date) if col_end_date else None
df_prod['is_ot']   = df_prod[col_brand_a].astype(str).str.strip().str.upper() == 'OT'
df_prod['is_ip']   = df_prod[col_ip_g].astype(str).str.strip().str.upper() == 'IP' if col_ip_g else False
df_prod['brand']   = df_prod[col_brand_a].astype(str).str.strip()
img_col = None
for c in df_prod.columns:
    if '圖' in str(c) or 'img' in str(c).lower() or 'photo' in str(c).lower():
        img_col = c; break
df_prod['img_url'] = df_prod[img_col].fillna('') if img_col else ''
# 商品條碼
barcode_col = None
for c in df_prod.columns:
    if '條碼' in str(c) or 'barcode' in str(c).lower():
        barcode_col = c; break
df_prod['barcode'] = df_prod[barcode_col].fillna('') if barcode_col else ''
if '年度' not in df_prod.columns: df_prod['年度'] = ''
if '季節' not in df_prod.columns: df_prod['季節'] = ''

prod_map = df_prod.set_index('商品編號').to_dict('index')
print(f"  商品數：{len(prod_map)}")

# ── 做貨備註 ──
lt_map = {}
lt_path = os.path.join(DATA_DIR, '做貨備註.xlsx')
if os.path.exists(lt_path):
    df_lt = read_file(lt_path)
    df_lt.columns = [str(c).strip() for c in df_lt.columns]
    for _, r in df_lt.iterrows():
        pid   = str(r.get('商品編號','')).strip()
        color = str(r.get('顏色','')).strip()
        days  = safe_int(r.get('做貨天數',30), 30)
        lt_map[(pid,color)] = days

def get_leadtime(pid, color):
    return lt_map.get((pid,color), lt_map.get((pid,''), 30))

# ── 採購總表（已到貨）──
arrived_wnums = set()
po_path = os.path.join(DATA_DIR, '採購總表_2026.xlsx')
if os.path.exists(po_path):
    df_po = read_file(po_path)
    df_po.columns = [str(c).strip() for c in df_po.columns]
    if '預購單號' in df_po.columns:
        arrived_wnums = set(df_po['預購單號'].dropna().astype(str).str.strip())
print(f"  已到貨小白單：{len(arrived_wnums)} 筆")

# ── 追加單 ──
transit_map = {}   # (pid,color,sz) → 在途量
total_ordered_map = {}  # (pid,color) → 已追加量
ro_path = os.path.join(DATA_DIR, '追加單_2026.xlsx')
if os.path.exists(ro_path):
    df_ro = read_file(ro_path)
    df_ro.columns = [str(c).strip() for c in df_ro.columns]
    for _, r in df_ro.iterrows():
        wnum  = str(r.get('小白單編號','')).strip()
        pid   = str(r.get('商品編號','')).strip()
        color = str(r.get('顏色名稱','')).strip()
        sz    = str(r.get('尺寸','')).strip()
        qty   = safe_int(r.get('TW_A(門市)頭單',0))
        ck    = (pid,color)
        total_ordered_map[ck] = total_ordered_map.get(ck,0) + qty
        if wnum not in arrived_wnums:
            key = (pid,color,sz)
            transit_map[key] = transit_map.get(key,0) + qty

# ── 銷售明細 ──
print("  讀取銷售資料...")
sales_files = all_files('銷售_*.csv') + all_files('銷售_*.xlsx')
df_sales_list = []
for f in sales_files:
    try:
        df_s = read_file(f)
        df_s.columns = [str(c).strip() for c in df_s.columns]
        if '品別' in df_s.columns:
            df_s = df_s[df_s['品別'].astype(str).str.strip() != '出清品']
        df_sales_list.append(df_s)
    except Exception as e:
        print(f"  ⚠ 銷售檔讀取失敗：{os.path.basename(f)}：{e}")

df_sales = pd.concat(df_sales_list, ignore_index=True) if df_sales_list else pd.DataFrame()
if not df_sales.empty:
    df_sales['商品編號'] = df_sales['商品編號'].astype(str).str.strip()
    df_sales['商品顏色'] = df_sales.get('商品顏色', pd.Series(dtype=str)).fillna('').astype(str).str.strip()
    df_sales['商品尺寸'] = df_sales.get('商品尺寸', pd.Series(dtype=str)).fillna('').astype(str).str.strip()
    df_sales['倉庫編號'] = df_sales['倉庫編號'].astype(str).str.strip()
    df_sales['訂單數量'] = df_sales['訂單數量'].apply(safe_int)
    df_sales['訂單日期'] = pd.to_datetime(df_sales['訂單日期'], errors='coerce')
    df_sales = df_sales[df_sales['訂單日期'].notna()]
    df_sales['_date'] = df_sales['訂單日期'].dt.date
print(f"  銷售筆數：{len(df_sales)}")

# ── 庫存快照 ──
print("  讀取庫存快照...")
inv_path = latest_file('周一庫存_*.xlsx') or latest_file('周一庫存_*.csv')
if inv_path:
    df_inv = read_file(inv_path)
    df_inv.columns = [str(c).strip() for c in df_inv.columns]
    df_inv['商品編號'] = df_inv['商品編號'].astype(str).str.strip()
    df_inv['倉庫編號'] = df_inv['倉庫編號'].astype(str).str.strip()
    df_inv['庫存數量'] = df_inv['庫存數量'].apply(safe_int)
    # 顏色欄
    if '顏色名稱' in df_inv.columns:
        df_inv['_color'] = df_inv['顏色名稱'].fillna('').astype(str).str.strip()
    elif '顏色' in df_inv.columns:
        df_inv['_color'] = df_inv['顏色'].fillna('').astype(str).str.strip()
    elif '商品顏色' in df_inv.columns:
        df_inv['_color'] = df_inv['商品顏色'].fillna('').astype(str).str.strip()
    else:
        df_inv['_color'] = ''
    # 尺碼欄
    sz_col = '商品尺寸' if '商品尺寸' in df_inv.columns else ('尺寸' if '尺寸' in df_inv.columns else None)
    df_inv['_size'] = df_inv[sz_col].fillna('').astype(str).str.strip() if sz_col else ''
    m = re.search(r'(\d{8})', os.path.basename(inv_path))
    inv_snap_date = m.group(1) if m else str(TODAY).replace('-','')
else:
    df_inv = pd.DataFrame()
    inv_snap_date = str(TODAY).replace('-','')

# ════════════════════════════════════════════════════════
# 3. 通路資格判斷
# ════════════════════════════════════════════════════════

def get_eligible_stores(pid):
    info = prod_map.get(pid, {})
    is_ot  = info.get('is_ot', False)
    is_ip  = info.get('is_ip', False)
    brand  = info.get('brand', '')
    eligible = set()
    for sid, line in store_line_map.items():
        if not sid or sid in ('nan','None',''): continue
        if sid in OUTLET_STORES:
            if is_ot: eligible.add(sid)
        elif sid in IP_STORES:
            if is_ip: eligible.add(sid)
        else:
            if brand and brand in line:
                eligible.add(sid)
    return eligible

# ════════════════════════════════════════════════════════
# 4. 計算指標
# ════════════════════════════════════════════════════════
print("⚙️  計算指標...")

SIZE_ORDER = ['XS','S','S+','M','M+','L','L+','XL','2XL','3XL',
              '35','36','37','38','39','40','41','42','43','F']

def sz_key(s):
    return SIZE_ORDER.index(s) if s in SIZE_ORDER else 99

this_monday  = natural_week_start(TODAY)
week_starts  = [this_monday - timedelta(weeks=i) for i in range(4,0,-1)]  # W-4→W-1

products_data = []

for _, prod in df_prod.iterrows():
    pid     = str(prod['商品編號']).strip()
    style   = str(prod.get('款號','')).strip()
    name    = str(prod.get('商品名稱','')).strip()
    cat     = str(prod.get('商品大分類','')).strip()
    brand   = str(prod.get('brand','')).strip()
    year    = str(prod.get('年度','')).strip()
    season  = str(prod.get('季節','')).strip()
    img_url = str(prod.get('img_url','')).strip()
    barcode = str(prod.get('barcode','')).strip()
    launch  = prod.get('launch_date')
    end_dt  = prod.get('end_date')

    if launch is None: continue

    eligible_stores = get_eligible_stores(pid)
    if not eligible_stores: continue

    # 銷售
    if df_sales.empty:
        p_sales = pd.DataFrame()
    else:
        p_sales = df_sales[(df_sales['商品編號']==pid) & df_sales['倉庫編號'].isin(eligible_stores)].copy()

    # 庫存
    if df_inv.empty:
        p_inv = pd.DataFrame()
    else:
        p_inv = df_inv[(df_inv['商品編號']==pid) & df_inv['倉庫編號'].isin(eligible_stores)].copy()

    # 顏色列表
    colors = []
    if not p_sales.empty:
        colors = [c for c in p_sales['商品顏色'].unique() if c and c!='nan']
    if not colors and not p_inv.empty:
        colors = [c for c in p_inv['_color'].unique() if c and c!='nan']
    if not colors: colors = ['']

    for color in colors:
        c_sales = p_sales[p_sales['商品顏色']==color] if not p_sales.empty else pd.DataFrame()
        c_inv   = p_inv[p_inv['_color']==color]       if not p_inv.empty   else pd.DataFrame()

        # 累銷
        cum_sales = int(c_sales['訂單數量'].sum()) if not c_sales.empty else 0

        # 庫存
        inv_by_sz = {}
        if not c_inv.empty:
            for _, ir in c_inv.iterrows():
                sz = str(ir['_size']).strip()
                inv_by_sz[sz] = inv_by_sz.get(sz,0) + safe_int(ir['庫存數量'])
        color_total_inv  = sum(inv_by_sz.values())
        color_outlet_inv = int(c_inv[c_inv['倉庫編號'].isin(OUTLET_STORES)]['庫存數量'].sum()) if not c_inv.empty else 0
        color_main_inv   = color_total_inv - color_outlet_inv
        # 門市現場庫存（排除大倉808）
        color_store_inv  = int(c_inv[~c_inv['倉庫編號'].str.startswith('8')]['庫存數量'].sum()) if not c_inv.empty else 0

        # 進貨量回推
        total_purchase = cum_sales + color_total_inv
        sellthru = cum_sales / total_purchase if total_purchase > 0 else 0.0

        # ── 有效 store-days ──
        if not c_sales.empty:
            sd = c_sales.groupby(['倉庫編號','_date'])['訂單數量'].sum().reset_index()
            store_days_sold = len(sd)
            # 有庫存但無銷售的店（近7天）
            inv_stores_w_stock = set(c_inv[c_inv['庫存數量']>0]['倉庫編號'].tolist()) if not c_inv.empty else set()
            extra = len(inv_stores_w_stock - set(sd['倉庫編號'].unique())) * 7
            total_store_days = store_days_sold + extra
        else:
            total_store_days = 0
        effective_spsd = cum_sales / total_store_days if total_store_days > 0 else 0.0

        # ── 供給受限 ──
        supply_warn = False
        if not c_sales.empty and launch:
            early_cut = launch + timedelta(days=14)
            early_st  = early_stores  = c_sales[c_sales['_date']<=early_cut]['倉庫編號'].nunique()
            recent_st = c_sales[c_sales['_date']>=(TODAY-timedelta(days=7))]['倉庫編號'].nunique()
            supply_warn = early_stores>0 and recent_st < early_stores*0.7

        # ── 近4週趨勢 ──
        weekly_sales = []
        for ws in week_starts:
            we = ws + timedelta(days=6)
            wqty = int(c_sales[(c_sales['_date']>=ws)&(c_sales['_date']<=we)]['訂單數量'].sum()) if not c_sales.empty else 0
            weekly_sales.append({'week': ws.strftime('%m/%d'), 'qty': wqty})

        # ── 尺碼分析 ──
        all_sizes_raw = set()
        if not c_sales.empty: all_sizes_raw.update(c_sales['商品尺寸'].unique())
        all_sizes_raw.update(inv_by_sz.keys())
        all_sizes_sorted = sorted([s for s in all_sizes_raw if s and s!='nan'], key=sz_key)

        sz_data = []
        sz_sellthru_map = {}
        for sz in all_sizes_sorted:
            sz_s = c_sales[c_sales['商品尺寸']==sz] if not c_sales.empty else pd.DataFrame()
            sz_cum   = int(sz_s['訂單數量'].sum()) if not sz_s.empty else 0
            sz_inv   = inv_by_sz.get(sz, 0)
            sz_pur   = sz_cum + sz_inv
            sz_st    = sz_cum / sz_pur if sz_pur > 0 else 0.0
            sz_sellthru_map[sz] = sz_st

            # (B) 達55%天數
            days_to_55 = None
            if not sz_s.empty and sz_pur > 0:
                target = sz_pur * 0.55
                cumsum = 0
                for _, sr in sz_s.sort_values('訂單日期').iterrows():
                    cumsum += safe_int(sr['訂單數量'])
                    if cumsum >= target:
                        days_to_55 = (sr['_date'] - launch).days + 1
                        break
            # (C) 尺碼有效日均銷
            sz_sd = len(sz_s.groupby(['倉庫編號','_date'])) if not sz_s.empty else 0
            sz_spsd = sz_cum / sz_sd if sz_sd > 0 else 0.0

            sz_data.append({
                'sz': sz, 'inv': sz_inv, 'cum_sales': sz_cum,
                'sellthru_a': round(sz_st,4),
                'days_to_55_b': days_to_55,
                'spsd_c': round(sz_spsd,4),
                'is_oos': sz_inv == 0,
            })

        # ── 尺碼排序異常 ──
        sz_pur_map   = {s['sz']: s['cum_sales']+s['inv'] for s in sz_data}
        total_pur_sz = sum(sz_pur_map.values())
        sz_ho_ratio  = {sz: v/total_pur_sz for sz,v in sz_pur_map.items()} if total_pur_sz>0 else {}
        headorder_rank = sorted(sz_ho_ratio, key=lambda x:(-sz_ho_ratio[x], sz_key(x)))
        sellthru_rank  = sorted(sz_sellthru_map, key=lambda x:(-sz_sellthru_map[x], sz_key(x)))
        size_shift_warn = headorder_rank != sellthru_rank and len(headorder_rank)>1

        # ── 評分 ──
        if sellthru>=0.7:    st_score=5
        elif sellthru>=0.5:  st_score=4
        elif sellthru>=0.3:  st_score=3
        elif sellthru>=0.15: st_score=2
        else:                st_score=1
        if cum_sales>=1000:  vol_score=5
        elif cum_sales>=500: vol_score=4
        elif cum_sales>=100: vol_score=3
        elif cum_sales>=30:  vol_score=2
        else:                vol_score=1
        decision_score = st_score * vol_score

        # ── 周轉天數 ──
        n_stores = len(eligible_stores - OUTLET_STORES - IP_STORES) or 1
        turnover_days = round(color_main_inv / (effective_spsd * n_stores), 1) if effective_spsd>0 else None

        # ── 尺碼完整率 ──
        complete_stores = 0
        if all_sizes_sorted and not c_inv.empty:
            for sid in (eligible_stores - OUTLET_STORES - IP_STORES):
                s_inv = c_inv[c_inv['倉庫編號']==sid]
                if all(s_inv[s_inv['_size']==sz]['庫存數量'].sum()>0 for sz in all_sizes_sorted):
                    complete_stores += 1
        main_stores = len(eligible_stores - OUTLET_STORES - IP_STORES) or 1
        size_complete_rate = complete_stores / main_stores

        # ── 已追加量 / 在途量 ──
        already_ordered = total_ordered_map.get((pid,color), 0)
        transit_total   = sum(transit_map.get((pid,color,sz['sz']),0) for sz in sz_data)
        transit_by_sz   = {sz['sz']: transit_map.get((pid,color,sz['sz']),0) for sz in sz_data}

        # ── 距下檔 ──
        days_left = (end_dt - TODAY).days if end_dt else 999

        products_data.append({
            'pid': pid, 'style': style, 'name': name,
            'brand': brand, 'cat': cat, 'year': year, 'season': season,
            'color': color, 'img_url': img_url, 'barcode': barcode,
            'launch_date': str(launch),
            'cum_sales': cum_sales,
            'sellthru': round(sellthru,4),
            'effective_spsd': round(effective_spsd,4),
            'decision_score': decision_score,
            'st_score': st_score, 'vol_score': vol_score,
            'supply_warn': supply_warn,
            'size_shift_warn': size_shift_warn,
            'size_complete_rate': round(size_complete_rate,4),
            'color_main_inv': color_main_inv,
            'color_total_inv': color_total_inv,
            'color_store_inv': color_store_inv,
            'turnover_days': turnover_days,
            'days_left': days_left if days_left<999 else None,
            'weekly_sales': weekly_sales,
            'already_ordered': already_ordered,
            'transit_total': transit_total,
            'eligible_store_count': len(eligible_stores),
            'main_store_count': main_stores,
            'sz_data': sz_data,
            'sz_sellthru_map': sz_sellthru_map,
            'headorder_rank': headorder_rank,
            'sellthru_rank': sellthru_rank,
            'transit_by_sz': transit_by_sz,
            'leadtime': get_leadtime(pid, color),
        })

print(f"  共 {len(products_data)} 筆款色")

# ════════════════════════════════════════════════════════
# 5. 追加量計算
# ════════════════════════════════════════════════════════

DEFAULT_T = {'大爆款':22,'爆款':16,'潛力款':12,'常青款':6}
DEFAULT_G = {'大爆款':30,'爆款':21,'潛力款':14,'常青款':30,'地雷':0}

def get_tag(score, t=None):
    t = t or DEFAULT_T
    if score>=t['大爆款']:  return '大爆款'
    if score>=t['爆款']:    return '爆款'
    if score>=t['潛力款']:  return '潛力款'
    if score>=t['常青款']:  return '常青款'
    return '地雷'

def calc_reorder(p, t=None, g=None):
    tag = get_tag(p['decision_score'], t)
    gd  = (g or DEFAULT_G).get(tag, 0)
    lt  = p['leadtime']
    dl  = p['days_left'] if p['days_left'] is not None else 999
    eff_guard = max(0, min(gd, dl - lt))
    near_end  = (eff_guard==0 and dl<999 and tag!='地雷')
    if eff_guard<=0:
        return {'tag':tag,'reorder':0,'eff_guard':0,'lt':lt,'near_end':near_end}
    n = p['main_store_count']
    need = p['effective_spsd'] * n * (lt + eff_guard)
    reorder = max(0, round(need - p['color_main_inv'] - p['transit_total']))
    return {'tag':tag,'reorder':reorder,'eff_guard':eff_guard,'lt':lt,'near_end':near_end}

def calc_sz_reorder(p, total_reorder):
    total_st = sum(p['sz_sellthru_map'].values())
    result = []
    for sd in p['sz_data']:
        ratio = p['sz_sellthru_map'].get(sd['sz'],0)/total_st if total_st>0 else 1/max(1,len(p['sz_data']))
        result.append({**sd,
            'reorder': max(0,round(total_reorder*ratio)),
            'transit': p['transit_by_sz'].get(sd['sz'],0)})
    return result

for p in products_data:
    cr = calc_reorder(p)
    p['tag']       = cr['tag']
    p['reorder']   = cr['reorder']
    p['eff_guard'] = cr['eff_guard']
    p['near_end']  = cr['near_end']
    p['sz_detail'] = calc_sz_reorder(p, cr['reorder'])

print(f"  建議追加：{sum(1 for p in products_data if p['reorder']>0)} 款色")

# ════════════════════════════════════════════════════════
# 6. Excel 匯出
# ════════════════════════════════════════════════════════
print("📊 產出 Excel...")

excel_rows = []
for p in products_data:
    base = {
        '圖片URL':p['img_url'],'品牌':p['brand'],'年度':p['year'],'季節':p['season'],
        '大分類':p['cat'],'商品編號':p['pid'],'款號':p['style'],'商品名稱':p['name'],
        '顏色':p['color'],'標籤':p['tag'],'決策分':p['decision_score'],
        '完銷率':f"{p['sellthru']*100:.1f}%",
        '有效日均銷':round(p['effective_spsd'],2),
        '庫存周轉天數':p['turnover_days'],
        '近4週W1銷量':p['weekly_sales'][0]['qty'] if len(p['weekly_sales'])>0 else 0,
        '近4週W2銷量':p['weekly_sales'][1]['qty'] if len(p['weekly_sales'])>1 else 0,
        '近4週W3銷量':p['weekly_sales'][2]['qty'] if len(p['weekly_sales'])>2 else 0,
        '近4週W4銷量':p['weekly_sales'][3]['qty'] if len(p['weekly_sales'])>3 else 0,
        '尺碼完整率':f"{p['size_complete_rate']*100:.1f}%",
        '距下檔天數':p['days_left'] if p['days_left'] is not None else '',
        '已追加量':p['already_ordered'],'在途量':p['transit_total'],
        '建議追加量':p['reorder'],
        '供給受限':'是' if p['supply_warn'] else '',
        '尺碼結構偏移':'是' if p['size_shift_warn'] else '',
        '追加量已壓縮':'是' if p['near_end'] else '',
    }
    for sd in p['sz_detail']:
        excel_rows.append({**base,
            '商品條碼':p['barcode'],'尺碼':sd['sz'],'現庫存':sd['inv'],
            '(A)完銷率':f"{sd['sellthru_a']*100:.1f}%",
            '(B)達55%天數':sd['days_to_55_b'] or '',
            '(C)有效日均銷':round(sd['spsd_c'],2),
            '尺碼建議追加':sd['reorder'],'尺碼在途量':sd['transit'],
            '斷碼':'是' if sd['is_oos'] else ''})

df_excel = pd.DataFrame(excel_rows) if excel_rows else pd.DataFrame()

# 總覽
summary_rows = []
if not df_excel.empty:
    for keys, grp in df_excel.groupby(['年度','季節','品牌','大分類']):
        ck = grp.groupby(['商品編號','顏色']).first()
        ro_skus = ck[ck['建議追加量'].apply(lambda x: safe_int(x)>0)]
        summary_rows.append({
            '年度':keys[0],'季節':keys[1],'品牌':keys[2],'大分類':keys[3],
            '建議追加色數':len(ro_skus),
            '建議追加量':grp['尺碼建議追加'].apply(safe_int).sum()
        })
df_summary = pd.DataFrame(summary_rows)

excel_out = os.path.join(BASE_DIR, 'output.xlsx')
with pd.ExcelWriter(excel_out, engine='xlsxwriter') as writer:
    if not df_excel.empty:
        df_excel.to_excel(writer, sheet_name='追加明細', index=False)
    if not df_summary.empty:
        df_summary.to_excel(writer, sheet_name='本週總表', index=False)
    wb = writer.book
    hdr_fmt = wb.add_format({'bold':True,'bg_color':'#F5F5F0','border':1,'text_wrap':True})
    for sn, df_s in [('追加明細',df_excel),('本週總表',df_summary)]:
        if df_s.empty: continue
        ws = writer.sheets[sn]
        for ci,col in enumerate(df_s.columns):
            ws.write(0,ci,col,hdr_fmt)
            ws.set_column(ci,ci,max(10,len(str(col))+2))
        ws.freeze_panes(1,0)
print("  ✅ output.xlsx")

# ════════════════════════════════════════════════════════
# 7. JSON → HTML
# ════════════════════════════════════════════════════════

def to_json_obj(p):
    return {
        'pid':p['pid'],'style':p['style'],'name':p['name'],
        'brand':p['brand'],'cat':p['cat'],'year':p['year'],'season':p['season'],
        'color':p['color'],'img_url':p['img_url'],
        'cum_sales':p['cum_sales'],'sellthru':p['sellthru'],
        'effective_spsd':p['effective_spsd'],'decision_score':p['decision_score'],
        'st_score':p['st_score'],'vol_score':p['vol_score'],
        'supply_warn':p['supply_warn'],'size_shift_warn':p['size_shift_warn'],
        'size_complete_rate':p['size_complete_rate'],
        'color_main_inv':p['color_main_inv'],
        'turnover_days':p['turnover_days'],
        'days_left':p['days_left'],
        'weekly_sales':p['weekly_sales'],
        'already_ordered':p['already_ordered'],
        'transit_total':p['transit_total'],
        'eligible_store_count':p['eligible_store_count'],
        'main_store_count':p['main_store_count'],
        'tag':p['tag'],'reorder':p['reorder'],
        'eff_guard':p['eff_guard'],'leadtime':p['leadtime'],
        'near_end':p['near_end'],
        'sz_detail':p['sz_detail'],
        'headorder_rank':p['headorder_rank'],
        'sz_sellthru_map':p['sz_sellthru_map'],
    }

snap_fmt = f"{inv_snap_date[:4]}-{inv_snap_date[4:6]}-{inv_snap_date[6:]}"
json_data = {
    'generated_at': str(TODAY),
    'inv_snap_date': snap_fmt,
    'products': [to_json_obj(p) for p in products_data],
    'defaults': {'thresholds': DEFAULT_T, 'guard_days': DEFAULT_G},
    'week_labels': [ws.strftime('%m/%d') for ws in week_starts],
}
json_str = json.dumps(json_data, ensure_ascii=False)

# ════════════════════════════════════════════════════════
# 8. 產出 index.html
# ════════════════════════════════════════════════════════
print("🌐 產出 index.html...")

html = f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>追加下單建議 {TODAY}</title>
<style>
*{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI','Microsoft JhengHei',sans-serif;
     font-size:13px;color:#1a1a1a;background:#ffffff;line-height:1.5}}
.app{{max-width:1600px;margin:0 auto;padding:16px}}
.header{{display:flex;justify-content:space-between;align-items:flex-end;
         margin-bottom:14px;border-bottom:1.5px solid #e8e6e0;padding-bottom:10px}}
.header h1{{font-size:17px;font-weight:600}}
.header .meta{{font-size:11px;color:#999}}
.cards{{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:8px;margin-bottom:14px}}
.card{{background:#fafaf8;border:0.5px solid #e8e6e0;border-radius:8px;padding:10px 12px}}
.card .lbl{{font-size:10px;color:#999;margin-bottom:3px}}
.card .val{{font-size:20px;font-weight:600}}
.card .sub{{font-size:10px;color:#bbb;margin-top:1px}}
.params{{background:#fafaf8;border:0.5px solid #e8e6e0;border-radius:8px;
         padding:12px;margin-bottom:12px}}
.params-title{{font-size:11px;font-weight:500;color:#777;margin-bottom:8px}}
.params-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(160px,1fr));gap:8px}}
.param-row{{display:flex;flex-direction:column;gap:3px}}
.param-row label{{font-size:10px;color:#999}}
.param-inner{{display:flex;align-items:center;gap:5px}}
.param-inner input[type=range]{{flex:1;accent-color:#333}}
.param-val{{font-size:12px;font-weight:500;min-width:32px}}
.toolbar{{display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px;align-items:center}}
.toolbar input,.toolbar select{{height:30px;border:0.5px solid #ddd;border-radius:6px;
  padding:0 8px;font-size:12px;background:#fff;color:#222;outline:none}}
.toolbar input:focus,.toolbar select:focus{{border-color:#888}}
.toolbar input{{width:130px}}
.toolbar select{{min-width:110px}}
.def-notes{{display:flex;flex-direction:column;gap:6px;margin-bottom:10px}}
.def-note{{background:#fafaf8;border-left:3px solid #d0cec8;border-radius:0 6px 6px 0;
           padding:7px 12px;font-size:11px;color:#555;line-height:1.6}}
.def-note b{{color:#222;font-weight:500}}
.def-note .dn-tag{{display:inline-block;padding:1px 7px;border-radius:8px;
                   font-size:10px;font-weight:500;margin-right:6px;vertical-align:middle}}
.dn-supply{{background:#FEE2E2;color:#991B1B}}
.dn-shift{{background:#FFF3E0;color:#C2410C}}
.btn{{height:30px;padding:0 14px;border:0.5px solid #ddd;border-radius:6px;
      background:#fff;font-size:12px;cursor:pointer;color:#333}}
.btn:hover{{background:#f5f5f2}}
.tabs{{display:flex;border-bottom:1px solid #e8e6e0;margin-bottom:12px}}
.tab{{padding:6px 16px;font-size:12px;cursor:pointer;color:#999;
      border-bottom:2px solid transparent;margin-bottom:-1px}}
.tab.active{{color:#111;border-bottom-color:#111;font-weight:500}}
.table-wrap{{overflow-x:auto;overflow-y:visible}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
thead tr{{border-bottom:1.5px solid #ddd}}
th{{text-align:left;padding:6px 7px;font-weight:500;font-size:10px;
    color:#999;white-space:nowrap;cursor:pointer;user-select:none}}
th:hover{{color:#555}}
td{{padding:7px 7px;border-bottom:0.5px solid #f0ede8;vertical-align:middle}}
tr.main-row:hover>td{{background:#fafaf8}}
tr.sub-row>td{{background:#fafaf8;padding:0}}
.badge{{display:inline-flex;align-items:center;font-size:10px;padding:1px 7px;
        border-radius:9px;font-weight:500;white-space:nowrap}}
.tag-mega{{background:#EEEDFE;color:#3C3489}}
.tag-hot{{background:#FAECE7;color:#993C1D}}
.tag-rising{{background:#E1F5EE;color:#085041}}
.tag-steady{{background:#EAF3DE;color:#3B6D11}}
.tag-slow{{background:#F1EFE8;color:#5F5E5A}}
.badge-brand{{background:#F5F5F2;color:#555;border:0.5px solid #e0ded8}}
.warns{{display:flex;gap:4px;flex-wrap:wrap;margin-top:3px}}
.wchip{{font-size:10px;padding:1px 6px;border-radius:6px;font-weight:400}}
.wchip.supply{{background:#FEE2E2;color:#991B1B}}
.wchip.shift{{background:#FFF3E0;color:#C2410C}}
.wchip.end{{background:#F3E8FF;color:#6B21A8}}
.score-chip{{display:inline-block;width:26px;height:26px;border-radius:50%;
             background:#f5f5f2;border:0.5px solid #ddd;text-align:center;
             line-height:26px;font-weight:600;font-size:11px}}
.tip{{position:relative;display:inline-flex;align-items:center}}
.tip-icon{{font-size:11px;color:#ccc;cursor:help;margin-left:2px}}

.prod-img{{width:40px;height:40px;object-fit:cover;border-radius:4px;
           border:0.5px solid #e8e6e0}}
.prod-img-ph{{width:40px;height:40px;border-radius:4px;border:0.5px solid #e8e6e0;
              background:#f5f5f2;display:flex;align-items:center;justify-content:center;
              color:#ccc;font-size:18px}}
.exp-btn{{background:none;border:none;cursor:pointer;color:#bbb;font-size:14px;padding:0 3px}}
.exp-btn:hover{{color:#555}}
.sub-inner{{padding:10px 10px 10px 50px}}
.sub-meta{{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:8px;font-size:11px;color:#666}}
.sub-meta b{{color:#333}}

.sz-table{{width:100%;border-collapse:collapse;font-size:11px}}
.sz-table th{{padding:4px 7px;font-weight:500;color:#999;text-align:left;
              border-bottom:0.5px solid #eee;font-size:10px}}
.sz-table td{{padding:4px 7px;border-bottom:0.5px solid #f5f3ef}}
.oos{{background:#FEE2E2;color:#991B1B;font-size:9px;padding:1px 5px;
      border-radius:5px;margin-left:4px}}
.to-urgent{{color:#DC2626;font-weight:600}}
.to-healthy{{color:#16A34A}}
.to-heavy{{color:#D97706}}
.to-slow{{color:#9CA3AF}}
.no-data{{text-align:center;color:#bbb;padding:36px;font-size:13px}}
</style>
</head>
<body>
<div class="app">

<div class="header">
  <div>
    <h1>追加下單建議</h1>
    <div class="meta">庫存快照：{snap_fmt}　產出：{TODAY}</div>
  </div>
  <button class="btn" onclick="window.location.href='output.xlsx'">⬇ 匯出 Excel</button>
</div>

<div class="cards">
  <div class="card"><div class="lbl">建議追加色數</div><div class="val" id="c1">—</div><div class="sub">款色需追加</div></div>
  <div class="card"><div class="lbl">建議追加總量</div><div class="val" id="c2">—</div><div class="sub">件</div></div>
  <div class="card"><div class="lbl">供給受限</div><div class="val" id="c3">—</div><div class="sub">趨勢數據僅供參考</div></div>
  <div class="card"><div class="lbl">尺碼結構偏移</div><div class="val" id="c4">—</div><div class="sub">建議人工確認</div></div>
</div>

<div class="params">
  <div class="params-title">評分門檻與保障天數（即時重算）</div>
  <div class="params-grid">
    <div class="param-row"><label>大爆款門檻 <span class="tip-icon" onclick="showTip(this,'決策分≥此值→大爆款，保障天數最長，全力防斷貨')">ⓘ</span></label><div class="param-inner"><input type="range" min="15" max="25" value="22" id="t1" oninput="sl(this,'tv1','分')"><span class="param-val" id="tv1">22分</span></div></div>
    <div class="param-row"><label>爆款門檻</label><div class="param-inner"><input type="range" min="10" max="21" value="16" id="t2" oninput="sl(this,'tv2','分')"><span class="param-val" id="tv2">16分</span></div></div>
    <div class="param-row"><label>潛力款門檻</label><div class="param-inner"><input type="range" min="6" max="15" value="12" id="t3" oninput="sl(this,'tv3','分')"><span class="param-val" id="tv3">12分</span></div></div>
    <div class="param-row"><label>常青款門檻</label><div class="param-inner"><input type="range" min="3" max="11" value="6" id="t4" oninput="sl(this,'tv4','分')"><span class="param-val" id="tv4">6分</span></div></div>
    <div class="param-row"><label>大爆款保障天數</label><div class="param-inner"><input type="range" min="14" max="60" value="30" id="d1" oninput="sl(this,'dv1','天')"><span class="param-val" id="dv1">30天</span></div></div>
    <div class="param-row"><label>爆款保障天數</label><div class="param-inner"><input type="range" min="7" max="45" value="21" id="d2" oninput="sl(this,'dv2','天')"><span class="param-val" id="dv2">21天</span></div></div>
    <div class="param-row"><label>潛力款保障天數</label><div class="param-inner"><input type="range" min="7" max="30" value="14" id="d3" oninput="sl(this,'dv3','天')"><span class="param-val" id="dv3">14天</span></div></div>
    <div class="param-row"><label>常青款保障天數</label><div class="param-inner"><input type="range" min="7" max="45" value="30" id="d4" oninput="sl(this,'dv4','天')"><span class="param-val" id="dv4">30天</span></div></div>
  </div>
</div>

<div class="tabs">
  <div class="tab active" id="tab-main" onclick="switchTab('main')">追加明細</div>
  <div class="tab" id="tab-sum" onclick="switchTab('sum')">本週總表</div>
</div>

<div id="sec-toolbar">
<div class="def-notes">
  <div class="def-note">
    <span class="dn-tag dn-supply">⚡ 供給受限</span>
    <b>近7天有貨可販售店數 &lt; 上架初期有貨店數的70%。</b>代表商品銷售下滑可能是因為沒貨可賣，而非市場需求下降，此時趨勢數據不可作為追加依據。
  </div>
  <div class="def-note">
    <span class="dn-tag dn-shift">↕ 尺碼結構偏移</span>
    <b>各尺碼完銷率與頭單進貨佔比排序不一致。</b>代表版型或市場偏好與預期有落差，建議人工確認追加尺碼比例。
  </div>
</div>
<div class="toolbar">
  <input type="text" id="fi" placeholder="商品編號" oninput="render()">
  <input type="text" id="fs" placeholder="款號" oninput="render()">
  <input type="text" id="fn" placeholder="商品名稱" oninput="render()">
  <select id="fb" onchange="render()"><option value="">全部品牌</option></select>
  <select id="fc" onchange="render()"><option value="">全部大分類</option></select>
  <select id="ft" onchange="render()"><option value="">全部標籤</option>
    <option>大爆款</option><option>爆款</option><option>潛力款</option><option>常青款</option><option>地雷</option>
  </select>
  <select id="fw" onchange="render()"><option value="">全部警示</option>
    <option value="supply">供給受限</option>
    <option value="shift">尺碼結構偏移</option>
    <option value="end">追加量已壓縮</option>
  </select>
</div>
</div>

<div id="sec-main">
<div class="table-wrap">
<table>
<thead><tr>
  <th></th><th>照片</th><th>品牌</th><th>大分類</th>
  <th>商品編號</th><th>款號</th><th>商品名稱</th><th>標籤</th>
  <th style="text-align:right">決策分<span class="tip-icon" onclick="showTip(this,'完銷率評分(1-5) × 銷量評分(1-5)，最高25分。')">ⓘ</span></th>
  <th style="text-align:right">完銷率<span class="tip-icon" onclick="showTip(this,'累銷量 ÷ 總進貨量(累銷+現庫存回推)。')">ⓘ</span></th>
  <th style="text-align:right">有效日均銷<span class="tip-icon" onclick="showTip(this,'累銷量÷有貨store-days。只計算可販售店且當天有庫存的天數為分母，排除缺貨期低估。')">ⓘ</span></th>
  <th style="text-align:right">庫存周轉<span class="tip-icon" onclick="showTip(this,'正價店現有庫存÷有效日均銷÷店數。<10天緊急、10-30天健康、30-60天偏重、>60天滯銷')">ⓘ</span></th>
  <th style="text-align:right">距下檔<span class="tip-icon" onclick="showTip(this,'距最後販售日天數，接近換季自動壓縮追加量。')">ⓘ</span></th>
  <th style="text-align:right">建議追加<span class="tip-icon" onclick="showTip(this,'有效日均銷×可販售店數×(LeadTime+有效保障天數)-現有庫存-在途量。')">ⓘ</span></th>
  <th style="text-align:center;min-width:110px">近4週趨勢</th>
</tr></thead>
<tbody id="tbody"></tbody>
</table>
</div>
</div>

<div id="sec-sum" style="display:none">
<div class="table-wrap">
<table>
<thead><tr><th>年度</th><th>季節</th><th>品牌</th><th>大分類</th>
  <th style="text-align:right">建議追加色數</th><th style="text-align:right">建議追加量</th>
</tr></thead>
<tbody id="tbody-sum"></tbody>
</table>
</div>
</div>

</div>
<script>
const DATA={json_str};
const PRODS=DATA.products;
const WLABELS=DATA.week_labels;
const TAG_S={{'大爆款':{{c:'tag-mega',i:'👑'}},'爆款':{{c:'tag-hot',i:'✨'}},
              '潛力款':{{c:'tag-rising',i:'🌑'}},'常青款':{{c:'tag-steady',i:'🛡️'}},'地雷':{{c:'tag-slow',i:'⚠️'}}}};

// Init filters
(()=>{{
  const brands=[...new Set(PRODS.map(p=>p.brand))].sort();
  const cats=[...new Set(PRODS.map(p=>p.cat))].sort();
  ['fb','fc'].map((id,i)=>{{
    const sel=document.getElementById(id);
    (i===0?brands:cats).forEach(v=>{{const o=document.createElement('option');o.value=o.textContent=v;sel.appendChild(o);}});
  }});
}})();

function sl(el,oid,sfx){{document.getElementById(oid).textContent=el.value+sfx;render();}}

function params(){{
  return {{t1:+t1.value,t2:+t2.value,t3:+t3.value,t4:+t4.value,
          d1:+d1.value,d2:+d2.value,d3:+d3.value,d4:+d4.value}};
}}
function getTag(sc,p){{
  if(sc>=p.t1)return'大爆款';if(sc>=p.t2)return'爆款';
  if(sc>=p.t3)return'潛力款';if(sc>=p.t4)return'常青款';return'地雷';
}}
function calcRO(prod,p){{
  const tag=getTag(prod.decision_score,p);
  const gmap={{'大爆款':p.d1,'爆款':p.d2,'潛力款':p.d3,'常青款':p.d4,'地雷':0}};
  const gd=gmap[tag],lt=prod.leadtime||30,dl=prod.days_left??999;
  const eg=Math.max(0,Math.min(gd,dl-lt));
  const nearEnd=eg===0&&dl<999&&tag!=='地雷';
  if(eg<=0)return{{tag,ro:0,eg,lt,nearEnd}};
  const need=prod.effective_spsd*prod.main_store_count*(lt+eg);
  const ro=Math.max(0,Math.round(need-prod.color_main_inv-prod.transit_total));
  return{{tag,ro,eg,lt,nearEnd}};
}}
function calcSzRO(prod,total){{
  const stMap=prod.sz_sellthru_map;
  const total_st=Object.values(stMap).reduce((a,b)=>a+b,0);
  return prod.sz_detail.map(s=>{{
    const r=total_st>0?(stMap[s.sz]||0)/total_st:1/Math.max(1,prod.sz_detail.length);
    return{{...s,ro:Math.max(0,Math.round(total*r))}};
  }});
}}
function toClass(v){{
  if(v==null)return'';if(v<10)return'to-urgent';if(v<=30)return'to-healthy';
  if(v<=60)return'to-heavy';return'to-slow';
}}

const expanded=new Set();
function toggle(k){{expanded.has(k)?expanded.delete(k):expanded.add(k);render();}}

let curTab='main';
function switchTab(t){{
  curTab=t;
  document.getElementById('sec-main').style.display=t==='main'?'':'none';
  document.getElementById('sec-toolbar').style.display=t==='main'?'':'none';
  document.getElementById('sec-sum').style.display=t==='sum'?'':'none';
  document.getElementById('tab-main').className='tab'+(t==='main'?' active':'');
  document.getElementById('tab-sum').className='tab'+(t==='sum'?' active':'');
  if(t==='sum')renderSum();
}}

function render(){{
  const p=params();
  const qi=fi.value.toLowerCase(),qs=fs.value.toLowerCase(),qn=fn.value.toLowerCase();
  const fb_v=fb.value,fc_v=fc.value,ft_v=ft.value,fw_v=fw.value;
  let skus=0,qty=0,sw=0,shw=0,html='';

  PRODS.forEach(prod=>{{
    if(qi&&!prod.pid.toLowerCase().includes(qi))return;
    if(qs&&!prod.style.toLowerCase().includes(qs))return;
    if(qn&&!prod.name.toLowerCase().includes(qn))return;
    if(fb_v&&prod.brand!==fb_v)return;
    if(fc_v&&prod.cat!==fc_v)return;
    const cr=calcRO(prod,p);
    if(ft_v&&cr.tag!==ft_v)return;
    if(fw_v==='supply'&&!prod.supply_warn)return;
    if(fw_v==='shift'&&!prod.size_shift_warn)return;
    if(fw_v==='end'&&!cr.nearEnd)return;

    if(cr.ro>0){{skus++;qty+=cr.ro;}}
    if(prod.supply_warn)sw++;
    if(prod.size_shift_warn)shw++;

    const ts=TAG_S[cr.tag]||TAG_S['地雷'];
    const k=prod.pid+'_'+prod.color;
    const isExp=expanded.has(k);
    const dl=prod.days_left;
    const dlStr=dl!=null?(dl<14?`<span style="color:#DC2626;font-weight:600">${{dl}}天</span>`:dl+'天'):'—';
    const roStr=cr.ro>0?`<b>${{cr.ro}}</b>件`:'<span style="color:#ccc">—</span>';
    const imgHtml=prod.img_url
      ?`<img class="prod-img" src="${{prod.img_url}}" alt="" onerror="this.style.display='none'">`
      :`<div class="prod-img-ph">📷</div>`;
    let warns='';
    if(prod.supply_warn)warns+='<span class="wchip supply">⚡ 供給受限</span>';
    if(prod.size_shift_warn)warns+='<span class="wchip shift">↕ 尺碼偏移</span>';
    if(cr.nearEnd)warns+='<span class="wchip end">⏳ 追加壓縮</span>';
    if(warns)warns='<div class="warns">'+warns+'</div>';

    html+=`<tr class="main-row">
      <td><button class="exp-btn" onclick="toggle('${{k}}')">${{isExp?'▾':'▸'}}</button></td>
      <td>${{imgHtml}}</td>
      <td><span class="badge badge-brand">${{prod.brand}}</span></td>
      <td style="color:#777">${{prod.cat}}</td>
      <td style="font-size:11px;color:#999">${{prod.pid}}<br><span style="color:#666">${{prod.color}}</span></td>
      <td style="font-size:11px;color:#777">${{prod.style}}</td>
      <td><div style="font-size:12px">${{prod.name}}</div>${{warns}}</td>
      <td><span class="badge ${{ts.c}}">${{ts.i}} ${{cr.tag}}</span></td>
      <td style="text-align:right"><span class="score-chip">${{prod.decision_score}}</span></td>
      <td style="text-align:right">${{(prod.sellthru*100).toFixed(1)}}%</td>
      <td style="text-align:right">${{prod.effective_spsd.toFixed(2)}}</td>
      <td style="text-align:right"><span class="${{toClass(prod.turnover_days)}}">${{prod.turnover_days!=null?prod.turnover_days+'天':'—'}}</span></td>
      <td style="text-align:right">${{dlStr}}</td>
      <td style="text-align:right">${{roStr}}</td>
      <td style="text-align:center;padding:3px 6px">${{sparkline(prod.weekly_sales)}}</td>
    </tr>`;



    if(isExp){{
      const szs=calcSzRO(prod,cr.ro);
      html+=`<tr class="sub-row"><td colspan="15"><div class="sub-inner">
        <div class="sub-meta">
          <span>年度：<b>${{prod.year}}</b></span>
          <span>季節：<b>${{prod.season}}</b></span>
          <span>尺碼完整率：<b style="color:${{prod.size_complete_rate<0.8?'#D97706':'inherit'}}">${{(prod.size_complete_rate*100).toFixed(0)}}%</b>
            <span class="tip-icon" onclick="showTip(this,'各主力尺碼都有庫存的店數÷可販售總店數。')">ⓘ</span>
          </span>
          <span>LeadTime：<b>${{cr.lt}}</b>天</span>
          <span>可販售店：<b>${{prod.eligible_store_count}}</b>間</span>
          ${{prod.size_shift_warn?`<span style="color:#C2410C">⚠ 頭單：${{prod.headorder_rank.join('>')}}</span>`:''}}
        </div>
        <table class="sz-table">
          <thead><tr>
            <th>尺碼</th>
            <th style="text-align:right">現庫存</th>
            <th style="text-align:right">(A) 完銷率 <span class="tip-icon" onclick="showTip(this,'主力尺碼依據，追加量分配以此為基準。')">ⓘ</span></th>
            <th style="text-align:right">(B) 達55%天數 <span class="tip-icon" onclick="showTip(this,'從上架日起賣到55%完銷所需天數。')">ⓘ</span></th>
            <th style="text-align:right">(C) 有效日均銷 <span class="tip-icon" onclick="showTip(this,'該尺碼在有貨期間每日均銷量，排除缺貨期低估。')">ⓘ</span></th>
            <th style="text-align:right">已追加量 <span class="tip-icon" onclick="showTip(this,'此尺碼在追加單中尚未到貨的在途量。')">ⓘ</span></th>
            <th style="text-align:right">建議追加</th>
          </tr></thead>
          <tbody>
            ${{szs.map(s=>`<tr>
              <td>${{s.sz}}${{s.is_oos?'<span class="oos">斷碼</span>':''}}</td>
              <td style="text-align:right">${{s.inv}}</td>
              <td style="text-align:right">${{(s.sellthru_a*100).toFixed(1)}}%</td>
              <td style="text-align:right">${{s.days_to_55_b!=null?s.days_to_55_b+'天':'—'}}</td>
              <td style="text-align:right">${{s.spsd_c.toFixed(2)}}</td>
              <td style="text-align:right;color:#888">${{s.transit>0?s.transit+'件':'—'}}</td>
              <td style="text-align:right;font-weight:500">${{s.ro>0?s.ro+'件':'—'}}</td>
            </tr>`).join('')}}
          </tbody>
        </table>
      </div></td></tr>`;
    }}
  }});

  document.getElementById('c1').textContent=skus;
  document.getElementById('c2').textContent=qty.toLocaleString();
  document.getElementById('c3').textContent=sw;
  document.getElementById('c4').textContent=shw;
  document.getElementById('tbody').innerHTML=html||'<tr><td colspan="15" class="no-data">無符合條件的商品</td></tr>';
}}

function renderSum(){{
  const p=params();
  const map={{}};
  PRODS.forEach(prod=>{{
    const cr=calcRO(prod,p);
    const k=`${{prod.year}}|${{prod.season}}|${{prod.brand}}|${{prod.cat}}`;
    if(!map[k])map[k]={{year:prod.year,season:prod.season,brand:prod.brand,cat:prod.cat,skus:0,qty:0}};
    if(cr.ro>0)map[k].skus++;
    map[k].qty+=cr.ro;
  }});
  document.getElementById('tbody-sum').innerHTML=
    Object.values(map).sort((a,b)=>a.year.localeCompare(b.year)||a.brand.localeCompare(b.brand)||a.cat.localeCompare(b.cat))
    .map(r=>`<tr><td>${{r.year}}</td><td>${{r.season}}</td><td>${{r.brand}}</td><td>${{r.cat}}</td>
      <td style="text-align:right">${{r.skus}}</td><td style="text-align:right">${{r.qty.toLocaleString()}} 件</td></tr>`)
    .join('')||'<tr><td colspan="6" class="no-data">—</td></tr>';
}}


function sparkline(weekly) {{
  if (!weekly || weekly.length === 0) return '';
  const vals = weekly.map(function(w) {{ return w.qty; }});
  const maxV = Math.max.apply(null, vals.concat([1]));
  const W = 110, H = 52, padX = 6, padTop = 14, padBot = 12, n = vals.length;
  const pts = vals.map(function(v, i) {{
    const x = padX + (n > 1 ? i / (n - 1) : 0.5) * (W - padX * 2);
    const y = padTop + (1 - v / maxV) * (H - padTop - padBot);
    return [x.toFixed(1), y.toFixed(1)];
  }});
  const last = vals[vals.length - 1];
  const prev = vals.length > 1 ? vals[vals.length - 2] : last;
  const color = last > prev ? '#16A34A' : last < prev ? '#DC2626' : '#9CA3AF';
  let svg = '<svg width="' + W + '" height="' + H + '" viewBox="0 0 ' + W + ' ' + H + '" xmlns="http://www.w3.org/2000/svg" style="display:block;margin:auto;overflow:visible">';
  svg += '<polyline points="' + pts.map(function(p) {{ return p[0]+','+p[1]; }}).join(' ') + '" fill="none" stroke="' + color + '" stroke-width="1.6" stroke-linejoin="round" stroke-linecap="round"/>';
  pts.forEach(function(pt, i) {{
    const x = parseFloat(pt[0]), y = parseFloat(pt[1]);
    const isLast = i === pts.length - 1;
    svg += '<circle cx="' + x + '" cy="' + y + '" r="' + (isLast ? 2.8 : 1.8) + '" fill="' + (isLast ? color : '#bbb') + '"/>';
    svg += '<text x="' + x + '" y="' + (y - 4) + '" text-anchor="middle" font-size="9" fill="' + (isLast ? color : '#aaa') + '" font-family="-apple-system,sans-serif">' + vals[i] + '</text>';
  }});
  svg += '</svg>';
  return svg;
}}


// 共用 tooltip，掛在 body，position:fixed 不受任何容器裁切
var _tip = document.createElement('div');
_tip.id = '_gtip';
_tip.style.cssText = 'display:none;position:fixed;background:#222;color:#fff;border-radius:6px;padding:7px 11px;font-size:11px;width:220px;line-height:1.6;z-index:99999;word-break:keep-all;box-shadow:0 2px 8px rgba(0,0,0,.25);pointer-events:none';
document.body.appendChild(_tip);

var _tipSrc = null;
function showTip(el, text) {{
  if (_tipSrc === el && _tip.style.display !== 'none') {{
    _tip.style.display = 'none'; _tipSrc = null; return;
  }}
  _tipSrc = el;
  _tip.textContent = text;
  _tip.style.display = 'block';
  _tip.style.visibility = 'hidden';
  var rect = el.getBoundingClientRect();
  var tw = _tip.offsetWidth, th = _tip.offsetHeight;
  var top = rect.top - th - 8;
  var left = rect.left + rect.width/2 - tw/2;
  if (top < 6) top = rect.bottom + 8;
  if (left < 6) left = 6;
  if (left + tw > window.innerWidth - 6) left = window.innerWidth - tw - 6;
  _tip.style.top = top + 'px';
  _tip.style.left = left + 'px';
  _tip.style.visibility = 'visible';
}}
document.addEventListener('click', function(e) {{
  if (!e.target.classList.contains('tip-icon')) {{
    _tip.style.display = 'none'; _tipSrc = null;
  }}
}});

render();
</script>
</body>
</html>"""

with open(os.path.join(BASE_DIR,'index.html'),'w',encoding='utf-8') as f:
    f.write(html)

print(f"✅ 完成！index.html 產出")
print(f"   款色：{len(products_data)}　建議追加：{sum(1 for p in products_data if p['reorder']>0)}")
