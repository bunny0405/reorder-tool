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
    s = str(v).strip()
    try:
        t = pd.to_datetime(s, errors='coerce')
        if t is not pd.NaT and t == t:
            return t.date()
    except: pass
    for fmt in ('%Y/%m/%d','%Y-%m-%d','%Y%m%d','%Y-%m-%d %H:%M:%S','%Y/%m/%d %H:%M:%S'):
        try: return datetime.strptime(s[:10], fmt[:8] if len(fmt)>8 else fmt).date()
        except: pass
    try: return datetime.strptime(s[:10], '%Y-%m-%d').date()
    except: pass
    return None

def natural_week_start(d):
    return d - timedelta(days=d.weekday())# ════════════════════════════════════════════════════════
# 2. 載入資料
# ════════════════════════════════════════════════════════
print("📂 載入資料...")

# ── 門市基本資料 ──
_store_path = os.path.join(DATA_DIR, '門市基本資料.xlsx')
try:
    df_store = pd.read_excel(_store_path, sheet_name='店數', dtype=str)
except:
    df_store = read_file(_store_path)

df_store.columns = [str(c).strip() for c in df_store.columns]
df_store = df_store.dropna(axis=1, how='all')
df_store.columns = [str(c).strip() for c in df_store.columns]
df_store = df_store[df_store['倉庫編號'].notna() & (df_store['倉庫編號'].str.strip() != '')]
df_store['倉庫編號'] = df_store['倉庫編號'].astype(str).str.strip()
if '販售支線' in df_store.columns:
    df_store['販售支線'] = df_store['販售支線'].fillna('')
else:
    df_store['販售支線'] = ''

OUTLET_STORES = {'AS51','AS63','AT08'}
IP_STORES     = {'AS62','AT10'}

store_line_map = dict(zip(df_store['倉庫編號'], df_store['販售支線']))
store_type_map = dict(zip(df_store['倉庫編號'], df_store.get('門市型態', pd.Series()).fillna(''))) if '門市型態' in df_store.columns else {}
all_store_ids  = set(store_line_map.keys()) - {'nan','None',''}
print(f"  門市數：{len(all_store_ids)}")

# ── 商品主檔 ──
df_prod = read_file(os.path.join(DATA_DIR, '新品基本資料.xlsx'))
df_prod.columns = [str(c).strip() for c in df_prod.columns]
col_brand_a = df_prod.columns[0]
col_ip_g    = df_prod.columns[6] if len(df_prod.columns) > 6 else None
col_end_date = '最後販售日' if '最後販售日' in df_prod.columns else None
col_reorder_weeks = '續賣週數' if '續賣週數' in df_prod.columns else None

df_prod['商品編號'] = df_prod['商品編號'].astype(str).str.strip()
df_prod['款號']     = df_prod['款號'].astype(str).str.strip()

_launch_col = '上架日' if ('上架日' in df_prod.columns and df_prod['上架日'].notna().sum() > 0) else '最新上架日期'
df_prod['launch_date'] = df_prod[_launch_col].apply(parse_date)

if col_end_date:
    df_prod['end_date'] = df_prod[col_end_date].apply(parse_date)
elif col_reorder_weeks:
    def calc_end(row):
        launch = parse_date(row.get(_launch_col,''))
        weeks = 0
        try: weeks = int(float(str(row.get('續賣週數',0))))
        except: pass
        if launch and weeks > 0:
            return launch + timedelta(weeks=weeks)
        return None
    df_prod['end_date'] = df_prod.apply(calc_end, axis=1)
else:
    df_prod['end_date'] = None

df_prod['is_ot']  = df_prod[col_brand_a].astype(str).str.strip().str.upper() == 'OT'
df_prod['is_ip']  = df_prod[col_ip_g].astype(str).str.strip().str.upper() == 'IP' if col_ip_g else False
df_prod['brand']  = df_prod[col_brand_a].astype(str).str.strip()
if '品牌' in df_prod.columns:
    df_prod['brand_display'] = df_prod['品牌'].fillna('').astype(str).str.strip()
else:
    df_prod['brand_display'] = df_prod['brand']

img_col = '商品內部圖片' if '商品內部圖片' in df_prod.columns else None
if not img_col:
    for c in df_prod.columns:
        if '圖' in str(c) or 'img' in str(c).lower():
            img_col = c; break
df_prod['img_url'] = df_prod[img_col].fillna('') if img_col else ''

barcode_col = '商品條碼' if '商品條碼' in df_prod.columns else None
if not barcode_col:
    for c in df_prod.columns:
        if '條碼' in str(c):
            barcode_col = c; break
df_prod['barcode'] = df_prod[barcode_col].fillna('') if barcode_col else ''

if '年度' not in df_prod.columns: df_prod['年度'] = ''
if '季節' not in df_prod.columns: df_prod['季節'] = ''
if '商品成分' not in df_prod.columns: df_prod['商品成分'] = ''
if '厚度' not in df_prod.columns: df_prod['厚度'] = ''
if '商品中分類' not in df_prod.columns: df_prod['商品中分類'] = ''

prod_map = df_prod.drop_duplicates(subset='商品編號', keep='last').set_index('商品編號').to_dict('index')
print(f"  商品數：{len(prod_map)}")# ── 做貨備註 ──
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
transit_map = {}
total_ordered_map = {}
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
        print(f"    讀入銷售：{os.path.basename(f)} ({len(df_s)} 列)")
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
inv_path = latest_file('周一庫存_*.xlsx') or latest_file('周一庫存_*.csv') or latest_file('周一庫存*.xlsx')
if not inv_path:
    inv_files = glob.glob(os.path.join(DATA_DIR, '*庫存*.xlsx'))
    inv_path = sorted(inv_files)[-1] if inv_files else None
if inv_path:
    df_inv = read_file(inv_path)
    df_inv.columns = [str(c).strip() for c in df_inv.columns]
    df_inv['商品編號'] = df_inv['商品編號'].astype(str).str.strip()
    df_inv['倉庫編號'] = df_inv['倉庫編號'].astype(str).str.strip()
    df_inv['庫存數量'] = df_inv['庫存數量'].apply(safe_int).clip(lower=0)
    # 顏色欄：「顏色」是中文名稱，優先用
    if '顏色' in df_inv.columns:
        df_inv['_color'] = df_inv['顏色'].fillna('').astype(str).str.strip()
    elif '顏色名稱' in df_inv.columns:
        df_inv['_color'] = df_inv['顏色名稱'].fillna('').astype(str).str.strip()
    elif '商品顏色' in df_inv.columns:
        df_inv['_color'] = df_inv['商品顏色'].fillna('').astype(str).str.strip()
    else:
        df_inv['_color'] = ''
    sz_col = '商品尺寸' if '商品尺寸' in df_inv.columns else ('尺寸' if '尺寸' in df_inv.columns else None)
    df_inv['_size'] = df_inv[sz_col].fillna('').astype(str).str.strip() if sz_col else ''
    m = re.search(r'(\d{8})', os.path.basename(inv_path))
    inv_snap_date = m.group(1) if m else str(TODAY).replace('-','')
else:
    df_inv = pd.DataFrame()
    inv_snap_date = str(TODAY).replace('-','')# ════════════════════════════════════════════════════════
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
week_starts  = [this_monday - timedelta(weeks=i) for i in range(4,0,-1)]

products_data = []

_df_prod_unique = df_prod.drop_duplicates(subset='商品編號', keep='last').reset_index(drop=True)
print(f"  新品基本資料原始筆數：{len(df_prod)}　去重複後唯一商品數：{len(_df_prod_unique)}")

for _, prod in _df_prod_unique.iterrows():
    pid     = str(prod['商品編號']).strip()
    style   = str(prod.get('款號','')).strip()
    name    = str(prod.get('商品名稱','')).strip()
    cat     = str(prod.get('商品大分類','')).strip()
    mid_cat = str(prod.get('商品中分類','')).strip()
    brand   = str(prod.get('brand_display', prod.get('brand',''))).strip()
    year    = str(prod.get('年度','')).strip()
    season  = str(prod.get('季節','')).strip()
    img_url = str(prod.get('img_url','')).strip()
    barcode = str(prod.get('barcode','')).strip()
    material= str(prod.get('商品成分','')).strip()
    thickness_raw = str(prod.get('厚度','')).strip()
    try: thickness = int(float(thickness_raw))
    except: thickness = -1
    launch  = prod.get('launch_date')
    end_dt  = prod.get('end_date')

    if launch is None: continue

    eligible_stores = get_eligible_stores(pid)
    if not eligible_stores: continue

    if df_sales.empty:
        p_sales = pd.DataFrame()
    else:
        p_sales = df_sales[(df_sales['商品編號']==pid) & df_sales['倉庫編號'].isin(eligible_stores)].copy()

    if df_inv.empty:
        p_inv = pd.DataFrame()
    else:
        p_inv = df_inv[(df_inv['商品編號']==pid) & df_inv['倉庫編號'].isin(eligible_stores)].copy()

    colors = []
    if not p_sales.empty:
        colors = [c for c in p_sales['商品顏色'].unique() if c and c!='nan']
    if not colors and not p_inv.empty:
        colors = [c for c in p_inv['_color'].unique() if c and c!='nan']
    if not colors: colors = ['']

    for color in colors:
        c_sales = p_sales[p_sales['商品顏色']==color] if not p_sales.empty else pd.DataFrame()
        c_inv   = p_inv[p_inv['_color']==color]       if not p_inv.empty   else pd.DataFrame()

        cum_sales = int(c_sales['訂單數量'].sum()) if not c_sales.empty else 0

        inv_by_sz = {}
        if not c_inv.empty:
            for _, ir in c_inv.iterrows():
                sz = str(ir['_size']).strip()
                inv_by_sz[sz] = inv_by_sz.get(sz,0) + safe_int(ir['庫存數量'])
        color_total_inv  = sum(inv_by_sz.values())
        color_outlet_inv = int(c_inv[c_inv['倉庫編號'].isin(OUTLET_STORES)]['庫存數量'].sum()) if not c_inv.empty else 0
        color_main_inv   = color_total_inv - color_outlet_inv
        color_store_inv  = int(c_inv[~c_inv['倉庫編號'].str.startswith('8')]['庫存數量'].sum()) if not c_inv.empty else 0

        total_purchase = cum_sales + color_total_inv
        sellthru = cum_sales / total_purchase if total_purchase > 0 else 0.0

        if not c_sales.empty:
            sd = c_sales.groupby(['倉庫編號','_date'])['訂單數量'].sum().reset_index()
            store_days_sold = len(sd)
            inv_stores_w_stock = set(c_inv[c_inv['庫存數量']>0]['倉庫編號'].tolist()) if not c_inv.empty else set()
            extra = len(inv_stores_w_stock - set(sd['倉庫編號'].unique())) * 7
            total_store_days = store_days_sold + extra
        else:
            total_store_days = 0
        effective_spsd = cum_sales / total_store_days if total_store_days > 0 else 0.0

        supply_warn = False
        if not c_sales.empty and launch:
            early_cut = launch + timedelta(days=14)
            early_stores  = c_sales[c_sales['_date']<=early_cut]['倉庫編號'].nunique()
            recent_stores = c_sales[c_sales['_date']>=(TODAY-timedelta(days=7))]['倉庫編號'].nunique()
            supply_warn = early_stores>0 and recent_stores < early_stores*0.7

        weekly_sales = []
        for ws in week_starts:
            we = ws + timedelta(days=6)
            wqty = int(c_sales[(c_sales['_date']>=ws)&(c_sales['_date']<=we)]['訂單數量'].sum()) if not c_sales.empty else 0
            weekly_sales.append({'week': ws.strftime('%m/%d'), 'qty': wqty})all_sizes_raw = set()
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

            days_to_55 = None
            if not sz_s.empty and sz_pur > 0:
                target = sz_pur * 0.55
                cumsum = 0
                for _, sr in sz_s.sort_values('訂單日期').iterrows():
                    cumsum += safe_int(sr['訂單數量'])
                    if cumsum >= target:
                        days_to_55 = (sr['_date'] - launch).days + 1
                        break

            sz_sd = len(sz_s.groupby(['倉庫編號','_date'])) if not sz_s.empty else 0
            sz_spsd = sz_cum / sz_sd if sz_sd > 0 else 0.0

            sz_data.append({
                'sz': sz, 'inv': sz_inv, 'cum_sales': sz_cum,
                'sellthru_a': round(sz_st,4),
                'days_to_55_b': days_to_55,
                'spsd_c': round(sz_spsd,4),
                'is_oos': sz_inv == 0,
            })

        sz_pur_map   = {s['sz']: s['cum_sales']+s['inv'] for s in sz_data}
        total_pur_sz = sum(sz_pur_map.values())
        sz_ho_ratio  = {sz: v/total_pur_sz for sz,v in sz_pur_map.items()} if total_pur_sz>0 else {}
        headorder_rank = sorted(sz_ho_ratio, key=lambda x:(-sz_ho_ratio[x], sz_key(x)))
        sellthru_rank  = sorted(sz_sellthru_map, key=lambda x:(-sz_sellthru_map[x], sz_key(x)))
        size_shift_warn = headorder_rank != sellthru_rank and len(headorder_rank)>1

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

        n_stores = len(eligible_stores - OUTLET_STORES - IP_STORES) or 1
        turnover_days = round(color_main_inv / (effective_spsd * n_stores), 1) if effective_spsd>0 else None

        complete_stores = 0
        if all_sizes_sorted and not c_inv.empty:
            for sid in (eligible_stores - OUTLET_STORES - IP_STORES):
                s_inv = c_inv[c_inv['倉庫編號']==sid]
                if all(s_inv[s_inv['_size']==sz]['庫存數量'].sum()>0 for sz in all_sizes_sorted):
                    complete_stores += 1
        main_stores = len(eligible_stores - OUTLET_STORES - IP_STORES) or 1
        size_complete_rate = complete_stores / main_stores

        already_ordered = total_ordered_map.get((pid,color), 0)
        transit_total   = sum(transit_map.get((pid,color,sz['sz']),0) for sz in sz_data)
        transit_by_sz   = {sz['sz']: transit_map.get((pid,color,sz['sz']),0) for sz in sz_data}

        days_left = (end_dt - TODAY).days if end_dt else 999

        products_data.append({
            'pid': pid, 'style': style, 'name': name,
            'brand': brand, 'cat': cat, 'mid_cat': mid_cat,
            'year': year, 'season': season,
            'color': color, 'img_url': img_url, 'barcode': barcode,
            'material': material, 'thickness': thickness,
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

print(f"  共 {len(products_data)} 筆款色")# ════════════════════════════════════════════════════════
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
print("  ✅ output.xlsx")# ════════════════════════════════════════════════════════
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

print("🌐 產出 index.html...")

with open(os.path.join(BASE_DIR,'index.html'),'r',encoding='utf-8') as f:
    old_html = f.read()

import re as _re
new_html = _re.sub(
    r'const DATA=\{.*?\};',
    f'const DATA={json_str};',
    old_html,
    flags=_re.DOTALL
)

with open(os.path.join(BASE_DIR,'index.html'),'w',encoding='utf-8') as f:
    f.write(new_html)

print(f"✅ 完成！index.html 產出")
print(f"   款色：{len(products_data)}　建議追加：{sum(1 for p in products_data if p['reorder']>0)}")
