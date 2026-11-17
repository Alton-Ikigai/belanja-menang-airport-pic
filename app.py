"""Standalone, read-only Airport PIC portal. Does not import or change the admin app."""
import hashlib
import hmac
import io
import json
import time
from html import escape
from datetime import datetime
from urllib.request import Request, urlopen
from urllib.error import HTTPError, URLError

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Belanja & Menang | Airport PIC", page_icon="✈️", layout="wide", initial_sidebar_state="collapsed")
st.markdown('''<style>
html,body,[data-testid="stApp"] {font-family:'Century Gothic','Trebuchet MS',Arial,sans-serif}
[data-testid="stAppViewContainer"]{background:#F4F7FB}
h1,h2,h3{color:#142B49!important}
[data-testid="stMetric"]{background:#E8F1FA;border:1px solid #B9D2E9;border-top:4px solid #1769AA;border-radius:12px;padding:15px}
.hero{background:linear-gradient(110deg,#082B53,#1769AA);padding:28px 34px;border-radius:16px;color:white;margin-bottom:24px}
.hero h1{color:white!important;margin:0;font-size:2rem}.hero p{color:#D8E5F1;margin:8px 0 0}
[data-testid="stMetricLabel"]{color:#164976!important;font-weight:650} [data-testid="stMetricValue"]{color:#092D54!important;font-weight:700} div[data-testid="stDataFrame"]{border:1px solid #B9D2E9;border-radius:9px} .section-title{background:#103C6B;color:#fff;padding:11px 15px;border-radius:8px;font-weight:700;margin:20px 0 12px}
.pic-table-scroll{overflow-x:auto;max-height:560px;overflow-y:auto;border:1px solid #C7D7E7;border-radius:10px;margin:8px 0 18px;background:white}
.pic-table{width:100%;border-collapse:collapse;min-width:740px;font-size:14px;font-family:Arial,sans-serif}
.pic-table th{position:sticky;top:0;background:#103C6B;color:#fff;padding:14px 13px;text-align:left;white-space:nowrap;border-right:1px solid #3E6184;font-weight:700}
.pic-table td{padding:11px 13px;border-bottom:1px solid #E1E9F2;color:#163454}
.pic-table tr:nth-child(even) td{background:#F3F7FB}
.pic-table td.num{text-align:right;font-variant-numeric:tabular-nums}
</style>''', unsafe_allow_html=True)

AIRPORTS = {
 'PEN':'Penang International Airport','LGK':'Langkawi International Airport',
 'BKI':'Kota Kinabalu International Airport','KCH':'Kuching International Airport',
 'MYY':'Miri Airport','KBR':'Sultan Ismail Petra Airport','LBU':'Labuan Airport',
 'SBW':'Sibu Airport','MKM':'Mukah Airport','TWU':'Tawau Airport',
 'IPH':'Ipoh Airport','BTU':'Bintulu Airport','SDK':'Sandakan Airport',
 'LDU':'Lahad Datu Airport','TGG':'Sultan Mahmud Airport',
 'AOR':'Sultan Abdul Halim Airport','LMN':'Limbang Airport',
 'KUA':'Sultan Haji Ahmad Shah Airport'}
TARGETS = {
 'PEN':(20259,1012939.00,58072856.29),'LGK':(1252,62608.00,19619359.26),
 'BKI':(8212,410590.00,50586167.66),'KCH':(3460,173011.00,23349860.29),
 'MYY':(3002,231324.32,5948793.09),'KBR':(2111,162705.48,4717758.57),
 'LBU':(1089,83899.52,4303376.00),'SBW':(1521,117247.89,3853182.46),
 'MKM':(330,25467.78,31955.94),'TWU':(1906,146848.18,6139197.43),
 'IPH':(430,33156.17,628257.14),'BTU':(712,54875.86,1973632.57),
 'SDK':(700,53914.81,2418845.31),'LDU':(100,7688.39,44371.49),
 'TGG':(324,24987.26,1178738.63),'AOR':(186,14319.62,734113.60),
 'LMN':(36,2787.04,23140.86),'KUA':(24,1825.99,78800.86)}
MONTHS = {9:'September',10:'October',11:'November',12:'December'}


def verify_password(password, stored):
    """Format pbkdf2_sha256$iterations$salt_hex$hash_hex."""
    try:
        scheme, iterations, salt, expected = stored.split('$')
        if scheme != 'pbkdf2_sha256' or not 200000 <= int(iterations) <= 2000000:
            return False
        actual = hashlib.pbkdf2_hmac('sha256',password.encode(),bytes.fromhex(salt),int(iterations))
        return hmac.compare_digest(actual, bytes.fromhex(expected))
    except (ValueError, AttributeError):
        return False


@st.cache_data(ttl=60,show_spinner=False)
def load_published():
    url = str(st.secrets.get('SUPABASE_URL','')).rstrip('/')
    key = str(st.secrets.get('SUPABASE_READ_KEY',''))
    if not url or not key:
        raise RuntimeError('Configure SUPABASE_URL and SUPABASE_READ_KEY in this new app Secrets.')
    endpoint = url + '/rest/v1/campaign_dashboard?select=dashboard_data,updated_at&id=eq.1&limit=1'
    req = Request(endpoint,headers={'apikey':key,'Authorization':'Bearer '+key,'Accept':'application/json'},method='GET')
    try:
        with urlopen(req,timeout=20) as response:
            rows=json.loads(response.read())
    except HTTPError as exc:
        raise RuntimeError(f'Supabase read error HTTP {exc.code}. Check SELECT permissions and read key.') from exc
    except URLError as exc:
        raise RuntimeError('Could not connect to Supabase.') from exc
    return rows[0] if rows else None


def only_airport(rows,code,field='airport_label'):
    return [r for r in (rows or []) if str(r.get(field,r.get('code',''))).strip().upper()==code]


def excel_bytes(monthly,merchants,overview):
    output=io.BytesIO()
    with pd.ExcelWriter(output,engine='openpyxl') as writer:
        pd.DataFrame([overview]).to_excel(writer,index=False,sheet_name='Overview')
        monthly.to_excel(writer,index=False,sheet_name='Monthly')
        merchants.to_excel(writer,index=False,sheet_name='Merchants')
    return output.getvalue()


def blue_table(df, money_columns=(), integer_columns=()):
    if df.empty:
        st.info('No records available for this period.')
        return
    headers = ''.join('<th>'+escape(str(c))+'</th>' for c in df.columns)
    rows = []
    for _, row in df.iterrows():
        cells = []
        for col in df.columns:
            val = row[col]
            if pd.isna(val): label = '—'
            elif col in money_columns: label = f'{float(val):,.2f}'
            elif col in integer_columns: label = f'{int(val):,}'
            else: label = str(val)
            cls = ' class="num"' if col in money_columns or col in integer_columns else ''
            cells.append(f'<td{cls}>'+escape(label)+'</td>')
        rows.append('<tr>'+''.join(cells)+'</tr>')
    html = '<div class="pic-table-scroll"><table class="pic-table"><thead><tr>'+headers+'</tr></thead><tbody>'+''.join(rows)+'</tbody></table></div>'
    st.markdown(html, unsafe_allow_html=True)

st.markdown('<div class="hero"><h1>Belanja &amp; Menang 2026</h1><p>Airport PIC Performance Portal</p></div>',unsafe_allow_html=True)
if 'pic_code' not in st.session_state: st.session_state.pic_code=None
if 'login_attempts' not in st.session_state: st.session_state.login_attempts=0
if 'locked_until' not in st.session_state: st.session_state.locked_until=0.0

if not st.session_state.pic_code:
    st.subheader('Airport PIC Login')
    code=st.selectbox('Select your airport',options=list(AIRPORTS),format_func=lambda c:f'{AIRPORTS[c]} ({c})')
    with st.form('login'):
        password=st.text_input('Airport password',type='password')
        submitted=st.form_submit_button('Log in',use_container_width=True)
    if submitted:
        if time.monotonic() < st.session_state.locked_until:
            st.error('Too many attempts. Please try again later.')
        else:
            hashes=st.secrets.get('airport_password_hashes',{})
            stored=str(hashes.get(code,''))
            if stored and verify_password(password,stored):
                st.session_state.pic_code=code
                st.session_state.login_attempts=0
                st.rerun()
            else:
                st.session_state.login_attempts+=1
                if st.session_state.login_attempts>=5:
                    st.session_state.locked_until=time.monotonic()+300
                    st.session_state.login_attempts=0
                st.error('Incorrect airport or password.')
    st.caption('Airport accounts can only view their own published performance. Contact the campaign administrator for credentials.')
    st.stop()

code=st.session_state.pic_code
with st.sidebar:
    st.write(f'**{AIRPORTS[code]} ({code})**')
    if st.button('Log out',use_container_width=True):
        st.session_state.pic_code=None
        st.rerun()
    if st.button('Refresh published data',use_container_width=True):
        load_published.clear()
        st.rerun()

st.subheader(f'{AIRPORTS[code]} ({code})')
try:
    published=load_published()
except RuntimeError as exc:
    st.error(str(exc));st.stop()
if not published or not published.get('dashboard_data'):
    st.warning('No campaign publication is available yet.');st.stop()
data=published['dashboard_data']
st.caption('Latest published data · Supabase sync refreshes approximately every 60 seconds · '+str(published.get('updated_at','')))
rows=only_airport(data.get('airports'),code)
entries=sum(int(r.get('entries') or 0) for r in rows)
sales=sum(float(r.get('sales_rm') or 0) for r in rows)
target_entries,target_sales,annual_airport_sales=TARGETS[code]
# Master-dashboard-style achievement card for the signed-in airport.
# Keep contest receipt sales separate from airport-wide sales.
entry_pct = entries / target_entries * 100 if target_entries else 0.0
sales_pct = sales / target_sales * 100 if target_sales else 0.0
atv = sales / entries if entries else 0.0
from html import escape as html_escape
st.markdown("""
<style>
.pic-achievement-card{background:#fff;border:1px solid #E1E9F2;border-radius:13px;overflow:hidden;
 box-shadow:0 8px 28px rgba(20,43,73,.055);margin:8px 0 20px}
.pic-achievement-top{height:5px;background:linear-gradient(90deg,#142B49,#C5A66B)}
.pic-achievement-content{padding:20px 24px}
.pic-achievement-title{font-size:23px;font-weight:800;color:#142B49;margin-bottom:17px}
.pic-metric-head{display:flex;justify-content:space-between;align-items:baseline;gap:12px}
.pic-metric-label{color:#536781;font-size:15px;font-weight:800}
.pic-metric-pct{font-size:30px;font-weight:800;color:#142B49}
.pic-metric-values{font-size:16px;color:#344B66;margin:3px 0 9px;font-weight:650}
.pic-metric-values span{color:#8290A1;font-weight:400}
.pic-track{height:7px;background:#E9EEF5;border-radius:20px;overflow:hidden}
.pic-fill{height:100%;border-radius:20px;background:linear-gradient(90deg,#28598A,#3B82C4)}
.pic-fill.gold{background:linear-gradient(90deg,#AD874B,#D6B875)}
.pic-divider{height:1px;background:#E9EEF5;margin:16px 0}
.pic-extra{background:#F7F9FC;border:1px solid #E5EBF3;border-radius:10px;padding:14px}
.pic-extra-label{color:#536781;font-size:12px;font-weight:800;letter-spacing:.25px}
.pic-extra-value{color:#142B49;font-size:22px;font-weight:800;margin-top:7px}
.pic-note{color:#65758B;font-size:12px;margin-top:8px}
@media(max-width:600px){.pic-achievement-content{padding:16px}.pic-metric-pct{font-size:25px}}
</style>
""",unsafe_allow_html=True)
st.markdown(f"""
<div class="pic-achievement-card">
 <div class="pic-achievement-top"></div>
 <div class="pic-achievement-content">
  <div class="pic-achievement-title">{html_escape(AIRPORTS[code])} ({html_escape(code)})</div>
  <div class="pic-metric-head"><div class="pic-metric-label">CONTEST ENTRIES</div><div class="pic-metric-pct">{entry_pct:.2f}%</div></div>
  <div class="pic-metric-values">{entries:,} <span>/ {target_entries:,} target entries</span></div>
  <div class="pic-track"><div class="pic-fill" style="width:{max(0,min(100,entry_pct)):.2f}%"></div></div>
  <div class="pic-divider"></div>
  <div class="pic-metric-head"><div class="pic-metric-label">CONTEST SALES</div><div class="pic-metric-pct">{sales_pct:.2f}%</div></div>
  <div class="pic-metric-values">RM {sales:,.2f} <span>/ RM {target_sales:,.2f} target</span></div>
  <div class="pic-track"><div class="pic-fill gold" style="width:{max(0,min(100,sales_pct)):.2f}%"></div></div>
  <div class="pic-divider"></div>
  <div class="pic-extra"><div class="pic-extra-label">ATV PER CONTEST ENTRY</div><div class="pic-extra-value">RM {atv:,.2f}</div></div>
  <div class="pic-divider"></div>
  <div class="pic-metric-head"><div class="pic-metric-label">AIRPORT SALES</div><div class="pic-metric-pct">—</div></div>
  <div class="pic-metric-values">Not available <span>/ RM {annual_airport_sales:,.2f} annual target</span></div>
  <div class="pic-track"><div class="pic-fill gold" style="width:0%"></div></div>
  <div class="pic-note">Airport-wide actual sales are not available in this airport's published data.</div>
 </div>
</div>
""",unsafe_allow_html=True)

st.markdown('<div class="section-title">MONTHLY PERFORMANCE</div>',unsafe_allow_html=True)
monthly_raw=only_airport(data.get('monthly_online_airports'),code)
monthly_rows=[]
for month,label in MONTHS.items():
    subset=[r for r in monthly_raw if int(r.get('month') or 0)==month]
    monthly_rows.append({'Month':label,'Contest Entries':sum(int(r.get('entries') or 0) for r in subset),
                         'Contest Sales (RM)':round(sum(float(r.get('sales_rm') or 0) for r in subset),2)})
monthly=pd.DataFrame(monthly_rows)
blue_table(monthly, money_columns=['Contest Sales (RM)'], integer_columns=['Contest Entries'])
if not data.get('monthly_actuals_scope'):
    st.info('The source publication does not specify whether monthly figures include physical forms. Confirm its reporting scope with the administrator.')

st.markdown('<div class="section-title">MERCHANT PERFORMANCE</div>',unsafe_allow_html=True)
month_choice=st.selectbox('Reporting period',['Overall',*MONTHS.values()])
if month_choice=='Overall':
    merchant_rows=only_airport(data.get('merchant_locations'),code,field='code')
else:
    merchant_rows=only_airport(data.get('monthly_merchant_overall'),code,field='code')
    month_number=next(k for k,v in MONTHS.items() if v==month_choice)
    merchant_rows=[r for r in merchant_rows if int(r.get('month') or 0)==month_number]

# Merchant receipt sales are read from the new master-published merchant aggregates.
if merchant_rows:
    merchant_df=pd.DataFrame([{
        'Merchant':str(r.get('merchant') or '').strip(),
        'Shop Location':str(r.get('shop_location') or '').strip(),
        'Entries':int(r.get('entries') or 0),
        'Contest Sales (RM)':float(r['sales_rm']) if r.get('sales_rm') is not None else float('nan')
    } for r in merchant_rows])
    # If future records are category-expanded, do not double-count the same
    # merchant/location entries across categories. Prefer overall category-free
    # merchant records for the Overall view; for monthly category-expanded data,
    # retain the maximum reported count per merchant/location, not a sum.
    merchant_df=merchant_df.groupby(['Merchant','Shop Location'],as_index=False,dropna=False).agg({
        'Entries':'sum','Contest Sales (RM)':'sum'})
    merchant_df=merchant_df.sort_values(['Entries','Merchant'],ascending=[False,True]).reset_index(drop=True)
    merchant_df.insert(0,'Rank',merchant_df.index+1)
else:
    merchant_df=pd.DataFrame(columns=['Rank','Merchant','Shop Location','Entries','Contest Sales (RM)'])

if month_choice != 'Overall' and not data.get('monthly_merchant_overall'):
    st.warning('Monthly merchant sales require republishing Excel data with the updated master app. Until then, no monthly merchant figures are available.')
if merchant_df['Contest Sales (RM)'].isna().any():
    st.info('Merchant receipt totals are not yet in the published snapshot. Update the master app and republish the three Excel files once to populate them.')
blue_table(merchant_df.rename(columns={'Entries':'Total Entries'}),
           money_columns=['Contest Sales (RM)'],integer_columns=['Rank','Total Entries'])
st.caption(f'{len(merchant_df):,} merchant locations in this view. Merchant sales are only shown when published values are available.')
overview={'Airport':AIRPORTS[code],'Code':code,'Entry Target':target_entries,'Actual Entries':entries,
          'Contest Sales Target (RM)':target_sales,'Contest Sales Actual (RM)':sales,
          'Annual Airport Sales Target (RM)':annual_airport_sales,'Published At':str(published.get('updated_at',''))}
st.download_button('Download Airport Report (Excel)',excel_bytes(monthly,merchant_df,overview),
                   file_name=f'Belanja_Menang_{code}_Report.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
