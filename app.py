"""Standalone, read-only Airport PIC portal. Does not import or change the admin app."""
import hashlib
import hmac
import io
import json
import time
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
[data-testid="stMetric"]{background:white;border:1px solid #E2E8F0;border-top:3px solid #C5A66B;border-radius:12px;padding:15px}
.hero{background:linear-gradient(110deg,#142B49,#284D73);padding:28px 34px;border-radius:16px;color:white;margin-bottom:24px}
.hero h1{color:white!important;margin:0;font-size:2rem}.hero p{color:#D8E5F1;margin:8px 0 0}
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
c1,c2,c3,c4=st.columns(4)
c1.metric('Contest Entry Target',f'{target_entries:,}')
c2.metric('Actual Contest Entries',f'{entries:,}')
c3.metric('Entry Achievement',f'{entries/target_entries*100:,.1f}%' if target_entries else '—')
c4.metric('Contest Sales (RM)',f'{sales:,.2f}')
a,b,c=st.columns(3)
a.metric('Contest Sales Target (RM)',f'{target_sales:,.2f}')
b.metric('Contest Sales Achievement',f'{sales/target_sales*100:,.1f}%' if target_sales else '—')
c.metric('Annual Airport Sales Target (RM)',f'{annual_airport_sales:,.2f}')
st.caption('Contest sales are receipt values from campaign entries. Annual airport sales targets are a separate measure; airport-wide actual sales are not inferred from receipts.')

st.markdown('### Monthly Performance')
monthly_raw=only_airport(data.get('monthly_online_airports'),code)
monthly_rows=[]
for month,label in MONTHS.items():
    subset=[r for r in monthly_raw if int(r.get('month') or 0)==month]
    monthly_rows.append({'Month':label,'Contest Entries':sum(int(r.get('entries') or 0) for r in subset),
                         'Contest Sales (RM)':round(sum(float(r.get('sales_rm') or 0) for r in subset),2)})
monthly=pd.DataFrame(monthly_rows)
st.dataframe(monthly,hide_index=True,use_container_width=True)
if not data.get('monthly_actuals_scope'):
    st.info('The source publication does not specify whether monthly figures include physical forms. Confirm its reporting scope with the administrator.')

st.markdown('### Merchant Performance')
month_choice=st.selectbox('Reporting period',['Overall',*MONTHS.values()])
category_choice=st.selectbox('Spending category',['All','Retail','Restaurant','Services'])
if month_choice=='Overall':
    merchant_rows=only_airport(data.get('merchant_locations'),code,field='code')
    if category_choice!='All':
        st.info('Overall merchant totals do not contain spending category labels in the existing publication. Select a month to filter by category.')
else:
    merchant_rows=only_airport(data.get('monthly_merchant_locations'),code,field='code')
    month_number=next(k for k,v in MONTHS.items() if v==month_choice)
    merchant_rows=[r for r in merchant_rows if int(r.get('month') or 0)==month_number]
    if category_choice!='All': merchant_rows=[r for r in merchant_rows if str(r.get('category','')).casefold()==category_choice.casefold()]
if merchant_rows:
    merchant_df=pd.DataFrame([{'Merchant':r.get('merchant',''),'Shop Location':r.get('shop_location',''),
                              'Category':r.get('category','—'),'Entries':int(r.get('entries') or 0)} for r in merchant_rows])
    merchant_df=merchant_df.groupby(['Merchant','Shop Location','Category'],as_index=False,dropna=False)['Entries'].sum()
    merchant_df=merchant_df.sort_values(['Entries','Merchant'],ascending=[False,True]).reset_index(drop=True)
    merchant_df.insert(0,'Rank',merchant_df.index+1)
else:
    merchant_df=pd.DataFrame(columns=['Rank','Merchant','Shop Location','Category','Entries'])
st.dataframe(merchant_df,hide_index=True,use_container_width=True)
st.caption(f'{len(merchant_df):,} merchant locations in this view. Entries count is aggregated from the published dataset.')
overview={'Airport':AIRPORTS[code],'Code':code,'Entry Target':target_entries,'Actual Entries':entries,
          'Contest Sales Target (RM)':target_sales,'Contest Sales Actual (RM)':sales,
          'Annual Airport Sales Target (RM)':annual_airport_sales,'Published At':str(published.get('updated_at',''))}
st.download_button('Download Airport Report (Excel)',excel_bytes(monthly,merchant_df,overview),
                   file_name=f'Belanja_Menang_{code}_Report.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
