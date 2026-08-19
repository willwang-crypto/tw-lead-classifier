"""  
Sales Ops · Data Quality Suite — Singapore  
Delivery Hero / Pandora · Digital Sales APAC

TAB 1  Lead Classification  
       Postal \+ Name dedup (unit used as filter)  
       Labels: P1 / P2 / P3 / P4 / Business Closed / Wrong Target Group

TAB 2  Generate Apify URLs

TAB 3  SF Account Audit  
       Find suspected duplicate pairs within Salesforce itself

Label reference:  
  P1  — New               No CRM match, Apify-confirmed restaurant  
  P2  — Please Check      No Apify result or no category found  
  P3  — Potential Match   Name 0.50–0.74 at same postal code  
  P4  — Duplicate         Name ≥ 0.75 at same postal code  
  Business Closed         Apify: Google confirms permanently/temporarily closed  
  Wrong Target Group      Apify: category not food-delivery eligible  
"""

import streamlit as st  
import pandas as pd  
import re  
import io  
import time  
import json  
from urllib.parse import unquote, quote, urlencode  
from urllib.request import urlopen  
from difflib import SequenceMatcher          \# used only for street\_similarity  
from itertools import combinations  
from openpyxl import Workbook  
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side  
from openpyxl.utils import get\_column\_letter  
from openpyxl.worksheet.datavalidation import DataValidation  
from shapely.wkt import loads as wkt\_loads  
from shapely.geometry import Point  
from rapidfuzz import fuzz

try:  
    from pypinyin import lazy\_pinyin, Style as PinyinStyle  
    \_PYPINYIN\_AVAILABLE \= True  
except ImportError:  
    \_PYPINYIN\_AVAILABLE \= False

\# ── Secrets & caches ─────────────────────────────────────────────  
APP\_PASSWORD        \= st.secrets\["APP\_PASSWORD"\]  
ONEMAP\_CREDENTIALS  \= {"email": st.secrets\["ONEMAP\_EMAIL"\],  
                        "password": st.secrets\["ONEMAP\_PASSWORD"\]}  
\_TOKEN\_CACHE        \= {"token": None, "expiry": 0}

def \_load\_logo():  
    import os, base64  
    p \= os.path.join(os.path.dirname(os.path.abspath(\_\_file\_\_)), "dh\_logo.png")  
    return base64.b64encode(open(p, "rb").read()).decode() if os.path.exists(p) else ""

DH\_LOGO\_B64 \= \_load\_logo()

def check\_password():  
    if st.session\_state.get("authenticated"):  
        return True  
    \_, col, \_ \= st.columns(\[1, 1.6, 1\])  
    with col:  
        st.markdown(f'''  
        \<div style="text-align:center;padding:2rem 0 1rem 0;"\>  
            \<img src="data:image/png;base64,{DH\_LOGO\_B64}"  
                 style="width:140px;margin-bottom:1.2rem;" /\>  
            \<h2 style="color:\#1A1A1A;font-size:1.4rem;font-weight:700;margin-bottom:0.2rem;"\>  
                Sales Ops · Data Quality Suite\</h2\>  
            \<p style="color:\#888;font-size:0.88rem;margin-bottom:1.4rem;"\>  
                Digital Sales APAC · Pandora / Delivery Hero\</p\>  
        \</div\>''', unsafe\_allow\_html=True)  
        pwd \= st.text\_input("Password", type="password",  
                            placeholder="Enter password to continue")  
        if st.button("Sign in", type="primary", use\_container\_width=True):  
            if pwd \== APP\_PASSWORD:  
                st.session\_state\["authenticated"\] \= True  
                st.rerun()  
            else:  
                st.error("Incorrect password.")  
    return False

st.set\_page\_config(page\_title="Sales Ops Suite",  
                   page\_icon="🎯", layout="wide",  
                   initial\_sidebar\_state="expanded")

\# ═════════════════════════════════════════════════════════════════  
\# STATUS GROUPS  (shared by both tabs)  
\# ═════════════════════════════════════════════════════════════════

ACTIVE\_PIPELINE \= \[  
    "active", "new", "collecting documents", "negotiation",  
    "menu processing", "onboarding", "quality check",  
\]  
WIN\_BACK        \= \["lost", "terminated"\]  
WIN\_BACK\_FAILED \= \["win back failed"\]

def has\_closed\_marker(name) \-\> bool:  
    """True if the account name ends with '- Closed' (common SF convention)."""  
    if not name or pd.isna(name):  
        return False  
    return bool(re.search(r'-\\s\*closed', str(name).strip().lower()))

def get\_status\_group(status, is\_closed\_name: bool \= False) \-\> str:  
    if is\_closed\_name:  
        return "win\_back"  
    s \= str(status).strip().lower()  
    if s in WIN\_BACK\_FAILED: return "win\_back\_failed"  
    if s in WIN\_BACK:        return "win\_back"  
    return "active\_pipeline"

def get\_crm\_label(match\_method: str, match\_score: float,  
                  status\_group: str, p2: float, p3: float) \-\> str:  
    """Map (match quality, CRM status) → lead priority label."""  
    if status\_group \== "win\_back\_failed":  
        return "WBF — Win-back Failed"  
    if status\_group \== "win\_back":  
        return "WB — Win-back"  
    \# active\_pipeline  
    if match\_method \== "GRID Exact" or match\_score \>= 1.0:  
        return "P4 — Duplicate (Exact)"  
    if match\_score \>= p3:  
        return "P3 — Duplicate"  
    if match\_score \>= p2:  
        return "P2 — Potential"  
    return "Needs Review"

def get\_risk\_and\_action(status\_a, status\_b):  
    """Risk level \+ recommended action for SF Account Audit pairs."""  
    def grp(s):  
        sl \= str(s).strip().lower()  
        if sl in WIN\_BACK\_FAILED: return "wbf"  
        if sl in WIN\_BACK:        return "winback"  
        return "pipeline"  
    ga, gb \= grp(status\_a), grp(status\_b)  
    pair   \= frozenset(\[ga, gb\])  
    if pair \== frozenset(\["pipeline"\]):  
        if (str(status\_a).strip().lower() \== "active"  
                and str(status\_b).strip().lower() \== "active"):  
            return "🔴 High",   "Merge required — two live active accounts"  
        return     "🔴 High",   "Being worked twice — check with rep"  
    if pair \== frozenset(\["pipeline", "winback"\]):  
        return     "🟡 Medium", "Review — possibly stale record from prior attempt"  
    if pair \== frozenset(\["winback"\]):  
        return     "🟢 Low",    "Low priority — both inactive"  
    if "wbf" in pair:  
        return     "🟢 Low",    "Informational — win-back already attempted"  
    return         "🟡 Medium", "Review recommended"

\# ═════════════════════════════════════════════════════════════════  
\# MARKET CONFIG  
\# ═════════════════════════════════════════════════════════════════

MARKETS \= {  
    "SG": {  
        "code": "SG", "name": "Singapore", "flag": "🇸🇬",  
        "char\_map": {}, "country\_suffix": "Singapore", "phone\_prefix": "65",  
    }  
}

\# ═════════════════════════════════════════════════════════════════  
\# FOOD DELIVERY ELIGIBILITY  
\# ═════════════════════════════════════════════════════════════════

FOOD\_DELIVERY\_ALLOWED \= {  
    "Restaurant","Fine dining restaurant","Family restaurant","Casual dining restaurant",  
    "Buffet restaurant","Brasserie","Bistro","Diner","Eatery","Pizza restaurant",  
    "Pizza delivery","Pizza takeaway","Kebab shop","Kebab restaurant",  
    "Doner kebab restaurant","Shawarma restaurant","Falafel restaurant",  
    "Pita restaurant","Sushi restaurant","Ramen restaurant","Noodle restaurant",  
    "Dumpling restaurant","Dim sum restaurant","Wonton restaurant","Steak house",  
    "Steakhouse","Grill restaurant","Barbecue restaurant","BBQ restaurant",  
    "Smokehouse","Rotisserie chicken restaurant","Sandwich shop",  
    "Submarine sandwich shop","Wrap restaurant","Salad shop","Bowl restaurant",  
    "Poke bar","Soup restaurant","Soup kitchen","Breakfast restaurant",  
    "Brunch restaurant","Pancake restaurant","Waffle restaurant",  
    "Dessert restaurant","Dessert shop","Ice cream shop","Ice cream parlor",  
    "Frozen yogurt shop","Donut shop","Doughnut shop","Crepe restaurant",  
    "Waffle house","Vegetarian restaurant","Vegan restaurant",  
    "Plant-based restaurant","Organic restaurant","Health food restaurant",  
    "Gluten-free restaurant","Halal restaurant","Kosher restaurant","Food hall",  
    "Food truck","Street food restaurant","Market restaurant",  
    "Home cooking restaurant","Traditional restaurant","Local restaurant",  
    "Neighborhood restaurant","Deli","Delicatessen","Charcuterie","Lunchroom",  
    "Snack bar","Juice bar","Smoothie bar","Açaí shop","Chocolate shop",  
    "Sweet shop","Candy store","Noodle shop","Pasta shop","Rice restaurant",  
    "Porridge restaurant","Congee restaurant","Hot pot restaurant",  
    "Fondue restaurant","Raclette restaurant","Teppanyaki restaurant",  
    "Okonomiyaki restaurant","Takoyaki restaurant","Yakitori restaurant",  
    "Izakaya","Robatayaki restaurant","Tempura restaurant","Tonkatsu restaurant",  
    "Udon restaurant","Soba restaurant","Gyoza restaurant","Pho restaurant",  
    "Banh mi restaurant","Spring roll restaurant","Satay restaurant",  
    "Rendang restaurant","Curry restaurant","Tandoori restaurant",  
    "Biryani restaurant","Dosa restaurant","Idli restaurant","Thali restaurant",  
    "Ceviche restaurant","Empanada restaurant","Arepas restaurant",  
    "Chimichanga restaurant","Meal delivery","Food delivery","Takeaway",  
    "Takeout restaurant","Take-out restaurant","Cloud kitchen","Ghost kitchen",  
    "Virtual restaurant",  
    \# SG-specific  
    "Zi char restaurant","Cze char restaurant","Tze char restaurant",  
    "Economy rice stall","Nasi lemak restaurant","Chicken rice restaurant",  
    "Laksa restaurant","Wonton mee restaurant","Char kway teow restaurant",  
    "Bak kut teh restaurant","Murtabak restaurant","Prata restaurant",  
    "Mixed rice restaurant","Hawker-style restaurant","Nasi padang restaurant",  
    "Hawker stall",  
    \# Café & bakery  
    "Cafe","Coffee shop","Coffee house","Coffeehouse","Espresso bar",  
    "Tea house","Bubble tea shop","Boba shop","Bakery","Patisserie",  
    "Pastry shop","Cake shop","Bread bakery","Artisan bakery","French bakery",  
    "Cookie shop","Cupcake shop","Bagel shop",  
    \# Fast food / QSR  
    "Fast food restaurant","Fast-food restaurant","Quick service restaurant",  
    "Hamburger restaurant","Burger restaurant","Hot dog restaurant",  
    "Fried chicken restaurant","Chicken restaurant","Chicken wings restaurant",  
    "Fish and chips restaurant","Seafood restaurant","Fish restaurant",  
    "Taco restaurant","Burrito restaurant",  
    \# Cuisines  
    "Italian restaurant","French restaurant","Chinese restaurant",  
    "Japanese restaurant","Thai restaurant","Indian restaurant",  
    "Mexican restaurant","Greek restaurant","Lebanese restaurant",  
    "Middle Eastern restaurant","Mediterranean restaurant","Asian restaurant",  
    "Korean restaurant","Vietnamese restaurant","Spanish restaurant",  
    "American restaurant","Ethiopian restaurant","Afghan restaurant",  
    "Pakistani restaurant","Nepalese restaurant","Sri Lankan restaurant",  
    "Bangladeshi restaurant","Indonesian restaurant","Filipino restaurant",  
    "Peruvian restaurant","Brazilian restaurant","Argentinian restaurant",  
    "Georgian restaurant","Uzbek restaurant","Syrian restaurant",  
    "Moroccan restaurant","Egyptian restaurant","Caribbean restaurant",  
    "Jamaican restaurant","Cuban restaurant","Portuguese restaurant",  
    "German restaurant","Austrian restaurant","Scandinavian restaurant",  
    "Nordic restaurant","Latin American restaurant","Fusion restaurant",  
    "International restaurant","European restaurant","Pan-Asian restaurant",  
    "Oriental restaurant",  
    \# SG dialect / local tags  
    "Hokkien","Teochew","Cantonese","Dim Sum","Hainanese","Hakka","Shanghai",  
    "Sichuan","Hunan","Jiang Su","Putian","Dong Bei",  
    "Hong Kong (Cha Chaan Teng)","Seafood BBQ","Western","Dessert","Tang Shui",  
    "Cakes","Malay","Indonesian","Indian","Nyonya","Peranakan","Taiwan",  
    "Taiwanese","Turkish","Roast Meat","Nasi Padang","Indian Muslim",  
    "South Indian","Eurasian","Hotpot","Buffet","Seafood","Fusion","British",  
    "Australian","Cajun","Caribbean","Cuban","Greek","Halal","Internation",  
    "Mala","Mookata","Muslim","Middle Eastern","Steakhouse","Swedish","Vegan",  
    "Izakaya","Local",  
}

\_DEFAULT\_EXCLUSION\_KW \= \[  
    "hawker centre","hawker center","kopitiam",  
    "food court","food centre","food center","eating house",  
    "club","hotel","cantine","canteen","brewery","liquor","wine",  
    "酒吧","酒店","美食广场","美食中心","食阁","咖啡店","巴刹",  
    "karaoke","卡拉ok","bbq space","event space","salon",  
    "coffee beans retailer","shopping mall","catering","roastery",  
    "vending machine","cyber cafe","grocery","supermarket",  
    "alcohol shop","fruit shop","events","food festival","pop-up",  
    "night club","pet bakery","酒馆","沙龙","自动贩卖机","商场","超市",  
    "cc restaurant","community club","army","saf","mindef",  
    "school canteen","hospital canteen","industrial canteen",  
    "mix & match",  
\]

def is\_food\_delivery\_eligible(category, exclusion\_kw: list) \-\> bool:  
    if not category or pd.isna(category):  
        return False  
    cat \= str(category).strip()  
    cat\_lower \= cat.lower()  
    if any(kw in cat\_lower for kw in exclusion\_kw):  
        return False  
    if cat in FOOD\_DELIVERY\_ALLOWED:  
        return True  
    for allowed in FOOD\_DELIVERY\_ALLOWED:  
        if allowed.lower() \== cat\_lower:  
            return True  
    food\_kw \= \[  
        "restaurant","cafe","café","bakery","kebab","pizza","sushi","burger",  
        "grill","bistro","brasserie","diner","eatery","kitchen","takeaway",  
        "takeout","delivery","patisserie","pastry","coffee","tea house",  
        "noodle","ramen","deli","snack","food truck","steakhouse","seafood",  
        "sandwich","dim sum","hotpot","steak house","餐厅","餐馆","小吃","烘焙","面包",  
    \]  
    return any(kw in cat\_lower for kw in food\_kw)

\# ═════════════════════════════════════════════════════════════════  
\# SINGAPORE UTILITY FUNCTIONS  
\# ═════════════════════════════════════════════════════════════════

SG\_UNIT\_RE   \= re.compile(r'\#\\s\*(\\d{1,2})\\s\*\[-–\]\\s\*(\\d{1,4}\[A-Za-z\]?)', re.IGNORECASE)  
SG\_POSTAL\_RE \= re.compile(r'\\b(0\[1-9\]|\[1-7\]\\d|8\[0-2\])\\d{4}\\b')  
SG\_NAME\_NOISE= re.compile(  
    r'\\b(pte\\.?\\s\*ltd\\.?|sdn\\.?\\s\*bhd\\.?|llp|incorporated|'  
    r'holdings?|group|enterprise\[s\]?|trading|singapore|sg)\\b',  
    re.IGNORECASE)

\# Common Singapore area names for geographic conflict detection  
\# in the zero-postal name-only fallback.  
SG\_AREAS \= {  
    "tampines","jurong","woodlands","yishun","hougang","sengkang","punggol",  
    "bishan","ang mo kio","bedok","pasir ris","changi","geylang","kallang",  
    "clementi","choa chu kang","sembawang","queenstown","toa payoh",  
    "buona vista","orchard","serangoon","bukit timah","bukit batok",  
    "paya lebar","novena","marine parade","marsiling","simei","boon lay",  
    "admiralty","bras basah","chinatown","dhoby ghaut","farrer park",  
    "harbourfront","hillview","jurong east","jurong west","kovan","lavender",  
    "little india","macpherson","newton","potong pasir","redhill","tiong bahru",  
    "upper thomson","vivocity","whampoa",  
}

\# ── Strip venue names and generic words before name comparison ────  
\# Parenthetical content (e.g. "(Tanglin Mall)", "(White Sands)") and  
\# generic words ("restaurant/s") inflate fuzzy scores between  
\# co-located but unrelated businesses. Strip before normalising.  
\_PAREN\_RE   \= re.compile(r'\[)\]\*', re.UNICODE)  
\_GENERIC\_RE \= re.compile(r'\\b(restaurants?)\\b', re.IGNORECASE)

def strip\_venue\_generic(name: str) \-\> str:  
    """Remove parenthetical venue tags and generic restaurant words.

    Examples:  
        'Western Grill (Tanglin Mall)'          → 'Western Grill'  
        'KFC (White Sands)'                     → 'KFC'  
        'DOMO Restaurant'                        → 'DOMO'  
        'Shahi Maharani North Indian Restaurant' → 'Shahi Maharani North Indian'  
    """  
    s \= \_PAREN\_RE.sub(' ', str(name or ''))  
    s \= \_GENERIC\_RE.sub(' ', s)  
    return re.sub(r'\\s+', ' ', s).strip()  
\_CHINESE\_RE  \= re.compile(r'\[\\u4e00-\\u9fff\\u3400-\\u4dbf\\uf900-\\ufaff\]')  
\_NA\_VALUES   \= {"","nan","none","n/a","na","nil","-","–",  
                "unknown","no name","na/","n.a.","n.a"}

def extract\_sg\_unit(text: str) \-\> str:  
    if not text or str(text).strip() in ("","nan"):  
        return ""  
    m \= SG\_UNIT\_RE.search(str(text))  
    if m:  
        return f"{m.group(1).zfill(2)}-{m.group(2).lower().zfill(3)}"  
    return ""

def extract\_sg\_postal(text: str) \-\> str:  
    if not text or str(text).strip() in ("","nan"):  
        return ""  
    m \= SG\_POSTAL\_RE.search(str(text))  
    return m.group(0) if m else ""

\# Placeholder / obviously fake postal codes that should be treated as blank  
\_FAKE\_POSTALS \= frozenset({"000000","000001","111111","123456","999999"})

def \_norm\_postal\_input(raw) \-\> str:  
    """  
    Normalise a raw postal column value before extraction:  
    \- Strips non-digits and trailing .0 (Excel float artefact)  
    \- Zero-pads 5-digit codes (Excel drops leading zero for 09xxxx postals)  
    \- Returns '' for known fake / placeholder codes

    Only applies to postal COLUMN values, not to full address strings.  
    """  
    s \= re.sub(r'\\D', '', str(raw or "").strip().replace(".0",""))  
    if len(s) \== 5:  
        s \= "0" \+ s  
    if s in \_FAKE\_POSTALS:  
        return ""  
    return s

def fix\_postal\_sg(postal) \-\> str:  
    """Normalise to 6-digit string for SF Account Audit preprocessing."""  
    s \= \_norm\_postal\_input(postal)  
    return s if s else str(postal).strip().replace(".0","")  \# fallback: return as-is

def is\_blank\_name(s) \-\> bool:  
    if s is None or (isinstance(s, float) and pd.isna(s)):  
        return True  
    return str(s).strip().lower() in \_NA\_VALUES

def has\_chinese(text: str) \-\> bool:  
    return bool(\_CHINESE\_RE.search(str(text or "")))

def to\_pinyin(text: str) \-\> str:  
    """  
    Converts Chinese characters to Hanyu Pinyin (no tones).  
    Non-Chinese text is passed through as WHOLE SEGMENTS (not char by char)  
    to prevent single-letter tokens causing spurious fuzzy matches.

    Examples:  
        '海底捞'               → 'hai di lao'  
        'ABC 餐厅'             → 'abc can ting'  
        'Master Tang 大堂炒饭'  → 'master tang da tang chao fan'  
    """  
    if not text or is\_blank\_name(text):  
        return ""  
    text \= str(text).strip()  
    if not \_PYPINYIN\_AVAILABLE:  
        return text.lower()  
    result  \= \[\]  
    segment \= \[\]                        \# buffer for consecutive non-Chinese chars  
    for char in text:  
        if \_CHINESE\_RE.match(char):  
            if segment:                 \# flush non-Chinese buffer as one word group  
                result.append("".join(segment))  
                segment \= \[\]  
            result.extend(lazy\_pinyin(char, style=PinyinStyle.NORMAL))  
        else:  
            segment.append(char)  
    if segment:  
        result.append("".join(segment)) \# flush any trailing non-Chinese text  
    return " ".join(result).lower().strip()

def norm\_name\_sg(s, char\_map: dict) \-\> tuple:  
    """Returns (latin\_norm, pinyin\_norm) — strips SG entity noise."""  
    if is\_blank\_name(s):  
        return "", ""  
    s \= str(s).strip()  
    for k, v in char\_map.items():  
        s \= s.replace(k, v)  
    s \= SG\_NAME\_NOISE.sub("", s)  
    s \= SG\_UNIT\_RE.sub("", s)  
    s \= re.sub(r'@\\w+', "", s)  
    s \= re.sub(r'\\s+', " ", s).strip()  
    return s.lower().strip(), (to\_pinyin(s) if has\_chinese(s) else "")

\# ═════════════════════════════════════════════════════════════════  
\# CORE HELPERS  
\# ═════════════════════════════════════════════════════════════════

def get\_onemap\_token():  
    from urllib.request import Request  
    now \= time.time()  
    if \_TOKEN\_CACHE\["token"\] and \_TOKEN\_CACHE\["expiry"\] \> (now \+ 3600):  
        return \_TOKEN\_CACHE\["token"\]  
    try:  
        payload \= json.dumps(ONEMAP\_CREDENTIALS).encode()  
        req \= Request("https://www.onemap.gov.sg/api/auth/post/getToken",  
                      data=payload,  
                      headers={"Content-Type":"application/json",  
                               "User-Agent":"SalesOpsSuite/2.0"},  
                      method="POST")  
        with urlopen(req, timeout=5) as resp:  
            data \= json.loads(resp.read().decode())  
            if "access\_token" in data:  
                \_TOKEN\_CACHE\["token"\]  \= data\["access\_token"\]  
                \_TOKEN\_CACHE\["expiry"\] \= now \+ 259200  
                return data\["access\_token"\]  
    except Exception as e:  
        st.error(f"OneMap token error: {e}")  
    return None

def norm\_name(s, char\_map: dict) \-\> str:  
    if pd.isna(s): return ""  
    s \= str(s).strip()  
    for k, v in char\_map.items():  
        s \= s.replace(k, v)  
    s \= SG\_NAME\_NOISE.sub("", s)  
    return re.sub(r'\\s+', ' ', s).strip().lower()

def name\_confidence(a, b, char\_map: dict) \-\> float:  
    """  
    Name similarity using rapidfuzz token\_sort\_ratio (handles word order).  
    Also compares Hanyu Pinyin when either string contains Chinese.  
    Returns 0.0–1.0.  
    """  
    a\_n \= norm\_name(a, char\_map)  
    b\_n \= norm\_name(b, char\_map)  
    if not a\_n or not b\_n:  
        return 0.0  
    \# token\_sort\_ratio: handles word order (A B C \== C B A)  
    \# token\_set\_ratio:  handles subset matching — critical for CRM names that  
    \#                   append location in brackets, e.g. "Yummy Taste" vs  
    \#                   "Yummy Taste (Rivervale Drive)" → 100% not 55%  
    ts\_score    \= fuzz.token\_sort\_ratio(a\_n, b\_n) / 100.0  
    tset\_score  \= fuzz.token\_set\_ratio(a\_n, b\_n)  / 100.0  
    latin\_score \= max(ts\_score, tset\_score)  
    if has\_chinese(str(a or "")) or has\_chinese(str(b or "")):  
        pin\_a \= to\_pinyin(str(a or ""))  
        pin\_b \= to\_pinyin(str(b or ""))  
        if pin\_a and pin\_b:  
            p\_ts         \= fuzz.token\_sort\_ratio(pin\_a, pin\_b) / 100.0  
            p\_tset       \= fuzz.token\_set\_ratio(pin\_a, pin\_b)  / 100.0  
            pinyin\_score \= max(p\_ts, p\_tset)  
            return round(max(latin\_score, pinyin\_score), 3\)  
    return round(latin\_score, 3\)

def norm\_phone(p, prefix: str) \-\> str:  
    if pd.isna(p): return ""  
    s \= str(p).replace("+","").replace(" ","").replace("-","").strip()  
    if s.endswith(".0"): s \= s\[:-2\]  
    s \= s.lstrip("0")  
    if not s.startswith(prefix): s \= prefix \+ s  
    local \= re.sub(r"\\D","",s)\[len(prefix):\]  
    if len(local) \!= 8 or local\[0\] not in ("6","8","9"): return ""  
    return s

def to\_e164(p, prefix: str) \-\> str:  
    n \= norm\_phone(p, prefix)  
    return "+" \+ n if n else ""

def norm\_url(u) \-\> str:  
    if pd.isna(u): return ""  
    u \= unquote(str(u).strip())  
    u \= u.split("?hl=")\[0\]  
    u \= u.split("\&hl=")\[0\]              \# searchPageUrl appends \&hl=en  
    u \= u.split("\&query\_place\_id=")\[0\]  \# strip Apify-appended place ID  
    return u.lower()

def address\_match(lead\_street, apify\_address, char\_map: dict) \-\> bool:  
    if pd.isna(lead\_street) or pd.isna(apify\_address): return False  
    ls \= norm\_name(str(lead\_street), char\_map)  
    aa \= norm\_name(str(apify\_address), char\_map)  
    pl \= set(re.findall(r"\\b\\d{6}\\b", ls))  
    pa \= set(re.findall(r"\\b\\d{6}\\b", aa))  
    if pl and pa and pl & pa:  
        ul, ua \= extract\_sg\_unit(ls), extract\_sg\_unit(aa)  
        if ul and ua: return ul \== ua  
        return True  
    ignore \= {"","no","sk","singapore","blk","block","新加坡","大厦","路","街"}  
    tl \= set(re.split(r"\\W+",ls)) \- ignore  
    ta \= set(re.split(r"\\W+",aa)) \- ignore  
    return len(tl & ta) \>= 2

def street\_similarity(street\_a, street\_b):  
    def \_n(s):  
        if not s or str(s).strip() in ("","nan"): return ""  
        s \= str(s).upper().strip()  
        s \= re.sub(r'\\b\\d+\[A-Z\]?\\b','',s)  
        s \= re.sub(r'\\b(BLK|BLOCK)\\b','',s)  
        return re.sub(r'\\s+',' ',s).strip()  
    na, nb \= \_n(street\_a), \_n(street\_b)  
    if not na or not nb: return None  
    return SequenceMatcher(None, na, nb).ratio()

def detect\_column(df: pd.DataFrame, candidates: list):  
    for c in candidates:  
        if c in df.columns: return c  
    lmap \= {col.lower(): col for col in df.columns}  
    for c in candidates:  
        if c.lower() in lmap: return lmap\[c.lower()\]  
    return None

def find\_header\_row(file, key\_col="GRID") \-\> int:  
    try:  
        raw \= pd.read\_excel(file, header=None, nrows=30)  
        for i, row in raw.iterrows():  
            if any(str(v).strip() \== key\_col for v in row.dropna()): return i  
    except Exception: pass  
    return 0

\# ═════════════════════════════════════════════════════════════════  
\# CACHED FILE LOADERS  
\# ═════════════════════════════════════════════════════════════════

@st.cache\_data(show\_spinner=False)  
def \_cached\_read(file\_bytes: bytes, filename: str) \-\> pd.DataFrame:  
    """Read raw bytes → DataFrame, cached so re-uploads don't re-parse.  
    Supports .xlsx, .xls (real binary), .xls (HTML Salesforce export), and .csv."""  
    from io import StringIO  
    if filename.endswith(".xlsx"):  
        return pd.read\_excel(io.BytesIO(file\_bytes))  
    if filename.endswith(".xls"):  
        \# Salesforce often exports HTML tables with an .xls extension.  
        \# Detect by checking if the file starts with an HTML tag.  
        header \= file\_bytes\[:512\].lstrip()  
        is\_html \= (header.startswith(b"\<") or  
                   b"\<html" in header.lower() or  
                   b"\<head" in header.lower() or  
                   b"\<table" in header.lower())  
        if is\_html:  
            try:  
                tables \= pd.read\_html(io.BytesIO(file\_bytes))  
                if tables:  
                    return tables\[0\]  
            except Exception:  
                pass  
            \# Last resort: tab-separated text (some Salesforce HTML exports)  
            for enc in ("utf-8", "utf-8-sig", "windows-1252", "latin-1"):  
                try:  
                    raw \= file\_bytes.decode(enc)  
                    sep \= "\\t" if raw.count("\\t") \> raw.count(",") else ","  
                    return pd.read\_csv(StringIO(raw), sep=sep,  
                                       on\_bad\_lines="skip", engine="python")  
                except Exception:  
                    continue  
        else:  
            try:  
                return pd.read\_excel(io.BytesIO(file\_bytes), engine="xlrd")  
            except Exception:  
                pass  
    for enc in ("utf-8","utf-8-sig","windows-1252","latin-1"):  
        try:  
            raw \= file\_bytes.decode(enc)  
            sep \= ";" if raw.count(";") \> raw.count(",") else ","  
            return pd.read\_csv(StringIO(raw), sep=sep, quotechar='"',  
                               on\_bad\_lines="skip", engine="python")  
        except Exception: continue  
    raw \= file\_bytes.decode("latin-1", errors="replace")  
    sep \= ";" if raw.count(";") \> raw.count(",") else ","  
    return pd.read\_csv(StringIO(raw), sep=sep, quotechar='"',  
                       on\_bad\_lines="skip", engine="python")

@st.cache\_data(show\_spinner="Indexing CRM accounts — runs once per session…")  
def \_cached\_preprocess\_crm(df\_raw: pd.DataFrame,  
                             name\_col, postal\_col, addr\_col,  
                             status\_col, char\_map\_tuple: tuple) \-\> pd.DataFrame:  
    """  
    Heavy CRM preprocessing (normalisation, pinyin, closed-name detection).  
    Cached: re-uploading the same file skips all processing.  
    """  
    char\_map \= dict(char\_map\_tuple)  
    df \= df\_raw.copy()  
    if name\_col:  
        df\["\_is\_closed\_name"\] \= df\[name\_col\].apply(has\_closed\_marker)  
        df\["\_name\_latin"\]     \= df\[name\_col\].apply(  
            lambda n: norm\_name\_sg(strip\_venue\_generic(n), char\_map)\[0\])  
        df\["\_name\_pinyin"\]    \= df\[name\_col\].apply(  
            lambda n: norm\_name\_sg(strip\_venue\_generic(n), char\_map)\[1\])  
    else:  
        df\["\_is\_closed\_name"\] \= False  
        df\["\_name\_latin"\]     \= ""  
        df\["\_name\_pinyin"\]    \= ""  
    if postal\_col:  
        df\["\_postal\_fixed"\]   \= df\[postal\_col\].apply(  
            lambda p: extract\_sg\_postal(\_norm\_postal\_input(p))  
                      or fix\_postal\_sg(p))  
    else:  
        df\["\_postal\_fixed"\]   \= ""  
    if addr\_col:  
        df\["\_unit\_extracted"\] \= df\[addr\_col\].apply(extract\_sg\_unit)  
    else:  
        df\["\_unit\_extracted"\] \= ""  
    if status\_col:  
        df\["\_status\_group"\]   \= df.apply(  
            lambda r: get\_status\_group(r\[status\_col\], r\["\_is\_closed\_name"\]), axis=1)  
    else:  
        df\["\_status\_group"\]   \= "active\_pipeline"  
    return df

def load\_leads(file\_bytes: bytes, filename: str, market\_cfg: dict):  
    df \= \_cached\_read(file\_bytes, filename)  
    if filename.endswith((".xlsx", ".xls")):  
        \# Try to detect header row for Excel files with GRID column  
        from io import BytesIO  
        engine \= "xlrd" if filename.endswith(".xls") else None  
        try:  
            raw\_xl \= pd.read\_excel(BytesIO(file\_bytes), header=None, nrows=30,  
                                   \*\*({"engine": engine} if engine else {}))  
            for i, row in raw\_xl.iterrows():  
                if any(str(v).strip() \== "GRID" for v in row.dropna()):  
                    df \= pd.read\_excel(BytesIO(file\_bytes), header=i,  
                                       \*\*({"engine": engine} if engine else {}))  
                    break  
        except Exception: pass  
        grid\_col \= next((c for c in df.columns if str(c).strip() \== "GRID"), None)  
        if grid\_col:  
            df \= df\[df\[grid\_col\].astype(str).str.match(r'^\[A-Z0-9\]{6,}$')\].copy()  
    col\_map \= {}  
    col\_map\["name"\]    \= detect\_column(df, \["Company / Account","Company","Account Name",  
                                             "Name","restaurant\_name","公司名称"\])  
    col\_map\["phone"\]   \= detect\_column(df, \["Phone","phone\_number","Mobile","电话"\])  
    col\_map\["street"\]  \= detect\_column(df, \["Street","Formatted Restaurant Address",  
                                             "restaurant\_address","Address",  
                                             "Block/Street Name","地址"\])  
    col\_map\["city"\]    \= detect\_column(df, \["Area","City","restaurant\_city","city",  
                                             "Restaurant City","城市"\])  
    col\_map\["grid"\]    \= detect\_column(df, \["GRID","grid","Grid"\])  
    col\_map\["lead\_id"\] \= detect\_column(df, \["Lead ID","lead\_id","LeadID","ID"\])  
    col\_map\["url"\]     \= detect\_column(df, \["GOOGLE URL","Google URL","google\_url",  
                                             "URL","Website"\])  
    col\_map\["lat"\]     \= detect\_column(df, \["Coordinates (Latitude)","restaurant\_lat",  
                                             "Latitude","lat","纬度"\])  
    col\_map\["lng"\]     \= detect\_column(df, \["Coordinates (Longitude)","restaurant\_long",  
                                             "Longitude","lng","经度"\])  
    col\_map\["zip"\]     \= detect\_column(df, \["Zip/Postal Code","Restaurant PostalCode",  
                                             "Zip","postal\_code","PostalCode",  
                                             "Postal Code","邮编"\])  
    return df, col\_map

def load\_crm(file\_bytes: bytes, filename: str, market\_cfg: dict):  
    char\_map \= market\_cfg\["char\_map"\]  
    df \= \_cached\_read(file\_bytes, filename)  
    col\_map \= {}  
    col\_map\["grid"\]   \= detect\_column(df, \["GRID\_\_c","GRID","Grid"\])  
    col\_map\["name"\]   \= detect\_column(df, \["Account Name","Name","name"\])  
    col\_map\["phone"\]  \= detect\_column(df, \["Phone","phone"\])  
    col\_map\["status"\] \= detect\_column(df, \["Account\_Status\_\_c","Account Status",  
                                            "AccountStatus"\])  
    col\_map\["reason"\] \= detect\_column(df, \["Status\_Reason\_\_c","Status Reason",  
                                            "StatusReason"\])  
    col\_map\["city"\]   \= detect\_column(df, \["BillingCity","Restaurant City","City"\])  
    col\_map\["postal"\] \= detect\_column(df, \["Restaurant PostalCode","PostalCode",  
                                            "Postal Code","Zip/Postal Code",  
                                            "BillingPostalCode","邮编"\])  
    col\_map\["street"\] \= detect\_column(df, \["Formatted Restaurant Address",  
                                            "BillingStreet","Street",  
                                            "restaurant\_address","Address",  
                                            "Block/Street Name","地址"\])  
    \# Heavy preprocessing — cached  
    df \= \_cached\_preprocess\_crm(  
        df,  
        name\_col   \= col\_map\["name"\],  
        postal\_col \= col\_map\["postal"\],  
        addr\_col   \= col\_map\["street"\],  
        status\_col \= col\_map\["status"\],  
        char\_map\_tuple \= tuple(sorted(char\_map.items())),  
    )  
    return df, col\_map

def load\_apify(file\_bytes: bytes, filename: str):  
    df \= \_cached\_read(file\_bytes, filename)  
    col\_map \= {}  
    col\_map\["grid"\]     \= detect\_column(df, \["GRID", "grid", "Grid"\])  
    col\_map\["title"\]    \= detect\_column(df, \["title"\])  
    col\_map\["url"\]      \= detect\_column(df, \["searchPageUrl", "search\_page\_url",  
                                              "searchUrl", "input\_url"\])  
    col\_map\["gm\_url"\]   \= detect\_column(df, \["url"\])  
    col\_map\["phone"\]    \= detect\_column(df, \["phone"\])  
    col\_map\["website"\]  \= detect\_column(df, \["website"\])  
    col\_map\["category"\] \= detect\_column(df, \["categoryName", "category\_name",  
                                              "category", "categories/0"\])  
    col\_map\["perm"\]     \= detect\_column(df, \["permanentlyClosed", "permanently\_closed",  
                                              "isClosed", "is\_closed"\])  
    col\_map\["temp"\]     \= detect\_column(df, \["temporarilyClosed", "temporarily\_closed",  
                                              "isTemporarilyClosed"\])  
    col\_map\["address"\]  \= detect\_column(df, \["address"\])  
    col\_map\["lat"\]      \= detect\_column(df, \["latitude","lat"\])  
    col\_map\["lng"\]      \= detect\_column(df, \["longitude","lng"\])  
    if col\_map\["url"\] and df\[col\_map\["url"\]\].notna().sum() \> 0:  
        df\["\_url\_norm"\] \= df\[col\_map\["url"\]\].apply(norm\_url)  
    else:  
        ss \= detect\_column(df, \["searchString"\])  
        if ss and df\[ss\].notna().sum() \> 0:  
            def \_ex(s):  
                if pd.isna(s): return ""  
                s \= str(s).strip()  
                return norm\_url(s.replace("Direct Detail URL:","").strip()  
                                if "Direct Detail URL:" in s else s)  
            df\["\_url\_norm"\] \= df\[ss\].apply(\_ex)  
            col\_map\["url"\] \= ss  
        elif col\_map\["gm\_url"\] and df\[col\_map\["gm\_url"\]\].notna().sum() \> 0:  
            df\["\_url\_norm"\] \= df\[col\_map\["gm\_url"\]\].apply(norm\_url)  
            col\_map\["url"\] \= col\_map\["gm\_url"\]  
    return df, col\_map

\# ═════════════════════════════════════════════════════════════════  
\# URL GENERATOR  
\# ═════════════════════════════════════════════════════════════════

def generate\_google\_urls(leads\_df, col\_map, market\_cfg, mode="coords"):  
    """  
    mode="text"   → URL uses Company / Account \+ Street \+ Postal as query text  
    mode="coords" → URL uses Coordinates (Latitude, Longitude) as query  
    Falls back to the other mode if the preferred data is missing.  
    """  
    name\_col   \= col\_map.get("name")  
    street\_col \= col\_map.get("street")  
    zip\_col    \= col\_map.get("zip")  
    lat\_col    \= col\_map.get("lat")  
    lng\_col    \= col\_map.get("lng")  
    url\_col    \= col\_map.get("url")  
    suffix     \= market\_cfg.get("country\_suffix","")

    urls, reused \= \[\], 0  
    for \_, row in leads\_df.iterrows():  
        name   \= str(row\[name\_col\]).strip()   if name\_col   and pd.notna(row.get(name\_col))   else ""  
        street \= str(row\[street\_col\]).strip() if street\_col and pd.notna(row.get(street\_col)) else ""  
        postal \= str(row\[zip\_col\]).strip()    if zip\_col    and pd.notna(row.get(zip\_col))    else ""  
        lat    \= row.get(lat\_col) if lat\_col else None  
        lng    \= row.get(lng\_col) if lng\_col else None  
        try:  
            lat \= float(lat) if lat is not None and str(lat) not in ("","nan") else None  
            lng \= float(lng) if lng is not None and str(lng) not in ("","nan") else None  
        except (ValueError, TypeError):  
            lat \= lng \= None

        \# Reuse existing URL if present  
        existing \= str(row.get(url\_col,"") or "") if url\_col else ""  
        if existing not in ("","nan") and "google.com/maps" in existing:  
            urls.append(existing.strip()); reused \+= 1; continue

        url \= ""

        if mode \== "text":  
            \# Option 1: Company / Account \+ Street \+ Postal  
            parts \= \[p for p in \[name, street, postal\] if p and p.lower() \!= "nan"\]  
            if parts:  
                q \= " ".join(parts)  
                url \= f"https://www.google.com/maps/search/?api=1\&query={quote(q)}"  
            elif lat and lng:  
                \# Fallback to coords if no text available  
                url \= f"https://www.google.com/maps/search/?api=1\&query={lat},{lng}"

        else:  
            \# Option 2: Company / Account \+ Coordinates (Latitude, Longitude)  
            if lat and lng:  
                if name:  
                    url \= f"https://www.google.com/maps/search/?api=1\&query={quote(f'{name},{lat},{lng}')}"  
                else:  
                    url \= f"https://www.google.com/maps/search/?api=1\&query={lat},{lng}"  
            else:  
                \# Fallback to text if no coords available  
                parts \= \[p for p in \[name, street, postal\] if p and p.lower() \!= "nan"\]  
                if parts:  
                    q \= " ".join(parts)  
                    url \= f"https://www.google.com/maps/search/?api=1\&query={quote(q)}"  
                elif name:  
                    url \= f"https://www.google.com/maps/search/?api=1\&query={quote(f'{name} {suffix}')}"

        urls.append(url)  
    return urls, reused

\# ═════════════════════════════════════════════════════════════════  
\# GEO / ZONE  
\# ═════════════════════════════════════════════════════════════════

def geocode\_address(street, postal\_code, country\_suffix, cache={}):  
    key \= f"{street}|{postal\_code}|{country\_suffix}"  
    if key in cache: return cache\[key\]  
    from urllib.request import Request  
    street\_c  \= str(street).strip()      if street      and str(street).strip()      not in ("","nan") else ""  
    postal\_c  \= str(postal\_code).strip() if postal\_code and str(postal\_code).strip() not in ("","nan") else ""  
    sg\_postal \= re.search(r'\\b\\d{6}\\b', f"{street\_c} {postal\_c}")  
    if sg\_postal:  
        token \= get\_onemap\_token()  
        if token:  
            try:  
                om \= (f"https://www.onemap.gov.sg/api/common/elastic/search"  
                      f"?searchVal={sg\_postal.group(0)}\&returnGeom=Y\&getAddrDetails=N\&pageNum=1")  
                req \= Request(om, headers={"Authorization": token,  
                                           "User-Agent": "SalesOpsSuite/2.0"})  
                with urlopen(req, timeout=5) as r:  
                    res \= json.loads(r.read().decode()).get("results",\[\])  
                    if res:  
                        lat, lng \= float(res\[0\]\["LATITUDE"\]), float(res\[0\]\["LONGITUDE"\])  
                        cache\[key\] \= (lat, lng); return (lat, lng)  
            except Exception: pass  
    parts \= \[p for p in \[street\_c, postal\_c, "Singapore"\] if p.strip()\]  
    if not parts: cache\[key\] \= (None, None); return (None, None)  
    time.sleep(1.1)  
    try:  
        req \= urlopen(  
            f"https://photon.komoot.io/api/?q={urlencode({'q':', '.join(parts)})}\&limit=1",  
            timeout=6)  
        feats \= json.loads(req.read().decode()).get("features",\[\])  
        if feats:  
            c \= feats\[0\]\["geometry"\]\["coordinates"\]  
            cache\[key\] \= (float(c\[1\]), float(c\[0\])); return cache\[key\]  
    except Exception: pass  
    cache\[key\] \= (None, None); return (None, None)

def load\_zones(file\_bytes=None, filename=None, market\_code=None):  
    import os  
    def \_parse(df):  
        zones, wc \= \[\], next((c for c in df.columns if "wkt" in c.lower()), None)  
        zn \= next((c for c in df.columns if "zone\_name" in c.lower()), None)  
        if not wc: return zones  
        for \_, r in df.iterrows():  
            w \= str(r.get(wc,"")).strip()  
            if not w or w.lower() \== "nan": continue  
            try: zones.append({"polygon": wkt\_loads(w),  
                                "zone\_name": str(r.get(zn,"")) if zn else "",  
                                "city\_name": ""})  
            except Exception: continue  
        return zones  
    if file\_bytes:  
        try:  
            df \= \_cached\_read(file\_bytes, filename or "zones.csv")  
            return \_parse(df)  
        except Exception: return \[\]  
    if market\_code \== "SG":  
        p \= os.path.join(os.path.dirname(os.path.abspath(\_\_file\_\_)),  
                         f"zones\_{market\_code}.json")  
        if os.path.exists(p):  
            try:  
                raw \= json.load(open(p, encoding="utf-8"))  
                return \[{"polygon": wkt\_loads(z\["wkt"\]),  
                         "zone\_name": z.get("zone\_name",""),  
                         "city\_name": z.get("city\_name","")}  
                        for z in raw if "wkt" in z\]  
            except Exception: return \[\]  
    return \[\]

def point\_in\_zones(lat, lng, zones):  
    if lat is None or lng is None: return None, None  
    try:  
        pt \= Point(float(lng), float(lat))  
        for z in zones:  
            if z\["polygon"\].contains(pt):  
                return z\["zone\_name"\], z\["city\_name"\]  
    except Exception: pass  
    return None, None

def check\_delivery\_zone(row, col\_map, zones, suffix, geocode\_enabled):  
    if not zones: return "No Zone Data","","",""  
    lat\_c, lng\_c \= col\_map.get("lat"), col\_map.get("lng")  
    str\_c, zip\_c \= col\_map.get("street"), col\_map.get("zip")  
    lat \= row.get(lat\_c) if lat\_c else None  
    lng \= row.get(lng\_c) if lng\_c else None  
    try:  
        lat \= float(lat) if lat is not None and not pd.isna(lat) else None  
        lng \= float(lng) if lng is not None and not pd.isna(lng) else None  
    except (ValueError, TypeError):  
        lat \= lng \= None  
    if lat and lng:  
        zn, zc \= point\_in\_zones(lat, lng, zones)  
        return (("Within Zone", zn, zc, "Coordinates")  
                if zn else ("Outside Zone","","","Coordinates"))  
    if geocode\_enabled:  
        street \= str(row.get(str\_c,"") or "") if str\_c else ""  
        postal \= str(row.get(zip\_c,"") or "") if zip\_c else ""  
        if street.strip() or postal.strip():  
            lat, lng \= geocode\_address(street, postal, suffix)  
            if lat:  
                zn, zc \= point\_in\_zones(lat, lng, zones)  
                return (("Within Zone", zn, zc, "Geocoded")  
                        if zn else ("Outside Zone","","","Geocoded"))  
            return "Geocoding Failed","","","Geocoded"  
    return "Outside Zone","","","No coordinates"

\# ═════════════════════════════════════════════════════════════════  
\# CLASSIFY LEADS — MAIN ENGINE  
\# ═════════════════════════════════════════════════════════════════

def classify\_leads(leads\_df, col\_map\_leads, crm\_df, col\_map\_crm,  
                   apify\_df, col\_map\_apify, market\_cfg,  
                   p2\_threshold=0.50, p3\_threshold=0.75,  
                   exclusion\_kw=None, zones=None, geocode\_enabled=True,  
                   progress\_cb=None):

    char\_map    \= market\_cfg\["char\_map"\]  
    prefix      \= market\_cfg\["phone\_prefix"\]  
    suffix      \= market\_cfg.get("country\_suffix","")  
    excl\_kw     \= exclusion\_kw or \_DEFAULT\_EXCLUSION\_KW

    \# ── Build CRM indexes ─────────────────────────────────────────  
    \# Each item stored as (raw\_name, lat\_norm, pin\_norm, row) so that  
    \# name normalisation is done ONCE here, not on every comparison.  
    \# crm\_postal\_all\_dict gives O(1) pool lookup per lead postal.  
    crm\_postal\_unit\_dict \= {}   \# "postal|unit" \-\> \[item, ...\]  
    crm\_postal\_all\_dict  \= {}   \# postal \-\> \[item, ...\] (ALL at that postal)  
    all\_crm\_items        \= \[\]   \# flat list for zero-postal fallback

    if crm\_df is not None:  
        postal\_c \= col\_map\_crm.get("postal")  
        addr\_c   \= col\_map\_crm.get("street")  
        name\_c\_  \= col\_map\_crm.get("name")

        for \_, r in crm\_df.iterrows():  
            raw\_n         \= str(r.get(name\_c\_,"") or "") if name\_c\_ else ""  
            lat\_n, pin\_n  \= norm\_name\_sg(strip\_venue\_generic(raw\_n), char\_map)  
            item          \= (raw\_n, lat\_n, pin\_n, r)  
            all\_crm\_items.append(item)

            postal\_raw \= \_norm\_postal\_input(r.get(postal\_c,"") if postal\_c else "")  
            addr\_raw   \= str(r.get(addr\_c,  "") or "") if addr\_c   else ""  
            crm\_postal \= extract\_sg\_postal(postal\_raw) or extract\_sg\_postal(addr\_raw)  
            crm\_unit   \= extract\_sg\_unit(addr\_raw)

            if crm\_postal:  
                crm\_postal\_all\_dict.setdefault(crm\_postal, \[\]).append(item)  
                if crm\_unit:  
                    crm\_postal\_unit\_dict.setdefault(  
                        f"{crm\_postal}|{crm\_unit}", \[\]).append(item)

    \# ── Build Apify GRID index ─────────────────────────────────────  
    apify\_dict \= {}  
    if apify\_df is not None:  
        gc \= detect\_column(apify\_df, \["GRID", "grid", "Grid"\])  
        if gc:  
            for \_, r in apify\_df.iterrows():  
                g \= str(r.get(gc,"") or "").strip()  
                if g and g not in apify\_dict:  
                    apify\_dict\[g\] \= r

    \# ── Column aliases ─────────────────────────────────────────────  
    name\_col\_l   \= col\_map\_leads.get("name")  
    phone\_col\_l  \= col\_map\_leads.get("phone")  
    street\_col\_l \= col\_map\_leads.get("street")  
    lat\_col\_l    \= col\_map\_leads.get("lat")  
    lng\_col\_l    \= col\_map\_leads.get("lng")  
    grid\_col\_l   \= col\_map\_leads.get("grid")  
    lead\_id\_col  \= col\_map\_leads.get("lead\_id")  
    zip\_col\_l    \= col\_map\_leads.get("zip")  
    reason\_col\_c \= col\_map\_crm.get("reason") if col\_map\_crm else None

    results \= \[\]  
    \_crm\_name\_col   \= col\_map\_crm.get("name","")   if col\_map\_crm else ""  
    \_crm\_grid\_col   \= col\_map\_crm.get("grid","")   if col\_map\_crm else ""  
    \_crm\_status\_col \= col\_map\_crm.get("status","") if col\_map\_crm else ""  
    \_crm\_street\_col \= col\_map\_crm.get("street","") if col\_map\_crm else ""  
    \_crm\_postal\_col \= col\_map\_crm.get("postal","") if col\_map\_crm else ""  
    \_n\_total        \= len(leads\_df)

    \# ── Scoring helper (defined ONCE — not recreated per lead) ────  
    def \_score(lead\_lat\_n, lead\_pin\_n, item):  
        """Score a pre-normalised CRM item against a lead.  
        Returns (score, crm\_raw, crm\_row).  
        Uses pre-computed lat\_n / pin\_n — avoids re-normalising on  
        every comparison (critical for 100k+ CRM accounts)."""  
        raw\_n, lat\_n, pin\_n, r \= item  
        ns \= 0.0  
        if lead\_lat\_n and lat\_n:  
            ns \= max(fuzz.token\_sort\_ratio(lead\_lat\_n, lat\_n),  
                     fuzz.token\_set\_ratio(lead\_lat\_n,  lat\_n)) / 100.0  
        if lead\_pin\_n or pin\_n:  
            lp, cp \= lead\_pin\_n or "", pin\_n or ""  
            if lp and cp:  
                ps \= max(fuzz.token\_sort\_ratio(lp, cp),  
                         fuzz.token\_set\_ratio(lp,  cp)) / 100.0  
                return round(max(ns, ps), 3), raw\_n, r  
        return round(ns, 3), raw\_n, r

    def \_pool(postal):  
        """O(1) pool lookup — returns all CRM items at this postal."""  
        return crm\_postal\_all\_dict.get(postal, \[\])

    for \_lead\_i, (\_, row) in enumerate(leads\_df.iterrows()):

        \# ── Extract lead fields ────────────────────────────────────  
        lead\_name\_raw \= row.get(name\_col\_l,"")   if name\_col\_l   else ""  
        lead\_street   \= str(row.get(street\_col\_l,"") or "") if street\_col\_l else ""  
        lead\_zip      \= str(row.get(zip\_col\_l,  "") or "") if zip\_col\_l    else ""

        lead\_lat \= row.get(lat\_col\_l) if lat\_col\_l else None  
        lead\_lng \= row.get(lng\_col\_l) if lng\_col\_l else None  
        try:  
            lead\_lat \= float(lead\_lat) if lead\_lat is not None and not pd.isna(lead\_lat) else None  
            lead\_lng \= float(lead\_lng) if lead\_lng is not None and not pd.isna(lead\_lng) else None  
        except (ValueError, TypeError):  
            lead\_lat \= lead\_lng \= None

        lead\_postal\_raw \= \_norm\_postal\_input(row.get(zip\_col\_l,"") if zip\_col\_l else "")  
        lead\_postal \= (extract\_sg\_postal(lead\_postal\_raw)  
                       or extract\_sg\_postal(lead\_street))  
        lead\_unit   \= extract\_sg\_unit(lead\_street)  
        lead\_lat\_n, lead\_pin\_n \= norm\_name\_sg(strip\_venue\_generic(lead\_name\_raw), char\_map)  
        lead\_blank  \= is\_blank\_name(lead\_name\_raw)  
        lead\_grid   \= str(row.get(grid\_col\_l,"") or "").strip() if grid\_col\_l else ""

        \# ── Initialise ─────────────────────────────────────────────  
        crm\_match          \= None  
        match\_method       \= ""  
        match\_score        \= 0.0  
        label              \= ""  
        dup\_grid \= dup\_name \= dup\_crm\_status \= dup\_reason \= dup\_method \= ""  
        dup\_address \= dup\_postal \= ""  
        prev\_occupant\_grid \= ""  
        prev\_occupant\_name \= ""

        \# ── Pre-flight: exclude Mix & Match immediately ────────────  
        if not lead\_blank and "mix & match" in str(lead\_name\_raw).lower():  
            label        \= "Wrong Target Group"  
            match\_method \= "Lead name contains 'Mix & Match' — excluded"

        \# ══════════════════════════════════════════════════════════  
        \# CRM DEDUP — Postal \+ Name (unit used as filter)  
        \#  
        \# Logic:  
        \#   1\. If lead has a unit AND CRM has an exact unit match →  
        \#      score THAT record. If name \< p2, it's a new tenant.  
        \#   2\. Otherwise → score all CRM accounts at same postal,  
        \#      take best name match.  
        \#   score ≥ p3\_threshold (0.75) → P4 Duplicate  
        \#   score  p2–p3 (0.50–0.74)  → P3 Potential Match  
        \#   score \< p2\_threshold (0.50) → no CRM match → Apify  
        \# ══════════════════════════════════════════════════════════  
        if not label and lead\_postal and not lead\_blank:

            if lead\_unit:  
                unit\_matches \= crm\_postal\_unit\_dict.get(f"{lead\_postal}|{lead\_unit}", \[\])

                if unit\_matches:  
                    best\_sc, best\_cand\_row, best\_raw \= 0.0, None, ""  
                    for item in unit\_matches:  
                        sc, crm\_raw, cand\_row \= \_score(lead\_lat\_n, lead\_pin\_n, item)  
                        if sc \> best\_sc:  
                            best\_sc, best\_cand\_row, best\_raw \= sc, cand\_row, crm\_raw

                    if best\_sc \>= p2\_threshold:  
                        crm\_match    \= best\_cand\_row  
                        match\_score  \= best\_sc  
                        match\_method \= (f"Postal+Unit+Name \[{lead\_postal} \#{lead\_unit}\] "  
                                        f"score={best\_sc:.2f}")  
                    else:  
                        prev\_occupant\_name \= best\_raw  
                        prev\_occupant\_grid \= str(  
                            best\_cand\_row.get(\_crm\_grid\_col,"") or "") \\  
                            if best\_cand\_row is not None else ""  
                        match\_method \= (f"New business at known address "  
                                        f"\[{lead\_postal} \#{lead\_unit}\] "  
                                        f"prev='{best\_raw}' score={best\_sc:.2f}")  
                else:  
                    \# No CRM record at lead unit → search all at same postal  
                    \# 0.90 minimum since unit is unconfirmed  
                    no\_unit\_min \= max(p2\_threshold, 0.90)  
                    pool \= \_pool(lead\_postal)  
                    best\_sc, best\_cand\_row, best\_raw \= 0.0, None, ""  
                    for item in pool:  
                        sc, crm\_raw, cand\_row \= \_score(lead\_lat\_n, lead\_pin\_n, item)  
                        if sc \> best\_sc:  
                            best\_sc, best\_cand\_row, best\_raw \= sc, cand\_row, crm\_raw  
                    if best\_cand\_row is not None and best\_sc \>= no\_unit\_min:  
                        crm\_match    \= best\_cand\_row  
                        match\_score  \= best\_sc  
                        match\_method \= (f"Postal+Name \[{lead\_postal}\] "  
                                        f"score={best\_sc:.2f}")  
            else:  
                \# Lead has no unit → postal \+ name only, 0.90 minimum  
                no\_unit\_min \= max(p2\_threshold, 0.90)  
                pool \= \_pool(lead\_postal)  
                best\_sc, best\_cand\_row, best\_raw \= 0.0, None, ""  
                for item in pool:  
                    sc, crm\_raw, cand\_row \= \_score(lead\_lat\_n, lead\_pin\_n, item)  
                    if sc \> best\_sc:  
                        best\_sc, best\_cand\_row, best\_raw \= sc, cand\_row, crm\_raw  
                if best\_cand\_row is not None and best\_sc \>= no\_unit\_min:  
                    crm\_match    \= best\_cand\_row  
                    match\_score  \= best\_sc  
                    match\_method \= (f"Postal+Name \[{lead\_postal}\] "  
                                    f"score={best\_sc:.2f} (no unit)")

        \# ── Assign CRM-based label ─────────────────────────────────  
        if crm\_match is not None:  
            label          \= ("P4 — Duplicate"      if match\_score \>= p3\_threshold  
                              else "P3 — Potential Match")  
            dup\_grid       \= str(crm\_match.get(\_crm\_grid\_col,"") or "")  
            dup\_name       \= str(crm\_match.get(\_crm\_name\_col,"") or "")  
            dup\_crm\_status \= str(crm\_match.get(\_crm\_status\_col,"") or "")  
            dup\_reason     \= str(crm\_match.get(reason\_col\_c,"") or "") if reason\_col\_c else ""  
            dup\_method     \= match\_method  
            dup\_address    \= str(crm\_match.get(\_crm\_street\_col,"") or "")  
            dup\_postal     \= str(crm\_match.get(\_crm\_postal\_col,"") or "")

        \# ══════════════════════════════════════════════════════════  
        \# ZERO-POSTAL FALLBACK — Name-only scan (full CRM)  
        \#  
        \# Triggered when the lead has no valid SG postal code  
        \# (000000 / 00000 / blank) — meaning the address data is  
        \# unreliable and the postal cascade produced no match.  
        \#  
        \# Rules:  
        \#   • Score ≥ 0.90 (hard minimum, not adjustable from sidebar)  
        \#   • Geographic conflict check: if CRM has an Area: line and  
        \#     that area conflicts with a known area in the lead text,  
        \#     reject the candidate.  
        \#   • Result is always P3 (no postal confirmation available).  
        \# ══════════════════════════════════════════════════════════  
        if crm\_match is None and not label and not lead\_blank and not lead\_postal:

            def \_extract\_crm\_area(addr\_raw: str) \-\> str:  
                m \= re.search(r'Area:\\s\*(.+)', str(addr\_raw), re.IGNORECASE)  
                return m.group(1).strip().lower() if m else ""

            def \_geo\_conflict(lead\_text: str, crm\_area: str) \-\> bool:  
                if not crm\_area: return False  
                lead\_low \= lead\_text.lower()  
                if crm\_area in lead\_low: return False  
                return any(area in lead\_low for area in SG\_AREAS if area \!= crm\_area)

            lead\_text\_for\_geo   \= f"{lead\_name\_raw} {lead\_street}".lower()  
            NAME\_ONLY\_THRESHOLD \= 0.90  
            best\_sc, best\_cand  \= 0.0, None

            for item in all\_crm\_items:  
                sc, \_, cand\_row \= \_score(lead\_lat\_n, lead\_pin\_n, item)  
                if sc \< NAME\_ONLY\_THRESHOLD or sc \<= best\_sc:  
                    continue  
                addr\_raw\_c \= str(cand\_row.get(\_crm\_street\_col,"") or "")  
                if \_geo\_conflict(lead\_text\_for\_geo, \_extract\_crm\_area(addr\_raw\_c)):  
                    continue  
                best\_sc, best\_cand \= sc, cand\_row

            if best\_cand is not None:  
                crm\_match      \= best\_cand  
                match\_score    \= best\_sc  
                match\_method   \= f"Name-only (zero postal) score={best\_sc:.2f}"  
                label          \= "P3 — Potential Match"  
                dup\_grid       \= str(crm\_match.get(\_crm\_grid\_col,"") or "")  
                dup\_name       \= str(crm\_match.get(\_crm\_name\_col,"") or "")  
                dup\_crm\_status \= str(crm\_match.get(\_crm\_status\_col,"") or "")  
                dup\_reason     \= str(crm\_match.get(reason\_col\_c,"") or "") if reason\_col\_c else ""  
                dup\_method     \= match\_method  
                dup\_address    \= str(crm\_match.get(\_crm\_street\_col,"") or "")  
                dup\_postal     \= str(crm\_match.get(\_crm\_postal\_col,"") or "")

        \# ══════════════════════════════════════════════════════════  
        \# APIFY VALIDATION — GRID exact lookup  
        \# ══════════════════════════════════════════════════════════  
        apy \= apify\_dict.get(lead\_grid) if lead\_grid else None  
        gm\_title \= gm\_cat \= gm\_biz\_status \= gm\_phone \= gm\_website \= gm\_url \= ""  
        match\_conf \= 0.0; match\_reason \= ""

        if apy is not None:  
            tc  \= col\_map\_apify.get("title");    gm\_title   \= str(apy\[tc\] or "")  if tc  and pd.notna(apy.get(tc))  else ""  
            cc  \= col\_map\_apify.get("category"); gm\_cat     \= str(apy\[cc\] or "")  if cc  and pd.notna(apy.get(cc))  else ""  
            pc  \= col\_map\_apify.get("phone");    gm\_phone   \= to\_e164(apy.get(pc,""), prefix) if pc else ""  
            wc  \= col\_map\_apify.get("website");  gm\_website \= str(apy\[wc\] or "")  if wc  and pd.notna(apy.get(wc))  else ""  
            uc  \= col\_map\_apify.get("gm\_url");   gm\_url     \= str(apy\[uc\] or "")  if uc  and pd.notna(apy.get(uc))  else ""  
            prc \= col\_map\_apify.get("perm");     perm \= str(apy.get(prc,"")).lower() \== "true" if prc else False  
            tmc \= col\_map\_apify.get("temp");     temp \= str(apy.get(tmc,"")).lower() \== "true" if tmc else False  
            gm\_biz\_status \= ("Permanently Closed" if perm else  
                             "Temporarily Closed" if temp else "Open")

            gm\_phone\_norm \= norm\_phone(apy.get(pc,""), prefix) if pc else ""  
            name\_sc       \= name\_confidence(lead\_name\_raw, gm\_title, char\_map)  
            addr\_col\_a    \= col\_map\_apify.get("address","")  
            addr\_hit      \= address\_match(lead\_street, apy.get(addr\_col\_a,""), char\_map)

            match\_conf \= round((min(name\_sc, 1.0) \* 0.6)  
                               \+ (0.4 if addr\_hit else 0.0), 3\)  
            \# GRID guarantees we have the right Apify row — confirmed by definition.  
            confirmed  \= True

            reasons \= \[\]  
            reasons.append(f"Name {name\_sc:.2f}" \+ (" ✓" if name\_sc \>= 0.5 else " ✗"))  
            if addr\_hit: reasons.append("Address ✓")  
            reasons.append("GRID match ✓")  
            match\_reason \= " | ".join(reasons)

            \# Apify sets label only when CRM cascade found no match.  
            \# Two safety checks prevent wrong Apify results from mislabelling:  
            \#   1\. If Apify says closed BUT address doesn't match the lead →  
            \#      likely a different branch/location → P2  
            \#   2\. If category is non-food AND name barely matches the lead →  
            \#      Apify returned the wrong business entirely → P2  
            if not label:  
                wrong\_location   \= (perm or temp) and not addr\_hit  
                wrong\_business   \= (not is\_food\_delivery\_eligible(gm\_cat, excl\_kw)  
                                    and name\_sc \< 0.2)  
                if wrong\_location or wrong\_business:  
                    label \= "P2 — Please Check"  
                elif not gm\_cat:                                      label \= "P2 — Please Check"  
                elif perm or temp:                                    label \= "Business Closed"  
                elif not is\_food\_delivery\_eligible(gm\_cat, excl\_kw): label \= "Wrong Target Group"  
                else:                                                 label \= "P1 — New"  
        else:  
            gm\_biz\_status \= "Not Found on Google"  
            match\_reason  \= "No Apify result"  
            if not label: label \= "P2 — Please Check"

        \# ── Zone check ─────────────────────────────────────────────  
        zone\_status \= zone\_name \= zone\_city \= zone\_method \= ""  
        if zones:  
            zone\_status, zone\_name, zone\_city, zone\_method \= check\_delivery\_zone(  
                row, col\_map\_leads, zones, suffix, geocode\_enabled)

        results.append({  
            "GRID":                   row.get(grid\_col\_l,"")   if grid\_col\_l  else "",  
            "Lead ID":                row.get(lead\_id\_col,"")  if lead\_id\_col else "",  
            "Company / Account":      lead\_name\_raw,  
            "City":                   "SINGAPORE",  
            "Street":                 lead\_street,  
            "Phone":                  to\_e164(row.get(phone\_col\_l,""), prefix) if phone\_col\_l else "",  
            "GM Title":               gm\_title,  
            "GM Category":            gm\_cat,  
            "GM Business Status":     gm\_biz\_status,  
            "GM Phone":               gm\_phone,  
            "GM Website":             gm\_website,  
            "GM URL":                 gm\_url,  
            "Match Confidence":       match\_conf,  
            "Match Reason":           match\_reason,  
            "Label":                  label,  
            "Match Score (CRM)":      round(match\_score, 3),  
            "Duplicate GRID":         dup\_grid,  
            "Duplicate CRM Name":     dup\_name,  
            "Duplicate CRM Address":  dup\_address,  
            "Duplicate CRM Postal":   dup\_postal,  
            "CRM Account Status":     dup\_crm\_status,  
            "CRM Status Reason":      dup\_reason,  
            "Duplicate Match Method": dup\_method,  
            "Previous Occupant GRID": prev\_occupant\_grid,  
            "Previous Occupant Name": prev\_occupant\_name,  
            "Delivery Zone Status":   zone\_status,  
            "Zone Name":              zone\_name,  
            "Zone City":              zone\_city,  
            "Zone Method":            zone\_method,  
        })

        if progress\_cb and (\_lead\_i % 10 \== 0 or \_lead\_i \== \_n\_total \- 1):  
            progress\_cb(\_lead\_i \+ 1, \_n\_total)

    return pd.DataFrame(results)

\# ═════════════════════════════════════════════════════════════════  
\# SF ACCOUNT AUDIT  (finds duplicates WITHIN Salesforce)  
\# ═════════════════════════════════════════════════════════════════

def find\_sf\_duplicates(df, name\_col, addr\_col, status\_col, grid\_col, id\_col, threshold):  
    """  
    Find suspected duplicate account pairs within the SF master.  
    Trigger: same 6-digit postal code \+ same extracted unit number  
             \+ name similarity ≥ threshold.  
    """  
    pairs \= \[\]  
    for postal, pg in df.groupby("\_postal\_fixed"):  
        with\_unit \= pg\[pg\["\_unit\_extracted"\] \!= ""\]  
        if with\_unit.empty: continue  
        for unit, ug in with\_unit.groupby("\_unit\_extracted"):  
            if len(ug) \< 2: continue  
            rows \= ug.reset\_index(drop=True)  
            for i, j in combinations(range(len(rows)), 2):  
                a, b   \= rows.iloc\[i\], rows.iloc\[j\]  
                score  \= fuzz.token\_sort\_ratio(a\["\_name\_latin"\], b\["\_name\_latin"\])  
                if score \< threshold: continue  
                risk, action \= get\_risk\_and\_action(a\[status\_col\], b\[status\_col\])  
                pairs.append({  
                    "ACCT\_A\_SF\_ID":       str(a\[id\_col\]),  
                    "ACCT\_A\_NAME":        str(a\[name\_col\]),  
                    "ACCT\_A\_STATUS":      str(a\[status\_col\]),  
                    "ACCT\_A\_ADDRESS":     str(a\[addr\_col\]) if addr\_col else "",  
                    "ACCT\_A\_POSTAL":      str(a\["\_postal\_fixed"\]),  
                    "ACCT\_A\_GRID":        str(a\[grid\_col\]) if grid\_col else "",  
                    "ACCT\_B\_SF\_ID":       str(b\[id\_col\]),  
                    "ACCT\_B\_NAME":        str(b\[name\_col\]),  
                    "ACCT\_B\_STATUS":      str(b\[status\_col\]),  
                    "ACCT\_B\_ADDRESS":     str(b\[addr\_col\]) if addr\_col else "",  
                    "ACCT\_B\_POSTAL":      str(b\["\_postal\_fixed"\]),  
                    "ACCT\_B\_GRID":        str(b\[grid\_col\]) if grid\_col else "",  
                    "SHARED\_UNIT":        unit,  
                    "NAME\_SCORE":         f"{score}%",  
                    "RISK\_LEVEL":         risk,  
                    "RECOMMENDED\_ACTION": action,  
                })  
    return pairs

\# ═════════════════════════════════════════════════════════════════  
\# EXCEL REPORT BUILDER  
\# ═════════════════════════════════════════════════════════════════

\# ─────────────────────────────────────────────────────────────────────  
\# CRM CHECK — standalone duplicate check (no GRID, no Apify needed)  
\# ─────────────────────────────────────────────────────────────────────

def crm\_check\_classify(rest\_df, rest\_cols, crm\_df, col\_map\_crm,  
                        char\_map, p2\_threshold, p3\_threshold):  
    """  
    CRM-only duplicate check for raw restaurant lists.  
    Input: name \+ street \+ postal only. No GRID or Apify needed.  
    Labels: P4 — Duplicate / P3 — Potential Match / Unverified  
    """  
    \# ── Build CRM indexes (pre-compute names once) ────────────────  
    crm\_postal\_unit\_dict \= {}   \# "postal|unit" \-\> \[item, ...\]  
    crm\_postal\_all\_dict  \= {}   \# postal \-\> \[item, ...\]  
    all\_crm\_items        \= \[\]

    crm\_name\_c \= col\_map\_crm.get("name")   if col\_map\_crm else None  
    crm\_grid\_c \= col\_map\_crm.get("grid")   if col\_map\_crm else None  
    crm\_stat\_c \= col\_map\_crm.get("status") if col\_map\_crm else None  
    crm\_reas\_c \= col\_map\_crm.get("reason") if col\_map\_crm else None  
    crm\_addr\_c \= col\_map\_crm.get("street") if col\_map\_crm else None

    if crm\_df is not None:  
        postal\_c \= col\_map\_crm.get("postal") if col\_map\_crm else None  
        addr\_c   \= crm\_addr\_c  
        for \_, r in crm\_df.iterrows():  
            raw\_n        \= str(r.get(crm\_name\_c,"") or "") if crm\_name\_c else ""  
            lat\_n, pin\_n \= norm\_name\_sg(strip\_venue\_generic(raw\_n), char\_map)  
            item         \= (raw\_n, lat\_n, pin\_n, r)  
            all\_crm\_items.append(item)  
            postal\_raw \= \_norm\_postal\_input(r.get(postal\_c,"") if postal\_c else "")  
            addr\_raw   \= str(r.get(addr\_c,  "") or "") if addr\_c   else ""  
            crm\_postal \= extract\_sg\_postal(postal\_raw) or extract\_sg\_postal(addr\_raw)  
            crm\_unit   \= extract\_sg\_unit(addr\_raw)  
            if crm\_postal:  
                crm\_postal\_all\_dict.setdefault(crm\_postal, \[\]).append(item)  
                if crm\_unit:  
                    crm\_postal\_unit\_dict.setdefault(  
                        f"{crm\_postal}|{crm\_unit}", \[\]).append(item)

    name\_col   \= rest\_cols.get("name")  
    street\_col \= rest\_cols.get("street")  
    postal\_col \= rest\_cols.get("postal")  
    grid\_col   \= rest\_cols.get("grid")

    \# ── Helpers (defined once) ────────────────────────────────────  
    def \_sc(lead\_lat, lead\_pin, item):  
        raw\_n, lat\_n, pin\_n, r \= item  
        ns \= 0.0  
        if lead\_lat and lat\_n:  
            ns \= max(fuzz.token\_sort\_ratio(lead\_lat, lat\_n),  
                     fuzz.token\_set\_ratio(lead\_lat,  lat\_n)) / 100.0  
        if lead\_pin or pin\_n:  
            lp, cp \= lead\_pin or "", pin\_n or ""  
            if lp and cp:  
                ps \= max(fuzz.token\_sort\_ratio(lp, cp),  
                         fuzz.token\_set\_ratio(lp,  cp)) / 100.0  
                return round(max(ns, ps), 3), raw\_n, r  
        return round(ns, 3), raw\_n, r

    def \_crm\_n(r): return str(r.get(crm\_name\_c,"") or "") if crm\_name\_c else ""

    def \_extract\_area(addr\_raw):  
        m \= re.search(r'Area:\\s\*(.+)', str(addr\_raw), re.IGNORECASE)  
        return m.group(1).strip().lower() if m else ""

    def \_geo\_conflict(lead\_text, crm\_area):  
        if not crm\_area: return False  
        lead\_low \= lead\_text.lower()  
        if crm\_area in lead\_low: return False  
        return any(area in lead\_low for area in SG\_AREAS if area \!= crm\_area)

    results \= \[\]

    for \_, row in rest\_df.iterrows():  
        name\_raw \= str(row.get(name\_col,"") or "") if name\_col else ""  
        street   \= str(row.get(street\_col,"") or "") if street\_col else ""  
        postal\_s \= str(row.get(postal\_col,"") or "") if postal\_col else ""  
        grid\_val \= str(row.get(grid\_col,"") or "") if grid\_col else ""  
        postal  \= extract\_sg\_postal(\_norm\_postal\_input(postal\_s)) or extract\_sg\_postal(street)  
        unit    \= extract\_sg\_unit(street)  
        blank   \= is\_blank\_name(name\_raw)  
        lat\_n, pin\_n \= norm\_name\_sg(strip\_venue\_generic(name\_raw), char\_map)

        crm\_match \= None; match\_score \= 0.0; match\_method \= ""

        if postal and not blank:  
            if unit:  
                unit\_matches \= crm\_postal\_unit\_dict.get(f"{postal}|{unit}", \[\])  
                pool \= unit\_matches if unit\_matches else crm\_postal\_all\_dict.get(postal, \[\])  
                tag  \= f"Postal+Unit+Name \[{postal} \#{unit}\]" if unit\_matches \\  
                       else f"Postal+Name \[{postal}\]"  
                \# No-unit fallback uses stricter threshold  
                min\_sc \= p2\_threshold if unit\_matches else max(p2\_threshold, 0.90)  
            else:  
                pool   \= crm\_postal\_all\_dict.get(postal, \[\])  
                tag    \= f"Postal+Name \[{postal}\] (no unit)"  
                min\_sc \= max(p2\_threshold, 0.90)  \# 0.90 minimum for no-unit

            best\_sc, best\_cand, best\_raw \= 0.0, None, ""  
            for item in pool:  
                sc, raw\_n, cand\_row \= \_sc(lat\_n, pin\_n, item)  
                if sc \> best\_sc:  
                    best\_sc, best\_cand, best\_raw \= sc, cand\_row, raw\_n  
            if best\_cand is not None and best\_sc \>= min\_sc:  
                crm\_match \= best\_cand; match\_score \= best\_sc  
                match\_method \= f"{tag} score={best\_sc:.2f}"  
            elif best\_cand is not None:  
                match\_method \= f"No name match at {tag} best='{best\_raw}' score={best\_sc:.2f}"  
            else:  
                match\_method \= f"No CRM accounts at postal {postal}"

        elif blank:  
            match\_method \= "Blank restaurant name — skipped"

        else:  
            \# Zero/no postal → name-only scan at ≥0.90  
            lead\_geo \= f"{name\_raw} {street}".lower()  
            best\_sc, best\_cand \= 0.0, None  
            for item in all\_crm\_items:  
                sc, \_, cand\_row \= \_sc(lat\_n, pin\_n, item)  
                if sc \< 0.90 or sc \<= best\_sc: continue  
                addr\_raw\_c \= str(cand\_row.get(crm\_addr\_c,"") or "") if crm\_addr\_c else ""  
                if \_geo\_conflict(lead\_geo, \_extract\_area(addr\_raw\_c)): continue  
                best\_sc, best\_cand \= sc, cand\_row  
            if best\_cand is not None:  
                crm\_match \= best\_cand; match\_score \= best\_sc  
                match\_method \= f"Name-only (zero postal) score={best\_sc:.2f}"  
            else:  
                match\_method \= "No postal — name-only scan: no match at ≥90%"

        \# ── Label ─────────────────────────────────────────────────  
        if crm\_match is not None:  
            label      \= ("P4 — Duplicate" if match\_score \>= p3\_threshold  
                          else "P3 — Potential Match")  
            dup\_grid   \= str(crm\_match.get(crm\_grid\_c,"") or "") if crm\_grid\_c else ""  
            dup\_name   \= \_crm\_n(crm\_match)  
            dup\_status \= str(crm\_match.get(crm\_stat\_c,"") or "") if crm\_stat\_c else ""  
            dup\_reas   \= str(crm\_match.get(crm\_reas\_c,"") or "") if crm\_reas\_c else ""  
            dup\_addr   \= str(crm\_match.get(crm\_addr\_c,"") or "") if crm\_addr\_c else ""  
            dup\_post   \= str(crm\_match.get(col\_map\_crm.get("postal",""),"") or "") \\  
                         if col\_map\_crm else ""  
        else:  
            label \= "Unverified"  
            dup\_grid \= dup\_name \= dup\_status \= dup\_reas \= dup\_addr \= dup\_post \= ""

        results.append({  
            "GRID":               grid\_val,  
            "Company / Account":  name\_raw,  
            "Street":             street,  
            "Zip/Postal Code":    postal\_s,  
            "Label":              label,  
            "Match Score":        match\_score if crm\_match is not None else "",  
            "Duplicate CRM Name": dup\_name,  
            "Duplicate GRID":     dup\_grid,  
            "Duplicate CRM Address": dup\_addr,  
            "Duplicate CRM Postal":  dup\_post,  
            "CRM Account Status": dup\_status,  
            "CRM Status Reason":  dup\_reas,  
            "Match Reason":       match\_method,  
        })

    return pd.DataFrame(results)

def build\_crm\_check\_excel(df: pd.DataFrame) \-\> bytes:  
    from io import BytesIO

    FILLS\_C \= {  
        "P4 — Duplicate":       PatternFill("solid", start\_color="FFC7CE"),  
        "P3 — Potential Match": PatternFill("solid", start\_color="FFF2CC"),  
        "Unverified":           PatternFill("solid", start\_color="DBEAFE"),  
    }  
    ALT\_C \= {  
        "P4 — Duplicate":       PatternFill("solid", start\_color="FFE0E0"),  
        "P3 — Potential Match": PatternFill("solid", start\_color="FFFAE0"),  
        "Unverified":           PatternFill("solid", start\_color="EFF6FF"),  
    }  
    FC\_C \= {  
        "P4 — Duplicate":       "9C0006",  
        "P3 — Potential Match": "7D5A00",  
        "Unverified":           "1E40AF",  
    }  
    HDR\_FILL \= PatternFill("solid", start\_color="1F4E79")

    def \_thin():  
        s \= Side(style="thin", color="D0D0D0")  
        return Border(left=s, right=s, top=s, bottom=s)

    heads  \= \["GRID","Company / Account","Street","Zip/Postal Code","Label",  
              "Match Score","Duplicate CRM Name","Duplicate GRID",  
              "Duplicate CRM Address","Duplicate CRM Postal",  
              "CRM Account Status","CRM Status Reason","Match Reason"\]  
    widths \= \[10,32,36,14,22,12,30,14,36,14,16,18,40\]  
    widths \= \[32,36,14,22,12,30,14,18,18,45\]

    def \_build\_sheet(ws, title, data):  
        ws\["A1"\] \= title  
        ws\["A1"\].font \= Font(name="Poppins", bold=True, size=12, color="1F4E79")  
        ws.merge\_cells(f"A1:{get\_column\_letter(len(heads))}1")  
        for ci, h in enumerate(heads, 1):  
            c \= ws.cell(row=3, column=ci, value=h)  
            c.font      \= Font(name="Poppins", bold=True, color="FFFFFF", size=9)  
            c.fill      \= HDR\_FILL  
            c.alignment \= Alignment(horizontal="center", vertical="center", wrap\_text=True)  
            c.border    \= \_thin()  
        ws.row\_dimensions\[3\].height \= 30  
        for ri, (\_, row) in enumerate(data.iterrows(), 4):  
            lbl  \= row\["Label"\]  
            fill \= FILLS\_C.get(lbl, PatternFill()) if ri % 2 \== 0 else ALT\_C.get(lbl, PatternFill())  
            fc   \= FC\_C.get(lbl, "000000")  
            ws.row\_dimensions\[ri\].height \= 20  
            for ci, key in enumerate(heads, 1):  
                val \= row.get(key,""); val \= "" if pd.isna(val) else val  
                c   \= ws.cell(row=ri, column=ci, value=val)  
                c.border \= \_thin()  
                if key \== "Label":  
                    c.font \= Font(name="Poppins", size=8, bold=True, color=fc)  
                    c.fill \= fill  
                    c.alignment \= Alignment(horizontal="center", vertical="center")  
                elif key \== "Match Score":  
                    try: c.value \= float(val); c.number\_format \= "0.00"  
                    except Exception: pass  
                    c.font \= Font(name="Poppins", size=8, bold=True)  
                    c.fill \= fill  
                    c.alignment \= Alignment(horizontal="center", vertical="center")  
                elif key \== "Match Reason":  
                    c.font \= Font(name="Poppins", size=8, color="595959")  
                    c.fill \= fill  
                    c.alignment \= Alignment(vertical="center", wrap\_text=True)  
                else:  
                    c.font \= Font(name="Poppins", size=8)  
                    c.fill \= fill  
                    c.alignment \= Alignment(vertical="center")  
        for i, w in enumerate(widths, 1):  
            ws.column\_dimensions\[get\_column\_letter(i)\].width \= w  
        ws.freeze\_panes \= "A4"  
        if len(data) \> 0:  
            ws.auto\_filter.ref \= f"A3:{get\_column\_letter(len(heads))}{3 \+ len(data)}"

    wb  \= Workbook()  
    ws1 \= wb.active; ws1.title \= "All Results"  
    \_build\_sheet(ws1, "CRM Duplicate Check — All Results", df)

    ws2 \= wb.create\_sheet("✅ Unverified — Create")  
    \_build\_sheet(ws2, "Unverified — No CRM Match (Safe to Create)",  
                 df\[df\["Label"\] \== "Unverified"\].reset\_index(drop=True))

    ws3 \= wb.create\_sheet("🟡 P3 — Review First")  
    \_build\_sheet(ws3, "P3 — Potential Match (Review Before Creating)",  
                 df\[df\["Label"\] \== "P3 — Potential Match"\].reset\_index(drop=True))

    ws4 \= wb.create\_sheet("🔴 P4 — Duplicates")  
    \_build\_sheet(ws4, "P4 — Duplicate (Already in Salesforce — Skip)",  
                 df\[df\["Label"\] \== "P4 — Duplicate"\].reset\_index(drop=True))

    out \= BytesIO()  
    wb.save(out)  
    return out.getvalue()

def build\_excel(df: pd.DataFrame, market\_name: str):  
    \# ── Color palette ──────────────────────────────────────────────  
    FILLS \= {  
        "P1 — New":            PatternFill("solid", start\_color="C6EFCE"),  
        "P2 — Please Check":   PatternFill("solid", start\_color="ECECEC"),  
        "P3 — Potential Match":PatternFill("solid", start\_color="FFF2CC"),  
        "P4 — Duplicate":      PatternFill("solid", start\_color="FFC7CE"),  
        "Business Closed":     PatternFill("solid", start\_color="FFEB9C"),  
        "Wrong Target Group":  PatternFill("solid", start\_color="FFDCA8"),  
    }  
    ALT \= {k: PatternFill("solid", start\_color=  
           {"P1 — New":"EBF7EB","P2 — Please Check":"F7F7F7",  
            "P3 — Potential Match":"FFFAE0","P4 — Duplicate":"FFE0E0",  
            "Business Closed":"FFF7D1","Wrong Target Group":"FFF0DC"}.get(k,"FFFFFF"))  
           for k in FILLS}  
    FONT\_C \= {  
        "P1 — New":"276221","P2 — Please Check":"595959",  
        "P3 — Potential Match":"7D5A00","P4 — Duplicate":"9C0006",  
        "Business Closed":"7D4E00","Wrong Target Group":"833C00",  
    }  
    HDR\_FILL \= PatternFill("solid", start\_color="1F4E79")  
    SEC\_FILL \= PatternFill("solid", start\_color="2E75B6")

    def \_thin():  
        s \= Side(style="thin", color="D0D0D0")  
        return Border(left=s, right=s, top=s, bottom=s)  
    def \_hdr(ws, r, c, v):  
        cell \= ws.cell(row=r, column=c, value=v)  
        cell.font      \= Font(name="Poppins", bold=True, color="FFFFFF", size=9)  
        cell.fill      \= HDR\_FILL  
        cell.alignment \= Alignment(horizontal="center", vertical="center", wrap\_text=True)  
        cell.border    \= \_thin()  
    def \_sec(ws, r, c, text, span=3):  
        ws.merge\_cells(start\_row=r, start\_column=c,  
                       end\_row=r, end\_column=c \+ span \- 1\)  
        cell \= ws.cell(row=r, column=c, value=f"  {text}")  
        cell.font      \= Font(name="Poppins", bold=True, color="FFFFFF", size=9)  
        cell.fill      \= SEC\_FILL  
        cell.alignment \= Alignment(vertical="center")  
        ws.row\_dimensions\[r\].height \= 22  
    def \_dc(ws, r, c, v, fill=None, bold=False, fmt=None,  
             align="center", color="000000"):  
        cell \= ws.cell(row=r, column=c, value=v)  
        cell.font      \= Font(name="Poppins", size=8, bold=bold, color=color)  
        cell.fill      \= fill if fill else PatternFill()  
        cell.alignment \= Alignment(horizontal=align, vertical="center")  
        cell.border    \= \_thin()  
        if fmt: cell.number\_format \= fmt

    def \_fill(lbl, ri):  
        base \= FILLS.get(lbl, PatternFill())  
        alt  \= ALT.get(lbl, PatternFill())  
        return base if ri % 2 \== 0 else alt

    def \_conf\_fill(score):  
        try: score \= float(score)  
        except Exception: return PatternFill()  
        if score \>= 0.8: return PatternFill("solid", start\_color="C6EFCE")  
        if score \>= 0.6: return PatternFill("solid", start\_color="FFEB9C")  
        return PatternFill("solid", start\_color="FFC7CE")

    labels\_order \= \[  
        "P1 — New", "P3 — Potential Match", "P4 — Duplicate",  
        "Business Closed", "Wrong Target Group", "P2 — Please Check",  
    \]  
    col\_headers \= \[  
        \# ── A–F: Agent workflow columns (blank on export, filled by agents) ──  
        "Agent","Due Date","Convert/Lost","Invalid Reason",  
        "Comments/Duplicate GRID","Feedback",  
        \# ── G onwards: classifier output ─────────────────────────────────────  
        "GRID","Lead ID","Company / Account","City","Street","Phone",   \# G–L  
        "GM Title","GM Category","GM Business Status",                  \# M–O  
        "GM Phone","GM Website","GM URL",                               \# P–R  
        "Match Confidence","Match Reason","Label","Match Score (CRM)",  \# S–V  
        "Duplicate GRID","Duplicate CRM Name",                          \# W–X  
        "Duplicate CRM Address","Duplicate CRM Postal",                 \# Y–Z (new)  
        "CRM Account Status","CRM Status Reason","Duplicate Match Method", \# AA–AC  
        "Previous Occupant GRID","Previous Occupant Name",              \# AD–AE  
        "Delivery Zone Status","Zone Name","Zone City","Zone Method",    \# AF–AI  
    \]  
    col\_w \= \[  
        10,12,14,18,22,20,              \# Agent cols A–F  
        10,18,32,12,36,16,              \# G–L (GRID … Phone)  
        28,24,18,16,28,48,              \# M–R (GM Title … GM URL)  
        14,38,24,14,14,30,              \# S–V (Match Conf … Match Score)  
        16,18,36,14,                    \# W–Z (Dup GRID, Dup Name, Dup Addr, Dup Postal)  
        16,18,40,14,30,18,18,14,14,    \# AA–AI (CRM Status … Zone Method)  
    \]

    DATA\_S \= 5  
    DATA\_E \= DATA\_S \+ len(df) \- 1  
    LBL\_R  \= f"'Classified Leads'\!U{DATA\_S}:U{DATA\_E}"   \# col 21 \= U (unchanged)  
    CAT\_R  \= f"'Classified Leads'\!N{DATA\_S}:N{DATA\_E}"   \# col 14 \= N (unchanged)  
    CRMS\_R \= f"'Classified Leads'\!AA{DATA\_S}:AA{DATA\_E}" \# col 27 \= AA (was Y, \+2)  
    METH\_R \= f"'Classified Leads'\!AC{DATA\_S}:AC{DATA\_E}" \# col 29 \= AC (was AA, \+2)  
    ZONE\_R \= f"'Classified Leads'\!AF{DATA\_S}:AF{DATA\_E}" \# col 32 \= AF (was AD, \+2)  
    has\_z  \= df\["Delivery Zone Status"\].astype(str).str.strip().ne("").any()

    wb  \= Workbook()

    \# ── Sheet 1: Classified Leads ─────────────────────────────────  
    ws1 \= wb.active; ws1.title \= "Classified Leads"  
    counts \= df\["Label"\].value\_counts()  
    ws1\["A1"\] \= f"Lead Classification Report  |  {market\_name}"  
    ws1\["A1"\].font \= Font(name="Poppins", bold=True, size=14, color="1F4E79")  
    ws1.merge\_cells("A1:AG1")  
    ws1\["A2"\] \= "Total: {:,}   |   {}".format(  
        len(df), "   ".join(f"{l}: {counts.get(l,0):,}" for l in labels\_order))  
    ws1\["A2"\].font \= Font(name="Poppins", italic=True, size=8, color="595959")  
    ws1.merge\_cells("A2:AG2")  
    for ci, h in enumerate(col\_headers, 1): \_hdr(ws1, 4, ci, h)  
    ws1.row\_dimensions\[4\].height \= 30

    AGENT\_KEYS \= {"Agent","Due Date","Convert/Lost","Invalid Reason",  
                  "Comments/Duplicate GRID","Feedback"}

    for ri, (\_, row) in enumerate(df.iterrows(), DATA\_S):  
        lbl  \= row\["Label"\]  
        fill \= \_fill(lbl, ri)  
        lc   \= FONT\_C.get(lbl,"000000")  
        ws1.row\_dimensions\[ri\].height \= 20   \# taller rows for Poppins 8pt

        for ci, key in enumerate(col\_headers, 1):  
            val \= row.get(key,""); val \= "" if pd.isna(val) else val

            \# ── Agent workflow columns A–F: plain white, blank ────  
            if key in AGENT\_KEYS:  
                c \= ws1.cell(row=ri, column=ci, value="")  
                c.border \= \_thin()  
                c.font   \= Font(name="Poppins", size=8)  
                c.fill   \= PatternFill("solid", start\_color="FFFFFF")  
                c.alignment \= Alignment(vertical="center")  
                if key \== "Due Date":  
                    c.number\_format \= "DD-MMM-YY"  
                continue

            \# ── Classifier output columns G onwards ───────────────  
            c \= ws1.cell(row=ri, column=ci, value=val)  
            c.border \= \_thin()  
            if key \== "Label":  
                c.font \= Font(name="Poppins", size=8, bold=True, color=lc)  
                c.fill \= fill; c.alignment \= Alignment(horizontal="center", vertical="center")  
            elif key in ("Match Confidence","Match Score (CRM)"):  
                try: c.value \= float(val); c.number\_format \= "0.00"  
                except Exception: pass  
                c.font \= Font(name="Poppins", size=8, bold=True)  
                c.fill \= \_conf\_fill(val)  
                c.alignment \= Alignment(horizontal="center", vertical="center")  
            elif key in ("Match Reason","Duplicate Match Method"):  
                c.font \= Font(name="Poppins", size=8, color="595959")  
                c.fill \= fill; c.alignment \= Alignment(vertical="center", wrap\_text=True)  
            elif key \== "Delivery Zone Status":  
                zc \= ("276221" if val \== "Within Zone" else  
                      "9C0006" if val \== "Outside Zone" else "595959")  
                zf \= (PatternFill("solid", start\_color="C6EFCE") if val \== "Within Zone"  
                      else PatternFill("solid", start\_color="FFC7CE") if val \== "Outside Zone"  
                      else PatternFill("solid", start\_color="EFEFEF"))  
                c.font \= Font(name="Poppins", size=8, bold=True, color=zc)  
                c.fill \= zf; c.alignment \= Alignment(horizontal="center", vertical="center")  
            else:  
                c.font \= Font(name="Poppins", size=8)  
                c.fill \= fill; c.alignment \= Alignment(vertical="center")

    for i, w in enumerate(col\_w, 1):  
        ws1.column\_dimensions\[get\_column\_letter(i)\].width \= w

    \# ── Dropdowns for agent columns ───────────────────────────────  
    dv\_convert \= DataValidation(  
        type="list", formula1='"Converted,Lost"', allow\_blank=True)  
    dv\_reason  \= DataValidation(  
        type="list",  
        formula1='"Duplicate,Invalid Data,Closed Down,Wrong Target Group,Other"',  
        allow\_blank=True)  
    ws1.add\_data\_validation(dv\_convert)  
    ws1.add\_data\_validation(dv\_reason)  
    dv\_convert.sqref \= f"C{DATA\_S}:C{DATA\_E}"  
    dv\_reason.sqref  \= f"D{DATA\_S}:D{DATA\_E}"

    ws1.freeze\_panes     \= "G5"  
    ws1.auto\_filter.ref  \= f"A4:AI{DATA\_E}"

    \# ── Sheet 2: Summary ──────────────────────────────────────────  
    ws2 \= wb.create\_sheet("Summary")  
    ws2\["A1"\] \= "Classification Summary"; ws2\["A1"\].font \= Font(name="Poppins", bold=True, size=14, color="1F4E79")  
    ws2\["A2"\] \= "Formula-driven — auto-updates when Classified Leads is edited."  
    ws2\["A2"\].font \= Font(name="Poppins", italic=True, size=8, color="595959")  
    ws2.merge\_cells("A2:J2")

    r \= 4  
    \_sec(ws2, r, 1, "CLASSIFICATION BREAKDOWN", 3); r \+= 1  
    for ci, h in enumerate(\["Label","Count","% of Total"\], 1): \_hdr(ws2, r, ci, h)  
    r \+= 1  
    s1 \= r; total\_r \= s1 \+ len(labels\_order)  
    for i, lbl in enumerate(labels\_order):  
        ri \= s1 \+ i  
        \_dc(ws2, ri, 1, lbl,  fill=FILLS.get(lbl), bold=True, align="left", color=FONT\_C.get(lbl,"000000"))  
        \_dc(ws2, ri, 2, f'=COUNTIF({LBL\_R},"{lbl}")', fill=FILLS.get(lbl))  
        \_dc(ws2, ri, 3, f'=IF(B{total\_r}=0,0,B{ri}/B{total\_r})', fill=FILLS.get(lbl), fmt="0.0%")  
    r \= total\_r  
    \_dc(ws2, r, 1, "TOTAL", bold=True, align="left")  
    \_dc(ws2, r, 2, f'=SUM(B{s1}:B{total\_r-1})', bold=True)  
    \_dc(ws2, r, 3, "100.0%", bold=True, fmt="0.0%")  
    r \+= 2

    \# Match method breakdown  
    \_sec(ws2, r, 1, "MATCH METHOD BREAKDOWN", 3); r \+= 1  
    for ci, h in enumerate(\["Method","Count","% of CRM Matches"\], 1): \_hdr(ws2, r, ci, h)  
    r \+= 1  
    crm\_labels \= \["P3 — Potential Match","P4 — Duplicate"\]  
    crm\_total  \= '+'.join(\[f'COUNTIF({LBL\_R},"{l}")' for l in crm\_labels\])  
    methods \= \[  
        ("Postal+Unit+Name", f'=COUNTIFS({METH\_R},"Postal+Unit\*")', FILLS\["P4 — Duplicate"\]),  
        ("Postal+Name",      f'=COUNTIFS({METH\_R},"Postal+Name\*")', ALT\["P4 — Duplicate"\]),  
        ("New at known addr",f'=COUNTIFS({METH\_R},"New business\*")',FILLS\["P3 — Potential Match"\]),  
    \]  
    for ml, mf, mfill in methods:  
        \_dc(ws2, r, 1, ml, fill=mfill, align="left")  
        \_dc(ws2, r, 2, mf, fill=mfill)  
        \_dc(ws2, r, 3, f'=IF(({crm\_total})=0,0,B{r}/({crm\_total}))', fill=mfill, fmt="0.0%")  
        r \+= 1  
    r \+= 1

    \# Top GM categories  
    \_sec(ws2, r, 1, "TOP GM CATEGORIES (Qualified Leads)", 3); r \+= 1  
    for ci, h in enumerate(\["Category","Count","% of Matched"\], 1): \_hdr(ws2, r, ci, h)  
    r \+= 1  
    top\_cats \= (df\[df\["GM Category"\].notna() & (df\["GM Category"\] \!= "")\]  
                \["GM Category"\].value\_counts().head(20).index.tolist())  
    matched\_f \= f'=COUNTIF({LBL\_R},"P1 — New")'  
    hid\_r \= r \+ len(top\_cats) \+ 1  
    ws2.cell(row=hid\_r, column=2).value \= matched\_f  
    ws2.cell(row=hid\_r, column=2).font  \= Font(color="FFFFFF", size=1)  
    for i, cat in enumerate(top\_cats):  
        fl \= (PatternFill("solid", start\_color="EBF3FB")  
              if i % 2 \== 0 else PatternFill("solid", start\_color="FFFFFF"))  
        \_dc(ws2, r, 1, cat, fill=fl, align="left")  
        \_dc(ws2, r, 2, f'=COUNTIF({CAT\_R},"{cat}")', fill=fl)  
        \_dc(ws2, r, 3, f'=IF(B{hid\_r}=0,0,B{r}/B{hid\_r})', fill=fl, fmt="0.0%")  
        r \+= 1  
    r \= hid\_r \+ 2

    \# Zone section  
    if has\_z:  
        \_sec(ws2, r, 1, "DELIVERY ZONE BREAKDOWN", 3); r \+= 1  
        for ci, h in enumerate(\["Zone Status","Count","% of Total"\], 1): \_hdr(ws2, r, ci, h)  
        r \+= 1  
        for zs, zcol, bg in \[("Within Zone","276221","C6EFCE"),  
                               ("Outside Zone","9C0006","FFC7CE"),  
                               ("Geocoding Failed","595959","D9D9D9")\]:  
            fl \= PatternFill("solid", start\_color=bg)  
            \_dc(ws2, r, 1, zs, fill=fl, bold=True, align="left", color=zcol)  
            \_dc(ws2, r, 2, f'=COUNTIF({ZONE\_R},"{zs}")', fill=fl)  
            \_dc(ws2, r, 3, f'=IF(B{total\_r}=0,0,B{r}/B{total\_r})', fill=fl, fmt="0.0%")  
            r \+= 1

    for col, w in zip("ABCDE",\[32,12,14,4,26\]):  
        ws2.column\_dimensions\[col\].width \= w

    \# ── Sheet 3: P1 New ───────────────────────────────────────────  
    def \_mini\_sheet(wb, title\_text, title\_color, filter\_fn, heads, widths, sheet\_name):  
        ws \= wb.create\_sheet(sheet\_name)  
        sub \= df\[filter\_fn(df)\].reset\_index(drop=True)  
        ws\["A1"\] \= f"{title\_text}  ({len(sub):,} leads)"  
        ws\["A1"\].font \= Font(name="Poppins", bold=True, size=12, color=title\_color)  
        ws.merge\_cells(f"A1:{get\_column\_letter(len(heads))}1")  
        for ci, h in enumerate(heads, 1): \_hdr(ws, 3, ci, h)  
        for ri, (\_, row) in enumerate(sub.iterrows(), 4):  
            lbl  \= row\["Label"\]  
            fill \= \_fill(lbl, ri)  
            lc   \= FONT\_C.get(lbl,"000000")  
            ws.row\_dimensions\[ri\].height \= 20  
            for ci, key in enumerate(heads, 1):  
                val \= row.get(key,""); val \= "" if pd.isna(val) else val  
                c   \= ws.cell(row=ri, column=ci, value=val)  
                c.border \= \_thin()  
                if key \== "Label":  
                    c.font \= Font(name="Poppins",size=8,bold=True,color=lc)  
                    c.fill \= fill; c.alignment \= Alignment(horizontal="center",vertical="center")  
                elif key in ("Match Confidence","Match Score (CRM)"):  
                    try: c.value \= float(val); c.number\_format \= "0.00"  
                    except Exception: pass  
                    c.font \= Font(name="Poppins",size=8,bold=True)  
                    c.fill \= \_conf\_fill(val)  
                    c.alignment \= Alignment(horizontal="center",vertical="center")  
                elif key in ("Match Reason","Duplicate Match Method"):  
                    c.font \= Font(name="Poppins",size=8,color="595959")  
                    c.fill \= fill; c.alignment \= Alignment(vertical="center",wrap\_text=True)  
                else:  
                    c.font \= Font(name="Poppins",size=8)  
                    c.fill \= fill; c.alignment \= Alignment(vertical="center")  
        for i, w in enumerate(widths, 1):  
            ws.column\_dimensions\[get\_column\_letter(i)\].width \= w  
        ws.freeze\_panes \= "A4"  
        if len(sub) \> 0:  
            ws.auto\_filter.ref \= f"A3:{get\_column\_letter(len(heads))}{3+len(sub)}"  
        return ws

    new\_h \= \["GRID","Lead ID","Company / Account","City","Street","Phone",  
             "GM Title","GM Category","GM Business Status",  
             "GM Phone","GM Website","GM URL","Match Confidence","Match Reason",  
             "Previous Occupant GRID","Previous Occupant Name",  
             "Delivery Zone Status","Zone Name"\]  
    \_mini\_sheet(wb, "✅ P1 — New Leads", "276221",  
                lambda d: d\["Label"\] \== "P1 — New",  
                new\_h, \[10,18,32,12,36,16,28,24,18,16,28,48,14,38,14,30,18,18\],  
                "✅ P1 — New")

    dup\_h \= \["GRID","Lead ID","Company / Account","City","Street","Phone",  
             "Label","Duplicate GRID","Duplicate CRM Name",  
             "Duplicate CRM Address","Duplicate CRM Postal",  
             "CRM Account Status","CRM Status Reason",  
             "Duplicate Match Method","Match Score (CRM)"\]  
    \_mini\_sheet(wb, "🔴 P4 — Duplicates", "9C0006",  
                lambda d: d\["Label"\] \== "P4 — Duplicate",  
                dup\_h, \[10,18,32,12,36,16,20,14,30,36,14,16,18,40,14\],  
                "🔴 P4 — Duplicate")

    pot\_h \= \["GRID","Lead ID","Company / Account","City","Street","Phone",  
             "Label","Duplicate GRID","Duplicate CRM Name",  
             "Duplicate CRM Address","Duplicate CRM Postal",  
             "CRM Account Status","Duplicate Match Method","Match Score (CRM)",  
             "GM Title","GM Category","Match Confidence","Match Reason"\]  
    \_mini\_sheet(wb, "🟡 P3 — Potential Match", "7D5A00",  
                lambda d: d\["Label"\] \== "P3 — Potential Match",  
                pot\_h, \[10,18,32,12,36,16,22,14,30,36,14,16,40,14,28,24,14,38\],  
                "🟡 P3 — Potential")

    chk\_h \= \["GRID","Lead ID","Company / Account","City","Street",  
             "GM Title","GM Category","GM Business Status",  
             "Match Confidence","Match Reason","GM URL"\]  
    \_mini\_sheet(wb, "⚪ P2 — Please Check", "595959",  
                lambda d: d\["Label"\] \== "P2 — Please Check",  
                chk\_h, \[10,18,32,12,36,28,24,18,14,38,48\],  
                "⚪ P2 — Please Check")

    other\_h \= \["GRID","Lead ID","Company / Account","City","Phone","Label",  
               "GM Title","GM Category","GM Business Status",  
               "Match Confidence","Match Reason","GM URL"\]  
    \_mini\_sheet(wb, "⚠️ Closed / Wrong Target Group", "7D4E00",  
                lambda d: d\["Label"\].isin(\["Business Closed","Wrong Target Group"\]),  
                other\_h, \[10,18,32,12,16,22,28,24,18,14,38,48\],  
                "⚠️ Closed \+ Wrong TG")

    buf \= io.BytesIO(); wb.save(buf); buf.seek(0)  
    return buf

\# ═════════════════════════════════════════════════════════════════  
\# STREAMLIT MAIN  
\# ═════════════════════════════════════════════════════════════════

def \_col\_map\_ui(df, auto\_map: dict, fields: list, key\_prefix: str) \-\> dict:  
    """  
    Render a column mapping expander.  
    fields: list of (internal\_key, display\_label, required: bool)  
    Returns updated mapping dict {key: column\_name\_or\_None}.  
    Auto-expands if any required field is unmapped.  
    """  
    all\_cols  \= \["(not mapped)"\] \+ list(df.columns)  
    missing   \= \[label for key, label, req in fields  
                 if req and auto\_map.get(key) is None\]  
    expanded  \= bool(missing)  
    icon      \= "⚠️" if missing else "✅"  
    label\_txt \= (f"Column mapping {icon} — missing: {', '.join(missing)}"  
                 if missing else f"Column mapping {icon} — all detected")

    result \= dict(auto\_map)  
    with st.expander(label\_txt, expanded=expanded):  
        cols\_ui \= st.columns(2)  
        for idx, (key, label, \_) in enumerate(fields):  
            current \= auto\_map.get(key)  
            default \= all\_cols.index(current) if current and current in all\_cols else 0  
            sel \= cols\_ui\[idx % 2\].selectbox(  
                label, all\_cols, index=default,  
                key=f"{key\_prefix}\_{key}")  
            result\[key\] \= sel if sel \!= "(not mapped)" else None  
    return result

\# ── Google Maps → Salesforce cuisine picklist mapping ─────────────  
APIFY\_CUISINE\_MAP \= {  
    "thai restaurant":"Thai","thai food":"Thai",  
    "japanese restaurant":"Japanese","ramen restaurant":"Japanese","izakaya":"Japanese",  
    "sushi restaurant":"Sushi",  
    "korean restaurant":"Korean",  
    "chinese restaurant":"Chinese",  
    "indian restaurant":"Indian",  
    "italian restaurant":"Italian",  
    "vietnamese restaurant":"Vietnamese",  
    "pizza restaurant":"Pizza","pizza delivery":"Pizza",  
    "malay restaurant":"Malay","nasi padang restaurant":"Malay",  
    "indonesian restaurant":"Indonesian",  
    "western restaurant":"Western",  
    "fast food restaurant":"Fast Food",  
    "seafood restaurant":"Seafood",  
    "vegetarian restaurant":"Vegetarian","vegan restaurant":"Vegetarian",  
    "american restaurant":"American",  
    "french restaurant":"French","patisserie":"French",  
    "mexican restaurant":"Mexican",  
    "mediterranean restaurant":"Mediterranean",  
    "middle eastern restaurant":"Middle Eastern",  
    "filipino restaurant":"Filipino",  
    "sandwich shop":"Sandwiches","sandwich restaurant":"Sandwiches",  
    "german restaurant":"German",  
    "greek restaurant":"Greek",  
    "spanish restaurant":"Spanish",  
    "turkish restaurant":"Turkish",  
    "lebanese restaurant":"Lebanese",  
    "malaysian restaurant":"Malaysian",  
    "bakery":"Cakes & Bakery","cake shop":"Cakes & Bakery","confectionery":"Cakes & Bakery",  
    "dessert shop":"Desserts","dessert restaurant":"Desserts","ice cream shop":"Desserts",  
    "bubble tea shop":"Non-alcoholic Drinks","juice bar":"Non-alcoholic Drinks",  
    "tea house":"Non-alcoholic Drinks",  
    "chicken restaurant":"Chicken","fried chicken restaurant":"Chicken",  
    "burger restaurant":"Burgers","hamburger restaurant":"Burgers",  
    "steakhouse":"Meat","barbecue restaurant":"Meat",  
    "halal restaurant":"Halal",  
    "health food restaurant":"Healthy Food","salad shop":"Healthy Food",  
    "singaporean restaurant":"Singaporean","hawker stall":"Singaporean",  
    "asian restaurant":"Asian","pan asian restaurant":"Asian",  
    "southeast asian restaurant":"South East Asian",  
    "international restaurant":"International",  
}

\_NAMING\_RE \= re.compile(r'^(.+?)\\s\*(\[)\]+)\\s\*$')

def \_map\_apify\_cuisine(category: str) \-\> str:  
    """Map Google Maps category to Salesforce cuisine picklist. Returns '' if unmappable."""  
    if not category:  
        return ""  
    cat \= str(category).strip().lower()  
    if cat in APIFY\_CUISINE\_MAP:  
        return APIFY\_CUISINE\_MAP\[cat\]  
    for key, val in APIFY\_CUISINE\_MAP.items():  
        if key in cat:  
            return val  
    return ""

def \_check\_naming\_convention(name: str):  
    """Check RESTAURANT NAME (LOCATION) structure only.  
    Valid: any text followed by (any text in brackets).  
    No case enforcement — Title Case and ALL CAPS are both acceptable.  
    Returns (ok: bool, detail: str)."""  
    if not name or (isinstance(name, float) and pd.isna(name)):  
        return False, "Account name is blank"  
    name \= str(name).strip()  
    m \= \_NAMING\_RE.match(name)  
    if not m:  
        return False, "Missing (Location) part — expected: NAME (LOCATION)"  
    if not m.group(1).strip():  
        return False, "Restaurant name part is empty"  
    if not m.group(2).strip():  
        return False, "Location part inside brackets is empty"  
    return True, "Pass"

def clean\_lead\_report(df: pd.DataFrame) \-\> pd.DataFrame:  
    """Strip Salesforce footer rows and rows with invalid GRID."""  
    FOOTER \= \["confidential", "copyright", "salesforce.com"\]  
    mask \= df\["Lead Source"\].astype(str).str.lower().apply(  
        lambda v: not any(f in v for f in FOOTER))  
    df \= df\[mask & df\["GRID"\].astype(str).str.strip().ne("nan")\].copy()  
    return df.reset\_index(drop=True)

def sample\_leads(df: pd.DataFrame, seed: int \= 42\) \-\> tuple:  
    """  
    Stratified 10% sample by Edited By × Lead Source × Lead Status.  
    Returns (sampled\_df, summary\_df).  
    """  
    import math  
    strata \= \["Edited By", "Lead Source", "Lead Status"\]  
    sampled, summary \= \[\], \[\]  
    for keys, grp in df.groupby(strata, dropna=False):  
        n \= len(grp)  
        if n \== 0:  
            continue  
        k \= min(max(1, math.ceil(n \* 0.10)), n)  
        sampled.append(grp.sample(n=k, random\_state=seed))  
        summary.append({  
            "Edited By": keys\[0\], "Lead Source": keys\[1\],  
            "Lead Status": keys\[2\], "Total Leads": n, "Sampled": k,  
        })  
    sampled\_df  \= pd.concat(sampled,  ignore\_index=True) if sampled  else df.iloc\[:0\]  
    summary\_df  \= pd.DataFrame(summary)  
    return sampled\_df, summary\_df

def run\_kpi\_checks(sampled\_df, apify\_df, crm\_df, col\_map\_crm,  
                   account\_df, account\_cols,  
                   zones, char\_map,  
                   p2\_threshold=0.50, p3\_threshold=0.75):  
    """  
    Run all KPI checks on sampled leads.  
    Returns (results\_df, agent\_summary\_df).  
    """

    def \_sv(row, col, default=""):  
        """Safe scalar from a pandas Series row.  
        Handles duplicate column names, pd.NA, NaT, and None."""  
        if not col or row is None:  
            return default  
        try:  
            if col not in row.index:  
                return default  
            val \= row\[col\]  
            if isinstance(val, pd.Series):   \# duplicate column → take first  
                val \= val.iloc\[0\] if len(val) \> 0 else default  
            if pd.isna(val):  
                return default  
            return val  
        except Exception:  
            return default

    \# ── Build CRM indexes ─────────────────────────────────────────  
    crm\_postal\_all\_dict  \= {}  
    crm\_postal\_unit\_dict \= {}  
    all\_crm\_items        \= \[\]  
    postal\_c \= col\_map\_crm.get("postal") if col\_map\_crm else None  
    addr\_c   \= col\_map\_crm.get("street") if col\_map\_crm else None  
    name\_c   \= col\_map\_crm.get("name")   if col\_map\_crm else None  
    grid\_c   \= col\_map\_crm.get("grid")   if col\_map\_crm else None

    if crm\_df is not None:  
        for \_, r in crm\_df.iterrows():  
            raw\_n      \= str(\_sv(r, name\_c)   or "")  
            lat\_n, pin\_n \= norm\_name\_sg(strip\_venue\_generic(raw\_n), char\_map)  
            item       \= (raw\_n, lat\_n, pin\_n, r)  
            all\_crm\_items.append(item)  
            postal\_raw \= \_norm\_postal\_input(\_sv(r, postal\_c))  
            addr\_raw   \= str(\_sv(r, addr\_c)   or "")  
            crm\_p  \= extract\_sg\_postal(postal\_raw) or extract\_sg\_postal(addr\_raw)  
            crm\_u  \= extract\_sg\_unit(addr\_raw)  
            if crm\_p:  
                crm\_postal\_all\_dict.setdefault(crm\_p, \[\]).append(item)  
                if crm\_u:  
                    crm\_postal\_unit\_dict.setdefault(f"{crm\_p}|{crm\_u}", \[\]).append(item)

    \# ── Build Apify dict (GRID → row) ─────────────────────────────  
    apify\_dict \= {}  
    if apify\_df is not None:  
        gc \= detect\_column(apify\_df, \["GRID","grid","Grid"\])  
        if gc:  
            for \_, r in apify\_df.iterrows():  
                g \= str(\_sv(r, gc) or "").strip()  
                if g and g not in apify\_dict:  
                    apify\_dict\[g\] \= r  
        \_apy\_title \= detect\_column(apify\_df, \["title","name"\])        or "title"  
        \_apy\_cat   \= detect\_column(apify\_df, \["categoryName","category"\]) or "categoryName"  
        \_apy\_perm  \= detect\_column(apify\_df, \["permanentlyClosed"\])   or "permanentlyClosed"  
        \_apy\_temp  \= detect\_column(apify\_df, \["temporarilyClosed"\])   or "temporarilyClosed"  
        \_apy\_phone \= detect\_column(apify\_df, \["phone","phoneNumber"\]) or "phone"  
        \_apy\_web   \= detect\_column(apify\_df, \["website"\])             or "website"  
    else:  
        \_apy\_title \= \_apy\_cat \= \_apy\_perm \= \_apy\_temp \= \_apy\_phone \= \_apy\_web \= ""

    \# ── Build account details dict (GRID → row) ───────────────────  
    account\_dict \= {}  
    if account\_df is not None:  
        ag \= account\_cols.get("grid")  
        if ag:  
            for \_, r in account\_df.iterrows():  
                g \= str(\_sv(r, ag) or "").strip()  
                if g and g not in account\_dict:  
                    account\_dict\[g\] \= r

    \# ── Scoring function ──────────────────────────────────────────  
    def \_sc(lead\_lat, lead\_pin, item):  
        raw\_n, lat\_n, pin\_n, r \= item  
        ns \= 0.0  
        if lead\_lat and lat\_n:  
            ns \= max(fuzz.token\_sort\_ratio(lead\_lat, lat\_n),  
                     fuzz.token\_set\_ratio(lead\_lat, lat\_n)) / 100.0  
        if lead\_pin or pin\_n:  
            lp, cp \= lead\_pin or "", pin\_n or ""  
            if lp and cp:  
                ps \= max(fuzz.token\_sort\_ratio(lp, cp),  
                         fuzz.token\_set\_ratio(lp, cp)) / 100.0  
                return round(max(ns, ps), 3), raw\_n, r  
        return round(ns, 3), raw\_n, r

    \# ── Dedup helper (excludes own GRID) ─────────────────────────  
    NO\_UNIT\_MIN \= max(p2\_threshold, 0.90)

    def \_find\_dup(lead\_grid, lead\_lat, lead\_pin, lead\_postal, lead\_unit):  
        def \_excl(pool):  
            return \[i for i in pool  
                    if str(\_sv(i\[3\], grid\_c) or "") \!= lead\_grid\]  
        best\_sc, best\_row, best\_raw \= 0.0, None, ""  
        if lead\_postal:  
            if lead\_unit:  
                um \= crm\_postal\_unit\_dict.get(f"{lead\_postal}|{lead\_unit}", \[\])  
                pool   \= \_excl(um) if um else \_excl(crm\_postal\_all\_dict.get(lead\_postal,\[\]))  
                min\_sc \= p2\_threshold if um else NO\_UNIT\_MIN  
            else:  
                pool   \= \_excl(crm\_postal\_all\_dict.get(lead\_postal, \[\]))  
                min\_sc \= NO\_UNIT\_MIN  
            for item in pool:  
                sc, raw\_n, cand \= \_sc(lead\_lat, lead\_pin, item)  
                if sc \> best\_sc:  
                    best\_sc, best\_row, best\_raw \= sc, cand, raw\_n  
            if best\_row is not None and best\_sc \>= min\_sc:  
                return best\_row, best\_sc, f"Postal+Name \[{lead\_postal}\] score={best\_sc:.2f}"  
        else:  
            for item in \_excl(all\_crm\_items):  
                sc, raw\_n, cand \= \_sc(lead\_lat, lead\_pin, item)  
                if sc \>= 0.90 and sc \> best\_sc:  
                    best\_sc, best\_row, best\_raw \= sc, cand, raw\_n  
            if best\_row is not None:  
                return best\_row, best\_sc, f"Name-only (zero postal) score={best\_sc:.2f}"  
        return None, 0.0, ""

    results \= \[\]

    for \_, row in sampled\_df.iterrows():  
        lead\_grid   \= str(row.get("GRID","") or "").strip()  
        lead\_src    \= str(row.get("Lead Source","") or "")  
        lead\_stat   \= str(row.get("Lead Status","") or "")  
        lost\_rsn    \= str(row.get("Lost Reason","") or "") \\  
                      if pd.notna(row.get("Lost Reason")) else ""  
        edited\_by   \= str(row.get("Edited By","") or "")  
        edit\_date   \= row.get("Edit Date")  
        street      \= str(row.get("Street","") or "")  
        postal\_raw  \= str(row.get("Zip/Postal Code","") or "")  
        dup\_id\_agent \= str(row.get("Duplicate Id","") or "") \\  
                       if pd.notna(row.get("Duplicate Id")) else ""

        lead\_postal \= extract\_sg\_postal(\_norm\_postal\_input(postal\_raw)) \\  
                      or extract\_sg\_postal(street)  
        lead\_unit   \= extract\_sg\_unit(street)

        try:  
            edit\_dt \= pd.to\_datetime(str(edit\_date), dayfirst=True, errors="coerce")  
        except Exception:  
            edit\_dt \= None

        \# ── Safe scalar extractor (alias for \_sv defined above) ───  
        \_gs \= \_sv

        \# ── Apify row ──────────────────────────────────────────────  
        apy \= apify\_dict.get(lead\_grid)  
        gm\_title \= gm\_cat \= gm\_biz\_status \= gm\_phone \= gm\_web \= ""  
        perm \= temp \= False  
        if apy is not None:  
            gm\_title \= str(\_gs(apy, \_apy\_title, ""))  
            gm\_cat   \= str(\_gs(apy, \_apy\_cat,   ""))  
            gm\_phone \= str(\_gs(apy, \_apy\_phone,  ""))  
            gm\_web   \= str(\_gs(apy, \_apy\_web,    ""))  
            perm\_v   \= \_gs(apy, \_apy\_perm, False)  
            temp\_v   \= \_gs(apy, \_apy\_temp, False)  
            perm     \= str(perm\_v).lower() in ("true","1","yes")  
            temp     \= str(temp\_v).lower() in ("true","1","yes")  
            gm\_biz\_status \= ("Permanently Closed" if perm else  
                             "Temporarily Closed"  if temp else "Open")

        \# ── Account details (converted only) ──────────────────────  
        acc \= account\_dict.get(lead\_grid) if lead\_stat \== "Converted" else None

        \# Lead name: from account (converted) or Apify title  
        an\_col    \= account\_cols.get("name") if account\_cols else None  
        lead\_name \= (str(\_gs(acc, an\_col) or "") if acc is not None and an\_col else "") or gm\_title  
        lead\_lat\_n, lead\_pin\_n \= norm\_name\_sg(strip\_venue\_generic(lead\_name), char\_map)

        \# Restaurant name for output:  
        \# Converted → Account Name from Account Details report  
        \# Lost      → Company column from lead report  
        if lead\_stat \== "Converted":  
            restaurant\_name \= str(\_gs(acc, an\_col) or "") if acc is not None and an\_col else ""  
        else:  
            \# Try both column names (XLS vs CSV export formats)  
            restaurant\_name \= (str(row.get("Company","") or "")  
                               or str(row.get("Company / Account","") or ""))

        \# ── Run checks ────────────────────────────────────────────  
        C \= {}

        \# C01 — Valid restaurant  
        if apy is None:  
            C\["C01 Valid Restaurant"\] \= "⚠️ No Apify result — unverifiable"  
        elif perm or temp:  
            C\["C01 Valid Restaurant"\] \= \\  
                f"❌ {'Permanently' if perm else 'Temporarily'} closed on Google Maps"  
        elif not gm\_cat:  
            C\["C01 Valid Restaurant"\] \= "⚠️ No category in Apify — unverifiable"  
        elif not is\_food\_delivery\_eligible(gm\_cat, \_DEFAULT\_EXCLUSION\_KW):  
            C\["C01 Valid Restaurant"\] \= f"❌ Non-food category: {gm\_cat}"  
        else:  
            C\["C01 Valid Restaurant"\] \= f"✅ Pass — {gm\_cat}"

        \# C03 — Dedup (needed for C02 and C05)  
        dup\_row, dup\_sc, dup\_meth \= \_find\_dup(  
            lead\_grid, lead\_lat\_n, lead\_pin\_n, lead\_postal, lead\_unit)  
        dup\_name\_found  \= str(\_gs(dup\_row, name\_c) or "") if dup\_row is not None and name\_c else ""  
        dup\_grid\_found  \= str(\_gs(dup\_row, grid\_c) or "") if dup\_row is not None and grid\_c else ""  
        dup\_created\_str \= ""  
        if dup\_row is not None:  
            for dc in \["Created Date","created\_date","CreatedDate"\]:  
                v \= \_gs(dup\_row, dc)  
                if v and str(v) not in ("","nan"):  
                    dup\_created\_str \= str(v); break

        if not lead\_name:  
            C\["C03 Duplicate"\] \= "⚠️ No lead name — dedup unverifiable"  
        elif dup\_row is not None:  
            C\["C03 Duplicate"\] \= \\  
                f"⚠️ Potential duplicate: {dup\_name\_found} ({dup\_grid\_found}) score={dup\_sc:.2f}"  
        else:  
            C\["C03 Duplicate"\] \= "✅ No duplicate found"

        \# C02 — Lost reason (lost only)  
        if lead\_stat \!= "Lost":  
            C\["C02 Lost Reason"\] \= "N/A — Converted"  
        elif lost\_rsn \== "Other":  
            C\["C02 Lost Reason"\] \= "⚠️ Unverifiable — likely test account"  
        elif lost\_rsn \== "Closed Down":  
            if apy is None:  
                C\["C02 Lost Reason"\] \= "⚠️ No Apify result — unverifiable"  
            elif perm or temp:  
                C\["C02 Lost Reason"\] \= "✅ Pass — Google confirms closed"  
            else:  
                C\["C02 Lost Reason"\] \= f"❌ Restaurant appears open on Google ({gm\_biz\_status})"  
        elif lost\_rsn \== "Wrong Target Group":  
            if apy is None:  
                C\["C02 Lost Reason"\] \= "⚠️ No Apify result — unverifiable"  
            elif not gm\_cat:  
                C\["C02 Lost Reason"\] \= "⚠️ No Apify category — unverifiable"  
            elif not is\_food\_delivery\_eligible(gm\_cat, \_DEFAULT\_EXCLUSION\_KW):  
                C\["C02 Lost Reason"\] \= f"✅ Pass — confirmed non-food: {gm\_cat}"  
            else:  
                C\["C02 Lost Reason"\] \= f"❌ Category appears food-eligible: {gm\_cat}"  
        elif lost\_rsn \== "Invalid Data":  
            if apy is None:  
                C\["C02 Lost Reason"\] \= "✅ Pass — no Google listing found (flag for manual check)"  
            else:  
                C\["C02 Lost Reason"\] \= \\  
                    f"⚠️ Has Google listing ({gm\_title}) — verify if truly no online presence"  
        elif lost\_rsn \== "No Delivery Service":  
            if not lead\_postal:  
                C\["C02 Lost Reason"\] \= "⚠️ No postal — zone check unverifiable"  
            elif not zones:  
                C\["C02 Lost Reason"\] \= f"⚠️ No zone data — verify postal {lead\_postal} manually"  
            else:  
                \# Try to check via geocoding if available  
                try:  
                    token \= get\_onemap\_token()  
                    coords \= geocode\_postal\_sg(lead\_postal, token) if token else None  
                    if coords:  
                        lat, lng \= coords  
                        z\_stat, z\_name, z\_city, \_ \= classify\_zone(lat, lng, zones)  
                        if z\_stat \== "Outside Zone":  
                            C\["C02 Lost Reason"\] \= f"✅ Pass — outside delivery zones"  
                        else:  
                            C\["C02 Lost Reason"\] \= \\  
                                f"❌ Postal {lead\_postal} is within zone '{z\_name}'"  
                    else:  
                        C\["C02 Lost Reason"\] \= \\  
                            f"⚠️ Could not geocode postal {lead\_postal} — verify manually"  
                except Exception:  
                    C\["C02 Lost Reason"\] \= \\  
                        f"⚠️ Zone check failed — verify postal {lead\_postal} manually"  
        elif lost\_rsn \== "Duplicate":  
            if dup\_row is not None:  
                C\["C02 Lost Reason"\] \= \\  
                    f"✅ Pass — duplicate confirmed: {dup\_name\_found} ({dup\_grid\_found})"  
                \# Update C03  
                C\["C03 Duplicate"\] \= \\  
                    f"✅ Confirmed duplicate: {dup\_name\_found} ({dup\_grid\_found}) score={dup\_sc:.2f}"  
            else:  
                C\["C02 Lost Reason"\] \= "❌ Lost as Duplicate but no duplicate found by tool"  
                C\["C03 Duplicate"\]   \= "❌ No duplicate found — lost reason may be incorrect"  
        else:  
            C\["C02 Lost Reason"\] \= f"⚠️ Unknown lost reason: {lost\_rsn}"

        \# C04 — Agent's Duplicate ID accuracy  
        if lead\_stat \!= "Lost" or lost\_rsn \!= "Duplicate":  
            C\["C04 Duplicate ID Accuracy"\] \= "N/A"  
        elif not dup\_id\_agent:  
            C\["C04 Duplicate ID Accuracy"\] \= "❌ Agent did not enter Duplicate ID"  
        elif dup\_grid\_found and dup\_id\_agent.strip() \== dup\_grid\_found.strip():  
            C\["C04 Duplicate ID Accuracy"\] \= f"✅ Match — agent: {dup\_id\_agent}"  
        elif dup\_grid\_found:  
            C\["C04 Duplicate ID Accuracy"\] \= \\  
                f"⚠️ Mismatch — agent: {dup\_id\_agent} | tool: {dup\_grid\_found} ({dup\_name\_found})"  
        else:  
            C\["C04 Duplicate ID Accuracy"\] \= \\  
                f"⚠️ No tool match to compare — agent entered: {dup\_id\_agent}"

        \# C05 — Wrongful conversion  
        if lead\_stat \!= "Converted":  
            C\["C05 Wrongful Conversion"\] \= "N/A — Lost lead"  
        elif dup\_row is None:  
            C\["C05 Wrongful Conversion"\] \= "✅ No duplicate found"  
        else:  
            try:  
                dup\_created\_dt \= pd.to\_datetime(dup\_created\_str, dayfirst=True, errors="coerce")  
            except Exception:  
                dup\_created\_dt \= None  
            if dup\_created\_dt is None or edit\_dt is None:  
                C\["C05 Wrongful Conversion"\] \= \\  
                    f"⚠️ Duplicate found ({dup\_name\_found}) — cannot compare dates, manual review"  
            elif dup\_created\_dt.date() \<= edit\_dt.date():  
                C\["C05 Wrongful Conversion"\] \= \\  
                    f"❌ Duplicate ({dup\_name\_found}) existed before conversion " \\  
                    f"\[dup created: {dup\_created\_dt.date()}, lead converted: {edit\_dt.date()}\]"  
            else:  
                C\["C05 Wrongful Conversion"\] \= \\  
                    f"✅ Duplicate created after conversion — not wrongful " \\  
                    f"\[dup: {dup\_created\_dt.date()}, converted: {edit\_dt.date()}\]"

        \# C06 — Naming convention (converted only)  
        if lead\_stat \!= "Converted":  
            C\["C06 Naming Convention"\] \= "N/A"  
        elif not lead\_name:  
            C\["C06 Naming Convention"\] \= "⚠️ No account name available"  
        else:  
            ok, detail \= \_check\_naming\_convention(lead\_name)  
            C\["C06 Naming Convention"\] \= "✅ Pass" if ok else f"❌ {detail}"

        \# Helpers for account detail checks  
        def \_acc\_check(key, label):  
            """Returns check string for a simple populated/blank field."""  
            if lead\_stat \!= "Converted": return "N/A"  
            if acc is None: return "⚠️ No account details uploaded"  
            col \= account\_cols.get(key)  
            val \= str(\_gs(acc, col) or "") if col else ""  
            if not val or val.lower() \== "nan":  
                return f"❌ {label} not populated"  
            return f"✅ Populated: {val}"

        \# C07 — Phone  
        if lead\_stat \!= "Converted":  
            C\["C07 Phone"\] \= "N/A"  
        elif acc is None:  
            C\["C07 Phone"\] \= "⚠️ No account details uploaded"  
        else:  
            ph\_col  \= account\_cols.get("phone")  
            sf\_ph   \= str(\_gs(acc, ph\_col) or "") if ph\_col else ""  
            if not sf\_ph or sf\_ph.lower() \== "nan":  
                C\["C07 Phone"\] \= "❌ Phone not populated in Salesforce"  
            elif gm\_phone:  
                sf\_n \= re.sub(r'\\D','', sf\_ph)  
                gm\_n \= re.sub(r'\\D','', gm\_phone)  
                if sf\_n \== gm\_n:  
                    C\["C07 Phone"\] \= f"✅ Match — {sf\_ph}"  
                else:  
                    C\["C07 Phone"\] \= \\  
                        f"⚠️ Mismatch — SF: {sf\_ph} | Google: {gm\_phone}"  
            else:  
                C\["C07 Phone"\] \= f"✅ Populated: {sf\_ph} (no Apify phone to compare)"

        C\["C08 Email"\]          \= \_acc\_check("email",          "Email")  
        C\["C10 Social Media"\]   \= \_acc\_check("social\_media",   "Social Media URL")  
        C\["C11 Parent Account"\] \= \_acc\_check("parent\_account", "Parent Account")  
        C\["C12 Business Office"\]= \_acc\_check("business\_office","Business Office")  
        C\["C14 Target Partner"\] \= \_acc\_check("target\_partner", "Target Partner")

        \# C09 — Website  
        if lead\_stat \!= "Converted":  
            C\["C09 Website"\] \= "N/A"  
        elif acc is None:  
            C\["C09 Website"\] \= "⚠️ No account details uploaded"  
        else:  
            ws\_col \= account\_cols.get("website")  
            sf\_ws  \= str(\_gs(acc, ws\_col) or "") if ws\_col else ""  
            if not sf\_ws or sf\_ws.lower() \== "nan":  
                C\["C09 Website"\] \= "❌ Website not populated"  
            elif gm\_web:  
                def \_nu(u):  
                    return re.sub(r'^https?://(www\\.)?','',  
                                  str(u).lower().strip()).rstrip('/')  
                if \_nu(sf\_ws) \== \_nu(gm\_web):  
                    C\["C09 Website"\] \= f"✅ Match — {sf\_ws}"  
                else:  
                    C\["C09 Website"\] \= \\  
                        f"⚠️ Mismatch — SF: {sf\_ws} | Google: {gm\_web}"  
            else:  
                C\["C09 Website"\] \= f"✅ Populated: {sf\_ws}"

        \# C13 — Delivery service  
        if lead\_stat \!= "Converted":  
            C\["C13 Delivery Service"\] \= "N/A"  
        elif acc is None:  
            C\["C13 Delivery Service"\] \= "⚠️ No account details uploaded"  
        else:  
            ds\_col \= account\_cols.get("delivery\_service")  
            sf\_ds  \= str(\_gs(acc, ds\_col) or "") if ds\_col else ""  
            if not sf\_ds or sf\_ds.lower() \== "nan":  
                C\["C13 Delivery Service"\] \= "❌ Delivery Service not populated"  
            else:  
                dsl  \= sf\_ds.lower()  
                h\_dh \= "dh delivery" in dsl  
                h\_ta \= "take away" in dsl or "takeaway" in dsl  
                if h\_dh and h\_ta:  
                    C\["C13 Delivery Service"\] \= f"✅ Pass — {sf\_ds}"  
                else:  
                    missing \= \[\]  
                    if not h\_dh: missing.append("DH Delivery")  
                    if not h\_ta: missing.append("Take Away")  
                    C\["C13 Delivery Service"\] \= \\  
                        f"❌ Missing: {', '.join(missing)} — current: {sf\_ds}"

        \# C15 — Restaurant category  
        if lead\_stat \!= "Converted":  
            C\["C15 Restaurant Category"\] \= "N/A"  
        elif acc is None:  
            C\["C15 Restaurant Category"\] \= "⚠️ No account details uploaded"  
        else:  
            cat\_col \= account\_cols.get("category")  
            sf\_cat  \= str(\_gs(acc, cat\_col) or "") if cat\_col else ""  
            if not sf\_cat or sf\_cat.lower() \== "nan":  
                C\["C15 Restaurant Category"\] \= "❌ Category not populated"  
            elif not gm\_cat:  
                C\["C15 Restaurant Category"\] \= \\  
                    f"⚠️ Cannot verify — no Apify category. SF: {sf\_cat}"  
            else:  
                mapped \= \_map\_apify\_cuisine(gm\_cat)  
                if not mapped:  
                    C\["C15 Restaurant Category"\] \= \\  
                        f"⚠️ Apify category '{gm\_cat}' unmappable — SF: {sf\_cat}"  
                elif mapped.lower() \== sf\_cat.lower():  
                    C\["C15 Restaurant Category"\] \= f"✅ Match — {sf\_cat}"  
                else:  
                    C\["C15 Restaurant Category"\] \= \\  
                        f"⚠️ Possible mismatch — SF: {sf\_cat} | Apify suggests: {mapped}"

        auto\_errs \= sum(1 for v in C.values() if str(v).startswith("❌"))

        results.append({  
            "GRID":              lead\_grid,  
            "Restaurant Name":   restaurant\_name,  
            "Lead Source":       lead\_src,  
            "Lead Status":     lead\_stat,  
            "Lost Reason":     lost\_rsn,  
            "Edited By":       edited\_by,  
            "Edit Date":       edit\_date,  
            "Street":          street,  
            "Zip/Postal Code": \_norm\_postal\_input(postal\_raw) or str(postal\_raw).replace(".0",""),  
            "GM Title":        gm\_title,  
            "GM Category":     gm\_cat,  
            "GM Status":       gm\_biz\_status,  
            \*\*C,  
            "Auto Error Count":      auto\_errs,  
            "Ops Feedback":          "",  
            "Updated Error Count":   "",  
            "Agent Acknowledgement": "",  
            "Agent Feedback":        "",  
        })

    results\_df \= pd.DataFrame(results)

    \# Agent summary  
    agg \= results\_df.groupby("Edited By", dropna=False).agg(  
        Leads\_Sampled=("GRID","count"),  
        Auto\_Errors=("Auto Error Count","sum")  
    ).reset\_index()  
    agg.columns \= \["Agent","Leads Sampled","Auto Error Count"\]  
    agg\["Updated Error Count"\] \= ""

    return results\_df, agg

def build\_kpi\_excel(results\_df: pd.DataFrame,  
                    agent\_df: pd.DataFrame,  
                    sampling\_df: pd.DataFrame) \-\> bytes:  
    """Build KPI scorecard Excel with 3 sheets."""  
    import io as \_io  
    from openpyxl import Workbook  
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side  
    from openpyxl.utils import get\_column\_letter

    wb  \= Workbook()  
    FNT \= "Poppins"

    CHECK\_COLS \= \[c for c in results\_df.columns  
                  if c.startswith("C") and c\[1:3\].isdigit()\]  
    AGENT\_COLS\_OUT \= \[  
        "Auto Error Count","Ops Feedback","Updated Error Count",  
        "Agent Acknowledgement","Agent Feedback",  
    \]  
    INFO\_COLS \= \["GRID","Restaurant Name","Lead Source","Lead Status","Lost Reason",  
                 "Edited By","Edit Date","Street","Zip/Postal Code",  
                 "GM Title","GM Category","GM Status"\]  
    ALL\_COLS  \= INFO\_COLS \+ CHECK\_COLS \+ AGENT\_COLS\_OUT

    \# Pre-compute check column letter range for COUNTIF formula  
    \_chk\_s \= get\_column\_letter(len(INFO\_COLS) \+ 1\)  
    \_chk\_e \= get\_column\_letter(len(INFO\_COLS) \+ len(CHECK\_COLS))

    def \_thin():  
        s \= Side(style="thin", color="D0D0D0")  
        return Border(left=s, right=s, top=s, bottom=s)

    HDR \= PatternFill("solid", start\_color="1A1A2E")  
    AGT \= PatternFill("solid", start\_color="EEF2FF")

    FILLS \= {  
        "✅": PatternFill("solid", start\_color="C6EFCE"),  
        "❌": PatternFill("solid", start\_color="FFC7CE"),  
        "⚠️": PatternFill("solid", start\_color="FFF2CC"),  
        "N/A": PatternFill("solid", start\_color="F0F0F0"),  
    }  
    FONTS \= {  
        "✅":"276221","❌":"9C0006","⚠️":"7D5A00","N/A":"595959",  
    }

    def \_cell\_fill(val):  
        v \= str(val)  
        for k in \["✅","❌","⚠️","N/A"\]:  
            if v.startswith(k):  
                return FILLS\[k\], FONTS\[k\]  
        return PatternFill(), "000000"

    \# ── Sheet 1: Sampled Leads ─────────────────────────────────────  
    ws1 \= wb.active; ws1.title \= "Sampled Leads"  
    ws1\["A1"\] \= f"KPI Sample Checker — {pd.Timestamp.now().strftime('%B %Y')}"  
    ws1\["A1"\].font \= Font(name=FNT, bold=True, size=14, color="1F4E79")  
    ws1.merge\_cells(f"A1:{get\_column\_letter(len(ALL\_COLS))}1")  
    ws1.row\_dimensions\[1\].height \= 24

    for ci, h in enumerate(ALL\_COLS, 1):  
        c \= ws1.cell(row=2, column=ci, value=h)  
        c.font      \= Font(name=FNT, bold=True, color="FFFFFF", size=9)  
        c.fill      \= HDR  
        c.border    \= \_thin()  
        c.alignment \= Alignment(horizontal="center", vertical="center", wrap\_text=True)  
    ws1.row\_dimensions\[2\].height \= 30

    for ri, (\_, row) in enumerate(results\_df\[ALL\_COLS\].iterrows(), 3):  
        ws1.row\_dimensions\[ri\].height \= 18  
        for ci, col in enumerate(ALL\_COLS, 1):  
            val \= row\[col\]; val \= "" if pd.isna(val) else val  
            c   \= ws1.cell(row=ri, column=ci, value=val)  
            c.border \= \_thin()  
            c.font   \= Font(name=FNT, size=8)

            if col in CHECK\_COLS:  
                fill, fc \= \_cell\_fill(val)  
                c.fill      \= fill  
                c.font      \= Font(name=FNT, size=8, color=fc)  
                c.alignment \= Alignment(vertical="center", wrap\_text=True)  
            elif col in AGENT\_COLS\_OUT:  
                c.fill      \= AGT  
                c.alignment \= Alignment(vertical="center")  
            elif col \== "Auto Error Count":  
                \# Dynamic COUNTIF formula — auto-updates when check cells are edited  
                c.value \= f'=COUNTIF({\_chk\_s}{ri}:{\_chk\_e}{ri},"❌\*")'  
                c.font      \= Font(name=FNT, size=8, bold=True, color="9C0006")  
                c.alignment \= Alignment(horizontal="center", vertical="center")  
            else:  
                c.alignment \= Alignment(vertical="center")

    \# Column widths  
    widths \= {  
        "GRID":14, "Restaurant Name":32, "Lead Source":16,  
        "Lead Status":12, "Lost Reason":18,  
        "Edited By":22, "Edit Date":16, "Street":36, "Zip/Postal Code":12,  
        "GM Title":28, "GM Category":22, "GM Status":18,  
        "Auto Error Count":12, "Ops Feedback":24, "Updated Error Count":14,  
        "Agent Acknowledgement":20, "Agent Feedback":28,  
    }  
    for ci, col in enumerate(ALL\_COLS, 1):  
        w \= widths.get(col, 42 if col in CHECK\_COLS else 14\)  
        ws1.column\_dimensions\[get\_column\_letter(ci)\].width \= w  
    ws1.freeze\_panes \= "A3"  
    ws1.auto\_filter.ref \= f"A2:{get\_column\_letter(len(ALL\_COLS))}{2+len(results\_df)}"

    \# ── Sheet 2: Agent Summary ─────────────────────────────────────  
    ws2 \= wb.create\_sheet("Agent Summary")  
    ws2\["A1"\] \= "Agent KPI Summary"  
    ws2\["A1"\].font \= Font(name=FNT, bold=True, size=13, color="1F4E79")  
    ws2.merge\_cells("A1:D1")  
    heads2 \= list(agent\_df.columns)  
    for ci, h in enumerate(heads2, 1):  
        c \= ws2.cell(row=2, column=ci, value=h)  
        c.font \= Font(name=FNT, bold=True, color="FFFFFF", size=9)  
        c.fill \= HDR; c.border \= \_thin()  
        c.alignment \= Alignment(horizontal="center", vertical="center")  
    for ri, (\_, r) in enumerate(agent\_df.iterrows(), 3):  
        for ci, col in enumerate(heads2, 1):  
            c \= ws2.cell(row=ri, column=ci, value=r\[col\])  
            c.font \= Font(name=FNT, size=8); c.border \= \_thin()  
            c.alignment \= Alignment(horizontal="center", vertical="center")  
    for ci, w in enumerate(\[28,14,16,16\], 1):  
        ws2.column\_dimensions\[get\_column\_letter(ci)\].width \= w

    \# ── Sheet 3: Sampling Breakdown ────────────────────────────────  
    ws3 \= wb.create\_sheet("Sampling Breakdown")  
    ws3\["A1"\] \= "Sampling Breakdown (10% per Agent × Source × Status)"  
    ws3\["A1"\].font \= Font(name=FNT, bold=True, size=13, color="1F4E79")  
    ws3.merge\_cells(f"A1:{get\_column\_letter(len(sampling\_df.columns))}1")  
    for ci, h in enumerate(sampling\_df.columns, 1):  
        c \= ws3.cell(row=2, column=ci, value=h)  
        c.font \= Font(name=FNT, bold=True, color="FFFFFF", size=9)  
        c.fill \= HDR; c.border \= \_thin()  
        c.alignment \= Alignment(horizontal="center", vertical="center")  
    for ri, (\_, r) in enumerate(sampling\_df.iterrows(), 3):  
        for ci, col in enumerate(sampling\_df.columns, 1):  
            c \= ws3.cell(row=ri, column=ci, value=r\[col\])  
            c.font \= Font(name=FNT, size=8); c.border \= \_thin()  
            c.alignment \= Alignment(horizontal="center", vertical="center")  
    for ci, w in enumerate(\[24,18,14,12,10\], 1):  
        ws3.column\_dimensions\[get\_column\_letter(ci)\].width \= w

    buf \= \_io.BytesIO()  
    wb.save(buf)  
    return buf.getvalue()

def main():  
    if not check\_password():  
        return

    st.markdown("""  
    \<style\>  
    html,body,\[class\*="css"\]{font-family:Arial,sans-serif;}  
    header\[data-testid="stHeader"\]{background:\#FFF;border-bottom:2px solid rgba(223,16,103,.25);}  
    \[data-testid="stSidebar"\]{background:\#FAFAFA;border-right:2px solid rgba(223,16,103,.3);}  
    \[data-testid="stSidebar"\] h1,\[data-testid="stSidebar"\] h2,  
    \[data-testid="stSidebar"\] h3{color:\#DF1067\!important;font-weight:700\!important;}  
    \[data-testid="stButton"\]\>button\[kind="primary"\]{  
        background:\#DF1067\!important;border:none\!important;color:\#fff\!important;  
        font-weight:600\!important;font-size:1rem\!important;  
        border-radius:8px\!important;padding:.6rem 1.5rem\!important;}  
    \[data-testid="stButton"\]\>button\[kind="primary"\]:hover{background:\#C00055\!important;}  
    \[data-testid="stDownloadButton"\]\>button{  
        background:\#DF1067\!important;border:none\!important;color:\#fff\!important;  
        font-weight:600\!important;border-radius:8px\!important;width:100%\!important;}  
    \[data-testid="stMetric"\]{background:\#FFF;border:1px solid \#EBEBEB;  
        border-top:3px solid \#DF1067;border-radius:8px;  
        padding:.9rem 1rem;box-shadow:0 1px 4px rgba(0,0,0,.06);}  
    \</style\>""", unsafe\_allow\_html=True)

    if not \_PYPINYIN\_AVAILABLE:  
        st.warning("⚠️ \*\*pypinyin\*\* not installed — Chinese name → Pinyin matching is off. "  
                   "Run \`pip install pypinyin\` and restart.")

    \# ── Sidebar ───────────────────────────────────────────────────  
    with st.sidebar:  
        st.markdown(  
            f'\<div style="text-align:center;padding:.5rem 0 .8rem;"\>'  
            f'\<img src="data:image/png;base64,{DH\_LOGO\_B64}" style="width:110px;"/\>\</div\>',  
            unsafe\_allow\_html=True)  
        st.header("⚙️ Settings")

        market\_code \= st.selectbox(  
            "Market",  
            options=\[None\] \+ list(MARKETS.keys()),  
            format\_func=lambda k: (  
                "— Select market —" if k is None  
                else f"{MARKETS\[k\]\['flag'\]} {MARKETS\[k\]\['name'\]} ({k})"))  
        market\_cfg \= MARKETS.get(market\_code)

        st.divider()  
        st.subheader("🎚 Match Thresholds")  
        p2\_threshold \= st.slider("P3 Potential Match starts at", 0.30, 0.65, 0.50, 0.05,  
                                  help="Name similarity at same postal ≥ this → P3 Potential Match")  
        p3\_threshold \= st.slider("P4 Duplicate starts at",  
                                  float(round(p2\_threshold \+ 0.05, 2)), 0.95, 0.75, 0.05,  
                                  help="Name similarity at same postal ≥ this → P4 Duplicate")  
        st.caption(f"P3 threshold: {p2\_threshold:.0%}  ·  P4 threshold: {p3\_threshold:.0%}")

        st.divider()  
        st.subheader("🚫 Exclusion Keywords")  
        st.caption("Leads whose Google Maps category contains any of these → Wrong Target Group.")  
        kw\_input \= st.text\_area("Keywords (one per line)",  
                                 value="\\n".join(\_DEFAULT\_EXCLUSION\_KW), height=180)  
        exclusion\_kw \= \[k.strip().lower() for k in kw\_input.split("\\n") if k.strip()\]  
        st.caption(f"{len(exclusion\_kw)} keywords active")

        st.divider()  
        with st.expander("🎯 Label reference"):  
            st.markdown("""  
| Label | What it means | Action |  
|---|---|---|  
| \*\*P1 — New\*\* | No CRM match. Google Maps confirms it's a restaurant. | ✅ Pitch delivery |  
| \*\*P3 — Potential Match\*\* | Name 50–74% similar to a CRM account at same postal. | 🔍 Verify before pitching |  
| \*\*P4 — Duplicate\*\* | Name ≥ 75% similar to a CRM account at same postal. | ❌ Skip — already in system |  
| \*\*Business Closed\*\* | Google Maps shows permanently or temporarily closed. | ❌ Skip |  
| \*\*Wrong Target Group\*\* | Google Maps category is not food delivery eligible. | ❌ Skip |  
| \*\*P2 — Please Check\*\* | No Google Maps result found, or result is unreliable. | ⚪ Manual verification needed |  
            """)

    if not market\_cfg:  
        st.info("👈 Select a market from the sidebar to get started.")  
        return

    tab1, tab2, tab3, tab4, tab5, tab6 \= st.tabs(\[  
        "📊 Classify Leads",  
        "🔗 Generate Apify URLs",  
        "🏢 SF Account Audit",  
        "🔍 CRM Check",  
        "📋 KPI Sample Checker",  
        "📖 How to Use",  
    \])

    \# ════════════════════════════════════════════════════════════════  
    \# TAB 1 — LEAD CLASSIFICATION  
    \# ════════════════════════════════════════════════════════════════  
    with tab1:  
        with st.expander("📎 How to get your files — click to expand"):  
            c1, c2 \= st.columns(2)  
            with c1:  
                st.markdown("\*\*Step 1 · Leads file (Salesforce)\*\*")  
                st.link\_button(  
                    "Open Leads Report →",  
                    "https://deliveryhero.lightning.force.com/"  
                    "lightning/r/Report/00ObO000006oALhUAM/view",  
                    use\_container\_width=True)  
            with c2:  
                st.markdown("\*\*Step 2 · CRM Export (Salesforce)\*\*")  
                st.link\_button(  
                    "Open Singapore CRM Report →",  
                    "https://deliveryhero.lightning.force.com/lightning/r/Report/"  
                    "00ObO000005IE85UAG/view?queryScope=userFolders",  
                    use\_container\_width=True)  
            st.info(  
                "🔗 \*\*For Apify Results:\*\* go to the \*\*Generate Apify URLs\*\* tab → "  
                "Step 1 generates your URLs → paste into Apify → "  
                "Step 2 adds the GRID column to your Apify export automatically.")

        st.divider()  
        col1, col2, col3 \= st.columns(3)  
        with col1:  
            st.subheader("1. Leads")  
            leads\_up \= st.file\_uploader("Upload leads (.xlsx or .csv)",  
                                         type=\["xlsx","xls","csv"\], key="leads")  
        with col2:  
            st.subheader("2. Apify Results")  
            apify\_up \= st.file\_uploader("Upload Apify output (.csv or .xlsx)",  
                                         type=\["csv","xlsx","xls"\], key="apify")  
        with col3:  
            st.subheader("3. CRM Export")  
            crm\_up   \= st.file\_uploader("Upload SF CRM export (.csv or .xlsx)",  
                                         type=\["csv","xlsx","xls"\], key="crm")

        st.divider()

        import os as \_os  
        \_builtin \= \_os.path.exists(\_os.path.join(  
            \_os.path.dirname(\_os.path.abspath(\_\_file\_\_)),  
            f"zones\_{market\_code}.json"))

        if \_builtin:  
            st.markdown(f"\*\*📍 Zones:\*\* Built-in zone data for \*\*{market\_cfg\['flag'\]} {market\_cfg\['name'\]}\*\* loaded.")  
            zone\_up \= st.file\_uploader("Upload custom zone file to override (optional)",  
                                        type=\["csv","xlsx","xls"\], key="zones")  
        else:  
            zone\_up \= st.file\_uploader("Upload delivery zone file (.csv / .xlsx) — optional",  
                                        type=\["csv","xlsx","xls"\], key="zones")  
        geocode\_on \= st.toggle("Geocode leads without coordinates", value=True,  
                                help="Uses OneMap SLA postal code API for sub-metre accuracy.")

        leads\_df \= col\_map\_leads \= None  
        crm\_df   \= col\_map\_crm   \= None  
        apify\_df \= col\_map\_apify \= None

        if leads\_up:  
            try:  
                leads\_df, col\_map\_leads \= load\_leads(  
                    leads\_up.read(), leads\_up.name, market\_cfg)  
                st.success(f"Leads loaded: {len(leads\_df):,} rows")  
                col\_map\_leads \= \_col\_map\_ui(  
                    leads\_df, col\_map\_leads,  
                    fields=\[  
                        ("name",    "Restaurant name",    True),  
                        ("street",  "Street / address",   False),  
                        ("zip",     "Postal code",        False),  
                        ("grid",    "GRID",               False),  
                        ("lead\_id", "Lead ID",            False),  
                        ("lat",     "Latitude",           False),  
                        ("lng",     "Longitude",          False),  
                    \],  
                    key\_prefix="t1\_leads")  
            except Exception as e:  
                st.error(f"Error loading leads: {e}")

        if crm\_up:  
            try:  
                crm\_df, col\_map\_crm \= load\_crm(  
                    crm\_up.read(), crm\_up.name, market\_cfg)  
                st.success(f"CRM loaded: {len(crm\_df):,} accounts")  
            except Exception as e:  
                st.error(f"Error loading CRM: {e}")

        if apify\_up:  
            try:  
                apify\_df, col\_map\_apify \= load\_apify(  
                    apify\_up.read(), apify\_up.name)  
                st.success(f"Apify loaded: {len(apify\_df):,} rows")  
            except Exception as e:  
                st.error(f"Error loading Apify: {e}")

        st.divider()

        if leads\_df is not None and st.button(  
                "▶ Run Classification", type="primary", use\_container\_width=True):  
            zones \= \[\]  
            if zone\_up:  
                zones \= load\_zones(zone\_up.read(), zone\_up.name)  
            elif \_builtin:  
                zones \= load\_zones(market\_code=market\_code)

            import time as \_time  
            \_prog\_bar  \= st.progress(0.0)  
            \_prog\_text \= st.empty()  
            \_t0        \= \_time.time()

            def \_progress\_cb(done, total):  
                pct     \= done / total  
                elapsed \= \_time.time() \- \_t0  
                if done \> 10:  
                    eta\_sec \= int((elapsed / done) \* (total \- done))  
                    eta\_str \= f"\~{eta\_sec//60}m {eta\_sec%60:02d}s remaining"  
                else:  
                    eta\_str \= "estimating…"  
                \_prog\_bar.progress(pct)  
                \_prog\_text.caption(  
                    f"Classifying {done:,} / {total:,} leads · {eta\_str}")

            result\_df \= classify\_leads(  
                leads\_df, col\_map\_leads,  
                crm\_df,   col\_map\_crm   or {},  
                apify\_df, col\_map\_apify or {},  
                market\_cfg,  
                p2\_threshold    \= p2\_threshold,  
                p3\_threshold    \= p3\_threshold,  
                exclusion\_kw    \= exclusion\_kw,  
                zones           \= zones,  
                geocode\_enabled \= geocode\_on,  
                progress\_cb     \= \_progress\_cb,  
            )  
            \_prog\_bar.progress(1.0)  
            \_prog\_text.caption(  
                f"Done — {len(result\_df):,} leads in "  
                f"{int(\_time.time()-\_t0)}s")

            counts \= result\_df\["Label"\].value\_counts()  
            st.success(f"✅ Done — {len(result\_df):,} leads classified.")

            \# Metrics row  
            m \= st.columns(6)  
            m\[0\].metric("✅ P1 New",           counts.get("P1 — New", 0))  
            m\[1\].metric("🔴 P4 Duplicate",     counts.get("P4 — Duplicate", 0))  
            m\[2\].metric("🟡 P3 Potential",     counts.get("P3 — Potential Match", 0))  
            m\[3\].metric("🏢 Business Closed",  counts.get("Business Closed", 0))  
            m\[4\].metric("❌ Wrong Target",     counts.get("Wrong Target Group", 0))  
            m\[5\].metric("⚪ Please Check",     counts.get("P2 — Please Check", 0))

            \# Per-priority tabs  
            label\_tabs \= st.tabs(\[  
                f"All ({len(result\_df)})",  
                f"✅ P1 New ({counts.get('P1 — New',0)})",  
                f"🔴 P4 Duplicate ({counts.get('P4 — Duplicate',0)})",  
                f"🟡 P3 Potential ({counts.get('P3 — Potential Match',0)})",  
                f"🏢 Closed ({counts.get('Business Closed',0)})",  
                f"❌ Wrong TG ({counts.get('Wrong Target Group',0)})",  
                f"⚪ Please Check ({counts.get('P2 — Please Check',0)})",  
            \])

            LABEL\_STYLE \= {  
                "P1 — New":            "background-color:\#d4edda",  
                "P3 — Potential Match":"background-color:\#fff3cd",  
                "P4 — Duplicate":      "background-color:\#f8d7da",  
                "Business Closed":     "background-color:\#ffeb9c",  
                "Wrong Target Group":  "background-color:\#ffdca8",  
                "P2 — Please Check":   "background-color:\#e2e3e5",  
            }  
            def \_style(row):  
                return \[LABEL\_STYLE.get(row\["Label"\],"")\] \* len(row)  
            def \_show(dff):  
                if dff.empty: st.info("No leads in this category.")  
                else: st.dataframe(dff.style.apply(\_style, axis=1), use\_container\_width=True)

            with label\_tabs\[0\]: \_show(result\_df)  
            with label\_tabs\[1\]: \_show(result\_df\[result\_df\["Label"\] \== "P1 — New"\])  
            with label\_tabs\[2\]: \_show(result\_df\[result\_df\["Label"\] \== "P4 — Duplicate"\])  
            with label\_tabs\[3\]: \_show(result\_df\[result\_df\["Label"\] \== "P3 — Potential Match"\])  
            with label\_tabs\[4\]: \_show(result\_df\[result\_df\["Label"\] \== "Business Closed"\])  
            with label\_tabs\[5\]: \_show(result\_df\[result\_df\["Label"\] \== "Wrong Target Group"\])  
            with label\_tabs\[6\]: \_show(result\_df\[result\_df\["Label"\] \== "P2 — Please Check"\])

            excel\_buf \= build\_excel(result\_df, f"{market\_cfg\['flag'\]} {market\_cfg\['name'\]}")  
            st.download\_button(  
                "⬇ Download Excel Report",  
                data=excel\_buf,  
                file\_name=f"ALG\_{pd.Timestamp.now().strftime('%d%m%y')}.xlsx",  
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",  
                use\_container\_width=True)

    \# ════════════════════════════════════════════════════════════════  
    \# TAB 2 — GENERATE APIFY URLS  
    \# ════════════════════════════════════════════════════════════════  
    with tab2:  
        st.subheader("Generate Google Maps URLs for Apify")

        \# ── Step 1: Generate URLs ─────────────────────────────────  
        st.markdown("\#\#\#\# Step 1 · Generate URLs")

        url\_mode \= st.radio(  
            "URL format",  
            options=\["text", "coords"\],  
            format\_func=lambda m: (  
                "📝  Company / Account \+ Street \+ Postal"  
                if m \== "text"  
                else "📍  Company / Account \+ Coordinates (Latitude, Longitude)"  
            ),  
            horizontal=False,  
            key="url\_mode",  
        )

        url\_up \= st.file\_uploader("Upload leads file (.xlsx or .csv)",  
                                   type=\["xlsx","xls","csv"\], key="url\_leads")  
        if url\_up and market\_cfg:  
            try:  
                url\_bytes \= url\_up.read()  
                url\_df, url\_col\_map \= load\_leads(url\_bytes, url\_up.name, market\_cfg)  
                urls, reused \= generate\_google\_urls(url\_df, url\_col\_map, market\_cfg,  
                                                    mode=url\_mode)  
                url\_df\["GOOGLE URL"\] \= urls  
                valid \= \[u for u in urls if u\]

                \# Store GRID → norm\_url mapping in session state for Step 2  
                grid\_col\_name \= url\_col\_map.get("grid")  
                if grid\_col\_name:  
                    url\_to\_grid \= {}  
                    for i, (\_, row) in enumerate(url\_df.iterrows()):  
                        if i \< len(urls) and urls\[i\]:  
                            g \= str(row.get(grid\_col\_name, "") or "").strip()  
                            if g:  
                                url\_to\_grid\[norm\_url(urls\[i\])\] \= g  
                    st.session\_state\["url\_to\_grid"\] \= url\_to\_grid

                st.info(f"{len(valid):,} URLs generated ({reused:,} reused from existing column).")  
                st.text\_area("Generated targets", "\\n".join(valid), height=180)  
                buf \= io.StringIO()  
                url\_df.to\_csv(buf, index=False)  
                st.download\_button("⬇ Download leads with URLs (.csv)",  
                                   buf.getvalue(),  
                                   f"leads\_with\_urls\_{market\_code}.csv",  
                                   mime="text/csv", use\_container\_width=True)  
            except Exception as e:  
                st.error(f"Error: {e}")

        st.divider()

        \# ── Step 2: Append GRID to Apify export ──────────────────  
        st.markdown("\#\#\#\# Step 2 · Add GRID to your Apify Export")  
        st.caption(  
            "After running Apify, upload your export here. "  
            "The tool matches each row via \`inputStartUrl\` and adds a \*\*GRID\*\* column.")

        \# Build lookup — prefer session state, fall back to manual URL CSV upload  
        lookup: dict \= dict(st.session\_state.get("url\_to\_grid", {}))

        if lookup:  
            st.success(f"✅ {len(lookup):,} GRIDs ready from Step 1 above.")  
        else:  
            st.info("No URLs generated this session yet. Upload your URL CSV below.")

        with st.expander("📂 Upload URL CSV (if you generated URLs in a previous session)"):  
            st.caption(  
                "Upload the CSV downloaded from Step 1, or your Google Sheet URL extractor. "  
                "The tool auto-detects the header row and the GRID \+ URL columns.")  
            url\_csv\_up \= st.file\_uploader(  
                "Upload leads+URL CSV (.csv or .xlsx)",  
                type=\["csv","xlsx","xls"\], key="url\_csv\_fallback")  
            if url\_csv\_up:  
                try:  
                    fb\_bytes \= url\_csv\_up.read()  
                    from io import StringIO as \_SIO

                    if url\_csv\_up.name.endswith((".xlsx", ".xls")):  
                        \# Excel / HTML-as-XLS — \_cached\_read handles both formats  
                        fb\_df \= \_cached\_read(fb\_bytes, url\_csv\_up.name)  
                        \# Google Sheets exported as XLS often have title rows  
                        \# before the real column headers. If "grid" and "url"  
                        \# aren't in the columns yet, scan the rows for them.  
                        \_has\_grid \= detect\_column(fb\_df, \["GRID","grid","Grid"\])  
                        \_has\_url  \= detect\_column(fb\_df, \[  
                            "GOOGLE URL","Google URL","google\_url",  
                            "URL (Name \+ Address \+ Postal)",  
                            "URL (Name \+ Lat \+ Long)",  
                            "URL (Name \+ Coordinates)","URL","url"\])  
                        if not (\_has\_grid and \_has\_url):  
                            for i, row in fb\_df.iterrows():  
                                vals \= \[str(v).lower() for v in row.tolist()\]  
                                if (any("grid" in v for v in vals)  
                                        and any("url" in v for v in vals)):  
                                    fb\_df.columns \= \[str(c) for c in fb\_df.iloc\[i\].tolist()\]  
                                    fb\_df \= fb\_df.iloc\[i+1:\].reset\_index(drop=True)  
                                    break  
                    else:  
                        \# ── Smart header detection for CSV ────────────────  
                        \# Google Sheet exports have title rows before the real headers.  
                        \# Scan every row to find one that contains both "grid" and "url".  
                        raw\_rows \= None  
                        for enc in ("utf-8","utf-8-sig","windows-1252","latin-1"):  
                            try:  
                                raw\_rows \= pd.read\_csv(  
                                    \_SIO(fb\_bytes.decode(enc)),  
                                    header=None, on\_bad\_lines="skip", engine="python")  
                                break  
                            except Exception:  
                                continue

                        header\_row \= 0  
                        if raw\_rows is not None:  
                            for i, row in raw\_rows.iterrows():  
                                vals \= \[str(v).lower() for v in row.tolist()\]  
                                if any("grid" in v for v in vals) and any("url" in v for v in vals):  
                                    header\_row \= i  
                                    break

                        fb\_df \= None  
                        for enc in ("utf-8","utf-8-sig","windows-1252","latin-1"):  
                            try:  
                                fb\_df \= pd.read\_csv(  
                                    \_SIO(fb\_bytes.decode(enc)),  
                                    header=header\_row, on\_bad\_lines="skip", engine="python")  
                                break  
                            except Exception:  
                                continue  
                        if fb\_df is None:  
                            fb\_df \= \_cached\_read(fb\_bytes, url\_csv\_up.name)

                    \# ── Column detection ───────────────────────────────  
                    \# Covers both tool-generated CSV and Google Sheet formats  
                    grid\_c \= detect\_column(fb\_df, \["GRID","grid","Grid"\])  
                    gurl\_c \= detect\_column(fb\_df, \[  
                        "GOOGLE URL","Google URL","google\_url",  
                        "URL (Name \+ Address \+ Postal)",  
                        "URL (Name \+ Lat \+ Long)",  
                        "URL (Name \+ Coordinates)",  
                        "URL","url",  
                    \])

                    if grid\_c and gurl\_c:  
                        for \_, row in fb\_df.iterrows():  
                            g \= str(row.get(grid\_c,"") or "").strip()  
                            u \= str(row.get(gurl\_c,"") or "").strip()  
                            if g and u and u.lower() not in ("nan",""):  
                                lookup\[norm\_url(u)\] \= g  
                        st.success(f"Loaded {len(lookup):,} GRIDs from uploaded file.")  
                    elif grid\_c and not gurl\_c:  
                        \# URL column name not recognised — find any column with http URLs  
                        auto\_url\_col \= next(  
                            (col for col in fb\_df.columns  
                             if fb\_df\[col\].dropna().astype(str).str.startswith("http").any()),  
                            None)  
                        if auto\_url\_col:  
                            for \_, row in fb\_df.iterrows():  
                                g \= str(row.get(grid\_c,"") or "").strip()  
                                u \= str(row.get(auto\_url\_col,"") or "").strip()  
                                if g and u and u.startswith("http"):  
                                    lookup\[norm\_url(u)\] \= g  
                            st.success(  
                                f"Loaded {len(lookup):,} GRIDs "  
                                f"(URL column auto-detected as '{auto\_url\_col}').")  
                        else:  
                            st.error("Found GRID column but could not identify the URL column.")  
                    else:  
                        st.error(  
                            f"Could not find GRID or URL columns. "  
                            f"Detected columns: {list(fb\_df.columns\[:8\])}")  
                except Exception as e:  
                    st.error(f"Error reading URL CSV: {e}")

        apify\_merge\_up \= st.file\_uploader(  
            "Upload Apify export (.csv or .xlsx)",  
            type=\["csv","xlsx","xls"\], key="apify\_merge")

        if apify\_merge\_up and lookup:  
            try:  
                merge\_bytes \= apify\_merge\_up.read()  
                apy\_df      \= \_cached\_read(merge\_bytes, apify\_merge\_up.name)

                \# Find the input URL column — inputStartUrl is the cleanest match  
                input\_url\_col \= detect\_column(apy\_df, \[  
                    "inputStartUrl", "input\_start\_url",  
                    "searchPageUrl",  "search\_page\_url",  
                    "searchUrl",      "inputUrl",  
                \])

                if input\_url\_col is None:  
                    st.error(  
                        "Could not find \`inputStartUrl\` (or \`searchPageUrl\`) column "  
                        "in the Apify export. Make sure you export with these columns enabled.")  
                else:  
                    grids   \= \[\]  
                    matched \= 0  
                    for \_, row in apy\_df.iterrows():  
                        raw     \= str(row.get(input\_url\_col, "") or "")  
                        key     \= norm\_url(raw)  
                        grid    \= lookup.get(key, "")  
                        if grid:  
                            matched \+= 1  
                        grids.append(grid)

                    result\_df \= apy\_df.copy()  
                    if "GRID" in result\_df.columns:  
                        result\_df\["GRID"\] \= grids   \# overwrite existing  
                    else:  
                        result\_df.insert(0, "GRID", grids)

                    unmatched \= len(result\_df) \- matched  
                    if unmatched \> 0:  
                        st.warning(  
                            f"⚠️ {matched} of {len(result\_df)} rows matched. "  
                            f"{unmatched} rows have no GRID — check that this Apify export "  
                            f"was generated from the same URL batch.")  
                    else:  
                        st.success(f"✅ All {len(result\_df)} rows matched successfully.")

                    buf2 \= io.StringIO()  
                    result\_df.to\_csv(buf2, index=False)  
                    st.download\_button(  
                        "⬇ Download Apify export with GRID",  
                        buf2.getvalue(),  
                        "apify\_with\_grid.csv",  
                        mime="text/csv",  
                        use\_container\_width=True)

            except Exception as e:  
                st.error(f"Error processing Apify export: {e}")

        elif apify\_merge\_up and not lookup:  
            st.warning("No GRID lookup available. Generate URLs in Step 1 first, "  
                       "or upload your URL CSV using the expander above.")

    \# ════════════════════════════════════════════════════════════════  
    \# TAB 3 — SF ACCOUNT AUDIT  
    \# ════════════════════════════════════════════════════════════════  
    with tab3:  
        st.markdown(  
            "Find suspected duplicate accounts \*\*within your Salesforce master\*\*. "  
            "Run this periodically for data hygiene — it's separate from lead dedup.")

        audit\_up \= st.file\_uploader(  
            "Upload Salesforce Master (for audit)", type=\["csv","xlsx","xls"\], key="audit")

        if audit\_up:  
            try:  
                df\_audit\_raw \= \_cached\_read(audit\_up.read(), audit\_up.name)  
                st.success(f"Loaded: {len(df\_audit\_raw):,} accounts")  
            except Exception as e:  
                st.error(f"Error: {e}"); st.stop()

            \# Auto-detect then let user confirm / override  
            auto\_audit \= {  
                "name":   detect\_column(df\_audit\_raw, \["Account Name","Name","name"\]),  
                "postal": detect\_column(df\_audit\_raw, \["Restaurant PostalCode","PostalCode",  
                                                        "Postal Code","Zip/Postal Code",  
                                                        "BillingPostalCode"\]),  
                "street": detect\_column(df\_audit\_raw, \["Formatted Restaurant Address",  
                                                        "BillingStreet","Street","Address"\]),  
                "status": detect\_column(df\_audit\_raw, \["Account Status","Account\_Status\_\_c"\]),  
                "grid":   detect\_column(df\_audit\_raw, \["GRID","grid","Grid"\]),  
                "sf\_id":  detect\_column(df\_audit\_raw, \["SF 18 Char ID","Id","SF\_ID",  
                                                        "Salesforce ID"\]),  
            }  
            audit\_map \= \_col\_map\_ui(  
                df\_audit\_raw, auto\_audit,  
                fields=\[  
                    ("name",   "Account Name",    True),  
                    ("postal", "Postal Code",     True),  
                    ("street", "Address",         False),  
                    ("status", "Account Status",  False),  
                    ("grid",   "GRID",            False),  
                    ("sf\_id",  "SF 18-Char ID",   False),  
                \],  
                key\_prefix="t3\_audit")  
            au\_name   \= audit\_map.get("name")  
            au\_post   \= audit\_map.get("postal")  
            au\_addr   \= audit\_map.get("street")  
            au\_status \= audit\_map.get("status")  
            au\_grid   \= audit\_map.get("grid")  
            au\_id     \= audit\_map.get("sf\_id")

            audit\_thresh \= st.slider(  
                "Name similarity threshold", 50, 100, 70, 5, key="audit\_thresh",  
                help="Pairs at or above this similarity at same postal+unit are flagged.")  
            st.caption(  
                f"Pairs with name similarity ≥ {audit\_thresh}% at the same postal code "  
                f"and unit number will be flagged as suspected duplicates.")

            if st.button("🔍 Run SF Account Audit", type="primary", key="run\_audit"):  
                with st.spinner("Preprocessing SF accounts…"):  
                    df\_audit \= \_cached\_preprocess\_crm(  
                        df\_audit\_raw,  
                        name\_col   \= au\_name,  
                        postal\_col \= au\_post,  
                        addr\_col   \= au\_addr,  
                        status\_col \= au\_status,  
                        char\_map\_tuple \= (),  
                    )

                progress\_a \= st.progress(0, text="Scanning for duplicates…")  
                with st.spinner("Scanning for suspected duplicates…"):  
                    pairs \= find\_sf\_duplicates(  
                        df\_audit, au\_name, au\_addr,  
                        au\_status, au\_grid, au\_id,  
                        audit\_thresh,  
                    )  
                progress\_a.progress(1.0, text="Done\!")

                if not pairs:  
                    st.success(  
                        f"✅ No suspected duplicates found at {audit\_thresh}% threshold. "  
                        f"Your SF data looks clean.")  
                else:  
                    pairs\_df \= pd.DataFrame(pairs)  
                    high   \= len(pairs\_df\[pairs\_df\["RISK\_LEVEL"\] \== "🔴 High"\])  
                    medium \= len(pairs\_df\[pairs\_df\["RISK\_LEVEL"\] \== "🟡 Medium"\])  
                    low    \= len(pairs\_df\[pairs\_df\["RISK\_LEVEL"\] \== "🟢 Low"\])

                    st.success(f"✅ Audit complete — {len(pairs\_df):,} suspected duplicate pairs found.")

                    m1, m2, m3, m4 \= st.columns(4)  
                    m1.metric("Total Pairs",    len(pairs\_df))  
                    m2.metric("🔴 High Risk",   high)  
                    m3.metric("🟡 Medium Risk", medium)  
                    m4.metric("🟢 Low Risk",    low)

                    \# Risk breakdown chart  
                    st.subheader("Risk Breakdown")  
                    st.bar\_chart(pd.DataFrame({  
                        "Risk": \["🔴 High","🟡 Medium","🟢 Low"\],  
                        "Pairs": \[high, medium, low\]  
                    }).set\_index("Risk"))

                    def \_audit\_style(row):  
                        c \= {"🔴 High":"background-color:\#f8d7da",  
                             "🟡 Medium":"background-color:\#fff3cd",  
                             "🟢 Low":"background-color:\#d4edda"}  
                        return \[c.get(row\["RISK\_LEVEL"\],"")\] \* len(row)  
                    def \_show\_a(dff):  
                        if dff.empty: st.info("No pairs in this category.")  
                        else: st.dataframe(dff.style.apply(\_audit\_style, axis=1),  
                                           use\_container\_width=True)

                    st.subheader("Suspected Duplicate Pairs")  
                    a\_tabs \= st.tabs(\[  
                        f"All ({len(pairs\_df)})",  
                        f"🔴 High ({high})",  
                        f"🟡 Medium ({medium})",  
                        f"🟢 Low ({low})",  
                    \])  
                    with a\_tabs\[0\]: \_show\_a(pairs\_df)  
                    with a\_tabs\[1\]: \_show\_a(pairs\_df\[pairs\_df\["RISK\_LEVEL"\] \== "🔴 High"\])  
                    with a\_tabs\[2\]: \_show\_a(pairs\_df\[pairs\_df\["RISK\_LEVEL"\] \== "🟡 Medium"\])  
                    with a\_tabs\[3\]: \_show\_a(pairs\_df\[pairs\_df\["RISK\_LEVEL"\] \== "🟢 Low"\])

                    st.download\_button(  
                        "📥 Download Audit Results",  
                        pairs\_df.to\_csv(index=False),  
                        "sf\_audit\_results.csv",  
                        mime="text/csv", use\_container\_width=True)

    \# ════════════════════════════════════════════════════════════════  
    \# TAB 4 — CRM CHECK (no GRID / no Apify)  
    \# ════════════════════════════════════════════════════════════════  
    with tab4:  \# CRM Check  
        st.subheader("🔍 Quick CRM Duplicate Check")  
        st.caption(  
            "Upload a restaurant list (e.g. from a government website) and your "  
            "Salesforce CRM export to check for existing duplicates. "  
            "No GRID or Apify needed.")

        cc1, cc2 \= st.columns(2)  
        with cc1:  
            st.markdown("\*\*Restaurant List\*\*")  
            st.caption("CSV or Excel — needs: Name, Street, Postal Code")  
            rest\_up \= st.file\_uploader(  
                "Upload restaurant list",  
                type=\["csv","xlsx","xls"\], key="crm\_check\_rest")  
        with cc2:  
            st.markdown("\*\*CRM All Accounts (Salesforce)\*\*")  
            st.caption("Same export used in the Classify Leads tab")  
            crm\_chk\_up \= st.file\_uploader(  
                "Upload CRM export",  
                type=\["csv","xlsx","xls"\], key="crm\_check\_crm")

        if rest\_up and crm\_chk\_up:  
            try:  
                rest\_bytes \= rest\_up.read()  
                rest\_raw   \= \_cached\_read(rest\_bytes, rest\_up.name)

                crm\_chk\_bytes \= crm\_chk\_up.read()  
                crm\_chk\_df, col\_map\_crm\_chk \= load\_crm(crm\_chk\_bytes, crm\_chk\_up.name,  
                                                         market\_cfg)

                \# ── Auto-detect restaurant list columns ───────────  
                auto\_rest \= {  
                    "name":   detect\_column(rest\_raw, \[  
                        "Company / Account","Account Name","Restaurant Name",  
                        "Name","Business Name","name"\]),  
                    "street": detect\_column(rest\_raw, \[  
                        "Street","Formatted Restaurant Address","Address",  
                        "Full Address","Formatted Address","street","address"\]),  
                    "postal": detect\_column(rest\_raw, \[  
                        "Zip/Postal Code","Restaurant PostalCode","Postal Code",  
                        "Postal","PostalCode","Zip","postal","zip"\]),  
                    "grid":   detect\_column(rest\_raw, \["GRID","grid","Grid"\]),  
                }  
                rest\_map \= \_col\_map\_ui(  
                    rest\_raw, auto\_rest,  
                    fields=\[  
                        ("name",   "Restaurant name",  True),  
                        ("street", "Street / address", False),  
                        ("postal", "Postal code",      False),  
                        ("grid",   "GRID",             False),  
                    \],  
                    key\_prefix="t5\_rest")  
                name\_c   \= rest\_map.get("name")  
                street\_c \= rest\_map.get("street")  
                postal\_c \= rest\_map.get("postal")  
                grid\_c   \= rest\_map.get("grid")  
                rest\_cols \= {"name": name\_c, "street": street\_c,  
                             "postal": postal\_c, "grid": grid\_c}

                if not name\_c:  
                    st.warning("Please map the restaurant name column above.")  
                else:  
                    st.info(  
                        f"\*\*{len(rest\_raw):,} restaurants\*\* loaded  ·  "  
                        f"Name: \`{name\_c}\`  ·  "  
                        f"Street: \`{street\_c or '—'}\`  ·  "  
                        f"Postal: \`{postal\_c or '—'}\`")

                    if st.button("▶ Run CRM Check", type="primary",  
                                 use\_container\_width=True, key="run\_crm\_check"):  
                        with st.spinner("Checking against CRM…"):  
                            \_char\_map \= market\_cfg.get("char\_map", {})  
                            check\_df \= crm\_check\_classify(  
                                rest\_raw, rest\_cols,  
                                crm\_chk\_df, col\_map\_crm\_chk,  
                                \_char\_map, p2\_threshold, p3\_threshold)  
                            st.session\_state\["crm\_check\_result"\] \= check\_df

                    if "crm\_check\_result" in st.session\_state:  
                        chk \= st.session\_state\["crm\_check\_result"\]  
                        cnts \= chk\["Label"\].value\_counts()

                        mc \= st.columns(3)  
                        mc\[0\].metric("✅ Unverified — Create",  
                                     cnts.get("Unverified", 0))  
                        mc\[1\].metric("🟡 P3 — Review first",  
                                     cnts.get("P3 — Potential Match", 0))  
                        mc\[2\].metric("🔴 P4 — Duplicate Skip",  
                                     cnts.get("P4 — Duplicate", 0))

                        CSTYLE \= {  
                            "Unverified":           "background-color:\#dbeafe",  
                            "P3 — Potential Match": "background-color:\#fff3cd",  
                            "P4 — Duplicate":       "background-color:\#f8d7da",  
                        }  
                        def \_cs(row):  
                            return \[CSTYLE.get(row\["Label"\],"")\] \* len(row)  
                        def \_cshow(dff):  
                            if dff.empty: st.info("No entries in this category.")  
                            else: st.dataframe(  
                                dff.style.apply(\_cs, axis=1),  
                                use\_container\_width=True)

                        ctabs \= st.tabs(\[  
                            f"All ({len(chk)})",  
                            f"✅ Unverified ({cnts.get('Unverified',0)})",  
                            f"🟡 P3 Potential ({cnts.get('P3 — Potential Match',0)})",  
                            f"🔴 P4 Duplicate ({cnts.get('P4 — Duplicate',0)})",  
                        \])  
                        with ctabs\[0\]: \_cshow(chk)  
                        with ctabs\[1\]: \_cshow(chk\[chk\["Label"\]=="Unverified"\])  
                        with ctabs\[2\]: \_cshow(chk\[chk\["Label"\]=="P3 — Potential Match"\])  
                        with ctabs\[3\]: \_cshow(chk\[chk\["Label"\]=="P4 — Duplicate"\])

                        excel\_chk \= build\_crm\_check\_excel(chk)  
                        st.download\_button(  
                            "📥 Download CRM Check Report",  
                            excel\_chk,  
                            f"CRM\_Check\_{pd.Timestamp.now().strftime('%d%m%y')}.xlsx",  
                            mime="application/vnd.openxmlformats-officedocument"  
                                 ".spreadsheetml.sheet",  
                            use\_container\_width=True)

            except Exception as e:  
                st.error(f"Error: {e}")

    \# ════════════════════════════════════════════════════════════════  
    \# TAB 5 — KPI SAMPLE CHECKER  
    \# ════════════════════════════════════════════════════════════════  
    with tab5:  
        st.subheader("📋 KPI Sample Checker")  
        st.caption("Monthly QA check — 10% stratified sample per Agent × Source × Status.")

        st.markdown("\*\*Step 1 · Upload Lead Status Change report\*\*")  
        st.link\_button(  
            "Open Lead Status Change Report →",  
            "https://deliveryhero.lightning.force.com/lightning/r/Report/"  
            "00ObO000007f12XUAQ/view?queryScope=userFolders",  
            use\_container\_width=False)  
        kpi\_up \= st.file\_uploader(  
            "Upload Lead Status Change report (.xlsx or .csv)",  
            type=\["xlsx","xls","csv"\], key="kpi\_leads")

        kpi\_sampled\_df \= None  
        kpi\_sampling\_df \= None

        if kpi\_up:  
            try:  
                kpi\_raw \= \_cached\_read(kpi\_up.read(), kpi\_up.name)  
                kpi\_raw \= clean\_lead\_report(kpi\_raw)  
                st.success(f"Loaded: {len(kpi\_raw):,} leads")  
                kpi\_sampled\_df, kpi\_sampling\_df \= sample\_leads(kpi\_raw)  
                st.success(f"Sample: {len(kpi\_sampled\_df):,} leads selected "  
                           f"({len(kpi\_sampling\_df):,} strata)")  
                with st.expander("📊 Sampling breakdown"):  
                    st.dataframe(kpi\_sampling\_df, use\_container\_width=True)

                \# Generate Apify URLs for sampled leads  
                st.divider()  
                st.markdown("\*\*Step 2 · Generate Google Maps URLs\*\*")  
                url\_mode\_kpi \= st.radio(  
                    "URL format", \["text","coords"\],  
                    format\_func=lambda m: (  
                        "📝 Street \+ Postal" if m \== "text"  
                        else "📍 Coordinates (if available)"),  
                    horizontal=True, key="kpi\_url\_mode")

                kpi\_urls \= \[\]  
                for \_, row in kpi\_sampled\_df.iterrows():  
                    \# Include restaurant name — strip garbled/non-ASCII chars  
                    \# (Chinese names may read as ??? from XLS encoding)  
                    raw\_name \= str(row.get("Company","") or  
                                   row.get("Company / Account","") or "")  
                    name    \= re.sub(r'\[^\\x20-\\x7E\]+', '', raw\_name).strip()  
                    street  \= str(row.get("Street","") or "")  
                    postal  \= \_norm\_postal\_input(  
                                  str(row.get("Zip/Postal Code","") or ""))  
                    if url\_mode\_kpi \== "text":  
                        parts \= \[p for p in \[name, street, postal\]  
                                 if p and p.lower() not in ("0","0.0","nan","")\]  
                        q \= " ".join(parts)  
                        url \= (f"https://www.google.com/maps/search/?api=1\&query={quote(q)}"  
                               if q else "")  
                    else:  
                        url \= ""  
                    kpi\_urls.append(url)

                kpi\_sampled\_df \= kpi\_sampled\_df.copy()  
                kpi\_sampled\_df\["GOOGLE URL"\] \= kpi\_urls

                \# Store GRID→norm\_url mapping in session state  
                url\_to\_grid \= {}  
                for (\_, r), u in zip(kpi\_sampled\_df.iterrows(), kpi\_urls):  
                    if not u:  
                        continue  
                    \# Safely extract GRID as a plain string  
                    g \= r\["GRID"\] if "GRID" in r.index else ""  
                    if isinstance(g, pd.Series):  
                        g \= g.iloc\[0\] if len(g) \> 0 else ""  
                    g \= str(g or "").strip()  
                    if g:  
                        url\_to\_grid\[norm\_url(u)\] \= g  
                st.session\_state\["kpi\_url\_to\_grid"\] \= url\_to\_grid  
                st.caption(f"✅ {len(url\_to\_grid):,} URL→GRID pairs stored in session")

                valid\_urls \= \[u for u in kpi\_urls if u\]  
                st.text\_area("URLs for Apify", "\\n".join(valid\_urls), height=150,  
                             key="kpi\_url\_preview")  
                buf\_kpi \= io.StringIO()  
                kpi\_sampled\_df.to\_csv(buf\_kpi, index=False)  
                st.download\_button(  
                    "⬇ Download sampled leads \+ URLs",  
                    buf\_kpi.getvalue(),  
                    f"kpi\_sample\_{pd.Timestamp.now().strftime('%d%m%y')}.csv",  
                    mime="text/csv", use\_container\_width=True)

            except Exception as e:  
                st.error(f"Error: {e}")

        st.divider()  
        st.markdown("\*\*Step 3 · Run Apify\*\*")  
        st.info(  
            "1. Copy the URLs from Step 2 and paste them as inputs into your Apify scraper\\n"  
            "2. Run the scraper — takes \~10–15 min for 100–200 leads\\n"  
            "3. Export the Apify results as CSV\\n"  
            "4. Come back here and upload the raw export in Step 4 below — "  
            "the tool will automatically match each row to a GRID via inputStartUrl"  
        )

        st.divider()  
        st.markdown("\*\*Step 4 · Upload supporting files & run checks\*\*")

        k1, k2, k3 \= st.columns(3)  
        with k1:  
            kpi\_apify\_up \= st.file\_uploader(  
                "Apify export — raw export from Apify (optional)",  
                type=\["csv","xlsx","xls"\], key="kpi\_apify")  
        with k2:  
            st.link\_button(  
                "Open CRM All Accounts →",  
                "https://deliveryhero.lightning.force.com/lightning/r/Report/"  
                "00ObO000005IE85UAG/view?queryScope=userFolders",  
                use\_container\_width=True)  
            kpi\_crm\_up \= st.file\_uploader(  
                "CRM All Accounts (required)",  
                type=\["csv","xlsx","xls"\], key="kpi\_crm")  
        with k3:  
            st.link\_button(  
                "Open Account Details Report →",  
                "https://deliveryhero.lightning.force.com/lightning/r/Report/"  
                "00ObO000007f1TxUAI/view",  
                use\_container\_width=True)  
            kpi\_acc\_up \= st.file\_uploader(  
                "Account Details — converted leads (required)",  
                type=\["csv","xlsx","xls"\], key="kpi\_acc")

        if kpi\_crm\_up and kpi\_acc\_up and kpi\_sampled\_df is not None:  
            \# ── Apify: auto-match GRID via session state URL→GRID map ──  
            kpi\_apify\_df \= None  
            if kpi\_apify\_up:  
                try:  
                    kpi\_apify\_raw \= \_cached\_read(kpi\_apify\_up.read(), kpi\_apify\_up.name)

                    \# Check if GRID already present  
                    existing\_gc \= detect\_column(kpi\_apify\_raw, \["GRID","grid","Grid"\])

                    if existing\_gc and existing\_gc \== "GRID":  
                        \# Already has GRID — use directly  
                        kpi\_apify\_df \= kpi\_apify\_raw  
                        st.success(f"Apify loaded: {len(kpi\_apify\_df):,} rows (GRID column present)")  
                    else:  
                        \# Match via inputStartUrl → GRID using session state  
                        url\_map \= st.session\_state.get("kpi\_url\_to\_grid", {})  
                        url\_col \= detect\_column(kpi\_apify\_raw,  
                                                \["inputStartUrl","searchPageUrl"\])  
                        if url\_col and url\_map:  
                            grids   \= \[\]  
                            for \_, r in kpi\_apify\_raw.iterrows():  
                                key \= norm\_url(str(r.get(url\_col,"") or ""))  
                                grids.append(url\_map.get(key,""))  
                            matched \= sum(1 for g in grids if g)  
                            if matched \== 0:  
                                st.error(  
                                    "❌ 0 GRIDs matched — the Apify file you uploaded "  
                                    "appears to be from a \*\*different batch\*\*.\\n\\n"  
                                    "\*\*What to do:\*\*\\n"  
                                    "1. Go back to Step 2 and download the URL CSV\\n"  
                                    "2. Paste those specific URLs into Apify\\n"  
                                    "3. Export and upload \*\*that\*\* Apify result here")  
                                kpi\_apify\_df \= None  
                            else:  
                                if existing\_gc:  
                                    kpi\_apify\_raw\["GRID"\] \= grids  
                                else:  
                                    kpi\_apify\_raw.insert(0, "GRID", grids)  
                                kpi\_apify\_df \= kpi\_apify\_raw  
                                st.success(  
                                    f"Apify loaded: {len(kpi\_apify\_df):,} rows — "  
                                    f"{matched:,} GRIDs matched via URL")  
                                if matched \< len(kpi\_apify\_df):  
                                    st.warning(  
                                        f"{len(kpi\_apify\_df)-matched:,} rows could not be matched.")  
                            \# Debug expander — helps diagnose mismatches  
                            with st.expander("🔍 URL matching debug (expand if GRIDs not matching)"):  
                                st.write(f"Session state URL map: \*\*{len(url\_map):,} keys\*\*")  
                                if url\_map:  
                                    sample\_key \= next(iter(url\_map))  
                                    st.write(f"Sample key in map: \`{sample\_key\[:80\]}...\`")  
                                    st.write(f"→ GRID: \`{url\_map\[sample\_key\]}\`")  
                                first\_apify\_url \= norm\_url(  
                                    str(kpi\_apify\_raw.iloc\[0\].get(url\_col,"") or ""))  
                                st.write(f"First Apify inputStartUrl (normalised): \`{first\_apify\_url\[:80\]}...\`")  
                                st.write(f"Key exists in map: \*\*{first\_apify\_url in url\_map}\*\*")  
                        else:  
                            \# Fallback: use as-is and warn  
                            kpi\_apify\_df \= kpi\_apify\_raw  
                            st.warning(  
                                "No URL→GRID map found in session state. "  
                                "Generate URLs in Step 1 first, then upload Apify results here. "  
                                "Apify data loaded but GRID matching may be incomplete.")  
                except Exception as e:  
                    st.error(f"Apify error: {e}")

            \# Load CRM  
            kpi\_crm\_raw     \= None  
            kpi\_col\_map\_crm \= {}  
            try:  
                kpi\_crm\_raw  \= \_cached\_read(kpi\_crm\_up.read(), kpi\_crm\_up.name)  
                auto\_crm \= {  
                    "name":   detect\_column(kpi\_crm\_raw, \["Account Name","Name","name"\]),  
                    "postal": detect\_column(kpi\_crm\_raw, \["Restaurant PostalCode","PostalCode",  
                                                           "Postal Code","Zip/Postal Code"\]),  
                    "street": detect\_column(kpi\_crm\_raw, \["Formatted Restaurant Address",  
                                                           "BillingStreet","Street","Address"\]),  
                    "grid":   detect\_column(kpi\_crm\_raw, \["GRID","grid","Grid"\]),  
                    "status": detect\_column(kpi\_crm\_raw, \["Account Status","Status"\]),  
                }  
                kpi\_col\_map\_crm \= \_col\_map\_ui(  
                    kpi\_crm\_raw, auto\_crm,  
                    fields=\[  
                        ("name",   "Account Name",   True),  
                        ("postal", "Postal Code",    True),  
                        ("street", "Address",        False),  
                        ("grid",   "GRID",           False),  
                        ("status", "Account Status", False),  
                    \], key\_prefix="kpi\_crm")  
                st.success(f"CRM loaded: {len(kpi\_crm\_raw):,} accounts")  
            except Exception as e:  
                st.error(f"CRM error: {e}")

            \# Load account details  
            kpi\_acc\_df   \= None  
            kpi\_acc\_cols \= {}  
            if kpi\_acc\_up:  
                try:  
                    kpi\_acc\_df  \= \_cached\_read(kpi\_acc\_up.read(), kpi\_acc\_up.name)  
                    auto\_acc \= {  
                        "grid":            detect\_column(kpi\_acc\_df, \["GRID","grid"\]),  
                        "name":            detect\_column(kpi\_acc\_df, \["Account Name","Name"\]),  
                        "phone":           detect\_column(kpi\_acc\_df, \["Phone"\]),  
                        "email":           detect\_column(kpi\_acc\_df, \["Account Email","Email"\]),  
                        "website":         detect\_column(kpi\_acc\_df, \["Website"\]),  
                        "social\_media":    detect\_column(kpi\_acc\_df, \["Social Media URL","Social Media"\]),  
                        "parent\_account":  detect\_column(kpi\_acc\_df, \["Parent Account"\]),  
                        "business\_office": detect\_column(kpi\_acc\_df, \["Business Office"\]),  
                        "delivery\_service":detect\_column(kpi\_acc\_df, \["Delivery Service","Category"\]),  
                        "target\_partner":  detect\_column(kpi\_acc\_df, \["Target Partner"\]),  
                        "category":        detect\_column(kpi\_acc\_df, \["Category","Restaurant Category"\]),  
                    }  
                    kpi\_acc\_cols \= \_col\_map\_ui(  
                        kpi\_acc\_df, auto\_acc,  
                        fields=\[  
                            ("grid",            "GRID",             True),  
                            ("name",            "Account Name",     True),  
                            ("phone",           "Phone",            False),  
                            ("email",           "Email",            False),  
                            ("website",         "Website",          False),  
                            ("social\_media",    "Social Media URL", False),  
                            ("parent\_account",  "Parent Account",   False),  
                            ("business\_office", "Business Office",  False),  
                            ("delivery\_service","Delivery Service", False),  
                            ("target\_partner",  "Target Partner",   False),  
                            ("category",        "Restaurant Category",False),  
                        \], key\_prefix="kpi\_acc")  
                    st.success(f"Account details loaded: {len(kpi\_acc\_df):,} rows")  
                except Exception as e:  
                    st.error(f"Account details error: {e}")

            if kpi\_crm\_raw is not None and st.button(  
                    "▶ Run KPI Checks", type="primary",  
                    use\_container\_width=True, key="run\_kpi"):  
                with st.spinner("Running KPI checks…"):  
                    try:  
                        kpi\_char\_map \= market\_cfg.get("char\_map", {})  
                        kpi\_zones    \= \[\]  
                        if \_builtin:  
                            kpi\_zones \= load\_zones(market\_code=market\_code)

                        kpi\_results\_df, kpi\_agent\_df \= run\_kpi\_checks(  
                            sampled\_df   \= kpi\_sampled\_df,  
                            apify\_df     \= kpi\_apify\_df,  
                            crm\_df       \= kpi\_crm\_raw,  
                            col\_map\_crm  \= kpi\_col\_map\_crm,  
                            account\_df   \= kpi\_acc\_df,  
                            account\_cols \= kpi\_acc\_cols,  
                            zones        \= kpi\_zones,  
                            char\_map     \= kpi\_char\_map,  
                            p2\_threshold \= p2\_threshold,  
                            p3\_threshold \= p3\_threshold,  
                        )  
                        st.success(  
                            f"Done — {len(kpi\_results\_df):,} leads checked, "  
                            f"{kpi\_results\_df\['Auto Error Count'\].sum():,} auto errors found")  
                        st.dataframe(kpi\_agent\_df, use\_container\_width=True)  
                        kpi\_excel \= build\_kpi\_excel(  
                            kpi\_results\_df, kpi\_agent\_df, kpi\_sampling\_df)  
                        st.download\_button(  
                            "📥 Download KPI Scorecard (.xlsx)",  
                            kpi\_excel,  
                            f"KPI\_{pd.Timestamp.now().strftime('%d%m%y')}.xlsx",  
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",  
                            use\_container\_width=True)  
                    except Exception as e:  
                        st.error(f"KPI check error: {e}")  
        elif kpi\_sampled\_df is None:  
            st.info("Upload the Lead Status Change report above to begin.")  
        else:  
            st.info("Upload CRM All Accounts and Account Details report to run checks.")

    \# ════════════════════════════════════════════════════════════════  
    \# TAB 6 — HOW TO USE  
    \# ════════════════════════════════════════════════════════════════  
    with tab6:  \# How to Use  
        st.markdown("""  
\#\# 📖 How to Use the Lead Classifier

This tool classifies new restaurant leads against your Salesforce CRM to identify  
duplicates, potential matches, and genuinely new business opportunities.

\---

\#\#\# Step 1 · Export your Leads from Salesforce  
1\. Click \*\*Open Leads Report →\*\* in the \*Classify Leads\* tab  
2\. Export the report as \*\*CSV\*\* or \*\*Excel (.xls / .xlsx)\*\*  
3\. Required columns: \`GRID\`, \`Company / Account\`, \`Street\`, \`Zip/Postal Code\`

\---

\#\#\# Step 2 · Generate URLs, Run Apify, Get GRID back  
Go to the \*\*🔗 Generate Apify URLs\*\* tab:

\*\*Step 2A — Generate URLs\*\*  
1\. Upload your leads file  
2\. Choose your URL format:  
   \- \*\*📝 Company / Account \+ Street \+ Postal\*\* — recommended  
   \- \*\*📍 Company / Account \+ Coordinates\*\* — use if leads have lat/lng  
3\. Copy the URLs from the \*Generated targets\* box and paste directly into Apify

\*\*Step 2B — Add GRID to your Apify Export\*\*  
1\. After Apify finishes (10–15 min), export the results as CSV  
2\. Come back to \*\*Step 2\*\* in the Generate Apify URLs tab  
3\. Upload the Apify export — the tool matches each row to a GRID via \`inputStartUrl\`  
4\. Download the enriched Apify file with GRID as the first column

\> 💡 Keep the Streamlit tab open while Apify runs — the GRID lookup stays in memory.  
\> If you close the browser, use the \*Upload URL CSV\* expander in Step 2B to reload.

\> ⚠️ Works with both URL formats and mixed batches — matching is always URL-based, never row-order.

\---

\#\#\# Step 3 · Export Salesforce All Accounts (CRM)  
1\. Click \*\*Open Singapore CRM Report →\*\* in the \*Classify Leads\* tab  
2\. Export as \*\*CSV\*\* or \*\*Excel\*\*  
3\. Required columns: \`GRID\`, \`Account Name\`, \`Account Status\`, \`Restaurant PostalCode\`, \`Formatted Restaurant Address\`

\---

\#\#\# Step 4 · Run Classification  
1\. Go to the \*\*📊 Classify Leads\*\* tab  
2\. Upload all three files:  
   \- \*\*Leads\*\* — from Step 1  
   \- \*\*Apify Results with GRID\*\* — from Step 2B  
   \- \*\*CRM Export\*\* — from Step 3  
3\. Click \*\*▶ Run Classification\*\*  
4\. Download the Excel report (filename: \`ALG\_DDMMYY\`)

\---

\#\#\# Agent Columns (A–F in the Excel Report)

The Classified Leads sheet has 6 editable columns for your agents to fill in:

| Column | Purpose |  
|---|---|  
| \*\*Agent\*\* | Assign the lead to an agent |  
| \*\*Due Date\*\* | Set a follow-up date |  
| \*\*Convert/Lost\*\* | Dropdown: \`Converted\` or \`Lost\` |  
| \*\*Invalid Reason\*\* | Dropdown: \`Duplicate\`, \`Invalid Data\`, \`Closed Down\`, \`Wrong Target Group\`, \`Other\` |  
| \*\*Comments/Duplicate GRID\*\* | Free text — notes or the duplicate GRID reference |  
| \*\*Feedback\*\* | Free text — any additional context |

\---

\#\#\# Classification Labels

| Label | What it means | What to do |  
|---|---|---|  
| \*\*✅ P1 — New\*\* | No match in CRM. Google confirms it's an open restaurant. | Pitch delivery to this lead |  
| \*\*🟡 P3 — Potential Match\*\* | Name is 50–74% similar to a CRM account at the same postal code. | Verify manually — could be a duplicate or a different restaurant |  
| \*\*🔴 P4 — Duplicate\*\* | Name is ≥ 75% similar to a CRM account at the same postal code. | Skip — already in the system |  
| \*\*🏢 Business Closed\*\* | Google Maps confirms the restaurant is permanently or temporarily closed. | Skip |  
| \*\*❌ Wrong Target Group\*\* | Google Maps category is not food delivery eligible. | Skip |  
| \*\*⚪ P2 — Please Check\*\* | No Google Maps result found, or the result doesn't match the lead. | Verify manually before acting |

\---

\#\#\# Matching Logic

1\. \*\*Postal code\*\* — only CRM accounts at the same 6-digit postal are compared  
2\. \*\*Unit number\*\* — if the lead has a unit (e.g. \#01-23), all CRM accounts at that unit are scored  
3\. \*\*Name similarity\*\* — fuzzy matching \+ Hanyu Pinyin for Chinese names  
   \- Score ≥ 75% → \*\*P4 Duplicate\*\*  
   \- Score 50–74% → \*\*P3 Potential Match\*\*  
   \- Score \< 50% → Apify check

\---

\#\#\# SF Account Audit (Tab 3\)  
Find suspected duplicate accounts within Salesforce itself. Run monthly for CRM hygiene.

\---

\#\#\# Tips  
\- Thresholds (P3 / P4) are adjustable in the sidebar  
\- Chinese names are auto-converted to Hanyu Pinyin  
\- Exclusion keywords control which Google Maps categories become Wrong Target Group  
\- \`mix & match\` and hawker stalls are handled correctly by default  
        """)

if \_\_name\_\_ \== "\_\_main\_\_":  
    main()

