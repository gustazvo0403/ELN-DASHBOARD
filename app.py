import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go

# ==========================================
# 1. 網頁基本設定與樣式
# ==========================================
st.set_page_config(page_title="大豐銀行 - 股權寶(ELD)收益模擬器", layout="wide", page_icon="📈")

st.markdown("""
    <style>
    div[data-testid="stMetricValue"] { font-size: 1.6rem !important; }
    .scenario-card { padding: 20px; border-radius: 10px; margin-bottom: 20px; border: 1px solid #e0e0e0; box-shadow: 2px 2px 10px rgba(0,0,0,0.05); }
    .card-green { border-left: 6px solid #00A36C; background-color: #f0fff4; }
    .card-yellow { border-left: 6px solid #FFC000; background-color: #fffff0; }
    .card-red { border-left: 6px solid #FF0000; background-color: #fff0f0; }
    .card-title { font-size: 1.25em; font-weight: bold; margin-bottom: 10px; }
    .card-data { font-size: 1.1em; line-height: 1.6; }
    .profit-text { color: #00A36C; font-weight: bold; }
    .loss-text { color: #FF0000; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("📈 股權寶 (ELD) 投資收益模擬器")
st.markdown("支援 **Excel 批量導入** 或 **手動輸入**。透過調整**計價日收市價**，直觀了解不同市場情況下的結算方式與真實損益。")

# ==========================================
# 2. 初始化與狀態管理 (Session State)
# ==========================================
keys_defaults = {
    "underlying": "阿里巴巴",
    "ric": "09988.HK",
    "val_date": "2026-04-29",
    "fix_date": "2026-06-10",
    "ref_price": 130.20,
    "strike_pct": 90.0,
    "shares": 4300,
    "principal": 498267.90,
    "extreme_drop": 100.00,
    "prev_excel_sel": None
}

for k, v in keys_defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ==========================================
# 3. 左側頂部：一鍵重置功能
# ==========================================
if st.sidebar.button("🔄 一鍵重置 (清空所有設定)", use_container_width=True):
    for k in keys_defaults.keys():
        st.session_state[k] = None
    st.rerun()

st.sidebar.markdown("---")

# ==========================================
# 4. 左側：資料載入區與 Excel 處理
# ==========================================
st.sidebar.header("📥 1. 數據載入方式")
data_mode = st.sidebar.radio("請選擇輸入方式：", ["✍️ 手動輸入 (預設案例)", "📁 批量導入 Excel"])

if data_mode == "📁 批量導入 Excel":
    st.sidebar.info("請上傳您的 ELD 報價 Excel 檔")
    uploaded_file = st.sidebar.file_uploader("上傳檔案", type=["xlsx", "xls"])
    
    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file, engine='openpyxl').dropna(how='all')
            
            if '代號 Code' in df.columns and '掛鈎股票 Underlying2' in df.columns:
                df = df.dropna(subset=['代號 Code', '掛鈎股票 Underlying2'])
                display_options = df['代號 Code'].astype(str) + " - " + df['掛鈎股票 Underlying2'].astype(str)
                
                selected_idx = st.sidebar.selectbox("🔍 請選擇要模擬的產品：", df.index, format_func=lambda x: display_options[x])
                
                if st.session_state.prev_excel_sel != selected_idx:
                    row = df.loc[selected_idx]
                    
                    st.session_state.underlying = str(row.get('掛鈎股票 Underlying2', ""))
                    st.session_state.ric = str(row.get('編號 RIC', ""))
                    st.session_state.ref_price = float(row.get('參考現價 REF INIT PRICE (HKD)', 0.0))
                    
                    raw_strike = float(row.get('折扣率 STRIKE(%)', 0.9))
                    st.session_state.strike_pct = raw_strike * 100 if raw_strike <= 1.0 else raw_strike
                    
                    st.session_state.shares = int(row.get('參考股數 REF NO OF SHARES', 0))
                    st.session_state.principal = float(row.get('參考交易金額 REF DEPOSIT AMT (HKD)', 0.0))
                    
                    st.session_state.val_date = str(row.get('生效日 VALUE DATE', ""))[:10]
                    st.session_state.fix_date = str(row.get('計價日 FIXING DATE', ""))[:10]
                    
                    if st.session_state.shares > 0:
                        bk_even = st.session_state.principal / st.session_state.shares
                        st.session_state.extreme_drop = round(bk_even * 0.85, 2)
                    
                    st.session_state.prev_excel_sel = selected_idx
                    st.rerun() 
                    
        except ImportError:
            st.sidebar.error("❌ 系統缺少讀取 Excel 的依賴套件。請執行：`pip install openpyxl`")
        except Exception as e:
            st.sidebar.error(f"讀取 Excel 失敗，請確認檔案格式。錯誤資訊: {e}")

# ==========================================
# 5. 左側：參數設定區 
# ==========================================
st.sidebar.header("⚙️ 2. 參數設定區 (可手動修改)")

st.sidebar.text_input("掛鈎股票名稱", key="underlying")
st.sidebar.text_input("股票代號 (RIC)", key="ric")

col_d1, col_d2 = st.sidebar.columns(2)
with col_d1:
    st.text_input("生效日", key="val_date")
with col_d2:
    st.text_input("計價日", key="fix_date")

st.sidebar.number_input("參考現價 (HKD)", step=1.0, key="ref_price")
st.sidebar.number_input("折扣率 (%)", step=1.0, key="strike_pct")
st.sidebar.number_input("參考股數 (股)", step=100, key="shares")
st.sidebar.number_input("交易金額/本金 (HKD) - 可修改", step=1000.0, key="principal")

st.sidebar.markdown("---")
st.sidebar.number_input("📉 情況三極端下跌收市價假設 (HKD)", step=1.0, key="extreme_drop")

required_keys = ["ref_price", "strike_pct", "shares", "principal", "extreme_drop"]
if any(st.session_state[k] is None for k in required_keys):
    st.info("👈 **目前無參數數據。** 請在左側手動輸入參數，或點擊左上方「批量導入 Excel」上傳檔案以開始模擬。")
    st.stop()

ref_price = st.session_state.ref_price
strike_rate = st.session_state.strike_pct / 100.0
shares = st.session_state.shares
principal = st.session_state.principal

strike_price = ref_price * strike_rate
maturity_amt = shares * strike_price
breakeven = principal / shares if shares else 0
max_profit = maturity_amt - principal

st.sidebar.markdown("---")
st.sidebar.markdown("### 📌 產品關鍵指標")
st.sidebar.info(f"""
**參考行使價**: HKD {strike_price:,.3f}  
**盈虧平衡點**: HKD {breakeven:,.3f}  
**到期本息總額**: HKD {maturity_amt:,.2f}  
**最大潛在利潤**: HKD {max_profit:,.2f}
""")

# ==========================================
# 6. 左側：新增圖表橫軸範圍設定區
# ==========================================
st.sidebar.markdown("---")
st.sidebar.header("📊 3. 圖表顯示設定")

# 以現價為基準，計算一個合理的拉動極限範圍 (0.1 倍 ~ 2.0 倍)
min_bound = max(0.0, float(ref_price * 0.1))
max_bound = float(ref_price * 2.0)

# 給定一個預設的視角區間 (0.4 倍 ~ 1.3 倍)
default_chart_min = float(ref_price * 0.4)
default_chart_max = float(ref_price * 1.3)

chart_range = st.sidebar.slider(
    "設定損益曲線圖「橫軸」區間範圍 (HKD)", 
    min_value=min_bound, 
    max_value=max_bound, 
    value=(default_chart_min, default_chart_max), 
    step=1.0
)
chart_min, chart_max = chart_range

# ==========================================
# 7. 主區域：互動滑動條 (動態適配設定的橫軸區間)
# ==========================================
st.markdown(f"### 🎚️ 模擬【{st.session_state.underlying}】計價日表現")

# 確保滑動條預設值不會超出左側自訂的區間範圍
default_closing = max(chart_min, min(float(ref_price), chart_max))

closing_price = st.slider(
    "請左右拉動橫軸，設定「計價日收市價」(HKD)，觀察對應的損益變化：", 
    min_value=float(chart_min), 
    max_value=float(chart_max), 
    value=default_closing, 
    step=0.5
)

# 邏輯計算
if closing_price >= strike_price:
    scenario = "A"
    settlement_type = "💰 現金結算 (不接貨)"
    settlement_val = maturity_amt
    delivery_val = 0
else:
    if closing_price >= breakeven:
        scenario = "B"
    else:
        scenario = "C"
    settlement_type = "📦 股票接貨"
    delivery_val = shares * closing_price
    settlement_val = delivery_val

pnl = settlement_val - principal
pnl_pct = (pnl / principal) * 100 if principal else 0

# ==========================================
# 8. 全新升級的動態資訊面板 (豐富版)
# ==========================================
if pnl >= 0:
    bg_color = "#f0fff4"
    border_color = "#ccffcc"
    text_color = "#00A36C"
    pnl_sign = "+"
    pnl_emoji = "📈"
else:
    bg_color = "#fff0f0"
    border_color = "#ffcccc"
    text_color = "#FF0000"
    pnl_sign = ""
    pnl_emoji = "📉"

if scenario == "A":
    delivery_title = "📦 接貨股票市值 (HKD)"
    delivery_main = "—"
    delivery_sub = "<span style='color: #00A36C; font-weight: bold;'>✅ 無須接貨，現金收回</span>"
else:
    tax = delivery_val * 0.001
    delivery_title = "📦 接貨股票市值 (HKD)"
    delivery_main = f"{delivery_val:,.2f}"
    delivery_sub = f"預估印花稅支出: <span style='color: #FF0000; font-weight: bold;'>-HKD {tax:,.2f}</span>"

html_panel = f"""
<div style="display: flex; flex-wrap: wrap; gap: 15px; background-color: {bg_color}; border: 1px solid {border_color}; padding: 20px; border-radius: 10px; margin-bottom: 20px;">
    
    <!-- 第一欄：結算方式與本金成本 -->
    <div style="flex: 1; min-width: 200px; border-right: 1px dashed {border_color}; padding-right: 15px;">
        <div style="color: #666; font-size: 0.9em; margin-bottom: 5px;">💼 結算方式 & 本金成本</div>
        <div style="font-size: 1.2em; font-weight: bold; color: #333; margin-bottom: 5px;">{settlement_type}</div>
        <div style="font-size: 0.95em; color: #555;">投入本金: <span style="color: #FF0000; font-weight: bold;">-HKD {principal:,.2f}</span></div>
    </div>
    
    <!-- 第二欄：當前結算總值 -->
    <div style="flex: 1; min-width: 180px; border-right: 1px dashed {border_color}; padding-right: 15px;">
        <div style="color: #666; font-size: 0.9em; margin-bottom: 5px;">💰 當前結算總值 (HKD)</div>
        <div style="font-size: 1.4em; font-weight: bold; color: #333;">{settlement_val:,.2f}</div>
        <div style="color: {text_color}; font-size: 0.9em; font-weight: bold;">{pnl_sign}{pnl:,.2f} (與本金落差)</div>
    </div>
    
    <!-- 第三欄：淨損益額 -->
    <div style="flex: 1; min-width: 180px; border-right: 1px dashed {border_color}; padding-right: 15px;">
        <div style="color: #666; font-size: 0.9em; margin-bottom: 5px;">⚖️ 淨損益額 / 損益率</div>
        <div style="font-size: 1.4em; font-weight: bold; color: {text_color};">{pnl_emoji} {pnl_sign}{pnl:,.2f}</div>
        <div style="color: {text_color}; font-size: 0.9em; font-weight: bold;">回報率: {pnl_sign}{pnl_pct:.2f}%</div>
    </div>
    
    <!-- 第四欄：接貨市值與印花稅 -->
    <div style="flex: 1; min-width: 180px;">
        <div style="color: #666; font-size: 0.9em; margin-bottom: 5px;">{delivery_title}</div>
        <div style="font-size: 1.4em; font-weight: bold; color: #333;">{delivery_main}</div>
        <div style="font-size: 0.9em;">{delivery_sub}</div>
    </div>
    
</div>
"""
st.markdown(html_panel, unsafe_allow_html=True)

# ==========================================
# 9. 互動損益折線圖 (Plotly)
# ==========================================
st.markdown("### 📊 股權寶到期損益曲線圖")

# 【重要更新】：圖表的 X 軸資料點範圍改由左側的 slider 變數控制
prices = np.linspace(chart_min, chart_max, 300)
pnls = np.where(prices >= strike_price, max_profit, (shares * prices) - principal)

fig = go.Figure()
fig.add_trace(go.Scatter(x=prices, y=pnls, mode='lines', name='淨損益', line=dict(color='#1E90FF', width=3)))

fig.add_vline(x=strike_price, line_dash="dash", line_color="orange", annotation_text=f"行使價 ({strike_price:.2f})", annotation_position="top right")
fig.add_vline(x=breakeven, line_dash="dash", line_color="red", annotation_text=f"盈虧平衡點 ({breakeven:.2f})", annotation_position="bottom right")
fig.add_vline(x=closing_price, line_width=2, line_color="purple", annotation_text=f"📍 當前收市價 ({closing_price:.2f})", annotation_position="top left")
fig.add_hline(y=0, line_width=1.5, line_color="black")

fig.update_layout(
    xaxis_title="計價日收市價 (HKD)", yaxis_title="淨損益 (HKD)",
    plot_bgcolor="rgba(245,245,245,0.8)", hovermode="x unified",
    height=400, margin=dict(l=20, r=20, t=30, b=20)
)
st.plotly_chart(fig, use_container_width=True)

# ==========================================
# 10. 三種潛在情況詳細數據 (靜態展示區)
# ==========================================
st.markdown("---")
st.markdown("### 📋 本次投資的三種潛在情況詳解")
st.markdown(f"*(基於當前設定本金 **HKD {principal:,.2f}**)*")

ex_price_2 = round((strike_price + breakeven) / 2, 2)
ex_val_2 = shares * ex_price_2
ex_pnl_2 = ex_val_2 - principal

ex_val_3 = shares * st.session_state.extreme_drop
ex_pnl_3 = ex_val_3 - principal

st.markdown(f"""
<div class="scenario-card card-green">
    <div class="card-title">🟢 情況一：看對方向 (收市價升穿或持平)</div>
    <div class="card-data">
        <b>市場情況：</b> 計價日收市價 <b>≥ HKD {strike_price:.3f}</b> (行使價)<br>
        <b>結算方式：</b> 現金結算 (銀行放棄行使期權)<br>
        <b>到期總值：</b> HKD {maturity_amt:,.2f}<br>
        <b>客戶損益：</b> 穩賺全數利息，淨利潤 <span class="profit-text">+HKD {max_profit:,.2f}</span>
    </div>
</div>

<div class="scenario-card card-yellow">
    <div class="card-title">🟡 情況二：輕微下跌 (跌穿行使價，但高於盈虧平衡點)</div>
    <div class="card-data">
        <b>市場情況：</b> <b>HKD {breakeven:.3f}</b> (盈虧平衡點) <b>≤</b> 計價日收市價 <b>< HKD {strike_price:.3f}</b><br>
        <b>結算方式：</b> 股票接貨 (必須以行使價買入 <b>{shares} 股</b>)<br>
        <b>具體案例：</b> 假設收市價跌至 HKD {ex_price_2:.2f}。該批股票當日市值為 HKD {ex_val_2:,.2f}。<br>
        <b>客戶損益：</b> 受惠於期權金補貼，雖然接貨但仍有微利，淨利潤 <span class="profit-text">+HKD {ex_pnl_2:,.2f}</span> <i>(未計入接貨印花稅)</i>
    </div>
</div>

<div class="scenario-card card-red">
    <div class="card-title">🔴 情況三：大幅下跌 (跌穿盈虧平衡點，承受虧損)</div>
    <div class="card-data">
        <b>市場情況：</b> 計價日收市價 <b>< HKD {breakeven:.3f}</b><br>
        <b>結算方式：</b> 股票接貨 (必須以行使價買入 <b>{shares} 股</b>)<br>
        <b>具體案例：</b> 假設收市價暴跌至您設定的極端價 <b>HKD {st.session_state.extreme_drop:.2f}</b>。股票市值縮水至 HKD {ex_val_3:,.2f}。<br>
        <b>客戶損益：</b> 產生實質資本虧損，淨損益 <span class="loss-text">-HKD {abs(ex_pnl_3):,.2f}</span> <i>(未計入接貨印花稅)</i>
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 11. 重要提示 (免責聲明)
# ==========================================
st.markdown("---")
st.info("""
**【重要提示】**：  
本產品非存款產品，不納入銀行存款保障範疇。本内容所載資料僅供參考之用，文件內任何資訊及分析並不構成對投資產品未來表現的任何保證。本文件所收錄之任何資訊、預測或意見的準確性、正確性或完整性並不保證，亦不會對因依賴有關資訊、預測或意見而引致的損失負任何責任。投資涉及風險，並受市場波動及投資風險所影響。如對本文內容的含意或所引致的影響有任何疑問，請徵詢獨立專業人士的意見。
""")
