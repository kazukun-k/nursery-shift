import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import json
import math
import calendar
import re
from datetime import datetime
import google.generativeai as genai

# ==========================================
# 🎨 1. アプリケーション基本設定・デザイン
# ==========================================
st.set_page_config(
    page_title="保育士シフト管理システム - SmartShift",
    page_icon="📛",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 保育園にマッチした温かみと高級感のあるカスタムCSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Nunito:wght@400;600;700;800&family=Zen+Maru+Gothic:wght@400;500;700&display=swap');
    
    /* フォント適用 */
    html, body, [class*="css"] {
        font-family: 'Zen Maru Gothic', 'Nunito', sans-serif;
    }
    
    /* メインヘッダーのデザイン */
    .main-title {
        color: #FF6B8B;
        font-size: 2.8rem;
        font-weight: 800;
        text-align: center;
        margin-bottom: 5px;
        text-shadow: 2px 2px 4px rgba(255, 107, 139, 0.15);
    }
    .sub-title {
        color: #70A1FF;
        font-size: 1.2rem;
        text-align: center;
        margin-bottom: 30px;
        font-weight: 500;
    }
    
    /* カード風のコンポーネント */
    .card {
        background-color: #FFFFFF;
        border-radius: 15px;
        padding: 20px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.05);
        border: 1px solid #FFE3E8;
        margin-bottom: 20px;
    }
    
    /* ボタンのカスタマイズ */
    div.stButton > button {
        background: linear-gradient(135deg, #FF6B8B 0%, #FF8E53 100%);
        color: white !important;
        border-radius: 25px;
        padding: 10px 25px;
        font-weight: 700;
        border: none;
        box-shadow: 0 4px 10px rgba(255, 107, 139, 0.3);
        transition: all 0.3s ease;
    }
    div.stButton > button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 15px rgba(255, 107, 139, 0.4);
        background: linear-gradient(135deg, #FF8E53 0%, #FF6B8B 100%);
    }
    
    /* サブヘッダーの装飾 */
    h2, h3 {
        color: #2F3542;
        border-left: 5px solid #FF6B8B;
        padding-left: 10px;
        margin-top: 20px;
    }
</style>
""", unsafe_allow_html=True)

# アプリヘッダー
st.markdown('<div class="main-title">📛 保育士シフト管理・作成システム</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">🤖 AI独自ルール解析 × 自動シフト生成エンジン搭載</div>', unsafe_allow_html=True)

# ==========================================
# 🔑 2. サイドバー設定 (APIキー & シフト基本設定)
# ==========================================
with st.sidebar:
    st.markdown("### ⚙️ システム設定")
    
    # APIキー管理
    st.subheader("🔑 Google Gemini APIキー")
    api_key_input = st.text_input(
        "Gemini APIキーを入力",
        type="password",
        placeholder="AIzaSy...",
        help="Google AI Studioで取得した無料のAPIキーを入力してください。未入力の場合は、AI機能のみ標準動作となります。"
    )
    
    if api_key_input:
        st.success("APIキーが入力されました！")
    else:
        st.warning("APIキーを入力すると、自由記述の園独自ルールをAIが自動解釈できるようになります。")
        
    st.markdown("""
    [👉 無料APIキーの取得手順はこちら](https://aistudio.google.com/)
    """)
    
    st.divider()
    
    # 対象年月設定
    st.subheader("📅 シフト対象年月")
    today = datetime.today()
    current_year = today.year
    current_month = today.month
    
    # 翌月をデフォルト値にする
    default_month = current_month + 1 if current_month < 12 else 1
    default_year = current_year if current_month < 12 else current_year + 1
    
    target_year = st.selectbox("対象年", range(current_year - 1, current_year + 3), index=1)
    target_month = st.selectbox("対象月", range(1, 13), index=default_month - 1)
    
    # 月の日数を計算
    num_days = calendar.monthrange(target_year, target_month)[1]
    
    st.info(f"設定期間: {target_year}年{target_month}月 ({num_days}日間)")

# ==========================================
# 🧩 3. 30分スロット定義 & 表記揺れパーサーの実装
# ==========================================
# 7:00〜20:00の30分刻み（合計 26 スロット）
TIME_SLOT_HOURS = []
for h in range(7, 20):
    TIME_SLOT_HOURS.append((h, 0, h, 30))
    TIME_SLOT_HOURS.append((h, 30, h + 1, 0))

# 30分スロット名リスト
time_slots = [f"{sh}:{sm:02d}～{eh}:{em:02d}" for sh, sm, eh, em in TIME_SLOT_HOURS]

# 正社員の固定シフトパターン (A〜E) の定義
REGULAR_PATTERNS = {
    "A": list(range(0, 18)),  # 7:00 - 16:00
    "B": list(range(2, 20)),  # 8:00 - 17:00
    "C": list(range(4, 22)),  # 9:00 - 18:00
    "D": list(range(6, 24)),  # 10:00 - 19:00
    "E": list(range(8, 26))   # 11:00 - 20:00
}

def parse_work_hours(tz_str):
    """
    "8:30-13:00 9:00-13:30 7:00-11:00" などの複数候補をパースし、
    それぞれの候補の時間スロットインデックス(0〜25)のリストのリストを返す。
    例: "8:30-13:00 9:00-13:30" -> [[3, 4, 5, 6, 7, 8, 9, 10, 11], [4, ...]]
    """
    if not tz_str or not isinstance(tz_str, str):
        return []
    
    if tz_str == "全日":
        return [list(range(26))]  # 7:00〜20:00 全てカバー可能
        
    # 表記揺れを統一
    s = tz_str.translate(str.maketrans({
        '０':'0', '１':'1', '２':'2', '３':'3', '４':'4', '５':'5', '６':'6', '７':'7', '８':'8', '９':'9',
        '；':':', ';':':', '：':':',
        '〜':'-', '～':'-', 'ー':'-', '－':'-', '—':'-', 'ー':'-', '─':'-'
    })).strip()
    
    # スペースやカンマ、読点で分割して各時間候補を取得
    parts = re.split(r'[\s,，、/]+', s)
    candidates = []
    
    for part in parts:
        if not part:
            continue
        
        # 1. H:MM - H:MM の抽出
        match = re.search(r'(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})', part)
        if match:
            sh, sm = int(match.group(1)), int(match.group(2))
            eh, em = int(match.group(3)), int(match.group(4))
            
            # 7:00からの30分スロットのインデックスに変換
            start_idx = (sh - 7) * 2 + (1 if sm >= 30 else 0)
            end_idx = (eh - 7) * 2 + (1 if em >= 30 else 0)
            
            start_idx = max(0, min(25, start_idx))
            end_idx = max(0, min(26, end_idx))
            
            candidates.append(list(range(start_idx, end_idx)))
            continue
            
        # 2. H - H (時のみ) の抽出
        match_h = re.search(r'(\d{1,2})\s*-\s*(\d{1,2})', part)
        if match_h:
            sh, eh = int(match_h.group(1)), int(match_h.group(2))
            start_idx = (sh - 7) * 2
            end_idx = (eh - 7) * 2
            
            start_idx = max(0, min(25, start_idx))
            end_idx = max(0, min(26, end_idx))
            
            candidates.append(list(range(start_idx, end_idx)))
            
    if not candidates:
        # デフォルト日中 9:00〜17:00 (インデックス 4〜19)
        return [list(range(4, 20))]
        
    return candidates

def get_covered_slots(shift_name):
    """
    シフト名（記号 A〜E または時間帯文字列）から、カバーする時間帯インデックスのリストを返す
    """
    if shift_name in REGULAR_PATTERNS:
        return REGULAR_PATTERNS[shift_name]
    elif shift_name == "公休":
        return []
    # 自由記述されたパート時間（例: "8:30-13:00"）をパース
    # 複数時間帯から構成されている場合は、最初の候補をカバー時間とする（補助用）
    candidates = parse_work_hours(shift_name)
    return candidates[0] if candidates else []

# ==========================================
# 📂 4. セッション状態の初期化
# ==========================================
if 'staff_list' not in st.session_state:
    st.session_state.staff_list = [
        {"名前": "山田 花子", "雇用形態": "正社員", "月上限日数": 20, "希望時間帯": "全日", "時短希望": "なし", "勤務可能曜日": "月,火,水,木,金,土"},
        {"名前": "鈴木 一郎", "雇用形態": "パート", "月上限日数": 12, "希望時間帯": "8:30-13:00 9:00-13:30 7:00-11:00", "時短希望": "あり", "勤務可能曜日": "月,火,水,木,金"},
        {"名前": "佐藤 美咲", "雇用形態": "パート", "月上限日数": 15, "希望時間帯": "15:00-19:00", "時短希望": "なし", "勤務可能曜日": "月,火,水,木,金"},
        {"名前": "田中 恵子", "雇用形態": "正社員", "月上限日数": 20, "希望時間帯": "全日", "時短希望": "なし", "勤務可能曜日": "月,火,水,木,金,土"},
        {"名前": "小林 翔太", "雇用形態": "正社員", "月上限日数": 22, "希望時間帯": "全日", "時短希望": "なし", "勤務可能曜日": "月,火,水,木,金,土"},
        {"名前": "高橋 陽子", "雇用形態": "パート", "月上限日数": 16, "希望時間帯": "9:00-17:00", "時短希望": "なし", "勤務可能曜日": "月,火,水,木,金"},
        {"名前": "渡辺 理恵", "雇用形態": "パート", "月上限日数": 10, "希望時間帯": "7:00-12:00 8:00-13:00", "時短希望": "なし", "勤務可能曜日": "月,火,水,木,金"},
        {"名前": "伊藤 直美", "雇用形態": "パート", "月上限日数": 14, "希望時間帯": "13:30-19:00 15:00-20:00", "時短希望": "なし", "勤務可能曜日": "月,火,水,木,金"},
    ]

if 'ai_rules' not in st.session_state:
    st.session_state.ai_rules = []

# 園児数テーブルのセッション管理（一括コピー対応のため）
if 'df_children_state' not in st.session_state:
    # 30分枠の園児数初期値
    init_kids_data = {
        "時間帯": time_slots,
        "0歳児数": [3, 3, 5, 5, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 6, 5, 5, 3, 3, 2, 2, 1, 1, 1, 1],
        "1-2歳児数": [5, 5, 10, 10, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 12, 10, 10, 8, 8, 5, 5, 2, 2, 1, 1],
        "3歳児数": [2, 2, 15, 15, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 18, 15, 15, 10, 10, 4, 4, 1, 1, 1, 1],
        "4歳以上児数": [5, 5, 20, 20, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 22, 20, 20, 15, 15, 8, 8, 2, 2, 1, 1]
    }
    st.session_state.df_children_state = pd.DataFrame(init_kids_data)

# --- Google Gemini API 呼び出し ---
def parse_rules_with_gemini(api_key, rule_text):
    if not api_key:
        return None
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        prompt = f"""
あなたは保育園のシフト管理システムのアシスタントです。
園長が入力した以下の「園独自の特殊ルール（日本語）」を解析し、時間帯別の必要人数に対する補正ルールを抽出してください。
なお、時間帯は30分刻み（7:00から20:00まで）で計算されています。

【ルール記述】
{rule_text}

【出力フォーマット】
必ず以下の構造のJSONオブジェクトのみを返してください。不要な説明やマークダウンタグ（```jsonなど）は含めず、純粋なJSONオブジェクト単体、もしくは ```json ... ``` で囲んで出力してください。
JSONには、抽出されたルールが含まれる "rules" キーのリストを設定してください。

出力スキーマ:
{{
  "rules": [
    {{
      "start_hour": 7,  // 開始時間（数値、例: 7）
      "end_hour": 9,    // 終了時間（数値、例: 9。これは9:00までを意味します）
      "min_staff": 2,   // 最低必要な保育士の人数（数値）
      "reason": "合同保育のため" // ルールの理由（文字列、例: 合同保育のため）
    }}
  ]
}}

※該当するルールがない、または解析できない場合は、空のリスト `{"rules": []}` を返してください。
"""
        response = model.generate_content(
            prompt,
            generation_config={"response_mime_type": "application/json"}
        )
        
        clean_text = response.text.strip()
        if clean_text.startswith("```json"):
            clean_text = clean_text[7:]
        if clean_text.endswith("```"):
            clean_text = clean_text[:-3]
        
        data = json.loads(clean_text)
        return data.get("rules", [])
    except Exception as e:
        st.error(f"Gemini APIでの解析中にエラーが発生しました: {str(e)}")
        return None

# --- タブ構成 ---
tab1, tab2, tab3 = st.tabs([
    "📊 園児数・必要人数計算", 
    "👥 職員条件設定", 
    "🗓️ シフト自動生成・Excel出力"
])

# ==========================================
# タブ1: 園児数・必要人数計算
# ==========================================
with tab1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.header("🏢 園児数設定とAI変則ルール")
    
    st.subheader("📝 園独自の特殊ルール（AI自動判定）")
    custom_rule_text = st.text_area(
        "【例】「7-9時は全学年合同で過ごすため最低2人配置」「17:00以降は遅番で最低2人必要」など、独自の配置制限を文章で記述してください。",
        value="7-9時は全学年合同で過ごすため最低2人配置してください。また、17-19時は合同で最低2人配置にしてください。",
        placeholder="ここに園独自のルールを自由に文章で入力してください。"
    )
    
    col_ai_btn, col_ai_status = st.columns([1, 3])
    with col_ai_btn:
        run_ai = st.button("AIルールを解析して反映する")
        
    if run_ai:
        if not api_key_input:
            st.warning("⚠️ APIキーが設定されていません。無料デモ用ルールを適用します。")
            if "7-9" in custom_rule_text and "合同" in custom_rule_text:
                st.session_state.ai_rules = [
                    {"start_hour": 7, "end_hour": 9, "min_staff": 2, "reason": "【デモ適用】7-9時合同保育のため最低2名"}
                ]
                if "17-19" in custom_rule_text or "17時" in custom_rule_text:
                    st.session_state.ai_rules.append(
                        {"start_hour": 17, "end_hour": 19, "min_staff": 2, "reason": "【デモ適用】17-19時合同保育のため最低2名"}
                    )
            st.success("デモルールを適用しました。")
        else:
            with st.spinner("Gemini AIがルールを解析中..."):
                extracted_rules = parse_rules_with_gemini(api_key_input, custom_rule_text)
                if extracted_rules:
                    st.session_state.ai_rules = extracted_rules
                    st.success(f"🎉 AI解析成功！ {len(extracted_rules)}件 of ルールを適用しました。")
                else:
                    st.info("特殊ルールは検出されませんでした。標準の配置基準を適用します。")
                    st.session_state.ai_rules = []
                    
    if st.session_state.ai_rules:
        st.info("💡 **現在適用中のAI補正ルール:**\n" + \
                "\n".join([f"- **{r['start_hour']}:00～{r['end_hour']}:00**: 最低 {r['min_staff']}人配置 ({r['reason']})" for r in st.session_state.ai_rules]))

    st.divider()

    st.subheader("🧒 時間帯別の園児数設定 (30分刻み)")
    
    # 園児数の一括コピーボタン
    if st.button("📋 1行目の園児数をすべての時間帯にコピーする"):
        first_row = st.session_state.df_children_state.iloc[0]
        for col in ["0歳児数", "1-2歳児数", "3歳児数", "4歳以上児数"]:
            st.session_state.df_children_state[col] = first_row[col]
        st.success("1行目のデータを全ての時間帯にコピーしました！下の表で微調整できます。")
        st.rerun()

    # 編集可能なデータエディタ
    edited_children_df = st.data_editor(
        st.session_state.df_children_state, 
        num_rows="fixed", 
        key="children_editor_v2"
    )
    # セッション状態を更新
    st.session_state.df_children_state = edited_children_df
    
    # 配置基準＋AI補正計算
    def calculate_required_staff_with_ai(row):
        # 30分スロットから開始時間をパース
        current_hour = int(row["時間帯"].split(":")[0])
        
        count = (row["0歳児数"]/3) + (row["1-2歳児数"]/6) + (row["3歳児数"]/20) + (row["4歳以上児数"]/30)
        base_staff = math.ceil(count)
        
        min_staff_req = 1
        for rule in st.session_state.ai_rules:
            if rule["start_hour"] <= current_hour < rule["end_hour"]:
                min_staff_req = max(min_staff_req, rule["min_staff"])
                
        final_staff = max(base_staff, min_staff_req)
        return final_staff

    edited_children_df["必要保育士数"] = edited_children_df.apply(calculate_required_staff_with_ai, axis=1)
    
    st.subheader("⏰ 計算結果：時間帯別の必要保育士数")
    result_df = edited_children_df[["時間帯", "必要保育士数"]].set_index("時間帯").T
    st.dataframe(result_df)
    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# タブ2: 職員条件設定
# ==========================================
with tab2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.header("👥 保育士・スタッフの勤務条件設定")
    
    type_staff = st.selectbox("雇用形態", ["正社員", "パート"], key="reg_type_staff")
    
    with st.form("add_staff_form"):
        st.subheader("➕ 職員の新規登録")
        col1, col2, col3 = st.columns(3)
        with col1:
            name = st.text_input("お名前", placeholder="例：山田 花子")
            max_days = st.number_input("月間勤務上限日数", min_value=1, max_value=31, value=20 if type_staff == "正社員" else 12)
        with col2:
            time_zone = st.text_input(
                "勤務可能時間帯（自由記述）", 
                value="全日" if type_staff == "正社員" else "8:30-13:00 9:00-13:30", 
                placeholder="例: 8:30-13:00 9:00-13:30",
                help="『8:30-13:00 9:00-13:30』のようにスペースで区切って複数パターン入力できます。自動生成時にその日最も必要な時間が割り当てられます。"
            )
            short_time = st.selectbox("時短希望", ["なし", "あり"])
        with col3:
            default_days = ["月", "火", "水", "木", "金", "土"] if type_staff == "正社員" else ["月", "火", "水", "木", "金"]
            available_days = st.multiselect(
                "勤務可能曜日（チェック項目）", 
                ["月", "火", "水", "木", "金", "土", "日"], 
                default=default_days,
                help="正社員は自動で月〜土、パートは月〜金が選ばれます。"
            )
            submit = st.form_submit_button("この条件で職員を追加")
            
        if submit and name:
            days_str = ",".join(available_days)
            st.session_state.staff_list.append({
                "名前": name, 
                "雇用形態": type_staff, 
                "月上限日数": max_days, 
                "希望時間帯": time_zone, 
                "時短希望": short_time,
                "勤務可能曜日": days_str
            })
            st.success(f"✅ {name} さんを登録しました。")

    st.divider()

    st.subheader("📋 現在の職員一覧")
    st.caption("※一覧のデータをダブルクリックで修正できます。希望時間帯にスペース区切りで複数時間を入力可能です。")
    
    df_staff = pd.DataFrame(st.session_state.staff_list)
    edited_staff_df = st.data_editor(df_staff, num_rows="dynamic", key="staff_editor")
    
    if st.button("変更を確定して職員リストを保存"):
        st.session_state.staff_list = edited_staff_df.to_dict('records')
        st.success("✅ 職員リストの変更を確定しました。")
    
    if st.button("⚠️ 一覧をクリアして初期化"):
        st.session_state.staff_list = []
        st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# タブ3: シフト自動生成・Excel出力
# ==========================================
with tab3:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.header("🗓️ シフト自動生成エンジン")
    st.write(f"**対象期間: {target_year}年{target_month}月** ({num_days}日間)")
    
    # 凡例の表示
    st.info("""
    📖 **正社員のシフト記号凡例:**
    - **A**: 7:00 - 16:00
    - **B**: 8:00 - 17:00
    - **C**: 9:00 - 18:00
    - **D**: 10:00 - 19:00
    - **E**: 11:00 - 20:00
    - **公休**: 休み
    *(※パート職員は、入力された複数候補から最適な時間が選ばれ、そのまま時間が書き込まれます。)*
    """)
    
    # --- シフト生成アルゴリズム (均等ペース配分ソルバー) ---
    def generate_auto_schedule(df_kids, staff_list, target_year, target_month, num_days):
        # 30分ごとの必要人数を取得
        req_by_slot = {}
        for idx, row in df_kids.iterrows():
            req_by_slot[idx] = row["必要保育士数"]
            
        # 各職員の勤務日数カウンター
        staff_work_days = {staff["名前"]: 0 for staff in staff_list}
        
        # スケジュール初期化
        days_columns = [f"{d}日" for d in range(1, num_days + 1)]
        schedule_data = []
        for staff in staff_list:
            row = {"名前": staff["名前"], "雇用形態": staff["雇用形態"]}
            for col in days_columns:
                row[col] = "公休"
            schedule_data.append(row)
            
        schedule_df = pd.DataFrame(schedule_data)
        schedule_df.set_index("名前", inplace=True)
        
        # 直近の出勤日履歴を追跡するための辞書 (名前 -> list of bool [Day1..DayN])
        work_history = {staff["名前"]: [False] * (num_days + 1) for staff in staff_list}
        
        # 日ごとの割り当て処理
        for d in range(1, num_days + 1):
            day_col = f"{d}日"
            weekday = calendar.weekday(target_year, target_month, d)
            is_weekend = weekday in [5, 6]
            
            # その日の曜日文字（月、火...）
            weekday_map = ["月", "火", "水", "木", "金", "土", "日"]
            day_of_week = weekday_map[weekday]
            
            # その日の30分スロット別必要人数（土日は半分に削減）
            daily_req = {}
            for idx in range(26):
                req = req_by_slot.get(idx, 1)
                if is_weekend:
                    req = max(1, math.ceil(req * 0.5))
                daily_req[idx] = req
                
            # 出勤候補スタッフの取得
            def get_eligible_candidates(assigned_today):
                candidates = []
                for s in staff_list:
                    s_name = s["名前"]
                    if s_name in assigned_today:
                        continue
                    if staff_work_days[s_name] >= s["月上限日数"]:
                        continue
                        
                    # 勤務可能曜日のチェック
                    available_days_str = s.get("勤務可能曜日", "月,火,水,木,金,土")
                    available_days_list = [x.strip() for x in available_days_str.split(",") if x.strip()]
                    if day_of_week not in available_days_list:
                        continue
                        
                    # 勤務ペースのチェック（後半の公休だらけを避けるためのペース管理）
                    # 直近7日間の出勤日数
                    start_day = max(1, d - 6)
                    recent_work_days = sum(work_history[s_name][start_day:d])
                    
                    # 1週間の目安上限ペース = (月上限日数 / 30) * 7 （四捨五入）
                    weekly_pace_limit = math.ceil(s["月上限日数"] * 7 / 30)
                    
                    # 直近の連勤数を取得
                    consecutive_days = 0
                    check_day = d - 1
                    while check_day >= 1 and work_history[s_name][check_day]:
                        consecutive_days += 1
                        check_day -= 1
                    
                    # ペース制限を超えている、または5連勤以上の場合は、他の人がいれば休みを優先する優先度を計算
                    over_pace = recent_work_days >= weekly_pace_limit
                    too_many_consecutive = consecutive_days >= 5
                    
                    candidates.append({
                        "staff": s,
                        "work_days": staff_work_days[s_name],
                        "recent_work": recent_work_days,
                        "consecutive": consecutive_days,
                        "over_pace": over_pace,
                        "too_many_consecutive": too_many_consecutive
                    })
                return candidates

            assigned_today = {} # 名前 -> シフト名/記号
            
            # 各時間枠の割り当て人数を管理するカウンター
            assigned_counts = {idx: 0 for idx in range(26)}
            
            # 割り当てループ (その日の必要人数が充足するまで、または出勤可能スタッフがいなくなるまで)
            while True:
                # 30分スロットごとの不足人数を再計算
                understaffed_slots = []
                for idx in range(26):
                    deficit = daily_req[idx] - assigned_counts[idx]
                    if deficit > 0:
                        understaffed_slots.append((idx, deficit))
                        
                if not understaffed_slots:
                    break # 全て充足していれば終了
                    
                # 候補となるスタッフのリストを取得
                candidates = get_eligible_candidates(assigned_today)
                if not candidates:
                    break # 出勤可能者がいなくなれば終了
                    
                # 最も不足している時間帯スロット（不足数最大の時間）を見つける
                understaffed_slots.sort(key=lambda x: x[1], reverse=True)
                target_slot = understaffed_slots[0][0]
                
                # このターゲットスロットをカバーできるスタッフを探し、スコアを計算
                best_match = None
                best_score = -999999
                best_pattern = None
                
                for cand in candidates:
                    s = cand["staff"]
                    s_name = s["名前"]
                    role = s["雇用形態"]
                    
                    # 割り当て可能なシフトパターンの選択
                    patterns = [] # (シフト名, カバーするスロットインデックス)
                    
                    if role == "正社員":
                        if s["希望時間帯"] == "全日":
                            # A〜Eのすべてのパターンを候補とする
                            for pat, slots in REGULAR_PATTERNS.items():
                                patterns.append((pat, slots))
                        else:
                            # 個別指定がある場合
                            custom_slots_list = parse_work_hours(s["希望時間帯"])
                            for slots in custom_slots_list:
                                patterns.append((s["希望時間帯"], slots))
                    else: # パート
                        custom_slots_list = parse_work_hours(s["希望時間帯"])
                        # 自由入力内の各候補時間（例: "8:30-13:00"）を個別のシフト候補とする
                        parts = re.split(r'[\s,，、/]+', s["希望時間帯"])
                        for i, slots in enumerate(custom_slots_list):
                            label = parts[i] if i < len(parts) else s["希望時間帯"]
                            patterns.append((label, slots))
                            
                    # 各パターンについてスコアを計算
                    for pat_name, slots in patterns:
                        if target_slot not in slots:
                            continue # ターゲットスロットをカバーできないパターンは除外
                            
                        # 不足時間帯をどれだけカバーできるか（カバー度）
                        overlap = sum(1 for sl in slots if daily_req[sl] > assigned_counts[sl])
                        
                        # スコア計算:
                        # + カバー度 (高いほど良い)
                        # - 月間上限に近さ (出勤日数が上限に遠い人を優先)
                        # - 出勤ペース制限 (ペースを守っている人を優先)
                        # - 連勤ペナルティ (連勤数が少ない人を優先)
                        score = overlap * 10
                        score += (s["月上限日数"] - cand["work_days"]) * 5
                        score -= cand["recent_work"] * 3
                        score -= cand["consecutive"] * 2
                        
                        if cand["over_pace"]:
                            score -= 15 # 週ペースオーバーへのペナルティ
                        if cand["too_many_consecutive"]:
                            score -= 30 # 5連勤以上の過度な連勤への重いペナルティ
                            
                        if score > best_score:
                            best_score = score
                            best_match = cand
                            best_pattern = (pat_name, slots)
                            
                # マッチするスタッフ＋パターンが見つかった場合、割り当てる
                if best_match and best_pattern:
                    s_name = best_match["staff"]["名前"]
                    pat_name, slots = best_pattern
                    
                    assigned_today[s_name] = pat_name
                    staff_work_days[s_name] += 1
                    work_history[s_name][d] = True
                    
                    # 割り当て人数カウンターを更新
                    for sl in slots:
                        assigned_counts[sl] += 1
                else:
                    # ターゲットスロットを誰もカバーできない場合、そのスロットは諦めて2番目に不足しているスロットを狙う
                    if len(understaffed_slots) > 1:
                        target_slot = understaffed_slots[1][0]
                        # 2番目のターゲットでリトライするためループを抜けない
                        # ただし、無限ループを防ぐため、候補者全員でターゲットをカバーできない場合は終了させる
                        break
                    else:
                        break
                        
            # 全員分の今日のシフトをデータフレームに書き込み
            for s_name, shift in assigned_today.items():
                schedule_df.at[s_name, day_col] = shift
                
        schedule_df.reset_index(inplace=True)
        return schedule_df

    # シフト生成ボタン
    if len(st.session_state.staff_list) == 0:
        st.warning("⚠️ 職員が登録されていません。「職員条件設定」タブから職員を追加してください。")
    else:
        if st.button("⚡ シフトを自動生成する"):
            with st.spinner("ペース配分および出勤曜日制限に基づき、最適なシフトを計算中..."):
                schedule_result = generate_auto_schedule(
                    edited_children_df, 
                    st.session_state.staff_list, 
                    target_year, 
                    target_month, 
                    num_days
                )
                st.session_state.schedule_result = schedule_result
                st.success("🎉 シフト表を自動作成しました！")

        # 生成結果があれば表示
        if 'schedule_result' in st.session_state:
            st.subheader("📅 作成されたシフト表 (手動微調整も可能)")
            st.caption("※セルをダブルクリックすることで、手動で任意の勤務時間や記号（A〜E）に直接書き換えることもできます。")
            
            edited_schedule = st.data_editor(
                st.session_state.schedule_result, 
                key="schedule_editor", 
                num_rows="fixed"
            )
            
            if st.button("シフト表の調整内容を保存"):
                st.session_state.schedule_result = edited_schedule
                st.success("✅ 調整されたシフトを保存しました！")

            # 統計情報の表示
            st.subheader("📊 職員ごとの出勤日数カウント")
            stat_data = []
            for idx, row in edited_schedule.iterrows():
                work_days_count = sum(1 for col in edited_schedule.columns if "日" in col and row[col] != "公休")
                stat_data.append({
                    "名前": row["名前"],
                    "雇用形態": row["雇用形態"],
                    "当月出勤日数": work_days_count
                })
            st.dataframe(pd.DataFrame(stat_data))

            # ==========================================
            # Excel出力セクション
            # ==========================================
            st.divider()
            st.subheader("💾 完成したシフト表をダウンロード")
            
            def generate_styled_excel(df_kids, df_teachers, df_schedule, year, month):
                wb = openpyxl.Workbook()
                
                # スタイル定義
                font_title = Font(name="Meiryo UI", size=14, bold=True, color="2F3542")
                font_header = Font(name="Meiryo UI", size=10, bold=True, color="FFFFFF")
                font_data = Font(name="Meiryo UI", size=9)
                font_bold = Font(name="Meiryo UI", size=9, bold=True)
                font_legend = Font(name="Meiryo UI", size=9, italic=True, color="57606F")
                
                fill_primary = PatternFill(start_color="FF6B8B", end_color="FF6B8B", fill_type="solid")
                fill_secondary = PatternFill(start_color="70A1FF", end_color="70A1FF", fill_type="solid")
                fill_weekend = PatternFill(start_color="F1F2F6", end_color="F1F2F6", fill_type="solid")
                fill_off = PatternFill(start_color="E4E7EB", end_color="E4E7EB", fill_type="solid")
                
                border_thin = Border(
                    left=Side(style='thin', color='CED6E0'),
                    right=Side(style='thin', color='CED6E0'),
                    top=Side(style='thin', color='CED6E0'),
                    bottom=Side(style='thin', color='CED6E0')
                )
                
                align_center = Alignment(horizontal="center", vertical="center")
                align_left = Alignment(horizontal="left", vertical="center")
                
                # --- シート1: シフト表 (メイン) ---
                ws_schedule = wb.active
                ws_schedule.title = "シフト表"
                ws_schedule.views.sheetView[0].showGridLines = True
                
                # タイトル
                ws_schedule.cell(row=1, column=1, value=f"📛 {year}年{month}月 保育士シフト表").font = font_title
                ws_schedule.row_dimensions[1].height = 30
                
                headers = list(df_schedule.columns)
                ws_schedule.append([])
                ws_schedule.append(headers + ["出勤日数"])
                ws_schedule.row_dimensions[3].height = 25
                
                # ヘッダースタイル
                for c_idx in range(1, len(headers) + 2):
                    cell = ws_schedule.cell(row=3, column=c_idx)
                    cell.fill = fill_primary
                    cell.font = font_header
                    cell.alignment = align_center
                    cell.border = border_thin
                
                # データ書き込み
                last_row_idx = 3
                for r_idx, row in enumerate(df_schedule.values, 4):
                    row_list = list(row)
                    ws_schedule.append(row_list)
                    ws_schedule.row_dimensions[r_idx].height = 22
                    last_row_idx = r_idx
                    
                    last_col_letter = get_column_letter(len(headers))
                    formula = f'=COUNTIF(C{r_idx}:{last_col_letter}{r_idx}, "<>公休")'
                    ws_schedule.cell(row=r_idx, column=len(headers) + 1, value=formula).font = font_bold
                    ws_schedule.cell(row=r_idx, column=len(headers) + 1).alignment = align_center
                    ws_schedule.cell(row=r_idx, column=len(headers) + 1).border = border_thin
                    
                    for c_idx in range(1, len(headers) + 1):
                        cell = ws_schedule.cell(row=r_idx, column=c_idx)
                        cell.font = font_data
                        cell.border = border_thin
                        
                        if c_idx == 1:
                            cell.alignment = align_left
                        else:
                            cell.alignment = align_center
                            
                        val = cell.value
                        if val == "公休":
                            cell.fill = fill_off
                            cell.font = Font(name="Meiryo UI", size=9, color="747D8C")
                        elif val in ["A", "B", "C", "D", "E"]:
                            cell.font = Font(name="Meiryo UI", size=9, bold=True)
                            cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid") # 正社員記号は薄黄色
                        elif val:
                            cell.font = Font(name="Meiryo UI", size=9, bold=True)
                            
                # 列幅調整
                for col in ws_schedule.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws_schedule.column_dimensions[col_letter].width = max(max_len + 3, 11)
                
                # シフト記号の凡例（シート下部に追加）
                legend_start_row = last_row_idx + 3
                ws_schedule.cell(row=legend_start_row, column=1, value="📖 シフト記号凡例:").font = font_bold
                legend_items = [
                    ("A", "7:00 - 16:00"),
                    ("B", "8:00 - 17:00"),
                    ("C", "9:00 - 18:00"),
                    ("D", "10:00 - 19:00"),
                    ("E", "11:00 - 20:00")
                ]
                for idx, (sym, tm) in enumerate(legend_items):
                    row_num = legend_start_row + 1 + idx
                    ws_schedule.cell(row=row_num, column=1, value=f"  {sym} : {tm}").font = font_legend
                
                # --- シート2: 必要人数要件 ---
                ws_kids = wb.create_sheet(title="時間帯別必要人数")
                ws_kids.views.sheetView[0].showGridLines = True
                
                ws_kids.cell(row=1, column=1, value="📊 時間帯別 園児数と必要人数").font = font_title
                ws_kids.append([])
                
                for r_idx, row in enumerate(openpyxl.utils.dataframe.dataframe_to_rows(df_kids, index=False, header=True), 3):
                    ws_kids.append(row)
                    ws_kids.row_dimensions[r_idx].height = 20
                    for c_idx in range(1, len(row) + 1):
                        cell = ws_kids.cell(row=r_idx, column=c_idx)
                        cell.border = border_thin
                        cell.alignment = align_center
                        if r_idx == 3:
                            cell.fill = fill_secondary
                            cell.font = font_header
                        else:
                            cell.font = font_data
                
                for col in ws_kids.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws_kids.column_dimensions[col_letter].width = max(max_len + 5, 12)
                
                # --- シート3: 職員マスター ---
                ws_staff = wb.create_sheet(title="職員マスター")
                ws_staff.views.sheetView[0].showGridLines = True
                
                ws_staff.cell(row=1, column=1, value="👥 職員（スタッフ）勤務条件一覧").font = font_title
                ws_staff.append([])
                
                for r_idx, row in enumerate(openpyxl.utils.dataframe.dataframe_to_rows(df_teachers, index=False, header=True), 3):
                    ws_staff.append(row)
                    ws_staff.row_dimensions[r_idx].height = 20
                    for c_idx in range(1, len(row) + 1):
                        cell = ws_staff.cell(row=r_idx, column=c_idx)
                        cell.border = border_thin
                        cell.alignment = align_center
                        if r_idx == 3:
                            cell.fill = fill_secondary
                            cell.font = font_header
                        else:
                            cell.font = font_data
                            if c_idx == 1:
                                cell.alignment = align_left
                
                for col in ws_staff.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws_staff.column_dimensions[col_letter].width = max(max_len + 5, 15)

                excel_buffer = BytesIO()
                wb.save(excel_buffer)
                excel_buffer.seek(0)
                return excel_buffer

            excel_data = generate_styled_excel(edited_children_df, df_staff, edited_schedule, target_year, target_month)
            
            st.download_button(
                label="📥 完成したシフト管理Excelファイルをダウンロード",
                data=excel_data,
                file_name=f"保育園シフト管理表_{target_year}年{target_month}月.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    st.markdown('</div>', unsafe_allow_html=True)
