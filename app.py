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
# 🔑 3. 時間帯表記揺れパーサーの実装
# ==========================================
def parse_work_hours(tz_str):
    """
    "8:00〜13:30" や "9;00-13;30" などの表記揺れを含む文字列から、
    カバーする時間のリスト(整数)を返す関数。
    例: "8:00〜13:30" -> [8, 9, 10, 11, 12, 13]
    """
    if not tz_str or not isinstance(tz_str, str):
        return []
    
    if tz_str == "全日":
        return list(range(7, 19))
        
    s = tz_str.translate(str.maketrans({
        '０':'0', '１':'1', '２':'2', '３':'3', '４':'4', '５':'5', '６':'6', '７':'7', '８':'8', '９':'9',
        '；':':', ';':':', '：':':',
        '〜':'-', '～':'-', 'ー':'-', '－':'-', '—':'-', 'ー':'-', '─':'-'
    })).strip()
    
    # 1. H:MM - H:MM 形式の抽出
    match = re.search(r'(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})', s)
    if match:
        start_h = int(match.group(1))
        start_m = int(match.group(2))
        end_h = int(match.group(3))
        end_m = int(match.group(4))
        
        if end_m > 0:
            return list(range(start_h, end_h + 1))
        else:
            return list(range(start_h, end_h))
            
    # 2. H - H 形式の抽出
    match_hour_only = re.search(r'(\d{1,2})\s*-\s*(\d{1,2})', s)
    if match_hour_only:
        start_h = int(match_hour_only.group(1))
        end_h = int(match_hour_only.group(2))
        return list(range(start_h, end_h))
        
    return list(range(9, 17))

def get_covered_slots(shift_name):
    if shift_name == "フル":
        return list(range(8, 18))
    elif shift_name == "フル早":
        return list(range(7, 16))
    elif shift_name == "フル遅":
        return list(range(10, 19))
    elif shift_name == "早":
        return list(range(7, 11))
    elif shift_name == "日":
        return list(range(9, 17))
    elif shift_name == "遅":
        return list(range(15, 19))
    elif shift_name == "公休":
        return []
    return parse_work_hours(shift_name)

# ==========================================
# 📂 4. セッション状態の初期化
# ==========================================
if 'staff_list' not in st.session_state:
    st.session_state.staff_list = [
        {"名前": "山田 花子", "雇用形態": "正社員", "月上限日数": 20, "希望時間帯": "全日", "時短希望": "なし", "勤務可能曜日": "月,火,水,木,金,土"},
        {"名前": "鈴木 一郎", "雇用形態": "パート", "月上限日数": 12, "希望時間帯": "8:00〜13:30", "時短希望": "あり", "勤務可能曜日": "月,火,水,木,金"},
        {"名前": "佐藤 美咲", "雇用形態": "パート", "月上限日数": 15, "希望時間帯": "15:00-19:00", "時短希望": "なし", "勤務可能曜日": "月,火,水,木,金"},
        {"名前": "田中 恵子", "雇用形態": "正社員", "月上限日数": 20, "希望時間帯": "全日", "時短希望": "なし", "勤務可能曜日": "月,火,水,木,金,土"},
        {"名前": "小林 翔太", "雇用形態": "正社員", "月上限日数": 22, "希望時間帯": "全日", "時短希望": "なし", "勤務可能曜日": "月,火,水,木,金,土"},
        {"名前": "高橋 陽子", "雇用形態": "パート", "月上限日数": 16, "希望時間帯": "9:00-17:00", "時短希望": "なし", "勤務可能曜日": "月,火,水,木,金"},
        {"名前": "渡辺 理恵", "雇用形態": "パート", "月上限日数": 10, "希望時間帯": "7:00〜12:00", "時短希望": "なし", "勤務可能曜日": "月,火,水,木,金"},
        {"名前": "伊藤 直美", "雇用形態": "パート", "月上限日数": 14, "希望時間帯": "13:30-19:00", "時短希望": "なし", "勤務可能曜日": "月,火,水,木,金"},
    ]

if 'ai_rules' not in st.session_state:
    st.session_state.ai_rules = []

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

    st.subheader("🧒 時間帯別の園児数設定")
    st.caption("※時間帯ごとの園児数を入力します。ダブルクリックで数値を直接変更できます。")

    time_slots = [f"{h}:00～{h+1}:00" for h in range(7, 19)]
    init_data = {
        "時間帯": time_slots,
        "0歳児数": [3, 5, 6, 6, 6, 6, 6, 6, 5, 3, 2, 1],
        "1-2歳児数": [5, 10, 12, 12, 12, 12, 12, 12, 10, 8, 5, 2],
        "3歳児数": [2, 15, 18, 18, 18, 18, 18, 18, 15, 10, 4, 1],
        "4歳以上児数": [5, 20, 22, 22, 22, 22, 22, 22, 20, 15, 8, 2]
    }
    df_children = pd.DataFrame(init_data)
    edited_df = st.data_editor(df_children, num_rows="dynamic", key="children_editor")
    
    def calculate_required_staff_with_ai(row):
        current_hour = int(row["時間帯"].split(":")[0])
        
        count = (row["0歳児数"]/3) + (row["1-2歳児数"]/6) + (row["3歳児数"]/20) + (row["4歳以上児数"]/30)
        base_staff = math.ceil(count)
        
        min_staff_req = 1
        for rule in st.session_state.ai_rules:
            if rule["start_hour"] <= current_hour < rule["end_hour"]:
                min_staff_req = max(min_staff_req, rule["min_staff"])
                
        final_staff = max(base_staff, min_staff_req)
        return final_staff

    edited_df["必要保育士数"] = edited_df.apply(calculate_required_staff_with_ai, axis=1)
    
    st.subheader("⏰ 計算結果：時間帯別の必要保育士数")
    result_df = edited_df[["時間帯", "必要保育士数"]].set_index("時間帯").T
    st.dataframe(result_df)
    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================
# タブ2: 職員条件設定
# ==========================================
with tab2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.header("👥 保育士・スタッフの勤務条件設定")
    
    # 雇用形態の選択をフォーム外に置くか、フォーム内でステート連動させるためにセッションを活用
    # 雇用形態の変更を即時検知するため、selectboxにkeyを持たせて状態管理
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
                value="全日", 
                placeholder="例: 8:00〜13:30, 9:00-17:00, 全日",
                help="『8:00〜13:30』や『9:00-13:30』のように任意の時間を入力できます。『全日』と書くと早番から遅番まで全ての時間帯の勤務候補になります。"
            )
            short_time = st.selectbox("時短希望", ["なし", "あり"])
        with col3:
            # 雇用形態（type_staff）に連動したデフォルト曜日を設定
            default_days = ["月", "火", "水", "木", "金", "土"] if type_staff == "正社員" else ["月", "火", "水", "木", "金"]
            available_days = st.multiselect(
                "勤務可能曜日（チェック項目）", 
                ["月", "火", "水", "木", "金", "土", "日"], 
                default=default_days,
                help="職員が出勤可能な曜日を選択します。正社員の場合は自動的に『月〜土』が選ばれ、パートの場合は『月〜金』が初期選択されます。"
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
            st.success(f"✅ {name} さんを登録しました。 (曜日: {days_str})")

    st.divider()

    st.subheader("📋 現在の職員一覧")
    st.caption("※一覧に表示されるデータをダブルクリックで直接修正・追加入力できます（曜日を変更する際はカンマ『,』で区切って入力してください）。")
    
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
    
    def generate_auto_schedule(df_kids, staff_list, target_year, target_month, num_days):
        req_by_hour = {}
        for idx, row in df_kids.iterrows():
            hour = int(row["時間帯"].split(":")[0])
            req_by_hour[hour] = row["必要保育士数"]
            
        staff_work_days = {staff["名前"]: 0 for staff in staff_list}
        
        days_columns = [f"{d}日" for d in range(1, num_days + 1)]
        schedule_data = []
        for staff in staff_list:
            row = {"名前": staff["名前"], "雇用形態": staff["雇用形態"]}
            for col in days_columns:
                row[col] = "公休"
            schedule_data.append(row)
            
        schedule_df = pd.DataFrame(schedule_data)
        schedule_df.set_index("名前", inplace=True)
        
        for d in range(1, num_days + 1):
            day_col = f"{d}日"
            weekday = calendar.weekday(target_year, target_month, d)
            is_weekend = weekday in [5, 6]
            
            # その日の曜日文字（月、火...）を取得
            weekday_map = ["月", "火", "水", "木", "金", "土", "日"]
            day_of_week = weekday_map[weekday]
            
            daily_req = {}
            for h in range(7, 19):
                req = req_by_hour.get(h, 1)
                if is_weekend:
                    req = max(1, math.ceil(req * 0.5))
                daily_req[h] = req
                
            def get_eligible_staff(assigned_today):
                eligible = []
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
                        continue # この曜日に出勤できないスタッフは割り当て候補から除外
                        
                    eligible.append((s, staff_work_days[s_name]))
                
                eligible.sort(key=lambda x: x[1])
                return [x[0] for x in eligible]

            assigned_today = {}
            
            # パス1：早朝（7:00および8:00）をカバーするスタッフ
            needed_morning = max(daily_req.get(7, 1), daily_req.get(8, 1))
            morning_candidates = []
            for s in get_eligible_staff(assigned_today):
                pref = s["希望時間帯"]
                covered = parse_work_hours(pref)
                if 7 in covered or 8 in covered:
                    morning_candidates.append(s)
            
            for s in morning_candidates[:needed_morning]:
                s_name = s["名前"]
                if s["雇用形態"] == "正社員" and s["希望時間帯"] == "全日":
                    assigned_today[s_name] = "フル早"
                else:
                    assigned_today[s_name] = s["希望時間帯"]
                staff_work_days[s_name] += 1
                
            # パス2：遅番（17:00および18:00）をカバーするスタッフ
            needed_evening = max(daily_req.get(17, 1), daily_req.get(18, 1))
            evening_candidates = []
            for s in get_eligible_staff(assigned_today):
                pref = s["希望時間帯"]
                covered = parse_work_hours(pref)
                if 17 in covered or 18 in covered:
                    evening_candidates.append(s)
            
            for s in evening_candidates[:needed_evening]:
                s_name = s["名前"]
                if s["雇用形態"] == "正社員" and s["希望時間帯"] == "全日":
                    assigned_today[s_name] = "フル遅"
                else:
                    assigned_today[s_name] = s["希望時間帯"]
                staff_work_days[s_name] += 1
                
            # パス3：日中時間帯（9:00〜16:00）
            max_mid_need = 0
            for h in range(9, 17):
                current_assigned = sum(1 for name, sh in assigned_today.items() if h in get_covered_slots(sh))
                need = max(0, daily_req[h] - current_assigned)
                if need > max_mid_need:
                    max_mid_need = need
                    
            mid_candidates = []
            for s in get_eligible_staff(assigned_today):
                pref = s["希望時間帯"]
                covered = parse_work_hours(pref)
                if any(hour in covered for hour in range(9, 17)):
                    mid_candidates.append(s)
            
            for s in mid_candidates[:max_mid_need]:
                s_name = s["名前"]
                if s["雇用形態"] == "正社員" and s["希望時間帯"] == "全日":
                    assigned_today[s_name] = "フル"
                else:
                    assigned_today[s_name] = s["希望時間帯"]
                staff_work_days[s_name] += 1
                
            for s_name, shift in assigned_today.items():
                schedule_df.at[s_name, day_col] = shift
                
        schedule_df.reset_index(inplace=True)
        return schedule_df

    if len(st.session_state.staff_list) == 0:
        st.warning("⚠️ 職員が登録されていません。「職員条件設定」タブから職員を追加してください。")
    else:
        if st.button("⚡ シフトを自動生成する"):
            with st.spinner("カレンダー要件、個別勤務時間および出勤曜日制限に基づき、最適なシフトを計算中..."):
                schedule_result = generate_auto_schedule(
                    edited_df, 
                    st.session_state.staff_list, 
                    target_year, 
                    target_month, 
                    num_days
                )
                st.session_state.schedule_result = schedule_result
                st.success("🎉 シフト表を自動作成しました！")

        if 'schedule_result' in st.session_state:
            st.subheader("📅 作成されたシフト表 (手動微調整も可能)")
            st.caption("※セルをダブルクリックすることで、手動で任意の勤務時間に直接書き換えることもできます。")
            
            edited_schedule = st.data_editor(
                st.session_state.schedule_result, 
                key="schedule_editor", 
                num_rows="fixed"
            )
            
            if st.button("シフト表の調整内容を保存"):
                st.session_state.schedule_result = edited_schedule
                st.success("✅ 調整されたシフトを保存しました！")

            st.subheader("📊 職員ごとの出勤日数カウント")
            stat_data = []
            for idx, row in edited_schedule.iterrows():
                work_days_count = sum(1 for col in edited_schedule.columns if "日" in col and row[col] != "公休")
                stat_data.append({
                    "名前": row["名前"],
                    "雇用形態": row["雇用形態"],
                    "登録された希望時間": row["希望時間帯"] if "希望時間帯" in edited_schedule.columns else "個別設定",
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
                
                font_title = Font(name="Meiryo UI", size=14, bold=True, color="2F3542")
                font_header = Font(name="Meiryo UI", size=10, bold=True, color="FFFFFF")
                font_data = Font(name="Meiryo UI", size=9)
                font_bold = Font(name="Meiryo UI", size=9, bold=True)
                
                fill_primary = PatternFill(start_color="FF6B8B", end_color="FF6B8B", fill_type="solid")
                fill_secondary = PatternFill(start_color="70A1FF", end_color="70A1FF", fill_type="solid")
                fill_weekend = PatternFill(start_color="F1F2F6", end_color="F1F2F6", fill_type="solid")
                fill_holiday = PatternFill(start_color="FFE3E8", end_color="FFE3E8", fill_type="solid")
                fill_off = PatternFill(start_color="E4E7EB", end_color="E4E7EB", fill_type="solid")
                
                border_thin = Border(
                    left=Side(style='thin', color='CED6E0'),
                    right=Side(style='thin', color='CED6E0'),
                    top=Side(style='thin', color='CED6E0'),
                    bottom=Side(style='thin', color='CED6E0')
                )
                
                align_center = Alignment(horizontal="center", vertical="center")
                align_left = Alignment(horizontal="left", vertical="center")
                
                ws_schedule = wb.active
                ws_schedule.title = "シフト表"
                ws_schedule.views.sheetView[0].showGridLines = True
                
                ws_schedule.cell(row=1, column=1, value=f"📛 {year}年{month}月 保育士シフト表").font = font_title
                ws_schedule.row_dimensions[1].height = 30
                
                headers = list(df_schedule.columns)
                ws_schedule.append([])
                ws_schedule.append(headers + ["出勤日数"])
                ws_schedule.row_dimensions[3].height = 25
                
                for c_idx in range(1, len(headers) + 2):
                    cell = ws_schedule.cell(row=3, column=c_idx)
                    cell.fill = fill_primary
                    cell.font = font_header
                    cell.alignment = align_center
                    cell.border = border_thin
                
                for r_idx, row in enumerate(df_schedule.values, 4):
                    row_list = list(row)
                    ws_schedule.append(row_list)
                    ws_schedule.row_dimensions[r_idx].height = 22
                    
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
                        elif val:
                            cell.font = Font(name="Meiryo UI", size=9, bold=True)
                            
                for col in ws_schedule.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = get_column_letter(col[0].column)
                    ws_schedule.column_dimensions[col_letter].width = max(max_len + 3, 11)
                
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

            excel_data = generate_styled_excel(edited_df, df_staff, edited_schedule, target_year, target_month)
            
            st.download_button(
                label="📥 完成したシフト管理Excelファイルをダウンロード",
                data=excel_data,
                file_name=f"保育園シフト管理表_{target_year}年{target_month}月.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    st.markdown('</div>', unsafe_allow_html=True)
