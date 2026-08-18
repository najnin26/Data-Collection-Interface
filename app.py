import uuid
from datetime import datetime
from pathlib import Path

import openpyxl
import streamlit as st
from openpyxl import Workbook, load_workbook

# ------------------------------------------------------------------------------
# App Config & Constants
# ------------------------------------------------------------------------------
st.set_page_config(
    page_title="Bangla Pragmatics Data Collection",
    page_icon="🗣️",
    layout="wide",
)

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
EXCEL_FILE = DATA_DIR / "bangla_pragmatics_responses.xlsx"

SCENARIOS = [
    (
        "Direct Request",
        (
            "আপনি আপনার ঘরে পড়াশোনা করছেন। পাশের ঘরে আপনার ছোট ভাই/বোন খুব"
            " জোরে গান বাজাচ্ছে এবং আপনি মনোযোগ দিতে পারছেন না।"
        ),
    ),
    (
        "Indirect Request",
        (
            "আপনি একজন বন্ধুর সঙ্গে একটি ঘরে বসে আছেন। জানালা খোলা এবং বাইরে"
            " অনেক ঠান্ডা।"
        ),
    ),
    (
        "Asking Someone to Move",
        (
            "আপনি বাসে বসতে যাচ্ছেন। পাশের আসনে একজন তার ব্যাগ রেখেছে এবং আপনি"
            " সেখানে বসতে চান।"
        ),
    ),
    (
        "Request for Help",
        "আপনার হাতে কয়েকটি ভারী ব্যাগ এবং আপনার বন্ধু কাছেই দাঁড়িয়ে আছে।",
    ),
    (
        "Indirect Request to Stop",
        (
            "গুরুত্বপূর্ণ পরীক্ষার প্রস্তুতির সময় আপনার বন্ধু বারবার আপনাকে"
            " মেসেজ পাঠাচ্ছে।"
        ),
    ),
    (
        "Direct Refusal",
        (
            "আপনার বন্ধু আপনাকে একটি পার্টিতে যেতে বলেছে, কিন্তু পরদিন সকালে"
            " আপনার গুরুত্বপূর্ণ পরীক্ষা।"
        ),
    ),
    (
        "Indirect Refusal",
        (
            "একজন বন্ধু জিজ্ঞেস করল: “কাল আমার সাথে ঘুরতে যাবে?” কিন্তু আপনার"
            " আগে থেকেই অন্য একটি কাজ আছে।"
        ),
    ),
    (
        "Polite Refusal",
        (
            "একজন সহকর্মী আপনাকে আজ অতিরিক্ত কাজ করতে বলেছে, কিন্তু আপনি"
            " ইতিমধ্যে অনেক কাজের চাপে আছেন।"
        ),
    ),
    (
        "Avoiding an Invitation",
        (
            "আপনাকে এমন একজন ব্যক্তি ডিনারে আমন্ত্রণ জানিয়েছে যাকে আপনি খুব বেশি"
            " চেনেন না, কিন্তু আপনি যেতে চান না।"
        ),
    ),
    (
        "Refusing to Lend",
        (
            "একজন বন্ধু আপনার ল্যাপটপ ধার চাইছে, কিন্তু আপনি ল্যাপটপটি দিতে চান"
            " না।"
        ),
    ),
    (
        "Sarcasm: Late Arrival",
        (
            "আপনার বন্ধু বিকেল ৫টায় আসার কথা বলেছিল, কিন্তু সে রাত ৭টায় এসে"
            " পৌঁছাল।"
        ),
    ),
    (
        "Sarcasm: Poor Performance",
        (
            "আপনার বন্ধু একটি কাজকে খুব সহজ বলেছিল, কিন্তু শেষ পর্যন্ত সে কাজটি"
            " সম্পূর্ণ করতে ব্যর্থ হয়েছে।"
        ),
    ),
    (
        "Sarcasm: Messy Room",
        (
            "আপনার ভাই/বোন ঘর পরিষ্কার করবে বলেছিল, কিন্তু ঘরটি আরও বেশি এলোমেলো"
            " করে ফেলেছে।"
        ),
    ),
    (
        "Sarcasm: Broken Promise",
        (
            "আপনার বন্ধু কয়েকবার আপনাকে ফোন করবে বলেছিল, কিন্তু কখনো ফোন করেনি।"
            " অবশেষে সে ফোন করল।"
        ),
    ),
    (
        "Sarcasm: Obvious Mistake",
        (
            "আপনার বন্ধু একটি খুব স্পষ্ট ভুল করেছে এবং বলছে, “আমি কখনো ভুল করি"
            " না।”"
        ),
    ),
    (
        "Irony: Rain",
        (
            "আপনি সারাদিনের জন্য একটি বাইরের অনুষ্ঠান আয়োজন করেছেন, কিন্তু"
            " অনুষ্ঠান শুরুর সময় প্রচণ্ড বৃষ্টি শুরু হলো।"
        ),
    ),
    (
        "Irony: Technology Failure",
        (
            "আপনি নতুন একটি ফোন কিনেছেন কারণ আপনি আশা করেছিলেন এটি খুব ভালো কাজ"
            " করবে, কিন্তু প্রথম দিনেই ফোনটি নষ্ট হয়ে গেল।"
        ),
    ),
    (
        "Irony: Perfect Timing",
        "আপনি বাসস্ট্যান্ডে পৌঁছানোর ঠিক পরেই আপনার বাসটি ছেড়ে দিল।",
    ),
    (
        "Irony: Group Project",
        (
            "আপনার গ্রুপের সদস্যরা প্রায় কোনো কাজ না করে শেষ মুহূর্তে ডেডলাইনের"
            " এক মিনিট আগে অ্যাসাইনমেন্ট জমা দিল।"
        ),
    ),
    (
        "Irony: Another Problem",
        (
            "আপনি অনেক কষ্টে একটি কঠিন কাজ শেষ করলেন, কিন্তু সঙ্গে সঙ্গেই আরেকটি"
            " বড় সমস্যা তৈরি হলো।"
        ),
    ),
    (
        "Rhetorical Question: Repeated Mistake",
        (
            "আপনার বন্ধু একই ভুল বারবার করছে, যদিও আপনি তাকে বিষয়টি অনেকবার"
            " বুঝিয়েছেন।"
        ),
    ),
    (
        "Rhetorical Question: Exam",
        (
            "একজন ব্যক্তি পরীক্ষার আগের দিন পড়াশোনা না করে অভিযোগ করছে যে তাকে"
            " কেন পড়তে হবে।"
        ),
    ),
    (
        "Rhetorical Question: Broken Object",
        (
            "আপনার ভাই/বোন কিছু ভেঙে ফেলেছে এবং এখন এমন আচরণ করছে যেন সে কিছুই"
            " জানে না।"
        ),
    ),
    (
        "Rhetorical Question: Tiredness",
        (
            "আপনি সারাদিন কাজ করার পর ক্লান্ত, কিন্তু আপনার বন্ধু বারবার জিজ্ঞেস"
            " করছে কেন আপনি ক্লান্ত।"
        ),
    ),
    (
        "Rhetorical Question: Ignored Advice",
        (
            "আপনি কাউকে বারবার সতর্ক করেছিলেন, সে আপনার কথা শোনেনি এবং এখন ঠিক"
            " সেই সমস্যায় পড়েছে।"
        ),
    ),
    (
        "Genuine Praise: Project",
        "আপনার বন্ধু একটি কঠিন প্রজেক্ট সফলভাবে সম্পন্ন করেছে।",
    ),
    (
        "Genuine Praise: Academic Result",
        "আপনার ছোট ভাই/বোন পরীক্ষায় খুব ভালো ফল করেছে।",
    ),
    (
        "Appreciation: Help",
        "একজন সহকর্মী আপনাকে একটি গুরুত্বপূর্ণ কাজ শেষ করতে সাহায্য করেছে।",
    ),
    (
        "Praise: Appearance",
        "আপনার বন্ধু নতুন একটি পোশাক পরেছে এবং তাকে খুব সুন্দর দেখাচ্ছে।",
    ),
    (
        "Praise: Presentation",
        "আপনি আপনার বন্ধুর একটি খুব ভালো প্রেজেন্টেশন শুনলেন।",
    ),
    (
        "Complaint: Noise",
        "আপনার প্রতিবেশী সপ্তাহে কয়েকবার গভীর রাত পর্যন্ত জোরে গান বাজায়।",
    ),
    (
        "Complaint: Late Delivery",
        (
            "আপনি অনলাইনে কিছু অর্ডার করেছিলেন, কিন্তু প্রতিশ্রুত সময়ের কয়েক দিন"
            " পরে সেটি এসেছে।"
        ),
    ),
    (
        "Complaint: Group Member",
        (
            "আপনার গ্রুপের একজন সদস্য বারবার তার দায়িত্বের কাজ শেষ করছে না।"
        ),
    ),
    (
        "Complaint: Restaurant",
        (
            "আপনি রেস্টুরেন্টে খাবার অর্ডার করেছেন, কিন্তু খাবারটি আপনার প্রত্যাশার"
            " তুলনায় খুব খারাপ।"
        ),
    ),
    (
        "Complaint: Cancelled Plans",
        "আপনার বন্ধু বারবার শেষ মুহূর্তে আপনাদের পরিকল্পনা বাতিল করছে।",
    ),
    ("Ambiguous: Good News", "আপনার বন্ধু সবেমাত্র একটি খুব ভালো খবর পেয়েছে।"),
    (
        "Ambiguous: Serious Mistake",
        "আপনার বন্ধু সবেমাত্র একটি গুরুতর ভুল করেছে।",
    ),
    (
        "Ambiguous: Uncertain Help",
        (
            "কেউ আপনাকে সাহায্য করবে বলেছে, কিন্তু সে সত্যিই সাহায্য করবে কি না তা"
            " নিয়ে আপনি নিশ্চিত নন।"
        ),
    ),
    (
        "Ambiguous: Waiting",
        (
            "আপনি এমন একজনের জন্য অপেক্ষা করছেন যে ইতিমধ্যে ৩০ মিনিট দেরি করেছে এবং"
            " সে এখন এসে পৌঁছেছে।"
        ),
    ),
    (
        "Ambiguous: Indirect No",
        (
            "কেউ আপনাকে এমন কিছু করতে বলেছে যা আপনি করতে চান না, কিন্তু আপনি"
            " সরাসরি “না” বলতে চান না।"
        ),
    ),
    (
        "Doubt",
        (
            "আপনার বন্ধু বলছে সে মাত্র ৩০ মিনিটে একটি অনেক বড় অ্যাসাইনমেন্ট শেষ"
            " করেছে।"
        ),
    ),
    (
        "Surprise",
        "বহু বছর পর আপনি হঠাৎ একজন পুরোনো বন্ধুর সঙ্গে দেখা করলেন।",
    ),
    (
        "Disappointment",
        (
            "একটি গুরুত্বপূর্ণ অনুষ্ঠানে আপনি কারও সহযোগিতা আশা করেছিলেন, কিন্তু সে"
            " আসেনি।"
        ),
    ),
    (
        "Sympathy",
        "আপনার বন্ধু জানাল যে সে একটি গুরুত্বপূর্ণ পরীক্ষায় ফেল করেছে।",
    ),
    (
        "Frustration",
        (
            "আপনি একই সমস্যা কয়েকবার বুঝিয়েছেন, কিন্তু অন্য ব্যক্তি এখনও"
            " বিষয়টি বুঝতে পারছে না।"
        ),
    ),
    (
        "Teacher–Student",
        (
            "আপনি একটি অ্যাসাইনমেন্ট দেরিতে জমা দিয়েছেন এবং শিক্ষক জিজ্ঞেস করলেন"
            " কেন দেরি হয়েছে।"
        ),
    ),
    (
        "Workplace",
        (
            "আপনার সুপারভাইজার আপনাকে আরও একটি কাজ দিলেন, অথচ আপনার হাতে ইতিমধ্যে"
            " কয়েকটি জরুরি কাজ আছে।"
        ),
    ),
    (
        "Family",
        (
            "আপনার পরিবার আপনাকে একটি অনুষ্ঠানে যেতে বলছে, কিন্তু আপনার"
            " গুরুত্বপূর্ণ একাডেমিক কাজ আছে।"
        ),
    ),
    (
        "Stranger",
        (
            "ভিড়ের মধ্যে একজন অপরিচিত ব্যক্তি আপনাকে ধাক্কা দিলেন এবং ক্ষমা"
            " চাইলেন না।"
        ),
    ),
    (
        "Keeping a Secret",
        (
            "আপনার বন্ধু আপনাকে একটি ব্যক্তিগত বিষয় বলেছে এবং কাউকে না বলতে"
            " অনুরোধ করেছে। পরে অন্য একজন বন্ধু বিষয়টি সম্পর্কে জানতে চাইল।"
        ),
    ),
]


# ------------------------------------------------------------------------------
# Helper Functions
# ------------------------------------------------------------------------------
def safe_text(x):
    return str(x).replace("\x00", "").strip()


def init_excel():
    if not EXCEL_FILE.exists():
        wb = Workbook()
        ws = wb.active
        ws.title = "Responses"
        headers = [
            "Response_ID",
            "Timestamp",
            "Participant_Age_Group",
            "Education",
            "Bangla_Usage",
            "Bangla_Variety",
            "Scenario_ID",
            "Scenario_Category",
            "Context",
            "Utterance",
            "Intended_Meaning",
            "Primary_Intent_Category",
        ]
        ws.append(headers)
        wb.save(EXCEL_FILE)


def save_response(
    participant,
    scenario_id,
    category,
    context,
    utterance,
    intended,
    intent_cat="Uncategorized",
):
    init_excel()
    wb = load_workbook(EXCEL_FILE)
    ws = wb["Responses"]
    ws.append([
        str(uuid.uuid4()),
        datetime.now().isoformat(timespec="seconds"),
        participant["age"],
        participant["education"],
        participant["usage"],
        participant["variety"],
        scenario_id,
        category,
        context,
        safe_text(utterance),
        safe_text(intended),
        safe_text(intent_cat),
    ])
    wb.save(EXCEL_FILE)


init_excel()

# ------------------------------------------------------------------------------
# Session State Initialization
# ------------------------------------------------------------------------------
if "started" not in st.session_state:
    st.session_state.started = False
if "participant" not in st.session_state:
    st.session_state.participant = None
if "index" not in st.session_state:
    st.session_state.index = 0
if "answers" not in st.session_state:
    st.session_state.answers = {}

# ------------------------------------------------------------------------------
# Main UI Logic
# ------------------------------------------------------------------------------
st.title("🗣️ Bangla Pragmatics Data Collection")
st.caption("Context + Natural Utterance + Intended Meaning")

if not st.session_state.started:
    st.markdown("""
    ### Research Information
    You will read everyday Bangla situations and provide:
    1. **What you would naturally say** in that scenario.
    2. **What you intended to mean** by saying that.

    There are no right or wrong answers. Please use natural everyday Bangla (Standard, Regional, or Casual).
    Do **not** enter your name, phone number, email, address, or other identifying information.
    """)

    with st.form("participant_form"):
        age = st.selectbox("Age group", [
            "Under 18",
            "18–24",
            "25–34",
            "35–44",
            "45 or above",
            "Prefer not to say",
        ])
        education = st.selectbox("Highest education", [
            "Secondary",
            "Higher Secondary",
            "Undergraduate",
            "Master's",
            "PhD",
            "Other",
            "Prefer not to say",
        ])
        usage = st.selectbox(
            "How frequently do you use Bangla in everyday communication?",
            [
                "Almost always",
                "Very frequently",
                "Frequently",
                "Sometimes",
                "Rarely",
            ],
        )
        variety = st.selectbox("Bangla variety primarily used", [
            "Standard Bangla",
            "Regional/Dialectal Bangla",
            "Mixture of Standard and Regional Bangla",
            "Other / Prefer not to say",
        ])
        consent = st.checkbox(
            "I voluntarily agree to participate and allow my anonymous"
            " responses to be used for academic research."
        )
        submitted = st.form_submit_button("Start Questionnaire")

        if submitted:
            if not consent:
                st.error("Please provide consent before starting.")
            else:
                st.session_state.participant = {
                    "age": age,
                    "education": education,
                    "usage": usage,
                    "variety": variety,
                }
                st.session_state.started = True
                st.rerun()

else:
    i = st.session_state.index

    if i < len(SCENARIOS):
        category, context = SCENARIOS[i]

        st.progress((i + 1) / len(SCENARIOS))
        st.subheader(f"Scenario {i + 1} of {len(SCENARIOS)}")
        st.info(f"**পরিস্থিতি (Situation):** {context}")

        # Standard Reference Example Box
        with st.expander("💡 কীভাবে উত্তর দেবেন? একটি উদাহরণ দেখুন (How to answer - Example)"):
            st.markdown("""
            **নমুনা পরিস্থিতি:** আপনি বন্ধুর সাথে ঘরে বসে আছেন, বাইরে খুব ঠান্ডা এবং জানালা খোলা।
            * **১. আপনি যা বলবেন (Utterance):** *"আজকে আবহাওয়াটা একটু বেশিই ঠান্ডা, না?"*
            * **২. আপনার আসল উদ্দেশ্য (Intended Meaning):** *"পরোক্ষভাবে বন্ধুকে জানালাটা বন্ধ করতে বলা।"*
            """)

        with st.form(f"scenario_form_{i}"):
            utterance = st.text_area(
                "১. এই পরিস্থিতিতে আপনি স্বাভাবিকভাবে কী বলবেন? (What would you naturally say?)",
                value=st.session_state.answers.get((i, "u"), ""),
                height=110,
                placeholder=(
                    "উদাহরণ: আপনি বাস্তবে যেভাবে বলতেন সেভাবে লিখুন (আঞ্চলিক বা"
                    " প্রমিত বাংলায়)..."
                ),
            )

            intended = st.text_area(
                "২. এই কথা বলে আপনি মূল কী বোঝাতে চেয়েছেন? (What was your intended meaning?)",
                value=st.session_state.answers.get((i, "m"), ""),
                height=110,
                placeholder=(
                    "উদাহরণ: আপনার কথার মাধ্যমে আসল কী উদ্দেশ্য, অনুভূতি বা অর্থ"
                    " প্রকাশ করতে চেয়েছেন তা ব্যাখ্যা করুন..."
                ),
            )

            # Optional high-level categorization helper
            intent_cat = st.selectbox(
                (
                    "৩. (ঐচ্ছিক) আপনার উদ্দেশ্যটি মূলত কোন ধরণের? (Optional:"
                    " Select primary communicative function)"
                ),
                [
                    "Not Specified",
                    "অনুরোধ (Request)",
                    "অস্বীকৃতি/না বলা (Refusal)",
                    "ব্যঙ্গ/ঠাট্টা (Sarcasm/Irony)",
                    "অভিযোগ (Complaint)",
                    "প্রশংসা/ধন্যবাদ (Praise/Appreciation)",
                    "অন্যান্য (Other)",
                ],
                index=st.session_state.answers.get((i, "cat_idx"), 0),
            )

            c1, c2 = st.columns(2)
            back = c1.form_submit_button("← Previous")
            next_btn = c2.form_submit_button("Save & Next →")

            if back:
                st.session_state.answers[(i, "u")] = utterance
                st.session_state.answers[(i, "m")] = intended
                st.session_state.answers[(i, "cat_idx")] = [
                    "Not Specified",
                    "অনুরোধ (Request)",
                    "অস্বীকৃতি/না বলা (Refusal)",
                    "ব্যঙ্গ/ঠাট্টা (Sarcasm/Irony)",
                    "অভিযোগ (Complaint)",
                    "প্রশংসা/ধন্যবাদ (Praise/Appreciation)",
                    "অন্যান্য (Other)",
                ].index(intent_cat)
                st.session_state.index = max(0, i - 1)
                st.rerun()

            if next_btn:
                if not utterance.strip() or not intended.strip():
                    st.error(
                        "অনুগ্রহ করে দুটি প্রশ্নেরই উত্তর দিন। (Please answer both text questions.)"
                    )
                else:
                    st.session_state.answers[(i, "u")] = utterance
                    st.session_state.answers[(i, "m")] = intended
                    st.session_state.answers[(i, "cat_idx")] = [
                        "Not Specified",
                        "অনুরোধ (Request)",
                        "অস্বীকৃতি/না বলা (Refusal)",
                        "ব্যঙ্গ/ঠাট্টা (Sarcasm/Irony)",
                        "অভিযোগ (Complaint)",
                        "প্রশংসা/ধন্যবাদ (Praise/Appreciation)",
                        "অন্যান্য (Other)",
                    ].index(intent_cat)

                    if not st.session_state.answers.get((i, "saved"), False):
                        save_response(
                            st.session_state.participant,
                            i + 1,
                            category,
                            context,
                            utterance,
                            intended,
                            intent_cat,
                        )
                        st.session_state.answers[(i, "saved")] = True

                    st.session_state.index = i + 1
                    st.rerun()

    else:
        st.success("Thank you! All your responses have been recorded.")
        st.balloons()
        st.markdown("### আপনার অংশগ্রহণের জন্য অসংখ্য ধন্যবাদ।")

        # if EXCEL_FILE.exists():
        #     st.download_button(
        #         "⬇️ Download Current Excel Dataset",
        #         data=EXCEL_FILE.read_bytes(),
        #         file_name="bangla_pragmatics_responses.xlsx",
        #         mime=(
        #             "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        #         ),
        #     )

        if st.button("Start a New Response"):
            st.session_state.started = False
            st.session_state.participant = None
            st.session_state.index = 0
            st.session_state.answers = {}
            st.rerun()