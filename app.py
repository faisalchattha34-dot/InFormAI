import streamlit as st

# ----------------------------
# PAGE CONFIG
# ----------------------------
st.set_page_config(
    page_title="Formlify – SaaS Dashboard",
    page_icon="📄",
    layout="wide"
)

# ----------------------------
# SESSION STATE (Mock Auth)
# ----------------------------
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# ----------------------------
# CSS (SaaS Look)
# ----------------------------
st.markdown("""
<style>
body {
    background-color: #f6f7fb;
}
.sidebar-title {
    font-size:22px;
    font-weight:600;
    margin-bottom:20px;
}
.card {
    background:white;
    padding:20px;
    border-radius:12px;
    box-shadow:0 4px 12px rgba(0,0,0,0.05);
}
.metric {
    font-size:28px;
    font-weight:700;
}
.label {
    color:#6b7280;
}
</style>
""", unsafe_allow_html=True)

# ----------------------------
# LOGIN PAGE
# ----------------------------
def login_page():
    st.markdown("## 🔐 Login to Formlify")
    col1, col2, col3 = st.columns([1,2,1])

    with col2:
        email = st.text_input("Email")
        password = st.text_input("Password", type="password")
        if st.button("Login", use_container_width=True):
            st.session_state.logged_in = True
            st.experimental_rerun()

# ----------------------------
# SIDEBAR
# ----------------------------
def sidebar():
    with st.sidebar:
        st.markdown("<div class='sidebar-title'>📄 Formlify</div>", unsafe_allow_html=True)
        page = st.radio(
            "Navigation",
            ["Dashboard", "Forms", "Members", "Responses", "Billing", "Settings"]
        )
        st.markdown("---")
        if st.button("🚪 Logout"):
            st.session_state.logged_in = False
            st.experimental_rerun()
    return page

# ----------------------------
# DASHBOARD PAGE
# ----------------------------
def dashboard_page():
    st.markdown("## 📊 Dashboard")

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown("<div class='card'><div class='metric'>3</div><div class='label'>Total Forms</div></div>", unsafe_allow_html=True)
    with c2:
        st.markdown("<div class='card'><div class='metric'>124</div><div class='label'>Responses</div></div>", unsafe_allow_html=True)
    with c3:
        st.markdown("<div class='card'><div class='metric'>89%</div><div class='label'>Completion Rate</div></div>", unsafe_allow_html=True)
    with c4:
        st.markdown("<div class='card'><div class='metric'>Pro</div><div class='label'>Current Plan</div></div>", unsafe_allow_html=True)

    st.markdown("### 📈 Activity Overview")
    st.info("Charts & analytics will appear here")

# ----------------------------
# FORMS PAGE
# ----------------------------
def forms_page():
    st.markdown("## 📝 Forms")
    st.button("➕ Create New Form")
    st.table({
        "Form Name": ["HR Hiring", "Student Admission"],
        "Status": ["Active", "Draft"],
        "Responses": [45, 0]
    })

# ----------------------------
# MEMBERS PAGE
# ----------------------------
def members_page():
    st.markdown("## 👥 Members")
    st.button("➕ Invite Member")
    st.table({
        "Email": ["user1@email.com", "user2@email.com"],
        "Status": ["Submitted", "Pending"]
    })

# ----------------------------
# RESPONSES PAGE
# ----------------------------
def responses_page():
    st.markdown("## 📥 Responses")
    st.info("Editable response table will be here")
    st.button("📤 Export Excel")

# ----------------------------
# BILLING PAGE
# ----------------------------
def billing_page():
    st.markdown("## 💳 Billing")
    st.markdown("""
    **Current Plan:** Pro  
    **Price:** $15 / month
    """)
    st.button("Upgrade Plan")

# ----------------------------
# SETTINGS PAGE
# ----------------------------
def settings_page():
    st.markdown("## ⚙️ Settings")
    st.text_input("Organization Name")
    st.button("Save Settings")

# ----------------------------
# MAIN ROUTER
# ----------------------------
if not st.session_state.logged_in:
    login_page()
else:
    page = sidebar()

    if page == "Dashboard":
        dashboard_page()
    elif page == "Forms":
        forms_page()
    elif page == "Members":
        members_page()
    elif page == "Responses":
        responses_page()
    elif page == "Billing":
        billing_page()
    elif page == "Settings":
        settings_page()
