import streamlit as st
import pandas as pd

# Load Excel data
st.set_page_config(page_title="Excel Dashboard", layout="wide")
st.title("📊 Customer Support Ticket Analysis (Excel)")

uploaded_file = "data/Customer_Support_Tickets_50K_Enhanced.xlsx"
df = pd.read_excel(uploaded_file)

# Show file preview
st.subheader("🧾 Data Preview")
st.dataframe(df.head())

# KPIs (example)
st.markdown("### 📈 Key Metrics")
col1, col2, col3 = st.columns(3)
col1.metric("Total Tickets", f"{len(df):,}")
col2.metric("Avg SLA Breach %", f"{df['SLA Breach %'].mean():.2f}%")
col3.metric("High Risk Count", df[df['Risk Score'] > 80].shape[0])

# Aggregated view
st.subheader("📊 Complaints by Category")
if 'Complaint Type' in df.columns:
    st.bar_chart(df['Complaint Type'].value_counts())

# Download
st.download_button(
    "📥 Download Original Excel",
    open(uploaded_file, 'rb'),
    file_name="Customer_Support_Tickets_50K_Enhanced.xlsx"
)
