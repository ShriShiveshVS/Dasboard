import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import matplotlib.pyplot as plt
import seaborn as sns
import calendar
from datetime import datetime
import io
from matplotlib.colors import ListedColormap
import os

# Streamlit layout
st.set_page_config(layout="wide")
st.title("📊 Worker Attendance Dashboard")

# Upload file
uploaded_file = st.file_uploader("📁 Upload Excel Attendance File", type=["xlsx"])

if uploaded_file:
    @st.cache_data
    def load_excel_file(uploaded_file):
        xl = pd.ExcelFile(uploaded_file)

        # Identify month sheets using keywords
        month_keywords = {
            "january": "January", "february": "February", "march": "March",
            "april": "April", "may": "May", "june": "June",
            "july": "July", "august": "August", "september": "September",
            "october": "October", "november": "November", "december": "December"
        }

        sheet_map = {}
        for sheet in xl.sheet_names:
            sheet_lower = sheet.lower()
            for key in month_keywords:
                if key in sheet_lower:
                    clean_name = month_keywords[key]
                    sheet_map[clean_name] = xl.parse(sheet, header=1)
                    break

        return sheet_map

    sheet_map = load_excel_file(uploaded_file)

    # Month selection
    all_months_ordered = ["January", "February", "March", "April", "May", "June",
                          "July", "August", "September", "October", "November", "December"]
    available_months = [m for m in all_months_ordered if m in sheet_map]
    months = ["Overall"] + available_months
    selected_month = st.selectbox("🗕️ Select Month", months)

    # Data preprocessing
    def preprocess(df):
        df.columns = [str(col).strip() for col in df.columns]
        day_cols = sorted([col for col in df.columns if col.isdigit()], key=lambda x: int(x))
        df[day_cols] = df[day_cols].fillna(0)
        df["Attendance %"] = df[day_cols].sum(axis=1) / len(day_cols) * 100
        df["Irregular"] = df["Attendance %"] < 60
        return df, day_cols

    # Load selected data
    if selected_month == "Overall":
        all_data = []
        for month in available_months:
            df_month, _ = preprocess(sheet_map[month])
            df_month["Month"] = month
            all_data.append(df_month)
        df = pd.concat(all_data, ignore_index=True)
    else:
        df, day_cols = preprocess(sheet_map[selected_month])
        df["Month"] = selected_month

    # Filters
    st.sidebar.header("🔍 Filters")
    skill_filter = st.sidebar.multiselect("Select Skill", df["Skill"].dropna().unique())
    skill_level_filter = st.sidebar.multiselect("Select Skill Level", df["Category"].dropna().unique())
    subcontractor_filter = st.sidebar.multiselect("Select Subcontractor", df["Subcontractor"].dropna().unique())
    worker_filter = st.sidebar.multiselect("Select Worker", df["Worker Name"].dropna().unique())

    filtered = df.copy()
    if skill_filter:
        filtered = filtered[filtered["Skill"].isin(skill_filter)]
    if skill_level_filter:
        filtered = filtered[filtered["Category"].isin(skill_level_filter)]
    if subcontractor_filter:
        filtered = filtered[filtered["Subcontractor"].isin(subcontractor_filter)]
    if worker_filter:
        filtered = filtered[filtered["Worker Name"].isin(worker_filter)]

    # Summary Metrics
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Workers", filtered.shape[0])
    col2.metric("Avg Attendance %", f"{filtered['Attendance %'].mean():.1f}%")
    col3.metric("Irregular Workers", filtered["Irregular"].sum())

    # Attendance % Chart
    st.subheader("📉 Attendance % by Subcontractor")
    fig = px.bar(
        filtered,
        x="Subcontractor",
        y="Attendance %",
        color="Irregular",
        color_discrete_map={True: "yellow", False: "green"},
    )
    fig.update_layout(xaxis_tickangle=-40)
    st.plotly_chart(fig, use_container_width=True)

    # Attendance Summary Table
    st.subheader("📋 Attendance Summary")
    if selected_month == "Overall":
        summary_day_cols = sorted({col for m in available_months for col in sheet_map[m].columns if str(col).isdigit()}, key=lambda x: int(x))
    else:
        summary_day_cols = sorted([col for col in sheet_map[selected_month].columns if str(col).isdigit()], key=lambda x: int(x))

    display_cols = [
        "Month", "WISA ID", "EIP ID", "Worker Name", "Skill", "Category", "Vendor Code", "Subcontractor",
        "Attendance %", "Irregular"
    ] + summary_day_cols

    display_cols = [col for col in display_cols if col in filtered.columns or col in summary_day_cols]

    st.dataframe(filtered[display_cols])

    # Export options
    csv = filtered[display_cols].to_csv(index=False).encode("utf-8")
    st.download_button("📅 Download as CSV", csv, file_name="attendance_summary.csv", mime="text/csv")

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
        filtered[display_cols].to_excel(writer, index=False, sheet_name='Summary')
    st.download_button("📅 Download as Excel", data=excel_buffer.getvalue(), file_name="attendance_summary.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Heatmap
    st.subheader("🗓️ Worker Calendar Heatmap")
    show_heatmap = st.checkbox("Show Worker Heatmap for All 12 Months")

    if show_heatmap:
        heatmap_df = df[df["Month"] == selected_month] if selected_month != "Overall" else df
        identifiers_df = heatmap_df[["Worker Name", "WISA ID", "EIP ID"]].drop_duplicates().astype(str)
        identifier = st.selectbox("Select Worker (Name/WISA/EIP)", identifiers_df.agg(" - ".join, axis=1))
        selected_name = identifier.split(" - ")[0]

        custom_cmap = ListedColormap(["#0a7500", "#4CAF50"])
        row_layout = [all_months_ordered[i:i+4] for i in range(0, 12, 4)]

        for row_months in row_layout:
            cols = st.columns(4)
            for i, month in enumerate(row_months):
                if month not in sheet_map:
                    with cols[i]:
                        st.warning(f"No data for {selected_name} in {month}")
                    continue

                df_month, _ = preprocess(sheet_map[month])
                worker_row = df_month[df_month["Worker Name"] == selected_name]
                if worker_row.empty:
                    with cols[i]:
                        st.warning(f"No data for {selected_name} in {month}")
                    continue

                worker_row = worker_row.iloc[0]
                month_number = all_months_ordered.index(month) + 1
                year = 2025
                _, num_days = calendar.monthrange(year, month_number)
                valid_cols = [str(day) for day in range(1, num_days + 1)]

                attendance = {
                    int(day): int(worker_row[day]) if pd.notna(worker_row[day]) else 0
                    for day in valid_cols if day in worker_row
                }

                cal = calendar.monthcalendar(year, month_number)
                heat_matrix, day_labels = [], []
                for week in cal:
                    row_data, labels = [], []
                    for day in week:
                        if day == 0:
                            row_data.append(np.nan)
                            labels.append("")
                        else:
                            row_data.append(float(attendance.get(day, 0)))
                            labels.append(str(day))
                    heat_matrix.append(row_data)
                    day_labels.append(labels)

                fig, ax = plt.subplots(figsize=(2.6, 2.6))
                sns.heatmap(
                    heat_matrix,
                    cmap=custom_cmap,
                    linewidths=0.4,
                    linecolor='white',
                    cbar=False,
                    square=True,
                    xticklabels=False,
                    yticklabels=False,
                    annot=day_labels,
                    fmt='',
                    annot_kws={"fontsize": 6, "color": "white"}
                )
                ax.set_title(f"{month}", fontsize=9, pad=3)
                plt.tight_layout(pad=0.1)
                with cols[i]:
                    st.pyplot(fig)
else:
    st.info("👆 Please upload an Excel file with monthly attendance sheets.")
