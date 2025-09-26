import pandas as pd
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Demand–Employee Matching", layout="wide")
st.title("📊 Demand → Employee Matching Tool")

uploaded_file = st.file_uploader("Upload Demand_to_employee_summary.xlsx", type=["xlsx"])
if uploaded_file:
    st.write("Processing...")

    def process_file(uploaded_file):
        df_raw = pd.read_excel(uploaded_file)
        rows = []
        for _, r in df_raw.iterrows():
            demand = r["Request profile number"]
            associates = str(r["Associates"]).split(",")
            for assoc in associates:
                assoc = assoc.strip()
                if not assoc: continue
                emp_id = assoc.split("(")[0].strip()
                pct = float(assoc.split("(")[1].replace("%)", "")) if "(" in assoc else 0.0
                rows.append([demand, int(emp_id), pct, r["Associates"]])
        df = pd.DataFrame(rows, columns=["Demand","Employee","Match_%","All_Employees"])

        assignments, used = {}, set()

        # --- Step 1: Unique Fix ---
        while True:
            counts = df.groupby("Demand")["Employee"].nunique()
            uniques = counts[counts == 1].index
            if uniques.empty: break
            unique_rows = df[df["Demand"].isin(uniques)].drop_duplicates("Demand")
            for _, row in unique_rows.iterrows():
                d, e = row["Demand"], int(row["Employee"])
                if d not in assignments:
                    assignments[d] = (e, row["Match_%"], None, row["All_Employees"], "Unique Fix")
                    used.add(e)
            df = df[~df["Employee"].isin(used)]

        # --- Step 2: Scoring ---
        if not df.empty:
            emp_counts = df.groupby("Employee")["Demand"].nunique().to_dict()
            df["num_factor"] = df["Employee"].map(emp_counts).fillna(1.0)
            w1, w2 = 0.7, 0.3
            df["Score"] = w1 * df["Match_%"] + w2 * (1.0 / df["num_factor"])
            df = df.sort_values("Score", ascending=False)

            # Unique employees first
            for _, row in df.iterrows():
                d, e = row["Demand"], int(row["Employee"])
                if d not in assignments and e not in used:
                    assignments[d] = (e, row["Match_%"], row["Score"], row["All_Employees"], "Score Unique")
                    used.add(e)

            # Remaining demands reuse employees (*)
            for _, row in df.iterrows():
                d, e = row["Demand"], int(row["Employee"])
                if d not in assignments:
                    assignments[d] = (str(e) + "*", row["Match_%"], row["Score"], row["All_Employees"], "Score Reuse")

        # --- Step 3: Final Output ---
        all_demands = df_raw["Request profile number"].unique()
        out_rows = []
        for d in all_demands:
            if d in assignments:
                e, match, score, all_emps, method = assignments[d]
                out_rows.append([d, e, match, score, all_emps, method])
            else:
                out_rows.append([d, None, None, None, "-", "Unassigned"])
        return pd.DataFrame(out_rows, columns=["Demand", "Assigned_Employee", "Match_%", "Score", "All_Employees", "Method"])

    final = process_file(uploaded_file)
    st.success("✅ Matching complete!")
    st.dataframe(final.head(20))

    # --- Download Excel ---
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final.to_excel(writer, index=False, sheet_name="Assignments")
    st.download_button(
        "📥 Download Results (Excel)",
        data=output.getvalue(),
        file_name="Starvation_Safe_Assignments.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
