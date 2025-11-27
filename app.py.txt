import streamlit as st
import pandas as pd

st.set_page_config(page_title="Access Comparison Tool", layout="wide")

st.title("🔐 Access Comparison Tool")

st.markdown("Upload Standard Access and Client Access Excel files to compare them.")

# Upload files
standard_file = st.file_uploader("Upload Standard Access Excel", type=["xlsx"])
client_file = st.file_uploader("Upload Client Access Excel", type=["xlsx"])

# Load rejected suggestions
def load_rejected():
    try:
        return pd.read_csv("rejected.csv")
    except:
        return pd.DataFrame(columns=["SecurityGroup", "Domain", "Suggestion"])

rejected_df = load_rejected()

if standard_file and client_file:
    # Read Excel sheets
    standard_df = pd.read_excel(standard_file)
    client_df = pd.read_excel(client_file)

    st.subheader("📄 Uploaded Data")
    st.write("Standard Access")
    st.dataframe(standard_df)

    st.write("Client Access")
    st.dataframe(client_df)

    # Merge for comparison
    merged = pd.merge(
        standard_df,
        client_df,
        on=["SecurityGroup", "Domain"],
        how="left",
        suffixes=("_Standard", "_Client")
    )

    # Identify mismatches
    merged["Suggestion"] = merged.apply(
        lambda row: f"Recommend Client Access = {row['AccessLevel_Standard']}"
        if row["AccessLevel_Standard"] != row["AccessLevel_Client"]
        else "",
        axis=1
    )

    # Remove suggestions that were rejected previously
    merged = merged[~merged["Suggestion"].isin(rejected_df["Suggestion"])]

    st.subheader("🔎 Suggestions")
    suggestions = merged[merged["Suggestion"] != ""]

    if suggestions.empty:
        st.success("No mismatches found!")
    else:
        for i, row in suggestions.iterrows():
            st.write(f"### {row['SecurityGroup']} — {row['Domain']}")
            st.write(f"Current Client Access: {row['AccessLevel_Client']}")
            st.write(f"Standard Access: **{row['AccessLevel_Standard']}**")
            st.write(f"Suggestion: **{row['Suggestion']}**")

            reject = st.button(f"Reject Suggestion {i}")

            if reject:
                # Add to rejected list
                new_row = {
                    "SecurityGroup": row["SecurityGroup"],
                    "Domain": row["Domain"],
                    "Suggestion": row["Suggestion"]
                }
                rejected_df = pd.concat([rejected_df, pd.DataFrame([new_row])])
                rejected_df.to_csv("rejected.csv", index=False)

                st.warning("Suggestion rejected and archived.")
                st.experimental_rerun()

    st.subheader("📁 Rejected Suggestions")
    st.dataframe(rejected_df)
