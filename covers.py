import pandas as pd
import matplotlib.pyplot as plt
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from datetime import datetime
import os
import streamlit as st

# Ensure openpyxl is used for Excel writing
import openpyxl

# Set directory to save reports
save_path = "C:/Users/Charles Vance/Documents/BettingReports/"  # Updated path
os.makedirs(save_path, exist_ok=True)  # Ensure directory exists

# Sample Data for df_matchups (Ensure real data is loaded here)
df_matchups = pd.DataFrame({
    "Matchup": ["Ohio St. vs Illinois", "Memphis vs Rice", "Nebraska vs Oregon"],
    "Predicted_Spread": [2.68, 3.71, 0.93],
    "Implied_ML_Odds": [-1071, -1483, -107]
})

# Function to generate daily betting report
def generate_betting_report(df_matchups):
    today = datetime.today().strftime("%Y-%m-%d")
    filename_excel = os.path.join(save_path, f"Daily_Betting_Report_{today}.xlsx")
    filename_pdf = os.path.join(save_path, f"Daily_Betting_Report_{today}.pdf")
    
    # Save to Excel using openpyxl engine
    df_matchups.to_excel(filename_excel, index=False, engine='openpyxl')
    print(f"Saved Excel report: {filename_excel}")
    
    # Generate PDF Report
    c = canvas.Canvas(filename_pdf, pagesize=letter)
    width, height = letter
    
    c.setFont("Helvetica-Bold", 16)
    c.drawString(200, height - 50, f"Daily Betting Report - {today}")
    
    c.setFont("Helvetica-Bold", 12)
    y_position = height - 80
    c.drawString(50, y_position, "Matchup")
    c.drawString(250, y_position, "Predicted Spread")
    c.drawString(400, y_position, "Implied ML Odds")
    y_position -= 20
    
    c.setFont("Helvetica", 12)
    for _, row in df_matchups.iterrows():
        c.drawString(50, y_position, row["Matchup"])
        c.drawString(250, y_position, f"{row['Predicted_Spread']:.2f}")
        c.drawString(400, y_position, f"{row['Implied_ML_Odds']:.2f}")
        y_position -= 20
        if y_position < 50:  # Create a new page if needed
            c.showPage()
            c.setFont("Helvetica", 12)
            y_position = height - 50
    
    c.save()
    print(f"Saved PDF report: {filename_pdf}")
    
    return filename_excel, filename_pdf

# Streamlit UI
st.title("Daily Betting Report")
st.write("### Today's Betting Predictions")
st.dataframe(df_matchups)

if st.button("Generate Report"):
    excel_file, pdf_file = generate_betting_report(df_matchups)
    st.success("Report Generated Successfully!")
    
    with open(excel_file, "rb") as file:
        st.download_button(label="Download Excel Report", data=file, file_name=excel_file)
    
    with open(pdf_file, "rb") as file:
        st.download_button(label="Download PDF Report", data=file, file_name=pdf_file)
