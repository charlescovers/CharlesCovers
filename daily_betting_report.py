import pandas as pd
import os
import streamlit as st
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from datetime import datetime

# Ensure openpyxl is used for Excel writing
import openpyxl

# Attempt to import matplotlib, handle error gracefully
try:
    import matplotlib.pyplot as plt
    MATPLOTLIB_AVAILABLE = True
except ModuleNotFoundError:
    MATPLOTLIB_AVAILABLE = False
    st.warning("Matplotlib is not installed. Plots will not be available.")

# Set directory to save reports
save_path = "C:/Users/Charles Vance/Documents/BettingReports/"  # Updated path
os.makedirs(save_path, exist_ok=True)  # Ensure directory exists

# Streamlit UI
st.title("Daily Betting Model")
st.write("### Upload Raw Data (KenPom & Betting Odds)")

# File uploader for user to manually input data
offense_file = st.file_uploader("Upload Offense Data CSV", type=["csv"])
defense_file = st.file_uploader("Upload Defense Data CSV", type=["csv"])
odds_file = st.file_uploader("Upload Betting Odds CSV", type=["csv"])

if offense_file and defense_file and odds_file:
    # Load Data
    df_offense = pd.read_csv(offense_file)
    df_defense = pd.read_csv(defense_file)
    df_odds = pd.read_csv(odds_file)
    
    # Merge Data for Model Processing
    df_matchups = pd.merge(df_offense, df_defense, on="Team", how="inner")
    df_matchups = pd.merge(df_matchups, df_odds, on="Team", how="inner")
    
    # Compute Predictive Model (Basic version for now)
    df_matchups["Predicted_Spread"] = (df_matchups["AdjO"] - df_matchups["AdjD"]) / 2
    df_matchups["Implied_ML_Odds"] = df_matchups["Predicted_Spread"].apply(lambda x: -100 * (x / 2.5) if x > 0 else 100 * (2.5 / abs(x)))
    
    st.write("### Computed Matchups & Predictions")
    st.dataframe(df_matchups)
    
    # Function to generate daily betting report
    def generate_betting_report(df_matchups):
        today = datetime.today().strftime("%Y-%m-%d")
        filename_excel = os.path.join(save_path, f"Daily_Betting_Report_{today}.xlsx")
        filename_pdf = os.path.join(save_path, f"Daily_Betting_Report_{today}.pdf")
        
        # Save to Excel using openpyxl engine
        df_matchups.to_excel(filename_excel, index=False, engine='openpyxl')
        
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
            c.drawString(50, y_position, row["Team"])
            c.drawString(250, y_position, f"{row['Predicted_Spread']:.2f}")
            c.drawString(400, y_position, f"{row['Implied_ML_Odds']:.2f}")
            y_position -= 20
            if y_position < 50:  # Create a new page if needed
                c.showPage()
                c.setFont("Helvetica", 12)
                y_position = height - 50
        
        c.save()
        
        return filename_excel, filename_pdf
    
    if st.button("Generate Report"):
        excel_file, pdf_file = generate_betting_report(df_matchups)
        st.success("Report Generated Successfully!")
        
        with open(excel_file, "rb") as file:
            st.download_button(label="Download Excel Report", data=file, file_name=excel_file)
        
        with open(pdf_file, "rb") as file:
            st.download_button(label="Download PDF Report", data=file, file_name=pdf_file)
