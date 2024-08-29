import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

def create_lollipop_chart(dataframe):
    # Extract relevant columns and clean data
    df_clean = dataframe.iloc[1:, [0, 1, 3]].copy()
    df_clean.columns = ['Sector', 'Loan_Amount', 'Position']
    df_clean['Loan_Amount'] = pd.to_numeric(df_clean['Loan_Amount'], errors='coerce')
    df_clean['Position'] = pd.to_numeric(df_clean['Position'], errors='coerce')

    # Set the common limit based on the first sector's eligible limit
    common_limit = pd.to_numeric(dataframe.iloc[1, 2], errors='coerce')

    # Convert loan amounts to crores (1 crore = 10 million)
    df_clean['Loan_Amount_Crores'] = df_clean['Loan_Amount'] / 1e7

    # Sort data by Loan Amount to improve visualization
    df_clean = df_clean.sort_values('Loan_Amount_Crores', ascending=False)

    # Create a transposed lollipop chart with a larger figure size
    plt.figure(figsize=(14, 10))  # Increased figure size for better readability

    # Set background color
    plt.gca().set_facecolor('Lightblue')

    # Draw the lollipop chart
    plt.vlines(x=df_clean['Sector'], ymin=0, ymax=df_clean['Loan_Amount_Crores'], color='magenta')
    plt.plot(df_clean['Sector'], df_clean['Loan_Amount_Crores'], "o", color='orange')

    # Add a horizontal line for the common eligible limit in crores
    plt.axhline(y=common_limit / 1e7, color='white', linestyle='--', label=' Eligible Limit (Crores)')

    # Labeling and visual adjustments
    plt.ylabel('Loan Amount (Crores)')
    plt.xlabel('Sector')
    plt.title('Sectorial Loan Amount vs  Eligible Limit (in Crores)')
    plt.legend()
    plt.grid(axis='y', linestyle='', alpha=0.7)

    # Adding labels to the lollipop points (with horizontal orientation)
    for index, row in df_clean.iterrows():
        plt.text(row['Sector'], row['Loan_Amount_Crores'] + 0.5, f'{row["Loan_Amount_Crores"]:.2f}', 
                 va='bottom', ha='center', color='blue', rotation=0)  # Horizontal labels

    # Show the limit amount on the right end of the red line
    plt.text(df_clean['Sector'].iloc[-1], common_limit / 1e7, f'{common_limit / 1e7:.2f} Cr', 
             color='red', ha='left', va='bottom', fontsize=10, fontweight='bold')

    plt.xticks(rotation=90)

    # Adjust the y-axis ticks to show increments of 50 crores, with more space between them
    plt.yticks(range(0, int(df_clean['Loan_Amount_Crores'].max()) + 100, 100))

    plt.tight_layout()
    return plt

# Streamlit app
st.title("Excel to Lollipop Chart Converter")

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    # Load the Excel file
    df_compare = pd.read_excel(uploaded_file, sheet_name='Report')
    st.write("Data loaded successfully!")

    # Generate the chart
    plt = create_lollipop_chart(df_compare)
    
    # Display the chart
    st.pyplot(plt)

    # Provide a download button for the PNG
    buf = BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)
    st.download_button(label="Download chart as PNG", data=buf, file_name="lollipop_chart.png", mime="image/png")
