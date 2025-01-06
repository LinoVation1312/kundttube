import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from io import BytesIO
import os

# Streamlit application configuration
st.set_page_config(
    page_title="Acoustic Analysis",
    page_icon=":chart_with_upwards_trend:",  # Icon chosen for the application
    layout="centered",
    initial_sidebar_state="expanded"
)

# Title and sidebar configuration
st.title("Interactive Acoustic Analysis Tool")
st.sidebar.title("Parameter Configuration")

# Feature to upload two Excel files
uploaded_file_1 = st.sidebar.file_uploader("Upload the first Excel file", type=["xlsx"])
uploaded_file_2 = st.sidebar.file_uploader("Upload the second Excel file", type=["xlsx"])

def load_data_from_excel(file):
    """
    Load data from an Excel file.
    """
    # Load the Excel file
    df = pd.read_excel(file, sheet_name=0, header=0)  # Read the first sheet (with titles in the first row)
    
    # Extract frequencies (column A)
    frequencies = df.iloc[:, 0].dropna().values  # Frequencies in the first column, ignoring empty values
    
    # Extract absorption data (all other columns)
    absorption_data = df.iloc[:, 1:].dropna(axis=0, how="all").values  # Remove rows where all values are NaN
    
    # Define thicknesses and densities (example, adapt according to your file)
    thicknesses = np.array([10, 20, 30])  # Thicknesses: 10, 20, 30 mm
    densities = np.array([75, 110, 150])  # Densities: 75, 110, 150 kg/m³
    
    return frequencies, thicknesses, densities, absorption_data

# Check if files have been uploaded
if uploaded_file_1 is not None and uploaded_file_2 is not None:
    # Load data from the two Excel files
    frequencies_1, thicknesses_1, densities_1, absorption_data_1 = load_data_from_excel(uploaded_file_1)
    frequencies_2, thicknesses_2, densities_2, absorption_data_2 = load_data_from_excel(uploaded_file_2)
    
    # Extract file names without the '.xlsx' extension
    file_name_1 = os.path.splitext(uploaded_file_1.name)[0]
    file_name_2 = os.path.splitext(uploaded_file_2.name)[0]
else:
    # Use default data if one or both files are not uploaded
    file_name_1 = "File_1"
    file_name_2 = "File_2"
    frequencies_1 = frequencies_2 = np.array([100, 500, 1000, 2000])
    thicknesses_1 = thicknesses_2 = np.array([10, 20, 30])
    densities_1 = densities_2 = np.array([75, 110, 150])
    absorption_data_1 = absorption_data_2 = np.array([
        [0.2, 0.4, 0.6, 0.8],
        [0.25, 0.45, 0.65, 0.85],
        [0.3, 0.5, 0.7, 0.9]
    ])

# Custom parameters via the interface
thickness_selected = st.sidebar.selectbox(
    "Choose thickness (mm)",
    options=[10, 20, 30],
    index=0
)

density_selected = st.sidebar.selectbox(
    "Choose density (kg/m³)",
    options=[75, 110, 150],
    index=0
)

# Initialize index variables only if files are uploaded
if uploaded_file_1 is not None:
    thickness_index_1 = np.where(thicknesses_1 == thickness_selected)[0][0]
    density_index_1 = np.where(densities_1 == density_selected)[0][0]

if uploaded_file_2 is not None:
    thickness_index_2 = np.where(thicknesses_2 == thickness_selected)[0][0]
    density_index_2 = np.where(densities_2 == density_selected)[0][0]

fig = None  # Initialize fig to None to prevent errors

# Extract absorption data for the selected frequency, thickness, and density
if uploaded_file_1 and uploaded_file_2:
    absorption_curve_1 = absorption_data_1[:, thickness_index_1 * len(densities_1) + density_index_1]
    absorption_curve_2 = absorption_data_2[:, thickness_index_2 * len(densities_2) + density_index_2]

    # Try plotting the absorption curves
    try:
        fig, ax = plt.subplots(figsize=(10, 8))

        # Change the background color of the graph
        fig.patch.set_facecolor('#6f6f7f')  # Dark gray background
        ax.set_facecolor('#3f3f4f')  # Dark gray axis background
        ax.tick_params(axis='both', colors='white')  # White tick color

        # Plot the curves
        ax.plot(frequencies_1, absorption_curve_1, label=file_name_1, color="b", marker="o", markersize=6)
        ax.plot(frequencies_2, absorption_curve_2, label=file_name_2, color="r", marker="x", markersize=6)

        # Add a title and labels
        ax.set_title(f"Absorption Curves for Thickness {thickness_selected} mm and Density {density_selected} kg/m³", color='white')
        ax.set_xlabel("Frequency (Hz)", color='white')
        ax.set_ylabel("Acoustic Absorption", color='white')
        ax.legend()

        # Enable grid
        ax.grid(True, linestyle="--", color='white', alpha=0.6)

        # Display the graph in Streamlit
        st.pyplot(fig)

    except ValueError as e:
        # Handle the error without displaying it intrusively
        st.markdown(
            f'<p style="position: fixed; bottom: 10px; right: 10px; font-size: 12px; color: red;">Dimension error: {str(e)}</p>',
            unsafe_allow_html=True
        )
else:
    st.warning("Please upload your Excel files to view the graph.")

# Function to save the graph as a PDF
def save_as_pdf(fig):
    """
    Save the current graph as a PDF and return it as a downloadable file.
    """
    pdf_buffer = BytesIO()
    if fig:
        fig.savefig(pdf_buffer, format="pdf")
        pdf_buffer.seek(0)
    return pdf_buffer

# Add a download button only if the figure exists
if fig:
    st.download_button(
        label="Download Comparison as PDF",
        data=save_as_pdf(fig),
        file_name="acoustic_comparison.pdf",
        mime="application/pdf"
    )
