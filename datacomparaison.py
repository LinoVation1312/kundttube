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
    Load data from an Excel file and validate the expected structure.
    """
    try:
        # Load the Excel file
        df = pd.read_excel(file, sheet_name=0, header=0)  # Read the first sheet (with titles in the first row)

        # Validate that the file has at least two columns (frequencies and absorption data)
        if df.shape[1] < 2:
            raise ValueError("The Excel file must have at least two columns: one for frequencies and one for absorption data.")

        # Extract frequencies (column A)
        frequencies = df.iloc[:, 0].dropna().values  # Frequencies in the first column, ignoring empty values
        
        # Extract absorption data (all other columns)
        absorption_data = df.iloc[:, 1:].dropna(axis=0, how="all").values  # Remove rows where all values are NaN
        
        # Validate that there are no missing frequency values
        if frequencies.size == 0:
            raise ValueError("The frequency column is empty. Please ensure that the first column contains frequency values.")

        # Define thicknesses and densities (example, adapt according to your file)
        thicknesses = np.array([10, 20, 30])  # Thicknesses: 10, 20, 30 mm
        densities = np.array([75, 110, 150])  # Densities: 75, 110, 150 kg/m³

        return frequencies, thicknesses, densities, absorption_data

    except Exception as e:
        # Display a helpful error message to the user
        st.error(f"Error loading file: {e}")
        st.warning(
            "Please ensure the file contains at least two columns: the first for frequencies (numeric values) "
            "and the others for absorption data (numeric values). The file should be in .xlsx format."
        )
        return None, None, None, None

# Add spinner for loading process
with st.spinner("Loading data from Excel files..."):
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
        fig, ax = plt.subplots(figsize=(12, 8))

        # Change the background color of the graph
        fig.patch.set_facecolor('#616161')  # Darker gray background
        ax.set_facecolor('#dfdfdf')  # Slightly lighter dark gray axis background
        ax.tick_params(axis='both', colors='white', labelsize=12)  # White tick color with larger labels

        # Plot the curves with enhanced styling
        ax.plot(frequencies_1, absorption_curve_1, label=file_name_1, color="#1f77b4", linestyle='-', marker="x", markersize=9, linewidth=2.2)
        ax.plot(frequencies_2, absorption_curve_2, label=file_name_2, color="#ff7f0e", linestyle='-', marker="x", markersize=9, linewidth=2.2)

        # Add a title with custom font and styling
        ax.set_title(f"Absorption Curves for Thickness {thickness_selected} mm and Density {density_selected} kg/m³", 
                     color='white', fontsize=18, fontweight='bold', fontname="Arial", pad=20)
        
        # Add axis labels with custom styling
        ax.set_xlabel("Frequency (Hz)", color='white', fontsize=16, fontweight='bold', fontname="Arial")
        ax.set_ylabel("Acoustic Absorption", color='white', fontsize=16, fontweight='bold', fontname="Arial")

        # Add a legend with improved readability
        ax.legend(
            fontsize=14,
            loc='upper left',
            facecolor='black',  # Keep the background of the legend black
            framealpha=0.7,     # Add transparency to the frame
            edgecolor='white',  # White edge for the legend frame
            labelcolor='white'  # Set the legend text color to white
        )

        # Adjust the grid for better visual appearance
        ax.grid(True, linestyle="--", color='black', alpha=0.8)  # Subtle lighter grid lines

        # Add annotations (optional, but this can highlight key data points)
        max_y_1 = np.max(absorption_curve_1)
        max_x_1 = frequencies_1[np.argmax(absorption_curve_1)]
        ax.annotate(f'Max: {max_y_1:.2f}', xy=(max_x_1, max_y_1), xytext=(max_x_1 + 50, max_y_1 + 0.05),
                    arrowprops=dict(arrowstyle='->', color='white'), fontsize=12, color='white')

        max_y_2 = np.max(absorption_curve_2)
        max_x_2 = frequencies_2[np.argmax(absorption_curve_2)]
        ax.annotate(f'Max: {max_y_2:.2f}', xy=(max_x_2, max_y_2), xytext=(max_x_2 + 50, max_y_2 + 0.05),
                    arrowprops=dict(arrowstyle='->', color='white'), fontsize=12, color='white')

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

# Function to save the graph as a JPEG
def save_as_jpeg(fig):
    """
    Save the current graph as a JPEG and return it as a downloadable file.
    """
    jpeg_buffer = BytesIO()
    if fig:
        fig.savefig(jpeg_buffer, format="jpeg", dpi=300)
        jpeg_buffer.seek(0)
    return jpeg_buffer

# Add download buttons for PDF and JPEG
if fig:
    st.download_button(
        label="Download Comparison as PDF",
        data=save_as_pdf(fig),
        file_name="acoustic_comparison.pdf",
        mime="application/pdf"
    )
    st.download_button(
        label="Download Comparison as JPEG",
        data=save_as_jpeg(fig),
        file_name="acoustic_comparison.jpeg",
        mime="image/jpeg"
    )
