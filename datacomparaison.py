import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from io import BytesIO
import os

# Streamlit application configuration
st.set_page_config(
    page_title="Acoustic Analysis",
    page_icon=":chart_with_upwards_trend:",
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
        df = pd.read_excel(file, sheet_name=0, header=0)
        if df.shape[1] < 2:
            raise ValueError("The Excel file must have at least two columns: one for frequencies and one for absorption data.")
        frequencies = df.iloc[:, 0].dropna().values
        absorption_data = df.iloc[:, 1:].dropna(axis=0, how="all").values
        if frequencies.size == 0:
            raise ValueError("The frequency column is empty. Please ensure that the first column contains frequency values.")
        thicknesses = np.array([10, 20, 30])
        densities = np.array([75, 110, 150])
        return frequencies, thicknesses, densities, absorption_data
    except Exception as e:
        st.error(f"Error loading file: {e}")
        st.warning("Please ensure the file contains at least two columns: frequencies and absorption data.")
        return None, None, None, None

# Add spinner for loading process
with st.spinner("Loading data from Excel files..."):
    if uploaded_file_1 is not None and uploaded_file_2 is not None:
        frequencies_1, thicknesses_1, densities_1, absorption_data_1 = load_data_from_excel(uploaded_file_1)
        frequencies_2, thicknesses_2, densities_2, absorption_data_2 = load_data_from_excel(uploaded_file_2)
        file_name_1 = os.path.splitext(uploaded_file_1.name)[0]
        file_name_2 = os.path.splitext(uploaded_file_2.name)[0]
    else:
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

thickness_selected = st.sidebar.selectbox("Choose thickness (mm)", options=[10, 20, 30], index=0)
density_selected = st.sidebar.selectbox("Choose density (kg/m³)", options=[75, 110, 150], index=0)

if uploaded_file_1 is not None:
    thickness_index_1 = np.where(thicknesses_1 == thickness_selected)[0][0]
    density_index_1 = np.where(densities_1 == density_selected)[0][0]

if uploaded_file_2 is not None:
    thickness_index_2 = np.where(thicknesses_2 == thickness_selected)[0][0]
    density_index_2 = np.where(densities_2 == density_selected)[0][0]

fig = None
freq_ticks = [80, 100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000, 5000, 6300]

if uploaded_file_1 and uploaded_file_2:
    absorption_curve_1 = absorption_data_1[:, thickness_index_1 * len(densities_1) + density_index_1]
    absorption_curve_2 = absorption_data_2[:, thickness_index_2 * len(densities_2) + density_index_2]

    try:
        def calculate_area(frequencies, absorption_curve):
            return np.sum((absorption_curve[:-1] + absorption_curve[1:]) / 2 * np.diff(frequencies))

        area_1 = calculate_area(frequencies_1, absorption_curve_1)
        area_2 = calculate_area(frequencies_2, absorption_curve_2)

        def format_filename(filename):
            replacements = {
                "_": " ",
                "-": " ",
                "FG": "Fiber Glass",
                "fg": "Fiber Glass",
                ".": " "
            }
            for old, new in replacements.items():
                filename = filename.replace(old, new)
            return " ".join(filename.split()).title()

        if area_1 > area_2:
            diff_percentage = ((area_1 - area_2) / area_2) * 100
            absorption_message = f"The {format_filename(file_name_1)} absorbs {diff_percentage:.2f}% more than the {format_filename(file_name_2)}."
        else:
            diff_percentage = ((area_2 - area_1) / area_1) * 100
            absorption_message = f"The {format_filename(file_name_2)} absorbs {diff_percentage:.2f}% more than the {format_filename(file_name_1)}."

        fig, ax = plt.subplots(figsize=(12, 8))
        fig.patch.set_facecolor('#616161')
        ax.set_facecolor('#dfdfdf')
        ax.tick_params(axis='both', colors='white', labelsize=12)

        ax.plot(frequencies_1, absorption_curve_1, label=format_filename(file_name_1), color="#1f77b4", linestyle='-', marker="x", markersize=9, linewidth=2.2)
        ax.plot(frequencies_2, absorption_curve_2, label=format_filename(file_name_2), color="#ff7f0e", linestyle='-', marker="x", markersize=9, linewidth=2.2)

        ax.set_xscale('log')
        ax.set_xticks(freq_ticks)
        ax.get_xaxis().set_major_formatter(plt.ScalarFormatter())
        ax.set_xticklabels(freq_ticks, rotation=25, ha="right", fontsize=12, color='white')

        ax.set_title(f"Absorption Curves for Thickness {thickness_selected} mm and Density {density_selected} kg/m³", 
                     color='white', fontsize=18, fontweight='bold', fontname="Arial", pad=20)
        ax.set_xlabel("Frequency (Hz)", color='white', fontsize=16, fontweight='bold', fontname="Arial")
        ax.set_ylabel("Alpha (Kundt)", color='white', fontsize=16, fontweight='bold', fontname="Arial")
        ax.legend(fontsize=14, loc='upper left', facecolor='black', framealpha=0.7, edgecolor='white', labelcolor='white')
        ax.grid(True, linestyle="--", color='black', alpha=0.8)
        ax.text(0.98, -0.12, "X-axis : semi-logarithmic scale", transform=ax.transAxes, fontsize=10, color="yellow", ha="right", va="center", style="italic")

        st.pyplot(fig)

        # Display the absorption message
        st.markdown(
            f'<p style="color: lightblue; font-size: 18px; text-align: center; font-weight: bold;">{absorption_message}</p>',
            unsafe_allow_html=True
        )

        # Display a warning if the selected thickness is 30 mm
        if thickness_selected == 30:
            st.markdown(
                '<p style="color: orange; font-size: 16px; text-align: center; font-style: italic;">'
                'Warning: Data may not be representative. This thickness is at the limit of the Kundt tube\'s operating range with this type of material.'
                '</p>',
                unsafe_allow_html=True
            )

                    # Display the absorption message
        st.markdown(
            f'<p style="color: lightblue; font-size: 18px; text-align: center; font-weight: bold;"GitHub Source : https://github.com/LinoVation1312/kundttube </p>',
            unsafe_allow_html=True
        )


    except ValueError as e:
        st.error(f"Dimension error: {e}")
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
