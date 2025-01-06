import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import streamlit as st
from io import BytesIO
import os
from scipy.interpolate import interp1d

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
        return frequencies, absorption_data
    except Exception as e:
        st.error(f"Error loading file: {e}")
        st.warning("Please ensure the file contains at least two columns: frequencies and absorption data.")
        return None, None

def align_frequencies(frequencies_ref, frequencies_to_align, absorption_to_align):
    """
    Align absorption data to the reference frequencies using linear interpolation.
    """
    interpolator = interp1d(frequencies_to_align, absorption_to_align, kind="linear", fill_value="extrapolate", axis=0)
    return interpolator(frequencies_ref)

# Load and process data
frequencies_1, absorption_data_1 = None, None
frequencies_2, absorption_data_2 = None, None

with st.spinner("Loading data from Excel files..."):
    if uploaded_file_1 is not None:
        frequencies_1, absorption_data_1 = load_data_from_excel(uploaded_file_1)
    if uploaded_file_2 is not None:
        frequencies_2, absorption_data_2 = load_data_from_excel(uploaded_file_2)

# Default values for testing (if no files are uploaded)
if frequencies_1 is None or frequencies_2 is None:
    st.warning("Using default test data since no files were uploaded.")
    frequencies_1 = frequencies_2 = np.array([100, 500, 1000, 2000])
    absorption_data_1 = absorption_data_2 = np.array([
        [0.2, 0.4, 0.6, 0.8],
        [0.25, 0.45, 0.65, 0.85],
        [0.3, 0.5, 0.7, 0.9]
    ])

# Sidebar configurations for thickness and density
thicknesses = np.array([10, 20, 30])
densities = np.array([75, 110, 150])

thickness_selected = st.sidebar.selectbox("Choose thickness (mm)", options=thicknesses, index=0)
density_selected = st.sidebar.selectbox("Choose density (kg/m³)", options=densities, index=0)

# Align frequencies and check dimensions
if frequencies_1 is not None and frequencies_2 is not None:
    if len(frequencies_1) != len(frequencies_2) or not np.array_equal(frequencies_1, frequencies_2):
        absorption_data_2 = align_frequencies(frequencies_1, frequencies_2, absorption_data_2)
        frequencies_2 = frequencies_1  # Align frequencies for consistency

    if absorption_data_1.shape[0] != absorption_data_2.shape[0]:
        st.error("Dimension error: The number of frequency points does not match between the two files.")
        st.stop()

# Select data for the chosen thickness and density
try:
    thickness_index = np.where(thicknesses == thickness_selected)[0][0]
    density_index = np.where(densities == density_selected)[0][0]

    absorption_curve_1 = absorption_data_1[:, thickness_index * len(densities) + density_index]
    absorption_curve_2 = absorption_data_2[:, thickness_index * len(densities) + density_index]
except Exception as e:
    st.error(f"Error selecting data for the chosen thickness and density: {e}")
    st.stop()

# Calculate the area under the curve
def calculate_area(frequencies, absorption_curve):
    return np.sum((absorption_curve[:-1] + absorption_curve[1:]) / 2 * np.diff(frequencies))

area_1 = calculate_area(frequencies_1, absorption_curve_1)
area_2 = calculate_area(frequencies_2, absorption_curve_2)

# Compare absorption areas
def format_filename(filename):
    replacements = {"_": " ", "-": " ", ".": " "}
    for old, new in replacements.items():
        filename = filename.replace(old, new)
    return " ".join(filename.split()).title()

file_name_1 = os.path.splitext(uploaded_file_1.name)[0] if uploaded_file_1 else "File 1"
file_name_2 = os.path.splitext(uploaded_file_2.name)[0] if uploaded_file_2 else "File 2"

if area_1 > area_2:
    diff_percentage = ((area_1 - area_2) / area_2) * 100
    absorption_message = f"The {format_filename(file_name_1)} absorbs {diff_percentage:.2f}% more than the {format_filename(file_name_2)}."
else:
    diff_percentage = ((area_2 - area_1) / area_1) * 100
    absorption_message = f"The {format_filename(file_name_2)} absorbs {diff_percentage:.2f}% more than the {format_filename(file_name_1)}."

# Plot absorption curves
fig, ax = plt.subplots(figsize=(12, 8))
ax.plot(frequencies_1, absorption_curve_1, label=format_filename(file_name_1), linestyle='-', marker='x')
ax.plot(frequencies_2, absorption_curve_2, label=format_filename(file_name_2), linestyle='-', marker='x')
ax.set_xscale('log')
ax.set_title(f"Absorption Curves for Thickness {thickness_selected} mm and Density {density_selected} kg/m³")
ax.set_xlabel("Frequency (Hz)")
ax.set_ylabel("Alpha (Kundt)")
ax.legend()
ax.grid(True)

st.pyplot(fig)

# Display absorption comparison
st.markdown(f"<p style='color: lightblue; font-size: 18px; text-align: center; font-weight: bold;'>{absorption_message}</p>", unsafe_allow_html=True)

# Save graph as PDF or JPEG
def save_as_pdf(fig):
    pdf_buffer = BytesIO()
    fig.savefig(pdf_buffer, format="pdf")
    pdf_buffer.seek(0)
    return pdf_buffer

def save_as_jpeg(fig):
    jpeg_buffer = BytesIO()
    fig.savefig(jpeg_buffer, format="jpeg", dpi=300)
    jpeg_buffer.seek(0)
    return jpeg_buffer

st.download_button("Download Comparison as PDF", data=save_as_pdf(fig), file_name="comparison.pdf", mime="application/pdf")
st.download_button("Download Comparison as JPEG", data=save_as_jpeg(fig), file_name="comparison.jpeg", mime="image/jpeg")

# Display GitHub link
st.markdown(
    '<p style="color: blue; font-size: 14px; text-align: center;">'
    'GitHub Link: <a href="https://github.com/LinoVation1312/kundttube" style="color: blue; text-decoration: none;" target="_blank">'
    'https://github.com/LinoVation1312/kundttube</a></p>',
    unsafe_allow_html=True
)
