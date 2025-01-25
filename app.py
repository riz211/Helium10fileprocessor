import streamlit as st
import pandas as pd
from io import BytesIO
import re
from openpyxl.styles import PatternFill
import os
from modules.blocked_items import BlockedItemsManager
from modules.data_processing import process_dataframes, extract_weight_with_packs
from modules.excel_utils import read_excel_file, calculate_shipping_cost, create_excel_export
from modules.tutorial import TutorialGuide
import time

# Configure the page
st.set_page_config(
    page_title="Helium 10 File Processor",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Initialize paths
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
ASSETS_DIR = os.path.join(BASE_DIR, "attached_assets")

# Create data directory if it doesn't exist
os.makedirs(DATA_DIR, exist_ok=True)

# Get file paths
blocked_brands_path = os.path.join(DATA_DIR, "Blocked_Brands.xlsx")
blocked_product_ids_path = os.path.join(DATA_DIR, "Blocked_Product_IDs.xlsx")
shipping_legend_path = os.path.join(ASSETS_DIR, "default_shipping_legend.xlsx")

# Initialize TutorialGuide
tutorial_guide = TutorialGuide()

# Custom CSS with added animations
st.markdown("""
    <style>
        .main > div {
            padding: 2rem 3rem;
        }
        .stTitle {
            font-size: 3rem !important;
            font-weight: 700 !important;
            color: #A682FF !important;
            margin-bottom: 2rem !important;
            text-align: center;
        }
        .stMarkdown {
            font-size: 1.1rem;
        }
        hr {
            margin: 2rem 0;
            border: none;
            border-top: 1px solid #E6E6FA;
        }
        .sidebar .stButton button {
            width: 100%;
            border-radius: 4px;
            background-color: #A682FF;
        }
        .sidebar .stMarkdown h1, .sidebar .stMarkdown h2, .sidebar .stMarkdown h3 {
            color: #A682FF;
        }
        .stDownloadButton button {
            background-color: #A682FF !important;
            color: white !important;
            padding: 0.5rem 1rem !important;
            border-radius: 4px !important;
        }
        .metrics-container {
            background-color: #F5F3FF;
            padding: 1.5rem;
            border-radius: 8px;
            margin: 1rem 0;
            animation: fadeIn 0.5s ease-in-out;
        }
        .stAlert.st-ae {
            background-color: #F5F3FF !important;
            border: 1px solid #A682FF !important;
            color: #262730 !important;
        }

        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }

        @keyframes slideIn {
            from { transform: translateX(-20px); opacity: 0; }
            to { transform: translateX(0); opacity: 1; }
        }

        .animate-fade-in {
            animation: fadeIn 0.5s ease-in-out;
        }

        .stDataFrame {
            animation: slideIn 0.6s ease-in-out;
        }

        .stProgress > div > div {
            background-color: #A682FF !important;
        }
    </style>
""", unsafe_allow_html=True)

# App title with elegant styling
st.title("Helium 10 File Processor")
st.markdown("<hr>", unsafe_allow_html=True)

# Add tutorial toggle button at the top of the sidebar
tutorial_guide.toggle_tutorial()

# Initialize BlockedItemsManager
blocked_items_manager = BlockedItemsManager(blocked_brands_path, blocked_product_ids_path)

# Load shipping legend
try:
    shipping_legend = pd.read_excel(shipping_legend_path)
    if not all(col in shipping_legend.columns for col in ["Weight Range Min (lb)", "Weight Range Max (lb)", "SHIPPING COST"]):
        st.error("Shipping legend file is missing required columns")
        shipping_legend = None
except Exception as e:
    st.error(f"Error loading shipping legend: {e}")
    shipping_legend = None

# Sidebar for Blocked Items Management
st.sidebar.header("Manage Blocked Items")

# Render tutorial if active
tutorial_guide.render_tutorial()

# Tab selection for Brands vs Product IDs
block_type = st.sidebar.radio("Select Block Type:", ["Brands", "Product IDs"])

if block_type == "Brands":
    with st.sidebar.form("Add Blocked Brands"):
        new_brand = st.text_input("Enter the brand to block")
        submit_button = st.form_submit_button("Add Brand")

        if submit_button:
            success, message = blocked_items_manager.add_brand(new_brand)
            if success:
                st.sidebar.success(message)
            else:
                st.sidebar.warning(message)

    # Bulk Upload for Blocked Brands
    st.sidebar.subheader("Bulk Upload Blocked Brands")
    bulk_brands_file = st.sidebar.file_uploader(
        "Upload Excel file with Blocked Brands",
        type=["xlsx"],
        key="bulk_brands"
    )

    if bulk_brands_file:
        try:
            bulk_brands = pd.read_excel(bulk_brands_file)
            success, message = blocked_items_manager.bulk_upload_brands(bulk_brands)
            if success:
                st.sidebar.success(message)
            else:
                st.sidebar.error(message)
        except Exception as e:
            st.sidebar.error(f"Error processing bulk upload: {e}")

else:  # Product IDs section
    with st.sidebar.form("Add Blocked Product IDs"):
        new_product_id = st.text_input("Enter the Product ID to block")
        block_reason = st.text_area("Reason for blocking (optional)", height=100)
        submit_button = st.form_submit_button("Add Product ID")

        if submit_button:
            success, message = blocked_items_manager.add_product_id(new_product_id, block_reason)
            if success:
                st.sidebar.success(message)
            else:
                st.sidebar.warning(message)

    # Bulk Upload for Blocked Product IDs
    st.sidebar.subheader("Bulk Upload Blocked Product IDs")
    bulk_products_file = st.sidebar.file_uploader(
        "Upload Excel file with Blocked Product IDs",
        type=["xlsx"],
        key="bulk_products"
    )

    if bulk_products_file:
        try:
            bulk_products = pd.read_excel(bulk_products_file)
            success, message = blocked_items_manager.bulk_upload_product_ids(bulk_products)
            if success:
                st.sidebar.success(message)
            else:
                st.sidebar.error(message)
        except Exception as e:
            st.sidebar.error(f"Error processing bulk upload: {e}")

# Display current blocked items based on selection
try:
    if block_type == "Brands":
        blocked_items_df = blocked_items_manager.get_blocked_brands()
        st.sidebar.subheader("Current Blocked Brands")
    else:
        blocked_items_df = blocked_items_manager.get_blocked_product_ids()
        st.sidebar.subheader("Current Blocked Product IDs")

    # Format the S.No column to remove commas and ensure it's displayed as a plain integer
    if "S.No" in blocked_items_df.columns:
        blocked_items_df["S.No"] = blocked_items_df["S.No"].astype(int)

    # Display the appropriate DataFrame
    st.sidebar.dataframe(blocked_items_df.style.format({"S.No": "{:.0f}"}), hide_index=True)

except Exception as e:
    st.sidebar.error(f"Error loading blocked items: {e}")


# Export section for blocked brands
st.sidebar.markdown("---")  # Add a visual separator
st.sidebar.subheader("Export Blocked Items")

if st.sidebar.button("Export Blocked Items List"):
    try:
        buffer, error, file_name = blocked_items_manager.export_blocked_items(block_type)
        if error:
            st.sidebar.error(error)
        else:
            st.sidebar.download_button(
                label="📥 Download Blocked Items",
                data=buffer.getvalue(),
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Click to download the current list of blocked items"
            )
            st.sidebar.success("Export file is ready for download!")
    except Exception as e:
        st.sidebar.error(f"Error preparing export: {str(e)}")

# Main content area
st.header("Upload Excel Files")
uploaded_files = st.file_uploader("Upload one or more Excel files", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    all_data = []

    with st.spinner('🔄 Initializing file processing...'):
        time.sleep(0.5)  # Small delay for visual feedback

    # Add a progress bar
    progress_bar = st.progress(0)
    status_text = st.empty()

    # Process each uploaded file with progress indication
    total_files = len(uploaded_files)
    for idx, uploaded_file in enumerate(uploaded_files):
        try:
            # Update progress
            progress = (idx + 1) / total_files
            progress_bar.progress(progress)
            status_text.text(f"Processing file {idx + 1} of {total_files}: {uploaded_file.name}")

            with st.spinner(f'📊 Processing {uploaded_file.name}...'):
                df, error = read_excel_file(uploaded_file)
                if error:
                    st.error(f"Error reading file {uploaded_file.name}: {error}")
                else:
                    all_data.append(df)
                    time.sleep(0.3)  # Small delay for visual feedback

        except Exception as e:
            st.error(f"Error processing file {uploaded_file.name}: {e}")

    # Clear progress bar and status text after completion
    progress_bar.empty()
    status_text.empty()

    if all_data:
        try:
            with st.spinner('🔍 Analyzing and combining data...'):
                # First combine all data
                combined_raw = pd.concat(all_data, ignore_index=True)

                # Process the combined data with the data_processing module
                processed_df, _ = process_dataframes([combined_raw], None, shipping_legend)

                # Then filter out blocked items
                combined_df, brands_removed, products_removed = blocked_items_manager.filter_data(processed_df)
                time.sleep(0.5)  # Small delay for visual feedback

            # Calculate and display metrics with animation
            with st.spinner('📊 Calculating metrics...'):
                total_input_listings = len(pd.concat(all_data, ignore_index=True))
                total_output_listings = len(combined_df)
                total_duplicates_removed = total_input_listings - total_output_listings - (brands_removed + products_removed)
                listings_no_weights = combined_df["ITEM WEIGHT (pounds)"].isnull().sum()
                low_price_items = len(combined_df[combined_df["RETAIL PRICE"] < 10])
                time.sleep(0.3)  # Small delay for visual feedback

            # Display metrics with enhanced styling
            st.markdown("### Metrics Summary")
            st.markdown('<div class="metrics-container">', unsafe_allow_html=True)
            st.markdown(f"""
            - **Total Listings in Input Files:** {total_input_listings}
            - **Total Listings in Output File:** {total_output_listings}
            - **Total Duplicates Removed:** {total_duplicates_removed}
            - **Blocked Brand Items Removed:** {brands_removed}
            - **Blocked Product IDs Removed:** {products_removed}
            - **Listings with No Weights:** {listings_no_weights}
            - **Review Orange Highlighted Rows:** {low_price_items}
            """, unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

            # Display processed data
            st.write("### Processed Data Preview")
            st.dataframe(combined_df, use_container_width=True)

            # Export functionality
            if st.button("Export to Excel"):
                try:
                    # Create formatted export with shipping legend
                    buffer = create_excel_export(combined_df, shipping_legend)

                    st.download_button(
                        label="Download Processed File",
                        data=buffer.getvalue(),
                        file_name="sellerchamp_batch_file.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success("File is ready for download!")
                    st.warning("⚠️ Important: After editing the missing weight items, don't forget to save the file as CSV!")
                except Exception as e:
                    st.error(f"Error creating export: {e}")

        except Exception as e:
            st.error(f"Error processing data: {e}")
else:
    st.markdown('<div class="metrics-container">Please upload one or more Excel files to begin processing.</div>', unsafe_allow_html=True)