import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io # Used to handle the uploaded Excel file in memory

# --- SET UP STREAMLIT PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Broker-Carrier Relationship Analyzer",
    layout="wide", # Use 'wide' layout for better visualization of the chart
    initial_sidebar_state="expanded" # Sidebar will be open by default
)

st.title("📊 Broker-Carrier Relationship Analyzer")
st.write("Upload your Excel file to visualize the relationships between brokers and carriers and explore detailed lists.")

# --- FILE UPLOADER ---
st.sidebar.header("Upload Data")
uploaded_file = st.sidebar.file_uploader(
    "Choose your 'Broker Carrier Listing.xlsx' file",
    type=["xlsx"] # Specify accepted file types
)

# Initialize variables to hold processed data and chart
# These will be populated only after a file is uploaded
df = None
broker_carrier_counts = {}
broker_carrier_lists = {}
brokers_sorted = []
counts_sorted = []
hover_texts_sorted = []

# --- Conditional execution based on file upload ---
if uploaded_file is not None:
    # Read the Excel file directly from the uploaded file object
    try:
        # Use io.BytesIO to ensure pandas can read the file-like object
        df = pd.read_excel(uploaded_file)
        st.sidebar.success("File uploaded and read successfully!")
    except Exception as e:
        st.sidebar.error(f"Error reading file: {e}. Please ensure it's a valid Excel (.xlsx) file.")
        st.stop() # Stop execution if file can't be read or is invalid

    # --- DATA PROCESSING LOGIC (from your original script) ---
    # Clean column names (strip whitespace)
    df.columns = df.columns.str.strip()

    # Dictionary to store broker-carrier relationships
    broker_carrier_counts = {}
    broker_carrier_lists = {}

    # Process each row to populate broker_carrier_counts and broker_carrier_lists
    for _, row in df.iterrows():
        broker = str(row['BROKERS']).strip() if pd.notna(row['BROKERS']) else None
        if not broker or broker.lower() in ['nan', 'none', '']:
            continue # Skip rows with invalid or empty broker names

        carriers_value = row['Brokers through']
        # Skip if carriers_value is NaN or contains specific noise strings
        if pd.isna(carriers_value) or str(carriers_value).strip().lower() in ['no data', 'n/a', 'aggregator', '']:
            continue

        # Split carriers by comma and clean whitespace for each
        carriers = [carrier.strip() for carrier in str(carriers_value).split(',') if carrier.strip()]

        # Store the count and list of carriers for each broker
        broker_carrier_counts[broker] = len(carriers)
        broker_carrier_lists[broker] = carriers

    # Prepare data for bar chart
    brokers = list(broker_carrier_counts.keys())
    counts = list(broker_carrier_counts.values())
    hover_texts = [f"Carriers: {', '.join(broker_carrier_lists[broker])}" for broker in brokers]

    # Sort brokers by count (descending)
    sorted_indices = sorted(range(len(counts)), key=lambda i: counts[i], reverse=True)
    brokers_sorted = [brokers[i] for i in sorted_indices]
    counts_sorted = [counts[i] for i in sorted_indices]
    hover_texts_sorted = [hover_texts[i] for i in sorted_indices]

    # --- PLOTLY CHART SECTION ---
    st.header("Broker Carrier Counts")
    st.write("This bar chart visualizes the number of carriers associated with each broker. Hover over a bar to see the list of carriers.")

    # Streamlit's selectbox for filtering the chart
    chart_selection = st.selectbox(
        "Select a broker to filter the chart view:",
        options=["All Brokers"] + brokers_sorted,
        index=0 # Default to 'All Brokers' when the app loads
    )

    # Prepare data for the selected chart view
    chart_brokers = []
    chart_counts = []
    chart_hover_texts = []

    if chart_selection == "All Brokers":
        chart_brokers = brokers_sorted
        chart_counts = counts_sorted
        chart_hover_texts = hover_texts_sorted
    else:
        # If a specific broker is selected, find their data
        try:
            idx = brokers_sorted.index(chart_selection)
            chart_brokers = [brokers_sorted[idx]]
            chart_counts = [counts_sorted[idx]]
            chart_hover_texts = [hover_texts_sorted[idx]]
        except ValueError:
            st.warning(f"Data for '{chart_selection}' not found after processing. Please try re-uploading the file.")


    # Create the Plotly bar chart
    fig = go.Figure()
    fig.add_trace(
        go.Bar(
            x=chart_brokers,
            y=chart_counts,
            text=chart_counts, # Display count on top of bars
            textposition='auto',
            hovertext=chart_hover_texts, # Detailed carriers on hover
            hoverinfo='text+y', # Show hovertext and y-value
            marker_color='lightblue'
        )
    )

    # Update layout for the bar chart
    fig.update_layout(
        title='Number of Carriers per Broker',
        xaxis_title='Broker',
        yaxis_title='Number of Carriers',
        xaxis={'tickangle': 45}, # Rotate x-axis labels for readability
        height=600,
        width=1200,
        margin=dict(b=200), # Adjust bottom margin for rotated labels
        showlegend=False # No legend needed for a single bar trace
    )
    
    # Display the Plotly chart in the Streamlit app
    st.plotly_chart(fig, use_container_width=True) # Makes chart responsive to container width

    # Add a horizontal divider for visual separation
    st.markdown("---")

    # --- BROKER-SPECIFIC CARRIER LIST SECTION ---
    st.header("Broker-Specific Carrier Details")
    st.write("Select a broker from the dropdown below to see a comprehensive list of all carriers associated with them.")

    # Streamlit's selectbox for the broker details display
    table_selection = st.selectbox(
        "Select a Broker to view its Associated Carriers:",
        options=["Select a Broker"] + brokers_sorted, # Add a default "Select" option
        index=0 # Default to "Select a Broker"
    )

    # Display carrier list based on selection
    if table_selection != "Select a Broker":
        st.subheader(f"Carriers for {table_selection}")
        if table_selection in broker_carrier_lists:
            carriers = broker_carrier_lists[table_selection]
            # Display carriers as a markdown list in Streamlit for clear readability
            for carrier in carriers:
                st.markdown(f"- **{carrier}**")
        else:
            st.warning(f"No associated carriers found for '{table_selection}'. This might be due to data cleaning or missing entries.")
    else:
        st.info("Please select a broker from the dropdown above to view their associated carriers.")

# --- Initial Message when no file is uploaded ---
else:
    st.info("Please upload your 'Broker Carrier Listing.xlsx' file in the sidebar to begin analysis. Use the sample file if you don't have your own data ready.")