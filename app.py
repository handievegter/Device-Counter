import streamlit as st
import pandas as pd
import io
import re
from typing import Optional, Dict

# Global dictionary to store unknown device mappings
unknown_device_mappings: Dict[str, str] = {}

# --- Classification Logic ---

def classify_device_type(device: str) -> Optional[str]:
    """
    Classify device into one of the known categories:
    - BAC-I, I-CAB, I-CAB H, I-CAB M, BEAME
    Based on presence of keywords in the device name.
    """
    if not isinstance(device, str):
        return None

    # Normalize
    norm = device.strip().lower().replace('-', '').replace('_', '').replace(' ', '')

    # BEAME
    if 'beame' in norm or 'blame' in norm:
        return 'BEAME'

    # BAC-I
    if 'baci' in norm or 'bai03' in norm:
        return 'BAC-I'

    # I-CAB M
    if re.search(r'\bicabm\b', norm) or 'icabm' in norm:
        return 'I-CAB M'

    # I-CAB H
    if re.search(r'\bicabh\b', norm) or 'icabh' in norm:
        return 'I-CAB H'

    # COMBO without BACI → treat as I-CAB
    if 'combo' in norm and 'baci' not in norm:
        return 'I-CAB'

    # General I-CAB
    if 'icab' in norm:
        return 'I-CAB'

    return None

def classify_device_type_with_overrides(device: str) -> Optional[str]:
    """
    Classify device using user overrides if present, else fall back to default classification.
    """
    if not isinstance(device, str):
        return None

    key = device.strip()
    if key in unknown_device_mappings:
        return unknown_device_mappings[key]
    return classify_device_type(device)

# --- Sheet Processing ---

def process_flexible_sheet(df: pd.DataFrame, sheet_name: str, classify_func=classify_device_type) -> pd.DataFrame:
    """
    Process any sheet with customer/device/qty columns.
    Uses classify_func() to count devices and adds them to customer rows.
    """

    df = df.copy()

    # Detect and optionally use first row as headers
    first_row = df.iloc[0].astype(str).str.lower().str.contains("customer|device|qty").any()
    if first_row:
        df.columns = [str(c).strip() for c in df.iloc[0]]
        df = df[1:].reset_index(drop=True)
    else:
        df.columns = [str(c).strip() for c in df.columns]

    # Drop duplicate or unnamed columns
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.loc[:, ~df.columns.str.lower().str.contains('^nan$|^unnamed')]

    # Find expected columns
    col_map: Dict[str, Optional[str]] = {
        'customer': None,
        'device': None,
        'qty': None
    }

    for col in df.columns:
        label = col.strip().lower()
        if 'customer' in label:
            col_map['customer'] = col
        elif 'device' in label:
            col_map['device'] = col
        elif 'qty' in label:
            col_map['qty'] = col

    if not all(col_map.values()):
        return df

    # Forward-fill customer names
    df[col_map['customer']] = df[col_map['customer']].ffill()

    categories = ['BAC-I', 'I-CAB', 'I-CAB H', 'I-CAB M', 'BEAME']
    for cat in categories:
        df[cat] = ''

    grouped = df.groupby(col_map['customer'])

    for customer, group in grouped:
        counts = {cat: 0 for cat in categories}

        for _, row in group.iterrows():
            device = row.get(col_map['device'], '')
            qty = row.get(col_map['qty'], 0)

            try:
                qty = int(qty)
            except:
                qty = 0

            category = classify_func(device)
            if category:
                counts[category] += qty

        # Write totals into first row of group
        first_row = group.index[0]
        for col, val in counts.items():
            df.at[first_row, col] = val

    # Clean up: blank Customer Code except in first row
    for customer, group in grouped:
        indexes = group.index.tolist()
        if len(indexes) > 1:
            for idx in indexes[1:]:
                df.at[idx, col_map['customer']] = ''

    # Add NEW QTY = sum of the 5 summary columns (only on first row)
    df['NEW QTY'] = ''
    for customer, group in grouped:
        first_row = group.index[0]
        total = 0
        for col in categories:
            try:
                val = df.at[first_row, col]
                total += int(val) if val not in [None, ''] else 0
            except:
                pass
        df.at[first_row, 'NEW QTY'] = total

    # Reorder columns: original 3 + NEW QTY + summary (I-CAB before BAC-I)
    final_order = [
        col_map['customer'],
        col_map['device'],
        col_map['qty'],
        'NEW QTY',
        'I-CAB',
        'BAC-I',
        'I-CAB H',
        'I-CAB M',
        'BEAME'
    ]
    df = df[final_order]

    return df


def process_sheet_if_applicable(df: pd.DataFrame, sheet_name: str, classify_func=classify_device_type) -> pd.DataFrame:
    """
    Process all sheets except 'De/Re/Maintenance'
    """
    if sheet_name.strip().lower() == "de/re/maintenance":
        return df
    return process_flexible_sheet(df, sheet_name, classify_func=classify_func)

def style_customer_rows(df: pd.DataFrame, customer_col: str):
    """
    Apply bold + underline style to rows where customer code is present.
    Works only for visual display in Streamlit.
    """
    def highlight_row(row):
        is_customer = bool(row[customer_col]) and not pd.isna(row[customer_col])
        style = 'font-weight: bold; text-decoration: underline;' if is_customer else ''
        return [style] * len(row)

    return df.style.apply(highlight_row, axis=1)

# --- Streamlit UI ---

st.set_page_config(page_title="Device Counter", layout="wide")
st.title("📊 Device Counter App")


uploaded_file = st.file_uploader("📤 Upload an Excel file", type=["xlsx"])

if uploaded_file:
    st.success("File uploaded successfully.")
    xls = pd.ExcelFile(uploaded_file)

    processed_sheets = {}

    # First pass: process sheets with default classification
    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        processed_df = process_sheet_if_applicable(df, str(sheet))
        processed_sheets[sheet] = processed_df

    # Collect unknown device types from all sheets
    unknown_devices_set = set()
    for sheet, df in processed_sheets.items():
        # Identify device column
        device_col = None
        # Try to find device column by checking columns
        for col in df.columns:
            if 'device' in col.lower():
                device_col = col
                break
        if not device_col:
            continue
        # Check each device in the sheet
        for device in df[device_col].dropna().unique():
            if device and classify_device_type(device) is None:
                unknown_devices_set.add(device.strip())

    if unknown_devices_set:
        st.warning("⚠️ Unknown device types detected. Please classify them below:")

        with st.form("device_classification_form"):
            for device in sorted(unknown_devices_set):
                if device not in unknown_device_mappings:
                    unknown_device_mappings[device] = "Select category"
            for device in sorted(unknown_devices_set):
                options = ["Select category", "BAC-I", "I-CAB", "I-CAB H", "I-CAB M", "BEAME"]
                choice = st.selectbox(f"Device: {device}", options, index=options.index(unknown_device_mappings.get(device, "Select category")), key=f"device_{device}")
                unknown_device_mappings[device] = choice

            submitted = st.form_submit_button("Submit classifications")

        if submitted:
            # Remove entries with 'Select category'
            unknown_device_mappings = {k: v for k, v in unknown_device_mappings.items() if v != 'Select category'}

            # Re-process sheets with overrides
            processed_sheets = {}
            for sheet in xls.sheet_names:
                df = xls.parse(sheet)
                processed_df = process_sheet_if_applicable(df, str(sheet), classify_func=classify_device_type_with_overrides)
                processed_sheets[sheet] = processed_df

    output = io.BytesIO()
    # Save to Excel in memory
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in processed_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)

            # Apply bold + underline to customer rows
            worksheet = writer.sheets[sheet_name]
            if "Customer Code" in df.columns:
                customer_col_index = df.columns.get_loc("Customer Code")  # 0-based index
                for row_idx, val in enumerate(df["Customer Code"], start=2):  # Excel rows start at 1, plus header
                    if pd.notna(val) and str(val).strip() != "":
                        cell = worksheet.cell(row=row_idx, column=customer_col_index + 1)
                        cell.font = cell.font.copy(bold=True, underline="single")
    output.seek(0)

    # Download button
    st.download_button(
        label="📥 Download Processed Excel",
        data=output,
        file_name="processed_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
