import streamlit as st
import pandas as pd
import io
import re

# --- Classification Logic ---

def classify_device_type(device: str) -> str:
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
    if 'baci' in norm:
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

# --- Sheet Processing ---

def process_flexible_sheet(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """
    Process any sheet with customer/device/qty columns.
    Uses classify_device_type() to count devices and adds them to customer rows.
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
    col_map = {
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
        st.warning(f"⚠️ Skipping sheet '{sheet_name}' — required columns not found.")
        return df

    # Forward-fill customer names
    df[col_map['customer']] = df[col_map['customer']].fillna(method='ffill')

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

            category = classify_device_type(device)
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
        try:
            total = sum(int(df.at[first_row, col] or 0) for col in categories)
            df.at[first_row, 'NEW QTY'] = total
        except:
            df.at[first_row, 'NEW QTY'] = ''

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


def process_sheet_if_applicable(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """
    Process all sheets except 'De/Re/Maintenance'
    """
    if sheet_name.strip().lower() == "de/re/maintenance":
        return df
    return process_flexible_sheet(df, sheet_name)

# --- Streamlit UI ---

st.set_page_config(page_title="Device Counter", layout="wide")
st.title("📊 Device Counter App")

uploaded_file = st.file_uploader("📤 Upload an Excel file", type=["xlsx"])

if uploaded_file:
    st.success("File uploaded successfully.")
    xls = pd.ExcelFile(uploaded_file)

    processed_sheets = {}

    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        processed_df = process_sheet_if_applicable(df, sheet)
        processed_sheets[sheet] = processed_df

        st.subheader(f"📄 Preview: {sheet}")
        st.dataframe(processed_df.head(10))

    # Save to Excel in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, df in processed_sheets.items():
            df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)

    # Download button
    st.download_button(
        label="📥 Download Processed Excel",
        data=output,
        file_name="processed_output.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
