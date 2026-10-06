import streamlit as st
import pandas as pd
import re
from datetime import datetime, timedelta
import os
import subprocess
import json
import streamlit.components.v1 as components

st.set_page_config(page_title="Jayshakti Farsan Mart - Barcode Printer", layout="wide")
st.title("🖨️ Jayshakti Farsan Mart - Barcode Printer")

# --- Default Preset Data ---
default_presets = {
    "Default Vertical (2-up)": {
        "layout_type": "2-up",
        "right_x_offset": 300,
        "elements": [
            {"id": "ProductName", "type": "text", "x": 10, "y": 30, "font": "3"},
            {"id": "NetWeight", "type": "text", "x": 10, "y": 90, "font": "2"},
            {"id": "MRP", "type": "text", "x": 10, "y": 120, "font": "2"},
            {"id": "Batch", "type": "text", "x": 10, "y": 150, "font": "2"},
            {"id": "BestBeforeTitle", "type": "text", "x": 10, "y": 180, "font": "2"},
            {"id": "BestBeforeDate", "type": "text", "x": 10, "y": 210, "font": "2"},
            {"id": "Barcode", "type": "barcode", "x": 160, "y": 80, "height": 50, "thickness": 2, "rotation": 0},
            {"id": "ShopName1", "type": "text", "x": 35, "y": 280, "font": "2"},
            {"id": "ShopName2", "type": "text", "x": 35, "y": 305, "font": "2"},
            {"id": "Address", "type": "text", "x": 10, "y": 340, "font": "1"}
        ]
    }
}

if not os.path.exists('presets.json'):
    with open('presets.json', 'w') as f:
        json.dump(default_presets, f, indent=4)

with open('presets.json', 'r') as f:
    presets = json.load(f)

# --- Component Declaration ---
_editor_component = components.declare_component("my_editor", path="editor_component")

# --- App Logic ---
@st.cache_data
def load_data(uploaded_file):
    if uploaded_file is None: return pd.DataFrame()
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)
        header_idx = 0
        for i, row in df_raw.iterrows():
            if "Item Name" in row.values:
                header_idx = i
                break
        df = pd.read_excel(uploaded_file, header=header_idx)
        return df.dropna(subset=['Item Name'])
    except Exception:
        return pd.DataFrame()

def parse_weight(item_name):
    match = re.search(r'(\d+)\s*(g|kg|ml|l|gm|grams)\b', str(item_name), re.IGNORECASE)
    if match: return f"{match.group(1)}{match.group(2).lower()}"
    return ""

def get_printers():
    try:
        result = subprocess.run(['lpstat', '-p'], capture_output=True, text=True)
        return [line.split()[1] for line in result.stdout.splitlines() if line.startswith('printer')]
    except Exception: return []

# --- Editor Dialog ---
@st.dialog("Canva-Style Sticker Designer", width="large")
def open_editor(base_name, current_item, orient, layout_data, live_data):
    st.markdown(f"**Editing Layout for:** {current_item if current_item else 'Global Template'}")
    
    result = _editor_component(layout_data=layout_data, live_data=live_data, default=None)
    
    if result is not None:
        try:
            import json
            with open('presets.json', 'r') as f:
                current_presets = json.load(f)
            
            # Check if user clicked 'Save for All Items'
            save_mode = result.pop('save_mode', 'item')
            
            if save_mode == 'global' or not current_item:
                save_key = base_name
            else:
                save_key = f"Override_{orient}_{current_item}"
                
            current_presets[save_key] = result
            
            with open('presets.json', 'w') as f:
                json.dump(current_presets, f, indent=4)
                
            st.success(f"Layout saved successfully!")
            st.session_state.editor_open = False
            st.rerun()
        except Exception as e:
            st.error(f"Error saving: {e}")

    if st.button("❌ Close Without Saving"):
        st.session_state.editor_open = False
        st.rerun()

# --- Main UI ---
st.sidebar.header("📁 Data Source")
uploaded_file = st.sidebar.file_uploader("Upload Item_Details_Report.xlsx (Overwrites saved)", type=["xlsx", "xls"])

DATA_FILE = "saved_items_data.xlsx"

if uploaded_file is not None:
    with open(DATA_FILE, "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.sidebar.success("New file saved!")

df = pd.DataFrame()
if os.path.exists(DATA_FILE):
    mod_time = os.path.getmtime(DATA_FILE)
    dt_str = datetime.fromtimestamp(mod_time).strftime("%d %b %Y, %I:%M %p")
    st.sidebar.info(f"Using saved data\n(Updated: {dt_str})")
    df = load_data(DATA_FILE)

if df.empty:
    st.info("Using sample data for preview.")
    df = pd.DataFrame([
        {"Item Name": "Bhel 200G", "Sell Price": 60, "Barcode": 10083},
        {"Item Name": "Aaloo Sev 250G", "Sell Price": 90, "Barcode": 10310},
        {"Item Name": "Ratlami Sev 250G", "Sell Price": 90, "Barcode": 10038}
    ])

if "batch_queue" not in st.session_state:
    st.session_state.batch_queue = []

tab_designer, tab_batch = st.tabs(["🏷️ Label Editor", "🛒 Batch Queue"])

with tab_designer:
    col1, col2 = st.columns([1, 1], gap="large")

    with col1:
        st.subheader("1. Select Product")
        
        # Category Filter
        if 'Category' in df.columns:
            categories = sorted(list(set([str(c).strip() for c in df['Category'].dropna() if str(c).strip()])))
            selected_category = st.selectbox("Filter by Category (Optional)", ["All Categories"] + categories)
        else:
            selected_category = "All Categories"
            
        if selected_category != "All Categories":
            filtered_df = df[df['Category'].astype(str).str.strip() == selected_category]
        else:
            filtered_df = df
            
        item_names = filtered_df['Item Name'].dropna().astype(str).tolist()
        selected_item = st.selectbox("Search and Select Item", [""] + item_names)

        item_data = {"name": "", "price": "", "barcode": ""}
        if selected_item:
            row = df[df['Item Name'] == selected_item].iloc[0]
            item_data = {
                "name": str(row.get("Item Name", "")),
                "price": row.get("Sell Price", ""),
                "barcode": str(row.get("Barcode", ""))
            }

        st.subheader("2. Label Details")

        default_name = item_data.get("name", "")
        name_input = st.text_area("Product Name (Press Enter for new line)", value=default_name, height=80)
        weight_input = st.text_input("Net Weight", value=parse_weight(default_name))
        mrp_input = st.text_input("MRP", value=f"MRP: ₹ {item_data.get('price', '')}" if item_data.get("price") else "MRP: ₹ ")
        batch_input = st.text_input("Batch No", value=datetime.now().strftime("%m%d"))
        expiry_date_input = st.text_input("Expiry Date", value=(datetime.now() + timedelta(days=60)).strftime("%d-%b-%Y").upper())
        barcode_input = st.text_input("Barcode Data", value=item_data.get("barcode", "").replace(".0", "") if pd.notnull(item_data.get("barcode")) else "")

        st.subheader("3. Sticker Orientation")
        selected_orient = st.radio("Select Orientation", ["Landscape", "Portrait", "Mini 6-in-1"], index=0, horizontal=True)

        if selected_orient == "Landscape":
            base_preset_name = "Default Landscape"
        elif selected_orient == "Portrait":
            base_preset_name = "Default Vertical (2-up)"
        else:
            base_preset_name = "Default Mini 6-in-1"
            
        current_preset = presets.get(base_preset_name, {}).copy()
        selected_preset_name = base_preset_name

        is_override = False
        if selected_item:
            override_key = f"Override_{selected_orient}_{selected_item}"
            if override_key in presets:
                current_preset = presets[override_key].copy()
                selected_preset_name = override_key
                is_override = True
                st.success(f"Loaded custom saved layout for '{selected_item}'")

        current_preset["orientation"] = selected_orient.lower()
        current_preset["elements"] = [e for e in current_preset.get("elements", []) if e["id"] != "VegLogo"]

        live_data = {
            "ProductName": name_input,
            "NetWeight": f"Net wt: {weight_input}" if not weight_input.startswith("Net wt") else weight_input,
            "MRP": mrp_input,
            "Batch": f"Batch: {batch_input}" if not batch_input.startswith("Batch") else batch_input,
            "BestBeforeTitle": "Expiry Date:",
            "BestBeforeDate": expiry_date_input,
            "Barcode": barcode_input,
            "ShopName1": "Shree Jayshakti",
            "ShopName2": "Farsan Mart",
            "Address": "Ahmedabadi pole, Raopura."
        }
        
        for i in range(1, 7):
            live_data[f"ProductName_{i}"] = name_input
            live_data[f"NetWeight_{i}"] = f"Net wt: {weight_input}" if not weight_input.startswith("Net wt") else weight_input
            live_data[f"MRP_{i}"] = mrp_input
            live_data[f"Barcode_{i}"] = barcode_input

        if st.button("✏️ Edit Barcode Layout", type="primary"):
            st.session_state.editor_open = True

        if st.session_state.get("editor_open", False):
            open_editor(base_preset_name, selected_item, selected_orient, current_preset, live_data)

    with col2:
        st.subheader("Live Preview (Read-Only)")

        def generate_static_preview(layout, live):
            canvas_w = 400
            canvas_h = 600
            html = f'<div style="position: relative; width: {canvas_w}px; height: {canvas_h}px; border: 2px solid #555; background-color: #fff; color: #000; box-shadow: 2px 2px 10px rgba(0,0,0,0.1); overflow: hidden;">'

            y_offset = layout.get("right_x_offset", 300)

            for side in [0, 1]:
                html += f'<div style="position: absolute; left: 0; top: 0; right: 0; bottom: 0; transform: {"translateY(" + str(y_offset) + "px)" if side == 1 else "none"};">'

                for el in layout["elements"]:
                    txt = el.get("text") or live.get(el["id"], "")
                    if el["id"] == "MRP": txt = txt.replace('₹', 'Rs.')

                    if el["type"] == "text":
                        h_map = {"1": 12, "2": 20, "3": 24, "4": 32, "5": 48}
                        w_map = {"1": 8, "2": 12, "3": 14, "4": 24, "5": 32}
                        h = h_map.get(str(el.get("font", "2")), 20)
                        w = w_map.get(str(el.get("font", "2")), 12)
                        align = el.get("align", "left")

                        lines = str(txt).split("\n")
                        for i, line in enumerate(lines):
                            if line.strip():
                                line_width = len(line.strip()) * w
                                x_adj = 0
                                if align == "center": x_adj = -(line_width / 2)
                                elif align == "right": x_adj = -line_width

                                rot = f"transform: rotate({el.get('rotation',0)}deg); transform-origin: top left;" if el.get("rotation",0) != 0 else ""




                                px = int(el.get("padX", 0))
                                py = int(el.get("padY", 0))

                                box_left = el["x"] + x_adj
                                box_top = el["y"] + (i * (h + 4))

                                if el.get("invert"):
                                    html += f"""
                                    <div style="position: absolute; left: {box_left}px; top: {box_top}px; width: {line_width}px; height: {h}px; display: flex; align-items: center; justify-content: flex-start; overflow: visible;">
                                        <span style="font-family: monospace; font-size: {h}px; font-weight: bold; color: white; background: black; white-space: nowrap; letter-spacing: 0px; padding: {py}px {px}px; margin-left: -{px}px; margin-top: -{py}px;">{line.strip()}</span>
                                    </div>"""
                                else:
                                    html += f"""
                                    <div style="position: absolute; left: {box_left}px; top: {box_top}px; width: {line_width}px; height: {h}px; display: flex; align-items: center; justify-content: flex-start;">
                                        <span style="font-family: monospace; font-size: {h}px; font-weight: bold; color: black; white-space: nowrap; letter-spacing: 0px;">{line.strip()}</span>
                                    </div>""" 

                    elif el["type"] == "barcode":
                        rot = f"transform: rotate({el.get('rotation',0)}deg); transform-origin: top left;" if el.get("rotation",0) != 0 else ""
                        thick = el.get("thickness", 2)
                        bw = 90 if thick == 1 else 160
                        align = el.get("align", "left")
                        x_adj = 0
                        if align == "center": x_adj = -(bw / 2)
                        elif align == "right": x_adj = -bw

                        html += f'<div style="position: absolute; left: {el["x"] + x_adj}px; top: {el["y"]}px; {rot}"><svg class="barcode-preview" jsbarcode-value="{txt}" jsbarcode-width="{thick}" jsbarcode-height="{el.get("height",50)}" jsbarcode-displayvalue="true" jsbarcode-fontsize="16" jsbarcode-margin="0"></svg></div>'

                html += "</div>" # close the side container

            html += f'<div style="position: absolute; top: {layout.get("right_x_offset", 300)}px; left: 0; right: 0; height: 1px; border-top: 2px dashed #bbb;"></div>'
            html += "</div>"
            html += '<script src="https://cdn.jsdelivr.net/npm/jsbarcode@3.11.5/dist/JsBarcode.all.min.js"></script><script>JsBarcode(".barcode-preview").init(); setTimeout(() => { document.querySelectorAll(".tspl-static").forEach(span => { let parent = span.parentElement; let targetW = parseInt(parent.style.width); if (span.offsetWidth > 0) { let scale = targetW / span.offsetWidth; span.style.transform = `scaleX(${scale})`; } }); }, 50);</script>'
            return html

        st.components.v1.html(generate_static_preview(current_preset, live_data), height=650)

    st.subheader("4. Print Action")
    printers = get_printers()
    if not printers:
        st.warning("No local printers found. Are you running on the server?")

    selected_printer = st.selectbox("Select Printer", printers) if printers else "TSC_RAW_PRINTER"
    quantity = st.number_input("Number of Sheets (Yields 2 stickers per sheet)", min_value=1, value=1)

    col_btn1, col_btn2 = st.columns(2)
    with col_btn1:
        print_now = st.button("🖨️ Print Now", type="primary", use_container_width=True)
    with col_btn2:
        add_batch = st.button("➕ Add to Batch Printing", use_container_width=True)

    if add_batch:
        st.session_state.batch_queue.append({
            "product_name": live_data.get("ProductName", "Unknown Product"),
            "quantity": quantity,
            "preset": current_preset.copy(),
            "live_data": live_data.copy(),
            "printer": selected_printer
        })
        st.success(f"Added {live_data.get('ProductName', 'Item')} (Qty: {quantity}) to Batch!")

    if print_now:
        if selected_printer:
            try:
                tspl = "SIZE 75 mm, 50 mm\r\nGAP 2 mm, 0 mm\r\nDIRECTION 1\r\nCLS\r\n"
                right_off = int(current_preset.get("right_x_offset", 300))

                for side in [0, 1]:
                    for el in current_preset["elements"]:
                        txt = el.get("text") or live_data.get(el["id"], "")
                        if el["id"] == "MRP": txt = txt.replace('₹', 'Rs.')

                        if el["type"] == "text":
                            h_map = {"1": 12, "2": 20, "3": 24, "4": 32, "5": 48}
                            w_map = {"1": 8, "2": 12, "3": 14, "4": 24, "5": 32}
                            char_w = w_map.get(str(el.get("font", "2")), 12)
                            char_h = h_map.get(str(el.get("font", "2")), 20)
                            line_height = char_h + 4
                            align = el.get("align", "left")

                            lines = str(txt).split("\n")
                            for i, line in enumerate(lines):
                                if line.strip():
                                    line_width = len(line.strip()) * char_w
                                    x_adj = 0
                                    if align == "center": x_adj = int(-(line_width / 2))
                                    elif align == "right": x_adj = int(-line_width)

                                    # ALWAYS rotate 90 degrees for Landscape preset!
                                    x_base = 300 if side == 0 else 600
                                    phys_x = int(x_base - el["y"] - (i * line_height))
                                    phys_y = int(el["x"] + x_adj)
                                    phys_x = max(0, phys_x)
                                    phys_y = max(0, phys_y)
                                    tspl += f'TEXT {phys_x},{phys_y},"{el.get("font","2")}",90,1,1,"{line.strip()}"\r\n'

                                    if el.get("invert"):
                                        px = int(el.get("padX", 0))
                                        py = int(el.get("padY", 0))
                                        rev_x = int(phys_x - char_h - py)
                                        rev_y = int(phys_y - px)
                                        rev_w = int(char_h + (2 * py))
                                        rev_h = int(line_width + (2 * px))
                                        if rev_y < 0:
                                            rev_h += rev_y
                                            rev_y = 0
                                        if rev_x < 0:
                                            rev_w += rev_x
                                            rev_x = 0
                                        tspl += f'REVERSE {rev_x},{rev_y},{rev_w},{rev_h}\r\n'

                        elif el["type"] == "barcode":
                            thick = int(el.get("thickness", 2))
                            bw = 90 if thick == 1 else 160
                            align = el.get("align", "left")
                            x_adj = 0
                            if align == "center": x_adj = int(-(bw / 2))
                            elif align == "right": x_adj = int(-bw)

                            x_base = 300 if side == 0 else 600
                            phys_x = int(x_base - el["y"])
                            phys_y = int(el["x"] + x_adj)
                            phys_rot = (int(el.get("rotation",0)) + 90) % 360
                            tspl += f'BARCODE {phys_x},{phys_y},"128",{el.get("height",50)},1,{phys_rot},{thick},{thick},"{txt}"\r\n'

                mid_x = int(current_preset.get("right_x_offset", 300))
                if mid_x > 0:
                    for dash_y in range(0, 400, 16):
                        tspl += f"BAR {mid_x},{dash_y},2,8\r\n"

                tspl += f"PRINT 1, {quantity}\r\n"

                temp_file = "temp_print_job.prn"
                with open(temp_file, "wb") as f:
                    f.write(tspl.encode('utf-8'))
                process = subprocess.run(['lp', '-d', selected_printer, '-o', 'raw', temp_file], capture_output=True, text=True)

                if process.returncode == 0:
                    st.success("Successfully sent to printer!")
                else:
                    st.error(f"Printing failed: {process.stderr}")
            except Exception as e:
                st.error(f"Error printing: {e}")

with tab_batch:
    st.subheader("🛒 Batch Print Queue")
    if not st.session_state.batch_queue:
        st.info("Your batch queue is empty. Go to the Label Editor and add some items!")
    else:
        # Display list
        for idx, item in enumerate(st.session_state.batch_queue):
            bc1, bc2, bc3 = st.columns([3, 1, 1])
            with bc1: st.write(f"**{item['product_name']}**")
            with bc2: st.write(f"Qty: {item['quantity']} sheets")
            with bc3:
                if st.button("🗑️ Remove", key=f"rem_{idx}"):
                    st.session_state.batch_queue.pop(idx)
                    st.rerun()
        st.markdown("---")
        bcol1, bcol2 = st.columns([1, 1])
        with bcol1:
            if st.button("🗑️ Clear Batch", use_container_width=True):
                st.session_state.batch_queue = []
                st.rerun()
        with bcol2:
            if st.button("🖨️ Print All Batch", type="primary", use_container_width=True):
                if not printers:
                    st.error("No printers available!")
                else:
                    try:
                        # The user might have selected different printers for different items, but let's use the first item's printer or a default.
                        batch_printer = st.session_state.batch_queue[0]["printer"] if st.session_state.batch_queue[0].get("printer") else printers[0]
                        # Master setup
                        full_tspl = "SIZE 75 mm, 50 mm\r\nGAP 2 mm, 0 mm\r\nDIRECTION 1\r\n"
                        
                        for item in st.session_state.batch_queue:
                            full_tspl += "CLS\r\n"
                            item_preset = item["preset"]
                            item_live = item["live_data"]
                            for side in [0, 1]:
                                for el in item_preset["elements"]:
                                    txt = el.get("text") or item_live.get(el["id"], "")
                                    if el["id"] == "MRP": txt = txt.replace('₹', 'Rs.')
                                    
                                    if el["type"] == "text":
                                        h_map = {"1": 12, "2": 20, "3": 24, "4": 32, "5": 48}
                                        w_map = {"1": 8, "2": 12, "3": 14, "4": 24, "5": 32}
                                        char_w = w_map.get(str(el.get("font", "2")), 12)
                                        char_h = h_map.get(str(el.get("font", "2")), 20)
                                        line_height = char_h + 4
                                        align = el.get("align", "left")
                                        
                                        lines = str(txt).split("\n")
                                        for i, line in enumerate(lines):
                                            if line.strip():
                                                line_width = len(line.strip()) * char_w
                                                x_adj = 0
                                                if align == "center": x_adj = int(-(line_width / 2))
                                                elif align == "right": x_adj = int(-line_width)
                                                
                                                x_base = 300 if side == 0 else 600
                                                phys_x = int(x_base - el["y"] - (i * line_height))
                                                phys_y = int(el["x"] + x_adj)
                                                phys_x = max(0, phys_x)
                                                phys_y = max(0, phys_y)
                                                full_tspl += f'TEXT {phys_x},{phys_y},"{el.get("font","2")}",90,1,1,"{line.strip()}"\r\n'
                                                
                                                if el.get("invert"):
                                                    px = int(el.get("padX", 0))
                                                    py = int(el.get("padY", 0))
                                                    rev_x = int(phys_x - char_h - py)
                                                    rev_y = int(phys_y - px)
                                                    rev_w = int(char_h + (2 * py))
                                                    rev_h = int(line_width + (2 * px))
                                                    if rev_y < 0:
                                                        rev_h += rev_y
                                                        rev_y = 0
                                                    if rev_x < 0:
                                                        rev_w += rev_x
                                                        rev_x = 0
                                                    full_tspl += f'REVERSE {rev_x},{rev_y},{rev_w},{rev_h}\r\n'
                                    
                                    elif el["type"] == "barcode":
                                        thick = int(el.get("thickness", 2))
                                        bw = 90 if thick == 1 else 160
                                        align = el.get("align", "left")
                                        x_adj = 0
                                        if align == "center": x_adj = int(-(bw / 2))
                                        elif align == "right": x_adj = int(-bw)
                                        
                                        x_base = 300 if side == 0 else 600
                                        phys_x = int(x_base - el["y"])
                                        phys_y = int(el["x"] + x_adj)
                                        phys_rot = (int(el.get("rotation",0)) + 90) % 360
                                        full_tspl += f'BARCODE {phys_x},{phys_y},"128",{el.get("height",50)},1,{phys_rot},{thick},{thick},"{txt}"\r\n'
                            
                            mid_x = int(item_preset.get("right_x_offset", 300))
                            if mid_x > 0:
                                for dash_y in range(0, 400, 16):
                                    full_tspl += f"BAR {mid_x},{dash_y},2,8\r\n"
                            
                            full_tspl += f'PRINT 1, {item["quantity"]}\r\n'
                        
                        temp_file = "temp_batch_job.prn"
                        with open(temp_file, "wb") as f:
                            f.write(full_tspl.encode("utf-8"))
                        import subprocess
                        process = subprocess.run(["lp", "-d", batch_printer, "-o", "raw", temp_file], capture_output=True, text=True)
                        if process.returncode == 0:
                            st.success("Batch successfully sent to printer!")
                            st.session_state.batch_queue = []
                            st.rerun()
                        else:
                            st.error(f"Batch printing failed: {process.stderr}")
                    except Exception as e:
                        st.error(f"Error printing batch: {e}")

def generate_tspl(layout, copies):
    try:
        quantity = int(copies)
        
        for p in st.session_state.printers:
            if p["name"] == layout.get("printer"):
                selected_printer = p["name"]
                break
        else:
            if not st.session_state.printers:
                st.error("No printers found.")
                return
            selected_printer = st.session_state.printers[0]["name"]
        
        tspl = "SIZE 75 mm, 50 mm\r\nGAP 3 mm, 0 mm\r\nCLS\r\n"
        
        for side in [0, 1]:
            for el in layout.get("elements", []):
                txt = str(st.session_state.live_data.get(el["id"], el.get("text", el["id"])))
                
                if el["type"] == "text":
                    font_str = str(el.get("font", "2"))
                    h = h_map.get(font_str, 20)
                    w = w_map.get(font_str, 12)
                    align = el.get("align", "left")
                    
                    lines = txt.split("\n")
                    line_height = h + 4
                    char_h = h
                    
                    for i, line in enumerate(lines):
                        if not line.strip(): continue
                        
                        line_width = len(line.strip()) * w
                        x_adj = 0
                        if align == "center": x_adj = int(-(line_width / 2))
                        elif align == "right": x_adj = int(-line_width)
                        
                        x_base = 300 if side == 0 else 600
                        phys_x = int(x_base - el["y"] - (i * line_height))
                        
                        if el.get("invert"):
                            phys_y = int(el["x"] + x_adj)
                            tspl += f'TEXT {phys_x},{phys_y},"{font_str}",90,1,1,"{line.strip()}"\r\n'
                            
                            px = int(el.get("padX", 0))
                            py = int(el.get("padY", 0))
                            rev_x = int(phys_x - char_h - py)
                            rev_y = int(phys_y - px)
                            rev_w = int(char_h + (2 * py))
                            rev_h = int(line_width + (2 * px))
                            if rev_y < 0:
                                rev_h += rev_y
                                rev_y = 0
                            if rev_x < 0:
                                rev_w += rev_x
                                rev_x = 0
                            tspl += f'REVERSE {rev_x},{rev_y},{rev_w},{rev_h}\r\n'
                        else:
                            phys_y = int(el["x"] + x_adj)
                            tspl += f'TEXT {phys_x},{phys_y},"{font_str}",90,1,1,"{line.strip()}"\r\n'

                elif el["type"] == "barcode":
                    thick = int(el.get("thickness", 2))
                    bw = 90 if thick == 1 else 160
                    align = el.get("align", "left")
                    x_adj = 0
                    if align == "center": x_adj = int(-(bw / 2))
                    elif align == "right": x_adj = int(-bw)
                    
                    x_base = 300 if side == 0 else 600
                    phys_x = int(x_base - el["y"])
                    phys_y = int(el["x"] + x_adj)
                    phys_rot = (int(el.get("rotation",0)) + 90) % 360
                    tspl += f'BARCODE {phys_x},{phys_y},"128",{el.get("height",50)},1,{phys_rot},{thick},{thick},"{txt}"\r\n'
        
        mid_x = int(layout.get("right_x_offset", 300))
        if mid_x > 0:
            for dash_y in range(0, 400, 16):
                tspl += f"BAR {mid_x},{dash_y},2,8\r\n"
        
        tspl += f"PRINT 1, {quantity}\r\n"
        
        temp_file = "temp_print_job.prn"
        with open(temp_file, "wb") as f:
            f.write(tspl.encode('utf-8'))
        process = subprocess.run(['lp', '-d', selected_printer, '-o', 'raw', temp_file], capture_output=True, text=True)
        
        if process.returncode == 0:
            st.success("Successfully sent to printer!")
        else:
            st.error(f"Printing failed: {process.stderr}")
    except Exception as e:
        st.error(f"Error printing: {e}")

def download_tspl(layout):
    tspl = "SIZE 75 mm, 50 mm\r\nGAP 3 mm, 0 mm\r\nCLS\r\n"
    h_map = {"1": 12, "2": 20, "3": 24, "4": 32, "5": 48}
    w_map = {"1": 8, "2": 12, "3": 14, "4": 24, "5": 32}
    
    for side in [0, 1]:
        for el in layout.get("elements", []):
            txt = str(st.session_state.live_data.get(el["id"], el.get("text", el["id"])))
            if el["id"] == "MRP": txt = txt.replace('₹', 'Rs.')
            
            if el["type"] == "text":
                font_str = str(el.get("font", "2"))
                h = h_map.get(font_str, 20)
                w = w_map.get(font_str, 12)
                align = el.get("align", "left")
                
                lines = txt.split("\n")
                line_height = h + 4
                char_h = h
                
                for i, line in enumerate(lines):
                    if not line.strip(): continue
                    line_width = len(line.strip()) * w
                    x_adj = 0
                    if align == "center": x_adj = int(-(line_width / 2))
                    elif align == "right": x_adj = int(-line_width)
                    
                    x_base = 300 if side == 0 else 600
                    phys_x = int(x_base - el["y"] - (i * line_height))
                    phys_y = int(el["x"] + x_adj)
                    
                    tspl += f'TEXT {phys_x},{phys_y},"{font_str}",90,1,1,"{line.strip()}"\r\n'
                    if el.get("invert"):
                        px = int(el.get("padX", 0))
                        py = int(el.get("padY", 0))
                        rev_x = int(phys_x - char_h - py)
                        rev_y = int(phys_y - px)
                        rev_w = int(char_h + (2 * py))
                        rev_h = int(line_width + (2 * px))
                        if rev_y < 0:
                            rev_h += rev_y
                            rev_y = 0
                        if rev_x < 0:
                            rev_w += rev_x
                            rev_x = 0
                        tspl += f'REVERSE {rev_x},{rev_y},{rev_w},{rev_h}\r\n'

            elif el["type"] == "barcode":
                thick = int(el.get("thickness", 2))
                bw = 90 if thick == 1 else 160
                align = el.get("align", "left")
                x_adj = 0
                if align == "center": x_adj = int(-(bw / 2))
                elif align == "right": x_adj = int(-bw)
                
                x_base = 300 if side == 0 else 600
                phys_x = int(x_base - el["y"])
                phys_y = int(el["x"] + x_adj)
                phys_rot = (int(el.get("rotation",0)) + 90) % 360
                tspl += f'BARCODE {phys_x},{phys_y},"128",{el.get("height",50)},1,{phys_rot},{thick},{thick},"{txt}"\r\n'
    
    mid_x = int(layout.get("right_x_offset", 300))
    if mid_x > 0:
        for dash_y in range(0, 400, 16):
            tspl += f"BAR {mid_x},{dash_y},2,8\r\n"
    
    tspl += "PRINT 1, 1\r\n"
    return tspl

def generate_tspl(layout, copies):
    try:
        quantity = int(copies)
        
        for p in st.session_state.printers:
            if p["name"] == layout.get("printer"):
                selected_printer = p["name"]
                break
        else:
            if not st.session_state.printers:
                st.error("No printers found.")
                return
            selected_printer = st.session_state.printers[0]["name"]
        
        tspl = download_tspl(layout)
        tspl = tspl.replace("PRINT 1, 1\r\n", f"PRINT 1, {quantity}\r\n")
        
        temp_file = "temp_print_job.prn"
        with open(temp_file, "wb") as f:
            f.write(tspl.encode('utf-8'))
        process = subprocess.run(['lp', '-d', selected_printer, '-o', 'raw', temp_file], capture_output=True, text=True)
        
        if process.returncode == 0:
            st.success("Successfully sent to printer!")
        else:
            st.error(f"Printing failed: {process.stderr}")
    except Exception as e:
        st.error(f"Error printing: {e}")
