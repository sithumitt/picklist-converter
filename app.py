import os
import io
import re
import pdfplumber
from docx import Document
import streamlit as st

# Set up the web page title, icon, and layout profile
st.set_page_config(
    page_title="පික් ලිස්ට් පරිවර්තකය", 
    page_icon="📋", 
    layout="centered",
    initial_sidebar_state="collapsed"
)

# Custom CSS for soft light blue palette, black fonts, and light green table hovers
st.markdown("""
    <style>
    /* Main background and global font color */
    .stApp {
        background-color: #F0F4F8;
    }
    
    html, body, [data-testid="stWidgetLabel"], p, div, h1, h2, h3, span {
        color: #000000 !important;
        font-family: 'Segoe UI', Helvetica, Arial, sans-serif;
    }
    
    /* Header title styling */
    .main-title {
        color: #0D47A1;
        font-size: 2.4rem;
        font-weight: 600;
        text-align: center;
        margin-bottom: 0.5rem;
    }
    
    .sub-title {
        color: #333333;
        font-size: 1.1rem;
        text-align: center;
        margin-bottom: 2.5rem;
    }
    
    /* Light Blue elegant card styling for data preview */
    .preview-card {
        background-color: #FFFFFF;
        border-radius: 8px;
        padding: 1.5rem;
        box-shadow: 0 4px 6px rgba(0,0,0,0.02), 0 1px 3px rgba(0,0,0,0.05);
        border: 2px solid #BBDEFB;
        margin-top: 1.5rem;
    }
    
    /* Force Dataframe tables to use light green highlight color during hovers instead of black */
    [data-testid="stTable"] tr:hover, 
    [data-testid="stDataFrame"] tr:hover,
    div[data-role="grid"] div[role="row"]:hover {
        background-color: #E8F5E9 !important;
    }
    
    /* Ensure internal canvas data viewer matches light cell styling accents */
    .glideDataEditor-canvas {
        --bg-color-hover: #E8F5E9 !important;
        --accent-color: #C8E6C9 !important;
    }
    
    /* Soft Blue Step indicators style */
    .step-container {
        background-color: #E3F2FD;
        padding: 1rem;
        border-left: 5px solid #2196F3;
        border-radius: 4px;
        margin-bottom: 1.5rem;
        color: #000000;
        font-weight: 500;
    }
    
    /* File Uploader custom styling adjustments */
    [data-testid="stFileUploaderDropzone"] {
        background-color: #FFFFFF;
        border: 2px dashed #90CAF9 !important;
        border-radius: 8px;
    }
    
    /* Custom style tables to maintain black text grid lines */
    table {
        color: #000000 !important;
    }
    
    /* Hide default Streamlit decorations */
    footer {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    </style>
""", unsafe_allow_html=True)

# Hardcoded Product Mapping Dictionary
PRODUCT_MAPPING = {
    "aloe vera drink 1l": "කෝමාරිකා බීම ලීටර් 1",
    "aloe vera drink 200ml": "කෝමාරිකා බීම 200",
    "aloe vera drink 500ml": "කෝමාරිකා බීම 500",
    "assortment 150gm": "ඇසෝඩ්මන්ට් බිස්කට් 150",
    "assortment 300gm": "ඇසෝඩ්මන්ට් බිස්කට් 300",
    "butter cookies 120gm": "බටර් කුකීස් 120",
    "cheese cuts tin 210gm": "චීස් බිස්කට් රතු ටින්",
    "chilli paste 60g": "චිලී පේස්ට් 60",
    "choco shorties 220gm": "චොකෝ ශෝටීස් 220",
    "chocolate chips cookies 120gm": "චොකලට් චිප් කුකීස් 120",
    "chocolate chips cookies 65gm": "චොකලට් චිප් කුකීස් 65",
    "chocolate cream 500gm": "චොකලට් ක්‍රීම්",
    "chocolate wafer 40gm": "චොකලට වේපස් 40",
    "chocolate wafer 90gm": "චොකලට වේපස් 90",
    "chocolate wafer 200gm": "චොකලට වේපස් 200",
    "chocolate wafer 360gm x 3pks": "චොකලට වේපස් 360",
    "cookie assortment blue 330gm": "කුකි ඇසෝඩ්මන්ට් නිල් 330",
    "coffee wafer 200gm": "කෝපි වේපර්ස්",
    "creamy choc 90gm": "ක්‍රීමි චොක් 90",
    "creamy choc 210gm": "ක්‍රීමි චොක් 210",
    "creamy choc 360gm x 4pks": "ක්‍රීමි චොක් 360",
    "creamy vanilla 360gm": "ක්‍රීම් වැනිලා 360",
    "falooda wafer 200gm": "ෆාලූඩා වේපස් 200",
    "ginger 80gm": "ඉඟුරු බිස්කට් 80",
    "ginger 250gm": "ඉඟුරු බිස්කට් 250",
    "green apple sparkling 250ml": "ඇපල් ස්පාක්ලින්",
    "kist cream cracker 125gm": "ක්‍රීම් ක්‍රැකර් 125",
    "kist cream cracker 500gm": "ක්‍රීම් ක්‍රැකර් 500",
    "kist gift assortment 400gm": "ගීෆ්ට් ඇසෝඩ්මන්ට් 400",
    "kithul treacle 340ml": "කිතුල් පැණි 340",
    "knuckles water 500ml": "වතුර 500",
    "knuckles water 1000ml": "වතුර 1000",
    "knuckles water 1500ml": "වතුර 1500",
    "lemon & mint nectar 1l": "ලෙමන් සහ මින්ට් බීම 1l",
    "lemon puff 100gm": "ලෙමන් පෆ් 100",
    "lemon puff 200gm": "ලෙමන් පෆ් 200",
    "magic choco fun 40gm": "චෝෆන් කහ පාට 40",
    "magic choco fun 100gm": "චෝෆන් කහ පාට 100",
    "magic choco nut 160gm": "චොකෝ නට් පොල් 160",
    "magic chocooh 100gm": "චොකො ඕ [සුදු පාට ]",
    "magic chocoblok chocolat 160gm": "චොකෝ බ්ලොක් චොකලට් 160",
    "magic chocoblok coffee 160gm": "චොකෝ බ්ලොක් කෝපි",
    "magic chocoblok vanilla 160gm": "චොකෝ බ්ලොක් වැනිලා",
    "magic chokstik choco 10gm": "චොකලට් චොක්ස්ටික්",
    "magic chokstik strawberry 10gm": "ස්ටෝබරි චොක්ස්ටික්",
    "magic chokstik vanilla 10gm": "වැනිලා චොක්ස්ටික්",
    "magic chonkz 120gm": "චෝන්ස් ලා නිල් 120",
    "magic chonkz 240gm": "චෝන්ස් ලා නිල් 240",
    "magic fingers 27gm": "ෆින්ගස් 27",
    "magic fingers 100gm": "ෆින්ගස් 100",
    "magic olo assortment 180gm": "ඔලෝ ඇසෝර්ට්මන් 180",
    "magic olo butter scotch 140gm": "ඕලෝ බටර් ස්කොච් 140",
    "magic olo chocolate 60gm": "ඕලෝ චොකලට් 60",
    "magic olo chocolate 140gm": "ඕලෝ චොකලට් 140",
    "magic olo classic 60gm": "ඔලෝ ක්ලැසික් 60",
    "magic olo classic 140gm": "ඔලෝ ක්ලැසික් 140",
    "magic olo strawberry 60gm": "ඔලෝ ස්ට්‍රෝබෙරි 60",
    "magic olo strawberry 140gm": "ඔලෝ ස්ට්‍රෝබෙරි 140",
    "magic olo white 140gm": "ඔලෝ වයිට් 140",
    "magic treats 90gm": "ටීට්ස් දම් පාට",
    "mango nectar 1l": "අඹ බීම 1l",
    "mango nectar 200ml": "අඹ බීම 200",
    "mango nectar 500ml": "අඹ බීම 500",
    "marie 50gm": "මාරි 50",
    "marie 100gm": "මාරි 100",
    "marie 200gm": "මාරි 200",
    "marie 500gm": "මාරි 500",
    "mayonnaise 200g pouch": "මයෝනාඊස් 200",
    "milk shorties 220gm": "කිරි ෂෝටීස් 220",
    "milki cookies 120gm": "මීල්කි කුකීස්",
    "mixed fruit jam cup 100g": "මිශ්‍ර ජෑම් කප් 100",
    "mixed fruit jam 200g": "මිශ්‍ර ජෑම් 200",
    "mixed fruit jam 300g": "මිශ්‍ර ජෑම් 300",
    "mixed fruit jam 510g": "මිශ්‍ර ජෑම් 510",
    "mixed fruit nectar 1l": "මිශ්‍ර පලතුරු බීම 1L",
    "mixed fruit nectar 200ml": "මිශ්‍ර පලතුරු බීම 200ML",
    "mixed fruit nectar 500ml": "මිශ්‍ර පලතුරු බීම 500ML",
    "nice 430gm": "නායිස් 430",
    "onion byte 30gm": "අනියන් බයිට් 30",
    "orange nectar 500ml": "දොඩම් බීම 500",
    "orange sparkling 250ml": "දොඩම් ස්පාක්ලින්",
    "passion fruit nectar 1l": "පැෂන්ෆෲට් නෙක්ටා 1L",
    "ride classic drink 250ml": "රයිට් නිල්",
    "ride redberry drink 250ml": "රයිට් රතු",
    "ride sugar free drink 250ml": "රයිට් සීනි නැති",
    "s/berry flv. melon jam200g": "ස්ටෝබරි ජෑම් 200",
    "s/berry flv. melon jam300g": "ස්ටෝබරි ජෑම් 300",
    "s/berry melon jam cup100g": "ස්ටෝබරි ජෑම් 100 C",
    "sesame cookies 120gm": "සෙසමිකුකීස්",
    "strawberry sparkling 250ml": "ස්ටෝබරි ස්පාක්ලින්",
    "strawberry wafer 40gm": "ස්ටෝබරි වේපස් 40",
    "strawberry wafer 90gm": "ස්ටෝබරි වේපස් 90",
    "strawberry wafer 200gm": "ස්ටෝබරි වේපස් 200",
    "strawberry wafer 360gm x 3pks": "ස්ටෝබරි වේපස් 360",
    "soya sauce squeeasy 180ml": "සෝස් පැකට් 180ml",
    "tomato sachet 15g": "සෝස් පැකට් 15",
    "tomato sauce 110gr - pouch": "සෝස් පැකට් 110",
    "tomato sauce 400g - pouch": "සෝස් පැකට් 400",
    "tomato sauce 400g ": "සෝස් වීදුරු බෝතලේ 400 ",
    "tomato sauce 200gr ": "සෝස් වීදුරු බෝතලේ 200 ",
    "vanilla wafer 40gm": "වැනිලා වේපස් 40",
    "vanilla wafer 90gm": "වැනිලා වේපස් 90",
    "vanilla wafer 200gm": "වැනිලා වේපස් 200",
    "vanilla wafer 360gm x 3pks": "වැනිලා වේපස් 360",
    "woodapple nectar 200ml": "දිවුල් බීම 200",
    "woodapple nectar 500ml": "දිවුල් බීම 500"
}

def clean_text_for_matching(text):
    """Removes slashes, dots, quotes, and extra whitespace for fuzzy matching."""
    if not text:
        return ""
    # Remove quotes, slashes, dots, and hyphens
    cleaned = text.replace('"', '').replace('/', ' ').replace('.', ' ').replace('-', ' ')
    # Standardize spaces to a single space, lowercase
    return " ".join(cleaned.lower().split())

# UI Header Layout Elements
st.markdown('<div class="main-title">📋 පික් ලිස්ට් එකේ බඩු පරිවර්තකය</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">ඔබේ Picklist PDF එක සිංහල Word ගොනුවක් බවට ක්ෂණිකව පරිවර්තනය කරන්න</div>', unsafe_allow_html=True)

# Step 1 Container Instruction Banner
st.markdown('<div class="step-container"><strong>පියවර 1:</strong> ඔබේ මුල් පිටපතේ PDF ගොනුව පහත කොටුවට එක් කරන්න (Upload PDF File)</div>', unsafe_allow_html=True)

# File Uploader
uploaded_file = st.file_uploader("", type=["pdf"], label_visibility="collapsed")

if uploaded_file is not None:
    with st.spinner("දත්ත විශ්ලේෂණය කරමින් පවතී. කරුණාකර රැඳී සිටින්න..."):
        # Initialize structured Word Document
        doc = Document()
        doc.add_heading('පික් ලිස්ට් එකේ බඩු', level=1)
        
        table = doc.add_table(rows=1, cols=3)
        table.style = 'Table Grid'
        
        hdr_cells = table.rows[0].cells
        hdr_cells[0].text = 'භාණ්ඩය'
        hdr_cells[1].text = 'කේස්'
        hdr_cells[2].text = 'කෑලි'
        
        matched_count = 0
        preview_data = []
        
        with pdfplumber.open(uploaded_file) as pdf:
            for page in pdf.pages:
                text_content = page.extract_text()
                if not text_content:
                    continue
                
                lines = text_content.split('\n')
                for line in lines:
                    normalized_line = line.replace('"', '').strip()
                    # Clean the line text completely for matching
                    matchable_line = clean_text_for_matching(normalized_line)
                    
                    for english_key, sinhala_val in PRODUCT_MAPPING.items():
                        # Clean the key completely as well
                        matchable_key = clean_text_for_matching(english_key)
                        
                        if matchable_key in matchable_line:
                            # 1. Primary Strategy: Look for standard batch identifiers (like CG1, CG2, IN1 etc.)
                            batch_match = re.search(r'\b([A-Z]{2}\d)\b', normalized_line)
                            
                            qty1, qty2 = "0", "0"
                            if batch_match:
                                pre_batch_text = normalized_line[:batch_match.start()].strip()
                                all_numbers = re.findall(r'\d+', pre_batch_text)
                                if len(all_numbers) >= 2:
                                    qty1 = all_numbers[-2]
                                    qty2 = all_numbers[-1]
                            else:
                                # 2. Fallback Strategy: Intelligently grab trailing numbers from the end of the line
                                all_numbers = re.findall(r'\b\d+\b', normalized_line)
                                if len(all_numbers) >= 2:
                                    qty1 = all_numbers[-2]
                                    qty2 = all_numbers[-1]
                            
                            row_cells = table.add_row().cells
                            row_cells[0].text = sinhala_val
                            row_cells[1].text = qty1
                            row_cells[2].text = qty2
                            
                            preview_data.append({"භාණ්ඩය": sinhala_val, "කේස්": qty1, "කෑලි": qty2})
                            matched_count += 1
                            break 

    if matched_count > 0:
        st.success(f"🎉 සාර්ථකයි! ගැළපෙන භාණ්ඩ පේළි {matched_count} ක් සාර්ථකව පරිවර්තනය කරන ලදී.")
        
        # Step 2 Container Instruction Banner
        st.markdown('<div class="step-container"><strong>පියවර 2:</strong> සකස් කරන ලද නව දත්ත පෙරදසුන පරීක්ෂා කර බාගත කරගන්න</div>', unsafe_allow_html=True)
        
        # Displaying preview inside the custom layout card
        st.markdown('<div class="preview-card">', unsafe_allow_html=True)
        st.dataframe(preview_data, use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
        
        st.write("")  # Spacer Element
        
        # Render output document to byte storage array stream
        doc_stream = io.BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        
        # Action download button
        st.download_button(
            label="📥 නිපදවන ලද Word ලිපිගොනුව බාගත කරගන්න (Download Document)",
            data=doc_stream,
            file_name="පික්_ලිස්ට්_එකේ_බඩු.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    else:
        st.error("⚠️ දෝෂයකි: අප්ලෝඩ් කරන ලද PDF ගොනුවේ තිබූ කිසිදු භාණ්ඩයක් අපගේ නාමාවලිය සමඟ ගැළපුණේ නැත. කරුණාකර වෙනත් ගොනුවක් උත්සාහ කරන්න.")
