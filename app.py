import streamlit as st
import io
import re
import pdfplumber
from docx import Document

# Set up the web page title and icon
st.set_page_config(page_title="පික් ලිස්ට් පරිවර්තකය", page_icon="📝", layout="centered")

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
    "mixed fruit nectar 1l": "මිශ්‍ර පලතුරු බීම 1L",
    "mixed fruit nectar 200ml": "මිශ්‍ර පලතුරු බීම 200ML",
    "mixed fruit nectar 500ml": "මිශ්‍ර පලතුරු බීම 500ML",
    "nice 430gm": "නායිස් 430",
    "onion byte 30gm": "අනියන් බයිට් 30",
    "orange nectar 500ml": "දොඩම් බීම 500",
    "orange sparkling 250ml": "දොඩම්",
    "passion fruit nectar 1l": "ස්පාක්ලින්",
    "ride classic drink 250ml": "රයිට් නිල්",
    "ride redberry drink 250ml": "රයිට් රතු",
    "ride sugar free drink 250ml": "රයිට් සීනි නැති",
    "s/berry flv. melon jam200g": "ස්ටෝබරි ජෑම් 200",
    "s/berry melon jam cup100g": "ස්ටෝබරි ජෑම් 100 C",
    "sesame cookies 120gm": "සෙසමිකුකීස්",
    "strawberry sparkling 250ml": "ස්ටෝබරි ස්පාක්ලින්",
    "strawberry wafer 40gm": "ස්ටෝබරි වේපස් 40",
    "strawberry wafer 90gm": "ස්ටෝබරි වේපස් 90",
    "strawberry wafer 200gm": "ස්ටෝබරි වේපස් 200",
    "strawberry wafer 360gm x 3pks": "ස්ටෝබරි වේපස් 360",
    "tomato sachet 15g": "සෝස් පැකට් 15",
    "tomato sauce 110gr - pouch": "සෝස් පැකට් 110",
    "tomato sauce 400g - pouch": "සෝස් පැකට් 400",
    "vanilla wafer 40gm": "වැනිලා වේපස් 40",
    "vanilla wafer 90gm": "වැනිලා වේපස් 90",
    "vanilla wafer 200gm": "වැනිලා වේපස් 200",
    "vanilla wafer 360gm x 3pks": "වැනිලා වේපස් 360",
    "woodapple nectar 200ml": "දිවුල් බීම 200",
    "woodapple nectar 500ml": "දිවුල් බීම 500"
}

st.title("📋 පික් ලිස්ට් එකේ බඩු පරිවර්තකය")
st.write("ඔබේ PDF ගොනුව මෙතැනට අප්ලෝඩ් කර Word (Docx) ගොනුව බාගත කරගන්න.")

# UI Element: File Uploader
uploaded_file = st.file_uploader("PDF එක තෝරන්න (Choose PDF file)", type=["pdf"])

if uploaded_file is not None:
    st.info("ගොනුව සකසමින් පවතී, කරුණාකර රැඳී සිටින්න...")
    
    # Initialize Word Document
    doc = Document()
    doc.add_heading('පික් ලිස්ට් එකේ බඩු', level=1)
    
    table = doc.add_table(rows=1, cols=3)
    table.style = 'Table Grid'
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'භාණ්ඩය'
    hdr_cells[1].text = 'කේස්'
    hdr_cells[2].text = 'කෑලි'
    
    matched_count = 0
    preview_data = [] # To show a sneak peek in the web UI
    
    # Read the PDF content from the upload stream
    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            text_content = page.extract_text()
            if not text_content:
                continue
            
            lines = text_content.split('\n')
            for line in lines:
                normalized_line = line.replace('"', '').strip()
                normalized_line_lower = normalized_line.lower()
                
                for english_key, sinhala_val in PRODUCT_MAPPING.items():
                    if english_key in normalized_line_lower:
                        batch_match = re.search(r'(CG\d)', normalized_line)
                        
                        qty1, qty2 = "0", "0"
                        if batch_match:
                            pre_batch_text = normalized_line[:batch_match.start()].strip()
                            all_numbers = re.findall(r'\d+', pre_batch_text)
                            
                            if len(all_numbers) >= 2:
                                qty1 = all_numbers[-2]
                                qty2 = all_numbers[-1]
                        
                        # Add to Word Doc table
                        row_cells = table.add_row().cells
                        row_cells[0].text = sinhala_val
                        row_cells[1].text = qty1
                        row_cells[2].text = qty2
                        
                        # Save a sample for the UI preview table
                        preview_data.append({"භාණ්ඩය": sinhala_val, "කේස්": qty1, "කෑලි": qty2})
                        matched_count += 1
                        break 

    if matched_count > 0:
        st.success(f"සාර්ථකයි! භාණ්ඩ පේළි {matched_count} ක් සොයාගන්නා ලදී.")
        
        # Display an interactive preview table directly on the webpage
        st.subheader("දත්ත පෙරදසුන (Data Preview)")
        st.table(preview_data[:10]) # Show the first 10 rows as preview
        if matched_count > 10:
            st.write(f"...සහ තවත් පේළි {matched_count - 10}ක් බාගත වන ගොනුවේ අඩංගු වේ.")
        
        # Save document structure to bytes buffer
        doc_stream = io.BytesIO()
        doc.save(doc_stream)
        doc_stream.seek(0)
        
        # UI Element: Download Button
        st.download_button(
            label="📄 Word File එක බාගත කරගන්න (Download)",
            data=doc_stream,
            file_name="පික්_ලිස්ට්_එකේ_බඩු.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.error("කිසිදු ගැළපෙන භාණ්ඩයක් සොයාගත නොහැකි විය. කරුණාකර නිවැරදි PDF එකක්දැයි පරීක්ෂා කරන්න.")
