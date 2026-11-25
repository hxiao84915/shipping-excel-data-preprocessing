import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

print("Starting to build full manual...")

doc = docx.Document()

# ---------------------------------------------------------------------------
# Page Setup: Standard A4, 1 inch margins
# ---------------------------------------------------------------------------
for section in doc.sections:
    section.page_width = Inches(8.27)  # A4
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1.0)
    section.bottom_margin = Inches(1.0)
    section.left_margin = Inches(1.0)
    section.right_margin = Inches(1.0)
    
    # Header & Footer setup
    header = section.header
    h_p = header.paragraphs[0]
    h_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    h_run = h_p.add_run("航運行政人員 MS Excel 基礎資料分析前處理實戰手冊")
    h_run.font.name = "Microsoft JhengHei"
    h_run.font.size = Pt(8.5)
    h_run.font.color.rgb = RGBColor(128, 128, 128)
    
    footer = section.footer
    f_p = footer.paragraphs[0]
    f_p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    f_run = f_p.add_run("頁碼：")
    f_run.font.name = "Microsoft JhengHei"
    f_run.font.size = Pt(9)
    f_run.font.color.rgb = RGBColor(100, 100, 100)
    
    fld_page = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
    f_p._p.append(fld_page)
    
    f_run_mid = f_p.add_run(" / ")
    f_run_mid.font.name = "Microsoft JhengHei"
    f_run_mid.font.size = Pt(9)
    f_run_mid.font.color.rgb = RGBColor(100, 100, 100)
    
    fld_numpages = parse_xml(r'<w:fldSimple %s w:instr="NUMPAGES"/>' % nsdecls('w'))
    f_p._p.append(fld_numpages)

# ---------------------------------------------------------------------------
# Style Helpers & XML manipulators
# ---------------------------------------------------------------------------
COLOR_PRIMARY = RGBColor(27, 54, 93)     # Deep Navy Blue #1B365D
COLOR_SECONDARY = RGBColor(0, 120, 212) # Blue #0078D4
COLOR_ACCENT = RGBColor(16, 124, 65)    # Excel Green #107C41
COLOR_DARK = RGBColor(32, 31, 30)       # #201F1E
COLOR_MUTED = RGBColor(96, 94, 92)      # #605E5C

def set_cell_background(cell, hex_color):
    shading_xml = f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>'
    cell._tc.get_or_add_tcPr().append(parse_xml(shading_xml))

def set_cell_margins(cell, top=120, bottom=120, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table, color="D9D9D9", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>
            <w:insideV w:val="none"/>
            <w:left w:val="none"/>
            <w:right w:val="none"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def add_p(doc, text="", font_size=10.5, bold=False, italic=False, color=COLOR_DARK, align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=4, line_spacing=1.2):
    p = doc.add_paragraph()
    p.alignment = align
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = line_spacing
    if text:
        r = p.add_run(text)
        r.font.name = "Microsoft JhengHei"
        r.font.size = Pt(font_size)
        r.font.bold = bold
        r.font.italic = italic
        r.font.color.rgb = color
    return p

def add_h1(doc, text, is_first=False):
    if not is_first:
        doc.add_page_break()
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = "Microsoft JhengHei"
    r.font.size = Pt(18)
    r.font.bold = True
    r.font.color.rgb = COLOR_PRIMARY
    return p

def add_h2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(11)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = "Microsoft JhengHei"
    r.font.size = Pt(14)
    r.font.bold = True
    r.font.color.rgb = COLOR_SECONDARY
    return p

def add_h3(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.keep_with_next = True
    r = p.add_run(text)
    r.font.name = "Microsoft JhengHei"
    r.font.size = Pt(11.5)
    r.font.bold = True
    r.font.color.rgb = COLOR_ACCENT
    return p

def add_callout(doc, box_type, title, text):
    # Types: NOTE (#EFF6FC, #0078D4), TIP (#E2F0D9, #107C41), WARNING (#FDE7E9, #A80000)
    conf = {
        "NOTE": {"bg": "EFF6FC", "border": "0078D4", "title": f"【航運實務知識】{title}", "title_col": RGBColor(0, 78, 140)},
        "TIP": {"bg": "F0F9F0", "border": "107C41", "title": f"【操作技巧提示】{title}", "title_col": RGBColor(16, 124, 65)},
        "WARNING": {"bg": "FFF0F0", "border": "C00000", "title": f"【常見操作防呆】{title}", "title_col": RGBColor(192, 0, 0)}
    }
    c = conf.get(box_type, conf["NOTE"])
    
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    set_cell_background(cell, c["bg"])
    set_cell_margins(cell, top=140, bottom=140, left=180, right=180)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders_xml = f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="none"/>
            <w:left w:val="single" w:sz="24" w:space="0" w:color="{c['border']}"/>
            <w:bottom w:val="none"/>
            <w:right w:val="none"/>
        </w:tcBorders>
    '''
    tcPr.append(parse_xml(borders_xml))
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(3)
    r_t = p.add_run(c["title"] + "\n")
    r_t.font.name = "Microsoft JhengHei"
    r_t.font.size = Pt(10)
    r_t.font.bold = True
    r_t.font.color.rgb = c["title_col"]
    
    r_body = p.add_run(text)
    r_body.font.name = "Microsoft JhengHei"
    r_body.font.size = Pt(9.5)
    r_body.font.color.rgb = COLOR_DARK
    
    # Add empty spacing after table
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(0)
    p_sp.paragraph_format.space_after = Pt(4)

def add_figure(doc, img_path, caption, width=Inches(6.2)):
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(2)
        run_img = p_img.add_run()
        run_img.add_picture(img_path, width=width)
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(2)
        p_cap.paragraph_format.space_after = Pt(8)
        r_cap = p_cap.add_run(caption)
        r_cap.font.name = "Microsoft JhengHei"
        r_cap.font.size = Pt(9)
        r_cap.font.bold = True
        r_cap.font.color.rgb = COLOR_MUTED
    else:
        print(f"Warning: image {img_path} not found!")

def add_table_data(doc, headers, data_rows, col_widths=None, alignments=None, caption=None):
    if caption:
        p_cap = doc.add_paragraph()
        p_cap.paragraph_format.space_before = Pt(6)
        p_cap.paragraph_format.space_after = Pt(2)
        p_cap.paragraph_format.keep_with_next = True
        r_cap = p_cap.add_run(caption)
        r_cap.font.name = "Microsoft JhengHei"
        r_cap.font.size = Pt(9.5)
        r_cap.font.bold = True
        r_cap.font.color.rgb = COLOR_PRIMARY
        
    num_rows = len(data_rows) + 1
    num_cols = len(headers)
    tbl = doc.add_table(rows=num_rows, cols=num_cols)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(tbl, color="D0D7DE", sz="4")
    
    # Header Row
    hdr_cells = tbl.rows[0].cells
    for i, title in enumerate(headers):
        cell = hdr_cells[i]
        set_cell_background(cell, "1B365D")
        set_cell_margins(cell, top=100, bottom=100, left=120, right=120)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        r = p.add_run(str(title))
        r.font.name = "Microsoft JhengHei"
        r.font.size = Pt(9)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        
    # Repeat header row across pages
    trPr = tbl.rows[0]._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    
    # Data Rows
    for r_idx, row_values in enumerate(data_rows):
        row_cells = tbl.rows[r_idx + 1].cells
        bg_col = "F7F9FB" if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_values):
            cell = row_cells[c_idx]
            set_cell_background(cell, bg_col)
            set_cell_margins(cell, top=70, bottom=70, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            
            align = WD_ALIGN_PARAGRAPH.LEFT
            if alignments and c_idx < len(alignments):
                align = alignments[c_idx]
            p.alignment = align
            
            r = p.add_run(str(val))
            r.font.name = "Microsoft JhengHei"
            r.font.size = Pt(8.5)
            r.font.color.rgb = COLOR_DARK
            
    # Set widths if specified
    if col_widths:
        for row in tbl.rows:
            for c_idx, w in enumerate(col_widths):
                if c_idx < len(row.cells):
                    row.cells[c_idx].width = Inches(w)
                    
    p_sp = doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(0)
    p_sp.paragraph_format.space_after = Pt(4)

print("Helper functions ready.")
