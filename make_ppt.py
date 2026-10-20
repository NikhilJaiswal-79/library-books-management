"""
Generates Library Book Management Project PPTX
Requires: pip install python-pptx
"""
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn

# ---------- Colors ----------
NAVY = RGBColor(0x0F, 0x17, 0x2A)
BLUE = RGBColor(0x3B, 0x82, 0xF6)
CYAN = RGBColor(0x06, 0xB6, 0xD4)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xF1, 0xF5, 0xF9)
GREY = RGBColor(0x64, 0x74, 0x8B)
EMERALD = RGBColor(0x10, 0xB9, 0x81)
ROSE = RGBColor(0xF4, 0x3F, 0x5E)
SLATE = RGBColor(0x1E, 0x29, 0x3B)

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]

def add_slide():
    return prs.slides.add_slide(BLANK)

def add_rect(slide, l, t, w, h, color, line=False):
    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, l, t, w, h)
    shp.fill.solid()
    shp.fill.fore_color.rgb = color
    if line:
        shp.line.color.rgb = line
    else:
        shp.line.fill.background()
    shp.shadow.inherit = False
    return shp

def add_text(slide, l, t, w, h, runs, size=18, color=GREY, bold=False, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, space_after=6, line_spacing=1.0):
    """runs: list of paragraphs; each paragraph is list of (text, dict_overrides)"""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    tf.vertical_anchor = anchor
    tf.margin_left = Inches(0.1)
    tf.margin_right = Inches(0.1)
    tf.margin_top = Inches(0.05)
    tf.margin_bottom = Inches(0.05)
    first = True
    for para in runs:
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.alignment = align
        p.space_after = Pt(space_after)
        p.line_spacing = line_spacing
        if isinstance(para, str):
            para = [(para, {})]
        for text, ov in para:
            r = p.add_run()
            r.text = text
            r.font.size = Pt(ov.get("size", size))
            r.font.bold = ov.get("bold", bold)
            r.font.italic = ov.get("italic", False)
            r.font.color.rgb = ov.get("color", color)
            r.font.name = "Calibri"
    return box

def add_title(slide, text, subtitle=None):
    add_rect(slide, 0, 0, prs.slide_width, Inches(0.12), BLUE)
    add_text(slide, Inches(0.6), Inches(0.35), Inches(12.2), Inches(0.9),
             [[(text, {"size": 34, "bold": True, "color": NAVY})]], space_after=0)
    if subtitle:
        add_text(slide, Inches(0.6), Inches(1.15), Inches(12.2), Inches(0.5),
                 [[(subtitle, {"size": 16, "color": GREY})]], space_after=0)

def add_footer(slide, page_num):
    add_text(slide, Inches(0.6), Inches(7.05), Inches(9), Inches(0.4),
             [[("Library Book Management System  |  Project Presentation", {"size": 10, "color": GREY})]], space_after=0)
    add_text(slide, Inches(12.3), Inches(7.05), Inches(0.6), Inches(0.4),
             [[(str(page_num), {"size": 10, "color": GREY, "bold": True})]], align=PP_ALIGN.RIGHT, space_after=0)

def add_bullets(slide, l, t, w, h, items, size=16, color=GREY, space_after=8, bullet_color=BLUE):
    """items: list of (text, level) or (text, level, bold)"""
    box = slide.shapes.add_textbox(l, t, w, h)
    tf = box.text_frame
    tf.word_wrap = True
    first = True
    for item in items:
        text = item[0]
        level = item[1] if len(item) > 1 else 0
        bold = item[2] if len(item) > 2 else False
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        first = False
        p.level = level
        p.space_after = Pt(space_after)
        p.line_spacing = 1.05
        prefix = "  " if level == 0 else "  "
        r = p.add_run()
        r.text = prefix + text
        r.font.size = Pt(size if level == 0 else size - 2)
        r.font.bold = bold
        r.font.color.rgb = NAVY if bold else color
        r.font.name = "Calibri"
    return box

def style_cell(cell, text, size=12, bold=False, color=WHITE, fill=None, align=PP_ALIGN.LEFT):
    cell.margin_left = Inches(0.06)
    cell.margin_right = Inches(0.06)
    cell.margin_top = Inches(0.03)
    cell.margin_bottom = Inches(0.03)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    if fill is not None:
        cell.fill.solid()
        cell.fill.fore_color.rgb = fill
    tf = cell.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = align
    r = p.add_run()
    r.text = text
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.color.rgb = color
    r.font.name = "Calibri"

def add_table(slide, l, t, w, h, data, col_widths=None, header_fill=NAVY, size=12, header_size=13):
    rows = len(data)
    cols = len(data[0])
    shape = slide.shapes.add_table(rows, cols, l, t, w, h)
    table = shape.table
    if col_widths:
        total = sum(col_widths)
        for i, cw in enumerate(col_widths):
            table.columns[i].width = Emu(int(w * cw / total))
    for ri, row in enumerate(data):
        for ci, val in enumerate(row):
            cell = table.cell(ri, ci)
            if ri == 0:
                style_cell(cell, str(val), size=header_size, bold=True, color=WHITE, fill=header_fill,
                           align=PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT)
            else:
                fill = LIGHT if ri % 2 == 0 else WHITE
                style_cell(cell, str(val), size=size, bold=False, color=NAVY, fill=fill,
                           align=PP_ALIGN.CENTER if ci > 0 else PP_ALIGN.LEFT)
    return table

# ============================================================
# SLIDE 1: TITLE
# ============================================================
s = add_slide()
add_rect(s, 0, 0, prs.slide_width, prs.slide_height, NAVY)
add_rect(s, 0, Inches(2.3), prs.slide_width, Inches(0.06), BLUE)
add_text(s, Inches(0.9), Inches(1.0), Inches(11.5), Inches(1.2),
         [[("Library Book Usage Analysis", {"size": 44, "bold": True, "color": WHITE})]], space_after=0)
add_text(s, Inches(0.9), Inches(2.0), Inches(11.5), Inches(0.6),
         [[("& Circulation System", {"size": 44, "bold": True, "color": CYAN})]], space_after=0)
add_text(s, Inches(0.9), Inches(2.7), Inches(11.5), Inches(0.6),
         [[("A Modern Library Circulation Management & Usage Analytics Web Platform", {"size": 20, "color": LIGHT})]], space_after=0)
add_text(s, Inches(0.9), Inches(3.5), Inches(11.5), Inches(0.5),
         [[("Tech Stack:  ", {"size": 16, "bold": True, "color": BLUE}),
           ("Python  Flask  Gunicorn  MySQL  Pandas  NumPy  Chart.js  HTML5/CSS3/JS (SPA)", {"size": 16, "color": LIGHT})]], space_after=0)
add_text(s, Inches(0.9), Inches(5.2), Inches(11.5), Inches(0.5),
         [[("Presented by:  ", {"size": 16, "bold": True, "color": BLUE}), ("[Your Name]", {"size": 16, "color": LIGHT})]], space_after=0)
add_text(s, Inches(0.9), Inches(5.7), Inches(11.5), Inches(0.5),
         [[("Demo:  ", {"size": 16, "bold": True, "color": BLUE}), ("http://127.0.0.1:8000", {"size": 16, "color": LIGHT})]], space_after=0)

# ============================================================
# SLIDE 2: PROJECT OVERVIEW
# ============================================================
s = add_slide()
add_title(s, "Project Overview", "What the system does and why it exists")
add_rect(s, Inches(0.6), Inches(1.6), Inches(6.0), Inches(5.2), SLATE)
add_rect(s, Inches(0.6), Inches(1.6), Inches(6.0), Inches(0.5), NAVY)
add_text(s, Inches(0.8), Inches(1.68), Inches(5.6), Inches(0.4),
         [[("WHAT IT IS", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
add_bullets(s, Inches(0.8), Inches(2.25), Inches(5.6), Inches(4.5), [
    ("Full-stack library circulation management system.", 0, True),
    ("Tracks book issues, returns, borrowers, and inventory in real time.", 0),
    ("Generates usage analytics with interactive charts and statistical reports.", 0),
    ("Replaces manual library registers with a digital, searchable platform.", 0),
], size=14, color=LIGHT)

add_rect(s, Inches(6.8), Inches(1.6), Inches(5.95), Inches(5.2), SLATE)
add_rect(s, Inches(6.8), Inches(1.6), Inches(5.95), Inches(0.5), NAVY)
add_text(s, Inches(7.0), Inches(1.68), Inches(5.5), Inches(0.4),
         [[("CORE PURPOSE & PAIN POINTS", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
add_bullets(s, Inches(7.0), Inches(2.25), Inches(5.6), Inches(4.5), [
    ("Manual record errors and lost paperwork eliminated.", 0),
    ("No overdue alerts \u2192 automated overdue detection and badges.", 0),
    ("No usage insights \u2192 7 interactive analytics charts + KPIs.", 0),
    ("Slow catalogue search \u2192 live search and category filters.", 0),
    ("No inventory visibility \u2192 real-time stock pills and allocation.", 0),
], size=14, color=LIGHT)
add_footer(s, 2)

# ============================================================
# SLIDE 3: TECHNOLOGY STACK
# ============================================================
s = add_slide()
add_title(s, "Technology Stack", "Backend, Database, Data Analytics, and Frontend")
data = [
    ["Layer", "Technology", "Purpose"],
    ["Language", "Python 3.10+", "Core backend logic & transactions"],
    ["Backend Framework", "Flask + Gunicorn", "Lightweight WSGI web framework for REST API"],
    ["Database", "MySQL Server 8.0+", "Relational data store (auto-schema on first run)"],
    ["DB Driver", "mysql-connector-python", "Connects Python to MySQL"],
    ["Configuration", "python-dotenv", "Loads credentials from .env (git-ignored)"],
    ["Data Analysis", "Pandas", "Category groupings, monthly circulation trends"],
    ["Statistics", "NumPy", "Mean, median, std dev, percentiles, IQR"],
    ["Visualization (Web)", "Chart.js (HTML5 Canvas)", "7 interactive vector charts \u2014 no static images"],
    ["Visualization (CLI)", "Matplotlib", "Saves static PNG charts to reports_and_charts/"],
    ["Frontend", "HTML5, CSS3, JavaScript (SPA)", "Glassmorphism UI, real-time stats, search filters"],
]
add_table(s, Inches(0.6), Inches(1.7), Inches(12.1), Inches(5.0), data,
          col_widths=[1.8, 2.4, 7.9], size=12, header_size=13)
add_footer(s, 3)

# ============================================================
# SLIDE 4: SYSTEM ARCHITECTURE
# ============================================================
s = add_slide()
add_title(s, "System Architecture", "Client \u2192 Flask Backend \u2192 MySQL Database")
layers = [
    ("CLIENT (Browser)", "HTML5/CSS3/JS SPA\nChart.js Canvas Charts\nFetch API \u2192 REST Endpoints", BLUE),
    ("FLASK BACKEND", "server.py \u2192 Routes + Static Files\noperations.py \u2192 CRUD & Circulation Logic\nanalytics.py \u2192 Pandas + NumPy\ndatabase.py \u2192 MySQL Connection & Init", CYAN),
    ("DATABASE LAYER", "MySQL Server 8.0+ (InnoDB)\nTables: books \u2502 borrowers \u2502 loans\nmysql-connector-python", EMERALD),
]
x = Inches(0.6)
for name, desc, color in layers:
    w = Inches(3.9)
    add_rect(s, x, Inches(1.7), w, Inches(4.6), color)
    add_rect(s, x, Inches(1.7), w, Inches(0.6), NAVY)
    add_text(s, x + Inches(0.15), Inches(1.78), w - Inches(0.3), Inches(0.5),
             [[(name, {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
    add_text(s, x + Inches(0.15), Inches(2.5), w - Inches(0.3), Inches(3.6),
             [[(desc, {"size": 12, "color": WHITE})]], space_after=4, line_spacing=1.1)
    x += Inches(4.15)
add_text(s, Inches(0.6), Inches(6.5), Inches(12.1), Inches(0.5),
         [[("Data Flow:  ", {"size": 13, "bold": True, "color": NAVY}),
           ("Browser clicks \u2192 JavaScript fetch() \u2192 Flask route \u2192 business logic \u2192 MySQL query \u2192 JSON response \u2192 UI updates", {"size": 13, "color": GREY})]], space_after=0)
add_footer(s, 4)

# ============================================================
# SLIDE 5: MYSQL SCHEMA
# ============================================================
s = add_slide()
add_title(s, "MySQL Database Schema", "3 tables created automatically via database.initialize_database()")

# Table: books
add_rect(s, Inches(0.6), Inches(1.7), Inches(3.85), Inches(5.0), SLATE)
add_rect(s, Inches(0.6), Inches(1.7), Inches(3.85), Inches(0.5), NAVY)
add_text(s, Inches(0.75), Inches(1.78), Inches(3.5), Inches(0.4),
         [[("TABLE: books", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
books = [
    ["Column", "Type", "Constraint"],
    ["book_id", "INT", "PK, AUTO_INCREMENT"],
    ["isbn", "VARCHAR(20)", "NOT NULL, UNIQUE"],
    ["title", "VARCHAR(150)", "NOT NULL"],
    ["author", "VARCHAR(100)", "NOT NULL"],
    ["category", "VARCHAR(50)", "NOT NULL"],
    ["copies_available", "INT", "DEFAULT 1"],
]
add_table(s, Inches(0.75), Inches(2.35), Inches(3.55), Inches(4.2), books,
          col_widths=[1.3, 1.0, 1.25], size=10, header_size=11)

# Table: borrowers
add_rect(s, Inches(4.65), Inches(1.7), Inches(3.85), Inches(5.0), SLATE)
add_rect(s, Inches(4.65), Inches(1.7), Inches(3.85), Inches(0.5), NAVY)
add_text(s, Inches(4.8), Inches(1.78), Inches(3.5), Inches(0.4),
         [[("TABLE: borrowers", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
borr = [
    ["Column", "Type", "Constraint"],
    ["borrower_id", "INT", "PK, AUTO_INCREMENT"],
    ["borrower_name", "VARCHAR(100)", "NOT NULL"],
    ["roll_no", "VARCHAR(50)", "NOT NULL, UNIQUE"],
]
add_table(s, Inches(4.8), Inches(2.35), Inches(3.55), Inches(2.6), borr,
          col_widths=[1.5, 1.0, 1.05], size=10, header_size=11)

# Table: loans
add_rect(s, Inches(8.7), Inches(1.7), Inches(3.95), Inches(5.0), SLATE)
add_rect(s, Inches(8.7), Inches(1.7), Inches(3.95), Inches(0.5), NAVY)
add_text(s, Inches(8.85), Inches(1.78), Inches(3.6), Inches(0.4),
         [[("TABLE: loans", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
loans = [
    ["Column", "Type", "Constraint"],
    ["loan_id", "INT", "PK, AUTO_INCREMENT"],
    ["book_id", "INT", "FK \u2192 books(book_id)"],
    ["borrower_id", "INT", "FK \u2192 borrowers(borrower_id)"],
    ["issue_date", "DATE", "NOT NULL"],
    ["due_date", "DATE", "NOT NULL"],
    ["return_date", "DATE", "NULL (active loan)"],
]
add_table(s, Inches(8.85), Inches(2.35), Inches(3.65), Inches(4.2), loans,
          col_widths=[1.4, 1.0, 1.25], size=10, header_size=11)
add_text(s, Inches(0.6), Inches(6.55), Inches(12.1), Inches(0.5),
         [[("Relationships:  ", {"size": 13, "bold": True, "color": NAVY}),
           ("loans.book_id \u2192 books.book_id (CASCADE)  |  loans.borrower_id \u2192 borrowers.borrower_id (CASCADE)   Engine: InnoDB   Charset: utf8mb4", {"size": 13, "color": GREY})]], space_after=0)
add_footer(s, 5)

# ============================================================
# SLIDE 6: FEATURE 1 - DASHBOARD
# ============================================================
s = add_slide()
add_title(s, "Feature 1 \u2014 Dashboard", "Real-Time KPIs & Active Loans")
add_rect(s, Inches(0.6), Inches(1.7), Inches(6.0), Inches(5.1), SLATE)
add_rect(s, Inches(0.6), Inches(1.7), Inches(6.0), Inches(0.5), NAVY)
add_text(s, Inches(0.8), Inches(1.78), Inches(5.6), Inches(0.4),
         [[("WHAT IT SHOWS", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
add_bullets(s, Inches(0.8), Inches(2.4), Inches(5.6), Inches(4.3), [
    ("KPI Tiles: Total Titles, Copies In Stock, Active Loans, Overdue Loans, Registered Members.", 0),
    ("Active Loans Priority Queue with overdue badges (\U0001f534 Overdue / \U0001f7e2 Active).", 0),
    ("Live Chart.js bar chart: Top Borrowed Books.", 0),
    ("Live Chart.js line chart: Monthly Circulation Trajectory.", 0),
], size=13, color=LIGHT)

add_rect(s, Inches(6.8), Inches(1.7), Inches(5.95), Inches(5.1), SLATE)
add_rect(s, Inches(6.8), Inches(1.7), Inches(5.95), Inches(0.5), NAVY)
add_text(s, Inches(7.0), Inches(1.78), Inches(5.5), Inches(0.4),
         [[("HOW IT'S IMPLEMENTED", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
add_bullets(s, Inches(7.0), Inches(2.4), Inches(5.6), Inches(4.3), [
    ("GET /api/kpis \u2192 analytics.get_library_kpis()", 0),
    ("GET /api/loans/active \u2192 operations.get_active_loans()", 0),
    ("GET /api/analytics/popular & /trends \u2192 chart data", 0),
    ("Overdue = today > due_date (calculated in Python)", 0),
    ("Chart.js renders on HTML5 Canvas \u2014 no static images", 0),
    ("SPA hash routing: view switches without page reload", 0),
], size=13, color=LIGHT)
add_footer(s, 6)

# ============================================================
# SLIDE 7: FEATURE 2 - BOOK CATALOGUE
# ============================================================
s = add_slide()
add_title(s, "Feature 2 \u2014 Book Catalogue", "Search, Filter, and Browse Library Titles")
add_rect(s, Inches(0.6), Inches(1.7), Inches(6.0), Inches(5.1), SLATE)
add_rect(s, Inches(0.6), Inches(1.7), Inches(6.0), Inches(0.5), NAVY)
add_text(s, Inches(0.8), Inches(1.78), Inches(5.6), Inches(0.4),
         [[("WHAT IT DOES", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
add_bullets(s, Inches(0.8), Inches(2.4), Inches(5.6), Inches(4.3), [
    ("Lists all books: ISBN, title, author, category, available copies.", 0),
    ("Instant live search by title, author, or ISBN.", 0),
    ("Category filter dropdown (auto-populated from DB).", 0),
    ("Stock pills: \u25cf In Stock (green) / \u25cf Out of Stock (red).", 0),
    ("Quick Issue shortcut \u2192 pre-selects book in Circulation Desk.", 0),
], size=13, color=LIGHT)

add_rect(s, Inches(6.8), Inches(1.7), Inches(5.95), Inches(5.1), SLATE)
add_rect(s, Inches(6.8), Inches(1.7), Inches(5.95), Inches(0.5), NAVY)
add_text(s, Inches(7.0), Inches(1.78), Inches(5.5), Inches(0.4),
         [[("HOW IT'S IMPLEMENTED", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
add_bullets(s, Inches(7.0), Inches(2.4), Inches(5.6), Inches(4.3), [
    ("GET /api/books \u2192 SELECT * FROM books ORDER BY book_id", 0),
    ("Frontend filters locally (filterBooks()) for instant response.", 0),
    ("Category list built dynamically from book records.", 0),
    ("quickIssueBook(id) switches view and auto-selects the book.", 0),
    ("Built with operations.get_all_books() in backend.", 0),
], size=13, color=LIGHT)
add_footer(s, 7)

# ============================================================
# SLIDE 8: FEATURE 3 - ADD NEW BOOK
# ============================================================
s = add_slide()
add_title(s, "Feature 3 \u2014 Add New Book", "Dedicated Registration with Live Preview")
add_rect(s, Inches(0.6), Inches(1.7), Inches(6.0), Inches(5.1), SLATE)
add_rect(s, Inches(0.6), Inches(1.7), Inches(6.0), Inches(0.5), NAVY)
add_text(s, Inches(0.8), Inches(1.78), Inches(5.6), Inches(0.4),
         [[("WHAT IT DOES", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
add_bullets(s, Inches(0.8), Inches(2.4), Inches(5.6), Inches(4.3), [
    ("Form: ISBN, Title, Author, Category, Available Copies.", 0),
    ("Category autocomplete datalist (CS, Fiction, Science, Math\u2026).", 0),
    ("Live Shelf Preview Card updates as you type.", 0),
    ("Instant save to database and catalogue refresh.", 0),
    ("Duplicate ISBN is rejected with a clear error message.", 0),
], size=13, color=LIGHT)

add_rect(s, Inches(6.8), Inches(1.7), Inches(5.95), Inches(5.1), SLATE)
add_rect(s, Inches(6.8), Inches(1.7), Inches(5.95), Inches(0.5), NAVY)
add_text(s, Inches(7.0), Inches(1.78), Inches(5.5), Inches(0.4),
         [[("HOW IT'S IMPLEMENTED", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
add_bullets(s, Inches(7.0), Inches(2.4), Inches(5.6), Inches(4.3), [
    ("updateLiveBookPreview() reads inputs on every keystroke.", 0),
    ("POST /api/books \u2192 JSON body \u2192 operations.add_book().", 0),
    ("Backend validates non-empty fields + unique ISBN constraint.", 0),
    ("SQL: INSERT INTO books (isbn, title, author, category, copies_available)", 0),
    ("On success: toast notification + loadBooks() refreshes catalogue.", 0),
], size=13, color=LIGHT)
add_footer(s, 8)

# ============================================================
# SLIDE 9: FEATURE 4 - BORROWERS
# ============================================================
s = add_slide()
add_title(s, "Feature 4 \u2014 Borrowers Directory", "Member Registration & Search")
add_rect(s, Inches(0.6), Inches(1.7), Inches(6.0), Inches(5.1), SLATE)
add_rect(s, Inches(0.6), Inches(1.7), Inches(6.0), Inches(0.5), NAVY)
add_text(s, Inches(0.8), Inches(1.78), Inches(5.6), Inches(0.4),
         [[("WHAT IT DOES", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
add_bullets(s, Inches(0.8), Inches(2.4), Inches(5.6), Inches(4.3), [
    ("Searchable roster by borrower name or roll number.", 0),
    ("Register new borrower via modal (name + roll no).", 0),
    ("Direct checkout: \u201cIssue Book To Member \u2192\u201d pre-selects member.", 0),
    ("Unique roll number enforcement (no duplicates).", 0),
], size=13, color=LIGHT)

add_rect(s, Inches(6.8), Inches(1.7), Inches(5.95), Inches(5.1), SLATE)
add_rect(s, Inches(6.8), Inches(1.7), Inches(5.95), Inches(0.5), NAVY)
add_text(s, Inches(7.0), Inches(1.78), Inches(5.5), Inches(0.4),
         [[("HOW IT'S IMPLEMENTED", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
add_bullets(s, Inches(7.0), Inches(2.4), Inches(5.6), Inches(4.3), [
    ("GET /api/borrowers \u2192 SELECT borrower_id, name, roll_no FROM borrowers", 0),
    ("POST /api/borrowers \u2192 INSERT INTO borrowers (borrower_name, roll_no)", 0),
    ("Duplicate roll_no caught via MySQL unique constraint.", 0),
    ("filterBorrowers() runs client-side for instant search.", 0),
    ("quickIssueToBorrower(id) auto-fills borrower in issue form.", 0),
], size=13, color=LIGHT)
add_footer(s, 9)

# ============================================================
# SLIDE 10: FEATURE 5 - CIRCULATION DESK
# ============================================================
s = add_slide()
add_title(s, "Feature 5 \u2014 Circulation Desk", "Issue Books & Process Returns")
add_rect(s, Inches(0.6), Inches(1.7), Inches(6.0), Inches(5.1), SLATE)
add_rect(s, Inches(0.6), Inches(1.7), Inches(6.0), Inches(0.5), NAVY)
add_text(s, Inches(0.8), Inches(1.78), Inches(5.6), Inches(0.4),
         [[("ISSUE BOOK", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
add_bullets(s, Inches(0.8), Inches(2.35), Inches(5.6), Inches(4.4), [
    ("Dropdown shows only in-stock books (copies_available > 0).", 0),
    ("Borrower picker, issue date, configurable loan period (default 14 days).", 0),
    ("Atomic transaction: insert loan + decrement copies.", 0),
    ("Rollback if stock insufficient during checkout.", 0),
], size=13, color=LIGHT)

add_rect(s, Inches(6.8), Inches(1.7), Inches(5.95), Inches(5.1), SLATE)
add_rect(s, Inches(6.8), Inches(1.7), Inches(5.95), Inches(0.5), NAVY)
add_text(s, Inches(7.0), Inches(1.78), Inches(5.5), Inches(0.4),
         [[("RETURN BOOK", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
add_bullets(s, Inches(7.0), Inches(2.35), Inches(5.6), Inches(4.4), [
    ("Active loans table with overdue alerts and badges.", 0),
    ("1-click Return Check-in modal with date validation.", 0),
    ("Rejects: missing loan, already returned, return < issue date.", 0),
    ("Increments copies_available; flags late returns.", 0),
], size=13, color=LIGHT)
add_footer(s, 10)

# ============================================================
# SLIDE 11: FEATURE 6 - ANALYTICS
# ============================================================
s = add_slide()
add_title(s, "Feature 6 \u2014 Analytics & Charts", "7 Interactive Chart.js Visuals + Pandas Tables")
data = [
    ["#", "Chart", "Type", "Logic"],
    ["1", "Top Most-Borrowed Books", "Horizontal Bar", "SQL GROUP BY + COUNT, Pandas"],
    ["2", "Monthly Circulation Trajectory", "Multi-line Area", "Pandas groupby Year-Month"],
    ["3", "Circulation Share by Genre", "Doughnut", "Pandas category share %"],
    ["4", "Loan Duration Distribution", "Histogram Bar", "NumPy bins: 1-7d, 8-14d, 15-21d, 22+d"],
    ["5", "Most Active Borrowers", "Bar", "SQL COUNT loans per borrower"],
    ["6", "Inventory Allocation", "Doughnut", "Available vs on-loan copies"],
    ["7", "Return Punctuality", "Doughnut", "On-time vs late vs active vs overdue"],
]
add_table(s, Inches(0.6), Inches(1.7), Inches(12.1), Inches(3.6), data,
          col_widths=[0.5, 3.2, 1.8, 6.6], size=11.5, header_size=12)
add_text(s, Inches(0.6), Inches(5.5), Inches(12.1), Inches(1.6),
         [[("NumPy Statistics (CLI mode):  ", {"size": 13, "bold": True, "color": NAVY}),
           ("Mean, median, standard deviation, min, max, Q1/Q3, IQR of completed loan durations.", {"size": 13, "color": GREY})], [
               ("Pandas Data Tables:  ", {"size": 13, "bold": True, "color": NAVY}),
           ("Top Borrowed Books leaderboard, Category circulation summary, Monthly loan history.", {"size": 13, "color": GREY})],
         ], space_after=6, line_spacing=1.1)
add_footer(s, 11)

# ============================================================
# SLIDE 12: REST API
# ============================================================
s = add_slide()
add_title(s, "REST API Reference", "All endpoints return JSON with consistent error format")
data = [
    ["Method", "Endpoint", "Description"],
    ["GET", "/api/status", "Health check & DB backend"],
    ["GET", "/api/kpis", "Headline circulation & stock KPIs"],
    ["GET", "/api/books", "List all books"],
    ["POST", "/api/books", "Add new book"],
    ["GET", "/api/borrowers", "List all borrowers"],
    ["POST", "/api/borrowers", "Register borrower"],
    ["GET", "/api/loans/active", "Active loans with overdue status"],
    ["POST", "/api/loans/issue", "Issue a book to a borrower"],
    ["POST", "/api/loans/return", "Record a book return"],
    ["GET", "/api/analytics/popular", "Most-borrowed books leaderboard"],
    ["GET", "/api/analytics/categories", "Category circulation summary"],
    ["GET", "/api/analytics/trends", "Monthly borrowing trajectory"],
    ["GET", "/api/analytics/statistics", "NumPy loan duration statistics"],
    ["GET", "/api/analytics/punctuality", "Return timeliness breakdown"],
    ["POST", "/api/seed", "Reset & reload sample data"],
]
add_table(s, Inches(0.6), Inches(1.7), Inches(12.1), Inches(5.0), data,
          col_widths=[1.0, 3.6, 7.5], size=11, header_size=12)
add_footer(s, 12)

# ============================================================
# SLIDE 13: CLI MODE
# ============================================================
s = add_slide()
add_title(s, "CLI Mode \u2014 Terminal Interface", "Run: python main.py  (or python main.py --demo)")
add_rect(s, Inches(0.6), Inches(1.7), Inches(6.0), Inches(5.1), SLATE)
add_rect(s, Inches(0.6), Inches(1.7), Inches(6.0), Inches(0.5), NAVY)
add_text(s, Inches(0.8), Inches(1.78), Inches(5.6), Inches(0.4),
         [[("MENU OPTIONS", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
add_bullets(s, Inches(0.8), Inches(2.4), Inches(5.6), Inches(4.3), [
    ("1. View Book Catalogue   2. Add New Book", 0),
    ("3. View Borrowers   4. Register Borrower", 0),
    ("5. Issue Book   6. Return Book", 0),
    ("7. View Active Loans   8. Most-Borrowed Report", 0),
    ("9. Overdue Loans   10. Category Summary (Pandas)", 0),
    ("11. Monthly Trends (Pandas)   12. Loan Stats (NumPy)", 0),
    ("13. Generate Charts (Matplotlib)   14. Seed Data   0. Exit", 0),
], size=12.5, color=LIGHT, space_after=5)

add_rect(s, Inches(6.8), Inches(1.7), Inches(5.95), Inches(5.1), SLATE)
add_rect(s, Inches(6.8), Inches(1.7), Inches(5.95), Inches(0.5), NAVY)
add_text(s, Inches(7.0), Inches(1.78), Inches(5.5), Inches(0.4),
         [[("AUTOMATED DEMO", {"size": 15, "bold": True, "color": WHITE})]], space_after=0)
add_bullets(s, Inches(7.0), Inches(2.4), Inches(5.6), Inches(4.3), [
    ("--demo flag runs a complete walkthrough:", 0),
    ("1. Seed sample data (12 books, 7 borrowers, 20 loans)", 1),
    ("2. Display catalogue & active loans", 1),
    ("3. Test loan return + duplicate rejection", 1),
    ("4. Show all Pandas reports & NumPy stats", 1),
    ("5. Generate Matplotlib charts", 1),
    ("Ideal for project viva/demo day.", 0, True),
], size=12.5, color=LIGHT, space_after=5)
add_footer(s, 13)

# ============================================================
# SLIDE 14: SAMPLE DATA
# ============================================================
s = add_slide()
add_title(s, "Sample Data & Seed Function", "Enter: python seed_data.py  or  POST /api/seed")
cards = [
    ("12 Books", "Computer Science, Fiction, Science, Mystery, Philosophy, History, Mathematics"),
    ("7 Borrowers", "Students with roll numbers (e.g., CS-2024-001)"),
    ("20 Loans", "Mix of returned on-time, returned late, active, and overdue"),
    ("Realistic Scenarios", "Overdue flags, duplicate return rejection, stock updates"),
]
x = Inches(0.6)
for title, desc in cards:
    add_rect(s, x, Inches(1.9), Inches(2.9), Inches(3.2), SLATE)
    add_rect(s, x, Inches(1.9), Inches(2.9), Inches(0.55), NAVY)
    add_text(s, x + Inches(0.15), Inches(1.98), Inches(2.6), Inches(0.45),
             [[(title, {"size": 16, "bold": True, "color": WHITE})]], space_after=0)
    add_text(s, x + Inches(0.15), Inches(2.65), Inches(2.6), Inches(2.3),
             [[(desc, {"size": 11.5, "color": LIGHT})]], space_after=0, line_spacing=1.1)
    x += Inches(3.05)
add_text(s, Inches(0.6), Inches(5.5), Inches(12.1), Inches(1.2),
         [[("Sample loans simulate real library activity: ", {"size": 13, "bold": True, "color": NAVY}),
           ("books issued weeks ago, some returned late, some still active with 11+ days overdue \u2014 so the dashboard and analytics have meaningful data to display immediately.", {"size": 13, "color": GREY})]], space_after=0, line_spacing=1.15)
add_footer(s, 14)

# ============================================================
# SLIDE 15: FILE STRUCTURE
# ============================================================
s = add_slide()
add_title(s, "Project File Structure", "Organized backend modules + frontend SPA assets")
structure = [
    "config.py            MySQL settings from .env + loan rules",
    "database.py          MySQL connection + schema initialization (3 tables)",
    "operations.py        CRUD + issue/return business logic + atomic transactions",
    "analytics.py         Pandas grouping, NumPy stats, Matplotlib charts",
    "seed_data.py         Sample data generator",
    "server.py            Flask app, REST routes, static file serving",
    "main.py              Terminal CLI menu + automated demo runner",
    "requirements.txt     Python dependencies",
    ".env.example         Template (no secrets committed)",
    ".gitignore           Excludes .env, __pycache__, charts",
    "README.md            Documentation",
    "static/index.html    6-view SPA: dashboard, books, addBook, borrowers, circulation, analytics",
    "static/css/style.css Glassmorphism design system",
    "static/js/app.js     Chart.js rendering, API fetch layer, routing, forms",
]
add_text(s, Inches(0.6), Inches(1.7), Inches(12.1), Inches(5.0),
         [[(line, {"size": 13.5, "color": NAVY if not line.startswith("static") and not line.startswith("requirements") and not line.startswith(".env") and not line.startswith(".gitignore") and not line.startswith("README") else GREY, "bold": not line.startswith("static") and not line.startswith("requirements") and not line.startswith(".env") and not line.startswith(".gitignore") and not line.startswith("README")})] for line in structure],
         space_after=7, line_spacing=1.05)
add_footer(s, 15)

# ============================================================
# SLIDE 16: HOW TO RUN
# ============================================================
s = add_slide()
add_title(s, "How to Run (Local Setup)", "Prerequisites: Python 3.10+ and MySQL Server 8.0+")
steps = [
    ("1", "Clone the repo", "git clone <repo-url> && cd library-book-management"),
    ("2", "Install dependencies", "pip install -r requirements.txt"),
    ("3", "Configure database", "Copy .env.example \u2192 .env; add MySQL host/user/password"),
    ("4", "Load sample data", "python seed_data.py"),
    ("5", "Start web server", "python server.py  \u2192  http://127.0.0.1:8000"),
    ("6", "Or run CLI", "python main.py"),
]
y = Inches(1.7)
for num, title, cmd in steps:
    add_rect(s, Inches(0.6), y, Inches(0.5), Inches(0.55), BLUE)
    add_text(s, Inches(0.6), y + Inches(0.08), Inches(0.5), Inches(0.4),
             [[(num, {"size": 16, "bold": True, "color": WHITE})]], align=PP_ALIGN.CENTER, space_after=0)
    add_text(s, Inches(1.25), y, Inches(2.6), Inches(0.55),
             [[(title, {"size": 14, "bold": True, "color": NAVY})]], space_after=0)
    add_text(s, Inches(4.0), y + Inches(0.05), Inches(8.6), Inches(0.5),
             [[(cmd, {"size": 12, "color": GREY})]], space_after=0)
    y += Inches(0.75)
add_text(s, Inches(0.6), Inches(6.4), Inches(12.1), Inches(0.6),
         [[("Note:  ", {"size": 12, "bold": True, "color": ROSE}),
           (".env is git-ignored \u2014 never commit real passwords. Tables are created automatically on first run.", {"size": 12, "color": GREY})]], space_after=0)
add_footer(s, 16)

# ============================================================
# SLIDE 17: IMPLEMENTATION HIGHLIGHTS
# ============================================================
s = add_slide()
add_title(s, "Key Implementation Highlights", "Engineering decisions behind the system")
data = [
    ["Feature", "Implementation"],
    ["Atomic Transactions", "Issue/return use INSERT + UPDATE in one transaction with ROLLBACK on failure"],
    ["Overdue Detection", "Compares date.today() vs due_date; flagged with \U0001f534 / \U0001f7e2 badges"],
    ["No Static Images", "All charts are Chart.js HTML5 Canvas, fetched live from API"],
    ["SPA Routing", "Hash-based client-side routing (#dashboard, #books\u2026) without reloads"],
    ["Security", "Secrets in .env (git-ignored); .env.example committed as safe template"],
    ["Auto Schema", "Tables created automatically on startup if missing"],
    ["Input Validation", "Book/borrower existence, stock checks, date order, duplicate ISBN/roll_no"],
]
add_table(s, Inches(0.6), Inches(1.7), Inches(12.1), Inches(4.8), data,
          col_widths=[2.6, 9.5], size=11.5, header_size=13)
add_footer(s, 17)

# ============================================================
# SLIDE 18: FUTURE ENHANCEMENTS
# ============================================================
s = add_slide()
add_title(s, "Future Enhancements", "Roadmap for scaling the project")
items = [
    ("User Authentication", "Role-based access: Librarian vs Member login"),
    ("Overdue Notifications", "Email/SMS alerts for overdue loans"),
    ("Book Reservations", "Hold queue for out-of-stock titles"),
    ("Fine Calculation", "Auto-compute late-return penalties"),
    ("Barcode/ISBN Lookup", "Auto-fetch book details via external API"),
    ("Report Export", "Export analytics to PDF / Excel"),
    ("Dockerization", "One-command containerized deployment"),
    ("Cloud Deployment", "Render / Railway with external MySQL (Aiven, PlanetScale)"),
]
add_bullets(s, Inches(0.6), Inches(1.8), Inches(12.1), Inches(5.0),
            [(f"{t} \u2014 {d}", 0, True) for t, d in items], size=15, space_after=12)
add_footer(s, 18)

prs.save("E:\\Library Book Management\\Library_Book_Management_Project_Flask.pptx")
print("PPTX saved successfully.")