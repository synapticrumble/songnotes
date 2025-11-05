"""
Bansuri Music Document Reformatter
===================================
This script reformats a Word document containing Bansuri flute music notations.

Features:
- Table of Contents (centered, bold) at the top
- Extracts song titles after separator lines (=====****=====)
- Chronological numbering of titles in TOC
- Page break after every separator line
- Clickable links from TOC to each song
- "Back to Table of Contents" link after each song title
- All text in Georgia font for optimal HTML rendering
- Fully Word-compatible bookmarks and internal links

Key Fix: Stores paragraph OBJECTS (not indices) before modifying document structure,
ensuring bookmarks and links are placed in the correct locations.
"""

import re
from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml.shared import OxmlElement, qn

# ------------------------------------------------------------
# Helper Functions
# ------------------------------------------------------------
def sanitize_bookmark_name(name):
    """Ensure valid Word bookmark name."""
    sanitized = re.sub(r'[^a-zA-Z0-9_]', '_', name)
    if sanitized and not sanitized[0].isalpha():
        sanitized = 'BM_' + sanitized
    return sanitized[:40]

def add_bookmark(paragraph, name, bookmark_id):
    """Insert a bookmark around paragraph text and return safe name."""
    safe_name = sanitize_bookmark_name(name)
    p = paragraph._p

    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), str(bookmark_id))
    start.set(qn('w:name'), safe_name)
    p.insert(0, start)

    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), str(bookmark_id))
    p.append(end)
    return safe_name

def add_hyperlink(paragraph, bookmark_name, text):
    """Create an internal hyperlink to a bookmark."""
    safe_bookmark = sanitize_bookmark_name(bookmark_name)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('w:anchor'), safe_bookmark)

    r = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rPr.append(u)
    color = OxmlElement('w:color'); color.set(qn('w:val'), '0563C1'); rPr.append(color)
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Georgia')
    rFonts.set(qn('w:hAnsi'), 'Georgia')
    rPr.append(rFonts)

    r.append(rPr)
    t = OxmlElement('w:t'); t.text = text; r.append(t)
    hyperlink.append(r)
    paragraph._p.append(hyperlink)

def insert_paragraph_after(doc, paragraph, text=""):
    """Insert a paragraph after the given one."""
    idx = None
    for i, p in enumerate(doc.paragraphs):
        if p == paragraph:
            idx = i
            break
    if idx is None or idx + 1 >= len(doc.paragraphs):
        return doc.add_paragraph(text)
    else:
        return doc.paragraphs[idx + 1].insert_paragraph_before(text)

# ------------------------------------------------------------
# Main Script
# ------------------------------------------------------------
input_path = r"P:\ShareDownloads\BansuriMusic.docx"
output_path = r"P:\ShareDownloads\songnotes\songs_reformatted.docx"

doc = Document(input_path)
sep_pattern = re.compile(r'^[=*]+\*+[=*]+$')

titles = []  # Will store (paragraph_object, title_text) tuples
bookmark_id = 1

# Step 1: Detect all song titles immediately after separator
# IMPORTANT: Store paragraph OBJECTS, not indices
for i, para in enumerate(doc.paragraphs[:-1]):
    if sep_pattern.match(para.text.strip()):
        j = i + 1
        # Skip blank paragraphs to find the title
        while j < len(doc.paragraphs) and not doc.paragraphs[j].text.strip():
            j += 1
        if j < len(doc.paragraphs):
            title_para = doc.paragraphs[j]
            title_text = title_para.text.strip()
            if title_text:
                # Store the paragraph object itself, not the index
                titles.append((title_para, title_text))

# Step 2: Add bookmarks to title paragraphs BEFORE modifying document structure
for title_para, title_text in titles:
    add_bookmark(title_para, title_text, bookmark_id)
    bookmark_id += 1

# Step 3: Add page breaks after each separator
for para in doc.paragraphs:
    if sep_pattern.match(para.text.strip()):
        para.add_run().add_break(WD_BREAK.PAGE)

# Step 4: Apply Georgia font globally
for p in doc.paragraphs:
    for r in p.runs:
        r.font.name = 'Georgia'

# Step 5: Insert Table of Contents at top
toc_heading = doc.paragraphs[0].insert_paragraph_before("Table of Contents")
toc_heading.alignment = 1  # Center alignment
toc_run = toc_heading.runs[0] if toc_heading.runs else toc_heading.add_run("Table of Contents")
toc_run.font.name = 'Georgia'
toc_run.font.bold = True
toc_run.font.underline = True

toc_bookmark_name = add_bookmark(toc_heading, "TableOfContents", bookmark_id)
bookmark_id += 1

# Step 6: Insert all TOC links directly after heading
last_toc_para = toc_heading
for i, (title_para, title_text) in enumerate(titles, start=1):
    toc_line = insert_paragraph_after(doc, last_toc_para, f"{i}. ")
    add_hyperlink(toc_line, title_text, title_text)
    for r in toc_line.runs:
        r.font.name = 'Georgia'
    last_toc_para = toc_line

# Page break after TOC section
pb = insert_paragraph_after(doc, last_toc_para, "")
pb.add_run().add_break(WD_BREAK.PAGE)

# Step 5: Add bookmarks and “Back to TOC” link for each title
for title_para, title_text in titles:
    back_para = insert_paragraph_after(doc, title_para, "")
    add_hyperlink(back_para, toc_bookmark_name, "Back to Table of Contents")
    for r in back_para.runs:
        r.font.name = 'Georgia'

# Save
doc.save(output_path)
print(f"✅ Reformatted successfully: {output_path}")


# ======================**************************************  ====================


# import re
# from docx import Document
# from docx.enum.text import WD_BREAK
# from docx.oxml.shared import OxmlElement, qn
# from docx.enum.dml import MSO_THEME_COLOR_INDEX

# def add_bookmark(paragraph, name):
#     """Insert a Word bookmark around an empty run in the given paragraph."""
#     run = paragraph.add_run()                 # create a new run at paragraph end
#     r = run._r                               # get the XML <w:r> element
#     # Start bookmark
#     start = OxmlElement('w:bookmarkStart')
#     start.set(qn('w:id'), '0')
#     start.set(qn('w:name'), name)
#     r.append(start)
#     # End bookmark
#     end = OxmlElement('w:bookmarkEnd')
#     end.set(qn('w:id'), '0')
#     end.set(qn('w:name'), name)
#     r.append(end)

# def add_hyperlink(paragraph, bookmark_name, text):
#     """
#     Add a hyperlink to a bookmark in `paragraph`. The visible text is `text`.
#     """
#     # Create <w:hyperlink> element with w:anchor set to the bookmark name
#     hyperlink = OxmlElement('w:hyperlink')
#     hyperlink.set(qn('w:anchor'), bookmark_name)
#     # Create <w:r> and <w:rPr> for the linked text run
#     r = OxmlElement('w:r')
#     rPr = OxmlElement('w:rPr')
#     r.append(rPr)
#     # Set the run text
#     r.text = text
#     hyperlink.append(r)
#     # Append the hyperlink to the paragraph
#     new_run = paragraph.add_run()
#     new_run._r.append(hyperlink)
#     # Style the link (optional): font, color, underline
#     new_run.font.name = "Georgia"
#     new_run.font.color.theme_color = MSO_THEME_COLOR_INDEX.HYPERLINK
#     new_run.font.underline = True

# # Load the existing document
# doc = Document("P:\\ShareDownloads\\BansuriMusic.docx")
# sep_pattern = re.compile(r'^=+\*+=+$')  # pattern for lines like =====***====

# # 1. Find all titles after separators
# titles = []  # list of (paragraph_obj, title_text)
# paras = doc.paragraphs
# for i, para in enumerate(paras[:-1]):  # skip last to avoid index error
#     if sep_pattern.match(para.text.strip()):
#         # The next paragraph is assumed to be the song title
#         title_para = paras[i+1]
#         title_text = title_para.text.strip()
#         if title_text:
#             titles.append((title_para, title_text))

# # 2. Insert page breaks after each separator line
# for para in paras:
#     if sep_pattern.match(para.text.strip()):
#         br_run = para.add_run()
#         br_run.add_break(WD_BREAK.PAGE)   # new page after this paragraph:contentReference[oaicite:6]{index=6}

# # 3. Convert all text to Georgia font
# #    (You can set the 'Normal' style font or do each run manually)
# for para in doc.paragraphs:
#     for run in para.runs:
#         run.font.name = 'Georgia'        # set font to Georgia:contentReference[oaicite:7]{index=7}:contentReference[oaicite:8]{index=8}

# # 4. Insert Table of Contents at top (with clickable links)
# # Insert a heading at the very beginning
# first_para = doc.paragraphs[0]
# toc_para = first_para.insert_paragraph_before("Table of Contents", style='Heading 1')
# toc_para.style.font.name = 'Georgia'
# # Insert link entries in reverse order so they appear in correct order
# original_first = first_para  # after insert, original content moved to index 1
# for title_para, title_text in reversed(titles):
#     # Insert an empty paragraph for the link
#     link_para = original_first.insert_paragraph_before("")  
#     add_hyperlink(link_para, bookmark_name=title_text, text=title_text)
# # Add bookmarks at each title for the hyperlinks to target:contentReference[oaicite:9]{index=9}
# for title_para, title_text in titles:
#     add_bookmark(title_para, title_text)

# # Save the reformatted document
# doc.save("P:\\ShareDownloads\\songnotes\\songs_reformatted.docx")
# print("✅ Document reformatted and saved as songs_reformatted.docx")