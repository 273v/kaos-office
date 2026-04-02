"""OOXML namespace constants and Clark notation helpers.

Pre-computed qualified names for O(1) element tag comparison with lxml.
Covers WordprocessingML, DrawingML, and OPC namespaces.
"""

# --- Namespace URIs (Transitional — what all real-world documents use) ---

# WordprocessingML
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

# Office Document Relationships
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# DrawingML
A = "http://schemas.openxmlformats.org/drawingml/2006/main"

# DrawingML WordprocessingDrawing
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"

# DrawingML Picture
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

# Markup Compatibility
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

# Content Types
CT = "http://schemas.openxmlformats.org/package/2006/content-types"

# Package Relationships
PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

# Dublin Core (for docProps/core.xml)
DC = "http://purl.org/dc/elements/1.1/"
DCTERMS = "http://purl.org/dc/terms/"
DCMITYPE = "http://purl.org/dc/dcmitype/"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"

# Extended Properties (for docProps/app.xml)
EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

# --- Strict namespace variants (ISO 29500) ---
W_STRICT = "http://purl.oclc.org/ooxml/wordprocessingml/main"
R_STRICT = "http://purl.oclc.org/ooxml/officeDocument/relationships"


def qn(ns: str, tag: str) -> str:
    """Create a Clark notation qualified name: {namespace}tag."""
    return f"{{{ns}}}{tag}"


# --- Pre-computed WordprocessingML tags ---
W_BODY = qn(W, "body")
W_P = qn(W, "p")
W_R = qn(W, "r")
W_T = qn(W, "t")
W_TBL = qn(W, "tbl")
W_TR = qn(W, "tr")
W_TC = qn(W, "tc")
W_PPR = qn(W, "pPr")
W_RPR = qn(W, "rPr")
W_PSTYLE = qn(W, "pStyle")
W_RSTYLE = qn(W, "rStyle")
W_SECTPR = qn(W, "sectPr")
W_SDT = qn(W, "sdt")
W_SDTCONTENT = qn(W, "sdtContent")
W_HYPERLINK = qn(W, "hyperlink")
W_BOOKMARK_START = qn(W, "bookmarkStart")
W_BOOKMARK_END = qn(W, "bookmarkEnd")
W_INS = qn(W, "ins")
W_DEL = qn(W, "del")
W_DEL_TEXT = qn(W, "delText")
W_MOVE_FROM = qn(W, "moveFrom")
W_MOVE_TO = qn(W, "moveTo")
W_COMMENT_RANGE_START = qn(W, "commentRangeStart")
W_COMMENT_RANGE_END = qn(W, "commentRangeEnd")
W_COMMENT_REFERENCE = qn(W, "commentReference")
W_FOOTNOTE_REFERENCE = qn(W, "footnoteReference")
W_ENDNOTE_REFERENCE = qn(W, "endnoteReference")
W_TAB = qn(W, "tab")
W_BR = qn(W, "br")
W_DRAWING = qn(W, "drawing")
W_NUMPR = qn(W, "numPr")
W_NUMID = qn(W, "numId")
W_ILVL = qn(W, "ilvl")
W_VAL = qn(W, "val")
W_B = qn(W, "b")
W_I = qn(W, "i")
W_U = qn(W, "u")
W_STRIKE = qn(W, "strike")
W_VERTALING = qn(W, "vertAlign")
W_STYLE = qn(W, "style")
W_STYLEID = qn(W, "styleId")
W_NAME = qn(W, "name")
W_BASED_ON = qn(W, "basedOn")
W_OUTLINE_LVL = qn(W, "outlineLvl")
W_ABSTRACT_NUM = qn(W, "abstractNum")
W_ABSTRACT_NUM_ID = qn(W, "abstractNumId")
W_NUM = qn(W, "num")
W_LVL = qn(W, "lvl")
W_NUM_FMT = qn(W, "numFmt")
W_DOCUMENT = qn(W, "document")
W_FOOTNOTES = qn(W, "footnotes")
W_FOOTNOTE = qn(W, "footnote")
W_ENDNOTES = qn(W, "endnotes")
W_ENDNOTE = qn(W, "endnote")
W_COMMENTS = qn(W, "comments")
W_COMMENT = qn(W, "comment")
W_TCPR = qn(W, "tcPr")
W_GRIDSPAN = qn(W, "gridSpan")
W_VMERGE = qn(W, "vMerge")
W_TBLGRID = qn(W, "tblGrid")
W_GRIDCOL = qn(W, "gridCol")
W_TYPE = qn(W, "type")
W_ID = qn(W, "id")
W_AUTHOR = qn(W, "author")
W_DATE = qn(W, "date")
W_ANCHOR = qn(W, "anchor")

# DrawingML tags
A_BLIP = qn(A, "blip")
A_GRAPHIC = qn(A, "graphic")
A_GRAPHIC_DATA = qn(A, "graphicData")

# WordprocessingDrawing tags
WP_DOCPR = qn(WP, "docPr")
WP_INLINE = qn(WP, "inline")
WP_ANCHOR = qn(WP, "anchor")
WP_EXTENT = qn(WP, "extent")

# Picture tags
PIC_PIC = qn(PIC, "pic")
PIC_BLIP_FILL = qn(PIC, "blipFill")

# Relationship attributes use the R namespace
R_EMBED = qn(R, "embed")
R_LINK = qn(R, "link")
R_ID = qn(R, "id")

# --- Relationship Type URIs ---
RT_OFFICE_DOCUMENT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)
RT_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
RT_NUMBERING = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
RT_FONT_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
RT_SETTINGS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
RT_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
RT_COMMENTS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
RT_FOOTNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
RT_ENDNOTES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
RT_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
RT_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
RT_HEADER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
RT_FOOTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
RT_CORE_PROPERTIES = (
    "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
)
RT_EXTENDED_PROPERTIES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
)

# --- MIME Types ---
DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PPTX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

# --- EMU Conversion ---
EMU_PER_INCH = 914400
EMU_PER_PIXEL_96DPI = 9525  # EMU_PER_INCH / 96


def emu_to_px(emu: int, dpi: int = 96) -> float:
    """Convert English Metric Units to pixels at given DPI."""
    return emu / (EMU_PER_INCH / dpi)
