"""OOXML namespace constants and Clark notation helpers.

Pre-computed qualified names for O(1) element tag comparison with lxml.
Covers WordprocessingML, PresentationML, DrawingML, and OPC namespaces.
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

# PresentationML
P = "http://schemas.openxmlformats.org/presentationml/2006/main"

# DrawingML Chart
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"

# DrawingML Diagram (SmartArt)
DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"

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
W_LVL_TEXT = qn(W, "lvlText")
W_START = qn(W, "start")
W_LVL_RESTART = qn(W, "lvlRestart")
W_LVL_OVERRIDE = qn(W, "lvlOverride")
W_START_OVERRIDE = qn(W, "startOverride")
W_IS_LGL = qn(W, "isLgl")
W_SUFF = qn(W, "suff")
W_LVL_JC = qn(W, "lvlJc")
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

# Header / footer / section-properties elements (Phase 4)
W_HDR = qn(W, "hdr")
W_FTR = qn(W, "ftr")
W_HEADER_REFERENCE = qn(W, "headerReference")
W_FOOTER_REFERENCE = qn(W, "footerReference")
W_PGSZ = qn(W, "pgSz")
W_PGMAR = qn(W, "pgMar")
W_TITLEPG = qn(W, "titlePg")
W_EVEN_AND_ODD_HEADERS = qn(W, "evenAndOddHeaders")
W_SETTINGS = qn(W, "settings")
R_ID_ATTR = qn(R, "id")

# Content types for header / footer / settings parts (relationship types
# RT_HEADER / RT_FOOTER / RT_SETTINGS are already defined below).
CT_HEADER = "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
CT_FOOTER = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
CT_SETTINGS = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"

# Twip conversion (1 twip = 1/20 point = 1/1440 inch)
TWIPS_PER_POINT = 20


def twips_to_pt(twips: int | float) -> float:
    """Convert OOXML twips (1/20 pt) to typographic points."""
    return float(twips) / TWIPS_PER_POINT


def pt_to_twips(points: int | float) -> int:
    """Convert typographic points to OOXML twips, rounded to int."""
    return round(float(points) * TWIPS_PER_POINT)


# --- Pre-computed PresentationML tags ---
P_SLD = qn(P, "sld")
P_CSLD = qn(P, "cSld")
P_SP_TREE = qn(P, "spTree")
P_SP = qn(P, "sp")
P_PIC = qn(P, "pic")
P_GRAPHIC_FRAME = qn(P, "graphicFrame")
P_GRP_SP = qn(P, "grpSp")
P_CXN_SP = qn(P, "cxnSp")
P_NV_SP_PR = qn(P, "nvSpPr")
P_NV_PIC_PR = qn(P, "nvPicPr")
P_NV_GRP_SP_PR = qn(P, "nvGrpSpPr")
P_NV_GRAPHIC_FRAME_PR = qn(P, "nvGraphicFramePr")
P_CNV_PR = qn(P, "cNvPr")
P_NV_PR = qn(P, "nvPr")
P_PH = qn(P, "ph")
P_TX_BODY = qn(P, "txBody")
P_SP_PR = qn(P, "spPr")
P_GRP_SP_PR = qn(P, "grpSpPr")
P_BLIP_FILL = qn(P, "blipFill")
P_XFRM = qn(P, "xfrm")
P_NOTES = qn(P, "notes")

# --- Pre-computed DrawingML tags (extended for PPTX) ---
A_BLIP = qn(A, "blip")
A_GRAPHIC = qn(A, "graphic")
A_GRAPHIC_DATA = qn(A, "graphicData")
A_P = qn(A, "p")
A_R = qn(A, "r")
A_T = qn(A, "t")
A_RPR = qn(A, "rPr")
A_PPR = qn(A, "pPr")
A_BR = qn(A, "br")
A_BODY_PR = qn(A, "bodyPr")
A_LST_STYLE = qn(A, "lstStyle")
A_TX_BODY = qn(A, "txBody")
A_TBL = qn(A, "tbl")
A_TBL_GRID = qn(A, "tblGrid")
A_GRID_COL = qn(A, "gridCol")
A_TR = qn(A, "tr")
A_TC = qn(A, "tc")
A_XFRM = qn(A, "xfrm")
A_OFF = qn(A, "off")
A_EXT = qn(A, "ext")
A_CH_OFF = qn(A, "chOff")
A_CH_EXT = qn(A, "chExt")
A_BU_CHAR = qn(A, "buChar")
A_BU_AUTO_NUM = qn(A, "buAutoNum")
A_BU_NONE = qn(A, "buNone")
A_HLINKCLICK = qn(A, "hlinkClick")

# --- Pre-computed Chart tags ---
C_CHART_SPACE = qn(C, "chartSpace")
C_CHART = qn(C, "chart")
C_TITLE = qn(C, "title")
C_TX = qn(C, "tx")
C_RICH = qn(C, "rich")
C_PLOT_AREA = qn(C, "plotArea")
C_SER = qn(C, "ser")
C_CAT = qn(C, "cat")
C_VAL = qn(C, "val")
C_STR_REF = qn(C, "strRef")
C_NUM_REF = qn(C, "numRef")
C_STR_CACHE = qn(C, "strCache")
C_NUM_CACHE = qn(C, "numCache")
C_PT = qn(C, "pt")
C_V = qn(C, "v")

# Chart type elements
C_BAR_CHART = qn(C, "barChart")
C_LINE_CHART = qn(C, "lineChart")
C_PIE_CHART = qn(C, "pieChart")
C_SCATTER_CHART = qn(C, "scatterChart")
C_AREA_CHART = qn(C, "areaChart")
C_RADAR_CHART = qn(C, "radarChart")
C_DOUGHNUT_CHART = qn(C, "doughnutChart")
C_BAR_3D_CHART = qn(C, "bar3DChart")
C_LINE_3D_CHART = qn(C, "line3DChart")
C_PIE_3D_CHART = qn(C, "pie3DChart")
C_AREA_3D_CHART = qn(C, "area3DChart")

CHART_TYPE_TAGS = frozenset(
    {
        C_BAR_CHART,
        C_LINE_CHART,
        C_PIE_CHART,
        C_SCATTER_CHART,
        C_AREA_CHART,
        C_RADAR_CHART,
        C_DOUGHNUT_CHART,
        C_BAR_3D_CHART,
        C_LINE_3D_CHART,
        C_PIE_3D_CHART,
        C_AREA_3D_CHART,
    }
)

# --- Pre-computed Diagram (SmartArt) tags ---
DGM_DATA_MODEL = qn(DGM, "dataModel")
DGM_PT_LST = qn(DGM, "ptLst")
DGM_PT = qn(DGM, "pt")
DGM_T = qn(DGM, "t")
DGM_REL_IDS = qn(DGM, "relIds")

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

# PresentationML relationship types
RT_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
RT_SLIDE_LAYOUT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
RT_SLIDE_MASTER = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
RT_NOTES_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
RT_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
RT_DIAGRAM_DATA = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"

# graphicData URI values (shape type identification)
GD_TABLE = "http://schemas.openxmlformats.org/drawingml/2006/table"
GD_CHART = "http://schemas.openxmlformats.org/drawingml/2006/chart"
GD_DIAGRAM = "http://schemas.openxmlformats.org/drawingml/2006/diagram"

# --- Pre-computed SpreadsheetML tags ---
SML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

SML_WORKBOOK = qn(SML, "workbook")
SML_SHEETS = qn(SML, "sheets")
SML_SHEET = qn(SML, "sheet")
SML_WORKBOOK_PR = qn(SML, "workbookPr")
SML_WORKSHEET = qn(SML, "worksheet")
SML_SHEET_DATA = qn(SML, "sheetData")
SML_ROW = qn(SML, "row")
SML_CELL = qn(SML, "c")
SML_VALUE = qn(SML, "v")
SML_FORMULA = qn(SML, "f")
SML_SST = qn(SML, "sst")
SML_SI = qn(SML, "si")
SML_T = qn(SML, "t")
SML_R = qn(SML, "r")
SML_MERGE_CELLS = qn(SML, "mergeCells")
SML_MERGE_CELL = qn(SML, "mergeCell")
SML_NUM_FMTS = qn(SML, "numFmts")
SML_NUM_FMT = qn(SML, "numFmt")
SML_CELL_XFS = qn(SML, "cellXfs")
SML_XF = qn(SML, "xf")
SML_DIMENSION = qn(SML, "dimension")
SML_INLINE_STR = qn(SML, "is")

# SpreadsheetML relationship types
RT_WORKSHEET = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
RT_SHARED_STRINGS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"
)
RT_STYLES = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"

# --- MIME Types ---
DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PPTX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.presentationml.presentation"

# --- EMU Conversion ---
EMU_PER_INCH = 914400
EMU_PER_PIXEL_96DPI = 9525  # EMU_PER_INCH / 96
EMU_PER_POINT = 12700  # EMU_PER_INCH / 72


def emu_to_px(emu: int, dpi: int = 96) -> float:
    """Convert English Metric Units to pixels at given DPI."""
    return emu / (EMU_PER_INCH / dpi)


def emu_to_pt(emu: int | float) -> float:
    """Convert English Metric Units to typographic points (1 pt = 1/72 inch)."""
    return float(emu) / EMU_PER_POINT
