from typing import Any, Dict, List, Optional, Tuple
import uuid

class SceneNode:
    """Base class for all scene graph nodes."""
    def __init__(
        self,
        bbox: Optional[Tuple[float, float, float, float]] = None, # (x0, y0, x1, y1)
        transform: Optional[Tuple[float, ...]] = None, # 6-tuple for affine transform
        z_order: int = 0,
        opacity: float = 1.0,
        clipping_path: Optional['VectorPath'] = None,
        parent: Optional['Group'] = None,
        confidence: float = 1.0,
        provenance: str = "unknown"
    ):
        self.id = str(uuid.uuid4())
        self.bbox = bbox
        self.transform = transform
        self.z_order = z_order
        self.opacity = opacity
        self.clipping_path = clipping_path
        self.parent = parent
        self.confidence = confidence
        self.provenance = provenance

class TextRun:
    """A run of text with uniform formatting."""
    def __init__(
        self,
        text: str,
        font_name: Optional[str] = None,
        font_size: Optional[float] = None,
        color: Optional[Tuple[int, int, int]] = None,
        bold: bool = False,
        italic: bool = False,
        underline: bool = False
    ):
        self.text = text
        self.font_name = font_name
        self.font_size = font_size
        self.color = color
        self.bold = bold
        self.italic = italic
        self.underline = underline

class TextParagraph:
    """A single paragraph of text containing multiple runs."""
    def __init__(self):
        self.runs: List[TextRun] = []
        self.alignment: str = "left"

class TextBlock(SceneNode):
    """A block of text, composed of paragraphs."""
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.paragraphs: List[TextParagraph] = []
        self.fill_color: Optional[Tuple[int, int, int]] = None
        self.line_color: Optional[Tuple[int, int, int]] = None

class Image(SceneNode):
    """A raster image."""
    def __init__(self, image_bytes: bytes, **kwargs):
        super().__init__(**kwargs)
        self.image_bytes = image_bytes

class VectorPath(SceneNode):
    """A vector drawing path consisting of segments."""
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.segments: List[Dict[str, Any]] = [] # e.g., [{'type': 'l', 'pts': [...]}]
        self.fill_color: Optional[Tuple[int, int, int]] = None
        self.stroke_color: Optional[Tuple[int, int, int]] = None
        self.stroke_width: float = 0.0

class Line(SceneNode):
    """A simple line segment."""
    def __init__(self, start: Tuple[float, float], end: Tuple[float, float], **kwargs):
        super().__init__(**kwargs)
        self.start = start
        self.end = end
        self.stroke_color: Optional[Tuple[int, int, int]] = None
        self.stroke_width: float = 0.0

class Rect(SceneNode):
    """A rectangle."""
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.fill_color: Optional[Tuple[int, int, int]] = None
        self.stroke_color: Optional[Tuple[int, int, int]] = None
        self.stroke_width: float = 0.0

class TableCell:
    """A single cell in a table."""
    def __init__(self, row: int, col: int, rowspan: int = 1, colspan: int = 1):
        self.row = row
        self.col = col
        self.rowspan = rowspan
        self.colspan = colspan
        self.text_block: Optional[TextBlock] = None
        self.fill_color: Optional[Tuple[int, int, int]] = None

class Table(SceneNode):
    """A table layout."""
    def __init__(self, rows: int, cols: int, **kwargs):
        super().__init__(**kwargs)
        self.rows = rows
        self.cols = cols
        self.cells: List[TableCell] = []

class Group(SceneNode):
    """A grouped collection of nodes."""
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.children: List[SceneNode] = []

class ChartCandidate(SceneNode):
    """A candidate region for a chart/diagram."""
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.elements: List[SceneNode] = [] # The underlying lines, text, etc.
        self.chart_type: str = "unknown"

class UnknownRegion(SceneNode):
    """An unclassified or fallback region."""
    def __init__(self, image_fallback: Optional[bytes] = None, **kwargs):
        super().__init__(**kwargs)
        self.image_fallback = image_fallback

class Page:
    """A single page containing the scene graph root."""
    def __init__(self, width: float, height: float, page_num: int):
        self.width = width
        self.height = height
        self.page_num = page_num
        self.root = Group()
        self.fallback_bytes: Optional[bytes] = None
