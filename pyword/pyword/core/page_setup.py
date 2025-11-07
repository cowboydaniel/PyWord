from dataclasses import dataclass
from enum import Enum, auto
from typing import Optional, Dict, Any
from PySide6.QtCore import QMarginsF

class PageOrientation(Enum):
    PORTRAIT = auto()
    LANDSCAPE = auto()

@dataclass
class PageMargins:
    """Represents page margins in points (1/72 of an inch)."""
    left: float = 72.0  # 1 inch
    right: float = 72.0
    top: float = 72.0
    bottom: float = 72.0
    header: float = 36.0  # 0.5 inch
    footer: float = 36.0
    gutter: float = 0.0
    
    def to_qmarginsf(self) -> QMarginsF:
        """Convert to QMarginsF (left, top, right, bottom)."""
        return QMarginsF(self.left, self.top, self.right, self.bottom)
    
    def to_dict(self) -> Dict[str, float]:
        """Convert to dictionary for serialization."""
        return {
            'left': self.left,
            'right': self.right,
            'top': self.top,
            'bottom': self.bottom,
            'header': self.header,
            'footer': self.footer,
            'gutter': self.gutter
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'PageMargins':
        """Create from dictionary."""
        return cls(
            left=float(data.get('left', 72.0)),
            right=float(data.get('right', 72.0)),
            top=float(data.get('top', 72.0)),
            bottom=float(data.get('bottom', 72.0)),
            header=float(data.get('header', 36.0)),
            footer=float(data.get('footer', 36.0)),
            gutter=float(data.get('gutter', 0.0))
        )

@dataclass
class PageSetup:
    """Represents page setup settings for a document."""
    paper_size: str = 'A4'  # Standard paper size (A4, Letter, etc.)
    orientation: PageOrientation = PageOrientation.PORTRAIT
    margins: PageMargins = PageMargins()
    page_width: float = 595.0  # A4 width in points (210mm)
    page_height: float = 842.0  # A4 height in points (297mm)
    
    def get_effective_page_size(self) -> tuple[float, float]:
        """Get the effective page size (width, height) based on orientation."""
        if self.orientation == PageOrientation.LANDSCAPE:
            return max(self.page_width, self.page_height), min(self.page_width, self.page_height)
        return min(self.page_width, self.page_height), max(self.page_width, self.page_height)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization."""
        return {
            'paper_size': self.paper_size,
            'orientation': self.orientation.name,
            'margins': self.margins.to_dict(),
            'page_width': self.page_width,
            'page_height': self.page_height
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'PageSetup':
        """Create from dictionary."""
        return cls(
            paper_size=data.get('paper_size', 'A4'),
            orientation=PageOrientation[data.get('orientation', 'PORTRAIT')],
            margins=PageMargins.from_dict(data.get('margins', {})),
            page_width=float(data.get('page_width', 595.0)),
            page_height=float(data.get('page_height', 842.0))
        )
