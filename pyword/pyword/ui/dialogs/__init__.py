"""Dialog modules for PyWord."""

from .base_dialog import BaseDialog
from .find_replace_dialog import FindReplaceDialog
from .font_dialog import FontDialog
from .options_dialog import OptionsDialog
from .page_setup_dialog import PageSetupDialog
from .paragraph_dialog import ParagraphDialog
from .print_dialog import PrintDialog
from .table_properties_dialog import TablePropertiesDialog
from .word_count_dialog import WordCountDialog

# New dialogs
from .new_document_dialog import NewDocumentDialog
from .print_preview_dialog import PrintPreviewDialog
from .insert_table_dialog import InsertTableDialog
from .insert_image_dialog import InsertImageDialog
from .insert_link_dialog import InsertLinkDialog, HyperlinkDialog
from .goto_dialog import GoToDialog
from .about_dialog import AboutDialog
from .style_dialog import StyleDialog
from .bullets_numbering_dialog import BulletsAndNumberingDialog
from .border_shading_dialog import BorderAndShadingDialog
from .columns_dialog import ColumnsDialog
from .tabs_dialog import TabsDialog
from .symbol_dialog import SymbolDialog

__all__ = [
    'BaseDialog',
    'FindReplaceDialog',
    'FontDialog',
    'OptionsDialog',
    'PageSetupDialog',
    'ParagraphDialog',
    'PrintDialog',
    'TablePropertiesDialog',
    'WordCountDialog',
    'NewDocumentDialog',
    'PrintPreviewDialog',
    'InsertTableDialog',
    'InsertImageDialog',
    'InsertLinkDialog',
    'HyperlinkDialog',
    'GoToDialog',
    'AboutDialog',
    'StyleDialog',
    'BulletsAndNumberingDialog',
    'BorderAndShadingDialog',
    'ColumnsDialog',
    'TabsDialog',
    'SymbolDialog',
]
