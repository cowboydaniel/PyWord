from PySide6.QtWidgets import QToolBar, QComboBox, QSpinBox
from PySide6.QtGui import QAction
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QIcon, QTextTableFormat, QTextCursor

class TableToolBar(QToolBar):
    def __init__(self, parent=None):
        super().__init__("Table Toolbar", parent)
        self.setIconSize(QSize(16, 16))
        self.setToolButtonStyle(Qt.ToolButtonIconOnly)
        self.setup_ui()
    
    def setup_ui(self):
        # Insert Table
        self.insert_table_action = QAction("Insert Table", self)
        self.insert_table_action.setIcon(self.style().standardIcon("SP_FileDialogDetailedView"))
        self.insert_table_action.setToolTip("Insert Table")
        
        # Delete Table
        self.delete_table_action = QAction("Delete Table", self)
        self.delete_table_action.setIcon(self.style().standardIcon("SP_TrashIcon"))
        self.delete_table_action.setToolTip("Delete Table")
        
        # Add table actions
        self.addAction(self.insert_table_action)
        self.addAction(self.delete_table_action)
        self.addSeparator()
        
        # Insert Rows
        self.insert_row_above_action = QAction("Insert Row Above", self)
        self.insert_row_above_action.setIcon(self.style().standardIcon("SP_ArrowUp"))
        self.insert_row_above_action.setToolTip("Insert Row Above")
        
        self.insert_row_below_action = QAction("Insert Row Below", self)
        self.insert_row_below_action.setIcon(self.style().standardIcon("SP_ArrowDown"))
        self.insert_row_below_action.setToolTip("Insert Row Below")
        
        # Delete Rows
        self.delete_row_action = QAction("Delete Rows", self)
        self.delete_row_action.setIcon(self.style().standardIcon("SP_DialogCloseButton"))
        self.delete_row_action.setToolTip("Delete Rows")
        
        # Add row actions
        self.addAction(self.insert_row_above_action)
        self.addAction(self.insert_row_below_action)
        self.addAction(self.delete_row_action)
        self.addSeparator()
        
        # Insert Columns
        self.insert_column_left_action = QAction("Insert Column Left", self)
        self.insert_column_left_action.setIcon(self.style().standardIcon("SP_ArrowLeft"))
        self.insert_column_left_action.setToolTip("Insert Column to the Left")
        
        self.insert_column_right_action = QAction("Insert Column Right", self)
        self.insert_column_right_action.setIcon(self.style().standardIcon("SP_ArrowRight"))
        self.insert_column_right_action.setToolTip("Insert Column to the Right")
        
        # Delete Columns
        self.delete_column_action = QAction("Delete Columns", self)
        self.delete_column_action.setIcon(self.style().standardIcon("SP_DialogCloseButton"))
        self.delete_column_action.setToolTip("Delete Columns")
        
        # Add column actions
        self.addAction(self.insert_column_left_action)
        self.addAction(self.insert_column_right_action)
        self.addAction(self.delete_column_action)
        self.addSeparator()
        
        # Merge Cells
        self.merge_cells_action = QAction("Merge Cells", self)
        self.merge_cells_action.setIcon(self.style().standardIcon("SP_FileLinkIcon"))
        self.merge_cells_action.setToolTip("Merge Cells")
        
        # Split Cells
        self.split_cells_action = QAction("Split Cells", self)
        self.split_cells_action.setIcon(self.style().standardIcon("SP_FileDialogDetailedView"))
        self.split_cells_action.setToolTip("Split Cells")
        
        # Add cell actions
        self.addAction(self.merge_cells_action)
        self.addAction(self.split_cells_action)
        self.addSeparator()
        
        # Table Properties
        self.table_properties_action = QAction("Table Properties", self)
        self.table_properties_action.setIcon(self.style().standardIcon("SP_ComputerIcon"))
        self.table_properties_action.setToolTip("Table Properties")
        
        # Add properties action
        self.addAction(self.table_properties_action)
        self.addSeparator()
        
        # Cell Alignment
        self.cell_align_top_left = QAction("Top Left", self)
        self.cell_align_top_center = QAction("Top Center", self)
        self.cell_align_top_right = QAction("Top Right", self)
        self.cell_align_middle_left = QAction("Middle Left", self)
        self.cell_align_middle_center = QAction("Middle Center", self)
        self.cell_align_middle_right = QAction("Middle Right", self)
        self.cell_align_bottom_left = QAction("Bottom Left", self)
        self.cell_align_bottom_center = QAction("Bottom Center", self)
        self.cell_align_bottom_right = QAction("Bottom Right", self)
        
        # Cell alignment menu
        self.cell_alignment_menu = QMenu(self)
        self.cell_alignment_menu.addAction(self.cell_align_top_left)
        self.cell_alignment_menu.addAction(self.cell_align_top_center)
        self.cell_alignment_menu.addAction(self.cell_align_top_right)
        self.cell_alignment_menu.addSeparator()
        self.cell_alignment_menu.addAction(self.cell_align_middle_left)
        self.cell_alignment_menu.addAction(self.cell_align_middle_center)
        self.cell_alignment_menu.addAction(self.cell_align_middle_right)
        self.cell_alignment_menu.addSeparator()
        self.cell_alignment_menu.addAction(self.cell_align_bottom_left)
        self.cell_alignment_menu.addAction(self.cell_align_bottom_center)
        self.cell_alignment_menu.addAction(self.cell_align_bottom_right)
        
        # Cell alignment button
        self.cell_alignment_button = QToolButton(self)
        self.cell_alignment_button.setPopupMode(QToolButton.MenuButtonPopup)
        self.cell_alignment_button.setMenu(self.cell_alignment_menu)
        self.cell_alignment_button.setText("Cell Alignment")
        self.cell_alignment_button.setToolTip("Cell Alignment")
        self.cell_alignment_button.setIcon(self.style().standardIcon("SP_FileDialogDetailedView"))
        self.addWidget(self.cell_alignment_button)
