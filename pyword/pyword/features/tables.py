from PySide6.QtWidgets import QTableWidget, QTableWidgetItem, QMenu, QHeaderView
from PySide6.QtCore import Qt, QSize
from PySide6.QtGui import QAction, QKeySequence

class TableManager:
    def __init__(self, parent):
        self.parent = parent
        self.current_table = None

    def insert_table(self, rows=3, cols=3):
        """Insert a new table at the current cursor position."""
        cursor = self.parent.textCursor()
        if not cursor.isNull():
            # Create table
            table_format = cursor.currentTableFormat()
            table_format.setBorder(1)
            table_format.setCellSpacing(0)
            table_format.setCellPadding(5)
            
            self.current_table = cursor.insertTable(rows, cols, table_format)
            
            # Set header row formatting
            header_format = cursor.charFormat()
            header_format.setFontWeight(Qt.Weight.Bold)
            header_format.setBackground(Qt.GlobalColor.lightGray)
            
            # Fill header cells
            for col in range(cols):
                cursor = self.current_table.cellAt(0, col).firstCursorPosition()
                cursor.setCharFormat(header_format)
                cursor.insertText(f"Column {col+1}")
            
            # Fill data cells
            for row in range(1, rows):
                for col in range(cols):
                    cursor = self.current_table.cellAt(row, col).firstCursorPosition()
                    cursor.insertText(f"Cell {row+1},{col+1}")
            
            return self.current_table
        return None

    def delete_table(self):
        """Delete the currently selected table."""
        cursor = self.parent.textCursor()
        if cursor.currentTable():
            cursor.currentTable().removeRows(0, cursor.currentTable().rows())
            return True
        return False

    def table_properties(self):
        """Show table properties dialog."""
        # TODO: Implement table properties dialog
        pass

    def resize_columns(self):
        """Auto-resize columns to fit content."""
        if self.current_table:
            self.current_table.resizeColumnsToContents()

    def insert_row(self):
        """Insert a new row in the current table."""
        cursor = self.parent.textCursor()
        table = cursor.currentTable()
        if table:
            table.appendRows(1)

    def insert_column(self):
        """Insert a new column in the current table."""
        cursor = self.parent.textCursor()
        table = cursor.currentTable()
        if table:
            table.appendColumns(1)

    def delete_row(self):
        """Delete the current row from the table."""
        cursor = self.parent.textCursor()
        table = cursor.currentTable()
        if table:
            cell = table.cellAt(cursor)
            if cell.isValid():
                table.removeRows(cell.row(), 1)

    def delete_column(self):
        """Delete the current column from the table."""
        cursor = self.parent.textCursor()
        table = cursor.currentTable()
        if table:
            cell = table.cellAt(cursor)
            if cell.isValid():
                table.removeColumns(cell.column(), 1)

    def show_context_menu(self, pos):
        """Show context menu for table operations."""
        menu = QMenu(self.parent)
        
        insert_row = menu.addAction("Insert Row Above")
        insert_row.triggered.connect(self.insert_row)
        
        insert_col = menu.addAction("Insert Column Left")
        insert_col.triggered.connect(self.insert_column)
        
        menu.addSeparator()
        
        delete_row = menu.addAction("Delete Row")
        delete_row.triggered.connect(self.delete_row)
        
        delete_col = menu.addAction("Delete Column")
        delete_col.triggered.connect(self.delete_column)
        
        menu.addSeparator()
        
        auto_fit = menu.addAction("AutoFit")
        auto_fit.triggered.connect(self.resize_columns)
        
        menu.exec(self.parent.viewport().mapToGlobal(pos))
