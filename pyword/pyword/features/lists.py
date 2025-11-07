from enum import Enum, auto
from PySide6.QtGui import QTextListFormat, QTextCursor, QTextBlockFormat, QTextCharFormat
from PySide6.QtCore import Qt

class ListType(Enum):
    BULLET = auto()
    NUMBERED = auto()
    LOWER_ALPHA = auto()
    UPPER_ALPHA = auto()
    LOWER_ROMAN = auto()
    UPPER_ROMAN = auto()

class ListManager:
    def __init__(self, text_edit):
        self.text_edit = text_edit
        self.list_formats = {
            ListType.BULLET: self._create_bullet_format(),
            ListType.NUMBERED: self._create_numbered_format(),
            ListType.LOWER_ALPHA: self._create_alpha_format(False),
            ListType.UPPER_ALPHA: self._create_alpha_format(True),
            ListType.LOWER_ROMAN: self._create_roman_format(False),
            ListType.UPPER_ROMAN: self._create_roman_format(True)
        }

    def _create_bullet_format(self):
        fmt = QTextListFormat()
        fmt.setStyle(QTextListFormat.Style.ListDisc)
        return fmt

    def _create_numbered_format(self):
        fmt = QTextListFormat()
        fmt.setStyle(QTextListFormat.Style.ListDecimal)
        return fmt

    def _create_alpha_format(self, uppercase):
        fmt = QTextListFormat()
        fmt.setStyle(QTextListFormat.Style.ListLowerAlpha if not uppercase 
                    else QTextListFormat.Style.ListUpperAlpha)
        return fmt

    def _create_roman_format(self, uppercase):
        fmt = QTextListFormat()
        fmt.setStyle(QTextListFormat.Style.ListLowerRoman if not uppercase 
                    else QTextListFormat.Style.ListUpperRoman)
        return fmt

    def toggle_list(self, list_type: ListType):
        cursor = self.text_edit.textCursor()
        cursor.beginEditBlock()

        # Check if current block is already in a list
        current_format = cursor.blockFormat()
        current_list = cursor.currentList()
        
        if current_list:
            # If same list type, remove list formatting
            if self._get_list_type(current_list) == list_type:
                self._remove_list_formatting(cursor)
            else:
                # Different list type, change format
                self._change_list_format(cursor, list_type)
        else:
            # Create new list
            self._create_new_list(cursor, list_type)
        
        cursor.endEditBlock()

    def _get_list_type(self, qlist):
        style = qlist.format().style()
        return {
            QTextListFormat.Style.ListDisc: ListType.BULLET,
            QTextListFormat.Style.ListCircle: ListType.BULLET,
            QTextListFormat.Style.ListSquare: ListType.BULLET,
            QTextListFormat.Style.ListDecimal: ListType.NUMBERED,
            QTextListFormat.Style.ListLowerAlpha: ListType.LOWER_ALPHA,
            QTextListFormat.Style.ListUpperAlpha: ListType.UPPER_ALPHA,
            QTextListFormat.Style.ListLowerRoman: ListType.LOWER_ROMAN,
            QTextListFormat.Style.ListUpperRoman: ListType.UPPER_ROMAN,
        }.get(style, ListType.BULLET)

    def _remove_list_formatting(self, cursor):
        # Get the list and its format
        current_list = cursor.currentList()
        if not current_list:
            return
            
        # Get the block format of the current list item
        block_format = cursor.blockFormat()
        
        # Create a new block format that removes list-specific formatting
        new_block_format = QTextBlockFormat()
        new_block_format.setIndent(0)  # Reset indentation
        
        # Apply the new format to all blocks in the list
        cursor.beginEditBlock()
        for i in range(current_list.count()):
            block = current_list.item(i)
            cursor.setPosition(block.position())
            cursor.setBlockFormat(new_block_format)
        cursor.endEditBlock()
        
        # Clear the list
        cursor.beginEditBlock()
        cursor.currentList().remove(cursor.block())
        cursor.endEditBlock()

    def _change_list_format(self, cursor, new_type: ListType):
        # Get the current list and its items
        current_list = cursor.currentList()
        if not current_list:
            return
            
        # Store the text of all items
        items = []
        for i in range(current_list.count()):
            block = current_list.item(i)
            cursor.setPosition(block.position())
            cursor.movePosition(QTextCursor.MoveOperation.EndOfBlock, QTextCursor.MoveMode.KeepAnchor)
            items.append(cursor.selectedText())
        
        # Remove the old list
        cursor.beginEditBlock()
        current_list.remove(cursor.block())
        cursor.endEditBlock()
        
        # Create new list with the new format
        cursor.beginEditBlock()
        cursor.movePosition(QTextCursor.MoveOperation.StartOfBlock)
        cursor.insertList(self.list_formats[new_type])
        cursor.endEditBlock()

    def _create_new_list(self, cursor, list_type: ListType):
        # If selection exists, apply list to each line in selection
        if cursor.hasSelection():
            start = cursor.selectionStart()
            end = cursor.selectionEnd()
            
            cursor.setPosition(start)
            cursor.movePosition(QTextCursor.MoveOperation.StartOfLine)
            start = cursor.position()
            
            cursor.setPosition(end)
            cursor.movePosition(QTextCursor.MoveOperation.EndOfLine)
            end = cursor.position()
            
            cursor.setPosition(start)
            cursor.setPosition(end, QTextCursor.MoveMode.KeepAnchor)
            
            selected_text = cursor.selectedText()
            lines = selected_text.split('\u2029')  # QTextDocument uses paragraph separators
            
            cursor.beginEditBlock()
            cursor.removeSelectedText()
            
            # Create list and add items
            cursor.insertList(self.list_formats[list_type])
            for i, line in enumerate(lines[1:], 1):
                cursor.insertText(line.strip())
                if i < len(lines) - 1:
                    cursor.insertList(self.list_formats[list_type])
        else:
            # No selection, just create a new list at cursor
            cursor.insertList(self.list_formats[list_type])
            
    def increase_indent(self):
        cursor = self.text_edit.textCursor()
        current_list = cursor.currentList()
        
        if current_list:
            fmt = current_list.format()
            fmt.setIndent(fmt.indent() + 1)
            current_list.setFormat(fmt)
            
    def decrease_indent(self):
        cursor = self.text_edit.textCursor()
        current_list = cursor.currentList()
        
        if current_list:
            fmt = current_list.format()
            if fmt.indent() > 1:  # Don't go below 1
                fmt.setIndent(fmt.indent() - 1)
                current_list.setFormat(fmt)
