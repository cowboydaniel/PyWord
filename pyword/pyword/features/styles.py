from PySide6.QtGui import QTextCharFormat, QTextBlockFormat, QTextListFormat, QTextFormat, QFont
from PySide6.QtCore import Qt

class DocumentStyles:
    def __init__(self):
        self.styles = {
            'Normal': self._create_normal_style(),
            'Heading 1': self._create_heading_style(18, True),
            'Heading 2': self._create_heading_style(16, True),
            'Heading 3': self._create_heading_style(14, True, italic=True),
            'Quote': self._create_quote_style(),
            'Code': self._create_code_style()
        }
        self.current_theme = 'Light'
        self.themes = {
            'Light': {
                'background': '#ffffff',
                'text': '#000000',
                'highlight': '#e6f3ff',
                'accent': '#0078d7'
            },
            'Dark': {
                'background': '#2d2d2d',
                'text': '#e0e0e0',
                'highlight': '#3e4a52',
                'accent': '#4fa6ed'
            },
            'Sepia': {
                'background': '#f4ecd8',
                'text': '#5b4636',
                'highlight': '#e4d4b1',
                'accent': '#9c786c'
            }
        }

    def _create_normal_style(self):
        fmt = QTextCharFormat()
        fmt.setFontPointSize(12)
        return fmt

    def _create_heading_style(self, size, bold=False, italic=False):
        fmt = QTextCharFormat()
        fmt.setFontPointSize(size)
        fmt.setFontWeight(QFont.Bold if bold else QFont.Normal)
        fmt.setFontItalic(italic)
        return fmt

    def _create_quote_style(self):
        block_fmt = QTextBlockFormat()
        block_fmt.setLeftMargin(25)
        block_fmt.setRightMargin(25)
        block_fmt.setTopMargin(10)
        block_fmt.setBottomMargin(10)
        
        char_fmt = QTextCharFormat()
        char_fmt.setFontItalic(True)
        
        # Return both formats as a tuple
        return (block_fmt, char_fmt)

    def _create_code_style(self):
        fmt = QTextCharFormat()
        fmt.setFontFamily('Courier')
        fmt.setBackground(Qt.GlobalColor.lightGray)
        fmt.setFontFixedPitch(True)
        return fmt

    def get_style(self, name):
        return self.styles.get(name, self.styles['Normal'])

    def set_theme(self, theme_name):
        if theme_name in self.themes:
            self.current_theme = theme_name
            return True
        return False

    def get_theme_colors(self):
        return self.themes.get(self.current_theme, self.themes['Light'])
