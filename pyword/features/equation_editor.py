"""Equation Editor for PyWord - create and edit mathematical equations."""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton,
                               QTextEdit, QLabel, QGroupBox, QGridLayout,
                               QToolButton, QTabWidget, QWidget, QScrollArea,
                               QSizePolicy, QLineEdit, QComboBox, QSplitter)
from PySide6.QtCore import Qt, Signal, QSize
from PySide6.QtGui import QFont, QTextCursor, QTextCharFormat, QTextImageFormat, QImage, QPainter


class EquationSymbol:
    """Represents a mathematical symbol or operator."""

    def __init__(self, latex: str, display: str, category: str = "general"):
        self.latex = latex
        self.display = display
        self.category = category


class EquationEditor(QDialog):
    """Dialog for creating and editing mathematical equations."""

    equation_inserted = Signal(str)  # Emits LaTeX equation

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Equation Editor")
        self.resize(800, 600)

        # Equation symbols organized by category
        self.symbols = self.initialize_symbols()

        self.setup_ui()

    def initialize_symbols(self) -> dict:
        """Initialize mathematical symbols organized by category."""
        symbols = {
            'basic': [
                EquationSymbol(r'+', '+', 'basic'),
                EquationSymbol(r'-', '-', 'basic'),
                EquationSymbol(r'\times', '×', 'basic'),
                EquationSymbol(r'\div', '÷', 'basic'),
                EquationSymbol(r'=', '=', 'basic'),
                EquationSymbol(r'\neq', '≠', 'basic'),
                EquationSymbol(r'\pm', '±', 'basic'),
                EquationSymbol(r'\mp', '∓', 'basic'),
            ],
            'relations': [
                EquationSymbol(r'<', '<', 'relations'),
                EquationSymbol(r'>', '>', 'relations'),
                EquationSymbol(r'\leq', '≤', 'relations'),
                EquationSymbol(r'\geq', '≥', 'relations'),
                EquationSymbol(r'\ll', '≪', 'relations'),
                EquationSymbol(r'\gg', '≫', 'relations'),
                EquationSymbol(r'\approx', '≈', 'relations'),
                EquationSymbol(r'\equiv', '≡', 'relations'),
                EquationSymbol(r'\sim', '∼', 'relations'),
                EquationSymbol(r'\propto', '∝', 'relations'),
            ],
            'greek': [
                EquationSymbol(r'\alpha', 'α', 'greek'),
                EquationSymbol(r'\beta', 'β', 'greek'),
                EquationSymbol(r'\gamma', 'γ', 'greek'),
                EquationSymbol(r'\delta', 'δ', 'greek'),
                EquationSymbol(r'\epsilon', 'ε', 'greek'),
                EquationSymbol(r'\theta', 'θ', 'greek'),
                EquationSymbol(r'\lambda', 'λ', 'greek'),
                EquationSymbol(r'\mu', 'μ', 'greek'),
                EquationSymbol(r'\pi', 'π', 'greek'),
                EquationSymbol(r'\sigma', 'σ', 'greek'),
                EquationSymbol(r'\tau', 'τ', 'greek'),
                EquationSymbol(r'\phi', 'φ', 'greek'),
                EquationSymbol(r'\omega', 'ω', 'greek'),
            ],
            'operators': [
                EquationSymbol(r'\sum', '∑', 'operators'),
                EquationSymbol(r'\prod', '∏', 'operators'),
                EquationSymbol(r'\int', '∫', 'operators'),
                EquationSymbol(r'\oint', '∮', 'operators'),
                EquationSymbol(r'\nabla', '∇', 'operators'),
                EquationSymbol(r'\partial', '∂', 'operators'),
                EquationSymbol(r'\infty', '∞', 'operators'),
                EquationSymbol(r'\lim', 'lim', 'operators'),
            ],
            'arrows': [
                EquationSymbol(r'\rightarrow', '→', 'arrows'),
                EquationSymbol(r'\leftarrow', '←', 'arrows'),
                EquationSymbol(r'\Rightarrow', '⇒', 'arrows'),
                EquationSymbol(r'\Leftarrow', '⇐', 'arrows'),
                EquationSymbol(r'\leftrightarrow', '↔', 'arrows'),
                EquationSymbol(r'\Leftrightarrow', '⇔', 'arrows'),
                EquationSymbol(r'\uparrow', '↑', 'arrows'),
                EquationSymbol(r'\downarrow', '↓', 'arrows'),
            ],
            'sets': [
                EquationSymbol(r'\in', '∈', 'sets'),
                EquationSymbol(r'\notin', '∉', 'sets'),
                EquationSymbol(r'\subset', '⊂', 'sets'),
                EquationSymbol(r'\supset', '⊃', 'sets'),
                EquationSymbol(r'\cup', '∪', 'sets'),
                EquationSymbol(r'\cap', '∩', 'sets'),
                EquationSymbol(r'\emptyset', '∅', 'sets'),
                EquationSymbol(r'\forall', '∀', 'sets'),
                EquationSymbol(r'\exists', '∃', 'sets'),
            ],
        }
        return symbols

    def setup_ui(self):
        """Initialize the equation editor UI."""
        layout = QVBoxLayout(self)

        # Splitter for symbols and preview
        splitter = QSplitter(Qt.Vertical)
        layout.addWidget(splitter)

        # Top section: Symbol palette and templates
        top_widget = QWidget()
        top_layout = QVBoxLayout(top_widget)

        # Tab widget for different symbol categories
        self.symbol_tabs = QTabWidget()
        self.create_symbol_tabs()
        top_layout.addWidget(self.symbol_tabs)

        # Templates section
        templates_group = QGroupBox("Templates")
        templates_layout = QGridLayout(templates_group)
        self.create_templates(templates_layout)
        top_layout.addWidget(templates_group)

        splitter.addWidget(top_widget)

        # Bottom section: LaTeX input and preview
        bottom_widget = QWidget()
        bottom_layout = QVBoxLayout(bottom_widget)

        # LaTeX input
        latex_label = QLabel("LaTeX Code:")
        bottom_layout.addWidget(latex_label)

        self.latex_input = QTextEdit()
        self.latex_input.setMaximumHeight(100)
        self.latex_input.setFont(QFont("Courier New", 10))
        self.latex_input.textChanged.connect(self.update_preview)
        bottom_layout.addWidget(self.latex_input)

        # Preview
        preview_label = QLabel("Preview:")
        bottom_layout.addWidget(preview_label)

        self.preview_area = QLabel()
        self.preview_area.setMinimumHeight(150)
        self.preview_area.setAlignment(Qt.AlignCenter)
        self.preview_area.setStyleSheet("""
            QLabel {
                background-color: white;
                border: 1px solid #d0d0d0;
                font-size: 16pt;
            }
        """)
        bottom_layout.addWidget(self.preview_area)

        splitter.addWidget(bottom_widget)

        # Buttons
        button_layout = QHBoxLayout()
        button_layout.addStretch()

        insert_button = QPushButton("Insert")
        insert_button.clicked.connect(self.insert_equation)
        button_layout.addWidget(insert_button)

        cancel_button = QPushButton("Cancel")
        cancel_button.clicked.connect(self.reject)
        button_layout.addWidget(cancel_button)

        layout.addLayout(button_layout)

    def create_symbol_tabs(self):
        """Create tabs for different symbol categories."""
        for category, symbols in self.symbols.items():
            tab = QWidget()
            scroll = QScrollArea()
            scroll.setWidgetResizable(True)
            scroll.setWidget(tab)

            layout = QGridLayout(tab)
            layout.setSpacing(5)

            row = 0
            col = 0
            max_cols = 8

            for symbol in symbols:
                button = QPushButton(symbol.display)
                button.setToolTip(symbol.latex)
                button.clicked.connect(lambda checked, s=symbol: self.insert_symbol(s))
                button.setMinimumSize(40, 40)
                layout.addWidget(button, row, col)

                col += 1
                if col >= max_cols:
                    col = 0
                    row += 1

            layout.setRowStretch(row + 1, 1)

            self.symbol_tabs.addTab(scroll, category.capitalize())

    def create_templates(self, layout: QGridLayout):
        """Create common equation templates."""
        templates = [
            (r'\frac{a}{b}', 'Fraction'),
            (r'x^{2}', 'Superscript'),
            (r'x_{i}', 'Subscript'),
            (r'\sqrt{x}', 'Square Root'),
            (r'\sqrt[n]{x}', 'nth Root'),
            (r'\sum_{i=1}^{n}', 'Sum'),
            (r'\int_{a}^{b}', 'Integral'),
            (r'\lim_{x \to \infty}', 'Limit'),
            (r'\begin{pmatrix} a & b \\ c & d \end{pmatrix}', 'Matrix'),
            (r'\begin{cases} x & \text{if } x > 0 \\ 0 & \text{otherwise} \end{cases}', 'Cases'),
        ]

        row = 0
        col = 0
        max_cols = 5

        for latex, name in templates:
            button = QPushButton(name)
            button.setToolTip(latex)
            button.clicked.connect(lambda checked, l=latex: self.insert_template(l))
            layout.addWidget(button, row, col)

            col += 1
            if col >= max_cols:
                col = 0
                row += 1

    def insert_symbol(self, symbol: EquationSymbol):
        """Insert a symbol into the LaTeX input."""
        cursor = self.latex_input.textCursor()
        cursor.insertText(symbol.latex + ' ')
        self.latex_input.setTextCursor(cursor)

    def insert_template(self, template: str):
        """Insert a template into the LaTeX input."""
        cursor = self.latex_input.textCursor()
        cursor.insertText(template)
        self.latex_input.setTextCursor(cursor)

    def update_preview(self):
        """Update the equation preview."""
        latex_code = self.latex_input.toPlainText()

        # In a real implementation, this would render the LaTeX using a library like matplotlib
        # For now, we'll just show the LaTeX code in a larger font
        preview_text = f"${latex_code}$" if latex_code else "Preview will appear here"
        self.preview_area.setText(preview_text)

    def insert_equation(self):
        """Insert the equation and close the dialog."""
        latex_code = self.latex_input.toPlainText()
        if latex_code:
            self.equation_inserted.emit(latex_code)
            self.accept()

    def get_equation(self) -> str:
        """Get the current equation as LaTeX."""
        return self.latex_input.toPlainText()

    def set_equation(self, latex: str):
        """Set the equation from LaTeX."""
        self.latex_input.setPlainText(latex)


class EquationRenderer:
    """Renders LaTeX equations to images."""

    def __init__(self):
        self.dpi = 150
        self.font_size = 12

    def render_equation(self, latex: str) -> QImage:
        """Render a LaTeX equation to a QImage.

        Note: In a real implementation, this would use a library like:
        - matplotlib with LaTeX support
        - sympy for rendering
        - Or an external LaTeX renderer

        For now, this returns a simple placeholder image.
        """
        # Create a placeholder image
        image = QImage(400, 100, QImage.Format_ARGB32)
        image.fill(Qt.white)

        painter = QPainter(image)
        painter.setPen(Qt.black)
        painter.setFont(QFont("Times New Roman", 14))
        painter.drawText(image.rect(), Qt.AlignCenter, f"${latex}$")
        painter.end()

        return image

    def render_to_html(self, latex: str) -> str:
        """Render equation to HTML with MathML or similar.

        Note: This would use a proper LaTeX to HTML/MathML converter.
        """
        return f'<span class="equation">${latex}$</span>'


class EquationManager:
    """Manages equations in a document."""

    def __init__(self, editor):
        self.editor = editor
        self.equations = {}  # Maps equation IDs to LaTeX code
        self.renderer = EquationRenderer()

    def insert_equation(self, latex: str):
        """Insert an equation at the current cursor position."""
        cursor = self.editor.textCursor()

        # Generate unique ID for this equation
        equation_id = f"eq_{len(self.equations)}"
        self.equations[equation_id] = latex

        # Render equation to image
        image = self.renderer.render_equation(latex)

        # Insert image into document
        image_format = QTextImageFormat()
        image_format.setName(equation_id)

        # Store image in document resources
        self.editor.document().addResource(
            QTextCursor.ImageResource,
            equation_id,
            image
        )

        # Insert the image
        cursor.insertImage(image_format)

    def edit_equation(self, equation_id: str) -> bool:
        """Open equation editor for an existing equation."""
        if equation_id not in self.equations:
            return False

        latex = self.equations[equation_id]

        # Open equation editor
        dialog = EquationEditor(self.editor)
        dialog.set_equation(latex)

        if dialog.exec() == QDialog.Accepted:
            new_latex = dialog.get_equation()
            self.equations[equation_id] = new_latex

            # Re-render equation
            image = self.renderer.render_equation(new_latex)
            self.editor.document().addResource(
                QTextCursor.ImageResource,
                equation_id,
                image
            )
            return True

        return False

    def show_equation_editor(self):
        """Show the equation editor dialog."""
        dialog = EquationEditor(self.editor)
        dialog.equation_inserted.connect(self.insert_equation)
        dialog.exec()

    def export_equations(self) -> dict:
        """Export all equations as LaTeX."""
        return self.equations.copy()

    def import_equations(self, equations: dict):
        """Import equations from LaTeX."""
        self.equations.update(equations)

        # Re-render all imported equations
        for equation_id, latex in equations.items():
            image = self.renderer.render_equation(latex)
            self.editor.document().addResource(
                QTextCursor.ImageResource,
                equation_id,
                image
            )
