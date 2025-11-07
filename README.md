# PyWord

A comprehensive Microsoft Word clone built with Python and PySide6, featuring a complete word processing experience with professional-grade functionality.

![PyWord](https://via.placeholder.com/1200x600/2c3e50/ecf0f1?text=PyWord+-+Professional+Word+Processor)

## Overview

PyWord is a full-featured word processor that replicates Microsoft Word's core functionality using Python and the Qt framework (PySide6). The application provides a modern, cross-platform document editing experience with support for multiple file formats, advanced formatting, collaboration tools, and enterprise features.

## Features

### Core Functionality
- **Document Management**: Create, open, save, and manage multiple documents with tabs
- **Rich Text Editing**: Comprehensive text formatting (bold, italic, underline, fonts, colors, sizes)
- **Paragraph Formatting**: Alignment, indentation, spacing, and advanced paragraph controls
- **Page Setup**: Custom margins, orientation, paper size, and print settings
- **Undo/Redo**: Full undo/redo support with history tracking

### Advanced Formatting
- **Styles & Themes**: Pre-defined and custom styles with theme support
- **Lists**: Bullets, numbering, and multi-level lists
- **Tables**: Create and format tables with advanced cell merging and styling
- **Columns**: Multi-column text layout
- **Headers & Footers**: Different first page and odd/even page support
- **Page Numbers**: Customizable page numbering with various formats
- **Sections**: Section breaks with independent formatting
- **Page Backgrounds**: Colors, gradients, and watermarks

### Document Navigation
- **Find & Replace**: Advanced search with regex support
- **Go To**: Jump to specific pages, sections, or elements
- **Document Map**: Navigation pane showing document structure
- **Split View**: Horizontal and vertical split views for easier editing
- **View Modes**: Print Layout, Web Layout, and Draft views

### Graphics & Media
- **Images**: Insert and format images with advanced positioning and text wrapping
- **Shapes**: Rectangles, circles, lines, arrows, and custom shapes
- **SmartArt**: Diagrams and organizational charts
- **Charts**: Data visualization with various chart types
- **3D Models**: Insert and manipulate 3D objects
- **Equation Editor**: Mathematical equation editing

### Collaboration
- **Track Changes**: Monitor and review document modifications
- **Comments**: Add threaded comments with author tracking
- **Document Comparison**: Compare two versions side-by-side
- **Review Tools**: Accept/reject changes, resolve comments

### References & Citations
- **Table of Contents**: Auto-generated with heading styles
- **Footnotes & Endnotes**: Academic reference support
- **Citations & Bibliography**: Multiple citation styles (APA, MLA, Chicago)
- **Captions**: Automatic numbering for figures, tables, and equations
- **Cross-References**: Link to headings, figures, and bookmarks

### Professional Features
- **Mail Merge**: Connect to data sources and generate personalized documents
- **Macros**: Automation through Python scripting
- **Security**: Document protection, password encryption, digital signatures
- **Document Properties**: Metadata and custom properties
- **Information Rights Management**: Access control and permissions

### Modern UI/UX
- **Ribbon Interface**: Organized toolbar with contextual tabs
- **Quick Access Toolbar**: Customizable shortcuts
- **Dark/Light Themes**: Multiple color schemes
- **Touch Support**: Optimized for touch screens
- **Multi-Monitor**: Support for multiple displays
- **Accessibility**: Screen reader support, high contrast, keyboard navigation, text-to-speech

### Performance & Reliability
- **Large Document Optimization**: Efficient handling of documents with thousands of pages
- **Memory Management**: Smart resource allocation and cleanup
- **Background Saving**: Non-blocking auto-save
- **Auto-Recovery**: Recover documents after crashes
- **Performance Monitoring**: Real-time performance metrics

### File Format Support
- **DOCX**: Full Microsoft Word document support (read/write)
- **ODT**: OpenDocument Text format (read/write)
- **PDF**: Export to PDF with formatting preservation
- **RTF**: Rich Text Format (read/write)
- **HTML**: Web page export (read/write)

## Installation

### Prerequisites
- Python 3.8 or higher
- pip package manager

### Install Dependencies

```bash
pip install -r requirements.txt
```

### Required Packages
- PySide6 (Qt for Python UI framework)
- python-docx (DOCX file support)
- Pillow (Image processing)
- odfpy (ODT file support)
- reportlab (PDF generation)
- beautifulsoup4 & lxml (HTML parsing)
- striprtf (RTF processing)

## Usage

### Running PyWord

From the project directory:

```bash
python -m pyword
```

Or run the main module directly:

```bash
python pyword/__main__.py
```

### Basic Operations

**Create New Document**: File → New (Ctrl+N)

**Open Document**: File → Open (Ctrl+O)

**Save Document**: File → Save (Ctrl+S)

**Save As**: File → Save As (Ctrl+Shift+S)

**Print**: File → Print (Ctrl+P)

### Keyboard Shortcuts

#### File Operations
- `Ctrl+N`: New document
- `Ctrl+O`: Open file
- `Ctrl+S`: Save
- `Ctrl+Shift+S`: Save As
- `Ctrl+P`: Print
- `Ctrl+Q`: Quit

#### Editing
- `Ctrl+Z`: Undo
- `Ctrl+Y` or `Ctrl+Shift+Z`: Redo
- `Ctrl+X`: Cut
- `Ctrl+C`: Copy
- `Ctrl+V`: Paste
- `Ctrl+A`: Select All
- `Ctrl+F`: Find
- `Ctrl+H`: Replace
- `Ctrl+G`: Go To

#### Formatting
- `Ctrl+B`: Bold
- `Ctrl+I`: Italic
- `Ctrl+U`: Underline

#### View
- `F12`: Toggle preview panel

## Project Structure

```
PyWord/
├── pyword/
│   ├── __main__.py                 # Application entry point
│   ├── core/                       # Core functionality
│   │   ├── application.py          # Main application window
│   │   ├── editor.py               # Text editor component
│   │   ├── document.py             # Document model
│   │   ├── file_formats.py         # File format handlers
│   │   ├── print_manager.py        # Print functionality
│   │   └── page_setup.py           # Page configuration
│   ├── features/                   # Feature modules
│   │   ├── accessibility.py        # Accessibility features
│   │   ├── advanced_images.py      # Image handling
│   │   ├── advanced_tables.py      # Advanced table features
│   │   ├── automation.py           # Macros and scripting
│   │   ├── captions.py             # Figure/table captions
│   │   ├── charts.py               # Chart creation
│   │   ├── citations.py            # Bibliography management
│   │   ├── columns.py              # Multi-column layout
│   │   ├── comments.py             # Document comments
│   │   ├── cross_references.py     # Cross-reference support
│   │   ├── document_comparison.py  # Document diff
│   │   ├── equation_editor.py      # Math equations
│   │   ├── footnotes.py            # Footnotes/endnotes
│   │   ├── headers_footers.py      # Header/footer management
│   │   ├── lists.py                # Bullet/numbered lists
│   │   ├── mail_merge.py           # Mail merge functionality
│   │   ├── models_3d.py            # 3D model support
│   │   ├── navigation.py           # Find/replace/go to
│   │   ├── page_backgrounds.py     # Page backgrounds
│   │   ├── page_numbers.py         # Page numbering
│   │   ├── performance.py          # Performance optimization
│   │   ├── references.py           # Table of contents
│   │   ├── review.py               # Review tools
│   │   ├── sections.py             # Section management
│   │   ├── security.py             # Document security
│   │   ├── shapes.py               # Shape insertion
│   │   ├── smartart.py             # SmartArt diagrams
│   │   ├── split_view.py           # Split view mode
│   │   ├── styles.py               # Style management
│   │   ├── tables.py               # Table functionality
│   │   ├── text_wrapping.py        # Text wrapping
│   │   └── track_changes.py        # Change tracking
│   └── ui/                         # User interface
│       ├── main_window.py          # Main window
│       ├── ribbon.py               # Ribbon interface
│       ├── quick_access_toolbar.py # Quick access toolbar
│       ├── theme_manager.py        # Theme system
│       ├── touch_support.py        # Touch interface
│       ├── multi_monitor.py        # Multi-display support
│       ├── dialogs/                # Dialog windows
│       ├── panels/                 # Side panels
│       └── toolbars/               # Toolbar components
├── requirements.txt                # Python dependencies
├── roadmap.md                      # Development roadmap
└── README.md                       # This file
```

## Development

### Architecture

PyWord follows a modular architecture:

- **Core**: Essential word processing functionality (document model, file I/O, printing)
- **Features**: Self-contained feature modules that extend core functionality
- **UI**: User interface components built with PySide6/Qt

Each feature module is designed to be independent and can be enabled/disabled as needed.

### Code Guidelines

- Follow PEP 8 style conventions
- Use type hints for better code clarity
- Document all public APIs with docstrings
- Write unit tests for new features
- Maintain backward compatibility when possible

### Development Status

PyWord has completed Phases 1-8 of development (see roadmap.md for details):

- ✅ Phase 1: Core Functionality (MVP)
- ✅ Phase 2: Intermediate Features
- ✅ Phase 3: Advanced Features
- ✅ Phase 4: Professional Features
- ✅ Phase 5: Enterprise Features
- ✅ Phase 6: Polish and Optimization
- ✅ Phase 7: Integration and Extensibility (File Formats)
- ✅ Phase 8: Advanced UI/UX

### Future Development

Planned enhancements include:
- Cloud storage integration (Dropbox, Google Drive, OneDrive)
- Real-time collaboration
- Mobile version
- Web version
- Plugin architecture
- AI-powered features (auto-formatting, writing suggestions)

## Building Standalone Executable

To create a standalone executable using PyInstaller:

```bash
# Install PyInstaller
pip install pyinstaller

# Build executable
pyinstaller --onefile --windowed --name PyWord pyword/__main__.py

# The executable will be in the dist/ directory
```

For a more complete build with all resources:

```bash
pyinstaller --name PyWord \
    --windowed \
    --onedir \
    --add-data "resources:resources" \
    pyword/__main__.py
```

## Contributing

Contributions are welcome! Please follow these steps:

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Write or update tests as needed
5. Ensure code follows PEP 8 style guidelines
6. Commit your changes (`git commit -m 'Add amazing feature'`)
7. Push to the branch (`git push origin feature/amazing-feature`)
8. Open a Pull Request

### Bug Reports

When reporting bugs, please include:
- PyWord version
- Operating system and version
- Python version
- Steps to reproduce the issue
- Expected vs actual behavior
- Screenshots if applicable

## Testing

Run the test suite:

```bash
# Run all tests
python -m pytest

# Run with coverage
python -m pytest --cov=pyword

# Run specific test file
python -m pytest tests/test_document.py
```

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Built with [PySide6](https://www.qt.io/qt-for-python) (Qt for Python)
- File format support from [python-docx](https://python-docx.readthedocs.io/), [odfpy](https://github.com/eea/odfpy), and [ReportLab](https://www.reportlab.com/)
- Inspired by Microsoft Word and other professional word processors

## Support

For questions, issues, or feature requests, please open an issue on the project's issue tracker.

---

**PyWord** - Professional word processing with Python
