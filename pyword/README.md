# Nebula Notes

A beautiful, modern markdown note-taking application built with PySide6.

![Nebula Notes Screenshot](https://via.placeholder.com/1200x800/282a36/f8f8f2?text=Nebula+Notes+Screenshot)

## Features

- **Markdown Support**: Write and preview markdown in real-time
- **Code Highlighting**: Syntax highlighting for code blocks in multiple languages
- **Dark Theme**: Easy on the eyes with a beautiful dark color scheme
- **Split View**: Edit and preview your notes simultaneously
- **Rich Text Formatting**: Toolbar and keyboard shortcuts for common formatting
- **File Management**: Create, open, save, and delete notes with ease
- **Searchable**: Quickly find your notes in the sidebar
- **Cross-Platform**: Works on Linux, Windows, and macOS

## Installation

1. Make sure you have Python 3.8+ installed
2. Install the required dependencies:

```bash
pip install -r requirements.txt
```

## Usage

Run the application:

```bash
python notes_app.py
```

### Keyboard Shortcuts

- `Ctrl+N`: New note
- `Ctrl+O`: Open file
- `Ctrl+S`: Save current note
- `Ctrl+Shift+S`: Save as...
- `Ctrl+Q`: Quit application
- `Ctrl+Z`: Undo
- `Ctrl+Shift+Z`: Redo
- `Ctrl+X`: Cut
- `Ctrl+C`: Copy
- `Ctrl+V`: Paste
- `F12`: Toggle preview panel
- `Ctrl+B`: Bold
- `Ctrl+I`: Italic

## Building a Standalone Executable (Optional)

You can use PyInstaller to create a standalone executable:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=icon.ico notes_app.py
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.
