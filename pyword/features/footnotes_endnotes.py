"""Footnote and endnote management for PyWord documents."""

from dataclasses import dataclass, field
from typing import Dict, List, Any
from PySide6.QtCore import QObject, Signal


@dataclass
class Note:
    """Represents a footnote or endnote."""
    number: int
    text: str
    position: int = 0  # Position in the document where the note reference appears

    def to_dict(self) -> Dict[str, Any]:
        return {
            'number': self.number,
            'text': self.text,
            'position': self.position
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Note':
        return cls(
            number=data.get('number', 0),
            text=data.get('text', ''),
            position=data.get('position', 0)
        )


class NoteManager(QObject):
    """Manages footnotes and endnotes in a document."""

    noteAdded = Signal(str, int)  # Signal emitted when a note is added (type, number)
    noteRemoved = Signal(str, int)  # Signal emitted when a note is removed

    def __init__(self, parent=None):
        super().__init__(parent)
        self.footnotes: List[Note] = []
        self.endnotes: List[Note] = []

    def add_footnote(self, text: str, position: int = 0) -> int:
        """Add a footnote and return its number."""
        number = len(self.footnotes) + 1
        footnote = Note(number=number, text=text, position=position)
        self.footnotes.append(footnote)
        self.noteAdded.emit('footnote', number)
        return number

    def add_endnote(self, text: str, position: int = 0) -> int:
        """Add an endnote and return its number (as roman numeral index)."""
        number = len(self.endnotes) + 1
        endnote = Note(number=number, text=text, position=position)
        self.endnotes.append(endnote)
        self.noteAdded.emit('endnote', number)
        return number

    def get_footnote(self, number: int) -> Note:
        """Get a footnote by number."""
        for note in self.footnotes:
            if note.number == number:
                return note
        return None

    def get_endnote(self, number: int) -> Note:
        """Get an endnote by number."""
        for note in self.endnotes:
            if note.number == number:
                return note
        return None

    def remove_footnote(self, number: int) -> bool:
        """Remove a footnote by number."""
        for i, note in enumerate(self.footnotes):
            if note.number == number:
                self.footnotes.pop(i)
                self._renumber_footnotes()
                self.noteRemoved.emit('footnote', number)
                return True
        return False

    def remove_endnote(self, number: int) -> bool:
        """Remove an endnote by number."""
        for i, note in enumerate(self.endnotes):
            if note.number == number:
                self.endnotes.pop(i)
                self._renumber_endnotes()
                self.noteRemoved.emit('endnote', number)
                return True
        return False

    def _renumber_footnotes(self):
        """Renumber all footnotes sequentially."""
        for i, note in enumerate(self.footnotes):
            note.number = i + 1

    def _renumber_endnotes(self):
        """Renumber all endnotes sequentially."""
        for i, note in enumerate(self.endnotes):
            note.number = i + 1

    def get_all_footnotes(self) -> List[Note]:
        """Get all footnotes."""
        return self.footnotes.copy()

    def get_all_endnotes(self) -> List[Note]:
        """Get all endnotes."""
        return self.endnotes.copy()

    def clear_all_footnotes(self):
        """Remove all footnotes."""
        self.footnotes.clear()

    def clear_all_endnotes(self):
        """Remove all endnotes."""
        self.endnotes.clear()

    def to_roman(self, num: int) -> str:
        """Convert a number to lowercase roman numerals."""
        val = [
            1000, 900, 500, 400,
            100, 90, 50, 40,
            10, 9, 5, 4,
            1
        ]
        syms = [
            'm', 'cm', 'd', 'cd',
            'c', 'xc', 'l', 'xl',
            'x', 'ix', 'v', 'iv',
            'i'
        ]
        roman_num = ''
        i = 0
        while num > 0:
            for _ in range(num // val[i]):
                roman_num += syms[i]
                num -= val[i]
            i += 1
        return roman_num

    def get_footnote_count(self) -> int:
        """Get the total number of footnotes."""
        return len(self.footnotes)

    def get_endnote_count(self) -> int:
        """Get the total number of endnotes."""
        return len(self.endnotes)

    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary for serialization."""
        return {
            'footnotes': [note.to_dict() for note in self.footnotes],
            'endnotes': [note.to_dict() for note in self.endnotes]
        }

    def from_dict(self, data: Dict[str, Any]):
        """Load from dictionary."""
        self.footnotes = [Note.from_dict(note_data) for note_data in data.get('footnotes', [])]
        self.endnotes = [Note.from_dict(note_data) for note_data in data.get('endnotes', [])]
