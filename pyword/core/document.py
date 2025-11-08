import os
import json
from dataclasses import dataclass, field, asdict
from enum import Enum, auto
from typing import Optional, List, Dict, Any, Union, Tuple, TYPE_CHECKING
from datetime import datetime
from pathlib import Path

# Import PageSetup (needed at runtime for default_factory)
from .page_setup import PageSetup

if TYPE_CHECKING:
    from ..features.headers_footers import HeaderFooterManager, HeaderFooterType
    from ..features.footnotes_endnotes import NoteManager

# Import file format manager
from .file_formats import format_manager

class DocumentType(Enum):
    """Supported document types."""
    TEXT = "text/plain"
    RTF = "text/rtf"
    HTML = "text/html"
    DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    ODT = "application/vnd.oasis.opendocument.text"
    
    @classmethod
    def from_extension(cls, extension: str) -> 'DocumentType':
        """Get document type from file extension."""
        extension = extension.lower().lstrip('.')
        return {
            'txt': cls.TEXT,
            'rtf': cls.RTF,
            'html': cls.HTML,
            'htm': cls.HTML,
            'docx': cls.DOCX,
            'odt': cls.ODT
        }.get(extension, cls.TEXT)

@dataclass
class DocumentSection:
    """Represents a section within a document."""
    name: str
    content: str = ""
    level: int = 1
    metadata: Dict[str, Any] = field(default_factory=dict)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert section to dictionary."""
        return {
            'name': self.name,
            'content': self.content,
            'level': self.level,
            'metadata': self.metadata
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'DocumentSection':
        """Create section from dictionary."""
        return cls(
            name=data.get('name', ''),
            content=data.get('content', ''),
            level=data.get('level', 1),
            metadata=data.get('metadata', {})
        )

@dataclass
class Document:
    """Represents a document in the word processor with support for rich content and sections."""
    file_path: Optional[str] = None
    title: str = "Untitled"
    content: str = ""  # Main content or raw content for simple documents
    sections: List[DocumentSection] = field(default_factory=list)
    modified: bool = False
    created_at: datetime = field(default_factory=datetime.now)
    modified_at: datetime = field(default_factory=datetime.now)
    document_type: DocumentType = DocumentType.RTF
    metadata: Dict[str, Any] = field(default_factory=dict)
    version: str = "1.0"
    author: str = ""
    keywords: List[str] = field(default_factory=list)
    template: Optional[str] = None
    read_only: bool = False
    page_setup: 'PageSetup' = field(default_factory=lambda: PageSetup())
    _content_type: str = field(init=False, default="")
    _encoding: str = field(init=False, default="utf-8")
    _header_footer_manager: Optional['HeaderFooterManager'] = field(init=False, default=None)
    _note_manager: Optional['NoteManager'] = field(init=False, default=None)
    
    def __post_init__(self):
        if self.file_path:
            path = Path(self.file_path)
            self.title = path.stem
            # Infer document type from file extension if not specified
            if not hasattr(self, '_content_type') or not self._content_type:
                self._content_type = DocumentType.from_extension(path.suffix).value
            
        # Initialize with a default section if none exists
        if not self.sections:
            self.sections = [DocumentSection(name="Main")]
            
        # Initialize header/footer manager
        from ..features.headers_footers import HeaderFooterManager, HeaderFooterType
        self._header_footer_manager = HeaderFooterManager(None)  # Document will be set when loaded in UI

        # Set default header/footer content
        self._header_footer_manager.set_header_footer_content(
            HeaderFooterType.HEADER,
            f"{self.title if self.title else 'Untitled Document'}\t\tPage \\n"
        )
        self._header_footer_manager.set_header_footer_content(
            HeaderFooterType.FOOTER,
            "Created with PyWord\t\\t\\d"
        )

        # Initialize note manager
        from ..features.footnotes_endnotes import NoteManager
        self._note_manager = NoteManager()
    
    @property
    def content_type(self) -> str:
        """Get the content type of the document."""
        if not self._content_type:
            self._content_type = self.document_type.value
        return self._content_type
    
    @content_type.setter
    def content_type(self, value: str):
        """Set the content type of the document."""
        self._content_type = value
        # Update document_type if it matches a known type
        for doc_type in DocumentType:
            if doc_type.value == value:
                self.document_type = doc_type
                break
    
    def add_section(self, name: str, content: str = "", level: int = 1, 
                   metadata: Optional[Dict[str, Any]] = None) -> DocumentSection:
        """Add a new section to the document."""
        section = DocumentSection(
            name=name,
            content=content,
            level=level,
            metadata=metadata or {}
        )
        self.sections.append(section)
        self.modified = True
        return section
    
    def get_section(self, name: str) -> Optional[DocumentSection]:
        """Get a section by name."""
        for section in self.sections:
            if section.name == name:
                return section
        return None
    
    def update_metadata(self, **kwargs):
        """Update document metadata."""
        self.metadata.update(kwargs)
        self.modified = True
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert document to dictionary for serialization."""
        return {
            'title': self.title,
            'content': self.content,
            'sections': [section.to_dict() for section in self.sections],
            'created_at': self.created_at.isoformat(),
            'modified_at': self.modified_at.isoformat(),
            'document_type': self.document_type.value,
            'metadata': self.metadata,
            'version': self.version,
            'author': self.author,
            'keywords': self.keywords,
            'template': self.template,
            'read_only': self.read_only
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'Document':
        """Create document from dictionary."""
        doc = cls(
            title=data.get('title', 'Untitled'),
            content=data.get('content', ''),
            document_type=DocumentType(data.get('document_type', DocumentType.RTF.value)),
            metadata=data.get('metadata', {}),
            version=data.get('version', '1.0'),
            author=data.get('author', ''),
            keywords=data.get('keywords', []),
            template=data.get('template'),
            read_only=data.get('read_only', False)
        )
        
        # Add sections if present
        for section_data in data.get('sections', []):
            doc.sections.append(DocumentSection.from_dict(section_data))
        
        # Set timestamps
        if 'created_at' in data:
            doc.created_at = datetime.fromisoformat(data['created_at'])
        if 'modified_at' in data:
            doc.modified_at = datetime.fromisoformat(data['modified_at'])
        
        return doc
    
    def save(self, file_path: Optional[Union[str, Path]] = None,
            format: Optional[DocumentType] = None) -> bool:
        """
        Save the document to a file.

        Args:
            file_path: Path to save the document to. If None, uses the current file_path.
            format: Document format to save as. If None, infers from file extension.

        Returns:
            bool: True if save was successful, False otherwise.
        """
        if file_path:
            self.file_path = str(file_path) if isinstance(file_path, Path) else file_path
            path = Path(self.file_path)
            self.title = path.stem

            # Update document type based on file extension if not explicitly provided
            if format is None:
                self.document_type = DocumentType.from_extension(path.suffix)

        if not self.file_path:
            return False

        path = Path(self.file_path)
        path.parent.mkdir(parents=True, exist_ok=True)

        try:
            # Update modification time
            self.modified_at = datetime.now()

            # Try to use format manager for supported formats
            extension = path.suffix.lower().lstrip('.')
            if format_manager.can_export(extension):
                success = format_manager.export_document(self, str(path))
                if success:
                    self.modified = False
                return success

            # Fallback to basic text-based save for unsupported formats
            if self.document_type == DocumentType.TEXT:
                with open(path, 'w', encoding=self._encoding) as f:
                    f.write(self.content)
            else:
                # For other formats without handlers, save as JSON
                with open(path.with_suffix('.json'), 'w', encoding=self._encoding) as f:
                    json.dump(self.to_dict(), f, indent=2, ensure_ascii=False)

            self.modified = False
            return True

        except Exception as e:
            # TODO: Log error properly
            print(f"Error saving document: {e}")
            return False
    
    def load(self, file_path: Union[str, Path]) -> bool:
        """
        Load document content from a file.

        Args:
            file_path: Path to the file to load.

        Returns:
            bool: True if load was successful, False otherwise.
        """
        try:
            path = Path(file_path) if isinstance(file_path, str) else file_path
            self.file_path = str(path)
            self.title = path.stem

            # Determine document type from file extension
            self.document_type = DocumentType.from_extension(path.suffix)

            # Try to use format manager for supported formats
            extension = path.suffix.lower().lstrip('.')
            if format_manager.can_import(extension):
                data = format_manager.import_document(str(path))
                if data:
                    # Update document with imported data
                    self.content = data.get('content', '')
                    self.title = data.get('title', path.stem)
                    self.author = data.get('author', '')
                    self.keywords = data.get('keywords', [])
                    self.metadata = data.get('metadata', {})

                    # Update timestamps if provided
                    if 'created_at' in data:
                        try:
                            self.created_at = datetime.fromisoformat(data['created_at'])
                        except (ValueError, TypeError) as e:
                            print(f"Warning: Failed to parse created_at timestamp: {e}")
                    if 'modified_at' in data:
                        try:
                            self.modified_at = datetime.fromisoformat(data['modified_at'])
                        except (ValueError, TypeError) as e:
                            print(f"Warning: Failed to parse modified_at timestamp: {e}")

                    # Create a default section with the content
                    self.sections = [DocumentSection(name="Main Content", content=self.content)]

                    self.modified = False
                    return True

            # Fallback to basic text loading for unsupported formats
            if self.document_type == DocumentType.TEXT:
                with open(path, 'r', encoding=self._encoding) as f:
                    self.content = f.read()
                self.sections = [DocumentSection(name="Main Content", content=self.content)]
            else:
                # Try to load as JSON
                json_path = path if path.suffix.lower() == '.json' else path.with_suffix('.json')
                if json_path.exists():
                    with open(json_path, 'r', encoding=self._encoding) as f:
                        data = json.load(f)

                    # Create document from dictionary
                    loaded_doc = self.from_dict(data)

                    # Update current document with loaded data
                    self.__dict__.update(loaded_doc.__dict__)
                else:
                    # If no JSON file exists, treat as plain text
                    with open(path, 'r', encoding=self._encoding, errors='ignore') as f:
                        self.content = f.read()
                    self.sections = [DocumentSection(name="Main Content", content=self.content)]

            self.modified = False
            self.modified_at = datetime.now()
            return True

        except Exception as e:
            # TODO: Log error properly
            print(f"Error loading document: {e}")
            return False
    
    @property
    def header_footer_manager(self) -> 'HeaderFooterManager':
        """Get the header/footer manager for this document."""
        if self._header_footer_manager is None:
            from ..features.headers_footers import HeaderFooterManager
            self._header_footer_manager = HeaderFooterManager(None)  # Document will be set when loaded in UI
        return self._header_footer_manager

    @property
    def note_manager(self) -> 'NoteManager':
        """Get the note manager for this document."""
        if self._note_manager is None:
            from ..features.footnotes_endnotes import NoteManager
            self._note_manager = NoteManager()
        return self._note_manager
        
    def get_metadata(self) -> Dict[str, Any]:
        """
        Get document metadata and statistics.
        
        Returns:
            Dict containing document metadata and statistics.
        """
        # Calculate word and character counts
        word_count = len(self.content.split())
        char_count = len(self.content)
        
        # Calculate section stats
        section_count = len(self.sections)
        
        # Calculate content stats from all sections
        for section in self.sections:
            word_count += len(section.content.split())
            char_count += len(section.content)
        
        return {
            'title': self.title,
            'file_path': self.file_path,
            'created_at': self.created_at,
            'modified_at': self.modified_at,
            'document_type': self.document_type.value,
            'word_count': word_count,
            'char_count': char_count,
            'section_count': section_count,
            'version': self.version,
            'author': self.author,
            'keywords': self.keywords,
            'template': self.template,
            'read_only': self.read_only,
            **self.metadata
        }

class DocumentManager:
    """
    Manages multiple documents in the application with support for:
    - Multiple open documents
    - Document switching
    - Document state management
    - Recent documents list
    """
    def __init__(self, max_recent_docs: int = 10):
        self.documents: List[Document] = []
        self.current_document_index: int = -1
        self.recent_documents: List[Dict[str, Any]] = []
        self.max_recent_docs = max_recent_docs
        self._listeners = []
    
    def add_listener(self, callback):
        """
        Add a callback to be notified of document changes.

        Args:
            callback: A callable that will be invoked with (event_type, **kwargs) when document events occur.
        """
        if callback not in self._listeners:
            self._listeners.append(callback)
    
    def remove_listener(self, callback):
        """
        Remove a document change listener.

        Args:
            callback: The callback function to remove from the listeners list.
        """
        if callback in self._listeners:
            self._listeners.remove(callback)
    
    def _notify_listeners(self, event_type: str, **kwargs):
        """
        Notify all listeners of a document event.

        Args:
            event_type: String identifying the type of event (e.g., 'document_created', 'document_saved').
            **kwargs: Additional event-specific data to pass to listeners.
        """
        for callback in self._listeners:
            try:
                callback(event_type, **kwargs)
            except Exception as e:
                print(f"Error in document listener: {e}")
    
    @property
    def current_document(self) -> Optional[Document]:
        """Get the current active document."""
        if 0 <= self.current_document_index < len(self.documents):
            return self.documents[self.current_document_index]
        return None
    
    @property
    def has_unsaved_changes(self) -> bool:
        """Check if any open documents have unsaved changes."""
        return any(doc.modified for doc in self.documents)
    
    @property
    def document_count(self) -> int:
        """Get the number of open documents."""
        return len(self.documents)
    
    def create_document(self, title: str = "Untitled", template: Optional[str] = None) -> Document:
        """
        Create a new document.
        
        Args:
            title: Title of the new document
            template: Optional path to a template file to use
            
        Returns:
            The newly created Document instance
        """
        doc = Document(title=title)
        
        # Apply template if provided
        if template:
            try:
                template_doc = Document()
                if template_doc.load(template):
                    # Copy template content and metadata
                    doc.content = template_doc.content
                    doc.sections = template_doc.sections.copy()
                    doc.metadata = template_doc.metadata.copy()
                    doc.document_type = template_doc.document_type
            except Exception as e:
                print(f"Error applying template: {e}")
        
        self.documents.append(doc)
        self.current_document_index = len(self.documents) - 1
        self._notify_listeners('document_created', document=doc)
        return doc
    
    def open_document(self, file_path: Union[str, Path], 
                     create_if_not_exists: bool = False) -> Optional[Document]:
        """
        Open an existing document.
        
        Args:
            file_path: Path to the document to open
            create_if_not_exists: If True, create a new document if the file doesn't exist
            
        Returns:
            The opened Document instance, or None if failed
        """
        path = Path(file_path) if isinstance(file_path, str) else file_path
        
        # Check if document is already open
        for i, doc in enumerate(self.documents):
            if doc.file_path and Path(doc.file_path).resolve() == path.resolve():
                self.current_document_index = i
                self._notify_listeners('document_activated', document=doc)
                return doc
        
        # Create new document if it doesn't exist and create_if_not_exists is True
        if not path.exists():
            if create_if_not_exists:
                doc = self.create_document(title=path.stem)
                doc.file_path = str(path)
                return doc
            return None
        
        # Try to load the document
        doc = Document(file_path=str(path))
        if doc.load(path):
            self.documents.append(doc)
            self.current_document_index = len(self.documents) - 1
            self._add_to_recent(doc)
            self._notify_listeners('document_opened', document=doc)
            return doc
        return None
    
    def _add_to_recent(self, doc: Document):
        """
        Add a document to the recent documents list.

        Maintains a list of recently opened documents, with the most recent at the beginning.
        Automatically removes duplicates and trims the list to max_recent_docs size.

        Args:
            doc: The Document instance to add to the recent list.
        """
        if not doc.file_path:
            return

        # Remove if already in the list
        self.recent_documents = [d for d in self.recent_documents
                               if d.get('path') != doc.file_path]

        # Add to the beginning of the list
        self.recent_documents.insert(0, {
            'path': doc.file_path,
            'title': doc.title,
            'timestamp': datetime.now().isoformat()
        })

        # Trim the list if it's too long
        if len(self.recent_documents) > self.max_recent_docs:
            self.recent_documents = self.recent_documents[:self.max_recent_docs]
    
    def close_document(self, index = None,
                      force: bool = False) -> bool:
        """
        Close a document by index, document object, or the current document.

        Args:
            index: Index of the document to close, a Document object, or None to close the current document.
            force: If True, close without checking for unsaved changes.

        Returns:
            bool: True if document was closed, False if operation was cancelled.
        """
        # Handle different input types
        if index is None:
            index = self.current_document_index
        elif not isinstance(index, int):
            # Assume it's a Document object, find its index
            try:
                index = self.documents.index(index)
            except (ValueError, AttributeError):
                return False

        if not 0 <= index < len(self.documents):
            return False
        
        doc = self.documents[index]
        
        # Check for unsaved changes if not forcing
        if not force and doc.modified:
            # Notify listeners to handle the unsaved changes (e.g., show a dialog)
            result = self._notify_listeners('document_closing', document=doc, index=index)
            if result is False:  # If any listener returns False, cancel the close
                return False
        
        # Remove the document
        self.documents.pop(index)
        
        # Update current document index
        if self.current_document_index >= len(self.documents):
            self.current_document_index = max(0, len(self.documents) - 1)
        
        self._notify_listeners('document_closed', document=doc, index=index)
        return True
    
    def close_all_documents(self, force: bool = False) -> bool:
        """
        Close all open documents.
        
        Args:
            force: If True, close all documents without saving.
            
        Returns:
            bool: True if all documents were closed, False if operation was cancelled.
        """
        # First check if we can close all documents
        if not force:
            for doc in self.documents:
                if doc.modified:
                    result = self._notify_listeners('document_closing', document=doc, all_docs=True)
                    if result is False:
                        return False
        
        # If we get here, either force is True or user confirmed to close all
        self.documents.clear()
        self.current_document_index = -1
        self._notify_listeners('all_documents_closed')
        return True
    
    def save_document(self, index: Optional[int] = None) -> bool:
        """
        Save a document by index or the current document.
        
        Args:
            index: Index of the document to save. If None, saves the current document.
            
        Returns:
            bool: True if save was successful, False otherwise.
        """
        doc = None
        if index is not None and 0 <= index < len(self.documents):
            doc = self.documents[index]
        else:
            doc = self.current_document
        
        if not doc:
            return False
        
        # If document has no file path, treat as Save As
        if not doc.file_path:
            return self.save_document_as("", index)
        
        # Save the document
        if doc.save():
            self._notify_listeners('document_saved', document=doc)
            return True
        return False
    
    def save_document_as(self, file_path: str, index: Optional[int] = None) -> bool:
        """
        Save a document with a new file path.
        
        Args:
            file_path: New file path to save the document to.
            index: Index of the document to save. If None, saves the current document.
            
        Returns:
            bool: True if save was successful, False otherwise.
        """
        doc = None
        if index is not None and 0 <= index < len(self.documents):
            doc = self.documents[index]
        else:
            doc = self.current_document
        
        if not doc:
            return False
        
        # If no file path provided, notify listeners to show save dialog
        if not file_path:
            result = self._notify_listeners('save_as_requested', document=doc)
            if result and isinstance(result, str):
                file_path = result
            else:
                return False
        
        # Save the document
        if doc.save(file_path):
            self._add_to_recent(doc)
            self._notify_listeners('document_saved_as', document=doc)
            return True
        return False
    
    def get_recent_documents(self, max_docs: Optional[int] = None) -> List[Dict[str, Any]]:
        """
        Get the list of recently opened documents.
        
        Args:
            max_docs: Maximum number of recent documents to return. If None, returns all.
            
        Returns:
            List of dictionaries containing recent document info.
        """
        if max_docs is not None:
            return self.recent_documents[:max_docs]
        return self.recent_documents
    
    def clear_recent_documents(self):
        """
        Clear the list of recently opened documents.

        Removes all entries from the recent documents list and notifies listeners.
        """
        self.recent_documents = []
        self._notify_listeners('recent_documents_cleared')
