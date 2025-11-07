"""
Automation System for PyWord.

This module provides automation and extensibility features including:
- Macros (recorded and scripted actions)
- Add-ins support (plugins and extensions)
- Custom XML data storage
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                               QPushButton, QTextEdit, QListWidget, QListWidgetItem,
                               QMessageBox, QComboBox, QCheckBox, QGroupBox,
                               QTableWidget, QTableWidgetItem, QHeaderView,
                               QTabWidget, QWidget, QSplitter, QFileDialog,
                               QTreeWidget, QTreeWidgetItem, QSpinBox, QFormLayout)
from PySide6.QtCore import Qt, Signal, QObject, QTimer
from PySide6.QtGui import QFont, QColor, QTextCursor, QKeySequence
import json
import uuid
from datetime import datetime
from typing import Dict, List, Any, Optional, Callable
from enum import Enum
import re
import xml.etree.ElementTree as ET
from xml.dom import minidom


class MacroType(Enum):
    """Types of macros."""
    RECORDED = "recorded"  # Recorded user actions
    SCRIPTED = "scripted"  # Python scripts
    KEYBOARD_SHORTCUT = "keyboard_shortcut"


class ActionType(Enum):
    """Types of recorded actions."""
    INSERT_TEXT = "insert_text"
    DELETE_TEXT = "delete_text"
    FORMAT_BOLD = "format_bold"
    FORMAT_ITALIC = "format_italic"
    FORMAT_UNDERLINE = "format_underline"
    CHANGE_FONT = "change_font"
    CHANGE_SIZE = "change_size"
    CHANGE_COLOR = "change_color"
    INSERT_PARAGRAPH = "insert_paragraph"
    MOVE_CURSOR = "move_cursor"
    SELECT_TEXT = "select_text"
    CUSTOM = "custom"


class MacroAction:
    """Represents a single action in a macro."""

    def __init__(self, action_type: ActionType, parameters: Dict[str, Any] = None):
        self.action_type = action_type
        self.parameters = parameters or {}
        self.timestamp = datetime.now()

    def to_dict(self) -> dict:
        """Convert action to dictionary."""
        return {
            'action_type': self.action_type.value,
            'parameters': self.parameters,
            'timestamp': self.timestamp.isoformat()
        }

    @staticmethod
    def from_dict(data: dict) -> 'MacroAction':
        """Create action from dictionary."""
        action = MacroAction(
            ActionType(data['action_type']),
            data['parameters']
        )
        action.timestamp = datetime.fromisoformat(data['timestamp'])
        return action


class Macro:
    """Represents a macro (sequence of actions or a script)."""

    def __init__(self, name: str, macro_type: MacroType = MacroType.RECORDED):
        self.id = str(uuid.uuid4())
        self.name = name
        self.macro_type = macro_type
        self.description = ""
        self.actions: List[MacroAction] = []
        self.script = ""
        self.created_date = datetime.now()
        self.modified_date = datetime.now()
        self.run_count = 0
        self.keyboard_shortcut = ""
        self.enabled = True
        self.auto_run_on_open = False
        self.auto_run_on_save = False

    def add_action(self, action: MacroAction):
        """Add an action to the macro."""
        self.actions.append(action)
        self.modified_date = datetime.now()

    def to_dict(self) -> dict:
        """Convert macro to dictionary."""
        return {
            'id': self.id,
            'name': self.name,
            'macro_type': self.macro_type.value,
            'description': self.description,
            'actions': [a.to_dict() for a in self.actions],
            'script': self.script,
            'created_date': self.created_date.isoformat(),
            'modified_date': self.modified_date.isoformat(),
            'run_count': self.run_count,
            'keyboard_shortcut': self.keyboard_shortcut,
            'enabled': self.enabled,
            'auto_run_on_open': self.auto_run_on_open,
            'auto_run_on_save': self.auto_run_on_save
        }

    @staticmethod
    def from_dict(data: dict) -> 'Macro':
        """Create macro from dictionary."""
        macro = Macro(data['name'], MacroType(data['macro_type']))
        macro.id = data['id']
        macro.description = data.get('description', '')
        macro.actions = [MacroAction.from_dict(a) for a in data.get('actions', [])]
        macro.script = data.get('script', '')
        macro.created_date = datetime.fromisoformat(data['created_date'])
        macro.modified_date = datetime.fromisoformat(data['modified_date'])
        macro.run_count = data.get('run_count', 0)
        macro.keyboard_shortcut = data.get('keyboard_shortcut', '')
        macro.enabled = data.get('enabled', True)
        macro.auto_run_on_open = data.get('auto_run_on_open', False)
        macro.auto_run_on_save = data.get('auto_run_on_save', False)
        return macro


class AddIn:
    """Represents an add-in/plugin."""

    def __init__(self, name: str, version: str = "1.0.0"):
        self.id = str(uuid.uuid4())
        self.name = name
        self.version = version
        self.description = ""
        self.author = ""
        self.enabled = True
        self.loaded = False
        self.install_date = datetime.now()
        self.file_path = ""
        self.entry_point = ""
        self.dependencies: List[str] = []
        self.permissions: List[str] = []
        self.settings: Dict[str, Any] = {}

    def to_dict(self) -> dict:
        """Convert add-in to dictionary."""
        return {
            'id': self.id,
            'name': self.name,
            'version': self.version,
            'description': self.description,
            'author': self.author,
            'enabled': self.enabled,
            'loaded': self.loaded,
            'install_date': self.install_date.isoformat(),
            'file_path': self.file_path,
            'entry_point': self.entry_point,
            'dependencies': self.dependencies,
            'permissions': self.permissions,
            'settings': self.settings
        }

    @staticmethod
    def from_dict(data: dict) -> 'AddIn':
        """Create add-in from dictionary."""
        addon = AddIn(data['name'], data['version'])
        addon.id = data['id']
        addon.description = data.get('description', '')
        addon.author = data.get('author', '')
        addon.enabled = data.get('enabled', True)
        addon.loaded = data.get('loaded', False)
        addon.install_date = datetime.fromisoformat(data['install_date'])
        addon.file_path = data.get('file_path', '')
        addon.entry_point = data.get('entry_point', '')
        addon.dependencies = data.get('dependencies', [])
        addon.permissions = data.get('permissions', [])
        addon.settings = data.get('settings', {})
        return addon


class CustomXMLPart:
    """Represents a custom XML data part."""

    def __init__(self, name: str, namespace: str = ""):
        self.id = str(uuid.uuid4())
        self.name = name
        self.namespace = namespace or f"http://pyword.example.com/{name}"
        self.xml_content = self._create_empty_xml()
        self.created_date = datetime.now()
        self.modified_date = datetime.now()
        self.schema_location = ""

    def _create_empty_xml(self) -> str:
        """Create an empty XML structure."""
        root = ET.Element("data", attrib={"xmlns": self.namespace})
        return self._prettify_xml(root)

    def _prettify_xml(self, elem: ET.Element) -> str:
        """Return a pretty-printed XML string."""
        rough_string = ET.tostring(elem, encoding='unicode')
        reparsed = minidom.parseString(rough_string)
        return reparsed.toprettyxml(indent="  ")

    def set_xml_content(self, xml_content: str) -> bool:
        """Set the XML content, validating it first."""
        try:
            ET.fromstring(xml_content)
            self.xml_content = xml_content
            self.modified_date = datetime.now()
            return True
        except ET.ParseError as e:
            return False

    def get_xml_content(self) -> str:
        """Get the XML content."""
        return self.xml_content

    def add_element(self, parent_path: str, element_name: str,
                   text: str = "", attributes: Dict[str, str] = None) -> bool:
        """Add an element to the XML."""
        try:
            root = ET.fromstring(self.xml_content)
            parent = root.find(parent_path) if parent_path else root

            if parent is None:
                return False

            new_element = ET.SubElement(parent, element_name, attrib=attributes or {})
            if text:
                new_element.text = text

            self.xml_content = self._prettify_xml(root)
            self.modified_date = datetime.now()
            return True
        except Exception as e:
            return False

    def update_element(self, path: str, text: str = None,
                      attributes: Dict[str, str] = None) -> bool:
        """Update an element in the XML."""
        try:
            root = ET.fromstring(self.xml_content)
            element = root.find(path)

            if element is None:
                return False

            if text is not None:
                element.text = text

            if attributes:
                for key, value in attributes.items():
                    element.set(key, value)

            self.xml_content = self._prettify_xml(root)
            self.modified_date = datetime.now()
            return True
        except Exception as e:
            return False

    def remove_element(self, path: str) -> bool:
        """Remove an element from the XML."""
        try:
            root = ET.fromstring(self.xml_content)
            parent_path = '/'.join(path.split('/')[:-1])
            element_name = path.split('/')[-1]

            parent = root.find(parent_path) if parent_path else root
            if parent is None:
                return False

            element = parent.find(element_name)
            if element is not None:
                parent.remove(element)
                self.xml_content = self._prettify_xml(root)
                self.modified_date = datetime.now()
                return True

            return False
        except Exception as e:
            return False

    def query_elements(self, xpath: str) -> List[ET.Element]:
        """Query elements using XPath."""
        try:
            root = ET.fromstring(self.xml_content)
            return root.findall(xpath)
        except Exception as e:
            return []

    def to_dict(self) -> dict:
        """Convert to dictionary."""
        return {
            'id': self.id,
            'name': self.name,
            'namespace': self.namespace,
            'xml_content': self.xml_content,
            'created_date': self.created_date.isoformat(),
            'modified_date': self.modified_date.isoformat(),
            'schema_location': self.schema_location
        }

    @staticmethod
    def from_dict(data: dict) -> 'CustomXMLPart':
        """Create from dictionary."""
        xml_part = CustomXMLPart(data['name'], data['namespace'])
        xml_part.id = data['id']
        xml_part.xml_content = data['xml_content']
        xml_part.created_date = datetime.fromisoformat(data['created_date'])
        xml_part.modified_date = datetime.fromisoformat(data['modified_date'])
        xml_part.schema_location = data.get('schema_location', '')
        return xml_part


class AutomationManager:
    """Manages automation features."""

    def __init__(self, parent):
        self.parent = parent
        self.macros: Dict[str, Macro] = {}
        self.addins: Dict[str, AddIn] = {}
        self.xml_parts: Dict[str, CustomXMLPart] = {}

        # Macro recording
        self.is_recording = False
        self.current_recording: Optional[Macro] = None

        # Add-in hooks
        self.addon_hooks: Dict[str, List[Callable]] = {
            'on_document_open': [],
            'on_document_save': [],
            'on_document_close': [],
            'on_text_insert': [],
            'on_text_delete': [],
            'on_format_change': []
        }

    # Macro Management Methods
    def create_macro(self, name: str, macro_type: MacroType = MacroType.RECORDED) -> Macro:
        """Create a new macro."""
        macro = Macro(name, macro_type)
        self.macros[macro.id] = macro
        return macro

    def delete_macro(self, macro_id: str) -> bool:
        """Delete a macro."""
        if macro_id in self.macros:
            del self.macros[macro_id]
            return True
        return False

    def get_macro(self, macro_id: str) -> Optional[Macro]:
        """Get a macro by ID."""
        return self.macros.get(macro_id)

    def get_all_macros(self) -> List[Macro]:
        """Get all macros."""
        return list(self.macros.values())

    def start_recording(self, name: str) -> Macro:
        """Start recording a macro."""
        self.current_recording = self.create_macro(name, MacroType.RECORDED)
        self.is_recording = True
        return self.current_recording

    def stop_recording(self) -> Optional[Macro]:
        """Stop recording a macro."""
        self.is_recording = False
        recorded_macro = self.current_recording
        self.current_recording = None
        return recorded_macro

    def record_action(self, action_type: ActionType, parameters: Dict[str, Any] = None):
        """Record an action to the current macro."""
        if self.is_recording and self.current_recording:
            action = MacroAction(action_type, parameters)
            self.current_recording.add_action(action)

    def run_macro(self, macro_id: str) -> bool:
        """Run a macro."""
        macro = self.get_macro(macro_id)
        if not macro or not macro.enabled:
            return False

        if macro.macro_type == MacroType.SCRIPTED:
            return self._run_scripted_macro(macro)
        else:
            return self._run_recorded_macro(macro)

    def _run_recorded_macro(self, macro: Macro) -> bool:
        """Execute a recorded macro."""
        try:
            for action in macro.actions:
                self._execute_action(action)
            macro.run_count += 1
            macro.modified_date = datetime.now()
            return True
        except Exception as e:
            return False

    def _run_scripted_macro(self, macro: Macro) -> bool:
        """Execute a scripted macro."""
        try:
            # Create a safe execution environment
            safe_globals = {
                'parent': self.parent,
                '__builtins__': {
                    'print': print,
                    'len': len,
                    'str': str,
                    'int': int,
                    'float': float,
                    'bool': bool,
                    'list': list,
                    'dict': dict,
                    'range': range,
                }
            }
            exec(macro.script, safe_globals)
            macro.run_count += 1
            macro.modified_date = datetime.now()
            return True
        except Exception as e:
            return False

    def _execute_action(self, action: MacroAction):
        """Execute a single macro action."""
        if not hasattr(self.parent, 'text_edit'):
            return

        text_edit = self.parent.text_edit
        cursor = text_edit.textCursor()

        if action.action_type == ActionType.INSERT_TEXT:
            cursor.insertText(action.parameters.get('text', ''))
        elif action.action_type == ActionType.DELETE_TEXT:
            count = action.parameters.get('count', 1)
            for _ in range(count):
                cursor.deleteChar()
        elif action.action_type == ActionType.FORMAT_BOLD:
            fmt = cursor.charFormat()
            fmt.setFontWeight(QFont.Weight.Bold if action.parameters.get('enabled', True)
                            else QFont.Weight.Normal)
            cursor.setCharFormat(fmt)
        elif action.action_type == ActionType.FORMAT_ITALIC:
            fmt = cursor.charFormat()
            fmt.setFontItalic(action.parameters.get('enabled', True))
            cursor.setCharFormat(fmt)
        elif action.action_type == ActionType.FORMAT_UNDERLINE:
            fmt = cursor.charFormat()
            fmt.setFontUnderline(action.parameters.get('enabled', True))
            cursor.setCharFormat(fmt)

    # Add-in Management Methods
    def install_addon(self, addon: AddIn) -> bool:
        """Install an add-in."""
        self.addins[addon.id] = addon
        return True

    def uninstall_addon(self, addon_id: str) -> bool:
        """Uninstall an add-in."""
        if addon_id in self.addins:
            addon = self.addins[addon_id]
            if addon.loaded:
                self.disable_addon(addon_id)
            del self.addins[addon_id]
            return True
        return False

    def enable_addon(self, addon_id: str) -> bool:
        """Enable an add-in."""
        addon = self.addins.get(addon_id)
        if not addon:
            return False

        addon.enabled = True
        addon.loaded = True
        return True

    def disable_addon(self, addon_id: str) -> bool:
        """Disable an add-in."""
        addon = self.addins.get(addon_id)
        if not addon:
            return False

        addon.enabled = False
        addon.loaded = False
        return True

    def get_addon(self, addon_id: str) -> Optional[AddIn]:
        """Get an add-in by ID."""
        return self.addins.get(addon_id)

    def get_all_addins(self) -> List[AddIn]:
        """Get all add-ins."""
        return list(self.addins.values())

    def register_hook(self, hook_name: str, callback: Callable) -> bool:
        """Register a hook callback."""
        if hook_name in self.addon_hooks:
            self.addon_hooks[hook_name].append(callback)
            return True
        return False

    def trigger_hook(self, hook_name: str, *args, **kwargs):
        """Trigger all callbacks for a hook."""
        if hook_name in self.addon_hooks:
            for callback in self.addon_hooks[hook_name]:
                try:
                    callback(*args, **kwargs)
                except Exception as e:
                    pass

    # Custom XML Management Methods
    def create_xml_part(self, name: str, namespace: str = "") -> CustomXMLPart:
        """Create a new custom XML part."""
        xml_part = CustomXMLPart(name, namespace)
        self.xml_parts[xml_part.id] = xml_part
        return xml_part

    def delete_xml_part(self, part_id: str) -> bool:
        """Delete a custom XML part."""
        if part_id in self.xml_parts:
            del self.xml_parts[part_id]
            return True
        return False

    def get_xml_part(self, part_id: str) -> Optional[CustomXMLPart]:
        """Get a custom XML part by ID."""
        return self.xml_parts.get(part_id)

    def get_xml_part_by_name(self, name: str) -> Optional[CustomXMLPart]:
        """Get a custom XML part by name."""
        for part in self.xml_parts.values():
            if part.name == name:
                return part
        return None

    def get_all_xml_parts(self) -> List[CustomXMLPart]:
        """Get all custom XML parts."""
        return list(self.xml_parts.values())

    # Serialization Methods
    def to_dict(self) -> dict:
        """Convert automation settings to dictionary."""
        return {
            'macros': {mid: m.to_dict() for mid, m in self.macros.items()},
            'addins': {aid: a.to_dict() for aid, a in self.addins.items()},
            'xml_parts': {xid: x.to_dict() for xid, x in self.xml_parts.items()}
        }

    def from_dict(self, data: dict):
        """Load automation settings from dictionary."""
        self.macros = {mid: Macro.from_dict(m)
                      for mid, m in data.get('macros', {}).items()}
        self.addins = {aid: AddIn.from_dict(a)
                      for aid, a in data.get('addins', {}).items()}
        self.xml_parts = {xid: CustomXMLPart.from_dict(x)
                         for xid, x in data.get('xml_parts', {}).items()}

    def get_automation_info(self) -> dict:
        """Get summary of automation settings."""
        return {
            'macro_count': len(self.macros),
            'enabled_macro_count': sum(1 for m in self.macros.values() if m.enabled),
            'addon_count': len(self.addins),
            'loaded_addon_count': sum(1 for a in self.addins.values() if a.loaded),
            'xml_part_count': len(self.xml_parts),
            'is_recording': self.is_recording
        }
