"""
Security System for PyWord.

This module provides comprehensive document security features including:
- Document protection (restrict editing, formatting)
- Password protection for documents
- Digital signatures
- Information rights management (IRM)
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                               QPushButton, QComboBox, QCheckBox, QGroupBox,
                               QMessageBox, QTextEdit, QListWidget, QListWidgetItem,
                               QInputDialog, QTableWidget, QTableWidgetItem, QHeaderView,
                               QRadioButton, QButtonGroup, QWidget, QFormLayout,
                               QSpinBox, QDateEdit, QTabWidget)
from PySide6.QtCore import Qt, Signal, QDateTime, QDate
from PySide6.QtGui import QFont, QColor
import hashlib
import secrets
import json
import uuid
from datetime import datetime, timedelta
from enum import Enum
from typing import Optional, List, Dict


class ProtectionType(Enum):
    """Types of document protection."""
    NONE = "none"
    READ_ONLY = "read_only"
    COMMENTS = "comments"  # Only comments allowed
    TRACKED_CHANGES = "tracked_changes"  # Only tracked changes
    FORMS = "forms"  # Only form fields editable
    NO_FORMATTING = "no_formatting"  # No formatting changes


class PermissionLevel(Enum):
    """Information rights management permission levels."""
    NONE = "none"
    VIEW = "view"
    EDIT = "edit"
    PRINT = "print"
    COPY = "copy"
    FULL_CONTROL = "full_control"


class DigitalSignature:
    """Represents a digital signature on a document."""

    def __init__(self, signer_name: str, certificate_id: str = None,
                 timestamp: datetime = None, reason: str = "", location: str = ""):
        self.id = str(uuid.uuid4())
        self.signer_name = signer_name
        self.certificate_id = certificate_id or str(uuid.uuid4())
        self.timestamp = timestamp or datetime.now()
        self.reason = reason
        self.location = location
        self.is_valid = True
        self.signature_data = self._generate_signature()

    def _generate_signature(self) -> str:
        """Generate a mock signature hash."""
        data = f"{self.signer_name}{self.certificate_id}{self.timestamp.isoformat()}"
        return hashlib.sha256(data.encode()).hexdigest()

    def to_dict(self) -> dict:
        """Convert signature to dictionary."""
        return {
            'id': self.id,
            'signer_name': self.signer_name,
            'certificate_id': self.certificate_id,
            'timestamp': self.timestamp.isoformat(),
            'reason': self.reason,
            'location': self.location,
            'is_valid': self.is_valid,
            'signature_data': self.signature_data
        }

    @staticmethod
    def from_dict(data: dict) -> 'DigitalSignature':
        """Create signature from dictionary."""
        sig = DigitalSignature(
            data['signer_name'],
            data['certificate_id'],
            datetime.fromisoformat(data['timestamp']),
            data.get('reason', ''),
            data.get('location', '')
        )
        sig.id = data['id']
        sig.is_valid = data.get('is_valid', True)
        sig.signature_data = data['signature_data']
        return sig


class UserPermission:
    """Represents permissions for a specific user or group."""

    def __init__(self, user_id: str, user_name: str, permissions: List[PermissionLevel]):
        self.user_id = user_id
        self.user_name = user_name
        self.permissions = permissions
        self.expiry_date = None
        self.granted_by = None
        self.granted_date = datetime.now()

    def has_permission(self, permission: PermissionLevel) -> bool:
        """Check if user has a specific permission."""
        if self.expiry_date and datetime.now() > self.expiry_date:
            return False
        return permission in self.permissions or PermissionLevel.FULL_CONTROL in self.permissions

    def to_dict(self) -> dict:
        """Convert to dictionary."""
        return {
            'user_id': self.user_id,
            'user_name': self.user_name,
            'permissions': [p.value for p in self.permissions],
            'expiry_date': self.expiry_date.isoformat() if self.expiry_date else None,
            'granted_by': self.granted_by,
            'granted_date': self.granted_date.isoformat()
        }

    @staticmethod
    def from_dict(data: dict) -> 'UserPermission':
        """Create from dictionary."""
        perm = UserPermission(
            data['user_id'],
            data['user_name'],
            [PermissionLevel(p) for p in data['permissions']]
        )
        if data.get('expiry_date'):
            perm.expiry_date = datetime.fromisoformat(data['expiry_date'])
        perm.granted_by = data.get('granted_by')
        perm.granted_date = datetime.fromisoformat(data['granted_date'])
        return perm


class SecurityManager:
    """Manages document security features."""

    def __init__(self, parent):
        self.parent = parent
        self.protection_type = ProtectionType.NONE
        self.protection_password_hash = None
        self.protection_salt = None

        # Password protection
        self.is_password_protected = False
        self.password_hash = None
        self.password_salt = None
        self.password_hint = ""

        # Digital signatures
        self.signatures: List[DigitalSignature] = []

        # Information Rights Management
        self.irm_enabled = False
        self.irm_permissions: List[UserPermission] = []
        self.irm_owner = None
        self.irm_policy_name = ""
        self.irm_policy_description = ""

        # Encryption
        self.encryption_algorithm = "AES-256"
        self.is_encrypted = False

    # Document Protection Methods
    def protect_document(self, protection_type: ProtectionType, password: str = None) -> bool:
        """Protect the document with specified protection type."""
        self.protection_type = protection_type

        if password:
            self.protection_salt = secrets.token_bytes(32)
            self.protection_password_hash = self._hash_password(password, self.protection_salt)

        self._apply_protection()
        return True

    def unprotect_document(self, password: str = None) -> bool:
        """Remove document protection."""
        if self.protection_password_hash:
            if not password:
                return False
            if not self._verify_password(password, self.protection_password_hash,
                                        self.protection_salt):
                return False

        self.protection_type = ProtectionType.NONE
        self.protection_password_hash = None
        self.protection_salt = None
        self._remove_protection()
        return True

    def _apply_protection(self):
        """Apply protection settings to the document."""
        if self.protection_type == ProtectionType.READ_ONLY:
            if hasattr(self.parent, 'text_edit'):
                self.parent.text_edit.setReadOnly(True)
        elif self.protection_type == ProtectionType.NO_FORMATTING:
            # Disable formatting controls
            pass

    def _remove_protection(self):
        """Remove protection from document."""
        if hasattr(self.parent, 'text_edit'):
            self.parent.text_edit.setReadOnly(False)

    # Password Protection Methods
    def set_password(self, password: str, hint: str = "") -> bool:
        """Set a password for the document."""
        if not password:
            return False

        self.password_salt = secrets.token_bytes(32)
        self.password_hash = self._hash_password(password, self.password_salt)
        self.password_hint = hint
        self.is_password_protected = True
        return True

    def verify_password(self, password: str) -> bool:
        """Verify the document password."""
        if not self.is_password_protected:
            return True
        return self._verify_password(password, self.password_hash, self.password_salt)

    def remove_password(self, password: str) -> bool:
        """Remove password protection."""
        if not self.verify_password(password):
            return False

        self.password_hash = None
        self.password_salt = None
        self.password_hint = ""
        self.is_password_protected = False
        return True

    def _hash_password(self, password: str, salt: bytes) -> str:
        """Hash a password with salt using PBKDF2."""
        return hashlib.pbkdf2_hmac('sha256', password.encode(), salt, 100000).hex()

    def _verify_password(self, password: str, hash_value: str, salt: bytes) -> bool:
        """Verify a password against its hash."""
        return self._hash_password(password, salt) == hash_value

    # Digital Signature Methods
    def add_signature(self, signer_name: str, reason: str = "",
                     location: str = "") -> DigitalSignature:
        """Add a digital signature to the document."""
        signature = DigitalSignature(signer_name, reason=reason, location=location)
        self.signatures.append(signature)
        return signature

    def remove_signature(self, signature_id: str) -> bool:
        """Remove a digital signature."""
        self.signatures = [s for s in self.signatures if s.id != signature_id]
        return True

    def verify_signature(self, signature_id: str) -> bool:
        """Verify a digital signature."""
        signature = next((s for s in self.signatures if s.id == signature_id), None)
        if not signature:
            return False

        # In a real implementation, this would verify against a certificate authority
        # For now, we just check if the signature data is consistent
        expected_sig = signature._generate_signature()
        return signature.signature_data == expected_sig

    def get_all_signatures(self) -> List[DigitalSignature]:
        """Get all signatures in the document."""
        return self.signatures.copy()

    def invalidate_signatures(self):
        """Invalidate all signatures (called when document is modified)."""
        for signature in self.signatures:
            signature.is_valid = False

    # Information Rights Management Methods
    def enable_irm(self, owner: str, policy_name: str = "",
                   policy_description: str = "") -> bool:
        """Enable Information Rights Management."""
        self.irm_enabled = True
        self.irm_owner = owner
        self.irm_policy_name = policy_name
        self.irm_policy_description = policy_description
        return True

    def disable_irm(self) -> bool:
        """Disable Information Rights Management."""
        self.irm_enabled = False
        self.irm_permissions.clear()
        return True

    def grant_permission(self, user_id: str, user_name: str,
                        permissions: List[PermissionLevel],
                        expiry_date: datetime = None) -> UserPermission:
        """Grant permissions to a user."""
        # Remove existing permissions for this user
        self.irm_permissions = [p for p in self.irm_permissions if p.user_id != user_id]

        # Add new permissions
        user_perm = UserPermission(user_id, user_name, permissions)
        user_perm.expiry_date = expiry_date
        user_perm.granted_by = self.irm_owner
        self.irm_permissions.append(user_perm)
        return user_perm

    def revoke_permission(self, user_id: str) -> bool:
        """Revoke all permissions from a user."""
        self.irm_permissions = [p for p in self.irm_permissions if p.user_id != user_id]
        return True

    def check_permission(self, user_id: str, permission: PermissionLevel) -> bool:
        """Check if a user has a specific permission."""
        if not self.irm_enabled:
            return True

        user_perm = next((p for p in self.irm_permissions if p.user_id == user_id), None)
        if not user_perm:
            return False

        return user_perm.has_permission(permission)

    def get_user_permissions(self, user_id: str) -> Optional[UserPermission]:
        """Get permissions for a specific user."""
        return next((p for p in self.irm_permissions if p.user_id == user_id), None)

    # Serialization Methods
    def to_dict(self) -> dict:
        """Convert security settings to dictionary."""
        return {
            'protection_type': self.protection_type.value,
            'protection_password_hash': self.protection_password_hash,
            'protection_salt': self.protection_salt.hex() if self.protection_salt else None,
            'is_password_protected': self.is_password_protected,
            'password_hash': self.password_hash,
            'password_salt': self.password_salt.hex() if self.password_salt else None,
            'password_hint': self.password_hint,
            'signatures': [s.to_dict() for s in self.signatures],
            'irm_enabled': self.irm_enabled,
            'irm_permissions': [p.to_dict() for p in self.irm_permissions],
            'irm_owner': self.irm_owner,
            'irm_policy_name': self.irm_policy_name,
            'irm_policy_description': self.irm_policy_description,
            'encryption_algorithm': self.encryption_algorithm,
            'is_encrypted': self.is_encrypted
        }

    def from_dict(self, data: dict):
        """Load security settings from dictionary."""
        self.protection_type = ProtectionType(data.get('protection_type', 'none'))
        self.protection_password_hash = data.get('protection_password_hash')
        salt_hex = data.get('protection_salt')
        self.protection_salt = bytes.fromhex(salt_hex) if salt_hex else None

        self.is_password_protected = data.get('is_password_protected', False)
        self.password_hash = data.get('password_hash')
        pwd_salt_hex = data.get('password_salt')
        self.password_salt = bytes.fromhex(pwd_salt_hex) if pwd_salt_hex else None
        self.password_hint = data.get('password_hint', '')

        self.signatures = [DigitalSignature.from_dict(s) for s in data.get('signatures', [])]

        self.irm_enabled = data.get('irm_enabled', False)
        self.irm_permissions = [UserPermission.from_dict(p)
                               for p in data.get('irm_permissions', [])]
        self.irm_owner = data.get('irm_owner')
        self.irm_policy_name = data.get('irm_policy_name', '')
        self.irm_policy_description = data.get('irm_policy_description', '')
        self.encryption_algorithm = data.get('encryption_algorithm', 'AES-256')
        self.is_encrypted = data.get('is_encrypted', False)

    def get_security_info(self) -> dict:
        """Get summary of security settings."""
        return {
            'protected': self.protection_type != ProtectionType.NONE,
            'protection_type': self.protection_type.value,
            'password_protected': self.is_password_protected,
            'signature_count': len(self.signatures),
            'valid_signature_count': sum(1 for s in self.signatures if s.is_valid),
            'irm_enabled': self.irm_enabled,
            'irm_user_count': len(self.irm_permissions),
            'encrypted': self.is_encrypted
        }
