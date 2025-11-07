"""
Security dialogs for PyWord.

Provides UI dialogs for document security features including:
- Document protection
- Password protection
- Digital signatures
- Information rights management
"""

from PySide6.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
                               QPushButton, QComboBox, QCheckBox, QGroupBox,
                               QMessageBox, QTextEdit, QListWidget, QListWidgetItem,
                               QTableWidget, QTableWidgetItem, QHeaderView,
                               QTabWidget, QWidget, QFormLayout, QDateEdit)
from PySide6.QtCore import Qt, Signal, QDate
from PySide6.QtGui import QFont
from ...features.security import (SecurityManager, ProtectionType, PermissionLevel,
                                   DigitalSignature, UserPermission)
from datetime import datetime, timedelta


class DocumentProtectionDialog(QDialog):
    """Dialog for protecting a document."""

    def __init__(self, security_manager: SecurityManager, parent=None):
        super().__init__(parent)
        self.security_manager = security_manager
        self.setWindowTitle("Protect Document")
        self.setModal(True)
        self.setMinimumWidth(500)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Protection type
        type_group = QGroupBox("Protection Type")
        type_layout = QVBoxLayout()

        self.protection_combo = QComboBox()
        self.protection_combo.addItem("None", ProtectionType.NONE)
        self.protection_combo.addItem("Read-Only", ProtectionType.READ_ONLY)
        self.protection_combo.addItem("Comments Only", ProtectionType.COMMENTS)
        self.protection_combo.addItem("Tracked Changes Only", ProtectionType.TRACKED_CHANGES)
        self.protection_combo.addItem("Forms Only", ProtectionType.FORMS)
        self.protection_combo.addItem("No Formatting Changes", ProtectionType.NO_FORMATTING)

        # Set current protection type
        for i in range(self.protection_combo.count()):
            if self.protection_combo.itemData(i) == self.security_manager.protection_type:
                self.protection_combo.setCurrentIndex(i)
                break

        type_layout.addWidget(QLabel("Select protection level:"))
        type_layout.addWidget(self.protection_combo)
        type_group.setLayout(type_layout)
        layout.addWidget(type_group)

        # Password
        password_group = QGroupBox("Password Protection (Optional)")
        password_layout = QFormLayout()

        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_edit.setPlaceholderText("Leave empty for no password")

        self.password_confirm_edit = QLineEdit()
        self.password_confirm_edit.setEchoMode(QLineEdit.EchoMode.Password)

        password_layout.addRow("Password:", self.password_edit)
        password_layout.addRow("Confirm:", self.password_confirm_edit)

        password_group.setLayout(password_layout)
        layout.addWidget(password_group)

        # Info label
        info_label = QLabel("Note: Document protection restricts editing but does not encrypt the document.\n"
                          "Use 'Set Password' for encryption.")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("QLabel { color: gray; font-style: italic; }")
        layout.addWidget(info_label)

        # Buttons
        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("OK")
        self.cancel_button = QPushButton("Cancel")
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addStretch()
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def accept(self):
        password = self.password_edit.text()
        password_confirm = self.password_confirm_edit.text()

        if password and password != password_confirm:
            QMessageBox.warning(self, "Error", "Passwords do not match!")
            return

        protection_type = self.protection_combo.currentData()

        try:
            self.security_manager.protect_document(protection_type, password if password else None)
            QMessageBox.information(self, "Success", "Document protection applied successfully!")
            super().accept()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to apply protection: {str(e)}")


class UnprotectDocumentDialog(QDialog):
    """Dialog for unprotecting a document."""

    def __init__(self, security_manager: SecurityManager, parent=None):
        super().__init__(parent)
        self.security_manager = security_manager
        self.setWindowTitle("Unprotect Document")
        self.setModal(True)
        self.setMinimumWidth(400)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        layout.addWidget(QLabel("This document is currently protected."))
        layout.addWidget(QLabel(f"Protection type: {self.security_manager.protection_type.value}"))

        if self.security_manager.protection_password_hash:
            layout.addWidget(QLabel("\nThis protection is password-protected."))
            password_layout = QFormLayout()
            self.password_edit = QLineEdit()
            self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
            password_layout.addRow("Password:", self.password_edit)
            layout.addLayout(password_layout)
        else:
            self.password_edit = None

        # Buttons
        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("Unprotect")
        self.cancel_button = QPushButton("Cancel")
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addStretch()
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def accept(self):
        password = self.password_edit.text() if self.password_edit else None

        if self.security_manager.unprotect_document(password):
            QMessageBox.information(self, "Success", "Document protection removed successfully!")
            super().accept()
        else:
            QMessageBox.critical(self, "Error", "Failed to remove protection. Incorrect password?")


class SetPasswordDialog(QDialog):
    """Dialog for setting document password."""

    def __init__(self, security_manager: SecurityManager, parent=None):
        super().__init__(parent)
        self.security_manager = security_manager
        self.setWindowTitle("Set Document Password")
        self.setModal(True)
        self.setMinimumWidth(400)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        layout.addWidget(QLabel("Encrypt this document with a password:"))

        form_layout = QFormLayout()

        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)

        self.password_confirm_edit = QLineEdit()
        self.password_confirm_edit.setEchoMode(QLineEdit.EchoMode.Password)

        self.hint_edit = QLineEdit()
        self.hint_edit.setPlaceholderText("Optional password hint")

        form_layout.addRow("Password:", self.password_edit)
        form_layout.addRow("Confirm:", self.password_confirm_edit)
        form_layout.addRow("Hint:", self.hint_edit)

        layout.addLayout(form_layout)

        # Info
        info_label = QLabel("\nNote: If you lose your password, it cannot be recovered.\n"
                          "Keep your password in a safe place.")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("QLabel { color: #cc0000; font-weight: bold; }")
        layout.addWidget(info_label)

        # Buttons
        button_layout = QHBoxLayout()
        self.ok_button = QPushButton("Set Password")
        self.cancel_button = QPushButton("Cancel")
        self.ok_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addStretch()
        button_layout.addWidget(self.ok_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def accept(self):
        password = self.password_edit.text()
        password_confirm = self.password_confirm_edit.text()
        hint = self.hint_edit.text()

        if not password:
            QMessageBox.warning(self, "Error", "Password cannot be empty!")
            return

        if password != password_confirm:
            QMessageBox.warning(self, "Error", "Passwords do not match!")
            return

        if len(password) < 6:
            QMessageBox.warning(self, "Error", "Password must be at least 6 characters long!")
            return

        if self.security_manager.set_password(password, hint):
            QMessageBox.information(self, "Success", "Password set successfully!\n\n"
                                  "This document is now encrypted.")
            super().accept()
        else:
            QMessageBox.critical(self, "Error", "Failed to set password!")


class DigitalSignatureDialog(QDialog):
    """Dialog for managing digital signatures."""

    def __init__(self, security_manager: SecurityManager, parent=None):
        super().__init__(parent)
        self.security_manager = security_manager
        self.setWindowTitle("Digital Signatures")
        self.setModal(True)
        self.setMinimumSize(600, 400)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # Signature list
        layout.addWidget(QLabel("Digital Signatures:"))

        self.signature_table = QTableWidget()
        self.signature_table.setColumnCount(5)
        self.signature_table.setHorizontalHeaderLabels(
            ["Signer", "Date", "Status", "Reason", "Location"])
        self.signature_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.signature_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)

        self.refresh_signature_list()

        layout.addWidget(self.signature_table)

        # Buttons
        button_layout = QHBoxLayout()

        self.add_button = QPushButton("Add Signature")
        self.add_button.clicked.connect(self.add_signature)

        self.remove_button = QPushButton("Remove Selected")
        self.remove_button.clicked.connect(self.remove_signature)

        self.verify_button = QPushButton("Verify Selected")
        self.verify_button.clicked.connect(self.verify_signature)

        self.close_button = QPushButton("Close")
        self.close_button.clicked.connect(self.accept)

        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.remove_button)
        button_layout.addWidget(self.verify_button)
        button_layout.addStretch()
        button_layout.addWidget(self.close_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def refresh_signature_list(self):
        self.signature_table.setRowCount(0)
        signatures = self.security_manager.get_all_signatures()

        for sig in signatures:
            row = self.signature_table.rowCount()
            self.signature_table.insertRow(row)

            self.signature_table.setItem(row, 0, QTableWidgetItem(sig.signer_name))
            self.signature_table.setItem(row, 1,
                QTableWidgetItem(sig.timestamp.strftime("%Y-%m-%d %H:%M:%S")))
            status = "Valid" if sig.is_valid else "Invalid"
            self.signature_table.setItem(row, 2, QTableWidgetItem(status))
            self.signature_table.setItem(row, 3, QTableWidgetItem(sig.reason))
            self.signature_table.setItem(row, 4, QTableWidgetItem(sig.location))

            # Store signature ID in the row
            self.signature_table.item(row, 0).setData(Qt.ItemDataRole.UserRole, sig.id)

    def add_signature(self):
        dialog = AddSignatureDialog(self.security_manager, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh_signature_list()

    def remove_signature(self):
        row = self.signature_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Error", "Please select a signature to remove!")
            return

        sig_id = self.signature_table.item(row, 0).data(Qt.ItemDataRole.UserRole)

        reply = QMessageBox.question(self, "Confirm",
            "Are you sure you want to remove this signature?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            self.security_manager.remove_signature(sig_id)
            self.refresh_signature_list()

    def verify_signature(self):
        row = self.signature_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Error", "Please select a signature to verify!")
            return

        sig_id = self.signature_table.item(row, 0).data(Qt.ItemDataRole.UserRole)

        if self.security_manager.verify_signature(sig_id):
            QMessageBox.information(self, "Verification",
                "The signature is valid and has not been tampered with.")
        else:
            QMessageBox.warning(self, "Verification",
                "The signature is invalid or the document has been modified after signing.")


class AddSignatureDialog(QDialog):
    """Dialog for adding a digital signature."""

    def __init__(self, security_manager: SecurityManager, parent=None):
        super().__init__(parent)
        self.security_manager = security_manager
        self.setWindowTitle("Add Digital Signature")
        self.setModal(True)
        self.setMinimumWidth(400)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        form_layout = QFormLayout()

        self.signer_edit = QLineEdit()
        self.signer_edit.setPlaceholderText("Your name")

        self.reason_edit = QLineEdit()
        self.reason_edit.setPlaceholderText("Reason for signing (optional)")

        self.location_edit = QLineEdit()
        self.location_edit.setPlaceholderText("Location (optional)")

        form_layout.addRow("Signer Name:", self.signer_edit)
        form_layout.addRow("Reason:", self.reason_edit)
        form_layout.addRow("Location:", self.location_edit)

        layout.addLayout(form_layout)

        # Info
        info_label = QLabel("\nBy signing this document, you certify that you have reviewed\n"
                          "and approve its contents.")
        info_label.setWordWrap(True)
        info_label.setStyleSheet("QLabel { color: gray; font-style: italic; }")
        layout.addWidget(info_label)

        # Buttons
        button_layout = QHBoxLayout()
        self.sign_button = QPushButton("Sign Document")
        self.cancel_button = QPushButton("Cancel")
        self.sign_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addStretch()
        button_layout.addWidget(self.sign_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def accept(self):
        signer_name = self.signer_edit.text().strip()

        if not signer_name:
            QMessageBox.warning(self, "Error", "Signer name is required!")
            return

        reason = self.reason_edit.text().strip()
        location = self.location_edit.text().strip()

        self.security_manager.add_signature(signer_name, reason, location)
        QMessageBox.information(self, "Success", "Digital signature added successfully!")
        super().accept()


class InformationRightsDialog(QDialog):
    """Dialog for managing information rights management."""

    def __init__(self, security_manager: SecurityManager, parent=None):
        super().__init__(parent)
        self.security_manager = security_manager
        self.setWindowTitle("Information Rights Management")
        self.setModal(True)
        self.setMinimumSize(700, 500)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # IRM status
        status_group = QGroupBox("IRM Status")
        status_layout = QVBoxLayout()

        self.irm_enabled_checkbox = QCheckBox("Enable Information Rights Management")
        self.irm_enabled_checkbox.setChecked(self.security_manager.irm_enabled)
        self.irm_enabled_checkbox.stateChanged.connect(self.toggle_irm)

        status_layout.addWidget(self.irm_enabled_checkbox)

        if self.security_manager.irm_enabled:
            owner_label = QLabel(f"Owner: {self.security_manager.irm_owner or 'Not set'}")
            status_layout.addWidget(owner_label)

        status_group.setLayout(status_layout)
        layout.addWidget(status_group)

        # Policy info
        policy_group = QGroupBox("Policy Information")
        policy_layout = QFormLayout()

        self.policy_name_edit = QLineEdit()
        self.policy_name_edit.setText(self.security_manager.irm_policy_name)
        self.policy_name_edit.setEnabled(self.security_manager.irm_enabled)

        self.policy_desc_edit = QTextEdit()
        self.policy_desc_edit.setPlainText(self.security_manager.irm_policy_description)
        self.policy_desc_edit.setMaximumHeight(60)
        self.policy_desc_edit.setEnabled(self.security_manager.irm_enabled)

        policy_layout.addRow("Policy Name:", self.policy_name_edit)
        policy_layout.addRow("Description:", self.policy_desc_edit)

        policy_group.setLayout(policy_layout)
        layout.addWidget(policy_group)

        # User permissions
        permissions_group = QGroupBox("User Permissions")
        permissions_layout = QVBoxLayout()

        self.permissions_table = QTableWidget()
        self.permissions_table.setColumnCount(3)
        self.permissions_table.setHorizontalHeaderLabels(["User", "Permissions", "Expiry"])
        self.permissions_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.permissions_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.permissions_table.setEnabled(self.security_manager.irm_enabled)

        self.refresh_permissions_list()

        permissions_layout.addWidget(self.permissions_table)

        # Permission buttons
        perm_button_layout = QHBoxLayout()

        self.add_perm_button = QPushButton("Add User")
        self.add_perm_button.clicked.connect(self.add_user_permission)
        self.add_perm_button.setEnabled(self.security_manager.irm_enabled)

        self.remove_perm_button = QPushButton("Remove User")
        self.remove_perm_button.clicked.connect(self.remove_user_permission)
        self.remove_perm_button.setEnabled(self.security_manager.irm_enabled)

        perm_button_layout.addWidget(self.add_perm_button)
        perm_button_layout.addWidget(self.remove_perm_button)
        perm_button_layout.addStretch()

        permissions_layout.addLayout(perm_button_layout)

        permissions_group.setLayout(permissions_layout)
        layout.addWidget(permissions_group)

        # Dialog buttons
        button_layout = QHBoxLayout()

        self.save_button = QPushButton("Save")
        self.save_button.clicked.connect(self.save_settings)

        self.close_button = QPushButton("Close")
        self.close_button.clicked.connect(self.accept)

        button_layout.addStretch()
        button_layout.addWidget(self.save_button)
        button_layout.addWidget(self.close_button)

        layout.addLayout(button_layout)

        self.setLayout(layout)

    def toggle_irm(self, state):
        enabled = state == Qt.CheckState.Checked.value

        if enabled and not self.security_manager.irm_owner:
            from PySide6.QtWidgets import QInputDialog
            owner, ok = QInputDialog.getText(self, "Set Owner",
                "Enter the document owner's name or email:")
            if ok and owner:
                self.security_manager.enable_irm(owner)
            else:
                self.irm_enabled_checkbox.setChecked(False)
                return
        elif not enabled:
            self.security_manager.disable_irm()

        # Update UI
        self.policy_name_edit.setEnabled(enabled)
        self.policy_desc_edit.setEnabled(enabled)
        self.permissions_table.setEnabled(enabled)
        self.add_perm_button.setEnabled(enabled)
        self.remove_perm_button.setEnabled(enabled)

    def refresh_permissions_list(self):
        self.permissions_table.setRowCount(0)
        permissions = self.security_manager.irm_permissions

        for perm in permissions:
            row = self.permissions_table.rowCount()
            self.permissions_table.insertRow(row)

            self.permissions_table.setItem(row, 0, QTableWidgetItem(perm.user_name))

            perm_str = ", ".join([p.value for p in perm.permissions])
            self.permissions_table.setItem(row, 1, QTableWidgetItem(perm_str))

            expiry_str = perm.expiry_date.strftime("%Y-%m-%d") if perm.expiry_date else "Never"
            self.permissions_table.setItem(row, 2, QTableWidgetItem(expiry_str))

            # Store user ID
            self.permissions_table.item(row, 0).setData(Qt.ItemDataRole.UserRole, perm.user_id)

    def add_user_permission(self):
        dialog = AddUserPermissionDialog(self.security_manager, self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            self.refresh_permissions_list()

    def remove_user_permission(self):
        row = self.permissions_table.currentRow()
        if row < 0:
            QMessageBox.warning(self, "Error", "Please select a user to remove!")
            return

        user_id = self.permissions_table.item(row, 0).data(Qt.ItemDataRole.UserRole)

        reply = QMessageBox.question(self, "Confirm",
            "Are you sure you want to revoke this user's permissions?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

        if reply == QMessageBox.StandardButton.Yes:
            self.security_manager.revoke_permission(user_id)
            self.refresh_permissions_list()

    def save_settings(self):
        self.security_manager.irm_policy_name = self.policy_name_edit.text()
        self.security_manager.irm_policy_description = self.policy_desc_edit.toPlainText()
        QMessageBox.information(self, "Success", "IRM settings saved successfully!")


class AddUserPermissionDialog(QDialog):
    """Dialog for adding user permissions."""

    def __init__(self, security_manager: SecurityManager, parent=None):
        super().__init__(parent)
        self.security_manager = security_manager
        self.setWindowTitle("Add User Permission")
        self.setModal(True)
        self.setMinimumWidth(400)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        form_layout = QFormLayout()

        self.user_name_edit = QLineEdit()
        self.user_name_edit.setPlaceholderText("User name or email")

        self.user_id_edit = QLineEdit()
        self.user_id_edit.setPlaceholderText("User ID (optional)")

        form_layout.addRow("User Name:", self.user_name_edit)
        form_layout.addRow("User ID:", self.user_id_edit)

        layout.addLayout(form_layout)

        # Permissions
        perm_group = QGroupBox("Permissions")
        perm_layout = QVBoxLayout()

        self.view_check = QCheckBox("View")
        self.edit_check = QCheckBox("Edit")
        self.print_check = QCheckBox("Print")
        self.copy_check = QCheckBox("Copy")
        self.full_control_check = QCheckBox("Full Control")

        self.view_check.setChecked(True)

        perm_layout.addWidget(self.view_check)
        perm_layout.addWidget(self.edit_check)
        perm_layout.addWidget(self.print_check)
        perm_layout.addWidget(self.copy_check)
        perm_layout.addWidget(self.full_control_check)

        perm_group.setLayout(perm_layout)
        layout.addWidget(perm_group)

        # Expiry
        expiry_group = QGroupBox("Expiration")
        expiry_layout = QVBoxLayout()

        self.no_expiry_check = QCheckBox("No expiration")
        self.no_expiry_check.setChecked(True)
        self.no_expiry_check.stateChanged.connect(self.toggle_expiry_date)

        self.expiry_date = QDateEdit()
        self.expiry_date.setCalendarPopup(True)
        self.expiry_date.setDate(QDate.currentDate().addDays(30))
        self.expiry_date.setEnabled(False)

        expiry_layout.addWidget(self.no_expiry_check)
        expiry_layout.addWidget(QLabel("Expiry Date:"))
        expiry_layout.addWidget(self.expiry_date)

        expiry_group.setLayout(expiry_layout)
        layout.addWidget(expiry_group)

        # Buttons
        button_layout = QHBoxLayout()
        self.add_button = QPushButton("Add")
        self.cancel_button = QPushButton("Cancel")
        self.add_button.clicked.connect(self.accept)
        self.cancel_button.clicked.connect(self.reject)
        button_layout.addStretch()
        button_layout.addWidget(self.add_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)

        self.setLayout(layout)

    def toggle_expiry_date(self, state):
        self.expiry_date.setEnabled(state != Qt.CheckState.Checked.value)

    def accept(self):
        user_name = self.user_name_edit.text().strip()
        if not user_name:
            QMessageBox.warning(self, "Error", "User name is required!")
            return

        user_id = self.user_id_edit.text().strip() or str(hash(user_name))

        # Collect permissions
        permissions = []
        if self.view_check.isChecked():
            permissions.append(PermissionLevel.VIEW)
        if self.edit_check.isChecked():
            permissions.append(PermissionLevel.EDIT)
        if self.print_check.isChecked():
            permissions.append(PermissionLevel.PRINT)
        if self.copy_check.isChecked():
            permissions.append(PermissionLevel.COPY)
        if self.full_control_check.isChecked():
            permissions.append(PermissionLevel.FULL_CONTROL)

        if not permissions:
            QMessageBox.warning(self, "Error", "At least one permission must be selected!")
            return

        # Get expiry date
        expiry = None
        if not self.no_expiry_check.isChecked():
            q_date = self.expiry_date.date()
            expiry = datetime(q_date.year(), q_date.month(), q_date.day())

        self.security_manager.grant_permission(user_id, user_name, permissions, expiry)
        QMessageBox.information(self, "Success", "User permission added successfully!")
        super().accept()
