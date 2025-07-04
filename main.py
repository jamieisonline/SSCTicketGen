import sys
import os
import datetime
import csv
from PyQt6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QPushButton, QComboBox, QStackedWidget, QTextEdit, QLineEdit, QHBoxLayout, QSpinBox, QTextBrowser, QMessageBox
)
from PyQt6.QtCore import Qt, QSize, QUrl
from PyQt6.QtMultimedia import QSoundEffect
from PyQt6.QtGui import QPixmap, QMovie
from jinja2 import Environment, FileSystemLoader

class TroubleshooterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.issue_code_map = self.load_issue_code_map()
        self.setWindowTitle("SM9 Ticket Generator")
        self.setMinimumSize(400, 350)
        self.setStyleSheet("""
            QWidget {
                background-color: #08132a;
                color: #ffffff;
            }
            QLineEdit, QTextEdit, QComboBox {
                background-color: #101c36;
                color: #ffffff;
                border: 1px solid #22345a;
            }
            QPushButton {
                background-color: #22345a;
                color: #ffffff;
                border-radius: 4px;
                padding: 6px 12px;
            }
            QPushButton:hover {
                background-color: #142042;
            }
            QLabel {
                color: #ffffff;
            }
        """)
        # Sound effect setup
        self.button_sound = QSoundEffect()
        try:
            self.button_sound.setSource(QUrl.fromLocalFile("buttonsound.wav"))
        except Exception:
            pass
        self.button_sound.setVolume(0.5)

        self.layout = QVBoxLayout(self)
        # Title section at the top
        self.title_label = QLabel("")
        self.title_label.setStyleSheet("font-size: 18px; font-weight: bold; margin-bottom: 10px;")
        # Prepare to load Windows Passwords steps if needed
        self.windows_passwords_steps = None
        try:
            with open("windows passwords.txt", "r", encoding="utf-8") as f:
                self.windows_passwords_steps = f.read()
        except Exception:
            self.windows_passwords_steps = None
        self.layout.addWidget(self.title_label)
        self.stacked = QStackedWidget()
        self.layout.addWidget(self.stacked)
        self.init_ui()
        # Footer layout for logo and copyright
        footer_layout = QHBoxLayout()
        # SSC Logo on the left
        self.logo_label = QLabel()
        try:
            logo_pixmap = QPixmap("SSC-Logo-Purple-Leaf.png")
            if not logo_pixmap.isNull():
                self.logo_label.setPixmap(logo_pixmap.scaled(60, 60, aspectRatioMode=Qt.AspectRatioMode.KeepAspectRatio, transformMode=Qt.TransformationMode.SmoothTransformation))
        except Exception:
            pass
        self.logo_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        footer_layout.addWidget(self.logo_label, alignment=Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        # Copyright text next to logo
        copyright_label = QLabel()
        copyright_label.setText('<a href="https://github.com/jamieisonline" style="color: #cccccc; text-decoration: none;">@Jamie Goodrich 2025</a>')
        copyright_label.setStyleSheet("font-size: 10px; color: #cccccc; margin-left: 8px;")
        copyright_label.setTextFormat(Qt.TextFormat.RichText)
        copyright_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextBrowserInteraction)
        copyright_label.setOpenExternalLinks(True)
        footer_layout.addWidget(copyright_label, alignment=Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        footer_layout.addStretch(1)
        self.layout.addLayout(footer_layout)

    def play_button_sound(self):
        if self.button_sound.isLoaded():
            self.button_sound.play()
        else:
            self.button_sound.setSource(QUrl.fromLocalFile("buttonsound.wav"))
            self.button_sound.play()

    def connect_with_sound(self, button, slot):
        def wrapper(*args, **kwargs):
            self.play_button_sound()
            return slot(*args, **kwargs)
        button.clicked.connect(wrapper)

    def add_back_button(self, layout, prev_index):
        back_btn = QPushButton("Back")
        self.connect_with_sound(back_btn, lambda _=None: self.stacked.setCurrentIndex(prev_index))
        layout.addWidget(back_btn)

    def update_title(self):
        branch = self.branch_combo.currentText() if hasattr(self, 'branch_combo') else ""
        region = self.region_combo.currentText() if hasattr(self, 'region_combo') else ""
        issue_code = getattr(self, 'selected_issue_code', "")
        issue = getattr(self, 'selected_issue', "")
        # Prefer code from CSV, fallback to issue name
        last = issue_code if issue_code else issue
        title = " - ".join(filter(None, [branch, region, last]))
        self.title_label.setText(title)

    def init_ui(self):
        # Combined Step 1-3: Language, Branch, Region selection
        combined_widget = QWidget()
        combined_layout = QVBoxLayout(combined_widget)
        # Language selection
        lang_label = QLabel("Select Language / Sélectionnez la langue:")
        self.lang_combo = QComboBox()
        self.lang_combo.addItems(["English", "Français"])
        combined_layout.addWidget(lang_label)
        combined_layout.addWidget(self.lang_combo)
        # Branch selection
        branch_label = QLabel("Select Branch:")
        self.branch_combo = QComboBox()
        self.branch_combo.addItems(["PSPC", "SSC"])
        combined_layout.addWidget(branch_label)
        combined_layout.addWidget(self.branch_combo)
        # Region code selection (dropdown)
        region_label = QLabel("Select Region:")
        self.region_combo = QComboBox()
        self.region_combo.addItems(["NCR", "QUE", "ATL", "WST", "ONT", "PAC"])  # Added PAC
        combined_layout.addWidget(region_label)
        combined_layout.addWidget(self.region_combo)
        # Asset/Serial/IMEI input
        asset_label = QLabel("Asset / Serial / IMEI Number:")
        self.asset_input = QLineEdit()
        self.asset_input.setPlaceholderText("Enter Asset, Serial, or IMEI Number")
        combined_layout.addWidget(asset_label)
        combined_layout.addWidget(self.asset_input)
        # Existing Ticket Number input
        ticketnum_label = QLabel("Existing Ticket Number:")
        self.ticketnum_input = QLineEdit()
        self.ticketnum_input.setPlaceholderText("Enter Existing Ticket Number (if any)")
        combined_layout.addWidget(ticketnum_label)
        combined_layout.addWidget(self.ticketnum_input)
        # Callback Number input
        callback_label = QLabel("Callback Number:")
        self.callback_input = QLineEdit()
        self.callback_input.setPlaceholderText("Enter Callback Number")
        combined_layout.addWidget(callback_label)
        combined_layout.addWidget(self.callback_input)
        # User count selector and VPN/Office toggle in a single row
        user_vpn_row = QHBoxLayout()
        user_count_label = QLabel("Users Affected:")
        user_count_label.setStyleSheet("font-size: 12px;")
        self.user_count_spin = QSpinBox()
        self.user_count_spin.setMinimum(1)
        self.user_count_spin.setMaximum(1000)
        self.user_count_spin.setValue(1)
        self.user_count_spin.setFixedWidth(60)
        user_vpn_row.addWidget(user_count_label)
        user_vpn_row.addWidget(self.user_count_spin)

        # Add spacing between the widgets
        user_vpn_row.addSpacing(20)

        vpn_office_label = QLabel("Connection Type:")
        vpn_office_label.setStyleSheet("font-size: 12px;")
        self.vpn_office_combo = QComboBox()
        self.vpn_office_combo.addItems(["VPN", "Office"])
        user_vpn_row.addWidget(vpn_office_label)
        user_vpn_row.addWidget(self.vpn_office_combo)
        user_vpn_row.addStretch(1)
        combined_layout.addLayout(user_vpn_row)
        # Next button to go to device type
        combined_next = QPushButton("Next")
        self.connect_with_sound(combined_next, self.goto_device_type)
        combined_layout.addWidget(combined_next)
        self.stacked.addWidget(combined_widget)

        # Step 4: Device type
        device_widget = QWidget()
        device_layout = QVBoxLayout(device_widget)
        # EU description input
        eu_label = QLabel("EU description of the problem:")
        self.eu_input = QLineEdit()
        self.eu_input.setPlaceholderText("Enter end user description here")
        self.eu_input.setMinimumHeight(50)  # Make the box taller
        device_layout.addWidget(eu_label)
        device_layout.addWidget(self.eu_input)
        device_label = QLabel("Device Type:")
        laptop_btn = QPushButton("Laptop")
        mobile_btn = QPushButton("Mobile")
        self.connect_with_sound(laptop_btn, lambda _=None: self.select_device("Laptop"))
        self.connect_with_sound(mobile_btn, lambda _=None: self.select_device("Mobile"))
        device_layout.addWidget(device_label)
        device_layout.addWidget(laptop_btn)
        device_layout.addWidget(mobile_btn)
        self.add_back_button(device_layout, 0)
        self.stacked.addWidget(device_widget)

        # Step 5: Issue type
        issue_widget = QWidget()
        issue_layout = QVBoxLayout(issue_widget)
        issue_label = QLabel("Issue Type:")
        hardware_btn = QPushButton("Hardware")
        software_btn = QPushButton("Software")
        self.connect_with_sound(hardware_btn, lambda _=None: self.goto_issue_list("Hardware"))
        self.connect_with_sound(software_btn, lambda _=None: self.goto_issue_list("Software"))
        issue_layout.addWidget(issue_label)
        issue_layout.addWidget(hardware_btn)
        issue_layout.addWidget(software_btn)
        self.add_back_button(issue_layout, 1)
        self.stacked.addWidget(issue_widget)

        # Step 6: Issue List (placeholder)
        self.issue_list_widget = QWidget()
        issue_list_layout = QVBoxLayout(self.issue_list_widget)
        # Clarification label for device and issue type
        self.issue_clarification_label = QLabel("")
        self.issue_clarification_label.setStyleSheet("font-size: 14px; font-weight: bold; margin-bottom: 6px;")
        issue_list_layout.addWidget(self.issue_clarification_label)
        self.issue_list_label = QLabel("Select Issue (placeholder):")
        # Placeholder: VPN, Email, Other for software; Battery, Screen, Other for hardware
        self.issue_btns = []
        issue_list_layout.addWidget(self.issue_list_label)
        self.issue_btn_layout = QVBoxLayout()
        issue_list_layout.addLayout(self.issue_btn_layout)
        self.add_back_button(issue_list_layout, 2)
        self.stacked.addWidget(self.issue_list_widget)

        # Step 7: Steps/Article (placeholder)
        self.steps_widget = QWidget()
        steps_layout = QVBoxLayout(self.steps_widget)
        self.steps_label = QTextBrowser()
        self.steps_label.setOpenExternalLinks(True)
        self.steps_label.setStyleSheet(
            "background-color: #101c36; color: #ffffff; border: 1px solid #22345a; font-size: 16px;"
        )
        self.steps_label.setMinimumHeight(345) # Make the box taller
        steps_layout.addWidget(self.steps_label)
        # Move next button to top
        next_btn = QPushButton("Next")
        self.connect_with_sound(next_btn, self.goto_ticket_page)
        steps_layout.addWidget(next_btn)
        # Move back button to bottom
        steps_layout.addStretch(1)
        back_btn_steps = QPushButton("Back")
        self.connect_with_sound(back_btn_steps, lambda _=None: self.stacked.setCurrentIndex(3))
        steps_layout.addWidget(back_btn_steps)
        self.stacked.addWidget(self.steps_widget)

        # Step 8: Editable Ticket Output
        ticket_widget = QWidget()
        ticket_layout = QVBoxLayout(ticket_widget)
        # Editable/copyable ticket title
        title_row = QHBoxLayout()
        self.ticket_title_edit = QLineEdit()
        self.ticket_title_edit.setPlaceholderText("Ticket Title")
        self.ticket_title_edit.setText("Generated Ticket")
        copy_title_btn = QPushButton("Copy Title", parent=ticket_widget)
        self.connect_with_sound(copy_title_btn, self.copy_ticket_title)
        title_row.addWidget(self.ticket_title_edit)
        title_row.addWidget(copy_title_btn)
        ticket_layout.addLayout(title_row)
        ticket_label = QLabel("Generated Ticket:")
        self.ticket_text = QTextEdit()
        self.ticket_text.setReadOnly(False)  # Editable
        copy_btn = QPushButton("Copy Ticket", parent=ticket_widget)
        self.connect_with_sound(copy_btn, self.copy_ticket)
        save_btn = QPushButton("Save to TXT", parent=ticket_widget)
        self.connect_with_sound(save_btn, self.save_ticket_to_txt)
        ticket_layout.addWidget(ticket_label)
        ticket_layout.addWidget(self.ticket_text)
        ticket_layout.addWidget(copy_btn)
        ticket_layout.addWidget(save_btn)
        # Add "New Ticket" button to the ticket page
        new_ticket_btn = QPushButton("New Ticket", parent=ticket_widget)
        def confirm_new_ticket(_=None):
            reply = QMessageBox.question(
                self,
                "Start New Ticket",
                "Are you sure you want to start a new ticket? All current information will be lost.",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                self.clear_all_fields()
                self.stacked.setCurrentIndex(0)
                self.update_title()
        self.connect_with_sound(new_ticket_btn, confirm_new_ticket)
        ticket_layout.addWidget(new_ticket_btn)
        # Move back button below the new ticket button
        back_btn_ticket = QPushButton("Back", parent=ticket_widget)
        def back_to_steps(_=None):
            self.stacked.setCurrentIndex(4)
        self.connect_with_sound(back_btn_ticket, back_to_steps)
        ticket_layout.addWidget(back_btn_ticket)
        self.stacked.addWidget(ticket_widget)

    def copy_ticket_title(self, event=None):
        clipboard = QApplication.instance().clipboard()
        clipboard.setText(self.ticket_title_edit.text())

    def goto_device_type(self, event=None):
        # Allow proceeding even if region is left blank
        self.selected_region = self.region_combo.currentText() if self.region_combo.currentIndex() != -1 else ""
        self.stacked.setCurrentIndex(1)
        self.update_title()

    def select_device(self, device):
        self.selected_device = device
        self.stacked.setCurrentIndex(2)
        self.update_title()

    def goto_issue_list(self, issue_type):
        self.selected_issue_type = issue_type
        # Clear previous buttons
        for i in reversed(range(self.issue_btn_layout.count())):
            self.issue_btn_layout.itemAt(i).widget().setParent(None)

        lang = self.lang_combo.currentText() if hasattr(self, 'lang_combo') else "English"
        branch = self.branch_combo.currentText().lower() if hasattr(self, 'branch_combo') else "pspc"
        device = getattr(self, 'selected_device', '').lower()
        issue_type_folder = issue_type.lower()

        base_path = os.path.dirname(__file__)
        issues_folder = os.path.join(base_path, "articles", branch, device, issue_type_folder)
        if lang == "Français":
            issues_folder = os.path.join(base_path, "articles", "french", branch, device, issue_type_folder)
        else:
            issues_folder = os.path.join(base_path, "articles", branch, device, issue_type_folder)

        issues = []
        if os.path.exists(issues_folder):
            for fname in os.listdir(issues_folder):
                if fname.endswith(".j2"):
                    issues.append(os.path.splitext(fname)[0])

        for issue in issues:
            pretty_label = issue.replace("-", " ").replace("_", " ").title()
            btn = QPushButton(pretty_label)
            self.connect_with_sound(btn, lambda _, iss=issue: self.select_issue(iss))
            self.issue_btn_layout.addWidget(btn)
            self.issue_btns.append(btn)

        # --- Add "Issue Not Listed" button ---
        issue_not_listed_btn = QPushButton("Issue Not Listed")
        issue_not_listed_btn.setStyleSheet("background-color: #e75480; color: #fff; font-weight: bold;")  # Pink button
        self.connect_with_sound(issue_not_listed_btn, self.handle_issue_not_listed)
        self.issue_btn_layout.addWidget(issue_not_listed_btn)
        # --------------------------------------

        self.issue_list_label.setText(f"Select {issue_type} Issue:")
        self.issue_clarification_label.setText(f"{self.selected_device} - {issue_type} Issues" if self.selected_device and issue_type else "")
        self.stacked.setCurrentIndex(3)
        self.update_title()

    def select_issue(self, issue):
        self.selected_issue = issue
        article_filename = f"{issue}.j2".lower()
        self.selected_issue_code = self.issue_code_map.get(article_filename, "")
        self.update_title()
        lang = self.lang_combo.currentText() if hasattr(self, 'lang_combo') else "English"
        branch = self.branch_combo.currentText().lower() if hasattr(self, 'branch_combo') else "pspc"
        device = getattr(self, 'selected_device', '').lower()
        issue_type = self.selected_issue_type.lower()
        base_path = os.path.dirname(__file__)

        # Choose article file based on language
        if lang == "Français":
            issue_file = os.path.join(base_path, "articles", "french", branch, device, issue_type, f"{issue}.j2")
        else:
            issue_file = os.path.join(base_path, "articles", branch, device, issue_type, f"{issue}.j2")

        article_text = ""
        steps_text = ""
        resolution_text = ""
        confluence_link = ""
        if os.path.exists(issue_file):
            with open(issue_file, "r", encoding="utf-8") as f:
                article_text = f.read()
            import re
            comments = re.findall(r"\{#(.*?)#\}", article_text, re.DOTALL)
            steps_text = re.sub(r"\{#.*?#\}", "", article_text, flags=re.DOTALL).strip()
            resolution_text = comments[0].strip() if comments else ""
            # Get the last comment as the confluence link if it looks like a URL
            if comments and comments[-1].strip().startswith("http"):
                confluence_link = comments[-1].strip()
        else:
            steps_text = "No article found for this issue."
            resolution_text = ""
            confluence_link = ""

        self.selected_steps = steps_text
        self.selected_resolution = resolution_text
        self.selected_confluence_link = confluence_link

        # Steps page: show each step on a new line, preserving numbering
        formatted = steps_text
        if confluence_link:
            formatted += f"\n\nMore information: <a href='{confluence_link}' style='color:#6cf;text-decoration:underline;'>click here</a>"
        self.steps_label.setHtml(formatted.replace('\n', '<br>'))
        self.stacked.setCurrentIndex(4)

    def load_article(self, branch, issue_key, context=None):
        # Build file path for Jinja2 template
        article_path = os.path.join(
            os.path.dirname(__file__),
            "articles",
            branch.lower(),
            f"{issue_key}.j2"
        )
        if os.path.exists(article_path):
            env = Environment(loader=FileSystemLoader(os.path.dirname(article_path)))
            template = env.get_template(f"{issue_key}.j2")
            # Pass context for dynamic fields, or empty dict if none
            return template.render(context or {})
        else:
            return "No article found for this issue."

    def format_article_as_bullets(self, text):
        # If it's an error message or already HTML, don't format
        if text.startswith("No article found") or "<ul>" in text or "<ol>" in text:
            return text
        # Split lines, ignore empty lines, and wrap in <ul>
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        if not lines:
            return ""
        html = "<ul style='margin-left: 0; font-size: 16px; line-height: 1.8em; text-align: center;'>"
        for line in lines:
            html += f"<li>{line}</li>"
        html += "</ul>"
        return html

    def goto_ticket_page(self, event=None):
        branch = self.branch_combo.currentText() if hasattr(self, 'branch_combo') else ""
        region = self.region_combo.currentText() if hasattr(self, 'region_combo') else ""
        issue_code = getattr(self, 'selected_issue_code', "")
        issue = getattr(self, 'selected_issue', "")
        last = issue_code if issue_code else issue
        ticket_title = " - ".join(filter(None, [branch, region, last]))
        if hasattr(self, 'ticket_title_edit'):
            self.ticket_title_edit.setText(ticket_title if ticket_title else "Generated Ticket")
        # Gather user input
        lang = self.lang_combo.currentText() if hasattr(self, 'lang_combo') else "English"
        branch = self.branch_combo.currentText() if hasattr(self, 'branch_combo') else ""
        device = getattr(self, 'selected_device', '').lower()
        region = self.region_combo.currentText() if hasattr(self, 'region_combo') else ""
        asset = self.asset_input.text() if hasattr(self, 'asset_input') else ""
        serial = asset
        callback = self.callback_input.text() if hasattr(self, 'callback_input') else ""
        ticketnum = self.ticketnum_input.text() if hasattr(self, 'ticketnum_input') else ""
        eu_desc = self.eu_input.text() if hasattr(self, 'eu_input') else ""
        issue = getattr(self, 'selected_issue', '')
        users_affected = self.user_count_spin.value() if hasattr(self, 'user_count_spin') else 1
        today = datetime.datetime.now().strftime("%-m/%-d/%Y") if sys.platform != "win32" else datetime.datetime.now().strftime("%#m/%#d/%Y")
        vpn = self.vpn_office_combo.currentText() if hasattr(self, 'vpn_office_combo') else "VPN"

        # Template selection
        if lang == "English" and device == "laptop":
            template_file = "eng-basic.j2"
        elif lang == "English" and device == "mobile":
            template_file = "eng-mobile.j2"
        elif lang == "Français" and device == "laptop":
            template_file = "fra-basic.j2"
        elif lang == "Français" and device == "mobile":
            template_file = "fra-mobile.j2"
        else:
            template_file = "eng-basic.j2"

        # Template directory (portable for PyInstaller)
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(__file__)
        template_dir = os.path.join(base_path, "articles", "templates")
        env = Environment(loader=FileSystemLoader(template_dir))
        template = env.get_template(template_file)

        # Dynamic fields for template
        if lang == "English":
            asset_field = f"Asset (PSPC/ INFC) : {asset}" if branch == "PSPC" else ""
            serial_field = f"Serial Number (SSC) : {serial}" if branch == "SSC" else ""
            is_new = "x" if not ticketnum else ""
            is_existing = "x" if ticketnum else ""
            existing_ticket = f"Existing Ticket# : {ticketnum}" if ticketnum else ""
            when_field = today
            users_field = users_affected
            vpn_or_core = "VPN" if vpn == "VPN" else "Core Network"
        else:
            asset_field = f"Bien (SPAC/ INFC) ou numéro de série (SSC) : {asset if branch == 'PSPC' else serial}"
            serial_field = ""  # French template uses the same field for both
            is_new = "x" if not ticketnum else ""
            is_existing = "x" if ticketnum else ""
            existing_ticket = f"Numéro de référence : {ticketnum}" if ticketnum else ""
            when_field = today
            users_field = users_affected
            vpn_or_core = "RPV" if vpn == "VPN" else "réseau central"

        # Render template
        ticket_content = template.render(
            lang=lang,
            region=region,
            asset=asset,
            serial=serial,
            callback=callback,
            ticketnum=ticketnum,
            eu_desc=eu_desc,
            issue=issue,
            asset_field=asset_field,
            serial_field=serial_field,
            is_new=is_new,
            is_existing=is_existing,
            existing_ticket=existing_ticket,
            when_field=when_field,
            users_field=users_field,
            vpn=vpn,
            vpn_or_core=vpn_or_core,
            first_encountered=when_field,
            error_msg="",
            steps=getattr(self, 'selected_steps', ''),
            resolution=getattr(self, 'selected_resolution', ''),
            confluence_article=getattr(self, 'selected_confluence_link', ''),
            ip="",
            app_login="",
            mailing_address=""
        )

        self.ticket_text.setPlainText(ticket_content)
        if hasattr(self, 'ticket_title_edit'):
            branch = self.branch_combo.currentText() if hasattr(self, 'branch_combo') else ""
            region = self.region_combo.currentText() if hasattr(self, 'region_combo') else ""
            issue_code = getattr(self, 'selected_issue_code', "")
            issue = getattr(self, 'selected_issue', "")
            last = issue_code if issue_code else issue
            ticket_title = " - ".join(filter(None, [branch, region, last]))
            self.ticket_title_edit.setText(ticket_title if ticket_title else "Generated Ticket")
        self.stacked.setCurrentIndex(5)

    def copy_ticket(self, event=None):
        clipboard = QApplication.instance().clipboard()
        clipboard.setText(self.ticket_text.toPlainText())

    def save_ticket_to_txt(self, event=None):
        from PyQt6.QtWidgets import QFileDialog, QMessageBox
        file_path, _ = QFileDialog.getSaveFileName(self, "Save Ticket", "ticket.txt", "Text Files (*.txt)")
        if file_path:
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(self.ticket_text.toPlainText())
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save file:\n{e}")
                
    def generate_ticket(self, lang, branch, region, asset, serial, callback, vpn, ticketnum, first_encountered, users_affected, eu_desc, error_msg, steps, root_cause, resolution, confluence_article, ip, app_login, mailing_address):
        # GENERAL:
        # IM: incident management: help with issue
        # SD: Service Request: request change
        # Group: {{ region }}
        
        # https://confluence.ssc-spc.gc.ca/pages/viewpage.action?pageId=231737721
        
        # -------------------------------------------------
        # Language of call : English[{{ 'x' if lang == 'English' else '' }}]; French[{{ 'x' if lang == 'Français' else '' }}]
        # Asset (PSPC/ INFC) : {{ asset }}
        # Serial Number (SSC) : {{ serial }}
        # Callback number : {{ callback }}
        # VPN or Core network : {{ vpn }}
        # -------------------------------------------------
        # IS THIS A NEW OR EXISTING ISSUE: NEW[] ; EXISTING[]
        # Existing Ticket# : {{ ticketnum }}
        # WHEN WAS THE ISSUE FIRST ENCOUNTERED : {{ first_encountered }}
        # NUMBER OF USERS AFFECTED : {{ users_affected }}
        # -------------------------------------------------
        # EU description of problem : {{ eu_desc }}
        
        # Error message/ code : {{ error_msg }}
        
        # Document steps taken to troubleshoot EUs issue : {{ steps }}
        
        # What is the root cause found after troubleshooting : {{ root_cause }}
        
        # Resolution or next steps : {{ resolution }}
        
        # Confluence Article used to support EU : {{ confluence_article }}
        
        # IP address (if network related) : {{ ip }}
        # Application login (if applicable) : {{ app_login }}
        # End User's Mailing Address (If KB states required): {{ mailing_address }}
        
        # -------------------------------------------------
        # English - Basic EUSD template June 1st 2023
        pass
                

    def load_issue_code_map(self):
        code_map = {}
        mapping_file = os.path.join(os.path.dirname(__file__), "articles", "Issue-codes.csv")
        try:
            with open(mapping_file, newline='', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                for row in reader:
                    if len(row) >= 2:
                        code_map[row[0].strip().lower()] = row[1].strip()
        except Exception as e:
            print(f"Failed to load issue code mapping: {e}")
        return code_map

    def handle_issue_not_listed(self, event=None):
        self.selected_issue = "Issue Not Listed"
        self.selected_issue_code = ""
        self.selected_steps = "No troubleshooting steps available. Please describe the issue in detail."
        self.selected_resolution = ""
        self.selected_confluence_link = ""
        self.update_title()
        self.goto_ticket_page()

    def closeEvent(self, event):
        reply = QMessageBox.question(
            self,
            "Confirm Exit",
            "Are you sure you want to close the application?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        if reply == QMessageBox.StandardButton.Yes:
            event.accept()
        else:
            event.ignore()

    def clear_all_fields(self):
        # Reset all user input fields and selections
        if hasattr(self, 'lang_combo'):
            self.lang_combo.setCurrentIndex(0)
        if hasattr(self, 'branch_combo'):
            self.branch_combo.setCurrentIndex(0)
        if hasattr(self, 'region_combo'):
            self.region_combo.setCurrentIndex(0)
        if hasattr(self, 'asset_input'):
            self.asset_input.clear()
        if hasattr(self, 'ticketnum_input'):
            self.ticketnum_input.clear()
        if hasattr(self, 'callback_input'):
            self.callback_input.clear()
        if hasattr(self, 'eu_input'):
            self.eu_input.clear()
        if hasattr(self, 'user_count_spin'):
            self.user_count_spin.setValue(1)
        if hasattr(self, 'vpn_office_combo'):
            self.vpn_office_combo.setCurrentIndex(0)
        if hasattr(self, 'ticket_title_edit'):
            self.ticket_title_edit.setText("Generated Ticket")
        if hasattr(self, 'ticket_text'):
            self.ticket_text.clear()
        # Reset selections
        self.selected_device = ""
        self.selected_issue_type = ""
        self.selected_issue = ""
        self.selected_issue_code = ""
        self.selected_steps = ""
        self.selected_resolution = ""

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = TroubleshooterApp()
    window.show()
    sys.exit(app.exec())
