import sys
import os
import glob
import tempfile
import html
from PyQt5 import QtCore, QtWidgets, QtGui
import email
from email import policy
from email.parser import BytesParser

# For .msg support. Make sure to install extract_msg via pip.
try:
    import extract_msg
except ImportError:
    extract_msg = None
    print("Warning: extract_msg module not found. .msg file support will be disabled.")

class EmailViewer(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        # Set the application title to Emlee.
        self.setWindowTitle("Emlee")
        self.resize(800, 600)
        
        # Set application icon (assuming icon.ico is in the same folder as main.py)
        icon_path = os.path.join(os.path.dirname(__file__), "icon.ico")
        self.setWindowIcon(QtGui.QIcon(icon_path))

        # Track current file and list of email files (.eml and .msg).
        self.current_email_path = None
        self.email_files_list = []
        self.current_index = -1

        # Mapping of attachment filename to temporary file path.
        self.attachments = {}

        self.init_ui()
        self.setAcceptDrops(True)

    def init_ui(self):
        # File menu with "Open File" action.
        open_action = QtWidgets.QAction("Open File", self)
        open_action.triggered.connect(self.open_file_dialog)
        file_menu = self.menuBar().addMenu("File")
        file_menu.addAction(open_action)

        # Create central widget and a vertical layout.
        central_widget = QtWidgets.QWidget(self)
        self.setCentralWidget(central_widget)
        main_layout = QtWidgets.QVBoxLayout(central_widget)
        main_layout.setContentsMargins(5, 5, 5, 5)
        main_layout.setSpacing(5)

        # --- Header Section ---
        # Increased fixed height to accommodate Subject.
        self.header_widget = QtWidgets.QWidget()
        self.header_widget.setFixedHeight(130)
        header_layout = QtWidgets.QGridLayout(self.header_widget)
        header_layout.setSpacing(5)
        self.label_from = QtWidgets.QLabel("From: ")
        self.label_to = QtWidgets.QLabel("To: ")
        self.label_cc = QtWidgets.QLabel("CC: ")
        self.label_bcc = QtWidgets.QLabel("BCC: ")
        self.label_subject = QtWidgets.QLabel("Subject: ")
        self.label_date = QtWidgets.QLabel("Date: ")

        # Layout rows: row 0 for From/CC, row 1 for To/BCC, row 2 for Subject, row 3 for Date.
        header_layout.addWidget(self.label_from, 0, 0)
        header_layout.addWidget(self.label_cc, 0, 1)
        header_layout.addWidget(self.label_to, 1, 0)
        header_layout.addWidget(self.label_bcc, 1, 1)
        header_layout.addWidget(self.label_subject, 2, 0, 1, 2)
        header_layout.addWidget(self.label_date, 3, 0, 1, 2)
        main_layout.addWidget(self.header_widget)

        # --- Email Body Section ---
        # QTextBrowser renders HTML and includes a scrollbar.
        self.body_text = QtWidgets.QTextBrowser()
        self.body_text.setOpenExternalLinks(True)
        main_layout.addWidget(self.body_text)

        # --- Attachments Section ---
        self.attachments_list = QtWidgets.QListWidget()
        self.attachments_list.setMaximumHeight(100)
        self.attachments_list.itemDoubleClicked.connect(self.open_attachment)
        main_layout.addWidget(self.attachments_list)

        # --- Navigation Buttons at Bottom Right ---
        nav_layout = QtWidgets.QHBoxLayout()
        nav_layout.addStretch()  # Push buttons to the right.
        prev_button = QtWidgets.QPushButton("Previous")
        next_button = QtWidgets.QPushButton("Next")
        prev_button.clicked.connect(self.load_previous)
        next_button.clicked.connect(self.load_next)
        nav_layout.addWidget(prev_button)
        nav_layout.addWidget(next_button)
        main_layout.addLayout(nav_layout)

    def open_file_dialog(self):
        options = QtWidgets.QFileDialog.Options()
        file_path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Open Email File", "",
            "Email Files (*.eml *.msg)", options=options)
        if file_path:
            self.load_email_file(file_path)

    def load_email_file(self, file_path):
        # Display a loading message before processing.
        self.body_text.setHtml("<p>Loading email, please wait...</p>")
        QtWidgets.QApplication.processEvents()

        self.current_email_path = file_path
        self.setWindowTitle(f"Emlee - {os.path.basename(file_path)}")

        # Update file list from the directory, including both .eml and .msg if supported.
        directory = os.path.dirname(file_path)
        if extract_msg:
            patterns = [os.path.join(directory, "*.eml"), os.path.join(directory, "*.msg")]
        else:
            patterns = [os.path.join(directory, "*.eml")]
        files = []
        for pattern in patterns:
            files.extend(glob.glob(pattern))
        self.email_files_list = sorted(files, key=lambda s: s.lower())
        try:
            self.current_index = self.email_files_list.index(file_path)
        except ValueError:
            self.current_index = -1

        # Clear previous attachments.
        self.attachments_list.clear()
        self.attachments = {}

        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".eml":
            self.load_eml(file_path)
        elif ext == ".msg":
            if extract_msg:
                self.load_msg(file_path)
            else:
                QtWidgets.QMessageBox.warning(self, "Error",
                                              "MSG file support not available. Please install extract_msg module.")
        else:
            QtWidgets.QMessageBox.warning(self, "Error", "Unsupported file format.")

        # Adjust document width to fill available space.
        self.body_text.document().setTextWidth(self.body_text.viewport().width())

    def load_eml(self, file_path):
        with open(file_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)

        # Update header labels.
        self.label_from.setText(f"From: {msg.get('From', '')}")
        self.label_to.setText(f"To: {msg.get('To', '')}")
        self.label_cc.setText(f"CC: {msg.get('Cc', '')}")
        self.label_bcc.setText(f"BCC: {msg.get('Bcc', '')}")
        self.label_subject.setText(f"Subject: {msg.get('Subject', '')}")
        self.label_date.setText(f"Date: {msg.get('Date', '')}")

        # Extract the email body while preserving formatting.
        body = ""
        html_body = None
        plain_body = None
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_disposition() is None:
                    if part.get_content_type() == "text/html":
                        html_body = part.get_content()
                    elif part.get_content_type() == "text/plain":
                        plain_body = part.get_content()
            if html_body:
                body = html_body
            elif plain_body:
                body = "<pre>" + html.escape(plain_body) + "</pre>"
        else:
            content_type = msg.get_content_type()
            content = msg.get_content()
            if content_type == "text/html":
                body = content
            else:
                body = "<pre>" + html.escape(content) + "</pre>"

        self.body_text.setHtml(body)
        self.body_text.document().setTextWidth(self.body_text.viewport().width())

        # Process attachments.
        for part in msg.walk():
            content_disp = part.get("Content-Disposition", "")
            if "attachment" in content_disp:
                filename = part.get_filename()
                if filename:
                    data = part.get_payload(decode=True)
                    temp_dir = tempfile.gettempdir()
                    temp_path = os.path.join(temp_dir, filename)
                    with open(temp_path, 'wb') as temp_file:
                        temp_file.write(data)
                    self.attachments[filename] = temp_path
                    item = QtWidgets.QListWidgetItem(filename)
                    icon = QtWidgets.QFileIconProvider().icon(QtCore.QFileInfo(temp_path))
                    item.setIcon(icon)
                    self.attachments_list.addItem(item)

    def load_msg(self, file_path):
        msg = extract_msg.Message(file_path)
        msg_sender = msg.sender or ""
        msg_to = msg.to or ""
        msg_cc = msg.cc or ""
        msg_bcc = msg.bcc or ""
        msg_date = msg.date or ""
        msg_subject = msg.subject or ""

        self.label_from.setText(f"From: {msg_sender}")
        self.label_to.setText(f"To: {msg_to}")
        self.label_cc.setText(f"CC: {msg_cc}")
        self.label_bcc.setText(f"BCC: {msg_bcc}")
        self.label_subject.setText(f"Subject: {msg_subject}")
        self.label_date.setText(f"Date: {msg_date}")

        # For the body, prefer HTML if available.
        if msg.htmlBody:
            body = msg.htmlBody.decode("utf-8", errors="replace") if isinstance(msg.htmlBody, bytes) else msg.htmlBody
        elif msg.body:
            text = msg.body.decode("utf-8", errors="replace") if isinstance(msg.body, bytes) else msg.body
            body = "<pre>" + html.escape(text) + "</pre>"
        else:
            body = ""
        self.body_text.setHtml(body)
        self.body_text.document().setTextWidth(self.body_text.viewport().width())

        # Process attachments from the MSG file.
        for att in msg.attachments:
            filename = att.longFilename if att.longFilename else att.shortFilename
            if filename:
                temp_dir = tempfile.gettempdir()
                temp_path = os.path.join(temp_dir, filename)
                try:
                    with open(temp_path, "wb") as f:
                        f.write(att.data)
                    self.attachments[filename] = temp_path
                    item = QtWidgets.QListWidgetItem(filename)
                    icon = QtWidgets.QFileIconProvider().icon(QtCore.QFileInfo(temp_path))
                    item.setIcon(icon)
                    self.attachments_list.addItem(item)
                except Exception as e:
                    print(f"Error saving attachment {filename}: {e}")

    def open_attachment(self, item):
        filename = item.text()
        if filename in self.attachments:
            file_path = self.attachments[filename]
            try:
                os.startfile(file_path)
            except Exception as e:
                QtWidgets.QMessageBox.warning(self, "Error",
                                              f"Could not open attachment:\n{str(e)}")

    def load_next(self):
        if self.email_files_list:
            next_index = (self.current_index + 1) % len(self.email_files_list)
            self.load_email_file(self.email_files_list[next_index])

    def load_previous(self):
        if self.email_files_list:
            prev_index = (self.current_index - 1) % len(self.email_files_list)
            self.load_email_file(self.email_files_list[prev_index])

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith(('.eml', '.msg')):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls:
            file_path = urls[0].toLocalFile()
            if file_path.lower().endswith(('.eml', '.msg')):
                self.load_email_file(file_path)

    def resizeEvent(self, event):
        # Update the document's text width whenever the window is resized.
        super().resizeEvent(event)
        self.body_text.document().setTextWidth(self.body_text.viewport().width())

def main():
    app = QtWidgets.QApplication(sys.argv)
    viewer = EmailViewer()
    viewer.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
