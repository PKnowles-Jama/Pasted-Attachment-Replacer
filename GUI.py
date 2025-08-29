import sys
import os
from PyQt6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, QLineEdit, QRadioButton, QLabel, QTextEdit, QHBoxLayout, QFrame, QFormLayout, QFileDialog)
from PyQt6.QtCore import Qt, QObject, pyqtSignal, QThread
from PyQt6.QtGui import QIcon, QPixmap
from PermanentHeader import permanent_header
from JamaLogin import JamaLogin
from Functions import update_jama_attachments

# Custom stream class to redirect stdout (print statements) to the QTextEdit widget
class Stream(QObject):
    """
    A custom stream class that redirects stdout to a signal.
    This allows us to capture print statements and display them in the GUI.
    """
    text_written = pyqtSignal(str)

    def write(self, text):
        self.text_written.emit(str(text))
    
    def flush(self):
        pass

# Worker class to run the long-running functions in a separate thread
class Worker(QObject):
    """
    A worker object to execute the update sequence.
    Inherits from QObject to allow signals.
    """
    finished = pyqtSignal()
    
    def __init__(self, basic_oauth, jama_username, jama_password, project_api_id, url, attachment_item_type_id, file_path):
        super().__init__()
        self.basic_oauth = basic_oauth
        self.jama_username = jama_username
        self.jama_password = jama_password
        self.project_api_id = project_api_id
        self.url = url
        self.attachment_item_type_id = attachment_item_type_id
        self.file_path = file_path


    def run(self):
        """
        Executes the two update functions in sequence.
        This method will be run in the new thread.
        """
        try:
            # Construct the V2 URL
            if not self.url.endswith("/"):
                jama_base_url = self.url
            else:
                jama_base_url = self.url
            
            # Execute JamaLogin
            session = JamaLogin(self.basic_oauth, self.jama_username, self.jama_password, jama_base_url)
            
            # Execute update_jama_attachments
            # Strip trailing slash from url to prevent double slashes
            clean_url = self.url.rstrip('/')
            update_jama_attachments(clean_url, session, int(self.project_api_id), int(self.attachment_item_type_id), self.file_path, self.basic_oauth)

        except Exception as e:
            print(f"An error occurred during the update sequence: {e}")
        
        # Emit the finished signal when done
        self.finished.emit()


class AttachmentReplacer(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Attachment Replacer") # Title of app in header bar
        
        # Open app in the center of the screen & set size
        screen = QApplication.primaryScreen() # Scrape the correct screen to open on
        screen_geometry = screen.geometry() # Determine the primary screen's geometry
        window_width = 800 # Width of the app
        window_height = 500 # Height of the app
        x = (screen_geometry.width() - window_width) // 2 # Calculate the halfway width
        y = (screen_geometry.height() - window_height) // 2 # Calculate the halfway height
        self.setGeometry(x, y, window_width, window_height) # Set the app's opening location and size

        # Set the icon
        script_dir = os.path.dirname(os.path.abspath(__file__)) # Get the file path for this code
        icon_path = os.path.join(script_dir, 'jama_logo_icon.png') # Add the icon's file name to the path
        self.setWindowIcon(QIcon(icon_path)) # Add the icon to the app's header bar

        # Initialize the main layout for the entire window
        self.main_app_layout = QVBoxLayout()
        self.setLayout(self.main_app_layout)

        # --- Permanent Top Section ---
        header_layout, separator = permanent_header('Attachment Replacer','jama_logo.png')
        self.main_app_layout.addLayout(header_layout)
        self.main_app_layout.addWidget(separator)

        # --- Dynamic Content Area ---
        # This is the layout that will be cleared and repopulated
        self.dynamic_content_layout = QVBoxLayout()
        self.main_app_layout.addLayout(self.dynamic_content_layout) # Add it to the main layout

        # Add a stretch to push content to the top
        self.main_app_layout.addStretch()

        self.SelectLoginMethod()

        # Redirect sys.stdout to our custom Stream class
        self.stream = Stream()
        sys.stdout = self.stream
        self.stream.text_written.connect(self.log_to_readout)

    def log_to_readout(self, text):
        """Append text to the readout log."""
        self.readout_log.insertPlainText(text)
        self.readout_log.verticalScrollBar().setValue(self.readout_log.verticalScrollBar().maximum())

    def clearLayout(self, layout):
        if layout is not None:
            while layout.count():
                item = layout.takeAt(0)
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()
                else:
                    self.clearLayout(item.layout())

    def SelectLoginMethod(self):
        # A set of radio buttons to allow the user to select Basic or oAuth login for Jama Connect
        form_layout = QFormLayout()

        self.basic = QRadioButton("Basic")
        self.oAuth = QRadioButton("oAuth")
        self.basic.setChecked(True)
        
        radio_button_layout = QHBoxLayout()
        radio_button_layout.addWidget(self.basic)
        radio_button_layout.addWidget(self.oAuth)
        radio_button_layout.addStretch()

        form_layout.addRow("Select Jama Connect Login Method:", radio_button_layout)
        
        self.submit_button = self.NextButton("Submit",True)
        
        self.dynamic_content_layout.addLayout(form_layout)
        self.dynamic_content_layout.addWidget(self.submit_button)
        self.dynamic_content_layout.addStretch() # Push content to the top within its own section

        self.submit_button.clicked.connect(self.CheckLoginMethod)

    def CheckLoginMethod(self):
        # Store the authentication method before clearing the layout
        self.basic_oauth = 'basic' if self.basic.isChecked() else 'oauth'
        
        self.clearLayout(self.dynamic_content_layout) # Clear only the dynamic content layout
        
        if self.basic.isChecked():
            self.LoginForm("Username","Password")
        elif self.oAuth.isChecked():
            self.LoginForm("Client ID","Client Secret")

    def LoginForm(self, UN, PW):
        form_layout = QFormLayout()

        self.Jama_label = QLabel("Jama Connect API Login Information")
        self.URL_label = QLabel("Jama Connect URL: ")
        self.URL_input = QLineEdit()
        self.URL_input.setPlaceholderText("Enter your Jama Connect instance's URL")
        form_layout.addRow(self.URL_label,self.URL_input)

        self.username_label = QLabel(UN + ": ")
        self.username_input = QLineEdit()
        self.username_input.setPlaceholderText("Enter your " + UN)
        form_layout.addRow(self.username_label, self.username_input)

        self.password_label = QLabel(PW + ": ")
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Enter your " + PW)
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        form_layout.addRow(self.password_label, self.password_input)

        self.project_api_id_label = QLabel("Project API ID: ")
        self.project_api_id_input = QLineEdit()
        self.project_api_id_input.setPlaceholderText("Enter the API ID of the specific project for updates")
        form_layout.addRow(self.project_api_id_label,self.project_api_id_input)

        self.attachement_api_id_label = QLabel("Attachment API ID: ")
        self.attachement_api_id_input = QLineEdit()
        self.attachement_api_id_input.setPlaceholderText("Enter the API ID of the Attachment item type (typically 22)")
        form_layout.addRow(self.attachement_api_id_label,self.attachement_api_id_input)

        self.select_file_button = QPushButton("Select ID List File")
        self.select_file_button.clicked.connect(self.select_file)
        self.select_file_button.setStyleSheet("background-color: #0052CC; color: white;")
        self.file_path_label = QLabel("No file selected")
        form_layout.addRow(self.select_file_button,self.file_path_label)

        self.login_button = self.NextButton("Run",True)
        self.save_logs_button = self.NextButton("Save Logs", True)
        self.save_logs_button.hide()
        
        # Add the readout log area
        self.readout_label = QLabel("Log:")
        self.readout_log = QTextEdit()
        self.readout_log.setReadOnly(True)

        # Add to the dynamic content layout
        self.dynamic_content_layout.addLayout(form_layout)
        self.dynamic_content_layout.addWidget(self.login_button)
        self.dynamic_content_layout.addWidget(self.readout_label)
        self.dynamic_content_layout.addWidget(self.readout_log)
        self.dynamic_content_layout.addWidget(self.save_logs_button)
        self.dynamic_content_layout.addStretch() # Push content to the top within its own section

        # Connect the login button to the new sequence function
        self.login_button.clicked.connect(self.start_update_sequence)
        self.save_logs_button.clicked.connect(self.save_logs)

    def select_file(self):
        #
        # This function does the following:
        #   Open file explorer.
        #   Collect user's file name & path.
        #   Only allow user to select the previously determined file type
        #
        file_dialog = QFileDialog()
        file_path, _ = file_dialog.getOpenFileName(self, "Select File", "", "Excel files (*.xlsx)")
        if file_path:
            self.file_path = file_path
            self.file_path_label.setText(f"Selected: {self.file_path}")
            self.select_file_button.setStyleSheet("background-color: #53575A; color: white;") # Set the button color

    def NextButton(self,label,Enable):
        self.next_button = QPushButton(label)
        self.next_button.setEnabled(Enable)
        if not Enable:
            self.next_button.setStyleSheet("background-color: #53575A; color: white;")
        else:
            self.next_button.setStyleSheet("background-color: #0052CC; color: white;")
        return self.next_button

    def start_update_sequence(self):
        """
        Gathers inputs and starts the update sequence in a new thread.
        """
        # Disable the button, hide the save button and clear the log for a new run
        self.login_button.setEnabled(False)
        self.login_button.setStyleSheet("background-color: #53575A; color: white;")
        self.save_logs_button.hide()
        self.readout_log.clear()
        
        print("Starting attachment update sequence...")

        # Use the stored authentication method
        basic_oauth = self.basic_oauth

        # Get all the input values from the GUI
        jama_username = self.username_input.text()
        jama_password = self.password_input.text()
        project_api_id = self.project_api_id_input.text()
        url = self.URL_input.text()
        attachment_item_type_id = self.attachement_api_id_input.text()

        # Create the thread and worker objects
        self.thread = QThread()
        self.worker = Worker(
            basic_oauth=basic_oauth,
            jama_username=jama_username,
            jama_password=jama_password,
            project_api_id=project_api_id,
            url=url,
            attachment_item_type_id=attachment_item_type_id,
            file_path=self.file_path
        )

        # Move the worker object to the thread
        self.worker.moveToThread(self.thread)

        # Connect signals and slots
        self.thread.started.connect(self.worker.run)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        
        # Connect the finished signal to re-enable the button
        self.worker.finished.connect(self.enable_run_button)

        # Start the thread
        self.thread.start()
    
    def enable_run_button(self):
        """Re-enables the 'Run' button and shows the 'Save Logs' button."""
        self.login_button.setEnabled(True)
        self.login_button.setStyleSheet("background-color: #0052CC; color: white;")
        self.save_logs_button.show()

    def save_logs(self):
        """Opens a file dialog and saves the log content to a text file."""
        # Open a save file dialog
        fileName, _ = QFileDialog.getSaveFileName(self, "Save Logs", "", "Text Files (*.txt);;All Files (*)")

        if fileName:
            # If a file name was selected, write the log content to it
            try:
                with open(fileName, 'w', encoding='utf-8') as f:
                    f.write(self.readout_log.toPlainText())
                print(f"Logs successfully saved to {fileName}")
            except Exception as e:
                print(f"Error saving file: {e}")


if __name__ == "__main__":
    # Run the app when running this file
    app = QApplication(sys.argv)
    window = AttachmentReplacer()
    window.show()
    sys.exit(app.exec())