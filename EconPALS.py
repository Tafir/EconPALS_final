import sys
from TopHat import *
from PyQt5 import QtGui
from PyQt5.QtWidgets import *
import matplotlib.pyplot as plt
import configparser


class App(QWidget):

    def __init__(self):
        super().__init__()
        self.title = 'EconPALS'
        self.left = 10
        self.top = 10
        self.width = 1200
        self.height = 600
        self.setWindowIcon(QtGui.QIcon('icon.png'))
        self.df = pd.DataFrame()
        self.settings_flag = False
        self.StudyPALS_imported = False
        self.initialise_settings()
        self.initialise_ui()

    #
    # ------------Regular view buttons------------
    #

    @staticmethod
    def session_error_dialog():
        msg = QMessageBox()

        msg.setText("Invalid week/semester combination")
        msg.setWindowTitle("Session error")
        msg.setStandardButtons(QMessageBox.Ok)

        msg.exec_()

    def emails_button(self):
        self.textEdit.clear()
        self.textEdit.setReadOnly(True)
        semester = int(self.cb_sem.currentText())
        try:
            if self.cb_week.currentText() == 'All':
                emails = get_emails(self.df, semester=semester)
            else:
                week = int(self.cb_week.currentText())
                emails = get_emails(self.df, week, semester)
            self.textEdit.insertPlainText(emails)
        except:
            self.session_error_dialog()

    def regulars_button(self):
        self.textEdit.clear()
        self.textEdit.setReadOnly(True)
        regulars = regulars_list(self.df, int(self.cb_reg.currentText()))
        self.textEdit.insertPlainText(regulars)

    def select_file(self):
        dlg = QFileDialog()
        dlg.setFileMode(QFileDialog.AnyFile)
        dlg.setNameFilters(["Excel files (*.xls)"])
        if dlg.exec_():
            filenames = dlg.selectedFiles()
            if filenames:
                self.file_path_box.setText(filenames[0])

    def studypals_select_file(self):
        dlg = QFileDialog()
        dlg.setFileMode(QFileDialog.AnyFile)
        dlg.setNameFilters(["Excel files (*.xls)"])
        if dlg.exec_():
            filenames = dlg.selectedFiles()
            if filenames:
                self.StudyPALS_file_path_box.setText(filenames[0])

    @staticmethod
    def file_error_dialog():
        msg = QMessageBox()

        msg.setText("Invalid filepath")
        msg.setWindowTitle("Filepath error")
        msg.setStandardButtons(QMessageBox.Ok)

        msg.exec_()

    def import_file(self):
        if self.file_path_box.text() != '':
            if self.settings_flag:
                self.open_settings_click()
            try:
                self.df = read_file(self.file_path_box.text(), self.leaders_uuns, self.sessions_to_drop, self.sessions_to_merge)
                if self.settings_flag is False:
                    self.mailing_list_button.show()
                    self.regulars_list_button.show()
                    self.get_attendance_button.show()
                    self.cb_sem.show()
                    self.cb_week.show()
                    self.cb_reg.show()
                    self.lb_sem.show()
                    self.lb_week.show()
                    self.graph_button.show()
                    self.settings_button.show()
                    self.export_button.show()
                    self.StudyPALS_file_path_box.show()
                    self.lb_StudyPALS.show()
                    self.StudyPALS_select_file_button.show()
                else:
                    self.open_settings()
            except:
                print(sys.exc_info())  # temporary workaround, will make the error dialogs better later
                self.file_error_dialog()

    def studypals_import_file(self):
        if self.StudyPALS_imported:
            self.import_file()
        if self.StudyPALS_file_path_box.text() != '':
            if self.settings_flag:
                self.open_settings_click()
            try:
                self.df = studypals_read_file(self.StudyPALS_file_path_box.text(), self.df, self.leaders_uuns)
                if self.settings_flag is False:
                    self.mailing_list_button.show()
                    self.regulars_list_button.show()
                    self.get_attendance_button.show()
                    self.cb_sem.show()
                    self.cb_week.show()
                    self.cb_reg.show()
                    self.lb_sem.show()
                    self.lb_week.show()
                    self.graph_button.show()
                    self.settings_button.show()
                    self.export_button.show()
                    self.StudyPALS_imported = True
                else:
                    self.open_settings()
            except:
                self.file_error_dialog()

    def graph_click(self):
        sessions, attendance = zip(*attendance_count(self.df))
        plt.plot(attendance)
        plt.ylabel('Attendance')
        plt.axvline(x=3, linestyle=':', color='black')
        for xc in range(8, len(sessions), 4):
            plt.axvline(x=xc, linestyle=':', color='black')
        plt.xticks(range(0, len(sessions)), sessions, rotation='vertical')
        plt.show()

    @staticmethod
    def get_attendance_error(e=None):
        # Revise the error system
        msg = QMessageBox()
        if e is None:
            msg.setText("Invalid username entered")
            msg.setWindowTitle("Invalid username error")
        else:
            msg.setText("An error occurred \n{}".format(e))
            msg.setWindowTitle("Error")
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def get_attendance_click(self):
        inp = QInputDialog.getText(self, "Choose an email", "Choose an email to look up")

        # check that user input something
        if inp[1]:
            self.textEdit.clear()
            try:
                attendance_text = get_attendance(self.df, inp[0])
                if attendance_text != "Invalid username entered":
                    self.textEdit.insertPlainText(attendance_text)
                else:
                    self.get_attendance_error()
            except:
                self.get_attendance_error(e=sys.exc_info()[0])

    #
    # ------------Settings mode buttons------------
    #

    def list_sessions_click(self):
        self.textEdit.clear()
        i = 1
        attendance = attendance_count(self.df)
        for column in self.df.columns.values:
            if 'w' in column:
                self.textEdit.insertPlainText(str(i) + ') ' + column + " --- " + str(attendance[i - 1][1]) + '\n')
                i = i + 1
        self.number_of_sessions = i
        self.delete_session_button.show()
        self.merge_sessions_button.show()
        self.add_session_button.show()

    @staticmethod
    def session_index_error_dialog(error):
        msg = QMessageBox()

        if error == 'range':
            msg.setText("Session index out of range")
            msg.setWindowTitle("Index error")
        else:
            msg.setText("Please enter an integer")
            msg.setWindowTitle("Index error")

        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def delete_session_click(self):
        inp = QInputDialog.getText(self, "Select a session to delete", "Choose a session index")

        try:
            session_index = int(inp[0])
            if self.number_of_sessions >= session_index > 0:
                self.df = self.df.drop(columns=self.df.columns.values[(session_index + 7)])
                self.list_sessions_click()
            else:
                self.session_index_error_dialog('range')
        except:
            self.session_index_error_dialog('type')

    def merge_sessions_click(self):
        inp = QInputDialog.getText(self, "Select 2 sessions to merge", "Input 2 indices separated by a comma")

        try:
            session1, session2 = inp[0].split(',')  # the [0] is getting the string out of the tuple (string, boolean)
            session1 = int(session1)
            session2 = int(session2)
            if (
                    0 < session1 <= self.number_of_sessions and 0 < session2 <= self.number_of_sessions):
                self.df = merge_sessions(self.df, self.df.columns.values[(session1 + 7)],
                                         self.df.columns.values[(session2 + 7)])  # from TopHat.py
                self.list_sessions_click()
            else:
                self.session_index_error_dialog('range')
        except:
            self.session_index_error_dialog('type')

    def add_session_click(self):
        inp = QInputDialog.getText(self, "Choose a session index", "Input an index to insert at")

        try:
            session_index = int(inp[0])
            if self.number_of_sessions >= session_index > 0:
                self.df.insert(session_index + 7, 'Lecture w', 'A')  # inputs an empty session at a given index
                self.df = rename_sessions(self.df, True)  # renames the sessions accordingly
                for i in range(0, len(self.df.columns)):
                    print(self.df.columns[i])
                self.list_sessions_click()
            else:
                self.session_index_error_dialog('range')
        except:
            self.session_index_error_dialog('type')

    #
    # ------------Export button------------
    #

    @staticmethod
    def export_error_dialog(e):
        msg = QMessageBox()

        msg.setText("Could not export the data frame {}".format(e))
        msg.setWindowTitle("Export Error")
        msg.setStandardButtons(QMessageBox.Ok)

        msg.exec_()

    @staticmethod
    def export_success_dialog(filename):
        msg = QMessageBox()

        msg.setText("Exported data frame to {}".format(filename))
        msg.setWindowTitle("Export Success")
        msg.setStandardButtons(QMessageBox.Ok)

        msg.exec_()

    def export_click(self):
        try:
            filename = 'TopHat_export.xls'
            self.df.to_excel(filename, sheet_name='Attendance')
            self.export_success_dialog(filename)
        except:
            self.export_error_dialog(e=sys.exc_info()[0])

    #
    # ------------Settings mode switch------------
    #

    def open_settings_click(self):
        self.textEdit.clear()
        if self.settings_flag is False:
            self.mailing_list_button.hide()
            self.regulars_list_button.hide()
            self.get_attendance_button.hide()
            self.cb_sem.hide()
            self.cb_week.hide()
            self.cb_reg.hide()
            self.lb_sem.hide()
            self.lb_week.hide()
            self.graph_button.hide()
            self.session_list_button.show()
            self.settings_flag = True
        else:
            self.mailing_list_button.show()
            self.regulars_list_button.show()
            self.get_attendance_button.show()
            self.cb_sem.show()
            self.cb_week.show()
            self.cb_reg.show()
            self.lb_sem.show()
            self.lb_week.show()
            self.graph_button.show()
            self.session_list_button.hide()
            self.delete_session_button.hide()
            self.merge_sessions_button.hide()
            self.add_session_button.hide()
            self.settings_flag = False

    #
    # ------------Settings Initialisation------------
    #

    @staticmethod
    def settings_error_dialog():
        msg = QMessageBox()

        msg.setText("Could not load settings")
        msg.setWindowTitle("Settings initialisation error")
        msg.setStandardButtons(QMessageBox.Ok)

        msg.exec_()

    @staticmethod
    def write_to_settings(missing_line=''):
        try:

            if missing_line:
                settings_file = open('settings.txt', 'a')
                settings_file.write(f"\n{missing_line} = ")
            else:
                settings_file = open('settings.txt', 'w+')
                settings_file.write('[Defaults]')
                settings_file.write("\nEconPALS_filepath = ")
                settings_file.write("\nStudyPALS_filepath = ")
                settings_file.write("\nleaders_uuns = ")
                settings_file.write("\nsessions_to_drop = ")
                settings_file.write("\nsessions_to_merge = ")

            settings_file.close()
        except:
            print("Could not write to settings")

    def initialise_settings(self):
        try:
            config = configparser.ConfigParser()
            config.read_file(open(r'settings.txt'))
            settings_EconPALS_filepath = config.get('Defaults', 'EconPALS_filepath')
            settings_StudyPALS_filepath = config.get('Defaults', 'StudyPALS_filepath')
            self.default_EconPALS_filepath = settings_EconPALS_filepath
            self.default_StudyPALS_filepath = settings_StudyPALS_filepath
            self.leaders_uuns = config.get('Defaults', 'leaders_uuns').split(',')
            self.sessions_to_drop = config.get('Defaults', 'sessions_to_drop').split(',')
            self.sessions_to_merge = config.get('Defaults', 'sessions_to_merge').split(',')

        except (configparser.NoSectionError, FileNotFoundError, configparser.MissingSectionHeaderError):
            self.write_to_settings()
            self.initialise_settings()
        except configparser.NoOptionError as e:
            self.write_to_settings(missing_line=str(e).split("'")[1])
            self.initialise_settings()

    #
    # ------------Initialise UI------------
    #

    def initialise_ui(self):
        #
        # ------------Initialise the main window------------
        #

        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)

        #
        # ------------Initialise the TextEdit widget------------
        #
        self.textEdit = QTextEdit(self)  # main output field
        self.textEdit.setGeometry(0, 0, 600, 600)

        #
        # ------------Initialise the file selection buttons, fields and labels ------------
        #
        self.lb_EconPALS = QLabel(self)
        self.lb_EconPALS.setText('EconPALS')
        self.lb_EconPALS.move(650, 15)

        self.file_path_box = QLineEdit(self)
        self.file_path_box.setGeometry(650, 50, 425, 50)
        self.file_path_box.setText(self.default_EconPALS_filepath)
        self.file_path_box.returnPressed.connect(self.import_file)
        self.file_path_box.setFocus()

        self.select_file_button = QPushButton('...', self)
        self.select_file_button.clicked.connect(self.select_file)
        self.select_file_button.setGeometry(1100, 50, 50, 50)

        self.lb_StudyPALS = QLabel(self)
        self.lb_StudyPALS.setText('StudyPALS')
        self.lb_StudyPALS.move(650, 115)
        self.lb_StudyPALS.hide()

        self.StudyPALS_file_path_box = QLineEdit(self)
        self.StudyPALS_file_path_box.setGeometry(650, 150, 425, 50)
        self.StudyPALS_file_path_box.setText(self.default_StudyPALS_filepath)
        self.StudyPALS_file_path_box.returnPressed.connect(self.studypals_import_file)
        self.StudyPALS_file_path_box.hide()

        self.StudyPALS_select_file_button = QPushButton('...', self)
        self.StudyPALS_select_file_button.clicked.connect(self.studypals_select_file)
        self.StudyPALS_select_file_button.setGeometry(1100, 150, 50, 50)
        self.StudyPALS_select_file_button.hide()

        #
        # ------------Initialise the main menu buttons, labels and combo boxes------------
        #
        self.mailing_list_button = QPushButton('Mailing list', self)
        self.mailing_list_button.clicked.connect(self.emails_button)
        self.mailing_list_button.setGeometry(650, 250, 200, 50)
        self.mailing_list_button.hide()

        self.regulars_list_button = QPushButton('Regulars', self)
        self.regulars_list_button.clicked.connect(self.regulars_button)
        self.regulars_list_button.setGeometry(650, 350, 200, 50)
        self.regulars_list_button.hide()

        self.get_attendance_button = QPushButton('Get attendance', self)
        self.get_attendance_button.clicked.connect(self.get_attendance_click)
        self.get_attendance_button.setGeometry(650, 450, 200, 50)
        self.get_attendance_button.hide()

        self.lb_sem = QLabel(self)
        self.lb_sem.setText('Semester')
        self.lb_sem.move(875, 215)
        self.lb_sem.hide()

        self.lb_week = QLabel(self)
        self.lb_week.setText('Week')
        self.lb_week.move(1025, 215)
        self.lb_week.hide()

        self.cb_sem = QComboBox(self)
        self.cb_sem.addItems(['1', '2'])
        self.cb_sem.setGeometry(875, 250, 125, 50)
        self.cb_sem.hide()

        self.cb_week = QComboBox(self)
        self.cb_week.addItems(['All', '2', '3', '4', '5', '6', '7', '8', '9', '10'])
        self.cb_week.setGeometry(1025, 250, 100, 50)
        self.cb_week.hide()

        self.cb_reg = QComboBox(self)
        self.cb_reg.addItems(['2', '3', '4', '5', '6', '7', '8', '9', '10'])
        self.cb_reg.setGeometry(875, 350, 125, 50)
        self.cb_reg.hide()

        self.graph_button = QPushButton('Graph', self)
        self.graph_button.clicked.connect(self.graph_click)
        self.graph_button.setGeometry(1050, 450, 150, 50)
        self.graph_button.hide()

        #
        # ------------Initialise the settings mode switch button------------
        #

        self.settings_button = QPushButton('Settings', self)
        self.settings_button.clicked.connect(self.open_settings_click)
        self.settings_button.setGeometry(1050, 550, 150, 50)
        self.settings_button.hide()

        #
        # ------------Initialise the settings mode buttons------------
        #

        self.session_list_button = QPushButton('List Sessions', self)
        self.session_list_button.clicked.connect(self.list_sessions_click)
        self.session_list_button.setGeometry(650, 250, 200, 50)
        self.session_list_button.hide()

        self.delete_session_button = QPushButton('Delete Session', self)
        self.delete_session_button.clicked.connect(self.delete_session_click)
        self.delete_session_button.setGeometry(875, 250, 200, 50)
        self.delete_session_button.hide()

        self.merge_sessions_button = QPushButton('Merge Sessions', self)
        self.merge_sessions_button.clicked.connect(self.merge_sessions_click)
        self.merge_sessions_button.setGeometry(650, 350, 200, 50)
        self.merge_sessions_button.hide()

        self.add_session_button = QPushButton('Add Session', self)
        self.add_session_button.clicked.connect(self.add_session_click)
        self.add_session_button.setGeometry(875, 350, 200, 50)
        self.add_session_button.hide()

        #
        # ------------Initialise the export button------------
        #
        self.export_button = QPushButton('Export to Excel', self)
        self.export_button.clicked.connect(self.export_click)
        self.export_button.setGeometry(800, 550, 225, 50)
        self.export_button.hide()

        self.show()


#
# ------------Main------------
#

if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setApplicationName("EconPALS")
    ex = App()
    sys.exit(app.exec_())
