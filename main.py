import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtPrintSupport import *

class BSWord(QMainWindow):
    def __init__(self):
        super(BSWord, self).__init__()
        self.editor=QTextEdit()
        self.editor.setFontPointSize(20)
        self.setCentralWidget(self.editor)
        self.font_size_box=QSpinBox()
        self.showMaximized()
        self.setWindowTitle('My BS word')
        self.create_tool_bar()

    def create_tool_bar(self):
        tool_bar=QToolBar()

        undo_action=QAction(QIcon('undo.png'),'undo',self)
        undo_action.triggered.connect(self.editor.undo)
        tool_bar.addAction(undo_action)

        redo_action=QAction(QIcon('redo.png'),'redo',self)
        redo_action.triggered.connect(self.editor.redo)
        tool_bar.addAction(redo_action)

        tool_bar.addSeparator()
        tool_bar.addSeparator()

        cut_action = QAction(QIcon('cut.png'), 'cut', self)
        cut_action.triggered.connect(self.editor.cut)
        tool_bar.addAction(cut_action)

        copy_action = QAction(QIcon('copy.png'), 'copy', self)
        copy_action.triggered.connect(self.editor.copy)
        tool_bar.addAction(copy_action)

        paste_action = QAction(QIcon('paste.png'), 'paste', self)
        paste_action.triggered.connect(self.editor.paste)
        tool_bar.addAction(paste_action)

        tool_bar.addSeparator()
        tool_bar.addSeparator()

        self.font_size_box.setValue(20)
        self.font_size_box.valueChanged.connect(self.set_size)
        tool_bar.addWidget(self.font_size_box)

        self.addToolBar(tool_bar)
        self.create_menu()

    def set_size(self):
        value=self.font_size_box.value()
        self.editor.setFontPointSize(value)

    def create_menu(self):
        menu=QMenuBar()
        file=QMenu('File',self)
        menu.addMenu(file)
        save_as_pdf_action=QAction('save_as_pdf',self)
        save_as_pdf_action.triggered.connect(self.save_as_pdf)
        file.addAction(save_as_pdf_action)
        edit = QMenu('edit', self)
        menu.addMenu(edit)
        view = QMenu('view', self)
        menu.addMenu(view)
        self.setMenuBar(menu)

    def save_as_pdf(self):
        file_path,_=QFileDialog.getSaveFileName(self, 'Export PDF', None, 'PDF Files (*.pdf)')
        printer=QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(file_path)
        self.editor.document().print_(printer)


app=QApplication(sys.argv)
window=BSWord()
window.show()
sys.exit(app.exec_())