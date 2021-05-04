import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtPrintSupport import *
import docx2txt

class Word(QMainWindow):
    def __init__(self):
        super(Word, self).__init__()
        self.editor=QTextEdit()
        self.setCentralWidget(self.editor)
        self.showMaximized()
        self.title = 'Microsoft Word'
        self.setWindowTitle(self.title)
        self.tool_bar()

    def tool_bar(self):
        tool=QToolBar()
        
        undo_action=QAction(QIcon('undo.png'),'undo',self)
        undo_action.triggered.connect(self.editor.undo)
        tool.addAction(undo_action)
        
        redo_action=QAction(QIcon('redo.png'),'redo',self)
        redo_action.triggered.connect(self.editor.redo)
        tool.addAction(redo_action)

        tool.addSeparator()

        cut_action = QAction(QIcon('cut.png'), 'cut', self)
        cut_action.triggered.connect(self.editor.cut)
        tool.addAction(cut_action)

        copy_action = QAction(QIcon('copy.png'), 'copy', self)
        copy_action.triggered.connect(self.editor.copy)
        tool.addAction(copy_action)

        paste_action = QAction(QIcon('paste.png'), 'paste', self)
        paste_action.triggered.connect(self.editor.paste)
        tool.addAction(paste_action)

        tool.addSeparator()

        self.font_combo=QComboBox()
        self.font_combo.addItems(["Courier Std", "Hellentic Typewriter Regular", "Helvetica", "Arial", "SansSerif", "Helvetica", "Times", "Monospace"])
        self.font_combo.activated.connect(self.font_change)
        tool.addWidget(self.font_combo)

        self.font_size=QSpinBox()
        self.font_size.setValue(12)
        self.font_size.valueChanged.connect(self.value_change)
        tool.addWidget(self.font_size)

        bold_action=QAction(QIcon('bold.png'),'bold',self)
        bold_action.triggered.connect(self.bold)
        tool.addAction(bold_action)

        italic=QAction(QIcon('italic.png'),'italic',self)
        italic.triggered.connect(self.italic_change)
        tool.addAction(italic)

        underline=QAction(QIcon('underline.png'),'underline',self)
        underline.triggered.connect(self.under)
        tool.addAction(underline)

        zoom_in=QAction(QIcon('zoom-in.png'),'zoom-in',self)
        zoom_in.triggered.connect(self.editor.zoomIn)
        tool.addAction(zoom_in)

        zoom_out = QAction(QIcon('zoom-out.png'), 'zoom-out', self)
        zoom_out.triggered.connect(self.editor.zoomOut)
        tool.addAction(zoom_out)

        right_alignment_action=QAction(QIcon('right-align.png'),'right-align',self)
        right_alignment_action.triggered.connect(lambda: self.editor.setAlignment(Qt.AlignRight))
        tool.addAction(right_alignment_action)

        left_alignment_action=QAction(QIcon('left-align.png'),'left-align',self)
        left_alignment_action.triggered.connect(lambda : self.editor.setAlignment(Qt.AlignLeft))
        tool.addAction(left_alignment_action)

        justification_action=QAction(QIcon('justification.png'),'justification',self)
        justification_action.triggered.connect(lambda: self.editor.setAlignment(Qt.AlignCenter))
        tool.addAction(justification_action)

        
        self.addToolBar(tool)

        self.create_menubar()

    def italic_change(self):
        ital=self.editor.fontItalic()
        self.editor.setFontItalic(not(ital))

    def bold(self):
        if self.editor.fontWeight()!=QFont.Bold:
            self.editor.setFontWeight(QFont.Bold)
            return
        self.editor.setFontWeight(QFont.Normal)

    def under(self):
        state=self.editor.fontUnderline()
        self.editor.setFontUnderline(not(state))

    def value_change(self):
        val=self.font_size.value()
        self.editor.setFontPointSize(val)

    def font_change(self):
        s=self.font_combo.currentText()
        self.editor.setCurrentFont(QFont(s))

    def create_menubar(self):
        menu=QMenuBar()

        file=QMenu('file',self)
        menu.addMenu(file)

        save_action=QAction('save',self)
        save_action.triggered.connect(self.save_file)
        file.addAction(save_action)

        rename_action=QAction('rename',self)
        rename_action.triggered.connect(self.saveas)
        file.addAction(rename_action)

        open_action=QAction('open',self)
        open_action.triggered.connect(self.open)
        file.addAction(open_action)

        pdf_action=QAction('save as pdf',self)
        pdf_action.triggered.connect(self.save_pdf)
        file.addAction(pdf_action)

        edit=QMenu('edit',self)
        menu.addMenu(edit)

        paste=QAction('paste',self)
        paste.triggered.connect(self.editor.paste)
        edit.addAction(paste)

        clear = QAction('clear', self)
        clear.triggered.connect(self.editor.clear)
        edit.addAction(clear)

        select = QAction('select', self)
        select.triggered.connect(self.editor.selectAll)
        edit.addAction(select)
        
        view=QMenu('view',self)
        menu.addMenu(view)

        fullscr=QAction('fullscr',self)
        fullscr.triggered.connect(lambda: self.showFullScreen())
        view.addAction(fullscr)
        
        norscr=QAction('norscr',self)
        norscr.triggered.connect(lambda: self.showNormal())
        view.addAction(norscr)

        minscr = QAction('minscr', self)
        minscr.triggered.connect(lambda: self.showMinimized())
        view.addAction(minscr)



        self.setMenuBar(menu)

    def save_file(self):
        print(self.path)
        if self.path=='':
            self.saveas
        text=self.editor.toPlainText()
        try:
            with open(self.path,'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
            print(e)

    def saveas(self):
        self.path, _ =QFileDialog.getSaveFileName(self,'save file',"",'text documents (*.text);Text documents (*.txt);All files(*.*)')
        if self.path=='':
            return
        text=self.editor.toPlainText()
        try:
            with open(self.path,'w') as f:
                f.write(text)
                self.update_title
        except Exception as e:
            print(e)

    def save_pdf(self):
        fname, _ =QFileDialog.getSaveFileName(self,'Export PDF',None,'PDF Files(*.pdf);;All Files()')
        print(fname)
        if fname!='':
            printer=QPrinter(QPrinter.HighResolution)
            printer.setOutputFormat(QPrinter.PdfFormat)
            printer.setOutputFileName(fname)
            self.editor.document().print_(printer)

    def open(self):
        self.path, _ =QFileDialog.getSaveFileName(self,'open file',"",'text documents(*.text);Text documents(*.txt);All Files()')
        try:
            text=docx2txt.preocess(self.path)
        except Exception as e:
            print(e)
        else:
            self.editor.setText(text)
            self.update_title()



    def update_title(self):
        self.setWindowTitle(self.title+' '+self.path)



app=QApplication(sys.argv)
window=Word()
window.show()
sys.exit(app.exec_())

