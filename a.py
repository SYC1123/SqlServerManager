from PyQt5.QtWidgets import QVBoxLayout, QLineEdit, QTextBrowser, QDialog


class Form(QDialog):

    def __init__(self, parent=None):
        super(Form, self).__init__(parent)
        self.browser = QTextBrowser()
        self.lineedit = QLineEdit("Type an expression and press Enter")
        self.lineedit.selectAll()
        layout = QVBoxLayout()  # 垂直盒式布局
        layout.addWidget(self.browser)
        layout.addWidget(self.lineedit)
        # layout = QGridLayout() #网格布局
        # layout.addWidget(self.browser，0, 0)
        # layout.addWidget(self.lineedit，0, 0)
        self.setLayout(layout)
        self.lineedit.setFocus()
        self.connect(self.lineedit, SIGNAL("returnPressed()"), self.updateUi)  # 信号绑定到槽
        self.setWindowTitle("Calculate")

    def updateUi(self):
        try:
            text = unicode(self.lineedit.text())
            self.browser.append("%s = <b>%s</b>" % (text, eval(text)))
        except:
            self.browser.append(
                "<font color=red>%s is invalid!</font>" % text)