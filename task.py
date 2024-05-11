import sys
import os
import openpyxl
from PySide6.QtWidgets import QApplication, QWidget, QLabel, QVBoxLayout, QPushButton, QFileDialog, QTableWidgetItem, QLineEdit, QDateEdit, QTableWidget, QComboBox, QHBoxLayout, QMenu, QMessageBox
from PySide6.QtCore import QDate, QTimer, Qt
from tapd import tapdTask, USER, COOKIE, STORY, PROJECT


class MainWindow(QWidget):
    sheet = ""
    dataIndex = 0

    def __init__(self):
        super().__init__()
        self.setWindowTitle('tapd批量task生成 by GKK')
        self.setMinimumHeight(600)
        # 左侧布局
        left_layout = QVBoxLayout()
        self.data_show = QTableWidget()
        self.data_show.setMinimumWidth(720)
        self.data_show.setColumnCount(6)
        for i, j in enumerate(["名称", "花费", "负责人", "开始", "结束", "状态/tid"]):
            self.data_show.setHorizontalHeaderItem(i, QTableWidgetItem(j))
        self.data_show.setColumnWidth(0, 300)
        self.data_show.setColumnWidth(1, 30)
        self.data_show.setColumnWidth(2, 70)
        self.data_show.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.data_show.customContextMenuRequested.connect(self.Menu)
        left_layout.addWidget(self.data_show)
        # 右侧布局
        right_layout = QVBoxLayout()
        #
        self.user_input = QLineEdit()
        self.user_input.setText(USER)
        user_layout = QHBoxLayout()
        user_layout.addWidget(QLabel('默认负责人'))
        user_layout.addWidget(self.user_input)
        right_layout.addLayout(user_layout)
        #
        self.cookie_input = QLineEdit()
        self.cookie_input.setText(COOKIE)
        cookie_layout = QHBoxLayout()
        cookie_layout.addWidget(QLabel('cookie'))
        cookie_layout.addWidget(self.cookie_input)
        right_layout.addLayout(cookie_layout)
        #
        self.story_input = QLineEdit()
        self.story_input.setText(STORY)
        story_layout = QHBoxLayout()
        story_layout.addWidget(QLabel('需求id'))
        story_layout.addWidget(self.story_input)
        right_layout.addLayout(story_layout)
        #
        self.project_input = QLineEdit()
        self.project_input.setText(PROJECT)
        project_layout = QHBoxLayout()
        project_layout.addWidget(QLabel('项目名称'))
        project_layout.addWidget(self.project_input)
        right_layout.addLayout(project_layout)
        #
        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        date_layout = QHBoxLayout()
        date_layout.addWidget(QLabel('开始日期'))
        date_layout.addWidget(self.date_edit)
        right_layout.addLayout(date_layout)
        # 文件选择框
        self.eFile_button = QPushButton('选择表格')
        self.eFile_button.clicked.connect(self.choose_file)
        right_layout.addWidget(self.eFile_button)
        # sheet
        self.combo = QComboBox(self)
        self.combo.currentTextChanged.connect(lambda x: setattr(self, "sheet", x))
        self.read_button = QPushButton('读取')
        self.read_button.clicked.connect(self.read)
        sheet_layout = QHBoxLayout()
        sheet_layout.addWidget(self.combo)
        sheet_layout.addWidget(self.read_button)
        right_layout.addLayout(sheet_layout)
        # 开始
        self.start_button = QPushButton('上传')
        self.start_button.clicked.connect(self.start)
        action_layout = QHBoxLayout()
        action_layout.addWidget(self.start_button)
        right_layout.addLayout(action_layout)

        # 添加右侧布局到主布局
        main_layout = QHBoxLayout()
        main_layout.addLayout(left_layout)
        main_layout.addLayout(right_layout)

        # 添加主布局和文本输出框到窗口
        layout = QVBoxLayout()
        layout.addLayout(main_layout)
        self.setLayout(layout)
        self.timer = QTimer()
        self.timer.timeout.connect(self.sender)

    def choose_file(self):
        self.eFile, _ = QFileDialog.getOpenFileName(self, '选择', '', 'Excel files (*.xlsx)')
        if self.eFile:
            self.eFile_button.setText(os.path.basename(self.eFile))
            ws = openpyxl.load_workbook(self.eFile)
            self.combo.addItems(ws.sheetnames)

    def read(self):
        self.tapd = tapdTask(self.story_input.text(), self.cookie_input.text())
        self.tapd.read(
            self.eFile,
            self.sheet,
            self.project_input.text(),
            self.user_input.text(),
            self.date_edit.text(),
        )
        self.data_show.setRowCount(len(self.tapd.datas))
        taskIds = self.tapd.taskIds()
        for i, row in enumerate(self.tapd.datas):
            for j, cel in enumerate(row):
                self.data_show.setItem(i, j, QTableWidgetItem(str(cel)))
                if j == 0 and cel in taskIds:
                    self.data_show.setItem(i, 5, QTableWidgetItem(taskIds[cel]))
        if len(taskIds) == 0:
            QMessageBox.information(self, "提醒", "没有发现历史任务，检查cookie/story是否正常？", QMessageBox.Ok | QMessageBox.Cancel)
        self.dataIndex = 0
        self.start_button.setText("上传")

    def start(self):
        if self.start_button.text() == "上传":
            self.timer.start(500)
            self.start_button.setText("暂停")
        else:
            self.timer.stop()
            self.start_button.setText("上传")

    def sender(self):
        if len(self.tapd.datas) > self.dataIndex:
            item = self.data_show.item(self.dataIndex, 5)
            if item and item.text() == "完成":
                self.dataIndex += 1
                return
            res = self.tapd.createOne(self.tapd.datas[self.dataIndex])
            self.data_show.setItem(self.dataIndex, 5, QTableWidgetItem(res))
            self.dataIndex += 1

    def Menu(self, pos):
        menu = QMenu()
        item1 = menu.addAction(u'完成')
        action = menu.exec(self.data_show.mapToGlobal(pos))
        if action == item1:
            for i in self.data_show.selectedIndexes():
                if i.column() == 5:
                    res = self.tapd.done(i.data())
                    self.data_show.setItem(i.row(), 5, QTableWidgetItem(res))
            self.tapd.save()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())
