import sys
import tkinter as tk
from tkinter import filedialog
import os
from PyQt6.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPlainTextEdit, QPushButton, QFileDialog, QMessageBox, QHBoxLayout, QSizePolicy
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QTextCursor, QFont
from docxxx import DocxReview

class StreamRedirector:
    def __init__(self, signal):
        self.signal = signal

    def write(self, text):
        self.signal.emit(text)

    def flush(self):
        pass

class WorkerThread(QThread):
    progress = pyqtSignal(str)
    finished = pyqtSignal()
    error = pyqtSignal(str)
    
    def __init__(self, file_path, model_key, chunk_size, max_retries, retry_delay, max_workers):
        super().__init__()
        self.file_path = file_path
        self.model_key = model_key
        self.chunk_size = chunk_size
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        self.max_workers = max_workers
    
    def run(self):
        redirector = StreamRedirector(self.progress)
        sys.stdout = redirector
        
        try:
            DocxReview(self.file_path, self.model_key, self.chunk_size, self.max_retries, self.retry_delay, self.max_workers).run()
        except RuntimeError as e:
            self.error.emit(str(e))
        except Exception as e:
            self.error.emit(f"处理文档时发生错误: {str(e)}")
        finally:
            sys.stdout = sys.__stdout__
            self.finished.emit()

class ModelProgressApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('Risdc-GPT/docx review2.0')
        self.setGeometry(100, 100, 800, 600)  # 增加窗口大小

        # 主布局
        mainLayout = QVBoxLayout()
        mainLayout.setSpacing(10)

        # 文件路径输入区域
        fileLayout = QHBoxLayout()
        self.pathLabel = QLabel('文件路径：')
        self.pathLabel.setStyleSheet("font-size: 14px;")
        self.pathInput = QLineEdit(self)
        self.pathInput.setStyleSheet("font-size: 14px; padding: 5px; border: 1px solid #ccc; border-radius: 5px;")
        self.browseButton = QPushButton('选择文件', self)
        self.browseButton.setStyleSheet("font-size: 14px; padding: 5px; background-color: #4CAF50; color: white; border: none; border-radius: 5px;")
        self.browseButton.clicked.connect(self.browseFile)
        fileLayout.addWidget(self.pathLabel)
        fileLayout.addWidget(self.pathInput)
        fileLayout.addWidget(self.browseButton)
        mainLayout.addLayout(fileLayout)

        # 参数输入区域
        paramsLayout = QHBoxLayout()
        
        # 大模型Key输入
        keyLayout = QVBoxLayout()
        self.keyLabel = QLabel('Qwen大模型Key：')
        self.keyLabel.setStyleSheet("font-size: 14px;")
        self.keyInput = QLineEdit(self)
        self.keyInput.setStyleSheet("font-size: 14px; padding: 5px; border: 1px solid #ccc; border-radius: 5px;")
        keyLayout.addWidget(self.keyLabel)
        keyLayout.addWidget(self.keyInput)
        paramsLayout.addLayout(keyLayout)

        # 单次分析字数输入
        chunkLayout = QVBoxLayout()
        self.chunkSizeLabel = QLabel('单次分析字数：')
        self.chunkSizeLabel.setStyleSheet("font-size: 14px;")
        self.chunkSizeInput = QLineEdit(self)
        self.chunkSizeInput.setStyleSheet("font-size: 14px; padding: 5px; border: 1px solid #ccc; border-radius: 5px;")
        self.chunkSizeInput.setText("1500")  # 默认值
        chunkLayout.addWidget(self.chunkSizeLabel)
        chunkLayout.addWidget(self.chunkSizeInput)
        paramsLayout.addLayout(chunkLayout)

        # 最大重试次数输入
        retriesLayout = QVBoxLayout()
        self.retriesLabel = QLabel('最大重试次数：')
        self.retriesLabel.setStyleSheet("font-size: 14px;")
        self.retriesInput = QLineEdit(self)
        self.retriesInput.setStyleSheet("font-size: 14px; padding: 5px; border: 1px solid #ccc; border-radius: 5px;")
        self.retriesInput.setText("3")  # 默认值
        retriesLayout.addWidget(self.retriesLabel)
        retriesLayout.addWidget(self.retriesInput)
        paramsLayout.addLayout(retriesLayout)

        # 重试延迟输入
        delayLayout = QVBoxLayout()
        self.delayLabel = QLabel('重试延迟（秒）：')
        self.delayLabel.setStyleSheet("font-size: 14px;")
        self.delayInput = QLineEdit(self)
        self.delayInput.setStyleSheet("font-size: 14px; padding: 5px; border: 1px solid #ccc; border-radius: 5px;")
        self.delayInput.setText("5")  # 默认值
        delayLayout.addWidget(self.delayLabel)
        delayLayout.addWidget(self.delayInput)
        paramsLayout.addLayout(delayLayout)

        # 最大工作线程数输入
        workersLayout = QVBoxLayout()
        self.workersLabel = QLabel('最大工作线程数：')
        self.workersLabel.setStyleSheet("font-size: 14px;")
        self.workersInput = QLineEdit(self)
        self.workersInput.setStyleSheet("font-size: 14px; padding: 5px; border: 1px solid #ccc; border-radius: 5px;")
        self.workersInput.setText("5")  # 默认值
        workersLayout.addWidget(self.workersLabel)
        workersLayout.addWidget(self.workersInput)
        paramsLayout.addLayout(workersLayout)

        mainLayout.addLayout(paramsLayout)

        # 进度展示区
        self.progressLabel = QLabel('进度：')
        self.progressLabel.setStyleSheet("font-size: 14px;")
        mainLayout.addWidget(self.progressLabel)

        self.progressDisplay = QPlainTextEdit(self)
        self.progressDisplay.setReadOnly(True)
        self.progressDisplay.setStyleSheet("font-size: 14px; padding: 5px; border: 1px solid #ccc; border-radius: 5px;")
        self.progressDisplay.setMinimumHeight(300)  # 设置最小高度
        font = QFont("Courier")
        font.setPointSize(12)
        self.progressDisplay.setFont(font)
        mainLayout.addWidget(self.progressDisplay)

        # 开始按钮
        self.startButton = QPushButton('开始处理', self)
        self.startButton.setStyleSheet("font-size: 16px; padding: 10px; background-color: #4CAF50; color: white; border: none; border-radius: 5px;")
        self.startButton.clicked.connect(self.startProcess)
        mainLayout.addWidget(self.startButton)

        self.setLayout(mainLayout)

    def browseFile(self):
        # 打开文件选择对话框，支持 .doc 和 .docx 文件
        file_path, _ = QFileDialog.getOpenFileName(
            self, 
            "选择文件", 
            "", 
            "Word Documents (*.doc *.docx);;All Files (*)"
        )
        if file_path:
            self.pathInput.setText(file_path)

    def startProcess(self):
        file_path = self.pathInput.text()
        model_key = self.keyInput.text()
        chunk_size = self.chunkSizeInput.text()
        max_retries = self.retriesInput.text()
        retry_delay = self.delayInput.text()
        max_workers = self.workersInput.text()

        if file_path and model_key and chunk_size.isdigit() and max_retries.isdigit() and retry_delay.isdigit() and max_workers.isdigit():
            chunk_size = int(chunk_size)
            max_retries = int(max_retries)
            retry_delay = int(retry_delay)
            max_workers = int(max_workers)
            
            # 清空之前的进度显示
            self.progressDisplay.clear()
            
            self.worker = WorkerThread(file_path, model_key, chunk_size, max_retries, retry_delay, max_workers)
            self.worker.progress.connect(self.updateProgress)
            self.worker.finished.connect(self.onFinished)
            self.worker.error.connect(self.onError)
            self.worker.start()
            
            self.startButton.setEnabled(False)
        else:
            QMessageBox.warning(self, "输入错误", "请确保所有字段都已正确填写。")

    def updateProgress(self, text):
        # 所有输出都作为新行添加
        self.progressDisplay.appendPlainText(text.strip())
        
        # 滚动到底部
        self.progressDisplay.verticalScrollBar().setValue(self.progressDisplay.verticalScrollBar().maximum())
        QApplication.processEvents()  # 确保UI及时更新

    def onFinished(self):
        self.progressDisplay.appendPlainText("处理完成！")
        self.startButton.setEnabled(True)

    def onError(self, error_message):
        QMessageBox.critical(self, "错误", f"处理过程中发生错误：\n{error_message}")
        self.startButton.setEnabled(True)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = ModelProgressApp()
    window.show()
    sys.exit(app.exec())
