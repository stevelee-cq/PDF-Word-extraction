import sys, os
import pdfplumber
import spacy
import nltk
from collections import Counter
from nltk.corpus import words as nltk_words
from nltk.corpus import stopwords
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton, QLabel, QVBoxLayout,
    QWidget, QFileDialog, QTextEdit, QCheckBox, QMessageBox, QProgressBar,
    QHBoxLayout, QLineEdit
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from wordcloud import WordCloud
import matplotlib.pyplot as plt

# ========== 资源初始化 ==========
nltk.download('words')
nltk.download('stopwords')
nlp = spacy.load("en_core_web_sm")
english_vocab = set(w.lower() for w in nltk_words.words())
stop_words = set(stopwords.words('english'))

# ========== 后台提取线程 ==========
class ExtractWorker(QThread):
    progress = pyqtSignal(int)
    result = pyqtSignal(Counter, str)  # Counter结果 + PDF路径

    def __init__(self, pdf_path, start_page, end_page):
        super().__init__()
        self.pdf_path = pdf_path
        self.start_page = start_page
        self.end_page = end_page

    def run(self):
        try:
            word_counter = Counter()
            with pdfplumber.open(self.pdf_path) as pdf:
                pages = pdf.pages[self.start_page - 1 : self.end_page]
                for i, page in enumerate(pages):
                    text = page.extract_text()
                    if text:
                        doc = nlp(text)
                        for token in doc:
                            if token.is_alpha and token.is_ascii and len(token) > 1:
                                lemma = token.lemma_.lower()
                                if lemma in english_vocab and lemma not in stop_words:
                                    word_counter[lemma] += 1
                    self.progress.emit(int((i + 1) / len(pages) * 100))
            self.result.emit(word_counter, self.pdf_path)
        except Exception as e:
            self.result.emit(Counter({"❌ 提取失败": 1}), self.pdf_path)

# ========== 主窗口类 ==========
class PDFWordExtractor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("英文PDF单词词频提取器（词云可视化 | 页码选择 | 美观夜间风格）")
        self.setGeometry(100, 100, 880, 700)
        self.set_dark_theme_style()

        # 控件定义
        self.label = QLabel("📄 请选择PDF文献：")
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)

        self.start_page_input = QLineEdit()
        self.start_page_input.setPlaceholderText("起始页（从1开始）")
        self.end_page_input = QLineEdit()
        self.end_page_input.setPlaceholderText("结束页")

        self.select_button = QPushButton("选择PDF")
        self.extract_button = QPushButton("提取并统计词频")
        self.wordcloud_button = QPushButton("生成词云图")
        self.wordcloud_button.setEnabled(False)
        self.save_checkbox = QCheckBox("提取后保存为 .txt 文件")
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)

        # 页码输入区域
        page_layout = QHBoxLayout()
        page_layout.addWidget(self.start_page_input)
        page_layout.addWidget(self.end_page_input)

        # 主布局
        layout = QVBoxLayout()
        layout.addWidget(self.label)
        layout.addWidget(self.select_button)
        layout.addLayout(page_layout)
        layout.addWidget(self.save_checkbox)
        layout.addWidget(self.extract_button)
        layout.addWidget(self.wordcloud_button)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.text_edit)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

        # 绑定事件
        self.select_button.clicked.connect(self.select_pdf)
        self.extract_button.clicked.connect(self.extract_words)
        self.wordcloud_button.clicked.connect(self.show_wordcloud)

        self.pdf_path = ""
        self.worker = None
        self.total_pages = 0
        self.word_counter = None

    def set_dark_theme_style(self):
        self.setStyleSheet("""
            QWidget {
                background-color: #1e1e1e;
                color: #ffffff;
                font-family: 'Segoe UI';
                font-size: 10.5pt;
            }
            QPushButton {
                background-color: #3A99D8;
                color: white;
                border-radius: 5px;
                padding: 6px;
            }
            QPushButton:hover {
                background-color: #5FB3E7;
            }
            QTextEdit {
                background-color: #282c34;
                border: 1px solid #444;
                color: #f8f8f2;
                font-family: Consolas;
                font-size: 10pt;
            }
            QLineEdit {
                background-color: #2e2e2e;
                border: 1px solid #555;
                border-radius: 4px;
                padding: 4px;
                color: white;
            }
            QProgressBar {
                border: 1px solid #555;
                border-radius: 5px;
                text-align: center;
                background-color: #2e2e2e;
                color: white;
            }
            QProgressBar::chunk {
                background-color: #5CB85C;
                border-radius: 5px;
            }
            QCheckBox {
                spacing: 5px;
            }
        """)

    def select_pdf(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "选择PDF文件", "", "PDF Files (*.pdf)")
        if file_path:
            self.pdf_path = file_path
            self.label.setText(f"📄 当前文件：{os.path.basename(file_path)}")
            try:
                with pdfplumber.open(self.pdf_path) as pdf:
                    self.total_pages = len(pdf.pages)
                self.text_edit.append(f"✅ 已加载文件（共 {self.total_pages} 页）：{self.pdf_path}\n")
            except Exception as e:
                self.text_edit.append(f"❌ 无法读取PDF页数：{e}")

    def extract_words(self):
        if not self.pdf_path:
            QMessageBox.warning(self, "⚠️ 未选择文件", "请先选择一个PDF文件")
            return

        try:
            start_page = int(self.start_page_input.text().strip())
            end_page = int(self.end_page_input.text().strip())
        except ValueError:
            QMessageBox.warning(self, "❌ 页码格式错误", "请输入有效的起始页和结束页（正整数）")
            return

        if start_page < 1 or end_page < start_page or end_page > self.total_pages:
            QMessageBox.warning(self, "❌ 页码范围错误", f"页码范围应在 1 ~ {self.total_pages} 之间，且起始页不大于结束页")
            return

        self.text_edit.clear()
        self.progress_bar.setValue(0)
        self.wordcloud_button.setEnabled(False)
        self.text_edit.append(f"⏳ 正在分析第 {start_page} 页至第 {end_page} 页内容...\n")

        self.worker = ExtractWorker(self.pdf_path, start_page, end_page)
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.result.connect(self.display_result)
        self.worker.start()

    def display_result(self, word_counter: Counter, pdf_path: str):
        if "❌ 提取失败" in word_counter:
            self.text_edit.append("❌ 提取失败，请检查PDF是否包含可识别文本内容。")
            self.wordcloud_button.setEnabled(False)
            return

        self.word_counter = word_counter
        self.wordcloud_button.setEnabled(True)

        total_unique = len(word_counter)
        total_count = sum(word_counter.values())

        self.text_edit.append(f"✅ 有效英文总词数：{total_count}，唯一词汇：{total_unique} 个\n")
        sorted_items = word_counter.most_common()

        for word, freq in sorted_items:
            self.text_edit.append(f"{word:<20} {freq}")

        if self.save_checkbox.isChecked():
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            out_path = os.path.join(os.path.dirname(pdf_path), base_name + "_词频统计.txt")
            with open(out_path, "w", encoding="utf-8") as f:
                f.write(f"总词数：{total_count}，唯一词汇：{total_unique} 个\n\n")
                for word, freq in sorted_items:
                    f.write(f"{word:<20} {freq}\n")
            self.text_edit.append(f"\n💾 结果已保存至：{out_path}")

        self.progress_bar.setValue(100)

    def show_wordcloud(self):
        if not self.word_counter or sum(self.word_counter.values()) == 0:
            QMessageBox.information(self, "提示", "无有效词频数据，请先提取词频。")
            return

        wc = WordCloud(
            width=800,
            height=400,
            background_color='white',
            max_words=200,
            colormap='viridis'
        )
        wc.generate_from_frequencies(self.word_counter)

        plt.figure(figsize=(10, 5))
        plt.imshow(wc, interpolation="bilinear")
        plt.axis("off")
        plt.title("词频词云", fontsize=16)
        plt.tight_layout()
        plt.show()

# ========== 启动程序 ==========
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFWordExtractor()
    window.show()
    sys.exit(app.exec_())
