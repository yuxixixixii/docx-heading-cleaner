from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import Qt, QUrl
from PySide6.QtGui import QDesktopServices
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from docx_heading_cleaner import DocxCleanerError, clean_docx_all, default_output_path


APP_TITLE = "Word 导航标题清理器"


@dataclass
class FileJob:
    input_path: Path
    output_path: Path
    status: str = "待处理"
    result: str = ""


class DropArea(QFrame):
    def __init__(self, window: "MainWindow") -> None:
        super().__init__()
        self.window = window
        self.setAcceptDrops(True)
        self.setObjectName("dropArea")
        self.setMinimumHeight(150)
        self.label = QLabel("把 .docx 文件拖到这里\n支持一次拖入多个文件")
        self.label.setAlignment(Qt.AlignCenter)
        layout = QVBoxLayout(self)
        layout.addWidget(self.label)

    def dragEnterEvent(self, event) -> None:  # noqa: N802 - Qt method name
        if self._has_supported_file(event.mimeData().urls()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event) -> None:  # noqa: N802 - Qt method name
        paths = [Path(url.toLocalFile()) for url in event.mimeData().urls() if url.isLocalFile()]
        self.window.add_files(paths)
        event.acceptProposedAction()

    def _has_supported_file(self, urls: list[QUrl]) -> bool:
        return any(_is_supported_docx(Path(url.toLocalFile())) for url in urls if url.isLocalFile())


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.jobs: list[FileJob] = []
        self.setWindowTitle(APP_TITLE)
        self.resize(900, 560)
        self._build_ui()

    def _build_ui(self) -> None:
        root = QWidget()
        layout = QVBoxLayout(root)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        intro = QLabel("清理误入 Word 左侧导航的正文段落，原文件不会被覆盖。")
        intro.setObjectName("intro")
        layout.addWidget(intro)

        self.drop_area = DropArea(self)
        layout.addWidget(self.drop_area)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(["文件", "状态", "输出位置", "结果"])
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.verticalHeader().setVisible(False)
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        layout.addWidget(self.table)

        button_row = QHBoxLayout()
        self.add_button = QPushButton("添加文件")
        self.run_button = QPushButton("开始清理")
        self.clear_button = QPushButton("清空列表")
        self.open_folder_button = QPushButton("打开输出文件夹")
        self.open_folder_button.setEnabled(False)

        self.add_button.clicked.connect(self.choose_files)
        self.run_button.clicked.connect(self.process_jobs)
        self.clear_button.clicked.connect(self.clear_jobs)
        self.open_folder_button.clicked.connect(self.open_selected_output_folder)

        button_row.addWidget(self.add_button)
        button_row.addWidget(self.run_button)
        button_row.addWidget(self.clear_button)
        button_row.addStretch(1)
        button_row.addWidget(self.open_folder_button)
        layout.addLayout(button_row)

        self.summary = QLabel("等待添加文件")
        layout.addWidget(self.summary)

        root.setStyleSheet(
            """
            QWidget {
                font-family: "Microsoft YaHei", "Segoe UI", sans-serif;
                font-size: 13px;
            }
            QMainWindow {
                background: #f6f7f9;
            }
            #intro {
                color: #30343b;
            }
            #dropArea {
                border: 2px dashed #7993b5;
                border-radius: 8px;
                background: #ffffff;
                color: #2d405f;
                font-size: 18px;
            }
            QTableWidget {
                background: #ffffff;
                border: 1px solid #d7dce3;
                gridline-color: #e5e8ed;
            }
            QPushButton {
                min-height: 32px;
                padding: 0 14px;
            }
            """
        )
        self.setCentralWidget(root)

    def choose_files(self) -> None:
        files, _ = QFileDialog.getOpenFileNames(self, "选择 Word 文件", "", "Word 文档 (*.docx)")
        self.add_files([Path(file) for file in files])

    def add_files(self, paths: list[Path]) -> None:
        added = 0
        ignored = 0
        known = {job.input_path.resolve() for job in self.jobs}

        for path in paths:
            resolved = path.expanduser().resolve()
            if not _is_supported_docx(resolved) or resolved in known:
                ignored += 1
                continue
            self.jobs.append(FileJob(input_path=resolved, output_path=default_output_path(resolved)))
            known.add(resolved)
            added += 1

        self.refresh_table()
        if ignored and not added:
            QMessageBox.information(self, APP_TITLE, "没有添加文件。请拖入 .docx 文件，且不要拖入 Word 临时文件。")
        elif ignored:
            QMessageBox.information(self, APP_TITLE, f"已添加 {added} 个文件，忽略 {ignored} 个不支持或重复的文件。")

    def process_jobs(self) -> None:
        pending = [job for job in self.jobs if job.status in {"待处理", "失败", "跳过"}]
        if not pending:
            QMessageBox.information(self, APP_TITLE, "请先添加需要清理的 .docx 文件。")
            return

        self._set_buttons_enabled(False)
        QApplication.setOverrideCursor(Qt.WaitCursor)

        succeeded = 0
        skipped = 0
        failed = 0
        cancelled = False

        try:
            for job in pending:
                if job.output_path.exists():
                    decision = self._ask_existing_output(job.output_path)
                    if decision == QMessageBox.Cancel:
                        cancelled = True
                        break
                    if decision == QMessageBox.No:
                        job.status = "跳过"
                        job.result = "输出文件已存在"
                        skipped += 1
                        self.refresh_table()
                        QApplication.processEvents()
                        continue
                    overwrite = True
                else:
                    overwrite = False

                job.status = "处理中"
                job.result = ""
                self.refresh_table()
                QApplication.processEvents()

                try:
                    report = clean_docx_all(job.input_path, job.output_path, overwrite=overwrite)
                except (DocxCleanerError, OSError) as exc:
                    job.status = "失败"
                    job.result = str(exc)
                    failed += 1
                else:
                    job.status = "成功"
                    job.result = (
                        f"清理段落 {report.changed_paragraphs}，"
                        f"样式副本 {report.cloned_styles}，"
                        f"直接大纲 {report.direct_outline_removed}"
                    )
                    succeeded += 1

                self.refresh_table()
                QApplication.processEvents()
        finally:
            QApplication.restoreOverrideCursor()
            self._set_buttons_enabled(True)

        self.open_folder_button.setEnabled(any(job.status == "成功" for job in self.jobs))
        if cancelled:
            self.summary.setText(f"已取消。成功 {succeeded}，跳过 {skipped}，失败 {failed}。")
        else:
            self.summary.setText(f"处理完成。成功 {succeeded}，跳过 {skipped}，失败 {failed}。")

    def _ask_existing_output(self, output_path: Path) -> QMessageBox.StandardButton:
        return QMessageBox.question(
            self,
            "输出文件已存在",
            f"文件已存在：\n{output_path}\n\n是否覆盖？",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel,
            QMessageBox.No,
        )

    def clear_jobs(self) -> None:
        self.jobs.clear()
        self.refresh_table()

    def refresh_table(self) -> None:
        self.table.setRowCount(len(self.jobs))
        for row, job in enumerate(self.jobs):
            values = [job.input_path.name, job.status, str(job.output_path), job.result]
            for col, value in enumerate(values):
                item = QTableWidgetItem(value)
                item.setToolTip(value)
                if col == 1:
                    item.setTextAlignment(Qt.AlignCenter)
                self.table.setItem(row, col, item)

        total = len(self.jobs)
        if total == 0:
            self.summary.setText("等待添加文件")
        else:
            self.summary.setText(f"已添加 {total} 个文件")
        self.open_folder_button.setEnabled(any(job.status == "成功" for job in self.jobs))

    def open_selected_output_folder(self) -> None:
        row = self.table.currentRow()
        if row < 0 or row >= len(self.jobs):
            successful = next((job for job in self.jobs if job.status == "成功"), None)
            folder = successful.output_path.parent if successful else None
        else:
            folder = self.jobs[row].output_path.parent

        if folder is not None:
            QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder)))

    def _set_buttons_enabled(self, enabled: bool) -> None:
        self.add_button.setEnabled(enabled)
        self.run_button.setEnabled(enabled)
        self.clear_button.setEnabled(enabled)
        self.open_folder_button.setEnabled(enabled and any(job.status == "成功" for job in self.jobs))


def _is_supported_docx(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() == ".docx" and not path.name.startswith("~$")


def main() -> int:
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
