"""
URL Downloader
讀取 Excel，將 URL 欄位的檔案批次下載至
「Excel 所在目錄 / Excel 檔名（不含副檔名）」資料夾。
"""

import sys
import os
import threading
from pathlib import Path
from urllib.parse import urlparse, unquote

import pandas as pd
import requests

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QLineEdit, QFileDialog, QTableWidget,
    QTableWidgetItem, QProgressBar, QTextEdit, QSplitter,
    QHeaderView, QMessageBox, QAbstractItemView, QGroupBox,
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QObject
from PyQt5.QtGui import QColor, QFont


# ──────────────────────────────────────────────────────────────────────────────
# Worker
# ──────────────────────────────────────────────────────────────────────────────

class DownloadWorker(QObject):
    """在背景執行緒中逐一下載 URL 清單中的檔案。"""

    progress = pyqtSignal(int, int)          # (current_index, total)
    row_done = pyqtSignal(int, str, str)     # (row_index, status, filename)
    log_msg  = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, rows: list[dict], dest_dir: str):
        super().__init__()
        self.rows     = rows          # [{"row": int, "url": str, "filename": str}]
        self.dest_dir = dest_dir
        self._cancel  = False

    def cancel(self):
        self._cancel = True

    def run(self):
        total = len(self.rows)
        for idx, item in enumerate(self.rows):
            if self._cancel:
                self.log_msg.emit("⚠️  下載已取消。")
                break

            row_idx  = item["row"]
            url      = item["url"].strip()
            filename = item["filename"]

            self.log_msg.emit(f"[{idx+1}/{total}] 下載: {url}")
            try:
                resp = requests.get(url, stream=True, timeout=30)
                resp.raise_for_status()

                # 嘗試從 Content-Disposition 取得檔名
                cd = resp.headers.get("Content-Disposition", "")
                if "filename=" in cd:
                    fname = cd.split("filename=")[-1].strip().strip('"').strip("'")
                    filename = fname or filename

                dest_path = os.path.join(self.dest_dir, filename)
                # 避免同名覆蓋
                dest_path = _unique_path(dest_path)

                with open(dest_path, "wb") as f:
                    for chunk in resp.iter_content(chunk_size=65536):
                        if chunk:
                            f.write(chunk)

                self.row_done.emit(row_idx, "完成", os.path.basename(dest_path))
                self.log_msg.emit(f"  ✔ 儲存至: {dest_path}")
            except Exception as exc:
                self.row_done.emit(row_idx, "失敗", str(exc))
                self.log_msg.emit(f"  ✘ 錯誤: {exc}")

            self.progress.emit(idx + 1, total)

        self.finished.emit()


def _unique_path(path: str) -> str:
    """若路徑已存在，自動在檔名後加上 _1, _2, … 以免覆蓋。"""
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    counter = 1
    while True:
        new_path = f"{base}_{counter}{ext}"
        if not os.path.exists(new_path):
            return new_path
        counter += 1


def _filename_from_url(url: str) -> str:
    """從 URL 推測合理的檔名；若無法判斷則回傳 'download'。"""
    try:
        parsed = urlparse(url)
        name   = unquote(os.path.basename(parsed.path))
        return name if name else "download"
    except Exception:
        return "download"


# ──────────────────────────────────────────────────────────────────────────────
# Main Window
# ──────────────────────────────────────────────────────────────────────────────

class MainWindow(QMainWindow):

    STATUS_COLOR = {
        "等待中": "#888888",
        "完成":   "#27ae60",
        "失敗":   "#e74c3c",
    }

    def __init__(self):
        super().__init__()
        self.record_df: pd.DataFrame | None = None

        self._worker: DownloadWorker | None = None
        self._thread: QThread | None = None
        self._setup_ui()

    # ── UI construction ────────────────────────────────────────────────────────

    def _setup_ui(self):
        self.setWindowTitle("URL 批次下載工具")
        self.resize(960, 680)

        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setSpacing(8)
        root_layout.setContentsMargins(12, 12, 12, 12)

        # ── Step 1: 讀取 Excel ──────────────────────────────────────────────
        xlsx_group = QGroupBox("Step 1：選擇 URL_contactor.xlsx")
        xlsx_layout = QHBoxLayout(xlsx_group)

        self.xlsx_path_edit = QLineEdit()
        self.xlsx_path_edit.setPlaceholderText("URL_contactor.xlsx 的檔案路徑…")
        self.xlsx_path_edit.setReadOnly(True)

        btn_browse_xlsx = QPushButton("瀏覽…")
        btn_browse_xlsx.setFixedWidth(80)
        btn_browse_xlsx.clicked.connect(self._browse_xlsx)

        self.btn_load = QPushButton("載入")
        self.btn_load.setFixedWidth(70)
        self.btn_load.clicked.connect(self._load_xlsx)

        xlsx_layout.addWidget(QLabel("檔案:"))
        xlsx_layout.addWidget(self.xlsx_path_edit)
        xlsx_layout.addWidget(btn_browse_xlsx)
        xlsx_layout.addWidget(self.btn_load)
        root_layout.addWidget(xlsx_group)

        # ── Step 2: 目的目錄 ────────────────────────────────────────────────
        dest_group = QGroupBox("Step 2：選擇下載目的目錄")
        dest_layout = QHBoxLayout(dest_group)

        self.dest_path_edit = QLineEdit()
        self.dest_path_edit.setPlaceholderText("下載目的目錄…")

        btn_browse_dest = QPushButton("瀏覽…")
        btn_browse_dest.setFixedWidth(80)
        btn_browse_dest.clicked.connect(self._browse_dest)

        dest_layout.addWidget(QLabel("目錄:"))
        dest_layout.addWidget(self.dest_path_edit)
        dest_layout.addWidget(btn_browse_dest)
        root_layout.addWidget(dest_group)

        # ── Step 3: 表格 + log ──────────────────────────────────────────────
        splitter = QSplitter(Qt.Vertical)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["URL", "推測檔名", "狀態"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeToContents)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        splitter.addWidget(self.table)

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setFont(QFont("Courier New", 9))
        self.log_box.setMaximumHeight(160)
        splitter.addWidget(self.log_box)

        root_layout.addWidget(splitter, 1)

        # ── Progress bar ────────────────────────────────────────────────────
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        root_layout.addWidget(self.progress_bar)

        # ── Action buttons ──────────────────────────────────────────────────
        btn_row = QHBoxLayout()
        self.btn_start = QPushButton("開始下載")
        self.btn_start.setFixedHeight(36)
        self.btn_start.setEnabled(False)
        self.btn_start.clicked.connect(self._start_download)

        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.setFixedHeight(36)
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self._cancel_download)

        self.btn_clear_log = QPushButton("清除 Log")
        self.btn_clear_log.setFixedHeight(36)
        self.btn_clear_log.clicked.connect(self.log_box.clear)

        btn_row.addWidget(self.btn_start)
        btn_row.addWidget(self.btn_cancel)
        btn_row.addStretch()
        btn_row.addWidget(self.btn_clear_log)
        root_layout.addLayout(btn_row)

        self._log("程式已啟動，請選擇 Excel 檔案後按「載入」。")

    # ── Slot helpers ───────────────────────────────────────────────────────────

    def _log(self, msg: str):
        self.log_box.append(msg)
        self.log_box.verticalScrollBar().setValue(
            self.log_box.verticalScrollBar().maximum()
        )

    def _browse_xlsx(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "選擇 Excel 檔案", "", "Excel 檔案 (*.xlsx *.xls)"
        )
        if path:
            self.xlsx_path_edit.setText(path)

    def _browse_dest(self):
        path = QFileDialog.getExistingDirectory(self, "選擇下載目的目錄", "")
        if path:
            self.dest_path_edit.setText(path)

    def _load_xlsx(self):
        path = self.xlsx_path_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "提示", "請先選擇 Excel 檔案。")
            return
        if not os.path.isfile(path):
            QMessageBox.critical(self, "錯誤", f"找不到檔案:\n{path}")
            return
        try:
            record_df = pd.read_excel(path,header=1)
   
        except Exception as exc:
            QMessageBox.critical(self, "讀取失敗", str(exc))
            return

        if "URL" not in record_df.columns:
            QMessageBox.critical(
                self, "欄位缺失",
                "Excel 中找不到 'URL' 欄位。\n"
                f"目前欄位: {list(record_df.columns)}"
            )
            return

        self.record_df = record_df
        self._populate_table()
        self._log(f"✔ 已載入 {len(record_df)} 筆資料（來自 {os.path.basename(path)}）。")
        self.btn_start.setEnabled(True)

    def _populate_table(self):
        df = self.record_df
        self.table.setRowCount(0)
        for _, row in df.iterrows():
            url      = str(row.get("URL", "")).strip()
            filename = str(row.get("Filename", _filename_from_url(url)))
            if not filename or filename == "nan":
                filename = _filename_from_url(url)

            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r, 0, QTableWidgetItem(url))
            self.table.setItem(r, 1, QTableWidgetItem(filename))
            status_item = QTableWidgetItem("等待中")
            status_item.setForeground(QColor(self.STATUS_COLOR["等待中"]))
            self.table.setItem(r, 2, status_item)

    def _collect_rows(self) -> list[dict]:
        rows = []
        for r in range(self.table.rowCount()):
            rows.append({
                "row":      r,
                "url":      self.table.item(r, 0).text(),
                "filename": self.table.item(r, 1).text(),
            })
        return rows

    # ── Download control ───────────────────────────────────────────────────────

    def _start_download(self):
        dest = self.dest_path_edit.text().strip()
        if not dest:
            QMessageBox.warning(self, "提示", "請先選擇下載目的目錄。")
            return
        os.makedirs(dest, exist_ok=True)

        rows = self._collect_rows()
        if not rows:
            QMessageBox.information(self, "提示", "沒有可下載的項目。")
            return

        # 重置狀態
        for r in range(self.table.rowCount()):
            item = self.table.item(r, 2)
            item.setText("等待中")
            item.setForeground(QColor(self.STATUS_COLOR["等待中"]))

        self.progress_bar.setMaximum(len(rows))
        self.progress_bar.setValue(0)
        self.btn_start.setEnabled(False)
        self.btn_cancel.setEnabled(True)

        self._thread = QThread()
        self._worker = DownloadWorker(rows, dest)
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.progress.connect(self._on_progress)
        self._worker.row_done.connect(self._on_row_done)
        self._worker.log_msg.connect(self._log)
        self._worker.finished.connect(self._on_finished)
        self._worker.finished.connect(self._thread.quit)

        self._thread.start()

    def _cancel_download(self):
        if self._worker:
            self._worker.cancel()
        self.btn_cancel.setEnabled(False)

    # ── Worker signals ─────────────────────────────────────────────────────────

    def _on_progress(self, current: int, total: int):
        self.progress_bar.setValue(current)
        self.setWindowTitle(f"URL 批次下載工具 [{current}/{total}]")

    def _on_row_done(self, row_idx: int, status: str, detail: str):
        status_item = self.table.item(row_idx, 2)
        status_item.setText(status)
        status_item.setForeground(QColor(self.STATUS_COLOR.get(status, "#000000")))
        if status == "完成":
            self.table.item(row_idx, 1).setText(detail)

    def _on_finished(self):
        self.btn_start.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.setWindowTitle("URL 批次下載工具")
        done  = sum(
            1 for r in range(self.table.rowCount())
            if self.table.item(r, 2).text() == "完成"
        )
        fail  = sum(
            1 for r in range(self.table.rowCount())
            if self.table.item(r, 2).text() == "失敗"
        )
        self._log(f"\n── 下載完畢：成功 {done} 筆，失敗 {fail} 筆 ──")


# ──────────────────────────────────────────────────────────────────────────────
# Entry point
# ──────────────────────────────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
