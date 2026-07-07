"""
URL Downloader
讀取 Excel，將 URL 欄位的檔案批次下載至
「Excel 所在目錄 / Excel 檔名（不含副檔名）」資料夾。
SharePoint URL 自動使用 Microsoft 帳號認證下載。
"""

import sys
import os
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
# SharePoint helpers
# ──────────────────────────────────────────────────────────────────────────────

def _is_sharepoint(url: str) -> bool:
    return "sharepoint.com" in url.lower()


def _parse_sharepoint_url(url: str) -> tuple[str, str]:
    """
    從 SharePoint 完整 URL 拆出 (site_url, server_relative_url)。
    例：
      https://tenant.sharepoint.com/sites/MySite/Folder/file.docx
      → site_url = "https://tenant.sharepoint.com/sites/MySite"
        rel_url  = "/sites/MySite/Folder/file.docx"
    """
    parsed = urlparse(url)
    path   = unquote(parsed.path)          # 解碼 %20 等
    base   = f"{parsed.scheme}://{parsed.netloc}"

    for prefix in ("/sites/", "/teams/", "/personal/"):
        idx = path.find(prefix)
        if idx == -1:
            continue
        seg_start = idx + len(prefix)
        seg_end   = path.find("/", seg_start)
        if seg_end == -1:
            seg_end = len(path)
        site_url = base + path[:seg_end]
        return site_url, path

    return base, path


def _download_sharepoint_file(url: str, dest_path: str,
                               username: str, password: str) -> None:
    """用 office365-rest-python-client 認證後下載 SharePoint 檔案。"""
    from office365.sharepoint.client_context import ClientContext
    from office365.runtime.auth.user_credential import UserCredential

    site_url, rel_url = _parse_sharepoint_url(url)
    cred = UserCredential(username, password)
    ctx  = ClientContext(site_url).with_credentials(cred)

    with open(dest_path, "wb") as f:
        ctx.web.get_file_by_server_relative_url(rel_url).download(f).execute_query()


# ──────────────────────────────────────────────────────────────────────────────
# Worker
# ──────────────────────────────────────────────────────────────────────────────

class DownloadWorker(QObject):
    """在背景執行緒中逐一下載 URL 清單中的檔案。"""

    progress = pyqtSignal(int, int)          # (current_index, total)
    row_done = pyqtSignal(int, str, str)     # (row_index, status, detail)
    log_msg  = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, rows: list[dict], dest_dir: str,
                 username: str = "", password: str = ""):
        super().__init__()
        self.rows     = rows
        self.dest_dir = dest_dir
        self.username = username
        self.password = password
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
                dest_path = _unique_path(os.path.join(self.dest_dir, filename))

                if _is_sharepoint(url):
                    if not self.username or not self.password:
                        raise ValueError(
                            "SharePoint URL 需要帳號密碼，請填寫 Step 2 的認證欄位。"
                        )
                    _download_sharepoint_file(
                        url, dest_path, self.username, self.password
                    )
                else:
                    self._download_http(url, dest_path)

                final_name = os.path.basename(dest_path)
                self.row_done.emit(row_idx, "完成", final_name)
                self.log_msg.emit(f"  ✔ 儲存至: {dest_path}")

            except Exception as exc:
                self.row_done.emit(row_idx, "失敗", str(exc))
                self.log_msg.emit(f"  ✘ 錯誤: {exc}")

            self.progress.emit(idx + 1, total)

        self.finished.emit()

    def _download_http(self, url: str, dest_path: str) -> None:
        resp = requests.get(url, stream=True, timeout=30)
        resp.raise_for_status()

        # Content-Disposition 可能含更精確的檔名
        cd = resp.headers.get("Content-Disposition", "")
        if "filename=" in cd:
            fname = cd.split("filename=")[-1].strip().strip('"').strip("'")
            if fname:
                dest_path = _unique_path(
                    os.path.join(os.path.dirname(dest_path), fname)
                )

        with open(dest_path, "wb") as f:
            for chunk in resp.iter_content(chunk_size=65536):
                if chunk:
                    f.write(chunk)


# ──────────────────────────────────────────────────────────────────────────────
# Utility
# ──────────────────────────────────────────────────────────────────────────────

def _unique_path(path: str) -> str:
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
        self._dest_dir: str = ""
        self._worker: DownloadWorker | None = None
        self._thread: QThread | None = None
        self._setup_ui()

    # ── UI ─────────────────────────────────────────────────────────────────────

    def _setup_ui(self):
        self.setWindowTitle("URL 批次下載工具")
        self.resize(960, 720)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(8)
        root.setContentsMargins(12, 12, 12, 12)

        # ── Step 1: 讀取 Excel ─────────────────────────────────────────────
        g1 = QGroupBox("Step 1：選擇 Excel 檔案")
        l1 = QHBoxLayout(g1)

        self.xlsx_path_edit = QLineEdit()
        self.xlsx_path_edit.setPlaceholderText("Excel 檔案路徑…")
        self.xlsx_path_edit.setReadOnly(True)

        btn_browse = QPushButton("瀏覽…")
        btn_browse.setFixedWidth(80)
        btn_browse.clicked.connect(self._browse_xlsx)

        self.btn_load = QPushButton("載入")
        self.btn_load.setFixedWidth(70)
        self.btn_load.clicked.connect(self._load_xlsx)

        l1.addWidget(QLabel("檔案:"))
        l1.addWidget(self.xlsx_path_edit)
        l1.addWidget(btn_browse)
        l1.addWidget(self.btn_load)
        root.addWidget(g1)

        # ── Step 2: SharePoint 認證 ────────────────────────────────────────
        g2 = QGroupBox("Step 2：SharePoint / Microsoft 帳號認證（一般 HTTP URL 可留空）")
        l2 = QHBoxLayout(g2)

        self.user_edit = QLineEdit()
        self.user_edit.setPlaceholderText("帳號（e-mail）")

        self.pass_edit = QLineEdit()
        self.pass_edit.setPlaceholderText("密碼")
        self.pass_edit.setEchoMode(QLineEdit.Password)

        l2.addWidget(QLabel("帳號:"))
        l2.addWidget(self.user_edit, 3)
        l2.addSpacing(16)
        l2.addWidget(QLabel("密碼:"))
        l2.addWidget(self.pass_edit, 2)
        root.addWidget(g2)

        # ── Step 3: 下載目錄（唯讀，自動產生）────────────────────────────
        g3 = QGroupBox("Step 3：下載目的目錄（自動產生）")
        l3 = QHBoxLayout(g3)

        self.dest_display = QLineEdit()
        self.dest_display.setReadOnly(True)
        self.dest_display.setPlaceholderText("載入 Excel 後自動設定…")
        self.dest_display.setStyleSheet("color: #555; background: #f5f5f5;")

        l3.addWidget(QLabel("目錄:"))
        l3.addWidget(self.dest_display)
        root.addWidget(g3)

        # ── 表格 + Log ─────────────────────────────────────────────────────
        splitter = QSplitter(Qt.Vertical)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["URL", "檔名", "狀態"])
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

        root.addWidget(splitter, 1)

        # ── Progress bar ───────────────────────────────────────────────────
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        root.addWidget(self.progress_bar)

        # ── Buttons ────────────────────────────────────────────────────────
        btn_row = QHBoxLayout()

        self.btn_start = QPushButton("開始下載")
        self.btn_start.setFixedHeight(36)
        self.btn_start.setEnabled(False)
        self.btn_start.clicked.connect(self._start_download)

        self.btn_cancel = QPushButton("取消")
        self.btn_cancel.setFixedHeight(36)
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self._cancel_download)

        btn_clear = QPushButton("清除 Log")
        btn_clear.setFixedHeight(36)
        btn_clear.clicked.connect(self.log_box.clear)

        btn_row.addWidget(self.btn_start)
        btn_row.addWidget(self.btn_cancel)
        btn_row.addStretch()
        btn_row.addWidget(btn_clear)
        root.addLayout(btn_row)

        self._log("程式已啟動。")
        self._log("• SharePoint URL → 請先填入 Step 2 帳號密碼再下載。")
        self._log("• 一般 HTTP(S) URL → Step 2 可留空直接下載。")

    # ── Helpers ────────────────────────────────────────────────────────────────

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
                f"Excel 中找不到 'URL' 欄位。\n目前欄位: {list(record_df.columns)}"
            )
            return

        xlsx_stem = Path(path).stem
        xlsx_dir  = Path(path).parent
        dest_dir  = str(xlsx_dir / xlsx_stem)

        self.record_df = record_df
        self._dest_dir = dest_dir
        self.dest_display.setText(dest_dir)
        self._populate_table()
        self._log(f"✔ 已載入 {len(record_df)} 筆（來自 {os.path.basename(path)}）。")
        self._log(f"   下載目錄: {dest_dir}")
        self.btn_start.setEnabled(True)

    def _populate_table(self):
        self.table.setRowCount(0)
        for _, row in self.record_df.iterrows():
            url      = str(row.get("URL", "")).strip()
            filename = str(row.get("Filename", ""))
            if not filename or filename == "nan":
                filename = _filename_from_url(url)

            r = self.table.rowCount()
            self.table.insertRow(r)
            self.table.setItem(r, 0, QTableWidgetItem(url))
            self.table.setItem(r, 1, QTableWidgetItem(filename))
            si = QTableWidgetItem("等待中")
            si.setForeground(QColor(self.STATUS_COLOR["等待中"]))
            self.table.setItem(r, 2, si)

    def _collect_rows(self) -> list[dict]:
        return [
            {
                "row":      r,
                "url":      self.table.item(r, 0).text(),
                "filename": self.table.item(r, 1).text(),
            }
            for r in range(self.table.rowCount())
        ]

    # ── Download control ───────────────────────────────────────────────────────

    def _start_download(self):
        if not self._dest_dir:
            QMessageBox.warning(self, "提示", "請先載入 Excel 檔案。")
            return

        rows = self._collect_rows()
        if not rows:
            QMessageBox.information(self, "提示", "沒有可下載的項目。")
            return

        # 若有 SharePoint URL 但未填帳號，警告後讓使用者決定是否繼續
        has_sp = any(_is_sharepoint(r["url"]) for r in rows)
        if has_sp and not self.user_edit.text().strip():
            ret = QMessageBox.question(
                self, "未填認證資訊",
                "清單中含有 SharePoint URL，但尚未填寫帳號密碼。\n"
                "SharePoint 檔案將會下載失敗。\n\n確定繼續？",
                QMessageBox.Yes | QMessageBox.No,
            )
            if ret == QMessageBox.No:
                return

        os.makedirs(self._dest_dir, exist_ok=True)

        for r in range(self.table.rowCount()):
            item = self.table.item(r, 2)
            item.setText("等待中")
            item.setForeground(QColor(self.STATUS_COLOR["等待中"]))

        self.progress_bar.setMaximum(len(rows))
        self.progress_bar.setValue(0)
        self.btn_start.setEnabled(False)
        self.btn_cancel.setEnabled(True)

        self._thread = QThread()
        self._worker = DownloadWorker(
            rows, self._dest_dir,
            username=self.user_edit.text().strip(),
            password=self.pass_edit.text(),
        )
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
        si = self.table.item(row_idx, 2)
        si.setText(status)
        si.setForeground(QColor(self.STATUS_COLOR.get(status, "#000000")))
        if status == "完成":
            self.table.item(row_idx, 1).setText(detail)

    def _on_finished(self):
        self.btn_start.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.setWindowTitle("URL 批次下載工具")
        done = sum(1 for r in range(self.table.rowCount())
                   if self.table.item(r, 2).text() == "完成")
        fail = sum(1 for r in range(self.table.rowCount())
                   if self.table.item(r, 2).text() == "失敗")
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
