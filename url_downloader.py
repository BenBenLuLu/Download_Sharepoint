"""
URL Batch Downloader
Reads an Excel file, downloads files from the URL column to
a folder named after the Excel file in the same directory.
SharePoint URLs are downloaded using Microsoft account credentials.
"""

import sys
import os
import subprocess
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
# Check office365 package availability
# ──────────────────────────────────────────────────────────────────────────────

try:
    from office365.sharepoint.client_context import ClientContext       # noqa: F401
    from office365.runtime.auth.user_credential import UserCredential   # noqa: F401
    _OFFICE365_AVAILABLE = True
except ModuleNotFoundError:
    _OFFICE365_AVAILABLE = False


def _ensure_office365() -> bool:
    """Try to install office365 via pip if missing; returns True on success."""
    global _OFFICE365_AVAILABLE
    if _OFFICE365_AVAILABLE:
        return True
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "Office365-REST-Python-Client"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        from office365.sharepoint.client_context import ClientContext       # noqa: F401
        from office365.runtime.auth.user_credential import UserCredential   # noqa: F401
        _OFFICE365_AVAILABLE = True
        return True
    except Exception:
        return False


# ──────────────────────────────────────────────────────────────────────────────
# SharePoint helpers
# ──────────────────────────────────────────────────────────────────────────────

def _is_sharepoint(url: str) -> bool:
    return "sharepoint.com" in url.lower()


def _parse_sharepoint_url(url: str) -> tuple[str, str]:
    """
    Split a full SharePoint URL into (site_url, server_relative_url).
    Example:
      https://tenant.sharepoint.com/sites/MySite/Folder/file.docx
      -> site_url = "https://tenant.sharepoint.com/sites/MySite"
         rel_url  = "/sites/MySite/Folder/file.docx"
    """
    parsed = urlparse(url)
    path   = unquote(parsed.path)
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
    """Authenticate with office365-rest-python-client and download a SharePoint file."""
    if not _ensure_office365():
        raise RuntimeError(
            "Missing required package. Please run:\n"
            "  pip install Office365-REST-Python-Client\n"
            "Then restart the application."
        )

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
    """Downloads files from the URL list in a background thread."""

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
                self.log_msg.emit("⚠️  Download cancelled.")
                break

            row_idx  = item["row"]
            url      = item["url"].strip()
            filename = item["filename"]

            self.log_msg.emit(f"[{idx+1}/{total}] Downloading: {url}")
            try:
                dest_path = _unique_path(os.path.join(self.dest_dir, filename))

                if _is_sharepoint(url):
                    if not self.username or not self.password:
                        raise ValueError(
                            "SharePoint URL requires credentials. "
                            "Please fill in Step 2 (username and password)."
                        )
                    _download_sharepoint_file(
                        url, dest_path, self.username, self.password
                    )
                else:
                    self._download_http(url, dest_path)

                final_name = os.path.basename(dest_path)
                self.row_done.emit(row_idx, "Done", final_name)
                self.log_msg.emit(f"  ✔ Saved to: {dest_path}")

            except Exception as exc:
                self.row_done.emit(row_idx, "Failed", str(exc))
                self.log_msg.emit(f"  ✘ Error: {exc}")

            self.progress.emit(idx + 1, total)

        self.finished.emit()

    def _download_http(self, url: str, dest_path: str) -> None:
        resp = requests.get(url, stream=True, timeout=30)
        resp.raise_for_status()

        # Use filename from Content-Disposition header if available
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
        "Pending": "#888888",
        "Done":    "#27ae60",
        "Failed":  "#e74c3c",
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
        self.setWindowTitle("URL Batch Downloader")
        self.resize(960, 720)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(8)
        root.setContentsMargins(12, 12, 12, 12)

        # ── Step 1: Load Excel ─────────────────────────────────────────────
        g1 = QGroupBox("Step 1: Select Excel File")
        l1 = QHBoxLayout(g1)

        self.xlsx_path_edit = QLineEdit()
        self.xlsx_path_edit.setPlaceholderText("Excel file path...")
        self.xlsx_path_edit.setReadOnly(True)

        btn_browse = QPushButton("Browse...")
        btn_browse.setFixedWidth(90)
        btn_browse.clicked.connect(self._browse_xlsx)

        self.btn_load = QPushButton("Load")
        self.btn_load.setFixedWidth(70)
        self.btn_load.clicked.connect(self._load_xlsx)

        l1.addWidget(QLabel("File:"))
        l1.addWidget(self.xlsx_path_edit)
        l1.addWidget(btn_browse)
        l1.addWidget(self.btn_load)
        root.addWidget(g1)

        # ── Step 2: SharePoint credentials ────────────────────────────────
        g2 = QGroupBox(
            "Step 2: SharePoint / Microsoft Account Credentials  "
            "(leave blank for regular HTTP URLs)"
        )
        l2 = QHBoxLayout(g2)

        self.user_edit = QLineEdit()
        self.user_edit.setPlaceholderText("Username (e-mail)")

        self.pass_edit = QLineEdit()
        self.pass_edit.setPlaceholderText("Password")
        self.pass_edit.setEchoMode(QLineEdit.Password)

        l2.addWidget(QLabel("Username:"))
        l2.addWidget(self.user_edit, 3)
        l2.addSpacing(16)
        l2.addWidget(QLabel("Password:"))
        l2.addWidget(self.pass_edit, 2)
        root.addWidget(g2)

        # ── Step 3: Output directory (read-only, auto-generated) ───────────
        g3 = QGroupBox("Step 3: Download Directory (auto-generated)")
        l3 = QHBoxLayout(g3)

        self.dest_display = QLineEdit()
        self.dest_display.setReadOnly(True)
        self.dest_display.setPlaceholderText("Auto-set after loading Excel...")
        self.dest_display.setStyleSheet("color: #555; background: #f5f5f5;")

        l3.addWidget(QLabel("Directory:"))
        l3.addWidget(self.dest_display)
        root.addWidget(g3)

        # ── Table + Log ────────────────────────────────────────────────────
        splitter = QSplitter(Qt.Vertical)

        self.table = QTableWidget(0, 3)
        self.table.setHorizontalHeaderLabels(["URL", "Filename", "Status"])
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

        self.btn_start = QPushButton("Start Download")
        self.btn_start.setFixedHeight(36)
        self.btn_start.setEnabled(False)
        self.btn_start.clicked.connect(self._start_download)

        self.btn_cancel = QPushButton("Cancel")
        self.btn_cancel.setFixedHeight(36)
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self._cancel_download)

        btn_clear = QPushButton("Clear Log")
        btn_clear.setFixedHeight(36)
        btn_clear.clicked.connect(self.log_box.clear)

        btn_row.addWidget(self.btn_start)
        btn_row.addWidget(self.btn_cancel)
        btn_row.addStretch()
        btn_row.addWidget(btn_clear)
        root.addLayout(btn_row)

        self._log("Application started.")
        self._log("• SharePoint URLs  → fill in Step 2 credentials before downloading.")
        self._log("• Regular HTTP(S) URLs → Step 2 can be left blank.")
        if not _OFFICE365_AVAILABLE:
            self._log("")
            self._log("⚠️  WARNING: office365 package not found. SharePoint download is unavailable.")
            self._log("   Please run the following command and restart the application:")
            self._log(f"   {sys.executable} -m pip install Office365-REST-Python-Client")

    # ── Helpers ────────────────────────────────────────────────────────────────

    def _log(self, msg: str):
        self.log_box.append(msg)
        self.log_box.verticalScrollBar().setValue(
            self.log_box.verticalScrollBar().maximum()
        )

    def _browse_xlsx(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if path:
            self.xlsx_path_edit.setText(path)

    def _load_xlsx(self):
        path = self.xlsx_path_edit.text().strip()
        if not path:
            QMessageBox.warning(self, "Notice", "Please select an Excel file first.")
            return
        if not os.path.isfile(path):
            QMessageBox.critical(self, "Error", f"File not found:\n{path}")
            return
        try:
            record_df = pd.read_excel(path)
        except Exception as exc:
            QMessageBox.critical(self, "Load Failed", str(exc))
            return

        if "URL" not in record_df.columns:
            QMessageBox.critical(
                self, "Missing Column",
                f"'URL' column not found in the Excel file.\n"
                f"Available columns: {list(record_df.columns)}"
            )
            return

        xlsx_stem = Path(path).stem
        xlsx_dir  = Path(path).parent
        dest_dir  = str(xlsx_dir / xlsx_stem)

        self.record_df = record_df
        self._dest_dir = dest_dir
        self.dest_display.setText(dest_dir)
        self._populate_table()
        self._log(f"✔ Loaded {len(record_df)} rows from {os.path.basename(path)}.")
        self._log(f"   Download directory: {dest_dir}")
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
            si = QTableWidgetItem("Pending")
            si.setForeground(QColor(self.STATUS_COLOR["Pending"]))
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
            QMessageBox.warning(self, "Notice", "Please load an Excel file first.")
            return

        rows = self._collect_rows()
        if not rows:
            QMessageBox.information(self, "Notice", "No items to download.")
            return

        has_sp = any(_is_sharepoint(r["url"]) for r in rows)
        if has_sp and not self.user_edit.text().strip():
            ret = QMessageBox.question(
                self, "Credentials Missing",
                "The list contains SharePoint URLs but no credentials were entered.\n"
                "SharePoint files will fail to download.\n\nContinue anyway?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if ret == QMessageBox.No:
                return

        os.makedirs(self._dest_dir, exist_ok=True)

        for r in range(self.table.rowCount()):
            item = self.table.item(r, 2)
            item.setText("Pending")
            item.setForeground(QColor(self.STATUS_COLOR["Pending"]))

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
        self.setWindowTitle(f"URL Batch Downloader [{current}/{total}]")

    def _on_row_done(self, row_idx: int, status: str, detail: str):
        si = self.table.item(row_idx, 2)
        si.setText(status)
        si.setForeground(QColor(self.STATUS_COLOR.get(status, "#000000")))
        if status == "Done":
            self.table.item(row_idx, 1).setText(detail)

    def _on_finished(self):
        self.btn_start.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.setWindowTitle("URL Batch Downloader")
        done = sum(1 for r in range(self.table.rowCount())
                   if self.table.item(r, 2).text() == "Done")
        fail = sum(1 for r in range(self.table.rowCount())
                   if self.table.item(r, 2).text() == "Failed")
        self._log(f"\n── Finished: {done} succeeded, {fail} failed ──")


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
