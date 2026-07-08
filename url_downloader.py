"""
URL Batch Downloader
Reads an Excel file, downloads files from the URL column to
a folder named after the Excel file in the same directory.
SharePoint URLs use Microsoft Device Code sign-in (supports MFA / Conditional Access).
"""

import sys
import os
import base64
import subprocess
import tempfile
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


# Public client IDs to try (Azure CLI is most widely consented in enterprises)
_MS_CLIENT_IDS = [
    "04b07795-8ddb-461a-bbee-04f9e1bf07b9",   # Azure CLI
    "14d82eec-204b-4c2f-b7e0-44b84b7b7ba1",   # Microsoft Graph PowerShell
    "1950a258-227b-4e31-a9cf-717495945fc2",   # Azure PowerShell
]
_GRAPH_SCOPES = ["https://graph.microsoft.com/.default"]


# ──────────────────────────────────────────────────────────────────────────────
# Package availability
# ──────────────────────────────────────────────────────────────────────────────

try:
    import msal                                                          # noqa: F401
    from office365.sharepoint.client_context import ClientContext       # noqa: F401
    from office365.runtime.auth.token_response import TokenResponse     # noqa: F401
    _OFFICE365_AVAILABLE = True
except ModuleNotFoundError:
    _OFFICE365_AVAILABLE = False


def _ensure_office365() -> bool:
    global _OFFICE365_AVAILABLE
    if _OFFICE365_AVAILABLE:
        return True
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install",
             "Office365-REST-Python-Client", "msal"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        import msal                                                          # noqa: F401
        from office365.sharepoint.client_context import ClientContext       # noqa: F401
        from office365.runtime.auth.token_response import TokenResponse     # noqa: F401
        _OFFICE365_AVAILABLE = True
        return True
    except Exception:
        return False


# ──────────────────────────────────────────────────────────────────────────────
# SharePoint auth – Interactive browser login + Microsoft Graph download
# ──────────────────────────────────────────────────────────────────────────────

class SharePointAuth:
    """MSAL session for SharePoint / Graph file downloads."""

    def __init__(self):
        self._app         = None
        self._client_id   = _MS_CLIENT_IDS[0]
        self._account     = None
        self._username    = ""
        self._tenant      = "organizations"
        self._scopes: list[str] = _GRAPH_SCOPES
        self._use_graph   = True

    @property
    def is_signed_in(self) -> bool:
        return self._account is not None

    @property
    def username(self) -> str:
        return self._username

    @property
    def use_graph(self) -> bool:
        return self._use_graph

    def _get_app(self, client_id: str) -> "msal.PublicClientApplication":
        import msal
        if self._app is None or client_id != self._client_id:
            self._client_id = client_id
            authority = f"https://login.microsoftonline.com/{self._tenant}"
            self._app = msal.PublicClientApplication(client_id, authority=authority)
        return self._app

    def _apply_result(self, result: dict, scopes: list[str], use_graph: bool) -> None:
        self._scopes    = scopes
        self._use_graph = use_graph
        accounts = self._get_app(self._client_id).get_accounts()
        self._account  = accounts[0] if accounts else None
        self._username = (
            self._account.get("username", "") if self._account
            else result.get("id_token_claims", {}).get("preferred_username", "")
        )

    def sign_in_interactive(self, sharepoint_host: str,
                            client_ids: list[str]) -> None:
        """
        Open the system browser for Microsoft login (supports MFA / Conditional Access).
        Tries Microsoft Graph scope first, then SharePoint scope, across several client IDs.
        """
        if not _ensure_office365():
            raise RuntimeError(
                "Missing packages. Run:\n"
                "  pip install Office365-REST-Python-Client msal"
            )

        host      = sharepoint_host.lower().strip()
        sp_scopes = [f"https://{host}/.default"]
        attempts  = [
            (_GRAPH_SCOPES, True),
            (sp_scopes,     False),
        ]
        errors: list[str] = []

        for cid in client_ids:
            app = self._get_app(cid)
            for scopes, use_graph in attempts:
                # Silent re-use
                for acct in app.get_accounts() or [None]:
                    try:
                        kwargs = {"scopes": scopes}
                        if acct:
                            kwargs["account"] = acct
                        result = app.acquire_token_silent(**kwargs)
                        if result and "access_token" in result:
                            self._apply_result(result, scopes, use_graph)
                            return
                    except Exception:
                        pass

                # Interactive browser
                try:
                    result = app.acquire_token_interactive(
                        scopes=scopes,
                        prompt="select_account",
                    )
                    if result and "access_token" in result:
                        self._apply_result(result, scopes, use_graph)
                        return
                    err = result.get("error_description", str(result))
                    errors.append(f"[{cid[:8]}…] {err}")
                except Exception as exc:
                    errors.append(f"[{cid[:8]}…] {exc}")

        hint = (
            "\n\nIf you see AADSTS700016, the Client ID is not registered in shl-group.com.\n"
            "Ask IT to follow AZURE_APP_SETUP.md and grant admin consent."
        )
        raise RuntimeError(
            "Sign-in failed.\n\n" + "\n".join(errors[-3:]) + hint
        )

    def get_bearer_token(self) -> str:
        if not self.is_signed_in:
            raise RuntimeError("Not signed in.")
        app    = self._get_app(self._client_id)
        result = app.acquire_token_silent(self._scopes, account=self._account)
        if not result or "access_token" not in result:
            raise RuntimeError("Session expired. Please sign in again.")
        return result["access_token"]

    def acquire_token(self, sharepoint_url: str):
        """TokenResponse wrapper for office365-rest-python-client."""
        from office365.runtime.auth.token_response import TokenResponse
        host   = urlparse(sharepoint_url).netloc
        scopes = [f"https://{host}/.default"]
        app    = self._get_app(self._client_id)
        result = app.acquire_token_silent(scopes, account=self._account)
        if not result or "access_token" not in result:
            result = app.acquire_token_silent(self._scopes, account=self._account)
        if not result or "access_token" not in result:
            raise RuntimeError("Session expired. Please sign in again.")
        return TokenResponse.from_json(result)


# ──────────────────────────────────────────────────────────────────────────────
# SharePoint URL helpers
# ──────────────────────────────────────────────────────────────────────────────

def _is_sharepoint(url: str) -> bool:
    return "sharepoint.com" in url.lower()


def _parse_sharepoint_url(url: str) -> tuple[str, str]:
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
        return base + path[:seg_end], path

    return base, path


def _tenant_hint_from_url(url: str) -> str:
    """shlgroup.sharepoint.com → shlgroup.onmicrosoft.com"""
    host = urlparse(url).netloc.lower()
    if ".sharepoint.com" in host:
        return host.split(".sharepoint.com")[0] + ".onmicrosoft.com"
    return "organizations"


def _sharepoint_host_from_tenant(tenant: str) -> str | None:
    """shlgroup.onmicrosoft.com → shlgroup.sharepoint.com"""
    t = tenant.strip().lower()
    if not t or t in ("organizations", "common", "consumers"):
        return None
    if t.endswith(".onmicrosoft.com"):
        return t.replace(".onmicrosoft.com", ".sharepoint.com")
    if ".sharepoint.com" in t:
        return t.split("/")[0]
    return f"{t}.sharepoint.com"


def _sharepoint_host_for_signin(tenant: str, sample_url: str | None) -> str:
    if sample_url and _is_sharepoint(sample_url):
        return urlparse(sample_url).netloc.lower()
    host = _sharepoint_host_from_tenant(tenant)
    if host:
        return host
    raise ValueError(
        "Cannot determine SharePoint host.\n"
        "Please load your Excel file first (Step 1 → Load),\n"
        "or set Tenant to e.g. shlgroup.onmicrosoft.com"
    )


def _encode_share_url(url: str) -> str:
    """Encode a SharePoint URL for the Graph /shares/ API."""
    encoded = base64.urlsafe_b64encode(url.encode("utf-8")).decode("utf-8").rstrip("=")
    return f"u!{encoded}"


def _download_via_graph(url: str, dest_path: str, token: str) -> None:
    share_id = _encode_share_url(url)
    api_url  = f"https://graph.microsoft.com/v1.0/shares/{share_id}/driveItem/content"
    resp = requests.get(
        api_url,
        headers={"Authorization": f"Bearer {token}"},
        stream=True, timeout=120,
    )
    resp.raise_for_status()
    fd, tmp = tempfile.mkstemp(suffix=".tmp")
    os.close(fd)
    try:
        with open(tmp, "wb") as f:
            for chunk in resp.iter_content(chunk_size=65536):
                if chunk:
                    f.write(chunk)
        _validate_download(tmp)
        os.replace(tmp, dest_path)
        tmp = None
    finally:
        if tmp and os.path.exists(tmp):
            os.remove(tmp)


def _download_sharepoint_file(url: str, dest_path: str, auth: SharePointAuth) -> None:
    if auth.use_graph:
        _download_via_graph(url, dest_path, auth.get_bearer_token())
        return

    from office365.sharepoint.client_context import ClientContext

    site_url, rel_url = _parse_sharepoint_url(url)
    ctx = ClientContext(site_url).with_access_token(
        lambda: auth.acquire_token(url)
    )
    fd, tmp = tempfile.mkstemp(suffix=".tmp")
    os.close(fd)
    try:
        with open(tmp, "wb") as f:
            ctx.web.get_file_by_server_relative_url(rel_url).download(f).execute_query()
        _validate_download(tmp)
        os.replace(tmp, dest_path)
        tmp = None
    finally:
        if tmp and os.path.exists(tmp):
            os.remove(tmp)


def _validate_download(path: str) -> None:
    size = os.path.getsize(path)
    if size == 0:
        raise RuntimeError("Downloaded file is empty (0 bytes). Authentication may have failed.")

    with open(path, "rb") as f:
        head = f.read(512).lstrip()

    # Reject HTML login/error pages saved as files
    if head.startswith(b"<!") or head.startswith(b"<html") or head.startswith(b"<HTML"):
        raise RuntimeError(
            "Downloaded content is an HTML page, not the actual file. "
            "Please sign in again with 'Sign in with Microsoft'."
        )


# ──────────────────────────────────────────────────────────────────────────────
# Workers
# ──────────────────────────────────────────────────────────────────────────────

class DownloadWorker(QObject):
    progress = pyqtSignal(int, int)
    row_done = pyqtSignal(int, str, str)
    log_msg  = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, rows: list[dict], dest_dir: str, auth: SharePointAuth):
        super().__init__()
        self.rows     = rows
        self.dest_dir = dest_dir
        self.auth     = auth
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

            if not url or url.lower() in ("nan", "none"):
                self.row_done.emit(row_idx, "Failed", "Empty URL")
                self.progress.emit(idx + 1, total)
                continue

            self.log_msg.emit(f"[{idx+1}/{total}] Downloading: {url}")
            dest_path = _unique_path(os.path.join(self.dest_dir, filename))

            try:
                if _is_sharepoint(url):
                    if not self.auth.is_signed_in:
                        raise RuntimeError(
                            "SharePoint URL requires sign-in. "
                            "Click 'Sign in with Microsoft' in Step 2."
                        )
                    _download_sharepoint_file(url, dest_path, self.auth)
                else:
                    self._download_http(url, dest_path)

                final_name = os.path.basename(dest_path)
                self.row_done.emit(row_idx, "Done", final_name)
                self.log_msg.emit(f"  ✔ Saved to: {dest_path} ({os.path.getsize(dest_path):,} bytes)")

            except Exception as exc:
                if os.path.exists(dest_path) and os.path.getsize(dest_path) == 0:
                    os.remove(dest_path)
                self.row_done.emit(row_idx, "Failed", str(exc))
                self.log_msg.emit(f"  ✘ Error: {exc}")

            self.progress.emit(idx + 1, total)

        self.finished.emit()

    def _download_http(self, url: str, dest_path: str) -> None:
        fd, tmp = tempfile.mkstemp(suffix=".tmp")
        os.close(fd)
        try:
            resp = requests.get(url, stream=True, timeout=60)
            resp.raise_for_status()

            cd = resp.headers.get("Content-Disposition", "")
            if "filename=" in cd:
                fname = cd.split("filename=")[-1].strip().strip('"').strip("'")
                if fname:
                    dest_path = _unique_path(
                        os.path.join(os.path.dirname(dest_path), fname)
                    )

            with open(tmp, "wb") as f:
                for chunk in resp.iter_content(chunk_size=65536):
                    if chunk:
                        f.write(chunk)

            _validate_download(tmp)
            os.replace(tmp, dest_path)
            tmp = None
        finally:
            if tmp and os.path.exists(tmp):
                os.remove(tmp)


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
        self._auth = SharePointAuth()
        self._worker: DownloadWorker | None = None
        self._thread: QThread | None = None
        self._setup_ui()

    def _setup_ui(self):
        self.setWindowTitle("URL Batch Downloader")
        self.resize(960, 740)

        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setSpacing(8)
        root.setContentsMargins(12, 12, 12, 12)

        # ── Step 1 ─────────────────────────────────────────────────────────
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

        # ── Step 2: SharePoint sign-in ─────────────────────────────────────
        g2 = QGroupBox("Step 2: SharePoint Sign-In  (required for SharePoint URLs)")
        l2v = QVBoxLayout(g2)

        row_tenant = QHBoxLayout()
        self.tenant_edit = QLineEdit("organizations")
        self.tenant_edit.setPlaceholderText("e.g. shlgroup.onmicrosoft.com")
        row_tenant.addWidget(QLabel("Tenant:"))
        row_tenant.addWidget(self.tenant_edit)
        l2v.addLayout(row_tenant)

        row_client = QHBoxLayout()
        self.client_id_edit = QLineEdit()
        self.client_id_edit.setPlaceholderText(
            "Required – App (client) ID from IT  (see IT Setup Guide)"
        )
        btn_it_help = QPushButton("IT Setup Guide")
        btn_it_help.setFixedWidth(110)
        btn_it_help.clicked.connect(self._show_it_guide)
        row_client.addWidget(QLabel("Client ID:"))
        row_client.addWidget(self.client_id_edit, 1)
        row_client.addWidget(btn_it_help)
        l2v.addLayout(row_client)

        note = QLabel(
            "<span style='color:#c0392b'>"
            "shl-group.com blocks public Microsoft apps. "
            "Client ID from IT is <b>required</b>."
            "</span>"
        )
        note.setWordWrap(True)
        l2v.addWidget(note)

        row_signin = QHBoxLayout()
        self.btn_signin = QPushButton("Sign in with Microsoft")
        self.btn_signin.setFixedHeight(32)
        self.btn_signin.clicked.connect(self._sign_in)
        self.signin_status = QLabel("Not signed in")
        self.signin_status.setStyleSheet("color: #888;")
        row_signin.addWidget(self.btn_signin)
        row_signin.addWidget(self.signin_status, 1)
        l2v.addLayout(row_signin)
        root.addWidget(g2)

        # ── Step 3 ─────────────────────────────────────────────────────────
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
        self.log_box.setMaximumHeight(180)
        splitter.addWidget(self.log_box)
        root.addWidget(splitter, 1)

        self.progress_bar = QProgressBar()
        root.addWidget(self.progress_bar)

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
        self._log("Step 1: Load Excel  →  Step 2: Sign in  →  Step 3: Start Download")
        self._log("• shl-group.com requires a Client ID from IT – click 'IT Setup Guide'.")
        if not _OFFICE365_AVAILABLE:
            self._log("")
            self._log("⚠️  office365/msal not found. Run:")
            self._log(f"   {sys.executable} -m pip install Office365-REST-Python-Client msal")

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
            record_df = pd.read_excel(path, header=1)
        except Exception as exc:
            QMessageBox.critical(self, "Load Failed", str(exc))
            return

        if "URL" not in record_df.columns:
            QMessageBox.critical(
                self, "Missing Column",
                f"'URL' column not found.\nAvailable columns: {list(record_df.columns)}"
            )
            return

        dest_dir = str(Path(path).parent / Path(path).stem)
        self.record_df = record_df
        self._dest_dir   = dest_dir
        self.dest_display.setText(dest_dir)
        self._populate_table()
        self._log(f"✔ Loaded {len(record_df)} rows from {os.path.basename(path)}.")
        self._log(f"   Download directory: {dest_dir}")
        self.btn_start.setEnabled(True)

        # Auto-fill tenant hint from first SharePoint URL
        for _, row in record_df.iterrows():
            url = str(row.get("URL", "")).strip()
            if _is_sharepoint(url):
                hint = _tenant_hint_from_url(url)
                self.tenant_edit.setText(hint)
                self._auth._tenant = hint
                break

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
            {"row": r,
             "url": self.table.item(r, 0).text(),
             "filename": self.table.item(r, 1).text()}
            for r in range(self.table.rowCount())
        ]

    def _first_sharepoint_url(self) -> str | None:
        for r in self._collect_rows():
            url = r["url"].strip()
            if url and url.lower() not in ("nan", "none") and _is_sharepoint(url):
                return url
        return None

    # ── Sign-in ──────────────────────────────────────────────────────────────

    def _show_it_guide(self):
        guide_path = Path(__file__).parent / "AZURE_APP_SETUP.md"
        text = guide_path.read_text(encoding="utf-8") if guide_path.is_file() else (
            "Ask IT to register an Azure AD app with:\n"
            "• Platform: Mobile & desktop → http://localhost\n"
            "• Permissions: Files.Read.All + Sites.Read.All (delegated)\n"
            "• Admin consent granted\n"
            "• Allow public client flows: Yes\n"
            "Then paste the Application (client) ID in the Client ID field."
        )
        QMessageBox.information(self, "IT Setup Guide", f"<pre>{text[:3000]}</pre>")

    def _reset_signin_ui(self):
        self.btn_signin.setEnabled(True)
        self.btn_signin.setText("Sign in with Microsoft")

    def _sign_in(self):
        try:
            sp_url  = self._first_sharepoint_url()
            tenant  = self.tenant_edit.text().strip() or "organizations"
            sp_host = _sharepoint_host_for_signin(tenant, sp_url)
        except ValueError as exc:
            QMessageBox.warning(self, "Cannot Sign In", str(exc))
            return

        client_id = self.client_id_edit.text().strip()
        if not client_id:
            QMessageBox.warning(
                self, "Client ID Required",
                "Your company (shl-group.com) blocks all public Microsoft apps.\n\n"
                "You must get an App (client) ID from IT first.\n\n"
                "Click 'IT Setup Guide' for the registration steps,\n"
                "then paste the Client ID and try again.",
            )
            return

        self._auth._tenant  = tenant
        self._auth._app     = None
        self._auth._account = None

        self.btn_signin.setEnabled(False)
        self.btn_signin.setText("Signing in...")
        self.signin_status.setText("Opening browser...")
        self.signin_status.setStyleSheet("color: #e67e22;")
        QApplication.processEvents()

        self._log(f"Opening browser for sign-in ({sp_host}) ...")
        self._log(f"Using Client ID: {client_id[:8]}...")

        try:
            self._auth.sign_in_interactive(sp_host, [client_id])
        except Exception as exc:
            self._reset_signin_ui()
            self.signin_status.setText("Sign-in failed")
            self.signin_status.setStyleSheet("color: #e74c3c;")
            self._log(f"✘ {exc}")
            QMessageBox.critical(self, "Sign-In Failed", str(exc))
            return

        self._reset_signin_ui()
        mode = "Graph API" if self._auth.use_graph else "SharePoint REST"
        self.signin_status.setText(f"Signed in as {self._auth.username}")
        self.signin_status.setStyleSheet("color: #27ae60; font-weight: bold;")
        self._log(f"✔ Signed in as {self._auth.username}  (via {mode})")

    # ── Download ───────────────────────────────────────────────────────────────

    def _start_download(self):
        if not self._dest_dir:
            QMessageBox.warning(self, "Notice", "Please load an Excel file first.")
            return

        rows = self._collect_rows()
        if not rows:
            QMessageBox.information(self, "Notice", "No items to download.")
            return

        has_sp = any(_is_sharepoint(r["url"]) for r in rows)
        if has_sp and not self._auth.is_signed_in:
            ret = QMessageBox.question(
                self, "Sign-In Required",
                "The list contains SharePoint URLs but you are not signed in.\n"
                "SharePoint files will fail.\n\nContinue anyway?",
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
        self._worker = DownloadWorker(rows, self._dest_dir, self._auth)
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


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = MainWindow()
    win.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
