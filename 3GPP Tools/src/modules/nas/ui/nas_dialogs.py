from pathlib import Path
from typing import Any, Dict, List, Optional

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from modules.nas.core.nas_db import NASDatabase, parse_version_tuple
from modules.nas.core.nas_threads import find_cached_spec_file
from modules.specifications.core.database import SpecsDatabase


PROTOCOL_SPEC_CATALOG: Dict[str, Dict[str, str]] = {
    "24.501": {
        "acronym": "5GS NAS",
        "wg": "CT1",
        "title": "5G Non-Access-Stratum (NAS) Protocol",
        "full_title": "Non-Access-Stratum (NAS) protocol for 5G System (5GS); Stage 3",
        "tech": "5G",
        "description": "Core Network signaling between UE and AMF/SMF for 5G registration, connection, mobility, and PDU sessions.",
    },
    "24.526": {
        "acronym": "5GS UE Policies (URSP)",
        "wg": "CT1",
        "title": "UE Policies for 5G System (URSP & ANDSP); Stage 3",
        "full_title": "User Equipment (UE) policies for 5G System (5GS); Stage 3",
        "tech": "5G",
        "description": "Stage 3 bit-level encoding of UE Route Selection Policy (URSP) rules, Traffic Descriptors, Route Selection Descriptors, and ANDSP/WLANSP.",
    },
    "24.301": {
        "acronym": "EPS NAS",
        "wg": "CT1",
        "title": "4G EPS Non-Access-Stratum (NAS) Protocol",
        "full_title": "Non-Access-Stratum (NAS) protocol for Evolved Packet System (EPS); Stage 3",
        "tech": "LTE",
        "description": "Core Network signaling between UE and MME for 4G LTE attach, tracking area update, and EPS bearer establishment.",
    },
    "24.008": {
        "acronym": "Core Network L3",
        "wg": "CT1",
        "title": "Core Network Protocols; Stage 3",
        "full_title": "Mobile radio interface Layer 3 specification; Core Network protocols; Stage 3",
        "tech": "3G/2G",
        "description": "Legacy 3G/2G Mobility Management (MM), GPRS Mobility Management (GMM), and Call Control (CC).",
    },
    "29.244": {
        "acronym": "PFCP",
        "wg": "CT4",
        "title": "Packet Forwarding Control Protocol (PFCP)",
        "full_title": "Interface between the Control Plane and the User Plane nodes (PFCP); Stage 3",
        "tech": "5G/Core",
        "description": "N4/Sxa/Sxb interface protocol between SMF/SGW-C and UPF/SGW-U for packet detection, forwarding, and usage reporting.",
    },
    "29.281": {
        "acronym": "GTPv1-U",
        "wg": "CT4",
        "title": "GTP User Plane Tunneling (GTPv1-U)",
        "full_title": "General Packet Radio System (GPRS) Tunnelling Protocol User Plane (GTPv1-U)",
        "tech": "5G/Core",
        "description": "Encapsulated user-plane IP and Ethernet tunneling over UDP between radio access nodes (gNodeB) and UPF/SGW-U.",
    },
    "38.331": {
        "acronym": "NR RRC",
        "wg": "RAN2",
        "title": "5G NR Radio Resource Control (RRC)",
        "full_title": "NR; Radio Resource Control (RRC) protocol specification",
        "tech": "5G",
        "description": "Air interface signaling: Broadcast of system information (SIBs), RRC connection establishment, reconfiguration, and handover.",
    },
    "36.331": {
        "acronym": "LTE RRC",
        "wg": "RAN2",
        "title": "4G LTE Radio Resource Control (RRC)",
        "full_title": "E-UTRA; Radio Resource Control (RRC); Protocol specification",
        "tech": "LTE",
        "description": "Air interface signaling for 4G LTE: SIB broadcast, paging, connection control, measurement reporting, and mobility.",
    },
    "38.413": {
        "acronym": "NGAP",
        "wg": "RAN3",
        "title": "NG Application Protocol (NGAP)",
        "full_title": "NG-RAN; NG Application Protocol (NGAP)",
        "tech": "5G",
        "description": "N2 interface control protocol between gNodeB and 5G Core (AMF) for PDU session resource management and UE context transfer.",
    },
    "38.415": {
        "acronym": "PDU Session UP",
        "wg": "RAN3",
        "title": "PDU Session User Plane Protocol",
        "full_title": "NG-RAN; PDU Session user plane protocol",
        "tech": "5G",
        "description": "Framing protocol for user-plane data over N3/N9, conveying QoS Flow Identifiers (QFI) and Reflective QoS Indicators (RQI).",
    },
    "38.423": {
        "acronym": "XnAP",
        "wg": "RAN3",
        "title": "Xn Application Protocol (XnAP)",
        "full_title": "NG-RAN; Xn Application Protocol (XnAP)",
        "tech": "5G",
        "description": "Direct inter-gNodeB interface protocol for dual connectivity, Xn-based handover, and inter-node coordination.",
    },
    "38.473": {
        "acronym": "F1AP",
        "wg": "RAN3",
        "title": "F1 Application Protocol (F1AP)",
        "full_title": "NG-RAN; F1 Application Protocol (F1AP)",
        "tech": "5G",
        "description": "Interface between Central Unit (CU) and Distributed Unit (DU) in disaggregated gNodeB deployments.",
    },
    "38.463": {
        "acronym": "E1AP",
        "wg": "RAN3",
        "title": "E1 Application Protocol (E1AP)",
        "full_title": "NG-RAN; E1 Application Protocol (E1AP)",
        "tech": "5G",
        "description": "Interface between CU Control Plane (CU-CP) and CU User Plane (CU-UP) nodes.",
    },
    "38.412": {
        "acronym": "NG-C Transport",
        "wg": "RAN3",
        "title": "NG Signaling Transport (NG-C)",
        "full_title": "NG-RAN; NG signalling transport",
        "tech": "5G",
        "description": "SCTP transport layer specification for NGAP signaling over the N2 reference point.",
    },
    "36.413": {
        "acronym": "S1AP",
        "wg": "RAN3",
        "title": "S1 Application Protocol (S1AP)",
        "full_title": "E-UTRA; S1 Application Protocol (S1AP)",
        "tech": "LTE",
        "description": "S1-MME interface control protocol between eNodeB and 4G Core (MME) for bearer management and handover.",
    },
    "23.501": {
        "acronym": "5GS Architecture",
        "wg": "SA2",
        "title": "System Architecture for 5G System (Stage 2)",
        "full_title": "System architecture for the 5G System (5GS); Stage 2",
        "tech": "5G",
        "description": "High-level Stage 2 architectural requirements, network functions, reference points, and roaming models.",
    },
    "23.502": {
        "acronym": "5GS Procedures",
        "wg": "SA2",
        "title": "Procedures for 5G System (Stage 2)",
        "full_title": "Procedures for the 5G System (5GS); Stage 2",
        "tech": "5G",
        "description": "Stage 2 call flows, message sequences, registration procedures, and PDU session establishment workflows.",
    },
}


def get_spec_metadata(spec_number: str) -> Dict[str, str]:
    """Returns rich metadata dictionary for a 3GPP specification number."""
    clean_num = str(spec_number or "").replace("TS", "").strip()
    if clean_num in PROTOCOL_SPEC_CATALOG:
        return PROTOCOL_SPEC_CATALOG[clean_num]

    # Heuristic fallback based on 3GPP specification series
    wg = "3GPP"
    tech = "Multi"
    if clean_num.startswith("24."):
        wg = "CT1"
        tech = "Core"
    elif clean_num.startswith("29."):
        wg = "CT4"
        tech = "Core"
    elif clean_num.startswith("38."):
        wg = "RAN2" if clean_num.startswith("38.3") else "RAN3"
        tech = "5G"
    elif clean_num.startswith("36."):
        wg = "RAN2" if clean_num.startswith("36.3") else "RAN3"
        tech = "LTE"
    elif clean_num.startswith("23."):
        wg = "SA2"
        tech = "Architecture"

    return {
        "acronym": f"TS {clean_num}",
        "wg": wg,
        "title": f"TS {clean_num} Specification",
        "full_title": f"3GPP TS {clean_num} Technical Specification",
        "tech": tech,
        "description": f"Technical Specification TS {clean_num} governed by 3GPP {wg}.",
    }


def get_spec_title(spec_number: str) -> str:
    """Returns backwards-compatible short title string with acronym."""
    clean_num = str(spec_number or "").replace("TS", "").strip()
    meta = get_spec_metadata(clean_num)
    return f"TS {clean_num} ({meta['acronym']})"


def get_spec_display_title(spec_number: str) -> str:
    """Returns a rich single-line display label: e.g. [CT1] TS 24.501 — 5GS NAS."""
    clean_num = str(spec_number or "").replace("TS", "").strip()
    meta = get_spec_metadata(clean_num)
    return f"[{meta['wg']}] TS {clean_num} — {meta['title']}"


def get_spec_tooltip(spec_number: str) -> str:
    """Returns multi-line informative tooltip describing the specification's 3GPP role."""
    clean_num = str(spec_number or "").replace("TS", "").strip()
    meta = get_spec_metadata(clean_num)
    return (
        f"3GPP TS {clean_num}: {meta['full_title']}\n"
        f"Working Group: {meta['wg']}  •  Domain: {meta['tech']}\n\n"
        f"Role: {meta['description']}"
    )


class NASVersionSelectDialog(QDialog):
    """Dialog allowing selection of NAS (24.501/24.301) and RRC/NGAP (38.331/36.331/38.413) versions."""

    def __init__(
            self,
            specs_db: SpecsDatabase,
            nas_db: NASDatabase,
            cache_dir: Path,
            parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.specs_db = specs_db
        self.nas_db = nas_db
        self.cache_dir = cache_dir
        self.selected_files_info: List[Dict[str, Any]] = []

        self.setWindowTitle("Select 3GPP Specification Version(s) to Ingest")
        self.resize(780, 520)
        self._setup_ui()
        self._load_available_versions()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        top_bar = QHBoxLayout()
        top_bar.addWidget(QLabel("Specification:"))

        self.spec_combo = QComboBox()
        # Populate with rich titles and WG clues
        for s_num in [
            "38.331",
            "36.331",
            "38.413",
            "38.423",
            "38.473",
            "38.463",
            "36.413",
            "38.412",
            "38.415",
            "29.244",
            "29.281",
            "24.501",
            "24.526",
            "24.301",
            "24.008",
        ]:
            meta = get_spec_metadata(s_num)
            label = f"TS {s_num} — {meta['title']} [{meta['wg']}]"
            self.spec_combo.addItem(label, s_num)

        self.spec_combo.currentIndexChanged.connect(self._on_combo_changed)
        top_bar.addWidget(self.spec_combo, 1)
        layout.addLayout(top_bar)

        # Explanatory card banner describing the selected specification
        self.spec_info_frame = QFrame()
        self.spec_info_frame.setStyleSheet("""
            QFrame {
                background-color: #F8FAFC;
                border: 1px solid #E2E8F0;
                border-left: 4px solid #1E5C99;
                border-radius: 4px;
                padding: 6px 10px;
            }
        """)
        spec_info_layout = QVBoxLayout(self.spec_info_frame)
        spec_info_layout.setContentsMargins(4, 2, 4, 2)
        self.spec_info_title = QLabel()
        self.spec_info_title.setStyleSheet("font-weight: bold; font-size: 12px; color: #0F172A;")
        self.spec_info_desc = QLabel()
        self.spec_info_desc.setStyleSheet("font-size: 11px; color: #475569;")
        self.spec_info_desc.setWordWrap(True)
        spec_info_layout.addWidget(self.spec_info_title)
        spec_info_layout.addWidget(self.spec_info_desc)
        layout.addWidget(self.spec_info_frame)

        info_lbl = QLabel(
            "Select one or more versions from the specification archive (use Ctrl+Click or Shift+Click).\n"
            "Files not currently cached locally will be downloaded automatically from the 3GPP FTP server."
        )
        info_lbl.setStyleSheet("color: #64748B; font-size: 11px; padding-bottom: 2px;")
        layout.addWidget(info_lbl)

        self.table = QTableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Version", "Filename", "Local Cache", "DB Status"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setSelectionMode(QTableWidget.ExtendedSelection)
        self.table.itemDoubleClicked.connect(self._on_item_double_clicked)
        self.table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #CBD5E1;
                border-radius: 4px;
                gridline-color: #F1F5F9;
            }
            QTableWidget::item {
                padding: 4px;
            }
        """)
        layout.addWidget(self.table)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.button(QDialogButtonBox.Ok).setText("📥 Fetch && Ingest Selected")
        btn_box.accepted.connect(self._on_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def _on_combo_changed(self):
        self._update_spec_banner()
        self._load_available_versions()

    def _update_spec_banner(self):
        spec_num = self.spec_combo.currentData()
        meta = get_spec_metadata(spec_num)
        self.spec_info_title.setText(f"TS {spec_num}: {meta['full_title']}")
        self.spec_info_desc.setText(
            f"Working Group: <b>{meta['wg']}</b> • Domain: <b>{meta['tech']}</b> • {meta['description']}"
        )

    def _load_available_versions(self):
        self._update_spec_banner()
        spec_num = self.spec_combo.currentData()
        spec_files = self.specs_db.search_files(spec_number=spec_num) if self.specs_db else []

        imported_entries = {
            (v["spec_number"], v["version"]) for v in self.nas_db.get_imported_versions()
        }

        spec_files = sorted(
            spec_files,
            key=lambda row: parse_version_tuple(row[5]),
            reverse=True,
        )

        self.table.setRowCount(0)

        for row_idx, row_data in enumerate(spec_files):
            _, s_num, _, _, filename, version, url, specification_upload_date, *rest = row_data
            self.table.insertRow(row_idx)

            v_item = QTableWidgetItem(f"v{version}")
            v_item.setData(
                Qt.UserRole,
                {
                    "spec_number": s_num,
                    "version": version,
                    "filename": filename,
                    "file_url": url,
                },
            )
            self.table.setItem(row_idx, 0, v_item)
            self.table.setItem(row_idx, 1, QTableWidgetItem(filename))

            cached_file = find_cached_spec_file(filename, s_num)
            if cached_file:
                cache_text = f"🟢 Cached ({cached_file.suffix[1:].upper()})"
                cache_item = QTableWidgetItem(cache_text)
                cache_item.setForeground(Qt.darkGreen)
            else:
                cache_item = QTableWidgetItem("🌐 Remote (FTP)")
                cache_item.setForeground(Qt.darkGray)
            self.table.setItem(row_idx, 2, cache_item)

            in_db = (s_num, version) in imported_entries
            db_item = QTableWidgetItem("✅ Ingested" if in_db else "⚪ Ready")
            if in_db:
                db_item.setForeground(Qt.blue)
            self.table.setItem(row_idx, 3, db_item)

    def _on_item_double_clicked(self, item: QTableWidgetItem):
        row = item.row()
        self.selected_files_info = [self.table.item(row, 0).data(Qt.UserRole)]
        self.accept()

    def _on_accept(self):
        selected_rows = sorted(list(set(item.row() for item in self.table.selectedItems())))
        if not selected_rows:
            QMessageBox.warning(
                self, "Selection Required", "Please select at least one specification version to import."
            )
            return

        self.selected_files_info = [
            self.table.item(row, 0).data(Qt.UserRole)
            for row in selected_rows
        ]
        self.accept()
