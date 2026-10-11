from __future__ import annotations

import csv
import io
import json
import math
import sys
from pathlib import Path

from openpyxl import load_workbook
from PySide6.QtCore import Qt, QStandardPaths
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QPlainTextEdit,
    QTabWidget,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

import core


class RRFWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("RRF Calculator — prototype")
        self.resize(980, 680)
        self.setAcceptDrops(True)

        self.input_path: Path | None = None
        self.default_dictionary = self._load_default_dictionary()
        self.dictionary_rows = self._load_user_dictionary()

        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_calculation_tab(), "Calculation")
        self.tabs.addTab(self._build_dictionary_tab(), "Class dictionary")
        self.setCentralWidget(self.tabs)

        self._fill_dictionary_table()
        self._log("Ready. Choose or drop an .xlsx file.")

    # -------------------- resources / persistence --------------------

    def _resource_path(self, filename: str) -> Path:
        if getattr(sys, "frozen", False):
            base = Path(getattr(sys, "_MEIPASS", Path(sys.executable).parent))
        else:
            base = Path(__file__).resolve().parent
        return base / filename

    def _load_default_dictionary(self) -> list[dict[str, str]]:
        with self._resource_path("default_classes.json").open("r", encoding="utf-8") as f:
            return json.load(f)

    def _config_dir(self) -> Path:
        path = Path(QStandardPaths.writableLocation(QStandardPaths.AppConfigLocation))
        path.mkdir(parents=True, exist_ok=True)
        return path

    def _user_dictionary_path(self) -> Path:
        return self._config_dir() / "classes.json"

    def _load_user_dictionary(self) -> list[dict[str, str]]:
        path = self._user_dictionary_path()
        if not path.exists():
            return [dict(row) for row in self.default_dictionary]
        try:
            with path.open("r", encoding="utf-8") as f:
                rows = json.load(f)
            self._validate_dictionary(rows)
            return rows
        except Exception:
            return [dict(row) for row in self.default_dictionary]

    def _save_user_dictionary(self) -> None:
        rows = self._dictionary_from_table()
        self._validate_dictionary(rows)
        with self._user_dictionary_path().open("w", encoding="utf-8") as f:
            json.dump(rows, f, ensure_ascii=False, indent=2)
        self.dictionary_rows = rows

    def _write_runtime_classes_txt(self) -> Path:
        """Generate the same semicolon dictionary format expected by the proven core."""
        rows = self._dictionary_from_table()
        self._validate_dictionary(rows)
        path = self._config_dir() / "classes_runtime.txt"
        with path.open("w", encoding="utf-8-sig", newline="") as f:
            writer = csv.DictWriter(
                f,
                fieldnames=["class", "H_formula", "summary_group"],
                delimiter=";",
            )
            writer.writeheader()
            writer.writerows(rows)
        self.dictionary_rows = rows
        return path

    # -------------------- interface --------------------

    def _build_calculation_tab(self) -> QWidget:
        root = QWidget()
        layout = QVBoxLayout(root)

        title = QLabel("RRF Calculator")
        title.setStyleSheet("font-size: 24px; font-weight: 600;")
        layout.addWidget(title)

        subtitle = QLabel(
            "Load a chromatographic XLSX file. The app validates Group Name/Volume, "
            "calculates C, H, Formula, RRF, RRF*V and w, %, then creates the Class summary heatmap."
        )
        subtitle.setWordWrap(True)
        layout.addWidget(subtitle)

        file_row = QHBoxLayout()
        self.file_label = QLabel("No file selected")
        self.file_label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        choose = QPushButton("Choose XLSX…")
        choose.clicked.connect(self._choose_file)
        file_row.addWidget(self.file_label, 1)
        file_row.addWidget(choose)
        layout.addLayout(file_row)

        button_row = QHBoxLayout()
        validate_btn = QPushButton("Validate")
        validate_btn.clicked.connect(self._validate_clicked)
        calculate_btn = QPushButton("Calculate & Save")
        calculate_btn.clicked.connect(self._calculate_clicked)
        button_row.addWidget(validate_btn)
        button_row.addWidget(calculate_btn)
        button_row.addStretch(1)
        layout.addLayout(button_row)

        self.status_label = QLabel("Status: waiting for a file")
        self.status_label.setStyleSheet("font-weight: 600;")
        layout.addWidget(self.status_label)

        self.log_box = QPlainTextEdit()
        self.log_box.setReadOnly(True)
        layout.addWidget(self.log_box, 1)
        return root

    def _build_dictionary_tab(self) -> QWidget:
        root = QWidget()
        layout = QVBoxLayout(root)

        info = QLabel(
            "Static class dictionary. Edit it when a new chemical class appears. "
            "summary_group remains in English and determines the columns of Class summary."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        self.dict_table = QTableWidget(0, 3)
        self.dict_table.setHorizontalHeaderLabels(["class", "H_formula", "summary_group"])
        self.dict_table.horizontalHeader().setStretchLastSection(True)
        layout.addWidget(self.dict_table, 1)

        controls = QHBoxLayout()
        add_btn = QPushButton("+ Add class")
        add_btn.clicked.connect(self._add_dictionary_row)
        delete_btn = QPushButton("Delete selected")
        delete_btn.clicked.connect(self._delete_dictionary_rows)
        reset_btn = QPushButton("Restore defaults")
        reset_btn.clicked.connect(self._restore_defaults)
        save_btn = QPushButton("Save dictionary")
        save_btn.clicked.connect(self._save_dictionary_clicked)

        controls.addWidget(add_btn)
        controls.addWidget(delete_btn)
        controls.addStretch(1)
        controls.addWidget(reset_btn)
        controls.addWidget(save_btn)
        layout.addLayout(controls)
        return root

    # -------------------- drag and drop --------------------

    def dragEnterEvent(self, event):
        urls = event.mimeData().urls()
        if urls and urls[0].toLocalFile().lower().endswith(".xlsx"):
            event.acceptProposedAction()

    def dropEvent(self, event):
        self._set_input_file(Path(event.mimeData().urls()[0].toLocalFile()))

    # -------------------- dictionary --------------------

    def _fill_dictionary_table(self) -> None:
        self.dict_table.setRowCount(0)
        for row in self.dictionary_rows:
            r = self.dict_table.rowCount()
            self.dict_table.insertRow(r)
            self.dict_table.setItem(r, 0, QTableWidgetItem(str(row["class"])))
            self.dict_table.setItem(r, 1, QTableWidgetItem(str(row["H_formula"])))
            self.dict_table.setItem(r, 2, QTableWidgetItem(str(row["summary_group"])))
        self.dict_table.resizeColumnsToContents()

    def _dictionary_from_table(self) -> list[dict[str, str]]:
        rows = []
        for r in range(self.dict_table.rowCount()):
            values = []
            for c in range(3):
                item = self.dict_table.item(r, c)
                values.append("" if item is None else item.text().strip())
            rows.append(
                {
                    "class": values[0],
                    "H_formula": values[1],
                    "summary_group": values[2],
                }
            )
        return rows

    def _validate_dictionary(self, rows: list[dict[str, str]]) -> None:
        seen = set()
        for i, row in enumerate(rows, start=1):
            class_name = core.normalize_text(row.get("class", ""))
            h_formula = core.normalize_text(row.get("H_formula", ""))
            summary_group = core.normalize_text(row.get("summary_group", ""))
            if not class_name or not h_formula or not summary_group:
                raise ValueError(
                    f"Dictionary row {i}: class, H_formula and summary_group must be filled."
                )
            key = class_name.casefold()
            if key in seen:
                raise ValueError(f"Duplicate class in dictionary: {class_name!r}")
            seen.add(key)
            test_h = core.safe_formula_eval(h_formula, 10)
            if not math.isfinite(test_h):
                raise ValueError(f"Invalid H_formula for {class_name!r}")

    def _add_dictionary_row(self) -> None:
        r = self.dict_table.rowCount()
        self.dict_table.insertRow(r)
        for c in range(3):
            self.dict_table.setItem(r, c, QTableWidgetItem(""))
        self.dict_table.setCurrentCell(r, 0)
        self.dict_table.editItem(self.dict_table.item(r, 0))

    def _delete_dictionary_rows(self) -> None:
        rows = sorted({index.row() for index in self.dict_table.selectedIndexes()}, reverse=True)
        for row in rows:
            self.dict_table.removeRow(row)

    def _save_dictionary_clicked(self) -> None:
        try:
            self._save_user_dictionary()
        except Exception as exc:
            QMessageBox.critical(self, "Dictionary error", str(exc))
            return
        self._log("Dictionary saved.")
        QMessageBox.information(self, "Dictionary", "Dictionary saved.")

    def _restore_defaults(self) -> None:
        answer = QMessageBox.question(
            self,
            "Restore defaults",
            "Replace the current dictionary with the built-in defaults?",
        )
        if answer != QMessageBox.Yes:
            return
        self.dictionary_rows = [dict(row) for row in self.default_dictionary]
        self._fill_dictionary_table()
        self._save_user_dictionary()
        self._log("Default dictionary restored.")

    # -------------------- file / validation --------------------

    def _choose_file(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Choose XLSX file",
            "",
            "Excel workbook (*.xlsx)",
        )
        if filename:
            self._set_input_file(Path(filename))

    def _set_input_file(self, path: Path) -> None:
        self.input_path = path
        self.file_label.setText(str(path))
        self.status_label.setText("Status: file selected")
        self._log(f"Selected: {path}")

    def _scan_file(self):
        if self.input_path is None:
            raise ValueError("Choose an XLSX file first.")

        rows = self._dictionary_from_table()
        self._validate_dictionary(rows)
        known = {core.normalize_text(r["class"]).casefold(): r for r in rows}

        wb = load_workbook(self.input_path, read_only=True, data_only=False)
        ws = wb.active
        header_row, headers = core.find_header_row_and_columns(ws)
        group_col = headers["group name"]
        volume_col = headers["volume"]

        errors = []
        warnings = []
        unknown: dict[str, list[int]] = {}
        total_rows = 0
        no_c_rows = []
        suffix_rows = []

        for row_idx in range(header_row + 1, ws.max_row + 1):
            group_name = ws.cell(row_idx, group_col).value
            volume = ws.cell(row_idx, volume_col).value

            if group_name is None and volume is None:
                continue
            total_rows += 1

            if group_name is None or core.normalize_text(group_name) == "":
                errors.append(f"row {row_idx}: empty Group Name")
                continue

            try:
                float(volume)
            except (TypeError, ValueError):
                errors.append(f"row {row_idx}: invalid Volume = {volume!r}")
                continue

            text = core.normalize_text(group_name)
            match = core.GROUP_RE.fullmatch(text)
            if not match:
                errors.append(f"row {row_idx}: cannot parse Group Name {text!r}")
                continue

            class_name = core.normalize_text(match.group("class"))
            key = class_name.casefold()
            if key not in known:
                unknown.setdefault(class_name, []).append(row_idx)
                continue

            if not match.group("C"):
                no_c_rows.append((row_idx, text))
                continue

            if match.group("suffix"):
                suffix_rows.append((row_idx, text))

            try:
                C = int(match.group("C"))
                H = core.safe_formula_eval(known[key]["H_formula"], C)
                if not math.isclose(H, round(H), abs_tol=1e-9):
                    errors.append(f"row {row_idx}: H is not integer for {text!r}")
            except Exception as exc:
                errors.append(f"row {row_idx}: {exc}")

        wb.close()

        for class_name, row_numbers in sorted(unknown.items()):
            preview = ", ".join(map(str, row_numbers[:8]))
            if len(row_numbers) > 8:
                preview += ", ..."
            errors.append(
                f"Unknown class {class_name!r}; rows: {preview}. Add it to the dictionary."
            )

        if no_c_rows:
            warnings.append(
                f"{len(no_c_rows)} row(s) have no explicit C number and will be skipped."
            )
        if suffix_rows:
            warnings.append(
                f"{len(suffix_rows)} row(s) contain '+' or '?'; numeric C will be used."
            )

        return {
            "ok": not errors,
            "rows": total_rows,
            "errors": errors,
            "warnings": warnings,
            "unknown": unknown,
        }

    def _show_scan(self, result) -> None:
        self._log("")
        self._log(f"Rows detected: {result['rows']}")
        self._log(f"Errors: {len(result['errors'])}")
        self._log(f"Warnings: {len(result['warnings'])}")
        for msg in result["errors"]:
            self._log("ERROR: " + msg)
        for msg in result["warnings"]:
            self._log("WARNING: " + msg)

        if result["ok"]:
            self.status_label.setText("Status: ready to calculate")
        else:
            self.status_label.setText("Status: action required")
            if result["unknown"]:
                self.tabs.setCurrentIndex(1)

    def _validate_clicked(self):
        try:
            result = self._scan_file()
        except Exception as exc:
            self.status_label.setText("Status: validation failed")
            self._log(f"ERROR: {exc}")
            QMessageBox.critical(self, "Validation error", str(exc))
            return None
        self._show_scan(result)
        return result

    # -------------------- calculation --------------------

    def _calculate_clicked(self) -> None:
        result = self._validate_clicked()
        if result is None or not result["ok"]:
            QMessageBox.warning(
                self,
                "Cannot calculate",
                "Fix validation errors first. Unknown classes can be added in Class dictionary.",
            )
            return

        default_output = self.input_path.with_name(self.input_path.stem + "_calculated.xlsx")
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Save calculated workbook",
            str(default_output),
            "Excel workbook (*.xlsx)",
        )
        if not filename:
            return

        output_path = Path(filename)
        if output_path.suffix.lower() != ".xlsx":
            output_path = output_path.with_suffix(".xlsx")

        try:
            self._save_user_dictionary()
            classes_path = self._write_runtime_classes_txt()

            # Capture console output from the existing proven calculator.
            buffer = io.StringIO()
            old_stdout = sys.stdout
            try:
                sys.stdout = buffer
                core.calculate_file(self.input_path, classes_path, output_path)
            finally:
                sys.stdout = old_stdout

            for line in buffer.getvalue().splitlines():
                self._log(line)

        except PermissionError:
            QMessageBox.critical(
                self,
                "Cannot save file",
                "The output workbook may be open in Excel. Close it and try again.",
            )
            return
        except Exception as exc:
            self._log(f"ERROR: {exc}")
            QMessageBox.critical(self, "Calculation error", str(exc))
            return

        self.status_label.setText("Status: finished")
        self._log(f"Saved: {output_path}")
        QMessageBox.information(
            self,
            "Done",
            f"Calculation completed successfully.\n\n{output_path}",
        )

    def _log(self, text: str) -> None:
        self.log_box.appendPlainText(text)


def main() -> None:
    app = QApplication(sys.argv)
    app.setApplicationName("RRF Calculator")
    window = RRFWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
