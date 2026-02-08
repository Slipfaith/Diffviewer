from __future__ import annotations

from PyQt6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFormLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QVBoxLayout,
    QWidget,
)

from core.qa_verify import QASheetConfig


class QAColumnMappingDialog(QDialog):
    _field_defs = [
        ("source_column", "Source text *"),
        ("original_column", "Original Translation *"),
        ("revised_column", "Revised Translation"),
        ("qa_mark_column", "QA mark (TP/FP) *"),
        ("segment_id_column", "Segment ID"),
        ("filename_column", "FileName"),
    ]

    def __init__(
        self,
        sheet_configs: list[QASheetConfig],
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setWindowTitle("QA Column Mapping")
        self.resize(640, 320)
        self._updating = False
        self._sheet_configs = [QASheetConfig.from_dict(item.to_dict()) for item in sheet_configs]

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 12, 12, 12)
        root.setSpacing(10)

        selector_row = QHBoxLayout()
        selector_row.addWidget(QLabel("Sheet:"))
        self.sheet_selector = QComboBox(self)
        for config in self._sheet_configs:
            self.sheet_selector.addItem(config.display_name())
        self.sheet_selector.currentIndexChanged.connect(self._load_sheet)
        selector_row.addWidget(self.sheet_selector, 1)
        root.addLayout(selector_row)

        form = QFormLayout()
        self._combos: dict[str, QComboBox] = {}
        for field_name, label_text in self._field_defs:
            combo = QComboBox(self)
            combo.currentIndexChanged.connect(self._on_mapping_changed)
            self._combos[field_name] = combo
            form.addRow(label_text, combo)
        root.addLayout(form)

        self.notes_label = QLabel("", self)
        self.notes_label.setWordWrap(True)
        root.addWidget(self.notes_label)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel,
            parent=self,
        )
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        root.addWidget(buttons)

        self._load_sheet(0)

    def sheet_configs(self) -> list[QASheetConfig]:
        return [QASheetConfig.from_dict(item.to_dict()) for item in self._sheet_configs]

    def _load_sheet(self, index: int) -> None:
        if index < 0 or index >= len(self._sheet_configs):
            return
        config = self._sheet_configs[index]
        self._updating = True
        try:
            for field_name, combo in self._combos.items():
                combo.blockSignals(True)
                combo.clear()
                combo.addItem("-- not set --", None)
                for column in config.columns:
                    combo.addItem(column.display_name(), column.column_letter)
                current_value = getattr(config.mapping, field_name)
                selected_idx = 0
                if current_value:
                    for item_idx in range(1, combo.count()):
                        if combo.itemData(item_idx) == current_value:
                            selected_idx = item_idx
                            break
                combo.setCurrentIndex(selected_idx)
                combo.blockSignals(False)
        finally:
            self._updating = False

        if config.notes:
            self.notes_label.setText("Notes: " + "; ".join(config.notes))
        else:
            self.notes_label.setText("Notes: auto mapping looks complete.")

    def _on_mapping_changed(self) -> None:
        if self._updating:
            return
        index = self.sheet_selector.currentIndex()
        if index < 0 or index >= len(self._sheet_configs):
            return
        config = self._sheet_configs[index]
        for field_name, combo in self._combos.items():
            setattr(config.mapping, field_name, combo.currentData())

    def accept(self) -> None:  # type: ignore[override]
        issues: list[str] = []
        has_complete_sheet = False
        for config in self._sheet_configs:
            mapping = config.mapping
            has_complete_sheet = has_complete_sheet or mapping.is_complete()

            selected = [
                mapping.source_column,
                mapping.original_column,
                mapping.revised_column,
                mapping.qa_mark_column,
                mapping.segment_id_column,
                mapping.filename_column,
            ]
            selected = [item for item in selected if item]
            if len(selected) != len(set(selected)):
                issues.append(
                    f"{config.display_name()}: the same Excel column is assigned to multiple fields."
                )

        if not has_complete_sheet:
            issues.append(
                "At least one sheet must have Source, Original Translation and QA mark mapped."
            )

        if issues:
            QMessageBox.warning(self, "Invalid mapping", "\n".join(issues[:12]))
            return
        super().accept()
