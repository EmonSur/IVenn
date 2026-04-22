from __future__ import annotations

import os
import re
import shutil
import tempfile
from importlib.resources import as_file, files

from lxml import etree
from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import QIcon, QPainter, QImage
from PySide6.QtSvgWidgets import QGraphicsSvgItem
from PySide6.QtWidgets import QCheckBox, QComboBox, QDialog, QFileDialog, QFrame, QGraphicsScene, QGraphicsView, QHBoxLayout, QLabel, QListWidget, QMenu, QPushButton, QVBoxLayout, QWidget

from ..core.themes import SET_COLOUR_THEMES

SVG_NS = {"svg": "http://www.w3.org/2000/svg"}


class RegionDetailsWindow(QDialog):
    """Window showing the elements in a selected region."""

    def __init__(self, region_label: str, elements: set[str], parent=None):
        """Initialise a window showing the elements in a selected region."""
        super().__init__(parent)
        self.setWindowTitle(f"Elements in {region_label}")

        layout = QVBoxLayout(self)

        heading = QLabel(f"Elements in {region_label}")
        layout.addWidget(heading)

        count = len(elements)

        stats_frame = QFrame()
        stats_frame.setFrameShape(QFrame.StyledPanel)
        stats_frame.setStyleSheet(
            "QFrame { background: #f5f6f8; border-radius: 4px; padding: 4px; }"
        )

        stats_layout = QVBoxLayout(stats_frame)
        stats_layout.setContentsMargins(10, 8, 10, 8)

        stats_label = QLabel(f"{count} element" + ("s" if count != 1 else ""))
        stats_layout.addWidget(stats_label)
        layout.addWidget(stats_frame)

        if count == 0:
            empty = QLabel("No elements in this region.")
            empty.setStyleSheet("color: #555;")
            layout.addWidget(empty)
            return

        list_widget = QListWidget()
        for element in sorted(elements, key=str):
            list_widget.addItem(str(element))
        layout.addWidget(list_widget)


class SetDetailsWindow(QDialog):
    """Window showing a set title, an optional description, and its elements."""

    def __init__(self, set_label: str, elements: set[str], description: str = "", parent=None):
        """Initialise a window showing a set title, an optional description, and its elements."""
        super().__init__(parent)
        self.setWindowTitle(f"Set {set_label}")

        layout = QVBoxLayout(self)

        description = str(description).strip()
        if description:
            
            description_heading = QLabel("Description")
            layout.addWidget(description_heading)
            
            desc_frame = QFrame()
            desc_frame.setFrameShape(QFrame.StyledPanel)
            desc_frame.setStyleSheet("QFrame { background: #f5f6f8; background-color: white; border-radius: 4px; padding: 4px; }")

            desc_layout = QVBoxLayout(desc_frame)
            desc_layout.setContentsMargins(10, 8, 10, 8)

            desc_label = QLabel(description)
            desc_label.setWordWrap(True)
            desc_layout.addWidget(desc_label)

            layout.addWidget(desc_frame)

        heading = QLabel(f"Elements in {set_label}")
        layout.addWidget(heading)

        if not elements:
            empty = QLabel("No elements in this set.")
            empty.setStyleSheet("color: #555;")
            layout.addWidget(empty)
            return

        list_widget = QListWidget()
        for element in sorted(elements, key=str):
            list_widget.addItem(str(element))
        layout.addWidget(list_widget)


class Viewer(QWidget):
    """Qt viewer for an IVenn controller."""

    def __init__(self, svg_path: str, controller=None):
        """Initialise the viewer window and load the first SVG diagram."""
        
        super().__init__()

        self.controller = controller
        self.current_svg_path = svg_path
        self.svg_item = None
        self.region_hitboxes: dict[str, QRectF] = {}
        self.label_hitboxes: dict[str, QRectF] = {}
        self._hover_region: str | None = None
        self._base_svg_path = svg_path
        self._hover_svg_path: str | None = None

        try:
            resource = files("ivenn").joinpath("assets", "venn-icon.ico")
            with as_file(resource) as icon_path:
                if os.path.exists(icon_path):
                    self.setWindowIcon(QIcon(str(icon_path)))
        except Exception:
            pass

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(8, 6, 8, 6)
        main_layout.setSpacing(6)

        # ----------------------------
        # Top navigation buttons
        # ----------------------------
        header_layout = QHBoxLayout()
        header_layout.setSpacing(8)

        self.nav_container = QWidget()
        nav_layout = QHBoxLayout(self.nav_container)
        nav_layout.setContentsMargins(0, 0, 0, 0)
        nav_layout.setSpacing(6)

        self.stop_btn = QPushButton("Stop")
        self.prev_btn = QPushButton("◀")
        self.next_btn = QPushButton("▶")

        self.stop_btn.setFixedWidth(56)
        self.prev_btn.setFixedWidth(32)
        self.next_btn.setFixedWidth(32)

        nav_layout.addWidget(self.stop_btn)
        nav_layout.addWidget(self.prev_btn)
        nav_layout.addWidget(self.next_btn)

        self.stop_btn.clicked.connect(self._stop)
        self.prev_btn.clicked.connect(self._prev)
        self.next_btn.clicked.connect(self._next)

        header_layout.addWidget(self.nav_container)
        header_layout.addStretch(1)
        main_layout.addLayout(header_layout)

        # ----------------------------
        # Diagram display area
        # ----------------------------
        self.scene = QGraphicsScene(self)
        self.view = QGraphicsView(self.scene)
        self.view.viewport().installEventFilter(self)
        self.view.setRenderHints(QPainter.Antialiasing | QPainter.SmoothPixmapTransform)
        self.view.setDragMode(QGraphicsView.ScrollHandDrag)
        self.view.setTransformationAnchor(QGraphicsView.AnchorUnderMouse)
        self.view.setResizeAnchor(QGraphicsView.AnchorUnderMouse)
        self.view.setMouseTracking(True)
        self.view.viewport().setMouseTracking(True)
        main_layout.addWidget(self.view, stretch=1)

        # ----------------------------
        # Bottom controls
        # ----------------------------
        bottom_layout = QHBoxLayout()
        bottom_layout.setSpacing(10)
        bottom_layout.addStretch(1)

        # Export controls
        export_group = QVBoxLayout()
        export_group.setSpacing(2)

        export_label = QLabel("Export")
        export_label.setAlignment(Qt.AlignCenter)
        export_label.setStyleSheet("font-size: 11px; color: #666;")

        export_btn = QPushButton("Export")
        export_btn.setFixedWidth(90)

        export_menu = QMenu(export_btn)
        export_svg_action = export_menu.addAction("As SVG")
        export_png_action = export_menu.addAction("As PNG")
        export_sets_excel_action = export_menu.addAction("Sets as Excel")
        export_intersections_excel_action = export_menu.addAction("Intersections as Excel")

        export_svg_action.triggered.connect(self._export_svg)
        export_png_action.triggered.connect(self._export_png)
        export_sets_excel_action.triggered.connect(self._export_sets_excel)
        export_intersections_excel_action.triggered.connect(self._export_intersections_excel)

        export_btn.setMenu(export_menu)
        export_group.addWidget(export_label)
        export_group.addWidget(export_btn)

        # Font controls
        font_group = QVBoxLayout()
        font_group.setSpacing(2)

        font_label = QLabel("Text size")
        font_label.setAlignment(Qt.AlignCenter)
        font_label.setStyleSheet("font-size: 11px; color: #666;")

        font_buttons = QHBoxLayout()
        font_buttons.setSpacing(4)

        font_minus = QPushButton("-")
        font_plus = QPushButton("+")
        self.font_value_label = QLabel(f"{int(round(14 * getattr(self.controller, 'font_scale', 1.0)))}px")
        self.font_value_label.setAlignment(Qt.AlignCenter)
        self.font_value_label.setFixedWidth(44)

        font_minus.setFixedWidth(28)
        font_plus.setFixedWidth(28)
        font_minus.clicked.connect(self._font_decrease)
        font_plus.clicked.connect(self._font_increase)

        font_buttons.addWidget(font_minus)
        font_buttons.addWidget(self.font_value_label)
        font_buttons.addWidget(font_plus)

        font_group.addWidget(font_label)
        font_group.addLayout(font_buttons)

        # Opacity controls
        opacity_group = QVBoxLayout()
        opacity_group.setSpacing(2)

        opacity_label = QLabel("Opacity")
        opacity_label.setAlignment(Qt.AlignCenter)
        opacity_label.setStyleSheet("font-size: 11px; color: #666;")

        opacity_buttons = QHBoxLayout()
        opacity_buttons.setSpacing(4)

        opacity_minus = QPushButton("-")
        opacity_plus = QPushButton("+")
        self.opacity_value_label = QLabel(f"{int(round(40 * getattr(self.controller, 'opacity_scale', 1.0)))}%"        )
        self.opacity_value_label.setAlignment(Qt.AlignCenter)
        self.opacity_value_label.setFixedWidth(44)

        opacity_minus.setFixedWidth(28)
        opacity_plus.setFixedWidth(28)
        opacity_minus.clicked.connect(self._opacity_decrease)
        opacity_plus.clicked.connect(self._opacity_increase)

        opacity_buttons.addWidget(opacity_minus)
        opacity_buttons.addWidget(self.opacity_value_label)
        opacity_buttons.addWidget(opacity_plus)

        opacity_group.addWidget(opacity_label)
        opacity_group.addLayout(opacity_buttons)

        # Theme controls
        theme_group = QVBoxLayout()
        theme_group.setSpacing(2)

        theme_label = QLabel("Theme")
        theme_label.setAlignment(Qt.AlignCenter)
        theme_label.setStyleSheet("font-size: 11px; color: #666;")

        self.theme_combo = QComboBox()

        if self.controller and hasattr(self.controller, "available_themes"):
            names = list(self.controller.available_themes())
        else:
            names = [name for name in SET_COLOUR_THEMES.keys() if not name.startswith("_")]
            names.sort()
            if "Default" in names:
                names.remove("Default")
                names.insert(0, "Default")

        if self.controller and getattr(self.controller, "custom_theme", None):
            if "Custom" not in names:
                names.append("Custom")

        self.theme_combo.addItems(names)
        self.theme_combo.currentTextChanged.connect(self._theme_changed)

        if self.controller:
            current_theme = getattr(self.controller, "theme", "Default")
            if current_theme == "Custom" and getattr(self.controller, "custom_theme", None):
                if self.theme_combo.findText("Custom") == -1:
                    self.theme_combo.addItem("Custom")

            index = self.theme_combo.findText(current_theme)
            if index != -1:
                self.theme_combo.blockSignals(True)
                self.theme_combo.setCurrentIndex(index)
                self.theme_combo.blockSignals(False)

        theme_group.addWidget(theme_label)
        theme_group.addWidget(self.theme_combo)

        # Percentage display controls
        percent_group = QVBoxLayout()
        percent_group.setSpacing(2)

        percent_label = QLabel("Display")
        percent_label.setAlignment(Qt.AlignCenter)
        percent_label.setStyleSheet("font-size: 11px; color: #666;")

        self.percent_checkbox = QCheckBox("Show %")
        self.percent_checkbox.toggled.connect(self._toggle_percentages)

        if self.controller:
            self.percent_checkbox.setChecked(
                bool(getattr(self.controller, "show_percentages", False))
            )

        percent_group.addWidget(percent_label)
        percent_group.addWidget(self.percent_checkbox)

        bottom_layout.addLayout(font_group)
        bottom_layout.addLayout(opacity_group)
        bottom_layout.addLayout(theme_group)
        bottom_layout.addLayout(export_group)
        bottom_layout.addLayout(percent_group)

        main_layout.addLayout(bottom_layout)

        self.load_svg(svg_path)
        self.view.scale(0.8, 0.8)
        self._update_nav_ui()

    def closeEvent(self, event):
        """Clean up temporary hover files when the viewer window closes."""
        self._cleanup_hover_svg()
        super().closeEvent(event)

    def load_svg(self, svg_path: str):
        """Load a new SVG diagram into the viewer and refresh its hitboxes."""
        self.current_svg_path = svg_path
        self.scene.clear()
        self._base_svg_path = svg_path
        self._set_hover_region(None)

        self.svg_item = QGraphicsSvgItem(svg_path)
        self.scene.addItem(self.svg_item)
        self.scene.setSceneRect(self.svg_item.boundingRect())

        self.region_hitboxes.clear()
        self.label_hitboxes.clear()
        self._extract_region_hitboxes(svg_path)
        self._update_nav_ui()

        if self.controller:
            has_custom = bool(getattr(self.controller, "custom_theme", None))
            # if there is a custom theme saved, include it in theme drop down
            if has_custom and self.theme_combo.findText("Custom") == -1:
                self.theme_combo.addItem("Custom")

    def _extract_region_hitboxes(self, svg_path: str):
        """Extract clickable hitboxes for region labels from the SVG file."""
        tree = etree.parse(svg_path)
        root = tree.getroot()

        transform_matrix = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
        group = root.find(".//svg:g[@transform]", namespaces=SVG_NS)
        if group is not None:
            transform_matrix = self._parse_svg_transform(group.get("transform", ""))

        for text in root.xpath(".//svg:text", namespaces=SVG_NS):
            node_id = text.get("id")
            if not node_id:
                continue

            x = float(text.get("x", 0))
            y = float(text.get("y", 0))

            viewbox = root.get("viewBox")
            if viewbox:
                vx, vy, _, _ = map(float, viewbox.split())
                x -= vx
                y -= vy

            a, b, c, d, e, f = transform_matrix
            tx = a * x + c * y + e
            ty = b * x + d * y + f

            style = text.get("style", "")
            font_match = re.search(r"font-size:\s*([-\d.]+)px", style)
            font_size = float(font_match.group(1)) if font_match else 14.0

            text_value = "".join(text.itertext()).strip()

            if node_id.startswith("label"):
                set_letter = node_id[-1].upper()

                width = max(40.0, len(text_value) * font_size * 0.62)
                height = font_size * 1.35

                self.label_hitboxes[set_letter] = QRectF(
                    tx - width / 2,
                    ty - height * 0.85,
                    width,
                    height,
                )
                continue

            if node_id.startswith("total"):
                continue

            region_id = node_id.lower()
            if not region_id.isalpha():
                continue

            self.region_hitboxes[region_id] = QRectF(tx, ty - 23 * 0.72, 25, 23)

    def _parse_svg_transform(self, transform: str) -> tuple[float, float, float, float, float, float]:
        """Parse an SVG transform string into a 2D affine transform matrix."""
        transform = (transform or "").strip()
        if not transform:
            return (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)

        matrix_match = re.search(
            r"matrix\(\s*([-\d.]+)\s*,\s*([-\d.]+)\s*,\s*([-\d.]+)\s*,\s*([-\d.]+)\s*,\s*([-\d.]+)\s*,\s*([-\d.]+)\s*\)",
            transform,
        )
        if matrix_match:
            return (
                float(matrix_match.group(1)),
                float(matrix_match.group(2)),
                float(matrix_match.group(3)),
                float(matrix_match.group(4)),
                float(matrix_match.group(5)),
                float(matrix_match.group(6)),
            )

        translate_match = re.search(
            r"translate\(\s*([-\d.]+)\s*(?:,\s*([-\d.]+)\s*)?\)",
            transform,
        )
        if translate_match:
            tx = float(translate_match.group(1))
            ty = float(translate_match.group(2)) if translate_match.group(2) is not None else 0.0
            return (1.0, 0.0, 0.0, 1.0, tx, ty)

        return (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)

    def eventFilter(self, obj, event):
        """Handle mouse clicks and hover events over the SVG viewport."""
        if obj is self.view.viewport():
            if event.type() == event.Type.MouseButtonPress and event.button() == Qt.LeftButton:
                scene_position = self.view.mapToScene(event.pos())

                for set_letter, rect in self.label_hitboxes.items():
                    if rect.contains(scene_position):
                        self._open_set_details(set_letter)
                        return True

                for region, rect in self.region_hitboxes.items():
                    if rect.contains(scene_position):
                        self._open_region_details(region)
                        return True

            elif event.type() == event.Type.MouseMove:
                if event.buttons() & Qt.LeftButton:
                    return False

                scene_position = self.view.mapToScene(event.pos())
                hovered = None
                for region, rect in self.region_hitboxes.items():
                    if rect.contains(scene_position):
                        hovered = region
                        break

                self._set_hover_region(hovered)
                return False

            elif event.type() == event.Type.Leave:
                self._set_hover_region(None)
                return False

        return super().eventFilter(obj, event)

    def _set_hover_region(self, region_key: str | None):
        """Update the currently hovered region and refresh highlight styling."""
        if region_key == self._hover_region:
            return

        self._hover_region = region_key

        # if no region is hovered over, return to base SVG
        if not region_key:
            if self.svg_item is not None:
                self.scene.removeItem(self.svg_item)
            self.svg_item = QGraphicsSvgItem(self._base_svg_path)
            self.scene.addItem(self.svg_item)
            self.scene.setSceneRect(self.svg_item.boundingRect())
            self._cleanup_hover_svg()
            return

        highlight_sets = {char.upper() for char in region_key if char.isalpha()}
        self._apply_hover_highlight(highlight_sets)

    def _cleanup_hover_svg(self) -> None:
        '''Delete the temporary SVG file if it still exists'''
        if not self._hover_svg_path:
            return

        try:
            if os.path.exists(self._hover_svg_path):
                os.remove(self._hover_svg_path)
        except Exception:
            pass
        finally:
            self._hover_svg_path = None

    def _apply_hover_highlight(self, highlight_sets: set[str]):
        """Create and display a temporary SVG with highlighted set outlines."""
        tree = etree.parse(self._base_svg_path)
        root = tree.getroot()

        set_id_re = re.compile(r"ellipse([A-F])$", re.IGNORECASE)

        for path in root.xpath(".//svg:path", namespaces=SVG_NS):
            path_id = path.get("id", "")
            match = set_id_re.match(path_id)
            if not match:
                continue

            letter = match.group(1).upper()
            style = path.get("style", "")

            parts = [part for part in style.split(";") if part and not part.strip().startswith(("stroke:", "stroke-width:"))]

            if letter in highlight_sets:
                if self.controller and hasattr(self.controller, "sets"):
                    n_way = len(self.controller.sets)
                else:
                    match_n = re.search(r"(\d)waydiagram", os.path.basename(self.current_svg_path).lower())
                    n_way = int(match_n.group(1)) if match_n else 0

                standalone_letters = set("ABCDEF")
                if self.controller and hasattr(self.controller, "union_states"):
                    current_unions = self.controller.union_states[self.controller.current_index]
                    grouped_letters = set()
                    for union in current_unions:
                        grouped_letters |= {char.upper() for char in union}
                    standalone_letters -= grouped_letters

                if n_way == 6:
                    stroke_width = 3.5 / 5 if letter in {"A", "B", "C"} and letter in standalone_letters else (3.5 * 2) / 3
                else:
                    stroke_width = (3.5 * 2) / 3

                parts.append("stroke:#000000")
                parts.append(f"stroke-width:{stroke_width}")

            path.set("style", ";".join(parts))

        self._cleanup_hover_svg()

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".svg")
        self._hover_svg_path = tmp.name
        tmp.close()

        tree.write(self._hover_svg_path)

        if self.svg_item is not None:
            self.scene.removeItem(self.svg_item)

        self.svg_item = QGraphicsSvgItem(self._hover_svg_path)
        self.scene.addItem(self.svg_item)
        self.scene.setSceneRect(self.svg_item.boundingRect())

    def _update_nav_ui(self):
        """Update the navigation controls to match the current union state."""
        if not self.controller:
            self.nav_container.setVisible(False)
            return

        if hasattr(self.controller, "_has_unions"):
            has_unions = bool(self.controller._has_unions())
        else:
            has_unions = hasattr(self.controller, "union_states") and len(self.controller.union_states) > 1

        self.nav_container.setVisible(has_unions)
        if not has_unions:
            return

        mode = getattr(self.controller, "union_mode", "list")
        if mode == "tree":
            self.prev_btn.setText("▼")
            self.next_btn.setText("▲")
        else:
            self.prev_btn.setText("◀")
            self.next_btn.setText("▶")

    def _export_svg(self):
        """Save the currently displayed diagram as an SVG file."""
        path, _ = QFileDialog.getSaveFileName(self, "Save SVG", "diagram.svg", "SVG files (*.svg)")
        if not path:
            return

        if self.controller and hasattr(self.controller, "export_svg"):
            self.controller.export_svg(path)
        elif self.current_svg_path:
            shutil.copyfile(self.current_svg_path, path)

    def _export_png(self):
        """Save the currently displayed diagram as a PNG file."""
        path, _ = QFileDialog.getSaveFileName(self, "Save PNG", "diagram.png", "PNG files (*.png)")
        if not path:
            return

        if self.controller and hasattr(self.controller, "export_png"):
            self.controller.export_png(path)
            return

        if not self.svg_item:
            return

        rect = self.svg_item.boundingRect()
        image = QImage(int(rect.width()), int(rect.height()), QImage.Format_ARGB32)
        image.fill(Qt.transparent)

        painter = QPainter(image)
        try:
            self.scene.render(painter)
        finally:
            painter.end()

        image.save(path)

    def _export_sets_excel(self):
        """Export the current sets to an Excel file."""

        if not self.controller:
            return

        path, _ = QFileDialog.getSaveFileName(self, "Save Sets Excel", "venn_sets.xlsx", "Excel files (*.xlsx)")
        if path:
            self.controller.export_sets(path)

    def _export_intersections_excel(self):
        """Export the current intersections to an Excel file."""
        if not self.controller:
            return

        path, _ = QFileDialog.getSaveFileName(self, "Save Intersections Excel", "venn_intersections.xlsx", "Excel files (*.xlsx)")
        if path:
            self.controller.export_intersections(path)

    def _open_region_details(self, region_key: str):
        """Open a window showing the elements contained in a selected region."""
        if not self.controller:
            return

        elements = self.controller.get_intersection(region_key)
        normalised_key = self.controller._normalise_intersection_lookup(region_key)
        region_label = self.controller._format_region_label(normalised_key)

        dialog = RegionDetailsWindow(region_label, elements, self)
        dialog.exec()

    def _open_set_details(self, set_letter: str):
        """Open a window showing the details of a selected set."""
        if not self.controller:
            return

        set_letter = set_letter.upper()

        elements = set()
        if hasattr(self.controller, "sets"):
            elements = set(self.controller.sets.get(set_letter, set()))

        set_label = set_letter
        if hasattr(self.controller, "labels"):
            set_label = self.controller.labels.get(set_letter, set_letter)

        description = ""
        if hasattr(self.controller, "descriptions"):
            description = str(self.controller.descriptions.get(set_letter, "")).strip()

        dialog = SetDetailsWindow(
            set_label=set_label,
            elements=elements,
            description=description,
            parent=self,
        )
        dialog.exec()
    
    def _stop(self):
        """Reset the viewer to the base diagram with no active unions."""
        if not self.controller:
            return

        if hasattr(self.controller, "reset_union_view"):
            self.controller.reset_union_view()
        elif hasattr(self.controller, "_stop"):
            self.controller._stop()

    def _prev(self):
        """Move to the previous union state in the viewer."""
        if not self.controller:
            return

        if hasattr(self.controller, "_prev"):
            self.controller._prev()
        elif hasattr(self.controller, "prev_state"):
            self.controller.prev_state()
        else:
            self.controller.prev_diagram()

    def _next(self):
        """Move to the next union state in the viewer."""
        if not self.controller:
            return

        if hasattr(self.controller, "_next"):
            self.controller._next()
        elif hasattr(self.controller, "next_state"):
            self.controller.next_state()
        else:
            self.controller.next_diagram()

    def _theme_changed(self, theme_name: str):
        """Apply a newly selected colour theme."""
        if not self.controller:
            return

        if theme_name == "Custom" and not getattr(self.controller, "custom_theme", None):
            self.theme_combo.blockSignals(True)
            self.theme_combo.setCurrentText("Default")
            self.theme_combo.blockSignals(False)
            return

        if hasattr(self.controller, "set_theme"):
            self.controller.set_theme(theme_name)

    def wheelEvent(self, event):
        """Zoom the diagram in or out using the mouse wheel."""
        factor = 1.15 if event.angleDelta().y() > 0 else 1 / 1.15
        self.view.scale(factor, factor)

    def _font_increase(self):
        """Increase the diagram text size."""
        if not self.controller or not hasattr(self.controller, "_font"):
            return

        current_scale = float(getattr(self.controller, "font_scale", 1.0))
        current_pt = 14.0 * current_scale
        new_pt = min(30.0, current_pt + 2.0)
        self.controller._font(new_pt / 14.0)
        self.font_value_label.setText(f"{int(round(new_pt))}px")

    def _font_decrease(self):
        """Decrease the diagram text size."""
        if not self.controller or not hasattr(self.controller, "_font"):
            return

        current_scale = float(getattr(self.controller, "font_scale", 1.0))
        current_pt = 14.0 * current_scale
        new_pt = max(6.0, current_pt - 2.0)
        self.controller._font(new_pt / 14.0)
        self.font_value_label.setText(f"{int(round(new_pt))}px")

    def _opacity_increase(self):
        """Increase the diagram opacity."""
        if not self.controller or not hasattr(self.controller, "_opacity"):
            return

        current_scale = float(getattr(self.controller, "opacity_scale", 1.0))
        current_percent = 40.0 * current_scale
        new_percent = min(80.0, current_percent + 10.0)
        self.controller._opacity(new_percent / 40.0)
        self.opacity_value_label.setText(f"{int(round(new_percent))}%")

    def _opacity_decrease(self):
        """Decrease the diagram opacity"""
        if not self.controller or not hasattr(self.controller, "_opacity"):
            return

        current_scale = float(getattr(self.controller, "opacity_scale", 1.0))
        current_percent = 40.0 * current_scale
        new_percent = max(20.0, current_percent - 10.0)
        self.controller._opacity(new_percent / 40.0)
        self.opacity_value_label.setText(f"{int(round(new_percent))}%")

    def _toggle_percentages(self, checked: bool):
        """Toggle whether region values are shown as percentages instead of counts."""
        if not self.controller:
            return

        self.controller._set_show_percentages(checked)