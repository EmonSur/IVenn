import os
import sys
import tempfile
import traceback
from importlib.resources import as_file, files

from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon, QImage, QPainter
from PySide6.QtSvg import QSvgRenderer
from PySide6.QtWidgets import QApplication, QMessageBox

from .viewer import Viewer


def _cleanup_previous_render(controller) -> None:
    """Delete the previously rendered temporary SVG file, if one exists."""
    old_path = getattr(controller, "_current_svg_path", None)
    if not old_path: # if no previous path, nothing to clean up
        return

    try:
        if os.path.exists(old_path):
            os.remove(old_path)
    except Exception:
        pass
    finally:
        controller._current_svg_path = None

def render_controller(controller, start: bool = False, base_dir: str | None = None) -> None:
    """Render the controller's current diagram in the Qt viewer."""
    try:
        app = QApplication.instance()
        if app is None:
            app = QApplication(sys.argv)

        try:
            resource = files("ivenn").joinpath("assets", "venn-icon.ico")
            with as_file(resource) as icon_path:
                if os.path.exists(icon_path):
                    app.setWindowIcon(QIcon(str(icon_path)))
        except Exception:
            pass

        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".svg")
        svg_path = tmp.name
        tmp.close()

        controller._render_svg(svg_path, base_dir=base_dir)

        _cleanup_previous_render(controller)
        controller._current_svg_path = svg_path

        if controller._viewer is None:
            controller._viewer = Viewer(svg_path, controller=controller)
            controller._viewer.showMaximized()
        else:
            controller._viewer.load_svg(svg_path)

        controller._viewer.setWindowTitle("IVenn")

        try:
            resource = files("ivenn").joinpath("assets", "venn-icon.ico")
            with as_file(resource) as icon_path:
                if os.path.exists(icon_path):
                    controller._viewer.setWindowIcon(QIcon(str(icon_path)))
        except Exception:
            pass

        if start:
            app.exec()

    except Exception as exc:
        traceback.print_exc()

        # try to get the existing viewer from the controller
        viewer = getattr(controller, "_viewer", None)
        if viewer is None:
            raise # no viewer to show an error in, so just reraise the error

        # show a popup error box inside the viewer
        QMessageBox.critical(
            viewer,
            "Render failed",
            f"{type(exc).__name__}: {exc}",
        )

def export_png(controller, path: str, base_dir: str | None = None, scale: float = 2.0) -> str:
    """Save the current diagram view as a PNG file."""

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".svg")
    svg_path = tmp.name
    tmp.close()

    try:
        controller._render_svg(svg_path, base_dir=base_dir)

        renderer = QSvgRenderer(svg_path)
        size = renderer.defaultSize()

        width = max(1, int(size.width() * scale))
        height = max(1, int(size.height() * scale))

        image = QImage(width, height, QImage.Format_ARGB32)
        image.fill(Qt.transparent)

        painter = QPainter(image)
        try:
            renderer.render(painter)
        finally:
            painter.end()

        saved = image.save(path)
        if not saved:
            raise IOError(f"Failed to save PNG to: {path}")

    finally:
        try:
            if os.path.exists(svg_path):
                os.remove(svg_path)
        except Exception:
            pass

    return path