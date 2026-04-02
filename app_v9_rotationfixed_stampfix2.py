import json
import copy
import re
import sys
import shutil
import tempfile
import uuid
from datetime import datetime
from email import policy
from email.parser import BytesParser
from pathlib import Path

import fitz  # PyMuPDF
from PIL import Image, ImageOps
from pypdf import PdfReader, PdfWriter

from PySide6.QtCore import Qt, QSize, QUrl, Signal, QSettings, QRect, QPoint
from PySide6.QtGui import (
    QAction,
    QDesktopServices,
    QDragEnterEvent,
    QDropEvent,
    QPixmap,
    QImage,
    QIcon,
    QKeySequence,
    QShortcut,
    QColor,
    QPainter,
    QPen,
    QBrush,
)
from PySide6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QFileDialog,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QListWidget,
    QListWidgetItem,
    QListView,
    QMainWindow,
    QMenu,
    QMessageBox,
    QScrollArea,
    QPushButton,
    QSlider,
    QSplitter,
    QStyle,
    QToolBar,
    QVBoxLayout,
    QWidget,
    QDialog,
    QPlainTextEdit,
    QDialogButtonBox,
    QTextEdit,
    QColorDialog,
)

APP_NAME = "Michelon PDF Sandbox"
SUPPORTED_IMPORTS = {".pdf", ".jpg", ".jpeg", ".png", ".doc", ".docx", ".eml", ".msg"}
WORD_EXTS = {".doc", ".docx"}
EMAIL_EXTS = {".eml", ".msg"}
PDF_WD_FORMAT = 17
A4_WIDTH = 595  # points
A4_HEIGHT = 842
PAGE_MARGIN = 28

try:
    import win32com.client  # type: ignore
    HAS_WIN32 = True
except Exception:
    HAS_WIN32 = False


def sandbox_root() -> Path:
    root = Path(tempfile.gettempdir()) / "michelon_pdf_sandbox"
    root.mkdir(parents=True, exist_ok=True)
    return root


def session_root() -> Path:
    root = sandbox_root() / f"session_{uuid.uuid4().hex[:8]}"
    root.mkdir(parents=True, exist_ok=True)
    return root


def safe_stem(name: str) -> str:
    name = name.replace("\n", " ").replace("\r", " ")
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    name = name.rstrip(". ")
    return name or "document"


def ensure_project_files_dir(project_dir: Path) -> Path:
    files_dir = project_dir / "files"
    files_dir.mkdir(parents=True, exist_ok=True)
    return files_dir


def format_datetime_value(value) -> str:
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d/%m/%Y %H:%M")
    text = str(value).strip()
    return text


class FileOps:
    @staticmethod
    def unique_path(folder: Path, preferred_name: str, suffix: str | None = None) -> Path:
        preferred_name = preferred_name.strip()

        if suffix is None:
            p = Path(preferred_name)
            final_suffix = p.suffix if p.suffix else ".pdf"
            raw_stem = p.stem if p.suffix else preferred_name
        else:
            final_suffix = suffix
            raw_stem = preferred_name

        stem = safe_stem(raw_stem)
        candidate = folder / f"{stem}{final_suffix}"
        i = 1
        while candidate.exists():
            candidate = folder / f"{stem}_{i}{final_suffix}"
            i += 1
        return candidate

    @staticmethod
    def unique_pdf_path(folder: Path, preferred_name: str) -> Path:
        return FileOps.unique_path(folder, preferred_name, ".pdf")

    @staticmethod
    def image_to_pdf(src: Path, dst: Path) -> Path:
        image = Image.open(src)
        image = ImageOps.exif_transpose(image)
        if image.mode != "RGB":
            image = image.convert("RGB")

        temp_image = dst.with_suffix(".tmp_image_for_pdf.png")
        image.save(temp_image, format="PNG")

        doc = fitz.open()
        page = doc.new_page(width=A4_WIDTH, height=A4_HEIGHT)
        img_w, img_h = image.size

        available_w = A4_WIDTH - (2 * PAGE_MARGIN)
        available_h = A4_HEIGHT - (2 * PAGE_MARGIN)
        scale = min(available_w / img_w, available_h / img_h)
        draw_w = img_w * scale
        draw_h = img_h * scale
        x0 = (A4_WIDTH - draw_w) / 2
        y0 = (A4_HEIGHT - draw_h) / 2
        rect = fitz.Rect(x0, y0, x0 + draw_w, y0 + draw_h)
        page.insert_image(rect, filename=str(temp_image), keep_proportion=True)
        doc.save(str(dst), garbage=4, deflate=True)
        doc.close()

        try:
            temp_image.unlink(missing_ok=True)
        except Exception:
            pass
        return dst

    @staticmethod
    def copy_word_to_sandbox(src: Path, folder: Path) -> Path:
        dst = FileOps.unique_path(folder, src.name, src.suffix.lower())
        shutil.copy2(src, dst)
        return dst

    @staticmethod
    def word_to_pdf(src: Path, dst: Path) -> Path:
        if not HAS_WIN32:
            raise RuntimeError(
                "Le support Word nécessite pywin32 et Microsoft Word installé sur Windows."
            )

        word = None
        doc = None
        try:
            word = win32com.client.DispatchEx("Word.Application")
            word.Visible = False
            word.DisplayAlerts = 0
            doc = word.Documents.Open(str(src.resolve()))
            doc.SaveAs(str(dst.resolve()), FileFormat=PDF_WD_FORMAT)
            return dst
        except Exception as e:
            raise RuntimeError(f"Conversion Word -> PDF impossible : {e}")
        finally:
            try:
                if doc is not None:
                    doc.Close(False)
            except Exception:
                pass
            try:
                if word is not None:
                    word.Quit()
            except Exception:
                pass

    @staticmethod
    def extract_eml_data(src: Path) -> dict:
        with open(src, "rb") as f:
            msg = BytesParser(policy=policy.default).parse(f)

        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                ctype = part.get_content_type()
                if ctype == "text/plain" and part.get_content_disposition() != "attachment":
                    try:
                        body = part.get_content()
                    except Exception:
                        payload = part.get_payload(decode=True) or b""
                        charset = part.get_content_charset() or "utf-8"
                        body = payload.decode(charset, errors="replace")
                    break
        else:
            try:
                body = msg.get_content()
            except Exception:
                payload = msg.get_payload(decode=True) or b""
                charset = msg.get_content_charset() or "utf-8"
                body = payload.decode(charset, errors="replace")

        attachments = []
        for part in msg.iter_attachments():
            filename = part.get_filename()
            if filename:
                attachments.append(filename)

        return {
            "subject": msg.get("subject", ""),
            "from": msg.get("from", ""),
            "to": msg.get("to", ""),
            "cc": msg.get("cc", ""),
            "date": msg.get("date", ""),
            "body": body or "",
            "attachments": attachments,
        }

    @staticmethod
    def extract_msg_data(src: Path) -> dict:
        if not HAS_WIN32:
            raise RuntimeError("Le support des fichiers .msg nécessite Outlook installé sur Windows.")

        outlook = None
        mail = None
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            mail = namespace.OpenSharedItem(str(src.resolve()))
            attachments = []
            try:
                for i in range(1, mail.Attachments.Count + 1):
                    attachments.append(str(mail.Attachments.Item(i).FileName))
            except Exception:
                pass

            return {
                "subject": str(getattr(mail, "Subject", "") or ""),
                "from": str(getattr(mail, "SenderName", "") or getattr(mail, "SenderEmailAddress", "") or ""),
                "to": str(getattr(mail, "To", "") or ""),
                "cc": str(getattr(mail, "CC", "") or ""),
                "date": format_datetime_value(getattr(mail, "SentOn", "")),
                "body": str(getattr(mail, "Body", "") or ""),
                "attachments": attachments,
            }
        except Exception as e:
            raise RuntimeError(f"Lecture du fichier .msg impossible : {e}")
        finally:
            try:
                if mail is not None:
                    mail.Close(1)
            except Exception:
                pass

    @staticmethod
    def email_data_to_pdf(data: dict, dst: Path) -> Path:
        doc = fitz.open()
        page = doc.new_page(width=A4_WIDTH, height=A4_HEIGHT)
        text_color = (0, 0, 0)
        blue = (0.11, 0.20, 0.46)
        cursor_y = 36
        left = 36
        right = A4_WIDTH - 36
        bottom = A4_HEIGHT - 36

        def new_page_if_needed(extra_height=0):
            nonlocal page, cursor_y
            if cursor_y + extra_height > bottom:
                page = doc.new_page(width=A4_WIDTH, height=A4_HEIGHT)
                cursor_y = 36

        def add_line(label: str, value: str, font_size=10, label_size=10):
            nonlocal cursor_y
            value = (value or "").strip()
            if not value:
                return
            new_page_if_needed(26)
            page.insert_text((left, cursor_y), f"{label}", fontsize=label_size, fontname="helv", color=blue)
            rect = fitz.Rect(left + 56, cursor_y - 11, right, cursor_y + 22)
            page.insert_textbox(rect, value, fontsize=font_size, fontname="helv", color=text_color)
            cursor_y += 24

        subject = (data.get("subject") or "Sans objet").strip() or "Sans objet"
        page.insert_text((left, cursor_y), subject, fontsize=15, fontname="helv", color=blue)
        cursor_y += 28

        add_line("De :", data.get("from", ""))
        add_line("À :", data.get("to", ""))
        add_line("Cc :", data.get("cc", ""))
        add_line("Date :", data.get("date", ""))

        attachments = data.get("attachments") or []
        if attachments:
            new_page_if_needed(22)
            page.insert_text((left, cursor_y), "Pièces jointes :", fontsize=10, fontname="helv", color=blue)
            cursor_y += 18
            for att in attachments:
                new_page_if_needed(16)
                rect = fitz.Rect(left + 10, cursor_y - 10, right, cursor_y + 18)
                page.insert_textbox(rect, f"• {att}", fontsize=9.5, fontname="helv", color=text_color)
                cursor_y += 16
            cursor_y += 6

        new_page_if_needed(28)
        page.insert_text((left, cursor_y), "Corps du message", fontsize=11, fontname="helv", color=blue)
        cursor_y += 18

        body = (data.get("body") or "").replace("\t", "    ").strip()
        if not body:
            body = "(aucun contenu texte exploitable)"

        paragraphs = body.splitlines() or [body]
        for para in paragraphs:
            para = para.rstrip()
            if not para:
                cursor_y += 8
                continue
            new_page_if_needed(28)
            rect = fitz.Rect(left, cursor_y - 10, right, bottom)
            chars_written = page.insert_textbox(rect, para, fontsize=10, fontname="helv", color=text_color, align=fitz.TEXT_ALIGN_LEFT)
            if chars_written < 0:
                # if paragraph does not fit, push to new page and try again
                page = doc.new_page(width=A4_WIDTH, height=A4_HEIGHT)
                cursor_y = 36
                rect = fitz.Rect(left, cursor_y - 10, right, bottom)
                page.insert_textbox(rect, para, fontsize=10, fontname="helv", color=text_color, align=fitz.TEXT_ALIGN_LEFT)
            cursor_y += max(14, min(120, ((len(para) // 80) + 1) * 12))

        doc.save(str(dst), garbage=4, deflate=True)
        doc.close()
        return dst

    @staticmethod
    def email_to_pdf(src: Path, dst: Path) -> Path:
        ext = src.suffix.lower()
        if ext == ".eml":
            data = FileOps.extract_eml_data(src)
            return FileOps.email_data_to_pdf(data, dst)
        if ext == ".msg":
            data = FileOps.extract_msg_data(src)
            return FileOps.email_data_to_pdf(data, dst)
        raise ValueError("Format email non pris en charge")

    @staticmethod
    def import_to_sandbox(src: Path, folder: Path) -> dict:
        ext = src.suffix.lower()
        if ext not in SUPPORTED_IMPORTS:
            raise ValueError(f"Format non pris en charge : {src.name}")

        if ext == ".pdf":
            dst = FileOps.unique_pdf_path(folder, src.name)
            shutil.copy2(src, dst)
            return {"kind": "pdf", "path": dst, "source_name": dst.name}

        if ext in {".jpg", ".jpeg", ".png"}:
            dst = FileOps.unique_pdf_path(folder, src.name)
            FileOps.image_to_pdf(src, dst)
            return {"kind": "pdf", "path": dst, "source_name": dst.name}

        if ext in WORD_EXTS:
            word_copy = FileOps.copy_word_to_sandbox(src, folder)
            pdf_preview = FileOps.unique_pdf_path(folder, src.stem)
            FileOps.word_to_pdf(word_copy, pdf_preview)
            return {
                "kind": "word",
                "path": word_copy,
                "preview_pdf": pdf_preview,
                "source_name": word_copy.name,
            }

        if ext in EMAIL_EXTS:
            dst = FileOps.unique_pdf_path(folder, src.stem)
            FileOps.email_to_pdf(src, dst)
            return {"kind": "pdf", "path": dst, "source_name": dst.name}

        raise ValueError(f"Format non pris en charge : {src.name}")

    @staticmethod
    def merge_pdfs(paths: list[Path], dst: Path) -> Path:
        writer = PdfWriter()
        for p in paths:
            reader = PdfReader(str(p))
            for page in reader.pages:
                writer.add_page(page)
        with open(dst, "wb") as f:
            writer.write(f)
        return dst

    @staticmethod
    def split_pdf_ranges(src: Path, ranges: list[tuple[int, int]], out_dir: Path) -> list[Path]:
        reader = PdfReader(str(src))
        total = len(reader.pages)
        out_paths = []
        for idx, (start, end) in enumerate(ranges, start=1):
            if start < 1 or end > total or start > end:
                raise ValueError(f"Plage invalide : {start}-{end}")
            writer = PdfWriter()
            for i in range(start - 1, end):
                writer.add_page(reader.pages[i])
            out_path = FileOps.unique_pdf_path(out_dir, f"{src.stem}_part_{idx}_{start}-{end}.pdf")
            with open(out_path, "wb") as f:
                writer.write(f)
            out_paths.append(out_path)
        return out_paths

    @staticmethod
    def split_pdf_every_x(src: Path, x: int, out_dir: Path) -> list[Path]:
        if x <= 0:
            raise ValueError("X doit être supérieur à 0")
        reader = PdfReader(str(src))
        total = len(reader.pages)
        ranges = []
        start = 1
        while start <= total:
            end = min(total, start + x - 1)
            ranges.append((start, end))
            start = end + 1
        return FileOps.split_pdf_ranges(src, ranges, out_dir)

    @staticmethod
    def _visible_rect_to_page_rect(page, rect: fitz.Rect) -> fitz.Rect:
        """Convertit un rectangle exprimé dans les coordonnées visibles de la page
        (tenant compte de la rotation) vers les coordonnées réelles de dessin du PDF.
        """
        if page.rotation:
            return rect * page.derotation_matrix
        return rect

    @staticmethod
    def _normalize_page_rotations_in_doc(doc):
        for page in doc:
            try:
                if page.rotation:
                    page.remove_rotation()
            except Exception:
                pass

    @staticmethod
    def add_page_numbers(src: Path, dst: Path) -> Path:
        doc = fitz.open(str(src))
        FileOps._normalize_page_rotations_in_doc(doc)
        total = len(doc)
        for i, page in enumerate(doc, start=1):
            visible = page.rect
            box = fitz.Rect(visible.width - 150, visible.height - 34, visible.width - 18, visible.height - 12)
            box = FileOps._visible_rect_to_page_rect(page, box)
            page.insert_textbox(
                box,
                f"Page {i} / {total}",
                fontsize=9,
                fontname="helv",
                color=(0, 0, 0),
                align=fitz.TEXT_ALIGN_RIGHT,
                overlay=True,
            )
        doc.save(str(dst), garbage=4, deflate=True)
        doc.close()
        return dst

    @staticmethod
    def add_piece_stamp(src: Path, dst: Path, piece_label: str, stamp_title: str = "Michelon Avocat") -> Path:
        doc = fitz.open(str(src))
        FileOps._normalize_page_rotations_in_doc(doc)
        title = stamp_title.strip() or "Michelon Avocat"
        border_color = (0.08, 0.08, 0.08)
        text_color = (0.11, 0.20, 0.46)

        for page in doc:
            visible = page.rect
            box_w = 150
            box_h = 64
            margin_top = 18
            margin_right = 18
            x0 = visible.width - margin_right - box_w
            y0 = margin_top
            x1 = visible.width - margin_right
            y1 = margin_top + box_h

            outer = FileOps._visible_rect_to_page_rect(page, fitz.Rect(x0, y0, x1, y1))
            inner = FileOps._visible_rect_to_page_rect(page, fitz.Rect(x0 + 2, y0 + 2, x1 - 2, y1 - 2))
            page.draw_rect(outer, color=border_color, width=0.9, overlay=True)
            page.draw_rect(inner, color=border_color, width=0.4, overlay=True)

            title_box = FileOps._visible_rect_to_page_rect(page, fitz.Rect(x0 + 8, y0 + 8, x1 - 8, y0 + 30))
            piece_box = FileOps._visible_rect_to_page_rect(page, fitz.Rect(x0 + 8, y0 + 28, x1 - 8, y1 - 8))

            page.insert_textbox(
                title_box,
                title,
                fontsize=12,
                fontname="Times-Bold",
                color=text_color,
                align=fitz.TEXT_ALIGN_CENTER,
                overlay=True,
            )
            page.insert_textbox(
                piece_box,
                f"Pièce n° : {piece_label}",
                fontsize=11,
                fontname="Times-Roman",
                color=text_color,
                align=fitz.TEXT_ALIGN_CENTER,
                overlay=True,
            )

        doc.save(str(dst), garbage=4, deflate=True)
        doc.close()
        return dst

    @staticmethod
    def rotate_pdf(src: Path, dst: Path, angle: int) -> Path:
        if angle % 90 != 0:
            raise ValueError("La rotation doit être un multiple de 90°")

        norm = angle % 360
        if norm == 0:
            shutil.copy2(src, dst)
            return dst

        in_doc = fitz.open(str(src))
        out_doc = fitz.open()

        try:
            for page_index in range(len(in_doc)):
                page = in_doc.load_page(page_index)
                rect = page.rect

                if norm in (90, 270):
                    new_width = rect.height
                    new_height = rect.width
                else:
                    new_width = rect.width
                    new_height = rect.height

                new_page = out_doc.new_page(width=new_width, height=new_height)
                new_page.show_pdf_page(
                    fitz.Rect(0, 0, new_width, new_height),
                    in_doc,
                    page_index,
                    rotate=norm,
                )

            out_doc.save(str(dst), garbage=4, deflate=True)
            return dst
        finally:
            out_doc.close()
            in_doc.close()

    @staticmethod
    def apply_rect_masks(src: Path, dst: Path, masks_by_page: dict[int, list[tuple[float, float, float, float, str]]]) -> Path:
        doc = fitz.open(str(src))
        for page_index, masks in masks_by_page.items():
            if not (0 <= page_index < len(doc)):
                continue
            page = doc.load_page(page_index)
            rect = page.rect
            for nx0, ny0, nx1, ny1, color_hex in masks:
                x0 = rect.x0 + nx0 * rect.width
                y0 = rect.y0 + ny0 * rect.height
                x1 = rect.x0 + nx1 * rect.width
                y1 = rect.y0 + ny1 * rect.height
                color = QColor(color_hex)
                rgb = (color.redF(), color.greenF(), color.blueF())
                mask_rect = fitz.Rect(x0, y0, x1, y1)
                page.draw_rect(mask_rect, color=rgb, fill=rgb, width=0, overlay=True)
        doc.save(str(dst), garbage=4, deflate=True)
        doc.close()
        return dst


class UndoManager:
    def __init__(self, root: Path, max_states: int = 20):
        self.root = root
        self.max_states = max_states
        self.stack: list[dict[str, str]] = []
        self.root.mkdir(parents=True, exist_ok=True)

    def clear(self):
        if self.root.exists():
            shutil.rmtree(self.root, ignore_errors=True)
        self.root.mkdir(parents=True, exist_ok=True)
        self.stack.clear()

    def can_undo(self) -> bool:
        return bool(self.stack)

    def push_snapshot(self, workdir: Path, manifest: dict, label: str):
        snap_dir = self.root / f"snap_{uuid.uuid4().hex[:10]}"
        files_dir = snap_dir / "files"
        files_dir.mkdir(parents=True, exist_ok=True)

        if workdir.exists():
            for child in workdir.iterdir():
                target = files_dir / child.name
                if child.is_dir():
                    shutil.copytree(child, target)
                else:
                    shutil.copy2(child, target)

        with open(snap_dir / "project.json", "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)

        self.stack.append({"dir": str(snap_dir), "label": label})
        while len(self.stack) > self.max_states:
            old = self.stack.pop(0)
            shutil.rmtree(old["dir"], ignore_errors=True)

    def pop_snapshot(self):
        if not self.stack:
            return None, None
        data = self.stack.pop()
        return Path(data["dir"]), data["label"]


class DocListWidget(QListWidget):
    files_dropped = Signal(list)
    order_changed = Signal()
    drag_started = Signal()

    def __init__(self):
        super().__init__()
        self._grid_w = 190
        self._grid_h = 150

        self.setViewMode(QListView.IconMode)
        self.setFlow(QListView.LeftToRight)
        self.setMovement(QListView.Snap)
        self.setResizeMode(QListView.Adjust)
        self.setWrapping(True)
        self.setUniformItemSizes(True)
        self.setSpacing(12)
        self.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.setDragEnabled(True)
        self.viewport().setAcceptDrops(True)
        self.setAcceptDrops(True)
        self.setDropIndicatorShown(True)
        self.setDragDropMode(QAbstractItemView.InternalMove)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setWordWrap(True)
        self.setTextElideMode(Qt.ElideMiddle)
        self.setContextMenuPolicy(Qt.CustomContextMenu)
        self.setLayoutMode(QListView.SinglePass)
        self.setBatchSize(100)
        self.apply_icon_layout(84)

    def apply_icon_layout(self, size: int):
        icon_size = QSize(size, size)
        self._grid_w = max(190, size + 105)
        self._grid_h = max(150, size + 80)
        self.setIconSize(icon_size)
        self.setGridSize(QSize(self._grid_w, self._grid_h))
        for i in range(self.count()):
            item = self.item(i)
            item.setSizeHint(QSize(self._grid_w, self._grid_h))
        self.refresh_grid()

    def refresh_grid(self):
        self.setUpdatesEnabled(False)
        try:
            self.doItemsLayout()
            self.viewport().update()
        finally:
            self.setUpdatesEnabled(True)

    def mimeData(self, items):
        mime = super().mimeData(items)
        urls = []
        for item in items:
            path = item.data(Qt.UserRole)
            if path and Path(path).exists():
                urls.append(QUrl.fromLocalFile(str(path)))
        if urls:
            mime.setUrls(urls)
        return mime

    def startDrag(self, supportedActions):
        self.drag_started.emit()
        super().startDrag(supportedActions)
        self.refresh_grid()
        self.order_changed.emit()

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls() or event.source() == self:
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls() or event.source() == self:
            event.acceptProposedAction()
        else:
            super().dragMoveEvent(event)

    def dropEvent(self, event: QDropEvent):
        if event.source() == self:
            super().dropEvent(event)
            self.refresh_grid()
            self.order_changed.emit()
            return

        urls = [u.toLocalFile() for u in event.mimeData().urls() if u.isLocalFile()]
        if urls:
            self.files_dropped.emit(urls)
            event.acceptProposedAction()
            self.refresh_grid()
            return

        super().dropEvent(event)
        self.refresh_grid()

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.refresh_grid()



class PreviewImageLabel(QLabel):
    def __init__(self, pane):
        super().__init__("Aperçu")
        self.pane = pane
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("background:#f5f5f5;")
        self.setMinimumWidth(400)
        self.setMouseTracking(True)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self.pane.current_path and self.pane.rotation == 0:
            point = self.pane.clamp_point_to_pixmap(event.position().toPoint())
            if point is not None:
                if self.pane.mask_mode_color:
                    self.pane.drag_start = point
                    self.pane.drag_current = point
                    self.pane.render_current_page()
                    event.accept()
                    return
                self.pane.select_mask_at(point)
                event.accept()
                return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.pane.drag_start is not None:
            point = self.pane.clamp_point_to_pixmap(event.position().toPoint())
            if point is not None:
                self.pane.drag_current = point
                self.pane.render_current_page()
                event.accept()
                return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton and self.pane.drag_start is not None:
            point = self.pane.clamp_point_to_pixmap(event.position().toPoint())
            if point is not None:
                self.pane.drag_current = point
                self.pane.commit_current_mask()
                event.accept()
                return
        super().mouseReleaseEvent(event)


class PreviewPane(QWidget):
    def __init__(self):
        super().__init__()
        self.current_path = None
        self.current_page = 0
        self.page_count = 0
        self.zoom = 1.2
        self.rotation = 0

        self.mask_mode_color: QColor | None = None
        self.drag_start: QPoint | None = None
        self.drag_current: QPoint | None = None
        self.masks_by_document: dict[str, dict[int, list[tuple[float, float, float, float, str]]]] = {}
        self.mask_undo_stack: list[dict[str, dict[int, list[tuple[float, float, float, float, str]]]]] = []
        self.max_mask_history = 50
        self.selected_mask_ref: tuple[str, int, int] | None = None

        layout = QVBoxLayout(self)
        controls = QHBoxLayout()

        self.prev_btn = QPushButton("◀")
        self.next_btn = QPushButton("▶")
        self.rot_left_btn = QPushButton("⟲")
        self.rot_right_btn = QPushButton("⟳")
        self.page_label = QLabel("Aucun document")
        self.zoom_in_btn = QPushButton("+")
        self.zoom_out_btn = QPushButton("-")

        for btn in (self.prev_btn, self.next_btn, self.rot_left_btn, self.rot_right_btn, self.zoom_out_btn, self.zoom_in_btn):
            btn.setFixedWidth(32)

        controls.addWidget(self.prev_btn)
        controls.addWidget(self.next_btn)
        controls.addWidget(self.rot_left_btn)
        controls.addWidget(self.rot_right_btn)
        controls.addWidget(self.page_label)
        controls.addStretch(1)
        controls.addWidget(self.zoom_out_btn)
        controls.addWidget(self.zoom_in_btn)

        self.image_label = PreviewImageLabel(self)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidget(self.image_label)
        self.scroll_area.setWidgetResizable(False)
        self.scroll_area.setAlignment(Qt.AlignCenter)
        self.scroll_area.setStyleSheet("background:#f5f5f5;border:1px solid #d0d0d0;")

        layout.addLayout(controls)
        layout.addWidget(self.scroll_area, 1)

        self.prev_btn.clicked.connect(self.prev_page)
        self.next_btn.clicked.connect(self.next_page)
        self.rot_left_btn.clicked.connect(self.rotate_left)
        self.rot_right_btn.clicked.connect(self.rotate_right)
        self.zoom_in_btn.clicked.connect(self.zoom_in)
        self.zoom_out_btn.clicked.connect(self.zoom_out)

    def clear(self):
        self.current_path = None
        self.current_page = 0
        self.page_count = 0
        self.rotation = 0
        self.drag_start = None
        self.drag_current = None
        self.selected_mask_ref = None
        self.page_label.setText("Aucun document")
        self.image_label.setText("Aperçu")
        self.image_label.setPixmap(QPixmap())
        self.image_label.adjustSize()

    def clear_mask_history(self):
        self.mask_undo_stack.clear()

    def can_undo_mask_state(self):
        return bool(self.mask_undo_stack)

    def snapshot_mask_state(self):
        self.mask_undo_stack.append(copy.deepcopy(self.masks_by_document))
        while len(self.mask_undo_stack) > self.max_mask_history:
            self.mask_undo_stack.pop(0)

    def undo_last_mask_state(self):
        if not self.mask_undo_stack:
            return False
        self.masks_by_document = self.mask_undo_stack.pop()
        self.selected_mask_ref = None
        self.drag_start = None
        self.drag_current = None
        self.disable_mask_mode()
        try:
            mw = self.window()
            if hasattr(mw, "update_undo_action_state"):
                mw.update_undo_action_state()
        except Exception:
            pass
        return True

    def show_pdf(self, path: str):
        self.current_path = path
        self.rotation = 0
        self.drag_start = None
        self.drag_current = None
        self.selected_mask_ref = None
        try:
            doc = fitz.open(path)
            self.page_count = len(doc)
            doc.close()
            self.current_page = 0
            self.render_current_page()
        except Exception as e:
            self.page_label.setText("Erreur aperçu")
            self.image_label.setText(str(e))

    def current_pixmap_rect(self) -> QRect:
        pm = self.image_label.pixmap()
        if pm is None or pm.isNull():
            return QRect()
        return QRect(0, 0, pm.width(), pm.height())

    def clamp_point_to_pixmap(self, point: QPoint):
        rect = self.current_pixmap_rect()
        if rect.isNull():
            return None
        x = min(max(point.x(), rect.left()), rect.right())
        y = min(max(point.y(), rect.top()), rect.bottom())
        return QPoint(x, y)

    def set_mask_mode(self, color: QColor):
        self.mask_mode_color = color
        self.drag_start = None
        self.drag_current = None
        self.selected_mask_ref = None
        self.render_current_page()

    def disable_mask_mode(self):
        self.mask_mode_color = None
        self.drag_start = None
        self.drag_current = None
        self.render_current_page()

    def get_masks_for_document(self, path: str):
        return copy.deepcopy(self.masks_by_document.get(path, {}))

    def clear_masks_for_document_no_history(self, path: str):
        if path in self.masks_by_document:
            self.masks_by_document[path] = {}
        if self.selected_mask_ref and self.selected_mask_ref[0] == path:
            self.selected_mask_ref = None

    def clear_masks_for_current_document(self):
        if not self.current_path:
            return
        doc_masks = self.masks_by_document.get(self.current_path, {})
        has_masks = any(doc_masks.get(p) for p in doc_masks)
        if not has_masks:
            self.drag_start = None
            self.drag_current = None
            self.selected_mask_ref = None
            self.render_current_page()
            return
        self.snapshot_mask_state()
        self.clear_masks_for_document_no_history(self.current_path)
        self.drag_start = None
        self.drag_current = None
        self.disable_mask_mode()

    def remove_document_masks(self, path: str):
        self.masks_by_document.pop(path, None)
        if self.selected_mask_ref and self.selected_mask_ref[0] == path:
            self.selected_mask_ref = None

    def move_document_masks(self, old_path: str, new_path: str):
        if old_path == new_path:
            return
        if old_path in self.masks_by_document:
            self.masks_by_document[new_path] = self.masks_by_document.pop(old_path)
        if self.selected_mask_ref and self.selected_mask_ref[0] == old_path:
            self.selected_mask_ref = (new_path, self.selected_mask_ref[1], self.selected_mask_ref[2])

    def commit_current_mask(self):
        if not (self.current_path and self.mask_mode_color and self.drag_start and self.drag_current):
            return
        pix_rect = self.current_pixmap_rect()
        if pix_rect.isNull():
            return
        x0 = min(self.drag_start.x(), self.drag_current.x()) - pix_rect.x()
        y0 = min(self.drag_start.y(), self.drag_current.y()) - pix_rect.y()
        x1 = max(self.drag_start.x(), self.drag_current.x()) - pix_rect.x()
        y1 = max(self.drag_start.y(), self.drag_current.y()) - pix_rect.y()
        self.drag_start = None
        self.drag_current = None
        if x1 - x0 < 2 or y1 - y0 < 2:
            self.render_current_page()
            return
        norm = (
            max(0.0, min(1.0, x0 / pix_rect.width())),
            max(0.0, min(1.0, y0 / pix_rect.height())),
            max(0.0, min(1.0, x1 / pix_rect.width())),
            max(0.0, min(1.0, y1 / pix_rect.height())),
        )
        self.snapshot_mask_state()
        doc_masks = self.masks_by_document.setdefault(self.current_path, {})
        page_masks = doc_masks.setdefault(self.current_page, [])
        page_masks.append((*norm, self.mask_mode_color.name()))
        self.selected_mask_ref = (self.current_path, self.current_page, len(page_masks) - 1)
        self.render_current_page()
        try:
            mw = self.window()
            if hasattr(mw, "update_undo_action_state"):
                mw.update_undo_action_state()
        except Exception:
            pass

    def _mask_rect_on_pixmap(self, mask_tuple, painted: QPixmap):
        nx0, ny0, nx1, ny1, _ = mask_tuple
        x = int(nx0 * painted.width())
        y = int(ny0 * painted.height())
        w = max(1, int((nx1 - nx0) * painted.width()))
        h = max(1, int((ny1 - ny0) * painted.height()))
        return QRect(x, y, w, h)

    def select_mask_at(self, point: QPoint):
        if not self.current_path or self.rotation != 0:
            self.selected_mask_ref = None
            self.render_current_page()
            return
        pix_rect = self.current_pixmap_rect()
        if pix_rect.isNull():
            return
        local = QPoint(point.x() - pix_rect.x(), point.y() - pix_rect.y())
        pm = self.image_label.pixmap()
        masks = self.masks_by_document.get(self.current_path, {}).get(self.current_page, [])
        found = None
        for idx in range(len(masks) - 1, -1, -1):
            rect = self._mask_rect_on_pixmap(masks[idx], pm)
            if rect.contains(local):
                found = (self.current_path, self.current_page, idx)
                break
        self.selected_mask_ref = found
        self.render_current_page()

    def delete_selected_mask(self):
        if not self.selected_mask_ref:
            return False
        doc_path, page_idx, mask_idx = self.selected_mask_ref
        page_masks = self.masks_by_document.get(doc_path, {}).get(page_idx, [])
        if not (0 <= mask_idx < len(page_masks)):
            self.selected_mask_ref = None
            self.render_current_page()
            return False
        self.snapshot_mask_state()
        del page_masks[mask_idx]
        if not page_masks:
            self.masks_by_document.get(doc_path, {}).pop(page_idx, None)
        self.selected_mask_ref = None
        self.render_current_page()
        try:
            mw = self.window()
            if hasattr(mw, "update_undo_action_state"):
                mw.update_undo_action_state()
        except Exception:
            pass
        return True

    def render_current_page(self):
        if not self.current_path:
            return
        doc = fitz.open(self.current_path)
        if not (0 <= self.current_page < len(doc)):
            doc.close()
            return
        page = doc.load_page(self.current_page)
        matrix = fitz.Matrix(self.zoom, self.zoom).prerotate(self.rotation)
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
        doc.close()

        base = QPixmap.fromImage(qimg)
        painted = QPixmap(base)
        painter = QPainter(painted)

        if self.current_path:
            masks = self.masks_by_document.get(self.current_path, {}).get(self.current_page, [])
            for idx, mask in enumerate(masks):
                rect = self._mask_rect_on_pixmap(mask, painted)
                color = QColor(mask[4])
                painter.fillRect(rect, QBrush(color))
                pen = QPen(color)
                pen.setWidth(1)
                painter.setPen(pen)
                painter.drawRect(rect)
                if self.selected_mask_ref == (self.current_path, self.current_page, idx):
                    sel_pen = QPen(QColor("yellow"))
                    sel_pen.setWidth(2)
                    sel_pen.setStyle(Qt.DashLine)
                    painter.setPen(sel_pen)
                    painter.drawRect(rect.adjusted(-2, -2, 2, 2))

        if self.drag_start is not None and self.drag_current is not None:
            pix_rect = self.current_pixmap_rect()
            p1 = QPoint(self.drag_start.x() - pix_rect.x(), self.drag_start.y() - pix_rect.y())
            p2 = QPoint(self.drag_current.x() - pix_rect.x(), self.drag_current.y() - pix_rect.y())
            temp_rect = QRect(p1, p2).normalized()
            color = self.mask_mode_color if self.mask_mode_color else QColor("red")
            fill = QColor(color)
            fill.setAlpha(160)
            painter.fillRect(temp_rect, QBrush(fill))
            pen = QPen(color)
            pen.setStyle(Qt.DashLine)
            pen.setWidth(2)
            painter.setPen(pen)
            painter.drawRect(temp_rect)

        painter.end()
        self.image_label.setPixmap(painted)
        self.image_label.resize(painted.size())
        self.image_label.adjustSize()
        mode_txt = ""
        if self.mask_mode_color:
            mode_txt = f" | Masquage actif : {self.mask_mode_color.name()}"
        elif self.selected_mask_ref and self.selected_mask_ref[0] == self.current_path and self.selected_mask_ref[1] == self.current_page:
            mode_txt = " | 1 masque sélectionné"
        self.page_label.setText(f"Page {self.current_page + 1} / {self.page_count} | Rotation aperçu : {self.rotation}°{mode_txt}")

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self.drag_start = None
            self.drag_current = None
            self.selected_mask_ref = None
            self.render_current_page()

    def next_page(self):
        if self.current_page + 1 < self.page_count:
            self.current_page += 1
            self.drag_start = None
            self.drag_current = None
            self.selected_mask_ref = None
            self.render_current_page()

    def zoom_in(self):
        self.zoom = min(3.0, self.zoom + 0.1)
        self.render_current_page()

    def zoom_out(self):
        self.zoom = max(0.5, self.zoom - 0.1)
        self.render_current_page()

    def rotate_right(self):
        self.rotation = (self.rotation + 90) % 360
        self.selected_mask_ref = None
        self.render_current_page()

    def rotate_left(self):
        self.rotation = (self.rotation - 90) % 360
        self.selected_mask_ref = None
        self.render_current_page()

class BordereauRenameDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Nommer les pièces selon bordereau")
        self.resize(700, 420)

        layout = QVBoxLayout(self)

        info = QLabel(
            "Colle ici ton bordereau, une ligne par pièce.\n"
            "Exemple :\n"
            "1. Contrat de travail\n"
            "2. Bulletins de salaire\n"
            "3. Arrêts de travail"
        )
        layout.addWidget(info)

        self.text_edit = QPlainTextEdit()
        self.text_edit.setPlaceholderText(
            "1. Contrat de travail\n"
            "2. Bulletins de salaire\n"
            "3. Arrêts de travail\n"
            "4. Lettre du 19 décembre 2023"
        )
        layout.addWidget(self.text_edit, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_lines(self):
        raw = self.text_edit.toPlainText()
        lines = []
        for line in raw.splitlines():
            line = line.strip()
            if not line:
                continue
            if re.fullmatch(r"\d+", line):
                continue
            lines.append(line)
        return lines


class PieceLabelDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Cachet des pièces - numérotation avancée")
        self.resize(700, 420)

        layout = QVBoxLayout(self)

        info = QLabel(
            "Saisis une ligne par document sélectionné.\n"
            "Exemples :\n"
            "1\n"
            "2\n"
            "3\n"
            "4.a\n"
            "4.b\n"
            "5"
        )
        layout.addWidget(info)

        self.text_edit = QPlainTextEdit()
        self.text_edit.setPlaceholderText(
            "1\n"
            "2\n"
            "3\n"
            "4.a\n"
            "4.b\n"
            "4.c\n"
            "5"
        )
        layout.addWidget(self.text_edit, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_labels(self):
        raw = self.text_edit.toPlainText()
        labels = []
        for line in raw.splitlines():
            line = line.strip()
            if not line:
                continue
            labels.append(line)
        return labels


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(1440, 900)
        self.workdir = session_root()
        self.project_dir: Path | None = None
        self.settings = QSettings("Michelon", "PDFSandbox")
        self.stamp_title = self.settings.value("stamp_title", "Michelon Avocat", type=str)
        self.undo_manager = UndoManager(sandbox_root() / f"undo_{uuid.uuid4().hex[:8]}")

        self.list_widget = DocListWidget()
        self.preview = PreviewPane()

        splitter = QSplitter()
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(8, 8, 8, 8)

        top_controls = QHBoxLayout()
        import_btn = QPushButton("Importer")
        undo_btn = QPushButton("Annuler")
        save_project_btn = QPushButton("Enregistrer")
        open_project_btn = QPushButton("Ouvrir")
        open_sandbox_btn = QPushButton("Sandbox")
        top_controls.addWidget(import_btn)
        top_controls.addWidget(undo_btn)
        top_controls.addWidget(save_project_btn)
        top_controls.addWidget(open_project_btn)
        top_controls.addWidget(open_sandbox_btn)
        top_controls.addStretch(1)
        top_controls.addWidget(QLabel("Taille icônes"))
        self.icon_slider = QSlider(Qt.Horizontal)
        self.icon_slider.setRange(56, 140)
        self.icon_slider.setValue(84)
        self.icon_slider.setFixedWidth(180)
        top_controls.addWidget(self.icon_slider)

        left_layout.addLayout(top_controls)
        left_layout.addWidget(self.list_widget, 1)

        splitter.addWidget(left_panel)
        splitter.addWidget(self.preview)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 4)
        self.setCentralWidget(splitter)

        toolbar = QToolBar("Actions")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        self.act_import = QAction("Importer", self)
        self.act_undo = QAction("Annuler", self)
        self.act_merge = QAction("Fusionner", self)
        self.act_split = QAction("Diviser", self)
        self.act_number = QAction("Ajouter numérotation", self)
        self.act_rotate_right = QAction("Rotation 90° droite", self)
        self.act_rotate_left = QAction("Rotation 90° gauche", self)
        self.act_stamp_quick = QAction("Cachet rapide", self)
        self.act_stamp_advanced = QAction("Cachet avancé", self)
        self.act_rename = QAction("Renommer", self)
        self.act_stamp_text = QAction("Texte du cachet", self)
        self.act_delete = QAction("Supprimer", self)
        self.act_bordereau = QAction("Nommer selon bordereau", self)
        self.act_save_project = QAction("Enregistrer dossier", self)
        self.act_open_project = QAction("Ouvrir dossier", self)
        self.act_open_sandbox = QAction("Ouvrir la sandbox", self)
        self.act_mask_black = QAction("Masquage noir", self)
        self.act_mask_white = QAction("Masquage blanc", self)
        self.act_mask_color = QAction("Masquage couleur…", self)
        self.act_apply_masks = QAction("Appliquer les masques", self)
        self.act_clear_masks = QAction("Effacer tous les masques du document", self)
        self.act_delete_mask = QAction("Supprimer le masque sélectionné", self)

        self.act_import.setShortcut("Ctrl+I")
        self.act_undo.setShortcut("Ctrl+Z")
        self.act_merge.setShortcut("Ctrl+F")
        self.act_split.setShortcut("Ctrl+D")
        self.act_number.setShortcut("Ctrl+N")
        self.act_rotate_right.setShortcut("Ctrl+Alt+Right")
        self.act_rotate_left.setShortcut("Ctrl+Alt+Left")
        self.act_stamp_quick.setShortcut("Ctrl+T")
        self.act_stamp_advanced.setShortcut("Ctrl+Shift+T")
        self.act_rename.setShortcut("Ctrl+R")
        self.act_bordereau.setShortcut("Ctrl+Shift+R")
        self.act_delete.setShortcut("Delete")
        self.act_save_project.setShortcut("Ctrl+S")
        self.act_open_project.setShortcut("Ctrl+O")
        self.act_mask_black.setShortcut("Ctrl+E")
        self.act_mask_white.setShortcut("Ctrl+Shift+E")
        self.act_apply_masks.setShortcut("Ctrl+M")
        self.act_clear_masks.setShortcut("Ctrl+Shift+M")
        self.act_delete_mask.setShortcut("Backspace")

        for act in (
            self.act_import, self.act_undo, self.act_merge, self.act_split, self.act_rename, self.act_bordereau,
            self.act_number, self.act_rotate_left, self.act_rotate_right, self.act_stamp_quick,
            self.act_stamp_advanced, self.act_mask_black, self.act_mask_white, self.act_apply_masks, self.act_delete_mask,
            self.act_save_project, self.act_open_project, self.act_delete,
        ):
            shortcut_text = act.shortcut().toString()
            if shortcut_text:
                act.setToolTip(f"{act.text()} ({shortcut_text})")
                act.setStatusTip(f"{act.text()} ({shortcut_text})")

        for act in (self.act_import, self.act_undo, self.act_merge, self.act_split, self.act_rename, self.act_bordereau):
            toolbar.addAction(act)
        toolbar.addSeparator()
        for act in (self.act_number, self.act_rotate_left, self.act_rotate_right, self.act_stamp_quick):
            toolbar.addAction(act)
        toolbar.addSeparator()
        for act in (self.act_mask_black, self.act_mask_white, self.act_apply_masks):
            toolbar.addAction(act)
        toolbar.addSeparator()
        for act in (self.act_save_project, self.act_open_project, self.act_delete):
            toolbar.addAction(act)

        menubar = self.menuBar()
        menu_file = menubar.addMenu("&Fichier")
        menu_file.addAction(self.act_import)
        menu_file.addSeparator()
        menu_file.addAction(self.act_open_project)
        menu_file.addAction(self.act_save_project)
        menu_file.addAction(self.act_open_sandbox)

        menu_edit = menubar.addMenu("&Édition")
        menu_edit.addAction(self.act_undo)
        menu_edit.addSeparator()
        menu_edit.addAction(self.act_rename)
        menu_edit.addAction(self.act_bordereau)
        menu_edit.addSeparator()
        menu_edit.addAction(self.act_delete)

        menu_pdf = menubar.addMenu("&PDF")
        menu_pdf.addAction(self.act_merge)
        menu_pdf.addAction(self.act_split)
        menu_pdf.addAction(self.act_number)
        menu_pdf.addSeparator()
        menu_pdf.addAction(self.act_rotate_left)
        menu_pdf.addAction(self.act_rotate_right)

        menu_piece = menubar.addMenu("&Pièces")
        menu_piece.addAction(self.act_stamp_quick)
        menu_piece.addAction(self.act_stamp_advanced)
        menu_piece.addAction(self.act_stamp_text)

        menu_mask = menubar.addMenu("&Masquage")
        menu_mask.addAction(self.act_mask_black)
        menu_mask.addAction(self.act_mask_white)
        menu_mask.addAction(self.act_mask_color)
        menu_mask.addSeparator()
        menu_mask.addAction(self.act_apply_masks)
        menu_mask.addAction(self.act_delete_mask)
        menu_mask.addAction(self.act_clear_masks)

        menu_view = menubar.addMenu("&Affichage")
        menu_view.addAction(self.act_open_sandbox)

        self.list_widget.files_dropped.connect(self.import_files)
        self.list_widget.drag_started.connect(lambda: self.snapshot_state("Réorganisation"))
        self.list_widget.order_changed.connect(self.refresh_preview)
        self.list_widget.customContextMenuRequested.connect(self.open_context_menu)
        self.list_widget.itemSelectionChanged.connect(self.refresh_preview)
        self.icon_slider.valueChanged.connect(self.update_icon_size)

        import_btn.clicked.connect(self.pick_files)
        undo_btn.clicked.connect(self.undo_last_action)
        save_project_btn.clicked.connect(self.save_project)
        open_project_btn.clicked.connect(self.open_project)
        open_sandbox_btn.clicked.connect(self.open_sandbox)

        self.act_import.triggered.connect(self.pick_files)
        self.act_undo.triggered.connect(self.undo_last_action)
        self.act_merge.triggered.connect(self.merge_selected)
        self.act_split.triggered.connect(self.split_selected)
        self.act_number.triggered.connect(self.number_selected)
        self.act_rotate_right.triggered.connect(lambda: self.rotate_selected_documents(90))
        self.act_rotate_left.triggered.connect(lambda: self.rotate_selected_documents(-90))
        self.act_stamp_quick.triggered.connect(self.stamp_selected_quick)
        self.act_stamp_advanced.triggered.connect(self.stamp_selected_advanced)
        self.act_rename.triggered.connect(self.rename_selected)
        self.act_stamp_text.triggered.connect(self.change_stamp_text)
        self.act_bordereau.triggered.connect(self.rename_selected_from_bordereau)
        self.act_delete.triggered.connect(self.delete_selected)
        self.act_save_project.triggered.connect(self.save_project)
        self.act_open_project.triggered.connect(self.open_project)
        self.act_open_sandbox.triggered.connect(self.open_sandbox)
        self.act_mask_black.triggered.connect(lambda: self.start_mask_mode(QColor("black")))
        self.act_mask_white.triggered.connect(lambda: self.start_mask_mode(QColor("white")))
        self.act_mask_color.triggered.connect(self.start_custom_mask_mode)
        self.act_apply_masks.triggered.connect(self.apply_masks_current_document)
        self.act_delete_mask.triggered.connect(self.delete_selected_mask)
        self.act_clear_masks.triggered.connect(self.clear_masks_current_document)

        QShortcut(QKeySequence(Qt.Key_Left), self, activated=self.preview.prev_page)
        QShortcut(QKeySequence(Qt.Key_Right), self, activated=self.preview.next_page)
        QShortcut(QKeySequence("Ctrl+Left"), self, activated=self.select_previous_document)
        QShortcut(QKeySequence("Ctrl+Right"), self, activated=self.select_next_document)
        QShortcut(QKeySequence("R"), self, activated=self.preview.rotate_right)
        QShortcut(QKeySequence("Shift+R"), self, activated=self.preview.rotate_left)

        self.statusBar().showMessage(f"Sandbox : {self.workdir} | Texte du cachet : {self.stamp_title}")
        self.update_undo_action_state()


    def update_undo_action_state(self):
        enabled = self.undo_manager.can_undo() or self.preview.can_undo_mask_state()
        self.act_undo.setEnabled(enabled)

    def snapshot_state(self, label: str):
        try:
            manifest = self.build_project_manifest()
            self.undo_manager.push_snapshot(self.workdir, manifest, label)
            self.update_undo_action_state()
        except Exception as e:
            QMessageBox.warning(self, "Historique", f"Impossible de créer l'état d'annulation : {e}")

    def _clear_directory_contents(self, folder: Path):
        folder.mkdir(parents=True, exist_ok=True)
        for child in folder.iterdir():
            if child.is_dir():
                shutil.rmtree(child, ignore_errors=True)
            else:
                child.unlink(missing_ok=True)

    def _copy_directory_contents(self, src: Path, dst: Path):
        dst.mkdir(parents=True, exist_ok=True)
        for child in src.iterdir():
            target = dst / child.name
            if child.is_dir():
                shutil.copytree(child, target)
            else:
                shutil.copy2(child, target)

    def _load_manifest_into_view(self, manifest: dict, files_dir: Path):
        self.clear_view_only()
        self.preview.masks_by_document.clear()
        self.preview.clear_mask_history()
        self.preview.disable_mask_mode()

        self.stamp_title = manifest.get("stamp_title", "Michelon Avocat")
        self.settings.setValue("stamp_title", self.stamp_title)

        icon_size = int(manifest.get("icon_size", 84))
        self.icon_slider.blockSignals(True)
        self.icon_slider.setValue(icon_size)
        self.icon_slider.blockSignals(False)
        self.list_widget.apply_icon_layout(icon_size)

        path_by_name = {}
        for entry in manifest.get("items", []):
            real_path = files_dir / entry["real_name"]
            preview_path = files_dir / entry.get("preview_name", entry["real_name"])
            if not real_path.exists():
                continue
            if not preview_path.exists():
                preview_path = real_path
            self.add_doc_item({
                "kind": entry.get("kind", "pdf"),
                "path": real_path,
                "preview_pdf": preview_path,
                "source_name": entry.get("text", real_path.name),
            })
            path_by_name[real_path.name] = str(real_path)

        for real_name, page_map in manifest.get("mask_objects", {}).items():
            real_path_str = path_by_name.get(real_name)
            if not real_path_str:
                continue
            loaded_page_map = {}
            for page_str, masks in page_map.items():
                try:
                    page_idx = int(page_str)
                except Exception:
                    continue
                loaded_page_map[page_idx] = [tuple(mask) for mask in masks]
            if loaded_page_map:
                self.preview.masks_by_document[real_path_str] = loaded_page_map

        self.preview.clear()

    def undo_last_action(self):
        if self.preview.can_undo_mask_state():
            try:
                if self.preview.undo_last_mask_state():
                    self.statusBar().showMessage("Annulation : masquage", 4000)
            finally:
                self.update_undo_action_state()
            return

        snapshot_dir, label = self.undo_manager.pop_snapshot()
        if snapshot_dir is None:
            return
        try:
            with open(snapshot_dir / "project.json", "r", encoding="utf-8") as f:
                manifest = json.load(f)
            self._clear_directory_contents(self.workdir)
            self._copy_directory_contents(snapshot_dir / "files", self.workdir)
            self._load_manifest_into_view(manifest, self.workdir)
            self.statusBar().showMessage(f"Annulation : {label}", 5000)
        except Exception as e:
            QMessageBox.critical(self, "Annulation", f"Impossible d'annuler la dernière action : {e}")
        finally:
            try:
                shutil.rmtree(snapshot_dir, ignore_errors=True)
            except Exception:
                pass
            self.update_undo_action_state()

    def keyPressEvent(self, event):
        if isinstance(self.focusWidget(), (QPlainTextEdit, QTextEdit)):
            super().keyPressEvent(event)
            return
        super().keyPressEvent(event)

    def clear_view_only(self):
        self.list_widget.clear()
        self.preview.clear()

    def iter_items_in_order(self):
        for i in range(self.list_widget.count()):
            yield self.list_widget.item(i)

    def build_project_manifest(self):
        items_data = []
        mask_objects = {}
        for item in self.iter_items_in_order():
            real_path = Path(item.data(Qt.UserRole))
            kind = item.data(Qt.UserRole + 1)
            preview_path = Path(item.data(Qt.UserRole + 2))
            entry = {
                "text": item.text(),
                "kind": kind,
                "real_name": real_path.name,
                "preview_name": preview_path.name,
            }
            items_data.append(entry)
            if kind == "pdf":
                doc_masks = self.preview.masks_by_document.get(str(real_path), {})
                if doc_masks:
                    mask_objects[real_path.name] = {str(page_idx): [list(mask) for mask in masks] for page_idx, masks in doc_masks.items() if masks}
        return {
            "version": 2,
            "stamp_title": self.stamp_title,
            "icon_size": self.icon_slider.value(),
            "items": items_data,
            "mask_objects": mask_objects,
        }

    def save_project(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Choisir le dossier du projet",
            str(self.project_dir if self.project_dir else Path.home())
        )
        if not folder:
            return

        project_dir = Path(folder)
        tmp_files_dir = project_dir / "__files_tmp__"
        manifest_path = project_dir / "project.json"

        if tmp_files_dir.exists():
            shutil.rmtree(tmp_files_dir, ignore_errors=True)
        tmp_files_dir.mkdir(parents=True, exist_ok=True)

        for item in self.iter_items_in_order():
            real_path = Path(item.data(Qt.UserRole))
            preview_path = Path(item.data(Qt.UserRole + 2))

            new_real = tmp_files_dir / real_path.name
            if new_real.exists():
                new_real = FileOps.unique_path(tmp_files_dir, real_path.name, real_path.suffix)
            shutil.copy2(real_path, new_real)

            if preview_path.resolve() != real_path.resolve():
                new_preview = tmp_files_dir / preview_path.name
                if new_preview.exists():
                    new_preview = FileOps.unique_pdf_path(tmp_files_dir, preview_path.name)
                shutil.copy2(preview_path, new_preview)

        manifest = self.build_project_manifest()

        final_files_dir = project_dir / "files"
        if final_files_dir.exists():
            shutil.rmtree(final_files_dir, ignore_errors=True)
        tmp_files_dir.rename(final_files_dir)

        with open(manifest_path, "w", encoding="utf-8") as f:
            json.dump(manifest, f, ensure_ascii=False, indent=2)

        self.project_dir = project_dir
        self.workdir = final_files_dir
        self.statusBar().showMessage(f"Dossier enregistré : {project_dir}", 5000)

    def open_project(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Ouvrir un dossier projet",
            str(self.project_dir if self.project_dir else Path.home())
        )
        if not folder:
            return

        project_dir = Path(folder)
        manifest_path = project_dir / "project.json"
        files_dir = project_dir / "files"

        if not manifest_path.exists():
            QMessageBox.warning(self, "Ouvrir dossier", "Aucun project.json trouvé dans ce dossier.")
            return
        if not files_dir.exists():
            QMessageBox.warning(self, "Ouvrir dossier", "Aucun sous-dossier 'files' trouvé dans ce dossier.")
            return

        try:
            with open(manifest_path, "r", encoding="utf-8") as f:
                manifest = json.load(f)
        except Exception as e:
            QMessageBox.critical(self, "Ouvrir dossier", f"Lecture du projet impossible : {e}")
            return

        self.clear_view_only()
        self.project_dir = project_dir
        self.workdir = files_dir
        self.stamp_title = manifest.get("stamp_title", "Michelon Avocat")
        self.settings.setValue("stamp_title", self.stamp_title)
        icon_size = int(manifest.get("icon_size", 84))
        self.icon_slider.setValue(icon_size)

        for entry in manifest.get("items", []):
            real_path = files_dir / entry["real_name"]
            preview_path = files_dir / entry.get("preview_name", entry["real_name"])
            if not real_path.exists():
                continue
            if not preview_path.exists():
                preview_path = real_path
            self.add_doc_item({
                "kind": entry.get("kind", "pdf"),
                "path": real_path,
                "preview_pdf": preview_path,
                "source_name": entry.get("text", real_path.name),
            })

        self.preview.clear()
        self.preview.clear_mask_history()
        self.undo_manager.clear()
        self.update_undo_action_state()
        self.statusBar().showMessage(f"Dossier ouvert : {project_dir}", 5000)

    def update_icon_size(self, size: int):
        self.list_widget.apply_icon_layout(size)

    def pick_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Importer des documents",
            str(Path.home()),
            "Documents (*.pdf *.jpg *.jpeg *.png *.doc *.docx *.eml *.msg)",
        )
        if files:
            self.import_files(files)

    def import_files(self, files: list[str]):
        if files:
            self.snapshot_state("Import")
        added = 0
        errors = []
        for f in files:
            try:
                src = Path(f)
                if not src.exists():
                    continue
                entry = FileOps.import_to_sandbox(src, self.workdir)
                self.add_doc_item(entry)
                added += 1
            except Exception as e:
                errors.append(f"{Path(f).name}: {e}")

        if added:
            self.statusBar().showMessage(
                f"{added} document(s) importé(s) | Texte du cachet : {self.stamp_title}",
                5000,
            )
        if errors:
            QMessageBox.warning(self, "Import partiel", "\n".join(errors))

    def add_doc_item(self, entry: dict, row: int | None = None):
        display_path = Path(entry["path"])
        preview_path = Path(entry.get("preview_pdf", entry["path"]))
        item = QListWidgetItem(entry.get("source_name", display_path.name))
        item.setData(Qt.UserRole, str(display_path))
        item.setData(Qt.UserRole + 1, entry.get("kind", "pdf"))
        item.setData(Qt.UserRole + 2, str(preview_path))
        item.setToolTip(str(display_path))
        item.setIcon(self.make_preview_icon(preview_path, self.list_widget.iconSize()))
        item.setSizeHint(self.list_widget.gridSize())
        if row is None:
            self.list_widget.addItem(item)
        else:
            self.list_widget.insertItem(row, item)
        self.list_widget.refresh_grid()
        return item

    def replace_item_with_pdf(self, item: QListWidgetItem, pdf_path: Path):
        old_path = str(item.data(Qt.UserRole))
        item.setText(pdf_path.name)
        item.setData(Qt.UserRole, str(pdf_path))
        item.setData(Qt.UserRole + 1, "pdf")
        item.setData(Qt.UserRole + 2, str(pdf_path))
        item.setToolTip(str(pdf_path))
        item.setIcon(self.make_preview_icon(pdf_path, self.list_widget.iconSize()))
        item.setSizeHint(self.list_widget.gridSize())
        self.preview.move_document_masks(old_path, str(pdf_path))
        self.list_widget.refresh_grid()

    def remove_items_and_files(self, items: list[QListWidgetItem]):
        rows = sorted((self.list_widget.row(item), item) for item in items)
        for row, item in reversed(rows):
            real_path = item.data(Qt.UserRole)
            preview_pdf = item.data(Qt.UserRole + 2)
            self.preview.remove_document_masks(str(real_path))
            try:
                if real_path and Path(real_path).exists():
                    Path(real_path).unlink(missing_ok=True)
            except Exception:
                pass
            try:
                if preview_pdf and Path(preview_pdf).exists() and preview_pdf != real_path:
                    Path(preview_pdf).unlink(missing_ok=True)
            except Exception:
                pass
            self.list_widget.takeItem(row)
        self.list_widget.refresh_grid()

    def make_preview_icon(self, preview_path: Path, size: QSize):
        try:
            doc = fitz.open(str(preview_path))
            page = doc.load_page(0)
            scale = max(0.20, min(0.55, size.width() / 300))
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            qimg = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format_RGB888).copy()
            doc.close()
            pixmap = QPixmap.fromImage(qimg).scaled(size, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            return QIcon(pixmap)
        except Exception:
            return self.style().standardIcon(QStyle.SP_FileIcon)

    def current_selected_items(self) -> list[QListWidgetItem]:
        return self.list_widget.selectedItems()

    def refresh_preview(self):
        items = self.current_selected_items()
        if not items:
            self.preview.clear()
            return
        preview_path = items[0].data(Qt.UserRole + 2)
        if preview_path:
            self.preview.show_pdf(preview_path)

    def open_sandbox(self):
        QDesktopServices.openUrl(QUrl.fromLocalFile(str(self.workdir)))

    def _current_single_pdf_item(self, title: str):
        items = self.current_selected_items()
        if len(items) != 1:
            QMessageBox.information(self, title, "Sélectionne un seul document PDF.")
            return None
        item = items[0]
        if item.data(Qt.UserRole + 1) != "pdf":
            QMessageBox.warning(self, title, "Cette fonction ne fonctionne que sur les PDF.")
            return None
        return item

    def start_mask_mode(self, color: QColor):
        item = self._current_single_pdf_item("Masquage")
        if item is None:
            return
        if self.preview.rotation != 0:
            QMessageBox.warning(self, "Masquage", "Remets d'abord la rotation d'aperçu à 0° avant de tracer des masques.")
            return
        self.preview.set_mask_mode(color)
        self.statusBar().showMessage(f"Mode masquage actif : {color.name()} | Trace un rectangle dans l'aperçu", 5000)

    def start_custom_mask_mode(self):
        color = QColorDialog.getColor(QColor("black"), self, "Choisir la couleur du masque")
        if not color.isValid():
            return
        self.start_mask_mode(color)

    def clear_masks_current_document(self):
        item = self._current_single_pdf_item("Masquage")
        if item is None:
            return
        if not self.preview.get_masks_for_document(str(Path(item.data(Qt.UserRole)))):
            self.preview.disable_mask_mode()
            self.statusBar().showMessage("Aucun masque à effacer", 3000)
            return
        self.preview.clear_masks_for_current_document()
        self.update_undo_action_state()
        self.statusBar().showMessage("Masques objets effacés", 4000)

    def delete_selected_mask(self):
        if self.preview.delete_selected_mask():
            self.update_undo_action_state()
            self.statusBar().showMessage("Masque supprimé", 3000)
        else:
            QMessageBox.information(self, "Masquage", "Aucun masque sélectionné sur la page courante.")

    def apply_masks_current_document(self):
        item = self._current_single_pdf_item("Masquage")
        if item is None:
            return
        src = Path(item.data(Qt.UserRole))
        masks_by_page = self.preview.get_masks_for_document(str(src))
        if not masks_by_page:
            QMessageBox.information(self, "Masquage", "Aucun masque à appliquer sur ce document.")
            return
        tmp_dst = FileOps.unique_pdf_path(self.workdir, f"{src.stem}_mask_apply")
        self.snapshot_state("Masquage appliqué")
        try:
            FileOps.apply_rect_masks(src, tmp_dst, masks_by_page)
            src.unlink(missing_ok=True)
            tmp_dst.rename(src)
            self.preview.clear_masks_for_document_no_history(str(src))
            self.preview.clear_mask_history()
            self.preview.disable_mask_mode()
            self.refresh_preview()
            self.update_undo_action_state()
            self.statusBar().showMessage("Masques appliqués au document", 5000)
        except Exception as e:
            try:
                if tmp_dst.exists():
                    tmp_dst.unlink(missing_ok=True)
            except Exception:
                pass
            QMessageBox.critical(self, "Masquage", str(e))

    def select_previous_document(self):
        current_row = self.list_widget.currentRow()
        if current_row > 0:
            self.list_widget.setCurrentRow(current_row - 1)
            item = self.list_widget.item(current_row - 1)
            if item:
                item.setSelected(True)

    def select_next_document(self):
        current_row = self.list_widget.currentRow()
        if current_row + 1 < self.list_widget.count():
            self.list_widget.setCurrentRow(current_row + 1)
            item = self.list_widget.item(current_row + 1)
            if item:
                item.setSelected(True)

    def open_context_menu(self, pos):
        menu = QMenu(self)
        selected = self.current_selected_items()
        kinds = {item.data(Qt.UserRole + 1) for item in selected}
        only_pdf = kinds <= {"pdf"}

        act_merge = menu.addAction("Fusionner")
        act_split = menu.addAction("Diviser")
        menu.addSeparator()
        act_num = menu.addAction("Ajouter numérotation")
        act_rotate_right = menu.addAction("Rotation 90° droite")
        act_rotate_left = menu.addAction("Rotation 90° gauche")
        act_stamp_quick = menu.addAction("Ajouter cachet des pièces (rapide)")
        act_stamp_advanced = menu.addAction("Ajouter cachet des pièces (avancé)")
        menu.addSeparator()
        act_mask_black = menu.addAction("Masquage noir")
        act_mask_white = menu.addAction("Masquage blanc")
        act_mask_color = menu.addAction("Masquage couleur…")
        act_apply_masks = menu.addAction("Appliquer les masques au PDF")
        act_delete_mask = menu.addAction("Supprimer le masque sélectionné")
        act_clear_masks = menu.addAction("Effacer tous les masques du document")
        menu.addSeparator()
        act_rename = menu.addAction("Renommer")
        act_stamp_text = menu.addAction("Modifier le texte du cachet")
        act_rename_bordereau = menu.addAction("Nommer les pièces selon bordereau")
        act_delete = menu.addAction("Supprimer")
        act_open = menu.addAction("Ouvrir dans l'explorateur")

        act_merge.setEnabled(len(selected) >= 2 and only_pdf)
        act_split.setEnabled(len(selected) == 1 and only_pdf)
        act_num.setEnabled(len(selected) >= 1 and only_pdf)
        act_rotate_right.setEnabled(len(selected) >= 1 and only_pdf)
        act_rotate_left.setEnabled(len(selected) >= 1 and only_pdf)
        act_stamp_quick.setEnabled(len(selected) >= 1 and only_pdf)
        act_stamp_advanced.setEnabled(len(selected) >= 1 and only_pdf)
        act_mask_black.setEnabled(len(selected) == 1 and only_pdf)
        act_mask_white.setEnabled(len(selected) == 1 and only_pdf)
        act_mask_color.setEnabled(len(selected) == 1 and only_pdf)
        act_apply_masks.setEnabled(len(selected) == 1 and only_pdf)
        act_delete_mask.setEnabled(len(selected) == 1 and only_pdf)
        act_clear_masks.setEnabled(len(selected) == 1 and only_pdf)
        act_rename.setEnabled(len(selected) == 1)
        act_stamp_text.setEnabled(True)
        act_rename_bordereau.setEnabled(len(selected) >= 1)
        act_delete.setEnabled(len(selected) >= 1)
        act_open.setEnabled(len(selected) == 1)

        chosen = menu.exec(self.list_widget.mapToGlobal(pos))
        if chosen == act_merge:
            self.merge_selected()
        elif chosen == act_split:
            self.split_selected()
        elif chosen == act_num:
            self.number_selected()
        elif chosen == act_rotate_right:
            self.rotate_selected_documents(90)
        elif chosen == act_rotate_left:
            self.rotate_selected_documents(-90)
        elif chosen == act_stamp_quick:
            self.stamp_selected_quick()
        elif chosen == act_stamp_advanced:
            self.stamp_selected_advanced()
        elif chosen == act_mask_black:
            self.start_mask_mode(QColor("black"))
        elif chosen == act_mask_white:
            self.start_mask_mode(QColor("white"))
        elif chosen == act_mask_color:
            self.start_custom_mask_mode()
        elif chosen == act_apply_masks:
            self.apply_masks_current_document()
        elif chosen == act_delete_mask:
            self.delete_selected_mask()
        elif chosen == act_clear_masks:
            self.clear_masks_current_document()
        elif chosen == act_rename:
            self.rename_selected()
        elif chosen == act_stamp_text:
            self.change_stamp_text()
        elif chosen == act_rename_bordereau:
            self.rename_selected_from_bordereau()
        elif chosen == act_delete:
            self.delete_selected()
        elif chosen == act_open and len(selected) == 1:
            path = selected[0].data(Qt.UserRole)
            if path:
                QDesktopServices.openUrl(QUrl.fromLocalFile(str(Path(path).parent)))

    def rename_item_file(self, item: QListWidgetItem, new_base_name: str):
        new_base_name = safe_stem(new_base_name)
        real_path = Path(item.data(Qt.UserRole))
        preview_path = Path(item.data(Qt.UserRole + 2))
        kind = item.data(Qt.UserRole + 1)

        if kind == "pdf":
            new_real_path = FileOps.unique_path(self.workdir, new_base_name, ".pdf")
            real_path.rename(new_real_path)
            item.setText(new_real_path.name)
            item.setData(Qt.UserRole, str(new_real_path))
            item.setData(Qt.UserRole + 2, str(new_real_path))
            item.setToolTip(str(new_real_path))
            item.setIcon(self.make_preview_icon(new_real_path, self.list_widget.iconSize()))
            item.setSizeHint(self.list_widget.gridSize())
            self.preview.move_document_masks(str(real_path), str(new_real_path))
            self.list_widget.refresh_grid()
            return

        if kind == "word":
            new_real_path = FileOps.unique_path(self.workdir, new_base_name, real_path.suffix.lower())
            real_path.rename(new_real_path)
            if preview_path.exists():
                new_preview_path = FileOps.unique_pdf_path(self.workdir, new_base_name)
                preview_path.rename(new_preview_path)
            else:
                new_preview_path = preview_path
            item.setText(new_real_path.name)
            item.setData(Qt.UserRole, str(new_real_path))
            item.setData(Qt.UserRole + 2, str(new_preview_path))
            item.setToolTip(str(new_real_path))
            item.setIcon(self.make_preview_icon(new_preview_path, self.list_widget.iconSize()))
            item.setSizeHint(self.list_widget.gridSize())
            self.list_widget.refresh_grid()
            return

    def rename_selected_from_bordereau(self):
        items = self.current_selected_items()
        if not items:
            QMessageBox.information(self, "Bordereau", "Sélectionne d'abord les documents à renommer.")
            return

        ordered_items = sorted(items, key=lambda it: self.list_widget.row(it))
        dialog = BordereauRenameDialog(self)
        if dialog.exec() != QDialog.Accepted:
            return

        lines = dialog.get_lines()
        if not lines:
            QMessageBox.warning(self, "Bordereau", "Aucune ligne n'a été saisie.")
            return
        if len(lines) != len(ordered_items):
            QMessageBox.warning(
                self,
                "Bordereau",
                f"Il y a {len(lines)} ligne(s) dans le bordereau pour {len(ordered_items)} document(s) sélectionné(s).\nLe nombre doit être identique.",
            )
            return

        self.snapshot_state("Renommage selon bordereau")
        errors = []
        for item, line in zip(ordered_items, lines):
            try:
                self.rename_item_file(item, line)
            except Exception as e:
                errors.append(f"{item.text()} -> {line} : {e}")

        self.refresh_preview()
        if errors:
            QMessageBox.warning(self, "Bordereau", "\n".join(errors))
        else:
            self.statusBar().showMessage("Renommage selon bordereau terminé", 5000)

    def apply_piece_labels(self, items: list[QListWidgetItem], labels: list[str]):
        self.snapshot_state("Cachet des pièces")
        errors = []
        for item, piece_label in zip(items, labels):
            src = Path(item.data(Qt.UserRole))
            safe_label = safe_stem(piece_label)
            dst = FileOps.unique_pdf_path(self.workdir, f"{src.stem}_piece_{safe_label}")
            try:
                FileOps.add_piece_stamp(src, dst, piece_label, self.stamp_title)
                src.unlink(missing_ok=True)
                self.replace_item_with_pdf(item, dst)
            except Exception as e:
                errors.append(f"{src.name}: {e}")

        self.refresh_preview()
        if errors:
            QMessageBox.warning(self, "Cachet", "\n".join(errors))
        else:
            self.statusBar().showMessage("Cachet des pièces ajouté", 5000)

    def stamp_selected_quick(self):
        items = self.current_selected_items()
        if not items:
            return
        if any(item.data(Qt.UserRole + 1) != "pdf" for item in items):
            QMessageBox.warning(self, "Cachet", "Le cachet de pièce ne fonctionne que sur les PDF.")
            return

        ordered_items = sorted(items, key=lambda it: self.list_widget.row(it))
        start_number, ok = QInputDialog.getInt(self, "Cachet rapide", "Numéro de départ :", 1, 1, 999999)
        if not ok:
            return
        labels = [str(start_number + i) for i in range(len(ordered_items))]
        self.apply_piece_labels(ordered_items, labels)

    def stamp_selected_advanced(self):
        items = self.current_selected_items()
        if not items:
            return
        if any(item.data(Qt.UserRole + 1) != "pdf" for item in items):
            QMessageBox.warning(self, "Cachet", "Le cachet de pièce ne fonctionne que sur les PDF.")
            return

        ordered_items = sorted(items, key=lambda it: self.list_widget.row(it))
        dialog = PieceLabelDialog(self)
        if dialog.exec() != QDialog.Accepted:
            return

        labels = dialog.get_labels()
        if not labels:
            QMessageBox.warning(self, "Cachet avancé", "Aucune numérotation n'a été saisie.")
            return
        if len(labels) != len(ordered_items):
            QMessageBox.warning(
                self,
                "Cachet avancé",
                f"Il y a {len(labels)} ligne(s) pour {len(ordered_items)} document(s) sélectionné(s).\nLe nombre doit être identique.",
            )
            return
        self.apply_piece_labels(ordered_items, labels)

    def change_stamp_text(self):
        text, ok = QInputDialog.getText(
            self,
            "Texte du cachet",
            "Texte affiché en haut du cachet :",
            text=self.stamp_title,
        )
        if not ok:
            return
        text = text.strip()
        if not text:
            QMessageBox.warning(self, "Texte du cachet", "Le texte du cachet ne peut pas être vide.")
            return
        self.snapshot_state("Texte du cachet")
        self.stamp_title = text
        self.settings.setValue("stamp_title", text)
        self.statusBar().showMessage(f"Texte du cachet mis à jour : {self.stamp_title}", 5000)

    def rename_selected(self):
        items = self.current_selected_items()
        if len(items) != 1:
            QMessageBox.information(self, "Renommer", "Sélectionne un seul document à renommer.")
            return

        item = items[0]
        real_path = Path(item.data(Qt.UserRole))
        preview_path = Path(item.data(Qt.UserRole + 2))
        current_stem = real_path.stem
        new_stem, ok = QInputDialog.getText(self, "Renommer", "Nouveau nom (sans extension) :", text=current_stem)
        if not ok:
            return
        new_stem = safe_stem(new_stem)
        if not new_stem:
            return

        self.snapshot_state("Renommage")
        try:
            new_real = FileOps.unique_path(self.workdir, new_stem, real_path.suffix)
            real_path.rename(new_real)
            if preview_path != real_path and preview_path.exists():
                new_preview = FileOps.unique_pdf_path(self.workdir, new_stem)
                preview_path.rename(new_preview)
            else:
                new_preview = new_real
            item.setText(new_real.name)
            item.setData(Qt.UserRole, str(new_real))
            item.setData(Qt.UserRole + 2, str(new_preview))
            item.setToolTip(str(new_real))
            item.setIcon(self.make_preview_icon(Path(new_preview), self.list_widget.iconSize()))
            item.setSizeHint(self.list_widget.gridSize())
            if item.data(Qt.UserRole + 1) == "pdf":
                self.preview.move_document_masks(str(real_path), str(new_real))
            self.list_widget.refresh_grid()
            self.refresh_preview()
            self.statusBar().showMessage("Document renommé", 3000)
        except Exception as e:
            QMessageBox.critical(self, "Renommer", str(e))

    def merge_selected(self):
        items = self.current_selected_items()
        if len(items) < 2:
            QMessageBox.information(self, "Fusion", "Sélectionne au moins 2 documents PDF.")
            return
        if any(item.data(Qt.UserRole + 1) != "pdf" for item in items):
            QMessageBox.warning(self, "Fusion", "La fusion ne fonctionne pour l'instant qu'avec des PDF.")
            return

        ordered = sorted(items, key=lambda i: self.list_widget.row(i))
        paths = [Path(i.data(Qt.UserRole)) for i in ordered]
        base_row = self.list_widget.row(ordered[0])
        name, ok = QInputDialog.getText(self, "Fusion", "Nom du PDF fusionné :", text="fusion")
        if not ok:
            return
        if not name.strip():
            name = "fusion"

        out_path = FileOps.unique_pdf_path(self.workdir, name)
        self.snapshot_state("Fusion")
        try:
            FileOps.merge_pdfs(paths, out_path)
            self.remove_items_and_files(ordered)
            new_item = self.add_doc_item({"kind": "pdf", "path": out_path, "source_name": out_path.name}, base_row)
            new_item.setSelected(True)
            self.list_widget.setCurrentItem(new_item)
            self.statusBar().showMessage("Fusion terminée", 5000)
        except Exception as e:
            QMessageBox.critical(self, "Fusion", str(e))

    def split_selected(self):
        items = self.current_selected_items()
        if len(items) != 1:
            QMessageBox.information(self, "Division", "Sélectionne un seul PDF à diviser.")
            return

        item = items[0]
        if item.data(Qt.UserRole + 1) != "pdf":
            QMessageBox.warning(self, "Division", "La division ne fonctionne que sur les PDF.")
            return

        src = Path(item.data(Qt.UserRole))
        mode, ok = QInputDialog.getItem(
            self,
            "Division",
            "Mode de division :",
            ["Plages personnalisées", "Tous les X pages"],
            editable=False,
        )
        if not ok:
            return

        self.snapshot_state("Division")
        try:
            if mode == "Plages personnalisées":
                text, ok = QInputDialog.getText(
                    self,
                    "Plages personnalisées",
                    "Exemple : 1-3,4-6,7-10",
                    text="1-2,3-4",
                )
                if not ok:
                    return
                ranges = self.parse_ranges(text)
                out_paths = FileOps.split_pdf_ranges(src, ranges, self.workdir)
            else:
                x, ok = QInputDialog.getInt(self, "Tous les X pages", "X =", 2, 1, 9999)
                if not ok:
                    return
                out_paths = FileOps.split_pdf_every_x(src, x, self.workdir)

            row = self.list_widget.row(item)
            self.remove_items_and_files([item])
            inserted = []
            for idx, path in enumerate(out_paths):
                inserted.append(self.add_doc_item({"kind": "pdf", "path": path, "source_name": path.name}, row + idx))
            if inserted:
                inserted[0].setSelected(True)
                self.list_widget.setCurrentItem(inserted[0])
            self.statusBar().showMessage("Division terminée", 5000)
        except Exception as e:
            QMessageBox.critical(self, "Division", str(e))

    def parse_ranges(self, text: str) -> list[tuple[int, int]]:
        text = text.strip()
        if not text:
            raise ValueError("Aucune plage saisie.")
        ranges = []
        for part in text.split(","):
            part = part.strip()
            if not part:
                continue
            if "-" in part:
                m = re.fullmatch(r"(\d+)\s*-\s*(\d+)", part)
                if not m:
                    raise ValueError(f"Format invalide : {part}")
                a, b = int(m.group(1)), int(m.group(2))
            else:
                if not re.fullmatch(r"\d+", part):
                    raise ValueError(f"Format invalide : {part}")
                a = b = int(part)
            if a > b:
                raise ValueError(f"Plage invalide : {part}")
            ranges.append((a, b))
        if not ranges:
            raise ValueError("Aucune plage valide.")
        return ranges

    def number_selected(self):
        items = self.current_selected_items()
        if not items:
            return
        if any(item.data(Qt.UserRole + 1) != "pdf" for item in items):
            QMessageBox.warning(self, "Numérotation", "La numérotation ne fonctionne que sur les PDF.")
            return

        self.snapshot_state("Numérotation")
        errors = []
        for item in items:
            src = Path(item.data(Qt.UserRole))
            dst = FileOps.unique_pdf_path(self.workdir, f"{src.stem}_num")
            try:
                FileOps.add_page_numbers(src, dst)
                src.unlink(missing_ok=True)
                self.replace_item_with_pdf(item, dst)
            except Exception as e:
                errors.append(f"{src.name}: {e}")

        self.refresh_preview()
        if errors:
            QMessageBox.warning(self, "Numérotation", "\n".join(errors))
        else:
            self.statusBar().showMessage("Numérotation ajoutée", 5000)

    def rotate_selected_documents(self, angle: int):
        items = self.current_selected_items()
        if not items:
            return
        if any(item.data(Qt.UserRole + 1) != "pdf" for item in items):
            QMessageBox.warning(self, "Rotation", "La rotation ne fonctionne que sur les PDF.")
            return
        self.snapshot_state("Rotation")
        errors = []
        suffix = "rot_d" if angle > 0 else "rot_g"
        for item in items:
            src = Path(item.data(Qt.UserRole))
            dst = FileOps.unique_pdf_path(self.workdir, f"{src.stem}_{suffix}")
            try:
                FileOps.rotate_pdf(src, dst, angle)
                src.unlink(missing_ok=True)
                self.replace_item_with_pdf(item, dst)
            except Exception as e:
                errors.append(f"{src.name}: {e}")
        self.refresh_preview()
        if errors:
            QMessageBox.warning(self, "Rotation", "\n".join(errors))
        else:
            self.statusBar().showMessage("Rotation appliquée", 5000)

    def delete_selected(self):
        items = self.current_selected_items()
        if not items:
            return
        message = f"Supprimer {len(items)} document(s) de la sandbox ?"
        if QMessageBox.question(self, "Suppression", message) != QMessageBox.Yes:
            return
        self.snapshot_state("Suppression")
        self.remove_items_and_files(items)
        self.preview.clear()
        self.statusBar().showMessage("Document(s) supprimé(s)", 5000)


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
