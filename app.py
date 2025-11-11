#!/usr/bin/env python3
import argparse
import sys
from pathlib import Path
import fitz  # PyMuPDF

from io import BytesIO
try:
    from PIL import Image
except Exception:
    Image = None

A4_WIDTH_PT = 595.2755905511812  # 210 mm at 72 dpi
A4_HEIGHT_PT = 841.8897637795277  # 297 mm at 72 dpi


def fit_rect_keep_aspect(img_w: int, img_h: int, page_w: float, page_h: float) -> fitz.Rect:
    """Return a Rect that fits the image entirely inside the page while keeping aspect ratio, centered."""
    scale = min(page_w / img_w, page_h / img_h)
    w = img_w * scale
    h = img_h * scale
    x = (page_w - w) / 2
    y = (page_h - h) / 2
    return fitz.Rect(x, y, x + w, y + h)


# Helper: encode Pixmap to JPEG with quality, fallback to Pillow if needed
def encode_pixmap_to_jpeg(pix: fitz.Pixmap, quality: int) -> bytes:
    """Return JPEG-encoded bytes of a Pixmap with quality control.
    Tries PyMuPDF's get_image_data first; falls back to Pillow if needed."""
    # Prefer PyMuPDF API that supports quality
    try:
        return pix.get_image_data(output="jpeg", quality=quality)
    except Exception:
        # Fall back to Pillow if available
        if Image is None:
            raise RuntimeError(
                "PyMuPDF get_image_data not available; install Pillow (`pip install Pillow`) for fallback."
            )
        mode = "L" if pix.colorspace.n == 1 else "RGB"
        img = Image.frombytes(mode, (pix.width, pix.height), pix.samples)
        buf = BytesIO()
        img.save(buf, format="JPEG", quality=quality, optimize=True)
        return buf.getvalue()


def rasterize_pdf_to_images_pdf(
    input_pdf: Path,
    dpi: int = 144,
    quality: int = 60,
    grayscale: bool = False,
    a4_portrait: bool = True,
) -> Path:
    if not input_pdf.exists():
        raise FileNotFoundError(f"Input not found: {input_pdf}")
    if input_pdf.suffix.lower() != ".pdf":
        raise ValueError("Input must be a .pdf file")

    out_path = input_pdf.with_name(f"{input_pdf.stem} (01).pdf")

    # Open source and destination PDFs
    src = fitz.open(input_pdf)
    dst = fitz.open()

    try:
        if src.needs_pass:
            raise RuntimeError("The input PDF is password-protected.")

        # Prepare page size (A4 portrait or landscape)
        if a4_portrait:
            page_w, page_h = A4_WIDTH_PT, A4_HEIGHT_PT
        else:
            page_w, page_h = A4_HEIGHT_PT, A4_WIDTH_PT

        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)

        for i in range(len(src)):
            page = src.load_page(i)
            # Render page to pixmap
            pix = page.get_pixmap(matrix=mat, colorspace=(fitz.csGRAY if grayscale else fitz.csRGB), alpha=False)

            # Encode to JPEG bytes with desired quality
            img_bytes = encode_pixmap_to_jpeg(pix, quality)

            # Create an A4 page and place image centered, preserving aspect ratio
            new_page = dst.new_page(width=page_w, height=page_h)
            rect = fit_rect_keep_aspect(pix.width, pix.height, page_w, page_h)
            new_page.insert_image(rect, stream=img_bytes)

        # Save optimized
        dst.save(out_path, deflate=True, clean=True, garbage=4)
        return out_path
    finally:
        src.close()
        dst.close()


if __name__ == "__main__":
    import tkinter as tk
    from tkinter import filedialog, messagebox

    def select_file():
        file_path = filedialog.askopenfilename(
            title="בחר קובץ PDF",
            filetypes=[("PDF files", "*.pdf")],
        )
        if not file_path:
            return
        input_path = Path(file_path)
        try:
            output_path = rasterize_pdf_to_images_pdf(input_path)
            messagebox.showinfo("הצלחה", f"הקובץ נוצר בהצלחה:\n{output_path}")
        except Exception as e:
            messagebox.showerror("שגיאה", f"אירעה שגיאה:\n{e}")

    root = tk.Tk()
    root.title("PDF Rasterizer")
    root.geometry("200x100")

    btn = tk.Button(root, text="בחר קובץ", command=select_file, font=("Arial", 14))
    btn.pack(expand=True)

    root.mainloop()