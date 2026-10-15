try:
    from pypdf import PdfReader, PdfWriter    # pypdf = ชื่อใหม่ของ PyPDF2 (ยังดูแลอยู่)
except ImportError:
    from PyPDF2 import PdfReader, PdfWriter
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.colors import Color
import io
import os
from utils import setup_logging, ensure_directory_exists

logger = setup_logging(__name__)

class PDFTasks:
    def __init__(self):
        logger.info("PDFTasks initialized.")

    def _create_text_watermark_pdf(self, watermark_text, temp_pdf_path="temp_watermark.pdf", font_size=50, color=(0,0,0,0.2), angle=45):
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=A4)
        width, height = A4
        r, g, b, a = color
        c.setFillColor(Color(r, g, b, alpha=a))
        c.setFont('Helvetica-Bold', font_size)
        c.translate(width/2, height/2)
        c.rotate(angle)
        text_width = c.stringWidth(watermark_text, 'Helvetica-Bold', font_size)
        c.drawString(-text_width/2, -font_size/4, watermark_text)
        c.showPage()
        c.save()
        buffer.seek(0)
        with open(temp_pdf_path, 'wb') as f:
            f.write(buffer.getvalue())
        logger.info(f"Generated temporary watermark PDF: {temp_pdf_path}")
        return temp_pdf_path

    def apply_watermark(self, input_pdf_path, output_pdf_path, watermark_text, font_size=50, color=(0.5,0.5,0.5,0.2), angle=45):
        if not os.path.exists(input_pdf_path):
            logger.error(f"Input PDF for watermarking not found: {input_pdf_path}")
            return False
        
        temp_watermark_pdf = os.path.join(os.path.dirname(output_pdf_path), "temp_text_watermark.pdf")
        ensure_directory_exists(os.path.dirname(output_pdf_path))
        try:
            self._create_text_watermark_pdf(watermark_text, temp_watermark_pdf, font_size, color, angle)
            watermark_reader = PdfReader(temp_watermark_pdf)
            watermark_page = watermark_reader.pages[0]
            reader = PdfReader(input_pdf_path)
            writer = PdfWriter()
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                page.merge_page(watermark_page)
                writer.add_page(page)
            with open(output_pdf_path, 'wb') as f:
                writer.write(f)
            logger.info(f"Watermark applied to '{input_pdf_path}'. Output saved to '{output_pdf_path}'.")
            return True
        except Exception as e:
            logger.error(f"Error applying watermark to '{input_pdf_path}': {e}")
            return False
        finally:
            if os.path.exists(temp_watermark_pdf):
                os.remove(temp_watermark_pdf)

    def merge_pdfs(self, pdf_paths, output_path):
        if not pdf_paths: return False
        writer = PdfWriter()
        ensure_directory_exists(os.path.dirname(output_path))
        for pdf_path in pdf_paths:
            if not os.path.exists(pdf_path): continue
            try:
                reader = PdfReader(pdf_path)
                for page in reader.pages:
                    writer.add_page(page)
            except Exception as e:
                logger.error(f"Error reading PDF '{pdf_path}': {e}")
        if len(writer.pages) == 0: return False
        try:
            with open(output_path, 'wb') as f:
                writer.write(f)
            return True
        except Exception as e:
            logger.error(f"Error writing merged PDF: {e}")
            return False

    def extract_text_from_pdf(self, pdf_path):
        if not os.path.exists(pdf_path): return None
        text_content = ""
        try:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                text_content += page.extract_text() + "\n--- Page End ---\n"
            return text_content
        except Exception as e:
            logger.error(f"Error extracting text: {e}")
            return None