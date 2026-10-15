# AllCode — Week 11 Agentic Document Automation
รวมซอร์สโค้ดทั้งหมดของโปรเจกต์ (สาริษฐ์ บุตรช่วง 663380365-3)

## `main.py`

```python
import sys
import os
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src'))

from config_parser import ConfigParser
from document_agent import DocumentAgent
from utils import setup_logging, ensure_directory_exists

logger = setup_logging(__name__)

PATH_KEYS = ['template_path', 'output_path', 'input_pdf', 'output_pdf', 'input_docx', 'output_text_file']
INPUT_KEYS = {'generate_word_report': ['template_path'], 'apply_pdf_watermark': ['input_pdf'],
              'extract_pdf_text': ['input_pdf']}


def check_required_files(config):
    """ตรวจสอบว่าไฟล์ input ทุกไฟล์ที่ task ต้องใช้มีอยู่จริง (ไม่อิงลำดับ task ใน config)"""
    missing = []
    for task in config['tasks']:
        for key in INPUT_KEYS.get(task['type'], []):
            if not os.path.exists(task[key]):
                missing.append(task[key])
        for p in task.get('input_pdfs', []):
            if not os.path.exists(p):
                missing.append(p)
    for path in missing:
        logger.critical(f"Missing input file: {path}")
    return not missing


def main():
    base_dir = os.path.dirname(os.path.abspath(__file__))
    config_file_path = os.path.join(base_dir, 'configs', 'document_automation_config.json')

    ensure_directory_exists(os.path.join(base_dir, 'documents'))
    ensure_directory_exists(os.path.join(base_dir, 'templates'))

    config_parser = ConfigParser(config_file_path)
    config = config_parser.load_config()

    # แปลง path ใน config ให้เป็น Absolute path (รวม output_path ของงานสร้าง Word ด้วย)
    for key, value in config.get('output_files', {}).items():
        config['output_files'][key] = os.path.join(base_dir, value)
    for task in config['tasks']:
        for k in PATH_KEYS:
            if k in task:
                task[k] = os.path.join(base_dir, task[k])
        if 'input_pdfs' in task:
            task['input_pdfs'] = [os.path.join(base_dir, p) for p in task['input_pdfs']]

    if not check_required_files(config):
        logger.critical("Missing required input files! Run 'python create_sample_documents.py' first.")
        return

    agent = DocumentAgent(config)
    agent.run()
    logger.info("Document automation process completed.")


if __name__ == "__main__":
    main()
```

## `create_sample_documents.py`

```python
# agentic-document-automator/create_sample_documents.py
# สร้างไฟล์ input ทั้งหมดที่ config ต้องใช้:
#   templates/report_template.docx  (เทมเพลต Word ที่มี {{PLACEHOLDER}})
#   documents/source_document.pdf    (ต้นฉบับสำหรับใส่ลายน้ำ / ดึงข้อความ)
#   documents/dummy_doc1.pdf, documents/dummy_doc2.pdf (สำหรับรวมไฟล์)
import os

import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from reportlab.lib.pagesizes import A4
from reportlab.lib.colors import HexColor
from reportlab.pdfgen import canvas

BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def create_word_template(path):
    doc = docx.Document()
    style = doc.styles["Normal"]
    style.font.name = "Calibri"
    style.font.size = Pt(11)

    title = doc.add_heading("{{REPORT_TITLE}}", level=0)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    meta = [("Date", "{{REPORT_DATE}}"), ("Author", "{{AUTHOR_NAME}}"), ("Company", "{{COMPANY_NAME}}")]
    for label, placeholder in meta:
        p = doc.add_paragraph()
        r = p.add_run(f"{label}: ")
        r.bold = True
        p.add_run(placeholder)
        p.paragraph_format.space_after = Pt(2)

    doc.add_heading("1. Introduction", level=1)
    doc.add_paragraph("{{INTRODUCTION_TEXT}}")

    doc.add_heading("2. Monthly Overview: {{MONTH_PERIOD}}", level=1)
    doc.add_paragraph("Here is a summary of key performance indicators:")
    doc.add_paragraph("{{TABLE_PLACEHOLDER}}")

    doc.add_heading("3. Conclusion", level=1)
    doc.add_paragraph("This report concludes the performance analysis. "
                      "All figures were filled in automatically by the Document Agent from the JSON configuration.")

    os.makedirs(os.path.dirname(path), exist_ok=True)
    doc.save(path)
    print(f"Created {path}")


def create_pdf(path, title, subtitle, sections, accent="#1F4E79"):
    """Simple multi-section A4 PDF with a coloured header bar."""
    c = canvas.Canvas(path, pagesize=A4)
    width, height = A4

    c.setFillColor(HexColor(accent))
    c.rect(0, height - 95, width, 95, fill=1, stroke=0)
    c.setFillColor(HexColor("#FFFFFF"))
    c.setFont("Helvetica-Bold", 22)
    c.drawString(50, height - 55, title)
    c.setFont("Helvetica", 12)
    c.drawString(50, height - 78, subtitle)

    y = height - 140
    c.setFillColor(HexColor("#000000"))
    for heading, lines in sections:
        c.setFont("Helvetica-Bold", 14)
        c.setFillColor(HexColor(accent))
        c.drawString(50, y, heading)
        y -= 22
        c.setFillColor(HexColor("#222222"))
        c.setFont("Helvetica", 11)
        for line in lines:
            c.drawString(60, y, line)
            y -= 17
        y -= 16

    c.setFont("Helvetica-Oblique", 9)
    c.setFillColor(HexColor("#777777"))
    c.drawString(50, 40, f"{os.path.basename(path)}  |  Generated by create_sample_documents.py")
    c.showPage()
    c.save()
    print(f"Created {path}")


def main():
    create_word_template(os.path.join(BASE_DIR, "templates", "report_template.docx"))

    docs = os.path.join(BASE_DIR, "documents")
    os.makedirs(docs, exist_ok=True)

    create_pdf(os.path.join(docs, "source_document.pdf"),
               "Quarterly Business Review", "Internal document - Q3 2026",
               [("Executive Summary", [
                   "Revenue grew steadily during the quarter, driven by new customer sign-ups.",
                   "Operating costs remained within the approved budget.",
                   "The automation project reduced manual document handling time by 40%."]),
                ("Key Metrics", [
                   "- Sales Revenue: $120,000 (up 10%)",
                   "- Customer Acquisition: 500 new customers",
                   "- Project Completion Rate: 95%"]),
                ("Next Steps", [
                   "1. Expand automated reporting to the finance team.",
                   "2. Add digital signatures to generated PDFs.",
                   "3. Review data retention policy for archived reports."])])

    create_pdf(os.path.join(docs, "dummy_doc1.pdf"),
               "Project Documentation - Part A", "System Architecture & Design",
               [("Chapter 1: Introduction", [
                   "The Agentic Document Automation System processes Word and PDF files",
                   "based on tasks declared in a JSON configuration file."]),
                ("Key Features", [
                   "- Template-based Word report generation",
                   "- PDF watermarking, merging and text extraction",
                   "- Audit logging of every task"])],
               accent="#2E7D32")

    create_pdf(os.path.join(docs, "dummy_doc2.pdf"),
               "Project Documentation - Part B", "Implementation & Testing",
               [("Chapter 2: Implementation", [
                   "DocumentAgent dispatches each task type to WordTasks or PDFTasks.",
                   "Results and errors are recorded in an audit report (JSON)."]),
                ("Chapter 3: Testing", [
                   "- Verified placeholders are replaced in the generated report",
                   "- Verified watermark appears on every page",
                   "- Verified merged PDF contains Part A followed by Part B"])],
               accent="#C62828")


if __name__ == "__main__":
    main()
```

## `src/__init__.py`

```python

```

## `src/utils.py`

```python
import logging
import os
import json
import datetime

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def setup_logging(name, level=logging.INFO, filename=None):
    logger = logging.getLogger(name)
    logger.setLevel(level)
    if not logger.handlers:
        ch = logging.StreamHandler()
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        ch.setFormatter(formatter)
        logger.addHandler(ch)
        if filename:
            fh = logging.FileHandler(filename)
            fh.setFormatter(formatter)
            logger.addHandler(fh)
    return logger

def ensure_directory_exists(path):
    try:
        os.makedirs(path, exist_ok=True)
        logging.info(f"Directory ensured: {path}")
    except OSError as e:
        logging.error(f"Error creating directory {path}: {e}")
        raise

def log_audit_entry(audit_log_list, status, message, task_type="N/A", details=None):
    log_entry = {
        "timestamp": datetime.datetime.now().isoformat(),
        "status": status,
        "task_type": task_type,
        "message": message,
        "details": details if details else {}
    }
    audit_log_list.append(log_entry)
    if status in ["ERROR", "CRITICAL"]:
        logging.error(f"AUDIT LOG - {status}: {message} | Task: {task_type} | Details: {details}")
    else:
        logging.info(f"AUDIT LOG - {status}: {message} | Task: {task_type}")
```

## `src/config_parser.py`

```python
import json
import os
from utils import setup_logging

logger = setup_logging(__name__)

class ConfigParser:
    def __init__(self, config_path):
        self.config_path = config_path
        self.config = {}

    def load_config(self):
        if not os.path.exists(self.config_path):
            logger.critical(f"Configuration file not found: {self.config_path}")
            raise FileNotFoundError(f"Configuration file not found: {self.config_path}")
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                self.config = json.load(f)
            self._validate_config()
            logger.info(f"Configuration loaded successfully from {self.config_path}")
            return self.config
        except json.JSONDecodeError as e:
            logger.critical(f"Invalid JSON format in config file: {e}")
            raise ValueError(f"Invalid JSON format in config file: {e}")

    def _validate_config(self):
        required_top_level_fields = ["global_data", "tasks"]
        for field in required_top_level_fields:
            if field not in self.config:
                raise ValueError(f"Missing required field in config: '{field}'")
        if not isinstance(self.config["tasks"], list):
            raise ValueError("Config field 'tasks' must be a list.")
        
        for i, task in enumerate(self.config["tasks"]):
            if "type" not in task:
                raise ValueError(f"Task {i} is missing 'type' field.")
            if task["type"] == "generate_word_report":
                if "template_path" not in task or "output_path" not in task:
                    raise ValueError(f"Task {i} requires 'template_path' and 'output_path'.")
            elif task["type"] == "apply_pdf_watermark":
                if "input_pdf" not in task or "output_pdf" not in task or "watermark_text" not in task:
                    raise ValueError(f"Task {i} requires 'input_pdf', 'output_pdf', and 'watermark_text'.")
        logger.info("Configuration validated.")
```

## `src/document_agent.py`

```python
import os
import json
import datetime
from utils import setup_logging, log_audit_entry, ensure_directory_exists
from pdf_tasks import PDFTasks
from word_tasks import WordTasks

logger = setup_logging(__name__)

class DocumentAgent:
    def __init__(self, config):
        self.config = config
        self.audit_log = []
        self.start_time = datetime.datetime.now()
        self.pdf_processor = PDFTasks()
        self.word_processor = WordTasks()

    def _process_task(self, task):
        task_type = task.get("type")
        task_name = task.get("name", task_type)
        log_audit_entry(self.audit_log, "INFO", f"Starting task '{task_name}'", task_type, task)
        try:
            success = False
            if task_type == "generate_word_report":
                placeholders = {**self.config.get("global_data", {}), **task.get("data", {})}
                success = self.word_processor.generate_report_from_template(
                    task["template_path"], task["output_path"], placeholders, task.get("table_data")
                )
            elif task_type == "apply_pdf_watermark":
                success = self.pdf_processor.apply_watermark(
                    task["input_pdf"], task["output_pdf"], task["watermark_text"],
                    task.get("font_size", 50), tuple(task.get("color", [0.5,0.5,0.5,0.2])), task.get("angle", 45)
                )
            elif task_type == "merge_pdfs":
                success = self.pdf_processor.merge_pdfs(task["input_pdfs"], task["output_pdf"])
            elif task_type == "extract_pdf_text":
                text = self.pdf_processor.extract_text_from_pdf(task["input_pdf"])
                if text:
                    ensure_directory_exists(os.path.dirname(task["output_text_file"]))
                    with open(task["output_text_file"], 'w', encoding='utf-8') as f: f.write(text)
                    success = True
            elif task_type == "extract_word_text":
                text = self.word_processor.extract_text_from_word(task["input_docx"])
                if text:
                    ensure_directory_exists(os.path.dirname(task["output_text_file"]))
                    with open(task["output_text_file"], 'w', encoding='utf-8') as f: f.write(text)
                    success = True
            
            if success: log_audit_entry(self.audit_log, "SUCCESS", f"Task '{task_name}' completed.", task_type)
            else: log_audit_entry(self.audit_log, "ERROR", f"Task '{task_name}' failed.", task_type)
        except Exception as e:
            log_audit_entry(self.audit_log, "ERROR", f"Error in '{task_name}': {e}", task_type)

    def _generate_audit_report(self, output_dir="documents"):
        ensure_directory_exists(output_dir)
        report_path = os.path.join(output_dir, f"audit_report_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.json")
        with open(report_path, 'w', encoding='utf-8') as f:
            json.dump(self.audit_log, f, indent=4, ensure_ascii=False)

    def run(self):
        log_audit_entry(self.audit_log, "INFO", "Document Agent started.", "agent_run")
        for task in self.config.get("tasks", []):
            self._process_task(task)
        duration = (datetime.datetime.now() - self.start_time).total_seconds()
        log_audit_entry(self.audit_log, "INFO", f"Finished in {duration:.2f}s.", "agent_run_summary")
        self._generate_audit_report(os.path.join(os.path.dirname(self.config['output_files']['generated_report']), 'reports'))
```

## `src/word_tasks.py`

```python
# agentic-document-automator/src/word_tasks.py
import copy
import os

import docx
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

from utils import setup_logging

logger = setup_logging(__name__)

TABLE_PLACEHOLDER = "{{TABLE_PLACEHOLDER}}"


class WordTasks:
    """Handles advanced Word document operations for the Document Agent."""

    def __init__(self):
        logger.info("WordTasks initialized.")

    @staticmethod
    def _as_placeholder(key):
        """'REPORT_TITLE' -> '{{REPORT_TITLE}}' (keys that already have braces are kept as-is)."""
        key = str(key)
        return key if key.startswith("{{") and key.endswith("}}") else "{{" + key + "}}"

    def _replace_placeholder_in_paragraph(self, paragraph, placeholder, value):
        """
        Replaces a placeholder in a paragraph.
        Word often splits text into several runs (e.g. '{{REPORT_' + 'TITLE}}'),
        so when the placeholder is not inside a single run, the paragraph text is
        rebuilt into the first run (keeping that run's formatting).
        """
        if placeholder not in paragraph.text:
            return False

        replaced = False
        for run in paragraph.runs:
            if placeholder in run.text:
                run.text = run.text.replace(placeholder, str(value))
                replaced = True

        if not replaced and paragraph.runs:  # placeholder spans multiple runs
            full_text = paragraph.text.replace(placeholder, str(value))
            paragraph.runs[0].text = full_text
            for run in paragraph.runs[1:]:
                run.text = ""
        logger.debug(f"Replaced '{placeholder}' with '{value}'.")
        return True

    def _build_table(self, document, table_data):
        """Creates a formatted table (header row bold + shaded) at the end of the body."""
        cols = len(table_data[0])
        table = document.add_table(rows=0, cols=cols)
        try:
            table.style = 'Table Grid'
        except Exception:
            logger.warning("Style 'Table Grid' not available, using default table style.")

        for r_idx, row_data in enumerate(table_data):
            row = table.add_row()
            for c_idx, cell_data in enumerate(row_data):
                cell = row.cells[c_idx]
                cell.text = str(cell_data)
                run = cell.paragraphs[0].runs[0] if cell.paragraphs[0].runs else None
                if r_idx == 0 and run is not None:
                    run.bold = True
                    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    self._shade_cell(cell, "1F4E79")
                if c_idx > 0:
                    cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
        return table

    @staticmethod
    def _shade_cell(cell, hex_color):
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn
        tc_pr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), hex_color)
        tc_pr.append(shd)

    def generate_report_from_template(self, template_path, output_path, data_placeholders, table_data=None):
        """
        Generates a Word document from a template, filling placeholders and optional table data.
        `data_placeholders` keys may be written as 'REPORT_TITLE' or '{{REPORT_TITLE}}'.
        `table_data` is a list of lists, where the first sub-list is the header row.
        """
        if not os.path.exists(template_path):
            logger.error(f"Word template not found: {template_path}")
            return False

        try:
            document = docx.Document(template_path)
            placeholders = {self._as_placeholder(k): v for k, v in data_placeholders.items()}

            # Replace placeholders in paragraphs
            for paragraph in document.paragraphs:
                for placeholder, value in placeholders.items():
                    self._replace_placeholder_in_paragraph(paragraph, placeholder, value)

            # Replace placeholders in existing table cells
            for table in document.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            for placeholder, value in placeholders.items():
                                self._replace_placeholder_in_paragraph(paragraph, placeholder, value)

            # Add dynamic table if provided
            if table_data:
                logger.info("Adding dynamic table to the document.")
                table = self._build_table(document, table_data)

                anchor = next((p for p in document.paragraphs if TABLE_PLACEHOLDER in p.text), None)
                if anchor is not None:
                    # move the table to where the placeholder paragraph is, then remove the placeholder
                    anchor._element.addnext(table._tbl)
                    anchor._element.getparent().remove(anchor._element)
                    logger.info("Table inserted at '{{TABLE_PLACEHOLDER}}'.")
                else:
                    logger.warning("Table placeholder '{{TABLE_PLACEHOLDER}}' not found in template. Table appended to end.")

            # warn about anything that was not filled
            leftovers = [p.text for p in document.paragraphs if "{{" in p.text]
            if leftovers:
                logger.warning(f"Unfilled placeholders remain: {leftovers}")

            os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
            document.save(output_path)
            logger.info(f"Report generated from template and saved to '{output_path}'.")
            return True

        except Exception as e:
            logger.error(f"Error generating report from template: {e}")
            return False

    def extract_text_from_word(self, docx_path):
        """Extracts all text from a Word document (paragraphs and tables, in document order)."""
        if not os.path.exists(docx_path):
            logger.error(f"Word document not found: {docx_path}")
            return None

        try:
            document = docx.Document(docx_path)
            full_text = []
            body = document.element.body
            for child in body.iterchildren():
                tag = child.tag.split('}')[-1]
                if tag == 'p':
                    full_text.append(docx.text.paragraph.Paragraph(child, document).text)
                elif tag == 'tbl':
                    table = docx.table.Table(child, document)
                    for row in table.rows:
                        full_text.append("\t".join(cell.text for cell in row.cells))  # Tab-separated cell values

            logger.info(f"Text extracted from Word document: {docx_path}")
            return "\n".join(full_text)

        except Exception as e:
            logger.error(f"Error extracting text from Word document '{docx_path}': {e}")
            return None
```

## `src/pdf_tasks.py`

```python
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
```

## `configs/document_automation_config.json`

```json
{
  "global_data": {
    "REPORT_DATE": "2026-10-15",
    "AUTHOR_NAME": "AI Document Agent",
    "COMPANY_NAME": "Automated Docs Inc."
  },
  "output_files": {
    "generated_report": "documents/generated_report.docx",
    "watermarked_pdf": "documents/watermarked_document.pdf",
    "merged_pdf": "documents/merged_docs.pdf"
  },
  "tasks": [
    {
      "name": "Generate Monthly Performance Report",
      "type": "generate_word_report",
      "template_path": "templates/report_template.docx",
      "output_path": "documents/generated_report.docx",
      "data": {
        "REPORT_TITLE": "Monthly Performance Summary",
        "INTRODUCTION_TEXT": "This automated report provides a summary of key performance indicators.",
        "MONTH_PERIOD": "September 2026"
      },
      "table_data": [
        ["Category", "Value", "Trend"],
        ["Sales Revenue", "$120,000", "Up 10%"],
        ["Customer Acquisition", "500", "Stable"],
        ["Project Completion Rate", "95%", "Excellent"]
      ]
    },
    {
      "name": "Apply Confidential Watermark to PDF",
      "type": "apply_pdf_watermark",
      "input_pdf": "documents/source_document.pdf",
      "output_pdf": "documents/watermarked_document.pdf",
      "watermark_text": "CONFIDENTIAL",
      "font_size": 70,
      "color": [1, 0, 0, 0.15],
      "angle": 55
    },
    {
      "name": "Merge Project Documentation",
      "type": "merge_pdfs",
      "input_pdfs": ["documents/dummy_doc1.pdf", "documents/dummy_doc2.pdf"],
      "output_pdf": "documents/merged_docs.pdf"
    },
    {
      "name": "Extract Text from Generated Word Report",
      "type": "extract_word_text",
      "input_docx": "documents/generated_report.docx",
      "output_text_file": "documents/extracted_word_text.txt"
    },
    {
      "name": "Extract Text from Source PDF",
      "type": "extract_pdf_text",
      "input_pdf": "documents/source_document.pdf",
      "output_text_file": "documents/extracted_pdf_text.txt"
    }
  ]
}
```

## `requirements.txt`

```text
pypdf>=4.0
reportlab>=4.0
python-docx>=1.1
```
