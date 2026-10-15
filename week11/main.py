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
