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