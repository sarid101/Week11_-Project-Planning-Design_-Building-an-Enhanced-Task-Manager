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