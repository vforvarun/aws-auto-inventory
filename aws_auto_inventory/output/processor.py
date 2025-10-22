"""
Output processor for AWS Auto Inventory.
"""
import os
import logging
from typing import List, Dict, Any

from .json_generator import JSONOutputGenerator
from .excel_generator import ExcelOutputGenerator


class OutputProcessor:
    """
    Main output processor that coordinates different output formats.
    """
    
    def __init__(self):
        """Initialize the output processor."""
        self.logger = logging.getLogger(__name__)
        self.json_generator = JSONOutputGenerator()
        self.excel_generator = ExcelOutputGenerator()
    
    def process(self, results: Dict[str, Any], output_dir: str, formats: List[str]) -> None:
        """
        Process scan results and generate output in specified formats.
        
        Args:
            results: Scan results dictionary
            output_dir: Directory to store output files
            formats: List of output formats ('json', 'excel')
        """
        # Ensure output directory exists
        os.makedirs(output_dir, exist_ok=True)
        
        self.logger.info(f"Processing output to {output_dir} in formats: {formats}")
        
        for format_type in formats:
            if format_type == "json":
                self._process_json(results, output_dir)
            elif format_type == "excel":
                self._process_excel(results, output_dir)
            else:
                self.logger.warning(f"Unknown output format: {format_type}")
    
    def _process_json(self, results: Dict[str, Any], output_dir: str) -> None:
        """Process JSON output."""
        try:
            self.json_generator.generate(results, output_dir)
            self.logger.info("JSON output generated successfully")
        except Exception as e:
            self.logger.error(f"Error generating JSON output: {e}")
            raise
    
    def _process_excel(self, results: Dict[str, Any], output_dir: str) -> None:
        """Process Excel output."""
        try:
            self.excel_generator.generate(results, output_dir)
            self.logger.info("Excel output generated successfully")
        except Exception as e:
            self.logger.error(f"Error generating Excel output: {e}")
            raise