"""
JSON output generator for AWS Auto Inventory.
"""
import json
import os
import logging
from datetime import datetime
from typing import Dict, Any


class JSONOutputGenerator:
    """
    Generates JSON output from scan results.
    """
    
    def __init__(self):
        """Initialize the JSON generator."""
        self.logger = logging.getLogger(__name__)
    
    def generate(self, results, output_dir: str) -> None:
        """
        Generate JSON output files from scan results.
        
        Args:
            results: Scan results (list of ScanResult objects or dictionary)
            output_dir: Directory to store output files
        """
        self.logger.info("Generating JSON output")
        
        # Convert ScanResult objects to dictionaries if needed
        processed_results = self._process_results(results)
        
        # Generate timestamp for filenames
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Generate main results file
        main_file = os.path.join(output_dir, f"aws_inventory_{timestamp}.json")
        self._write_json_file(processed_results, main_file)
        
        # Generate summary file
        summary = self._generate_summary(processed_results)
        summary_file = os.path.join(output_dir, f"aws_inventory_summary_{timestamp}.json")
        self._write_json_file(summary, summary_file)
        
        self.logger.info(f"JSON files generated: {main_file}, {summary_file}")
    
    def _process_results(self, results):
        """
        Process scan results and convert ScanResult objects to dictionaries.
        
        Args:
            results: Raw scan results (list of ScanResult objects or dictionary)
            
        Returns:
            Processed results as dictionary
        """
        if isinstance(results, list):
            # Handle list of ScanResult objects
            processed = {}
            for scan_result in results:
                if hasattr(scan_result, 'to_dict'):
                    result_dict = scan_result.to_dict()
                    processed[scan_result.inventory_name] = result_dict
                else:
                    # Fallback for non-ScanResult objects
                    processed[f"result_{len(processed)}"] = scan_result
            return processed
        else:
            # Already a dictionary
            return results
    
    def _write_json_file(self, data: Dict[str, Any], filepath: str) -> None:
        """
        Write data to JSON file.
        
        Args:
            data: Data to write
            filepath: Path to output file
        """
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, default=str, ensure_ascii=False)
        except Exception as e:
            self.logger.error(f"Error writing JSON file {filepath}: {e}")
            raise
    
    def _generate_summary(self, results: Dict[str, Any]) -> Dict[str, Any]:
        """
        Generate summary statistics from scan results.
        
        Args:
            results: Processed scan results dictionary
            
        Returns:
            Summary dictionary
        """
        summary = {
            "scan_timestamp": datetime.now().isoformat(),
            "total_inventories": 0,
            "total_regions": 0,
            "total_services": 0,
            "total_resources": 0,
            "inventories": {},
            "services": {}
        }
        
        # Process results to generate summary
        for inventory_name, inventory_data in results.items():
            if isinstance(inventory_data, dict):
                summary["total_inventories"] += 1
                inventory_resources = 0
                inventory_regions = set()
                inventory_services = set()
                
                # Check if this is organization results or account results
                if "organization_results" in inventory_data:
                    # Organization scan
                    for account_result in inventory_data["organization_results"]:
                        inventory_resources += self._count_account_resources(
                            account_result, inventory_regions, inventory_services, summary["services"]
                        )
                elif "account_results" in inventory_data:
                    # Single account scan
                    for region_result in inventory_data["account_results"]:
                        inventory_resources += self._count_region_resources(
                            region_result, inventory_regions, inventory_services, summary["services"]
                        )
                
                summary["total_regions"] += len(inventory_regions)
                summary["total_services"] += len(inventory_services)
                summary["total_resources"] += inventory_resources
                
                summary["inventories"][inventory_name] = {
                    "total_resources": inventory_resources,
                    "regions": list(inventory_regions),
                    "services": list(inventory_services)
                }
        
        return summary
    
    def _count_account_resources(self, account_result: Dict[str, Any], regions: set, services: set, service_summary: Dict[str, int]) -> int:
        """Count resources in an account result."""
        total_resources = 0
        # This would need to be implemented based on AccountResult structure
        # For now, return 0 as we don't have organization scanning implemented
        return total_resources
    
    def _count_region_resources(self, region_result: Dict[str, Any], regions: set, services: set, service_summary: Dict[str, int]) -> int:
        """Count resources in a region result."""
        total_resources = 0
        
        if "region" in region_result:
            regions.add(region_result["region"])
        
        if "services" in region_result:
            for service_result in region_result["services"]:
                if isinstance(service_result, dict):
                    service_name = service_result.get("service", "unknown")
                    services.add(service_name)
                    
                    # Count resources in the service result
                    if service_result.get("success", False) and "result" in service_result:
                        result_data = service_result["result"]
                        resource_count = self._count_service_resources(result_data)
                        total_resources += resource_count
                        
                        # Track service statistics
                        if service_name not in service_summary:
                            service_summary[service_name] = 0
                        service_summary[service_name] += resource_count
        
        return total_resources
    
    def _count_service_resources(self, result_data: Any) -> int:
        """Count resources in a service result."""
        if isinstance(result_data, list):
            return len(result_data)
        elif isinstance(result_data, dict):
            # Look for common AWS response patterns
            for key in ["Reservations", "Buckets", "Roles", "Functions", "Instances", "Items"]:
                if key in result_data:
                    value = result_data[key]
                    if isinstance(value, list):
                        # For EC2 reservations, count instances within reservations
                        if key == "Reservations":
                            instance_count = 0
                            for reservation in value:
                                if isinstance(reservation, dict) and "Instances" in reservation:
                                    instance_count += len(reservation["Instances"])
                            return instance_count
                        else:
                            return len(value)
            # If no known keys found, try to count any list values
            for value in result_data.values():
                if isinstance(value, list):
                    return len(value)
        
        return 0