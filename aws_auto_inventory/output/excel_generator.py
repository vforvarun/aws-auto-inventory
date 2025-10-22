"""
Excel output generator for AWS Auto Inventory.
"""
import os
import logging
from datetime import datetime
from typing import Dict, Any, List

try:
    import pandas as pd
    PANDAS_AVAILABLE = True
except ImportError:
    pd = None
    PANDAS_AVAILABLE = False


class ExcelOutputGenerator:
    """
    Generates Excel output from scan results.
    """
    
    def __init__(self):
        """Initialize the Excel generator."""
        self.logger = logging.getLogger(__name__)
        
        if not PANDAS_AVAILABLE:
            self.logger.warning("pandas not available. Excel output will be disabled.")
    
    def generate(self, results, output_dir: str) -> None:
        """
        Generate Excel output files from scan results.
        
        Args:
            results: Scan results (list of ScanResult objects or dictionary)
            output_dir: Directory to store output files
        """
        if not PANDAS_AVAILABLE:
            self.logger.error("Cannot generate Excel output: pandas not installed")
            raise ImportError("pandas is required for Excel output. Install with: pip install pandas openpyxl")
        
        self.logger.info("Generating Excel output")
        
        # Convert ScanResult objects to dictionaries if needed
        processed_results = self._process_results(results)
        
        # Generate timestamp for filename
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_file = os.path.join(output_dir, f"aws_inventory_{timestamp}.xlsx")
        
        # Convert results to DataFrames and write to Excel
        self._write_excel_file(processed_results, excel_file)
        
        self.logger.info(f"Excel file generated: {excel_file}")
    
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
    
    def _write_excel_file(self, results: Dict[str, Any], filepath: str) -> None:
        """
        Write results to Excel file with multiple sheets.
        
        Args:
            results: Scan results dictionary
            filepath: Path to output file
        """
        try:
            with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
                # Create summary sheet
                summary_df = self._create_summary_dataframe(results)
                summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
                # Create detailed sheets for each service
                service_data = self._extract_service_data(results)
                
                for service_name, service_resources in service_data.items():
                    if service_resources:
                        # Limit sheet name length and sanitize
                        sheet_name = self._sanitize_sheet_name(service_name)
                        df = pd.DataFrame(service_resources)
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                
        except Exception as e:
            self.logger.error(f"Error writing Excel file {filepath}: {e}")
            raise
    
    def _create_summary_dataframe(self, results: Dict[str, Any]):
        """
        Create summary DataFrame from results.
        
        Args:
            results: Processed scan results dictionary
            
        Returns:
            Summary DataFrame
        """
        summary_data = []
        
        for inventory_name, inventory_data in results.items():
            if isinstance(inventory_data, dict):
                # Check if this is organization results or account results
                if "organization_results" in inventory_data:
                    # Organization scan - not implemented yet
                    pass
                elif "account_results" in inventory_data:
                    # Single account scan
                    for region_result in inventory_data["account_results"]:
                        if isinstance(region_result, dict) and "region" in region_result:
                            region_name = region_result["region"]
                            
                            if "services" in region_result:
                                for service_result in region_result["services"]:
                                    if isinstance(service_result, dict):
                                        service_name = service_result.get("service", "unknown")
                                        success = service_result.get("success", False)
                                        
                                        if success and "result" in service_result:
                                            resource_count = self._count_service_resources(service_result["result"])
                                        else:
                                            resource_count = 0
                                        
                                        summary_data.append({
                                            'Inventory': inventory_name,
                                            'Region': region_name,
                                            'Service': service_name,
                                            'Resource Count': resource_count,
                                            'Status': 'Success' if success else 'Failed'
                                        })
        
        return pd.DataFrame(summary_data)
    
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
    
    def _extract_service_data(self, results: Dict[str, Any]) -> Dict[str, List[Dict[str, Any]]]:
        """
        Extract and flatten service data for Excel sheets.
        
        Args:
            results: Processed scan results dictionary
            
        Returns:
            Dictionary mapping service names to lists of resource data
        """
        service_data = {}
        
        for inventory_name, inventory_data in results.items():
            if isinstance(inventory_data, dict):
                # Check if this is organization results or account results
                if "organization_results" in inventory_data:
                    # Organization scan - not implemented yet
                    pass
                elif "account_results" in inventory_data:
                    # Single account scan
                    for region_result in inventory_data["account_results"]:
                        if isinstance(region_result, dict) and "region" in region_result:
                            region_name = region_result["region"]
                            
                            if "services" in region_result:
                                for service_result in region_result["services"]:
                                    if isinstance(service_result, dict):
                                        service_name = service_result.get("service", "unknown")
                                        success = service_result.get("success", False)
                                        
                                        if success and "result" in service_result:
                                            result_data = service_result["result"]
                                            
                                            if service_name not in service_data:
                                                service_data[service_name] = []
                                            
                                            # Process the result data
                                            resources = self._flatten_service_result(result_data, inventory_name, region_name, service_name)
                                            service_data[service_name].extend(resources)
        
        return service_data
    
    def _flatten_service_result(self, result_data: Any, inventory: str, region: str, service: str) -> List[Dict[str, Any]]:
        """
        Flatten service result data into a list of resource dictionaries.
        
        Args:
            result_data: The service result data
            inventory: Inventory name
            region: Region name
            service: Service name
            
        Returns:
            List of flattened resource dictionaries
        """
        resources = []
        
        if isinstance(result_data, list):
            # Direct list of resources
            for i, resource in enumerate(result_data):
                if isinstance(resource, dict):
                    flattened = {
                        'Inventory': inventory,
                        'Region': region,
                        'Service': service,
                        **self._flatten_dict(resource)
                    }
                    resources.append(flattened)
                else:
                    # Non-dict resource
                    resources.append({
                        'Inventory': inventory,
                        'Region': region,
                        'Service': service,
                        'Resource': str(resource)
                    })
        elif isinstance(result_data, dict):
            # Handle special cases like EC2 reservations
            if service == "ec2" and "Reservations" in result_data:
                for reservation in result_data["Reservations"]:
                    if isinstance(reservation, dict) and "Instances" in reservation:
                        for instance in reservation["Instances"]:
                            if isinstance(instance, dict):
                                flattened = {
                                    'Inventory': inventory,
                                    'Region': region,
                                    'Service': service,
                                    **self._flatten_dict(instance)
                                }
                                resources.append(flattened)
            else:
                # Try to find list values in the dict
                for key, value in result_data.items():
                    if isinstance(value, list):
                        for item in value:
                            if isinstance(item, dict):
                                flattened = {
                                    'Inventory': inventory,
                                    'Region': region,
                                    'Service': service,
                                    **self._flatten_dict(item)
                                }
                                resources.append(flattened)
        
        return resources
    
    def _flatten_dict(self, d: Dict[str, Any], parent_key: str = '', sep: str = '_') -> Dict[str, Any]:
        """
        Flatten a nested dictionary.
        
        Args:
            d: Dictionary to flatten
            parent_key: Parent key for nested keys
            sep: Separator for nested keys
            
        Returns:
            Flattened dictionary
        """
        from datetime import datetime
        
        items = []
        for k, v in d.items():
            new_key = f"{parent_key}{sep}{k}" if parent_key else k
            if isinstance(v, dict):
                items.extend(self._flatten_dict(v, new_key, sep=sep).items())
            elif isinstance(v, list):
                # Convert lists to strings for Excel compatibility
                items.append((new_key, str(v)))
            elif isinstance(v, datetime):
                # Convert datetime to string for Excel compatibility
                items.append((new_key, str(v)))
            else:
                # Convert any other complex objects to strings
                try:
                    # Try to use the value as-is first
                    items.append((new_key, v))
                except:
                    # If that fails, convert to string
                    items.append((new_key, str(v)))
        return dict(items)
    
    def _sanitize_sheet_name(self, name: str) -> str:
        """
        Sanitize sheet name for Excel compatibility.
        
        Args:
            name: Original sheet name
            
        Returns:
            Sanitized sheet name
        """
        # Excel sheet names cannot exceed 31 characters and cannot contain certain characters
        invalid_chars = ['\\', '/', '*', '[', ']', ':', '?']
        sanitized = name
        
        for char in invalid_chars:
            sanitized = sanitized.replace(char, '_')
        
        # Truncate to 31 characters
        if len(sanitized) > 31:
            sanitized = sanitized[:31]
        
        return sanitized