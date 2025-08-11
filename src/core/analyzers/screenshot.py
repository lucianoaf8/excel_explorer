"""
Screenshot Analyzer - Captures Excel sheets exactly as they appear in Excel
Uses xlwings to control Excel application and take screenshots
"""

import os
import time
from typing import Dict, Any, List, Optional
from pathlib import Path
import tempfile
from datetime import datetime

try:
    import xlwings as xw
    from PIL import Image
    import win32clipboard
    from io import BytesIO
    XLWINGS_AVAILABLE = True
except ImportError:
    XLWINGS_AVAILABLE = False

from .base import BaseAnalyzer


class ScreenshotAnalyzer(BaseAnalyzer):
    """Captures screenshots of Excel sheets using actual Excel application"""
    
    def __init__(self, config: Optional[Dict[str, Any]] = None):
        """
        Initialize screenshot analyzer
        
        Args:
            config: Configuration dictionary
        """
        super().__init__(config)
        self.screenshot_config = config.get('screenshot', {}) if config else {}
        self.output_dir = None
        self.screenshots = []
        
    def analyze(self, workbook: Any) -> Dict[str, Any]:
        """
        Capture screenshots of all sheets in the workbook
        
        Args:
            workbook: openpyxl workbook (we'll get the file path from it)
            
        Returns:
            Dictionary containing screenshot metadata and paths
        """
        start_time = time.time()
        
        # Check if screenshots are enabled
        if not self.screenshot_config.get('enabled', False):
            return {
                'status': 'disabled',
                'message': 'Screenshot capture is disabled in configuration',
                'analysis_duration': 0
            }
        
        if not XLWINGS_AVAILABLE:
            return {
                'error': 'xlwings not installed. Install with: pip install xlwings pillow pywin32',
                'analysis_duration': 0,
                'status': 'skipped'
            }
        
        # For this analyzer, we need the actual file path
        # This should be passed through the orchestrator
        file_path = self.screenshot_config.get('file_path')
        if not file_path:
            return {
                'error': 'File path not provided for screenshot capture',
                'analysis_duration': 0,
                'status': 'failed'
            }
        
        try:
            # Create output directory for screenshots
            self._create_output_directory(file_path)
            
            # Open Excel application (hidden by default)
            app = xw.App(visible=self.screenshot_config.get('show_excel', False))
            app.display_alerts = False
            
            try:
                # Open the workbook in Excel
                wb = app.books.open(file_path)
                
                # Capture each sheet
                sheet_screenshots = []
                for sheet in wb.sheets:
                    screenshot_info = self._capture_sheet(sheet)
                    if screenshot_info:
                        sheet_screenshots.append(screenshot_info)
                
                # Close workbook without saving
                wb.close()
                
            finally:
                # Quit Excel application
                app.quit()
            
            return {
                'total_sheets_captured': len(sheet_screenshots),
                'screenshots': sheet_screenshots,
                'output_directory': str(self.output_dir),
                'analysis_duration': time.time() - start_time,
                'status': 'success'
            }
            
        except Exception as e:
            return {
                'error': f'Screenshot capture failed: {str(e)}',
                'analysis_duration': time.time() - start_time,
                'status': 'failed'
            }
    
    def _create_output_directory(self, file_path: str):
        """Create organized output directory for screenshots"""
        file_name = Path(file_path).stem
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Use configured output path or default
        base_output = self.screenshot_config.get('output_dir', 'output/screenshots')
        self.output_dir = Path(base_output) / f"{file_name}_{timestamp}"
        self.output_dir.mkdir(parents=True, exist_ok=True)
    
    def _capture_sheet(self, sheet) -> Optional[Dict[str, Any]]:
        """
        Capture a single sheet as screenshot
        
        Args:
            sheet: xlwings sheet object
            
        Returns:
            Dictionary with screenshot information
        """
        try:
            sheet_name = sheet.name
            safe_name = "".join(c if c.isalnum() or c in ('_', '-') else '_' for c in sheet_name)
            
            # Activate the sheet
            sheet.activate()
            time.sleep(0.5)  # Brief pause to ensure sheet is rendered
            
            # Determine the used range
            used_range = sheet.used_range
            if not used_range:
                return None
            
            # Configure capture area
            capture_range = self._get_capture_range(sheet, used_range)
            
            # Use CopyPicture to capture exact appearance
            # Appearance: 1 = xlScreen (as shown on screen)
            # Format: 2 = xlBitmap
            capture_range.api.CopyPicture(Appearance=1, Format=2)
            
            # Get image from clipboard
            image = self._get_image_from_clipboard()
            
            if image:
                # Save the image
                output_path = self.output_dir / f"{safe_name}.png"
                image.save(str(output_path), 'PNG', quality=95)
                
                return {
                    'sheet_name': sheet_name,
                    'file_path': str(output_path),
                    'width': image.width,
                    'height': image.height,
                    'capture_area': str(capture_range.address)
                }
            
        except Exception as e:
            print(f"Failed to capture sheet {sheet.name}: {e}")
            return None
    
    def _get_capture_range(self, sheet, used_range):
        """
        Determine the range to capture based on configuration
        
        Args:
            sheet: xlwings sheet object
            used_range: The used range of the sheet
            
        Returns:
            Range object to capture
        """
        capture_mode = self.screenshot_config.get('capture_mode', 'used_range')
        
        if capture_mode == 'full_sheet':
            # Capture a larger predefined area
            return sheet.range('A1:Z100')
        elif capture_mode == 'print_area':
            # Try to use print area if defined
            try:
                print_area = sheet.api.PageSetup.PrintArea
                if print_area:
                    return sheet.range(print_area)
            except:
                pass
        
        # Default to used range
        return used_range
    
    def _get_image_from_clipboard(self) -> Optional[Image.Image]:
        """
        Retrieve image from Windows clipboard
        
        Returns:
            PIL Image object or None
        """
        try:
            win32clipboard.OpenClipboard()
            try:
                # Check if clipboard contains bitmap
                if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_DIB):
                    data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
                    # Convert DIB to PIL Image
                    return self._dib_to_image(data)
            finally:
                win32clipboard.CloseClipboard()
        except Exception as e:
            print(f"Failed to get image from clipboard: {e}")
            return None
    
    def _dib_to_image(self, dib_data) -> Image.Image:
        """
        Convert DIB (Device Independent Bitmap) data to PIL Image
        
        Args:
            dib_data: DIB data from clipboard
            
        Returns:
            PIL Image object
        """
        # This is a simplified version - you might need to handle different DIB formats
        # For now, we'll use a more robust approach with win32ui
        import win32ui
        import win32con
        
        # Create a device context
        hdc = win32ui.CreateDCFromHandle(win32ui.GetActiveWindow().GetDC())
        
        # Create a compatible DC and bitmap
        dc = hdc.CreateCompatibleDC()
        
        # For simplicity, let's try another approach using PIL directly
        # This assumes the clipboard has been properly formatted
        win32clipboard.OpenClipboard()
        try:
            from PIL import ImageGrab
            img = ImageGrab.grabclipboard()
            return img
        finally:
            win32clipboard.CloseClipboard()


class ScreenshotUtility:
    """Utility class for standalone screenshot operations"""
    
    @staticmethod
    def capture_excel_file(file_path: str, output_dir: Optional[str] = None, 
                          show_excel: bool = False) -> Dict[str, Any]:
        """
        Standalone function to capture screenshots of an Excel file
        
        Args:
            file_path: Path to Excel file
            output_dir: Output directory for screenshots
            show_excel: Whether to show Excel window during capture
            
        Returns:
            Dictionary with capture results
        """
        config = {
            'screenshot': {
                'file_path': file_path,
                'output_dir': output_dir or 'output/screenshots',
                'show_excel': show_excel,
                'capture_mode': 'used_range'
            }
        }
        
        analyzer = ScreenshotAnalyzer(config)
        # Pass a dummy workbook since we're using file_path from config
        return analyzer.analyze(None)
    
    @staticmethod
    def capture_specific_range(file_path: str, sheet_name: str, 
                              range_address: str, output_path: str) -> bool:
        """
        Capture a specific range from a sheet
        
        Args:
            file_path: Path to Excel file
            sheet_name: Name of the sheet
            range_address: Excel range (e.g., 'A1:D10')
            output_path: Where to save the screenshot
            
        Returns:
            True if successful
        """
        if not XLWINGS_AVAILABLE:
            print("xlwings not available")
            return False
        
        try:
            app = xw.App(visible=False)
            app.display_alerts = False
            
            try:
                wb = app.books.open(file_path)
                sheet = wb.sheets[sheet_name]
                
                # Capture specific range
                target_range = sheet.range(range_address)
                target_range.api.CopyPicture(Appearance=1, Format=2)
                
                # Get from clipboard
                from PIL import ImageGrab
                img = ImageGrab.grabclipboard()
                
                if img:
                    img.save(output_path, 'PNG')
                    return True
                    
            finally:
                wb.close()
                app.quit()
                
        except Exception as e:
            print(f"Error capturing range: {e}")
            return False
        
        return False