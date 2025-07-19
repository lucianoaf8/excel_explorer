"""
Modern Excel Explorer GUI with circular progress and embedded reports
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import json
import time
import math
import sys
import os
from pathlib import Path
from datetime import datetime, timedelta
import webbrowser
from typing import Dict, Any, Optional
from PIL import Image, ImageTk

from src.core import SimpleExcelAnalyzer
from src.reports import ReportGenerator
from src.reports.structured_text_report import StructuredTextReportGenerator
from src.reports.comprehensive_text_report import ComprehensiveTextReportGenerator


class CircularProgress(tk.Canvas):
    """Custom circular progress indicator"""
    
    def __init__(self, parent, size=120, thickness=8, color="#2E86AB", bg_color="#E0E0E0"):
        super().__init__(parent, width=size, height=size, highlightthickness=0)
        self.size = size
        self.thickness = thickness
        self.color = color
        self.bg_color = bg_color
        self.progress = 0.0
        self.center = size // 2
        self.radius = (size - thickness) // 2
        self.configure(bg='white')
        self.draw_progress()
    
    def set_progress(self, progress: float):
        """Set progress value (0.0 to 1.0)"""
        self.progress = max(0.0, min(1.0, progress))
        self.draw_progress()
    
    def draw_progress(self):
        """Draw the circular progress indicator"""
        self.delete("all")
        
        # Background circle
        self.create_oval(
            self.center - self.radius, self.center - self.radius,
            self.center + self.radius, self.center + self.radius,
            outline=self.bg_color, width=self.thickness, fill=""
        )
        
        # Progress arc
        if self.progress > 0:
            extent = 360 * self.progress
            self.create_arc(
                self.center - self.radius, self.center - self.radius,
                self.center + self.radius, self.center + self.radius,
                start=90, extent=-extent, outline=self.color,
                width=self.thickness, style="arc"
            )
        
        # Center text
        percentage = int(self.progress * 100)
        self.create_text(
            self.center, self.center,
            text=f"{percentage}%",
            font=("Segoe UI", 14, "bold"),
            fill="#333333"
        )


class ModernStyle:
    """Enhanced UI styling constants"""
    PRIMARY = "#2E86AB"
    SECONDARY = "#A23B72" 
    SUCCESS = "#4CAF50"
    WARNING = "#FF9800"
    ERROR = "#F44336"
    BACKGROUND = "#F8F9FA"
    SURFACE = "#FFFFFF"
    TEXT_PRIMARY = "#212121"
    TEXT_SECONDARY = "#757575"
    
    FONT_TITLE = ("Segoe UI", 18, "bold")
    FONT_HEADING = ("Segoe UI", 12, "bold")
    FONT_BODY = ("Segoe UI", 10)
    FONT_MONO = ("Consolas", 9)


class ProgressTracker:
    """Enhanced progress tracking with sub-steps"""
    
    def __init__(self, progress_var: tk.StringVar, detail_var: tk.StringVar, 
                 timer_var: tk.StringVar, circular_progress: CircularProgress):
        self.progress_var = progress_var
        self.detail_var = detail_var
        self.timer_var = timer_var
        self.circular_progress = circular_progress
        
        self.modules = [
            'health_checker', 'structure_mapper', 'data_profiler', 'formula_analyzer',
            'visual_cataloger', 'connection_inspector', 'pivot_intelligence', 'doc_synthesizer'
        ]
        self.total_steps = len(self.modules) * 4  # 4 sub-steps per module
        self.completed_steps = 0
        self.current_module_index = 0
        self.start_time = None
        
    def start_analysis(self):
        """Start the analysis timer"""
        self.start_time = time.time()
        self.completed_steps = 0
        self.current_module_index = 0
        self.circular_progress.set_progress(0.0)
        
    def start_module(self, module_name: str, description: str):
        """Signal start of module execution"""
        self.current_module_index = self.modules.index(module_name) if module_name in self.modules else 0
        self.progress_var.set(f"🔍 {module_name.replace('_', ' ').title()}")
        self.detail_var.set(description)
        self._update_progress()
        
    def update_step(self, module_name: str, step_description: str):
        """Update sub-step progress"""
        self.detail_var.set(f"  ↳ {step_description}")
        self.completed_steps += 1
        self._update_progress()
        
    def complete_module(self, module_name: str, success: bool = True):
        """Signal completion of module"""
        status = "✅" if success else "❌"
        self.detail_var.set(f"{status} {module_name.replace('_', ' ').title()} completed")
        # Ensure module is fully completed (4 steps)
        target_steps = (self.current_module_index + 1) * 4
        self.completed_steps = max(self.completed_steps, target_steps)
        self._update_progress()
        
    def _update_progress(self):
        """Update progress indicator and timer"""
        # Update circular progress
        progress = self.completed_steps / self.total_steps
        self.circular_progress.set_progress(progress)
        
        # Update timer
        if self.start_time:
            elapsed = time.time() - self.start_time
            self.timer_var.set(f"⏱️ {self._format_time(elapsed)}")
    
    def _format_time(self, seconds: float) -> str:
        """Format elapsed time"""
        if seconds < 60:
            return f"{seconds:.1f}s"
        elif seconds < 3600:
            return f"{int(seconds // 60)}m {int(seconds % 60)}s"
        else:
            hours = int(seconds // 3600)
            minutes = int((seconds % 3600) // 60)
            return f"{hours}h {minutes}m"
    
    def set_complete(self):
        """Signal analysis completion"""
        self.completed_steps = self.total_steps
        self.circular_progress.set_progress(1.0)
        self.progress_var.set("🎉 Analysis Complete!")
        self.detail_var.set("All modules executed successfully")
        
    def set_error(self, message: str):
        """Display error state"""
        self.progress_var.set("❌ Analysis Failed")
        self.detail_var.set(f"Error: {message}")


class ExcelExplorerApp:
    """Enhanced application with circular progress and embedded reports"""
    
    def __init__(self, root: tk.Tk):
        self.root = root
        self.setup_window()
        self.setup_variables()
        self.setup_ui()
        self.explorer = None
        self.analysis_thread = None
        self.current_results = None
        self.timer_thread = None
        
    def setup_window(self):
        """Configure main window"""
        self.root.title("Excel Explorer - Advanced Analysis Tool")
        self.root.geometry("1400x900")
        self.root.minsize(1000, 700)
        self.root.configure(bg=ModernStyle.BACKGROUND)
        
        # Set window icon for taskbar and title bar
        try:
            # Get paths to both logo versions
            logo_no_bg_path = Path(__file__).parent.parent.parent / "assets" / "logos" / "logo_1_no_bg.png"
            logo_with_bg_path = Path(__file__).parent.parent.parent / "assets" / "logos" / "icon_bg.png"
            
            if logo_no_bg_path.exists():
                # Windows-specific: Set application ID for taskbar grouping
                if sys.platform == "win32":
                    try:
                        import ctypes
                        # Set application user model ID for proper taskbar icon
                        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("ExcelExplorer.App.1.0")
                    except Exception:
                        pass
                
                # Load the no-background logo for title bar and GUI
                icon_image_no_bg = Image.open(logo_no_bg_path)
                
                # Create multiple sizes for title bar (no background)
                icon_sizes = [16, 32, 48, 64]
                self.icons = []
                
                for size in icon_sizes:
                    resized_icon = icon_image_no_bg.resize((size, size), Image.Resampling.LANCZOS)
                    photo_icon = ImageTk.PhotoImage(resized_icon)
                    self.icons.append(photo_icon)
                
                # Set the no-background icon for title bar
                self.root.iconphoto(True, *self.icons)
                
                # Windows-specific taskbar icon handling with background
                if sys.platform == "win32" and logo_with_bg_path.exists():
                    try:
                        # Use the background version for taskbar
                        taskbar_icon_image = Image.open(logo_with_bg_path)
                        
                        # Convert to ICO format for Windows taskbar
                        import tempfile
                        with tempfile.NamedTemporaryFile(suffix='.ico', delete=False) as tmp_ico:
                            # Create ICO with multiple sizes using background version
                            taskbar_icon_image.save(tmp_ico.name, format='ICO', sizes=[(16,16), (32,32), (48,48), (64,64)])
                            self.root.iconbitmap(tmp_ico.name)
                            # Clean up temp file after a delay
                            self.root.after(1000, lambda: self._cleanup_temp_file(tmp_ico.name))
                    except Exception as e:
                        print(f"Windows taskbar ICO handling failed: {e}")
                        # Fallback to no-background version for taskbar
                        try:
                            import tempfile
                            with tempfile.NamedTemporaryFile(suffix='.ico', delete=False) as tmp_ico:
                                icon_image_no_bg.save(tmp_ico.name, format='ICO', sizes=[(16,16), (32,32), (48,48), (64,64)])
                                self.root.iconbitmap(tmp_ico.name)
                                self.root.after(1000, lambda: self._cleanup_temp_file(tmp_ico.name))
                        except Exception as fallback_e:
                            print(f"Fallback ICO handling also failed: {fallback_e}")
                        
        except Exception as e:
            print(f"Failed to set window icon: {e}")
        
        # Center window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (1400 // 2)
        y = (self.root.winfo_screenheight() // 2) - (900 // 2)
        self.root.geometry(f"1400x900+{x}+{y}")
    
    def _cleanup_temp_file(self, filepath):
        """Clean up temporary ICO file"""
        try:
            if os.path.exists(filepath):
                os.unlink(filepath)
        except Exception:
            pass
        
    def setup_variables(self):
        """Initialize tkinter variables"""
        self.selected_file = tk.StringVar()
        self.progress_text = tk.StringVar(value="Ready to analyze Excel files")
        self.progress_detail = tk.StringVar(value="Select a file to begin analysis")
        self.timer_text = tk.StringVar(value="⏱️ 0.0s")
        self.analysis_running = tk.BooleanVar(value=False)
        # Path to the last auto-exported report
        self.auto_report_path: Optional[str] = None
        
    def setup_ui(self):
        """Create enhanced UI layout"""
        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="25")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header section
        self.create_header(main_frame)
        
        # Content area with two columns
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(20, 0))
        
        # Left column - File selection and progress
        left_frame = ttk.Frame(content_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 15))
        
        self.create_file_section(left_frame)
        self.create_progress_section(left_frame)
        self.create_action_buttons(left_frame)
        
        # Right column - Tabbed interface
        right_frame = ttk.Frame(content_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        self.create_tabbed_interface(right_frame)
        
    def create_header(self, parent):
        """Create enhanced header"""
        header_frame = ttk.Frame(parent)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Title with logo
        title_frame = ttk.Frame(header_frame)
        title_frame.pack(anchor=tk.W)
        
        # Add logo image
        try:
            logo_path = Path(__file__).parent.parent.parent / "assets" / "logos" / "logo_1_no_bg.png"
            if logo_path.exists():
                # Load and resize logo for header
                logo_image = Image.open(logo_path)
                # Resize to fit header (keeping aspect ratio)
                logo_image = logo_image.resize((80, 80), Image.Resampling.LANCZOS)
                self.logo = ImageTk.PhotoImage(logo_image)
                
                # Display logo
                logo_label = ttk.Label(title_frame, image=self.logo)
                logo_label.pack(side=tk.LEFT, padx=(0, 15))
        except Exception as e:
            print(f"Failed to load logo: {e}")
        
        # Title text (without emoji since we have the actual logo)
        title_container = ttk.Frame(title_frame)
        title_container.pack(side=tk.LEFT, fill=tk.Y)
        
        title_label = ttk.Label(
            title_container, 
            text="Excel Explorer",
            font=ModernStyle.FONT_TITLE,
            foreground=ModernStyle.PRIMARY
        )
        title_label.pack(anchor=tk.W)
        
        # Subtitle
        subtitle_label = ttk.Label(
            title_container,
            text="Advanced Excel file analysis with real-time progress tracking and comprehensive reporting",
            font=ModernStyle.FONT_BODY,
            foreground=ModernStyle.TEXT_SECONDARY
        )
        subtitle_label.pack(anchor=tk.W, pady=(5, 0))
        
    def create_file_section(self, parent):
        """Create file selection UI"""
        file_frame = ttk.LabelFrame(parent, text="📁 File Selection", padding="20")
        file_frame.pack(fill=tk.X, pady=(0, 20))
        
        # File path display with better styling
        ttk.Label(file_frame, text="Selected File:", font=ModernStyle.FONT_BODY).pack(anchor=tk.W)
        
        path_frame = ttk.Frame(file_frame)
        path_frame.pack(fill=tk.X, pady=(8, 15))
        
        self.file_entry = ttk.Entry(
            path_frame, 
            textvariable=self.selected_file,
            font=ModernStyle.FONT_BODY,
            state="readonly"
        )
        self.file_entry.pack(fill=tk.X)
        
        # Browse button with icon
        select_btn = ttk.Button(
            file_frame,
            text="🗂️ Browse Excel Files",
            command=self.select_file,
            style="Accent.TButton"
        )
        select_btn.pack(anchor=tk.W)
        
    def create_progress_section(self, parent):
        """Create enhanced progress section with circular indicator"""
        progress_frame = ttk.LabelFrame(parent, text="📈 Analysis Progress", padding="20")
        progress_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Progress indicator container
        indicator_frame = ttk.Frame(progress_frame)
        indicator_frame.pack(pady=(0, 20))
        
        # Circular progress indicator
        self.circular_progress = CircularProgress(indicator_frame, size=100, thickness=8)
        self.circular_progress.pack()
        
        # Timer display
        timer_label = ttk.Label(
            indicator_frame,
            textvariable=self.timer_text,
            font=ModernStyle.FONT_HEADING,
            foreground=ModernStyle.PRIMARY
        )
        timer_label.pack(pady=(10, 0))
        
        # Status text
        status_label = ttk.Label(
            progress_frame,
            textvariable=self.progress_text,
            font=ModernStyle.FONT_HEADING,
            foreground=ModernStyle.TEXT_PRIMARY
        )
        status_label.pack(anchor=tk.W)
        
        # Detail text
        detail_label = ttk.Label(
            progress_frame,
            textvariable=self.progress_detail,
            font=ModernStyle.FONT_BODY,
            foreground=ModernStyle.TEXT_SECONDARY,
            wraplength=280
        )
        detail_label.pack(anchor=tk.W, pady=(5, 0))
        
    def create_action_buttons(self, parent):
        """Create action button panel"""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Analyze button
        self.analyze_btn = ttk.Button(
            button_frame,
            text="🚀 Start Analysis",
            command=self.start_analysis,
            style="Accent.TButton"
        )
        self.analyze_btn.pack(fill=tk.X, pady=(0, 10))
        
        # Stop button
        self.stop_btn = ttk.Button(
            button_frame,
            text="⏹️ Stop Analysis",
            command=self.stop_analysis,
            state=tk.DISABLED
        )
        self.stop_btn.pack(fill=tk.X, pady=(0, 10))
        
        # Clear logs button
        clear_btn = ttk.Button(
            button_frame,
            text="🗑️ Clear Logs",
            command=self.clear_logs
        )
        clear_btn.pack(fill=tk.X)
        
    def create_tabbed_interface(self, parent):
        """Create enhanced tabbed interface"""
        self.notebook = notebook = ttk.Notebook(parent)
        notebook.pack(fill=tk.BOTH, expand=True)
        
        # Analysis logs tab
        log_frame = ttk.Frame(notebook)
        notebook.add(log_frame, text="📋 Analysis Logs")
        
        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            font=ModernStyle.FONT_MONO,
            height=20
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # Results preview tab
        results_frame = ttk.Frame(notebook)
        notebook.add(results_frame, text="📊 Results Summary")
        
        self.results_text = scrolledtext.ScrolledText(
            results_frame,
            wrap=tk.WORD,
            font=ModernStyle.FONT_MONO,
            height=20
        )
        self.results_text.pack(fill=tk.BOTH, expand=True, padx=15, pady=15)
        
        # Report display tab
        self.report_frame = report_frame = ttk.Frame(notebook)
        notebook.add(report_frame, text="📄 Analysis Report")
        
        # Search frame
        search_frame = ttk.Frame(report_frame)
        search_frame.pack(fill=tk.X, padx=15, pady=(15, 5))
        
        ttk.Label(search_frame, text="🔍 Search:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        self.search_entry.pack(side=tk.LEFT, padx=(5, 5))
        self.search_entry.bind('<KeyRelease>', self._on_search_change)
        
        search_btn = ttk.Button(search_frame, text="Find", command=self._search_report)
        search_btn.pack(side=tk.LEFT, padx=(0, 5))
        
        clear_search_btn = ttk.Button(search_frame, text="Clear", command=self._clear_search)
        clear_search_btn.pack(side=tk.LEFT)
        
        # Report will be populated after analysis
        self.report_text = scrolledtext.ScrolledText(
            report_frame,
            wrap=tk.WORD,
            font=ModernStyle.FONT_BODY,
            height=20
        )
        self.report_text.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)
        
        # Action buttons in report tab
        export_frame = ttk.Frame(report_frame)
        export_frame.pack(fill=tk.X, padx=15, pady=(0, 15))

        self.export_text_btn = ttk.Button(
            export_frame,
            text="📝 Export Text Report",
            command=self.export_text_report,
            state=tk.DISABLED
        )
        self.export_text_btn.pack(side=tk.RIGHT)

        self.export_markdown_btn = ttk.Button(
            export_frame,
            text="📄 Export Markdown",
            command=self.export_markdown_report,
            state=tk.DISABLED
        )
        self.export_markdown_btn.pack(side=tk.RIGHT, padx=(0, 5))

        self.open_report_btn = ttk.Button(
            export_frame,
            text="🌐 Open HTML Report",
            command=self.open_last_report,
            state=tk.DISABLED
        )
        self.open_report_btn.pack(side=tk.RIGHT, padx=(0, 5))

        self.export_btn = ttk.Button(
            export_frame,
            text="💾 Export HTML Report",
            command=self.export_report,
            state=tk.DISABLED
        )
        self.export_btn.pack(side=tk.RIGHT, padx=(0, 5))
        
    def select_file(self):
        """Open file selection dialog"""
        filetypes = [
            ('Excel files', '*.xlsx *.xls *.xlsm'),
            ('All files', '*.*')
        ]
        
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=filetypes
        )
        
        if filename:
            self.selected_file.set(filename)
            self.log_message(f"📁 Selected file: {Path(filename).name}")
            
    def start_analysis(self):
        """Start Excel file analysis with enhanced progress tracking"""
        if not self.selected_file.get():
            messagebox.showwarning("No File Selected", "Please select an Excel file first.")
            return
            
        if not Path(self.selected_file.get()).exists():
            messagebox.showerror("File Not Found", "The selected file does not exist.")
            return
            
        # UI state for analysis
        self.analysis_running.set(True)
        self.analyze_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.DISABLED)
        self.export_text_btn.config(state=tk.DISABLED)
        self.export_markdown_btn.config(state=tk.DISABLED)
        
        # Clear previous results
        self.results_text.delete(1.0, tk.END)
        self.report_text.delete(1.0, tk.END)
        self.current_results = None
        
        # Initialize progress tracker
        self.progress_tracker = ProgressTracker(
            self.progress_text, self.progress_detail, 
            self.timer_text, self.circular_progress
        )
        self.progress_tracker.start_analysis()
        
        # Start timer thread
        self.start_timer()
        
        # Start analysis thread
        self.analysis_thread = threading.Thread(target=self._run_analysis, daemon=True)
        self.analysis_thread.start()
        
    def start_timer(self):
        """Start the timer thread"""
        self.timer_thread = threading.Thread(target=self._update_timer, daemon=True)
        self.timer_thread.start()
        
    def _update_timer(self):
        """Update timer every second"""
        start_time = time.time()
        while self.analysis_running.get():
            elapsed = time.time() - start_time
            self.root.after(0, lambda: self.timer_text.set(f"⏱️ {self._format_time(elapsed)}"))
            time.sleep(1)
            
    def _format_time(self, seconds: float) -> str:
        """Format elapsed time"""
        if seconds < 60:
            return f"{seconds:.1f}s"
        elif seconds < 3600:
            return f"{int(seconds // 60)}m {int(seconds % 60)}s"
        else:
            hours = int(seconds // 3600)
            minutes = int((seconds % 3600) // 60)
            return f"{hours}h {minutes}m"
        
    def _run_analysis(self):
        """Run analysis in background thread"""
        try:
            # Initialize explorer
            self.log_message("🔧 Initializing Excel Explorer...")
            self.explorer = SimpleExcelAnalyzer()
            
            # Run analysis with progress tracking
            file_path = self.selected_file.get()
            self.log_message(f"🚀 Starting analysis of {Path(file_path).name}")
            
            # Execute analysis
            results = self.explorer.analyze(file_path, progress_callback=self._progress_callback)
            
            # Analysis complete
            self.current_results = results
            self._analysis_complete(results)
            
        except Exception as e:
            self._analysis_error(str(e))
            
    def _progress_callback(self, module_name: str, status: str, detail: str = ""):
        """Handle progress updates from analysis"""
        self.root.after(0, lambda: self._update_progress(module_name, status, detail))
        
    def _update_progress(self, module_name: str, status: str, detail: str):
        """Update progress UI on main thread"""
        if status == "starting":
            self.progress_tracker.start_module(module_name, detail)
            self.log_message(f"🔍 Starting {module_name.replace('_', ' ').title()}: {detail}")
        elif status == "step":
            self.progress_tracker.update_step(module_name, detail)
            self.log_message(f"  ↳ {detail}")
        elif status == "complete":
            self.progress_tracker.complete_module(module_name, True)
            self.log_message(f"✅ Completed {module_name.replace('_', ' ').title()}")
        elif status == "error":
            self.progress_tracker.complete_module(module_name, False)
            self.log_message(f"❌ ERROR in {module_name}: {detail}")
            
    def _analysis_complete(self, results: Dict[str, Any]):
        """Handle successful analysis completion"""
        self.root.after(0, lambda: self._finalize_analysis(results, success=True))
        
    def _analysis_error(self, error_message: str):
        """Handle analysis error"""
        self.root.after(0, lambda: self._finalize_analysis(error_message, success=False))
        
    def _finalize_analysis(self, data, success: bool):
        """Finalize analysis on main thread"""
        # Update UI state
        self.analysis_running.set(False)
        self.analyze_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        
        if success:
            # Complete progress
            self.progress_tracker.set_complete()
            
            # Enable export
            self.export_btn.config(state=tk.NORMAL)
            self.export_text_btn.config(state=tk.NORMAL)
            self.export_markdown_btn.config(state=tk.NORMAL)
            
            # Convert AnalysisResults to dict if needed
            if hasattr(data, 'file_info'):  # It's an AnalysisResults object
                results_dict = {
                    'file_info': data.file_info,
                    'analysis_metadata': data.analysis_metadata,
                    'module_results': data.module_results,
                    'execution_summary': data.execution_summary,
                    'resource_usage': data.resource_usage,
                    'recommendations': data.recommendations
                }
            else:
                results_dict = data
            
            # In _finalize_analysis, validate metrics before display
            if success and isinstance(data, dict):
                exec_summary = results_dict.get('execution_summary', {})
                if exec_summary.get('total_modules', 0) == 0:
                    self.log_message("⚠️ Warning: No module execution data found")
            
            # Display results summary
            summary = self._create_results_summary(results_dict)
            self.results_text.insert(tk.END, summary)
            
            # Generate and display report
            self._display_embedded_report(results_dict)
            
            # Auto-export full HTML report
            self._auto_export_report(results_dict)
            
            # Switch to report tab automatically
            if hasattr(self, 'notebook') and hasattr(self, 'report_frame'):
                self.notebook.select(self.report_frame)
            
            self.log_message("🎉 Analysis completed successfully!")
            messagebox.showinfo("Analysis Complete", "Excel file analysis completed successfully!")
            
        else:
            # Display error
            self.progress_tracker.set_error(data)
            self.log_message(f"❌ Analysis failed: {data}")
            messagebox.showerror("Analysis Failed", f"Analysis failed: {data}")
            
    def _display_embedded_report(self, results: Dict[str, Any]):
        """Display structured text report in the GUI tab"""
        try:
            # Use the new structured text report generator
            text_report_generator = StructuredTextReportGenerator()
            report_text = text_report_generator.generate_report(results)
            
            self.report_text.delete(1.0, tk.END)
            self.report_text.insert(tk.END, report_text)
            
        except Exception as e:
            self.log_message(f"⚠️ Report display failed: {e}")
            self.report_text.delete(1.0, tk.END)
            self.report_text.insert(tk.END, "Report generation failed. Use Export HTML Report for full report.")
            
    def _create_text_report(self, results: Dict[str, Any]) -> str:
        """Create text-based report for GUI display"""
        report = []
        report.append("═" * 80)
        report.append("📊 EXCEL ANALYSIS REPORT")
        report.append("═" * 80)
        report.append("")
        
        # File information
        if 'file_info' in results:
            file_info = results['file_info']
            report.append("📁 FILE INFORMATION")
            report.append("-" * 50)
            report.append(f"Name: {file_info.get('name', 'Unknown')}")
            report.append(f"Size: {file_info.get('size_mb', 0):.2f} MB")
            report.append(f"Path: {file_info.get('path', 'Unknown')}")
            report.append("")
        
        # Analysis metadata
        if 'analysis_metadata' in results:
            metadata = results['analysis_metadata']
            report.append("⏱️ ANALYSIS SUMMARY")
            report.append("-" * 50)
            report.append(f"Duration: {metadata.get('total_duration_seconds', 0):.1f} seconds")
            report.append(f"Success Rate: {metadata.get('success_rate', 0):.1%}")
            report.append(f"Quality Score: {metadata.get('quality_score', 0):.1%}")
            timestamp = datetime.fromtimestamp(metadata.get('timestamp', time.time()))
            report.append(f"Completed: {timestamp.strftime('%Y-%m-%d %H:%M:%S')}")
            report.append("")
        
        # Module execution summary
        if 'execution_summary' in results:
            exec_summary = results['execution_summary']
            report.append("🔧 MODULE EXECUTION")
            report.append("-" * 50)
            report.append(f"Total Modules: {exec_summary.get('total_modules', 0)}")
            report.append(f"Successful: {exec_summary.get('successful_modules', 0)}")
            report.append(f"Failed: {exec_summary.get('failed_modules', 0)}")
            
            statuses = exec_summary.get('module_statuses', {})
            for module, status in statuses.items():
                icon = "✅" if status == "success" else "❌"
                report.append(f"  {icon} {module.replace('_', ' ').title()}")
            report.append("")
        
        # Key findings from modules
        if 'module_results' in results:
            module_results = results['module_results']
            report.append("🔍 KEY FINDINGS")
            report.append("-" * 50)
            
            # Health checker findings
            if 'health_checker' in module_results:
                health = module_results['health_checker']
                report.append("Health Check:")
                if hasattr(health, 'corruption_detected'):
                    report.append(f"  • Corruption Detected: {'Yes' if health.corruption_detected else 'No'}")
                if hasattr(health, 'security_issues'):
                    report.append(f"  • Security Issues: {len(health.security_issues) if health.security_issues else 0}")
                
            # Structure findings
            if 'structure_mapper' in module_results:
                structure = module_results['structure_mapper']
                report.append("Structure Analysis:")
                if hasattr(structure, 'total_sheets'):
                    report.append(f"  • Total Sheets: {structure.total_sheets}")
                if hasattr(structure, 'total_cells_with_data'):
                    report.append(f"  • Cells with Data: {structure.total_cells_with_data:,}")
                    
            # Data profiling findings
            if 'data_profiler' in module_results:
                data_profile = module_results['data_profiler']
                report.append("Data Analysis:")
                if hasattr(data_profile, 'total_regions'):
                    report.append(f"  • Data Regions: {data_profile.total_regions}")
                if hasattr(data_profile, 'data_quality_score'):
                    report.append(f"  • Data Quality: {data_profile.data_quality_score:.1%}")
            
            report.append("")
        
        # Recommendations
        if 'recommendations' in results and results['recommendations']:
            report.append("🎯 RECOMMENDATIONS")
            report.append("-" * 50)
            for i, rec in enumerate(results['recommendations'], 1):
                report.append(f"{i}. {rec}")
            report.append("")
        
        # Resource usage
        if 'resource_usage' in results:
            resource = results['resource_usage']
            current = resource.get('current_usage', {})
            report.append("💾 RESOURCE USAGE")
            report.append("-" * 50)
            report.append(f"Memory Used: {current.get('current_mb', 0):.1f} MB")
            report.append(f"Peak Memory: {current.get('peak_mb', 0):.1f} MB")
            report.append(f"CPU Usage: {current.get('cpu_percent', 0):.1f}%")
            report.append("")
        
        report.append("=" * 80)
        report.append("For detailed technical analysis, export the full HTML report.")
        report.append("=" * 80)
        
        return "\n".join(report)
    
    def _create_results_summary(self, results: Dict[str, Any]) -> str:
        """Create summary for results tab"""
        summary = []
        summary.append("🎯 QUICK ANALYSIS SUMMARY")
        summary.append("=" * 60)
        summary.append("")
        
        # Key metrics
        if 'analysis_metadata' in results:
            metadata = results['analysis_metadata']
            summary.append(f"⏱️ Analysis Time: {metadata.get('total_duration_seconds', 0):.1f} seconds")
            summary.append(f"✅ Success Rate: {metadata.get('success_rate', 0):.1%}")
            summary.append(f"⭐ Quality Score: {metadata.get('quality_score', 0):.1%}")
            summary.append("")
        
        # File stats
        if 'file_info' in results:
            file_info = results['file_info']
            summary.append(f"📁 File: {file_info.get('name', 'Unknown')}")
            summary.append(f"📏 Size: {file_info.get('size_mb', 0):.2f} MB")
            summary.append("")
        
        # Module status
        if 'execution_summary' in results:
            exec_summary = results['execution_summary']
            summary.append("🔧 Module Results:")
            statuses = exec_summary.get('module_statuses', {})
            for module, status in statuses.items():
                icon = "✅" if status == "success" else "❌"
                summary.append(f"  {icon} {module.replace('_', ' ').title()}")
            summary.append("")
        
        summary.append("📄 Full report available in the 'Analysis Report' tab")
        
        return "\n".join(summary)
    
    def open_last_report(self):
        """Open the last auto-exported HTML report in default browser"""
        if self.auto_report_path and Path(self.auto_report_path).exists():
            webbrowser.open(f"file://{Path(self.auto_report_path).absolute()}")
        else:
            messagebox.showwarning("Report Not Found", "No exported report available to open.")

    def _auto_export_report(self, results: Dict[str, Any]):
        """Automatically export HTML report to output/reports folder"""
        try:
            # Use output/reports directory relative to project root
            project_root = Path(__file__).resolve().parent.parent.parent
            reports_dir = project_root / "output" / "reports"
            reports_dir.mkdir(parents=True, exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = Path(self.selected_file.get()).stem if self.selected_file.get() else "report"
            file_name = f"{base_name}_{timestamp}.html"
            output_path = reports_dir / file_name
            
            report_generator = ReportGenerator()
            report_generator.generate_html_report(results, str(output_path))
            self.auto_report_path = str(output_path)
            
            # Enable open button
            self.open_report_btn.config(state=tk.NORMAL)
            self.log_message(f"💾 Report automatically exported to: {output_path}")
            
        except Exception as e:
            self.log_message(f"⚠️ Auto-export failed: {e}")

    def export_report(self):
        """Export full HTML report"""
        if not self.current_results:
            messagebox.showwarning("No Results", "No analysis results available to export.")
            return
            
        try:
            # Ask user for save location
            output_file = filedialog.asksaveasfilename(
                title="Export Analysis Report",
                defaultextension=".html",
                filetypes=[
                    ('HTML files', '*.html'),
                    ('All files', '*.*')
                ]
            )
            
            if output_file:
                # Generate report
                self.log_message("📄 Generating HTML report...")
                report_generator = ReportGenerator()
                report_path = report_generator.generate_html_report(self.current_results, output_file)
                
                self.log_message(f"💾 Report exported to: {report_path}")
                
                # Ask if user wants to open report
                if messagebox.askyesno("Export Complete", f"Report exported successfully!\n\nOpen in browser?"):
                    webbrowser.open(f"file://{Path(report_path).absolute()}")
                    
        except Exception as e:
            self.log_message(f"❌ Export failed: {e}")
            messagebox.showerror("Export Error", f"Failed to export report: {e}")
    
    def export_text_report(self):
        """Export structured text report"""
        if not self.current_results:
            messagebox.showwarning("No Results", "No analysis results available to export.")
            return
            
        try:
            # Suggest filename based on analyzed file
            suggested_name = "report.txt"
            if self.selected_file.get():
                file_path = Path(self.selected_file.get())
                suggested_name = f"{file_path.stem}_report.txt"
            
            # Ask user for save location
            output_file = filedialog.asksaveasfilename(
                title="Export Text Report",
                defaultextension=".txt",
                initialfile=suggested_name,
                filetypes=[
                    ('Text files', '*.txt'),
                    ('All files', '*.*')
                ]
            )
            
            if output_file:
                # Use comprehensive text report generator
                text_report_generator = ComprehensiveTextReportGenerator()
                text_report_generator.generate_text_report(self.current_results, output_file)
                
                self.log_message(f"💾 Text report exported to: {output_file}")
                
                # Ask if user wants to open the file
                open_file = messagebox.askyesno(
                    "Export Complete", 
                    f"Text report exported successfully to:\n{output_file}\n\nWould you like to open the file?"
                )
                
                if open_file:
                    self._open_file_in_system(output_file)
                    
        except Exception as e:
            self.log_message(f"❌ Text export failed: {e}")
            messagebox.showerror("Export Error", f"Failed to export text report: {e}")
    
    def export_markdown_report(self):
        """Export structured markdown report"""
        if not self.current_results:
            messagebox.showwarning("No Results", "No analysis results available to export.")
            return
            
        try:
            # Suggest filename based on analyzed file
            suggested_name = "report.md"
            if self.selected_file.get():
                file_path = Path(self.selected_file.get())
                suggested_name = f"{file_path.stem}_report.md"
            
            # Ask user for save location
            output_file = filedialog.asksaveasfilename(
                title="Export Markdown Report",
                defaultextension=".md",
                initialfile=suggested_name,
                filetypes=[
                    ('Markdown files', '*.md'),
                    ('All files', '*.*')
                ]
            )
            
            if output_file:
                # Use comprehensive text report generator
                text_report_generator = ComprehensiveTextReportGenerator()
                text_report_generator.generate_markdown_report(self.current_results, output_file)
                
                self.log_message(f"💾 Markdown report exported to: {output_file}")
                
                # Ask if user wants to open the file
                open_file = messagebox.askyesno(
                    "Export Complete", 
                    f"Markdown report exported successfully to:\n{output_file}\n\nWould you like to open the file?"
                )
                
                if open_file:
                    self._open_file_in_system(output_file)
                    
        except Exception as e:
            self.log_message(f"❌ Markdown export failed: {e}")
            messagebox.showerror("Export Error", f"Failed to export markdown report: {e}")
    
    def _open_file_in_system(self, file_path: str):
        """Open a file using the system's default application"""
        try:
            file_path = Path(file_path)
            if not file_path.exists():
                messagebox.showerror("File Not Found", f"The file does not exist:\n{file_path}")
                return
                
            # Use different commands based on the operating system
            if sys.platform == 'win32':
                os.startfile(str(file_path))
            elif sys.platform == 'darwin':  # macOS
                os.system(f'open "{file_path}"')
            else:  # Linux and others
                os.system(f'xdg-open "{file_path}"')
                
            self.log_message(f"📄 Opened file: {file_path.name}")
            
        except Exception as e:
            self.log_message(f"❌ Failed to open file: {e}")
            messagebox.showerror("Open Error", f"Failed to open file:\n{e}")
    
    def _search_report(self):
        """Search for text in the report"""
        search_term = self.search_var.get().strip()
        if not search_term:
            return
            
        # Clear previous search highlights
        self.report_text.tag_remove("search_highlight", "1.0", tk.END)
        
        # Find all occurrences
        start = "1.0"
        count = 0
        
        while True:
            pos = self.report_text.search(search_term, start, tk.END, nocase=True)
            if not pos:
                break
                
            # Calculate end position
            end = f"{pos}+{len(search_term)}c"
            
            # Highlight the match
            self.report_text.tag_add("search_highlight", pos, end)
            
            # Move to next position
            start = end
            count += 1
        
        # Configure highlight tag
        self.report_text.tag_config("search_highlight", background="yellow", foreground="black")
        
        # Show first match
        if count > 0:
            first_match = self.report_text.search(search_term, "1.0", tk.END, nocase=True)
            self.report_text.see(first_match)
            self.log_message(f"🔍 Found {count} occurrences of '{search_term}'")
        else:
            self.log_message(f"🔍 No occurrences of '{search_term}' found")
    
    def _clear_search(self):
        """Clear search highlights and text"""
        self.search_var.set("")
        self.report_text.tag_remove("search_highlight", "1.0", tk.END)
        
    def _on_search_change(self, event=None):
        """Handle search text change"""
        # Auto-search as user types (with debouncing)
        if hasattr(self, '_search_timer'):
            self.root.after_cancel(self._search_timer)
        
        if self.search_var.get().strip():
            self._search_timer = self.root.after(500, self._search_report)  # 500ms delay
        else:
            self._clear_search()
    
    def stop_analysis(self):
        """Stop running analysis"""
        if self.analysis_thread and self.analysis_thread.is_alive():
            self.analysis_running.set(False)
            self.progress_tracker.set_error("Analysis stopped by user")
            self.log_message("⏹️ Analysis stop requested")
            
    def clear_logs(self):
        """Clear log display"""
        self.log_text.delete(1.0, tk.END)
        
    def log_message(self, message: str):
        """Add message to log display"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_entry = f"[{timestamp}] {message}\n"
        
        # Insert on main thread
        if threading.current_thread() == threading.main_thread():
            self.log_text.insert(tk.END, log_entry)
            self.log_text.see(tk.END)
        else:
            self.root.after(0, lambda: self._insert_log(log_entry))
            
    def _insert_log(self, log_entry: str):
        """Insert log entry on main thread"""
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)


if __name__ == "__main__":
    """Direct execution - redirect to main.py"""
    print("⚠️  Direct execution deprecated. Use: python main.py")
    print("Launching GUI mode...")
    root = tk.Tk()
    app = ExcelExplorerApp(root)
    root.mainloop()
