import customtkinter as ctk
import json
import os
import re
import sys
import threading
from tkinter import filedialog

# Try to import openpyxl for Excel export
try:
    from openpyxl import Workbook
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# Check platform
IS_MACOS = sys.platform == "darwin"
IS_WINDOWS = sys.platform == "win32"

# Try to import sqlparse for fallback formatting
try:
    import sqlparse
    SQLPARSE_AVAILABLE = True
except ImportError:
    SQLPARSE_AVAILABLE = False

# Try to import PIL for icon creation
PIL_AVAILABLE = False
try:
    from PIL import Image, ImageDraw
    PIL_AVAILABLE = True
except ImportError:
    pass

# Try to import tray dependencies (Windows only - macOS uses dock)
TRAY_AVAILABLE = False
pystray = None

if not IS_MACOS and PIL_AVAILABLE:
    try:
        import pystray
        TRAY_AVAILABLE = True
    except ImportError:
        pass

# Set the appearance (dark/light mode and color theme)
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Data file for persistent history
DATA_DIR = os.path.join(os.path.expanduser("~"), ".teamutils")
HISTORY_FILE = os.path.join(DATA_DIR, "clipboard_history.json")


def create_tray_icon_image():
    """Create a simple icon for the system tray"""
    size = 64
    image = Image.new("RGB", (size, size), color=(30, 30, 30))
    draw = ImageDraw.Draw(image)
    # Draw a simple "T" for TeamUtils
    draw.rectangle([16, 12, 48, 20], fill=(52, 131, 235))  # Top bar
    draw.rectangle([28, 12, 36, 52], fill=(52, 131, 235))  # Vertical bar
    return image


class TeamUtilsApp(ctk.CTk):
    """Main application window"""

    def __init__(self):
        super().__init__()

        # Window setup
        self.title("Tırnakla")
        self.geometry("800x400")  # Wider for side-by-side layout

        # Set app icon
        self._set_app_icon()

        # Clipboard history storage
        self.clipboard_history = []
        self.last_clipboard = ""
        self.max_history = 50  # Keep last 50 items

        # System tray
        self.tray_icon = None
        self.is_quitting = False

        # Load saved history
        self.load_history()

        # Create the tab container
        self.tabview = ctk.CTkTabview(self)
        self.tabview.pack(expand=True, fill="both", padx=10, pady=10)

        # Add tabs
        self.tab_quoter = self.tabview.add("SQL Quoter")
        self.tab_formatter = self.tabview.add("SQL Formatter")
        self.tab_converter = self.tabview.add("Report Converter")
        self.tab_clipboard = self.tabview.add("Clipboard History")

        # Setup each tab
        self.setup_quoter_tab()
        self.setup_formatter_tab()
        self.setup_converter_tab()
        self.setup_clipboard_tab()

        # Start monitoring clipboard
        self.monitor_clipboard()

        # Setup system tray (Windows) or dock minimize (macOS)
        if TRAY_AVAILABLE:
            self.setup_tray()
            self.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)
        elif IS_MACOS:
            # macOS: minimize to dock on close (click dock icon to restore)
            self.protocol("WM_DELETE_WINDOW", self.minimize_to_dock)
        else:
            # No tray available, just quit normally but save history first
            self.protocol("WM_DELETE_WINDOW", self.quit_app)

    def _set_app_icon(self):
        """Set the application icon from custom image"""
        if not PIL_AVAILABLE:
            return

        try:
            # Determine base path (works for both dev and PyInstaller bundle)
            if getattr(sys, 'frozen', False):
                # Running as compiled executable
                base_path = sys._MEIPASS
            else:
                # Running as script
                base_path = os.path.dirname(os.path.abspath(__file__))

            icon_source = os.path.join(base_path, "assets", "icon.png")

            if os.path.exists(icon_source):
                # Load and resize the custom image
                image = Image.open(icon_source)
                image = image.resize((256, 256), Image.Resampling.LANCZOS)

                # Convert to RGBA if needed
                if image.mode != 'RGBA':
                    image = image.convert('RGBA')
            else:
                # Fallback: create a simple blue icon
                image = Image.new("RGBA", (256, 256), color=(52, 131, 235, 255))

            # Convert to PhotoImage for tkinter
            from PIL import ImageTk
            self._icon_image = ImageTk.PhotoImage(image)
            self.iconphoto(True, self._icon_image)

        except Exception:
            # If icon creation fails, just continue without it
            pass

    def setup_quoter_tab(self):
        """Setup the SQL Quoter tab"""
        # Container for side-by-side layout
        container = ctk.CTkFrame(self.tab_quoter, fg_color="transparent")
        container.pack(expand=True, fill="both", padx=10, pady=10)

        # Configure grid columns to be equal width
        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=0)  # Button column
        container.grid_columnconfigure(2, weight=1)
        container.grid_rowconfigure(2, weight=1)  # Main content row

        # TOP ROW - Template input and options
        template_frame = ctk.CTkFrame(container, fg_color="transparent")
        template_frame.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        template_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(template_frame, text="Template:").grid(row=0, column=0, sticky="w", padx=(0, 5))
        self.template_input = ctk.CTkEntry(template_frame, placeholder_text="e.g. SELECT %s FROM t_order (leave empty for quote mode)")
        self.template_input.grid(row=0, column=1, sticky="ew", padx=5)

        # Options frame
        options_frame = ctk.CTkFrame(template_frame, fg_color="transparent")
        options_frame.grid(row=0, column=2, sticky="e", padx=(10, 0))

        self.quote_var = ctk.BooleanVar(value=True)
        self.trim_var = ctk.BooleanVar(value=True)
        self.comma_var = ctk.BooleanVar(value=True)

        self.quote_checkbox = ctk.CTkCheckBox(options_frame, text="Quote", variable=self.quote_var, width=60)
        self.quote_checkbox.pack(side="left", padx=5)
        self.trim_checkbox = ctk.CTkCheckBox(options_frame, text="Trim", variable=self.trim_var, width=60)
        self.trim_checkbox.pack(side="left", padx=5)
        self.comma_checkbox = ctk.CTkCheckBox(options_frame, text="Comma", variable=self.comma_var, width=70)
        self.comma_checkbox.pack(side="left", padx=5)

        # File mode button
        file_btn = ctk.CTkButton(options_frame, text="File Mode", width=80, command=self.file_mode_convert)
        file_btn.pack(side="left", padx=(15, 5))

        # LEFT SIDE - Input
        input_label = ctk.CTkLabel(container, text="Paste your values (one per line):")
        input_label.grid(row=1, column=0, sticky="w", pady=(0, 5))

        self.quoter_input = ctk.CTkTextbox(container)
        self.quoter_input.grid(row=2, column=0, sticky="nsew", pady=5)

        # MIDDLE - Convert button
        convert_btn = ctk.CTkButton(
            container,
            text="→",
            width=50,
            command=self.convert_to_sql
        )
        convert_btn.grid(row=2, column=1, padx=15)

        # RIGHT SIDE - Output
        output_label = ctk.CTkLabel(container, text="Result (click to copy):")
        output_label.grid(row=1, column=2, sticky="w", pady=(0, 5))

        self.quoter_output = ctk.CTkTextbox(container)
        self.quoter_output.grid(row=2, column=2, sticky="nsew", pady=5)

        # Bind click on output to copy
        self.quoter_output.bind("<Button-1>", self.copy_output_to_clipboard)

    def setup_formatter_tab(self):
        """Setup the SQL Formatter tab"""
        # Container for side-by-side layout
        container = ctk.CTkFrame(self.tab_formatter, fg_color="transparent")
        container.pack(expand=True, fill="both", padx=10, pady=10)

        # Configure grid columns to be equal width
        container.grid_columnconfigure(0, weight=1)
        container.grid_columnconfigure(1, weight=0)  # Button column
        container.grid_columnconfigure(2, weight=1)
        container.grid_rowconfigure(1, weight=1)

        # TOP ROW - Labels and mode toggle
        input_label = ctk.CTkLabel(container, text="Paste your SQL query:")
        input_label.grid(row=0, column=0, sticky="w", pady=(0, 5))

        # Mode toggle (River style vs Standard)
        self.formatter_mode = ctk.StringVar(value="river")
        mode_frame = ctk.CTkFrame(container, fg_color="transparent")
        mode_frame.grid(row=0, column=1, columnspan=2, sticky="e", pady=(0, 5))

        ctk.CTkLabel(mode_frame, text="Style:").pack(side="left", padx=(0, 5))
        self.river_radio = ctk.CTkRadioButton(
            mode_frame, text="River", variable=self.formatter_mode,
            value="river", font=ctk.CTkFont(size=11)
        )
        self.river_radio.pack(side="left", padx=2)
        self.standard_radio = ctk.CTkRadioButton(
            mode_frame, text="Standard", variable=self.formatter_mode,
            value="standard", font=ctk.CTkFont(size=11)
        )
        self.standard_radio.pack(side="left", padx=2)

        # LEFT SIDE - Input

        self.formatter_input = ctk.CTkTextbox(container, font=("Courier", 12))
        self.formatter_input.grid(row=1, column=0, sticky="nsew", pady=5)

        # MIDDLE - Format button
        format_btn = ctk.CTkButton(
            container,
            text="→",
            width=50,
            command=self.format_sql
        )
        format_btn.grid(row=1, column=1, padx=15)

        # RIGHT SIDE - Output
        output_label = ctk.CTkLabel(container, text="Formatted SQL (auto-copied):")
        output_label.grid(row=0, column=2, sticky="w", pady=(0, 5))

        self.formatter_output = ctk.CTkTextbox(container, font=("Courier", 12))
        self.formatter_output.grid(row=1, column=2, sticky="nsew", pady=5)

        # Bind click on output to copy
        self.formatter_output.bind("<Button-1>", self.copy_formatter_output)

    def setup_converter_tab(self):
        """Setup the Report Converter tab"""
        container = ctk.CTkFrame(self.tab_converter, fg_color="transparent")
        container.pack(expand=True, fill="both", padx=10, pady=10)

        # Configure grid
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(2, weight=1)  # Preview area expands

        # TOP ROW - Controls
        controls_frame = ctk.CTkFrame(container, fg_color="transparent")
        controls_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        # Select File button
        self.converter_file_btn = ctk.CTkButton(
            controls_frame, text="Select File", width=100,
            command=self.converter_select_file
        )
        self.converter_file_btn.pack(side="left", padx=(0, 10))

        # File path label
        self.converter_file_label = ctk.CTkLabel(controls_frame, text="No file selected", text_color="gray")
        self.converter_file_label.pack(side="left", padx=(0, 15))

        # Delimiter dropdown
        ctk.CTkLabel(controls_frame, text="Delimiter:").pack(side="left", padx=(10, 5))
        self.delimiter_var = ctk.StringVar(value="Auto")
        self.delimiter_dropdown = ctk.CTkOptionMenu(
            controls_frame,
            variable=self.delimiter_var,
            values=["Auto", "Tab", "Pipe |", "Comma", "Fixed-width"],
            width=110,
            command=self.converter_refresh_preview
        )
        self.delimiter_dropdown.pack(side="left", padx=(0, 15))

        # Skip rows
        ctk.CTkLabel(controls_frame, text="Skip rows:").pack(side="left", padx=(10, 5))
        self.skip_rows_var = ctk.StringVar(value="0")
        self.skip_rows_entry = ctk.CTkEntry(controls_frame, width=50, textvariable=self.skip_rows_var)
        self.skip_rows_entry.pack(side="left", padx=(0, 15))
        self.skip_rows_entry.bind("<KeyRelease>", lambda e: self.converter_refresh_preview())

        # Convert button
        self.convert_btn = ctk.CTkButton(
            controls_frame, text="Convert to Excel", width=120,
            command=self.converter_start_conversion, state="disabled"
        )
        self.convert_btn.pack(side="right", padx=(10, 0))

        # INFO ROW - Column count and status
        info_frame = ctk.CTkFrame(container, fg_color="transparent")
        info_frame.grid(row=1, column=0, sticky="ew", pady=(0, 5))

        self.converter_info_label = ctk.CTkLabel(info_frame, text="", font=ctk.CTkFont(size=12))
        self.converter_info_label.pack(side="left")

        self.converter_status_label = ctk.CTkLabel(info_frame, text="", text_color="green")
        self.converter_status_label.pack(side="right")

        # PREVIEW AREA
        preview_label = ctk.CTkLabel(container, text="Preview (first 20 rows):", anchor="w")
        preview_label.grid(row=2, column=0, sticky="nw", pady=(5, 2))

        self.converter_preview = ctk.CTkTextbox(container, font=("Courier", 11))
        self.converter_preview.grid(row=2, column=0, sticky="nsew", pady=(25, 0))

        # PROGRESS BAR (hidden initially)
        self.converter_progress_frame = ctk.CTkFrame(container, fg_color="transparent")
        self.converter_progress_frame.grid(row=3, column=0, sticky="ew", pady=(10, 0))

        self.converter_progress = ctk.CTkProgressBar(self.converter_progress_frame)
        self.converter_progress.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.converter_progress.set(0)

        self.converter_progress_label = ctk.CTkLabel(self.converter_progress_frame, text="")
        self.converter_progress_label.pack(side="right")

        # Hide progress bar initially
        self.converter_progress_frame.grid_remove()

        # Store file info
        self.converter_file_path = None
        self.converter_sample_lines = []
        self.converter_total_lines = 0

    def converter_select_file(self):
        """Select a file for conversion"""
        file_path = filedialog.askopenfilename(
            title="Select SQL Report File",
            filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not file_path:
            return

        self.converter_file_path = file_path

        # Update UI
        filename = os.path.basename(file_path)
        self.converter_file_label.configure(text=filename, text_color="white")
        self.convert_btn.configure(state="normal")

        # Count lines and get sample
        self.converter_status_label.configure(text="Reading file...")
        self.update()

        try:
            with open(file_path, 'r', encoding='utf-8', errors='replace') as f:
                self.converter_sample_lines = []
                self.converter_total_lines = 0
                for i, line in enumerate(f):
                    if i < 50:  # Keep first 50 for sampling
                        self.converter_sample_lines.append(line)
                    self.converter_total_lines += 1

            self.converter_status_label.configure(text=f"{self.converter_total_lines:,} lines")
            self.converter_refresh_preview()

        except Exception as e:
            self.converter_status_label.configure(text=f"Error: {str(e)}", text_color="red")

    def converter_refresh_preview(self, *args):
        """Refresh the preview based on current settings"""
        if not self.converter_sample_lines:
            return

        # Get delimiter
        delimiter = self._get_selected_delimiter()

        # Get skip rows
        try:
            skip_rows = int(self.skip_rows_var.get())
        except ValueError:
            skip_rows = 0

        # Parse and display preview
        preview_lines = []
        col_widths = []
        parsed_rows = []

        for i, line in enumerate(self.converter_sample_lines):
            if i < skip_rows:
                continue
            if i >= skip_rows + 20:  # Show 20 rows max
                break

            line = line.rstrip('\n\r')
            if not line.strip():
                continue

            if delimiter:
                columns = line.split(delimiter)
            else:
                columns = re.split(r'  +', line)

            columns = [col.strip() for col in columns]
            parsed_rows.append(columns)

            # Track column widths
            for j, col in enumerate(columns):
                if j >= len(col_widths):
                    col_widths.append(0)
                col_widths[j] = max(col_widths[j], min(len(col), 25))  # Cap at 25 chars

        # Build preview text
        if parsed_rows:
            num_cols = max(len(row) for row in parsed_rows)
            self.converter_info_label.configure(text=f"Detected {num_cols} columns")

            # Format as table
            for row in parsed_rows:
                formatted_cols = []
                for j, col in enumerate(row):
                    width = col_widths[j] if j < len(col_widths) else 15
                    # Truncate long values
                    if len(col) > width:
                        col = col[:width-2] + ".."
                    formatted_cols.append(col.ljust(width))
                preview_lines.append(" │ ".join(formatted_cols))

        # Update preview
        self.converter_preview.delete("1.0", "end")
        self.converter_preview.insert("1.0", "\n".join(preview_lines))

    def _get_selected_delimiter(self):
        """Get the delimiter based on dropdown selection"""
        selection = self.delimiter_var.get()
        if selection == "Tab":
            return '\t'
        elif selection == "Pipe |":
            return '|'
        elif selection == "Comma":
            return ','
        elif selection == "Fixed-width":
            return None  # Will use regex split
        else:  # Auto
            return self._detect_delimiter(self.converter_sample_lines)

    def converter_start_conversion(self):
        """Start the conversion in a background thread"""
        if not self.converter_file_path:
            return

        if not OPENPYXL_AVAILABLE:
            self.converter_status_label.configure(text="Error: openpyxl not installed", text_color="red")
            return

        # Get output file
        output_file = filedialog.asksaveasfilename(
            title="Save Excel File As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not output_file:
            return

        # Show progress bar
        self.converter_progress_frame.grid()
        self.converter_progress.set(0)
        self.convert_btn.configure(state="disabled")
        self.converter_status_label.configure(text="Converting...", text_color="orange")

        # Run conversion in background thread
        thread = threading.Thread(
            target=self._run_conversion,
            args=(self.converter_file_path, output_file),
            daemon=True
        )
        thread.start()

    def _run_conversion(self, input_file, output_file):
        """Run the actual conversion (in background thread)"""
        try:
            delimiter = self._get_selected_delimiter()
            try:
                skip_rows = int(self.skip_rows_var.get())
            except ValueError:
                skip_rows = 0

            MAX_ROWS_PER_SHEET = 1000000
            wb = Workbook(write_only=True)
            ws = None
            sheet_num = 1
            row_count = 0
            total_rows = 0
            header_row = None
            data_line_num = 0

            with open(input_file, 'r', encoding='utf-8', errors='replace') as f:
                for line_num, line in enumerate(f):
                    # Skip specified rows
                    if line_num < skip_rows:
                        continue

                    line = line.rstrip('\n\r')
                    if not line.strip():
                        continue

                    # Create new sheet if needed
                    if ws is None or row_count >= MAX_ROWS_PER_SHEET:
                        ws = wb.create_sheet(title=f"Data_{sheet_num}")
                        sheet_num += 1
                        row_count = 0
                        if header_row is not None:
                            ws.append(header_row)
                            row_count += 1

                    # Parse columns
                    if delimiter:
                        columns = line.split(delimiter)
                    else:
                        columns = re.split(r'  +', line)
                    columns = [col.strip() for col in columns]

                    # Store header
                    if data_line_num == 0:
                        header_row = columns

                    ws.append(columns)
                    row_count += 1
                    total_rows += 1
                    data_line_num += 1

                    # Update progress
                    if total_rows % 10000 == 0:
                        progress = min(total_rows / max(self.converter_total_lines - skip_rows, 1), 1.0)
                        self.after(0, lambda p=progress, t=total_rows: self._update_progress(p, t))

            # Save
            wb.save(output_file)

            # Done
            sheets = sheet_num - 1
            self.after(0, lambda: self._conversion_complete(total_rows, sheets, output_file))

        except Exception as e:
            self.after(0, lambda: self._conversion_error(str(e)))

    def _update_progress(self, progress, rows):
        """Update progress bar (called from main thread)"""
        self.converter_progress.set(progress)
        self.converter_progress_label.configure(text=f"{rows:,} rows")

    def _conversion_complete(self, total_rows, sheets, output_file):
        """Handle conversion complete (called from main thread)"""
        self.converter_progress.set(1.0)
        self.converter_progress_label.configure(text=f"{total_rows:,} rows")
        self.converter_status_label.configure(
            text=f"✓ Saved to {os.path.basename(output_file)} ({sheets} sheet{'s' if sheets > 1 else ''})",
            text_color="green"
        )
        self.convert_btn.configure(state="normal")

    def _conversion_error(self, error_msg):
        """Handle conversion error (called from main thread)"""
        self.converter_status_label.configure(text=f"Error: {error_msg}", text_color="red")
        self.convert_btn.configure(state="normal")
        self.converter_progress_frame.grid_remove()

    def setup_clipboard_tab(self):
        """Setup the Clipboard History tab"""
        # Header with clear button
        header = ctk.CTkFrame(self.tab_clipboard, fg_color="transparent")
        header.pack(fill="x", padx=10, pady=(10, 5))

        title_label = ctk.CTkLabel(
            header,
            text="Click any item to copy it back to clipboard",
            font=ctk.CTkFont(size=14)
        )
        title_label.pack(side="left")

        clear_btn = ctk.CTkButton(
            header,
            text="Clear History",
            width=100,
            fg_color="gray",
            command=self.clear_clipboard_history
        )
        clear_btn.pack(side="right")

        # Scrollable frame for history items
        self.history_frame = ctk.CTkScrollableFrame(self.tab_clipboard)
        self.history_frame.pack(expand=True, fill="both", padx=10, pady=10)

        # Refresh UI with loaded history
        self.refresh_history_ui()

    def setup_tray(self):
        """Setup the system tray icon (Windows only)"""
        try:
            icon_image = create_tray_icon_image()

            menu = pystray.Menu(
                pystray.MenuItem("Show", self.show_from_tray, default=True),
                pystray.MenuItem("Quit", self.quit_app)
            )

            self.tray_icon = pystray.Icon("TeamUtils", icon_image, "Team Utils", menu)

            # Run tray icon in a separate thread
            tray_thread = threading.Thread(target=self.tray_icon.run, daemon=True)
            tray_thread.start()
        except Exception as e:
            print(f"Tray setup failed: {e}")
            self.tray_icon = None
            self.protocol("WM_DELETE_WINDOW", self.quit_app)

    def minimize_to_tray(self):
        """Hide window to system tray instead of closing (Windows)"""
        self.withdraw()  # Hide the window

    def minimize_to_dock(self):
        """Minimize to dock instead of closing (macOS)"""
        self.iconify()  # Minimize to dock

    def show_from_tray(self):
        """Restore window from system tray"""
        self.after(0, self._show_window)

    def _show_window(self):
        """Actually show the window (must run in main thread)"""
        self.deiconify()  # Show the window
        self.lift()  # Bring to front
        self.focus_force()  # Give focus

    def quit_app(self):
        """Properly quit the application"""
        self.is_quitting = True
        self.save_history()
        if self.tray_icon:
            self.tray_icon.stop()
        self.after(0, self.destroy)

    def load_history(self):
        """Load clipboard history from file"""
        try:
            if os.path.exists(HISTORY_FILE):
                with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                    self.clipboard_history = json.load(f)
                # Set last_clipboard to most recent to avoid re-adding
                if self.clipboard_history:
                    self.last_clipboard = self.clipboard_history[0]
        except Exception:
            self.clipboard_history = []

    def save_history(self):
        """Save clipboard history to file"""
        try:
            os.makedirs(DATA_DIR, exist_ok=True)
            with open(HISTORY_FILE, "w", encoding="utf-8") as f:
                json.dump(self.clipboard_history, f, ensure_ascii=False, indent=2)
        except Exception:
            pass  # Fail silently

    def monitor_clipboard(self):
        """Check clipboard for changes every 500ms"""
        if self.is_quitting:
            return

        try:
            current = self.clipboard_get()
            if current and current != self.last_clipboard:
                self.last_clipboard = current
                self.add_to_history(current)
        except Exception:
            # Clipboard might be empty or contain non-text data
            pass

        # Schedule next check
        self.after(500, self.monitor_clipboard)

    def add_to_history(self, text):
        """Add text to clipboard history and update UI"""
        # Don't add duplicates (remove old one if exists)
        if text in self.clipboard_history:
            self.clipboard_history.remove(text)

        # Add to beginning of list
        self.clipboard_history.insert(0, text)

        # Trim to max size
        if len(self.clipboard_history) > self.max_history:
            self.clipboard_history = self.clipboard_history[:self.max_history]

        # Save to file
        self.save_history()

        # Refresh the UI
        self.refresh_history_ui()

    def refresh_history_ui(self):
        """Rebuild the history list UI"""
        # Clear existing widgets
        for widget in self.history_frame.winfo_children():
            widget.destroy()

        # Add a button for each history item
        for text in self.clipboard_history:
            # Truncate display text if too long
            display_text = text[:100] + "..." if len(text) > 100 else text
            # Replace newlines with spaces for display
            display_text = display_text.replace("\n", " ↵ ")

            btn = ctk.CTkButton(
                self.history_frame,
                text=display_text,
                anchor="w",
                fg_color="gray25",
                hover_color="gray35",
                command=lambda t=text: self.copy_history_item(t)
            )
            btn.pack(fill="x", pady=2)

    def copy_history_item(self, text):
        """Copy a history item back to clipboard"""
        self.clipboard_clear()
        self.clipboard_append(text)
        # Update last_clipboard so we don't re-add it
        self.last_clipboard = text

    def clear_clipboard_history(self):
        """Clear all clipboard history"""
        self.clipboard_history.clear()
        self.save_history()
        self.refresh_history_ui()

    def convert_to_sql(self):
        """Convert input lines to SQL-friendly format with single quotes or template substitution"""
        # Get input text
        input_text = self.quoter_input.get("1.0", "end-1c")
        template = self.template_input.get().strip()
        should_quote = self.quote_var.get()
        should_trim = self.trim_var.get()
        should_comma = self.comma_var.get()

        # Split into lines and process
        if should_trim:
            lines = [line.strip() for line in input_text.splitlines() if line.strip()]
        else:
            lines = [line for line in input_text.splitlines() if line.strip()]

        if not lines:
            return

        # Check if we're in template mode (template contains %s)
        if template and '%s' in template:
            # Template mode: substitute each value into template
            formatted_lines = []
            for line in lines:
                value = line
                if should_quote:
                    value = f"'{line}'"
                formatted_lines.append(template.replace('%s', value))
            result = "\n".join(formatted_lines)
        else:
            # Standard mode: optionally wrap in quotes and add commas
            formatted_lines = []
            for i, line in enumerate(lines):
                value = f"'{line}'" if should_quote else line
                # Add comma to all but last line (if comma enabled)
                if should_comma and i < len(lines) - 1:
                    value += ","
                formatted_lines.append(value)
            result = "\n".join(formatted_lines)

        # Clear output and insert result
        self.quoter_output.delete("1.0", "end")
        self.quoter_output.insert("1.0", result)

        # Auto-copy to clipboard
        self.clipboard_clear()
        self.clipboard_append(result)
        self.last_clipboard = result  # Prevent adding to history

    def file_mode_convert(self):
        """Process large files directly without loading into textboxes"""
        # Get settings
        template = self.template_input.get().strip()
        should_quote = self.quote_var.get()
        should_trim = self.trim_var.get()
        should_comma = self.comma_var.get()

        # Select input file
        input_file = filedialog.askopenfilename(
            title="Select Input File",
            filetypes=[("Text files", "*.txt"), ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not input_file:
            return

        # Select output file
        output_file = filedialog.asksaveasfilename(
            title="Save Output As",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("SQL files", "*.sql"), ("All files", "*.*")]
        )
        if not output_file:
            return

        # Process file
        try:
            with open(input_file, 'r', encoding='utf-8') as infile:
                lines = infile.readlines()

            # Count non-empty lines for comma logic
            if should_trim:
                non_empty_lines = [line.strip() for line in lines if line.strip()]
            else:
                non_empty_lines = [line.rstrip('\n\r') for line in lines if line.strip()]

            total_lines = len(non_empty_lines)

            with open(output_file, 'w', encoding='utf-8') as outfile:
                for i, line in enumerate(non_empty_lines):
                    # Template mode
                    if template and '%s' in template:
                        value = line
                        if should_quote:
                            value = f"'{line}'"
                        outfile.write(template.replace('%s', value) + '\n')
                    else:
                        # Standard mode
                        value = f"'{line}'" if should_quote else line
                        if should_comma and i < total_lines - 1:
                            value += ","
                        outfile.write(value + '\n')

            # Show success in output box
            self.quoter_output.delete("1.0", "end")
            self.quoter_output.insert("1.0", f"✓ Processed {total_lines:,} lines\n\nInput: {input_file}\nOutput: {output_file}")

        except Exception as e:
            self.quoter_output.delete("1.0", "end")
            self.quoter_output.insert("1.0", f"Error: {str(e)}")

    def _detect_delimiter(self, sample_lines):
        """Detect the delimiter used in the file"""
        if not sample_lines:
            return '\t'

        # Count occurrences of common delimiters
        tab_counts = [line.count('\t') for line in sample_lines]
        pipe_counts = [line.count('|') for line in sample_lines]
        comma_counts = [line.count(',') for line in sample_lines]

        # Check for consistency (same count across lines suggests delimiter)
        def is_consistent(counts):
            if not counts or counts[0] == 0:
                return False
            return len(set(counts)) <= 2  # Allow small variance

        # Prefer tab, then pipe, then comma
        if is_consistent(tab_counts) and tab_counts[0] > 0:
            return '\t'
        if is_consistent(pipe_counts) and pipe_counts[0] > 0:
            return '|'
        if is_consistent(comma_counts) and comma_counts[0] > 0:
            return ','

        # Default to None (will use fixed-width splitting)
        return None

    def format_sql(self):
        """Format SQL query using selected formatter"""
        # Get input text
        input_text = self.formatter_input.get("1.0", "end-1c")

        if not input_text.strip():
            return

        # Choose formatter based on mode
        mode = self.formatter_mode.get()

        if mode == "river":
            formatted = self.river_format_sql(input_text)
        elif mode == "standard" and SQLPARSE_AVAILABLE:
            formatted = sqlparse.format(
                input_text,
                reindent=True,
                keyword_case='upper',
                indent_width=4
            )
        else:
            # Fallback to river if sqlparse not available
            formatted = self.river_format_sql(input_text)

        # Clear output and insert result
        self.formatter_output.delete("1.0", "end")
        self.formatter_output.insert("1.0", formatted)

        # Auto-copy to clipboard
        self.clipboard_clear()
        self.clipboard_append(formatted)
        self.last_clipboard = formatted  # Prevent adding to history

    def river_format_sql(self, sql):
        """
        River-style SQL formatter where keywords right-align.

        Example output:
        SELECT column1
             , column2
             , column3
          FROM table1
          JOIN table2
            ON table1.id = table2.id
         WHERE condition1 = 'value'
           AND condition2 = 'value'
        """
        # Keyword width - keywords right-align to end at this column
        KW_WIDTH = 10

        # Main clause keywords (these start new lines, right-aligned)
        # Note: 'WITH' at the start of a query is CTE, but WITH (NOLOCK) is a table hint - handle separately
        CLAUSE_KEYWORDS = [
            'DECLARE', 'SET',  # T-SQL variable declarations
            'SELECT', 'FROM', 'WHERE', 'ORDER BY', 'GROUP BY', 'HAVING',
            'UNION ALL', 'UNION', 'EXCEPT', 'INTERSECT', 'LIMIT', 'OFFSET',
            'INSERT INTO', 'INSERT', 'UPDATE', 'DELETE FROM', 'DELETE',
            'VALUES', 'INTO'
        ]

        # Join keywords
        JOIN_KEYWORDS = [
            'LEFT OUTER JOIN', 'RIGHT OUTER JOIN', 'FULL OUTER JOIN',
            'CROSS APPLY', 'OUTER APPLY',
            'INNER JOIN', 'LEFT JOIN', 'RIGHT JOIN', 'FULL JOIN',
            'CROSS JOIN', 'JOIN'
        ]

        def right_align_keyword(keyword, width=KW_WIDTH):
            """Right-align a keyword to the specified width"""
            padding = max(0, width - len(keyword))
            return ' ' * padding + keyword

        def split_respecting_parens(text, delimiter=','):
            """Split text by delimiter, respecting parentheses and quotes"""
            result = []
            current = ""
            paren_depth = 0
            in_single_quote = False
            in_double_quote = False

            i = 0
            while i < len(text):
                char = text[i]

                # Handle quotes
                if char == "'" and not in_double_quote:
                    in_single_quote = not in_single_quote
                elif char == '"' and not in_single_quote:
                    in_double_quote = not in_double_quote

                # Handle parentheses (only when not in quotes)
                if not in_single_quote and not in_double_quote:
                    if char == '(':
                        paren_depth += 1
                    elif char == ')':
                        paren_depth -= 1
                    elif char == delimiter and paren_depth == 0:
                        result.append(current)
                        current = ""
                        i += 1
                        continue

                current += char
                i += 1

            if current.strip():
                result.append(current)

            return result

        def find_keyword_boundary(text, keywords, start=0):
            """Find the next keyword boundary, respecting parentheses and quotes"""
            text_upper = text.upper()
            paren_depth = 0
            in_single_quote = False
            in_double_quote = False

            i = start
            while i < len(text):
                char = text[i]

                # Handle quotes
                if char == "'" and not in_double_quote:
                    in_single_quote = not in_single_quote
                elif char == '"' and not in_single_quote:
                    in_double_quote = not in_double_quote

                # Handle parentheses
                if not in_single_quote and not in_double_quote:
                    if char == '(':
                        paren_depth += 1
                    elif char == ')':
                        paren_depth -= 1

                    # Only check for keywords at depth 0
                    if paren_depth == 0:
                        # Must be at word boundary
                        if i == 0 or not text[i-1].isalnum():
                            for kw in keywords:
                                if text_upper[i:].startswith(kw):
                                    # Check end boundary
                                    end_pos = i + len(kw)
                                    if end_pos >= len(text) or not text[end_pos].isalnum():
                                        # Special case: WITH followed by ( is a table hint, not a CTE keyword
                                        if kw == 'WITH':
                                            after_with = text[end_pos:].lstrip()
                                            if after_with.startswith('('):
                                                continue  # Skip this, it's WITH (NOLOCK) etc
                                        return i

                i += 1

            return -1

        def format_case_when(text, indent):
            """Format CASE WHEN statements with WHEN/ELSE/END on separate lines"""
            result = []
            remaining = text
            case_indent = ' ' * indent
            when_indent = ' ' * (indent + 2)

            # Find CASE
            case_pos = remaining.upper().find('CASE')
            if case_pos == -1:
                return text

            # Add everything before CASE
            if case_pos > 0:
                result.append(remaining[:case_pos].rstrip())

            remaining = remaining[case_pos:]

            # Process CASE ... END
            # Find matching END (respecting nested CASE)
            depth = 0
            i = 0
            case_content = ""
            while i < len(remaining):
                upper_rest = remaining[i:].upper()
                if upper_rest.startswith('CASE') and (i == 0 or not remaining[i-1].isalnum()):
                    if i + 4 >= len(remaining) or not remaining[i+4].isalnum():
                        depth += 1
                        case_content += remaining[i:i+4]
                        i += 4
                        continue
                elif upper_rest.startswith('END') and (i == 0 or not remaining[i-1].isalnum()):
                    if i + 3 >= len(remaining) or not remaining[i+3].isalnum():
                        depth -= 1
                        if depth == 0:
                            case_content += 'END'
                            remaining = remaining[i+3:]
                            break
                        case_content += remaining[i:i+3]
                        i += 3
                        continue
                case_content += remaining[i]
                i += 1

            # Now format the case_content
            # Split on WHEN, THEN, ELSE, END
            parts = []
            temp = case_content
            current = ""
            j = 0
            while j < len(temp):
                upper_temp = temp[j:].upper()
                matched_kw = None
                for kw in ['WHEN', 'THEN', 'ELSE', 'END']:
                    if upper_temp.startswith(kw):
                        if j + len(kw) >= len(temp) or not temp[j+len(kw)].isalnum():
                            if j == 0 or not temp[j-1].isalnum():
                                matched_kw = kw
                                break
                if matched_kw:
                    if current.strip():
                        parts.append(current.strip())
                    parts.append(matched_kw)
                    current = ""
                    j += len(matched_kw)
                else:
                    current += temp[j]
                    j += 1
            if current.strip():
                parts.append(current.strip())

            # Build formatted output
            formatted_case = []
            i = 0
            while i < len(parts):
                part = parts[i]
                if part == 'CASE':
                    formatted_case.append('CASE')
                elif part == 'WHEN':
                    if i + 1 < len(parts) and parts[i+1] not in ['WHEN', 'THEN', 'ELSE', 'END']:
                        formatted_case.append(when_indent + 'WHEN ' + parts[i+1])
                        i += 1
                    else:
                        formatted_case.append(when_indent + 'WHEN')
                elif part == 'THEN':
                    if i + 1 < len(parts) and parts[i+1] not in ['WHEN', 'THEN', 'ELSE', 'END']:
                        formatted_case.append(' THEN ' + parts[i+1])
                        i += 1
                    else:
                        formatted_case.append(' THEN')
                elif part == 'ELSE':
                    if i + 1 < len(parts) and parts[i+1] not in ['WHEN', 'THEN', 'ELSE', 'END']:
                        formatted_case.append(when_indent + 'ELSE ' + parts[i+1])
                        i += 1
                    else:
                        formatted_case.append(when_indent + 'ELSE')
                elif part == 'END':
                    formatted_case.append(case_indent + 'END')
                i += 1

            # Join - THEN should stay on same line as WHEN
            final_lines = []
            temp_line = ""
            for item in formatted_case:
                if item.strip().startswith('THEN'):
                    temp_line += item
                else:
                    if temp_line:
                        final_lines.append(temp_line)
                    temp_line = item
            if temp_line:
                final_lines.append(temp_line)

            result.extend(final_lines)

            # Add any remaining text after END
            if remaining.strip():
                result.append(remaining.strip())

            return '\n'.join(result) if result else text

        def format_subqueries(text, current_indent):
            """Find and format subqueries (SELECT after opening paren)"""
            result = ""
            i = 0
            while i < len(text):
                # Look for (SELECT pattern
                if text[i] == '(':
                    # Check if SELECT follows
                    rest = text[i+1:].lstrip()
                    if rest.upper().startswith('SELECT'):
                        # Find matching closing paren
                        depth = 1
                        j = i + 1
                        while j < len(text) and depth > 0:
                            if text[j] == '(':
                                depth += 1
                            elif text[j] == ')':
                                depth -= 1
                            j += 1

                        # Extract subquery content (without outer parens)
                        subquery = text[i+1:j-1].strip()

                        # Format subquery with increased indent
                        sub_indent = ' ' * (current_indent + 4)
                        formatted_sub = process_sql(subquery, indent_level=1)

                        # Add formatted subquery with proper indentation
                        sub_lines = formatted_sub.split('\n')
                        result += '(\n'
                        for line in sub_lines:
                            result += sub_indent + line.strip() + '\n'
                        result += ' ' * current_indent + ')'

                        i = j
                        continue

                result += text[i]
                i += 1

            return result

        def format_select_columns(columns_text, base_indent):
            """Format SELECT columns with leading commas"""
            columns = split_respecting_parens(columns_text.strip(), ',')
            if not columns:
                return columns_text.strip()

            lines = []
            comma_indent = ' ' * (base_indent - 2)  # Comma goes 2 chars before content

            for i, col in enumerate(columns):
                col = col.strip()
                if not col:
                    continue

                # Check if column contains CASE WHEN
                if 'CASE' in col.upper():
                    col = format_case_when(col, base_indent)

                # Check for subqueries
                if '(SELECT' in col.upper() or '( SELECT' in col.upper():
                    col = format_subqueries(col, base_indent)

                if i == 0:
                    lines.append(col)
                else:
                    lines.append(comma_indent + ', ' + col)

            return '\n'.join(lines)

        def format_conditions(conditions_text, base_indent):
            """Format WHERE/ON conditions with AND/OR on separate lines"""
            result_lines = []
            remaining = conditions_text.strip()

            # First, find the first condition (before any AND/OR)
            first_and = find_keyword_boundary(remaining, ['AND'], 0)
            first_or = find_keyword_boundary(remaining, ['OR'], 0)

            if first_and == -1 and first_or == -1:
                # No AND/OR, just return as-is
                return remaining

            # Find the first split point
            if first_and == -1:
                first_split = first_or
            elif first_or == -1:
                first_split = first_and
            else:
                first_split = min(first_and, first_or)

            # Add first condition
            first_cond = remaining[:first_split].strip()
            if first_cond:
                result_lines.append(first_cond)

            remaining = remaining[first_split:]

            # Now process remaining AND/OR conditions
            while remaining:
                remaining = remaining.strip()
                if not remaining:
                    break

                # Check if we start with AND or OR
                remaining_upper = remaining.upper()
                kw = None
                if remaining_upper.startswith('AND') and (len(remaining) == 3 or not remaining[3].isalnum()):
                    kw = 'AND'
                elif remaining_upper.startswith('OR') and (len(remaining) == 2 or not remaining[2].isalnum()):
                    kw = 'OR'

                if kw:
                    # Skip past the keyword
                    remaining = remaining[len(kw):].strip()

                    # Find next AND/OR
                    next_and = find_keyword_boundary(remaining, ['AND'], 0)
                    next_or = find_keyword_boundary(remaining, ['OR'], 0)

                    if next_and == -1 and next_or == -1:
                        # No more, take rest
                        kw_formatted = right_align_keyword(kw, base_indent)
                        result_lines.append(kw_formatted + ' ' + remaining)
                        break
                    else:
                        if next_and == -1:
                            next_split = next_or
                        elif next_or == -1:
                            next_split = next_and
                        else:
                            next_split = min(next_and, next_or)

                        cond = remaining[:next_split].strip()
                        kw_formatted = right_align_keyword(kw, base_indent)
                        result_lines.append(kw_formatted + ' ' + cond)
                        remaining = remaining[next_split:]
                else:
                    # Shouldn't happen, but just add what's left
                    result_lines.append(remaining)
                    break

            return '\n'.join(result_lines)

        def process_sql(sql_text, indent_level=0):
            """Process SQL text and return formatted lines"""
            lines = []
            remaining = sql_text.strip()
            base_indent = indent_level * 4

            while remaining:
                remaining = remaining.strip()
                if not remaining:
                    break

                # Check for -- comments at start
                if remaining.startswith('--'):
                    newline_pos = remaining.find('\n')
                    if newline_pos == -1:
                        lines.append(' ' * base_indent + remaining)
                        break
                    else:
                        lines.append(' ' * base_indent + remaining[:newline_pos])
                        remaining = remaining[newline_pos + 1:]
                        continue

                matched = False
                remaining_upper = remaining.upper()

                # Try to match clause keywords first
                for kw in CLAUSE_KEYWORDS:
                    if remaining_upper.startswith(kw) and (
                        len(remaining) == len(kw) or not remaining[len(kw)].isalnum()
                    ):
                        # Found a clause keyword
                        kw_formatted = right_align_keyword(kw, KW_WIDTH + base_indent)
                        after_kw = remaining[len(kw):].strip()

                        # Find next clause/join keyword
                        # For repeatable statements (DECLARE, SET), include them in search
                        repeatable = ['DECLARE', 'SET']
                        if kw in repeatable:
                            search_keywords = CLAUSE_KEYWORDS + JOIN_KEYWORDS
                        else:
                            search_keywords = [k for k in (CLAUSE_KEYWORDS + JOIN_KEYWORDS) if k != kw]

                        # Also check for semicolon as statement terminator (respecting quotes)
                        semicolon_pos = -1
                        in_sq = False
                        in_dq = False
                        for idx, ch in enumerate(after_kw):
                            if ch == "'" and not in_dq:
                                in_sq = not in_sq
                            elif ch == '"' and not in_sq:
                                in_dq = not in_dq
                            elif ch == ';' and not in_sq and not in_dq:
                                semicolon_pos = idx
                                break

                        next_boundary = find_keyword_boundary(after_kw, search_keywords, 0)

                        # Use semicolon if it comes first
                        if semicolon_pos != -1 and (next_boundary == -1 or semicolon_pos < next_boundary):
                            next_boundary = semicolon_pos + 1  # Include the semicolon

                        if next_boundary == -1:
                            content = after_kw
                            remaining = ""
                        else:
                            content = after_kw[:next_boundary].strip()
                            remaining = after_kw[next_boundary:]

                        # Format based on keyword type
                        if kw == 'SELECT':
                            col_formatted = format_select_columns(content, KW_WIDTH + base_indent + 1)
                            first_line, *rest_lines = col_formatted.split('\n') if '\n' in col_formatted else [col_formatted]
                            lines.append(kw_formatted + ' ' + first_line)
                            lines.extend(rest_lines)
                        elif kw == 'WHERE':
                            cond_lines = format_conditions(content, KW_WIDTH + base_indent)
                            first_line, *rest_lines = cond_lines.split('\n') if '\n' in cond_lines else [cond_lines]
                            lines.append(kw_formatted + ' ' + first_line)
                            lines.extend(rest_lines)
                        elif kw in ['DECLARE', 'SET']:
                            # DECLARE/SET get a blank line after them
                            lines.append(kw_formatted + ' ' + content)
                            lines.append('')  # Blank line after
                        else:
                            lines.append(kw_formatted + ' ' + content)

                        matched = True
                        break

                if matched:
                    continue

                # Try JOIN keywords
                for kw in JOIN_KEYWORDS:
                    if remaining_upper.startswith(kw) and (
                        len(remaining) == len(kw) or not remaining[len(kw)].isalnum()
                    ):
                        kw_formatted = right_align_keyword(kw, KW_WIDTH + base_indent)
                        after_kw = remaining[len(kw):].strip()

                        # Find ON or next clause/join
                        on_pos = find_keyword_boundary(after_kw, ['ON'], 0)
                        next_clause = find_keyword_boundary(
                            after_kw,
                            [k for k in (CLAUSE_KEYWORDS + JOIN_KEYWORDS) if k != 'ON'],
                            0
                        )

                        if on_pos != -1 and (next_clause == -1 or on_pos < next_clause):
                            # ON is part of this JOIN
                            table_part = after_kw[:on_pos].strip()
                            on_rest = after_kw[on_pos + 2:].strip()  # Skip "ON"

                            # Find end of ON condition
                            next_after_on = find_keyword_boundary(
                                on_rest,
                                [k for k in (CLAUSE_KEYWORDS + JOIN_KEYWORDS + ['AND', 'OR'])],
                                0
                            )

                            if next_after_on == -1:
                                on_content = on_rest
                                remaining = ""
                            else:
                                on_content = on_rest[:next_after_on].strip()
                                remaining = on_rest[next_after_on:]

                            lines.append(kw_formatted + ' ' + table_part)
                            on_formatted = right_align_keyword('ON', KW_WIDTH + base_indent)
                            lines.append(on_formatted + ' ' + on_content)
                        else:
                            # No ON, just table
                            if next_clause == -1:
                                content = after_kw
                                remaining = ""
                            else:
                                content = after_kw[:next_clause].strip()
                                remaining = after_kw[next_clause:]

                            lines.append(kw_formatted + ' ' + content)

                        matched = True
                        break

                if matched:
                    continue

                # Try AND/OR keywords (for additional conditions)
                for kw in ['AND', 'OR']:
                    if remaining_upper.startswith(kw) and (
                        len(remaining) == len(kw) or not remaining[len(kw)].isalnum()
                    ):
                        kw_formatted = right_align_keyword(kw, KW_WIDTH + base_indent)
                        after_kw = remaining[len(kw):].strip()

                        # Find next keyword
                        next_boundary = find_keyword_boundary(
                            after_kw,
                            CLAUSE_KEYWORDS + JOIN_KEYWORDS + ['AND', 'OR'],
                            0
                        )

                        if next_boundary == -1:
                            content = after_kw
                            remaining = ""
                        else:
                            content = after_kw[:next_boundary].strip()
                            remaining = after_kw[next_boundary:]

                        lines.append(kw_formatted + ' ' + content)
                        matched = True
                        break

                if matched:
                    continue

                # No keyword matched - add whatever is left
                if remaining:
                    lines.append(' ' * base_indent + remaining)
                    break

            return '\n'.join(lines)

        # Main formatting logic
        return process_sql(sql)

    def copy_output_to_clipboard(self, event=None):
        """Copy the quoter output text to clipboard"""
        output_text = self.quoter_output.get("1.0", "end-1c")
        if output_text:
            self.clipboard_clear()
            self.clipboard_append(output_text)
            self.last_clipboard = output_text

    def copy_formatter_output(self, event=None):
        """Copy the formatter output text to clipboard"""
        output_text = self.formatter_output.get("1.0", "end-1c")
        if output_text:
            self.clipboard_clear()
            self.clipboard_append(output_text)
            self.last_clipboard = output_text


# Run the app
if __name__ == "__main__":
    app = TeamUtilsApp()
    app.mainloop()
