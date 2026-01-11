"""
Converter Suite - Main Application UI
Modern CustomTkinter interface with category-based navigation.
"""

import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import threading
import logging
from typing import Optional, Callable

# Configure CustomTkinter
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# Try to import drag-and-drop support
DND_AVAILABLE = False
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    pass

from ..converters.registry import registry, ConverterRegistry
from ..converters.base import ConversionCategory, ConversionOptions
from ..converters.image_converter import ImageConverter
from ..converters.video_converter import VideoConverter
from ..converters.audio_converter import AudioConverter
from ..converters.document_converter import DocumentConverter
from ..converters.presentation_converter import PresentationConverter
from ..utils.dependency_checker import dependency_checker

logger = logging.getLogger(__name__)


class ConverterApp(ctk.CTk if not DND_AVAILABLE else type('ConverterApp', (ctk.CTk, TkinterDnD.DnDWrapper), {})):
    """
    Main application window for Converter Suite.
    Features category-based navigation and unified conversion interface.
    """
    
    # Uppercase names for techy look
    CATEGORIES = [
        ("📄", "DOCUMENTS", ConversionCategory.DOCUMENT),
        ("🖼️", "IMAGES", ConversionCategory.IMAGE),
        ("🎬", "VIDEO", ConversionCategory.VIDEO),
        ("🎵", "AUDIO", ConversionCategory.AUDIO),
        ("📊", "PRESENTATIONS", ConversionCategory.PRESENTATION),
    ]
    
    def __init__(self):
        super().__init__()
        
        if DND_AVAILABLE:
            self.TkdndVersion = TkinterDnD._require(self)
        
        self.title("Converter Suite Pro")
        self.geometry("1200x750")
        self.minsize(900, 600)
        
        # State
        self.current_category: Optional[ConversionCategory] = None
        self.output_dir = Path.home() / "ConvertedFiles"
        self.output_dir.mkdir(exist_ok=True)
        self.selected_files: list[Path] = []
        self.output_format = tk.StringVar(value="")
        self.is_converting = False
        
        # Register converters
        self._register_converters()
        
        # Setup UI
        self._setup_ui()
        
        # Select first category by default
        self._select_category(ConversionCategory.DOCUMENT)
    
    def _register_converters(self):
        """Register all available converters."""
        registry.clear()
        registry.register(ImageConverter())
        registry.register(VideoConverter())
        registry.register(AudioConverter())
        registry.register(DocumentConverter())
        registry.register(PresentationConverter())
    
    def _setup_ui(self):
        """Setup the main UI layout."""
        # Configure grid
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Sidebar with darker techy look
        self._create_sidebar()
        
        # Main content area
        self._create_main_area()
    
    def _create_sidebar(self):
        """Create the sidebar with category navigation."""
        # Darker gray background for sidebar
        self.sidebar = ctk.CTkFrame(self, width=260, corner_radius=0, fg_color=("gray90", "gray10"))
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(10, weight=1)
        
        # Logo with tech font
        logo_frame = ctk.CTkFrame(self.sidebar, fg_color="transparent")
        logo_frame.grid(row=0, column=0, padx=20, pady=(30, 40), sticky="ew")
        
        logo = ctk.CTkLabel(
            logo_frame,
            text="CORE CONVERTER",
            font=ctk.CTkFont(family="Consolas", size=20, weight="bold"),
            text_color=("gray20", "gray90")
        )
        logo.pack(anchor="w")
        
        logo_sub = ctk.CTkLabel(
            logo_frame,
            text="v0.1.3 // SUITE",
            font=ctk.CTkFont(family="Consolas", size=10),
            text_color="gray50"
        )
        logo_sub.pack(anchor="w")

        logo_brand = ctk.CTkLabel(
            logo_frame,
            text="by graeLabs",
            font=ctk.CTkFont(family="Consolas", size=9, slant="italic"),
            text_color="gray40"
        )
        logo_brand.pack(anchor="w", pady=(2, 0))
        
        # Category buttons
        self.category_buttons: dict[ConversionCategory, ctk.CTkButton] = {}
        
        for i, (icon, name, category) in enumerate(self.CATEGORIES):
            btn = ctk.CTkButton(
                self.sidebar,
                text=f"{icon} {name}",
                font=ctk.CTkFont(family="Roboto Mono", size=13),
                anchor="w",
                height=45,
                corner_radius=4,
                fg_color="transparent",
                text_color=("gray40", "gray60"),
                hover_color=("gray80", "gray15"),
                border_spacing=15,
                command=lambda c=category: self._select_category(c)
            )
            btn.grid(row=i+1, column=0, padx=15, pady=2, sticky="ew")
            self.category_buttons[category] = btn
        
        # Separator line
        separator = ctk.CTkFrame(self.sidebar, height=1, fg_color=("gray80", "gray20"))
        separator.grid(row=len(self.CATEGORIES)+1, column=0, padx=20, pady=30, sticky="ew")
        
        # Dependency status section
        dep_label = ctk.CTkLabel(
            self.sidebar,
            text="SYSTEM STATUS",
            font=ctk.CTkFont(family="Consolas", size=11, weight="bold"),
            text_color="gray50"
        )
        dep_label.grid(row=len(self.CATEGORIES)+2, column=0, padx=20, pady=(0, 10), sticky="w")
        
        self.dep_status = ctk.CTkLabel(
            self.sidebar,
            text="",
            font=ctk.CTkFont(family="Consolas", size=10),
            text_color="gray60",
            wraplength=220,
            justify="left"
        )
        self.dep_status.grid(row=len(self.CATEGORIES)+3, column=0, padx=20, pady=0, sticky="w")
        
        # Update dependency status
        self._update_dep_status()
        
        # Theme toggle at bottom
        self.theme_switch = ctk.CTkSwitch(
            self.sidebar,
            text="DARK MODE",
            font=ctk.CTkFont(family="Consolas", size=11),
            command=self._toggle_theme,
            onvalue="Dark",
            offvalue="Light"
        )
        self.theme_switch.select()
        self.theme_switch.grid(row=15, column=0, padx=20, pady=30, sticky="s")
    
    def _create_main_area(self):
        """Create the main content area."""
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=30, pady=30)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)
        
        # Typewriter style header
        self.header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        
        self.category_title = ctk.CTkLabel(
            self.header_frame,
            text="SELECT_CATEGORY",
            font=ctk.CTkFont(family="Consolas", size=28, weight="bold")
        )
        self.category_title.pack(anchor="w")
        
        self.category_subtitle = ctk.CTkLabel(
            self.header_frame,
            text="",
            font=ctk.CTkFont(family="Roboto", size=13),
            text_color="gray60"
        )
        self.category_subtitle.pack(anchor="w", pady=(5, 0))
        
        # Content frame (split layout)
        self.content_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.content_frame.grid(row=1, column=0, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=3) # Drop zone wider
        self.content_frame.grid_columnconfigure(1, weight=2) # Settings narrower
        self.content_frame.grid_rowconfigure(0, weight=1)
        
        # Left: Drop zone
        self._create_drop_zone()
        
        # Right: Settings panel
        self._create_settings_panel()
        
        # Bottom: Progress and log
        self._create_progress_area()
    
    def _create_drop_zone(self):
        """Create the file drop zone with hover effects."""
        self.drop_frame = ctk.CTkFrame(
            self.content_frame,
            border_width=2,
            border_color=("gray70", "gray30"),
            fg_color=("gray95", "gray15"),
            corner_radius=8
        )
        self.drop_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 20), pady=0)
        
        # Center container
        center_frame = ctk.CTkFrame(self.drop_frame, fg_color="transparent")
        center_frame.place(relx=0.5, rely=0.45, anchor="center")
        
        self.drop_label = ctk.CTkLabel(
            center_frame,
            text="DROP FILES HERE",
            font=ctk.CTkFont(family="Consolas", size=18, weight="bold"),
            text_color="gray50"
        )
        self.drop_label.pack(pady=(0, 5))
        
        self.drop_sublabel = ctk.CTkLabel(
             center_frame,
             text="or click to browse",
             font=ctk.CTkFont(size=14),
             text_color="gray60"
        )
        self.drop_sublabel.pack()

        # File list (bottom part of drop zone)
        self.file_list = ctk.CTkTextbox(
            self.drop_frame,
            height=120,
            state="disabled",
            fg_color="transparent",
            font=ctk.CTkFont(family="Consolas", size=12),
            text_color="gray70"
        )
        self.file_list.place(relx=0.5, rely=0.8, anchor="center", relwidth=0.9)
        
        # Click bindings
        self.drop_frame.bind("<Button-1>", lambda e: self._browse_files())
        self.drop_label.bind("<Button-1>", lambda e: self._browse_files())
        self.drop_sublabel.bind("<Button-1>", lambda e: self._browse_files())
        
        # Hover bindings
        self.drop_frame.bind("<Enter>", self._on_drag_enter)
        self.drop_frame.bind("<Leave>", self._on_drag_leave)
        # Bind children to propagate (simple approximation)
        self.drop_label.bind("<Enter>", self._on_drag_enter)
        self.drop_sublabel.bind("<Enter>", self._on_drag_enter)
        
        # Drag and drop
        if DND_AVAILABLE:
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self._on_drop)
            
    def _on_drag_enter(self, event):
        """Handle mouse enter on drop zone."""
        # Change border to accent color
        self.drop_frame.configure(border_color=("#3B8ED0", "#1F6AA5"), border_width=3)
        self.drop_label.configure(text_color=("#3B8ED0", "#1F6AA5"))

    def _on_drag_leave(self, event):
        """Handle mouse leave on drop zone."""
        # Reset border
        self.drop_frame.configure(border_color=("gray70", "gray30"), border_width=2)
        self.drop_label.configure(text_color="gray50")
    
    def _create_settings_panel(self):
        """Create the conversion settings panel."""
        self.settings_frame = ctk.CTkFrame(self.content_frame, fg_color=("gray90", "gray15"), corner_radius=10)
        self.settings_frame.grid(row=0, column=1, sticky="nsew", padx=0, pady=0)
        
        # Output format
        format_label = ctk.CTkLabel(
            self.settings_frame,
            text="OUTPUT FORMAT",
            font=ctk.CTkFont(family="Consolas", size=12, weight="bold"),
            text_color="gray50"
        )
        format_label.pack(anchor="w", padx=20, pady=(25, 5))
        
        self.format_menu = ctk.CTkOptionMenu(
            self.settings_frame,
            variable=self.output_format,
            values=["Select format..."],
            width=200,
            height=35,
            font=ctk.CTkFont(family="Consolas", size=13),
            dropdown_font=ctk.CTkFont(family="Consolas", size=13)
        )
        self.format_menu.pack(anchor="w", padx=20, pady=(0, 15))
        
        # Output directory
        dir_label = ctk.CTkLabel(
            self.settings_frame,
            text="DESTINATION",
            font=ctk.CTkFont(family="Consolas", size=12, weight="bold"),
            text_color="gray50"
        )
        dir_label.pack(anchor="w", padx=20, pady=(15, 5))
        
        self.dir_entry = ctk.CTkEntry(
            self.settings_frame,
            width=200,
            state="readonly",
            font=ctk.CTkFont(family="Consolas", size=12),
            border_color=("gray70", "gray30")
        )
        self.dir_entry.pack(anchor="w", padx=20, pady=(0, 5))
        self._update_dir_display()
        
        dir_browse_btn = ctk.CTkButton(
            self.settings_frame,
            text="Browse...",
            width=100,
            height=28,
            font=ctk.CTkFont(size=12),
            fg_color=("gray80", "gray25"),
            text_color=("gray20", "gray80"),
            hover_color=("gray70", "gray30"),
            command=self._browse_output_dir
        )
        dir_browse_btn.pack(anchor="w", padx=20, pady=(0, 20))
        
        # Quality slider
        quality_label = ctk.CTkLabel(
            self.settings_frame,
            text="QUALITY / COMPRESSION",
            font=ctk.CTkFont(family="Consolas", size=12, weight="bold"),
            text_color="gray50"
        )
        quality_label.pack(anchor="w", padx=20, pady=(15, 5))
        
        self.quality_var = tk.IntVar(value=85)
        self.quality_slider = ctk.CTkSlider(
            self.settings_frame,
            from_=1,
            to=100,
            variable=self.quality_var,
            width=200,
            progress_color=("#3B8ED0", "#1F6AA5")
        )
        self.quality_slider.pack(anchor="w", padx=20, pady=(0, 5))
        
        self.quality_value_label = ctk.CTkLabel(
            self.settings_frame,
            text="85%",
            font=ctk.CTkFont(family="Consolas", size=12)
        )
        self.quality_value_label.pack(anchor="w", padx=20, pady=(0, 20))
        
        self.quality_var.trace_add("write", lambda *_: self.quality_value_label.configure(
            text=f"{self.quality_var.get()}%"
        ))
        
        # Convert button
        self.convert_btn = ctk.CTkButton(
            self.settings_frame,
            text="INITIALIZE CONVERSION",
            font=ctk.CTkFont(family="Consolas", size=14, weight="bold"),
            height=45,
            command=self._start_conversion
        )
        self.convert_btn.pack(side="bottom", padx=20, pady=25, fill="x")
    
    def _create_progress_area(self):
        """Create the progress display area."""
        self.progress_frame = ctk.CTkFrame(self.main_frame, height=120, fg_color="transparent")
        self.progress_frame.grid(row=2, column=0, sticky="ew", pady=(20, 0))
        self.progress_frame.grid_columnconfigure(0, weight=1)
        
        # Log box with monospace font
        self.log_box = ctk.CTkTextbox(
            self.progress_frame,
            height=80,
            state="disabled",
            font=ctk.CTkFont(family="Consolas", size=11),
            fg_color=("gray95", "gray10"),
            border_width=1,
            border_color=("gray80", "gray25"),
            text_color="gray70"
        )
        self.log_box.pack(fill="x", padx=0, pady=(0, 10))
        
        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(self.progress_frame, height=6)
        self.progress_bar.pack(fill="x", padx=0, pady=(0, 5))
        self.progress_bar.set(0)
        
        self.progress_label = ctk.CTkLabel(
            self.progress_frame,
            text="SYSTEM READY",
            font=ctk.CTkFont(family="Consolas", size=11),
            text_color="gray50"
        )
        self.progress_label.pack(anchor="w", padx=0, pady=(0, 0))
    
    def _select_category(self, category: ConversionCategory):
        """Select a conversion category."""
        self.current_category = category
        
        # Update button states for Techy Look
        for cat, btn in self.category_buttons.items():
            if cat == category:
                # Active: Accent color text, highlighted background
                btn.configure(
                    fg_color=("gray85", "gray20"),
                    text_color=("#3B8ED0", "#1F6AA5"),
                    font=ctk.CTkFont(family="Roboto Mono", size=13, weight="bold")
                )
            else:
                # Inactive: Gray text, transparent bg
                btn.configure(
                    fg_color="transparent",
                    text_color=("gray40", "gray60"),
                    font=ctk.CTkFont(family="Roboto Mono", size=13)
                )
        
        # Update header
        for icon, name, cat in self.CATEGORIES:
            if cat == category:
                self.category_title.configure(text=name) # Just name, tech style
                break
        
        # Get converter for category
        converters = registry.get_converters(category)
        if converters:
            converter = converters[0]
            formats = converter.supported_output_formats
            self.output_format.set(formats[0] if formats else "")
            self.format_menu.configure(values=formats)
            
            input_fmts = ", ".join(f".{f}" for f in converter.supported_input_formats[:7])
            if len(converter.supported_input_formats) > 7:
                input_fmts += "..."
            self.category_subtitle.configure(text=f"INPUT: [{input_fmts}]")
        
        # Clear selected files
        self.selected_files.clear()
        self._update_file_list()
    
    def _update_dep_status(self):
        """Update dependency status displays."""
        deps = dependency_checker.check_all()
        status_lines = []
        
        for name, info in deps.items():
            status = "ONLINE" if info.available else "OFFLINE"
            # Using simple text indicators
            icon = "[+]" if info.available else "[-]"
            status_lines.append(f"{icon} {name.upper()}: {status}")
        
        self.dep_status.configure(text="\n".join(status_lines))
    
    def _toggle_theme(self):
        """Toggle between dark and light mode."""
        mode = self.theme_switch.get()
        ctk.set_appearance_mode(mode)
    
    def _browse_files(self):
        """Open file browser dialog."""
        if not self.current_category:
            return
        
        converters = registry.get_converters(self.current_category)
        if not converters:
            return
        
        # Build file type filter
        extensions = converters[0].supported_input_formats
        file_types = [("Supported files", " ".join(f"*.{ext}" for ext in extensions))]
        
        files = filedialog.askopenfilenames(filetypes=file_types)
        if files:
            self.selected_files = [Path(f) for f in files]
            self._update_file_list()
    
    def _browse_output_dir(self):
        """Browse for output directory."""
        dir_path = filedialog.askdirectory(initialdir=self.output_dir)
        if dir_path:
            self.output_dir = Path(dir_path)
            self._update_dir_display()
    
    def _update_dir_display(self):
        """Update the output directory display."""
        self.dir_entry.configure(state="normal")
        self.dir_entry.delete(0, "end")
        self.dir_entry.insert(0, str(self.output_dir))
        self.dir_entry.configure(state="readonly")
    
    def _update_file_list(self):
        """Update the file list display."""
        self.file_list.configure(state="normal")
        self.file_list.delete("0.0", "end")
        
        if self.selected_files:
            for f in self.selected_files:
                self.file_list.insert("end", f"> {f.name}\n")
            self.drop_label.configure(text=f"{len(self.selected_files)} FILE(S) LOADED")
            self.drop_sublabel.configure(text="Click to add more")
        else:
            self.drop_label.configure(text="DROP FILES HERE")
            self.drop_sublabel.configure(text="or click to browse")
        
        self.file_list.configure(state="disabled")
    
    def _on_drop(self, event):
        """Handle file drop."""
        if DND_AVAILABLE:
            files = self.tk.splitlist(event.data)
            self.selected_files.extend(Path(f) for f in files if Path(f).is_file())
            self._update_file_list()
    
    def _log(self, message: str):
        """Log a message to the log box."""
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{message}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
    
    def _start_conversion(self):
        """Start the conversion process."""
        if not self.selected_files:
            messagebox.showwarning("No Files", "Please select files to convert.")
            return
        
        if not self.output_format.get():
            messagebox.showwarning("No Format", "Please select an output format.")
            return
        
        if self.is_converting:
            return
        
        self.is_converting = True
        self.convert_btn.configure(state="disabled", text="PROCESSING...")
        
        # Run conversion in background thread
        thread = threading.Thread(target=self._convert_files, daemon=True)
        thread.start()
    
    def _convert_files(self):
        """Convert files in background thread."""
        try:
            converter = registry.find_converter(
                self.selected_files[0].suffix.lstrip('.'),
                self.output_format.get()
            )
            
            if not converter:
                self.after(0, lambda: self._log("[!] No suitable converter found"))
                return
            
            options = ConversionOptions(
                output_dir=self.output_dir,
                quality=self.quality_var.get(),
                overwrite=True
            )
            
            total = len(self.selected_files)
            for i, file_path in enumerate(self.selected_files):
                progress = i / total
                self.after(0, lambda p=progress: self.progress_bar.set(p))
                self.after(0, lambda f=file_path.name: self.progress_label.configure(
                    text=f"PROCESSING: {f}"
                ))
                
                result = converter.convert(file_path, self.output_format.get(), options)
                
                if result.success:
                    self.after(0, lambda r=result: self._log(f"[OK] {r.input_path.name}"))
                else:
                    self.after(0, lambda r=result: self._log(f"[FAIL] {r.input_path.name}: {r.error_message}"))
            
            self.after(0, lambda: self.progress_bar.set(1.0))
            self.after(0, lambda: self.progress_label.configure(text="OPERATION COMPLETE"))
            self.after(0, lambda: self._log(f"--- Completed {total} file(s) ---"))
            
        except Exception as e:
            self.after(0, lambda: self._log(f"[ERROR] {e}"))
        finally:
            self.after(0, self._conversion_complete)
    
    def _conversion_complete(self):
        """Called when conversion is complete."""
        self.is_converting = False
        self.convert_btn.configure(state="normal", text="INITIALIZE CONVERSION")


if __name__ == "__main__":
    app = ConverterApp()
    app.mainloop()
