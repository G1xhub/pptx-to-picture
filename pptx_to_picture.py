import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
from PIL import Image
import tempfile
import win32com.client
import pythoncom
import threading

try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False
    DND_FILES = None

class PPTXConverter:
    def __init__(self):
        if DND_AVAILABLE:
            self.root = TkinterDnD.Tk()
        else:
            self.root = tk.Tk()
        self.root.title("PPTX to Picture - Drag & Drop")
        self.root.geometry("1400x800")
        self.root.configure(bg='#1e1e1e')
        self.root.resizable(True, True)

        self.output_dir = os.path.expanduser("~/Pictures")
        self.image_format = tk.StringVar(value="PNG")
        self.quality = tk.IntVar(value=95)
        self.number_slides = tk.BooleanVar(value=True)

        self.setup_ui()

    def setup_ui(self):
        main_frame = tk.Frame(self.root, bg='#1e1e1e')
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)

        left_panel = tk.LabelFrame(main_frame, text=" Drag & Drop Input Files", fg='white', bg='#2d2d2d', font=('Arial', 12, 'bold'))
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        formats_text = "Supported:\n* .pptx .ppt .ppsx .pps\n* .pdf\n* .odp\n\n**Each slide/page -> individual image!**\n\nDrag files here or click button"
        tk.Label(left_panel, text=formats_text, bg='#2d2d2d', fg='#cccccc', font=('Arial', 10), justify=tk.LEFT).pack(pady=20, padx=20)

        self.drop_zone = tk.Frame(left_panel, height=250, bg='#3d3d3d', relief=tk.GROOVE, bd=3)
        self.drop_zone.pack(fill=tk.BOTH, expand=True, pady=10)
        tk.Label(self.drop_zone, text="DROP FILES HERE", bg='#3d3d3d', fg='white', font=('Arial', 20, 'bold')).pack(expand=True)

        if DND_AVAILABLE:
            self.drop_zone.drop_target_register(DND_FILES)
            self.drop_zone.dnd_bind('<<Drop>>', self.on_drop)

        tk.Button(left_panel, text=" SELECT FILES", command=self.select_files, bg='#0078d4', fg='white', font=('Arial', 12, 'bold'), pady=10).pack(pady=20)

        right_panel = tk.LabelFrame(main_frame, text=" Output Settings", fg='white', bg='#2d2d2d', font=('Arial', 12, 'bold'))
        right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))

        tk.Label(right_panel, text="Format:", bg='#2d2d2d', fg='white', font=('Arial', 11, 'bold')).pack(anchor=tk.W, padx=20, pady=(20, 5))
        format_frame = tk.Frame(right_panel, bg='#2d2d2d')
        format_frame.pack(fill=tk.X, padx=20, pady=5)
        for fmt in ['PNG', 'JPG', 'BMP']:
            tk.Radiobutton(format_frame, text=fmt, variable=self.image_format, value=fmt, bg='#2d2d2d', fg='white', selectcolor='#0078d4').pack(side=tk.LEFT, padx=20)

        tk.Label(right_panel, text="JPG Quality:", bg='#2d2d2d', fg='white', font=('Arial', 11, 'bold')).pack(anchor=tk.W, padx=20, pady=(20, 5))
        self.quality_scale = tk.Scale(right_panel, from_=1, to=100, orient=tk.HORIZONTAL, variable=self.quality, bg='#2d2d2d', fg='white', troughcolor='#3d3d3d')
        self.quality_scale.pack(fill=tk.X, padx=20, pady=5)

        tk.Checkbutton(right_panel, text="Number slides (slide_1, slide_2...)", variable=self.number_slides, bg='#2d2d2d', fg='white', selectcolor='#0078d4').pack(anchor=tk.W, padx=20, pady=10)

        tk.Label(right_panel, text="Output Directory:", bg='#2d2d2d', fg='white', font=('Arial', 11, 'bold')).pack(anchor=tk.W, padx=20, pady=(20, 5))
        dir_frame = tk.Frame(right_panel, bg='#2d2d2d')
        dir_frame.pack(fill=tk.X, padx=20, pady=5)
        self.dir_label = tk.Label(dir_frame, text=self.output_dir, bg='#3d3d3d', fg='#cccccc', relief=tk.SUNKEN, anchor=tk.W, padx=10, pady=5)
        self.dir_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        tk.Button(dir_frame, text="Browse", command=self.browse_dir, bg='#0078d4', fg='white').pack(side=tk.RIGHT)

        tk.Label(right_panel, text="Log:", bg='#2d2d2d', fg='white', font=('Arial', 11, 'bold')).pack(anchor=tk.W, padx=20, pady=(20, 5))
        self.log = tk.Text(right_panel, height=15, bg='#3d3d3d', fg='white', state=tk.DISABLED)
        self.log.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)

        if not DND_AVAILABLE:
            self.log_msg("Warning: tkinterdnd2 not installed. Drag & Drop disabled.")
            self.log_msg("Install with: pip install tkinterdnd2")

    def log_msg(self, msg):
        self.log.config(state=tk.NORMAL)
        self.log.insert(tk.END, msg + '\n')
        self.log.see(tk.END)
        self.log.config(state=tk.DISABLED)

    def browse_dir(self):
        dir = filedialog.askdirectory(initialdir=self.output_dir)
        if dir:
            self.output_dir = dir
            self.dir_label.config(text=dir)

    def on_drop(self, event):
        if DND_AVAILABLE:
            files = self.root.tk.splitlist(event.data)
            for f in files:
                if os.path.isfile(f):
                    threading.Thread(target=self.convert, args=(f,), daemon=True).start()

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PPTX", "*.pptx *.ppt *.ppsx *.pps *.pdf *.odp")])
        for f in files:
            threading.Thread(target=self.convert, args=(f,), daemon=True).start()

    def convert(self, file_path):
        try:
            self.log_msg(f"Converting: {os.path.basename(file_path)}")
            
            abs_path = os.path.abspath(file_path)
            abs_path = os.path.normpath(abs_path)
            abs_path = os.path.realpath(abs_path)
            
            if not os.path.isfile(abs_path):
                raise FileNotFoundError(f"Datei nicht gefunden: {abs_path}")
            if not os.access(abs_path, os.R_OK):
                raise PermissionError(f"Keine Leseberechtigung für Datei: {abs_path}")
            
            # Validate file extension - only allow PowerPoint/PDF/ODP files
            valid_extensions = ('.pptx', '.ppt', '.ppsx', '.pps', '.pdf', '.odp')
            file_ext = os.path.splitext(abs_path)[1].lower()
            if file_ext not in valid_extensions:
                raise ValueError(f"Ungültiger Dateityp: {file_ext}. Unterstützte Formate: {', '.join(valid_extensions)}")

            pp = None
            pres = None

            try:
                try:
                    pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
                except:
                    pythoncom.CoInitialize()
                pp = win32com.client.Dispatch("PowerPoint.Application")
                pp.Visible = True
                # Normalize path for PowerPoint COM - ensure it's an absolute normalized path
                pptx_path = os.path.abspath(abs_path)
                pptx_path = os.path.normpath(pptx_path)
                pptx_path = str(pptx_path)
                # PowerPoint COM expects positional parameters with integer values (1/0) instead of boolean (True/False)
                # Presentations.Open(FileName, ReadOnly, Untitled, WithWindow)
                try:
                    pres = pp.Presentations.Open(pptx_path, 1, 0, 0)
                except Exception as open_error:
                    # Fallback: try with just the filename
                    try:
                        pres = pp.Presentations.Open(pptx_path)
                    except Exception as open_error2:
                        raise open_error2

                base = os.path.splitext(os.path.basename(file_path))[0]
                temp_dir = tempfile.gettempdir()

                for i in range(1, pres.Slides.Count + 1):
                    temp_img = os.path.join(temp_dir, f"temp_{i}.jpg")
                    pres.Slides(i).Export(temp_img, "JPG")

                    img = Image.open(temp_img)
                    num = f"_slide_{i}" if self.number_slides.get() else ""
                    final = os.path.join(self.output_dir, f"{base}{num}.{self.image_format.get().lower()}")

                    if self.image_format.get() == "PNG":
                        img.save(final, "PNG")
                    elif self.image_format.get() == "JPG":
                        img.save(final, "JPEG", quality=self.quality.get())
                    else:
                        img.save(final, self.image_format.get())

                    os.remove(temp_img)
                    self.log_msg(f"  Saved: {os.path.basename(final)}")

                self.log_msg(f"Done: {os.path.basename(file_path)}")

            finally:
                if pres:
                    try:
                        pres.Close()
                    except:
                        pass
                if pp:
                    try:
                        pp.Quit()
                    except:
                        pass
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass

        except Exception as e:
            import traceback
            self.log_msg(f"Error: {str(e)}")
            self.log_msg(f"Traceback: {traceback.format_exc()}")

if __name__ == "__main__":
    app = PPTXConverter()
    app.root.mainloop()
