import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from PIL import Image, ImageTk
import tempfile
import win32com.client
import pythoncom
import threading

# Configuration for CustomTkinter
ctk.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

DND_AVAILABLE = False
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    pass

class PPTXConverter(ctk.CTk, TkinterDnD.DnDWrapper if DND_AVAILABLE else object):
    def __init__(self):
        super().__init__()
        
        if DND_AVAILABLE:
            self.TkdndVersion = TkinterDnD._require(self)

        self.title("PPTX to Picture - Modern Converter")
        self.geometry("1100x700")
        
        # Configure grid layout (1x2)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.output_dir = os.path.expanduser("~/Pictures")
        self.image_format = tk.StringVar(value="PNG")
        self.quality = tk.IntVar(value=95)
        self.number_slides = tk.BooleanVar(value=True)
        self.current_preview_file = None
        self.preview_photo = None

        self.setup_ui()

    def setup_ui(self):
        self.create_sidebar()
        self.create_main_area()

    def create_sidebar(self):
        self.sidebar_frame = ctk.CTkFrame(self, width=250, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(8, weight=1)

        logo_label = ctk.CTkLabel(self.sidebar_frame, text="PPTX > Picture", font=ctk.CTkFont(size=20, weight="bold"))
        logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        # Output Format
        lbl_format = ctk.CTkLabel(self.sidebar_frame, text="Output Format:", anchor="w")
        lbl_format.grid(row=1, column=0, padx=20, pady=(10, 0), sticky="w")
        
        self.format_menu = ctk.CTkSegmentedButton(self.sidebar_frame, values=["PNG", "JPG", "BMP"], 
                                                  command=self.change_format_callback)
        self.format_menu.set("PNG")
        self.format_menu.grid(row=2, column=0, padx=20, pady=(5, 10), sticky="ew")

        # Quality (only valid/visible for JPG technically, but well keep it simple)
        lbl_quality = ctk.CTkLabel(self.sidebar_frame, text="JPG Quality:", anchor="w")
        lbl_quality.grid(row=3, column=0, padx=20, pady=(10, 0), sticky="w")
        
        self.quality_slider = ctk.CTkSlider(self.sidebar_frame, from_=1, to=100, variable=self.quality, number_of_steps=99)
        self.quality_slider.grid(row=4, column=0, padx=20, pady=(5, 10), sticky="ew")

        # Checkboxes
        self.chk_numbering = ctk.CTkSwitch(self.sidebar_frame, text="Number Slides", variable=self.number_slides, command=self.update_settings_preview)
        self.chk_numbering.grid(row=5, column=0, padx=20, pady=10, sticky="w")

        # Output Directory
        lbl_dir = ctk.CTkLabel(self.sidebar_frame, text="Output Directory:", anchor="w")
        lbl_dir.grid(row=6, column=0, padx=20, pady=(10, 0), sticky="w")
        
        self.entry_dir = ctk.CTkEntry(self.sidebar_frame, placeholder_text=self.output_dir)
        self.entry_dir.insert(0, self.output_dir)
        self.entry_dir.configure(state="readonly") # Make read-only so they have to use browse
        self.entry_dir.grid(row=7, column=0, padx=20, pady=(5, 5), sticky="ew")
        
        btn_browse = ctk.CTkButton(self.sidebar_frame, text="Browse Folder", command=self.browse_dir)
        btn_browse.grid(row=8, column=0, padx=20, pady=5, sticky="n")

        # Convert / Convert Button (maybe big at bottom?)
        # Let's put a help text at bottom of sidebar
        lbl_help = ctk.CTkLabel(self.sidebar_frame, text="Supported:\nPPTX, PPT, PDF, ODP", 
                                font=ctk.CTkFont(size=12), text_color="gray")
        lbl_help.grid(row=9, column=0, padx=20, pady=20, sticky="s")


    def create_main_area(self):
        self.main_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        
        self.main_frame.grid_rowconfigure(0, weight=1) # Drop zone
        self.main_frame.grid_rowconfigure(1, weight=1) # Preview/Log
        self.main_frame.grid_columnconfigure(0, weight=1)

        # -- Top: Drag & Drop Zone --
        self.drop_frame = ctk.CTkFrame(self.main_frame, border_width=2, border_color="gray")
        self.drop_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 10))
        
        self.drop_label = ctk.CTkLabel(self.drop_frame, text="DRAG & DROP FILES HERE\n\nor click to select", 
                                       font=ctk.CTkFont(size=20, weight="bold"))
        self.drop_label.place(relx=0.5, rely=0.5, anchor="center")

        # Make the whole frame clickable for file selection
        self.drop_label.bind("<Button-1>", lambda e: self.select_files())
        self.drop_frame.bind("<Button-1>", lambda e: self.select_files())

        if DND_AVAILABLE:
            self.drop_frame.drop_target_register(DND_FILES)
            self.drop_frame.dnd_bind('<<Drop>>', self.on_drop)

        # -- Bottom: Preview and Log Split --
        self.bottom_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        self.bottom_frame.grid(row=1, column=0, sticky="nsew")
        self.bottom_frame.grid_columnconfigure(0, weight=1) # Preview Image
        self.bottom_frame.grid_columnconfigure(1, weight=1) # File List / Log
        self.bottom_frame.grid_rowconfigure(0, weight=1)

        # Preview Image Area
        self.preview_box = ctk.CTkFrame(self.bottom_frame)
        self.preview_box.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        self.lbl_preview_img = ctk.CTkLabel(self.preview_box, text="No Preview Available")
        self.lbl_preview_img.pack(fill="both", expand=True, padx=10, pady=10)

        # Right side of bottom: File List + Log
        self.right_bottom_box = ctk.CTkFrame(self.bottom_frame)
        self.right_bottom_box.grid(row=0, column=1, sticky="nsew")
        
        # Output File Preview List
        ctk.CTkLabel(self.right_bottom_box, text="Output Files Preview:").pack(anchor="w", padx=10, pady=(5,0))
        self.file_list_box = ctk.CTkTextbox(self.right_bottom_box, height=100)
        self.file_list_box.pack(fill="x", padx=10, pady=5)
        self.file_list_box.configure(state="disabled")

        # Log
        ctk.CTkLabel(self.right_bottom_box, text="Processing Log:").pack(anchor="w", padx=10, pady=(5,0))
        self.log_box = ctk.CTkTextbox(self.right_bottom_box)
        self.log_box.pack(fill="both", expand=True, padx=10, pady=5)
        self.log_box.configure(state="disabled")

    def change_format_callback(self, value):
        self.image_format.set(value)
        self.update_settings_preview()

    def log_msg(self, msg):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", msg + '\n')
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def clear_preview(self):
        self.lbl_preview_img.configure(image=None, text="No Preview Available")
        self.file_list_box.configure(state="normal")
        self.file_list_box.delete("0.0", "end")
        self.file_list_box.configure(state="disabled")
        self.current_preview_file = None

    def update_preview(self, file_path):
        try:
            file_ext = os.path.splitext(file_path)[1].lower()

            if file_ext == '.odp':
                self.lbl_preview_img.configure(image=None, text=f"ODP File:\n{os.path.basename(file_path)}\n(No Preview)")
                slide_count = self.get_slide_count(file_path)
                self.generate_output_list(file_path, slide_count)
                self.current_preview_file = file_path
                return

            self.extract_preview_image(file_path)
            slide_count = self.get_slide_count(file_path)
            self.generate_output_list(file_path, slide_count)
            self.current_preview_file = file_path

        except Exception as e:
            self.lbl_preview_img.configure(image=None, text=f"Preview error:\n{str(e)}")
            self.file_list_box.configure(state="normal")
            self.file_list_box.delete("0.0", "end")
            self.file_list_box.configure(state="disabled")

    def extract_preview_image(self, file_path):
        temp_dir = tempfile.gettempdir()
        temp_img = os.path.join(temp_dir, "preview.jpg")
        file_ext = os.path.splitext(file_path)[1].lower()

        try:
            if file_ext in ('.pptx', '.ppt', '.ppsx', '.pps'):
                pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
                pp = win32com.client.Dispatch("PowerPoint.Application")
                # Keep hidden is safer, but sometimes causes issues. Try hidden first.
                try:
                    pp.Visible = True # Sometimes needed for export to work reliably
                    pp.WindowState = 2 # Minimize
                except:
                    pass

                try:
                    pptx_path = os.path.abspath(os.path.normpath(file_path))
                    pres = pp.Presentations.Open(pptx_path, 1, 0, 0)
                    pres.Slides(1).Export(temp_img, "JPG")
                    pres.Close()
                finally:
                    try:
                        pp.Quit()
                    except:
                        pass
                    pythoncom.CoUninitialize()

            elif file_ext == '.pdf':
                from pdf2image import convert_from_path
                images = convert_from_path(file_path, first_page=1, last_page=1, size=(400, None))
                if images:
                    images[0].save(temp_img, "JPEG")
                else:
                    raise Exception("No pages found in PDF")
            else:
                self.lbl_preview_img.configure(image=None, text="Format not supported for preview")
                return

            if os.path.exists(temp_img):
                img = Image.open(temp_img)
                # Resize keeping aspect ratio to fit in preview box (approx 300x200)
                # We can use CTkImage for high DPI support
                ctk_img = ctk.CTkImage(light_image=img, dark_image=img, size=(300, 225))
                
                self.lbl_preview_img.configure(image=ctk_img, text="")
                self.preview_photo = ctk_img # Keep ref
                
                # Cleanup temp later, or let system handle it
                # os.remove(temp_img) 
            else:
                 self.lbl_preview_img.configure(image=None, text="Preview generation failed")

        except Exception as e:
            self.lbl_preview_img.configure(image=None, text=f"Preview error:\n{str(e)}")

    def get_slide_count(self, file_path):
        file_ext = os.path.splitext(file_path)[1].lower()

        if file_ext in ('.pptx', '.ppt', '.ppsx', '.pps'):
            # Reuse logic (simplified for minimal com activity)
            pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
            pp = win32com.client.Dispatch("PowerPoint.Application")
            try: 
                pp.Visible = True
                pp.WindowState = 2
            except: pass
            
            try:
                pres = pp.Presentations.Open(os.path.abspath(file_path), 1, 0, 0)
                count = pres.Slides.Count
                pres.Close()
                return count
            except:
                return 1
            finally:
                try: pp.Quit()
                except: pass
                pythoncom.CoUninitialize()

        elif file_ext == '.pdf':
            try:
                from pdf2image import pdfinfo_from_path
                info = pdfinfo_from_path(file_path)
                return info["Pages"]
            except:
                return 1
        return 1

    def generate_output_list(self, file_path, slide_count):
        self.file_list_box.configure(state="normal")
        self.file_list_box.delete("0.0", "end")

        base = os.path.splitext(os.path.basename(file_path))[0]
        ext = self.image_format.get().lower()

        text = ""
        for i in range(1, slide_count + 1):
            num = f"_slide_{i}" if self.number_slides.get() else ""
            filename = f"{base}{num}.{ext}"
            text += filename + "\n"
        
        self.file_list_box.insert("0.0", text)
        self.file_list_box.configure(state="disabled")

    def update_settings_preview(self):
        if hasattr(self, 'current_preview_file') and self.current_preview_file:
            self.generate_output_list(self.current_preview_file, self.get_slide_count(self.current_preview_file))

    def browse_dir(self):
        dir = filedialog.askdirectory(initialdir=self.output_dir)
        if dir:
            self.output_dir = dir
            self.entry_dir.configure(state="normal")
            self.entry_dir.delete(0, "end")
            self.entry_dir.insert(0, dir)
            self.entry_dir.configure(state="readonly")
            self.update_settings_preview()

    def on_drop(self, event):
        if DND_AVAILABLE:
            files = self.Tk.splitlist(self, event.data)
            for f in files:
                if os.path.isfile(f):
                    self.update_preview(f)
                    threading.Thread(target=self.convert, args=(f,), daemon=True).start()

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("Files", "*.pptx *.ppt *.ppsx *.pps *.pdf *.odp")])
        for f in files:
            self.update_preview(f)
            threading.Thread(target=self.convert, args=(f,), daemon=True).start()

    def convert(self, file_path):
        try:
            self.log_msg(f"START: {os.path.basename(file_path)}")
            
            abs_path = os.path.abspath(file_path)
            
            # PowerPoint / Conversion Logic (Similar to before)
            file_ext = os.path.splitext(abs_path)[1].lower()
            
            if file_ext in ('.pptx', '.ppt', '.ppsx', '.pps'):
                pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
                pp = win32com.client.Dispatch("PowerPoint.Application")
                try: 
                    pp.Visible = True
                    pp.WindowState = 2 # Minimize
                except: pass

                pres = None
                try:
                    pres = pp.Presentations.Open(abs_path, 1, 0, 0)
                    base = os.path.splitext(os.path.basename(file_path))[0]
                    temp_dir = tempfile.gettempdir()

                    for i in range(1, pres.Slides.Count + 1):
                        temp_img = os.path.join(temp_dir, f"temp_{i}.jpg")
                        pres.Slides(i).Export(temp_img, "JPG")

                        img = Image.open(temp_img)
                        num = f"_slide_{i}" if self.number_slides.get() else ""
                        final = os.path.join(self.output_dir, f"{base}{num}.{self.image_format.get().lower()}")

                        fmt = self.image_format.get()
                        if fmt == "JPG":
                            img.save(final, "JPEG", quality=self.quality.get())
                        else:
                            img.save(final, fmt)

                        os.remove(temp_img)
                        self.log_msg(f"  > Saved: {os.path.basename(final)}")

                    pres.Close()
                finally:
                    try: pp.Quit()
                    except: pass
                    pythoncom.CoUninitialize()
            
            elif file_ext == '.pdf':
                from pdf2image import convert_from_path
                base = os.path.splitext(os.path.basename(file_path))[0]
                images = convert_from_path(abs_path)
                
                for i, img in enumerate(images):
                    num = f"_slide_{i+1}" if self.number_slides.get() else ""
                    final = os.path.join(self.output_dir, f"{base}{num}.{self.image_format.get().lower()}")
                    
                    fmt = self.image_format.get()
                    if fmt == "JPG":
                        img.save(final, "JPEG", quality=self.quality.get())
                    else:
                        img.save(final, fmt)
                    self.log_msg(f"  > Saved: {os.path.basename(final)}")

            self.log_msg(f"DONE: {os.path.basename(file_path)}")

        except Exception as e:
            import traceback
            self.log_msg(f"ERROR: {str(e)}")
            print(traceback.format_exc())

if __name__ == "__main__":
    app = PPTXConverter()
    app.mainloop()
