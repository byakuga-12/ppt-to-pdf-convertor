import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import subprocess
import os
import threading

# Try to import drag-and-drop support
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# ─────────────────────────────────────────────
# PATH TO LIBREOFFICE ON WINDOWS
# If this doesn't work, find where LibreOffice
# is installed and update this path.
# ─────────────────────────────────────────────
LIBREOFFICE_PATH = r"C:\Program Files\LibreOffice\program\soffice.exe"


def convert_to_pdf(pptx_path):
    """
    Converts a .pptx file to .pdf using LibreOffice.
    The PDF is saved in the same folder as the .pptx file.
    """
    output_dir = os.path.dirname(pptx_path)  # Same folder as input file

    # Run LibreOffice silently in the background
    # SAL_USE_VCLPLUGIN=svp suppresses the printer connection popup on Windows
    env = os.environ.copy()
    env["SAL_USE_VCLPLUGIN"] = "svp"

    result = subprocess.run([
        LIBREOFFICE_PATH,
        "--headless",           # No LibreOffice window
        "--norestore",          # Don't show restore dialog
        "--nologo",             # No splash screen
        "--nofirststartwizard", # Skip first start wizard
        "--convert-to", "pdf",  # Convert to PDF
        pptx_path,              # Input file
        "--outdir", output_dir  # Output folder
    ], capture_output=True, text=True, env=env)

    # Check if conversion worked
    if result.returncode != 0:
        raise Exception(result.stderr)

    # Return path to the output PDF
    pdf_name = os.path.splitext(os.path.basename(pptx_path))[0] + ".pdf"
    return os.path.join(output_dir, pdf_name)


class App(tk.Tk if not DND_AVAILABLE else TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        # ── Window setup ──
        self.title("PPTX → PDF Converter")
        self.geometry("500x420")
        self.resizable(False, False)
        self.configure(bg="#0f0f0f")

        self.selected_file = None  # Stores the chosen file path

        self._build_ui()

        # Enable drag and drop if available
        if DND_AVAILABLE:
            self.drop_zone.drop_target_register(DND_FILES)
            self.drop_zone.dnd_bind("<<Drop>>", self._on_drop)

    def _build_ui(self):
        # ── Title ──
        tk.Label(
            self,
            text="PPTX → PDF",
            font=("Courier New", 22, "bold"),
            bg="#0f0f0f",
            fg="#00ff99"
        ).pack(pady=(30, 4))

        tk.Label(
            self,
            text="Convert PowerPoint files to PDF — offline, instantly.",
            font=("Courier New", 9),
            bg="#0f0f0f",
            fg="#555555"
        ).pack()

        # ── Drop Zone ──
        self.drop_zone = tk.Label(
            self,
            text="drag & drop .pptx here\nor click 'Browse'",
            font=("Courier New", 11),
            bg="#1a1a1a",
            fg="#444444",
            width=40,
            height=6,
            relief="flat",
            cursor="hand2"
        )
        self.drop_zone.pack(pady=24, padx=30)
        self.drop_zone.bind("<Button-1>", lambda e: self._browse())

        # ── Status label (shows selected file name) ──
        self.status_label = tk.Label(
            self,
            text="no file selected",
            font=("Courier New", 9),
            bg="#0f0f0f",
            fg="#555555",
            wraplength=440
        )
        self.status_label.pack()

        # ── Progress Bar (hidden until conversion starts) ──
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "green.Horizontal.TProgressbar",
            troughcolor="#1a1a1a",
            background="#00ff99",
            bordercolor="#0f0f0f",
            lightcolor="#00ff99",
            darkcolor="#00ff99"
        )
        self.progress = ttk.Progressbar(
            self,
            style="green.Horizontal.TProgressbar",
            orient="horizontal",
            length=380,
            mode="indeterminate"  # Bouncing animation since we can't track exact %
        )
        self.progress.pack(pady=(10, 0))
        self.progress.pack_forget()  # Hide until needed

        # ── Convert Button ──
        self.convert_btn = tk.Button(
            self,
            text="Convert to PDF",
            font=("Courier New", 11, "bold"),
            bg="#00ff99",
            fg="#0f0f0f",
            activebackground="#00cc77",
            activeforeground="#0f0f0f",
            relief="flat",
            padx=20,
            pady=10,
            cursor="hand2",
            command=self._start_conversion
        )
        self.convert_btn.pack(pady=20)

    def _browse(self):
        """Opens a file picker dialog to choose a .pptx file."""
        path = filedialog.askopenfilename(
            title="Select a PowerPoint file",
            filetypes=[("PowerPoint files", "*.pptx *.ppt")]
        )
        if path:
            self._set_file(path)

    def _on_drop(self, event):
        """Handles drag and drop — cleans up the path tkinter gives us."""
        path = event.data.strip().strip("{}")  # Remove curly braces Windows adds
        if path.lower().endswith((".pptx", ".ppt")):
            self._set_file(path)
        else:
            messagebox.showerror("Wrong file type", "Please drop a .pptx or .ppt file.")

    def _set_file(self, path):
        """Stores the selected file and updates the UI."""
        self.selected_file = path
        filename = os.path.basename(path)
        self.status_label.config(text=f"✓ {filename}", fg="#00ff99")
        self.drop_zone.config(text=f"{filename}", fg="#00ff99")

    def _start_conversion(self):
        """Starts the conversion in a background thread so the UI doesn't freeze."""
        if not self.selected_file:
            messagebox.showwarning("No file", "Please select a .pptx file first.")
            return

        # Check LibreOffice is actually installed
        if not os.path.exists(LIBREOFFICE_PATH):
            messagebox.showerror(
                "LibreOffice not found",
                f"LibreOffice not found at:\n{LIBREOFFICE_PATH}\n\n"
                "Please install LibreOffice or update the path in converter.py"
            )
            return

        # Disable button and show progress bar
        self.convert_btn.config(text="Converting...", state="disabled", bg="#333333", fg="#888888")
        self.status_label.config(text="Converting, please wait...", fg="#ffaa00")
        self.progress.pack(pady=(10, 0))   # Show the progress bar
        self.progress.start(12)            # Start the bouncing animation (12ms interval)

        # Run conversion in background thread (keeps UI responsive)
        thread = threading.Thread(target=self._convert_thread)
        thread.daemon = True
        thread.start()

    def _convert_thread(self):
        """Runs in background. Calls convert_to_pdf and updates UI when done."""
        try:
            output_path = convert_to_pdf(self.selected_file)
            # Update UI back on main thread
            self.after(0, self._on_success, output_path)
        except Exception as e:
            self.after(0, self._on_error, str(e))

    def _on_success(self, output_path):
        """Called when conversion succeeds."""
        self.progress.stop()               # Stop animation
        self.progress.pack_forget()        # Hide progress bar
        self.convert_btn.config(text="Convert to PDF", state="normal", bg="#00ff99", fg="#0f0f0f")
        self.status_label.config(text=f"✓ Saved: {os.path.basename(output_path)}", fg="#00ff99")
        messagebox.showinfo("Done!", f"PDF saved to:\n{output_path}")
        # Reset for next file
        self.selected_file = None
        self.drop_zone.config(text="drag & drop .pptx here\nor click 'Browse'", fg="#444444")

    def _on_error(self, error_msg):
        """Called when conversion fails."""
        self.progress.stop()               # Stop animation
        self.progress.pack_forget()        # Hide progress bar
        self.convert_btn.config(text="Convert to PDF", state="normal", bg="#00ff99", fg="#0f0f0f")
        self.status_label.config(text="Conversion failed.", fg="#ff4444")
        messagebox.showerror("Error", f"Conversion failed:\n{error_msg}")


# ── Run the app ──
if __name__ == "__main__":
    app = App()
    app.mainloop()
