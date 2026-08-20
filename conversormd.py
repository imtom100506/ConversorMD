import base64
import ctypes
import os
import sys
import threading
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD

# Corrección de pixelado / Escalado High-DPI en Windows
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# Logo de GitHub en Base64 (24x24 px) integrado para no depender de archivos externos
GITHUB_ICON_B64 = """
iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAABmJLR0QA/wD/AP+gvaeTAAAACXBI
WXMAAAsTAAALEwEAmpwYAAAAB3RJTUUH5QgKDRYw2mR9egAAAkNJREFUSMft1U1rE1EUBuDn3nEz
mSRNJqm12k/8bKKilbqwULuwoIu6q36CbnSnC3fvD+gvELsRxY2giIuCiFh040dQKVbEWtpmmvky
mTv3HhcdaVNrkzaC4Flfc86c87znwVku5P/OwtgT3w8UfD/Zc5xzrZTWVghRFEJc3Xv05OQ2nKPt
B53H14WUD4QQLYV4qK29uHrv4cv/hZ1fuvh1T6mXWmsZpQxKKTp169n+Y8/ff7uVbB/88M/1rVj3
iXOu1Xg81u12m9vtdhJFUdxsNs8+ePP9wU8x71fKvZVS03g8Nq1Wy7TbbdNoNLb2fD4XkUgkIqU8
8tP1x33ft77vp+Px2Ph+kEqlTq6urp76/m71iTGWSSaTSS6XMyEEiqLAsixYlgWtdZBKpUqZTCYj
hLhPjFnr9dpUKhVTW61WPdd1E9Vq9ZSU8gSAvNZaCSGglIIxBsYYGGPY2toSpVKpdX19fSOVSq2l
02krhEAIgUKhwDAMmKZpSilPAJg/L/rRzXw+X9FaE+ccQggYY7Bte2u73W7uK01MTEzkMpmMUEoh
hEAQBEgpIYTQ1trrAG4AuHHz9s2jWmsSQuzJcM4RxzHiOE7jOH60Vz50gW01hBAd53zfcRydTCZ7
xhgEgY84jl8BuA/geR7yeeY/QgjTNM131lre6/W2fd/fdV2XZ13318m/mE8A7/kQx/ELz/NeCyGm
AfiEkD1yzhFC4DiODsPw3n33y9H/7R8YmYkF3c6kfgAAAABJRU5ErkJggg==
"""


class MarkdownConverterApp:
    def __init__(self, root: TkinterDnD.Tk):
        self.root = root
        self.root.title("ConversorMD")
        self.root.geometry("680x620")
        self.root.minsize(620, 560)
        self.root.configure(bg="#f8fafc")

        self.engine = None
        self.files_queue = set()
        self.output_dir = None
        self.is_processing = False

        self.supported_extensions = {
            ".pdf", ".docx", ".pptx", ".xlsx", ".html", 
            ".htm", ".txt", ".csv", ".json", ".xml"
        }

        self._setup_theme()
        self._build_ui()

    def _setup_theme(self):
        self.style = ttk.Style(self.root)
        if "vista" in self.style.theme_names():
            self.style.theme_use("vista")
        elif "clam" in self.style.theme_names():
            self.style.theme_use("clam")

        self.style.configure(".", font=("Segoe UI", 9), background="#f8fafc")
        
        self.style.configure(
            "Accent.TButton",
            font=("Segoe UI", 10, "bold"),
            background="#2563eb",
            foreground="#ffffff",
            padding=(10, 8)
        )
        self.style.map("Accent.TButton",
            background=[("active", "#1d4ed8"), ("disabled", "#94a3b8")],
            foreground=[("disabled", "#f1f5f9")]
        )

        self.style.configure("Standard.TButton", font=("Segoe UI", 9), padding=(8, 5))
        self.style.configure("TLabel", background="#f8fafc", foreground="#1e293b")
        self.style.configure("TLabelframe", background="#f8fafc")
        self.style.configure("TLabelframe.Label", font=("Segoe UI", 9, "bold"), foreground="#334155", background="#f8fafc")

    def _open_github(self, event=None):
        webbrowser.open_new_tab("https://github.com/imtom100506")

    def _build_ui(self):
        # 1. Cabecera
        header_frame = tk.Frame(self.root, bg="#ffffff", height=60, bd=0, highlightthickness=1, highlightbackground="#e2e8f0")
        header_frame.pack(fill="x", side="top")
        
        # Bloque Izquierdo: Título y Créditos
        title_box = tk.Frame(header_frame, bg="#ffffff")
        title_box.pack(side="left", padx=20, pady=12)

        lbl_app_name = tk.Label(
            title_box, 
            text="ConversorMD", 
            font=("Segoe UI", 13, "bold"), 
            fg="#0f172a", 
            bg="#ffffff"
        )
        lbl_app_name.pack(side="left")

        lbl_sub = tk.Label(
            title_box, 
            text="Developed by Tom", 
            font=("Segoe UI", 9, "italic"), 
            fg="#64748b", 
            bg="#ffffff"
        )
        lbl_sub.pack(side="left", padx=(8, 0), pady=(3, 0))

        # Bloque Derecho: Botón interactivo de GitHub
        try:
            self.github_icon = tk.PhotoImage(data=GITHUB_ICON_B64)
            btn_github = tk.Label(
                header_frame, 
                image=self.github_icon, 
                bg="#ffffff", 
                cursor="hand2"
            )
        except Exception:
            btn_github = tk.Label(
                header_frame, 
                text="GitHub ↗", 
                font=("Segoe UI", 9, "bold"),
                fg="#2563eb", 
                bg="#ffffff", 
                cursor="hand2"
            )

        btn_github.pack(side="right", padx=20, pady=12)
        btn_github.bind("<Button-1>", self._open_github)
        btn_github.bind("<Enter>", lambda e: btn_github.configure(bg="#f1f5f9"))
        btn_github.bind("<Leave>", lambda e: btn_github.configure(bg="#ffffff"))

        # Contenedor principal
        main_container = tk.Frame(self.root, bg="#f8fafc")
        main_container.pack(fill="both", expand=True, padx=20, pady=15)

        # 2. Zona de Arrastrar y Soltar (Solo Archivos)
        self.drop_frame = tk.Frame(
            main_container, 
            bg="#f1f5f9", 
            bd=1, 
            relief="solid", 
            cursor="hand2",
            highlightthickness=1,
            highlightbackground="#cbd5e1"
        )
        self.drop_frame.pack(fill="x", pady=(0, 10), ipady=18)
        self.drop_frame.drop_target_register(DND_FILES)
        self.drop_frame.dnd_bind("<<Drop>>", self._on_drop)
        self.drop_frame.dnd_bind("<<DropEnter>>", lambda e: self.drop_frame.configure(bg="#e2e8f0", highlightbackground="#2563eb"))
        self.drop_frame.dnd_bind("<<DropLeave>>", lambda e: self.drop_frame.configure(bg="#f1f5f9", highlightbackground="#cbd5e1"))

        self.drop_label = tk.Label(
            self.drop_frame,
            text="Arrastra y suelta tus archivos aquí\n(PDF, Word, PowerPoint, Excel, etc.)",
            bg="#f1f5f9",
            fg="#334155",
            font=("Segoe UI", 10, "bold")
        )
        self.drop_label.pack(expand=True)

        # 3. Botón de búsqueda de archivos
        btn_frame = ttk.Frame(main_container)
        btn_frame.pack(fill="x", pady=5)

        btn_files = ttk.Button(
            btn_frame, 
            text="Examinar Archivos...", 
            style="Standard.TButton", 
            command=self._browse_files
        )
        btn_files.pack(fill="x", expand=True)

        # 4. Carpeta de destino
        dest_frame = ttk.LabelFrame(main_container, text=" Ubicación de guardado ", padding=10)
        dest_frame.pack(fill="x", pady=10)

        self.lbl_dest = ttk.Label(
            dest_frame, 
            text="Destino: Misma carpeta que los archivos originales", 
            foreground="#2563eb",
            wraplength=460
        )
        self.lbl_dest.pack(side="left", fill="x", expand=True)

        btn_dest = ttk.Button(dest_frame, text="Cambiar Carpeta...", style="Standard.TButton", command=self._choose_output_dir)
        btn_dest.pack(side="right", padx=(10, 0))

        # 5. Estado y consola de logs
        status_frame = ttk.Frame(main_container)
        status_frame.pack(fill="x", pady=(5, 2))

        self.lbl_queue_count = ttk.Label(
            status_frame, 
            text="Archivos en cola: 0", 
            font=("Segoe UI", 9, "bold")
        )
        self.lbl_queue_count.pack(side="left")

        btn_clear = ttk.Button(status_frame, text="Limpiar Cola", style="Standard.TButton", command=self._clear_queue)
        btn_clear.pack(side="right")

        log_container = tk.Frame(main_container, bd=1, relief="solid", highlightthickness=0, bg="#e2e8f0")
        log_container.pack(fill="both", expand=True, pady=5)

        self.log_text = tk.Text(
            log_container, 
            height=6, 
            state="disabled", 
            wrap="word", 
            font=("Consolas", 9),
            bg="#ffffff",
            fg="#0f172a",
            bd=0,
            padx=8,
            pady=8
        )
        self.log_text.pack(fill="both", expand=True)

        # 6. Barra de progreso y botón de conversión
        self.progress = ttk.Progressbar(main_container, mode="determinate")
        self.progress.pack(fill="x", pady=5)

        self.btn_convert = ttk.Button(
            main_container, 
            text="Convertir a Markdown", 
            style="Accent.TButton", 
            command=self._start_conversion_thread
        )
        self.btn_convert.pack(fill="x", pady=(5, 0))

    def _log(self, message: str):
        self.log_text.config(state="normal")
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def _add_paths_to_queue(self, paths: list):
        added = 0
        for p_str in paths:
            path = Path(p_str)
            if path.is_file() and path.suffix.lower() in self.supported_extensions:
                if path not in self.files_queue:
                    self.files_queue.add(path)
                    added += 1

        self.lbl_queue_count.config(text=f"Archivos en cola: {len(self.files_queue)}")
        if added > 0:
            self._log(f"Se agregaron {added} archivo(s).")
        else:
            self._log("No se encontraron archivos compatibles.")

    def _on_drop(self, event):
        self.drop_frame.configure(bg="#f1f5f9", highlightbackground="#cbd5e1")
        raw_files = self.root.tk.splitlist(event.data)
        self._add_paths_to_queue(raw_files)

    def _browse_files(self):
        types = [
            ("Archivos soportados", "*.pdf *.docx *.pptx *.xlsx *.html *.htm *.txt *.csv *.json *.xml"),
            ("Todos los archivos", "*.*")
        ]
        files = filedialog.askopenfilenames(title="Seleccionar archivos", filetypes=types)
        if files:
            self._add_paths_to_queue(files)

    def _choose_output_dir(self):
        selected = filedialog.askdirectory(title="Seleccionar carpeta de destino")
        if selected:
            self.output_dir = Path(selected)
            self.lbl_dest.config(text=f"Destino: {self.output_dir}")
        else:
            self.output_dir = None
            self.lbl_dest.config(text="Destino: Misma carpeta que los archivos originales")

    def _clear_queue(self):
        if self.is_processing:
            return
        self.files_queue.clear()
        self.lbl_queue_count.config(text="Archivos en cola: 0")
        self.progress["value"] = 0
        self._log("Cola limpiada.")

    def _start_conversion_thread(self):
        if not self.files_queue:
            messagebox.showwarning("Atención", "No hay archivos en la cola para convertir.")
            return

        if self.is_processing:
            return

        self.is_processing = True
        self.btn_convert.config(state="disabled")
        threading.Thread(target=self._process_queue, daemon=True).start()

    def _process_queue(self):
        if self.engine is None:
            self._log("Inicializando motor de conversión...")
            from markitdown import MarkItDown
            self.engine = MarkItDown()

        file_list = list(self.files_queue)
        total = len(file_list)
        self.progress["maximum"] = total
        self.progress["value"] = 0

        success = 0
        errors = 0

        self._log(f"\n--- Procesando {total} archivo(s) ---")

        for idx, file_path in enumerate(file_list, 1):
            try:
                self._log(f"[{idx}/{total}] Convirtiendo: {file_path.name}...")
                result = self.engine.convert(str(file_path))

                target_dir = self.output_dir if self.output_dir else file_path.parent
                target_dir.mkdir(parents=True, exist_ok=True)
                
                dest_file = target_dir / f"{file_path.stem}.md"
                dest_file.write_text(result.text_content, encoding="utf-8")

                success += 1
                self._log(f"  -> OK: {dest_file.name}")
            except Exception as e:
                errors += 1
                self._log(f"  -> ERROR en {file_path.name}: {str(e)}")

            self.progress["value"] = idx

        self._log(f"--- Completado. Exitosos: {success} | Errores: {errors} ---\n")
        self.files_queue.clear()
        self.lbl_queue_count.config(text="Archivos en cola: 0")
        self.is_processing = False
        self.btn_convert.config(state="normal")
        messagebox.showinfo("Finalizado", f"Conversión terminada.\nCorrectos: {success}\nFallidos: {errors}")


if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = MarkdownConverterApp(root)
    root.mainloop()