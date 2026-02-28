import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import re
import codecs
import pandas as pd


# ---------- Abrir archivo detectando BOM ----------
def open_text_file(path):
    with open(path, "rb") as f:
        raw = f.read()

    if raw.startswith(codecs.BOM_UTF16_LE) or raw.startswith(codecs.BOM_UTF16_BE):
        text = raw.decode("utf-16")
    elif raw.startswith(codecs.BOM_UTF8):
        text = raw.decode("utf-8")
    else:
        text = raw.decode("latin-1")

    return text.splitlines()


# ---------- Parser ----------
def parse_material_master(file_path):
    records = []
    buffer = None

    lines = open_text_file(file_path)

    for line in lines:
        line = line.strip()

        if not line:
            continue
        if line.startswith("Material") or line.startswith("Material description"):
            continue
        if re.match(r"\d{2}/\d{2}/\d{4}", line):
            continue

        # Línea técnica
        if re.match(r"^[A-Z0-9\-]+", line) and (
            "\t" in line or re.search(r"\s{2,}", line)
        ):
            buffer = re.split(r"\s{2,}|\t+", line)
            continue

        # Línea descripción
        if buffer:
            while len(buffer) < 5:
                buffer.append(None)

            records.append({
                "Material": buffer[0],
                "Plnt": buffer[1],
                "Matl grp": buffer[2],
                "CTyp": buffer[3],
                "Ctrl key": buffer[4],
                "Material description": line
            })

            buffer = None

    return records


# ---------- Variables globales ----------
parsed_records = []

# Loading modal window handle
loading_window = None


def show_loading(message="Procesando..."):
    global loading_window
    if loading_window:
        return
    loading_window = tk.Toplevel(root)
    loading_window.title("")
    loading_window.transient(root)
    loading_window.resizable(False, False)
    loading_window.grab_set()
    loading_window.protocol("WM_DELETE_WINDOW", lambda: None)

    lbl = tk.Label(loading_window, text=message, padx=20, pady=10)
    lbl.pack(fill="x", expand=True)

    pb = ttk.Progressbar(loading_window, mode="indeterminate", length=200)
    pb.pack(padx=10, pady=(0, 12))
    pb.start(10)

    # Center the loading window over root
    root.update_idletasks()
    x = root.winfo_rootx() + (root.winfo_width() // 2) - (loading_window.winfo_reqwidth() // 2)
    y = root.winfo_rooty() + (root.winfo_height() // 2) - (loading_window.winfo_reqheight() // 2)
    loading_window.geometry(f"+{x}+{y}")


def hide_loading():
    global loading_window
    if loading_window:
        try:
            loading_window.grab_release()
        except Exception:
            pass
        loading_window.destroy()
        loading_window = None


# ---------- Abrir archivo ----------
def open_file():
    global parsed_records
    file_path = filedialog.askopenfilename(
        title="Seleccionar archivo TXT",
        filetypes=[("Archivos de texto", "*.txt")]
    )

    if not file_path:
        return

    show_loading("Cargando archivo...")

    def worker():
        try:
            records = parse_material_master(file_path)

            def finish():
                global parsed_records
                parsed_records = records
                output.delete("1.0", tk.END)

                if not parsed_records:
                    output.insert(tk.END, "No se encontraron registros válidos.\n")
                    hide_loading()
                    return

                output.insert(tk.END, f"Registros encontrados: {len(parsed_records)}\n\n")
                for r in parsed_records:
                    output.insert(tk.END, "-" * 60 + "\n")
                    for k, v in r.items():
                        output.insert(tk.END, f"{k}: {v}\n")

                hide_loading()

            root.after(0, finish)

        except Exception as e:
            root.after(0, lambda: (hide_loading(), messagebox.showerror("Error", str(e))))

    threading.Thread(target=worker, daemon=True).start()


# ---------- Exportar a Excel ----------
def export_to_excel():
    if not parsed_records:
        messagebox.showwarning(
            "Sin datos",
            "Primero selecciona y procesa un archivo TXT."
        )
        return
    file_path = filedialog.asksaveasfilename(
        title="Guardar Excel",
        defaultextension=".xlsx",
        filetypes=[("Excel", "*.xlsx")]
    )

    if not file_path:
        return

    show_loading("Exportando a Excel...")

    def worker():
        try:
            df = pd.DataFrame(parsed_records)

            df = df[
                [
                    "Material",
                    "Plnt",
                    "Matl grp",
                    "CTyp",
                    "Ctrl key",
                    "Material description",
                ]
            ]

            df.to_excel(file_path, index=False)

            root.after(0, lambda: (hide_loading(), messagebox.showinfo("Éxito", "Archivo Excel exportado correctamente.")))

        except Exception as e:
            root.after(0, lambda: (hide_loading(), messagebox.showerror("Error", str(e))))

    threading.Thread(target=worker, daemon=True).start()


# ---------- Interfaz ----------
root = tk.Tk()
root.title("Material Master Parser (SAP)")
root.geometry("820x560")

frame = tk.Frame(root)
frame.pack(pady=10)

btn_open = tk.Button(
    frame,
    text="Seleccionar archivo TXT",
    font=("Segoe UI", 10, "bold"),
    command=open_file
)
btn_open.pack(side=tk.LEFT, padx=5)

btn_export = tk.Button(
    frame,
    text="Exportar a Excel",
    font=("Segoe UI", 10, "bold"),
    command=export_to_excel
)
btn_export.pack(side=tk.LEFT, padx=5)

output = scrolledtext.ScrolledText(
    root,
    wrap=tk.WORD,
    font=("Consolas", 10)
)
output.pack(expand=True, fill="both", padx=10, pady=10)

root.mainloop()
