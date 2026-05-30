import tkinter as tk
from tkinter import messagebox, filedialog
from threading import Thread

# import your backend file
from scan_url_to_pdf_Version1 import run_scan_ui

def start_scan():
    url = url_entry.get().strip()

    try:
        max_pages = int(pages_entry.get())
        max_depth = int(depth_entry.get())
    except ValueError:
        messagebox.showerror("Invalid Input", "Pages and Depth must be numbers")
        return
    if max_pages > 200:
        messagebox.showwarning("Warning", "High page count may be slow")

    if not url:
        messagebox.showerror("Missing Input", "URL is required")
        return

    output_path = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
    )

    if not output_path:
        return

    status_label.config(text="Scanning...")

    def task():
        try:
            scanned, matched = run_scan_ui(
                url, max_pages, max_depth, output_path
            )
            status_label.config(
                text=f"Done! Scanned {scanned} pages, {matched} matched."
            )
        except Exception as e:
            status_label.config(text="Error")
            messagebox.showerror("Error", str(e))

    Thread(target=task, daemon=True).start()


# UI Setup
root = tk.Tk()
root.title("Keyword URL Scanner")
root.geometry("420x260")

# URL
tk.Label(root, text="URL").pack(pady=(10, 0))
url_entry = tk.Entry(root, width=50)
url_entry.pack()

# Max Pages
tk.Label(root, text="Max Pages").pack(pady=(10, 0))
pages_entry = tk.Entry(root)
pages_entry.insert(0, "100")
pages_entry.pack()

# Max Depth
tk.Label(root, text="Max Depth").pack(pady=(10, 0))
depth_entry = tk.Entry(root)
depth_entry.insert(0, "3")
depth_entry.pack()

# Button
tk.Button(root, text="Start Scan", command=start_scan).pack(pady=15)

# Status
status_label = tk.Label(root, text="")
status_label.pack()

root.mainloop()