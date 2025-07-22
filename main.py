# pip install pywin32 pillow customtkinter
import customtkinter as ctk
import json, os, re, win32clipboard, base64
import sys

if sys.platform == "win32":
    import ctypes
    import ctypes.wintypes

def set_win_titlebar_color(window, color="#404040", text_color="#fafafa"):
    # Only works on Windows 10+ with dark mode enabled
    try:
        hwnd = int(window.wm_frame(), 16) if hasattr(window, "wm_frame") else window.winfo_id()
        DWMWA_USE_IMMERSIVE_DARK_MODE = 20
        DWMWA_CAPTION_COLOR = 35
        DWMWA_TEXT_COLOR = 36
        color_rgb = int(color.lstrip("#"), 16)
        text_rgb = int(text_color.lstrip("#"), 16)
        # Convert to COLORREF (0x00bbggrr)
        colorref = (color_rgb & 0xFF) << 16 | (color_rgb & 0xFF00) | (color_rgb >> 16 & 0xFF)
        textref = (text_rgb & 0xFF) << 16 | (text_rgb & 0xFF00) | (text_rgb >> 16 & 0xFF)
        dwmapi = ctypes.windll.dwmapi
        hwnd = ctypes.wintypes.HWND(hwnd)
        # Set title bar color
        dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_CAPTION_COLOR, ctypes.byref(ctypes.wintypes.DWORD(colorref)), ctypes.sizeof(ctypes.wintypes.DWORD))
        # Set title bar text color
        dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_TEXT_COLOR, ctypes.byref(ctypes.wintypes.DWORD(textref)), ctypes.sizeof(ctypes.wintypes.DWORD))
        # Enable dark mode for title bar
        dark = ctypes.wintypes.BOOL(1)
        dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ctypes.byref(dark), ctypes.sizeof(dark))
    except Exception:
        pass

class SDTemplatesApp:
    def __init__(self, root):
        self.root = root
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.json_path = os.path.join(self.script_dir, 'templates.json')
        self.pictures_dir = os.path.join(self.script_dir, 'Pictures')
        self.icon_path = os.path.join(self.script_dir, 'app.ico')
        self.data = self.load_data()
        self.category_var = ctk.StringVar()
        self.template_var = ctk.StringVar()
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")
        self.setup_ui()
        if sys.platform == "win32":
            set_win_titlebar_color(self.root, "#404040", "#fafafa")

    def load_data(self):
        with open(self.json_path, 'r', encoding='utf-8') as f:
            return json.load(f)

    def setup_ui(self):
        self.root.title("SDTemplates")
        self.root.geometry("600x220")
        if os.path.exists(self.icon_path):
            self.root.iconbitmap(self.icon_path)
        frame = ctk.CTkFrame(self.root, corner_radius=10)
        frame.pack(expand=True, fill='both', padx=20, pady=20)
        ctk.CTkLabel(frame, text="Kategori:", font=("Segoe UI", 14, "bold"), text_color="#fafafa").grid(row=0, column=0, sticky='w', padx=(0,8), pady=6)
        ctk.CTkLabel(frame, text="Mal:", font=("Segoe UI", 14, "bold"), text_color="#fafafa").grid(row=1, column=0, sticky='w', padx=(0,8), pady=6)

        self.category_menu = ctk.CTkOptionMenu(
            frame, variable=self.category_var, values=sorted(self.data.keys()), width=220,
            command=self.on_category_change
        )
        self.category_menu.grid(row=0, column=1, sticky='ew', padx=0, pady=6)
        self.template_menu = ctk.CTkOptionMenu(
            frame, variable=self.template_var, values=[], width=220
        )
        self.template_menu.grid(row=1, column=1, sticky='ew', padx=0, pady=6)

        copy_btn = ctk.CTkButton(
            frame,
            text="Kopier til utklippstavlen",
            command=self.copy_to_clipboard,
            fg_color="#3a28b7",
            hover_color="#2a1e7a",
            font=("Segoe UI", 14, "bold"),
            text_color="#fafafa",
            corner_radius=8,
            width=0  # Let the button size to fit content
        )
        copy_btn.grid(row=2, column=0, columnspan=2, pady=20, sticky='', padx=20)
        frame.columnconfigure(1, weight=1)
        self.category_var.set(sorted(self.data.keys())[0] if self.data else "")
        self.on_category_change(self.category_var.get())

    def on_category_change(self, value=None):
        cat = self.category_var.get()
        templates = sorted(self.data.get(cat, {}).keys())
        self.template_menu.configure(values=templates)
        if templates:
            self.template_var.set(templates[0])
        else:
            self.template_var.set("")

    def build_html(self, text, image_tags):
        html_parts = ["<p>" + text.replace("\n", "<br>") + "</p>"]
        for img_file in image_tags:
            img_path = os.path.join(self.pictures_dir, img_file.strip())
            if os.path.exists(img_path):
                try:
                    ext = os.path.splitext(img_path)[1].replace('.', '')
                    if ext.lower() == 'jpg':
                        ext = 'jpeg'
                    with open(img_path, 'rb') as f:
                        img_base64 = base64.b64encode(f.read()).decode('utf-8')
                    html_parts.append(f'<br><img src="data:image/{ext};base64,{img_base64}" />')
                except Exception as e:
                    self.custom_messagebox("Bildefeil", f"Feil ved bilde '{img_file}': {e}", icon="error")
        return ''.join(html_parts)

    def set_html_clipboard(self, html_content, plain_text):
        prefix = "<!DOCTYPE html><html><body>\r\n"
        fragment_start = "<!--StartFragment-->"
        fragment_end = "<!--EndFragment-->"
        suffix = "\r\n</body></html>"
        html_fragment = fragment_start + html_content + fragment_end
        full_html = prefix + html_fragment + suffix
        header_template = (
            "Version:1.0\r\n"
            "StartHTML:{starthtml:08d}\r\n"
            "EndHTML:{endhtml:08d}\r\n"
            "StartFragment:{startfragment:08d}\r\n"
            "EndFragment:{endfragment:08d}\r\n"
        )
        header = header_template.format(
            starthtml=0, endhtml=0, startfragment=0, endfragment=0
        )
        full_data = header + full_html
        header_bytes = header.encode('utf-8')
        start_html = len(header_bytes)
        end_html = start_html + len(full_html.encode('utf-8'))
        fragment_start_token = fragment_start.encode('utf-8')
        fragment_end_token = fragment_end.encode('utf-8')
        html_bytes = full_html.encode('utf-8')
        fragment_start_index = html_bytes.index(fragment_start_token) + len(fragment_start_token)
        fragment_end_index = html_bytes.index(fragment_end_token)
        start_fragment = start_html + fragment_start_index
        end_fragment = start_html + fragment_end_index
        header = header_template.format(
            starthtml=start_html,
            endhtml=end_html,
            startfragment=start_fragment,
            endfragment=end_fragment
        )
        final_data = header + full_html
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, plain_text)
        cf_html = win32clipboard.RegisterClipboardFormat("HTML Format")
        win32clipboard.SetClipboardData(cf_html, final_data.encode('utf-8'))
        win32clipboard.CloseClipboard()

    def custom_messagebox(self, title, message, icon="info"):
        win = ctk.CTkToplevel(self.root)
        win.title(title)
        win.geometry("340x180")
        win.resizable(False, False)
        frame = ctk.CTkFrame(win, corner_radius=10)
        frame.pack(expand=True, fill="both", padx=16, pady=16)
        icon_text = {"info": "ℹ️", "warning": "⚠️", "error": "❌"}.get(icon, "")
        ctk.CTkLabel(frame, text=icon_text, font=("Segoe UI Emoji", 24, "bold"), text_color="#fafafa").pack(pady=(18, 0))
        ctk.CTkLabel(frame, text=message, font=("Segoe UI", 11, "bold"), wraplength=260, justify="center", text_color="#fafafa").pack(pady=(8, 0), padx=10)
        ok_btn = ctk.CTkButton(frame, text="OK", command=win.destroy, font=("Segoe UI", 14, "bold"), text_color="#fafafa")
        ok_btn.pack(pady=12)
        win.grab_set()
        win.focus()
        win.wait_window()

    def copy_to_clipboard(self):
        category = self.category_var.get()
        template_key = self.template_var.get()
        if not (category and template_key):
            self.custom_messagebox("Mangler valg", "Velg både kategori og mal.", icon="warning")
            return
        raw_template = self.data[category][template_key]
        image_tags = re.findall(r"<img:(.*?)>", raw_template)
        placeholders = sorted(set(re.findall(r"{(.*?)}", raw_template)))
        # Dynamically calculate width based on longest placeholder
        min_entry_width = 220
        min_window_width = 420
        padding = 64  # increased padding for left/right
        if placeholders:
            max_ph_len = max((len(ph) for ph in placeholders), default=0)
            entry_width_px = max(min_entry_width, max_ph_len * 8 + 80)
            window_width_px = max(min_window_width, max_ph_len * 8 + 220 + padding)
        else:
            entry_width_px = min_entry_width
            window_width_px = min_window_width
        def apply_inputs(text, replacements):
            for key, val in replacements.items():
                text = text.replace(f"{{{key}}}", val)
            return text
        def finalize(replacements):
            processed_text = apply_inputs(raw_template, replacements)
            clean_text = re.sub(r"<img:.*?>", "", processed_text).strip()
            plain_text = re.sub(r'<[^>]+>', '', clean_text).strip()
            html_content = self.build_html(clean_text, image_tags)
            self.set_html_clipboard(html_content, plain_text)
            self.custom_messagebox("Kopiert", f"{template_key} er kopiert til utklippstavlen.", icon="info")
        if placeholders:
            input_window = ctk.CTkToplevel(self.root)
            input_window.title("Fyll inn verdier")
            input_window.geometry(f"{window_width_px}x{120+40*len(placeholders)}")
            input_window.resizable(False, False)
            frame = ctk.CTkFrame(input_window, corner_radius=10)
            frame.pack(expand=True, fill="both", padx=24, pady=16)
            entries = {}
            for i, ph in enumerate(placeholders):
                ctk.CTkLabel(frame, text=f"{ph}:", font=("Segoe UI", 14, "bold"), text_color="#fafafa").grid(row=i, column=0, sticky='w', padx=10, pady=(24,5) if i == 0 else 5)
                entry = ctk.CTkEntry(frame, width=entry_width_px, font=("Segoe UI", 14, "bold"), text_color="#fafafa")
                entry.grid(row=i, column=1, padx=10, pady=(24,5) if i == 0 else 5, sticky='ew')
                entries[ph] = entry
            frame.columnconfigure(1, weight=1)
            def on_submit():
                user_values = {ph: entries[ph].get().strip() for ph in placeholders}
                if any(not v for v in user_values.values()):
                    self.custom_messagebox("Manglende verdi", "Alle felt må fylles ut.", icon="warning")
                    return
                input_window.destroy()
                finalize(user_values)
            btn = ctk.CTkButton(frame, text="Kopier", command=on_submit, font=("Segoe UI", 14, "bold"), text_color="#fafafa")
            btn.grid(row=len(placeholders), column=0, columnspan=2, pady=10)
            input_window.after(10, lambda: list(entries.values())[0].focus())
            input_window.grab_set()
            input_window.focus()
            self.root.wait_window(input_window)
        else:
            finalize({})

if __name__ == "__main__":
    root = ctk.CTk()
    SDTemplatesApp(root)
    root.mainloop()