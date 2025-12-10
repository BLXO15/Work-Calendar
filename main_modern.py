#By Matheus Caella Santis
import tkinter as tk
from tkinter import ttk

# Reuse the exact app behavior/structure from the existing implementation
from main import CalendarApp
from main import get_calendar_service, sync_from_google_calendar, ensure_google_credentials
from tkinter import font as tkfont
# update checking removed per user request
try:
    from PIL import Image, ImageDraw, ImageFont, ImageTk
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False


def apply_modern_style(root: tk.Tk) -> None:
    background = "#2a2a2a"
    surface = "#3d3d3d"
    text_primary = "#FFFFFF"
    text_muted = "#303030"
    stroke = "#1f1f1f"
    primary = "#e4941d"  # orange accents

    # Window background
    try:
        root.configure(bg=background)
    except Exception:
        pass

    style = ttk.Style()
    try:
        style.theme_use("clam")
    except Exception:
        pass

    for element in ("TFrame", "TLabelframe", "TLabel"):
        style.configure(element, background=background, foreground=text_primary)

    # Checkbutton styling (ttk.Checkbutton)
    try:
        style.configure("TCheckbutton", background=background, foreground=text_primary)
        style.map(
            "TCheckbutton",
            background=[("active", surface), ("!active", background)],
            foreground=[("active", text_primary), ("!active", text_primary)],
        )
    except Exception:
        pass

    # Group containers
    style.configure(
        "TLabelframe",
        relief="flat",
        borderwidth=0,
        padding=8,
        background=background,
        foreground=text_primary,
    )
    style.configure(
        "TLabelframe.Label",
        background=background,
        foreground=text_muted,
        padding=2,
    )

    # Buttons: flat, no outline, smaller padding (top buttons felt too big)
    style.configure(
        "TButton",
        relief="flat",
        borderwidth=0,
        padding=(8, 4),
        background=surface,
        foreground=text_primary,
        focusthickness=0,
        highlightthickness=0,
    )
    style.map(
        "TButton",
        background=[("pressed", surface), ("active", "#3F3F3F")],
        relief=[("pressed", "flat"), ("!pressed", "flat")],
        highlightthickness=[("focus", 0)],
    )

    style.configure(
        "Accent.TButton",
        background=primary,
        foreground="#363636",
        borderwidth=0,
        relief="flat",
        padding=(8, 4),
    )
    style.map(
        "Accent.TButton",
        background=[("pressed", "#2563eb"), ("active", "#60a5fa")],
        relief=[("pressed", "flat"), ("!pressed", "flat")],
    )

    # Entries
    style.configure("TEntry",
                   fieldbackground=surface,
                   background=surface,
                   foreground=text_primary,
                   bordercolor=stroke,
                   lightcolor=stroke,
                   darkcolor=stroke,
                   relief="flat",
                   padding=6)
    style.map("TEntry",
             fieldbackground=[("readonly", surface), ("focus", surface)],
             bordercolor=[("focus", primary), ("!focus", stroke)])

    # Treeview
    style.configure(
        "Treeview",
        background=surface,
        fieldbackground=surface,
        foreground=text_primary,
        borderwidth=0,
        relief="flat",
        rowheight=26,
    )
    style.configure(
        "Treeview.Heading",
        background=background,
        foreground=text_muted,
        relief="flat",
        borderwidth=0,
        padding=(6, 8),
    )
    style.map(
        "Treeview",
        background=[("selected", "#444444")],
        foreground=[("selected", text_primary)],
    )


    style.configure("TNotebook", background=background, borderwidth=0)
    style.configure("TNotebook.Tab", background=surface, padding=(10, 6))
    style.map("TNotebook.Tab", background=[("selected", surface), ("!selected", background)])

    root.option_add("*Button.relief", "flat")
    root.option_add("*Button.borderWidth", 0)
    root.option_add("*Button.highlightThickness", 0)
    root.option_add("*Button.background", surface)
    root.option_add("*Button.activeBackground", "#3d3d3d")
    root.option_add("*Button.foreground", text_primary)
    # Remove outlines from labels (weekends in calendar use tk.Label). They will just be grey
    # as set by the app, without a ridge/border.
    root.option_add("*Label.relief", "flat")
    root.option_add("*Label.background", background)
    root.option_add("*Label.foreground", text_primary)
    # Also apply defaults for Checkbutton widgets (ttk and tk)
    root.option_add("*Checkbutton.background", background)
    root.option_add("*Checkbutton.foreground", text_primary)
    
    # Override ALL scrollbar styling globally - Google style
    style.configure("TScrollbar",
                   background="#414141",
                   troughcolor="transparent",
                   borderwidth=0,
                   arrowcolor="#9ca3af",
                   darkcolor="#616161",
                   lightcolor="#555555",
                   relief="flat",
                   width=8)
    style.map("TScrollbar",
             background=[("active", "#474747"), ("pressed", "#444444")],
             arrowcolor=[("active", text_primary), ("pressed", text_primary)])
    
    # Also try to override the theme's scrollbar
    style.configure("Vertical.TScrollbar", 
                   background="#505050",
                   troughcolor="transparent",
                   borderwidth=0,
                   width=0)
    style.configure("Horizontal.TScrollbar", 
                   background="#474747",
                   troughcolor="transparent", 
                   borderwidth=0,
                   width=0)
    
    # Configure Combobox without arrows
    style.configure("TCombobox",
                   fieldbackground="#2a2a2a",
                   background="#2a2a2a",
                   foreground=text_primary,
                   bordercolor=stroke,
                   lightcolor=stroke,
                   darkcolor=stroke,
                   relief="flat",
                   padding=6,
                   arrowcolor="#2a2a2a",  # Make arrow same color as background (invisible)
                   arrowsize=0)  # Make arrow size 0
    style.map("TCombobox",
             fieldbackground=[("readonly", "#414141"), ("focus", "#2a2a2a")],
             bordercolor=[("focus", primary), ("!focus", stroke)],
             arrowcolor=[("active", "#3d3d3d"), ("pressed", "#2a2a2a")])  # Keep arrows invisible


def main():
    # Ensure Google credentials are present before launching the app
    ensure_google_credentials()
    root = tk.Tk()
    apply_modern_style(root)
    
    # Hide native title bar and add custom titlebar (from main_modern.py)
    root.overrideredirect(True)
    
    # Outer container to hold titlebar + content
    outer = tk.Frame(root, bg=root.cget('bg'))
    outer.pack(fill='both', expand=True)

    # make titlebar thinner
    TITLEBAR_HEIGHT = 24
    titlebar = tk.Frame(outer, bg=root.cget('bg'), height=TITLEBAR_HEIGHT)
    titlebar.pack(side='top', fill='x')
    titlebar.pack_propagate(False)

    # Dragging support
    drag = {'x': 0, 'y': 0}
    def _start_move(e):
        drag['x'] = e.x_root - root.winfo_x()
        drag['y'] = e.y_root - root.winfo_y()
    def _on_move(e):
        x = e.x_root - drag['x']
        y = e.y_root - drag['y']
        root.geometry(f"+{x}+{y}")

    titlebar.bind('<Button-1>', _start_move)
    titlebar.bind('<B1-Motion>', _on_move)

    # Sync button (left side of titlebar)
    sync_left = tk.Frame(titlebar, bg=root.cget('bg'))
    sync_left.pack(side='left', padx=(8, 0))
    
    # Will populate sync_btn after app is created
    sync_btn_ref = [None]  # Mutable reference to update later

    # Window controls (minimize, maximize, close)
    ctrl = tk.Frame(titlebar, bg="#141414")
    ctrl.pack(side='right')

    def _close():
        try:
            root.destroy()
        except Exception:
            root.quit()

    # smaller close button to match thinner titlebar
    small_font = tkfont.Font(size=9)
    btn_close = tk.Button(ctrl, text='✕', command=_close, bg="#555555", fg="#fff",
                          bd=0, activebackground="#c0392b", padx=6, pady=2, font=small_font)
    btn_close.pack(side='left', pady=(TITLEBAR_HEIGHT - 18)//2)

    # Content area for the application UI placed under the custom titlebar
    content = tk.Frame(outer, bg=root.cget('bg'))
    content.pack(fill='both', expand=True)
    
    # Lock window resizing but allow programmatic resizing
    root.resizable(False, False)
    
    # Always on top
    root.wm_attributes('-topmost', 1)
    try:
        import ctypes
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd == 0:
            hwnd = root.winfo_id()
        ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 3)
    except Exception:
        pass
    
    # Launch the same app structure/behavior, then apply modern tweaks
    app = CalendarApp(root)
    
    # Reparent app.main_frame into the content container (under the titlebar)
    try:
        if hasattr(app, 'main_frame'):
            app.main_frame.pack_forget()
            app.main_frame.pack(in_=content, fill='both', expand=True)
    except Exception:
        pass

    # Create sync button for titlebar
    cmd_sync = getattr(app, 'manual_sync', None)
    if not callable(cmd_sync):
        cmd_sync = (lambda: None)

    def run_sync_inline():
        if getattr(run_sync_inline, "_busy", False):
            return
        run_sync_inline._busy = True
        
        def launch_sync():
            import threading
            def do_sync():
                try:
                    service = get_calendar_service()
                    deleted_count, modified_count = sync_from_google_calendar(service, app.tasks, app.profiles)
                    def after_sync():
                        try:
                            # Refresh the view only if it is currently visible.
                            # Avoid switching views based solely on state variables
                            # which might cause the UI to navigate unexpectedly.
                            if getattr(app, 'all_tasks_frame', None) and app.all_tasks_frame.winfo_ismapped():
                                app.show_all_tasks()
                            elif getattr(app, 'calendar_frame', None) and app.calendar_frame.winfo_ismapped():
                                app.show_calendar()
                            elif getattr(app, 'day_view_frame', None) and app.day_view_frame.winfo_ismapped():
                                # If day view is visible, refresh it without forcing a selection change
                                try:
                                    if getattr(app, 'selected_day', None):
                                        app.show_day_view(app.selected_day)
                                except Exception:
                                    pass
                        except Exception:
                            pass
                        # Show checkmark then reset to arrow
                        try:
                            sync_btn_ref[0].configure(text='✓', bg='#16a34a')
                        except Exception:
                            pass
                        def _reset():
                            try:
                                sync_btn_ref[0].configure(text='↻', bg='#22c55e')
                            finally:
                                setattr(run_sync_inline, "_busy", False)
                        root.after(800, _reset)
                    root.after(0, after_sync)
                except Exception:
                    # Show error X then reset
                    def _err():
                        try:
                            sync_btn_ref[0].configure(text='✗', bg='#ef4444')
                        except Exception:
                            pass
                        def _reset():
                            try:
                                sync_btn_ref[0].configure(text='↻', bg='#22c55e')
                            finally:
                                setattr(run_sync_inline, "_busy", False)
                        root.after(800, _reset)
                    root.after(0, _err)
            threading.Thread(target=do_sync, daemon=True).start()
        
        launch_sync()

    # Create the actual sync button in the titlebar
    sync_btn = tk.Button(
        sync_left,
        text='↻',
        command=run_sync_inline,
        bg='#22c55e',
        fg='#ffffff',
        activebackground='#16a34a',
        bd=0,
        relief='flat',
        highlightthickness=0,
        padx=6,
        pady=2,
        font=tkfont.Font(size=8)
    )
    sync_btn.pack(side='left')
    sync_btn_ref[0] = sync_btn

    # Hover effect
    def _sync_hover_enter(e):
        try: sync_btn.configure(bg='#16a34a')
        except Exception: pass
    def _sync_hover_leave(e):
        try: sync_btn.configure(bg='#22c55e')
        except Exception: pass
    sync_btn.bind('<Enter>', _sync_hover_enter)
    sync_btn.bind('<Leave>', _sync_hover_leave)

    # Update-checking code removed per user request (no external GitHub linkage)

    # Add action buttons to titlebar (All Tasks, Profiles, Export PDF)
    # Create a middle section in titlebar for these buttons
    title_middle = tk.Frame(titlebar, bg=root.cget('bg'))
    title_middle.pack(side='left', padx=(12, 0))

    # Get the app commands (will be populated later after app is created)
    title_btn_refs = {'all_tasks': None, 'profiles': None, 'export': None}

    # 1) Remove weekend borders and add hover grey on day buttons after calendar render
    original_show_calendar = app.show_calendar

    def modern_show_calendar():
        original_show_calendar()
        try:
            # Remove weekend label borders (they are tk.Label with bg lightgray)
            for child in app.calendar_frame.winfo_children():
                # Weekday header row and day cells are mixed; weekend day cells in original code use tk.Label
                if isinstance(child, tk.Label):
                    try:
                        # Normalize weekend/greyed labels to a darker muted style
                        cur_bg = None
                        try:
                            cur_bg = child.cget("bg")
                        except Exception:
                            cur_bg = None
                        if cur_bg in ("lightgray", "#474747", "#3a3a3a"):
                            child.configure(relief="flat", bd=0, highlightthickness=0, bg="#2e2e2e", fg="#bfbfbf")
                    except Exception:
                        pass
                # Add hover effect for tk.Button day cells
                if isinstance(child, tk.Button):
                    def _on_enter(e, btn=child, orig=child.cget("background")):
                        try:
                            btn._orig_bg = orig
                            btn.configure(background="#636363", activebackground="#6b6b6b")
                        except Exception:
                            pass
                    def _on_leave(e, btn=child):
                        try:
                            bg = getattr(btn, "_orig_bg", None) or "SystemButtonFace"
                            btn.configure(background=bg, activebackground=bg)
                        except Exception:
                            pass
                    # Bind once
                    if not getattr(child, "_modern_bound", False):
                        child.bind("<Enter>", _on_enter)
                        child.bind("<Leave>", _on_leave)
                        child._modern_bound = True
        except Exception:
            pass

    # 2) Apply calendar styling
    app.show_calendar = modern_show_calendar

    # 3) Also apply on initial render
    app.show_calendar()

    # 3) Replace topbar buttons with a modern, small rounded look and add a left sync icon
    # Build modern topbar controls
    topbar = app.topbar
    # Safe background getter for ttk/tk widgets
    def _bg(widget, fallback="#4d4d4d"):
        try:
            return widget.cget('background')
        except Exception:
            try:
                return widget.cget('bg')
            except Exception:
                try:
                    return widget.winfo_toplevel().cget('background')
                except Exception:
                    return fallback
    # If topbar is missing for any reason, bail out gracefully
    if not hasattr(app, 'topbar'):
        root.mainloop()
        return
    # Capture original commands
    cmd_all = getattr(app, 'show_all_tasks', None)
    cmd_profiles = getattr(app, 'show_profiles_page', None)
    cmd_export = getattr(app, 'export_pdf', None)
    cmd_sync = getattr(app, 'manual_sync', None)
    if not callable(cmd_sync):
        cmd_sync = (lambda: None)

    # Create titlebar buttons (All Tasks, Profiles, Export PDF)
    title_btn_font = tkfont.Font(size=8)
    
    if cmd_all:
        btn = tk.Button(title_middle, text='All Tasks', command=cmd_all, bg='#505050', fg='#fff',
                       bd=0, relief='flat', highlightthickness=0, padx=8, pady=1, font=title_btn_font)
        btn.pack(side='left', padx=3)
        title_btn_refs['all_tasks'] = btn
    
    if cmd_profiles:
        btn = tk.Button(title_middle, text='Profiles', command=cmd_profiles, bg="#505050", fg='#fff',
                       bd=0, relief='flat', highlightthickness=0, padx=8, pady=1, font=title_btn_font)
        btn.pack(side='left', padx=3)
        title_btn_refs['profiles'] = btn
    
    if cmd_export:
        btn = tk.Button(title_middle, text='Export', command=cmd_export, bg="#505050", fg='#fff',
                       bd=0, relief='flat', highlightthickness=0, padx=8, pady=1, font=title_btn_font)
        btn.pack(side='left', padx=3)
        title_btn_refs['export'] = btn

    for w in list(topbar.winfo_children()):
        if isinstance(w, ttk.Button):
            try:
                w.destroy()
            except Exception:
                pass

        def rounded_button(parent, text, command=None, fg="#C4C4C4", bg="#3d3d3d", hover="#7a7a7a"):
            container = tk.Frame(parent, background=_bg(parent))
            radius = 12
            padx, pady = 12, 6
            font_obj = tkfont.nametofont("TkDefaultFont")
            tw = font_obj.measure(text) + 2 * (padx + radius)
            th = font_obj.metrics('linespace') + 2 * (pady + int(radius/2))
            if PIL_AVAILABLE:
                scale = 4
                img = Image.new('RGBA', (tw*scale, th*scale), (0,0,0,0))
                draw = ImageDraw.Draw(img)
                rect = [0, 0, tw*scale, th*scale]
                draw.rounded_rectangle(rect, radius=radius*scale, fill=bg)
                try:
                    fnt = ImageFont.truetype("SegoeUI.ttf", int(font_obj.metrics('linespace')*scale))
                except Exception:
                    fnt = ImageFont.load_default()
                text_w, text_h = draw.textsize(text, font=fnt)
                draw.text(((tw*scale-text_w)/2, (th*scale-text_h)/2), text, fill=fg, font=fnt)
                img = img.resize((tw, th), Image.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                lbl = tk.Label(container, image=photo, background=_bg(parent))
                lbl.image = photo
                lbl.pack()
                def set_bg(color):
                    im2 = Image.new('RGBA', (tw*scale, th*scale), (0,0,0,0))
                    d2 = ImageDraw.Draw(im2)
                    d2.rounded_rectangle([0,0,tw*scale,th*scale], radius=radius*scale, fill=color)
                    d2.text(((tw*scale-text_w)/2, (th*scale-text_h)/2), text, fill=fg, font=fnt)
                    im2 = im2.resize((tw, th), Image.LANCZOS)
                    lbl.image = ImageTk.PhotoImage(im2)
                    lbl.configure(image=lbl.image)
                def on_enter(e): set_bg(hover)
                def on_leave(e): set_bg(bg)
                def on_click(e):
                    if callable(command): command()
                lbl.bind('<Enter>', on_enter)
                lbl.bind('<Leave>', on_leave)
                lbl.bind('<Button-1>', on_click)
                return container
            else:

                c = tk.Canvas(container, width=tw, height=th, highlightthickness=0, bd=0, background=_bg(parent))
                c.pack()
                def draw(color):
                    c.delete('all')
                    r = radius
                    c.create_arc(0, 0, 2*r, 2*r, start=90, extent=90, fill=color, outline=color)
                    c.create_arc(tw-2*r, 0, tw, 2*r, start=0, extent=90, fill=color, outline=color)
                    c.create_arc(0, th-2*r, 2*r, th, start=180, extent=90, fill=color, outline=color)
                    c.create_arc(tw-2*r, th-2*r, tw, th, start=270, extent=90, fill=color, outline=color)
                    c.create_rectangle(r, 0, tw-r, th, fill=color, outline=color)
                    c.create_rectangle(0, r, tw, th-r, fill=color, outline=color)
                    c.create_text(tw/2, th/2, text=text, fill=fg, font=font_obj)
                draw(bg)
                c.bind('<Enter>', lambda e: draw(hover))
                c.bind('<Leave>', lambda e: draw(bg))
                c.bind('<Button-1>', lambda e: command() if callable(command) else None)
                return container

    def square_icon_button(parent, command=None, size=28, bg="#22c55e", fg="#ffffff", symbol='↻'):
        container = tk.Frame(parent, background=_bg(parent))
        if PIL_AVAILABLE:
            scale = 4
            base = Image.new('RGBA', (size*scale, size*scale), (0,0,0,0))
            draw = ImageDraw.Draw(base)
            # Draw rounded rectangle instead of circle
            draw.rounded_rectangle([0,0,size*scale,size*scale], radius=6*scale, fill=bg)
            try:
                fnt = ImageFont.truetype("SegoeUI-Semibold.ttf", int(size*scale*0.55))
            except Exception:
                fnt = ImageFont.load_default()
            # robust text measure
            try:
                bbox = draw.textbbox((0,0), symbol, font=fnt)
                tw, th = bbox[2]-bbox[0], bbox[3]-bbox[1]
            except Exception:
                try:
                    tw, th = fnt.getbbox(symbol)[2:]
                except Exception:
                    tw, th = draw.textlength(symbol, font=fnt), int(size*scale*0.5)
            draw.text(((size*scale-tw)/2, (size*scale-th)/2), symbol, fill=fg, font=fnt)
            base = base.resize((size, size), Image.LANCZOS)
            photo = ImageTk.PhotoImage(base)
            lbl = tk.Label(container, image=photo, background=_bg(parent))
            lbl.image = photo
            lbl._base_img = base
            lbl.pack()
            def on_enter(e):
                im2 = Image.new('RGBA', (size, size), (0,0,0,0))
                d2 = ImageDraw.Draw(im2)
                d2.rounded_rectangle([0,0,size,size], radius=6, fill="#16a34a")
                # paste arrow from base preserving alpha
                im2.alpha_composite(lbl._base_img)
                lbl.image = ImageTk.PhotoImage(im2)
                lbl.configure(image=lbl.image)
            def on_leave(e):
                lbl.image = ImageTk.PhotoImage(lbl._base_img)
                lbl.configure(image=lbl.image)
            def on_click(e):
                if callable(command): command()
            lbl.bind('<Enter>', on_enter)
            lbl.bind('<Leave>', on_leave)
            lbl.bind('<Button-1>', on_click)
            # helpers
            def update_symbol(new_symbol=None, new_bg=None):
                # rebuild base image with optional bg/symbol
                bg_color = new_bg or bg
                canvas = Image.new('RGBA', (size*scale, size*scale), (0,0,0,0))
                d = ImageDraw.Draw(canvas)
                d.rounded_rectangle([0,0,size*scale,size*scale], radius=6*scale, fill=bg_color)
                sym = new_symbol or symbol
                try:
                    bbox2 = d.textbbox((0,0), sym, font=fnt)
                    tw2, th2 = bbox2[2]-bbox2[0], bbox2[3]-bbox2[1]
                except Exception:
                    try:
                        tw2, th2 = fnt.getbbox(sym)[2:]
                    except Exception:
                        tw2, th2 = d.textlength(sym, font=fnt), int(size*scale*0.5)
                d.text(((size*scale-tw2)/2, (size*scale-th2)/2), sym, fill=fg, font=fnt)
                img = canvas.resize((size, size), Image.LANCZOS)
                lbl._base_img = img
                lbl.image = ImageTk.PhotoImage(img)
                lbl.configure(image=lbl.image)
                lbl.update_idletasks()
            container.update_symbol = update_symbol
            return container
        else:
            c = tk.Canvas(container, width=size, height=size, highlightthickness=0, bd=0, background=_bg(parent))
            c.pack()
            x0, y0, x1, y1 = 0, 0, size, size
            rect = c.create_rectangle(x0, y0, x1, y1, fill=bg, outline=bg)
            txt = c.create_text(size/2, size/2, text=symbol, fill=fg, font=("Segoe UI", int(size/2)))
            c.bind('<Enter>', lambda e: c.itemconfig(rect, fill="#16a34a", outline="#16a34a"))
            c.bind('<Leave>', lambda e: c.itemconfig(rect, fill=bg, outline=bg))
            c.bind('<Button-1>', lambda e: command() if callable(command) else None)
            def update_symbol(new_symbol=None, new_bg=None):
                if new_symbol is not None:
                    c.itemconfig(txt, text=new_symbol)
                if new_bg is not None:
                    c.itemconfig(rect, fill=new_bg, outline=new_bg)
                c.update_idletasks()
            container.update_symbol = update_symbol
            return container

    # Add a hamburger (three lines) menu that pops a dropdown with previous actions
    def hamburger_button(parent, command=None, width=32, height=28, bg="#3F3F3F", fg="#ffffff"):
        container = tk.Frame(parent, background=_bg(parent))
        c = tk.Canvas(container, width=width, height=height, highlightthickness=0, bd=0, background=_bg(parent))
        c.pack()
        # rounded rectangle background (approximate with regular rectangle)
        c.create_rectangle(0, 0, width, height, fill=bg, outline=bg)
        # three lines
        pad = 6
        line_w = width - 2*pad
        cx = pad
        cy1 = height/2 - 6
        cy2 = height/2
        cy3 = height/2 + 6
        for cy in (cy1, cy2, cy3):
            c.create_rectangle(cx, cy-1.5, cx+line_w, cy+1.5, fill=fg, outline=fg)
        def on_enter(e):
            c.delete("all")
            c.create_rectangle(0, 0, width, height, fill="#636363", outline="#585858")
            for cy in (cy1, cy2, cy3):
                c.create_rectangle(cx, cy-1.5, cx+line_w, cy+1.5, fill=fg, outline=fg)
        def on_leave(e):
            c.delete("all")
            c.create_rectangle(0, 0, width, height, fill=bg, outline=bg)
            for cy in (cy1, cy2, cy3):
                c.create_rectangle(cx, cy-1.5, cx+line_w, cy+1.5, fill=fg, outline=fg)
        def on_click(e):
            if callable(command):
                command()
        c.bind('<Enter>', on_enter)
        c.bind('<Leave>', on_leave)
        c.bind('<Button-1>', on_click)
        return container

    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception:
        # Capture full traceback to a file so the user can share it for debugging
        import traceback, sys
        tb = traceback.format_exc()
        try:
            with open('startup_error.log', 'w', encoding='utf-8') as f:
                f.write(tb)
        except Exception:
            pass
        # Try to show a simple message box to the user
        try:
            import tkinter as _tk
            from tkinter import messagebox as _mb
            _root = _tk.Tk()
            _root.withdraw()
            _mb.showerror('Startup Error', "An error occurred while starting the app. A file 'startup_error.log' was written to the current folder.")
            _root.destroy()
        except Exception:
            pass
        sys.exit(1)

