import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageTk
import os
import math
IMAGE_FOLDER = "graphics/ships/"


class EditTorpedoDialog:
    def __init__(self, parent, current_list, available_types):
        self.top = tk.Toplevel(parent)
        self.top.title("Edit Torpedo Start")
        self.result = None
        self.available_types = available_types
        
        # Store rows, each row is a dict with variables.
        self.rows = []
        self.frm_rows = ttk.Frame(self.top)
        self.frm_rows.grid(row=0, column=0, columnspan=3, padx=10, pady=10)
        
        tk.Label(self.top, text="Torpedo Type").grid(row=1, column=0, padx=5, pady=5)
        tk.Label(self.top, text="Starting Count").grid(row=1, column=1, padx=5, pady=5)
        
        if current_list:
            for item in current_list:
                torpedo_type = list(item.keys())[0]
                count = item[torpedo_type]
                self.add_row(torpedo_type, str(count))
        else:
            self.add_row("", "0")
            
        # After binding key release events and calling self.update_beam_preview(), add:
        btn_add = ttk.Button(self.top, text="Add Row", command=lambda: self.add_row())
        btn_add.grid(row=2, column=0, padx=5, pady=5, sticky="w")

        # Create a dedicated frame for the OK and Cancel buttons.
        button_frame = ttk.Frame(self.top)
        button_frame.grid(row=3, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

        # Configure the button_frame columns to distribute space equally.
        button_frame.columnconfigure(0, weight=1)
        button_frame.columnconfigure(1, weight=1)

        btn_ok = ttk.Button(button_frame, text="OK", command=self.on_ok)
        btn_ok.grid(row=0, column=0, padx=5, pady=5, sticky="e")

        btn_cancel = ttk.Button(button_frame, text="Cancel", command=self.on_cancel)
        btn_cancel.grid(row=0, column=1, padx=5, pady=5, sticky="w")

        
        self.top.grab_set()
        self.top.wait_window(self.top)
        
    def add_row(self, torpedo_type: str = "", count: str = "0"):
        row_index = len(self.rows)
        type_var = tk.StringVar(value=torpedo_type)
        count_var = tk.StringVar(value=count)
        combobox_type = ttk.Combobox(self.frm_rows, textvariable=type_var, values=self.available_types, width=12)
        combobox_type.grid(row=row_index, column=0, padx=5, pady=2)
        combobox_type.config(state="normal")  # editable
        entry_count = ttk.Entry(self.frm_rows, textvariable=count_var, width=8)
        entry_count.grid(row=row_index, column=1, padx=5, pady=2)
        btn_remove = ttk.Button(self.frm_rows, text="Remove", command=lambda idx=row_index: self.remove_row(idx))
        btn_remove.grid(row=row_index, column=2, padx=5, pady=2)
        self.rows.append({
            "type_var": type_var,
            "count_var": count_var,
            "combobox_type": combobox_type,
            "entry_count": entry_count,
            "btn_remove": btn_remove
        })
        
    def remove_row(self, idx):
        row = self.rows[idx]
        if row is None:
            return
        row["combobox_type"].destroy()
        row["entry_count"].destroy()
        row["btn_remove"].destroy()
        self.rows[idx] = None
        
    def on_ok(self):
        result = []
        types_seen = set()
        for row in self.rows:
            if row is None:
                continue
            ttype = row["type_var"].get().strip()
            count_str = row["count_var"].get().strip()
            if ttype:
                if ttype in types_seen:
                    messagebox.showerror("Error", f"Duplicate torpedo type '{ttype}' found. Each type must be unique.")
                    return
                types_seen.add(ttype)
                try:
                    count = int(count_str)
                except ValueError:
                    messagebox.showerror("Error", f"Invalid count for torpedo type '{ttype}'. Must be an integer.")
                    return
                result.append({ttype: count})
        self.result = result
        self.top.destroy()
        
    def on_cancel(self):
        self.top.destroy()
        
import tkinter as tk
from tkinter import ttk, messagebox
from tksheet import Sheet  # Make sure to install tksheet via pip install tksheet


# --- Edit Beam Ports Dialog Class ---
class EditBeamPortsDialog:
    def __init__(self, parent, current_list, ship_artfileroot=None, open_3d_callback=None):
        # Save the ship's artfileroot, if provided.
        self.ship_artfileroot = ship_artfileroot
        
        self.top = tk.Toplevel(parent)
        self.top.title("Edit Beam Ports")
        self.result = None
        self.rows = []

        # --- Create preview frame and canvas BEFORE adding rows ---
        preview_frame = ttk.Frame(self.top)
        preview_frame.grid(row=0, column=1, rowspan=4, padx=10, pady=10, sticky="nsew")
        self.preview_canvas = tk.Canvas(preview_frame, width=400, height=400, background="black")
        self.preview_canvas.pack(expand=True, fill="both")
        # Redraw when the canvas is resized so the image stays centered/scaled
        self.preview_canvas.bind("<Configure>", lambda e: self.update_beam_preview())
        self.load_preview_base_image()  # Load the ship image (or fallback image)
        # Ensure the two main columns (left form / right preview) resize well
        self.top.grid_columnconfigure(0, weight=1)
        self.top.grid_columnconfigure(1, weight=1)
        self.top.grid_rowconfigure(1, weight=1)
        
        # Set up header for beam port rows.
        self.headers = ["X", "Y", "Z", "Color", "ArcColor", "Cycle Time", 
                        "Damage Coeff", "Range", "ArcWidth", "Barrel Angle", " ", " "]
        self.header_frame = ttk.Frame(self.top)
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=2, pady=2)
        for col, text in enumerate(self.headers):
            lbl = ttk.Label(self.header_frame, text=text, font=("Arial", 10, "bold"))
            lbl.grid(row=0, column=col, padx=2, pady=2, sticky="ew")
            self.header_frame.grid_columnconfigure(col, weight=1, uniform="col")
        
        self.frm_rows = ttk.Frame(self.top)
        self.frm_rows.grid(row=1, column=0, sticky="ew", padx=2, pady=2)
        
        # Populate rows from the current_list (or add a default row if none provided).
        if current_list:
            for entry in current_list:
                pos = entry.get("position", [0, 0, 0])
                row_data = {
                    "x": tk.StringVar(value=str(pos[0])),
                    "y": tk.StringVar(value=str(pos[1])),
                    "z": tk.StringVar(value=str(pos[2])),
                    "color": tk.StringVar(value=entry.get("color", "")),
                    "arccolor": tk.StringVar(value=entry.get("arccolor", "red")),
                    "cycle_time": tk.StringVar(value=str(entry.get("cycle_time", 0))),
                    "damage_coeff": tk.StringVar(value=str(entry.get("damage_coeff", 0))),
                    "range": tk.StringVar(value=str(entry.get("range", 0))),
                    "arcwidth": tk.StringVar(value=str(entry.get("arcwidth", 0))),
                    "barrel_angle": tk.StringVar(value=str(entry.get("barrel_angle", 0)))
                }
                self.add_row(row_data)
        else:
            self.add_row()
        
        # Add buttons for row control.
        # Controls row (left column): Add Row + Open 3D Viewer
        controls_frame = ttk.Frame(self.top)
        controls_frame.grid(row=2, column=0, padx=5, pady=5, sticky="ew")
        controls_frame.columnconfigure(0, weight=1)
        controls_frame.columnconfigure(1, weight=1)
        btn_add = ttk.Button(controls_frame, text="Add Row", command=lambda: self.add_row())
        btn_add.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        if open_3d_callback is not None:
            btn_3d = ttk.Button(controls_frame, text="Open 3D Viewer",
                                command=lambda cb=open_3d_callback: cb(parent=self.top))
            btn_3d.grid(row=0, column=1, padx=5, pady=5, sticky="e")
        btn_ok = ttk.Button(self.top, text="OK", command=self.on_ok)
        btn_ok.grid(row=3, column=0, padx=5, pady=5, sticky="e")
        btn_cancel = ttk.Button(self.top, text="Cancel", command=self.on_cancel)
        btn_cancel.grid(row=3, column=0, padx=5, pady=5, sticky="w")
        
        # Bind key release events on arc-related fields to update the preview.
        for row in self.rows:
            for field in ["range", "arcwidth", "barrel_angle", "arccolor"]:
                entry_widget = row["widgets"].get(field)
                if entry_widget:
                    entry_widget.bind("<KeyRelease>", lambda e: self.update_beam_preview())
        
        self.update_beam_preview()
        self.top.grab_set()
        self.top.wait_window(self.top)
    
    def load_preview_base_image(self):
        try:
            # Resolve under <cosmos_root>/data if the editor set it for us
            data_dir = os.environ.get("COSMOS_DATA_DIR", os.getcwd())
            base_dir = os.path.join(data_dir, IMAGE_FOLDER)
            if self.ship_artfileroot:
                norm_art = os.path.normpath(str(self.ship_artfileroot).lstrip("\\/"))
                image_path_1024 = os.path.join(base_dir, norm_art + "1024.png")
                image_path_256 = os.path.join(base_dir, norm_art + "256.png")
                image_path = image_path_1024 if os.path.exists(image_path_1024) else image_path_256
                print("Attempting to load image from:", os.path.abspath(image_path))
            else:
                image_path = os.path.join(base_dir, "unknown1024.png")
            
            if os.path.exists(image_path):
                # Keep the original PIL image; we’ll scale it per current canvas size
                base_img = Image.open(image_path).convert("RGBA").rotate(180)
                self.preview_base_img = base_img
                self.preview_base = None  # will be created in update_beam_preview()
            else:
                self.preview_base_img = None
                self.preview_base = None
        except Exception as e:
            print("Error loading preview base image:", e)
            self.preview_base = None

    def update_beam_preview(self):
        self.preview_canvas.delete("all")
        # Actual current size (winfo_*) with sane minimums to avoid 0 sizes
        canvas_w = max(2, int(self.preview_canvas.winfo_width() or 0))
        canvas_h = max(2, int(self.preview_canvas.winfo_height() or 0))
        # Center point for both image and arcs
        cx, cy = canvas_w // 2, canvas_h // 2
        # (Re)build a scaled PhotoImage so the ship texture is centered and sized to the panel
        if getattr(self, "preview_base_img", None) is not None:
            # Fit image into ~90% of the canvas area while keeping aspect
            desired_w = max(1, int(canvas_w * 0.9))
            desired_h = max(1, int(canvas_h * 0.9))
            img = self.preview_base_img.copy()
            img.thumbnail((desired_w, desired_h))
            self.preview_base = ImageTk.PhotoImage(img)
            # Draw centered
            self.preview_canvas.create_image(cx, cy, anchor="center", image=self.preview_base)
        scale = 0.1  # Adjust scale as needed.

        # Draw an overlay for each beam row.
        for row in self.rows:
            if row is None:
                continue
            data = row["data"]
            try:
                barrel_angle = float(data["barrel_angle"].get() or 0)
                arc_width = float(data["arcwidth"].get() or 0)
                port_range = float(data["range"].get() or 0)
            except Exception as e:
                print("Conversion error:", e)
                continue

            arc_color = data.get("arccolor").get() or "red"
            pixel_radius = port_range * scale
            # Compute center angle (Tkinter's coordinate system starts at 3 o’clock).
            center_angle = 90 - barrel_angle
            # Define bounding box for the arc or circle.
            x0, y0 = cx - pixel_radius, cy - pixel_radius
            x1, y1 = cx + pixel_radius, cy + pixel_radius
            
            if arc_width >= 360:
                # Draw a full circle.
                self.preview_canvas.create_oval(x0, y0, x1, y1, outline=arc_color, width=2)
                # Optional: draw a radial line to indicate the barrel angle.
                angle_rad = math.radians(center_angle)
                x_line = cx + pixel_radius * math.cos(angle_rad)
                y_line = cy - pixel_radius * math.sin(angle_rad)
                self.preview_canvas.create_line(cx, cy, x_line, y_line, fill=arc_color, width=2)
            else:
                # Draw arc.
                start_angle = center_angle - (arc_width / 2)
                self.preview_canvas.create_arc(x0, y0, x1, y1,
                                               start=start_angle, extent=arc_width,
                                               style="arc", outline=arc_color, width=2)
                # Draw the two radial lines at the start and end of the arc.
                start_rad = math.radians(start_angle)
                end_rad = math.radians(start_angle + arc_width)
                x_start = cx + pixel_radius * math.cos(start_rad)
                y_start = cy - pixel_radius * math.sin(start_rad)
                x_end = cx + pixel_radius * math.cos(end_rad)
                y_end = cy - pixel_radius * math.sin(end_rad)
                self.preview_canvas.create_line(cx, cy, x_start, y_start, fill=arc_color, width=2)
                self.preview_canvas.create_line(cx, cy, x_end, y_end, fill=arc_color, width=2)

    
    def add_row(self, row_data=None):
        row_index = len(self.rows)
        if row_data is None:
            row_data = {
                "x": tk.StringVar(value="0"),
                "y": tk.StringVar(value="0"),
                "z": tk.StringVar(value="0"),
                "color": tk.StringVar(value=""),
                "arccolor": tk.StringVar(value="red"),
                "cycle_time": tk.StringVar(value="0"),
                "damage_coeff": tk.StringVar(value="0"),
                "range": tk.StringVar(value="0"),
                "arcwidth": tk.StringVar(value="0"),
                "barrel_angle": tk.StringVar(value="0")
            }
        entries = {}
        col_fields = ["x", "y", "z", "color", "arccolor", "cycle_time", "damage_coeff", "range", "arcwidth", "barrel_angle"]
        for col, field in enumerate(col_fields):
            ent = ttk.Entry(self.frm_rows, textvariable=row_data[field], width=8)
            ent.grid(row=row_index, column=col, padx=2, pady=2, sticky="ew")
            self.frm_rows.grid_columnconfigure(col, weight=1, uniform="col")
            entries[field] = ent
            if field in ["range", "arcwidth", "barrel_angle", "arccolor"]:
                ent.bind("<KeyRelease>", lambda e: self.update_beam_preview())
        
        btn_paste = ttk.Button(self.frm_rows, text="Paste Pos", command=lambda idx=row_index: self.paste_position(idx))
        btn_paste.grid(row=row_index, column=len(col_fields), padx=2, pady=2, sticky="ew")
        self.frm_rows.grid_columnconfigure(len(col_fields), weight=1, uniform="col")
        
        btn_remove = ttk.Button(self.frm_rows, text="Remove", command=lambda idx=row_index: self.remove_row(idx))
        btn_remove.grid(row=row_index, column=len(col_fields)+1, padx=2, pady=2, sticky="ew")
        self.frm_rows.grid_columnconfigure(len(col_fields)+1, weight=1, uniform="col")
        
        self.rows.append({
            "data": row_data,
            "widgets": entries,
            "btn_paste": btn_paste,
            "btn_remove": btn_remove
        })
        self.update_beam_preview()

    def paste_position(self, idx):
        """Accept legacy format "[x,y,z]," or plain raw numbers "x, y, z" / "x y z" / "x,y,z"."""
        try:
            clip = self.top.clipboard_get().strip()
            # Drop trailing comma if present (legacy)
            if clip.endswith(','):
                clip = clip[:-1].strip()
            # If bracketed, strip brackets
            if clip.startswith('[') and clip.endswith(']'):
                clip = clip[1:-1].strip()
            # Split by comma or whitespace
            parts = [p for p in clip.replace(',', ' ').split() if p]
            if len(parts) != 3:
                messagebox.showerror("Error", "Clipboard does not contain three numbers (x, y, z).")
                return
            x, y, z = parts[0], parts[1], parts[2]
            self.rows[idx]["data"]["x"].set(x)
            self.rows[idx]["data"]["y"].set(y)
            self.rows[idx]["data"]["z"].set(z)
            self.update_beam_preview()
        except Exception as e:
            messagebox.showerror("Error", f"Error pasting position: {e}")

    def remove_row(self, idx):
        row = self.rows[idx]
        if row is None:
            return
        for widget in row["widgets"].values():
            widget.destroy()
        row["btn_paste"].destroy()
        row["btn_remove"].destroy()
        self.rows[idx] = None
        self.update_beam_preview()

    def on_ok(self):
        result = []
        for row in self.rows:
            if row is None:
                continue
            data = row["data"]
            try:
                x = float(data["x"].get().strip())
                y = float(data["y"].get().strip())
                z = float(data["z"].get().strip())
                cycle_time = int(data["cycle_time"].get().strip())
                damage_coeff = float(data["damage_coeff"].get().strip())
                range_val = int(data["range"].get().strip())
                arcwidth = int(data["arcwidth"].get().strip())
                barrel_angle = int(data["barrel_angle"].get().strip())
            except ValueError:
                messagebox.showerror("Error", "Numeric values in beam ports must be valid numbers.")
                return
            entry = {
                "position": [x, y, z],
                "color": data["color"].get().strip(),
                "arccolor": data["arccolor"].get().strip(),
                "cycle_time": cycle_time,
                "damage_coeff": damage_coeff,
                "range": range_val,
                "arcwidth": arcwidth,
                "barrel_angle": barrel_angle
            }
            result.append(entry)
        self.result = result
        self.top.destroy()
        
    def on_cancel(self):
        self.top.destroy()


# --- Edit Exhaust Ports Dialog Class ---
class EditExhaustPortsDialog:
    def __init__(self, parent, current_list, open_3d_callback=None):
        self.top = tk.Toplevel(parent)
        self.top.title("Edit Exhaust Ports")
        self.result = None
        self.rows = []
        
        # Define headers for the 4 data columns plus 2 action columns.
        self.headers = ["X", "Y", "Z", "Color", "Paste", "Remove"]
        
        # Create a header frame
        self.header_frame = ttk.Frame(self.top)
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=2, pady=2)
        for col, text in enumerate(self.headers):
            lbl = ttk.Label(self.header_frame, text=text, font=("Arial", 10, "bold"))
            lbl.grid(row=0, column=col, padx=2, pady=2, sticky="ew")
            # Ensure all columns have uniform width
            self.header_frame.grid_columnconfigure(col, weight=1, uniform="col")
        
        # Frame for data rows
        self.frm_rows = ttk.Frame(self.top)
        self.frm_rows.grid(row=1, column=0, sticky="ew", padx=2, pady=2)
        
        # Populate with existing data, if any; otherwise add one blank row.
        if current_list:
            for entry in current_list:
                pos = entry.get("position", [0, 0, 0])
                row_data = {
                    "x": tk.StringVar(value=str(pos[0])),
                    "y": tk.StringVar(value=str(pos[1])),
                    "z": tk.StringVar(value=str(pos[2])),
                    "color": tk.StringVar(value=entry.get("color", ""))
                }
                self.add_row(row_data)
        else:
            self.add_row()

        btn_add = ttk.Button(self.top, text="Add Row", command=lambda: self.add_row())
        btn_add.grid(row=2, column=0, sticky="w", padx=5, pady=5)
        # Optional: Open 3D Viewer button on the same row
        if open_3d_callback is not None:
            btn_3d = ttk.Button(self.top, text="Open 3D Viewer",
                                command = lambda cb=open_3d_callback: cb(parent=self.top))
            btn_3d.grid(row=2, column=1, sticky="e", padx=5, pady=5)
        btn_ok = ttk.Button(self.top, text="OK", command=self.on_ok)
        btn_ok.grid(row=3, column=0, sticky="e", padx=5, pady=5)
        btn_cancel = ttk.Button(self.top, text="Cancel", command=self.on_cancel)
        btn_cancel.grid(row=3, column=0, sticky="w", padx=5, pady=5)
        
        self.top.grab_set()
        self.top.wait_window(self.top)
    
    def add_row(self, row_data=None):
        row_index = len(self.rows)
        if row_data is None:
            row_data = {
                "x": tk.StringVar(value="0"),
                "y": tk.StringVar(value="0"),
                "z": tk.StringVar(value="0"),
                "color": tk.StringVar(value="")
            }
        entries = {}
        # Define the 4 data columns for Exhaust.
        col_fields = ["x", "y", "z", "color"]
        for col, field in enumerate(col_fields):
            ent = ttk.Entry(self.frm_rows, textvariable=row_data[field], width=8)
            ent.grid(row=row_index, column=col, padx=2, pady=2, sticky="ew")
            self.frm_rows.grid_columnconfigure(col, weight=1, uniform="col")
            entries[field] = ent
        
        # Column for Paste button.
        btn_paste = ttk.Button(self.frm_rows, text="Paste Pos", 
                                command=lambda idx=row_index: self.paste_position(idx))
        btn_paste.grid(row=row_index, column=len(col_fields), padx=2, pady=2, sticky="ew")
        self.frm_rows.grid_columnconfigure(len(col_fields), weight=1, uniform="col")
        
        # Column for Remove button.
        btn_remove = ttk.Button(self.frm_rows, text="Remove", 
                                 command=lambda idx=row_index: self.remove_row(idx))
        btn_remove.grid(row=row_index, column=len(col_fields)+1, padx=2, pady=2, sticky="ew")
        self.frm_rows.grid_columnconfigure(len(col_fields)+1, weight=1, uniform="col")
        
        self.rows.append({
            "data": row_data,
            "widgets": entries,
            "btn_paste": btn_paste,
            "btn_remove": btn_remove
        })
    
    def paste_position(self, idx):
        try:
            clip = self.top.clipboard_get().strip()
            if clip.endswith(','):
                clip = clip[:-1].strip()
            if clip.startswith("[") and clip.endswith("]"):
                content = clip[1:-1].strip()
                if content.endswith(','):
                    content = content[:-1].strip()
                parts = [p for p in content.split(",") if p.strip() != ""]
                if len(parts) != 3:
                    messagebox.showerror("Error", "Clipboard does not contain three numbers.")
                    return
                x, y, z = parts[0].strip(), parts[1].strip(), parts[2].strip()
                self.rows[idx]["data"]["x"].set(x)
                self.rows[idx]["data"]["y"].set(y)
                self.rows[idx]["data"]["z"].set(z)
            else:
                messagebox.showerror("Error", "Clipboard content not in expected format (e.g. [ x, y, z ]).")
        except Exception as e:
            messagebox.showerror("Error", f"Error pasting position: {e}")
    
    def remove_row(self, idx):
        row = self.rows[idx]
        if row is None:
            return
        for widget in row["widgets"].values():
            widget.destroy()
        row["btn_paste"].destroy()
        row["btn_remove"].destroy()
        self.rows[idx] = None
        
    def on_ok(self):
        result = []
        for row in self.rows:
            if row is None:
                continue
            data = row["data"]
            try:
                x = float(data["x"].get().strip())
                y = float(data["y"].get().strip())
                z = float(data["z"].get().strip())
            except ValueError:
                messagebox.showerror("Error", "Position values in exhaust ports must be valid numbers.")
                return
            entry = {
                "position": [x, y, z],
                "color": data["color"].get().strip()
            }
            result.append(entry)
        self.result = result
        self.top.destroy()
        
    def on_cancel(self):
        self.top.destroy()


# --- New Ship Dialog Class ---
class NewShipDialog:
    def __init__(self, parent, existing_sides):
        self.top = tk.Toplevel(parent)
        self.top.title("New Ship")
        self.result = None

        self.var_side = tk.StringVar(value=existing_sides[0] if existing_sides else "Unknown")
        self.var_name = tk.StringVar(value="New Ship")
        self.var_key = tk.StringVar()
        self.var_artfileroot = tk.StringVar(value="unknown")
        self.var_meshscale = tk.StringVar(value="1")
        self.var_radarscale = tk.StringVar(value="1")
        self.var_exclusionradius = tk.StringVar(value="0")

        def update_key(*args):
            side = self.var_side.get().strip()
            name = self.var_name.get().strip()
            if side and name:
                generated = f"{side}_{name}".lower().replace(" ", "_")
                self.var_key.set(generated)
        self.var_side.trace("w", update_key)
        self.var_name.trace("w", update_key)

        tk.Label(self.top, text="Side:").grid(row=0, column=0, sticky="e", padx=5, pady=5)
        self.combo_side = ttk.Combobox(self.top, textvariable=self.var_side)
        self.combo_side['values'] = existing_sides
        self.combo_side.grid(row=0, column=1, padx=5, pady=5)
        self.combo_side.config(state="normal")

        tk.Label(self.top, text="Name:").grid(row=1, column=0, sticky="e", padx=5, pady=5)
        self.entry_name = tk.Entry(self.top, textvariable=self.var_name)
        self.entry_name.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(self.top, text="Key:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
        self.entry_key = tk.Entry(self.top, textvariable=self.var_key)
        self.entry_key.grid(row=2, column=1, padx=5, pady=5)

        tk.Label(self.top, text="ArtFileRoot:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
        self.entry_artfileroot = tk.Entry(self.top, textvariable=self.var_artfileroot)
        self.entry_artfileroot.grid(row=3, column=1, padx=5, pady=5)

        tk.Label(self.top, text="MeshScale:").grid(row=4, column=0, sticky="e", padx=5, pady=5)
        self.entry_meshscale = tk.Entry(self.top, textvariable=self.var_meshscale)
        self.entry_meshscale.grid(row=4, column=1, padx=5, pady=5)

        tk.Label(self.top, text="RadarScale:").grid(row=5, column=0, sticky="e", padx=5, pady=5)
        self.entry_radarscale = tk.Entry(self.top, textvariable=self.var_radarscale)
        self.entry_radarscale.grid(row=5, column=1, padx=5, pady=5)

        tk.Label(self.top, text="ExclusionRadius:").grid(row=6, column=0, sticky="e", padx=5, pady=5)
        self.entry_exclusionradius = tk.Entry(self.top, textvariable=self.var_exclusionradius)
        self.entry_exclusionradius.grid(row=6, column=1, padx=5, pady=5)

        btn_ok = ttk.Button(self.top, text="OK", command=self.on_ok)
        btn_ok.grid(row=7, column=0, padx=5, pady=5)
        btn_cancel = ttk.Button(self.top, text="Cancel", command=self.on_cancel)
        btn_cancel.grid(row=7, column=1, padx=5, pady=5)

        self.top.grab_set()
        self.top.wait_window(self.top)

    def on_ok(self):
        key = self.var_key.get().strip()
        name = self.var_name.get().strip()
        side = self.var_side.get().strip()
        artfileroot = self.var_artfileroot.get().strip()
        if not key or not name or not side or not artfileroot:
            messagebox.showerror("Error", "Side, Name, Key, and ArtFileRoot are required.")
            return
        try:
            meshscale = float(self.var_meshscale.get().strip())
            radarscale = float(self.var_radarscale.get().strip())
            exclusionradius = float(self.var_exclusionradius.get().strip())
        except ValueError:
            messagebox.showerror("Error", "MeshScale, RadarScale, and ExclusionRadius must be numbers.")
            return
        self.result = {
            "key": key,
            "name": name,
            "side": side,
            "artfileroot": artfileroot,
            "meshscale": meshscale,
            "radarscale": radarscale,
            "exclusionradius": exclusionradius,
            "hullpoints": 0,
            "long_desc": "",
            "tubecount": 0,
            "baycount": 0,
            "internalmapscale": 1.0,
            "internalmapw": 0,
            "internalmaph": 0,
            "internalsymmetry": 1,
            "turn_rate": 0.0,
            "speed_coeff": 1.0,
            "scan_strength_coeff": 1.0,
            "ship_energy_cost": 1.0,
            "warp_energy_cost": 1.0,
            "jump_energy_cost": 1.0,
            "roles": "",
            "drone_launch_timer": 0
        }
        self.top.destroy()

    def on_cancel(self):
        self.top.destroy()
