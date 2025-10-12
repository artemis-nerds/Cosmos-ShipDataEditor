import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
from PIL import Image, ImageTk
# dialogs and OpenGL viewer will be loaded from a sibling OrionData/ folder (see below)
import os
import hjson
import json  # used for strict JSON output
from datetime import datetime
try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None
import math
import re
import subprocess
import platform
import sys
import pyopengltk
import importlib
import importlib.util

# --- Load dialogs and OpenGL viewer from a sibling OrionData/ folder (no package needed) ---
def _orion_sibling_dir():
    """Return absolute path to OrionData/ next to this script or the PyInstaller .exe."""
    try:
        if getattr(sys, "frozen", False):
            return os.path.join(os.path.dirname(sys.executable), "OrionData")
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), "OrionData")
    except Exception:
        return os.path.join(os.getcwd(), "OrionData")

def _load_module_from(path, modname):
    spec = importlib.util.spec_from_file_location(modname, path)
    if spec and spec.loader:
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        return mod
    raise ImportError(f"Could not load module at {path}")

_ORION_DIR = _orion_sibling_dir()

# dialogs.py (required)
_DLG_PATH = os.path.join(_ORION_DIR, "dialogs.py")
_dlg_mod  = _load_module_from(_DLG_PATH, "orion_dialogs")
EditTorpedoDialog      = _dlg_mod.EditTorpedoDialog
EditBeamPortsDialog    = _dlg_mod.EditBeamPortsDialog
EditExhaustPortsDialog = _dlg_mod.EditExhaustPortsDialog
NewShipDialog          = _dlg_mod.NewShipDialog

# obj_view_gl.py (optional)
ObjTexturedGLFrame = None
OBJ_VIEW_GL_IMPORT_ERROR = None
try:
    _VIEW_PATH = os.path.join(_ORION_DIR, "obj_view_gl.py")
    _view_mod  = _load_module_from(_VIEW_PATH, "orion_obj_view_gl")
    ObjTexturedGLFrame = _view_mod.ObjTexturedGLFrame
except Exception as e:
    OBJ_VIEW_GL_IMPORT_ERROR = str(e)

    yaml = None

try:
    from ruamel.yaml import YAML
    from ruamel.yaml.scalarstring import (
        SingleQuotedScalarString as SQS,
        DoubleQuotedScalarString as DQS,
        PlainScalarString  as PSS,
    )
except Exception:
    YAML = None
    SQS = DQS = PSS = str

# Try both package and top-level imports for the 3D viewer.
# This avoids "unresolved reference" when the file lives under the shipEditor/ package.

def _app_base_dir():
    """Best-guess base dir for this app (works for script and PyInstaller)."""
    try:
        if getattr(sys, "frozen", False):
            return os.path.dirname(sys.executable)
        return os.path.dirname(os.path.abspath(__file__))
    except Exception:
        return os.getcwd()


EDITOR_VERSION = "0.9.5"
GAME_VERSION = "Artemis Cosmos v1.2.1"
HJSON_PATH = "shipData.json"
YAML_PATH = "shipData.yaml"
IMAGE_FOLDER = "graphics/ships/"
IMAGE_SUFFIX = "256.png"

    # NOTE: All reads/writes should go through _resolve_data_path so we can work inside
    # a user-chosen Cosmos *root* (folder with the game executable). Data lives in <root>/data/.

def normalize_angle(angle):
    return ((angle + 180) % 360) - 180

def beam_overlaps_sector(barrel_angle, arc_width, sector_start, sector_end):

    # Normalize the beam's center
    ba = normalize_angle(barrel_angle)
    half_width = arc_width / 2.0
    # Compute the beam arc endpoints (normalize them)
    beam_start = normalize_angle(ba - half_width)
    beam_end = normalize_angle(ba + half_width)
    
    # Define a helper that tests if a given angle is within an interval on the circle.
    def is_between(angle, start, end):
        angle = normalize_angle(angle)
        start = normalize_angle(start)
        end = normalize_angle(end)
        if start <= end:
            return start <= angle <= end
        else:
            # Interval crosses the -180/180 boundary.
            return angle >= start or angle <= end
    
    # For simplicity, we check if either endpoint or the midpoint of the beam is in the sector.
    beam_mid = normalize_angle(ba)
    return (is_between(beam_start, sector_start, sector_end) or
            is_between(beam_end, sector_start, sector_end) or
            is_between(beam_mid, sector_start, sector_end))


def calculate_damage_statistics(ship):
    total_dpm = 0
    forward_dpm = 0
    rear_dpm = 0
    beam_ports = ship.get("hull_port_sets", {}).get("beam Primary Beams", [])
    for beam in beam_ports:
        try:
            cycle_time = float(beam.get("cycle_time", 0))
            if cycle_time <= 0:
                print("Warning: cycle_time is 0 or negative in beam:", beam)
            damage_coeff = float(beam.get("damage_coeff", 0))
            shots_per_minute = math.floor(60 / cycle_time) if cycle_time > 0 else 0
            beam_dpm = shots_per_minute * damage_coeff
            total_dpm += beam_dpm

            barrel_angle = float(beam.get("barrel_angle", 0))
            arc_width = float(beam.get("arcwidth", 0))
            if beam_overlaps_sector(barrel_angle, arc_width, -55, 55):
                forward_dpm += beam_dpm
            if beam_overlaps_sector(barrel_angle, arc_width, 125, 235):
                rear_dpm += beam_dpm
        except Exception as e:
            print("Error computing damage statistics for a beam port:", e)
    return {
        "total_dpm": total_dpm,
        "forward_dpm": forward_dpm,
        "rear_dpm": rear_dpm
    }


class CreateToolTip(object):
    def __init__(self, widget, text='widget info'):
        self.waittime = 500     # miliseconds
        self.wraplength = 180   # pixels
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None
    def enter(self, event=None):
        self.schedule()
    def leave(self, event=None):
        self.unschedule()
        self.hidetip()
    def schedule(self):
        self.unschedule()
        self.id = self.widget.after(self.waittime, self.showtip)
    def unschedule(self):
        _id = self.id
        self.id = None
        if _id:
            self.widget.after_cancel(_id)
    def showtip(self, event=None):
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 27
        y = y + self.widget.winfo_rooty() + cy + 27
        self.tw = tk.Toplevel(self.widget)
        self.tw.wm_overrideredirect(True)
        self.tw.wm_geometry("+%d+%d" % (x, y))
        label = tk.Label(self.tw, text=self.text, justify='left',
                         background="#ffffe0", relief='solid', borderwidth=1,
                         wraplength = self.wraplength)
        label.pack(ipadx=1)
    def hidetip(self):
        tw = self.tw
        self.tw = None
        if tw:
            tw.destroy()

# Dictionary mapping field keys to help tips.
help_texts = {
    "name": "Name - Name",
    "key": "Key - Unique Key",
    "side": "Side - Faction that this ship/object belongs to",
    "artfileroot": "Artfileroot - Where can we find the .obj and 4 .png files?",
    "meshscale": "Meshscale - Scaling for the model in world",
    "radarscale": "Radarscale - Scaling for the 2D top down view on the Radar",
    "exclusionradius": "Exclusionradius - Spherical Collision point of this object",
    "hullpoints": "Hullpoints - How strong the hull is per system (4 systems so hp = value * 4)",
    "long_desc": "Long_desc - Scientific description text",
    "tubecount": "Tubecount - How many torpedo tubes this ship has (players)",
    "baycount": "Baycount - How many shuttles/fighters this ship can have (needs role carrier to launch for AI ships)",
    "internalmapscale": "Internalmapscale - Scaling for the internal damage control map",
    "internalmapw": "Internalmapw - Width of internal map in nodes (port-starboard)",
    "internalmaph": "Internalmaph - Height of internal map in nodes (forward-aft)",
    "internalsymmetry": "Internalsymmetry - Unknown parameter",
    "turn_rate": "Turn_rate - Maximum rate of turn",
    "speed_coeff": "Speed_coeff - Speed coefficient",
    "scan_strength_coeff": "Scan_strength_coeff - Scanning speed coefficient",
    "ship_energy_cost": "Ship_energy_cost - Energy drain when operating",
    "warp_energy_cost": "Warp_energy_cost - Additional energy drain during warp",
    "jump_energy_cost": "Jump_energy_cost - Energy cost to jump (affected by jump distance)",
    "roles": "Roles - Affects AI behavior and selection for generated missions",
    "drone_launch_timer": "Drone_launch_timer - How often the ship launches drones (for AI torpedoes)",
    "shields_front": "Shields_front - For stations: omnidirectional shield; for ships: forward shield strength",
    "shields_rear": "Shields_rear - Rear shield strength for ships",
    "health": "Health - Monster Health",
    "heal_rate": "Heal_rate - Monster health regeneration rate"
}

# --- Edit Torpedo Start Dialog Class ---

# --- Main Ship Editor ---
class ShipEditor:

    # -------------------- path helpers & Orion picker --------------------
    def _get_app_dir(self):
        """Directory of the running executable (PyInstaller) or this script."""
        try:
            if getattr(sys, "frozen", False):
                return os.path.dirname(sys.executable)
            return os.path.dirname(os.path.abspath(__file__))
        except Exception:
            return os.getcwd()

    def _get_data_dir(self):
        """Return <cosmos_root>/data if a root is known; else CWD."""
        cosmos_root = getattr(self, "_cosmos_root", None)
        return os.path.join(cosmos_root, "data") if cosmos_root else os.getcwd()

    def _resolve_data_path(self, filename):
        """Resolve a data file path under <cosmos_root>/data (or CWD when no root)."""
        base = self._get_data_dir()
        return os.path.join(base, filename) if not os.path.isabs(filename) else filename

    def _image_full_path(self, artfileroot: str, size: str = IMAGE_SUFFIX):
        """Resolve a ship image to <cosmos_root>/data/graphics/ships/<root><size>."""
        img_dir = os.path.join(self._get_data_dir(), IMAGE_FOLDER)
        return os.path.join(img_dir, f"{artfileroot}{size}")


    def _abs(self, p: str) -> str:
        """Return an absolute, normalized path for error/info messages."""
        try:
            return os.path.abspath(p)
        except Exception:
            return p

    def _load_orion_paths(self):
        """
        Look for OrionData/Orion.hjson next to the editor executable.
        Return (orion_file_path, ordered list of (label, path)) or (None, []).
        """
        base = self._get_app_dir()
        orion_dir = os.path.join(base, "OrionData")
        orion_file = os.path.join(orion_dir, "Orion.hjson")
        if not os.path.exists(orion_file):
            return None, []
        try:
            with open(orion_file, "r", encoding="utf-8") as f:
                data = hjson.load(f, object_pairs_hook=dict)
        except Exception as e:
            # Don't block startup on a malformed Orion.hjson — we can still let the user browse.
            print(f"Warning: failed to parse Orion.hjson: {e}")
            return orion_file, []

        paths = []
        pf = data.get("Pathfiles")
        # Accept both: Pathfiles: [{ Name: "C:/path", Name2: "D:/path" }, ...]
        if isinstance(pf, list):
            for obj in pf:
                if isinstance(obj, dict):
                    for label, pth in obj.items():
                        if isinstance(pth, str) and pth.strip():
                            paths.append((label.strip(), os.path.normpath(pth.strip())))
        return orion_file, paths

    def _choose_cosmos_dir(self):
        """
        Present a chooser dialog with entries from Orion.hjson (if any) plus a
        'Browse…' button. Returns the selected directory or None.
        If Orion.hjson is parseable, 'Browse…' also offers to save the new entry.
        """
        orion_file, options = self._load_orion_paths()

        # If no options found, just prompt to browse.
        if not options:
            picked = filedialog.askdirectory(title="Locate your Cosmos install (folder with the executable)")
            if not picked:
                return None
            picked = os.path.normpath(picked)

            # If Orion.hjson does NOT exist (orion_file is None), and the picked folder
            # contains data/shipData.yaml/json, create OrionData/Orion.hjson and seed Pathfiles.
            if orion_file is None:
                base = self._get_app_dir()
                orion_dir = os.path.join(base, "OrionData")
                orion_path = os.path.join(orion_dir, "Orion.hjson")
                try:
                    data_dir = os.path.join(picked, "data")
                    has_yaml = os.path.exists(os.path.join(data_dir, YAML_PATH))
                    has_hjson = os.path.exists(os.path.join(data_dir, HJSON_PATH))
                    if has_yaml or has_hjson:
                        label = simpledialog.askstring(
                            "Name this path",
                            "Give this install a name (e.g., 'Cosmos Main'):",
                            parent=self.master
                        )
                        if label:
                            os.makedirs(orion_dir, exist_ok=True)
                            data = {"Pathfiles": [{label: picked}]}
                            with open(orion_path, "w", encoding="utf-8") as f:
                                hjson.dump(data, f, ensure_ascii=False, indent=2)
                            print(f"Created Orion.hjson at {orion_path} with entry '{label}': {picked}")
                except Exception as e:
                    # Non-fatal: we still proceed with the chosen directory for this session.
                    print(f"Note: failed to create Orion.hjson: {e}")
            return picked

        # Build a simple picker window
        win = tk.Toplevel(self.master)
        win.title("Select Cosmos Directory")
        win.geometry("560x360+120+120")
        win.transient(self.master)
        win.grab_set()

        ttk.Label(win, text="Choose a Cosmos installation folder:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10, pady=(10, 6))
        frame = ttk.Frame(win)
        frame.pack(fill="both", expand=True, padx=10, pady=(0, 8))

        listbox = tk.Listbox(frame)
        listbox.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(frame, orient="vertical", command=listbox.yview)
        sb.pack(side="right", fill="y")
        listbox.config(yscrollcommand=sb.set)

        for label, pth in options:
            listbox.insert(tk.END, f"{label} — {pth}")

        chosen = {"value": None}

        def use_selected():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("Pick one", "Please select an entry or click Browse…")
                return
            _, pth = options[sel[0]]
            chosen["value"] = pth
            win.destroy()

        def browse_new():
            directory = filedialog.askdirectory(title="Locate your Cosmos install (folder with the executable)")
            if not directory:
                return
            # Offer to save this into Orion.hjson (only if we successfully read it).
            label = simpledialog.askstring("Name this path", "Give this install a name (e.g., 'Cosmos Main'):", parent=win)
            save_ok = False
            if label:
                try:
                    with open(orion_file, "r", encoding="utf-8") as f:
                        data = hjson.load(f, object_pairs_hook=dict)
                    if not isinstance(data.get("Pathfiles"), list):
                        data["Pathfiles"] = []
                    # Append to the first mapping object if present, else create one.
                    if data["Pathfiles"] and isinstance(data["Pathfiles"][0], dict):
                        data["Pathfiles"][0][label] = directory
                    else:
                        data["Pathfiles"].append({label: directory})
                    with open(orion_file, "w", encoding="utf-8") as f:
                        hjson.dump(data, f, ensure_ascii=False, indent=2)
                    # Update the UI list
                    options.append((label, os.path.normpath(directory)))
                    listbox.insert(tk.END, f"{label} — {os.path.normpath(directory)}")
                    save_ok = True
                except Exception as e:
                    print(f"Note: could not update Orion.hjson: {e}")
            if not save_ok:
                messagebox.showinfo("Using selection", "We'll use this folder for this session.")
            chosen["value"] = os.path.normpath(directory)
            win.destroy()

        btns = ttk.Frame(win)
        btns.pack(fill="x", padx=10, pady=(4, 10))
        ttk.Button(btns, text="Use Selected", command=use_selected).pack(side="left")
        ttk.Button(btns, text="Browse…", command=browse_new).pack(side="left", padx=8)
        ttk.Button(btns, text="Cancel", command=win.destroy).pack(side="right")

        listbox.bind("<Double-1>", lambda _e: use_selected())
        win.wait_window()
        return chosen["value"]


    def setup_easter_egg(self):
        # Maintain a list to record key presses
        self.easter_keys = []
        # Define your secret code sequence, e.g., Up Up Down Down Left Right Left Right B A
        self.secret_sequence = ["Up", "Up", "Down", "Down"]
        self.master.bind("<Key>", self.check_key_sequence)

    # Moved check_key_sequence out of setup_easter_egg
    def check_key_sequence(self, event):
        self.easter_keys.append(event.keysym)
        if len(self.easter_keys) > len(self.secret_sequence):
            self.easter_keys.pop(0)
        if self.easter_keys == self.secret_sequence:
            messagebox.showinfo("Easter Egg!", "You have discovered the secret spaceship cache!")
            self.activate_easter_egg_mode()
            self.easter_keys = []

    def activate_easter_egg_mode(self):
        import random
        egg_window = tk.Toplevel(self.master)
        egg_window.title("Easter Egg!")
        egg_window.attributes("-topmost", True)
        transparent_color = "#123456"
        egg_window.configure(bg=transparent_color)
        egg_window.attributes("-transparentcolor", transparent_color)
        self.master.update_idletasks()
        w = self.master.winfo_width() or 800
        h = self.master.winfo_height() or 600
        egg_window.geometry(f"{w}x{h}+{self.master.winfo_rootx()}+{self.master.winfo_rooty()}")
        canvas = tk.Canvas(egg_window, bg=transparent_color, highlightthickness=0)
        canvas.pack(fill="both", expand=True)
        egg_window.update_idletasks()
        canvas_width = canvas.winfo_width() or w
        canvas_height = canvas.winfo_height() or h

        try:
            try:
                resample_mode = Image.Resampling.LANCZOS
            except AttributeError:
                resample_mode = Image.ANTIALIAS
            fish_img = Image.open("OrionData/Fish.png")
            fish_img = fish_img.resize((50, 50), resample_mode)
            fish_photo = ImageTk.PhotoImage(fish_img)
        except Exception as e:
            messagebox.showerror("Error", f"Error loading fish image: {e}")
            egg_window.destroy()
            return

        egg_window.fish_photo = fish_photo

        num_fish = 10
        fish_objects = []
        for _ in range(num_fish):
            x = random.randint(0, canvas_width)
            y = random.randint(0, canvas_height)
            dx = random.choice([-3, -2, -1, 1, 2, 3])
            dy = random.choice([-3, -2, -1, 1, 2, 3])
            fish_id = canvas.create_image(x, y, image=fish_photo)
            fish_objects.append({"id": fish_id, "dx": dx, "dy": dy})

        def animate():
            for obj in fish_objects:
                coords = canvas.coords(obj["id"])
                if not coords:
                    continue
                x, y = coords
                new_x = x + obj["dx"]
                new_y = y + obj["dy"]
                if new_x < 0:
                    new_x = canvas_width
                elif new_x > canvas_width:
                    new_x = 0
                if new_y < 0:
                    new_y = canvas_height
                elif new_y > canvas_height:
                    new_y = 0
                canvas.coords(obj["id"], new_x, new_y)
            egg_window.after(50, animate)
        animate()
        egg_window.after(30000, egg_window.destroy)

    def load_data(self):
        """Load ship data from YAML (comment-preserving) if available, else legacy HJSON.
        If neither exists locally, fall back to OrionData/Orion.hjson → Pathfiles picker to locate a Cosmos *root*.
        When a root is chosen, ship files and graphics are looked up under <root>/data.
        """
        try:
            self._cosmos_root = None  # reset any previous root
            # First try current working directory (legacy behavior for dev)
            local_yaml = os.path.join(os.getcwd(), YAML_PATH)
            local_hjson = os.path.join(os.getcwd(), HJSON_PATH)
            if os.path.exists(local_yaml):
                # --- YAML/HJSON path ---
                self.data_path = local_yaml
                with open(local_yaml, "r", encoding="utf-8") as f:
                    ytext = f.read()
                self._raw_text = ytext  # keep original text for surgical mode
                self._yaml_tab_fix_applied = False

                if self._looks_hjsonish(ytext):
                    # Surgical path: parse with hjson for UI data, but keep raw text for saves
                    try:
                        parsed = hjson.loads(ytext, object_pairs_hook=dict)
                    except Exception as e:
                        raise RuntimeError(f"HJSON-style parse failed: {e}")
                    self._save_mode = "surgical"
                    ship_list_key = "#ship-list" if "#ship-list" in parsed else "ship-list"
                    self.ships_data = list(parsed.get(ship_list_key, []))
                    self.header_data = {k: v for k, v in parsed.items() if k != ship_list_key}
                else:
                    # True YAML: ruamel round-trip
                    if YAML is None:
                        raise RuntimeError("ruamel.yaml is required for round-trip YAML. Install with: pip install ruamel.yaml")
                    y = YAML(typ="rt")
                    self._configure_yaml_emitter(y)
                    try:
                        doc = y.load(ytext) or {}
                    except Exception:
                        if "\t" in ytext:
                            fixed = re.sub(r"^(?P<t>\t+)", lambda m: "  " * len(m.group("t")), ytext, flags=re.M)
                            doc = y.load(fixed) or {}
                            self._yaml_tab_fix_applied = True
                        else:
                            raise
                    self._save_mode = "yaml"
                    self._yaml_rt = y
                    self._yaml_doc = doc
                    ship_list_key = "#ship-list" if "#ship-list" in doc else "ship-list"
                    if ship_list_key not in doc or doc.get(ship_list_key) is None:
                        from ruamel.yaml.comments import CommentedSeq, CommentedMap
                        if not isinstance(doc, dict):
                            doc = CommentedMap(); self._yaml_doc = doc
                        doc[ship_list_key] = CommentedSeq()
                    self.ships_data = self._yaml_doc[ship_list_key]
                    self.header_data = {k: v for k, v in self._yaml_doc.items() if k != ship_list_key}

            else:
                # Nothing found locally — consult Orion.hjson
                cosmos_dir = self._choose_cosmos_dir()
                if not cosmos_dir:
                    raise FileNotFoundError(
                        "Could not find required data files and no Cosmos folder was selected.\n"
                        f"Checked:\n  {self._abs(local_yaml)}\n  {self._abs(local_hjson)}"
                    )
                # From now on, resolve data files under <cosmos_root>/data
                self._cosmos_root = cosmos_dir
                data_dir = os.path.join(self._cosmos_root, "data")
                chosen_yaml = os.path.join(data_dir, YAML_PATH)
                chosen_hjson = os.path.join(data_dir, HJSON_PATH)
                if os.path.exists(chosen_yaml):
                    self.data_path = chosen_yaml

                    with open(chosen_yaml, "r", encoding="utf-8") as f:
                        ytext = f.read()
                    self._raw_text = ytext
                    self._yaml_tab_fix_applied = False

                    if self._looks_hjsonish(ytext):
                        # Surgical path for found file
                        try:
                            parsed = hjson.loads(ytext, object_pairs_hook=dict)
                        except Exception as e:
                            raise RuntimeError(f"HJSON-style parse failed: {e}")
                        self._save_mode = "surgical"
                        ship_list_key = "#ship-list" if "#ship-list" in parsed else "ship-list"
                        self.ships_data = list(parsed.get(ship_list_key, []))
                        self.header_data = {k: v for k, v in parsed.items() if k != ship_list_key}
                    else:
                        if YAML is None:
                            raise RuntimeError("ruamel.yaml is required for round-trip YAML. Install with: pip install ruamel.yaml")
                        y = YAML(typ="rt")
                        self._configure_yaml_emitter(y)
                        try:
                            doc = y.load(ytext) or {}
                        except Exception:
                            if "\t" in ytext:
                                fixed = re.sub(r"^(?P<t>\t+)", lambda m: "  " * len(m.group("t")), ytext, flags=re.M)
                                doc = y.load(fixed) or {}
                                self._yaml_tab_fix_applied = True
                            else:
                                raise
                        self._save_mode = "yaml"
                        self._yaml_rt = y
                        self._yaml_doc = doc
                        ship_list_key = "#ship-list" if "#ship-list" in doc else "ship-list"
                        if ship_list_key not in doc or doc.get(ship_list_key) is None:
                            from ruamel.yaml.comments import CommentedSeq, CommentedMap
                            if not isinstance(doc, dict):
                                doc = CommentedMap(); self._yaml_doc = doc
                            doc[ship_list_key] = CommentedSeq()
                        self.ships_data = self._yaml_doc[ship_list_key]
                        self.header_data = {k: v for k, v in self._yaml_doc.items() if k != ship_list_key}

                elif os.path.exists(chosen_hjson):
                    raise FileNotFoundError(
                        "Found shipData.json but HJSON is no longer supported for ship data. "
                        "Please convert your data to shipData.yaml backed by ruamel-compatible YAML."
                    )

                else:
                    raise FileNotFoundError(
                        "Could not find required data files inside the selected folder.\n"
                        f"Selected folder (game root): {self._abs(cosmos_dir)}\n"
                        f"Checked in data dir:\n  {self._abs(chosen_yaml)}\n  {self._abs(chosen_hjson)}"
                     )
                # Ensure at least one ship exists (optional convenience)
                if not self.ships_data:
                    default_ship = {
                        "name": "Unnamed Ship",
                        "key": "unnamed",
                        "side": "Independent",
                        "artfileroot": "unknown",
                        "meshscale": 1.0,
                        "radarscale": 1.0,
                        "exclusionradius": 0,
                    }
                    self.ships_data.append(default_ship)

            # Set an env var so dialogs can resolve images under <root>/data, too.
            try:
                os.environ["COSMOS_DATA_DIR"] = os.path.abspath(self._get_data_dir())
            except Exception:
                pass
            mode = "surgical" if self._save_mode == "surgical" else "YAML"
            print(f"Loaded {len(self.ships_data)} ships from {self.data_path} ({mode}).")
        except Exception as e:
            where = getattr(self, "data_path", self._resolve_data_path(HJSON_PATH))
            awhere = self._abs(where)
            print(f"Error loading data from {awhere}: {e}")
            messagebox.showerror("Error", f"Failed to load data from:\n{awhere}\n\n{e}")
            return

    # --- YAML emitter config (minimal): preserve quotes only ---
    def _configure_yaml_emitter(self, y):
        """
        Configure ruamel.yaml for stable round-trip. We keep this minimal so the
        serializer doesn’t reformat your file:
          - preserve_quotes keeps quoting as authored
        """
        try:
            y.preserve_quotes = True
            # Allow full unicode in output
            y.allow_unicode = True
            # We WANT flow-mapped (brace) style with one item per line.
            # Setting a very small width forces line breaks after commas in flow style.
            try:
                y.width = 1
            except Exception:
                pass
            try:
                # Avoid compact seq/map formatting
                y.compact(seq_seq=False, seq_map=False)
            except Exception:
                pass
            # Match original line endings (Windows vs Unix) to avoid noisy diffs
            lb = getattr(self, "_yaml_line_break", "\n")
            try:
                # ruamel expects "\n" or "\r\n"
                y.line_break = lb
            except Exception:
                pass
        except Exception:
            # Future ruamel versions may change attributes — ignore gracefully.
            pass


    # -------------------- HJSON-ish detection & surgical save helpers --------------------
    def _looks_hjsonish(self, text: str) -> bool:
        """
        Treat file as HJSON/JSON only when the ship list uses a bracketed array
        (…: [ … ]). Avoid regex: scan after the key to the first significant char.
        """
        keys = ('"#ship-list"', "'#ship-list'", '#ship-list',
                '"ship-list"',  "'ship-list'",  'ship-list')
        n = len(text)
        for k in keys:
            p = text.find(k)
            if p == -1:
                continue
            c = text.find(':', p + len(k))
            if c == -1:
                continue
            i = c + 1
            while i < n:
                ch = text[i]
                # whitespace
                if ch.isspace():
                    i += 1; continue
                # // or # line comments
                if text.startswith('//', i) or ch == '#':
                    nl = text.find('\n', i)
                    i = n if nl == -1 else nl + 1
                    continue
                # /* block comments */
                if text.startswith('/*', i):
                    end = text.find('*/', i + 2)
                    i = n if end == -1 else end + 2
                    continue
                # first significant char
                return ch == '['
        return False

    def _repr_hjson_scalar(self, v):
        """Conservative JSON/HJSON scalar repr without changing surrounding whitespace."""
        if isinstance(v, dict) or isinstance(v, list):
            # Handled by pretty printers (use wrappers below)
            raise TypeError("_repr_hjson_scalar received non-scalar")
        if isinstance(v, bool):
            return "true" if v else "false"
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            # Keep floats as floats (avoid integerization surprises)
            return str(v)
        # strings/other → JSON-quoted
        s = str(v).replace("\\", "\\\\").replace('"', '\\"')
        return f'"{s}"'

    # ---------- JSON/HJSON pretty printers (respect nesting) ----------
    def _repr_hjson_value_pretty(self, v, indent: str, step: str = "  ") -> str:
        """Render any JSON-like value (dict/list/scalar) with indentation."""
        if isinstance(v, dict):
            return self._repr_hjson_map_pretty(v, indent, step)
        if isinstance(v, list):
            return self._repr_hjson_list_pretty(v, indent, step)
        # scalar
        try:
            return self._repr_hjson_scalar(v)
        except TypeError:
            # dict/list would land here if called wrongly; be safe
            return self._repr_hjson_value_pretty(str(v), indent, step)

    def _repr_hjson_map_pretty(self, d: dict, indent: str, step: str = "  ") -> str:
        """Multi-line pretty map with keys as JSON strings."""
        lines = ["{"]
        inner = indent + step
        for i, (k, val) in enumerate(d.items()):
            key_str = '"' + str(k).replace('"', '\\"') + '"'
            val_str = self._repr_hjson_value_pretty(val, inner, step)
            # If value is multi-line, keep as-is; just prefix first line
            if "\n" in val_str:
                first, *rest = val_str.splitlines()
                lines.append(f"{inner}{key_str}: {first}")
                for r in rest:
                    lines.append(f"{inner}{r}")
                lines[-1] += ","  # comma after the block value
            else:
                lines.append(f"{inner}{key_str}: {val_str},")
        # Remove trailing comma if we actually added any lines
        if len(lines) > 1 and lines[-1].endswith(","):
            lines[-1] = lines[-1][:-1]
        lines.append(indent + "}")
        return "\n".join(lines)

    def _repr_hjson_list_pretty(self, arr: list, indent: str, step: str = "  ") -> str:
        """Multi-line pretty list; every element ends with a comma."""
        lines = ["["]
        inner = indent + step
        for v in arr:
            v_str = self._repr_hjson_value_pretty(v, inner, step)
            if "\n" in v_str:
                first, *rest = v_str.splitlines()
                lines.append(f"{inner}{first}")
                for r in rest:
                    lines.append(f"{inner}{r}")
                lines[-1] += ","
            else:
                lines.append(f"{inner}{v_str},")
        lines.append(indent + "]")
        return "\n".join(lines)


    def _repr_hjson_list(self, arr):
        # Simple list rendering: [a, b, c] with JSON scalars
        parts = []
        for v in arr:
            parts.append(self._repr_hjson_scalar(v))
        return "[" + ", ".join(parts) + "]"


    # -------- Orion banner: replace-or-insert after the first '{' --------
    def _format_editor_banner_value(self) -> str:
        """Return the text that goes inside the quotes for #OrionShipEditor."""
        try:
            tz = ZoneInfo("Europe/London") if ZoneInfo else None
        except Exception:
            tz = None
        ts = datetime.now(tz).strftime("%Y-%m-%d %H:%M:%S")
        return f" This file edited on {ts} by the Orion Ship Editor Program Version  {EDITOR_VERSION} for {GAME_VERSION}"

    def _upsert_editor_banner_simple(self, text: str) -> str:
        """
        Replace existing "#OrionShipEditor": "..." line (anywhere), else insert a new one
        immediately after the first '{' in the file. Keeps the file's newline style.
        """
        import re
        lb = "\r\n" if "\r\n" in text else "\n"
        banner = self._format_editor_banner_value().replace("\\", "\\\\").replace('"', '\\"')
        # 1) Replace existing line if present (preserve indentation)
        rx = re.compile(r'^(?P<indent>[ \t]*)"?#OrionShipEditor"\s*:\s*"(?:\\.|[^"\\])*"\s*,?\s*$',
                        flags=re.M)
        if rx.search(text):
            return rx.sub(lambda m: f'{m.group("indent")}"#OrionShipEditor": "{banner}",', text, count=1)
        # 2) Insert after the first '{'
        brace = text.find("{")
        if brace == -1:
            return text  # not a JSON/HJSON-ish file; leave untouched
        insert = f'{lb}  "#OrionShipEditor": "{banner}",'
        return text[:brace+1] + insert + text[brace+1:]

    def _find_ship_block_span(self, text: str, id_value: str):
        """
        Robustly find the { ... } block inside the ship list whose Key or name equals id_value.
        Strategy: narrow to the ship-list array, iterate top-level {…} entries by scanning braces
        (quote-aware), parse each block with hjson to compare identifiers, return exact (start,end).
        """
        region = self._extract_ship_list_region(text)
        if not region:
            return None
        arr_start, arr_end = region
        for b_start, b_end in self._iter_top_level_flow_maps(text, arr_start, arr_end):
            block = text[b_start:b_end]
            try:
                data = hjson.loads(block, object_pairs_hook=dict)
            except Exception:
                # Skip blocks we can’t parse; they might be comments or malformed entries
                continue
            ident = str(data.get("Key") or data.get("key") or data.get("name") or "")
            if ident == str(id_value):
                return (b_start, b_end)
        return None

    def _extract_ship_list_region(self, text: str):
        """
        Return (start,end) of the '#ship-list' / 'ship-list' array *including* brackets.
        Robust to quoted keys, comments, and newlines between ':' and '[' — no regex.
        """
        keys = ('"#ship-list"', "'#ship-list'", '#ship-list',
                '"ship-list"',  "'ship-list'",  'ship-list')
        n = len(text)
        key_pos = -1
        key_len = 0
        for k in keys:
            p = text.find(k)
            if p != -1:
                key_pos = p
                key_len = len(k)
                break
        if key_pos == -1:
            return None
        colon = text.find(':', key_pos + key_len)
        if colon == -1:
            return None
        # advance to first significant token after the colon
        i = colon + 1
        def skip_line_comment(j: int) -> int:
            nl = text.find('\n', j)
            return n if nl == -1 else nl + 1
        def skip_block_comment(j: int) -> int:
            end = text.find('*/', j + 2)
            return n if end == -1 else end + 2
        while i < n:
            ch = text[i]
            if ch.isspace():
                i += 1; continue
            if text.startswith('//', i) or ch == '#':
                i = skip_line_comment(i); continue
            if text.startswith('/*', i):
                i = skip_block_comment(i); continue
            if ch == '[':
                break
            # unexpected token
            return None
        if i >= n or text[i] != "[":
            return None
        # Balance brackets to find matching ']'
        arr_start = i
        depth = 0
        in_dq = False
        in_sq = False
        esc = False
        k = i
        end = None
        while k < n:
            ch = text[k]
            if not (in_dq or in_sq):
                if text[k:k+2] == "//" or text[k] == "#":
                    k = skip_line_comment(k); continue
                if text[k:k+2] == "/*":
                    k = skip_block_comment(k); continue
            if in_dq or in_sq:
                if esc:
                    esc = False
                elif ch == '\\\\':
                    esc = True
                elif in_dq and ch == '"':
                    in_dq = False
                elif in_sq and ch == "'":
                    in_sq = False
            else:
                if ch == '"':
                    in_dq = True
                elif ch == "'":
                    in_sq = True
                elif ch == '[':
                    depth += 1
                elif ch == ']':
                    depth -= 1
                    if depth == 0:
                        end = k + 1
                        break
            k += 1
        if end is None:
            return None
        return (arr_start, end)

    def _iter_top_level_flow_maps(self, text: str, start: int, end: int):
        """
        Yield (b_start,b_end) for each top-level '{...}' directly inside [start,end) range.
        Quote-aware; ignores nested braces until balanced.
        """
        i = start
        in_q = None
        depth = 0
        b_start = None
        while i < end:
            ch = text[i]
            if in_q:
                if ch == in_q and text[i-1] != '\\':
                    in_q = None
            else:
                if ch in ('"', "'"):
                    in_q = ch
                elif ch == '{':
                    if depth == 0:
                        b_start = i
                    depth += 1
                elif ch == '}':
                    depth -= 1
                    if depth == 0 and b_start is not None:
                        yield (b_start, i + 1)
                        b_start = None
            i += 1

    # ---------- Robust object finder: scan whole file for key=<value> and return the { ... } span ----------
    def _iter_object_spans_for_key(self, raw: str, wanted: str):
        """
        Yield (start, end) offsets of top-level {...} objects that contain:
            key/Key : "wanted"  OR  'wanted'  OR  bare wanted
        Tolerates whitespace and //, #, and /* */ comments between ':' and the value.
        Balances braces and respects strings.
        """
        import re

        # Allow comments after colon
        comment = r'(?:\s*(?://[^\n]*|#[^\n]*|/\*.*?\*/))*'
        # Value can be "double-quoted", 'single-quoted', or a bare token (letters/digits/_-.)
        val_alt = r'(?:"(?P<dval>[^"\\]*(?:\\.[^"\\]*)*)"|\'(?P<sval>[^\'\\]*(?:\\.[^\'\\]*)*)\'|(?P<bval>[A-Za-z0-9_\-\.]+))'
        # Key name: key/Key, possibly quoted.  IMPORTANT: no prefix constraint.
        # This lets us match keys even if they appear after comments, weird spacing, etc.
        key_pat = re.compile(
            rf'(?:"?[Kk]ey"?)\s*:\s*{comment}(?P<val>{val_alt})',
            flags=re.DOTALL | re.MULTILINE,
        )

        def skip_comment(i):
            n = len(raw)
            if raw[i:i+2] == "//" or raw[i] == "#":
                while i < n and raw[i] != "\n":
                    i += 1
            return i

        def find_object_bounds(anchor: int):
            """
            Find the { ... } bounds containing 'anchor'.
            1) Try precise backward brace-walk to the opening '{.
            2) If that fails, FALL BACK to nearest line-anchored.
            Then do a robust forward balance with comment/string handling.
            """
            n = len(raw)
            # ---------------- Backward precise walk ----------------
            i = anchor
            depth = 0
            in_str = False
            in_dq = False
            in_sq = False
            esc = False
            start = None
            while i >= 0:
                ch = raw[i]
                if in_dq or in_sq:
                    if esc:
                        esc = False
                    elif ch == '\\':
                        esc = True
                    elif in_dq and ch == '"':
                        in_dq = False
                    elif in_sq and ch == "'":
                        in_sq = False
                else:
                    if ch == '"':
                        in_dq = True
                    elif ch == "'":
                        in_sq = True
                    elif ch == '}':
                        depth += 1
                    elif ch == '{':
                        if depth == 0:
                            start = i
                            break
                        depth -= 1
                i -= 1
            # ---------------- Fallback to nearest line-anchored '{' ----------------
            if start is None:
                # Find the nearest "{” that starts a line (ignores bracket noise in comments above)
                import re
                up_to = raw[:anchor]
                m_last = None
                for m in re.finditer(r'^\s*\{', up_to, flags=re.M):
                    m_last = m
                if m_last:
                    start = m_last.end() - 1  # point at the '{'
                else:
                    return None

            # ---------------- Forward robust balance ----------------
            i = start
            depth = 0
            in_dq = False
            in_sq = False
            esc = False

            def skip_line_comment(j):
                # Skip //... or #... to end-of-line
                while j < n and raw[j] != "\n":
                    j += 1
                return j
            def skip_block_comment(j):
                # Skip /* ... */ safely
                j += 2  # after '/*'
                while j < n-1:
                    if raw[j] == '*' and raw[j+1] == '/':
                        return j + 2
                    j += 1
                return n

            while i < n:
                ch = raw[i]
                # comments only outside strings
                if not (in_dq or in_sq):
                    if raw[i:i+2] == "//" or raw[i] == "#":
                        i = skip_line_comment(i); continue
                    if raw[i:i+2] == "/*":
                        i = skip_block_comment(i); continue
                if in_dq or in_sq:
                    if esc:
                        esc = False
                    elif ch == '\\':
                        esc = True
                    elif in_dq and ch == '"':
                        in_dq = False
                    elif in_sq and ch == "'":
                        in_sq = False
                else:
                    if ch == '"':
                        in_dq = True
                    elif ch == "'":
                        in_sq = True
                    elif ch == '{':
                        depth += 1
                    elif ch == '}':
                        depth -= 1
                        if depth == 0:
                            return (start, i + 1)
                i += 1
            return None

        idx = 0
        for m in key_pat.finditer(raw):
            idx += 1
            mv = m.group("dval") or m.group("sval") or m.group("bval") or ""
            if self._debug_surgical:
                snippet = raw[max(0, m.start()-40): m.end()+40].replace("\n", "\\n")
                print(f"[surgical] key-match#{idx}: found value='{mv}' wanted='{wanted}' at {m.start()}..{m.end()}  ctx='{snippet[:160]}...'")
            if mv != wanted:
                continue
            bounds = find_object_bounds(m.start())
            if bounds:
                if self._debug_surgical:
                    print(f"[surgical] key='{mv}' -> object bounds {bounds}")
                yield bounds
            elif self._debug_surgical:
                print(f"[surgical] key='{mv}' but failed to find balanced object around index {m.start()}")


    def _patch_scalar_in_block(self, block: str, key_names, value_str: str) -> tuple[str, int]:
        """
        Replace the value of the first occurrence of any key in key_names within the given { ... } block.
        Handles quoted/bare keys, optional comments/whitespace after ':', and
        replaces the *entire* scalar value correctly even if it contains commas
        (e.g., "station,carrier") or escaped quotes.
        """
        import re
        key_alt = "|".join(re.escape(k) for k in key_names)
        comment = r'(?:\s*(?://[^\n]*|#[^\n]*|/\*.*?\*/))*'
        # Value token can be:
        #   - double-quoted string with escapes: " ... "
        #   - single-quoted string with escapes: ' ... '
        #   - bare token up to comma/brace/newline/bracket
        value_tok = r'(?:"(?:\\.|[^"\\])*"|\'(?:\\.|[^\'\\])*\'+|[^,\r\n}\]]+)'
        # group(1): key + colon + comments/space; group(2): complete scalar token
        pattern = rf'((?:"(?:{key_alt})"|(?:{key_alt}))\s*:\s*{comment})({value_tok})'
        rx = re.compile(pattern, flags=re.M | re.S)
        def _repl(m):
            # Use function replacement to avoid \1 + digit being parsed as group 10, etc.
            return m.group(1) + value_str
        new_block, n = rx.subn(_repl, block, count=1)
        if getattr(self, "_debug_surgical", False):
            print(f"[surgical]  patched key(s) {key_names} -> {value_str} (replacements={n})")
        return new_block, n


    def _patch_list_in_block(self, block: str, key_names, new_list_str: str | None, new_list_obj: list | None = None) -> tuple[str, int]:
        """
        Replace the entire [ ... ] list for the first occurrence of any key in key_names inside { ... } block.
        Robust to whitespace/comments after ':', and to multi-line pretty-printed lists.
        Preserves trailing comma and any comments after the list by only replacing the bracketed region.
        """
        import re
        key_alt = "|".join(re.escape(k) for k in key_names)
        comment = r'(?:\s*(?://[^\n]*|#[^\n]*|/\*.*?\*/))*'
        # Find the key position and land just after colon+comments
        head_rx = re.compile(rf'(?:"(?:{key_alt})"|(?:{key_alt}))\s*:\s*{comment}', flags=re.M | re.S)
        mh = head_rx.search(block)
        if not mh:
            return block, 0
        pos = mh.end()
        n = len(block)
        # Determine current indentation (spaces before the '[' line)
        line_start = block.rfind("\n", 0, pos) + 1
        base_indent = block[line_start:pos]
        base_indent = re.match(r'[ \t]*', base_indent).group(0)
        # Skip whitespace/comments until we hit the opening '['
        i = pos
        def skip_line_comment(j):
            while j < n and block[j] != "\n":
                j += 1
            return j
        def skip_block_comment(j):
            j += 2
            while j < n - 1:
                if block[j] == "*" and block[j+1] == "/":
                    return j + 2
                j += 1
            return n
        while i < n:
            ch = block[i]
            if block[i:i+2] == "//" or block[i] == "#":
                i = skip_line_comment(i); continue
            if block[i:i+2] == "/*":
                i = skip_block_comment(i); continue
            if ch.isspace():
                i += 1; continue
            if ch == '[':
                break
            # If we encounter something else (e.g., a scalar), fall back to scalar replacement
            if new_list_str is None and new_list_obj is not None:
                # If the file unexpectedly holds a scalar, replace scalar with a proper list string
                rendered = self._repr_hjson_list_pretty(new_list_obj, base_indent)
                return self._patch_scalar_in_block(block, key_names, rendered)
            return self._patch_scalar_in_block(block, key_names, new_list_str or "[]")
        if i >= n or block[i] != '[':
            return block, 0
        # Now balance brackets to find the matching ']'
        start = i
        depth = 0
        in_dq = False
        in_sq = False
        esc = False
        j = i
        while j < n:
            ch = block[j]
            # Handle comments only when outside strings
            if not (in_dq or in_sq):
                if block[j:j+2] == "//" or block[j] == "#":
                    j = skip_line_comment(j); continue
                if block[j:j+2] == "/*":
                    j = skip_block_comment(j); continue
            if in_dq or in_sq:
                if esc:
                    esc = False
                elif ch == '\\':
                    esc = True
                elif in_dq and ch == '"':
                    in_dq = False
                elif in_sq and ch == "'":
                    in_sq = False
            else:
                if ch == '"':
                    in_dq = True
                elif ch == "'":
                    in_sq = True
                elif ch == '[':
                    depth += 1
                elif ch == ']':
                    depth -= 1
                    if depth == 0:
                        end = j + 1
                        # Replace only the [ ... ] slice; keep any following comma/comments intact
                        # If caller provided an object, render with file's indentation.
                        if new_list_str is None and new_list_obj is not None:
                            rendered = self._repr_hjson_list_pretty(new_list_obj, base_indent)
                        else:
                            rendered = new_list_str if new_list_str is not None else "[]"
                        new_block = block[:start] + rendered + block[end:]
                        if getattr(self, "_debug_surgical", False):
                            print(f"[surgical]  patched list key(s) {key_names} -> {new_list_str} at slice {start}:{end}")
                        return new_block, 1
            j += 1
        return block, 0

    # ---------- New ship insertion & deletion (surgical mode) ----------
    def _deduce_key_style(self, example_block_dict: dict) -> dict:
        """
        Inspect an existing parsed ship dict and decide preferred casing for common keys.
        Returns a mapping like {"key": "Key", "name": "name", "side": "side", "artfileroot": "artfileroot"}.
        """
        style = {
            "key": "Key" if "Key" in example_block_dict else ("key" if "key" in example_block_dict else "Key"),
            "name": "Name" if "Name" in example_block_dict else ("name" if "name" in example_block_dict else "name"),
            "side": "Side" if "Side" in example_block_dict else ("side" if "side" in example_block_dict else "side"),
            "artfileroot": "Artfileroot" if "Artfileroot" in example_block_dict else ("artfileroot" if "artfileroot" in example_block_dict else "artfileroot"),
        }
        return style

    def _ship_list_indentation(self, raw: str, arr_start: int, arr_end: int) -> tuple[str, str, str]:
        """
        Return (base_indent, item_indent, linebreak) inferred from the #ship-list region.
        """
        lb = "\r\n" if "\r\n" in raw else "\n"
        # Indent of the line containing '['
        line_start = raw.rfind("\n", 0, arr_start) + 1
        base_indent = raw[line_start:arr_start]
        # Try to infer item indent from the first block; default to base + two spaces.
        item_indent = base_indent + "  "
        for b_start, _ in self._iter_top_level_flow_maps(raw, arr_start, arr_end):
            ls = raw.rfind("\n", 0, b_start) + 1
            item_indent = raw[ls:b_start]
            break
        return base_indent, item_indent, lb

    def _render_ship_block(self, ship_dict: dict, item_indent: str) -> str:
        """
        Render a single ship map as multi-line JSON/HJSON-like text using existing pretty printers.
        """
        # Ensure simple ordering for readability: Key/name/side/artfileroot first if present.
        ordering = ["Key", "key", "Name", "name", "Side", "side", "Artfileroot", "artfileroot"]
        d = dict(ship_dict)
        ordered = {}
        for k in ordering:
            if k in d:
                ordered[k] = d.pop(k)
        ordered.update(d)
        return self._repr_hjson_map_pretty(ordered, item_indent[:-2] if len(item_indent) >= 2 else "")

    def _surgical_insert_ship_block(self, ship: dict, cluster_by_side: bool = True) -> None:
        """
        Insert a new ship block into the raw text inside #ship-list.
        If cluster_by_side is True, place the new entry after the last ship with the same 'side' (casefold);
        otherwise append to the end of the array.
        """
        if not self._raw_text:
            raise RuntimeError("No raw text captured for surgical insert.")
        region = self._extract_ship_list_region(self._raw_text)
        if not region:
            raise RuntimeError("Could not find #ship-list region for insertion.")
        arr_start, arr_end = region  # arr_start points at '[', arr_end just after ']'
        base_indent, item_indent, lb = self._ship_list_indentation(self._raw_text, arr_start, arr_end)

        # Probe an existing block for key casing preferences
        key_style = None
        for bs, be in self._iter_top_level_flow_maps(self._raw_text, arr_start, arr_end):
            try:
                sample = hjson.loads(self._raw_text[bs:be], object_pairs_hook=dict)
                key_style = self._deduce_key_style(sample)
                break
            except Exception:
                continue
        if key_style is None:
            key_style = {"key": "Key", "name": "name", "side": "side", "artfileroot": "artfileroot"}

        # Apply key style to the new ship copy
        s = dict(ship)
        if "Key" in s and key_style["key"] == "key":
            s["key"] = s.pop("Key")
        if "key" in s and key_style["key"] == "Key":
            s["Key"] = s.pop("key")
        if "Name" in s and key_style["name"] == "name":
            s["name"] = s.pop("Name")
        if "name" in s and key_style["name"] == "Name":
            s["Name"] = s.pop("name")
        if "Side" in s and key_style["side"] == "side":
            s["side"] = s.pop("Side")
        if "side" in s and key_style["side"] == "Side":
            s["Side"] = s.pop("side")
        if "Artfileroot" in s and key_style["artfileroot"] == "artfileroot":
            s["artfileroot"] = s.pop("Artfileroot")
        if "artfileroot" in s and key_style["artfileroot"] == "Artfileroot":
            s["Artfileroot"] = s.pop("artfileroot")

        # Decide insertion location
        same_side_last_end = None
        new_side = (s.get("side") or s.get("Side") or "").casefold()
        if cluster_by_side and new_side:
            for bs, be in self._iter_top_level_flow_maps(self._raw_text, arr_start, arr_end):
                try:
                    d = hjson.loads(self._raw_text[bs:be], object_pairs_hook=dict)
                except Exception:
                    continue
                side_val = (d.get("side") or d.get("Side") or "").casefold()
                if side_val == new_side:
                    same_side_last_end = be

        insertion_pos = None
        trailing_comma_needed = False
        if same_side_last_end is not None:
            # Insert just after the last same-side block.
            insertion_pos = same_side_last_end
            trailing_comma_needed = True
        else:
            # Append before the closing bracket
            insertion_pos = arr_end - 1  # points at ']'
            # Add comma if the array is non-empty (prev non-space != '[')
            i = insertion_pos - 1
            while i > arr_start and self._raw_text[i].isspace():
                i -= 1
            if self._raw_text[i] != '[':
                trailing_comma_needed = True

        block_txt = self._render_ship_block(s, item_indent)
        prefix = self._raw_text[:insertion_pos]
        suffix = self._raw_text[insertion_pos:]
        add = ""
        if trailing_comma_needed:
            add += "," + lb
        add += item_indent + block_txt.strip()
        if same_side_last_end is not None:
            add += ","
        add += lb

        self._raw_text = prefix + add + suffix

    def _surgical_delete_ship_by_key(self, key_value: str) -> bool:
        """
        Remove the ship block with given Key from the raw text inside #ship-list.
        Returns True if a block was removed.
        """
        if not self._raw_text:
            return False
        region = self._extract_ship_list_region(self._raw_text)
        if not region:
            return False
        arr_start, arr_end = region
        matches = list(self._iter_object_spans_for_key(self._raw_text, str(key_value)))
        if not matches:
            return False
        start, end = matches[0]
        # Trim whitespace around to decide where the separating comma lives
        i = start - 1
        while i > arr_start and self._raw_text[i].isspace():
            i -= 1
        removed = False
        # Case A: there's a comma right before the block -> remove from that comma
        if i >= arr_start and self._raw_text[i] == ',':
            self._raw_text = self._raw_text[:i] + self._raw_text[end:]
            removed = True
        else:
            # Case B: try to consume a trailing comma after the block
            j = end
            n = len(self._raw_text)
            while j < n and self._raw_text[j].isspace():
                j += 1
            if j < n and self._raw_text[j] == ',':
                j += 1
            self._raw_text = self._raw_text[:start] + self._raw_text[j:]
            removed = True
        return removed
    def _surgical_save_current_ship(self, ship: dict, updates: dict):
        """
        Perform an in-place textual patch for the current ship inside self._raw_text.
        Only scalars and simple lists are handled here (e.g., shields).
        """
        if not self._raw_text:
            raise RuntimeError("No raw text captured for surgical save.")

        # Identify by unique Key
        key_val = ship.get("Key") or ship.get("key")
        name_val = ship.get("name")
        span = None
        debug_lines = []
        def _log(line):
            if getattr(self, "_debug_surgical", False):
                print(line)
                debug_lines.append(line)

        if key_val:
            wanted = str(key_val)
            _log(f"[surgical] Saving ship with Key='{wanted}'. Beginning search in raw text...")
            matches = list(self._iter_object_spans_for_key(self._raw_text, wanted))
            _log(f"[surgical] Matches for Key='{wanted}': {len(matches)}")
            if matches:
                # Show a short preview of each candidate block (first line + first few fields)
                if getattr(self, "_debug_surgical", False):
                    for i, (bs, be) in enumerate(matches, 1):
                        block = self._raw_text[bs:be]
                        preview = block.splitlines()[0][:200]
                        _log(f"[surgical]  candidate#{i} span=({bs},{be}) first-line='{preview}'")
                span = matches[0]
        # Optional fallback: locate by name if Key is missing
        if not span and name_val:
            matches = list(self._iter_object_spans_for_key(self._raw_text, str(name_val)))
            if matches:
                span = matches[0]
        if not span:
            # Not found → treat as a brand-new ship entry: insert a new block,
            # then run the patcher against that fresh block (so current edits are applied).
            # Merge updates into a temporary copy for insertion.
            tmp = dict(ship)
            tmp.update(updates)
            self._surgical_insert_ship_block(tmp, cluster_by_side=True)
            # Recompute span (on the inserted block)
            if key_val:
                matches = list(self._iter_object_spans_for_key(self._raw_text, str(key_val)))
                span = matches[0] if matches else None
            elif name_val:
                matches = list(self._iter_object_spans_for_key(self._raw_text, str(name_val)))
                span = matches[0] if matches else None
            if not span:
                raise RuntimeError("Inserted new ship, but could not re-locate it for patching.")

        start, end = span
        block = self._raw_text[start:end]

        # Apply updates (deterministic order helps testing)
        for k in sorted(updates.keys()):
            v = updates[k]
            key_forms = [k]
            if k == "Key": key_forms.append("key")
            if k == "name": key_forms.append("Name")
            if k == "side": key_forms.append("Side")
            if k == "artfileroot": key_forms.append("Artfileroot")
            if isinstance(v, list):
                # Lists may be of scalars or dicts. When dicts present, let the patcher
                # render with proper indentation based on file context.
                contains_dict = any(isinstance(x, dict) for x in v)
                if contains_dict or key_forms[0] == "torpedostart":
                    new_block, n = self._patch_list_in_block(block, key_forms, None, new_list_obj=v)
                else:
                    vstr = self._repr_hjson_list(v)  # existing compact list rendering
                    new_block, n = self._patch_list_in_block(block, key_forms, vstr)
            else:
                vstr = self._repr_hjson_scalar(v)
                new_block, n = self._patch_scalar_in_block(block, key_forms, vstr)
            if n > 0:
                block = new_block
            elif getattr(self, "_debug_surgical", False):
                print(f"[surgical]  no match to patch for key {k} (list={isinstance(v, list)})")

        # Splice back
        self._raw_text = self._raw_text[:start] + block + self._raw_text[end:]
        # Stamp/update the Orion banner right after the first '{'
        self._raw_text = self._upsert_editor_banner_simple(self._raw_text)
        # Write raw text back unchanged except for patched values and banner
        with open(self._resolve_data_path(YAML_PATH), "w", encoding="utf-8") as f:
            f.write(self._raw_text)

    # --- Round-trip: update nodes in place preserving style/ordering/comments ---
    def _rt_update_preserve(self, cm, updates: dict):
        """
        Update keys in a ruamel CommentedMap in-place, preserving:
          - key order
          - comments
          - scalar quoting style (single/double/plain) when replacing strings
        """
        try:
            from ruamel.yaml.comments import CommentedMap as _CMap, CommentedSeq as _CSeq
        except Exception:
            _CMap, _CSeq = dict, list
        for k, new in updates.items():
            if isinstance(cm, dict) and k in cm:
                old = cm[k]
                # Recurse on same-type containers
                if isinstance(old, _CMap) and isinstance(new, dict):
                    self._rt_update_preserve(old, new)
                elif isinstance(old, _CSeq) and isinstance(new, (list, tuple)):
                    if list(old) != list(new):
                        old.clear()
                        old.extend(new)
                # Preserve quoting style for strings
                elif isinstance(old, SQS):
                    cm[k] = SQS(str(new))
                elif isinstance(old, DQS):
                    cm[k] = DQS(str(new))
                elif isinstance(old, PSS):
                    cm[k] = PSS(str(new))
                else:
                    cm[k] = new
            else:
                cm[k] = new


    # --- Key normalization (names & casing) -----------------------------------
    def _rename_key_preserve(self, cm, old, new):
        """
        Rename a key on a ruamel CommentedMap while preserving order and comments.
        Uses rename_key when available; otherwise, emulates it.
        """
        try:
            # ruamel >= 0.18 provides rename_key with comment preservation
            cm.rename_key(old, new)
            return
        except Exception:
            pass
        try:
            from ruamel.yaml.comments import CommentedMap
        except Exception:
            CommentedMap = dict
        if not isinstance(cm, CommentedMap):
            # plain dict fallback
            if old in cm:
                cm[new] = cm.pop(old)
            return
        if old not in cm:
            return
        # Manual preserve: capture value and comments
        val = cm[old]
        # Insert new key right after old's position
        keys = list(cm.keys())
        idx = keys.index(old)
        # remove old, insert new
        del cm[old]
        # Rebuild in order
        rebuilt = CommentedMap()
        for i, k in enumerate(keys):
            if i == idx:
                rebuilt[new] = val
            if k != old:
                rebuilt[k] = cm[k] if k in cm else None
        # transfer comments if possible
        try:
            rebuilt.yaml_set_start_comment(cm.ca.comment[0] if cm.ca and cm.ca.comment else None)
        except Exception:
            pass
        cm.clear(); cm.update(rebuilt)

    def _normalize_ship_keys(self, ship_cm):
        """
        Ensure field names match desired external style:
          Key  (capital K), name, side, artfileroot  (lowercase)
        Accept common variants on input (e.g., 'key', 'Name', 'Side', 'Artfileroot').
        """
        try:
            from ruamel.yaml.comments import CommentedMap
        except Exception:
            CommentedMap = dict
        if not isinstance(ship_cm, CommentedMap):
            return
        # 1) Normalize 'key' → 'Key' (capital K)
        if 'key' in ship_cm and 'Key' not in ship_cm:
            self._rename_key_preserve(ship_cm, 'key', 'Key')
        # If 'Name'/'Side'/'Artfileroot' appear, normalize to lowercase variants
        for old, new in [('Name','name'), ('Side','side'), ('Artfileroot','artfileroot')]:
            if old in ship_cm and new not in ship_cm:
                self._rename_key_preserve(ship_cm, old, new)

        # Unquote simple keys like "name" → name by renaming to the same spelling,
        # which drops the original quoting style during round-trip.
        simple = []
        try:
            simple = list(ship_cm.keys())
        except Exception:
            pass
        for k in simple:
            if isinstance(k, str) and k.strip() and k == re.sub(r'[^A-Za-z0-9_]', '', k):
                # Already normalized 'Key' above if it was 'key'
                target = k
                # Skip if it's already plain 'Key' that we want capitalized
                if k == 'key':
                    target = 'Key'
                # Renaming k→target when equal triggers a re-key with plain style
                self._rename_key_preserve(ship_cm, k, target)

    # --- HJSON-like layout helpers: braces + blank lines between keys ---
    def _set_flow_style_recursive(self, node):
        """
        Force flow/brace style on mappings and keep sequences block style
        (so lists stay readable). Recurse into children.
        """
        try:
            from ruamel.yaml.comments import CommentedMap, CommentedSeq
        except Exception:
            CommentedMap, CommentedSeq = dict, list

        if isinstance(node, CommentedMap):
            try:
                node.fa.set_flow_style()
            except Exception:
                try:
                    node.fa.flow_style = True
                except Exception:
                    pass
            for v in node.values():
                self._set_flow_style_recursive(v)
        elif isinstance(node, CommentedSeq):
            # Keep sequences in block style for readability
            try:
                node.fa.set_block_style()
            except Exception:
                try:
                    node.fa.flow_style = False
                except Exception:
                    pass
            for v in node:
                self._set_flow_style_recursive(v)

    def _blank_line_between_keys(self, cm):
        """
        Insert a blank line before each key by attaching an empty 'before' comment.
        Yields a visual empty line between key/value pairs on dump.
        """
        try:
            from ruamel.yaml.comments import CommentedMap
        except Exception:
            CommentedMap = dict
        if not isinstance(cm, CommentedMap):
            return
        keys = list(cm.keys())
        for i, k in enumerate(keys):
            if i == 0:
                continue  # no leading blank line before first key
            try:
                # A single newline -> one blank line between pairs.
                # Use "\n\n" if you need two blank lines.
                cm.yaml_set_comment_before_after_key(k, before="\n")
            except Exception:
                pass

    def _apply_hjson_layout(self):
        """
        Make the document render with JSON/HJSON curly braces on mappings and
        add a blank line between each key/value pair at top-level and ship maps.
        """
        root = getattr(self, "_yaml_doc", None)
        if root is None:
            return
        # 1) Force brace style on mappings throughout
        self._set_flow_style_recursive(root)
        # 2) Ensure blank line between top-level keys
        self._blank_line_between_keys(root)
        # 3) Ensure per-ship maps get spacing too
        ship_list_key = "#ship-list" if "#ship-list" in root else "ship-list"
        try:
            ships = root.get(ship_list_key, [])
        except Exception:
            ships = []
        for ship in ships:
            # normalize field names for consistent casing before spacing/layout
            self._normalize_ship_keys(ship)
            self._blank_line_between_keys(ship)



    # --- Text post-process to enforce HJSON braces with one-per-line items ---
    def _postprocess_hjson_text(self, text: str) -> str:
        """
        After ruamel dumps the YAML (in flow/brace style), enforce:
          - one key/value per line inside ship maps,
          - keep commas at line ends,
          - leave comments/blank lines intact.
        This looks only inside '#ship-list'/'ship-list' arrays.
        """
        # Regex strategy:
        # Find flow maps within ship list entries: { ... } and replace ", " between pairs with ",\n"
        # We keep it conservative to avoid touching nested lists/maps: only replace at top level commas inside a single {}
        def fix_flow_map(m):
            inner = m.group(1)
            # Replace ", " that separate pairs at top depth (no nested braces/brackets)
            out = []
            depth = 0
            i = 0
            while i < len(inner):
                ch = inner[i]
                if ch in '"\'':
                    # copy quoted string verbatim
                    q = ch
                    out.append(ch)
                    i += 1
                    while i < len(inner):
                        out.append(inner[i])
                        if inner[i] == '\\':
                            i += 2
                            continue
                        if inner[i] == q:
                            i += 1
                            break
                        i += 1
                    continue
                elif ch in '{[':
                    depth += 1; out.append(ch); i += 1; continue
                elif ch in '}]':
                    depth = max(0, depth-1); out.append(ch); i += 1; continue
                elif ch == ',' and depth == 0:
                    # ensure comma at EOL followed by newline
                    out.append(',\n')
                    # skip any single space after comma
                    i += 1
                    if i < len(inner) and inner[i] == ' ':
                        i += 1
                    continue
                else:
                    out.append(ch); i += 1
            return '{' + ''.join(out).strip() + '}'

        # Only run the expensive pass within ship list array text regions:
        ship_list_pattern = re.compile(r'(#ship-list|ship-list)\s*:\s*\[\s*(.*?)\s*\]', re.DOTALL)
        def per_ship_fix(arr_m):
            arr_body = arr_m.group(2)
            # Replace top-level { ... } occurrences in the array body
            fixed = re.sub(r'\{(.*?)\}', lambda m: fix_flow_map(m), arr_body, flags=re.DOTALL)
            return f"{arr_m.group(1)}: [\n{fixed}\n]"
        try:
            return ship_list_pattern.sub(per_ship_fix, text)
        except Exception:
            return text


    # ---- YAML error pretty-printer ----
    def _format_yaml_error(self, err, text):
        """Return a readable message with line/column and a small code frame."""
        try:
            mark = getattr(err, "problem_mark", None)
            if not mark:
                return str(err)
            line, col = mark.line, mark.column
            lines = text.splitlines()
            start = max(0, line - 2)
            end = min(len(lines), line + 3)
            snippet = "\n".join(f"{i+1:5d}| {lines[i]}" for i in range(start, end))
            pointer = " " * (col + 7) + "^"
            base = getattr(err, "problem", str(err)) or str(err)
            return f"YAML error at line {line+1}, column {col+1}:\n{base}\n\n{snippet}\n{pointer}"
        except Exception:
            return str(err)

    def draw_beam_field_overlay(self):
        # Use adjusted image center if available.
        if hasattr(self.ship_canvas, "image_center"):
            cx, cy = self.ship_canvas.image_center
        else:
            w = self.ship_canvas.winfo_width() or 400
            h = self.ship_canvas.winfo_height() or 400
            cx, cy = w // 2, h // 2
        if cy == 0:  # safeguard
            cy = 200
        scale = cy / 2500.0  # Adjust the scale as needed.
        
        beam_ports = self.ships_data[self.current_ship_index].get("hull_port_sets", {}).get("beam Primary Beams", [])
        for beam in beam_ports:
            try:
                barrel_angle = float(beam.get("barrel_angle", 0))
                arc_width = float(beam.get("arcwidth", 0))
                port_range = float(beam.get("range", 0))
            except Exception as e:
                print("Error converting beam port values:", e)
                continue

            arc_color = beam.get("arccolor", "red")
            pixel_radius = port_range * scale
            center_angle = 90 - barrel_angle
            x0, y0 = cx - pixel_radius, cy - pixel_radius
            x1, y1 = cx + pixel_radius, cy + pixel_radius
            
            if arc_width >= 360:
                self.ship_canvas.create_oval(x0, y0, x1, y1, outline=arc_color, width=2)
                angle_rad = math.radians(center_angle)
                x_line = cx + pixel_radius * math.cos(angle_rad)
                y_line = cy - pixel_radius * math.sin(angle_rad)
                self.ship_canvas.create_line(cx, cy, x_line, y_line, fill=arc_color, width=2)
            else:
                start_angle = center_angle - (arc_width / 2)
                self.ship_canvas.create_arc(x0, y0, x1, y1, start=start_angle, extent=arc_width,
                                             style="arc", outline=arc_color, width=2)
                start_rad = math.radians(start_angle)
                end_rad = math.radians(start_angle + arc_width)
                x_start = cx + pixel_radius * math.cos(start_rad)
                y_start = cy - pixel_radius * math.sin(start_rad)
                x_end = cx + pixel_radius * math.cos(end_rad)
                y_end = cy - pixel_radius * math.sin(end_rad)
                self.ship_canvas.create_line(cx, cy, x_start, y_start, fill=arc_color, width=2)
                self.ship_canvas.create_line(cx, cy, x_end, y_end, fill=arc_color, width=2)

    def draw_damage_statistics(self):
        """
        Draw the damage statistics on the ship canvas.
        Displays:
          - Total DPM in the bottom-right corner.
          - Forward DPM in the top-right corner.
          - Rear DPM in the bottom-left corner.
        """
        stats = calculate_damage_statistics(self.ships_data[self.current_ship_index])
        total_dpm = stats["total_dpm"]
        forward_dpm = stats["forward_dpm"]
        rear_dpm = stats["rear_dpm"]

        canvas_w = int(self.ship_canvas.winfo_width() or 400)
        canvas_h = int(self.ship_canvas.winfo_height() or 400)

        # Total DPM (bottom-right)
        self.ship_canvas.create_text(
            canvas_w - 10, canvas_h - 10,
            text=f"Total DPM: {total_dpm:.0f}",
            anchor="se",
            fill="white",
            font=("Arial", 16, "bold")
        )
        # Forward DPM (top-right)
        self.ship_canvas.create_text(
            canvas_w - 10, 10,
            text=f"Forward DPM: {forward_dpm:.0f}",
            anchor="ne",
            fill="white",
            font=("Arial", 16, "bold")
        )
        # Rear DPM (bottom-left)
        self.ship_canvas.create_text(
            10, canvas_h - 10,
            text=f"Rear DPM: {rear_dpm:.0f}",
            anchor="sw",
            fill="white",
            font=("Arial", 16, "bold")
        )


    def __init__(self, master):
        print("Initializing Ship Data Editor...")
        self.master = master
        self.master.title("Ship Data Editor")
        self.selected_side = None
        self.selected_side_key = None  # normalized (casefold) side key
        self._side_groups = {}  # normalized_key -> {"display": str, "ships": [ship,...]}
        self._invalid_side_keys = set()  # ship["key"] for ships forced into Unknown
        self._side_list_order = []  # index -> normalized_key
        self._ship_list_order = []  # index -> ship dict for the current side
        self.ships_data = []
        self.ships_data = []
        # Option B state:
        self._save_mode = "yaml"   # "yaml" (ruamel) or "surgical" (HJSON-ish text patch)
        self._raw_text = None      # original file text for surgical mode
        self._debug_surgical = False  # enable extra diagnostics for surgical mode
        self.raw_hjson = ""
        self.current_ship_index = 0
        self.image_label = None
        self.image_cache = None

        # Optionally initialize the easter egg key sequence.
        self.setup_easter_egg()

        try:
            print("Loading data...")
            self.load_data()
            print("Building GUI...")
            self.build_gui()
            print("Populating side selection...")
            self.populate_side_selection()
        except Exception as e:
            print(f"Error during initialization: {e}")
            messagebox.showerror("Error", f"An error occurred during initialization: {e}")

    def copy_ship(self):
        try:
            # Retrieve the currently selected ship.
            ship = self.ships_data[self.current_ship_index]
            # Convert it to a JSON string (without the header, for clarity).
            ship_str = json.dumps(ship, indent=2, ensure_ascii=False)
            # Clear the clipboard and append the string.
            self.master.clipboard_clear()
            self.master.clipboard_append(ship_str)
            messagebox.showinfo("Copied", "Current ship data copied to clipboard!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while copying ship data: {e}")


    def build_gui(self):
        scale_factor = 1.25

        # --- Header Section (Logo, Title, Developer Logo, and Version Info) ---
        frm_header = ttk.Frame(self.master)
        frm_header.pack(fill="x", padx=10, pady=5)

        # Determine appropriate resampling mode (for Pillow v10+)
        try:
            resample_mode = Image.Resampling.LANCZOS
        except AttributeError:
            resample_mode = Image.ANTIALIAS

        # Load the main logo strictly from <cosmos_root>/data/graphics/CosmosLogo.png
        logo_path = os.path.join(self._get_data_dir(), "graphics", "CosmosLogo.png")
        if os.path.exists(logo_path):
            try:
                logo_img = Image.open(logo_path)
                logo_img = logo_img.resize((165, 125), resample_mode)
                self.logo_image = ImageTk.PhotoImage(logo_img)
            except Exception as e:
                print(f"Error loading Cosmos logo at {self._abs(logo_path)}:", e)
                self.logo_image = None
        else:
            # Not found under <cosmos_root>/data/graphics — display nothing.
            self.logo_image = None

        # Load the developer logo (Fish.png) and resize to 125x125.
        try:
            dev_img = Image.open("OrionData/Fish.png")
            dev_img = dev_img.resize((125, 125), resample_mode)
            self.dev_logo_image = ImageTk.PhotoImage(dev_img)
        except Exception as e:
            print("Error loading Fish.png:", e)
            self.dev_logo_image = None

        # Layout the header items:
        # Main logo on the left.
        if self.logo_image:
            lbl_logo = ttk.Label(frm_header, image=self.logo_image)
            lbl_logo.pack(side="left", padx=5)
        
        # Title in the center.
        lbl_title = ttk.Label(frm_header, text="Ship Data Editor", font=("Arial", 24, "bold"))
        lbl_title.pack(side="left", padx=10, pady=5)
        
        # Developer logo on the right.
        if self.dev_logo_image:
            lbl_dev = ttk.Label(frm_header, image=self.dev_logo_image)
            lbl_dev.pack(side="right", padx=5)
            # Bind a left-click event on the developer logo (Fish image) to trigger the easter egg.
            lbl_dev.bind("<Button-1>", lambda event: self.activate_easter_egg_mode())

        
        # Orion image above the version info (fixed path like dev logo)
        try:
            orion_img = Image.open("OrionData/Orion500.png")
            orion_img = orion_img.resize((170, 170), resample_mode)
            self.orion_logo_image = ImageTk.PhotoImage(orion_img)
        except Exception as e:
            print("Error loading Orion500.png:", e)
            self.orion_logo_image = None

        if self.orion_logo_image:
            # Pack at bottom so it sits ABOVE the version label (which is packed last).
            lbl_orion = ttk.Label(frm_header, image=self.orion_logo_image)
            lbl_orion.pack(side="left", pady=(0, 2))

        # Add versioning information below the title (stays at the very bottom).
        # You can use a smaller font for the version text.
        lbl_version = ttk.Label(frm_header, text=f"Editor {EDITOR_VERSION} | {GAME_VERSION}", font=("Arial", 12))
        lbl_version.pack(side="bottom", pady=5)

        # --- Main Application Layout ---
        frm_main = ttk.Frame(self.master)
        frm_main.pack(padx=10, pady=10, expand=True, fill="both")

        # Top toolbar with buttons
        frm_top = ttk.Frame(frm_main)
        frm_top.grid(row=0, column=0, columnspan=4, pady=(0, 10), sticky="ew")
        for i in range(5):
            frm_top.columnconfigure(i, weight=1)
        btn_save = ttk.Button(frm_top, text="Save", command=self.save_changes, style="TButton")
        btn_save.grid(row=0, column=0, padx=5, pady=5)
        btn_reload = ttk.Button(frm_top, text="Reload", command=self.reload_data, style="TButton")
        btn_reload.grid(row=0, column=1, padx=5, pady=5)
        btn_new = ttk.Button(frm_top, text="New Ship", command=self.new_ship, style="TButton")
        btn_new.grid(row=0, column=2, padx=5, pady=5)
        btn_verify = ttk.Button(frm_top, text="Verify", command=self.verify_json, style="TButton")
        btn_verify.grid(row=0, column=3, padx=5, pady=5)


        # Left panel: Faction/Side list
        frm_left = ttk.Frame(frm_main)
        frm_left.grid(row=1, column=0, padx=5, pady=5, sticky="nsew")
        lbl_side = ttk.Label(frm_left, text="Faction/Side", font=("Arial", int(12 * scale_factor), "bold"))
        lbl_side.pack(pady=5)
        CreateToolTip(lbl_side, help_texts.get("side", ""))
        self.side_listbox = tk.Listbox(frm_left, height=15, font=("Arial", int(10 * scale_factor)))
        self.side_listbox.pack(expand=True, fill="both", padx=5, pady=5)
        self.side_listbox.bind('<<ListboxSelect>>', self.on_side_selected)
        
        # Center panel: Ship/Station list
        frm_center = ttk.Frame(frm_main)
        frm_center.grid(row=1, column=1, padx=5, pady=5, sticky="nsew")
        lbl_ship = ttk.Label(frm_center, text="Ship/Station", font=("Arial", int(12 * scale_factor), "bold"))
        lbl_ship.pack(pady=5)
        CreateToolTip(lbl_ship, "Select a ship to edit its details")
        self.ship_listbox = tk.Listbox(frm_center, height=15, font=("Arial", int(10 * scale_factor)))
        self.ship_listbox.pack(expand=True, fill="both", padx=5, pady=5)
        self.ship_listbox.bind('<<ListboxSelect>>', self.on_ship_selected)
        
        # Right panel: Editable Fields with Tooltips on both Labels and Input Fields
        frm_fields = ttk.Frame(frm_main)
        frm_fields.grid(row=1, column=2, padx=5, pady=5, sticky="nsew")
        self.fields = {}
        field_names = [
            "name", "key", "side", "artfileroot", "meshscale", "radarscale", "exclusionradius",
            "hullpoints", "long_desc", "tubecount", "baycount", "internalmapscale", "internalmapw", 
            "internalmaph", "internalsymmetry", "turn_rate", "speed_coeff", "scan_strength_coeff", 
            "ship_energy_cost", "warp_energy_cost", "jump_energy_cost", "roles", "drone_launch_timer",
            "shields_front", "shields_rear", "health", "heal_rate"
        ]

        for i, key in enumerate(field_names):
            lbl = ttk.Label(frm_fields, text=key.capitalize() + ":", font=("Arial", int(10 * scale_factor), "bold"))
            lbl.grid(row=i, column=0, sticky="e", padx=5, pady=2)
            # Attach tooltip to the label
            if key in help_texts:
                CreateToolTip(lbl, help_texts[key])
            entry = ttk.Entry(frm_fields, font=("Arial", int(10 * scale_factor)))
            entry.grid(row=i, column=1, sticky="ew", padx=5, pady=2)
            # Also attach tooltip to the entry widget.
            if key in help_texts:
                CreateToolTip(entry, help_texts[key])
            self.fields[key] = entry

        # Add the "Edit Torpedo Start" button on the same row as the "tubecount" field.
        # (We assume "tubecount" is in the field_names list.)
        if "tubecount" in field_names:
            tubecount_index = field_names.index("tubecount")
            btn_edit_torpedo = ttk.Button(frm_fields, text="Edit Torpedo Start", command=self.edit_torpedo, style="TButton")
            btn_edit_torpedo.grid(row=tubecount_index, column=2, padx=5, pady=5)

        # Below the field table, add more buttons
        btn_edit_beam = ttk.Button(frm_fields, text="Edit Beam Ports", command=self.edit_beam_ports, style="TButton")
        btn_edit_beam.grid(row=len(field_names) + 1, column=0, padx=5, pady=5, sticky="e")

        btn_edit_exhaust = ttk.Button(frm_fields, text="Edit Exhaust Ports", command=self.edit_exhaust_ports, style="TButton")
        btn_edit_exhaust.grid(row=len(field_names) + 1, column=1, padx=5, pady=5, sticky="w")

        # 3D textured preview (OpenGL)
        btn_view_tex = ttk.Button(frm_fields, text="Textured 3D", command=self.open_textured_3d, style="TButton")
        btn_view_tex.grid(row=len(field_names) + 1, column=2, padx=5, pady=5)
    
        #btn_copy = ttk.Button(frm_top, text="Copy Ship", command=self.copy_ship, style="TButton")
        #btn_copy.grid(row=len(field_names) + 2, column=0, padx=5, pady=5)
        
        btn_copy = ttk.Button(frm_fields, text="Copy Ship", command=self.copy_ship, style="TButton")
        btn_copy.grid(row=len(field_names) + 2, column=0, padx=5, pady=5, sticky="e")
        
        btn_delete = ttk.Button(frm_fields, text="Delete Ship", command=self.delete_ship, style="TButton")
        btn_delete.grid(row=len(field_names) + 2, column=1, padx=5, pady=5, sticky="w")

        frm_fields.columnconfigure(1, weight=1)
        # Rightmost panel: Ship image
        frm_image = ttk.Frame(frm_main)
        frm_image.grid(row=1, column=3, padx=5, pady=5, sticky="nsew")
        self.ship_canvas = tk.Canvas(frm_image, background="black", highlightthickness=0)
        self.ship_canvas.pack(expand=True, fill="both", padx=10, pady=10)

        
        frm_main.columnconfigure(0, weight=1)
        frm_main.columnconfigure(1, weight=1)
        frm_main.columnconfigure(2, weight=2)
        frm_main.columnconfigure(3, weight=2)
        frm_main.rowconfigure(1, weight=1)
        
    def _build_side_groups(self):
        """
        Build case-insensitive side groups.
        - sides are grouped by .casefold()
        - display keeps the first seen value, but if it starts lowercase, we uppercase only the first character
        - blank / non-string sides are pushed into the 'Unknown' bucket and tracked for highlighting
        """
        groups = {}
        invalid_keys = set()

        def norm_key(s):
            return s.casefold().strip()

        for ship in list(self.ships_data or []):
            side_val = ship.get("side", "")
            name = ship.get("name", "Unnamed").strip()
            key = ship.get("key", name).strip()

            # Validate side value
            if not isinstance(side_val, str) or not side_val.strip():
                invalid_keys.add(key)
                # Push to Unknown
                ukey = "unknown"
                if ukey not in groups:
                    groups[ukey] = {"display": "Unknown", "ships": []}
                groups[ukey]["ships"].append(ship)
                continue

            raw = side_val.strip()
            nk = norm_key(raw)
            if nk not in groups:
                # Choose a friendly display: keep original casing except ensure leading letter is uppercase if it was lowercase.
                disp = raw[0].upper() + raw[1:] if raw and raw[0].islower() else raw
                groups[nk] = {"display": disp, "ships": []}
            groups[nk]["ships"].append(ship)

        return groups, invalid_keys

    def populate_side_selection(self):
        # Rebuild groups and remember invalids
        self._side_groups, self._invalid_side_keys = self._build_side_groups()

        # Fill the left list using friendly display names
        self.side_listbox.delete(0, tk.END)
        ordered_keys = sorted(self._side_groups.keys(), key=lambda k: self._side_groups[k]["display"].casefold())
        self._side_list_order = ordered_keys[:]  # keep a stable index→key map
        for k in ordered_keys:
            self.side_listbox.insert(tk.END, self._side_groups[k]["display"])

        # If any invalid/blank sides were found, show an error summary
        if self._invalid_side_keys:
            bad = sorted(self._invalid_side_keys)
            preview = ", ".join(bad[:6]) + ("…" if len(bad) > 6 else "")
            messagebox.showerror(
                "Invalid Side Values",
                f"{len(bad)} ship(s) had a blank or invalid 'side'. They were moved to 'Unknown'.\n\nExamples: {preview}"
            )

        if ordered_keys:
            self.side_listbox.select_set(0)
            self.on_side_selected()
            if not self.ship_listbox.curselection():
                self.ship_listbox.select_set(0)
                self.on_ship_selected()

    def on_side_selected(self, event=None):
        try:
            sel = self.side_listbox.curselection()
            if not sel:
                print("No side selected!")
                return
            idx = sel[0]
            if idx < 0 or idx >= len(self._side_list_order):
                messagebox.showerror("Error", "Invalid side selection.")
                return

            # Resolve normalized key and display from our index→key map
            key = self._side_list_order[idx]
            group = self._side_groups.get(key, {"display": "Unknown", "ships": []})
            self.selected_side_key = key
            self.selected_side = group["display"]
            print(f"Selected side: '{self.selected_side}' (key='{self.selected_side_key}')")

            # Populate ship list from the chosen group
            ships_for_side = list(group.get("ships", []))
            ships_for_side.sort(key=lambda s: s.get("name", ""))

            # keep a stable index→ship map so selection doesn't depend on label text
            self._ship_list_order = ships_for_side[:]

            self.ship_listbox.delete(0, tk.END)
            for row, s in enumerate(self._ship_list_order):
                name = s.get("name", "")
                self.ship_listbox.insert(tk.END, name)
                # Red (and bold if supported) for entries forced into Unknown
                ship_key = s.get("key", name)
                if ship_key in getattr(self, "_invalid_side_keys", set()):
                    try:
                        self.ship_listbox.itemconfig(row, fg="red")
                        self.ship_listbox.itemconfig(row, font=("Arial", 10, "bold"))
                    except Exception:
                        txt = self.ship_listbox.get(row)
                        if not txt.startswith("⚠ "):
                            self.ship_listbox.delete(row)
                            self.ship_listbox.insert(row, f"⚠ {txt}")

            if ships_for_side:
                self.ship_listbox.select_set(0)
                self.on_ship_selected()
        except Exception as e:
            print(f"Error in on_side_selected: {e}")
            messagebox.showerror("Error", f"An error occurred while selecting a side: {e}")

    def on_ship_selected(self, event=None):
        try:
            selected_index = self.ship_listbox.curselection()
            if not selected_index:
                print("No ship selected!")
                return
            idx = selected_index[0]
            if not self.selected_side_key:
                print("⚠️ No cached side selected!")
                messagebox.showerror("Error", "No side selected.")
                return

            if idx < 0 or idx >= len(self._ship_list_order):
                messagebox.showerror("Error", "Invalid ship selection.")
                return

            # Directly resolve ship by index (robust against label prefixes like '⚠ ')
            ship = self._ship_list_order[idx]
            print(f"→ Matched ship key: {ship.get('key')} (side: {ship.get('side')}, name: {ship.get('name')})")
            self.current_ship_index = self.ships_data.index(ship)
            # Loop through fields and update their values:
            for key, entry in self.fields.items():

                # Special handling for shields: assume you have two fields "shields_front" and "shields_rear"
                if key == "shields_front":
                    if ship.get("side", "").lower() != "monster":
                        shield_list = ship.get("shields", [])
                        front = str(shield_list[0]) if len(shield_list) > 0 else ""
                        entry.config(state="normal")
                        entry.delete(0, tk.END)
                        entry.insert(0, front)
                    else:
                        entry.config(state="disabled")
                        entry.delete(0, tk.END)
                elif key == "shields_rear":
                    if ship.get("side", "").lower() != "monster":
                        shield_list = ship.get("shields", [])
                        rear = str(shield_list[1]) if len(shield_list) > 1 else ""
                        # Always allow editing rear shields (even if currently single-valued)
                        entry.config(state="normal")
                        entry.delete(0, tk.END)
                        entry.insert(0, rear)
                    else:
                        entry.config(state="disabled")
                        entry.delete(0, tk.END)
                # Special handling for monster fields: "health" and "heal_rate"
                elif key == "health":
                    if ship.get("side", "").lower() == "monster":
                        entry.config(state="normal")
                        entry.delete(0, tk.END)
                        entry.insert(0, str(ship.get("health", "")))
                    else:
                        entry.config(state="disabled")
                        entry.delete(0, tk.END)
                elif key == "heal_rate":
                    if ship.get("side", "").lower() == "monster":
                        entry.config(state="normal")
                        entry.delete(0, tk.END)
                        entry.insert(0, str(ship.get("heal_rate", "")))
                    else:
                        entry.config(state="disabled")
                        entry.delete(0, tk.END)
                else:
                    value = ship.get(key, "")
                    entry.config(state="normal")
                    entry.delete(0, tk.END)
                    entry.insert(0, str(value))
            # Load the preview image once, after fields are populated
            self.load_ship_image(ship.get("artfileroot", ""))
        except Exception as e:
            print(f"⚠️ Error in on_ship_selected: {e}")
            messagebox.showerror("Error", f"An error occurred while selecting the ship: {e}")

    def load_ship_image(self, artfileroot):
        try:
            artfileroot = os.path.normpath(artfileroot)
            base_dir = os.path.join(self._get_data_dir(), IMAGE_FOLDER)
            image_path_1024 = os.path.join(base_dir, artfileroot + "1024.png")
            image_path_256 = os.path.join(base_dir, artfileroot + "256.png")
            image_path = image_path_1024 if os.path.exists(image_path_1024) else image_path_256
            self.ship_canvas.delete("all")
            if os.path.exists(image_path):
                # Get current canvas dimensions:
                canvas_w = int(self.ship_canvas.winfo_width() or 400)
                canvas_h = int(self.ship_canvas.winfo_height() or 400)
                
                # If canvas dimensions are too small, try again later.
                if canvas_w < 50 or canvas_h < 50:
                    print("Canvas dimensions too small, retrying load_ship_image in 100ms")
                    self.master.after(100, lambda: self.load_ship_image(artfileroot))
                    return
                
                # Determine desired image size—say, 80% of canvas dimensions:
                desired_w = canvas_w * 0.8
                desired_h = canvas_h * 0.8
                
                # Load and scale the image while preserving the aspect ratio:
                img = Image.open(image_path)
                img.thumbnail((desired_w, desired_h))
                img = img.rotate(180)
                img = img.convert("RGBA")
                self.image_cache = ImageTk.PhotoImage(img)
                
                # Compute the canvas center:
                cx, cy = canvas_w // 2, canvas_h // 2
                offset_x = 1  # adjust if necessary
                offset_y = 1  # adjust if necessary
                self.ship_canvas.create_image(cx + offset_x, cy + offset_y, anchor="center", image=self.image_cache)
                
                # Save the center for overlays:
                self.ship_canvas.image_center = (cx + offset_x, cy + offset_y)
                self.ship_canvas.image_width = img.width
                self.ship_canvas.image_height = img.height
                
                self.draw_beam_field_overlay()
                self.draw_damage_statistics()
            else:
                self.ship_canvas.create_text(200, 200, text="Image not found", fill="white")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred while loading the image: {e}")



    def save_changes(self):
        try:
            ship = self.ships_data[self.current_ship_index]
            # Define required numeric fields with defaults.
            int_fields = {
                "hullpoints": 0,
                "tubecount": 0,
                "baycount": 0,
                "internalmapw": 0,
                "internalmaph": 0,
                "internalsymmetry": 1
            }
            float_fields = {
                "meshscale": 1.0,
                "radarscale": 1.0,
                "exclusionradius": 0.0,
                "internalmapscale": 1.0,
                "turn_rate": 0.0,
                "speed_coeff": 1.0,
                "scan_strength_coeff": 1.0,
                "ship_energy_cost": 1.0,
                "warp_energy_cost": 1.0,
                "jump_energy_cost": 1.0,
                "drone_launch_timer": 0.0
            }
            
            new_values = {}
            for key, entry in self.fields.items():
                value = entry.get().strip()
                # Process the shields fields specially.
                if key == "shields_front":
                    if ship.get("side", "").lower() != "monster":
                        if value == "":
                            continue
                        else:
                            try:
                                front = int(value)
                            except ValueError:
                                messagebox.showerror("Error", "Shields (front) must be an integer")
                                return
                            rear_val = self.fields.get("shields_rear").get().strip() if "shields_rear" in self.fields else ""
                            shield_list = ship.get("shields", [])
                            if rear_val == "":
                                # Preserve original shape:
                                #  - if original had one value, keep a single-element list
                                #  - if original had two, keep existing rear
                                if len(shield_list) <= 1:
                                    new_values["shields"] = [front]
                                else:
                                    existing_rear = shield_list[1] if len(shield_list) > 1 else 0
                                    try:
                                        existing_rear_int = int(existing_rear)
                                    except Exception:
                                        existing_rear_int = 0
                                    new_values["shields"] = [front, existing_rear_int]
                            else:
                                try:
                                    rear = int(rear_val)
                                except ValueError:
                                    messagebox.showerror("Error", "Shields (rear) must be an integer")
                                    return
                                new_values["shields"] = [front, rear]
                    # Do nothing for monster ships.
                elif key in int_fields:
                    if value == "":
                        new_values[key] = int_fields[key]
                    else:
                        try:
                            new_values[key] = int(value)
                        except ValueError:
                            messagebox.showerror("Error", f"{key} must be an integer")
                            return
                elif key in float_fields:
                    if value == "":
                        new_values[key] = float_fields[key]
                    else:
                        try:
                            new_values[key] = float(value)
                        except ValueError:
                            messagebox.showerror("Error", f"{key} must be a number (float)")
                            return
                else:
                    new_values[key] = value

                # For YAML: coerce to the *original types* ...
                def _coerce_like(old, val):
                    try:
                        # Containers: recurse
                        if isinstance(old, list) and isinstance(val, list):
                            return [_coerce_like(o, v) if i < len(old) else v
                                    for i, (o, v) in enumerate(zip(old, val))] + val[len(old):]
                        if isinstance(old, dict) and isinstance(val, dict):
                            out = dict(val)
                            for k in val.keys() & old.keys():
                                out[k] = _coerce_like(old[k], val[k])
                            return out
                        # Scalars: align numeric types
                        if isinstance(old, bool):
                            # Keep booleans as-is if user typed truthy/falsey
                            if isinstance(val, str):
                                return val.lower() in ("true", "yes", "1")
                            return bool(val)
                        if isinstance(old, int):
                            # If float but integral, or string digits -> int
                            if isinstance(val, float) and val.is_integer():
                                return int(val)
                            if isinstance(val, str):
                                return int(float(val)) if val.replace('.', '', 1).isdigit() else val
                            if isinstance(val, int):
                                return val
                        if isinstance(old, float):
                            # Cast ints/strings to float if the original was float
                            if isinstance(val, (int, float)):
                                return float(val)
                            if isinstance(val, str):
                                return float(val)
                        # Otherwise, leave value as given (strings keep their style via _rt_update_preserve)
                        return val
                    except Exception:
                        return val

                # Apply coercion against the current YAML node values
                for k in list(new_values.keys()):
                    if k in ship:
                        new_values[k] = _coerce_like(ship[k], new_values[k])



            if self._save_mode == "surgical":
                # Only change text where needed, preserve everything else verbatim.
                # Include collections edited via dialogs so they get patched too.
                try:
                    # 1) torpedoes (top-level list) — ensure list of dicts, not strings
                    if isinstance(ship.get("torpedostart"), (list, tuple)):
                        import ast
                        normalized = []
                        for item in ship["torpedostart"]:
                            if isinstance(item, dict):
                                normalized.append(item)
                            elif isinstance(item, str):
                                # Try to parse python-ish dict string (e.g., "{'Homing': 8}")
                                parsed = None
                                try:
                                    parsed = ast.literal_eval(item)
                                except Exception:
                                    try:
                                        # fallback: very permissive hjson parse
                                        parsed = hjson.loads(item)
                                    except Exception:
                                        parsed = {"_raw": item}
                                normalized.append(parsed if isinstance(parsed, dict) else {"_raw": str(item)})
                            else:
                                # If dialog gave tuples like ("Homing", 8)
                                try:
                                    key, val = item
                                    normalized.append({str(key): int(val)})
                                except Exception:
                                    normalized.append({"_raw": str(item)})
                        new_values["torpedostart"] = normalized


                    # 2) hull_port_sets sublists edited in dialogs
                    hps = ship.get("hull_port_sets") or {}
                    if isinstance(hps.get("beam Primary Beams"), (list, tuple)):
                        # patch the array after the "beam Primary Beams" key inside the block
                        new_values["beam Primary Beams"] = list(hps["beam Primary Beams"])
                    if isinstance(hps.get("exhaust"), (list, tuple)):
                        # patch the array after the "exhaust" key inside the block
                        new_values["exhaust"] = list(hps["exhaust"])

                    self._surgical_save_current_ship(ship, new_values)
                except Exception as e:
                    if getattr(self, "_debug_surgical", False):
                        print(f"[surgical] ERROR during save: {e}")
                    raise
                self.data_path = self._resolve_data_path(YAML_PATH)
                messagebox.showinfo("Saved", "Changes saved (surgical HJSON patch; layout/comments preserved).")
            else:
                # Round-trip YAML path (unchanged)
                self._rt_update_preserve(ship, new_values)
                self._normalize_ship_keys(ship)
                self._apply_hjson_layout()
                if YAML is None or getattr(self, "_yaml_doc", None) is None:
                    raise RuntimeError("ruamel.yaml round-trip context not initialized.")
                y = getattr(self, "_yaml_rt", None) or YAML(typ="rt")
                self._configure_yaml_emitter(y)
                import io
                buf = io.StringIO()
                y.dump(self._yaml_doc, buf)
                dumped = buf.getvalue()
                dumped = self._postprocess_hjson_text(dumped)
                # Stamp/update the Orion banner right after the first '{'
                dumped = self._upsert_editor_banner_simple(dumped)
                with open(self._resolve_data_path(YAML_PATH), "w", encoding="utf-8") as f:
                    f.write(dumped)
                self.data_path = self._resolve_data_path(YAML_PATH)
                messagebox.showinfo("Saved", "Changes saved to YAML (comments preserved).")

        except Exception as e:
            print(f"Error saving changes: {e}")
            messagebox.showerror("Error", f"An error occurred while saving changes: {e}")


    def reload_data(self):
        try:
            self.load_data()
            self.populate_side_selection()
        except Exception as e:
            print(f"Error reloading data: {e}")
            messagebox.showerror("Error", f"An error occurred while reloading data: {e}")

    def verify_json(self):
        """Verify the YAML file using ruamel.yaml (round-trip)."""
        try:
            yp = self._resolve_data_path(YAML_PATH)
            ap = self._abs(yp)
            with open(yp, "r", encoding="utf-8") as f:
                ytext = f.read()
                # remember line endings for emitter
                self._yaml_line_break = "\r\n" if "\r\n" in ytext else "\n"
            if YAML is None:
                raise RuntimeError("ruamel.yaml is required. Install with: pip install ruamel.yaml")
            YAML(typ="rt").load(ytext)
            messagebox.showinfo("Verify", "YAML verified successfully!")
            self.scan_art_assets()
        except Exception as e:
            messagebox.showerror("Verification Error", f"Error verifying data at:\n{ap}\n\n{e}")

    # --- Asset Verification (artfileroot) ---
    def scan_art_assets(self):
        """
        For each ship, verify that the following files exist (preserving any subfolders in artfileroot):
          <artfileroot>.obj
          <artfileroot>_diffuse.png
          <artfileroot>_emissive.png
          <artfileroot>_normal.png
          <artfileroot>_specular.png

        If any are missing, show a warning dialog with per-ship rows and an
        'Open Folder' button to jump to the directory.
        """
        try:
            data_dir = self._get_data_dir()
            roots = [
                os.path.join(data_dir, IMAGE_FOLDER),                 # usually "<root>/data/graphics/ships/"
                os.path.join(data_dir, "graphics", "ships"),
                os.path.join(data_dir, "graphics", "models"),
            ]

            problems = []   # list of dicts {label, dir_hint}

            for ship in list(self.ships_data or []):
                afr = str(ship.get("artfileroot", "")).strip()
                ship_name = str(ship.get("name", ship.get("key", "Unknown"))).strip() or "Unknown"
                if not afr:
                    problems.append({
                        "label": f"{ship_name}: artfileroot is empty / missing",
                        "dir_hint": IMAGE_FOLDER
                    })
                    continue

                # Build candidate paths for each required file type, across all roots
                required_specs = {
                    "obj":       [os.path.join(r, afr + ".obj") for r in roots],
                    "diffuse":   [os.path.join(r, afr + "_diffuse.png")   for r in roots],
                    "emissive":  [os.path.join(r, afr + "_emissive.png")  for r in roots],
                    "normal":    [os.path.join(r, afr + "_normal.png")    for r in roots],
                    "specular":  [os.path.join(r, afr + "_specular.png")  for r in roots],
                }

                missing = []
                first_existing_path = None

                # Check presence; remember one existing file to infer a good folder to open
                for kind, candidates in required_specs.items():
                    exists = False
                    for p in candidates:
                        if os.path.exists(p):
                            exists = True
                            if first_existing_path is None:
                                first_existing_path = p
                            break
                    if not exists:
                        # Record the exact filename (without root) for clarity
                        if kind == "obj":
                            missing.append(f"{os.path.basename(afr)}.obj")
                        else:
                            missing.append(f"{os.path.basename(afr)}_{kind}.png")

                if missing:
                    # Choose a sensible folder to open: where something exists, or a default
                    if first_existing_path:
                        dir_hint = os.path.dirname(first_existing_path)
                    else:
                        # Default to the ships graphics folder while preserving subfolders of afr
                        base = os.path.join(data_dir, IMAGE_FOLDER)
                        dir_hint = os.path.join(base, os.path.dirname(afr)) if os.path.dirname(afr) else base
                    problems.append({
                        "label": f"{ship_name} ({afr}): missing {', '.join(missing)}",
                        "dir_hint": os.path.normpath(dir_hint)
                    })

            if not problems:
                messagebox.showinfo("Assets", "All art assets present for all ships.")
                return

            # Build a small warning dialog with a list and an Open Folder button
            win = tk.Toplevel(self.master)
            win.title("Asset Warnings")
            win.geometry("820x420+100+100")
            win.transient(self.master)
            win.grab_set()

            ttk.Label(win, text="The following ships have missing art assets:", font=("Arial", 11, "bold")).pack(anchor="w", padx=10, pady=(10, 4))

            frame = ttk.Frame(win)
            frame.pack(fill="both", expand=True, padx=10, pady=5)

            # Use a listbox for simple per-row selection
            listbox = tk.Listbox(frame, selectmode="browse")
            listbox.pack(side="left", fill="both", expand=True)

            scrollbar = ttk.Scrollbar(frame, orient="vertical", command=listbox.yview)
            scrollbar.pack(side="right", fill="y")
            listbox.config(yscrollcommand=scrollbar.set)

            # Map list index -> dir to open
            dir_map = {}
            for i, item in enumerate(problems):
                listbox.insert(tk.END, item["label"])
                dir_map[i] = item["dir_hint"]

            btns = ttk.Frame(win)
            btns.pack(fill="x", padx=10, pady=(6, 10))

            def open_selected():
                sel = listbox.curselection()
                if not sel:
                    messagebox.showwarning("Open Folder", "Select a row first.")
                    return
                self._open_folder(dir_map[int(sel[0])])

            def copy_report():
                try:
                    win.clipboard_clear()
                    win.clipboard_append("\n".join([p["label"] for p in problems]))
                    messagebox.showinfo("Copied", "Report copied to clipboard.")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not copy report: {e}")

            ttk.Button(btns, text="Open Folder…", command=open_selected).pack(side="left")
            ttk.Button(btns, text="Copy Report", command=copy_report).pack(side="left", padx=8)
            ttk.Button(btns, text="Close", command=win.destroy).pack(side="right")

            # Double-click convenience
            listbox.bind("<Double-1>", lambda _e: open_selected())

        except Exception as e:
            messagebox.showerror("Asset Scan Error", f"An error occurred while scanning art assets:\n{e}")

    def _open_folder(self, path):
        """Open the given folder in the platform's file explorer."""
        try:
            path = os.path.abspath(path)
            if platform.system() == "Windows":
                os.startfile(path)  # type: ignore[attr-defined]
            elif platform.system() == "Darwin":
                subprocess.Popen(["open", path])
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("Open Folder", f"Could not open folder:\n{path}\n\n{e}")



    def save_as_yaml(self):
        """Always write YAML (conversion/export). If loaded via ruamel, preserve comments."""
        try:
            if YAML is not None and getattr(self, "_yaml_doc", None) is not None:
                y = getattr(self, "_yaml_rt", None) or YAML(typ="rt")
                self._configure_yaml_emitter(y)
                with open(self._resolve_data_path(YAML_PATH), "w", encoding="utf-8") as f:
                    y.dump(self._yaml_doc, f)
                self.data_path = self._resolve_data_path(YAML_PATH)
                self.data_format = "yaml"
                messagebox.showinfo("Saved", f"Saved YAML with comments preserved:\n{YAML_PATH}")
            else:
                if yaml is None:
                    raise RuntimeError("Cannot save YAML: PyYAML is not installed. Run: pip install pyyaml")
                root = dict(self.header_data)
                root["#ship-list"] = self.ships_data
                with open(self._resolve_data_path(YAML_PATH), "w", encoding="utf-8") as f:
                    yaml.safe_dump(root, f, sort_keys=False, allow_unicode=True, default_flow_style=False)
                self.data_path = self._resolve_data_path(YAML_PATH)
                self.data_format = "yaml"
                messagebox.showinfo("Saved", f"Exported to YAML:\n{YAML_PATH}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save YAML:\n{e}")

    def new_ship(self):
        existing_sides = sorted(set(ship["side"] for ship in self.ships_data))
        dialog = NewShipDialog(self.master, existing_sides)
        if dialog.result:
            new_ship = dialog.result
            self.ships_data.append(new_ship)
            print(f"New ship created with key: {new_ship['key']}")
            self.populate_side_selection()

    def delete_ship(self):
        try:
            confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete this ship?")
            if not confirm:
                return
            ship = self.ships_data[self.current_ship_index]
            print(f"Deleting ship: {ship['key']}")
            self.ships_data.remove(ship)
            # Persist deletion to disk:
            if getattr(self, "_save_mode", "surgical") == "surgical":
                ident = ship.get("Key") or ship.get("key") or ship.get("name")
                if ident is not None and self._surgical_delete_ship_by_key(str(ident)):
                    # Stamp/update banner and write out
                    self._raw_text = self._upsert_editor_banner_simple(self._raw_text or "")
                    with open(self._resolve_data_path(YAML_PATH), "w", encoding="utf-8") as f:
                        f.write(self._raw_text)
                else:
                    messagebox.showwarning("Delete", "Entry removed from the list, but it was not found in the file.")
            else:
                # YAML round-trip: write immediately for consistency
                if YAML is not None and getattr(self, "_yaml_doc", None) is not None:
                    self._apply_hjson_layout()
                    y = getattr(self, "_yaml_rt", None) or YAML(typ="rt")
                    self._configure_yaml_emitter(y)
                    import io
                    buf = io.StringIO()
                    y.dump(self._yaml_doc, buf)
                    dumped = buf.getvalue()
                    dumped = self._postprocess_hjson_text(dumped)
                    dumped = self._upsert_editor_banner_simple(dumped)
                    with open(self._resolve_data_path(YAML_PATH), "w", encoding="utf-8") as f:
                        f.write(dumped)
            self.populate_side_selection()
        except Exception as e:
            print(f"Error deleting ship: {e}")
            messagebox.showerror("Error", f"An error occurred while deleting the ship: {e}")

    def edit_torpedo(self):
        ship = self.ships_data[self.current_ship_index]
        current_torpedo = ship.get("torpedostart", [])
        available_types = set()
        for s in self.ships_data:
            if "torpedostart" in s:
                for d in s["torpedostart"]:
                    for t in d.keys():
                        available_types.add(t)
        available_types = sorted(list(available_types))
        if not available_types:
            available_types = ["Homing", "Nuke", "EMP", "Mine"]
        dialog = EditTorpedoDialog(self.master, current_torpedo, available_types)
        if dialog.result is not None:
            ship["torpedostart"] = dialog.result
            print(f"Updated torpedostart: {ship['torpedostart']}")
            messagebox.showinfo("Updated", "Torpedo start values updated.")

    def edit_beam_ports(self):
        ship = self.ships_data[self.current_ship_index]
        hull_ports = ship.get("hull_port_sets", {})
        current_beam = hull_ports.get("beam Primary Beams", [])
        artfileroot = ship.get("artfileroot", "")
        dialog = EditBeamPortsDialog(self.master, current_beam, ship_artfileroot=artfileroot, open_3d_callback=self.open_textured_3d)
        if dialog.result is not None:
            if "hull_port_sets" not in ship:
                ship["hull_port_sets"] = {}
            ship["hull_port_sets"]["beam Primary Beams"] = dialog.result
            print(f"Updated beam ports: {ship['hull_port_sets']['beam Primary Beams']}")
            # messagebox.showinfo("Updated", "Beam port values updated.")
            # *** NEW: Update the main editor overlay after beam edits ***
            self.load_ship_image(artfileroot)


    def edit_exhaust_ports(self):
        ship = self.ships_data[self.current_ship_index]
        hull_ports = ship.get("hull_port_sets", {})
        current_exhaust = hull_ports.get("exhaust", [])
        dialog = EditExhaustPortsDialog(self.master, current_exhaust, open_3d_callback=self.open_textured_3d)
        if dialog.result is not None:
            if "hull_port_sets" not in ship:
                ship["hull_port_sets"] = {}
            ship["hull_port_sets"]["exhaust"] = dialog.result
            print(f"Updated exhaust ports: {ship['hull_port_sets']['exhaust']}")
            messagebox.showinfo("Updated", "Exhaust port values updated.")


    # --- 3D Textured viewer (OpenGL) integration ---
    def find_obj_and_texture(self, artfileroot):
        """
        Locate the OBJ and a fallback texture (PNG) for the given artfileroot.
        This respects your existing IMAGE_FOLDER and graphics tree layout.
        """
        import os
        # Normalize afr and strip any leading slash so joins don't get discarded.
        afr = str(artfileroot or "").strip().lstrip("\\/")
        data_dir = self._get_data_dir()
        # Candidate OBJ locations (preserve any subfolders in artfileroot)
        obj_candidates = [
            os.path.join(data_dir, IMAGE_FOLDER, afr + ".obj"),
            os.path.join(data_dir, "graphics", "ships", afr + ".obj"),
            os.path.join(data_dir, "graphics", "models", afr + ".obj"),
        ]
        obj_path = next((p for p in obj_candidates if os.path.exists(p)), None)

        # Texture rule: <artfileroot>_diffuse.png
        tex_candidates = []
        if obj_path:
            # First try next to the OBJ (keeps nested folders like "A/Ted_diffuse.png")
            obj_dir = os.path.dirname(obj_path)
            base_name = os.path.basename(artfileroot)
            tex_candidates.append(os.path.join(obj_dir, base_name + "_diffuse.png"))

        # Also try common graphics roots while preserving subfolders in artfileroot
        tex_candidates.extend([
            os.path.join(data_dir, IMAGE_FOLDER, afr + "_diffuse.png"),
            os.path.join(data_dir, "graphics", "ships", afr + "_diffuse.png"),
            os.path.join(data_dir, "graphics", "models", afr + "_diffuse.png"),
        ])

        tex_path = next((p for p in tex_candidates if os.path.exists(p)), None)
        # Stash for debugging UI if needed
        self._last_obj_candidates = [os.path.abspath(p) for p in obj_candidates]
        self._last_tex_candidates = [os.path.abspath(p) for p in tex_candidates]
        return obj_path, tex_path

    def open_textured_3d(self, parent=None):
        """
        Open a Toplevel window with an OpenGL textured preview of the ship's OBJ.
        """
        if ObjTexturedGLFrame is None:
            msg = (
                "OpenGL viewer not available.\n\n"
                "Please install dependencies:\n"
                "  pip install pyopengltk PyOpenGL Pillow\n"
                "and ensure OrionData/obj_view_gl.py exists next to this app.\n"
            )
            if OBJ_VIEW_GL_IMPORT_ERROR:
                msg += f"\nImport error details:\n{OBJ_VIEW_GL_IMPORT_ERROR}"
            # Extra hint for case-sensitive filesystems
            msg += (
                "\n\nHints:\n"
                " • If you’re on macOS/Linux, ensure the file is named exactly 'obj_view_gl.py' (lowercase).\n"
                " • Ensure there is a file \"shipEditor/obj_view_gl.py\" next to \"shipEditor/dialogs.py\", and that the \"shipEditor\" package is importable.\n"
            )
            messagebox.showerror("Textured 3D", msg)
            return

        ship = self.ships_data[self.current_ship_index]
        artfileroot = ship.get("artfileroot", "")
        obj_path, tex_path = self.find_obj_and_texture(artfileroot)
        if not obj_path:
            tried = "\n  ".join(getattr(self, "_last_obj_candidates", []))
            messagebox.showerror(
                "Textured 3D",
                "Could not find OBJ for artfileroot '{}'.\n\nSearched:\n  {}".format(artfileroot, tried or "(no candidates)")
            )
            return
        win = tk.Toplevel(parent or self.master)
        win.title(f"Textured 3D – {ship.get('name','')}")
        win.geometry("900x700")
        try:
            # Toolbar with toggle buttons for overlays + picked position readout
            toolbar = ttk.Frame(win)
            toolbar.pack(side="top", fill="x")
            picked_var = tk.StringVar(value="Pick with Left Click")
            ttk.Label(toolbar, text="Picked Pos:").pack(side="left", padx=(6,2))
            picked_entry = ttk.Entry(toolbar, textvariable=picked_var, width=40)
            picked_entry.configure(state="readonly")
            picked_entry.pack(side="left", padx=(0,10), pady=4)

            # Create the GL viewer below the toolbar
            viewer = ObjTexturedGLFrame(
                win,
                obj_path=obj_path,
                fallback_texture=tex_path,
                width=900,
                height=700
            )
            viewer.pack(side="top", fill="both", expand=True)

            # Variables bound to viewer flags
            beams_var = tk.BooleanVar(value=True)
            exhaust_var = tk.BooleanVar(value=False)

            def toggle_beams():
                viewer.show_beams = bool(beams_var.get())
                viewer.after(0, viewer.redraw)

            def toggle_exhaust():
                viewer.show_exhaust = bool(exhaust_var.get())
                viewer.after(0, viewer.redraw)

            btn_beams = ttk.Checkbutton(
                toolbar, text="Beam Ports", variable=beams_var, command=toggle_beams
            )
            btn_exhaust = ttk.Checkbutton(
                toolbar, text="Exhaust Ports", variable=exhaust_var, command=toggle_exhaust
            )
            btn_beams.pack(side="left", padx=6, pady=4)
            btn_exhaust.pack(side="left", padx=6, pady=4)

            # Provide an overlay provider stub (replace with real 3D port mapping when available)
            def _overlay_provider():
                """
                Build overlay gizmos from HJSON port sets in *model (OBJ) space*.
                Viewer will normalize these to match the mesh.
                """
                beams = []
                exhaust = []
                ports = ship.get("hull_port_sets", {}) if isinstance(ship, dict) else {}

                # --- BEAMS ---
                # Direction: use barrel_angle (degrees, yaw about +Y) if present; else +Z.
                # Length: scale from 'range' (normalized units ~ 0.15..0.60).
                for p in (ports.get("beam Primary Beams", []) or []):
                    pos = tuple(p.get("position", (0.0, 0.0, 0.0)))
                    ang = float(p.get("barrel_angle", 0.0))
                    r  = float(p.get("range", 1200.0))
                    # yaw about +Y
                    rad = math.radians(ang)
                    dx, dy, dz = math.sin(rad), 0.0, math.cos(rad)
                    # normalize d just in case
                    mag = max((dx*dx + dy*dy + dz*dz) ** 0.5, 1e-6)
                    dx, dy, dz = dx/mag, dy/mag, dz/mag
                    # 0.15..0.60 based on range (tweak to taste)
                    L = max(0.15, min(0.60, 0.15 + (r / 3000.0) * 0.45))
                    # prefer arccolor then color; pass through as 'col'
                    col = p.get("arccolor") or p.get("color")
                    beams.append({"p": pos, "d": (dx, dy, dz), "len": L, "col": col})

                # --- EXHAUST ---
                # Direction: -Z (aft) unless you add per-port orientation.
                for p in (ports.get("exhaust", []) or []):
                    pos = tuple(p.get("position", (0.0, 0.0, 0.0)))
                    col = p.get("color")  # often "orange"
                    exhaust.append({"p": pos, "d": (0.0, 0.0, -1.0), "len": 0.30, "col": col})

                return {"beams": beams, "exhaust": exhaust}

            viewer.set_overlay_provider(_overlay_provider)
            viewer.after(50, viewer.redraw)

            # Receive picks from the viewer, copy to clipboard as "x, y, z", and show on the toolbar
            def _on_pick(pos):
                x, y, z = pos
                raw = f"{x:.6f}, {y:.6f}, {z:.6f}"
                legacy = f"[{x:.6f},{y:.6f},{z:.6f}],"
                # Copy legacy format to clipboard so the dialog's Paste Pos accepts it
                try:
                    win.clipboard_clear()
                    win.clipboard_append(legacy)
                except Exception:
                    pass
                # Display raw (more readable) in the viewer toolbar
                picked_var.set(raw)
            viewer.set_pick_callback(_on_pick)

        except Exception as e:
            win.destroy()
            messagebox.showerror("Textured 3D", f"Failed to open viewer:\\n{e}")


if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = ShipEditor(root)
        root.mainloop()
    except Exception as e:
        print(f"Error in main: {e}")
        messagebox.showerror("Error", f"An error occurred: {e}")
