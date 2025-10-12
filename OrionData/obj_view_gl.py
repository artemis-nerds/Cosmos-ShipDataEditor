
import os
import math
import tkinter as tk
from PIL import Image
from pyopengltk import OpenGLFrame
from OpenGL.GL import *
from OpenGL.GLU import *
from OpenGL.error import GLError

def _triangulate(indices):
    # Fan triangulation: (0,i,i+1)
    tris = []
    for i in range(1, len(indices) - 1):
        tris.append((indices[0], indices[i], indices[i+1]))
    return tris

class ObjModel:
    def __init__(self):
        self.v = []     # [(x,y,z)]
        self.vt = []    # [(u,v)]
        self.vn = []    # [(nx,ny,nz)]
        # triangles as list of [(vi, ti, ni), (vi, ti, ni), (vi, ti, ni)]
        self.triangles = []
        self.bounds_center = (0.0, 0.0, 0.0)
        self.bounds_radius = 1.0
        self.center = (0.0, 0.0, 0.0)   # original OBJ centroid before normalization
        self.scale  = 1.0               # normalization scale (1 / max_distance)
        self.triangles = []
        self.bounds_center = (0.0, 0.0, 0.0)
        self.bounds_radius = 1.0
        self.center = (0.0, 0.0, 0.0)  # original OBJ centroid before normalization
        self.scale = 1.0  # normalization scale (1 / max_distance)
        self.texture_path = None

    def load(self, obj_path):
        self.__init__()
        base_dir = os.path.dirname(obj_path)
        mtl_libs = []
        active_mtl = None
        mtl_map = {}  # name -> { 'map_Kd': path }

        with open(obj_path, "r", encoding="utf-8", errors="ignore") as f:
            faces = []
            for line in f:
                if not line or line.startswith("#"):
                    continue
                parts = line.strip().split()
                if not parts:
                    continue
                tag = parts[0]
                if tag == "v" and len(parts) >= 4:
                    self.v.append(tuple(map(float, parts[1:4])))
                elif tag == "vt" and len(parts) >= 3:
                    u = float(parts[1]); v = float(parts[2])
                    self.vt.append((u, v))
                elif tag == "vn" and len(parts) >= 4:
                    self.vn.append(tuple(map(float, parts[1:4])))
                elif tag == "f" and len(parts) >= 4:
                    fidx = []
                    def _idx(val, n):
                        """OBJ indices: positive are 1-based; negative are relative to the end."""
                        if not val:
                            return None
                        i = int(val)
                        if i > 0:
                            return i - 1
                        if i < 0:
                            return n + i
                        return None
                    for p in parts[1:]:
                        vi = ti = ni = None
                        if "/" in p:
                            toks = p.split("/")
                            if len(toks) == 3:
                                vi = _idx(toks[0], len(self.v))
                                ti = _idx(toks[1], len(self.vt))
                                ni = _idx(toks[2], len(self.vn))
                            elif len(toks) == 2:
                                vi = _idx(toks[0], len(self.v))
                                ti = _idx(toks[1], len(self.vt))
                            else:  # e.g., 'v//n'
                                vi = _idx(toks[0], len(self.v))
                        else:
                            vi = _idx(p, len(self.v))
                        fidx.append((vi, ti, ni, active_mtl))
                    faces.append(fidx)
                elif tag == "mtllib" and len(parts) >= 2:
                    mtl_libs.extend(parts[1:])
                elif tag == "usemtl" and len(parts) >= 2:
                    active_mtl = parts[1]

        # Parse first mtl that has map_Kd
        for mtl in mtl_libs:
            mpath = os.path.join(base_dir, mtl)
            if not os.path.exists(mpath):
                continue
            try:
                with open(mpath, "r", encoding="utf-8", errors="ignore") as mf:
                    name = None
                    props = {}
                    for line in mf:
                        parts = line.strip().split()
                        if not parts:
                            continue
                        if parts[0] == "newmtl" and len(parts) >= 2:
                            if name is not None:
                                mtl_map[name] = props
                            name = parts[1]
                            props = {}
                        elif parts[0] == "map_Kd" and len(parts) >= 2:
                            props["map_Kd"] = " ".join(parts[1:])
                    if name is not None:
                        mtl_map[name] = props
            except Exception:
                pass

        # Triangulate and build final list
        tris = []
        for poly in faces:
            for a, b, c in _triangulate(poly):
                tris.append([a, b, c])

        # Normalize / compute bounds (also record center/scale for overlay normalization)
        if self.v:
            xs, ys, zs = zip(*self.v)
            cx = sum(xs)/len(xs)
            cy = sum(ys)/len(ys)
            cz = sum(zs)/len(zs)
            centered = [(x-cx, y-cy, z-cz) for x,y,z in self.v]
            max_d = max((x*x+y*y+z*z)**0.5 for x,y,z in centered) or 1.0
            self.center = (cx, cy, cz)
            self.scale  = 1.0 / max_d
            self.v = [(x*self.scale, y*self.scale, z*self.scale) for x,y,z in centered]
            self.bounds_center = (0.0, 0.0, 0.0)
            self.bounds_radius = 1.0

        # Convert maps to triangles
        self.triangles = []
        for tri in tris:
            out = []
            for (vi, ti, ni, mname) in tri:
                vx = self.v[vi] if vi is not None and 0 <= vi < len(self.v) else (0.0, 0.0, 0.0)
                vt = self.vt[ti] if ti is not None and 0 <= ti < len(self.vt) else (0.0, 0.0)
                vn = self.vn[ni] if ni is not None and 0 <= ni < len(self.vn) else (0.0, 0.0, 1.0)
                out.append((vx, vt, vn, mname))
            self.triangles.append(out)

        # Try to pick a diffuse texture: pick the first map_Kd in the used material set
        used_mtls = [m for tri in self.triangles for (_,_,_,m) in tri if m]
        tex = None
        for m in used_mtls:
            props = mtl_map.get(m, {})
            if "map_Kd" in props:
                tex = os.path.join(base_dir, props["map_Kd"])
                break
        self.texture_path = tex if tex and os.path.exists(tex) else None


class ObjTexturedGLFrame(OpenGLFrame):
    """
    Controls:
      - LEFT click: pick position on the mesh
      - RIGHT mouse drag (or middle on some Macs): orbit (yaw/pitch)
      - Mouse wheel: zoom in/out
      - W: toggle wireframe
      - R: reset view
    """
    def __init__(self, master, obj_path, fallback_texture=None, bg=(0.06,0.06,0.06,1.0), **kwargs):

        super().__init__(master, **kwargs)
        self.bg = bg
        self.model = ObjModel()
        self.model.load(obj_path)
        self.texture_id = None
        self.fallback_texture = fallback_texture
        self.yaw = 30.0
        self.pitch = -15.0
        self.zoom = 2.5
        self.wireframe = False
        self._drag_last = None
        # --- UV helpers / toggles ---
        self.flip_v = False   # flip V (most OBJs need this in OpenGL)
        self.flip_u = False  # optional U flip
        self.swap_uv = False # swap U<->V (rare but some exports do this)
        self.wrap_clamp = False  # False=REPEAT (default), True=CLAMP_TO_EDGE
        self._gl_ready = False # GL context readiness flag
        self.show_beams = True
        self.show_exhaust = False
        self.overlay_provider = None  # callable returning dict with "beams"/"exhaust"
        self.animate = 1
        self.after_idle(self.redraw)
        self.bind("<Enter>", lambda e: self.focus_set()) # Stop animation when widget is destroyed to avoid stray GL calls
        self.bind("<Destroy>", lambda e: self._on_destroy())


        # ---- Input bindings ----
        # Left click: pick 3D position on the mesh
        self.bind("<Button-1>", self._on_pick_click)
        # Right-drag orbit (Windows/Linux use Button-3 for right click)
        self.bind("<ButtonPress-3>", self._on_drag_start)
        self.bind("<B3-Motion>", self._on_drag_move)
        # Some Tk builds on macOS map secondary click to Button-2 (middle)
        self.bind("<ButtonPress-2>", self._on_drag_start)
        self.bind("<B2-Motion>", self._on_drag_move)
        # Scroll wheel zoom
        self.bind("<MouseWheel>", self._on_wheel)            # Windows/Linux
        self.bind("<Button-4>", lambda e: self._zoom_dir(1)) # X11/mac legacy
        self.bind("<Button-5>", lambda e: self._zoom_dir(-1))
        # Keyboard
        self.bind_all("<Key>", self._on_key)

    # ---- Picking API ----
    def set_pick_callback(self, callback):
        """Provide a function (x,y,z)->None to receive picked coords in ORIGINAL OBJ space."""
        self.on_pick = callback

    def _on_pick_click(self, event):
        """Left-click: compute intersection with mesh and report in original OBJ space."""
        hit = self._pick(event.x, event.y)
        if not hit:
            self.bell()
            return
        # Denormalize back to original OBJ coordinates (inverse of loader's center/scale)
        (x, y, z) = hit
        c = getattr(self.model, "center", (0.0, 0.0, 0.0))
        s = getattr(self.model, "scale",  1.0)
        ox, oy, oz = (x / s + c[0], y / s + c[1], z / s + c[2])
        if callable(getattr(self, "on_pick", None)):
            try:
                self.on_pick((ox, oy, oz))
                return
            except Exception:
                pass
        # Fallback: copy to clipboard as "x, y, z"
        try:
            txt = f"{ox:.6f}, {oy:.6f}, {oz:.6f}"
            self.clipboard_clear(); self.clipboard_append(txt)
        except Exception:
            pass

    def _pick(self, sx, sy):
        """Return intersection point in NORMALIZED model space or None."""
        # Screen -> camera ray (no GL context needed)
        w = max(1, int(self.winfo_width() or 0))
        h = max(1, int(self.winfo_height() or 0))
        if w < 2 or h < 2:
            return None
        # Normalized device coords [-1,1]
        nx = (2.0 * sx / float(w)) - 1.0
        ny = 1.0 - (2.0 * sy / float(h))
        # Camera space ray using fov=45deg
        aspect = w / float(h)
        tanh = math.tan(math.radians(45.0 * 0.5))
        rx, ry, rz = nx * aspect * tanh, ny * tanh, -1.0
        # Rotate into model space: apply inverse view rotations
        yaw  = math.radians(self.yaw)
        pitch= math.radians(self.pitch)
        # inverse rotations are R_y(-yaw) then R_x(-pitch)
        cy, sy = math.cos(-yaw),   math.sin(-yaw)
        cx, sx = math.cos(-pitch), math.sin(-pitch)
        # apply R_x(-pitch)
        ry2 =  ry * cx - rz * sx
        rz2 =  ry * sx + rz * cx
        rx2 =  rx
        # then R_y(-yaw)
        rx3 =  rx2 * cy + rz2 * sy
        rz3 = -rx2 * sy + rz2 * cy
        # camera position in model space (inverse of T(0,0,-zoom) then rotations)
        cam = [0.0, 0.0, self.zoom]
        # rotate camera position by inverse rotations too
        # apply R_x(-pitch)
        cy0 =  cam[1] * cx - cam[2] * sx
        cz0 =  cam[1] * sx + cam[2] * cx
        cx0 =  cam[0]
        # then R_y(-yaw)
        cx1 =  cx0 * cy + cz0 * sy
        cz1 = -cx0 * sy + cz0 * cy
        ro = (cx1, cy0, cz1)              # ray origin (model space)
        rd = self._normalize((rx3, ry2, rz3))  # ray dir (model space)
        return self._raycast(ro, rd)

    def _normalize(self, v):
        x,y,z = v
        m = math.sqrt(x*x + y*y + z*z) or 1.0
        return (x/m, y/m, z/m)

    def _raycast(self, ro, rd):
        """Intersect ray with all model triangles (normalized space). Return nearest hit or None."""
        hit_t = None
        hit_p = None
        # Möller–Trumbore
        EPS = 1e-8
        for tri in self.model.triangles:
            v0 = tri[0][0]; v1 = tri[1][0]; v2 = tri[2][0]
            e1 = (v1[0]-v0[0], v1[1]-v0[1], v1[2]-v0[2])
            e2 = (v2[0]-v0[0], v2[1]-v0[1], v2[2]-v0[2])
            pvec = (rd[1]*e2[2]-rd[2]*e2[1], rd[2]*e2[0]-rd[0]*e2[2], rd[0]*e2[1]-rd[1]*e2[0])
            det = e1[0]*pvec[0] + e1[1]*pvec[1] + e1[2]*pvec[2]
            if -EPS < det < EPS:
                continue
            invDet = 1.0 / det
            tvec = (ro[0]-v0[0], ro[1]-v0[1], ro[2]-v0[2])
            u = (tvec[0]*pvec[0] + tvec[1]*pvec[1] + tvec[2]*pvec[2]) * invDet
            if u < 0.0 or u > 1.0:
                continue
            qvec = (tvec[1]*e1[2]-tvec[2]*e1[1], tvec[2]*e1[0]-tvec[0]*e1[2], tvec[0]*e1[1]-tvec[1]*e1[0])
            v = (rd[0]*qvec[0] + rd[1]*qvec[1] + rd[2]*qvec[2]) * invDet
            if v < 0.0 or u + v > 1.0:
                continue
            t = (e2[0]*qvec[0] + e2[1]*qvec[1] + e2[2]*qvec[2]) * invDet
            if t > EPS and (hit_t is None or t < hit_t):
                hit_t = t
                hit_p = (ro[0] + rd[0]*t, ro[1] + rd[1]*t, ro[2] + rd[2]*t)
        return hit_p

    def set_overlay_provider(self, provider):
        """Set a callback that returns overlay data.
        provider() -> {
            'beams': [{'p':(x,y,z), 'd':(dx,dy,dz), 'len':L}],
            'exhaust': [{'p':(x,y,z), 'd':(dx,dy,dz), 'len':L}],
        }
        Coordinates must be in OBJ model space (centered & normalized by loader).
        """
        self.overlay_provider = provider

    # ---- OpenGL hooks ----
    def initgl(self):
        r,g,b,a = self.bg
        glClearColor(r, g, b, a)
        glEnable(GL_DEPTH_TEST)
        glEnable(GL_CULL_FACE)
        glCullFace(GL_BACK)
        glEnable(GL_BLEND)
        glBlendFunc(GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA)
        self._gl_ready = True

        # texture
        path = self.model.texture_path or self.fallback_texture
        if path and os.path.exists(path):
            self.texture_id = self._load_texture(path)
            glEnable(GL_TEXTURE_2D)
             # Apply initial wrap mode
            glBindTexture(GL_TEXTURE_2D, self.texture_id)
            glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE if self.wrap_clamp else GL_REPEAT)
            glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE if self.wrap_clamp else GL_REPEAT)
        else:
            self.texture_id = None
            glDisable(GL_TEXTURE_2D)

        # simple lighting off (flat look). You can enable later if you add normals & lights.

    def redraw(self):
        # Guard against redraws without a valid GL context or with tiny sizes
        if not self._gl_ready or not self.winfo_exists() or not self.winfo_ismapped():
            return
        w = int(self.winfo_width() or 0)
        h = int(self.winfo_height() or 0)
        if w < 2 or h < 2:
            return
        aspect = w / float(h)
        try:
            glViewport(0, 0, w, h)
            glClear(GL_COLOR_BUFFER_BIT | GL_DEPTH_BUFFER_BIT)
        except GLError:
            # Context not current (e.g., during teardown / reparent); skip this tick
            return

        # Projection
        glMatrixMode(GL_PROJECTION)
        glLoadIdentity()
        gluPerspective(45.0, aspect, 0.05, 100.0)

        # ModelView
        glMatrixMode(GL_MODELVIEW)
        glLoadIdentity()
        glTranslatef(0, 0, -self.zoom)
        glRotatef(self.pitch, 1, 0, 0)
        glRotatef(self.yaw, 0, 1, 0)

        if self.wireframe:
            glDisable(GL_TEXTURE_2D)
            glPolygonMode(GL_FRONT_AND_BACK, GL_LINE)
        else:
            if self.texture_id is not None:
                glEnable(GL_TEXTURE_2D)
                glBindTexture(GL_TEXTURE_2D, self.texture_id)
            glPolygonMode(GL_FRONT_AND_BACK, GL_FILL)

        # Draw
        glBegin(GL_TRIANGLES)
        for tri in self.model.triangles:
            for (vx, vt, vn, m) in tri:
                # Compute (u,v) with runtime toggles
                u, v = vt[0], vt[1]
                if self.swap_uv:
                    u, v = v, u
                if self.flip_u:
                    u = 1.0 - u
                if self.flip_v:
                    v = 1.0 - v
                glTexCoord2f(u, v)
                glVertex3f(vx[0], vx[1], vx[2])
        glEnd()

        # Optional: overlay lines on top for clarity in textured mode
        if not self.wireframe:
            glDisable(GL_TEXTURE_2D)
            glPolygonMode(GL_FRONT_AND_BACK, GL_LINE)
            glLineWidth(1.0)
            glColor4f(1,1,1,0.2)
            glBegin(GL_TRIANGLES)
            for tri in self.model.triangles:
                for (vx, vt, vn, m) in tri:
                    glVertex3f(vx[0], vx[1], vx[2])
            glEnd()
            glColor4f(1,1,1,1)
        self._draw_overlays()

    # ---- Input ----
    def _on_drag_start(self, e):
        self._drag_last = (e.x, e.y)

    def _on_drag_move(self, e):
        if not self._drag_last:
            return
        dx = e.x - self._drag_last[0]
        dy = e.y - self._drag_last[1]
        self._drag_last = (e.x, e.y)
        self.yaw += dx * 0.5
        self.pitch += dy * 0.5
        self.pitch = max(-89.9, min(89.9, self.pitch))
        self.after(0, self.redraw)

    def _on_wheel(self, e):
        direction = 1 if e.delta > 0 else -1
        self._zoom_dir(direction)

    def _zoom_dir(self, direction):
        factor = 0.9 if direction > 0 else 1.1
        self.zoom = max(0.5, min(10.0, self.zoom * factor))
        self.after(0, self.redraw)

    def _on_key(self, e):
        if e.char.lower() == 'w':
            self.wireframe = not self.wireframe
        elif e.char.lower() == 'r':
            self.yaw, self.pitch, self.zoom = 30.0, -15.0, 2.5
        #elif e.char.lower() == 'v':
        #    self.flip_v = not self.flip_v
        #elif e.char.lower() == 'u':
        #    self.flip_u = not self.flip_u
        #elif e.char.lower() == 's':
        #    self.swap_uv = not self.swap_uv
        #elif e.char.lower() == 'c':
            # Toggle wrap mode at runtime
         #   self.wrap_clamp = not self.wrap_clamp
         #   if self.texture_id is not None:
         #       glBindTexture(GL_TEXTURE_2D, self.texture_id)
         #       glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE if self.wrap_clamp else GL_REPEAT)
         #       glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE if self.wrap_clamp else GL_REPEAT)
        self.after(0, self.redraw)

    def _on_destroy(self):
        # Prevent any further scheduled redraws from hitting a dead context
        self.animate = 0
        self._gl_ready = False

    # ---- Texture helper ----
    def _load_texture(self, path):
        # Load texture with PIL
        img = Image.open(path).convert("RGBA")
        img_data = img.tobytes("raw", "RGBA", 0, -1)
        width, height = img.size

        tex_id = glGenTextures(1)
        glBindTexture(GL_TEXTURE_2D, tex_id)

        glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_LINEAR_MIPMAP_LINEAR)
        glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_LINEAR)
        # Wrap mode applied in initgl() and toggleable at runtime with 'c'
        glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_CLAMP_TO_EDGE if self.wrap_clamp else GL_REPEAT)
        glTexParameteri(GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_CLAMP_TO_EDGE if self.wrap_clamp else GL_REPEAT)

        glTexImage2D(GL_TEXTURE_2D, 0, GL_RGBA, width, height, 0, GL_RGBA, GL_UNSIGNED_BYTE, img_data)
        glGenerateMipmap(GL_TEXTURE_2D)
        return tex_id

    def _draw_overlays(self):
        """Render beam/exhaust gizmos if provided.
        Draws simple lines: origin -> origin + dir * len.
        """
        data = self.overlay_provider() if callable(self.overlay_provider) else None
        if not data:
            return
        glDisable(GL_TEXTURE_2D)
        glLineWidth(2.0)
        # Map model-native coords -> normalized viewer space (matches normalized mesh)
        c = getattr(self.model, "center", (0.0, 0.0, 0.0))
        s = getattr(self.model, "scale", 1.0)
        def _norm_pt(pt):
            x, y, z = pt
            return ((x - c[0]) * s, (y - c[1]) * s, (z - c[2]) * s)

        # Beams in green
        if self.show_beams and data.get("beams"):
            glColor4f(0.2, 1.0, 0.2, 1.0)
            glBegin(GL_LINES)
            for g in data.get("beams", []):
                (x,y,z) = g.get("p", (0,0,0))
                (dx,dy,dz) = g.get("d", (0,0,1))
                L = float(g.get("len", 0.3))
                nx, ny, nz = _norm_pt((x, y, z))
                glVertex3f(nx, ny, nz)
                glVertex3f(nx + dx*L, ny + dy*L, nz + dz*L)
            glEnd()

        # Exhaust in orange
        if self.show_exhaust and data.get("exhaust"):
            glColor4f(1.0, 0.6, 0.2, 1.0)
            glBegin(GL_LINES)
            for g in data.get("exhaust", []):
                (x,y,z) = g.get("p", (0,0,0))
                (dx,dy,dz) = g.get("d", (0,0,1))
                L = float(g.get("len", 0.3))
                nx, ny, nz = _norm_pt((x, y, z))
                glVertex3f(nx, ny, nz)
                glVertex3f(nx + dx*L, ny + dy*L, nz + dz*L)
            glEnd()

        glColor4f(1,1,1,1)


# Convenience window to test quickly
def open_textured_viewer(obj_path, fallback_texture=None, title="Textured 3D Preview", size=(800,600)):
    root = tk.Toplevel() if tk._default_root is not None else tk.Tk()
    root.title(title)
    w,h = size
    root.geometry(f"{w}x{h}")
    frame = ObjTexturedGLFrame(root, obj_path=obj_path, fallback_texture=fallback_texture, width=w, height=h)
    frame.pack(fill="both", expand=True)
    frame.after(50, frame.redraw)
    root.mainloop() if isinstance(root, tk.Tk) else None
