"""PDF → XLSX — перетворює PDF-таблиці (накази, додатки) в Excel з тією ж розміткою.

Перетягніть PDF у вікно (або на сам exe) — поруч з'явиться .xlsx з тим самим
ім'ям. Рушій: pdf2xlsx_engine.py.
"""

import ctypes
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog

from PIL import Image, ImageChops, ImageDraw, ImageFont, ImageTk

import pdf2xlsx_engine as engine

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except Exception:  # без drag&drop програма все одно працює
    TkinterDnD = None

APP_TITLE = (os.path.splitext(os.path.basename(sys.executable))[0]
             if getattr(sys, "frozen", False) else "PDF в XLSX")
FONT_DIR = os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts")
BG = "#2b2e33"
LOG_BG, LOG_FG = "#1e2024", "#e8eaed"
OK_FG, ERR_FG, DIM_FG = "#7fd67f", "#ff8a80", "#9aa0a6"


def resource_dir():
    return getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))


def read_brand():
    """Необов'язковий підпис у шапці: brand.txt поруч зі скриптом (вшивається в exe)."""
    try:
        with open(os.path.join(resource_dir(), "brand.txt"), encoding="utf-8-sig") as f:
            return f.read().strip()
    except OSError:
        return ""


BRAND = read_brand()


# ---------- «глянцевий» інтерфейс у стилі iOS 4 (як у «PDF Штамп») ----------

def load_font(size, *names):
    for n in names + ("segoeui.ttf", "arial.ttf"):
        try:
            return ImageFont.truetype(os.path.join(FONT_DIR, n), size)
        except OSError:
            pass
    return ImageFont.load_default(size)


def vgradient(w, h, stops):
    col = Image.new("RGBA", (1, max(1, h)))
    px = col.load()
    for y in range(max(1, h)):
        t = y / max(1, h - 1)
        for i in range(len(stops) - 1):
            (p0, c0), (p1, c1) = stops[i], stops[i + 1]
            if t <= p1 or i == len(stops) - 2:
                k = 0 if p1 == p0 else min(max((t - p0) / (p1 - p0), 0), 1)
                px[0, y] = tuple(round(a + (b - a) * k) for a, b in zip(c0, c1))
                break
    return col.resize((max(1, w), max(1, h)), Image.NEAREST)


def rmask(w, h, r):
    k = 4
    m = Image.new("L", (max(1, w) * k, max(1, h) * k), 0)
    ImageDraw.Draw(m).rounded_rectangle((0, 0, w * k - 1, h * k - 1), max(0, r) * k, fill=255)
    return m.resize((max(1, w), max(1, h)), Image.LANCZOS)


def masked(fill, mask):
    if not isinstance(fill, Image.Image):
        fill = Image.new("RGBA", mask.size, fill)
    out = fill.copy()
    out.putalpha(ImageChops.multiply(mask, fill.getchannel("A")))
    return out


def gray(v, a=255):
    return (v, v, v, a)


PAL_DARK = {"stops": [gray(112), gray(72), gray(50), gray(34)], "border": gray(8), "hi": gray(150, 200)}
PAL_DARK_HOVER = {"stops": [gray(128), gray(86), gray(62), gray(44)], "border": gray(8), "hi": gray(170, 210)}
PAL_PRESSED = {"stops": [gray(52), gray(40), gray(30), gray(24)], "border": gray(0), "hi": gray(70, 160)}
PAL_BLUE = {"stops": [(135, 175, 240, 255), (76, 128, 220, 255), (44, 100, 205, 255), (38, 88, 190, 255)],
            "border": (12, 30, 70, 255), "hi": (190, 215, 255, 220)}
PAL_BLUE_HOVER = {"stops": [(150, 188, 246, 255), (90, 142, 230, 255), (56, 112, 216, 255), (46, 98, 200, 255)],
                  "border": (12, 30, 70, 255), "hi": (205, 225, 255, 230)}


def glossy(w, h, r, pal, S):
    one = max(1, round(S))
    img = Image.new("RGBA", (w, h), (0, 0, 0, 0))
    img.alpha_composite(masked((255, 255, 255, 34), rmask(w, h, r)))
    img.alpha_composite(masked(pal["border"], rmask(w, h - one, r)))
    bw, bh = w - 2 * one, h - 3 * one
    s = pal["stops"]
    body = vgradient(bw, bh, [(0, s[0]), (0.5, s[1]), (0.5, s[2]), (1, s[3])])
    ImageDraw.Draw(body).line((0, 0, bw, 0), fill=pal["hi"], width=one)
    img.alpha_composite(masked(body, rmask(bw, bh, max(1, r - one))), (one, one))
    return img


def emboss(draw, xy, text, font, fill="white", shadow=(0, 0, 0, 170), dy=-1):
    x, y = xy
    draw.text((x, y + dy), text, font=font, fill=shadow)
    draw.text((x, y), text, font=font, fill=fill)


def on_bg(img, color=BG):
    """Кнопки — Label з картинкою: підкладаємо фон вікна під прозорі кути."""
    base = Image.new("RGBA", img.size, color)
    base.alpha_composite(img)
    return base


class GlossyButton:
    """Кнопка-картинка: текст намальований у самому зображенні, тож його не
    «з'їсть» зламана системна палітра (Windhawk), як було зі звичайними кнопками."""

    def __init__(self, parent, text, command, S, blue=False, width=None):
        self.command, self.enabled = command, True
        f = load_font(round(13 * S), "segoeuib.ttf")
        h = round(34 * S)
        w = width or round(f.getlength(text) + 40 * S)
        pals = (PAL_BLUE, PAL_BLUE_HOVER) if blue else (PAL_DARK, PAL_DARK_HOVER)
        self.imgs = {}
        for key, pal in (("n", pals[0]), ("h", pals[1]), ("p", PAL_PRESSED)):
            im = glossy(w, h, round(7 * S), pal, S)
            d = ImageDraw.Draw(im)
            tw = f.getlength(text)
            ty = h / 2 - sum(f.getbbox("Ag")[1::2]) / 2 - S
            emboss(d, ((w - tw) / 2, ty), text, f, dy=-max(1, round(S)))
            self.imgs[key] = ImageTk.PhotoImage(on_bg(im))
        dis = glossy(w, h, round(7 * S), PAL_PRESSED, S)
        ImageDraw.Draw(dis).text(((w - f.getlength(text)) / 2, h / 2 - sum(f.getbbox("Ag")[1::2]) / 2 - S),
                                 text, font=f, fill=gray(120))
        self.imgs["d"] = ImageTk.PhotoImage(on_bg(dis))
        self.w = tk.Label(parent, image=self.imgs["n"], bd=0, bg=BG, cursor="hand2")
        self.inside = False
        self.w.bind("<Enter>", lambda e: self._set(True, None))
        self.w.bind("<Leave>", lambda e: self._set(False, None))
        self.w.bind("<ButtonPress-1>", lambda e: self.enabled and self.w.configure(image=self.imgs["p"]))
        self.w.bind("<ButtonRelease-1>", self._release)

    def _set(self, inside, _):
        self.inside = inside
        if self.enabled:
            self.w.configure(image=self.imgs["h" if inside else "n"])

    def _release(self, e):
        if not self.enabled:
            return
        self.w.configure(image=self.imgs["h" if self.inside else "n"])
        if self.inside:
            self.command()

    def set_enabled(self, on):
        self.enabled = on
        self.w.configure(image=self.imgs["n" if on else "d"], cursor="hand2" if on else "arrow")


# ------------------------------------------------------------------ вікно

class App:
    def __init__(self, root, S):
        self.root, self.S = root, S
        self.jobs = queue.Queue()
        self.events = queue.Queue()
        self.busy = False
        self.last_dir = None
        self.hover_drop = False

        root.title(APP_TITLE)
        root.configure(bg=BG)
        root.geometry("%dx%d" % (round(600 * S), round(520 * S)))
        root.minsize(round(460 * S), round(400 * S))

        self.header = tk.Canvas(root, height=round(64 * S), highlightthickness=0, bd=0, bg="#222")
        self.header.pack(fill=tk.X)
        self.header.bind("<Configure>", lambda e: self.draw_header())

        body = tk.Frame(root, bg=BG)
        body.pack(fill=tk.BOTH, expand=True, padx=round(16 * S), pady=(round(14 * S), round(10 * S)))

        self.drop = tk.Canvas(body, height=round(150 * S), highlightthickness=0, bd=0, bg=BG, cursor="hand2")
        self.drop.pack(fill=tk.X)
        self.drop.bind("<Configure>", lambda e: self.draw_drop())
        self.drop.bind("<Button-1>", lambda e: self.pick_files())
        self.drop.bind("<Enter>", lambda e: self._drop_hover(True))
        self.drop.bind("<Leave>", lambda e: self._drop_hover(False))

        self.log = tk.Text(body, height=5, bg=LOG_BG, fg=LOG_FG, bd=0, highlightthickness=1,
                           highlightbackground="#111", relief=tk.FLAT, wrap=tk.WORD,
                           font=("Segoe UI", 10), padx=round(10 * S), pady=round(8 * S),
                           cursor="arrow", state=tk.DISABLED)
        for tag, col in (("ok", OK_FG), ("err", ERR_FG), ("dim", DIM_FG)):
            self.log.tag_configure(tag, foreground=col)

        bar = tk.Frame(body, bg=BG)
        bar.pack(fill=tk.X)
        self.btn_pick = GlossyButton(bar, "Вибрати PDF…", self.pick_files, S, blue=True)
        self.btn_pick.w.pack(side=tk.LEFT)
        self.btn_folder = GlossyButton(bar, "Відкрити теку", self.open_folder, S)
        self.btn_folder.w.pack(side=tk.RIGHT)
        self.btn_folder.set_enabled(False)
        bar.pack_forget()
        bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.log.pack(fill=tk.BOTH, expand=True, pady=(round(12 * S), round(10 * S)))

        self.say("Перетягніть PDF у вікно — поруч з'явиться .xlsx з тим самим ім'ям "
                 "і тією ж розміткою. Можна кілька файлів або цілу теку.", "dim")

        if TkinterDnD is not None:
            for w in (root, self.drop, self.log):
                try:
                    w.drop_target_register(DND_FILES)
                    w.dnd_bind("<<Drop>>", self.on_drop)
                    w.dnd_bind("<<DropEnter>>", lambda e: self._drop_hover(True) or e.action)
                    w.dnd_bind("<<DropLeave>>", lambda e: self._drop_hover(False) or e.action)
                except Exception:
                    pass

        threading.Thread(target=self.worker, daemon=True).start()
        root.after(100, self.poll)

    # --- малювання ---
    def draw_header(self):
        S = self.S
        W, H = max(self.header.winfo_width(), 200), round(64 * S)
        one = max(1, round(S))
        img = vgradient(W, H, [(0, gray(92)), (0.5, gray(56)), (0.5, gray(40)), (1, gray(22))])
        d = ImageDraw.Draw(img)
        d.line((0, 0, W, 0), fill=gray(140), width=one)
        d.line((0, H - one, W, H - one), fill=gray(0), width=one)
        f1 = load_font(round(19 * S), "segoeuib.ttf")
        f2 = load_font(round(11 * S), "segoeui.ttf")
        x = round(18 * S)
        emboss(d, (x, round(9 * S)), "Конвертер PDF в XLSX", f1, dy=-one)
        sub = "розмітка один-в-один" + ("  ·  " + BRAND if BRAND else "")
        emboss(d, (x + one, round(38 * S)), sub, f2, fill=gray(190), dy=-one)
        self.header_img = ImageTk.PhotoImage(img)
        self.header.delete("all")
        self.header.create_image(0, 0, image=self.header_img, anchor=tk.NW)

    def draw_drop(self):
        S = self.S
        W, H = max(self.drop.winfo_width(), 200), round(150 * S)
        r = round(14 * S)
        hot = self.hover_drop and not self.busy
        top, bot = ((62, 70, 84), (44, 50, 60)) if hot else ((52, 55, 61), (38, 40, 45))
        img = Image.new("RGBA", (W, H), BG)
        body = vgradient(W, H, [(0, top + (255,)), (1, bot + (255,))])
        img.alpha_composite(masked(body, rmask(W, H, r)))
        # пунктирна рамка
        d = ImageDraw.Draw(img)
        col = (120, 165, 240, 255) if hot else gray(105)
        inset, dash, gap = round(6 * S), round(10 * S), round(7 * S)
        w = max(1, round(2 * S))
        x0, y0, x1, y1 = inset, inset, W - inset, H - inset
        for xa in range(x0 + r, x1 - r, dash + gap):
            xb = min(xa + dash, x1 - r)
            d.line((xa, y0, xb, y0), fill=col, width=w)
            d.line((xa, y1, xb, y1), fill=col, width=w)
        for ya in range(y0 + r, y1 - r, dash + gap):
            yb = min(ya + dash, y1 - r)
            d.line((x0, ya, x0, yb), fill=col, width=w)
            d.line((x1, ya, x1, yb), fill=col, width=w)
        for (ax, ay, a0) in ((x0, y0, 180), (x1 - 2 * r, y0, 270), (x0, y1 - 2 * r, 90), (x1 - 2 * r, y1 - 2 * r, 0)):
            d.arc((ax, ay, ax + 2 * r, ay + 2 * r), a0, a0 + 90, fill=col, width=w)
        # піктограма: аркуш PDF → таблиця
        cx = W / 2
        fi = load_font(round(30 * S), "seguiemj.ttf", "segoeui.ttf")
        f1 = load_font(round(15 * S), "segoeuib.ttf")
        f2 = load_font(round(11 * S), "segoeui.ttf")
        if self.busy:
            t1, t2, icon = "Перетворюю…", "зачекайте, файли в черзі", "\u23F3"
        else:
            t1, t2, icon = "Перетягніть PDF сюди", "або натисніть, щоб вибрати файли", "\U0001F4C4 \u2192 \U0001F4CA"
        iw = d.textlength(icon, font=fi)
        d.text((cx - iw / 2, round(22 * S)), icon, font=fi, embedded_color=True, fill="white")
        tw = d.textlength(t1, font=f1)
        d.text((cx - tw / 2, round(74 * S)), t1, font=f1, fill="white")
        tw = d.textlength(t2, font=f2)
        d.text((cx - tw / 2, round(100 * S)), t2, font=f2, fill=gray(170))
        self.drop_img = ImageTk.PhotoImage(img)
        self.drop.delete("all")
        self.drop.create_image(0, 0, image=self.drop_img, anchor=tk.NW)

    def _drop_hover(self, on):
        self.hover_drop = on
        self.draw_drop()

    # --- журнал ---
    def say(self, text, tag=None, mark=None):
        self.log.configure(state=tk.NORMAL)
        if self.log.index("end-1c") != "1.0":
            self.log.insert(tk.END, "\n")
        if mark:
            self.log.mark_set(mark, "end-1c")
            self.log.mark_gravity(mark, tk.LEFT)
        self.log.insert(tk.END, text, tag or ())
        self.log.see(tk.END)
        self.log.configure(state=tk.DISABLED)

    # --- вхідні файли ---
    def pick_files(self):
        if self.busy:
            return
        files = filedialog.askopenfilenames(
            parent=self.root, title="Виберіть PDF",
            filetypes=[("PDF", "*.pdf"), ("Усі файли", "*.*")],
            initialdir=self.last_dir or os.path.expanduser("~"))
        self.add(files)

    def on_drop(self, e):
        self._drop_hover(False)
        self.add(self.root.tk.splitlist(e.data))
        return e.action

    def add(self, paths):
        pdfs = []
        for p in paths:
            if os.path.isdir(p):
                pdfs += sorted(os.path.join(p, f) for f in os.listdir(p) if f.lower().endswith(".pdf"))
            elif p.lower().endswith(".pdf"):
                pdfs.append(p)
            elif p:
                self.say("✖ %s — це не PDF, пропускаю" % os.path.basename(p), "err")
        for p in pdfs:
            self.jobs.put(os.path.normpath(p))
        if pdfs:
            self.busy = True
            self.btn_pick.set_enabled(False)
            self.draw_drop()

    # --- фонова робота ---
    def worker(self):
        while True:
            pdf = self.jobs.get()
            self.events.put(("start", pdf, None))
            try:
                out, pages = engine.convert_file(pdf)
                self.events.put(("ok", pdf, (out, pages)))
            except engine.NotConvertible as e:
                self.events.put(("err", pdf, str(e)))
            except PermissionError:
                self.events.put(("err", pdf, "немає доступу до теки або файл відкритий в іншій програмі"))
            except Exception as e:
                self.events.put(("err", pdf, "помилка: %s" % e))
            self.events.put(("done", None, self.jobs.empty()))

    def poll(self):
        try:
            while True:
                kind, pdf, data = self.events.get_nowait()
                name = os.path.basename(pdf) if pdf else ""
                if kind == "start":
                    self.say("…  %s" % name, "dim", mark="pending")
                elif kind in ("ok", "err") and "pending" in self.log.mark_names():
                    self.log.configure(state=tk.NORMAL)
                    self.log.delete("pending linestart", "pending lineend +1c")
                    self.log.mark_unset("pending")
                    if self.log.get("end-2c") == "\n":
                        self.log.delete("end-2c")
                    self.log.configure(state=tk.DISABLED)
                if kind == "ok":
                    out, pages = data
                    self.say("✔ %s  →  %s  (%d стор.)" % (name, os.path.basename(out), pages), "ok")
                    self.last_dir = os.path.dirname(out)
                    self.last_out = out
                    self.btn_folder.set_enabled(True)
                elif kind == "err":
                    self.say("✖ %s — %s" % (name, data), "err")
                elif kind == "done" and data:
                    self.busy = False
                    self.btn_pick.set_enabled(True)
                    self.draw_drop()
        except queue.Empty:
            pass
        self.root.after(100, self.poll)

    def open_folder(self):
        out = getattr(self, "last_out", None)
        if out and os.path.exists(out):
            subprocess.Popen('explorer /select,"%s"' % out)
        elif self.last_dir:
            os.startfile(self.last_dir)


def main():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    root = None
    if TkinterDnD is not None:
        try:
            root = TkinterDnD.Tk()
        except Exception:
            root = None
    if root is None:
        root = tk.Tk()
    try:
        root.iconbitmap(os.path.join(resource_dir(), "icon.ico"))
    except Exception:
        pass
    try:
        S = root.winfo_fpixels("1i") / 96
    except Exception:
        S = 1.0
    app = App(root, S)
    if len(sys.argv) > 1:          # файли, кинуті на сам exe / «Відкрити за допомогою»
        root.after(300, lambda: app.add(sys.argv[1:]))
    root.mainloop()


if __name__ == "__main__":
    main()
