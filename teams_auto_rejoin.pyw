"""
Microsoft Teams - Otomatik Yeniden Katılma
===========================================
Toplantıdan atılınca otomatik tekrar katılır.

✓ Fareye DOKUNMAZ
✓ Ekrana hiçbir şey GELMEZ
✓ Tam ekran oyun oynarken bile çalışır
✓ Sadece atılınca tıklar
✓ Toplantı penceresi otomatik minimize olur
✓ Arayüz ile kontrol

Kullanım:
    python teams_auto_rejoin.py
    İlk çalıştırmada gerekli paketler otomatik kurulur.
"""

# ── Otomatik Paket Kurulumu ──────────────────────────────
import subprocess
import sys

def install_packages():
    packages = {
        "pywinauto": "pywinauto",
        "win32gui": "pywin32",
    }
    missing = []
    for module, pip_name in packages.items():
        try:
            __import__(module)
        except ImportError:
            missing.append(pip_name)

    if missing:
        print(f"\n⏳ Gerekli paketler kuruluyor: {', '.join(missing)}")
        for pkg in missing:
            print(f"   Kuruluyor: {pkg}...")
            subprocess.check_call(
                [sys.executable, "-m", "pip", "install", pkg],
                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            )
        print("✓ Paketler kuruldu!\n")

install_packages()

# ── İçe Aktarmalar ───────────────────────────────────────
import time
import logging
import threading
import tkinter as tk
from tkinter import scrolledtext
from datetime import datetime
from pathlib import Path

from pywinauto import Desktop
import win32gui
import win32con

# ── Ayarlar ──────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).parent.resolve()
LOG_DIR = SCRIPT_DIR / "logs"

CHECK_INTERVAL = 3
BUTTON_TEXT = "Anında toplantı"
COOLDOWN = 10

LOG_DIR.mkdir(exist_ok=True)


# ── Pencere Takibi ───────────────────────────────────────
def get_teams_windows():
    windows = []
    def cb(hwnd, _):
        title = win32gui.GetWindowText(hwnd)
        if title and "microsoft teams" in title.lower():
            windows.append((hwnd, title))
    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass
    return windows


def has_meeting_window(windows):
    for hwnd, title in windows:
        t = title.lower()
        if "daraltılmış" in t or "daraltilmis" in t or "meeting" in t or "call" in t:
            return True
    return False


# ── Görünmez Tıklama ────────────────────────────────────
def invisible_click():
    try:
        import ctypes
        app = Desktop(backend="uia")
        wins = app.windows(title_re=".*Microsoft Teams.*")
        if not wins:
            return False

        teams, best = None, 0
        for w in wins:
            try:
                h = w.handle
                pl = win32gui.GetWindowPlacement(h)
                np = pl[4]
                s = (np[2] - np[0]) * (np[3] - np[1])
                if s > best:
                    best, teams = s, w
            except Exception:
                continue
        if not teams:
            teams = wins[0]

        hwnd = teams.handle
        placement = win32gui.GetWindowPlacement(hwnd)
        orig_pos = placement[4]
        w = max(orig_pos[2] - orig_pos[0], 1000)
        h = max(orig_pos[3] - orig_pos[1], 700)

        # Şeffaf yap
        orig_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE,
                               orig_style | win32con.WS_EX_LAYERED | win32con.WS_EX_TOOLWINDOW)
        try:
            ctypes.windll.user32.SetLayeredWindowAttributes(hwnd, 0, 0, 0x02)
        except Exception:
            pass

        # Ekran dışında restore
        off = (-5000, -5000, -5000 + w, -5000 + h)
        new_pl = list(placement)
        new_pl[1] = win32con.SW_SHOWNOACTIVATE
        new_pl[4] = off
        win32gui.SetWindowPlacement(hwnd, tuple(new_pl))
        time.sleep(0.5)

        # Buton tıkla
        success = False
        try:
            btns = teams.descendants(title=BUTTON_TEXT, control_type="Button")
            if btns:
                btns[0].invoke()
                success = True
            else:
                for item in teams.descendants():
                    if item.window_text() == BUTTON_TEXT:
                        item.invoke()
                        success = True
                        break
        except Exception:
            pass

        # Geri koy
        time.sleep(0.3)
        rpl = list(placement)
        rpl[1] = win32con.SW_SHOWMINIMIZED
        rpl[4] = orig_pos
        win32gui.SetWindowPlacement(hwnd, tuple(rpl))
        win32gui.SetWindowLong(hwnd, win32con.GWL_EXSTYLE, orig_style)
        try:
            ctypes.windll.user32.SetLayeredWindowAttributes(hwnd, 0, 255, 0x02)
        except Exception:
            pass

        return success
    except Exception:
        return False


# ── İzleme Motoru ────────────────────────────────────────
class TeamsMonitor:
    def __init__(self, gui):
        self.gui = gui
        self.running = False
        self.paused = False
        self.rejoin_count = 0
        self.last_click = 0
        self.had_meeting = False

    def log(self, msg, level="INFO"):
        ts = datetime.now().strftime("%H:%M:%S")
        self.gui.add_log(f"[{ts}] {msg}")

    def _minimize_meeting(self):
        """Toplantı penceresini arka plana at, Sohbet penceresini gizle."""
        try:
            for hwnd, title in get_teams_windows():
                t = title.lower()
                if "daraltılmış" in t or "daraltilmis" in t:
                    win32gui.SetWindowPos(hwnd, win32con.HWND_BOTTOM, 0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
                    self.log("✓ Toplantı penceresi arka plana atıldı")
        except Exception:
            pass
        self._hide_sohbet()

    def _hide_sohbet(self):
        """Sohbet penceresini arka plana at (görev çubuğunda kalır)."""
        try:
            for hwnd, title in get_teams_windows():
                if "sohbet" in title.lower() or "chat" in title.lower():
                    win32gui.SetWindowPos(hwnd, win32con.HWND_BOTTOM, 0, 0, 0, 0,
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | win32con.SWP_NOACTIVATE)
                    self.log("✓ Sohbet penceresi arka plana atıldı")
                    return
        except Exception:
            pass

    def _wait_and_minimize(self):
        self.log("Toplantı penceresi bekleniyor...")
        for _ in range(15):
            time.sleep(2)
            if has_meeting_window(get_teams_windows()):
                self.log("✓ Toplantı penceresi açıldı!")
                time.sleep(1)
                self._minimize_meeting()
                self.had_meeting = True
                return
        self.log("⚠ Toplantı penceresi 30sn içinde açılmadı")

    def start(self):
        self.running = True
        self.paused = False
        threading.Thread(target=self._loop, daemon=True).start()
        self.log("İzleme başladı")
        self.gui.set_status("Çalışıyor", "#00c853")

    def stop(self):
        self.running = False
        self.log(f"Durduruldu. Toplam katılma: {self.rejoin_count}")
        self.gui.set_status("Durduruldu", "#ff5252")

    def pause(self):
        self.paused = not self.paused
        if self.paused:
            self.log("⏸ Duraklatıldı")
            self.gui.set_status("Duraklatıldı", "#ffa726")
        else:
            self.log("▶ Devam ediyor")
            self.gui.set_status("Çalışıyor", "#00c853")

    def rejoin_now(self):
        self.log("Manuel katılma...")
        success = invisible_click()
        if success:
            self.rejoin_count += 1
            self.log(f"✓ Tıklandı! (toplam: {self.rejoin_count})")
            self.gui.update_count(self.rejoin_count)
            self._wait_and_minimize()
        else:
            self.log("✗ Buton bulunamadı")
        self.last_click = time.time()

    def leave_meeting(self):
        """Toplantıdan çık — daraltılmış penceredeki 'Ayrıl' butonuna tıkla."""
        self.log("Toplantıdan çıkılıyor...")
        try:
            app = Desktop(backend="uia")
            wins = app.windows(title_re=".*[Dd]aralt.*")
            if not wins:
                self.log("✗ Toplantı penceresi bulunamadı")
                return
            for w in wins:
                try:
                    btns = w.descendants(title="Ayrıl", control_type="Button")
                    if btns:
                        btns[0].invoke()
                        self.log("✓ Toplantıdan çıkıldı!")
                        self.gui.set_status("Toplantıdan çıkıldı", "#ffa726")
                        return
                except Exception:
                    pass
            self.log("✗ Ayrıl butonu bulunamadı")
        except Exception as e:
            self.log(f"Hata: {e}")

    def _loop(self):
        windows = get_teams_windows()
        self.had_meeting = has_meeting_window(windows)
        if self.had_meeting:
            self.log("✓ Toplantı penceresi mevcut")
        else:
            self.log("Toplantı penceresi yok — bekleniyor")

        while self.running:
            if self.paused:
                time.sleep(1)
                continue

            try:
                windows = get_teams_windows()
                has_mtg = has_meeting_window(windows)

                if self.had_meeting and not has_mtg:
                    if time.time() - self.last_click >= COOLDOWN:
                        self.log("⚠ TOPLANTIDAN ATILDINIZ!")
                        self.gui.set_status("Yeniden katılınıyor...", "#ffa726")

                        if invisible_click():
                            self.rejoin_count += 1
                            self.log(f"✓ Tıklandı! (toplam: {self.rejoin_count})")
                            self.gui.update_count(self.rejoin_count)
                            self._wait_and_minimize()
                            self.gui.set_status("Çalışıyor", "#00c853")
                        else:
                            self.log("✗ Katılma başarısız")
                            self.gui.set_status("Başarısız!", "#ff5252")
                        self.last_click = time.time()

                elif has_mtg and not self.had_meeting:
                    self.log("✓ Toplantı penceresi tekrar mevcut")
                    self._minimize_meeting()
                    self.gui.set_status("Çalışıyor", "#00c853")

                self.had_meeting = has_mtg

            except Exception as e:
                self.log(f"Hata: {e}")

            time.sleep(CHECK_INTERVAL)


# ── Arayüz (GUI) ────────────────────────────────────────
class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Teams Oto-Katılma")
        self.root.geometry("520x450")
        self.root.configure(bg="#1e1e2e")
        self.root.resizable(False, False)

        # Simge ayarla
        try:
            self.root.iconbitmap(default="")
        except Exception:
            pass

        self.monitor = None
        self._build_ui()

    def _build_ui(self):
        bg = "#1e1e2e"
        fg = "#cdd6f4"
        accent = "#7c3aed"

        # ── Başlık ──
        header = tk.Frame(self.root, bg=accent, height=60)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(header, text="⚡ Teams Oto-Katılma",
                 font=("Segoe UI", 16, "bold"), bg=accent, fg="white").pack(pady=12)

        # ── Durum Paneli ──
        status_frame = tk.Frame(self.root, bg="#2a2a3e", height=70)
        status_frame.pack(fill="x", padx=15, pady=(15, 5))
        status_frame.pack_propagate(False)

        left = tk.Frame(status_frame, bg="#2a2a3e")
        left.pack(side="left", padx=15, fill="y", expand=True)

        tk.Label(left, text="Durum:", font=("Segoe UI", 10),
                 bg="#2a2a3e", fg="#888").pack(anchor="w", pady=(10, 0))
        self.status_label = tk.Label(left, text="Hazır",
                                     font=("Segoe UI", 14, "bold"),
                                     bg="#2a2a3e", fg="#ffa726")
        self.status_label.pack(anchor="w")

        right = tk.Frame(status_frame, bg="#2a2a3e")
        right.pack(side="right", padx=15, fill="y", expand=True)

        tk.Label(right, text="Katılma Sayısı:", font=("Segoe UI", 10),
                 bg="#2a2a3e", fg="#888").pack(anchor="e", pady=(10, 0))
        self.count_label = tk.Label(right, text="0",
                                    font=("Segoe UI", 20, "bold"),
                                    bg="#2a2a3e", fg=accent)
        self.count_label.pack(anchor="e")

        # ── Butonlar ──
        btn_frame = tk.Frame(self.root, bg=bg)
        btn_frame.pack(fill="x", padx=15, pady=10)

        btn_style = {"font": ("Segoe UI", 11, "bold"), "relief": "flat",
                     "cursor": "hand2", "height": 1, "bd": 0}

        self.start_btn = tk.Button(btn_frame, text="▶  Başlat", bg="#00c853",
                                   fg="white", command=self._start, **btn_style)
        self.start_btn.pack(side="left", expand=True, fill="x", padx=(0, 3))

        self.pause_btn = tk.Button(btn_frame, text="⏸  Duraklat", bg="#ffa726",
                                   fg="white", command=self._pause,
                                   state="disabled", **btn_style)
        self.pause_btn.pack(side="left", expand=True, fill="x", padx=3)

        self.stop_btn = tk.Button(btn_frame, text="⏹  Durdur", bg="#ff5252",
                                  fg="white", command=self._stop,
                                  state="disabled", **btn_style)
        self.stop_btn.pack(side="left", expand=True, fill="x", padx=(3, 0))

        # ── Şimdi Katıl / Toplantıdan Çık ──
        action_frame = tk.Frame(self.root, bg=bg)
        action_frame.pack(fill="x", padx=15, pady=(0, 10))

        self.rejoin_btn = tk.Button(action_frame, text="🔄  Şimdi Katıl",
                                    bg=accent, fg="white",
                                    font=("Segoe UI", 11, "bold"),
                                    relief="flat", cursor="hand2",
                                    command=self._rejoin_now, state="disabled",
                                    height=1, bd=0)
        self.rejoin_btn.pack(side="left", expand=True, fill="x", padx=(0, 3))

        self.leave_btn = tk.Button(action_frame, text="📞  Toplantıdan Çık",
                                   bg="#c62828", fg="white",
                                   font=("Segoe UI", 11, "bold"),
                                   relief="flat", cursor="hand2",
                                   command=self._leave, state="disabled",
                                   height=1, bd=0)
        self.leave_btn.pack(side="left", expand=True, fill="x", padx=(3, 0))

        # ── Log Alanı ──
        log_label = tk.Label(self.root, text="📋 Olay Günlüğü",
                             font=("Segoe UI", 10), bg=bg, fg="#888",
                             anchor="w")
        log_label.pack(fill="x", padx=15)

        self.log_area = scrolledtext.ScrolledText(
            self.root, height=10,
            font=("Consolas", 9), bg="#11111b", fg="#a6adc8",
            insertbackground="#a6adc8", relief="flat", bd=0,
            state="disabled", wrap="word",
        )
        self.log_area.pack(fill="both", expand=True, padx=15, pady=(3, 15))

        # Tag renkleri
        self.log_area.tag_config("success", foreground="#00c853")
        self.log_area.tag_config("warning", foreground="#ffa726")
        self.log_area.tag_config("error", foreground="#ff5252")

    def add_log(self, msg):
        """Log mesajı ekle (thread-safe)."""
        def _add():
            self.log_area.configure(state="normal")
            tag = None
            if "✓" in msg:
                tag = "success"
            elif "⚠" in msg or "ATILDINIZ" in msg:
                tag = "warning"
            elif "✗" in msg or "Hata" in msg:
                tag = "error"
            self.log_area.insert("end", msg + "\n", tag)
            self.log_area.see("end")
            self.log_area.configure(state="disabled")
        self.root.after(0, _add)

    def set_status(self, text, color):
        def _set():
            self.status_label.configure(text=text, fg=color)
        self.root.after(0, _set)

    def update_count(self, count):
        def _set():
            self.count_label.configure(text=str(count))
        self.root.after(0, _set)

    def _start(self):
        if self.monitor and self.monitor.running:
            return
        self.monitor = TeamsMonitor(self)
        self.monitor.start()
        self.start_btn.configure(state="disabled")
        self.pause_btn.configure(state="normal")
        self.stop_btn.configure(state="normal")
        self.rejoin_btn.configure(state="normal")
        self.leave_btn.configure(state="normal")

    def _pause(self):
        if self.monitor:
            self.monitor.pause()
            if self.monitor.paused:
                self.pause_btn.configure(text="▶  Devam", bg="#00c853")
            else:
                self.pause_btn.configure(text="⏸  Duraklat", bg="#ffa726")

    def _stop(self):
        if self.monitor:
            self.monitor.stop()
        self.start_btn.configure(state="normal")
        self.pause_btn.configure(state="disabled", text="⏸  Duraklat", bg="#ffa726")
        self.stop_btn.configure(state="disabled")
        self.rejoin_btn.configure(state="disabled")
        self.leave_btn.configure(state="disabled")

    def _rejoin_now(self):
        if self.monitor:
            threading.Thread(target=self.monitor.rejoin_now, daemon=True).start()

    def _leave(self):
        if self.monitor:
            threading.Thread(target=self.monitor.leave_meeting, daemon=True).start()

    def run(self):
        self.add_log("[Başlat] butonuna basarak izlemeyi başlatın")
        windows = get_teams_windows()
        if windows:
            self.add_log(f"✓ {len(windows)} Teams penceresi bulundu")
        else:
            self.add_log("⚠ Teams penceresi bulunamadı — Teams açık olmalı")
        self.root.mainloop()


# ── Başlat ───────────────────────────────────────────────
if __name__ == "__main__":
    app = App()
    app.run()
