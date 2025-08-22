import os, time, random, threading, ctypes, re
from ctypes import wintypes
import tkinter as tk
from tkinter import ttk, messagebox

import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import JavascriptException, WebDriverException
from webdriver_manager.chrome import ChromeDriverManager

# ===================== Config =====================
URL_DEFAULT = "https://www.edutyping.com/student/lesson/110170"
WPM_DEFAULT = 42
LOAD_TIMEOUT = 40
# Hotkeys (global, Windows only)
MOD_SHIFT = 0x0004
VK_F11 = 0x7A
VK_F12 = 0x7B
WM_HOTKEY = 0x0312

# ================= Selenium helpers =================

def driver_setup(url):
    opts = webdriver.ChromeOptions()
    opts.add_argument("--start-maximized")
    opts.add_experimental_option("excludeSwitches", ["enable-automation"]) 
    opts.add_experimental_option("useAutomationExtension", False)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {"source": """
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        try { window.focus(); } catch(e){}
    """})
    driver.get(url)
    return driver


def smart_wait(driver, secs=LOAD_TIMEOUT):
    t0 = time.time()
    while time.time() - t0 < secs:
        try:
            if driver.execute_script("return document.readyState") == "complete":
                return
        except WebDriverException:
            pass
        time.sleep(0.25)


def focus_typing_area(driver):
    js_focus = """
    (function(){
      const centerClick = () => {
        const x = Math.floor(window.innerWidth/2), y = Math.floor(window.innerHeight*0.28);
        const t = document.elementFromPoint(x,y);
        if (t) {
          t.dispatchEvent(new MouseEvent('mousedown',{bubbles:true,clientX:x,clientY:y}));
          t.dispatchEvent(new MouseEvent('mouseup',{bubbles:true,clientX:x,clientY:y}));
          t.dispatchEvent(new MouseEvent('click',{bubbles:true,clientX:x,clientY:y}));
        }
      };
      window.focus();
      if (document.body) { document.body.tabIndex = -1; document.body.focus(); }
      centerClick();
      return document.hasFocus();
    })();
    """
    ok = False
    for _ in range(4):
        try:
            ok = bool(driver.execute_script(js_focus))
        except JavascriptException:
            ok = False
        if ok:
            break
        time.sleep(0.3)
    if not ok:
        ActionChains(driver).move_by_offset(5, 5).click().perform()


def send_keys_hard(driver, ch):
    try:
        ActionChains(driver).send_keys(ch).perform()
        return
    except WebDriverException:
        pass
    js = """
      (function(c){
        function fire(type){
          const e = new KeyboardEvent(type, {key:c, bubbles:true});
          document.dispatchEvent(e);
        }
        fire('keydown'); fire('keypress'); fire('keyup');
      })(arguments[0]);
    """
    try:
        driver.execute_script(js, ch)
    except JavascriptException:
        pass


def human_type(driver, text, wpm=WPM_DEFAULT, status_cb=None):
    cps = (wpm * 5) / 60.0
    base_delay = 1.0 / max(0.001, cps)
    for i, ch in enumerate(text):
        send_keys_hard(driver, ch)
        if i % 25 == 0 and status_cb:
            status_cb(f"Typed {i}/{len(text)} chars… (Esc in app to cancel)")
        if random.random() < 0.02:
            focus_typing_area(driver)
        jitter = random.uniform(-0.35, 0.25) * base_delay
        time.sleep(max(0.01, base_delay + jitter))

def normalize_text(s: str) -> str:
    s = s.replace("
", "
").replace("
", "
")
    s = "".join(ch for ch in s if ch not in "​‌‍")
    lines = s.split("
")
    non_empty = [ln for ln in lines if ln.strip()]
    if non_empty:
        single_count = sum(1 for ln in non_empty if len(ln.strip()) == 1)
        if single_count / len(non_empty) >= 0.6:
            words, buf = [], []
            for ln in lines:
                t = ln.strip()
                if not t:
                    if buf:
                        words.append("".join(buf))
                        buf = []
                else:
                    buf.append(t)
            if buf:
                words.append("".join(buf))
            s = " ".join(words)
        else:
            s = re.sub(r"[ 	]*
[ 	]*", " ", s)
    s = re.sub(r" {2,}", " ", s)
    return s.strip()

# ================= Tk App =================
class TypingApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("EduTyping Auto Typer")
        self.geometry("720x420")
        self.minsize(680, 380)

        # Selenium driver
        self.driver = None
        self.armed = False
        self.hotkey_thread = None
        self.stop_flag = False

        # UI
        self.url_var = tk.StringVar(value=URL_DEFAULT)
        self.wpm_var = tk.IntVar(value=WPM_DEFAULT)
        self.status_var = tk.StringVar(value="1) Click 'Open Lesson' and log in. 2) Click inside typing box.")

        top = ttk.Frame(self, padding=12)
        top.pack(fill=tk.BOTH, expand=True)

        row1 = ttk.Frame(top)
        row1.pack(fill=tk.X, pady=(0,8))
        ttk.Label(row1, text="Lesson URL:").pack(side=tk.LEFT)
        ttk.Entry(row1, textvariable=self.url_var).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=6)
        ttk.Button(row1, text="Open Lesson", command=self.open_lesson).pack(side=tk.LEFT)

        row2 = ttk.Frame(top)
        row2.pack(fill=tk.X, pady=(0,8))
        ttk.Label(row2, text="WPM:").pack(side=tk.LEFT)
        ttk.Spinbox(row2, from_=10, to=180, textvariable=self.wpm_var, width=6).pack(side=tk.LEFT, padx=6)
        ttk.Button(row2, text="Arm (Shift+F11)", command=self.arm_hotkey).pack(side=tk.LEFT, padx=4)
        ttk.Button(row2, text="Fix Text", command=self.fix_text).pack(side=tk.LEFT, padx=4)
        ttk.Button(row2, text="Type Now", command=self.type_now).pack(side=tk.LEFT)
        ttk.Button(row2, text="Stop", command=self.stop_typing).pack(side=tk.LEFT, padx=4)

        ttk.Label(top, text="Text to type:").pack(anchor="w")
        self.txt = tk.Text(top, height=10, wrap="word")
        self.txt.pack(fill=tk.BOTH, expand=True)

        self.status = ttk.Label(top, textvariable=self.status_var)
        self.status.pack(anchor="w", pady=(8,0))

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    # ----- actions -----
    def set_status(self, s):
        self.status_var.set(s)
        self.update_idletasks()

    def open_lesson(self):
        try:
            if not self.driver:
                self.driver = driver_setup(self.url_var.get().strip() or URL_DEFAULT)
                smart_wait(self.driver)
            else:
                self.driver.get(self.url_var.get().strip() or URL_DEFAULT)
            self.set_status("Lesson opened. Log in if needed, then click inside the typing area.")
        except Exception as e:
            messagebox.showerror("Error opening lesson", str(e))

    def clean_text(self, raw_text):
        # Remove linebreaks between single letters, join into words
        lines = raw_text.splitlines()
        if all(len(line.strip()) <= 2 for line in lines if line.strip()):
            # Looks like one char per line → join
            return re.sub(r"\s+", "", raw_text)
        # Otherwise collapse multiple spaces/newlines into single space
        return re.sub(r"\s+", " ", raw_text)

    def type_now(self):
        if not self.driver:
            messagebox.showwarning("Open Lesson", "Open the lesson first.")
            return
        raw = self.txt.get("1.0", tk.END)
        text = normalize_text(raw)
        if not text:
            self.set_status("No text provided; not typing.")
            return
        self.txt.delete("1.0", tk.END)
        self.txt.insert("1.0", text)
        self.stop_flag = False
        focus_typing_area(self.driver)
        wpm = max(10, int(self.wpm_var.get()))
        self.set_status(f"Typing {len(text)} chars at ~{wpm} WPM…")
        threading.Thread(target=self._type_thread, args=(text, wpm), daemon=True).start()

    def _type_thread(self, text, wpm):
        try:
            def cb(msg):
                self.status_var.set(msg)
            human_type(self.driver, text, wpm, status_cb=cb)
            self.set_status("Done typing. Browser stays open.")
        except Exception as e:
            self.set_status(f"Typing error: {e}")
    def fix_text(self):
        raw = self.txt.get("1.0", tk.END)
        fixed = normalize_text(raw)
        self.txt.delete("1.0", tk.END)
        self.txt.insert("1.0", fixed)
        self.set_status("Text normalized.")

    def stop_typing(self):
        self.stop_flag = True
        self.set_status("Stop requested. You can close the browser manually if needed.")

    # ----- hotkey -----
    def arm_hotkey(self):
        if self.hotkey_thread and self.hotkey_thread.is_alive():
            self.set_status("Hotkey already armed. Press Shift+F11 to trigger or Shift+F12 to cancel.")
            return
        self.armed = True
        self.set_status("Hotkey armed: press Shift+F11 to type current text. Shift+F12 to cancel.")
        self.hotkey_thread = threading.Thread(target=self._hotkey_loop, daemon=True)
        self.hotkey_thread.start()

    def _hotkey_loop(self):
        user32 = ctypes.windll.user32
        user32.RegisterHotKey(None, 1, MOD_SHIFT, VK_F11)
        user32.RegisterHotKey(None, 2, MOD_SHIFT, VK_F12)
        msg = wintypes.MSG()
        try:
            while self.armed:
                if user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
                    if msg.message == WM_HOTKEY:
                        if msg.wParam == 1:
                            self.after(0, self.type_now)
                        elif msg.wParam == 2:
                            self.after(0, self.stop_typing)
        finally:
            user32.UnregisterHotKey(None, 1)
            user32.UnregisterHotKey(None, 2)

    def on_close(self):
        self.armed = False
        try:
            if self.driver:
                pass
        except Exception:
            pass
        self.destroy()


if __name__ == "__main__":
    app = TypingApp()
    app.mainloop()
