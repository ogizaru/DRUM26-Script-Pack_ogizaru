#!/usr/bin/env python
# ============================================================================
# 自動インストール機能 (ドラッグ＆ドロップ等での実行時)
# ============================================================================
def __install_this_script__():
    import os, shutil, sys, inspect
    
    log_path = os.path.expanduser("~/resolve_install_debug.txt")
    def log(msg):
        try:
            with open(log_path, "a") as f:
                f.write(str(msg) + "\n")
        except: pass

    log("=== Python Installer Debug Start (Post_Importer) ===")

    def get_current_path():
        try:
            if "__file__" in globals():
                p = os.path.abspath(globals()["__file__"])
                log("__file__ in globals: " + str(p))
                return p
            p = os.path.abspath(__file__)
            log("__file__ in locals: " + str(p))
            return p
        except Exception as e:
             log("__file__ error: " + str(e))
             
        try:
            for f in inspect.stack():
                p = f[1]
                log("Inspect Frame: " + str(p))
                if p and os.path.exists(p) and p.endswith(".py"):
                    return os.path.abspath(p)
        except Exception as e:
             log("inspect error: " + str(e))
             
        gl = globals()
        for k, v in gl.items():
            if type(v) == str and os.path.exists(v) and v.endswith(".py"):
                log("Found in globals: " + k + " = " + v)
                return os.path.abspath(v)
        return None

    current_path = get_current_path()
    log("current_path: " + str(current_path))
    if not current_path: return True 
        
    current_path = current_path.replace("\\", "/")
    filename = os.path.basename(current_path)
    log("filename: " + filename)
    
    def get_bmd_module():
        import sys
        try: import DaVinciResolveScript as bmd; return bmd
        except ImportError: return sys.modules.get("fusionscript")

    bmd = get_bmd_module()
    log("bmd module found: " + str(bmd is not None))
    
    fuset = None
    try:
         if 'fu' in globals(): fuset = fu; log("fuset from globals")
         elif 'resolve' in globals(): fuset = resolve.GetFusion(); log("fuset from resolve")
         elif bmd: fuset = bmd.scriptapp("Resolve").GetFusion(); log("fuset from bmd scriptapp")
    except Exception as e:
         log("fuset error: " + str(e))
         return True
    
    log("fuset found: " + str(fuset is not None))
    if not fuset: return True
    
    import platform as py_platform
    is_win = (py_platform.system() == "Windows")
    log("is_win: " + str(is_win))
    all_dir = None

    if is_win:
        appdata = os.environ.get("APPDATA")
        if not appdata: return True 
        all_dir = os.path.join(appdata, "Blackmagic Design", "DaVinci Resolve", "Support", "Fusion", "Scripts", "Utility")
    else:
        all_dir = os.path.expanduser("~/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Utility/")

    log("all_dir: " + str(all_dir))

    # 【判定強化】パスの正規化比較
    src_n = current_path.replace("\\", "/").lower()
    a_n = all_dir.replace("\\", "/").lower() + "/" + filename.lower()
    log("src_n: " + src_n)
    log("a_n: " + a_n)

    if src_n == a_n:
        log("Already installed")
        return True # すでにインストール済み

    ui = getattr(fuset, "UIManager", None)
    log("ui found: " + str(ui is not None))
    if not ui: return True
    dispatcher = bmd.UIDispatcher(ui) if bmd and hasattr(bmd, "UIDispatcher") else None
    log("dispatcher found: " + str(dispatcher is not None))
    if not dispatcher: return True

    log("Creating Window...")

    win = dispatcher.AddWindow(
        {"ID": "InstallWin", "WindowTitle": "インストールの確認", "Geometry": [400, 200, 450, 260]},
        ui.VGroup([
            ui.Label({"Text": "以下の内容でインストールを実行します：", "Weight": 0}),
            ui.VGroup([
                ui.Label({"Text": "■ スクリプト名: " + filename, "Weight": 0}),
                ui.VGap(2),
                ui.Label({"Text": "■ インストール先: ", "Weight": 0}),
                ui.TextEdit({"Text": all_dir, "ReadOnly": True, "Weight": 1}),
                ui.VGap(4),
                ui.Label({"Text": "■ 自動セットアップされるライブラリ: ", "Weight": 0}),
                ui.Label({"Text": "   - Playwright (SNSブラウザ操作)", "Weight": 0}),
                ui.Label({"Text": "   - 専用ブラウザバイナリ", "Weight": 0}),
            ]),
            ui.HGroup({"Weight": 0}, [
                ui.Button({"ID": "BtnOk", "Text": "インストール"}),
                ui.Button({"ID": "BtnCancel", "Text": "キャンセル"})
            ])
        ])
    )
    
    result = {"install": False}
    def on_ok(ev): result["install"] = True; log("BtnOk Clicked"); dispatcher.ExitLoop()
    def on_close(ev): log("Window Closed"); dispatcher.ExitLoop()
        
    win.On.BtnOk.Clicked = on_ok
    win.On.BtnCancel.Clicked = on_close
    win.On.InstallWin.Close = on_close
    win.Show(); dispatcher.RunLoop(); win.Hide()
    
    if result["install"]:
        target_dir = all_dir
        target_path = os.path.join(target_dir, filename)
        log("Target path for install: " + target_path)
        if not os.path.exists(target_dir):
            try: os.makedirs(target_dir); log("Directory created")
            except Exception as e: log("makedirs error: " + str(e))
        try:
            shutil.copy(current_path, target_path)
            log("File copied")
            
            # --- 【追加】自動ライブラリインストール ---
            log("Installing requirements (playwright)...")
            try:
                import subprocess
                # 1. pip install playwright (--user を付けて安全にインストール)
                subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", "playwright"])
                # 2. playwright install
                subprocess.check_call([sys.executable, "-m", "playwright", "install"])
                log("Playwright requirements installed successfully")
            except Exception as e:
                log("Failed to install requirements: " + str(e))
            
            cwin = dispatcher.AddWindow(
                {"ID": "CWin", "WindowTitle": "完了", "Geometry": [400, 200, 250, 100]},
                ui.VGroup([
                    ui.Label({"Text": "インストールが完了しました。", "WordWrap": True, "Alignment": {"AlignHCenter": True, "AlignVCenter": True}, "Weight": 1}),
                    ui.Button({"ID": "BtnOk", "Text": "OK", "Weight": 0})
                ])
            )
            cwin.On.BtnOk.Clicked = on_close; cwin.Show(); dispatcher.RunLoop(); cwin.Hide()
            return False
        except Exception as e:
            log("Install Error: " + str(e))
            ewin = dispatcher.AddWindow(
                {"ID": "EWin", "WindowTitle": "エラー", "Geometry": [400, 200, 350, 120]},
                ui.VGroup([
                    ui.Label({"Text": "失敗しました。ProgramData の場合は管理者権限が必要です。", "WordWrap": True, "Alignment": {"AlignHCenter": True, "AlignVCenter": True}, "Weight": 1}),
                    ui.Button({"ID": "BtnOk", "Text": "OK"})
                ])
            )
            ewin.On.BtnOk.Clicked = on_close; ewin.Show(); dispatcher.RunLoop(); ewin.Hide()
    return True

if not __install_this_script__():
    import sys; sys.exit()
# ============================================================================
"""
Resolve X Importer v0.4.0 beta
Developed by DaVinci Resolve Addon and DCTL maker V3 & OGIZARU

Requirements:
- DaVinci Resolve Studio or Free (v16.2+)
- Python 3.1+
- libraries: playwright
  (Install via: pip install playwright)

Update v0.4.0 beta:
- Changed default save location from Temp to user's "Pictures/Resolve_X_Imports" to prevent Media Offline.
- Added "Browse" button to select custom save directory.
"""

import sys
import os
import time
import re
import subprocess
import platform

# --- Resolve API Path Setup ---
def setup_resolve_env():
    system_platform = sys.platform
    path_to_add = ""
    if system_platform.startswith("win"):
        path_to_add = os.path.expandvars(r"%PROGRAMDATA%\Blackmagic Design\DaVinci Resolve\Support\Developer\Scripting\Modules")
    elif system_platform.startswith("darwin"):
        path_to_add = "/Library/Application Support/Blackmagic Design/DaVinci Resolve/Developer/Scripting/Modules"
    elif system_platform.startswith("linux"):
        path_to_add = "/opt/resolve/Developer/Scripting/Modules"

    if path_to_add and os.path.exists(path_to_add):
        if path_to_add not in sys.path:
            sys.path.append(path_to_add)

setup_resolve_env()

# --- Resolve API Import ---
dvr_script = None
resolve = None
fusion = None

try:
    import DaVinciResolveScript as dvr_script
    resolve = dvr_script.scriptapp("Resolve")
    if resolve:
        fusion = resolve.Fusion()
except ImportError:
    try:
        resolve = resolve
        fusion = resolve.Fusion()
    except NameError:
        pass

# --- Playwright Setup ---
try:
    from playwright.sync_api import sync_playwright, Error as PlaywrightError
    PLAYWRIGHT_AVAILABLE = True
except ImportError:
    PLAYWRIGHT_AVAILABLE = False
    PlaywrightError = Exception

# --- Helper Functions ---

def install_playwright_browsers():
    print("Launching Playwright installer...")
    if sys.platform.startswith("win"):
        subprocess.Popen("start cmd /k playwright install", shell=True)
    elif sys.platform.startswith("darwin"):
        subprocess.Popen(["osascript", "-e", 'tell application "Terminal" to do script "playwright install"'])
    else:
        print("Please run 'playwright install' in your terminal.")

def get_default_save_path():
    """Returns a persistent path: ~/Pictures/Resolve_X_Imports"""
    home = os.path.expanduser("~")
    # Windows/Mac compatible 'Pictures' folder
    pictures_dir = os.path.join(home, "Pictures")
    if not os.path.exists(pictures_dir):
        # Fallback to home if Pictures doesn't exist
        pictures_dir = home
    
    target_dir = os.path.join(pictures_dir, "Resolve_X_Imports")
    if not os.path.exists(target_dir):
        try:
            os.makedirs(target_dir)
        except OSError:
            return home # Fallback
    return target_dir

# --- Class Definitions ---

class XCapture:
    """X (Twitter) Post Capturing Logic"""
    def __init__(self):
        pass # No longer using temp_dir in init
        
    def validate_url(self, url):
        if not url: return False
        pattern = r"https?://(www\.)?(twitter|x)\.com/[a-zA-Z0-9_]+/status/[0-9]+"
        return re.match(pattern, url) is not None

    def capture_post(self, url, save_dir, theme_mode="light"):
        if not PLAYWRIGHT_AVAILABLE:
            raise RuntimeError("Playwright library is not installed.")

        # Ensure filename is unique but identifiable
        filename = f"x_post_{int(time.time())}.png"
        output_path = os.path.join(save_dir, filename)

        # Theme Configuration
        night_mode_val = "0"
        color_scheme = "light"
        if theme_mode == "dark":
            night_mode_val = "2"
            color_scheme = "dark"

        with sync_playwright() as p:
            try:
                browser = p.chromium.launch(headless=True)
            except PlaywrightError as e:
                if "Executable doesn't exist" in str(e) or "playwright install" in str(e):
                    raise RuntimeError("BROWSERS_MISSING")
                raise e

            context = browser.new_context(
                viewport={'width': 600, 'height': 1000},
                device_scale_factor=2,
                color_scheme=color_scheme
            )
            
            context.add_cookies([{
                "name": "night_mode",
                "value": night_mode_val,
                "domain": ".x.com",
                "path": "/"
            }, {
                "name": "night_mode",
                "value": night_mode_val,
                "domain": ".twitter.com",
                "path": "/"
            }])

            page = context.new_page()

            try:
                print(f"Navigating to: {url}")
                page.goto(url, wait_until="domcontentloaded")
                
                selector = 'article[data-testid="tweet"]'
                try:
                    page.wait_for_selector(selector, timeout=20000)
                except Exception as e:
                    body_text = page.text_content("body") or ""
                    if "Age-restricted" in body_text or "年齢制限のある" in body_text:
                        raise RuntimeError("Age-Restricted content. Cannot load without login.")
                    elif "deleted" in body_text.lower() or "削除されました" in body_text:
                        raise RuntimeError("This Post has been deleted.")
                    elif "log in" in body_text.lower() or "ログイン" in body_text:
                        raise RuntimeError("Login required. Post might be private or restricted.")
                    else:
                        raise RuntimeError("Timeout: Failed to load tweet.")
                
                tweet_element = page.locator(selector).first
                tweet_element.scroll_into_view_if_needed()
                
                # Smart Image Wait
                page.evaluate("""(selector) => {
                    const tweet = document.querySelector(selector);
                    if (!tweet) return;
                    const images = tweet.querySelectorAll('img');
                    return Promise.all(Array.from(images).map(img => {
                        if (img.complete) return Promise.resolve();
                        return new Promise(resolve => {
                            img.onload = resolve;
                            img.onerror = resolve;
                        });
                    }));
                }""", selector)
                
                time.sleep(1.5)

                # Dynamic Viewport Resizing
                box = tweet_element.bounding_box()
                if box:
                    content_height = box['height']
                    content_y = box['y']
                    required_height = content_y + content_height + 100
                    
                    if required_height > 1000:
                        page.set_viewport_size({"width": 600, "height": int(required_height)})
                        tweet_element.scroll_into_view_if_needed()
                        time.sleep(0.5)

                tweet_element.screenshot(path=output_path)
                print(f"Captured to: {output_path}")

            except Exception as e:
                browser.close()
                raise e

            browser.close()
            
        return output_path

class ResolveHandler:
    """DaVinci Resolve Timeline Operations"""
    def __init__(self):
        if resolve:
            self.project_manager = resolve.GetProjectManager()
            self.current_project = self.project_manager.GetCurrentProject()
            self.media_pool = self.current_project.GetMediaPool()
        else:
            self.project_manager = None

    def is_ready(self):
        return self.project_manager is not None

    def import_to_timeline(self, file_path):
        if not self.is_ready():
            return False, "Resolve API not connected."
        if not os.path.exists(file_path):
            return False, "File not found."

        imported_items = self.media_pool.ImportMedia([file_path])
        if not imported_items:
            return False, "Failed to import media."
        
        clip = imported_items[0]
        if self.media_pool.AppendToTimeline([clip]):
            return True, "Imported and appended to timeline."
        else:
            return True, "Imported to Media Pool."

class InputMethod:
    """Handles User Input via Fusion UI (Primary) or Tkinter (Fallback)"""
    def __init__(self):
        self.fusion = fusion
        self.use_fusion_ui = False
        self.default_path = get_default_save_path()
        
        if dvr_script is not None and self.fusion is not None:
            try:
                self.ui = self.fusion.UIManager
                self.dispatcher = dvr_script.UIDispatcher(self.ui)
                if self.dispatcher:
                    self.use_fusion_ui = True
            except Exception as e:
                print(f"Fusion UI init failed: {e}")
                self.use_fusion_ui = False
    
    def get_user_input(self, callback_func):
        if self.use_fusion_ui:
            self._show_fusion_ui(callback_func)
        else:
            self._show_tkinter_ui(callback_func)

    def _open_folder_dialog(self):
        # Always use Tkinter for directory selection as it's reliable
        try:
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            path = filedialog.askdirectory(initialdir=self.default_path, title="Select Save Folder")
            root.destroy()
            return path
        except ImportError:
            return None

    def _show_fusion_ui(self, callback_func):
        win_id = "com.davinci.x_importer"
        
        layout = self.ui.VGroup({'Spacing': 10}, [
            self.ui.Label({'Text': 'X Post Importer v0.4', 'Font': self.ui.Font({'PixelSize': 14, 'Bold': True})}),
            
            # URL Input
            self.ui.Label({'Text': 'Post URL:', 'Weight': 0}),
            self.ui.LineEdit({'ID': 'LineUrl', 'Text': '', 'PlaceholderText': 'https://x.com/...'}),
            
            # Save Path Input
            self.ui.Label({'Text': 'Save Location:', 'Weight': 0}),
            self.ui.HGroup({'Weight': 0}, [
                self.ui.LineEdit({'ID': 'LinePath', 'Text': self.default_path, 'ReadOnly': True, 'Weight': 1}),
                self.ui.Button({'ID': 'BtnBrowse', 'Text': 'Browse', 'Weight': 0}),
            ]),
            
            # Theme Input
            self.ui.Label({'Text': 'Theme:', 'Weight': 0}),
            self.ui.ComboBox({'ID': 'ComboTheme', 'Weight': 0}),
            
            self.ui.VGap(5),
            self.ui.Button({'ID': 'BtnImport', 'Text': 'Import to Timeline', 'Height': 30, 'Weight': 0}),
            self.ui.Label({'ID': 'LblStatus', 'Text': 'Ready', 'Alignment': {'AlignHCenter': True}, 'Weight': 1}),
        ])

        window = self.dispatcher.AddWindow({
            'ID': win_id,
            'Geometry': [100, 100, 420, 300],
            'WindowTitle': "Resolve X Importer v0.4.0 beta",
        }, [layout])

        combo = window.Find('ComboTheme')
        combo.AddItem("Light Mode")
        combo.AddItem("Dark Mode")

        def on_close(ev): self.dispatcher.ExitLoop()
        
        def on_browse(ev):
            new_path = self._open_folder_dialog()
            if new_path:
                window.Find("LinePath").Text = new_path

        def on_import(ev):
            url = window.Find("LineUrl").Text
            save_path = window.Find("LinePath").Text
            status_lbl = window.Find("LblStatus")
            
            theme_idx = window.Find("ComboTheme").CurrentIndex
            theme_mode = "light"
            if theme_idx == 1: theme_mode = "dark"

            status_lbl.Text = "Capturing..."
            time.sleep(0.1)
            
            success, msg = callback_func(url, save_path, theme_mode)
            
            if msg == "BROWSERS_MISSING":
                status_lbl.Text = "Error: Browsers missing."
                install_playwright_browsers()
                status_lbl.Text = "Installer launched."
            else:
                status_lbl.Text = msg

        window.On[win_id].Close = on_close
        window.On['BtnImport'].Clicked = on_import
        window.On['BtnBrowse'].Clicked = on_browse
        window.Show()
        self.dispatcher.RunLoop()

    def _show_tkinter_ui(self, callback_func):
        try:
            import tkinter as tk
            from tkinter import simpledialog, messagebox, filedialog
            
            root = tk.Tk()
            root.title("Resolve X Importer v0.4.0 beta")
            root.geometry("450x300")
            root.attributes("-topmost", True)
            
            # URL
            tk.Label(root, text="Post URL:").pack(pady=2)
            url_var = tk.StringVar()
            tk.Entry(root, textvariable=url_var, width=55).pack(pady=2)
            
            # Path
            tk.Label(root, text="Save Location:").pack(pady=2)
            path_frame = tk.Frame(root)
            path_frame.pack(fill=tk.X, padx=20)
            path_var = tk.StringVar(value=self.default_path)
            tk.Entry(path_frame, textvariable=path_var).pack(side=tk.LEFT, fill=tk.X, expand=True)
            
            def browse_folder():
                p = filedialog.askdirectory(initialdir=path_var.get())
                if p: path_var.set(p)
            
            tk.Button(path_frame, text="...", command=browse_folder, width=3).pack(side=tk.LEFT, padx=5)

            # Theme
            tk.Label(root, text="Theme:").pack(pady=2)
            theme_var = tk.StringVar(value="dark")
            frame_theme = tk.Frame(root)
            frame_theme.pack(pady=2)
            tk.Radiobutton(frame_theme, text="Light", variable=theme_var, value="light").pack(side=tk.LEFT)
            tk.Radiobutton(frame_theme, text="Dark", variable=theme_var, value="dark").pack(side=tk.LEFT)
            
            def run_import():
                url = url_var.get()
                theme = theme_var.get()
                path = path_var.get()
                if not url: return
                
                root.title("Processing...")
                root.update()
                
                success, msg = callback_func(url, path, theme)
                
                if msg == "BROWSERS_MISSING":
                    do_install = messagebox.askyesno("Error", "Browsers missing. Install now?")
                    if do_install:
                        install_playwright_browsers()
                        root.destroy()
                elif success:
                    messagebox.showinfo("Success", msg)
                else:
                    messagebox.showerror("Error", msg)
                root.title("Resolve X Importer v0.4.0 beta")

            tk.Button(root, text="Import", command=run_import, bg="#DDDDDD").pack(pady=10)
            root.mainloop()

        except ImportError:
            pass

# --- Main Logic ---

def main():
    if resolve is None:
        print("Fatal Error: Could not connect to DaVinci Resolve.")
        return

    capture = XCapture()
    handler = ResolveHandler()
    input_ui = InputMethod()

    def process_request(url, save_dir, theme):
        if not PLAYWRIGHT_AVAILABLE:
            return False, "Error: 'playwright' library not installed."
        
        if not capture.validate_url(url):
            return False, "Invalid URL format."

        # ディレクトリ存在確認
        if not os.path.exists(save_dir):
            try:
                os.makedirs(save_dir)
            except OSError:
                return False, "Invalid Save Directory."
            
        try:
            img_path = capture.capture_post(url, save_dir, theme_mode=theme)
            success, msg = handler.import_to_timeline(img_path)
            return success, msg
        except RuntimeError as e:
            if str(e) == "BROWSERS_MISSING":
                return False, "BROWSERS_MISSING"
            return False, f"Error: {str(e)}"
        except Exception as e:
            print(f"Details: {e}")
            return False, f"Error: {str(e)}"

    input_ui.get_user_input(process_request)

if __name__ == "__main__":
    main()