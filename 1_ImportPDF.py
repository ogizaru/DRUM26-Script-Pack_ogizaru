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

    log("=== Python Installer Debug Start ===")

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
        {"ID": "InstallWin", "WindowTitle": "インストールの確認", "Geometry": [400, 200, 450, 250]},
        ui.VGroup([
            ui.Label({"Text": "以下の内容でインストールを実行します：", "Weight": 0}),
            ui.VGroup([
                ui.Label({"Text": "■ スクリプト名: " + filename, "Weight": 0}),
                ui.VGap(2),
                ui.Label({"Text": "■ インストール先: ", "Weight": 0}),
                ui.TextEdit({"Text": all_dir, "ReadOnly": True, "Weight": 1}),
                ui.VGap(4),
                ui.Label({"Text": "■ 自動セットアップされるライブラリ: ", "Weight": 0}),
                ui.Label({"Text": "   - PyMuPDF (PDF解析)", "Weight": 0}),
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
            log("Installing requirements (pymupdf)...")
            try:
                import subprocess
                # pip install pymupdf (--user を付けて安全にインストール)
                subprocess.check_call([sys.executable, "-m", "pip", "install", "--user", "pymupdf"])
                log("PyMuPDF installed successfully")
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
                    ui.Label({"Text": "失敗しました。権限などを確認してください。", "WordWrap": True, "Alignment": {"AlignHCenter": True, "AlignVCenter": True}, "Weight": 1}),
                    ui.Button({"ID": "BtnOk", "Text": "OK"})
                ])
            )
            ewin.On.BtnOk.Clicked = on_close; ewin.Show(); dispatcher.RunLoop(); ewin.Hide()
    return True

if not __install_this_script__():
    import sys; sys.exit()
# ============================================================================
"""
Import PDF v1.0.1
Developed by Google Antigravity & OGIZARU

Requirements:
- DaVinci Resolve Studio or Free (v16.2+)
- Python 3.1+
- libraries: playwright
  (Install via: pip install playwright)
"""

import sys
import os

# -----------------------------------------------------------------------------
# Dependency Check & Imports
# -----------------------------------------------------------------------------
try:
    import fitz  # PyMuPDF
except ImportError:
    print("Error: PyMuPDF (fitz) is not installed.")
    print("Please run: pip install pymupdf")
    sys.exit()

# -----------------------------------------------------------------------------
# Resolve / Fusion API Setup
# -----------------------------------------------------------------------------
try:
    resolve
except NameError:
    try:
        import DaVinciResolveScript as dvr_script
        resolve = dvr_script.scriptapp("Resolve")
    except ImportError:
        resolve = None

if not resolve:
    print("Could not connect to DaVinci Resolve.")
    sys.exit()

try:
    fusion = resolve.Fusion()
except:
    print("Could not get Fusion object.")
    sys.exit()

# -----------------------------------------------------------------------------
# Simplified Workflow (No Custom UI Window)
# -----------------------------------------------------------------------------
print("Starting PDF Import...")

# Request File directly
# RequestFile returns the path string or None
path_map = fusion.RequestFile(
    '',
    '',
    {
        "FReqS_Title": 'Select PDF File',
        "FReqS_Filter": 'PDF Files (*.pdf)|*.pdf',
    }
)

# RequestFile might return a map object in some versions or just a string?
# In Python, it usually returns a dictionary if multiple files allowed, or string?
# Let's handle the return type safely.
pdf_path = None
if isinstance(path_map, dict):
    # Depending on keys 'Filename' etc or just the key itself
    # Usually RequestFile returns { "Filename": "path" } or similar?
    # Or actually, in standard Fusion scripting: 
    # Returns: (string) path or (table) {[1]=path}
    # In Python, often looks like 'C:\\path\\to\\file.pdf' (string)
    # Let's inspect it if we can't assume.
    print(f"DEBUG: RequestFile returned type: {type(path_map)}")
    print(f"DEBUG: Value: {path_map}")
    # Try to extract reasonable path
    if 'Filename' in path_map:
        pdf_path = path_map['Filename']
    else:
        # Just grab the first value?
        pdf_path = list(path_map.values())[0] if path_map else None
elif isinstance(path_map, str):
    pdf_path = path_map
else:
    print(f"Unknown return type from RequestFile: {type(path_map)}")

if not pdf_path:
    print("No file selected.")
    sys.exit()

print(f"Selected: {pdf_path}")

# Default Settings (since we removed UI)
dpi = 150
output_fmt = "png"

# -----------------------------------------------------------------------------
# Logic
# -----------------------------------------------------------------------------
if not os.path.exists(pdf_path):
    print("File does not exist.")
    sys.exit()

# Create Output Directory
parent_dir = os.path.dirname(pdf_path)
filename = os.path.splitext(os.path.basename(pdf_path))[0]
output_dir = os.path.join(parent_dir, filename + "_images")

if not os.path.exists(output_dir):
    try:
        os.makedirs(output_dir)
    except OSError as e:
        print(f"Error creating directory: {e}")
        sys.exit()

print(f"Converting '{pdf_path}' to images in '{output_dir}'...")

# Conversion using PyMuPDF (fitz)
try:
    doc = fitz.open(pdf_path)
    img_paths = []
    
    for i in range(len(doc)):
        page = doc.load_page(i)
        print(f"Processing page {i+1}/{len(doc)}")
        
        # Zoom matrix for DPI
        # 72 dpi is 1.0 scale
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        
        pix = page.get_pixmap(matrix=mat)
        
        out_file = os.path.join(output_dir, f"{filename}_page_{i+1:03d}.{output_fmt}")
        pix.save(out_file)
        img_paths.append(out_file)
        
    doc.close()
    print("Conversion complete.")
    
    # Import to Media Pool
    project_manager = resolve.GetProjectManager()
    project = project_manager.GetCurrentProject()
    media_pool = project.GetMediaPool()
    
    # ImportMedia
    added_items = []
    if img_paths:
        print("Importing to Media Pool...")
        # Import one by one to avoid being treated as an image sequence (video clip)
        for p in img_paths:
            items = media_pool.ImportMedia([p])
            if items:
                added_items.extend(items)
        
        print(f"Imported {len(added_items)} items.")

    # Add to Timeline
    if added_items:
        timeline = project.GetCurrentTimeline()
        if not timeline:
            print("Creating new timeline...")
            timeline = media_pool.CreateEmptyTimeline("Imported PDF")
        
        if timeline:
            print("Adding to timeline...")
            media_pool.AppendToTimeline(added_items)
            print("Added to timeline.")

except Exception as e:
    print(f"Error during conversion or import: {e}")
    import traceback
    traceback.print_exc()

print("Done.")
