#!/usr/bin/env python
# ============================================================================
# 自動インストール機能 (ドラッグ＆ドロップ等での実行時)
# ============================================================================
def __install_this_script__():
    import os, shutil, sys, inspect
    
    def get_current_path():
        try:
            if "__file__" in globals(): return os.path.abspath(globals()["__file__"])
            return os.path.abspath(__file__)
        except: pass
        try:
            for f in inspect.stack():
                p = f[1]
                if p and os.path.exists(p) and p.endswith(".py"):
                    return os.path.abspath(p)
        except: pass
        gl = globals()
        for k, v in gl.items():
            if type(v) == str and os.path.exists(v) and v.endswith(".py"):
                return os.path.abspath(v)
        return None

    current_path = get_current_path()
    if not current_path: return True 
        
    current_path = current_path.replace("\\", "/")
    filename = os.path.basename(current_path)
    
    def get_bmd_module():
        import sys
        try: import DaVinciResolveScript as bmd; return bmd
        except ImportError: return sys.modules.get("fusionscript")

    bmd = get_bmd_module()
    fuset = None
    try:
        if 'fu' in globals(): fuset = fu
        elif 'resolve' in globals(): fuset = resolve.GetFusion()
        elif bmd: fuset = bmd.scriptapp("Resolve").GetFusion()
    except: return True
    if not fuset: return True
    
    import platform as py_platform
    is_win = (py_platform.system() == "Windows")
    all_dir = None

    if is_win:
        appdata = os.environ.get("APPDATA")
        if not appdata: return True 
        all_dir = os.path.join(appdata, "Blackmagic Design", "DaVinci Resolve", "Support", "Fusion", "Scripts", "Utility")
    else:
        all_dir = os.path.expanduser("~/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Utility/")

    # 【判定強化】パスの正規化比較
    src_n = current_path.replace("\\", "/").lower()
    a_n = all_dir.replace("\\", "/").lower() + "/" + filename.lower()

    if src_n == a_n:
        return True # すでにインストール済み

    ui = getattr(fuset, "UIManager", None)
    if not ui: return True
    dispatcher = bmd.UIDispatcher(ui) if bmd and hasattr(bmd, "UIDispatcher") else None
    if not dispatcher: return True

    win = dispatcher.AddWindow(
        {"ID": "InstallWin", "WindowTitle": "インストールの確認", "Geometry": [400, 200, 450, 220]},
        ui.VGroup([
            ui.Label({"Text": "以下の内容でインストールを実行します：", "Weight": 0}),
            ui.VGroup([
                ui.Label({"Text": "■ スクリプト名: " + filename, "Weight": 0}),
                ui.VGap(2),
                ui.Label({"Text": "■ インストール先: ", "Weight": 0}),
                ui.TextEdit({"Text": all_dir, "ReadOnly": True, "Weight": 1}),
            ]),
            ui.HGroup({"Weight": 0}, [
                ui.Button({"ID": "BtnOk", "Text": "インストール"}),
                ui.Button({"ID": "BtnCancel", "Text": "キャンセル"})
            ])
        ])
    )
    
    result = {"install": False}
    def on_ok(ev): result["install"] = True; dispatcher.ExitLoop()
    def on_close(ev): dispatcher.ExitLoop()
        
    win.On.BtnOk.Clicked = on_ok
    win.On.BtnCancel.Clicked = on_close
    win.On.InstallWin.Close = on_close
    win.Show(); dispatcher.RunLoop(); win.Hide()
    
    if result["install"]:
        target_dir = all_dir
        target_path = os.path.join(target_dir, filename)
        if not os.path.exists(target_dir):
            try: os.makedirs(target_dir)
            except: pass
        try:
            shutil.copy(current_path, target_path)
            cwin = dispatcher.AddWindow(
                {"ID": "CWin", "WindowTitle": "完了", "Geometry": [400, 200, 250, 100]},
                ui.VGroup([
                    ui.Label({"Text": "インストールが完了しました。", "WordWrap": True, "Alignment": {"AlignHCenter": True, "AlignVCenter": True}, "Weight": 1}),
                    ui.Button({"ID": "BtnOk", "Text": "OK", "Weight": 0})
                ])
            )
            cwin.On.BtnOk.Clicked = on_close; cwin.Show(); dispatcher.RunLoop(); cwin.Hide()
            return False
        except:
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
# -*- coding: utf-8 -*-

"""
Script Name: Web_Image_Loader (WIL) v1.0
Description: Web上の画像をURLから直接ダウンロードし、タイムラインの再生ヘッド位置に配置します。
Author: DaVinci Resolve Addon and DCTL maker V3 & OGIZARU
"""

import os
import sys
import hashlib
import time
import urllib.request
import urllib.parse
import ssl
import platform

# --- Resolve API Setup ---
try:
    resolve = bmd.scriptapp("Resolve")
    fusion = resolve.Fusion()
    ui = fusion.UIManager
    dispatcher = bmd.UIDispatcher(ui)
except NameError:
    print("Error: This script must be run inside DaVinci Resolve.")
    sys.exit(1)

# --- Configuration ---
CACHE_DIR_NAME = "Resolve_WIL_Cache"
CACHE_PATH = os.path.join(os.path.expanduser("~/Documents"), CACHE_DIR_NAME)

# --- Helper Functions ---

def ensure_cache_dir():
    """キャッシュディレクトリの存在を確認し、なければ作成する"""
    if not os.path.exists(CACHE_PATH):
        try:
            os.makedirs(CACHE_PATH)
            print(f"Created cache directory: {CACHE_PATH}")
        except OSError as e:
            print(f"Error creating directory: {e}")
            return False
    return True

def get_safe_filename(url):
    """URLから安全かつ一意なファイル名を生成する"""
    parsed = urllib.parse.urlparse(url)
    domain = parsed.netloc.replace('.', '_')
    path = parsed.path
    ext = os.path.splitext(path)[1]
    if not ext or len(ext) > 5:
        ext = ".jpg"
    
    url_hash = hashlib.md5(url.encode('utf-8')).hexdigest()[:8]
    timestamp = int(time.time())
    
    filename = f"{domain}_{url_hash}_{timestamp}{ext}"
    filename = "".join([c for c in filename if c.isalnum() or c in ('-', '_', '.')])
    
    return os.path.join(CACHE_PATH, filename)

def download_image(url):
    """URLから画像をダウンロードし、保存されたパスを返す"""
    if not ensure_cache_dir():
        return None

    save_path = get_safe_filename(url)
    ctx = ssl.create_default_context()
    ctx.check_hostname = False
    ctx.verify_mode = ssl.CERT_NONE

    try:
        print(f"Downloading: {url}")
        with urllib.request.urlopen(url, context=ctx) as response, open(save_path, 'wb') as out_file:
            out_file.write(response.read())
        print(f"Saved to: {save_path}")
        return save_path
    except Exception as e:
        print(f"Download Error: {e}")
        return None

def get_current_timeline_context():
    """現在のプロジェクト、タイムライン、メディアプールを取得"""
    project_manager = resolve.GetProjectManager()
    project = project_manager.GetCurrentProject()
    if not project:
        return None, None, None, None
    
    timeline = project.GetCurrentTimeline()
    if not timeline:
        return project, None, None, None
    
    media_pool = project.GetMediaPool()
    timecode = timeline.GetCurrentTimecode()
    
    return project, timeline, media_pool, timecode

# --- Main Logic Classes ---

class WIL_Controller:
    def __init__(self, output_text_field, track_combo):
        self.output = output_text_field
        self.track_combo = track_combo

    def log(self, message):
        """ログ出力"""
        print(message)
        current_text = self.output.Text
        self.output.Text = f"{message}\n{current_text}"

    def load_tracks(self):
        """タイムライン上のビデオトラックを取得してコンボボックスに設定"""
        project, timeline, _, _ = get_current_timeline_context()
        self.track_combo.Clear()
        
        if not timeline:
            self.track_combo.AddItem("No Timeline Active")
            return

        track_count = timeline.GetTrackCount("video")
        if track_count == 0:
            self.track_combo.AddItem("No Video Tracks")
            return

        for i in range(1, track_count + 1):
            # トラック名を取得（APIが対応していない場合はV1, V2...とする）
            try:
                name = timeline.GetTrackName("video", i)
            except:
                name = ""
            
            if not name:
                name = f"Video {i}"
            
            self.track_combo.AddItem(f"{i}: {name}")
        
        # デフォルトで一番上のトラック（V1）を選択（インデックス0）
        # ただし、トラックが存在する場合のみ
        if track_count > 0:
            # Fusion UIのComboBoxにはSetCurrentIndexがない場合があるため、動作確認が必要。
            # 通常はAddItemした順に0,1,2...となる。
            pass

    def add_image_to_timeline(self, url):
        """画像をダウンロードし、選択されたトラックに配置"""
        project, timeline, media_pool, _ = get_current_timeline_context()
        if not timeline:
            self.log("Error: No active timeline found.")
            return

        # コンボボックスから選択されたインデックスを取得 (0始まり)
        selected_idx = self.track_combo.CurrentIndex
        # トラック番号は1始まりなので +1
        track_number = selected_idx + 1

        # ダウンロード処理
        file_path = download_image(url)
        if not file_path:
            self.log("Error: Failed to download image.")
            return

        # インポート
        imported_items = media_pool.ImportMedia([file_path])
        if not imported_items:
            self.log("Error: Failed to import media.")
            return
        
        item = imported_items[0]
        item.SetMetadata("SourceURL", url)
        item.SetMetadata("WIL_Managed", "True")

        # 配置
        clip_info = {
            "mediaPoolItem": item,
            "trackIndex": track_number,
            "mediaType": 1 
        }
        
        if media_pool.AppendToTimeline([clip_info]):
            self.log(f"Success: Added to Track {track_number}")
        else:
            self.log("Error: Failed to append to timeline. (Check track lock?)")

    def refresh_all_images(self):
        """画像の再ロード"""
        _, _, media_pool, _ = get_current_timeline_context()
        if not media_pool: return

        self.log("Starting Refresh Process...")
        folder = media_pool.GetRootFolder()
        clips = folder.GetClipList()
        count = 0
        
        for clip in clips:
            url = clip.GetMetadata("SourceURL")
            if url:
                self.log(f"Updating: {clip.GetName()}...")
                new_path = download_image(url)
                if new_path:
                    if clip.ReplaceClip(new_path):
                        self.log(f"  -> Replaced.")
                        count += 1
                    else:
                        self.log("  -> Replace failed.")
                else:
                    self.log("  -> Download failed.")
        self.log(f"Refresh Complete. Updated {count} clips.")

    def clear_cache(self):
        """キャッシュクリア"""
        if not os.path.exists(CACHE_PATH): return
        
        abs_cache = os.path.abspath(CACHE_PATH)
        abs_docs = os.path.abspath(os.path.expanduser("~/Documents"))
        if not abs_cache.startswith(abs_docs) or CACHE_DIR_NAME not in abs_cache:
            self.log("Security Halt: Invalid cache path.")
            return

        files = os.listdir(CACHE_PATH)
        deleted = 0
        for f in files:
            full_path = os.path.join(CACHE_PATH, f)
            if f.lower().endswith(('.jpg', '.jpeg', '.png', '.bmp', '.gif')):
                try:
                    os.remove(full_path)
                    deleted += 1
                except OSError: pass
        self.log(f"Cache Cleared. Deleted {deleted} files.")

# --- UI Setup ---

def create_ui():
    win = dispatcher.AddWindow({
        'ID': 'WIL_Window',
        'WindowTitle': 'Web Image Loader V3',
        'Geometry': [600, 300, 400, 450],
        'Spacing': 10,
    }, [
        ui.VGroup([
            ui.Label({'ID': 'Lbl_URL', 'Text': 'Image URL:', 'Weight': 0}),
            ui.LineEdit({'ID': 'Input_URL', 'Text': '', 'PlaceholderText': 'Paste URL here...', 'Weight': 0}),
            
            # --- 変更箇所: ドロップダウンUI ---
            ui.HGroup({'Weight': 0}, [
                ui.Label({'Text': 'Target Track:', 'Weight': 0}),
                ui.ComboBox({'ID': 'Combo_Track', 'Weight': 1}), # コンボボックスに変更
                ui.Button({'ID': 'Btn_ReloadTracks', 'Text': '↻', 'Weight': 0, 'ToolTip': 'Reload Track List'}), # トラックリスト更新ボタン
            ]),

            ui.VGap(5),
            ui.Button({'ID': 'Btn_Add', 'Text': 'Download & Add to Timeline', 'Weight': 0, 'Height': 30}),
            ui.VGap(10),
            
            ui.HGroup({'Weight': 0}, [
                ui.Button({'ID': 'Btn_Refresh', 'Text': 'Refresh All Images', 'Weight': 1}),
                ui.Button({'ID': 'Btn_Clear', 'Text': 'Clear Cache Folder', 'Weight': 1}),
            ]),
            
            ui.Label({'Text': 'Log:', 'Weight': 0}),
            ui.TextEdit({'ID': 'Txt_Log', 'ReadOnly': True, 'Weight': 1, 'Font': 'Monospace'})
        ])
    ])
    return win

# --- Event Loop ---

win = create_ui()
itm = win.GetItems()

# コントローラー初期化時にコンボボックスUIを渡す
controller = WIL_Controller(itm['Txt_Log'], itm['Combo_Track'])

def OnAdd(ev):
    url = itm['Input_URL'].Text
    if not url:
        controller.log("Please enter a URL.")
        return
    controller.add_image_to_timeline(url)

def OnRefresh(ev):
    controller.refresh_all_images()

def OnClear(ev):
    controller.clear_cache()

def OnReloadTracks(ev):
    controller.load_tracks()
    controller.log("Track list reloaded.")

def OnShow(ev):
    # ウインドウ表示時にトラックリストを読み込む
    controller.load_tracks()

def OnClose(ev):
    dispatcher.ExitLoop()

win.On.Btn_Add.Clicked = OnAdd
win.On.Btn_Refresh.Clicked = OnRefresh
win.On.Btn_Clear.Clicked = OnClear
win.On.Btn_ReloadTracks.Clicked = OnReloadTracks
win.On.WIL_Window.Show = OnShow # 表示時のイベントフック
win.On.WIL_Window.Close = OnClose

controller.log(f"Initialized. Cache Path: {CACHE_PATH}")
win.Show()
dispatcher.RunLoop()
win.Hide()