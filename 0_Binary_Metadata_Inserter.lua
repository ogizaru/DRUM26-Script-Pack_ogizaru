-- ============================================================================
-- 自動インストール機能 (ドラッグ＆ドロップ等での実行時)
-- ============================================================================
local function __install_this_script__()
    local info = debug.getinfo(1, "S")
    if not info or not info.source then return true end
    
    local source = info.source
    if source:sub(1,1) == "@" then source = source:sub(2) else return true end
    
    source = source:gsub("\\", "/")
    local filename = source:match("([^/]+)$")
    if not filename then return true end

    -- Fusion オブジェクトの取得
    local fuset = nil
    if type(fu) == "table" then fuset = fu
    elseif type(resolve) == "userdata" or type(resolve) == "table" then
        if type(resolve.Fusion) == "function" then fuset = resolve:Fusion()
        elseif type(resolve.GetFusion) == "function" then fuset = resolve:GetFusion() end
    end
    if not fuset then return true end

    -- パス組み立て (Windows / Mac)
    local platform = (package.config:sub(1,1) == '\\') and "Windows" or "Mac"
    local all_dir -- デフォルトのインストール先
    local check_dirs = {}

    if platform == "Windows" then
        local appdata = os.getenv("APPDATA")
        local progdata = os.getenv("PROGRAMDATA") or "C:\\ProgramData"
        
        if appdata then
            local u_dir = appdata .. "\\Blackmagic Design\\DaVinci Resolve\\Support\\Fusion\\Scripts\\Utility\\"
            table.insert(check_dirs, u_dir)
            all_dir = u_dir
        end
        local s_dir = progdata .. "\\Blackmagic Design\\DaVinci Resolve\\Fusion\\Scripts\\Utility\\"
        table.insert(check_dirs, s_dir)
        if not all_dir then all_dir = s_dir end
    else
        local home = os.getenv("HOME")
        if home then
            local u_dir = home .. "/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Utility/"
            table.insert(check_dirs, u_dir)
            all_dir = u_dir
        end
        local s_dir = "/Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Utility/"
        table.insert(check_dirs, s_dir)
        if not all_dir then all_dir = s_dir end
    end

    if not all_dir then return true end

    -- 【判定強化】パスの正規化比較
    local function norm(p) return p:gsub("\\", "/"):lower() end
    local src_n = norm(source)

    local is_installed = false
    for _, dir in ipairs(check_dirs) do
        if src_n == norm(dir .. filename) then
            is_installed = true
            break
        end
    end

    if is_installed then
        return true -- すでにインストール済み
    end

    local ui = fuset.UIManager
    local disp = bmd.UIDispatcher(ui)
    
    local win = disp:AddWindow({
        ID = "InstallWin", WindowTitle = "インストールの確認", Geometry = { 400, 200, 450, 220 },
        WindowFlags = { Window = true, Dialog = true, Modal = true },
        ui:VGroup{
            Spacing = 8,
            ui:Label{ Text = "以下の内容でインストールを実行します：", Weight = 0 },
            ui:VGroup{
                Weight = 1,
                ui:Label{ Text = "■ スクリプト名: " .. filename, Weight = 0 },
                ui:VGap(2),
                ui:Label{ Text = "■ インストール先: ", Weight = 0 },
                ui:TextEdit{ ID = "TxtPath", Text = all_dir, ReadOnly = true, Weight = 1 },
            },
            ui:HGroup{
                Weight = 0,
                ui:Button{ ID = "BtnOk", Text = "インストール" },
                ui:Button{ ID = "BtnCancel", Text = "キャンセル" }
            }
        }
    })
    
    local is_install = false
    function win.On.BtnOk.Clicked(ev) is_install = true; disp:ExitLoop() end
    function win.On.BtnCancel.Clicked(ev) disp:ExitLoop() end
    function win.On.InstallWin.Close(ev) disp:ExitLoop() end
    
    win:Show(); disp:RunLoop(); win:Hide()
    
    if is_install then
        local target_path = all_dir .. filename
        if platform == "Windows" then
            local native_dir = all_dir:gsub("/", "\\")
            if native_dir:sub(-1) == "\\" then native_dir = native_dir:sub(1, -2) end
            os.execute('if not exist "' .. native_dir .. '" mkdir "' .. native_dir .. '"')
            
            local native_src = source:gsub("/", "\\")
            os.execute('copy "' .. native_src .. '" "' .. native_dir .. '\\"')
        else
            os.execute('mkdir -p "' .. all_dir .. '"')
            os.execute('cp "' .. source .. '" "' .. all_dir .. '"')
        end
        
        -- 完了ダイアログの表示
        local cwin = disp:AddWindow({
            ID = "CWin", WindowTitle = "完了", Geometry = { 400, 200, 250, 100 },
            WindowFlags = { Window = true, Dialog = true, Modal = true },
            ui:VGroup{
                ui:Label{ Text = "完了しました。", WordWrap = true, Alignment = {AlignHCenter = true, AlignVCenter = true}, Weight = 1 },
                ui:Button{ ID = "BtnOk", Text = "OK", Weight = 0 }
            }
        })
        function cwin.On.BtnOk.Clicked(ev) disp:ExitLoop() end
        cwin:Show(); disp:RunLoop(); cwin:Hide()
        return false
    end
    return true
end

if not __install_this_script__() then return end
-- ============================================================================
--[[
    Resolve Exif & Metadata Importer v1.0.3
    Author: OGIZARU & DaVinci Resolve Addon and DCTL maker V3 (Custom Gem)
    Description: Reads EXIF & API Metadata. Default output is HORIZONTAL (single line).
    Target: DaVinci Resolve Studio / Free 17+
    Language: Lua 5.1
]]

local resolve = resolve
local fusion = resolve:Fusion()
local ui = fusion.UIManager
local disp = bmd.UIDispatcher(ui)

--------------------------------------------------------------------------------
-- 1. DEFINITIONS & CONFIG
--------------------------------------------------------------------------------

local CONFIG_FILE = fusion:MapPath("Scripts:/Utility/Binary_Metadata_Inserter_Config.conf")

local DEFAULT_CONFIG = {
    last_state = {},
    presets = {},
    options = {
        pad_width = false
    }
}

local function SerializeTable(val, name, skipnewlines, depth)
    skipnewlines = skipnewlines or false
    depth = depth or 0
    local tmp = string.rep(" ", depth)
    
    if name then
        if type(name) == "number" then
            tmp = tmp .. "[" .. tostring(name) .. "] = "
        elseif type(name) == "string" then
             if string.match(name, "^[a-zA-Z_][a-zA-Z0-9_]*$") then
                tmp = tmp .. name .. " = " 
             else
                tmp = tmp .. "[\"" .. name .. "\"] = "
             end
        else
            tmp = tmp .. "[\"" .. tostring(name) .. "\"] = "
        end
    end
    
    if type(val) == "table" then
        tmp = tmp .. "{" .. (not skipnewlines and "\n" or "")
        for k, v in pairs(val) do
            tmp =  tmp .. SerializeTable(v, k, skipnewlines, depth + 1) .. "," .. (not skipnewlines and "\n" or "")
        end
        tmp = tmp .. string.rep(" ", depth) .. "}"
    elseif type(val) == "number" then
        tmp = tmp .. tostring(val)
    elseif type(val) == "string" then
        tmp = tmp .. string.format("%q", val)
    elseif type(val) == "boolean" then
        tmp = tmp .. (val and "true" or "false")
    else
        tmp = tmp .. "\"[inserializeable datatype:" .. type(val) .. "]\""
    end
    return tmp
end


local function LoadConfig()
    local f = io.open(CONFIG_FILE, "r")
    if not f then return DEFAULT_CONFIG end
    local content = f:read("*a")
    f:close()
    if not content then return DEFAULT_CONFIG end
    local chunk = loadstring(content)
    if chunk then
        setfenv(chunk, {})
        local success, result = pcall(chunk)
        if success and type(result) == "table" then 
            -- Ensure structure
            result.presets = result.presets or {}
            result.options = result.options or {}
            result.last_state = result.last_state or {}
            return result 
        end
    end
    return DEFAULT_CONFIG
end

local function SaveConfig(in_data, in_presets, in_options)
    local data = {
        last_state = in_data,
        presets = in_presets,
        options = in_options
    }
    local f = io.open(CONFIG_FILE, "w")
    if not f then return end
    f:write("return " .. SerializeTable(data))
    f:close()
end



local STANDARD_EXIF_TAGS = {
    -- Basic
    [0x010F] = {name = "メーカー", name_en = "Make", key = "Make"},
    [0x0110] = {name = "モデル", name_en = "Model", key = "Model"},
    [0x0132] = {name = "更新日時", name_en = "Date Time", key = "DateTime"},
    [0x9003] = {name = "撮影日時", name_en = "Date Time Original", key = "DateTimeOriginal"},
    [0x9004] = {name = "作成日時", name_en = "Create Date", key = "CreateDate"},
    [0xA002] = {name = "画像幅", name_en = "Image Width", key = "ExifImageWidth"},
    [0xA003] = {name = "画像高さ", name_en = "Image Height", key = "ExifImageHeight"},
    -- Exposure
    [0x829A] = {name = "シャッタースピード", name_en = "Shutter Speed", key = "ExposureTime", fmt = "shutter"},
    [0x829D] = {name = "絞り値", name_en = "F-Number", key = "FNumber", fmt = "fnumber"},
    [0x8822] = {name = "露出プログラム", name_en = "Exposure Program", key = "ExposureProgram", map = {[0]="未検出",[1]="マニュアル", [2]="プログラムAE", [3]="絞り優先", [4]="シャッター優先", [5]="スロー", [6]="アクション", [7]="ポートレート", [8]="風景", [9]="バルブ"}},
    [0x8827] = {name = "ISO感度", name_en = "ISO", key = "ISO"},
    [0x9204] = {name = "露出補正", name_en = "Exposure Bias", key = "ExposureBiasValue", fmt = "bias"},
    [0x9205] = {name = "開放絞り値", name_en = "Max Aperture", key = "MaxApertureValue", fmt = "fnumber"},
    [0x9207] = {name = "測光モード", name_en = "Metering Mode", key = "MeteringMode", map = {[0]="未検出", [1]="平均", [2]="中央重点", [3]="スポット", [4]="マルチスポット測光", [5]="ESP", [6]="分割測光", [255]="その他"}},
    -- Lens
    [0x920A] = {name = "焦点距離", name_en = "Focal Length", key = "FocalLength", fmt = "focal"},
    [0xA405] = {name = "35mm換算焦点距離", name_en = "Focal Length (35mm)", key = "FocalLengthIn35mmFormat", fmt = "focal_simple"},
    [0xA404] = {name = "デジタルズーム", name_en = "Digital Zoom Ratio", key = "DigitalZoomRatio"},
    [0xA434] = {name = "レンズモデル", name_en = "Lens Model", key = "LensModel"},
    -- Config
    [0xA403] = {name = "ホワイトバランス", name_en = "White Balance", key = "WhiteBalance", map = {[0]="自動", [1]="マニュアル"}},
    [0xA408] = {name = "コントラスト", name_en = "Contrast", key = "Contrast", map = {[0]="標準", [1]="低", [2]="高"}},
    [0xA409] = {name = "彩度", name_en = "Saturation", key = "Saturation", map = {[0]="標準", [1]="低", [2]="高"}},
    [0xA40A] = {name = "シャープネス", name_en = "Sharpness", key = "Sharpness", map = {[0]="標準", [1]="弱", [2]="強"}},
    [0x0131] = {name = "ソフトウェア", name_en = "Software", key = "Software"},
    [0x8298] = {name = "著作権者", name_en = "Copyright", key = "Copyright"},
}

local MAKERNOTE_TAGS = {
    Panasonic = {
        [0x00ee] = {name = "ダイナミックレンジブースト(動画)", name_en = "Dynamic Range Boost", key = "DynamicRangeBoost", map = {[0]="OFF", [1]="ON"}},
        [0x00e9] = {name = "被写体認識AF", name_en = "AI Subject Detection", key = "AISubjectDetection", map = {[0]="OFF", [1]="人物瞳/顔/体", [2]="動物", [3]="人物瞳/顔", [4]="動物体", [5]="動物瞳/体", [6]="車全体", [7]="バイク・自転車全体", [8]="車主要部優先", [9]="バイク・自転車ヘルメット優先", [10]="鉄道先頭車両優先", [11]="鉄道主要部優先", [12]="飛行機全体", [13]="飛行機機首優先"}},
        [0x0051] = {name = "レンズ名(パナソニック)", name_en = "Lens Function Type", key = "LensType"},
        [0x0044] = {name = "色温度(K)", name_en = "Color Temp (K)", key = "ColorTemp"},
        [0x0089] = {name = "フォトスタイル", name_en = "Photo Style", key = "PhotoStyle", map = {[1]="スタンダード", [2]="ヴィヴィッド", [3]="ナチュラル", [4]="モノクローム", [5]="風景", [6]="人物", [11]="L.モノクローム", [12]="709ライク", [15]="L.モノクロームD", [16]="フラットまたはカスタム", [17]="V-Log", [18]="シネライクD2", [19]="シネライクV2", [20]="L.モノクロームS", [21]="L.クラシックネオ", [22]="LEICAモノクローム", [24]="シネライクA2"}},
        [0x00d2] = {name = "粒状効果", name_en = "Grain Effect", key = "MonochromeGrainEffect", map = {[0]="なし", [1]="低", [2]="中", [3]="高"}},
        [0x00f1] = {name = "リアルタイムLUT1", name_en = "Realtime LUT1", key = "LUT1Name"},
        [0x00f4] = {name = "リアルタイムLUT2", name_en = "Realtime LUT2", key = "LUT2Name"},
        [0x0025] = {name = "シリアル番号", name_en = "Internal Serial Number", key = "InternalSerialNumber"},
        [0x001F] = {name = "撮影モード", name_en = "Shooting Mode", key = "ShootingMode"},
    },
    Olympus = {
        [0x0104] = {name = "ボディーファームウェア", name_en = "Body Firmware", key = "BodyFirmwareVersion", fmt = "om_firmware"},
        [0x0203] = {name = "レンズモデル", name_en = "Lens Model", key = "LensModel"},
        [0x2010] = {name = "Equipment"}, 
        [0x2020] = {name = "CameraSettings"}, 
        [0x2030] = {name = "RawDevelopment"}, 
        [0x2031] = {name = "RawDevelopment2"}, 
        [0x2040] = {name = "ImageProcessing"}, 
        [0x2050] = {name = "FocusInfo"}, 
    },
    OlympusCS = {
        [0x0301] = {name = "フォーカスモード", name_en = "Focus Mode", key = "FocusMode", fmt = "focus_mode_legacy"}, 
    },
    OlympusRD2 = {
        [0x0126] = {name = "ハイライトシャドウコントロール", name_en = "Highlight & Shadow Control", key = "HighlightShadowControl", fmt = "om_tone_control"},
    },
    OlympusEq = {
        [0x0301] = {name = "フォーカスモード", name_en = "Focus Mode", key = "FocusMode", fmt = "focus_mode_legacy"}, 
        [0x0310] = {name = "フォーカスモード", name_en = "Focus Mode", key = "FocusMode", fmt = "focus_mode_modern"}, 
        [0x0309] = {name = "AI被写体認識AF", name_en = "Subject Detection AF", key = "AISubjectTrackingMode", map = {[0]="オフ",[256]="モータースポーツ(未検出)",[257]="モータースポーツ(レーシングカー)",[258]="モータースポーツ(車)",[259]="モータースポーツ(バイク)",[512]="飛行機(未検出)",[513]="飛行機(旅客機・輸送機)",[514]="飛行機(小型機・戦闘機)",[515]="飛行機(ヘリコプター)",[768]="鉄道(未検出)",[769]="鉄道",[1024]="鳥(未検出)",[1025]="鳥",[1280]="動物(未検出)",[1281]="動物",[1536]="人物(未検出)",[1537]="人物"}},
    },
    OlympusIP = {
        [0x0301] = {name = "フォーカスモード", name_en = "Focus Mode", key = "FocusMode", fmt = "focus_mode_om1"}, 
        [0x0501] = {name = "色温度", name_en = "White Balance Temp", key = "WhiteBalanceTemperature"},
        [0x0520] = {name = "ピクチャーモード", name_en = "Picture Mode", key = "PictureMode", fmt = "om_picture_mode"},
        [0x0521] = {name = "彩度(OM)", name_en = "Saturation (OM)", key = "OlympusPictureModeSaturation", fmt = "om_first_number"},
        [0x0523] = {name = "コントラスト(OM)", name_en = "Contrast (OM)", key = "OlympusPictureModeContrast", fmt = "om_first_number"},
        [0x0524] = {name = "シャープネス(OM)", name_en = "Sharpness (OM)", key = "OlympusPictureModeSharpness", fmt = "om_first_number"},
        [0x0525] = {name = "モノクロフィルター", name_en = "Monochrome Filter", key = "PictureModeBWFilter", map = {[0]="なし",[1]="ニュートラル",[2]="黄",[3]="オレンジ",[4]="赤",[5]="緑"}},
        [0x0526] = {name = "調色", name_en = "Picture Mode Tone", key = "PictureModeTone", map = {[0]="なし",[1]="ニュートラル",[2]="セピア",[3]="青",[4]="紫",[5]="緑"}},
        [0x0527] = {name = "ノイズフィルター", name_en = "Noise Filter", key = "NoiseFilter", fmt = "om_noise_filter"},
        [0x0529] = {name = "アートフィルター", name_en = "Art Filter", key = "ArtFilter", fmt = "om_art_filter"},
        [0x052b] = {name = "ハイライトコントロール(旧)", name_en = "Highlight Control (Legacy)"},
        [0x052c] = {name = "シャドウコントロール(旧)", name_en = "Shadow Control (Legacy)"},
        [0x052d] = {name = "ミッドトーンコントロール(旧)", name_en = "Midtone Control (Legacy)"},
        [0x052e] = {name = "ハイライトシャドウコントロール", name_en = "Highlight & Shadow Control", key = "HighlightShadowControl", fmt = "om_tone_control"},
        [0x0537] = {name = "モノクロプロファイル設定", name_en = "Monochrome Profile Settings", key = "MonochromeProfileSettings", map = {[0]="カラーフィルターなし",[1]="黄",[2]="オレンジ",[3]="赤",[4]="マゼンタ",[5]="ブルー",[6]="シアン",[7]="緑",[8]="黄色/緑"}},
        [0x0538] = {name = "粒状効果", name_en = "Film Grain Effect", key = "FilmGrainEffect", map = {[0]="なし",[1]="低",[2]="中",[3]="高"}},
        [0x053a] = {name = "モノクロビネッティング", name_en = "Monochrome Vignetting", key = "MonochromeVignetting"},
        [0x0604] = {name = "手ぶれ補正モード", name_en = "Image Stabilization", key = "ImageStabilization", map = {[0]="OFF",[1]="S-IS1(全方向)",[2]="S-IS2(縦方向)",[3]="S-IS3(横方向)",[4]="S-IS Auto"}},
        [0x0804] = {name = "撮影設定", name_en = "Shooting Method", key = "StackedImage", fmt = "stacked_image"},
    },
    OlympusFI = {
        -- FocusInfo 領域は再帰パスのために維持（不要なデバッグタグは削除）
    },
}

local API_PROPS = {
    {name = "FPS", name_en = "FPS", key = "FPS"},
    {name = "解像度", name_en = "Resolution", key = "Resolution"},
    {name = "形式", name_en = "Format", key = "Format"},
    {name = "コーデック", name_en = "Video Codec", key = "Video Codec"},
    {name = "ビデオビット深度", name_en = "Video Bit Depth", key = "Video Bit Depth"},
    {name = "オーディオCH", name_en = "Audio Channels", key = "Audio Channels"},
    {name = "オーディオビット深度", name_en = "Audio Bit Depth", key = "Audio Bit Depth"},
    -- 常用カメラプロパティの追加
    {name = "レンズ", name_en = "Lens", key = "Lens"},
    {name = "カメラモデル", name_en = "Camera Type", key = "Camera Type"},
    {name = "ISO感度", name_en = "ISO", key = "ISO"},
    {name = "絞り値", name_en = "Aperture", key = "Aperture"},
    {name = "シャッター", name_en = "Shutter", key = "Shutter"},
    {name = "ホワイトバランス", name_en = "White Balance", key = "White Balance"},
    {name = "ガンマ", name_en = "Gamma", key = "Gamma"},
    {name = "カラースペース", name_en = "Color Space", key = "Color Space"},
}

-- Binary Reader with FFI support for Unicode paths
local ffi = require("ffi")
ffi.cdef[[
    typedef void FILE;
    FILE* _wfopen(const uint16_t* filename, const uint16_t* mode);
    int fclose(FILE* stream);
    size_t fread(void* ptr, size_t size, size_t count, FILE* stream);
    int fseek(FILE* stream, long offset, int origin);
    long ftell(FILE* stream);
    int MultiByteToWideChar(unsigned int CodePage, unsigned long dwFlags, const char* lpMultiByteStr, int cbMultiByte, uint16_t* lpWideCharStr, int cchWideChar);
]]

local BinaryReader = {}
BinaryReader.__index = BinaryReader

local function to_wstring(str)
    if not str then return nil end
    local size = ffi.C.MultiByteToWideChar(65001, 0, str, -1, nil, 0) -- CP_UTF8 = 65001
    if size == 0 then return nil end
    local buf = ffi.new("uint16_t[?]", size)
    ffi.C.MultiByteToWideChar(65001, 0, str, -1, buf, size)
    return buf
end

function BinaryReader.new(file_path)
    if not file_path then return nil end
    
    local platform = (package.config:sub(1,1) == '\\') and "Windows" or "Mac"
    local data = nil
    
    if platform == "Windows" then
        local wpath = to_wstring(file_path)
        local wmode = to_wstring("rb")
        local f = ffi.C._wfopen(wpath, wmode)
        
        if f == nil then return nil end
        
        ffi.C.fseek(f, 0, 2) -- SEEK_END
        local size = ffi.C.ftell(f)
        ffi.C.fseek(f, 0, 0) -- SEEK_SET
        
        if size < 0 or size > 32 * 1024 * 1024 then size = 32 * 1024 * 1024 end
        
        local buf = ffi.new("uint8_t[?]", size)
        local read_bytes = ffi.C.fread(buf, 1, size, f)
        ffi.C.fclose(f)
        
        if read_bytes == 0 then return nil end
        data = ffi.string(buf, read_bytes)
    else
        -- Mac: 標準機能で読み込む (UTF-8パスがそのまま扱えます)
        local f = io.open(file_path, "rb")
        if not f then return nil end
        data = f:read("*all")
        f:close()
        if not data or #data == 0 then return nil end
    end
    
    local self = setmetatable({}, BinaryReader)
    self.data = data
    self.pos = 1
    self.is_le = true
    return self
end
function BinaryReader:seek(pos) self.pos = pos end
function BinaryReader:set_endian(is_little) self.is_le = is_little end
function BinaryReader:u8()
    if self.pos > #self.data then return 0 end
    local b = string.byte(self.data, self.pos)
    self.pos = self.pos + 1
    return b
end
function BinaryReader:i8()
    local u = self:u8()
    if u >= 128 then return u - 256 else return u end
end
function BinaryReader:u16()
    if self.pos + 1 > #self.data then return 0 end
    local b1, b2 = string.byte(self.data, self.pos, self.pos + 1)
    self.pos = self.pos + 2
    if self.is_le then return b1 + b2 * 256 else return b2 + b1 * 256 end
end
function BinaryReader:i16()
    local u = self:u16()
    if u >= 32768 then return u - 65536 else return u end
end
function BinaryReader:u32()
    if self.pos + 3 > #self.data then return 0 end
    local b1, b2, b3, b4 = string.byte(self.data, self.pos, self.pos + 3)
    self.pos = self.pos + 4
    if self.is_le then return b1 + b2*256 + b3*65536 + b4*16777216 else return b4 + b3*256 + b2*65536 + b1*16777216 end
end
function BinaryReader:i32()
    local u = self:u32()
    if u >= 2147483648 then return u - 4294967296 else return u end
end
function BinaryReader:find(pattern_str, start_pos)
    return string.find(self.data, pattern_str, start_pos or 1, true)
end

-- Formatters
local Formatters = {}
function Formatters.fraction_reduce(n, d)
    if d == 0 then return "0" end
    local function gcd(a, b) while b ~= 0 do a, b = b, a % b end return a end
    local common = gcd(n, d)
    if d/common == 1 then return tostring(n/common) end
    return string.format("%d/%d", n/common, d/common)
end
function Formatters.shutter(n, d)
    if not d or d == 0 then return tostring(n) end
    if n == 0 then return "0" end
    if n < d then return string.format("1/%d", math.floor(d/n + 0.5)) else return string.format("%g\"", n/d) end
end
function Formatters.fnumber(n, d)
    if not d or d == 0 then return tostring(n) end
    return string.format("F%.1f", n/d)
end
function Formatters.focal(n, d)
    if not d or d == 0 then return tostring(n) end
    return string.format("%dmm", math.floor(n/d + 0.5))
end
function Formatters.focal_simple(n, d) 
    if d and d > 0 then
        return string.format("%dmm", math.floor(n/d + 0.5))
    end
    local num = tonumber(string.match(tostring(n), "^%-?%d+"))
    if num then return string.format("%dmm", num) end
    return tostring(n)
end
function Formatters.bias(n, d)
    if not d or d == 0 then return "0" end
    local val = n/d
    if val > 0 then return string.format("+%.1f", val) elseif val == 0 then return "±0" else return string.format("%.1f", val) end
end
local function format_focus_mode_shared(val, m)
    if not val then return nil end
    local s_val = tostring(val)
    local n = tonumber(s_val:match("%d+"))
    if n and m[n] then return m[n] end
    return s_val
end

function Formatters.focus_mode_legacy(val)
    local m = { [0] = "S-AF", [1] = "S-AF", [2] = "C-AF", [10] = "MF" }
    return format_focus_mode_shared(val, m)
end

function Formatters.focus_mode_modern(val)
    local m = {
        [0] = "S-AF", [1] = "S-AF", [2] = "S-AF+MF", [3] = "C-AF", [4] = "C-AF+MF", [5] = "C-AF+TR", [6] = "C-AF+TR+MF", [10] = "MF"
    }
    return format_focus_mode_shared(val, m)
end

function Formatters.focus_mode_om1(val)
    if not val then return nil end
    local s = tostring(val)
    -- Multi-component tag Handling (e.g., "10 84")
    local n1, n2 = s:match("(%d+)%s+(%d+)")
    local mode_val = tonumber(n2) or tonumber(n1) or tonumber(s:match("%d+"))
    if not mode_val then return s end
    
    local m = {
        [16] = "MF",
        [81] = "S-AF",
        [84] = "C-AF",
        [100] = "C-AF+TR",
        [51] = "S-AF (瞳優先)", 
        [54] = "C-AF+TR", 
        [592] = "星空AF",
        [112] = "プリセットMF"
    }
    return m[mode_val] or s
end
function Formatters.stacked_image(val)
    if not val then return nil end
    local v = tostring(val)
    local m = {
        ["0 0"] = "通常撮影",
        ["3 2"] = "ライブND (ND2)", ["3 4"] = "ライブND (ND4)", ["3 8"] = "ライブND (ND8)", ["3 16"] = "ライブND (ND16)", ["3 32"] = "ライブND (ND32)", ["3 64"] = "ライブND (ND64)", ["3 128"] = "ライブND (ND128)",
        ["5 4"] = "HDR1", ["6 4"] = "HDR2", 
        ["8 8"] = "三脚ハイレゾショット",
        ["11 12"] = "手持ちハイレゾショット(12枚)", ["11 16"] = "手持ちハイレゾショット(16枚)",
        ["13 2"] = "ライブGND (GND2)", ["13 4"] = "ライブGND (GND4)", ["13 8"] = "ライブGND (GND8)"
    }
    if m[v] then return m[v] end
    
    local maj, min = v:match("^(%d+)%s+(%d+)$")
    
    -- Heuristic for packed u32
    if maj and tonumber(maj) > 65535 then
        local n1 = tonumber(maj)
        local n2 = tonumber(min)
        -- LE Packing: n1 = (High << 16) | Low
        local high = bit.rshift(n1, 16)
        local low = bit.band(n1, 0xFFFF)
        local l2 = (n2 and bit.band(n2, 0xFFFF)) or 0
        
        -- Normal Shot Indicator (0xFFFC = -4) in Low Word
        if low == 65532 then return "なし" end
        
        -- Use High Word as Major ID
        local maj_cand = high
        local min_cand = low 
        
        local k = maj_cand .. " " .. min_cand
        if m[k] then return m[k] end
        


        min = tostring(l2)
        
        if maj_cand == 1 then 
             if tonumber(min) > 0 then return "ライブコンポジット(" .. min .. "Images)" else return "なし" end
        end
        if maj_cand == 4 then 
             if tonumber(min) > 0 then return "ライブタイム・バルブ(" .. min .. " images)" else return "なし" end
        end
        if maj_cand == 9 then return "フォーカスブラケット(" .. min .. "images)" end

    end
    
    if maj then
        local min_n = tonumber(min) or 0
        if maj == "1" then 

            if min_n > 0 then return "ライブコンポジット(" .. min .. "Images)" else return "なし" end
        end
        if maj == "4" then 
            if min_n > 0 then return "ライブタイム・バルブ(" .. min .. " images)" else return "なし" end
        end
        if maj == "9" then return "フォーカスブラケット(" .. min .. "images)" end
    end
    return val
end
function Formatters.om_firmware(val)
    local v = tostring(val)
    local parts = {}
    for p in v:gmatch("%d+") do table.insert(parts, p) end
    if #parts >= 3 then
       local v1 = tonumber(parts[1])
       local v3 = tonumber(parts[3])
       if v1 and v3 then
           local maj = bit.rshift(v1, 8)
           local min = bit.rshift(v3, 8)
           return maj .. "." .. min
       end
    end
    local n = tonumber(v:match("%d+"))
    if not n then return val end
    local h = string.format("%x", n)
    if #h < 4 then 
        if n > 1000 then return string.sub(v, 1, 1) .. "." .. string.sub(v, 2, 2) end
        return val 
    end
    return string.sub(h, 1, 1) .. "." .. string.sub(h, 2, 2)
end
function Formatters.om_picture_mode(val)
    if not val then return nil end
    local s = tostring(val)
    -- Hybrid mapping: User Samples (1-8, 256) + ExifTool Profiles (12-19)
    local m = {
        ["1"]="Vivid", ["2"]="Natural", ["3"]="Flat", ["4"]="Portrait", ["5"]="i-Finish", ["6"]="Monotone", ["7"]="Color Creator", ["8"]="Underwater", ["256"]="Monotone",
        ["9"]="e-Portrait", ["10"]="Art Filter", 
        ["12"]="Monotone Profile 1", ["13"]="Monotone Profile 2", ["14"]="Monotone Profile 3", ["15"]="Monotone Profile 4", 
        ["16"]="Color Profile 1", ["17"]="Color Profile 2", ["18"]="Color Profile 3", ["19"]="Color Profile 4"
    }
    for p in s:gmatch("%d+") do
        local k = tostring(tonumber(p))
        if m[k] then return m[k] end
    end
    return val
end
function Formatters.om_art_filter(val)
    if not val then return nil end
    local s = tostring(val)
    local parts = {}
    for p in s:gmatch("%S+") do table.insert(parts, p) end
    
    local m = {
        [0]="なし", [1]="ポップアートI", [2]="ポップアートII", [3]="ファンタジックフォーカス", 
        [4]="デイドリームI", [5]="デイドリームII", [6]="ライトトーン", 
        [7]="ラフモノクロームI", [8]="ラフモノクロームII", [9]="トイフォトI", 
        [10]="トイフォトII", [11]="トイフォトIII", [12]="ジオラマI", 
        [13]="ジオラマII", [14]="クロスプロセスI", [15]="クロスプロセスII", 
        [16]="ジェントルセピア", [17]="ドラマチックトーンI", [18]="ドラマチックトーンII", 
        [19]="リーニュクレールI", [20]="リーニュクレールII", 
        [21]="ウォーターカラーI", [22]="ウォーターカラーII", 
        [23]="ヴィンテージI", [24]="ヴィンテージII", [25]="ヴィンテージIII", 
        [26]="パートカラーI", [27]="パートカラーII", [28]="パートカラーIII", 
        [29]="ブリーチバイパスI", [30]="ブリーチバイパスII", 
        [31]="ネオノスタルジー"
    }
    
    for _, p in ipairs(parts) do 
        local n = tonumber(p)
        if n and m[n] then return m[n] end 
    end
    return val
end

function Formatters.om_first_number(val)
    if not val then return nil end
    local first_num = string.match(tostring(val), "^%-?%d+")
    return first_num and tonumber(first_num) or val
end

function Formatters.om_noise_filter(val)
    if not val then return nil end
    local first = tonumber(string.match(tostring(val), "^%-?%d+"))
    if not first then return val end
    local m = {[-2] = "Off", [-1] = "Low", [0] = "Standard", [1] = "High"}
    return m[first] or val
end

function Formatters.om_signed_number(val)
    if not val then return nil end
    local n = tonumber(string.match(tostring(val), "^%-?%d+"))
    if not n then return val end
    if n > 0 then return "+" .. n else return tostring(n) end
end

function Formatters.om_tone_control(val)
    if not val then return nil end
    local s = tostring(val)
    local h, m, sh = 0, 0, 0
    local found = false
    
    if s:find(";") then
        -- Format: Highlights; 0; -7; 7; Shadows; 0; -7; 7; Midtones; 0; -7; 7
        local h_str = s:match("Highlights;%s*([^;]+)")
        local sh_str = s:match("Shadows;%s*([^;]+)")
        local m_str = s:match("Midtones;%s*([^;]+)")
        if h_str then h = tonumber(h_str) or 0; found = true end
        if sh_str then sh = tonumber(sh_str) or 0; found = true end
        if m_str then m = tonumber(m_str) or 0; found = true end
    else
        -- Format: -31999 2 -14 14 -31998 0 -14 14 -31997 1 14 14
        local parts = {}
        for p in s:gmatch("%-?%d+") do table.insert(parts, tonumber(p)) end
        for i=1, #parts-3, 4 do
            local id = parts[i]
            local v  = parts[i+1]
            if id == -31999 then h = v; found = true
            elseif id == -31998 then sh = v; found = true
            elseif id == -31997 then m = v; found = true
            end
        end
    end
    
    if not found then return nil end
    
    local function fmt_val(n)
        return (n > 0 and "+" or "") .. n
    end
    return string.format("(%s,%s,%s)", fmt_val(h), fmt_val(m), fmt_val(sh))
end

-- EXIF Parser
local function ParseTIFF(reader, tiff_start, results, initial_table, initial_context, root_tiff_start, depth, visited)
    if not root_tiff_start then root_tiff_start = tiff_start end
    depth = depth or 0
    visited = visited or {}
    
    if depth > 3 then return end
    
    reader:seek(tiff_start)
    local bo = reader:u16()
    if bo == 0x4949 then reader:set_endian(true) elseif bo == 0x4D4D then reader:set_endian(false) else return end
    
    local magic = reader:u16()
    -- Standard is 42. OM System uses 0x04. Panasonic uses 0x55. Be lenient for MakerNotes.
    if magic ~= 42 and magic ~= 0x4 and magic ~= 0x55 then 

    end
    
    local ifd_offset = reader:u32()
    
    -- Fix for OM System invalid offset (e.g. 0x02000008 -> 8)
    if ifd_offset > #reader.data then
         if (ifd_offset % 65536) < #reader.data then
             ifd_offset = ifd_offset % 65536
         end
    end
    
    


    local function ReadIFD(offset, tag_table, vendor_context, nest_level, mn_base)
        nest_level = nest_level or 0
        if nest_level > 20 then return end
        if offset == 0 then return end
        
        
        local file_offset = tiff_start + offset
        if visited[file_offset] then return end 
        visited[file_offset] = true
        
        if file_offset > #reader.data or file_offset < 1 then
             -- Try Root Base (if offset is relative to Main Header)
             local root_file_offset = root_tiff_start + offset
             if root_file_offset <= #reader.data and root_file_offset > 0 then
                 -- Adjusted base logic: We effectively seek to root_file_offset.
                 -- But we need to keep `tiff_start` consistent for subsequent relative offsets?
                 -- Actually if this IFD is relative to Root, its SubIFDs likely are too?
                 -- For now, just correct the seek position. 
                 file_offset = root_file_offset
                 -- Note: If we found it via Root, we might want to flag it?
             else
                 return -- Invalid
             end
        end

        reader:seek(file_offset)
        
        -- MakerNote Header Detection for Vendor Switching
        if vendor_context == "Detect" then
            local check_pos = tiff_start + offset
            local max_check = math.min(#reader.data, check_pos + 12)
            local header_sig = string.sub(reader.data, check_pos, max_check)
            local header_sig = string.sub(reader.data, check_pos, max_check)
            
            if string.find(header_sig, "Panasonic") or string.find(header_sig, "LEICA") then
                offset = offset + 12
                tag_table = MAKERNOTE_TAGS.Panasonic
                vendor_context = "Panasonic"
            -- Olympus: "OLYMP\0" (8 bytes) or "OLYMPUS\0\0" (12 bytes)
            elseif string.find(header_sig, "OLYMP") or string.find(header_sig, "OM SYSTEM") or string.find(header_sig, "OM Digital") then
                if not mn_base then mn_base = offset end
                if string.find(header_sig, "OM SYSTEM") then
                     offset = offset + 16
                elseif string.find(header_sig, "OLYMPUS") or string.find(header_sig, "OM Digital") then 
                     offset = offset + 12 
                else 
                     offset = offset + 8 
                end
                tag_table = MAKERNOTE_TAGS.Olympus
                vendor_context = "Olympus"
            else
                 -- Fallback: If no header known, check Make tag? 
                 -- For now, if no header, assume it's directly the IFD (Canon/Nikon Type3)
                 -- Checking "Make" from main IFD is better, but here we are local.
                 -- If the first tag matches a known pattern?
                 -- Just assume Canon or Nikon Type 3 if we can't tell, or use fallback lookup?
            -- Simplified: If results["Make"] contains "Canon", use Canon table.
                  if type(make_val) == "string" then
                      if make_val:find("OM Digital") or make_val:find("OLYMPUS") then
                          tag_table = MAKERNOTE_TAGS.Olympus
                          vendor_context = "Olympus"
                      end
                  end
             end
            -- print("[DEBUG] MakerNote Header: " .. (header_sig:gsub("%z", ".")) .. " | Vendor: " .. (vendor_context or "Unknown"))
            
            -- Check for embedded TIFF structure (Common in OM System / Fujifilm / Etc)
            reader:seek(tiff_start + offset)
            if reader.pos + 2 <= #reader.data then
                local bo_check = reader:u16()
                if bo_check == 0x4949 or bo_check == 0x4D4D then
                     -- print("[DEBUG] Embedded TIFF found in MakerNote. Recursing with new base.")
                     ParseTIFF(reader, tiff_start + offset, results, tag_table, vendor_context, root_tiff_start, depth + 1, visited)
                     return
                end
            end
            reader:seek(tiff_start + offset)
        end
        
        local count = reader:u16()
        if count > 500 then return end

        for i=1, count do
            local tag = reader:u16()
            local type = reader:u16()
            local cnt = reader:u32()
            local val_off = reader:u32()
            local next_p = reader.pos
            

            
            local val = nil
            local type_sz = {1,1,2,4,8,1,1,2,4,8,4,8,4} -- 13 is IFD (4 bytes)
            local sz = (type_sz[type] or 1) * cnt
            
            if sz > 4 then reader:seek(tiff_start + (mn_base or 0) + val_off) else reader:seek(next_p - 4) end
            
            if type == 2 then 
                if sz > 0 and reader.pos+sz <= #reader.data then
                    val = string.sub(reader.data, reader.pos, reader.pos+sz-1):gsub("[%z%s]+$", "")
                end
            elseif type == 1 or type == 6 then
                if cnt > 1 and cnt < 100 then
                    local vals = {}
                    for k=1, cnt do table.insert(vals, type == 6 and reader:i8() or reader:u8()) end
                    val = table.concat(vals, " ")
                else
                    val = type == 6 and reader:i8() or reader:u8()
                end
            elseif type == 3 or type == 8 then 
                if cnt > 1 and cnt < 100 then
                    local vals = {}
                    for k=1, cnt do table.insert(vals, type == 8 and reader:i16() or reader:u16()) end
                    val = table.concat(vals, " ")
                else
                    val = type == 8 and reader:i16() or reader:u16() 
                end
            elseif type == 4 or type == 9 then 
                if cnt > 1 and cnt < 100 then
                    local vals = {}
                    for k=1, cnt do table.insert(vals, type == 9 and reader:i32() or reader:u32()) end
                    val = table.concat(vals, " ")
                else
                    val = type == 9 and reader:i32() or reader:u32()
                end
            elseif type == 5 or type == 10 then 
                local n, d
                if type == 10 then
                    n, d = reader:i32(), reader:i32()
                else
                    n, d = reader:u32(), reader:u32()
                end
                -- Use tag_table definition if available
                local def = tag_table and tag_table[tag]
                if def and def.fmt and Formatters[def.fmt] then val = Formatters[def.fmt](n, d)
                else val = Formatters.fraction_reduce(n, d) end
            elseif type == 7 then 
                 if sz > 0 and sz < 128 then
                    val = string.sub(reader.data, reader.pos, reader.pos+sz-1):gsub("[%z%s]+$", "")
                    if val:match("[^%g%s]") then val = nil end
                 end
            end
            
            -- Recursion for ExifOffset / GPS / MakerNote / SubIFDs
            if tag == 0x8769 then ReadIFD(val_off, STANDARD_EXIF_TAGS, "Exif", nest_level + 1) -- ExifOffset
            elseif tag == 0x8825 then ReadIFD(val_off, STANDARD_EXIF_TAGS, "GPS", nest_level + 1) -- GPS
            elseif tag == 0x927C then 
                ReadIFD(val_off, nil, "Detect", nest_level + 1) -- MakerNote
            elseif vendor_context == "Olympus" and (type == 13 or (tag >= 0x2010 and tag <= 0x2050)) then
                local next_table = tag_table
                local next_context = vendor_context
                if tag == 0x2010 then next_table = MAKERNOTE_TAGS.OlympusEq; next_context = "Equipment"
                elseif tag == 0x2020 then next_table = MAKERNOTE_TAGS.OlympusCS; next_context = "CameraSettings"
                elseif tag == 0x2030 then next_table = MAKERNOTE_TAGS.OlympusRD; next_context = "RawDevelopment"
                elseif tag == 0x2031 then next_table = MAKERNOTE_TAGS.OlympusRD2; next_context = "RawDevelopment"
                elseif tag == 0x2040 then next_table = MAKERNOTE_TAGS.OlympusIP; next_context = "ImageProcessing"
                elseif tag == 0x2050 then next_table = MAKERNOTE_TAGS.OlympusFI; next_context = "FocusInfo"
                end
                -- IFD offsets in Olympus MakerNotes are relative to mn_base
                ReadIFD((mn_base or 0) + val_off, next_table, next_context, nest_level + 1, mn_base)
            end

            -- Value mapping
            local def = tag_table and tag_table[tag]
            if def then 
                if def.map and val and type ~= 5 and type ~= 10 then
                    local map_val = val
                    if _G.type(val) == "string" then
                        local first_num = string.match(val, "^%-?%d+")
                        if first_num then map_val = tonumber(first_num) end
                    end
                    if def.map[tonumber(map_val)] then
                        val = def.map[tonumber(map_val)]
                    end
                elseif def.fmt and val and type ~= 5 and type ~= 10 and Formatters[def.fmt] then
                    val = Formatters[def.fmt](val)
                end
            end
            
            -- Emergency Fallbacks for Olympus tags
            if tag == 0x0529 then val = Formatters.om_art_filter(val) end
            if tag == 0x0520 then val = Formatters.om_picture_mode(val) end
            if tag == 0x0104 then val = Formatters.om_firmware(val) end
            if tag == 0x0804 then val = Formatters.stacked_image(val) end
            
            if def and def.key and val then
                -- IFD Priority: RawDevelopment(4) > ImageProcessing(3) > CameraSettings(2) > Equipment(2) > Root(1)
                local priorities = { ["Root"]=1, ["Equipment"]=2, ["CameraSettings"]=2, ["RawDevelopment"]=4, ["ImageProcessing"]=3, ["FocusInfo"]=5 }
                local context_p = priorities[vendor_context] or 0
                local current = results[def.key]
                
                if not current or context_p >= (current.prio or 0) then
                    results[def.key] = {label=def.name, value=val, prio=context_p}
                end
            end
            reader:seek(next_p)
        end
    end
    ReadIFD(ifd_offset, initial_table or STANDARD_EXIF_TAGS, initial_context or "Root", 0)
end

local function ScanNRAWFile(path)
    local r = BinaryReader.new(path)
    if not r then return {} end
    
    local NRAW_TAGS_MAP = {
        [0x0002] = {key = "Model", name = "モデル"},
        [0x0003] = {key = "Firmware", name = "ファームウェア"},
        [0x0005] = {key = "WhiteBalanceTemperature", name = "ホワイトバランス"},
    }
    
    local results = {}
    local sig = string.char(0x4E, 0x43, 0x54, 0x47) -- "NCTG"
    local idx = r:find(sig)
    if not idx then return {} end
    
    local pos = idx + 4
    while pos < #r.data - 8 do
        r:seek(pos)
        r:set_endian(false)
        local header = r:u16()
        
        if header == 0x0110 or header == 0x0210 or header == 0x0200 then
            local tag = r:u16()
            local type_ = r:u16()
            local count = r:u16()
            
            if type_ < 1 or type_ > 12 or count > 512 then
                pos = pos + 1
            else
                local data_size = 0
                if type_ == 1 or type_ == 2 or type_ == 7 then data_size = count
                elseif type_ == 3 then data_size = count * 2
                elseif type_ == 4 or type_ == 9 then data_size = count * 4
                elseif type_ == 5 or type_ == 10 then data_size = count * 8
                else data_size = count end
                
                local start_data = r.pos
                if start_data + data_size - 1 > #r.data then
                     pos = pos + 1
                else
                    local val = nil
                    if type_ == 2 then
                        if data_size > 0 then
                            val = r.data:sub(r.pos, r.pos + data_size - 1):gsub("[%z%s]+$", "")
                        end
                    elseif type_ == 5 or type_ == 10 then
                        local n = r:u32()
                        local d = r:u32()
                        local def = NRAW_TAGS_MAP[tag] or STANDARD_EXIF_TAGS[tag]
                        if def and def.fmt and Formatters[def.fmt] then val = Formatters[def.fmt](n, d)
                        else val = Formatters.fraction_reduce(n, d) end
                    elseif type_ == 3 or type_ == 8 then
                        val = count > 1 and "Array" or r:u16()
                    elseif type_ == 4 or type_ == 9 then
                        val = count > 1 and "Array" or r:u32()
                    end
                    
                    local def = NRAW_TAGS_MAP[tag] or STANDARD_EXIF_TAGS[tag]
                    if def and def.key and val then
                        results[def.key] = {label = def.name, value = val, prio = 1}
                    end
                    pos = start_data + data_size
                end
            end
        else
            pos = pos + 1
        end
    end
    
    -- 3. Scrape static strings for Model (Nikon NEV specific focus)
    if idx and path:lower():match("%.nev$") then
        local chunk = r.data:sub(idx + 4, math.min(#r.data, idx + 4096))
        local m_maker, m_model = chunk:match("(NIKON)%s+(Z%s-[A-Za-z0-9_]+)")
        if m_maker and m_model then
            local val = (m_maker .. " " .. m_model):gsub("[%z]+$", ""):gsub("%s+$", "")
            if #val > 6 and #val < 32 then
                results["Model"] = {label = "モデル", value = val, prio = 2}
            end
        else
            local model_only = chunk:match("(NIKON%s+Z%s-[A-Za-z0-9_]+)")
            if model_only then
                 local val = model_only:gsub("[%z]+$", ""):gsub("%s+$", "")
                 if #val > 6 and #val < 32 then
                     results["Model"] = {label = "モデル", value = val, prio = 2}
                 end
            end
        end
    end
    
    return results
end

local function ScanFile(path)
    local r = BinaryReader.new(path)
    if not r then return {} end
    
    local best_res = {}
    local max_tags = 0
    
    local function TryParse(offset, skip)
         local temp_res = {}
         ParseTIFF(r, offset + skip, temp_res)
         local cnt = 0
         for _ in pairs(temp_res) do cnt = cnt + 1 end
         if cnt > max_tags then
             max_tags = cnt
             best_res = temp_res
         end
    end

    local sig = string.char(0x45,0x78,0x69,0x66,0x00,0x00)
    local start_p = 1
    while true do
        local p = r:find(sig, start_p)
        if not p then break end
        TryParse(p, 6)
        start_p = p + 1
    end
    
    local sig_le = string.char(0x49,0x49,0x2A,0x00)
    local sig_be = string.char(0x4D,0x4D,0x00,0x2A)
    local start_le, start_be = 1, 1
    
    while true do
        local p = r:find(sig_le, start_le)
        if not p then break end
        TryParse(p, 0)
        start_le = p + 1
    end
    
    while true do
        local p = r:find(sig_be, start_be)
        if not p then break end
        TryParse(p, 0)
        start_be = p + 1
    end

    return best_res
end

local KEY_TO_LABELS = {}
local function BuildLabelMap()
    local function Add(tbl)
        for _, def in pairs(tbl) do
             if def.key then
                 KEY_TO_LABELS[def.key] = {jp = def.name, en = def.name_en or def.key or def.name}
             end
        end
    end
    Add(STANDARD_EXIF_TAGS)
    for _, t in pairs(MAKERNOTE_TAGS) do Add(t) end
    
    for _, def in ipairs(API_PROPS) do
        -- For API, key is e.g. "Video Codec"
        -- We will use "API_" .. def.key as lookup key
        local k = "API_" .. def.key
        KEY_TO_LABELS[k] = {jp = def.name, en = def.name_en or def.key}
    end
    
    -- Special Layout Items
    KEY_TO_LABELS["RET"] = {jp = "[ ↵ 改行 ]", en = "[ ↵ Return ]"}
    KEY_TO_LABELS["BLANK"] = {jp = "[ □ 空白 ]", en = "[ □ Space ]"}
    KEY_TO_LABELS["BAR"] = {jp = "[ | 縦線 ]", en = "[ | Bar ]"}
    KEY_TO_LABELS["CUSTOM"] = {jp = "[ ✎ カスタム ]", en = "[ ✎ Custom ]"}
end
BuildLabelMap()

--------------------------------------------------------------------------------
-- 3. UI & APP LOGIC
--------------------------------------------------------------------------------

local win = disp:AddWindow({
    ID = "MyWin",
    WindowTitle = "Exif & Metadata Importer v1.0.0",
    Geometry = {100, 100, 500, 680},
    Spacing = 8,
    
    ui:VGroup{
        -- HEADER
        ui:VGroup{
            Weight = 0,
            ui:HGroup{
                Weight = 0,
                ui:Label{ID = "LblTitle", Text = "<h3>Metadata Importer</h3>", Weight = 1},
                ui:Label{ID = "LblStatus", Text = "Ready", Alignment = {AlignHRight = true}, Weight = 2},
            },
            ui:VGap(4),
            ui:HGroup{
                Weight = 0,
                ui:Label{Text = "Mode:", Weight = 0},
                ui:ComboBox{ID = "ComboMode", Text = "モード選択", Weight = 1.5},
                ui:ComboBox{ID = "ComboColor", Text = "Color", Weight = 1, Enabled = false},
            },
            ui:VGap(4),
            -- PRESETS
            ui:HGroup{
                Weight = 0,
                ui:Label{Text = "Preset:", Weight = 0},
                ui:ComboBox{ID = "ComboPresets", Weight = 1.5},
                ui:LineEdit{ID = "LinePresetName", PlaceholderText = "New Preset Name", Weight = 1.5},
            },
            ui:HGroup{
                Weight = 0,
                ui:Button{ID = "BtnLoadPreset", Text = "Load", Weight = 1},
                ui:Button{ID = "BtnSavePreset", Text = "Save", Weight = 1},
                ui:Button{ID = "BtnDeletePreset", Text = "Delete", Weight = 1},
            },
            ui:VGap(4),
            ui:Button{ID = "BtnLoad", Text = "情報読み込み (API & EXIF)", Height = 30},
            ui:HGroup{
                 Weight = 0,
                 ui:Button{ID = "BtnSelectAll", Text = "全選択", Weight = 1},
                 ui:Button{ID = "BtnSelectNone", Text = "全解除", Weight = 1},
                 ui:Button{ID = "BtnSortChecked", Text = "選択項目を上へ", Weight = 1},
            },
            ui:VGap(4),
        },

        
        -- MAIN LIST & SIDE TOOLBAR
        ui:HGroup{
            Weight = 1,
            ui:Tree{ID = "TreeExif", SelectionMode = "Single", SortingEnabled = false, Weight = 1}, 
            
            ui:VGroup{
                Weight = 0,
                ui:VGap(0, 1),
                ui:Label{Text = "移動", Alignment = {AlignHCenter=true}},
                ui:Button{ID = "BtnUp", Text = "↑", Width=22, Height=30},
                ui:Button{ID = "BtnDown", Text = "↓", Width=22, Height=30},
                ui:VGap(10),
                ui:VGap(10),
                ui:Label{Text = "追加", Alignment = {AlignHCenter=true}},
                ui:Button{ID = "BtnAddRet", Text = "＋改行", Width=50, Height=30}, -- Newline
                ui:Button{ID = "BtnAddBlank", Text = "＋空白", Width=50, Height=30}, -- Spacer
                ui:Button{ID = "BtnAddBar", Text = "＋縦線", Width=50, Height=30}, -- Bar
                ui:VGap(5),
                ui:LineEdit{ID = "LineCustom", PlaceholderText = "自由入力", Width=50, Height=24},
                ui:Button{ID = "BtnAddCustom", Text = "＋追加", Width=50, Height=30}, -- Custom
                ui:VGap(10),
                ui:Button{ID = "BtnRemove", Text = "削除", Width=50, Height=30}, -- Remove Item
                ui:VGap(0, 1),
            }
        },
        
        -- FOOTER
        ui:VGroup{
            Weight = 0,
            ui:VGap(5),

            ui:HGroup{
                Weight = 0,
                ui:CheckBox{ID = "ChkLabel", Text = "ラベル名を含める (例: ISO: 800)", Checked = true},
                ui:HGap(10),
                ui:CheckBox{ID = "ChkEnglish", Text = "English Labels", Checked = false},
            },
            ui:CheckBox{ID = "ChkPadding", Text = "文字数（幅）を統一する", Checked = false},

            ui:HGroup{
                Weight = 0,
                ui:Button{ID = "BtnWrite", Text = "一括書き込み (Text+)", Height = 40, Weight = 3},
                ui:Button{ID = "BtnClose", Text = "閉じる", Height = 40, Weight = 1},
            },
        },
    }
})

local itm = win:GetItems()

-- Init
itm.ComboMode:AddItem("再生ヘッド直下 (最前面)")
itm.ComboMode:AddItem("クリップカラー (一括)")
local clip_colors = {"Orange", "Apricot", "Yellow", "Lime", "Olive", "Green", "Teal", "Cyan", "Blue", "Purple", "Violet", "Pink", "Tan", "Beige", "Brown", "Chocolate"}
for _, c in ipairs(clip_colors) do itm.ComboColor:AddItem(c) end

local display_data = {} 
local ui_items = {}
local loaded_config = LoadConfig()
local current_presets = loaded_config.presets
local current_options = loaded_config.options
loaded_config.last_state = {}

-- Init Options
if current_options.pad_width then itm.ChkPadding.Checked = true end
if current_options.use_english then itm.ChkEnglish.Checked = true end

-- Init Presets
local function RefreshPresetList()
    itm.ComboPresets:Clear()
    local names = {}
    for k in pairs(current_presets) do table.insert(names, k) end
    table.sort(names)
    for _, name in ipairs(names) do itm.ComboPresets:AddItem(name) end
end
RefreshPresetList()


-- UI Refresh
local function RefreshTree(select_index)
    itm.TreeExif:Clear()
    ui_items = {} 
    
    local header = itm.TreeExif:NewItem()
    header.Text[0] = "項目名"
    header.Text[1] = "値"
    header.Text[2] = "ソース/ID"
    itm.TreeExif:SetHeaderItem(header)
    itm.TreeExif:SetColumnCount(3)
    itm.TreeExif.ColumnWidth[0] = 130
    itm.TreeExif.ColumnWidth[1] = 150
    itm.TreeExif.ColumnWidth[2] = 60
    
    for i, data in ipairs(display_data) do
        local item = itm.TreeExif:NewItem()
        item.Text[0] = data.label
        item.Text[1] = tostring(data.value)
        item.Text[2] = data.id_str
        
        if data.checked then item.CheckState[0] = "Checked"
        else item.CheckState[0] = "Unchecked" end
        
        -- Style for layout objects
        if data.id_str == "RET" or data.id_str == "BLANK" or data.id_str == "BAR" then
             item.TextColor[0] = {R=0.3, G=0.3, B=1.0, A=1.0} -- Highlight Blueish
        elseif data.id_str == "CUSTOM" then
             item.TextColor[0] = {R=0.2, G=0.8, B=0.2, A=1.0} -- Highlight Greenish
        end
        
        itm.TreeExif:AddTopLevelItem(item)
        table.insert(ui_items, item)
        
        if i == select_index then
            item.Selected = true
            itm.TreeExif.CurrentItem = item
            if itm.TreeExif.ScrollToItem then
                itm.TreeExif:ScrollToItem(item, 0) -- EnsureVisible
            end
        end
    end
end

local function SyncCheckStates()
    for i, item in ipairs(ui_items) do
        if display_data[i] then
            display_data[i].checked = (item.CheckState[0] == "Checked")
        end
    end
end

local function GetSelectedIndex()
    for i, item in ipairs(ui_items) do
        if item.Selected then return i end
    end
    return nil -- Default append to end
end

-- Helpers
-- Helpers
local function TimecodeToFrame(tc, fps, force_df)
    if not tc then return 0 end
    local h, m, s, sep, f = tc:match("(%d+):(%d+):(%d+)([:;])(%d+)")
    if not h then return 0 end
    
    local h_n, m_n, s_n, f_n = tonumber(h), tonumber(m), tonumber(s), tonumber(f)
    
    -- Round FPS for Timebase
    local timebase = math.floor(fps + 0.5)
    
    -- Detect Drop Frame from separator OR force flag
    -- Semicolon usually indicates Drop Frame
    local is_drop_frame = force_df or (sep == ";")

    -- Drop Frame Calculation (Typically for 29.97 or 59.94)
    if is_drop_frame then
        local total_minutes = h_n * 60 + m_n
        local total_frames = total_minutes * 60 * timebase + s_n * timebase + f_n
        
        -- Drop 2 frames every minute, except every 10th minute
        local drop_count = 2 
        if timebase == 60 then drop_count = 4 end 
        
        local dropped_frames = drop_count * (total_minutes - math.floor(total_minutes / 10))
        return total_frames - dropped_frames
    end

    return (h_n * 3600 + m_n * 60 + s_n) * timebase + f_n
end

local function FrameToTimecode(frame, fps)
    local timebase = math.floor(fps + 0.5)
    local total_seconds = math.floor(frame / timebase)
    local f = math.floor(frame % timebase)
    local s = math.floor(total_seconds % 60)
    local m = math.floor((total_seconds / 60) % 60)
    local h = math.floor(total_seconds / 3600)
    return string.format("%02d:%02d:%02d:%02d", h, m, s, f)
end

local function FindTextPlusTool(comp)
    if not comp then return nil end
    
    -- 1. Try specific "Template" name (Standard for Text+ Titles)
    local t = comp:FindTool("Template")
    if t then return t end
    
    -- 2. Try "Text1"
    t = comp:FindTool("Text1")
    if t then return t end
    
    -- 3. Fallback: Scan for any TextPlus type
    local tools = comp:GetToolList(false, "TextPlus")
    if tools and #tools > 0 then return tools[1] end
    
    return nil
end



local function GetCurrentProjectAndTimeline()
    -- Refresh Resolve Object
    resolve = bmd and bmd.scriptapp("Resolve") or resolve
    
    local pm = resolve:GetProjectManager()
    local project = pm:GetCurrentProject()
    if not project then return nil, nil end
    local timeline = project:GetCurrentTimeline()
    if timeline then project:SetCurrentTimeline(timeline) end
    return project, timeline
end

local function SetTrackLock(timeline, track_type, index, locked)
    if timeline.SetTrackLock then
        return timeline:SetTrackLock(track_type, index, locked)
    end
    return false
end

local function GetTrackLock(timeline, track_type, index)
    if timeline.GetTrackLock then
        return timeline:GetTrackLock(track_type, index)
    end
    return false -- Default assumption if cannot read
end

local function IsTrackEnabled(timeline, track_type, index)
    if timeline.GetIsTrackEnabled then
        return timeline:GetIsTrackEnabled(track_type, index)
    elseif timeline.GetTrackEnabled then
        return timeline:GetTrackEnabled(track_type, index)
    end
    return true -- Logic: Default to true if API not found to avoid blocking
end

local function IsClipEnabled(clip)
    if clip.GetClipEnabled then
        local status = clip:GetClipEnabled()
        if status ~= nil then return status end
    end
    
    local props = clip:GetClipProperty()
    if props then
        if props["ClipEnabled"] ~= nil then return props["ClipEnabled"] end
        if props["Enabled"] ~= nil then return props["Enabled"] end
    end
    
    return true -- Default to true if cannot determine
end

local function GetTargetClips()
    local project, timeline = GetCurrentProjectAndTimeline()
    if not timeline then return {}, {} end
    
    local results = {}
    local touched_tracks = {} -- Store track indices strings to avoid dups via keys
    
    if itm.ComboMode.CurrentIndex == 0 then
        -- Playhead Mode
        local fps_str = timeline:GetSetting("timelineFrameRate") or "24"
        local fps = tonumber(fps_str)
        local is_df_setting = timeline:GetSetting("timelineDropFrameTimecode")
        local is_drop_frame = (is_df_setting == "1" or is_df_setting == 1 or is_df_setting == true)
        
        local tc = timeline:GetCurrentTimecode()
        -- Pass force_df as logic OR separator check in function
        local current_frame = TimecodeToFrame(tc, fps, is_drop_frame)
        
        print(string.format("[DEBUG] TC: %s | FPS: %s | DF: %s | Frame: %d", tc, fps_str, tostring(is_drop_frame), current_frame))
        
        local track_count = timeline:GetTrackCount("video")
        
        for t = track_count, 1, -1 do
            if IsTrackEnabled(timeline, "video", t) then 
                local items = timeline:GetItemListInTrack("video", t)
                if items then
                    for i, clip in ipairs(items) do
                        -- Strict Check
                        local start_f = clip:GetStart()
                        local end_f = clip:GetEnd()
                        
                        -- print(string.format("[DEBUG] Check Track %d Clip %d: Start %d End %d", t, i, start_f, end_f))
                        
                        if IsClipEnabled(clip) and current_frame >= start_f and current_frame < end_f then
                            print(string.format("[DEBUG] FOUND Match! Track %d Item %d", t, i))
                            table.insert(results, clip)
                            touched_tracks[t] = true
                            -- Continue searching other tracks instead of returning immediately
                        end
                    end
                end
            end
        end
    else
        -- Color Mode
        local target_color = itm.ComboColor.CurrentText
        local track_count = timeline:GetTrackCount("video")
        for t = 1, track_count do
            if IsTrackEnabled(timeline, "video", t) then
                local items = timeline:GetItemListInTrack("video", t)
                if items then
                    for i, clip in ipairs(items) do
                        if IsClipEnabled(clip) and clip:GetClipColor() == target_color then 
                            table.insert(results, clip)
                            touched_tracks[t] = true
                        end
                    end
                end
            end
        end
    end
    return results, touched_tracks
end

-- Handlers
function win.On.ComboMode.CurrentIndexChanged(ev)
    itm.ComboColor.Enabled = (ev.Index == 1)
end

function win.On.BtnLoad.Clicked(ev)
    local clips = GetTargetClips()
    if #clips == 0 then
        itm.LblStatus.Text = "<font color='red'>Not Found</font>"
        return
    end
    
    -- Smart Select: Find first clip with a valid File Path
    local selected_clip = clips[1]
    local selected_path = ""
    
    for _, c in ipairs(clips) do
        local mi = c:GetMediaPoolItem()
        local p = mi and mi:GetClipProperty("File Path") or ""
        if p and p ~= "" then
            selected_clip = c
            selected_path = p
            break
        end
    end
    
    local clip = selected_clip
    local mediaItem = clip:GetMediaPoolItem()
    local path = selected_path 
    
    print("[DEBUG] Selected Clip Path: " .. path)
    
    itm.LblStatus.Text = "Scanning..."
    
    local ex_data = {}
    if path ~= "" then 
        if path:lower():match("%.nev$") or path:lower():match("%.r3d$") then
            ex_data = ScanNRAWFile(path)
        else
            ex_data = ScanFile(path)
        end
    end
    
    display_data = {}
    local found_data_map = {}
    
    -- 1. Gather all DEFINED tags first (Standard + MakerNotes) - RB Mode Priority
    local preferred_keys = {
        Model=true, LensModel=true, FocalLength=true, FNumber=true, ExposureTime=true, 
        ISO=true, ExposureBiasValue=true, CreateDate=true, FocusMode=true, PictureMode=true,
        WhiteBalanceTemperature=true, AISubjectTrackingMode=true, ImageStabilization=true,
        StackedImage=true, ArtFilter=true
    }
    
    -- Iterate Standard
    for id, def in pairs(STANDARD_EXIF_TAGS) do
        if ex_data[def.key] then
            local lbl = def.name
            if itm.ChkEnglish.Checked then lbl = def.name_en or def.key or def.name end
            
            local item = {
                label = lbl,
                value = ex_data[def.key].value,
                id_str = def.key,
                checked = preferred_keys[def.key] or false
            }
            if not found_data_map[item.id_str] then found_data_map[item.id_str] = item end
        end
    end
    -- Iterate MakerNotes
    for vendor, tags in pairs(MAKERNOTE_TAGS) do
        for id, def in pairs(tags) do
             if ex_data[def.key] then
                local lbl = def.name
                if itm.ChkEnglish.Checked then lbl = def.name_en or def.key or def.name end

                local item = {
                    label = lbl,
                    value = ex_data[def.key].value,
                    id_str = def.key,
                    checked = preferred_keys[def.key] or false
                }
                if not found_data_map[item.id_str] then found_data_map[item.id_str] = item end
            end
        end
    end
    -- 1.5. Gather remaining items from ex_data (for NRAW/custom tags)
    for k, data in pairs(ex_data) do
        if not found_data_map[k] then
            local item = {
                label = data.label or k,
                value = data.value,
                id_str = k,
                checked = preferred_keys[k] or false
            }
            if k == "Model" or k == "WhiteBalanceTemperature" then item.checked = true end
            found_data_map[k] = item
        end
    end
    
    -- 2. Gather from API for remaining/missing fields - Fallback Role
    if mediaItem then
        local props = mediaItem:GetClipProperty()
        if type(props) == "table" then
            for k, val in pairs(props) do
                if val and val ~= "" then
                    local api_key = "API_" .. k
                    if not found_data_map[api_key] and not found_data_map[k] then
                        local lbl = k
                        for _, p in ipairs(API_PROPS) do
                            if p.key == k then lbl = p.name; break end
                        end
                        if itm.ChkEnglish.Checked then lbl = k end
                        
                        local item = {
                            label = lbl,
                            value = val,
                            id_str = api_key,
                            checked = true
                        }
                        found_data_map[api_key] = item
                    end
                end
            end
        else
            -- フォールバック
            for _, prop in ipairs(API_PROPS) do
                local api_key = "API_" .. prop.key
                if not found_data_map[api_key] and not found_data_map[prop.key] then
                    local val = mediaItem:GetClipProperty(prop.key)
                    if val and val ~= "" then
                        local lbl = prop.name
                        if itm.ChkEnglish.Checked then lbl = prop.name_en or prop.key end
                        
                        local item = {
                            label = lbl,
                            value = val,
                            id_str = api_key,
                            checked = true
                        }
                        found_data_map[api_key] = item
                    end
                end
            end
        end
    end
    
    -- 2. Load Config from last_state
    local saved_order = loaded_config.last_state
    local used_keys = {}
    local count = 0
    
    if saved_order and type(saved_order) == "table" then
        for _, saved_item in ipairs(saved_order) do
            if saved_item.id_str == "RET" or saved_item.id_str == "BLANK" or saved_item.id_str == "BAR" or saved_item.id_str == "CUSTOM" then
                -- For CUSTOM, we need to ensure 'value' is restored
                if saved_item.id_str == "CUSTOM" and not saved_item.value then
                    saved_item.value = saved_item.label:match("%[ ✎ (.*) %]") or ""
                end
                table.insert(display_data, saved_item)
            else
                local key = saved_item.id_str
                -- Compatibility fix for old API generic ID
                if key == "API" then 
                     -- Old format: key="API", label="FPS" -> Try mapping "API_FPS"?
                     -- Or just ignore old bad data?
                     -- Let's try to infer from label if possible, but label might be Japanese based on when saved.
                     -- Actually, old `found_data_map` keys were "API_"..name.
                     -- New keys are "API_"..key. Often same.
                     -- If we can't find it, we skip.
                end
                
                -- Check standard map
                if found_data_map[key] and not used_keys[key] then
                    local real_item = found_data_map[key]
                    real_item.checked = saved_item.checked
                    table.insert(display_data, real_item)
                    used_keys[key] = true
                    count = count + 1
                end
            end
        end
    end

    -- 3. Append new items
    for _, prop in ipairs(API_PROPS) do
        local key = "API_" .. prop.key
        if found_data_map[key] and not used_keys[key] then
            table.insert(display_data, found_data_map[key])
            used_keys[key] = true
            count = count + 1
        end
    end
    
    
    local sorted_tags = {}
    -- Helper to collect
    local function Collect(tbl)
        for _, def in pairs(tbl) do
            table.insert(sorted_tags, def)
        end
    end
    Collect(STANDARD_EXIF_TAGS)
    for _, vendor_table in pairs(MAKERNOTE_TAGS) do
        Collect(vendor_table)
    end
    -- Sort by KEY to make finding easier, or Label?
    -- User probably likes Label order? Or ID?
    -- Previous code sorted by ID (Hex).
    -- Since we have duplicates ID, sorting by ID is confusing.
    -- Let's sort by Name (Label).
    table.sort(sorted_tags, function(a,b) return a.name < b.name end)
    
    for _, def in ipairs(sorted_tags) do
        local key = def.key
        if found_data_map[key] and not used_keys[key] then
            table.insert(display_data, found_data_map[key])
            used_keys[key] = true
            count = count + 1
        end
    end
    
    RefreshTree()
    if #clips > 1 then
        itm.LblStatus.Text = string.format("Found %d clips (Show #1)", #clips)
    else
        itm.LblStatus.Text = count.." Items"
    end
end

local function UpdateDisplayLanguage(is_english)
    for _, item in ipairs(display_data) do
        local def_map = KEY_TO_LABELS[item.id_str]
        if def_map then
            item.label = is_english and def_map.en or def_map.jp
        end
    end
    RefreshTree()
end

function win.On.ChkEnglish.Clicked(ev)
    UpdateDisplayLanguage(itm.ChkEnglish.Checked)
end

function win.On.BtnAddRet.Clicked(ev)
    SyncCheckStates()
    local idx = GetSelectedIndex()
    local new_item = {
        label = "[ ↵ 改行 ]",
        value = "（次の行へ）",
        id_str = "RET",
        checked = true
    }
    
    if idx then
        table.insert(display_data, idx + 1, new_item)
        RefreshTree(idx + 1)
    else
        table.insert(display_data, new_item)
        RefreshTree(#display_data)
    end
end

function win.On.BtnAddBlank.Clicked(ev)
    SyncCheckStates()
    local idx = GetSelectedIndex()
    local new_item = {
        label = "[ □ 空白 ]",
        value = "（スペース）",
        id_str = "BLANK",
        checked = true
    }
    
    if idx then
        table.insert(display_data, idx + 1, new_item)
        RefreshTree(idx + 1)
    else
        table.insert(display_data, new_item)
        RefreshTree(#display_data)
    end
end

function win.On.BtnAddBar.Clicked(ev)
    SyncCheckStates()
    local idx = GetSelectedIndex()
    local new_item = {
        label = "[ | 縦線 ]",
        value = "（ | ）",
        id_str = "BAR",
        checked = true
    }
    
    if idx then
        table.insert(display_data, idx + 1, new_item)
        RefreshTree(idx + 1)
    else
        table.insert(display_data, new_item)
        RefreshTree(#display_data)
    end
end

function win.On.BtnAddCustom.Clicked(ev)
    local text = itm.LineCustom.Text
    if not text or text == "" then return end
    
    SyncCheckStates()
    local idx = GetSelectedIndex()
    local new_item = {
        label = "[ ✎ " .. text .. " ]",
        value = text,
        id_str = "CUSTOM",
        checked = true
    }
    
    if idx then
        table.insert(display_data, idx + 1, new_item)
        RefreshTree(idx + 1)
    else
        table.insert(display_data, new_item)
        RefreshTree(#display_data)
    end
    itm.LineCustom.Text = "" -- Clear input
end

function win.On.BtnRemove.Clicked(ev)
    local idx = GetSelectedIndex()
    if idx then
        local item = display_data[idx]
        if item.id_str == "RET" or item.id_str == "BLANK" or item.id_str == "BAR" or item.id_str == "CUSTOM" then
             SyncCheckStates()
             table.remove(display_data, idx)
             if #display_data == 0 then
                 RefreshTree()
             else
                 if idx > #display_data then idx = #display_data end
                 RefreshTree(idx)
             end
             itm.LblStatus.Text = "Deleted Item"
        else
             itm.LblStatus.Text = "<font color='red'>Cannot delete EXIF/API items</font>"
        end
    end
end

function win.On.BtnUp.Clicked(ev)
    local idx = GetSelectedIndex()
    if idx and idx > 1 then
        SyncCheckStates()
        local temp = display_data[idx]
        display_data[idx] = display_data[idx-1]
        display_data[idx-1] = temp
        RefreshTree(idx-1)
    end
end

function win.On.BtnDown.Clicked(ev)
    local idx = GetSelectedIndex()
    if idx and idx > 0 and idx < #display_data then
        SyncCheckStates()
        local temp = display_data[idx]
        display_data[idx] = display_data[idx+1]
        display_data[idx+1] = temp
        RefreshTree(idx+1)
    end
end

function win.On.BtnSelectAll.Clicked(ev)
    for i, d in ipairs(display_data) do d.checked = true end
    RefreshTree()
end

function win.On.BtnSelectNone.Clicked(ev)
    for i, d in ipairs(display_data) do d.checked = false end
    RefreshTree()
end

function win.On.BtnSortChecked.Clicked(ev)
    SyncCheckStates()
    local checked = {}
    local unchecked = {}
    for _, d in ipairs(display_data) do
        if d.checked then
            table.insert(checked, d)
        else
            table.insert(unchecked, d)
        end
    end
    display_data = {}
    for _, d in ipairs(checked) do table.insert(display_data, d) end
    for _, d in ipairs(unchecked) do table.insert(display_data, d) end
    RefreshTree()
    itm.LblStatus.Text = "Sorted Checked Items"
end

function win.On.BtnClose.Clicked(ev) 
    SyncCheckStates()
    current_options.pad_width = (itm.ChkPadding.Checked == "Checked") or itm.ChkPadding.Checked
    current_options.use_english = (itm.ChkEnglish.Checked == "Checked") or itm.ChkEnglish.Checked
    SaveConfig({}, current_presets, current_options)
    disp:ExitLoop() 
end

function win.On.BtnSavePreset.Clicked(ev)
    local name = itm.LinePresetName.Text
    if not name or name == "" then
        name = itm.ComboPresets.CurrentText
    end
    if name and name ~= "" then
        SyncCheckStates()
        local preset_data = {}
        for _, d in ipairs(display_data) do
             if d.id_str == "CUSTOM" then
                 -- For Custom items, likely want to save the Value text
                 table.insert(preset_data, {
                     id_str = d.id_str,
                     label = d.label,
                     value = d.value, -- Save user text
                     checked = d.checked
                 })
             else
                 table.insert(preset_data, {
                     id_str = d.id_str,
                     label = d.label,
                     checked = d.checked
                 })
             end
        end
        current_presets[name] = preset_data
        RefreshPresetList()
        itm.LblStatus.Text = "Saved: " .. name
    else
        itm.LblStatus.Text = "Enter Name"
    end
end

function win.On.BtnLoadPreset.Clicked(ev)
    local name = itm.ComboPresets.CurrentText
    if current_presets[name] then
        loaded_config.last_state = current_presets[name]
        itm.LblStatus.Text = "Loaded: " .. name .. " (Click Load info)"
        win.On.BtnLoad.Clicked({})
    end
end

function win.On.BtnDeletePreset.Clicked(ev)
    local name = itm.ComboPresets.CurrentText
    if name and current_presets[name] then
        current_presets[name] = nil
        RefreshPresetList()
        itm.LblStatus.Text = "Deleted: " .. name
    end
end

function win.On.BtnWrite.Clicked(ev)
    local project, timeline = GetCurrentProjectAndTimeline()
    if not timeline then return end
    
    SyncCheckStates()
    
    -- Template data check
    local has_checked = false
    for _, d in ipairs(display_data) do if d.checked then has_checked = true break end end
    if not has_checked then return end
    
    local clips, source_tracks = GetTargetClips()
    if #clips == 0 then return end
    
    local fps_str = timeline:GetSetting("timelineFrameRate") or "24"
    local fps = tonumber(fps_str)
    
    -- AUTO-LOCK LOGIC START
    local track_count = timeline:GetTrackCount("video")
    local audio_track_count = timeline:GetTrackCount("audio")
    local lock_states = {}
    local audio_lock_states = {}
    
    local lock_success = pcall(function()
        -- 1. Save Video States & Lock All
        for t = 1, track_count do
            lock_states[t] = GetTrackLock(timeline, "video", t)
            SetTrackLock(timeline, "video", t, true) -- Lock everything first
        end
        
        -- 2. Save Audio States & Lock All (全オーディオトラックをロック)
        for t = 1, audio_track_count do
            audio_lock_states[t] = GetTrackLock(timeline, "audio", t)
            SetTrackLock(timeline, "audio", t, true)
        end
        
        local target_write_track = track_count -- Assume top
        
        if source_tracks[track_count] then
            for t = 1, track_count do
                if source_tracks[t] then
                    SetTrackLock(timeline, "video", t, true)
                else
                    SetTrackLock(timeline, "video", t, false)
                end
            end
        else
             SetTrackLock(timeline, "video", track_count, false)
        end
    end)
    -- AUTO-LOCK LOGIC END
    
    local success_count = 0
    

    
    local main_success, main_err = pcall(function()
    
    -- Store metadata values locally
    local clip_meta_data = {} 
    local file_cache = {}
    
    for idx, target_clip in ipairs(clips) do
        local mediaItem = target_clip:GetMediaPoolItem()
        local path = mediaItem and mediaItem:GetClipProperty("File Path") or ""
        
        local ex_data = {}
        if path ~= "" then 
            if file_cache[path] then
                ex_data = file_cache[path]
            else
                if path:lower():match("%.nev$") or path:lower():match("%.r3d$") then
                    ex_data = ScanNRAWFile(path)
                else
                    ex_data = ScanFile(path)
                end
                file_cache[path] = ex_data
            end
        end
        
        local raw_values = {} 
        for i, tmpl in ipairs(display_data) do
            if tmpl.checked then
                 local val = nil
                 if tmpl.id_str == "RET" or tmpl.id_str == "BLANK" or tmpl.id_str == "BAR" then
                 elseif tmpl.id_str == "CUSTOM" then
                     val = tmpl.value -- Use stored custom text
                 elseif string.match(tmpl.id_str, "^API_") then
                     local api_key = string.sub(tmpl.id_str, 5)
                     val = mediaItem and mediaItem:GetClipProperty(api_key) or "none"
                 else
                     -- Lookup in ex_data using the Key (stored in id_str)
                     local key = tmpl.id_str
                     if ex_data[key] then
                         val = ex_data[key].value
                     end
                 end
                 if not val or val == "" then val = "none" end
                 raw_values[i] = val
            end
        end
        clip_meta_data[idx] = raw_values
    end
    
    -- CALCULATE MAX LENGTHS
    local max_lengths = {} 
    if itm.ChkPadding.Checked then
        for _, values in pairs(clip_meta_data) do
            for i, val in pairs(values) do
                 local ulen = utf8 and utf8.len(val) or string.len(val)
                if not max_lengths[i] or ulen > max_lengths[i] then
                    max_lengths[i] = ulen
                end
            end
        end
    end
    
    for idx, target_clip in ipairs(clips) do
        local values = clip_meta_data[idx]
        if values then
            local lines = {}
            for i, tmpl in ipairs(display_data) do
                if tmpl.checked then
                    if tmpl.id_str == "RET" then
                         table.insert(lines, "\n")
                    elseif tmpl.id_str == "BLANK" then
                         table.insert(lines, " ") 
                    elseif tmpl.id_str == "BAR" then
                         table.insert(lines, "|") 
                    else
                        local val = values[i]
                        if itm.ChkPadding.Checked and max_lengths[i] then
                             local ulen = utf8 and utf8.len(val) or string.len(val)
                             local pad = max_lengths[i] - ulen
                             if pad > 0 then val = val .. string.rep(" ", pad) end
                        end
                        local txt_segment = (itm.ChkLabel.Checked and tmpl.id_str ~= "CUSTOM" and (tmpl.label..": "..val)) or val
                        table.insert(lines, txt_segment)
                    end
                end
            end
            local final_txt = table.concat(lines, "")
            
            local start_frame = target_clip:GetStart()
            local start_tc = FrameToTimecode(start_frame, fps)
            timeline:SetCurrentTimecode(start_tc)
            
            local new_clip = nil
            if timeline.InsertFusionTitleIntoTimeline then
                new_clip = timeline:InsertFusionTitleIntoTimeline("Text+")
            end
            if not new_clip and timeline.InsertTitleIntoTimeline then
                 new_clip = timeline:InsertTitleIntoTimeline("Text+")
                 if not new_clip then new_clip = timeline:InsertTitleIntoTimeline("Text") end
            end
            
            if new_clip then
                pcall(function()
                     new_clip:SetClipColor("Tan")
                     local comp = new_clip:GetFusionCompByIndex(1)
                     if comp then
                         local tool = FindTextPlusTool(comp)
                         if tool then
                             if tool.SetInput then tool:SetInput("StyledText", final_txt) end
                             tool.StyledText = final_txt
                             -- Removed keyframing to prevent timing drift
                             tool.Size = 0.04
                         end
                     end
                    local dur = target_clip:GetDuration()
                    new_clip:SetStart(start_frame)
                    new_clip:SetDuration(dur)
                end)
                success_count = success_count + 1
            end
        end
    end
    end) -- main_success pcall end

    if not main_success then
        print("Processing interrupted: " .. tostring(main_err))
    end

    -- RESTORE LOCK STATE
    if lock_success then
        -- Restore Video
        for t = 1, track_count do
            SetTrackLock(timeline, "video", t, lock_states[t])
        end
        -- Restore Audio
        for t = 1, audio_track_count do
            SetTrackLock(timeline, "audio", t, audio_lock_states[t])
        end
    end
    
    itm.LblStatus.Text = string.format("Finished: %d/%d", success_count, #clips)
end

function win.On.MyWin.Close(ev)
    SyncCheckStates()
    current_options.pad_width = (itm.ChkPadding.Checked == "Checked") or itm.ChkPadding.Checked
    current_options.use_english = (itm.ChkEnglish.Checked == "Checked") or itm.ChkEnglish.Checked
    SaveConfig({}, current_presets, current_options)
    disp:ExitLoop()
end



win:Show()
disp:RunLoop()
win:Hide()