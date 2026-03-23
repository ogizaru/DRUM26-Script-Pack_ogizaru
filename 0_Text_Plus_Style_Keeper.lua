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
    Text+ Style Keeper v1.0.0
    For DaVinci Resolve Edit Page
    Author: OGIZARU & DaVinci Resolve Addon and DCTL maker V3 (Custom Gem)
    Date: 2026-01-05
]]--

local ui = fu.UIManager
local disp = bmd.UIDispatcher(ui)

-- [UI FIX] Adjusted height to 500px for a more compact UI
local width, height = 400, 500

-- ============================================================================
-- GLOBAL VARIABLES
-- ============================================================================
local selectedStyleNameDisplay = nil
local HEX_PREFIX = "Hex_"

-- [BATCH COPY CONSTANTS]
local COLOR_LIST = {
    "Orange", "Apricot", "Yellow", "Lime", "Olive", "Green", "Teal", 
    "Navy", "Blue", "Purple", "Violet", "Pink", "Tan", "Beige", "Brown", "Chocolate"
}


-- ============================================================================
-- PATH & ENCODING HELPERS
-- ============================================================================
local platform = (package.config:sub(1,1) == '\\') and "Windows" or "Mac"
local SEP = (platform == "Windows") and "\\" or "/"

local function getScriptPath()
    local info = debug.getinfo(1, "S")
    local source = info.source
    if source:sub(1, 1) == "@" then
        source = source:sub(2)
    end
    if platform == "Windows" then
        return source:match("(.*\\)")
    else
        return source:match("(.*/)")
    end
end

local BASE_DIR = getScriptPath()
if not BASE_DIR or BASE_DIR == "" then
    BASE_DIR = fu:MapPath("Scripts:/Utility/")
end

local STYLE_DIR = BASE_DIR .. "TextPlusStyles"

local function toHex(str)
    return (str:gsub('.', function (c)
        return string.format('%02X', string.byte(c))
    end))
end

local function fromHex(str)
    return (str:gsub('..', function (cc)
        return string.char(tonumber(cc, 16))
    end))
end

local function getFileNameFromInput(inputName)
    return HEX_PREFIX .. toHex(inputName) .. ".style"
end

local function getDisplayNameFromFile(filename)
    local namePart = filename:match("(.+)%.style$") or filename
    if namePart:sub(1, #HEX_PREFIX) == HEX_PREFIX then
        local hexPart = namePart:sub(#HEX_PREFIX + 1)
        return fromHex(hexPart)
    else
        return namePart
    end
end

local function ensureDirectoryExists(path)
    local hasLfs, lfs = pcall(require, "lfs")
    if hasLfs and lfs then
        if lfs.attributes(path, "mode") ~= "directory" then
            lfs.mkdir(path)
        end
        return
    end

    -- Fast native check to prevent os.execute (which flashes a cmd window)
    local dummy = path .. SEP .. ".dir_test"
    local f = io.open(dummy, "w")
    if f then
        f:close()
        os.remove(dummy)
        return -- Directory already exists
    end

    -- Attempt to create silently using internal API if mapped
    if type(bmd) == "table" and type(bmd.createdir) == "function" then
        pcall(bmd.createdir, path)
        local f2 = io.open(dummy, "w")
        if f2 then f2:close(); os.remove(dummy); return end
    end

    local cmd
    if platform == "Windows" then
        cmd = 'if not exist "' .. path .. '" mkdir "' .. path .. '"'
    else
        cmd = 'mkdir -p "' .. path .. '"'
    end
    os.execute(cmd)
end

if STYLE_DIR and STYLE_DIR ~= "" then
    ensureDirectoryExists(STYLE_DIR)
end

-- ============================================================================
-- UTILITY FUNCTIONS
-- ============================================================================

local function fileExists(path)
    local f = io.open(path, "rb")
    if f then
        f:close()
        return true
    else
        return false
    end
end

local function updateStyleIndexFile(filesTable)
    local indexPath = STYLE_DIR .. SEP .. "index.txt"
    local f = io.open(indexPath, "wb")
    if f then
        for _, filename in ipairs(filesTable) do
            f:write(filename .. "\n")
        end
        f:close()
    end
end

local function getStyleFiles(forceRefresh)
    local files = {}
    local indexPath = STYLE_DIR .. SEP .. "index.txt"
    
    -- Try loading from index file to avoid io.popen flashes on Windows
    if not forceRefresh then
        local f = io.open(indexPath, "rb")
        if f then
            local content = f:read("*a")
            f:close()
            local indexValid = true
            for filename in content:gmatch("[^\r\n]+") do
                if filename:match("%.style$") and fileExists(STYLE_DIR .. SEP .. filename) then
                    table.insert(files, filename)
                end
            end
            return files
        end
    end

    local hasLfs, lfs = pcall(require, "lfs")
    if hasLfs and lfs then
        for filename in lfs.dir(STYLE_DIR) do
            if filename:match("%.style$") then
                table.insert(files, filename)
            end
        end
        updateStyleIndexFile(files)
        return files
    end

    -- Internal BMD fallback just in case
    if type(bmd) == "table" and type(bmd.readdir) == "function" then
        local ok, dirs = pcall(bmd.readdir, STYLE_DIR .. SEP .. "*")
        if ok and type(dirs) == "table" then
            for _, v in ipairs(dirs) do
                local name = type(v) == "table" and v.Name or v
                if type(name) == "string" and name:match("%.style$") then
                    table.insert(files, name)
                end
            end
            updateStyleIndexFile(files)
            return files
        end
    end

    local pfile
    if platform == "Windows" then
        pfile = io.popen('dir "' .. STYLE_DIR .. '\\*.style" /b')
    else
        pfile = io.popen('ls "' .. STYLE_DIR .. '"/*.style')
    end

    if pfile then
        for filename in pfile:lines() do
            if platform ~= "Windows" then
                filename = filename:match("^.+/(.+)$") or filename
            end
            if filename and filename:match("%.style$") then
                table.insert(files, filename)
            end
        end
        pfile:close()
    end
    
    -- Save index for next time to prevent flashing
    updateStyleIndexFile(files)
    
    return files
end

local function getToolAndSyncFocus()
    local resolve = Resolve()
    local project = resolve:GetProjectManager():GetCurrentProject()
    if not project then return nil, nil, "エラー: プロジェクトが開かれていません。" end
    
    local timeline = project:GetCurrentTimeline()
    if not timeline then return nil, nil, "エラー: タイムラインが開かれていません。" end
    
    project:SetCurrentTimeline(timeline)
    
    local item = timeline:GetCurrentVideoItem()
    if not item then return nil, nil, "エラー: クリップが選択されていません。" end

    if item.GetFusionCompCount and item:GetFusionCompCount() == 0 then return nil, nil, "エラー: クリップにFusionコンポジションがありません。" end
    
    local comp = item:GetFusionCompByIndex(1)
    if not comp then return nil, nil, "エラー: Fusionコンポジションを取得できません。" end
    
    local tool = comp:FindTool("Template")
    if not tool then
        local toolList = comp:GetToolList(false, "TextPlus")
        if toolList then
            for _, t in pairs(toolList) do
                tool = t
                break
            end
        end
    end
    
    if not tool then return nil, nil, "エラー: 選択されたクリップはText+ではありません。" end
    
    return tool, comp, nil
end

-- ============================================================================
-- BATCH COPY HELPERS & FIXES
-- ============================================================================

-- Deep Copy Table
local function DeepCopy(orig)
    local orig_type = type(orig)
    local copy
    if orig_type == 'table' then
        copy = {}
        for orig_key, orig_value in next, orig, nil do
            copy[DeepCopy(orig_key)] = DeepCopy(orig_value)
        end
        setmetatable(copy, DeepCopy(getmetatable(orig)))
    else -- number, string, boolean, etc
        copy = orig
    end
    return copy
end

local function FindTextPlusTool(comp)
    if not comp then return nil end
    local tools = comp:GetToolList(false, "TextPlus")
    for id, tool in pairs(tools) do
        return tool
    end
    local allTools = comp:GetToolList(false)
    for id, tool in pairs(allTools) do
        local name = tool.Name
        if string.find(name, "Template") or string.find(name, "Text") then
            return tool
        end
    end
    return nil
end

local function GetToolSettings(tool)
    if tool.GetSettings then return tool:GetSettings() end
    if tool.SaveSettings then return tool:SaveSettings() end
    return nil
end

-- File-Based Style Application (Fix for Shading 2+ Corruption)
local function ApplyStyleViaTempFile(targetTool, sourceTool)
    -- 1. Setup paths (Unique to bypass cache)
    math.randomseed(os.time())
    local uniqueID = os.time() .. "_" .. math.random(100000)
    local tempName = "temp_style_transfer_" .. uniqueID .. ".setting"
    local tempPath = STYLE_DIR .. SEP .. tempName
    
    -- 2. Save Source to File
    if not sourceTool.SaveSettings then return false end
    sourceTool:SaveSettings(tempPath)
    
    if not fileExists(tempPath) then return false end
    
    -- 3. Read File Content
    local f = io.open(tempPath, "r")
    if not f then return false end
    local content = f:read("*all")
    f:close()
    
    -- 4. Patch Tool Name
    -- Pattern: Look for the first key in the Tools table. 
    -- Typically: Tools = ordered() { [SourceToolName] = TextPlus {
    -- We want to replace [SourceToolName] with [TargetToolName]
    
    -- Escape magic characters in names for pattern matching
    local escapedSourceName = sourceTool.Name:gsub("([%^%$%(%)%%%.%[%]%*%+%-%?])", "%%%1")
    
    -- Try specific replacement first
    local patchedContent, count = content:gsub("([\r\n]%s*)%[" .. escapedSourceName .. "%](%s*=%s*TextPlus)", "%1[" .. targetTool.Name .. "]%2")
    
    -- If strict replacement failed (unlikely if sourceTool.Name is correct), try broad replacement
    if count == 0 then
         patchedContent = content:gsub("([\r\n]%s*)%[.-%](%s*=%s*TextPlus)", "%1[" .. targetTool.Name .. "]%2", 1)
    end
    
    -- 5. Write Patched Content
    f = io.open(tempPath, "w")
    if not f then return false end
    f:write(patchedContent)
    f:close()
    
    -- 6. Load to Target
    local success = false
    if targetTool.LoadSettings then
        targetTool:LoadSettings(tempPath)
        success = true
    end
    
    -- 7. Cleanup
    os.remove(tempPath)
    
    return success
end


-- Variant for applying from an existing .style file
local function ApplyFileWithPatching(targetTool, filepath)
    if not fileExists(filepath) then return false end
    
    -- 1. Read File Content
    local f = io.open(filepath, "r")
    if not f then return false end
    local content = f:read("*all")
    f:close()
    
    -- 2. Find Source Name in File (First key in Tools table)
    -- Pattern: Tools = ordered() { [SourceToolName] = TextPlus
    local sourceName = content:match("[\r\n]%s*%[([\"']?.-[\"']?)%]%s*=%s*TextPlus")
    
    -- If we found a name, we can patch it
    local patchedContent = content
    if sourceName then
        if sourceName ~= ('"' .. targetTool.Name .. '"') and sourceName ~= ("'" .. targetTool.Name .. "'") and sourceName ~= targetTool.Name then
             -- Escape magic characters
            local escapedSourceName = sourceName:gsub("([%^%$%(%)%%%.%[%]%*%+%-%?])", "%%%1")
            patchedContent = content:gsub("([\r\n]%s*)%[" .. escapedSourceName .. "%](%s*=%s*TextPlus)", "%1[" .. targetTool.Name .. "]%2")
        end
    else
        -- Fallback: Use blind replacement if pattern match failed for some reason
        patchedContent = content:gsub("([\r\n]%s*)%[.-%](%s*=%s*TextPlus)", "%1[" .. targetTool.Name .. "]%2", 1)
    end
    
    -- 3. Write to Temp (Unique to bypass cache)
    math.randomseed(os.time())
    local uniqueID = os.time() .. "_" .. math.random(100000)
    local tempName = "temp_style_apply_" .. uniqueID .. ".setting"
    local tempPath = STYLE_DIR .. SEP .. tempName
    
    f = io.open(tempPath, "w")
    if not f then return false end
    f:write(patchedContent)
    f:close()
    
    -- 4. Load
    local success = false
    if targetTool.LoadSettings then
        targetTool:LoadSettings(tempPath)
        success = true
    end
    
    os.remove(tempPath)
    return success
end





-- ============================================================================
-- MODAL DIALOGS
-- ============================================================================
local function showConfirmDialog(title, message, onConfirm)
    local dWidth, dHeight = 350, 150
    local dialog = disp:AddWindow({
        ID = 'ConfirmWin',
        WindowTitle = title,
        Geometry = { 
            (width - dWidth)/2 + 100, (height - dHeight)/2 + 100, 
            dWidth, dHeight 
        },
        WindowFlags = { Window = true, Dialog = true, Modal = true, CustomizeWindowHint = true },
        
        ui:VGroup{
            ID = 'root',
            ui:Label{ Text = message, Alignment = {AlignHCenter = true, AlignVCenter = true}, WordWrap = true, Weight = 1 },
            ui:HGroup{
                Weight = 0,
                ui:Button{ ID = 'BtnYes', Text = 'はい / 確認' },
                ui:Button{ ID = 'BtnNo', Text = 'キャンセル' },
            }
        }
    })
    
    function dialog.On.BtnYes.Clicked(ev)
        dialog:Hide()
        onConfirm()
    end
    function dialog.On.BtnNo.Clicked(ev)
        dialog:Hide()
    end
    function dialog.On.ConfirmWin.Close(ev)
        dialog:Hide()
    end
    dialog:Show()
end

-- ============================================================================
-- MAIN LOGIC
-- ============================================================================

local function saveCurrentStyle(displayName, force)
    local filename = getFileNameFromInput(displayName)
    local filepath = STYLE_DIR .. SEP .. filename
    
    if not force and fileExists(filepath) then
        showConfirmDialog(
            "上書き確認", 
            "スタイル '" .. displayName .. "' はすでに存在します。\n上書きしますか？", 
            function() saveCurrentStyle(displayName, true) end
        )
        return nil
    end

    local tool, comp, err = getToolAndSyncFocus()
    if err then return err end
    
    if not tool.SaveSettings then return "エラー: ツールがSaveSettingsをサポートしていません。" end
    
    tool:SaveSettings(filepath)
    
    if fileExists(filepath) then
        return "成功: '" .. displayName .. "' を保存しました"
    else
        return "エラー: スタイルファイルの書き込みに失敗しました。"
    end
end

local function applySelectedStyle(displayName)
    local tool, comp, err = getToolAndSyncFocus()
    if err then return err end
    
    local filename = getFileNameFromInput(displayName)
    local filepath = STYLE_DIR .. SEP .. filename
    if not fileExists(filepath) then
        local legacyPath = STYLE_DIR .. SEP .. displayName .. ".style"
        if fileExists(legacyPath) then
            filepath = legacyPath
        else
            return "エラー: スタイルファイルが見つかりません。"
        end
    end
    
    local currentText = ""
    if tool.GetInput then
        currentText = tool:GetInput("StyledText") or ""
    end
    
    comp:StartUndo("Apply Text+ Style")
    
    -- Use File Patching to avoid corruption
    if not ApplyFileWithPatching(tool, filepath) then
        comp:EndUndo(true)
        return "エラー: スタイルファイルの適用に失敗しました。"
    end
    
    if tool.SetInput then
        local dummyText = " "
        if currentText == " " then dummyText = "" end
        tool:SetInput("StyledText", dummyText)
        tool:SetInput("StyledText", currentText)
    end
    
    comp:EndUndo(true)
    
    return "成功: '" .. displayName .. "' を適用しました"
end

local function exportAllStyles()
    local targetDir = fu:RequestDir()
    if not targetDir or targetDir == "" then return "エクスポートをキャンセルしました。" end
    
    local files = getStyleFiles()
    local count = 0
    
    for _, f in ipairs(files) do
        local src = STYLE_DIR .. SEP .. f
        local dstName = f 
        local dst = targetDir .. dstName
        
        local lastChar = string.sub(targetDir, -1)
        if lastChar ~= "/" and lastChar ~= "\\" then
            dst = targetDir .. SEP .. dstName
        end
        
        -- Native Lua binary file copy
        local infile = io.open(src, "rb")
        if infile then
            local content = infile:read("*a")
            infile:close()
            local outfile = io.open(dst, "wb")
            if outfile then
                outfile:write(content)
                outfile:close()
                count = count + 1
            end
        end
    end
    
    return count .. " 個のスタイルをエクスポートしました (Safe Mode)。"
end

local function processBatchCopy(targetColor, sourceTool, sourceStylePath)
    local resolve = Resolve()
    local projectManager = resolve:GetProjectManager()
    local project = projectManager:GetCurrentProject()
    
    if not project then return "エラー: プロジェクトが開かれていません。" end
    local timeline = project:GetCurrentTimeline()
    if not timeline then return "エラー: タイムラインが開かれていません。" end

    -- 1. Validate Source
    if not sourceTool and not sourceStylePath then
        return "重大なエラー: ソースが選択されていません。"
    end
    
    local trackCount = timeline:GetTrackCount("video")
    local successCount = 0
    local failCount = 0

    -- 2. Iterate all clips
    for i = 1, trackCount do
        local items = timeline:GetItemListInTrack("video", i)
        if items then
            for j, item in ipairs(items) do
                if item:GetClipColor() == targetColor then
                    local itemName = item:GetName()
                    local comp = item:GetFusionCompByIndex(1)
                    
                    if comp then
                        local targetTool = FindTextPlusTool(comp)
                        if targetTool then
                            -- A. Save original text
                            local originalText = nil
                            if targetTool.StyledText and targetTool.StyledText[1] then
                                originalText = targetTool.StyledText[1]
                            end
                            
                            -- B. Apply Settings (File-Based)
                            local applied = false
                            if sourceStylePath then
                                applied = ApplyFileWithPatching(targetTool, sourceStylePath)
                            elseif sourceTool then
                                applied = ApplyStyleViaTempFile(targetTool, sourceTool)
                            end
                            
                            -- C. Restore Text
                            if originalText then
                                targetTool.StyledText = originalText
                            end
                            
                            if applied then
                                successCount = successCount + 1
                            else
                                failCount = failCount + 1
                            end
                        else
                            failCount = failCount + 1
                        end
                    else
                        -- No comp
                    end
                end
            end
        end
    end

    return "完了。 成功: " .. successCount .. " / スキップ・失敗: " .. failCount
end


-- ============================================================================
-- GUI
-- ============================================================================
local win = disp:AddWindow({
    ID = 'MyWin',
    WindowTitle = 'Text+ スタイルキーパー v1.0.0',
    Geometry = {100, 100, width, height},
    
    ui:VGroup{
        ID = 'root',
        Spacing = 4,

        -- 1. HEADER & LIST
        ui:Label{ ID = 'LblTitle', Text = '保存済みのスタイル', Weight = 0, Font = ui:Font{PixelSize=14, Bold=true} },
        -- List takes all available expansion space
        ui:Tree{ ID = 'StyleList', Weight = 1, AlternatingRowColors = true, RootIsDecorated = false },
        
        -- 2. BATCH COPY & APPLY
        ui:HGroup{
            Weight = 0, Spacing = 4,
            ui:Label{ Text = 'ソース:', Weight = 0 },
            ui:ComboBox{ ID = 'ComboSource', Weight = 1 },
            ui:Label{ Text = '対象カラー:', Weight = 0 },
            ui:ComboBox{ ID = 'ComboColor', Weight = 1 },
        },
        ui:HGroup{
            Weight = 0, Spacing = 4,
            ui:Button{ 
                ID = 'BtnApply', 
                Text = 'ヘッド位置のクリップに適用', 
                Weight = 1,  
            },
            ui:Button{ 
                ID = 'BtnBatchCopy', 
                Text = '一致するクリップに適用', 
                Weight = 1, 
            },
        },

        -- 3. SAVE/DELETE CONTROLS
        -- Removed the empty VGap(8) from before, and just using VGap(12) to serve as a wider separator.
        ui:VGap(12),
        ui:HGroup{
            Weight = 0, Spacing = 2, -- Reduced spacing to minimal 2 pixels
            ui:Label{ Text = '名前:', Weight = 0 },
            ui:LineEdit{ ID = 'NameEdit', Text = '', PlaceholderText = 'スタイル名を入力...', Weight = 1 },
            ui:HGroup{
                Weight = 0, Spacing = 2,
                ui:Button{ ID = 'BtnSave', Text = '保存', Weight = 0, MinimumSize = {44, 26}, MaximumSize = {44, 26} },
                ui:Button{ ID = 'BtnDelete', Text = '削除', Weight = 0, MinimumSize = {44, 26}, MaximumSize = {44, 26} },
            }
        },
        
        ui:VGap(4),
        -- 4. FOOTER
        ui:HGroup{
            Weight = 0, Spacing = 4,
            ui:Button{ ID = 'BtnExport', Text = 'バックアップを出力' },
            ui:Button{ ID = 'BtnRefresh', Text = '再読み込み' },
        },
        
        -- STATUS LABEL
        ui:Label{ 
            ID = 'StatusLbl', 
            Text = '準備完了', 
            Alignment = {AlignHCenter = true, AlignVCenter = true}, 
            Font = ui:Font{PixelSize=10},
            Weight = 0,
            MaximumSize = {9999, 20} -- Force max height to 20px
        }
    }
})

local itm = win:GetItems()

local function setStatus(msg)
    if msg then itm.StatusLbl.Text = msg end
end

local function refreshList(forceRefresh)
    itm.StyleList:Clear()
    selectedStyleNameDisplay = nil
    itm.NameEdit.Text = ""
    local files = getStyleFiles(forceRefresh)
    
    local header = itm.StyleList:NewItem()
    header.Text[0] = "スタイル名"
    itm.StyleList:SetHeaderItem(header)
    itm.StyleList.ColumnCount = 1
    
    for _, f in ipairs(files) do
        local row = itm.StyleList:NewItem()
        row.Text[0] = getDisplayNameFromFile(f)
        itm.StyleList:AddTopLevelItem(row)
    end
end

function win.On.BtnRefresh.Clicked(ev)
    refreshList(true)
    setStatus("リストを再読み込みしました。")
end

function win.On.StyleList.ItemClicked(ev)
    if ev.item then
        local displayTxt = ev.item.Text[0]
        selectedStyleNameDisplay = displayTxt
        itm.NameEdit.Text = displayTxt
    end
end

function win.On.BtnSave.Clicked(ev)
    local name = itm.NameEdit.Text
    if name == "" then
        setStatus("エラー: 名前が入力されていません。")
        return
    end
    local res = saveCurrentStyle(name, false)
    setStatus(res)
    refreshList(true) -- Force refresh to update index
end

function win.On.BtnApply.Clicked(ev)
    if not selectedStyleNameDisplay then
        setStatus("エラー: スタイルが選択されていません。")
        return
    end
    local res = applySelectedStyle(selectedStyleNameDisplay)
    setStatus(res)
end

function win.On.BtnDelete.Clicked(ev)
    if not selectedStyleNameDisplay then 
        setStatus("エラー: 選択されていません。")
        return 
    end
    
    showConfirmDialog(
        "削除の確認",
        "'" .. selectedStyleNameDisplay .. "' を削除しますか？",
        function()
            local filename = getFileNameFromInput(selectedStyleNameDisplay)
            local filepath = STYLE_DIR .. SEP .. filename
            if not fileExists(filepath) then
                filepath = STYLE_DIR .. SEP .. selectedStyleNameDisplay .. ".style"
            end
            os.remove(filepath)
            setStatus("削除しました。")
            refreshList(true) -- Force refresh to update index
        end
    )
end

function win.On.BtnExport.Clicked(ev)
    local res = exportAllStyles()
    setStatus(res)
end

function win.On.BtnBatchCopy.Clicked(ev)
    local colorName = itm.ComboColor.CurrentText
    if not colorName or colorName == "" then
        setStatus("エラー: カラーが選択されていません。")
        return
    end

    local sourceMode = itm.ComboSource.CurrentIndex
    local sourceTool = nil
    local sourceStylePath = nil
    local confirmMsg = ""

    -- 0: Current track clip, 1: Saved style
    if sourceMode == 0 then
        local tool, comp, err = getToolAndSyncFocus()
        if err then
            setStatus(err)
            return
        end
        sourceTool = tool
        confirmMsg = "'" .. tool.Name .. "' のスタイルを\nすべての '" .. colorName .. "' クリップにコピーしますか？"
    else
        if not selectedStyleNameDisplay then
            setStatus("エラー: コピー元のスタイルが選択されていません。")
            return
        end
        local filename = getFileNameFromInput(selectedStyleNameDisplay)
        sourceStylePath = STYLE_DIR .. SEP .. filename
        if not fileExists(sourceStylePath) then
            sourceStylePath = STYLE_DIR .. SEP .. selectedStyleNameDisplay .. ".style"
            if not fileExists(sourceStylePath) then
                setStatus("エラー: スタイルファイルが見つかりません。")
                return
            end
        end
        confirmMsg = "保存済みスタイル '" .. selectedStyleNameDisplay .. "' を\nすべての '" .. colorName .. "' クリップに適用しますか？"
    end

    showConfirmDialog(
        "バッチコピーの確認",
        confirmMsg .. "\n\n【警告】\n「元に戻す/やり直し」操作を行うと、\nシェーディング要素などが崩れる可能性があります。",
        function()
            local res = processBatchCopy(colorName, sourceTool, sourceStylePath)
            setStatus(res)
        end
    )
end

function win.On.MyWin.Close(ev)
    disp:ExitLoop()
end

-- Initialize Color List
for _, colorName in ipairs(COLOR_LIST) do
    itm.ComboColor:AddItem(colorName)
end

-- Initialize Source List
itm.ComboSource:AddItem("現在のタイムライン位置")
itm.ComboSource:AddItem("選択した保存済みスタイル")

refreshList(true)
win:Show()
disp:RunLoop()
win:Hide()